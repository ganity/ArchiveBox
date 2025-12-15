use anyhow::{anyhow, Context, Result};
use docx_rs::*;
use once_cell::sync::Lazy;
use quick_xml::events::Event;
use quick_xml::Reader;
use regex::Regex;
use rust_xlsxwriter::{Format, FormatAlign, Workbook};
use serde::{Deserialize, Serialize};
use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};
use std::process::Command;
use tauri::{Manager, State};
use time::OffsetDateTime;
use uuid::Uuid;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};
use lopdf::{Document as PdfDocument, Object as PdfObject};

static RE_FIELD: Lazy<Regex> = Lazy::new(|| {
    Regex::new(
        r"(?m)^\s*(指令编号|指令标题|下发时间|指令内容)\s*[:：]\s*(?P<v>.*?)\s*$",
    )
    .expect("valid regex")
});

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
struct WordFields {
    instruction_no: String,
    title: String,
    issued_at: String,
    content: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct ZipSummary {
    id: String,
    filename: String,
    source_path: String,
    stored_path: String,
    extracted_dir: String,
    #[serde(default)]
    include_original_zip: bool,
    status: String,
    word: WordFields,
    has_video: bool,
    has_sample: bool,
    video_entries: Vec<String>,
    video_files: Vec<String>,
    image_files: Vec<String>,
    pdf_files: Vec<String>,
    pdf_image_files: Vec<String>,
    pdf_page_screenshot_files: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct BatchSummary {
    batch_id: String,
    created_at: i64,
    zips: Vec<ZipSummary>,
}

#[derive(Default)]
struct AppState {
    last_batch_id: std::sync::Mutex<Option<String>>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
struct ExportZipSelection {
    zip_id: String,
    include: bool,
    include_original_zip: bool,
    selected_video_indices: Vec<usize>,
    selected_image_indices: Vec<usize>,
    selected_pdf_indices: Vec<usize>,
    selected_pdf_image_indices: Vec<usize>,
    selected_pdf_page_screenshot_indices: Vec<usize>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
struct ExportBundleSelection {
    zips: Vec<ExportZipSelection>,
}

fn app_data_dir(app: &tauri::AppHandle) -> Result<PathBuf> {
    let dir = app
        .path()
        .app_data_dir()
        .context("无法获取AppData目录")?
        .join("ArchiveBox");
    fs::create_dir_all(&dir)?;
    Ok(dir)
}

fn batch_dir(app: &tauri::AppHandle, batch_id: &str) -> Result<PathBuf> {
    let dir = app_data_dir(app)?.join("batches").join(batch_id);
    fs::create_dir_all(&dir)?;
    Ok(dir)
}

fn prompt_save_path(default_name: String, ext: &str, filter_label: &str) -> Result<PathBuf, String> {
    let chosen = rfd::FileDialog::new()
        .add_filter(filter_label, &[ext])
        .set_file_name(&default_name)
        .save_file();
    let Some(path) = chosen else {
        return Err("已取消".to_string());
    };
    let path = ensure_extension(path, ext);
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(err_to_string)?;
    }
    Ok(path)
}

fn ensure_extension(path: PathBuf, ext: &str) -> PathBuf {
    match path.extension().and_then(|e| e.to_str()) {
        Some(_) => path,
        None => path.with_extension(ext),
    }
}

fn default_export_excel_name(now: OffsetDateTime) -> String {
    format!(
        "导出结果_{}{:02}{:02}_{:02}{:02}{:02}.xlsx",
        now.year(),
        now.month() as u8,
        now.day(),
        now.hour(),
        now.minute(),
        now.second()
    )
}

fn default_export_bundle_name(now: OffsetDateTime) -> String {
    format!(
        "汇总包_{}{:02}{:02}_{:02}{:02}{:02}.zip",
        now.year(),
        now.month() as u8,
        now.day(),
        now.hour(),
        now.minute(),
        now.second()
    )
}

#[tauri::command]
fn pick_zip_files() -> Result<Vec<String>, String> {
    let files = rfd::FileDialog::new()
        .add_filter("ZIP", &["zip"])
        .pick_files()
        .unwrap_or_default();
    Ok(files
        .into_iter()
        .map(|p| p.to_string_lossy().to_string())
        .collect())
}

#[tauri::command]
fn import_zips(app: tauri::AppHandle, state: State<'_, AppState>, paths: Vec<String>) -> Result<BatchSummary, String> {
    let now = OffsetDateTime::now_utc();
    let batch_id = format!("batch_{}", now.unix_timestamp());
    let batch_dir = batch_dir(&app, &batch_id).map_err(err_to_string)?;

    let mut zips = Vec::new();
    for p in paths {
        let source_path = PathBuf::from(&p);
        let filename = source_path
            .file_name()
            .and_then(|s| s.to_str())
            .unwrap_or("UNKNOWN.zip")
            .to_string();
        let zip_id = Uuid::new_v4().to_string();

        let stored_zip_path = {
            let dir = batch_dir.join("zips").join(&zip_id);
            fs::create_dir_all(&dir).map_err(err_to_string)?;
            let dst = dir.join(&filename);
            fs::copy(&source_path, &dst).map_err(err_to_string)?;
            dst
        };

        let mut summary = ZipSummary {
            id: zip_id.clone(),
            filename: filename.clone(),
            source_path: p.clone(),
            stored_path: stored_zip_path.to_string_lossy().to_string(),
            extracted_dir: String::new(),
            include_original_zip: false,
            status: "processing".to_string(),
            word: WordFields::default(),
            has_video: false,
            has_sample: false,
            video_entries: vec![],
            video_files: vec![],
            image_files: vec![],
            pdf_files: vec![],
            pdf_image_files: vec![],
            pdf_page_screenshot_files: vec![],
        };

        let zip_scan = match scan_zip(&stored_zip_path) {
            Ok(v) => v,
            Err(e) => {
                summary.status = format!("failed: {e:#}");
                zips.push(summary);
                continue;
            }
        };

        let (word, video_entries) = match extract_word_and_videos(&stored_zip_path, &zip_scan) {
            Ok(v) => v,
            Err(e) => {
                summary.status = format!("failed: {e:#}");
                zips.push(summary);
                continue;
            }
        };

        summary.word = word;
        summary.has_sample = zip_scan.has_sample;
        summary.video_entries = video_entries;
        summary.has_video = !summary.video_entries.is_empty();

        // 解压用于预览（视频/图片/PDF）
        if let Err(e) = extract_preview_files(&batch_dir, &zip_id, &stored_zip_path, &zip_scan, &mut summary) {
            summary.status = format!("failed: {e:#}");
            zips.push(summary);
            continue;
        }

        summary.status = "completed".to_string();

        zips.push(summary);
    }

    let batch = BatchSummary {
        batch_id: batch_id.clone(),
        created_at: now.unix_timestamp(),
        zips,
    };

    let meta_path = batch_dir.join("batch.json");
    fs::write(&meta_path, serde_json::to_vec_pretty(&batch).map_err(err_to_string)?)
        .map_err(err_to_string)?;

    *state.last_batch_id.lock().unwrap() = Some(batch_id);
    Ok(batch)
}

#[tauri::command]
fn export_excel(app: tauri::AppHandle, batch_id: String) -> Result<String, String> {
    let batch_dir = batch_dir(&app, &batch_id).map_err(err_to_string)?;
    let batch: BatchSummary = read_batch(&batch_dir).map_err(err_to_string)?;
    export_excel_impl(&app, &batch)
}

#[tauri::command]
fn export_excel_with_selection(
    app: tauri::AppHandle,
    batch_id: String,
    zip_ids: Vec<String>,
) -> Result<String, String> {
    let batch_dir = batch_dir(&app, &batch_id).map_err(err_to_string)?;
    let mut batch: BatchSummary = read_batch(&batch_dir).map_err(err_to_string)?;
    if !zip_ids.is_empty() {
        batch.zips.retain(|z| zip_ids.contains(&z.id));
    }
    export_excel_impl(&app, &batch)
}

fn export_excel_impl(app: &tauri::AppHandle, batch: &BatchSummary) -> Result<String, String> {
    let now = OffsetDateTime::now_utc();
    let _ = app;
    let out = prompt_save_path(default_export_excel_name(now), "xlsx", "Excel")?;

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let header_format = Format::new().set_bold().set_align(FormatAlign::Center);
    let headers = [
        "序号",
        "日期",
        "编码",
        "标题",
        "类型",
        "样本（视频OR图文）",
        "是否有样本",
        "是否多批次任务",
        "下发时间",
        "任务执行",
        "备注",
    ];
    for (i, h) in headers.iter().enumerate() {
        worksheet
            .write_string_with_format(0, i as u16, *h, &header_format)
            .map_err(err_to_string)?;
    }

    for (idx, z) in batch.zips.iter().enumerate() {
        let row = (idx + 1) as u32;
        let date = format!(
            "{:04}{:02}{:02}",
            now.year(),
            now.month() as u8,
            now.day()
        );
        let sample_kind = if !z.video_files.is_empty() || !z.video_entries.is_empty() {
            "视频"
        } else if !z.image_files.is_empty()
            || !z.pdf_image_files.is_empty()
            || !z.pdf_page_screenshot_files.is_empty()
        {
            "图文"
        } else {
            ""
        };
        let has_sample = if sample_kind.is_empty() { "否" } else { "是" };

        worksheet
            .write_number(row, 0, (idx + 1) as f64)
            .map_err(err_to_string)?;
        worksheet
            .write_string(row, 1, &date)
            .map_err(err_to_string)?;
        worksheet
            .write_string(row, 2, z.word.instruction_no.trim())
            .map_err(err_to_string)?;
        worksheet
            .write_string(row, 3, z.word.title.trim())
            .map_err(err_to_string)?;
        worksheet
            .write_string(row, 4, z.word.title.trim())
            .map_err(err_to_string)?;
        worksheet
            .write_string(row, 5, sample_kind)
            .map_err(err_to_string)?;
        worksheet
            .write_string(row, 6, has_sample)
            .map_err(err_to_string)?;
        worksheet
            .write_string(row, 7, "否")
            .map_err(err_to_string)?;
        worksheet
            .write_string(row, 8, z.word.issued_at.trim())
            .map_err(err_to_string)?;
        worksheet
            .write_string(row, 9, "已完成")
            .map_err(err_to_string)?;
        worksheet.write_string(row, 10, "").map_err(err_to_string)?;
    }

    workbook
        .save(out.to_string_lossy().as_ref())
        .map_err(err_to_string)?;
    Ok(out.to_string_lossy().to_string())
}

#[tauri::command]
fn export_bundle_zip(app: tauri::AppHandle, batch_id: String) -> Result<String, String> {
    let batch_dir = batch_dir(&app, &batch_id).map_err(err_to_string)?;
    let batch: BatchSummary = read_batch(&batch_dir).map_err(err_to_string)?;

    let now = OffsetDateTime::now_utc();
    let out = prompt_save_path(default_export_bundle_name(now), "zip", "ZIP")?;

    let docx_bytes = build_summary_docx(&batch).map_err(err_to_string)?;
    let bundle_bytes = build_bundle_zip_bytes(&batch, &docx_bytes).map_err(err_to_string)?;

    fs::write(&out, bundle_bytes).map_err(err_to_string)?;
    Ok(out.to_string_lossy().to_string())
}

#[tauri::command]
fn export_bundle_zip_with_selection(
    app: tauri::AppHandle,
    batch_id: String,
    selection: ExportBundleSelection,
) -> Result<String, String> {
    let batch_dir = batch_dir(&app, &batch_id).map_err(err_to_string)?;
    let batch: BatchSummary = read_batch(&batch_dir).map_err(err_to_string)?;
    let batch = apply_bundle_selection(&batch, selection).map_err(err_to_string)?;

    if batch.zips.is_empty() {
        return Err("未选择任何ZIP用于导出".to_string());
    }

    let now = OffsetDateTime::now_utc();
    let out = prompt_save_path(default_export_bundle_name(now), "zip", "ZIP")?;

    let docx_bytes = build_summary_docx(&batch).map_err(err_to_string)?;
    let bundle_bytes = build_bundle_zip_bytes(&batch, &docx_bytes).map_err(err_to_string)?;
    fs::write(&out, bundle_bytes).map_err(err_to_string)?;
    Ok(out.to_string_lossy().to_string())
}

fn read_batch(batch_dir: &Path) -> Result<BatchSummary> {
    let path = batch_dir.join("batch.json");
    let data = fs::read(&path).with_context(|| format!("读取批次信息失败: {}", path.display()))?;
    let batch: BatchSummary = serde_json::from_slice(&data)?;
    Ok(batch)
}

fn apply_bundle_selection(batch: &BatchSummary, selection: ExportBundleSelection) -> Result<BatchSummary> {
    let mut out = Vec::new();

    for z in &batch.zips {
        let sel = selection.zips.iter().find(|s| s.zip_id == z.id);
        let Some(sel) = sel else { continue };
        if !sel.include {
            continue;
        }

        let mut z2 = z.clone();
        z2.include_original_zip = sel.include_original_zip;
        let mut selected_videos = Vec::new();
        for &idx in &sel.selected_video_indices {
            if let Some(p) = z.video_files.get(idx) {
                selected_videos.push(p.clone());
            }
        }
        z2.video_files = selected_videos;

        let mut selected_images = Vec::new();
        for &idx in &sel.selected_image_indices {
            if let Some(p) = z.image_files.get(idx) {
                selected_images.push(p.clone());
            }
        }
        z2.image_files = selected_images;

        let mut selected_pdfs = Vec::new();
        for &idx in &sel.selected_pdf_indices {
            if let Some(p) = z.pdf_files.get(idx) {
                selected_pdfs.push(p.clone());
            }
        }
        z2.pdf_files = selected_pdfs;

        let mut selected_pdf_imgs = Vec::new();
        for &idx in &sel.selected_pdf_image_indices {
            if let Some(p) = z.pdf_image_files.get(idx) {
                selected_pdf_imgs.push(p.clone());
            }
        }
        z2.pdf_image_files = selected_pdf_imgs;

        let mut selected_pdf_screens = Vec::new();
        for &idx in &sel.selected_pdf_page_screenshot_indices {
            if let Some(p) = z.pdf_page_screenshot_files.get(idx) {
                selected_pdf_screens.push(p.clone());
            }
        }
        z2.pdf_page_screenshot_files = selected_pdf_screens;

        out.push(z2);
    }

    Ok(BatchSummary {
        batch_id: batch.batch_id.clone(),
        created_at: batch.created_at,
        zips: out,
    })
}

#[derive(Debug, Clone)]
struct ZipScan {
    docx_entry: String,
    video_entries: Vec<String>,
    image_entries: Vec<String>,
    pdf_entries: Vec<String>,
    has_sample: bool,
}

fn scan_zip(zip_path: &Path) -> Result<ZipScan> {
    let f = fs::File::open(zip_path)?;
    let mut zip = ZipArchive::new(f)?;

    let mut docx_entry: Option<String> = None;
    let mut has_sample = false;
    let mut video_entries = Vec::new();
    let mut image_entries = Vec::new();
    let mut pdf_entries = Vec::new();

    for i in 0..zip.len() {
        let name = zip.by_index(i)?.name().to_string();
        let lower = name.to_ascii_lowercase();

        if lower.ends_with(".docx") {
            if docx_entry.is_some() {
                return Err(anyhow!("ZIP内发现多个docx，不符合前提"));
            }
            docx_entry = Some(name);
            continue;
        }

        if lower.ends_with("/") || lower.ends_with(".ds_store") {
            continue;
        }

        // Word之外都算样本
        has_sample = true;

        if lower.ends_with(".mp4") {
            video_entries.push(name);
        } else if lower.ends_with(".pdf") {
            pdf_entries.push(name);
        } else if lower.ends_with(".png")
            || lower.ends_with(".jpg")
            || lower.ends_with(".jpeg")
            || lower.ends_with(".gif")
        {
            image_entries.push(name);
        }
    }

    Ok(ZipScan {
        docx_entry: docx_entry.ok_or_else(|| anyhow!("ZIP内未找到docx"))?,
        video_entries,
        image_entries,
        pdf_entries,
        has_sample,
    })
}

fn extract_word_and_videos(zip_path: &Path, scan: &ZipScan) -> Result<(WordFields, Vec<String>)> {
    let f = fs::File::open(zip_path)?;
    let mut zip = ZipArchive::new(f)?;

    let docx_bytes = {
        let mut file = zip.by_name(&scan.docx_entry)?;
        let mut buf = Vec::new();
        file.read_to_end(&mut buf)?;
        buf
    };
    let fields = extract_fields_from_docx(&docx_bytes)?;
    Ok((fields, scan.video_entries.clone()))
}

fn extract_preview_files(
    batch_dir: &Path,
    zip_id: &str,
    zip_path: &Path,
    scan: &ZipScan,
    summary: &mut ZipSummary,
) -> Result<()> {
    let root = batch_dir.join("zips").join(zip_id).join("extracted");
    summary.extracted_dir = root.to_string_lossy().to_string();
    let videos_dir = root.join("videos");
    let images_dir = root.join("images");
    let pdf_dir = root.join("pdf");
    let pdf_images_dir = root.join("pdf_images");
    fs::create_dir_all(&videos_dir)?;
    fs::create_dir_all(&images_dir)?;
    fs::create_dir_all(&pdf_dir)?;
    fs::create_dir_all(&pdf_images_dir)?;

    let f = fs::File::open(zip_path)?;
    let mut zip = ZipArchive::new(f)?;

    for entry in &scan.video_entries {
        let mut file = zip.by_name(entry)?;
        let name = safe_basename(entry);
        let out = unique_path(&videos_dir, &name);
        let mut buf = Vec::new();
        file.read_to_end(&mut buf)?;
        fs::write(&out, buf)?;
        summary.video_files.push(out.to_string_lossy().to_string());
    }

    for entry in &scan.image_entries {
        let mut file = zip.by_name(entry)?;
        let name = safe_basename(entry);
        let out = unique_path(&images_dir, &name);
        let mut buf = Vec::new();
        file.read_to_end(&mut buf)?;
        fs::write(&out, buf)?;
        summary.image_files.push(out.to_string_lossy().to_string());
    }

    for entry in &scan.pdf_entries {
        let mut file = zip.by_name(entry)?;
        let name = safe_basename(entry);
        let out = unique_path(&pdf_dir, &name);
        let mut buf = Vec::new();
        file.read_to_end(&mut buf)?;
        fs::write(&out, buf)?;
        summary.pdf_files.push(out.to_string_lossy().to_string());

        // 提取“页面可见图片”（当前实现：优先支持 XObject Image + DCTDecode(JPEG)）
        match extract_visible_images_from_pdf(&out) {
            Ok(imgs) => {
                for (idx, bytes, ext) in imgs {
                    let file_name = match ext.as_str() {
                        "jpg" | "jpeg" => format!("{idx}.jpg"),
                        "png" => format!("{idx}.png"),
                        _ => format!("{idx}.{ext}"),
                    };
                    let out_img = unique_path(&pdf_images_dir, &file_name);
                    fs::write(&out_img, bytes)?;
                    summary
                        .pdf_image_files
                        .push(out_img.to_string_lossy().to_string());
                }
            }
            Err(_) => {
                // PDF解析失败不应阻塞导入；用户仍可系统打开PDF
            }
        }
    }

    Ok(())
}

fn extract_visible_images_from_pdf(pdf_path: &Path) -> Result<Vec<(String, Vec<u8>, String)>> {
    let bytes = fs::read(pdf_path)?;
    let doc = PdfDocument::load_mem(&bytes)?;
    let mut out = Vec::new();

    // 收集每页 content stream 中使用到的 XObject 名称（Do 操作）
    let pages = doc.get_pages();
    for (_page_no, page_id) in pages {
        let page_obj = match doc.get_object(page_id) {
            Ok(o) => o,
            Err(_) => continue,
        };
        let page_dict = match page_obj.as_dict() {
            Ok(d) => d,
            Err(_) => continue,
        };

        let resources_obj = match page_dict.get(b"Resources") {
            Ok(r) => r,
            Err(_) => continue,
        };
        let resources_obj = doc.dereference(resources_obj).map(|(_, o)| o)?;
        let res_dict = match resources_obj.as_dict() {
            Ok(d) => d,
            Err(_) => continue,
        };

        let xobject_obj = match res_dict.get(b"XObject") {
            Ok(x) => x,
            Err(_) => continue,
        };
        let xobject_obj = doc.dereference(xobject_obj).map(|(_, o)| o)?;
        let xobj_dict = match xobject_obj.as_dict() {
            Ok(d) => d,
            Err(_) => continue,
        };

        let content_data = match doc.get_page_content(page_id) {
            Ok(d) => d,
            Err(_) => continue,
        };

        let content = match lopdf::content::Content::decode(&content_data) {
            Ok(c) => c,
            Err(_) => continue,
        };

        for op in content.operations {
            if op.operator != "Do" {
                continue;
            }
            if op.operands.len() != 1 {
                continue;
            }
            let name = match op.operands[0].as_name() {
                Ok(n) => n,
                Err(_) => continue,
            };

            let obj = match xobj_dict.get(name) {
                Ok(o) => o,
                Err(_) => continue,
            };
            let resolved = match doc.dereference(obj) {
                Ok((_, r)) => r,
                Err(_) => continue,
            };
            let PdfObject::Stream(stream) = resolved else { continue };

            let subtype = match stream.dict.get(b"Subtype").and_then(|o| o.as_name()) {
                Ok(v) => v,
                Err(_) => continue,
            };
            if subtype != b"Image" {
                continue;
            }

            let (data, ext) = match image_stream_to_bytes(&doc, &stream) {
                Some(v) => v,
                None => continue,
            };
            out.push((format!("pdfimg_{}", out.len() + 1), data, ext));
        }
    }

    Ok(out)
}

fn image_stream_to_bytes(doc: &PdfDocument, stream: &lopdf::Stream) -> Option<(Vec<u8>, String)> {
    let filter = stream.dict.get(b"Filter").ok()?;

    // 常见：DCTDecode = JPEG
    if let Ok(name) = filter.as_name() {
        if name == b"DCTDecode" {
            return Some((stream.content.clone(), "jpg".to_string()));
        }
        if name == b"FlateDecode" {
            // 解压后的内容通常是原始像素数据，无法直接作为常见图片格式保存；先跳过
            let _ = doc;
            return None;
        }
    }
    if let Ok(arr) = filter.as_array() {
        if let Some(first) = arr.first().and_then(|o| o.as_name().ok()) {
            if first == b"DCTDecode" {
                return Some((stream.content.clone(), "jpg".to_string()));
            }
        }
    }
    None
}

fn unique_path(dir: &Path, file_name: &str) -> PathBuf {
    let base = Path::new(file_name)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("file");
    let ext = Path::new(file_name).extension().and_then(|s| s.to_str());

    let mut n = 0usize;
    loop {
        let candidate = if n == 0 {
            match ext {
                Some(ext) => dir.join(format!("{base}.{ext}")),
                None => dir.join(base),
            }
        } else {
            match ext {
                Some(ext) => dir.join(format!("{base}_{n}.{ext}")),
                None => dir.join(format!("{base}_{n}")),
            }
        };
        if !candidate.exists() {
            return candidate;
        }
        n += 1;
    }
}

fn extract_fields_from_docx(docx_bytes: &[u8]) -> Result<WordFields> {
    let cursor = Cursor::new(docx_bytes);
    let mut zip = ZipArchive::new(cursor)?;
    let mut document_xml = zip
        .by_name("word/document.xml")
        .context("docx缺少word/document.xml")?;
    let mut xml = String::new();
    document_xml.read_to_string(&mut xml)?;

    let text = extract_paragraph_texts(&xml)?;
    let mut map: std::collections::BTreeMap<String, Vec<String>> = std::collections::BTreeMap::new();
    for cap in RE_FIELD.captures_iter(&text) {
        let key = cap.get(1).unwrap().as_str().to_string();
        let value = cap.name("v").unwrap().as_str().trim().to_string();
        map.entry(key).or_default().push(value);
    }

    fn first_nonempty(values: Option<&Vec<String>>) -> String {
        let Some(values) = values else { return String::new() };
        for v in values {
            let t = v.trim();
            if !t.is_empty() {
                return t.to_string();
            }
        }
        String::new()
    }

    fn best_instruction_content(values: Option<&Vec<String>>) -> String {
        let Some(values) = values else { return String::new() };
        let mut best: Option<&str> = None;
        let mut best_len = 0usize;
        for v in values {
            let t = v.trim();
            if t.is_empty() {
                continue;
            }
            // 过滤已知的表头/噪音（你反馈的“序号”就是表格表头）
            if t == "序号" {
                continue;
            }
            // 优先取“更像正文”的值（更长）
            if t.len() > best_len {
                best = Some(t);
                best_len = t.len();
            }
        }
        best.unwrap_or("").to_string()
    }

    Ok(WordFields {
        instruction_no: first_nonempty(map.get("指令编号")),
        title: first_nonempty(map.get("指令标题")),
        issued_at: first_nonempty(map.get("下发时间")),
        content: best_instruction_content(map.get("指令内容")),
    })
}

fn extract_paragraph_texts(document_xml: &str) -> Result<String> {
    let mut reader = Reader::from_str(document_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut current = String::new();
    let mut out = String::new();
    let mut in_paragraph = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"w:p" {
                    in_paragraph = true;
                    current.clear();
                }
            }
            Ok(Event::Empty(e)) => {
                if in_paragraph && e.name().as_ref() == b"w:br" {
                    current.push('\n');
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:p" {
                    in_paragraph = false;
                    let line = normalize_text(&current);
                    if !line.trim().is_empty() {
                        out.push_str(line.trim_end());
                        out.push('\n');
                    }
                }
            }
            Ok(Event::Text(e)) => {
                if in_paragraph {
                    current.push_str(&e.unescape()?.to_string());
                }
            }
            Ok(Event::Eof) => break,
            Err(err) => return Err(err.into()),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn normalize_text(s: &str) -> String {
    s.replace('\u{00A0}', " ")
        .replace('\u{3000}', " ")
        .replace('：', ":")
}

fn build_summary_docx(batch: &BatchSummary) -> Result<Vec<u8>> {
    let mut docx = Docx::new();
    docx = docx.add_paragraph(
        Paragraph::new().add_run(Run::new().add_text("汇总文档").bold()),
    );

    for z in &batch.zips {
        let chapter = z.filename.trim_end_matches(".zip");
        let zip_folder = format!("attachments/{}/", z.id);
        docx = docx.add_paragraph(Paragraph::new().add_run(
            Run::new().add_text(format!("ZIP章节: {chapter}")).bold(),
        ));
        docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!(
            "指令编号:  {}",
            z.word.instruction_no
        ))));
        docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!(
            "指令标题:  {}",
            z.word.title
        ))));
        docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!(
            "下发时间:  {}",
            z.word.issued_at
        ))));
        if !z.word.content.trim().is_empty() {
            docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!(
                "指令内容:  {}",
                z.word.content
            ))));
        }

        if !z.image_files.is_empty() {
            docx = docx.add_paragraph(
                Paragraph::new().add_run(Run::new().add_text("图片").bold()),
            );
            for img_path in &z.image_files {
                let bytes = fs::read(img_path)
                    .with_context(|| format!("读取图片失败: {}", img_path))?;
                let pic = Pic::new(&bytes);
                docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_image(pic)));
            }
        }

        if !z.pdf_image_files.is_empty() {
            docx = docx.add_paragraph(
                Paragraph::new().add_run(Run::new().add_text("PDF页面图片:").bold()),
            );
            for img_path in &z.pdf_image_files {
                let bytes = fs::read(img_path)
                    .with_context(|| format!("读取PDF图片失败: {}", img_path))?;
                let pic = Pic::new(&bytes);
                docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_image(pic)));
            }
        }

        if !z.pdf_page_screenshot_files.is_empty() {
            docx = docx.add_paragraph(
                Paragraph::new().add_run(Run::new().add_text("PDF页面截图:").bold()),
            );
            for img_path in &z.pdf_page_screenshot_files {
                let bytes = fs::read(img_path)
                    .with_context(|| format!("读取PDF页面截图失败: {}", img_path))?;
                let pic = Pic::new(&bytes);
                docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_image(pic)));
            }
        }

        docx = docx.add_paragraph(
            Paragraph::new().add_run(Run::new().add_text("附件清单:").bold()),
        );

        // 仅提供“本ZIP附件文件夹”链接
        let folder_link = Hyperlink::new(&zip_folder, HyperlinkType::External)
            .add_run(Run::new().add_text(zip_folder.clone()).style("Hyperlink"));
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("附件目录："))
                .add_hyperlink(folder_link),
        );

        for video in &z.video_files {
            docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!(
                "- {}",
                safe_basename(video)
            ))));
        }
        if z.include_original_zip {
            docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text(format!(
                "- {}",
                z.filename
            ))));
        }
        if z.video_files.is_empty() && !z.include_original_zip {
            docx = docx
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("- （无）")));
        }
        docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text("— — —")));
    }

    let mut out = Cursor::new(Vec::<u8>::new());
    docx.build()
        .pack(&mut out)
        .map_err(|e| anyhow!("docx生成失败: {e:?}"))?;
    Ok(out.into_inner())
}

fn build_bundle_zip_bytes(batch: &BatchSummary, docx_bytes: &[u8]) -> Result<Vec<u8>> {
    let now = OffsetDateTime::now_utc();
    let zip_dt = zip::DateTime::from_date_and_time(
        now.year() as u16,
        now.month() as u8,
        now.day() as u8,
        now.hour() as u8,
        now.minute() as u8,
        now.second() as u8,
    )
    .unwrap_or_else(|_| zip::DateTime::default());

    let file_options = FileOptions::<()>::default()
        .compression_method(CompressionMethod::Deflated)
        .unix_permissions(0o644)
        .last_modified_time(zip_dt);
    let dir_options = FileOptions::<()>::default()
        .compression_method(CompressionMethod::Deflated)
        .unix_permissions(0o755)
        .last_modified_time(zip_dt);

    let mut out = Cursor::new(Vec::<u8>::new());
    let mut writer = ZipWriter::new(&mut out);

    writer.start_file("汇总文档.docx", file_options)?;
    writer.write_all(docx_bytes)?;

    writer.add_directory("attachments/", dir_options)?;

    for z in &batch.zips {
        let zip_dir = format!("attachments/{}/", z.id);
        writer.add_directory(&zip_dir, dir_options)?;

        let zip_path = if !z.stored_path.trim().is_empty() {
            PathBuf::from(&z.stored_path)
        } else {
            PathBuf::from(&z.source_path)
        };
        if z.include_original_zip {
            let zip_bytes = fs::read(&zip_path)
                .with_context(|| format!("读取ZIP失败: {}", zip_path.display()))?;
            writer.start_file(format!("{zip_dir}{}", z.filename), file_options)?;
            writer.write_all(&zip_bytes)?;
        }

        for video_path in &z.video_files {
            let bytes = fs::read(video_path)
                .with_context(|| format!("读取视频失败: {}", video_path))?;
            writer.start_file(
                format!("{zip_dir}{}", safe_basename(video_path)),
                file_options,
            )?;
            writer.write_all(&bytes)?;
        }

        for pdf_path in &z.pdf_files {
            let bytes = fs::read(pdf_path)
                .with_context(|| format!("读取PDF失败: {}", pdf_path))?;
            writer.start_file(format!("{zip_dir}{}", safe_basename(pdf_path)), file_options)?;
            writer.write_all(&bytes)?;
        }
    }

    writer.finish()?;
    Ok(out.into_inner())
}

fn safe_basename(name: &str) -> String {
    let n = name.replace('\\', "/");
    n.rsplit('/').next().unwrap_or(&n).to_string()
}

fn err_to_string<E: std::fmt::Display>(e: E) -> String {
    e.to_string()
}

#[tauri::command]
fn open_path(path: String) -> Result<(), String> {
    let p = PathBuf::from(path);
    if !p.exists() {
        return Err("路径不存在".to_string());
    }
    open_in_os(&p).map_err(err_to_string)?;
    Ok(())
}

fn open_in_os(path: &Path) -> Result<()> {
    #[cfg(target_os = "macos")]
    {
        let status = Command::new("open").arg(path).status()?;
        if !status.success() {
            return Err(anyhow!("open 返回非0状态码"));
        }
        return Ok(());
    }

    #[cfg(target_os = "windows")]
    {
        let s = path
            .to_str()
            .ok_or_else(|| anyhow!("路径包含无法显示的字符"))?;
        let cmd = format!("start \"\" \"{s}\"");
        let status = Command::new("cmd").args(["/C", &cmd]).status()?;
        if !status.success() {
            return Err(anyhow!("cmd start 返回非0状态码"));
        }
        return Ok(());
    }

    #[cfg(all(not(target_os = "macos"), not(target_os = "windows")))]
    {
        let status = Command::new("xdg-open").arg(path).status()?;
        if !status.success() {
            return Err(anyhow!("xdg-open 返回非0状态码"));
        }
        return Ok(());
    }
}

#[tauri::command]
fn get_preview_image_data(
    app: tauri::AppHandle,
    batch_id: String,
    zip_id: String,
    index: usize,
) -> Result<String, String> {
    let batch_dir = batch_dir(&app, &batch_id).map_err(err_to_string)?;
    let batch: BatchSummary = read_batch(&batch_dir).map_err(err_to_string)?;
    let z = batch
        .zips
        .iter()
        .find(|z| z.id == zip_id)
        .ok_or_else(|| "ZIP不存在".to_string())?;
    let path = z
        .image_files
        .get(index)
        .ok_or_else(|| "图片索引越界".to_string())?;
    let bytes = fs::read(path).map_err(err_to_string)?;
    let mime = guess_image_mime(path);
    let b64 = base64_encode(&bytes);
    Ok(format!("data:{mime};base64,{b64}"))
}

#[tauri::command]
fn get_preview_pdf_image_data(
    app: tauri::AppHandle,
    batch_id: String,
    zip_id: String,
    index: usize,
) -> Result<String, String> {
    let batch_dir = batch_dir(&app, &batch_id).map_err(err_to_string)?;
    let batch: BatchSummary = read_batch(&batch_dir).map_err(err_to_string)?;
    let z = batch
        .zips
        .iter()
        .find(|z| z.id == zip_id)
        .ok_or_else(|| "ZIP不存在".to_string())?;
    let path = z
        .pdf_image_files
        .get(index)
        .ok_or_else(|| "PDF图片索引越界".to_string())?;
    let bytes = fs::read(path).map_err(err_to_string)?;
    let mime = guess_image_mime(path);
    let b64 = base64_encode(&bytes);
    Ok(format!("data:{mime};base64,{b64}"))
}

#[tauri::command]
fn save_pdf_page_screenshots(
    app: tauri::AppHandle,
    batch_id: String,
    zip_id: String,
    pdf_name: String,
    screenshots: Vec<String>,
) -> Result<Vec<String>, String> {
    let batch_dir = batch_dir(&app, &batch_id).map_err(err_to_string)?;
    let mut batch: BatchSummary = read_batch(&batch_dir).map_err(err_to_string)?;

    let zip = batch
        .zips
        .iter_mut()
        .find(|z| z.id == zip_id)
        .ok_or_else(|| "ZIP不存在".to_string())?;

    let out_dir = batch_dir
        .join("zips")
        .join(&zip.id)
        .join("extracted")
        .join("pdf_screens")
        .join(sanitize_file_stem(&pdf_name));
    fs::create_dir_all(&out_dir).map_err(err_to_string)?;

    let mut saved = Vec::new();
    for (i, s) in screenshots.iter().enumerate() {
        let bytes = decode_data_url_base64(s).map_err(err_to_string)?;
        let out_path = out_dir.join(format!("page_{:03}.png", i + 1));
        fs::write(&out_path, bytes).map_err(err_to_string)?;
        let p = out_path.to_string_lossy().to_string();
        saved.push(p.clone());
        zip.pdf_page_screenshot_files.push(p);
    }

    let meta_path = batch_dir.join("batch.json");
    fs::write(&meta_path, serde_json::to_vec_pretty(&batch).map_err(err_to_string)?)
        .map_err(err_to_string)?;

    Ok(saved)
}

fn sanitize_file_stem(name: &str) -> String {
    let base = Path::new(name)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("pdf");
    base.chars()
        .map(|c| {
            if c.is_ascii_alphanumeric() || c == '-' || c == '_' {
                c
            } else {
                '_'
            }
        })
        .collect::<String>()
}

fn decode_data_url_base64(data_url: &str) -> Result<Vec<u8>> {
    let (meta, b64) = data_url
        .split_once(',')
        .ok_or_else(|| anyhow!("data url格式不正确"))?;
    if !meta.contains(";base64") {
        return Err(anyhow!("data url不是base64格式"));
    }
    base64_decode(b64)
}

fn base64_decode(s: &str) -> Result<Vec<u8>> {
    fn val(c: u8) -> Option<u8> {
        match c {
            b'A'..=b'Z' => Some(c - b'A'),
            b'a'..=b'z' => Some(c - b'a' + 26),
            b'0'..=b'9' => Some(c - b'0' + 52),
            b'+' => Some(62),
            b'/' => Some(63),
            _ => None,
        }
    }

    let bytes = s.as_bytes();
    let mut out = Vec::with_capacity((bytes.len() / 4) * 3);
    let mut buf = [0u8; 4];
    let mut n = 0usize;

    for &b in bytes {
        if b == b'=' {
            buf[n] = 64;
            n += 1;
        } else if let Some(v) = val(b) {
            buf[n] = v;
            n += 1;
        } else if b == b'\n' || b == b'\r' || b == b' ' || b == b'\t' {
            continue;
        } else {
            return Err(anyhow!("base64包含非法字符"));
        }

        if n == 4 {
            out.push((buf[0] << 2) | (buf[1] >> 4));
            if buf[2] != 64 {
                out.push((buf[1] << 4) | (buf[2] >> 2));
            }
            if buf[3] != 64 {
                out.push((buf[2] << 6) | buf[3]);
            }
            n = 0;
        }
    }

    Ok(out)
}

fn guess_image_mime(path: &str) -> &'static str {
    let lower = path.to_ascii_lowercase();
    if lower.ends_with(".png") {
        "image/png"
    } else if lower.ends_with(".jpg") || lower.ends_with(".jpeg") {
        "image/jpeg"
    } else if lower.ends_with(".gif") {
        "image/gif"
    } else {
        "application/octet-stream"
    }
}

fn base64_encode(bytes: &[u8]) -> String {
    const TABLE: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::with_capacity(((bytes.len() + 2) / 3) * 4);
    let mut i = 0usize;
    while i + 3 <= bytes.len() {
        let b0 = bytes[i];
        let b1 = bytes[i + 1];
        let b2 = bytes[i + 2];
        out.push(TABLE[(b0 >> 2) as usize] as char);
        out.push(TABLE[(((b0 & 0b0000_0011) << 4) | (b1 >> 4)) as usize] as char);
        out.push(TABLE[(((b1 & 0b0000_1111) << 2) | (b2 >> 6)) as usize] as char);
        out.push(TABLE[(b2 & 0b0011_1111) as usize] as char);
        i += 3;
    }

    let rem = bytes.len() - i;
    if rem == 1 {
        let b0 = bytes[i];
        out.push(TABLE[(b0 >> 2) as usize] as char);
        out.push(TABLE[((b0 & 0b0000_0011) << 4) as usize] as char);
        out.push('=');
        out.push('=');
    } else if rem == 2 {
        let b0 = bytes[i];
        let b1 = bytes[i + 1];
        out.push(TABLE[(b0 >> 2) as usize] as char);
        out.push(TABLE[(((b0 & 0b0000_0011) << 4) | (b1 >> 4)) as usize] as char);
        out.push(TABLE[((b1 & 0b0000_1111) << 2) as usize] as char);
        out.push('=');
    }

    out
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .manage(AppState::default())
        .invoke_handler(tauri::generate_handler![
            pick_zip_files,
            import_zips,
            export_excel,
            export_excel_with_selection,
            export_bundle_zip,
            export_bundle_zip_with_selection,
            open_path,
            get_preview_image_data,
            get_preview_pdf_image_data,
            save_pdf_page_screenshots
        ])
        .setup(|app| {
            if cfg!(debug_assertions) {
                app.handle().plugin(
                    tauri_plugin_log::Builder::default()
                        .level(log::LevelFilter::Info)
                        .build(),
                )?;
            }
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fixture_zip(name: &str) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .parent()
            .unwrap()
            .join("zip_simples")
            .join(name)
    }

    #[test]
    fn parse_fields_from_fixture_zip() {
        let zip_path = fixture_zip("202512110007-ZL1.zip");
        let scan = scan_zip(&zip_path).expect("scan_zip");
        let (fields, videos) = extract_word_and_videos(&zip_path, &scan).expect("extract");
        assert!(!fields.instruction_no.is_empty());
        assert!(!fields.title.is_empty());
        assert!(!fields.issued_at.is_empty());
        // 视频可能有也可能没有，至少不应崩
        assert_eq!(videos.len(), scan.video_entries.len());
    }

    #[test]
    fn build_bundle_zip_has_per_zip_attachments_dir_and_docx_links() {
        let zip_path = fixture_zip("202512110028-ZL1.zip");
        let scan = scan_zip(&zip_path).expect("scan_zip");
        let (fields, _videos) = extract_word_and_videos(&zip_path, &scan).expect("extract");

        let tmp_root = std::env::temp_dir().join(format!("archivebox_test_{}", Uuid::new_v4()));
        fs::create_dir_all(&tmp_root).unwrap();
        let batch_dir = tmp_root.join("batch");
        fs::create_dir_all(&batch_dir).unwrap();

        let zip_id = Uuid::new_v4().to_string();
        let stored_zip = tmp_root.join(format!("{}.zip", zip_id));
        fs::copy(&zip_path, &stored_zip).unwrap();

        let mut zip_summary = ZipSummary {
            id: zip_id.clone(),
            filename: "202512110028-ZL1.zip".to_string(),
            source_path: zip_path.to_string_lossy().to_string(),
            stored_path: stored_zip.to_string_lossy().to_string(),
            extracted_dir: String::new(),
            include_original_zip: true,
            status: "completed".to_string(),
            word: fields,
            has_video: !scan.video_entries.is_empty(),
            has_sample: scan.has_sample,
            video_entries: scan.video_entries.clone(),
            video_files: vec![],
            image_files: vec![],
            pdf_files: vec![],
            pdf_image_files: vec![],
            pdf_page_screenshot_files: vec![],
        };

        extract_preview_files(&batch_dir, &zip_id, &stored_zip, &scan, &mut zip_summary)
            .expect("extract_preview_files");

        let batch = BatchSummary {
            batch_id: "batch_test".to_string(),
            created_at: OffsetDateTime::now_utc().unix_timestamp(),
            zips: vec![zip_summary.clone()],
        };

        let docx_bytes = build_summary_docx(&batch).expect("build_summary_docx");
        assert!(!docx_bytes.is_empty());

        // docx 内应有指向 attachments/<zipId>/ 的链接关系
        let mut docx_zip = ZipArchive::new(Cursor::new(&docx_bytes)).expect("docx zip");
        let mut rels = String::new();
        docx_zip
            .by_name("word/_rels/document.xml.rels")
            .expect("rels exists")
            .read_to_string(&mut rels)
            .unwrap();
        assert!(
            rels.contains(&format!("attachments/{}/", zip_id)),
            "rels should contain per-zip attachments link"
        );

        let bundle = build_bundle_zip_bytes(&batch, &docx_bytes).expect("bundle");
        let mut out_zip = ZipArchive::new(Cursor::new(bundle)).expect("bundle zip");

        // attachments/ 目录权限应为 0755，且必须包含 attachments/<zipId>/
        let attachments_dir_mode = out_zip
            .by_name("attachments/")
            .expect("attachments dir")
            .unix_mode()
            .unwrap_or(0);
        assert_eq!(attachments_dir_mode & 0o777, 0o755);
        let per_zip_dir_mode = out_zip
            .by_name(&format!("attachments/{}/", zip_id))
            .expect("per zip dir")
            .unix_mode()
            .unwrap_or(0);
        assert_eq!(per_zip_dir_mode & 0o777, 0o755);

        // 每个ZIP目录下必须包含原始ZIP
        out_zip
            .by_name(&format!("attachments/{}/{}", zip_id, zip_summary.filename))
            .expect("zip copied into per zip dir");
    }
}
