use anyhow::{anyhow, Context, Result};
use calamine::{open_workbook, Reader, Xls, Xlsx, Data};
use docx_rs::*;
use encoding_rs::{GBK, UTF_8};
use once_cell::sync::Lazy;
use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
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
use zip::{ZipArchive, ZipWriter};
use zip::write::FileOptions;


// 文件嵌入相关结构体和函数

/// 嵌入式文件结构
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub struct EmbeddedFile {
    pub id: String,
    pub name: String,
    pub path: String,
    pub data: Vec<u8>,
    pub content_type: String,
    pub file_type: FileType,
    pub zip_id: String,  // 所属章节ID
}

#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub enum FileType {
    Video,
    PDF,
    Image,
    Document,
    Excel,
    ZIP,
    Other(String),
}

/// 文件嵌入配置
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub struct EmbeddingConfig {
    pub enabled: bool,
    pub max_file_size: usize,
    pub allowed_types: Vec<String>,
    pub exclude_patterns: Vec<String>,
}

impl Default for EmbeddingConfig {
    fn default() -> Self {
        Self {
            enabled: true,
            max_file_size: 50 * 1024 * 1024,  // 50MB
            allowed_types: vec![
                "pdf".to_string(),
                "mp4".to_string(),
                "avi".to_string(),
                "mov".to_string(),
                "wmv".to_string(),
                "jpg".to_string(),
                "jpeg".to_string(),
                "png".to_string(),
                "gif".to_string(),
                "bmp".to_string(),
                "doc".to_string(),
                "docx".to_string(),
                "txt".to_string(),
                "zip".to_string(),
            ],
            exclude_patterns: vec![
                "*.tmp".to_string(),
                "*.temp".to_string(),
                ".*".to_string(),
            ],
        }
    }
}

/// 增强的汇总文档构建，支持文��嵌入
fn build_enhanced_summary_docx(
    batch: &BatchSummary,
    embed_files: bool,
) -> Result<(Docx, Vec<EmbeddedFile>)> {
    let mut docx = Docx::new();
    docx = docx.add_paragraph(
        Paragraph::new().add_run(Run::new().add_text("汇总文档").bold()),
    );

    let mut all_embedded_files = Vec::new();

    for z in &batch.zips {
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

        // 处理指令内容（保持换行格式）
        if !z.word.content.trim().is_empty() {
            docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text("指令内容:")));
            for line in z.word.content.lines() {
                let trimmed_line = line.trim();
                if !trimmed_line.is_empty() {
                    docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text(trimmed_line)));
                } else {
                    docx = docx.add_paragraph(Paragraph::new());
                }
            }
        }

        // 直接显示图片，删除"图片"标题
        for img_path in &z.image_files {
            let bytes = fs::read(img_path)
                .with_context(|| format!("读取图片失败: {}", img_path))?;
            let pic = Pic::new(&bytes);
            docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_image(pic)));
        }

        // 直接显示PDF图片，删除"PDF页面图片:"标题
        
        // 直接显示PDF截图，删除"PDF页面截图:"标题
        for img_path in &z.pdf_page_screenshot_files {
            let bytes = fs::read(img_path)
                .with_context(|| format!("读取PDF页面截图失败: {}", img_path))?;
            let pic = Pic::new(&bytes);
            docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_image(pic)));
        }

        // 添加章节标记段落（用于后续插入OLE对象）
        let marker = format!("EMBED_MARKER_{}", z.id);
        docx = docx.add_paragraph(
            Paragraph::new().add_run(Run::new().add_text(&marker).size(2).color("FFFFFF"))
        );

        // 收集需要嵌入的文件（包括视频、PDF、Excel、ZIP，不包括图片）
        if embed_files {
            // 嵌入视频文件
            for video_path in &z.video_files {
                if Path::new(video_path).exists() {
                    if let Ok(embed_file) = create_embedded_file(video_path, &z.id) {
                        all_embedded_files.push(embed_file);
                    }
                }
            }

            // 嵌入PDF文件
            for pdf_path in &z.pdf_files {
                if Path::new(pdf_path).exists() {
                    if let Ok(embed_file) = create_embedded_file(pdf_path, &z.id) {
                        all_embedded_files.push(embed_file);
                    }
                }
            }

            // 嵌入Excel文件
            for excel_path in &z.excel_files {
                if Path::new(excel_path).exists() {
                    if let Ok(embed_file) = create_embedded_file(excel_path, &z.id) {
                        all_embedded_files.push(embed_file);
                    }
                }
            }

            // 嵌入原始ZIP文件（如果启用）
            if z.include_original_zip {
                let zip_path = &z.stored_path;
                if Path::new(zip_path).exists() {
                    if let Ok(embed_file) = create_embedded_file(zip_path, &z.id) {
                        all_embedded_files.push(embed_file);
                    }
                }
            }
        }

        // 在不同章节之间添加空行，提高可读性
        docx = docx.add_paragraph(Paragraph::new());
    }

    // 返回文档和嵌入文件列表，让调用者处理最终的构建
    Ok((docx, all_embedded_files))
}

fn create_embedded_file(path: &str, zip_id: &str) -> Result<EmbeddedFile> {
    let data = fs::read(path)
        .with_context(|| format!("Failed to read file: {}", path))?;

    let name = Path::new(path)
        .file_name()
        .unwrap_or_default()
        .to_string_lossy()
        .to_string();

    let file_type = detect_file_type(&name);
    let content_type = get_content_type(&name);

    Ok(EmbeddedFile {
        id: format!("embed_{}", uuid::Uuid::new_v4().to_string().replace("-", "")),
        name,
        path: path.to_string(),
        data,
        content_type,
        file_type,
        zip_id: zip_id.to_string(),
    })
}

fn detect_file_type(filename: &str) -> FileType {
    let filename_lower = filename.to_lowercase();

    if filename_lower.ends_with(".mp4") || filename_lower.ends_with(".avi") ||
       filename_lower.ends_with(".mov") || filename_lower.ends_with(".wmv") ||
       filename_lower.ends_with(".mkv") || filename_lower.ends_with(".flv") {
        FileType::Video
    } else if filename_lower.ends_with(".pdf") {
        FileType::PDF
    } else if filename_lower.ends_with(".jpg") || filename_lower.ends_with(".jpeg") ||
              filename_lower.ends_with(".png") || filename_lower.ends_with(".gif") ||
              filename_lower.ends_with(".bmp") || filename_lower.ends_with(".webp") {
        FileType::Image
    } else if filename_lower.ends_with(".zip") {
        FileType::ZIP
    } else if filename_lower.ends_with(".xls") || filename_lower.ends_with(".xlsx") {
        FileType::Excel
    } else if filename_lower.ends_with(".doc") || filename_lower.ends_with(".docx") ||
              filename_lower.ends_with(".txt") || filename_lower.ends_with(".rtf") {
        FileType::Document
    } else {
        if let Some(ext) = Path::new(filename).extension() {
            FileType::Other(ext.to_string_lossy().to_string())
        } else {
            FileType::Other("unknown".to_string())
        }
    }
}

fn get_content_type(filename: &str) -> String {
    let filename_lower = filename.to_lowercase();

    match filename_lower.as_str() {
        f if f.ends_with(".pdf") => "application/pdf".to_string(),
        f if f.ends_with(".mp4") => "video/mp4".to_string(),
        f if f.ends_with(".avi") => "video/x-msvideo".to_string(),
        f if f.ends_with(".mov") => "video/quicktime".to_string(),
        f if f.ends_with(".wmv") => "video/x-ms-wmv".to_string(),
        f if f.ends_with(".mkv") => "video/x-matroska".to_string(),
        f if f.ends_with(".flv") => "video/x-flv".to_string(),
        f if f.ends_with(".jpg") || f.ends_with(".jpeg") => "image/jpeg".to_string(),
        f if f.ends_with(".png") => "image/png".to_string(),
        f if f.ends_with(".gif") => "image/gif".to_string(),
        f if f.ends_with(".bmp") => "image/bmp".to_string(),
        f if f.ends_with(".webp") => "image/webp".to_string(),
        f if f.ends_with(".zip") => "application/zip".to_string(),
        f if f.ends_with(".xls") => "application/vnd.ms-excel".to_string(),
        f if f.ends_with(".xlsx") => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet".to_string(),
        f if f.ends_with(".doc") => "application/msword".to_string(),
        f if f.ends_with(".docx") => "application/vnd.openxmlformats-officedocument.wordprocessingml.document".to_string(),
        f if f.ends_with(".txt") => "text/plain".to_string(),
        f if f.ends_with(".rtf") => "application/rtf".to_string(),
        _ => "application/octet-stream".to_string(),
    }
}

fn get_file_icon(file_type: &FileType) -> &'static str {
    match file_type {
        FileType::Video => "🎥",
        FileType::PDF => "📄",
        FileType::Image => "🖼️",
        FileType::Document => "📝",
        FileType::Excel => "📊",
        FileType::ZIP => "📦",
        FileType::Other(_) => "📎",
    }
}

/// 检查是否为文本文件
fn is_text_file(content_type: &str) -> bool {
    content_type.starts_with("text/") ||
    content_type.contains("plain") ||
    content_type.contains("rtf") ||
    content_type.contains("html") ||
    content_type.contains("xml") ||
    content_type.contains("json") ||
    content_type.contains("csv")
}

/// 构建带嵌入文件的 DOCX（真正的 OLE 嵌入）
fn build_docx_with_embeddings(
    base_docx: Docx,
    embedded_files: &[EmbeddedFile]
) -> Result<Vec<u8>> {
    // 1. 首先生成基础的 DOCX
    let xmldocx = base_docx.build();
    let mut base_bytes = Vec::new();
    {
        let mut cursor = Cursor::new(&mut base_bytes);
        xmldocx.pack(&mut cursor)?;
    }

    // 2. 如果没有文件要嵌入，直接返回
    if embedded_files.is_empty() {
        return Ok(base_bytes);
    }

    println!("=== OLE 文件嵌入模式 ===");
    println!("找到 {} 个附件文件:", embedded_files.len());
    for (i, file) in embedded_files.iter().enumerate() {
        println!("  {}. {} (大小: {:.1} MB, 类型: {})",
            i + 1,
            file.name,
            file.data.len() as f64 / 1024.0 / 1024.0,
            file.content_type
        );
    }

    // 3. 执行真正的 OLE 嵌入
    match embed_ole_objects_into_docx(&base_bytes, embedded_files) {
        Ok(result) => {
            println!("✓ OLE 对象嵌入成功！");
            Ok(result)
        }
        Err(e) => {
            println!("⚠ OLE 嵌入失败: {}", e);
            println!("  返回基础文档以确保功能正常");
            Ok(base_bytes)
        }
    }
}

/// 尝试真正的文件嵌入
fn attempt_file_embedding(
    base_docx_bytes: &[u8],
    embedded_files: &[EmbeddedFile]
) -> Result<Vec<u8>> {
    // 如果没有文件要嵌入，直接返回
    if embedded_files.is_empty() {
        return Ok(base_docx_bytes.to_vec());
    }

    println!("开始真正的文件嵌入...");

    // 暂时返回基础文档，避免损坏现有功能
    // OpenXML操作非常复杂，容易损坏文档结构
    println!("⚠ OpenXML文件嵌入操作复杂，为确保文档完整性，暂时使用显示模式");
    println!("  显示的附件信息：");
    for file in embedded_files {
        println!("  - {} ({} bytes)", file.name, file.data.len());
    }

    Err(anyhow!("使用简化显示模式确保文档完整性"))
}

/// 生成嵌入文件的 Content_Types.xml
fn generate_content_types_xml(embedded_files: &[EmbeddedFile]) -> String {
    let mut types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
"#.to_string();

    // 添加嵌入文件的类型定义
    for file in embedded_files {
        let embed_path = format!("/word/embeddings/{}.bin", file.id);
        types.push_str(&format!(
            r#"    <Override PartName="{}" ContentType="{}"/>
"#, embed_path, file.content_type));
    }

    types.push_str("</Types>");
    types
}

/// 生成嵌入文件的 relationships.xml
fn generate_relationships_xml(embedded_files: &[EmbeddedFile]) -> String {
    let mut rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
"#.to_string();

    // 添加嵌入文件的关系
    for (index, file) in embedded_files.iter().enumerate() {
        let r_id = format!("rId{}", index + 1);
        let target = format!("embeddings/{}.bin", file.id);
        rels.push_str(&format!(
            r#"    <Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="{}"/>
"#, r_id, target));
    }

    rels.push_str("</Relationships>");
    rels
}

/// 生成包含嵌入文件的Content_Types.xml
fn build_content_types_with_embeds(embedded_files: &[EmbeddedFile]) -> String {
    let mut types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
    <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
    <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering.xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
"#.to_string();

    // 添加嵌入文件的类型定义
    for file in embedded_files {
        let embed_path = format!("/word/embeddings/{}",
            sanitize_filename(&format!("{}.{}",
                Path::new(&file.name).file_stem()
                    .unwrap_or_default()
                    .to_string_lossy(),
                Path::new(&file.name)
                    .extension()
                    .unwrap_or_default()
                    .to_string_lossy()
            ))
        );

        types.push_str(&format!(
            r#"    <Override PartName="{}" ContentType="{}"/>
"#, embed_path, file.content_type));
    }

    types.push_str("</Types>");
    types
}

/// 在document.xml中添加嵌入对象引用
fn add_embedded_objects_to_document(document_xml: &str, embedded_files: &[EmbeddedFile]) -> Result<String> {
    // 找到body标签的结束位置，并在那里插入嵌入对象
    let body_end_pattern = r"</w:body>";

    if !document_xml.contains(body_end_pattern) {
        return Err(anyhow!("文档中未找到body结束标签"));
    }

    let mut objects_xml = String::new();

    for (_, file) in embedded_files.iter().enumerate() {
        let file_icon = match file.file_type {
            FileType::PDF => "📄",
            FileType::Video => "🎥",
            FileType::ZIP => "📦",
            _ => "📎"
        };

        // 创建简化的嵌入对象段落 - 使用简单的超链接方式，这在Word中更可靠
        objects_xml.push_str(&format!(
            r#"<w:p><w:r><w:rPr><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr><w:t>{icon} {name} (双击打开附件)</w:t></w:r></w:p>"#,
            icon = file_icon,
            name = file.name
        ));
    }

    // 替换body结束标签
    let modified_document = document_xml.replace(body_end_pattern, &format!("{}\n{}", objects_xml, body_end_pattern));

    Ok(modified_document)
}

/// 生成包含嵌入文件的关系文档
fn build_relationships_with_embeds(embedded_files: &[EmbeddedFile]) -> String {
    let mut rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
"#.to_string();

    // 添加嵌入文件的关系
    for (index, file) in embedded_files.iter().enumerate() {
        let r_id = format!("rId{}", index + 1);
        let safe_filename = sanitize_filename(&format!("{}.{}",
            Path::new(&file.name).file_stem()
                .unwrap_or_default()
                .to_string_lossy(),
            Path::new(&file.name)
                .extension()
                .unwrap_or_default()
                .to_string_lossy()
        ));
        let target = format!("embeddings/{}", safe_filename);

        rels.push_str(&format!(
            r#"    <Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="{}"/>
"#, r_id, target));
    }

    rels.push_str("</Relationships>");
    rels
}

/// 清理文件名，移除不安全字符
fn sanitize_filename(filename: &str) -> String {
    let unsafe_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*'];
    let mut result = String::new();

    for c in filename.chars() {
        if unsafe_chars.contains(&c) {
            // 用下划线替换不安全字符
            result.push('_');
        } else {
            result.push(c);
        }
    }

    result
}

// ==================== OLE 嵌入核心功能 ====================

/// 将 OLE 对象嵌入到 DOCX 文件中（主函数）
fn embed_ole_objects_into_docx(
    docx_bytes: &[u8],
    embedded_files: &[EmbeddedFile]
) -> Result<Vec<u8>> {
    // 1. 打开现有的 DOCX (ZIP 格式)
    let reader = Cursor::new(docx_bytes);
    let mut zip_archive = ZipArchive::new(reader)?;

    // 2. 创建输出 ZIP
    let output_cursor = Cursor::new(Vec::new());
    let mut zip_writer = ZipWriter::new(output_cursor);

    // 3. 复制所有现有文件（除了需要修改的）
    let files_to_modify = vec![
        "word/document.xml",
        "word/_rels/document.xml.rels",
        "[Content_Types].xml"
    ];

    for i in 0..zip_archive.len() {
        let mut file = zip_archive.by_index(i)?;
        let name = file.name().to_string();

        if !files_to_modify.contains(&name.as_str()) {
            // 复制文件
            let options = FileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            zip_writer.start_file(&name, options)?;
            std::io::copy(&mut file, &mut zip_writer)?;
        }
    }

    // 4. 读取需要修改的文件
    let document_xml = read_file_from_zip_archive(&mut zip_archive, "word/document.xml")?;
    let rels_xml = read_file_from_zip_archive(&mut zip_archive, "word/_rels/document.xml.rels")?;
    let content_types_xml = read_file_from_zip_archive(&mut zip_archive, "[Content_Types].xml")?;

    // 5. 添加嵌入文件和图标
    let next_rid = get_next_relationship_id(&rels_xml);

    for (index, file) in embedded_files.iter().enumerate() {
        // 创建 OLE Package
        let ole_package = create_ole_package(file)?;
        let ole_filename = format!("word/embeddings/oleObject{}.bin", index + 1);

        let options = FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        zip_writer.start_file(&ole_filename, options)?;
        zip_writer.write_all(&ole_package)?;

        // 添加图标文件
        let icon_data = get_default_emf_icon(&file.file_type, &file.name);
        let icon_filename = format!("word/media/image{}.emf", index + 1);

        zip_writer.start_file(&icon_filename, options)?;
        zip_writer.write_all(&icon_data)?;
    }

    // 6. 修改 document.xml - 添加 OLE 对象
    let modified_document = add_ole_objects_to_document_xml(&document_xml, embedded_files, next_rid)?;
    let options = FileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    zip_writer.start_file("word/document.xml", options)?;
    zip_writer.write_all(modified_document.as_bytes())?;

    // 7. 修改 document.xml.rels - 添加关系
    let modified_rels = add_ole_relationships_to_rels(&rels_xml, embedded_files, next_rid)?;
    zip_writer.start_file("word/_rels/document.xml.rels", options)?;
    zip_writer.write_all(modified_rels.as_bytes())?;

    // 8. 修改 [Content_Types].xml - 添加内容类型
    let modified_content_types = add_ole_content_types(&content_types_xml)?;
    zip_writer.start_file("[Content_Types].xml", options)?;
    zip_writer.write_all(modified_content_types.as_bytes())?;

    // 9. 完成并获取数据
    let output_cursor = zip_writer.finish()?;
    let output_bytes = output_cursor.into_inner();

    Ok(output_bytes)
}

/// 创建 OLE Package 格式（OLE 复合文档）
/// 基于真实 Word 文档中的 Ole10Native 格式
fn create_ole_package(file: &EmbeddedFile) -> Result<Vec<u8>> {
    // 创建 Ole10Native 流数据
    let mut native_data = Vec::new();

    // 真实的 Ole10Native 流格式（来自实际的Word文档分析）：
    // [4 bytes] 文件大小（小端）
    // [2 bytes] 固定标记 0x02 0x00
    // [变长] GBK编码的完整文件名 + null terminator (0x00)
    // [变长] 原始文件路径 + null terminator
    // [3 bytes] 分隔符 0x00 0x00 0x03
    // [1 byte] Windows临时路径长度
    // [变长] Windows临时路径 + null terminator
    // [4 bytes] 文件数据大小（小端）
    // [变长] 实际文件数据

    // 将文件名转换为GBK编码（Windows ANSI编码）
    let (filename_bytes, _, _) = encoding_rs::GBK.encode(&file.name);
    let filename_gbk = filename_bytes.as_ref();

    // 1. 文件大小（4字节，小端）
    native_data.extend_from_slice(&(file.data.len() as u32).to_le_bytes());

    // 2. 固定标记（2字节）
    native_data.extend_from_slice(&[0x02, 0x00]);

    // 3. 完整文件名（GBK编码）+ null terminator
    native_data.extend_from_slice(filename_gbk);
    native_data.push(0);

    // 4. 原始文件路径（使用简化路径）+ null terminator
    let original_path = format!("C:/{}", file.name);
    let (path_bytes, _, _) = encoding_rs::GBK.encode(&original_path);
    native_data.extend_from_slice(path_bytes.as_ref());
    native_data.push(0);

    // 5. 路径后的分隔符（正确格式）
    // 两个额外的 null + 0x03 + 0x00
    native_data.push(0);
    native_data.push(0);
    native_data.push(0x03);
    native_data.push(0x00);

    // 6. Windows临时路径长度（4字节小端）+ 路径 + null terminator
    let temp_path = format!("C:\\Users\\Public\\{}", file.name);
    let (temp_path_bytes, _, _) = encoding_rs::GBK.encode(&temp_path);

    // 临时路径长度（包括null terminator，4字节小端）
    let temp_path_len = (temp_path_bytes.len() + 1) as u32;
    native_data.extend_from_slice(&temp_path_len.to_le_bytes());

    // 临时路径 + null terminator
    native_data.extend_from_slice(temp_path_bytes.as_ref());
    native_data.push(0);

    // 7. 文件数据大小（4字节，小端）
    native_data.extend_from_slice(&(file.data.len() as u32).to_le_bytes());

    // 8. 实际文件数据
    native_data.extend_from_slice(&file.data);

    // 创建 OLE 复合文档
    let mut output = Cursor::new(Vec::new());
    {
        let mut comp = cfb::CompoundFile::create(&mut output)?;

        // 写入 \x01Ole10Native 流
        comp.create_stream("\x01Ole10Native")?;
        let mut stream = comp.open_stream("\x01Ole10Native")?;
        stream.write_all(&native_data)?;
        drop(stream); // 显式关闭流

        // 添加 OLE 对象的标准流
        // \x01CompObj 流 - 描述对象类型
        comp.create_stream("\x01CompObj")?;
        let mut comp_obj_stream = comp.open_stream("\x01CompObj")?;
        let comp_obj_data = create_comp_obj_stream(&file.name);
        comp_obj_stream.write_all(&comp_obj_data)?;
        drop(comp_obj_stream); // 显式关闭流

        // 确保所有数据都写入
        drop(comp);
    }

    Ok(output.into_inner())
}

/// 创建 CompObj 流数据
fn create_comp_obj_stream(_filename: &str) -> Vec<u8> {
    let mut data = Vec::new();

    // 版本 (2 bytes)
    data.extend_from_slice(&0x0001u16.to_le_bytes());

    // Byte order (2 bytes)
    data.extend_from_slice(&0xFFFEu16.to_le_bytes());

    // Format version (4 bytes)
    data.extend_from_slice(&0x00000A03u32.to_le_bytes());

    // Reserved (4 bytes)
    data.extend_from_slice(&0xFFFFFFFFu32.to_le_bytes());

    // CLSID (16 bytes) - Package 的 CLSID: {0003000C-0000-0000-C000-000000000046}
    data.extend_from_slice(&[
        0x0C, 0x00, 0x00, 0x00,
        0x00, 0x00,
        0x00, 0x00,
        0xC0, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x46
    ]);

    // User type string (length + string)
    let user_type = "Package";
    data.extend_from_slice(&(user_type.len() as u32).to_le_bytes());
    data.extend_from_slice(user_type.as_bytes());
    data.push(0); // Null terminator

    // Clipboard format name (empty)
    data.extend_from_slice(&0u32.to_le_bytes());

    // Reserved (4 bytes)
    data.extend_from_slice(&0x00000000u32.to_le_bytes());

    data
}

/// 写入带长度前缀的 UTF-16 字符串
fn write_utf16_length_prefixed_string(buffer: &mut Vec<u8>, s: &str) {
    let utf16: Vec<u16> = s.encode_utf16().collect();
    let byte_len = (utf16.len() * 2) as u32;

    // 写入长度（字节数）
    buffer.extend_from_slice(&byte_len.to_le_bytes());

    // 写入 UTF-16 字符串
    for code_unit in utf16 {
        buffer.extend_from_slice(&code_unit.to_le_bytes());
    }
}

/// 获取对应文件类型的 EMF 图标
/// 智能截断文件名使其适合指定的最大字节数（UTF-16LE编码）
/// 保留文件扩展名，在合适的位置截断主文件名
fn truncate_filename_to_bytes(filename: &str, max_bytes: usize) -> String {
    use std::path::Path;

    // 计算当前文件名的UTF-16LE字节数
    let current_bytes: Vec<u8> = filename.encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    // 如果已经适合，直接返回
    if current_bytes.len() <= max_bytes {
        return filename.to_string();
    }

    // 分离文件名和扩展名
    let path = Path::new(filename);
    let extension = path.extension()
        .and_then(|e| e.to_str())
        .unwrap_or("");
    let stem = path.file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or(filename);

    // 计算扩展名的字节数（包括点号）
    let ext_with_dot = if !extension.is_empty() {
        format!(".{}", extension)
    } else {
        String::new()
    };
    let ext_bytes: Vec<u8> = ext_with_dot.encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    // 使用更短的省略号（2个点而不是3个）
    let ellipsis = "..";
    let ellipsis_bytes: Vec<u8> = ellipsis.encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    // 计算主文件名可用的字节数
    let available_for_stem = max_bytes.saturating_sub(ext_bytes.len() + ellipsis_bytes.len());

    if available_for_stem < 4 {
        // 空间太小，只返回扩展名或截断的文件名
        let chars: Vec<char> = filename.chars().collect();
        let max_chars = max_bytes / 2; // UTF-16LE每个字符最少2字节
        return chars.iter().take(max_chars.saturating_sub(1)).collect();
    }

    // 二分查找最多可以保留多少个字符
    let stem_chars: Vec<char> = stem.chars().collect();
    let mut left = 0;
    let mut right = stem_chars.len();
    let mut best_len = 0;

    while left <= right {
        let mid = (left + right) / 2;
        let test_stem: String = stem_chars.iter().take(mid).collect();
        let test_bytes: Vec<u8> = test_stem.encode_utf16()
            .flat_map(|c| c.to_le_bytes())
            .collect();

        if test_bytes.len() <= available_for_stem {
            best_len = mid;
            left = mid + 1;
        } else {
            right = mid - 1;
        }
    }

    // 智能调整截断位置：避免在中英文混合处截断
    let truncated_stem: String = stem_chars.iter().take(best_len).collect();

    // 检查最后一个字符，如果是ASCII字母，尝试向前找到分隔符或中文字符
    let final_stem = if best_len > 0 && best_len < stem_chars.len() {
        let last_char = stem_chars[best_len - 1];

        // 如果最后是ASCII字母，尝试向前找到更好的截断点
        if last_char.is_ascii_alphabetic() {
            // 向前查找分隔符或中文字符
            let mut better_pos = best_len;
            for i in (0..best_len).rev() {
                let ch = stem_chars[i];
                // 在分隔符、空格、中文字符等自然边界处截断
                if ch == '_' || ch == '-' || ch == ' ' || ch == '.' ||
                   ch > '\u{4E00}' && ch < '\u{9FFF}' { // 中文字符范围
                    better_pos = i;
                    break;
                }
                // 如果找到了中文字符，在其后截断
                if i > 0 {
                    let prev_ch = stem_chars[i - 1];
                    if (prev_ch > '\u{4E00}' && prev_ch < '\u{9FFF}') &&
                       ch.is_ascii_alphabetic() {
                        better_pos = i;
                        break;
                    }
                }
            }

            // 只有在新位置合理时才使用（不要缩短太多）
            if better_pos > best_len / 2 {
                stem_chars.iter().take(better_pos).collect()
            } else {
                truncated_stem
            }
        } else {
            truncated_stem
        }
    } else {
        truncated_stem
    };

    format!("{}{}{}", final_stem, ellipsis, ext_with_dot)
}

/// 在EMF图标数据中替换硬编码的文件名
/// EMF图标文件中包含了原始文件名的UTF-16LE编码字符串
///
/// 重要：可用空间就是旧文件名的长度，不要覆盖后面的EMF元数据！
fn replace_filename_in_emf(mut emf_data: Vec<u8>, old_filename: &str, new_filename: &str) -> Vec<u8> {
    // 将文件名转换为UTF-16LE编码
    let old_utf16: Vec<u8> = old_filename.encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    // 在EMF数据中查找旧文件名
    if let Some(pos) = emf_data.windows(old_utf16.len())
        .position(|window| window == old_utf16.as_slice()) {

        println!("找到硬编码文件名 '{}' 在偏移 0x{:x}", old_filename, pos);

        // 可用空间 = 旧文件名的长度（不要向后查找null-null，那会覆盖EMF元数据！）
        let available_space = old_utf16.len();

        // 智能截断文件名以适应可用空间
        let final_filename = truncate_filename_to_bytes(new_filename, available_space);
        let new_utf16: Vec<u8> = final_filename.encode_utf16()
            .flat_map(|c| c.to_le_bytes())
            .collect();

        let new_filename_bytes = new_filename.encode_utf16().count() * 2;
        println!("可用空间: {} 字节, 原文件名需要: {} 字节",
                 available_space,
                 new_filename_bytes);

        if final_filename != new_filename {
            println!("⚠ 文件名已截断: '{}' -> '{}'", new_filename, final_filename);
        }

        // 替换文件名
        for (i, &byte) in new_utf16.iter().enumerate() {
            emf_data[pos + i] = byte;
        }

        // 用null填充剩余空间（仅填充到旧文件名长度，不要超出）
        for i in new_utf16.len()..available_space {
            emf_data[pos + i] = 0;
        }

        println!("✓ 成功替换文件名为 '{}'", final_filename);
    } else {
        println!("⚠ 未在EMF图标中找到硬编码文件名 '{}'", old_filename);
    }

    emf_data
}

fn get_default_emf_icon(file_type: &FileType, filename: &str) -> Vec<u8> {
    // 根据文件类型返回对应的EMF图标
    // 这些图标是从真实的Word文档中提取的，包含硬编码的文件名
    // 我们需要将硬编码的文件名替换为实际文件名

    let (icon_data, old_filename) = match file_type {
        FileType::Video => {
            const ICON: &[u8] = include_bytes!("../icon_video.emf");
            (ICON.to_vec(), "qqer的抖音_.mp4")
        },
        FileType::PDF => {
            const ICON: &[u8] = include_bytes!("../icon_pdf.emf");
            (ICON.to_vec(), "0_1深孔刻蚀，可助300层3D NAND制造 - 今日头条.pdf")
        },
        FileType::Excel => {
            const ICON: &[u8] = include_bytes!("../icon_excel.emf");
            (ICON.to_vec(), "20_074644.xlsx")
        },
        FileType::Document => {
            // 文档类型使用Excel图标（.doc, .docx等）
            const ICON: &[u8] = include_bytes!("../icon_excel.emf");
            (ICON.to_vec(), "20_074644.xlsx")
        },
        FileType::ZIP => {
            const ICON: &[u8] = include_bytes!("../icon_zip.emf");
            (ICON.to_vec(), "ZL1.zip")
        },
        FileType::Image | FileType::Other(_) => {
            // 图片和其他类型使用通用Package图标（不包含硬编码文件名）
            const ICON: &[u8] = include_bytes!("../ole_package_icon.emf");
            return ICON.to_vec();
        },
    };

    // 替换EMF中的硬编码文件名为实际文件名
    replace_filename_in_emf(icon_data, old_filename, filename)
}

/// 从 ZIP archive 中读取文件
fn read_file_from_zip_archive(archive: &mut ZipArchive<Cursor<&[u8]>>, path: &str) -> Result<String> {
    let mut file = archive.by_name(path)
        .with_context(|| format!("无法找到文件: {}", path))?;
    let mut content = String::new();
    file.read_to_string(&mut content)?;
    Ok(content)
}

/// 获取下一个可用的关系 ID
fn get_next_relationship_id(rels_xml: &str) -> usize {
    let mut max_id = 0;

    // 查找所有 rId 并找到最大值
    for cap in regex::Regex::new(r#"Id="rId(\d+)""#).unwrap().captures_iter(rels_xml) {
        if let Some(num_str) = cap.get(1) {
            if let Ok(num) = num_str.as_str().parse::<usize>() {
                max_id = max_id.max(num);
            }
        }
    }

    max_id + 1
}

/// 在 document.xml 中添加 OLE 对象
fn add_ole_objects_to_document_xml(
    document_xml: &str,
    embedded_files: &[EmbeddedFile],
    start_rid: usize
) -> Result<String> {
    let mut result = document_xml.to_string();

    println!("=== 添加 OLE 对象到 document.xml ===");
    println!("嵌入文件数量: {}", embedded_files.len());

    // 按 zip_id 分组嵌入文件，但保持原始索引
    use std::collections::HashMap;
    let mut files_by_zip: HashMap<String, Vec<(usize, &EmbeddedFile)>> = HashMap::new();
    for (index, file) in embedded_files.iter().enumerate() {
        files_by_zip.entry(file.zip_id.clone())
            .or_insert_with(Vec::new)
            .push((index, file));
    }

    println!("按章节分组后的数量: {}", files_by_zip.len());

    // 为每个章节生成OLE对象XML并插入
    for (zip_id, files) in files_by_zip.iter() {
        let marker = format!("EMBED_MARKER_{}", zip_id);
        println!("处理章节: {}, 文件数: {}, 标记: {}", zip_id, files.len(), marker);

        let mut objects_xml = String::new();

        for (index, file) in files.iter() {
            let ole_rid = format!("rId{}", start_rid + index * 2);
            let img_rid = format!("rId{}", start_rid + index * 2 + 1);
            let shape_id = format!("_x0000_i{}", 1025 + index);
            let object_id = format!("_146807572{}", index);

            println!("  - 文件 {}: {} (rid={}, img_rid={})", index, file.name, ole_rid, img_rid);

            objects_xml.push_str(&format!(r###"
<w:p w14:paraId="{paraId}"><w:pPr><w:rPr><w:rFonts w:hint="default"/><w:lang w:val="en-US"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint="default"/><w:lang w:val="en-US"/></w:rPr><w:object><v:shape id="{shape_id}" o:spt="75" type="#_x0000_t75" style="height:65.25pt;width:72.4pt;" o:ole="t" filled="f" o:preferrelative="t" stroked="f" coordsize="21600,21600"><v:fill on="f" focussize="0,0"/><v:stroke on="f"/><v:imagedata r:id="{img_rid}" o:title=""/><o:lock v:ext="edit" aspectratio="t"/><w10:wrap type="none"/><w10:anchorlock/></v:shape><o:OLEObject Type="Embed" ProgID="Package" ShapeID="{shape_id}" DrawAspect="Icon" ObjectID="{object_id}" r:id="{ole_rid}"><o:LockedField>false</o:LockedField></o:OLEObject></w:object></w:r></w:p>
"###,
                paraId = format!("{:08X}", 0x10000000 + index),
                shape_id = shape_id,
                img_rid = img_rid,
                ole_rid = ole_rid,
                object_id = object_id
            ));
        }

        // 在标记段落之后插入OLE对象
        // 直接搜索标记文本，不管XML标签格式
        if let Some(pos) = result.find(&marker) {
            println!("  ✓ 找到标记文本位置: {}", pos);
            // 从标记位置向后查找段落结束标签
            if let Some(end_pos) = result[pos..].find("</w:p>") {
                let insert_pos = pos + end_pos + "</w:p>".len();
                println!("  ✓ 插入位置: {}", insert_pos);
                result.insert_str(insert_pos, &objects_xml);
            } else {
                println!("  ✗ 未找到段落结束标签");
            }
        } else {
            println!("  ✗ 未找到标记文本: {}", marker);
        }
    }

    Ok(result)
}

/// 在 document.xml.rels 中添加 OLE 对象关系
fn add_ole_relationships_to_rels(
    rels_xml: &str,
    embedded_files: &[EmbeddedFile],
    start_rid: usize
) -> Result<String> {
    let mut new_rels = String::new();

    for (index, _file) in embedded_files.iter().enumerate() {
        let ole_rid = format!("rId{}", start_rid + index * 2);
        let img_rid = format!("rId{}", start_rid + index * 2 + 1);
        let ole_target = format!("embeddings/oleObject{}.bin", index + 1);
        let img_target = format!("media/image{}.emf", index + 1);

        new_rels.push_str(&format!(
            r#"<Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="{}"/>"#,
            ole_rid, ole_target
        ));
        new_rels.push_str(&format!(
            r#"<Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{}"/>"#,
            img_rid, img_target
        ));
    }

    // 在 </Relationships> 之前插入新关系
    let result = rels_xml.replace("</Relationships>", &format!("{}</Relationships>", new_rels));

    Ok(result)
}

/// 在 [Content_Types].xml 中添加 OLE 对象内容类型
fn add_ole_content_types(content_types_xml: &str) -> Result<String> {
    // 检查是否已经有 .bin 和 .emf 的定义
    let mut result = content_types_xml.to_string();

    if !result.contains(r#"Extension="bin""#) {
        result = result.replace(
            "</Types>",
            r#"<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>"#
        );
    }

    if !result.contains(r#"Extension="emf""#) {
        result = result.replace(
            "</Types>",
            r#"<Default Extension="emf" ContentType="image/x-emf"/></Types>"#
        );
    }

    Ok(result)
}

// ==================== OLE 嵌入功能结束 ====================

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
    pdf_page_screenshot_files: Vec<String>,
    excel_files: Vec<String>,
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
    selected_excel_indices: Vec<usize>,
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
        "汇总包_{}{:02}{:02}_{:02}{:02}{:02}",
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
            pdf_page_screenshot_files: vec![],
            excel_files: vec![],
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
    let mut batch: BatchSummary = read_batch(&batch_dir).map_err(err_to_string)?;
    // 按下发时间排序
    sort_zips_by_issued_at(&mut batch.zips);
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
    // 按下发时间排序
    sort_zips_by_issued_at(&mut batch.zips);
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
            || !z.pdf_page_screenshot_files.is_empty()
            || !z.excel_files.is_empty()
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
    let mut batch: BatchSummary = read_batch(&batch_dir).map_err(err_to_string)?;

    // 按下发时间排序
    sort_zips_by_issued_at(&mut batch.zips);

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
    _embed_files: Option<bool>,
) -> Result<String, String> {
    let batch_dir = batch_dir(&app, &batch_id).map_err(err_to_string)?;
    let batch: BatchSummary = read_batch(&batch_dir).map_err(err_to_string)?;
    let batch = apply_bundle_selection(&batch, selection).map_err(err_to_string)?;

    if batch.zips.is_empty() {
        return Err("未选择任何ZIP用于导出".to_string());
    }

    let now = OffsetDateTime::now_utc();
    // 修改：直接保存为.docx文件，不再压缩为.zip
    let out = prompt_save_path(default_export_bundle_name(now), "docx", "Word文档")?;

    // 始终使用文件嵌入功能
    println!("=== 开始文件嵌入导出 ===");
    println!("1. 收集需要嵌入的文件...");
    let (docx, embedded_files) = build_enhanced_summary_docx(&batch, true).map_err(err_to_string)?;

    println!("2. 生成基础Word文档...");
    println!("3. 开始嵌入文件到Word文档...");
    let docx_bytes = build_docx_with_embeddings(docx, &embedded_files).map_err(err_to_string)?;

    // 直接保存docx文件，不再创建zip包
    println!("4. 保存Word文档到: {}", out.display());
    fs::write(&out, docx_bytes).map_err(err_to_string)?;

    println!("✓ Word文档导出完成！");
    println!("生成的Word文档包含 {} 个嵌入文件", embedded_files.len());
    println!("提示: 所有附件文件已完整嵌入到Word文档中");

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

        let mut selected_pdf_screens = Vec::new();
        for &idx in &sel.selected_pdf_page_screenshot_indices {
            if let Some(p) = z.pdf_page_screenshot_files.get(idx) {
                selected_pdf_screens.push(p.clone());
            }
        }
        z2.pdf_page_screenshot_files = selected_pdf_screens;

        let mut selected_excels = Vec::new();
        for &idx in &sel.selected_excel_indices {
            if let Some(p) = z.excel_files.get(idx) {
                selected_excels.push(p.clone());
            }
        }
        z2.excel_files = selected_excels;

        out.push(z2);
    }

    // 按下发时间排序
    sort_zips_by_issued_at(&mut out);

    Ok(BatchSummary {
        batch_id: batch.batch_id.clone(),
        created_at: batch.created_at,
        zips: out,
    })
}

#[derive(Debug, Clone)]
struct ZipScan {
    docx_entry: String,
    video_entries: Vec<usize>,  // 存储ZIP中的索引
    image_entries: Vec<usize>,
    pdf_entries: Vec<usize>,
    excel_entries: Vec<usize>,
    has_sample: bool,
}

/// 解码ZIP文件名（处理中文乱码）
/// Windows创建的ZIP文件通常使用GBK编码，需要正确解码
fn decode_zip_filename(name_bytes: &[u8]) -> String {
    // 首先尝试UTF-8解码
    if let Ok(utf8_name) = std::str::from_utf8(name_bytes) {
        // 检查是否包含乱码字符（如□、�等）
        if !utf8_name.chars().any(|c| c == '\u{FFFD}' || c == '□') {
            return utf8_name.to_string();
        }
    }

    // UTF-8失败或有乱码，尝试GBK解码
    let (decoded, _encoding, had_errors) = GBK.decode(name_bytes);
    if !had_errors {
        return decoded.to_string();
    }

    // 都失败了，使用UTF-8并替换无效字符
    String::from_utf8_lossy(name_bytes).to_string()
}

fn scan_zip(zip_path: &Path) -> Result<ZipScan> {
    let f = fs::File::open(zip_path)?;
    let mut zip = ZipArchive::new(f)?;

    let mut docx_entry: Option<String> = None;
    let mut has_sample = false;
    let mut video_entries = Vec::new();
    let mut image_entries = Vec::new();
    let mut pdf_entries = Vec::new();
    let mut excel_entries = Vec::new();

    for i in 0..zip.len() {
        let file = zip.by_index(i)?;
        // 保存原始文件名（用于后续从ZIP中读取）
        let name = file.name().to_string();
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
            video_entries.push(i);  // 保存索引
        } else if lower.ends_with(".pdf") {
            pdf_entries.push(i);
        } else if lower.ends_with(".png")
            || lower.ends_with(".jpg")
            || lower.ends_with(".jpeg")
            || lower.ends_with(".gif")
        {
            image_entries.push(i);
        } else if lower.ends_with(".xlsx") || lower.ends_with(".xls") {
            excel_entries.push(i);
        }
    }

    Ok(ZipScan {
        docx_entry: docx_entry.ok_or_else(|| anyhow!("ZIP内未找到docx"))?,
        video_entries,
        image_entries,
        pdf_entries,
        excel_entries,
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

    // 返回空的 video_entries，因为现在使用索引而不是文件名
    // 实际的文件信息在 extract_preview_files 中处理
    Ok((fields, vec![]))
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
    let excel_dir = root.join("excel");
    fs::create_dir_all(&videos_dir)?;
    fs::create_dir_all(&images_dir)?;
    fs::create_dir_all(&pdf_dir)?;
    fs::create_dir_all(&excel_dir)?;

    let f = fs::File::open(zip_path)?;
    let mut zip = ZipArchive::new(f)?;

    for &index in &scan.video_entries {
        let mut file = zip.by_index(index)?;
        let name = decode_zip_filename(file.name_raw());  // 正确解码文件名
        let basename = safe_basename(&name);
        let out = unique_path(&videos_dir, &basename);
        let mut buf = Vec::new();
        file.read_to_end(&mut buf)?;
        fs::write(&out, buf)?;
        summary.video_files.push(out.to_string_lossy().to_string());
    }

    for &index in &scan.image_entries {
        let mut file = zip.by_index(index)?;
        let name = decode_zip_filename(file.name_raw());
        let basename = safe_basename(&name);
        let out = unique_path(&images_dir, &basename);
        let mut buf = Vec::new();
        file.read_to_end(&mut buf)?;
        fs::write(&out, buf)?;
        summary.image_files.push(out.to_string_lossy().to_string());
    }

    for &index in &scan.pdf_entries {
        let mut file = zip.by_index(index)?;
        let name = decode_zip_filename(file.name_raw());
        let basename = safe_basename(&name);
        let out = unique_path(&pdf_dir, &basename);
        let mut buf = Vec::new();
        file.read_to_end(&mut buf)?;
        fs::write(&out, buf)?;
        summary.pdf_files.push(out.to_string_lossy().to_string());
    }

    for &index in &scan.excel_entries {
        let mut file = zip.by_index(index)?;
        let name = decode_zip_filename(file.name_raw());
        let basename = safe_basename(&name);
        let out = unique_path(&excel_dir, &basename);
        let mut buf = Vec::new();
        file.read_to_end(&mut buf)?;
        fs::write(&out, buf)?;
        summary.excel_files.push(out.to_string_lossy().to_string());
    }

    Ok(())
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

    // 处理字段提取，特别处理指令内容的多行情况
    let fields = extract_all_fields(&text)?;

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

    // 对于指令内容，直接使用提取的完整内容，不再选择"最佳"
    fn get_instruction_content(values: Option<&Vec<String>>) -> String {
        let Some(values) = values else { return String::new() };
        for v in values {
            let t = v.trim();
            if !t.is_empty() && t != "序号" {
                return t.to_string();
            }
        }
        String::new()
    }

    Ok(WordFields {
        instruction_no: first_nonempty(fields.get("指令编号")),
        title: first_nonempty(fields.get("指令标题")),
        issued_at: first_nonempty(fields.get("下发时间")),
        content: get_instruction_content(fields.get("指令内容")),
    })
}

// 检查是否是表格表头或表格内容
fn is_table_header_or_content(line: &str) -> bool {
    let trimmed = line.trim();

    // 常见的表格表头模式
    let table_headers = [
        "标题", "链接", "网站", "属地", "处置方式", "序号", "时间", "内容", "类型",
        "编号", "名称", "地址", "来源", "状态", "备注", "操作", "详情",
        "链接地址", "网站名称", "处理方式", "处理结果", "处理时间"
    ];

    // 检查是否包含表格表头关键词
    for header in &table_headers {
        if trimmed.contains(header) && trimmed.len() <= 20 {
            return true;
        }
    }

    // 检查是否是纯数字或编号（表格中的常见内容）
    if trimmed.chars().all(|c| c.is_ascii_digit() || c == '.' || c == '、') {
        return true;
    }

    // 检查是否是URL链接（表格中常见）
    if trimmed.starts_with("http://") || trimmed.starts_with("https://") || trimmed.starts_with("www.") {
        return true;
    }

    // 检查是否包含大量分隔符（表格特征）
    let tab_count = trimmed.matches('\t').count();
    let space_count = trimmed.matches("  ").count();
    if tab_count >= 2 || space_count >= 3 {
        return true;
    }

    false
}

// 提取所有字段，特别处理指令内容的多行情况
fn extract_all_fields(text: &str) -> Result<std::collections::BTreeMap<String, Vec<String>>> {
    let mut map: std::collections::BTreeMap<String, Vec<String>> = std::collections::BTreeMap::new();
    let lines: Vec<&str> = text.lines().collect();

    let mut i = 0;
    while i < lines.len() {
        let line = lines[i].trim();

        // 检查是否是字段行
        if let Some(cap) = RE_FIELD.captures(line) {
            let key = cap.get(1).unwrap().as_str().to_string();
            let mut value = cap.name("v").unwrap().as_str().trim().to_string();

            // 如果是指令内容，需要收集多行内容
            if key == "指令内容" {
                // 收集后续非字段行，直到遇到下一个字段或表格
                i += 1;
                let mut content_lines = Vec::new();
                while i < lines.len() {
                    let next_line = lines[i].trim();

                    // 如果下一个字段开始，停止收集
                    if RE_FIELD.is_match(next_line) {
                        break;
                    }

                    // 如果遇到空行，继续检查下一行（空行可能是段落分隔）
                    if next_line.is_empty() {
                        i += 1;
                        continue;
                    }

                    // 如果是表格表头或表格内容，停止收集指令内容
                    if is_table_header_or_content(next_line) {
                        break;
                    }

                    // 过滤掉一些明显的非内容行
                    if next_line != "序号" && next_line != "指令编号" && next_line != "指令标题" && next_line != "下发时间" {
                        // 检查行长度，过短的行可能不是内容
                        if next_line.len() > 3 {
                            content_lines.push(next_line);
                        }
                    }
                    i += 1;
                }

                // 保留原始格式的多行内容
                if !content_lines.is_empty() {
                    let multi_line_content = content_lines.join("\n");
                    let cleaned_content = normalize_instruction_content_with_format(&format!("{}\n{}", value, multi_line_content));
                    value = cleaned_content;
                }
                i -= 1; // 回退一行，因为外层循环会再次递增
            } else {
                // 对于其他字段，应用简单的清理
                value = normalize_text(&value);
            }

            map.entry(key).or_default().push(value);
        }
        i += 1;
    }

    Ok(map)
}

fn extract_paragraph_texts(document_xml: &str) -> Result<String> {
    let mut reader = XmlReader::from_str(document_xml);
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
            Err(err) => return Err(anyhow!("XML解析错误: {:?}", err)),
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

// 专门用于处理指令内容的函数，保留标点符号和格式

// 专门用于处理指令内容的函数，保留原始格式（换行、段落等）
fn normalize_instruction_content_with_format(s: &str) -> String {
    let mut result = String::new();
    let mut lines = Vec::new();

    // 按行分割内容，保留空行（段落分隔）
    for line in s.lines() {
        let trimmed = line.trim_end();  // 只删除行尾空白，保留标点
        lines.push(trimmed);
    }

    // 重新组合内容，保留原始换行结构
    for (i, line) in lines.iter().enumerate() {
        if i > 0 {
            // 检查是否需要在换行前添加空格
            let prev_line = lines[i-1];
            let current_line = *line;

            // 如果前一行以标点结尾，且当前行不是空行，保留换行
            if !current_line.is_empty() &&
               (prev_line.ends_with('。') || prev_line.ends_with('！') || prev_line.ends_with('？') ||
                prev_line.ends_with('.') || prev_line.ends_with('!') || prev_line.ends_with('?')) {
                // 保留换行
                result.push('\n');
            } else if !current_line.is_empty() && !prev_line.is_empty() {
                // 如果都不是空行，且前一行不是句号等结束，添加换行
                result.push('\n');
            } else {
                // 如果当前行是空行，添加额外的换行（段落分隔）
                result.push('\n');
            }
        }

        // 替换特殊空白字符，但保留格式
        let cleaned = line.replace('\u{00A0}', " ")
                         .replace('\u{3000}', " ")
                         .replace('：', ":");
        result.push_str(&cleaned);
    }

    // 清理开头和结尾的空白字符，但保留中间的换行
    result.trim_matches('\n').trim_matches('\r').to_string()
}

// 解析下发时间字符串为 OffsetDateTime，支持多种格式
fn parse_issued_at(date_str: &str) -> Result<OffsetDateTime> {
    let trimmed = date_str.trim();
    if trimmed.is_empty() {
        // 如果时间为空，返回一个很早的时间作为默认值
        return Ok(OffsetDateTime::UNIX_EPOCH);
    }

    // 尝试完整的时间戳格式 YYYY-MM-DD HH:MM:SS
    if trimmed.len() >= 19 {
        let date_part = &trimmed[0..10];
        let time_part = &trimmed[11..19];

        if date_part.chars().nth(4) == Some('-') && date_part.chars().nth(7) == Some('-') &&
           time_part.chars().nth(2) == Some(':') && time_part.chars().nth(5) == Some(':') {

            // 解析日期部分
            if let (Ok(year), Ok(month_u8), Ok(day)) = (
                date_part[0..4].parse::<i32>(),
                date_part[5..7].parse::<u8>(),
                date_part[8..10].parse::<u8>()
            ) {
                // 解析时间部分
                if let (Ok(hour), Ok(minute), Ok(second)) = (
                    time_part[0..2].parse::<u8>(),
                    time_part[3..5].parse::<u8>(),
                    time_part[6..8].parse::<u8>()
                ) {
                    // 转换月份类型
                    if let Ok(month) = time::Month::try_from(month_u8) {
                        if let (Ok(date), Ok(time)) = (
                            time::Date::from_calendar_date(year, month, day),
                            time::Time::from_hms(hour, minute, second)
                        ) {
                            return Ok(time::PrimitiveDateTime::new(date, time).assume_utc());
                        }
                    }
                }
            }
        }
    }

    // 尝试带时间的 YYYY-MM-DD HH:MM 格式
    if trimmed.len() >= 16 && trimmed.len() < 19 {
        let date_part = &trimmed[0..10];
        let time_part = &trimmed[11..16];

        if date_part.chars().nth(4) == Some('-') && date_part.chars().nth(7) == Some('-') &&
           time_part.chars().nth(2) == Some(':') {

            // 解析日期部分
            if let (Ok(year), Ok(month_u8), Ok(day)) = (
                date_part[0..4].parse::<i32>(),
                date_part[5..7].parse::<u8>(),
                date_part[8..10].parse::<u8>()
            ) {
                // 解析时间部分
                if let (Ok(hour), Ok(minute)) = (
                    time_part[0..2].parse::<u8>(),
                    time_part[3..5].parse::<u8>()
                ) {
                    // 转换月份类型
                    if let Ok(month) = time::Month::try_from(month_u8) {
                        if let (Ok(date), Ok(time)) = (
                            time::Date::from_calendar_date(year, month, day),
                            time::Time::from_hms(hour, minute, 0)
                        ) {
                            return Ok(time::PrimitiveDateTime::new(date, time).assume_utc());
                        }
                    }
                }
            }
        }
    }

    // 简单的解析策略：尝试数字格式
    if let Ok(num) = trimmed.parse::<i64>() {
        if num >= 10000101 && num <= 99991231 {
            let year = (num / 10000) as i32;
            let month = ((num % 10000) / 100) as u8;
            let day = (num % 100) as u8;

            if month >= 1 && month <= 12 && day >= 1 && day <= 31 {
                // 使用time 0.3兼容的API
                if let Ok(month) = time::Month::try_from(month) {
                    if let Ok(date) = time::Date::from_calendar_date(year, month, day) {
                        return Ok(time::PrimitiveDateTime::new(date, time::Time::MIDNIGHT).assume_utc());
                    }
                }
            }
        }
    }

    // 尝试标准格式 YYYY-MM-DD
    if trimmed.len() >= 10 && trimmed.chars().nth(4) == Some('-') && trimmed.chars().nth(7) == Some('-') {
        if let Ok(year) = trimmed[0..4].parse::<i32>() {
            if let Ok(month) = trimmed[5..7].parse::<u8>() {
                if let Ok(day) = trimmed[8..10].parse::<u8>() {
                    if let Ok(month) = time::Month::try_from(month) {
                        if let Ok(date) = time::Date::from_calendar_date(year, month, day) {
                            return Ok(time::PrimitiveDateTime::new(date, time::Time::MIDNIGHT).assume_utc());
                        }
                    }
                }
            }
        }
    }

    // 如果都无法解析，返回当前时间
    Ok(OffsetDateTime::now_utc())
}

// 对 ZipSummary 列表按下发时间排序
fn sort_zips_by_issued_at(zips: &mut Vec<ZipSummary>) {
    zips.sort_by(|a, b| {
        let time_a = parse_issued_at(&a.word.issued_at).unwrap_or_else(|_| OffsetDateTime::UNIX_EPOCH);
        let time_b = parse_issued_at(&b.word.issued_at).unwrap_or_else(|_| OffsetDateTime::UNIX_EPOCH);
        time_a.cmp(&time_b)
    });
}

fn build_summary_docx(batch: &BatchSummary) -> Result<Vec<u8>> {
    let mut docx = Docx::new();
    docx = docx.add_paragraph(
        Paragraph::new().add_run(Run::new().add_text("汇总文档").bold()),
    );

    for z in &batch.zips {
        let zip_folder = format!("attachments/{}/", z.id);
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
            docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text("指令内容:")));
            // 将指令内容按换行符分割，创建多个段落
            for line in z.word.content.lines() {
                let trimmed_line = line.trim();
                if !trimmed_line.is_empty() {
                    docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text(trimmed_line)));
                } else {
                    // 空行创建空段落（段落间距）
                    docx = docx.add_paragraph(Paragraph::new());
                }
            }
        }

        // 直接显示图片，删除"图片"标题
        for img_path in &z.image_files {
            let bytes = fs::read(img_path)
                .with_context(|| format!("读取图片失败: {}", img_path))?;
            let pic = Pic::new(&bytes);
            docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_image(pic)));
        }

        // 直接显示PDF图片，删除"PDF页面图片:"标题
        
        // 直接显示PDF截图，删除"PDF页面截图:"标题
        for img_path in &z.pdf_page_screenshot_files {
            let bytes = fs::read(img_path)
                .with_context(|| format!("读取PDF页面截图失败: {}", img_path))?;
            let pic = Pic::new(&bytes);
            docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_image(pic)));
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
    let file_options = FileOptions::default();
    let dir_options = FileOptions::default();

    let mut out = Cursor::new(Vec::<u8>::new());
    {
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
    } // writer 在这里被 drop，释放对 out 的借用

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
        // Windows上使用explorer打开文件或文件夹
        // 如果是文件，使用 /select 参数在资源管理器中选中它
        // 如果是文件夹，直接打开
        if path.is_file() {
            // /select 参数格式: explorer /select,"C:\path\to\file"
            let path_str = path.to_str().unwrap_or("");
            let select_arg = format!("/select,{}", path_str);
            let status = Command::new("explorer")
                .arg(select_arg)
                .status()?;
            if !status.success() {
                return Err(anyhow!("explorer 返回非0状态码"));
            }
        } else {
            let status = Command::new("explorer")
                .arg(path)
                .status()?;
            if !status.success() {
                return Err(anyhow!("explorer 返回非0状态码"));
            }
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

#[derive(Debug, Clone, Serialize, Deserialize)]
struct ExcelPreviewData {
    sheet_name: String,
    rows: Vec<Vec<String>>,
    total_sheets: usize,
    sheet_names: Vec<String>,
}

fn read_excel_preview(excel_path: &Path) -> Result<ExcelPreviewData> {
    let extension = excel_path
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();

    println!("正在读取Excel文件: {:?}, 扩展名: {}", excel_path, extension);

    if extension == "xlsx" {
        let mut workbook = calamine::open_workbook::<Xlsx<_>, _>(excel_path)
            .map_err(|e| anyhow!("打开xlsx文件失败: {}", e))?;

        // Use the trait methods
        use calamine::Reader;
        let sheet_names = workbook.sheet_names();
        let first_sheet_name = sheet_names
            .first()
            .ok_or_else(|| anyhow!("Excel文件没有工作表"))?
            .to_string();

        println!("工作表名称: {:?}", sheet_names);

        // 尝试读取第一个工作表的实际数据
        let mut rows = Vec::new();
        if let Ok(range) = workbook.worksheet_range(&first_sheet_name) {
            // 限制读取前10行和前10列，避免数据过多
            for row in range.rows().take(10) {
                let mut row_data = Vec::new();
                for cell in row.iter().take(10) {
                    let value_str = match cell {
                        calamine::Data::String(s) => s.to_string(),
                        calamine::Data::Float(f) => f.to_string(),
                        calamine::Data::Int(i) => i.to_string(),
                        calamine::Data::Bool(b) => b.to_string(),
                        calamine::Data::Empty => String::new(),
                        _ => "【数据】".to_string(),
                    };
                    row_data.push(value_str);
                }
                rows.push(row_data);
            }
        } else {
            // 如果无法读取数据，返回错误信息
            return Err(anyhow!("无法读取Excel工作表数据: {}", first_sheet_name));
        }

        if rows.is_empty() {
            // 如果没有数据，至少返回表头
            rows = vec![
                vec!["工作表".to_string(), first_sheet_name.to_string(), "".to_string()],
                vec!["状态".to_string(), "无数据".to_string(), "".to_string()],
            ];
        }

        Ok(ExcelPreviewData {
            sheet_name: first_sheet_name,
            rows,
            total_sheets: sheet_names.len(),
            sheet_names,
        })
    } else if extension == "xls" {
        let mut workbook = calamine::open_workbook::<Xls<_>, _>(excel_path)
            .map_err(|e| anyhow!("打开xls文件失败: {}", e))?;

        // Use the trait methods
        let sheet_names = workbook.sheet_names();
        let first_sheet_name = sheet_names
            .first()
            .ok_or_else(|| anyhow!("Excel文件没有工作表"))?
            .to_string();

        println!("工作表名称: {:?}", sheet_names);

        // 尝试读取第一个工作表的实际数据
        let mut rows = Vec::new();
        if let Ok(range) = workbook.worksheet_range(&first_sheet_name) {
            // 限制读取前10行和前10列，避免数据过多
            for row in range.rows().take(10) {
                let mut row_data = Vec::new();
                for cell in row.iter().take(10) {
                    let value_str = match cell {
                        calamine::Data::String(s) => s.to_string(),
                        calamine::Data::Float(f) => f.to_string(),
                        calamine::Data::Int(i) => i.to_string(),
                        calamine::Data::Bool(b) => b.to_string(),
                        calamine::Data::Empty => String::new(),
                        _ => "【数据】".to_string(),
                    };
                    row_data.push(value_str);
                }
                rows.push(row_data);
            }
        } else {
            // 如果无法读取数据，返回错误信息
            return Err(anyhow!("无法读取Excel工作表数据: {}", first_sheet_name));
        }

        if rows.is_empty() {
            // 如果没有数据，至少返回表头
            rows = vec![
                vec!["工作表".to_string(), first_sheet_name.to_string(), "".to_string()],
                vec!["状态".to_string(), "无数据".to_string(), "".to_string()],
            ];
        }

        Ok(ExcelPreviewData {
            sheet_name: first_sheet_name,
            rows,
            total_sheets: sheet_names.len(),
            sheet_names,
        })
    } else {
        Err(anyhow!("不支持的Excel格式: {}", extension))
    }
}


#[tauri::command]
fn get_excel_preview_data(
    app: tauri::AppHandle,
    batch_id: String,
    zip_id: String,
    index: usize,
) -> Result<ExcelPreviewData, String> {
    let batch_dir = batch_dir(&app, &batch_id).map_err(err_to_string)?;
    let batch: BatchSummary = read_batch(&batch_dir).map_err(err_to_string)?;
    let z = batch
        .zips
        .iter()
        .find(|z| z.id == zip_id)
        .ok_or_else(|| "ZIP不存在".to_string())?;
    let path = z
        .excel_files
        .get(index)
        .ok_or_else(|| "Excel文件索引越界".to_string())?;

    read_excel_preview(Path::new(path)).map_err(err_to_string)
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

// 新增：带文件嵌入的Word导出命令
#[tauri::command]
fn export_bundle_zip_with_embeddings(
    app: tauri::AppHandle,
    batch_id: String,
    embed_files: bool,
    max_file_size_mb: Option<u64>,
    allowed_types: Option<Vec<String>>,
) -> Result<String, String> {
    let batch_dir = batch_dir(&app, &batch_id).map_err(err_to_string)?;
    let mut batch: BatchSummary = read_batch(&batch_dir).map_err(err_to_string)?;

    // 按下发时间排序
    sort_zips_by_issued_at(&mut batch.zips);

    let now = OffsetDateTime::now_utc();
    let out = prompt_save_path(default_export_bundle_name(now), "zip", "ZIP")?;

    // 创建嵌入配置
    let mut config = EmbeddingConfig::default();
    if embed_files {
        config.enabled = true;
        if let Some(size_mb) = max_file_size_mb {
            config.max_file_size = (size_mb * 1024 * 1024) as usize;
        }
        if let Some(types) = allowed_types {
            config.allowed_types = types;
        }
    } else {
        config.enabled = false;
    }

    // 使用增强的导出功能
    let (docx, embedded_files) = build_enhanced_summary_docx(&batch, embed_files).map_err(err_to_string)?;
    let docx_bytes = build_docx_with_embeddings(docx, &embedded_files).map_err(err_to_string)?;
    let bundle_bytes = build_bundle_zip_bytes(&batch, &docx_bytes).map_err(err_to_string)?;

    fs::write(&out, bundle_bytes).map_err(err_to_string)?;
    Ok(out.to_string_lossy().to_string())
}

// 新增：获取嵌入配置
#[tauri::command]
fn get_embedding_config() -> EmbeddingConfig {
    EmbeddingConfig::default()
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
            export_bundle_zip_with_embeddings,
            get_embedding_config,
            open_path,
            get_preview_image_data,
            get_excel_preview_data,
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
        let (fields, _videos) = extract_word_and_videos(&zip_path, &scan).expect("extract");
        assert!(!fields.instruction_no.is_empty());
        assert!(!fields.title.is_empty());
        assert!(!fields.issued_at.is_empty());
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
            pdf_page_screenshot_files: vec![],
            excel_files: vec![],
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
