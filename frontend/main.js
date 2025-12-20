const tauriGlobal = window.__TAURI__ ?? {};
const { core } = tauriGlobal;

function getInvoke() {
  if (core?.invoke) return core.invoke;
  if (tauriGlobal?.invoke) return tauriGlobal.invoke;
  if (typeof window.__TAURI_INVOKE__ === "function") return window.__TAURI_INVOKE__;
  if (typeof window.__TAURI_INTERNALS__?.invoke === "function") return window.__TAURI_INTERNALS__.invoke;
  return null;
}

function getConvertFileSrc() {
  if (core?.convertFileSrc) return core.convertFileSrc;
  if (typeof window.__TAURI_INTERNALS__?.convertFileSrc === "function")
    return window.__TAURI_INTERNALS__.convertFileSrc;
  return null;
}

async function invoke(cmd, args) {
  const fn = getInvoke();
  if (!fn) throw new Error("Tauri API不可用（invoke缺失，请确认已启用 withGlobalTauri 或使用内置 __TAURI_INVOKE__）");
  return await fn(cmd, args);
}

const el = {
  pickZipsBtn: document.getElementById("pickZipsBtn"),
  exportExcelBtn: document.getElementById("exportExcelBtn"),
  exportBundleBtn: document.getElementById("exportBundleBtn"),
  zipList: document.getElementById("zipList"),
  details: document.getElementById("details"),
  status: document.getElementById("status"),
  imageModal: document.getElementById("imageModal"),
  imageModalBackdrop: document.getElementById("imageModalBackdrop"),
  imageModalCloseBtn: document.getElementById("imageModalCloseBtn"),
  imageModalOpenBtn: document.getElementById("imageModalOpenBtn"),
  imageModalImg: document.getElementById("imageModalImg"),
  imageModalTitle: document.getElementById("imageModalTitle"),
  searchInput: document.getElementById("searchInput"),
  selectAllBtn: document.getElementById("selectAllBtn"),
  invertSelBtn: document.getElementById("invertSelBtn"),
};

let state = {
  batchId: null,
  zips: [],
  selectedZipId: null,
  selection: {},
  imageDataCache: {},
  modal: { path: null, title: "" },
  pdfRendering: { busy: false },
  filter: "",
};

function setStatus(text) {
  el.status.textContent = text;
}

function openImageModal({ title, path, src }) {
  state.modal = { title, path };
  el.imageModalTitle.textContent = title ?? "";
  el.imageModalImg.src = src ?? "";
  el.imageModal.classList.remove("hidden");
}

function closeImageModal() {
  state.modal = { path: null, title: "" };
  el.imageModalImg.src = "";
  el.imageModal.classList.add("hidden");
}

el.imageModalBackdrop.onclick = closeImageModal;
el.imageModalCloseBtn.onclick = closeImageModal;
window.addEventListener("keydown", (e) => {
  if (e.key === "Escape") closeImageModal();
});
el.imageModalOpenBtn.onclick = async () => {
  try {
    if (!state.modal?.path) return;
    await invoke("open_path", { path: state.modal.path });
  } catch (e) {
    setStatus(`打开失败：${e?.message ?? e}`);
  }
};

function basename(p) {
  if (!p) return "";
  const normalized = String(p).replaceAll("\\", "/");
  const parts = normalized.split("/");
  return parts[parts.length - 1] || normalized;
}

function fileSrc(p) {
  if (!p) return "";
  const convertFileSrc = getConvertFileSrc();
  if (convertFileSrc) return convertFileSrc(p);
  return `file://${encodeURI(p)}`;
}

function getPdfJs() {
  // pdf.min.js (UMD) exports `pdfjsLib` global
  return window.pdfjsLib ?? null;
}

function ensurePdfJs() {
  const pdfjsLib = getPdfJs();
  if (!pdfjsLib) throw new Error("PDF.js 未加载（pdfjsLib 缺失）");
  if (!pdfjsLib.GlobalWorkerOptions?.workerSrc) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = "./vendor/pdfjs/pdf.worker.min.js";
  }
  return pdfjsLib;
}

async function renderPdfToPngDataUrls(pdfPath, { maxPages = 50 } = {}) {
  const pdfjsLib = ensurePdfJs();
  const url = fileSrc(pdfPath);
  const ab = await fetch(url).then((r) => {
    if (!r.ok) throw new Error(`读取PDF失败：${r.status}`);
    return r.arrayBuffer();
  });
  const loadingTask = pdfjsLib.getDocument({ data: ab });
  const doc = await loadingTask.promise;
  const numPages = Math.min(doc.numPages, maxPages);

  const out = [];
  for (let p = 1; p <= numPages; p++) {
    const page = await doc.getPage(p);
    const viewport1 = page.getViewport({ scale: 1 });
    const maxDim = 1200;
    const scale = Math.min(2.0, Math.max(1.0, maxDim / Math.max(viewport1.width, viewport1.height)));
    const viewport = page.getViewport({ scale });

    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    await page.render({ canvasContext: ctx, viewport }).promise;
    out.push(canvas.toDataURL("image/png"));
    page.cleanup();
  }
  await doc.cleanup();
  return out;
}

async function autoGeneratePdfScreenshots() {
  if (!state.batchId) return;
  const pdfjsLib = getPdfJs();
  if (!pdfjsLib) {
    setStatus("提示：PDF.js未加载，无法自动生成PDF页面截图");
    return;
  }
  if (state.pdfRendering.busy) return;
  state.pdfRendering.busy = true;
  try {
    for (const z of state.zips) {
      if (z.status?.startsWith?.("failed")) continue;
      if (!z.pdf_files?.length) continue;
      // 如果已有截图，先不重复生成（避免重复与耗时）
      if ((z.pdf_page_screenshot_files?.length ?? 0) > 0) continue;

      for (const pdfPath of z.pdf_files) {
        setStatus(`正在自动生成PDF页面截图：${z.filename} / ${basename(pdfPath)} …`);
        const dataUrls = await renderPdfToPngDataUrls(pdfPath, { maxPages: 20 });
        const saved = await invoke("save_pdf_page_screenshots", {
          batchId: state.batchId,
          zipId: z.id,
          pdfName: basename(pdfPath),
          screenshots: dataUrls,
        });
        z.pdf_page_screenshot_files = [...(z.pdf_page_screenshot_files ?? []), ...saved];
        const sel = state.selection[z.id];
        sel.pdfScreens = [...(sel.pdfScreens ?? []), ...saved.map(() => true)];
        if (state.selectedZipId === z.id) {
          await renderDetails();
        }
      }
    }
    setStatus("PDF页面截图自动生成完成");
  } catch (e) {
    console.error(e);
    setStatus(`PDF页面截图生成失败：${e?.message ?? e}`);
  } finally {
    state.pdfRendering.busy = false;
  }
}

function initSelectionsForBatch() {
  state.selection = {};
  state.imageDataCache = {};
  for (const z of state.zips) {
    state.selection[z.id] = {
      include: true,
      includeOriginalZip: false,
      videos: (z.video_files ?? []).map(() => true),
      images: (z.image_files ?? []).map(() => true),
      pdfFiles: (z.pdf_files ?? []).map(() => false),
      pdfScreens: (z.pdf_page_screenshot_files ?? []).map(() => true),
      excels: (z.excel_files ?? []).map(() => true),
    };
    state.imageDataCache[z.id] = {};
  }
}

function selectedIndices(flags) {
  const out = [];
  for (let i = 0; i < flags.length; i++) if (flags[i]) out.push(i);
  return out;
}

el.searchInput.oninput = (e) => {
  state.filter = e.target.value.trim().toLowerCase();
  renderList();
};

function getVisibleZips() {
  if (!state.filter) return state.zips;
  return state.zips.filter(z => z.filename.toLowerCase().includes(state.filter));
}

el.selectAllBtn.onclick = () => {
  const targets = getVisibleZips();
  for (const z of targets) {
    if (state.selection[z.id]) state.selection[z.id].include = true;
  }
  renderList();
};

el.invertSelBtn.onclick = () => {
  const targets = getVisibleZips();
  for (const z of targets) {
    if (state.selection[z.id]) state.selection[z.id].include = !state.selection[z.id].include;
  }
  renderList();
};

function renderList() {
  el.zipList.innerHTML = "";
  let visibleZips = getVisibleZips();

  for (const z of visibleZips) {
    const row = document.createElement("div");
    row.className = "list-row";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = state.selection?.[z.id]?.include ?? true;
    checkbox.onclick = (e) => e.stopPropagation(); // prevent row click
    checkbox.onchange = () => {
      state.selection[z.id].include = checkbox.checked;
      renderList();
    };

    const item = document.createElement("button");
    item.className = "list-item" + (z.id === state.selectedZipId ? " active" : "");
    item.textContent = z.filename;
    item.title = z.filename;

    // Row click
    row.onclick = () => {
      state.selectedZipId = z.id;
      renderList();
      renderDetails();
    };

    // Delete Button
    const delBtn = document.createElement("button");
    delBtn.className = "item-del-btn";
    delBtn.innerHTML = "×"; // Use innerHTML for better char rendering if needed
    delBtn.title = "移除此项";
    delBtn.onclick = (e) => {
      e.stopPropagation(); // Definitely stop propagation

      // Update State
      state.zips = state.zips.filter(x => x.id !== z.id);
      delete state.selection[z.id];
      delete state.imageDataCache[z.id];

      // Update Selection
      if (state.selectedZipId === z.id) {
        state.selectedZipId = state.zips[0]?.id || null;
      }
      renderDetails();
      renderList();
      setStatus(`已移除 ${z.filename}`);
    };

    row.appendChild(checkbox);
    row.appendChild(item);
    row.appendChild(delBtn);
    el.zipList.appendChild(row);
  }
  const anyIncluded = state.zips.some((z) => state.selection?.[z.id]?.include);
  el.exportExcelBtn.disabled = !state.batchId || state.zips.length === 0 || !anyIncluded;
  el.exportBundleBtn.disabled = !state.batchId || state.zips.length === 0 || !anyIncluded;
}

async function loadImageData(zipId, index) {
  if (!state.imageDataCache[zipId]) state.imageDataCache[zipId] = {};
  if (state.imageDataCache[zipId][index]) return state.imageDataCache[zipId][index];
  const dataUrl = await invoke("get_preview_image_data", {
    batchId: state.batchId,
    zipId,
    index,
  });
  state.imageDataCache[zipId][index] = dataUrl;
  return dataUrl;
}

async function loadExcelPreviewData(zipId, index) {
  if (!state.imageDataCache[zipId]) state.imageDataCache[zipId] = {};
  const key = `excel:${index}`;
  if (state.imageDataCache[zipId][key]) return state.imageDataCache[zipId][key];

  const data = await invoke("get_excel_preview_data", {
    batchId: state.batchId,
    zipId,
    index,
  });
  state.imageDataCache[zipId][key] = data;
  return data;
}

function renderExcelTable(excelData) {
  const container = document.createElement("div");
  container.className = "excel-preview";

  const info = document.createElement("div");
  info.className = "excel-info small";
  info.textContent = `工作表: ${excelData.sheet_name} (共${excelData.total_sheets}个表，显示前10行)`;
  container.appendChild(info);

  const table = document.createElement("table");
  table.className = "excel-table";

  for (const row of excelData.rows) {
    const tr = document.createElement("tr");
    for (const cell of row) {
      const td = document.createElement("td");
      td.textContent = cell || "";
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }

  container.appendChild(table);
  return container;
}

function section(title) {
  const root = document.createElement("div");
  root.className = "section";
  const h = document.createElement("h3");
  h.textContent = title;
  root.appendChild(h);
  return root;
}

function addSelectAllInvert(root, { onAll, onInvert, disabled }) {
  const row = document.createElement("div");
  row.className = "row";
  const allBtn = document.createElement("button");
  allBtn.textContent = "全选";
  allBtn.disabled = !!disabled;
  allBtn.onclick = onAll;
  const invBtn = document.createElement("button");
  invBtn.textContent = "反选";
  invBtn.disabled = !!disabled;
  invBtn.onclick = onInvert;
  row.appendChild(allBtn);
  row.appendChild(invBtn);
  root.appendChild(row);
}

function setAll(flags, value) {
  for (let i = 0; i < flags.length; i++) flags[i] = value;
}

function invertAll(flags) {
  for (let i = 0; i < flags.length; i++) flags[i] = !flags[i];
}

function addKV(root, kvs) {
  const grid = document.createElement("div");
  grid.className = "kv";
  for (const [k, v] of kvs) {
    const kEl = document.createElement("div");
    kEl.className = "k";
    kEl.textContent = k;
    const vEl = document.createElement("div");
    vEl.textContent = v ?? "";
    grid.appendChild(kEl);
    grid.appendChild(vEl);
  }
  root.appendChild(grid);
}

async function renderDetails() {
  const z = state.zips.find((x) => x.id === state.selectedZipId);
  el.details.innerHTML = "";
  if (!z) return;
  const sel = state.selection[z.id];

  const meta = section("基本信息");

  // 1. Header with Title and Badges
  const header = document.createElement("div");
  header.className = "info-header";

  const title = document.createElement("div");
  title.className = "info-title";
  title.textContent = z.filename;
  title.title = z.filename; // tooltip for long names
  header.appendChild(title);

  const statusBadge = document.createElement("span");
  statusBadge.className = `badge status-${z.status === 'completed' ? 'success' : 'default'}`;
  statusBadge.textContent = z.status;
  header.appendChild(statusBadge);

  if (z.has_sample) {
    const sampleBadge = document.createElement("span");
    sampleBadge.className = "badge warning";
    sampleBadge.textContent = "含样本";
    header.appendChild(sampleBadge);
  }
  meta.appendChild(header);

  // 2. Stats Grid
  const statsGrid = document.createElement("div");
  statsGrid.className = "stats-grid";

  const stats = [
    { label: "视频", value: z.video_files?.length ?? 0, icon: "🎬" },
    { label: "图片", value: z.image_files?.length ?? 0, icon: "🖼️" },
    { label: "PDF", value: z.pdf_files?.length ?? 0, icon: "📄" },
    { label: "Excel", value: z.excel_files?.length ?? 0, icon: "📊" },
    { label: "PDF截图", value: z.pdf_page_screenshot_files?.length ?? 0, icon: "📸" },
  ];

  stats.forEach(stat => {
    const box = document.createElement("div");
    box.className = "stat-box";
    const val = document.createElement("div");
    val.className = "stat-value";
    val.textContent = stat.value;
    const lbl = document.createElement("div");
    lbl.className = "stat-label";
    lbl.textContent = `${stat.icon} ${stat.label}`;
    box.appendChild(val);
    box.appendChild(lbl);
    statsGrid.appendChild(box);
  });
  meta.appendChild(statsGrid);

  // 3. Actions Row
  const actionsBar = document.createElement("div");
  actionsBar.className = "actions-bar";

  const originalZipCb = document.createElement("input");
  originalZipCb.type = "checkbox";
  originalZipCb.id = "cb_original_zip";
  originalZipCb.checked = sel.includeOriginalZip ?? false;
  originalZipCb.onchange = () => {
    sel.includeOriginalZip = originalZipCb.checked;
  };

  const originalZipLabel = document.createElement("label");
  originalZipLabel.htmlFor = "cb_original_zip";
  originalZipLabel.className = "checkbox-label";
  originalZipLabel.textContent = "导出原始ZIP";

  const openExtracted = document.createElement("button");
  openExtracted.textContent = "📂 解压目录";
  openExtracted.onclick = async () => {
    try {
      if (!z.extracted_dir) throw new Error("无解压目录");
      await invoke("open_path", { path: z.extracted_dir });
    } catch (e) {
      setStatus(`打开失败：${e?.message ?? e}`);
    }
  };

  const openStored = document.createElement("button");
  openStored.textContent = "📦 ZIP副本";
  openStored.onclick = async () => {
    try {
      await invoke("open_path", { path: z.stored_path || z.source_path });
    } catch (e) {
      setStatus(`打开失败：${e?.message ?? e}`);
    }
  };

  const leftGroup = document.createElement("div");
  leftGroup.className = "action-group";
  leftGroup.appendChild(originalZipCb);
  leftGroup.appendChild(originalZipLabel);

  const rightGroup = document.createElement("div");
  rightGroup.className = "action-group";
  rightGroup.appendChild(openExtracted);
  rightGroup.appendChild(openStored);

  actionsBar.appendChild(leftGroup);
  actionsBar.appendChild(rightGroup);
  meta.appendChild(actionsBar);

  el.details.appendChild(meta);

  const word = section("Word字段（固定模板抽取）");
  addKV(word, [
    ["指令编号", z.word?.instruction_no ?? ""],
    ["指令标题", z.word?.title ?? ""],
    ["下发时间", z.word?.issued_at ?? ""],
    ["指令内容", z.word?.content ?? ""],
  ]);
  el.details.appendChild(word);

  const images = section("图片预览（从ZIP直接解压）");
  const thumbs = document.createElement("div");
  thumbs.className = "thumbs";
  const imageFiles = z.image_files ?? [];
  addSelectAllInvert(images, {
    disabled: imageFiles.length === 0,
    onAll: () => {
      setAll(sel.images, true);
      renderDetails();
    },
    onInvert: () => {
      invertAll(sel.images);
      renderDetails();
    },
  });
  if (!imageFiles.length) {
    const p = document.createElement("div");
    p.className = "small";
    p.textContent = "无图片";
    images.appendChild(p);
  } else {
    for (let i = 0; i < imageFiles.length; i++) {
      const card = document.createElement("div");
      card.className = "thumb";
      const row = document.createElement("div");
      row.className = "row";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = sel.images[i];
      cb.onchange = () => {
        sel.images[i] = cb.checked;
      };
      const name = document.createElement("div");
      name.textContent = basename(imageFiles[i]);
      name.className = "small";
      const openBtn = document.createElement("button");
      openBtn.textContent = "打开";
      openBtn.onclick = async () => {
        try {
          await invoke("open_path", { path: imageFiles[i] });
        } catch (e) {
          setStatus(`打开失败：${e?.message ?? e}`);
        }
      };
      row.appendChild(cb);
      row.appendChild(openBtn);
      row.appendChild(name);
      card.appendChild(row);

      const img = document.createElement("img");
      img.alt = basename(imageFiles[i]);
      img.src = "";
      img.style.cursor = "pointer";
      card.appendChild(img);

      thumbs.appendChild(card);

      loadImageData(z.id, i)
        .then((dataUrl) => {
          img.src = dataUrl;
          img.onclick = () => openImageModal({ title: basename(imageFiles[i]), path: imageFiles[i], src: dataUrl });
        })
        .catch(() => {
          img.src = "";
        });
    }
    images.appendChild(thumbs);
  }
  el.details.appendChild(images);

  const videos = section("视频（可预览；失败可系统打开）");
  const videoFiles = z.video_files ?? [];
  addSelectAllInvert(videos, {
    disabled: videoFiles.length === 0,
    onAll: () => {
      setAll(sel.videos, true);
      renderDetails();
    },
    onInvert: () => {
      invertAll(sel.videos);
      renderDetails();
    },
  });
  if (!videoFiles.length) {
    const p = document.createElement("div");
    p.className = "small";
    p.textContent = "无视频";
    videos.appendChild(p);
  } else {
    for (let i = 0; i < videoFiles.length; i++) {
      const row = document.createElement("div");
      row.className = "row";

      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = sel.videos[i];
      cb.onchange = () => {
        sel.videos[i] = cb.checked;
      };

      const openBtn = document.createElement("button");
      openBtn.textContent = "系统打开";
      openBtn.onclick = async () => {
        try {
          await invoke("open_path", { path: videoFiles[i] });
        } catch (e) {
          setStatus(`打开失败：${e?.message ?? e}`);
        }
      };

      const label = document.createElement("div");
      label.textContent = basename(videoFiles[i]);
      label.className = "small";

      row.appendChild(cb);
      row.appendChild(openBtn);
      row.appendChild(label);
      videos.appendChild(row);

      // 尝试内嵌播放（不保证所有环境可用）
      const v = document.createElement("video");
      v.controls = true;
      v.style.width = "100%";
      v.style.maxHeight = "180px";
      v.src = fileSrc(videoFiles[i]);
      videos.appendChild(v);
    }
  }
  el.details.appendChild(videos);

  const pdfs = section("PDF文件（系统打开）");
  const pdfFiles = z.pdf_files ?? [];
  addSelectAllInvert(pdfs, {
    disabled: pdfFiles.length === 0,
    onAll: () => {
      setAll(sel.pdfFiles, true);
      renderDetails();
    },
    onInvert: () => {
      invertAll(sel.pdfFiles);
      renderDetails();
    },
  });
  if (!pdfFiles.length) {
    const p = document.createElement("div");
    p.className = "small";
    p.textContent = "无PDF";
    pdfs.appendChild(p);
  } else {
    for (let i = 0; i < pdfFiles.length; i++) {
      const row = document.createElement("div");
      row.className = "row";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = sel.pdfFiles[i];
      cb.onchange = () => {
        sel.pdfFiles[i] = cb.checked;
      };

      const openBtn = document.createElement("button");
      openBtn.textContent = "系统打开";
      openBtn.onclick = async () => {
        try {
          await invoke("open_path", { path: pdfFiles[i] });
        } catch (e) {
          setStatus(`打开失败：${e?.message ?? e}`);
        }
      };
      const label = document.createElement("div");
      label.textContent = basename(pdfFiles[i]);
      label.className = "small";
      row.appendChild(cb);
      row.appendChild(openBtn);
      row.appendChild(label);
      pdfs.appendChild(row);
    }
  }
  el.details.appendChild(pdfs);

  const pdfScreens = section("PDF页面截图");
  const pdfScreenFiles = z.pdf_page_screenshot_files ?? [];
  addSelectAllInvert(pdfScreens, {
    disabled: pdfScreenFiles.length === 0,
    onAll: () => {
      setAll(sel.pdfScreens, true);
      renderDetails();
    },
    onInvert: () => {
      invertAll(sel.pdfScreens);
      renderDetails();
    },
  });
  if (!pdfScreenFiles.length) {
    const p = document.createElement("div");
    p.className = "small";
    p.textContent = "无PDF页面截图（导入后会自动生成；若PDF较大可能需要等待）";
    pdfScreens.appendChild(p);
  } else {
    const thumbs3 = document.createElement("div");
    thumbs3.className = "thumbs";
    for (let i = 0; i < pdfScreenFiles.length; i++) {
      const card = document.createElement("div");
      card.className = "thumb";
      const row = document.createElement("div");
      row.className = "row";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = sel.pdfScreens[i];
      cb.onchange = () => {
        sel.pdfScreens[i] = cb.checked;
      };
      const name = document.createElement("div");
      name.textContent = basename(pdfScreenFiles[i]);
      name.className = "small";
      const openBtn = document.createElement("button");
      openBtn.textContent = "打开";
      openBtn.onclick = async () => {
        try {
          await invoke("open_path", { path: pdfScreenFiles[i] });
        } catch (e) {
          setStatus(`打开失败：${e?.message ?? e}`);
        }
      };
      row.appendChild(cb);
      row.appendChild(openBtn);
      row.appendChild(name);
      card.appendChild(row);

      const img = document.createElement("img");
      img.alt = basename(pdfScreenFiles[i]);
      img.src = fileSrc(pdfScreenFiles[i]);
      img.style.cursor = "pointer";
      img.onclick = () =>
        openImageModal({ title: basename(pdfScreenFiles[i]), path: pdfScreenFiles[i], src: fileSrc(pdfScreenFiles[i]) });
      card.appendChild(img);

      thumbs3.appendChild(card);
    }
    pdfScreens.appendChild(thumbs3);
  }
  el.details.appendChild(pdfScreens);

  // Excel section
  const excels = section("Excel文件");
  const excelFiles = z.excel_files ?? [];
  addSelectAllInvert(excels, {
    disabled: excelFiles.length === 0,
    onAll: () => {
      setAll(sel.excels, true);
      renderDetails();
    },
    onInvert: () => {
      invertAll(sel.excels);
      renderDetails();
    },
  });
  if (!excelFiles.length) {
    const p = document.createElement("div");
    p.className = "small";
    p.textContent = "无Excel文件";
    excels.appendChild(p);
  } else {
    for (let i = 0; i < excelFiles.length; i++) {
      const card = document.createElement("div");
      card.className = "excel-card";

      const row = document.createElement("div");
      row.className = "row";

      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = sel.excels[i];
      cb.onchange = () => {
        sel.excels[i] = cb.checked;
      };

      const name = document.createElement("div");
      name.textContent = basename(excelFiles[i]);
      name.className = "small";

      const openBtn = document.createElement("button");
      openBtn.textContent = "系统打开";
      openBtn.onclick = async () => {
        try {
          await invoke("open_path", { path: excelFiles[i] });
        } catch (e) {
          setStatus(`打开失败：${e?.message ?? e}`);
        }
      };

      const previewBtn = document.createElement("button");
      previewBtn.textContent = "预览";
      previewBtn.onclick = async () => {
        try {
          const previewContainer = card.querySelector(".excel-preview-container");
          if (previewContainer.style.display === "block") {
            previewContainer.style.display = "none";
            previewBtn.textContent = "预览";
          } else {
            if (!previewContainer.hasChildNodes()) {
              setStatus("正在加载Excel预览...");
              const data = await loadExcelPreviewData(z.id, i);
              const table = renderExcelTable(data);
              previewContainer.appendChild(table);
              setStatus("Excel预览加载完成");
            }
            previewContainer.style.display = "block";
            previewBtn.textContent = "收起";
          }
        } catch (e) {
          setStatus(`预览失败：${e?.message ?? e}`);
        }
      };

      row.appendChild(cb);
      row.appendChild(openBtn);
      row.appendChild(previewBtn);
      row.appendChild(name);
      card.appendChild(row);

      const previewContainer = document.createElement("div");
      previewContainer.className = "excel-preview-container";
      previewContainer.style.display = "none";
      card.appendChild(previewContainer);

      excels.appendChild(card);
    }
  }
  el.details.appendChild(excels);
}

el.pickZipsBtn.onclick = async () => {
  try {
    setStatus("正在选择ZIP…");
    const paths = await invoke("pick_zip_files", {});
    if (!paths?.length) {
      setStatus("已取消");
      return;
    }
    setStatus(`已选择${paths.length}个ZIP，正在导入解析…`);
    const result = await invoke("import_zips", { paths });
    state.batchId = result.batch_id;
    state.zips = result.zips;
    state.selectedZipId = state.zips[0]?.id ?? null;
    initSelectionsForBatch();
    renderList();
    await renderDetails();
    setStatus(`导入完成：批次 ${state.batchId}，正在自动生成PDF页面截图…`);
    autoGeneratePdfScreenshots();
  } catch (e) {
    console.error(e);
    setStatus(`错误：${e?.message ?? e}`);
  }
};

el.exportExcelBtn.onclick = async () => {
  try {
    if (!state.batchId) return;
    setStatus("正在导出Excel…");
    const zipIds = state.zips
      .filter((z) => state.selection[z.id]?.include ?? true)
      .map((z) => z.id);
    const outPath = await invoke("export_excel_with_selection", {
      batchId: state.batchId,
      zipIds,
    });
    setStatus(`Excel已导出：${outPath}`);
  } catch (e) {
    console.error(e);
    setStatus(`导出失败：${e?.message ?? e}`);
  }
};

el.exportBundleBtn.onclick = async () => {
  try {
    if (!state.batchId) return;

    setStatus("正在导出Word文档（文件嵌入模式）...");

    const selection = {
      zips: state.zips.map((z) => ({
        zip_id: z.id,
        include: state.selection[z.id]?.include ?? true,
        include_original_zip: state.selection[z.id]?.includeOriginalZip ?? false,
        selected_video_indices: selectedIndices(state.selection[z.id]?.videos ?? []),
        selected_image_indices: selectedIndices(state.selection[z.id]?.images ?? []),
        selected_pdf_indices: selectedIndices(state.selection[z.id]?.pdfFiles ?? []),
        selected_excel_indices: selectedIndices(state.selection[z.id]?.excels ?? []),
        selected_pdf_page_screenshot_indices: selectedIndices(state.selection[z.id]?.pdfScreens ?? []),
      })),
    };

    const outPath = await invoke("export_bundle_zip_with_selection", {
      batchId: state.batchId,
      selection,
      embedFiles: true,
    });

    setStatus(`Word文档已导出：${outPath}`);
  } catch (e) {
    console.error(e);
    setStatus(`导出失败：${e?.message ?? e}`);
  }
};

renderList();
renderDetails();
