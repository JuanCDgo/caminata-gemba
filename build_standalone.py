"""Build a 100% self-contained, offline single-file HTML app: Caminata_Gemba_App.html."""

from __future__ import annotations

import base64
from pathlib import Path

BASE = Path(__file__).resolve().parent
JSZIP = (BASE / "vendor" / "jszip.min.js").read_text(encoding="utf-8")
ENGINE = (BASE / "static" / "browser_pptx.js").read_text(encoding="utf-8")
STYLES = (BASE / "static" / "styles.css").read_text(encoding="utf-8")
LOGO_B64 = base64.b64encode((BASE / "static" / "logo.png").read_bytes()).decode("ascii")
LOGO_URI = f"data:image/png;base64,{LOGO_B64}"

APP_JS = r"""
(() => {
  const LOGO_URI = window.__GEMBA_LOGO__;
  const SECTOR_OWNERS = {
    "Embalaje - Pulido": "Tino Varela",
    "LIVING": "ARIEL CARANDINO",
    "MAQUINADO": "ARIEL DEHESA",
    "ARMADO": "LUCAS LUQUE",
    "MAQUINADO - ARMADO": "ARIEL DEHESA – LUCAS LUQUE",
    "MAQUINADO – ESQUELETOS": "ARIEL DEHESA – ARIEL CARANDINO",
  };

  const state = {
    zip: null,
    fileName: "Caminata Gemba.pptx",
    fileSizeMB: 0,
    info: null,
    dirty: false,
    antesFiles: [],
    despuesDrafts: new Map(), // slideIndex -> { file?, previewUrl?, note }
    mediaUrlCache: new Map(),
  };

  // DOM refs
  const fileSummary = document.getElementById("file-summary");
  const statusBanner = document.getElementById("status-banner");
  const pptxInput = document.getElementById("pptx-upload-input");
  const pptxDropCard = document.getElementById("pptx-drop-card");
  const appWorkspace = document.getElementById("app-workspace");
  const btnDownloadPptx = document.getElementById("btn-download-pptx");

  const tabBtnAntes = document.getElementById("tab-btn-antes");
  const tabBtnDespues = document.getElementById("tab-btn-despues");
  const panelAntes = document.getElementById("panel-antes");
  const panelDespues = document.getElementById("panel-despues");

  const walkForm = document.getElementById("walk-form");
  const walkDate = document.getElementById("walk-date");
  const walkSector = document.getElementById("walk-sector");
  const walkSchedule = document.getElementById("walk-schedule");
  const walkOwner = document.getElementById("walk-owner");
  const coverCanvas = document.getElementById("cover-canvas");
  const antesDropzone = document.getElementById("antes-dropzone");
  const antesFilesInput = document.getElementById("antes-files-input");
  const antesGrid = document.getElementById("antes-grid");
  const antesCountBadge = document.getElementById("antes-count-badge");
  const btnClearAntes = document.getElementById("btn-clear-antes");
  const btnSubmitWalk = document.getElementById("btn-submit-walk");

  const walkSelector = document.getElementById("walk-selector");
  const filterPending = document.getElementById("filter-pending");
  const walkCoverStrip = document.getElementById("walk-cover-strip");
  const walkCoverImg = document.getElementById("walk-cover-img");
  const walkStats = document.getElementById("walk-stats");
  const despuesList = document.getElementById("despues-list");
  const despuesSaveBar = document.getElementById("despues-save-bar");
  const pendingChangesText = document.getElementById("pending-changes-text");
  const btnSaveDespues = document.getElementById("btn-save-despues");

  const logoImg = new Image();
  logoImg.src = LOGO_URI;
  logoImg.onload = () => drawCover();

  function showStatus(msg, isError = false) {
    statusBanner.textContent = msg;
    statusBanner.className = `status-banner ${isError ? "err" : "ok"}`;
    statusBanner.hidden = false;
  }

  function clearStatus() {
    statusBanner.hidden = true;
  }

  function setDefaultDate() {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, "0");
    const d = String(now.getDate()).padStart(2, "0");
    walkDate.value = `${y}-${m}-${d}`;
  }

  function formatDisplayDate(isoDate) {
    if (!isoDate || !isoDate.includes("-")) return "DD / MM / AAAA";
    const [y, m, d] = isoDate.split("-");
    return `${d} / ${m} / ${y}`;
  }

  function roundRect(ctx, x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + w, y, x + w, y + h, r);
    ctx.arcTo(x + w, y + h, x, y + h, r);
    ctx.arcTo(x, y + h, x, y, r);
    ctx.arcTo(x, y, x + w, y, r);
    ctx.closePath();
  }

  function drawCover() {
    const ctx = coverCanvas.getContext("2d");
    const W = coverCanvas.width;
    const H = coverCanvas.height;

    ctx.clearRect(0, 0, W, H);
    ctx.fillStyle = "#FFFFFF";
    roundRect(ctx, 4, 4, W - 8, H - 8, 36);
    ctx.fill();
    ctx.strokeStyle = "#CBD5E1";
    ctx.lineWidth = 4;
    ctx.stroke();

    if (logoImg.complete && logoImg.naturalWidth > 0) {
      const targetH = 120;
      const targetW = (logoImg.naturalWidth / logoImg.naturalHeight) * targetH;
      ctx.drawImage(logoImg, (W - targetW) / 2, 118, targetW, targetH);
    }

    ctx.fillStyle = "#1D63ED";
    ctx.font = '700 34px "Segoe UI", Calibri, sans-serif';
    ctx.textAlign = "center";
    ctx.fillText("CAMINATA GEMBA · REGISTRO DE FECHA", W / 2, 300);

    ctx.fillStyle = "#0F172A";
    ctx.font = '700 132px "Segoe UI", Calibri, sans-serif';
    ctx.fillText(formatDisplayDate(walkDate.value), W / 2, 452);

    const pillX = 105;
    const pillY = 540;
    const pillW = W - 210;
    const pillH = 160;
    ctx.fillStyle = "#F8FAFC";
    roundRect(ctx, pillX, pillY, pillW, pillH, 28);
    ctx.fill();
    ctx.strokeStyle = "#CBD5E1";
    ctx.lineWidth = 2.5;
    ctx.stroke();

    const colW = pillW / 3;
    ctx.strokeStyle = "#CBD5E1";
    ctx.lineWidth = 2.5;
    for (let i = 1; i <= 2; i++) {
      const dx = pillX + colW * i;
      ctx.beginPath();
      ctx.moveTo(dx, pillY + 38);
      ctx.lineTo(dx, pillY + pillH - 38);
      ctx.stroke();
    }

    const cols = [
      { label: "UBICACIÓN", value: (walkSector.value || "SECTOR").trim() },
      { label: "TURNO & HORARIO", value: (walkSchedule.value || "Mañana · 08:00 - 09:00 AM").trim() },
      { label: "ENCARGADO DEL SECTOR", value: (walkOwner.value || "ENCARGADO").trim() },
    ];

    cols.forEach((col, i) => {
      const cx = pillX + colW * i + colW / 2;
      ctx.fillStyle = "#64748B";
      ctx.font = '600 24px "Segoe UI", Calibri, sans-serif';
      ctx.fillText(col.label, cx, pillY + 66);

      let fontSize = 31;
      ctx.fillStyle = "#0F172A";
      while (fontSize > 20) {
        ctx.font = `700 ${fontSize}px "Segoe UI", Calibri, sans-serif`;
        if (ctx.measureText(col.value).width <= colW - 44) break;
        fontSize -= 2;
      }
      ctx.fillText(col.value, cx, pillY + 114);
    });
  }

  async function compressPhoto(file, maxPx = 1400, quality = 0.82) {
    const bitmap = await createImageBitmap(file, { imageOrientation: "from-image" });
    let w = bitmap.width;
    let h = bitmap.height;
    if (w > maxPx || h > maxPx) {
      if (w >= h) {
        h = Math.round((h * maxPx) / w);
        w = maxPx;
      } else {
        w = Math.round((w * maxPx) / h);
        h = maxPx;
      }
    }
    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(0, 0, w, h);
    ctx.drawImage(bitmap, 0, 0, w, h);
    bitmap.close();
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/jpeg", quality));
    const bytes = new Uint8Array(await blob.arrayBuffer());
    return { bytes, width: w, height: h };
  }

  async function canvasToPngBytes(canvas) {
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
    return {
      bytes: new Uint8Array(await blob.arrayBuffer()),
      width: canvas.width,
      height: canvas.height,
    };
  }

  function clearMediaCache() {
    for (const url of state.mediaUrlCache.values()) {
      URL.revokeObjectURL(url);
    }
    state.mediaUrlCache.clear();
  }

  async function getMediaUrl(mediaPath) {
    if (!mediaPath || !state.zip) return null;
    if (state.mediaUrlCache.has(mediaPath)) return state.mediaUrlCache.get(mediaPath);
    const f = state.zip.file(mediaPath);
    if (!f) return null;
    const blob = await f.async("blob");
    const url = URL.createObjectURL(blob);
    state.mediaUrlCache.set(mediaPath, url);
    return url;
  }

  function updateSummary() {
    if (!state.info) {
      fileSummary.textContent = "Cargá tu archivo Caminata Gemba.pptx para empezar";
      btnDownloadPptx.disabled = true;
      return;
    }
    btnDownloadPptx.disabled = false;
    const dirtyTag = state.dirty ? " · (cambios sin descargar)" : "";
    fileSummary.textContent = `${state.fileName} · ${state.info.walks.length} caminatas · ${state.info.total_slides} diapositivas${dirtyTag}`;
  }

  async function loadPptxFile(file) {
    if (!file || !file.name.toLowerCase().endsWith(".pptx")) {
      showStatus("Seleccioná un archivo .pptx válido (por ejemplo Caminata Gemba.pptx).", true);
      return;
    }
    clearStatus();
    fileSummary.textContent = "Leyendo presentación…";
    try {
      clearMediaCache();
      state.despuesDrafts.clear();
      const buf = await file.arrayBuffer();
      state.zip = await JSZip.loadAsync(buf);
      state.fileName = file.name;
      state.fileSizeMB = (file.size / (1024 * 1024)).toFixed(2);
      state.dirty = false;
      state.info = await GembaPPTX.inspectZip(state.zip);

      pptxDropCard.hidden = true;
      appWorkspace.hidden = false;
      updateSummary();
      populateWalkSelector();
      if (!panelDespues.hidden) await renderDespuesPanel();
      showStatus(
        `Archivo cargado: ${state.fileName} (${state.info.walks.length} caminatas · ${state.info.total_slides} diapositivas · ${state.fileSizeMB} MB).`
      );
    } catch (err) {
      showStatus("No se pudo abrir el archivo .pptx: " + err.message, true);
    }
  }

  async function downloadUpdatedPptx() {
    if (!state.zip) return;
    btnDownloadPptx.disabled = true;
    const prevText = btnDownloadPptx.textContent;
    btnDownloadPptx.textContent = "Generando .pptx…";
    try {
      const blob = await state.zip.generateAsync({
        type: "blob",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = state.fileName || "Caminata Gemba.pptx";
      document.body.append(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
      state.dirty = false;
      state.fileSizeMB = (blob.size / (1024 * 1024)).toFixed(2);
      updateSummary();
      showStatus(`Descargado "${state.fileName}" (${state.fileSizeMB} MB). Listo para subir/reemplazar en Google Drive.`);
    } catch (err) {
      showStatus("Error al generar el .pptx: " + err.message, true);
    } finally {
      btnDownloadPptx.disabled = false;
      btnDownloadPptx.textContent = prevText;
    }
  }

  // PPTX file loading events
  pptxInput.addEventListener("change", (e) => {
    loadPptxFile(e.target.files[0]);
    pptxInput.value = "";
  });

  pptxDropCard.addEventListener("dragover", (e) => {
    e.preventDefault();
    pptxDropCard.classList.add("dragover");
  });
  pptxDropCard.addEventListener("dragleave", () => pptxDropCard.classList.remove("dragover"));
  pptxDropCard.addEventListener("drop", (e) => {
    e.preventDefault();
    pptxDropCard.classList.remove("dragover");
    loadPptxFile(e.dataTransfer.files[0]);
  });

  btnDownloadPptx.addEventListener("click", downloadUpdatedPptx);

  // Tabs
  function switchTab(mode) {
    const isAntes = mode === "antes";
    tabBtnAntes.classList.toggle("active", isAntes);
    tabBtnAntes.setAttribute("aria-expanded", String(isAntes));
    tabBtnDespues.classList.toggle("active", !isAntes);
    tabBtnDespues.setAttribute("aria-expanded", String(!isAntes));
    panelAntes.hidden = !isAntes;
    panelDespues.hidden = isAntes;
    if (!isAntes) renderDespuesPanel();
  }

  tabBtnAntes.addEventListener("click", () => switchTab("antes"));
  tabBtnDespues.addEventListener("click", () => switchTab("despues"));

  // Form events
  walkDate.addEventListener("input", drawCover);
  walkSchedule.addEventListener("input", drawCover);
  walkOwner.addEventListener("input", drawCover);
  walkSector.addEventListener("input", () => {
    const matched = SECTOR_OWNERS[walkSector.value.trim()];
    if (matched) walkOwner.value = matched;
    drawCover();
  });

  // ANTES photos
  function addAntesFiles(fileList) {
    for (const file of fileList) {
      if (!file.type.startsWith("image/")) continue;
      state.antesFiles.push({
        id: Math.random().toString(36).slice(2),
        file,
        url: URL.createObjectURL(file),
      });
    }
    renderAntesGrid();
  }

  function renderAntesGrid() {
    antesCountBadge.textContent = `${state.antesFiles.length} ${state.antesFiles.length === 1 ? "foto" : "fotos"}`;
    antesGrid.innerHTML = "";
    state.antesFiles.forEach((item, idx) => {
      const card = document.createElement("div");
      card.className = "photo-card";

      const img = document.createElement("img");
      img.src = item.url;
      img.alt = `Foto ${idx + 1}`;

      const footer = document.createElement("div");
      footer.className = "photo-card-footer";
      const label = document.createElement("span");
      label.textContent = `#${idx + 1}`;

      const actions = document.createElement("div");
      actions.className = "photo-card-actions";

      const btnLeft = document.createElement("button");
      btnLeft.type = "button";
      btnLeft.className = "icon-btn";
      btnLeft.title = "Mover a la izquierda";
      btnLeft.textContent = "←";
      btnLeft.disabled = idx === 0;
      btnLeft.addEventListener("click", () => {
        [state.antesFiles[idx - 1], state.antesFiles[idx]] = [state.antesFiles[idx], state.antesFiles[idx - 1]];
        renderAntesGrid();
      });

      const btnRight = document.createElement("button");
      btnRight.type = "button";
      btnRight.className = "icon-btn";
      btnRight.title = "Mover a la derecha";
      btnRight.textContent = "→";
      btnRight.disabled = idx === state.antesFiles.length - 1;
      btnRight.addEventListener("click", () => {
        [state.antesFiles[idx + 1], state.antesFiles[idx]] = [state.antesFiles[idx], state.antesFiles[idx + 1]];
        renderAntesGrid();
      });

      const btnDel = document.createElement("button");
      btnDel.type = "button";
      btnDel.className = "icon-btn danger";
      btnDel.title = "Quitar foto";
      btnDel.textContent = "×";
      btnDel.addEventListener("click", () => {
        URL.revokeObjectURL(item.url);
        state.antesFiles.splice(idx, 1);
        renderAntesGrid();
      });

      actions.append(btnLeft, btnRight, btnDel);
      footer.append(label, actions);
      card.append(img, footer);
      antesGrid.append(card);
    });
  }

  antesFilesInput.addEventListener("change", (e) => {
    addAntesFiles(e.target.files);
    antesFilesInput.value = "";
  });
  antesDropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    antesDropzone.classList.add("dragover");
  });
  antesDropzone.addEventListener("dragleave", () => antesDropzone.classList.remove("dragover"));
  antesDropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    antesDropzone.classList.remove("dragover");
    addAntesFiles(e.dataTransfer.files);
  });

  btnClearAntes.addEventListener("click", () => {
    state.antesFiles.forEach((f) => URL.revokeObjectURL(f.url));
    state.antesFiles = [];
    renderAntesGrid();
  });

  walkForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    clearStatus();

    if (!state.zip) {
      showStatus("Primero cargá tu archivo Caminata Gemba.pptx.", true);
      return;
    }
    if (!walkForm.checkValidity()) {
      const firstInvalid = walkForm.querySelector(":invalid");
      if (firstInvalid) firstInvalid.focus();
      showStatus("Completá todos los campos obligatorios de la portada.", true);
      return;
    }
    if (state.antesFiles.length === 0) {
      showStatus("Agregá al menos una foto del ANTES.", true);
      return;
    }

    btnSubmitWalk.disabled = true;
    btnSubmitWalk.textContent = "Procesando fotos y generando diapositivas…";

    try {
      drawCover();
      const coverData = await canvasToPngBytes(coverCanvas);
      const compressedPhotos = [];
      for (const item of state.antesFiles) {
        compressedPhotos.push(await compressPhoto(item.file));
      }

      const res = await GembaPPTX.appendWalk(state.zip, coverData, compressedPhotos, true);
      state.antesFiles.forEach((f) => URL.revokeObjectURL(f.url));
      state.antesFiles = [];
      renderAntesGrid();

      clearMediaCache();
      state.info = await GembaPPTX.inspectZip(state.zip);
      state.dirty = true;
      updateSummary();
      populateWalkSelector();

      showStatus(
        `¡Listo! Se agregó la portada + ${res.added_photo_slides} diapositivas (total ${res.total_slides}). Hacé clic en "Descargar .pptx" para guardar el archivo actualizado.`
      );
    } catch (err) {
      showStatus("Error al agregar la caminata: " + err.message, true);
    } finally {
      btnSubmitWalk.disabled = false;
      btnSubmitWalk.textContent = "Agregar caminata al PowerPoint";
    }
  });

  // DESPUÉS panel
  function populateWalkSelector() {
    walkSelector.innerHTML = "";
    const walks = state.info?.walks || [];
    if (walks.length === 0) {
      const opt = document.createElement("option");
      opt.value = "";
      opt.textContent = "No se encontraron caminatas";
      walkSelector.append(opt);
      return;
    }
    [...walks].reverse().forEach((w) => {
      const pending = w.slides.filter((s) => !s.has_despues).length;
      const opt = document.createElement("option");
      opt.value = String(w.walk_index);
      opt.textContent = `Caminata #${w.walk_index} (diap. ${w.cover_slide_number || "?"}) — ${pending} pendientes de ${w.slides.length}`;
      walkSelector.append(opt);
    });
  }

  function updateSaveBar() {
    let count = 0;
    for (const draft of state.despuesDrafts.values()) {
      if (draft.file || (draft.note && draft.note.trim())) count++;
    }
    despuesSaveBar.hidden = count === 0;
    pendingChangesText.textContent = `${count} ${count === 1 ? "solución lista" : "soluciones listas"} para aplicar`;
  }

  async function renderDespuesPanel() {
    const walks = state.info?.walks || [];
    const selectedIdx = Number(walkSelector.value);
    const walk = walks.find((w) => w.walk_index === selectedIdx);

    if (!walk) {
      walkCoverStrip.hidden = true;
      despuesList.innerHTML = '<div class="empty-state">No hay caminatas cargadas todavía.</div>';
      updateSaveBar();
      return;
    }

    const total = walk.slides.length;
    const done = walk.slides.filter((s) => s.has_despues).length;
    const pending = total - done;

    const coverUrl = await getMediaUrl(walk.cover_media_path);
    if (coverUrl) {
      walkCoverImg.src = coverUrl;
      walkCoverStrip.hidden = false;
    } else {
      walkCoverStrip.hidden = true;
    }
    walkStats.innerHTML = `<strong>${done} de ${total} resueltas</strong> · ${pending} pendientes`;

    const onlyPending = filterPending.checked;
    const slidesToShow = walk.slides.filter(
      (s) => !onlyPending || !s.has_despues || state.despuesDrafts.has(s.slide_index)
    );

    despuesList.innerHTML = "";
    if (slidesToShow.length === 0) {
      despuesList.innerHTML =
        '<div class="empty-state">Todas las diapositivas de esta caminata ya tienen su DESPUÉS cargado. Destildá "Mostrar solo pendientes" si querés reemplazar alguna.</div>';
      updateSaveBar();
      return;
    }

    for (const s of slidesToShow) {
      const draft = state.despuesDrafts.get(s.slide_index) || {};
      const card = document.createElement("div");
      card.className = `slide-pair-card ${draft.file || draft.note ? "modified" : ""}`;

      // Left (ANTES)
      const colLeft = document.createElement("div");
      colLeft.className = "pair-col";
      colLeft.innerHTML = `<h3><span>Diapositiva #${s.slide_number} · ANTES</span></h3>`;
      const antesBox = document.createElement("div");
      antesBox.className = "pair-img-box";
      const antesUrl = await getMediaUrl(s.antes_media_path);
      if (antesUrl) {
        const img = document.createElement("img");
        img.src = antesUrl;
        img.alt = `Antes diapositiva ${s.slide_number}`;
        antesBox.append(img);
      } else {
        antesBox.textContent = "Sin foto ANTES";
      }
      colLeft.append(antesBox);

      // Right (DESPUÉS)
      const colRight = document.createElement("div");
      colRight.className = "pair-col";
      const h3 = document.createElement("h3");
      const titleSpan = document.createElement("span");
      titleSpan.textContent = "DESPUÉS";
      const pill = document.createElement("span");
      pill.className = s.has_despues ? "pill-ok" : "pill-pending";
      pill.textContent = draft.file || draft.note ? "Cambio pendiente" : s.has_despues ? "Cargado" : "Pendiente";
      h3.append(titleSpan, pill);

      const dropLabel = document.createElement("label");
      dropLabel.className = "despues-drop";
      const fileInput = document.createElement("input");
      fileInput.type = "file";
      fileInput.accept = "image/*";
      fileInput.className = "visually-hidden";

      function setDraftFile(file) {
        if (!file || !file.type.startsWith("image/")) return;
        const current = state.despuesDrafts.get(s.slide_index) || {
          slide_path: s.slide_path,
          rels_path: s.rels_path,
          note: s.notes || "",
        };
        if (current.previewUrl) URL.revokeObjectURL(current.previewUrl);
        current.file = file;
        current.previewUrl = URL.createObjectURL(file);
        state.despuesDrafts.set(s.slide_index, current);
        renderDespuesPanel();
      }

      fileInput.addEventListener("change", (ev) => setDraftFile(ev.target.files[0]));
      dropLabel.addEventListener("dragover", (ev) => {
        ev.preventDefault();
        dropLabel.classList.add("dragover");
      });
      dropLabel.addEventListener("dragleave", () => dropLabel.classList.remove("dragover"));
      dropLabel.addEventListener("drop", (ev) => {
        ev.preventDefault();
        dropLabel.classList.remove("dragover");
        setDraftFile(ev.dataTransfer.files[0]);
      });

      const existingDespuesUrl = await getMediaUrl(s.despues_media_path);
      if (draft.previewUrl) {
        const prev = document.createElement("img");
        prev.src = draft.previewUrl;
        prev.alt = "Vista previa DESPUÉS";
        dropLabel.append(fileInput, prev);
      } else if (existingDespuesUrl) {
        const ex = document.createElement("img");
        ex.src = existingDespuesUrl;
        ex.alt = `Después diapositiva ${s.slide_number}`;
        dropLabel.append(fileInput, ex);
      } else {
        const ph = document.createElement("div");
        ph.innerHTML = "<strong>Arrastrá la foto del DESPUÉS acá</strong><br><span class='hint'>o hacé clic para buscarla</span>";
        dropLabel.append(fileInput, ph);
      }

      const noteInput = document.createElement("input");
      noteInput.type = "text";
      noteInput.className = "note-input";
      noteInput.placeholder = "Nota opcional (ej. Se acomodó el cartel)";
      noteInput.value = draft.note !== undefined ? draft.note : s.notes || "";
      noteInput.addEventListener("input", () => {
        const current = state.despuesDrafts.get(s.slide_index) || {
          slide_path: s.slide_path,
          rels_path: s.rels_path,
        };
        current.note = noteInput.value;
        state.despuesDrafts.set(s.slide_index, current);
        card.classList.add("modified");
        updateSaveBar();
      });

      colRight.append(h3, dropLabel, noteInput);
      card.append(colLeft, colRight);
      despuesList.append(card);
    }

    updateSaveBar();
  }

  walkSelector.addEventListener("change", () => renderDespuesPanel());
  filterPending.addEventListener("change", () => renderDespuesPanel());

  btnSaveDespues.addEventListener("click", async () => {
    clearStatus();
    btnSaveDespues.disabled = true;
    btnSaveDespues.textContent = "Aplicando soluciones…";
    try {
      const updates = [];
      for (const draft of state.despuesDrafts.values()) {
        if (!draft.file && !(draft.note && draft.note.trim())) continue;
        const photo = draft.file ? await compressPhoto(draft.file) : null;
        updates.push({
          slide_path: draft.slide_path,
          rels_path: draft.rels_path,
          photo,
          note: draft.note || "",
        });
      }
      const res = await GembaPPTX.updateDespuesSlides(state.zip, updates);
      for (const draft of state.despuesDrafts.values()) {
        if (draft.previewUrl) URL.revokeObjectURL(draft.previewUrl);
      }
      state.despuesDrafts.clear();
      clearMediaCache();
      const prevWalk = walkSelector.value;
      state.info = await GembaPPTX.inspectZip(state.zip);
      state.dirty = true;
      updateSummary();
      populateWalkSelector();
      if (prevWalk && [...walkSelector.options].some((o) => o.value === prevWalk)) {
        walkSelector.value = prevWalk;
      }
      await renderDespuesPanel();
      showStatus(
        `¡Listo! Se aplicaron ${res.updated_slides} soluciones. Hacé clic en "Descargar .pptx" arriba a la derecha para guardar tu archivo actualizado.`
      );
    } catch (err) {
      showStatus("Error al aplicar soluciones: " + err.message, true);
    } finally {
      btnSaveDespues.disabled = false;
      btnSaveDespues.textContent = "Aplicar soluciones al PowerPoint";
    }
  });

  setDefaultDate();
  drawCover();
  updateSummary();
})();
"""

EXTRA_CSS = """
.start-drop-card {
  border: 2.5px dashed var(--blue);
  background: #f0f9ff;
  border-radius: 16px;
  padding: 3rem 1.5rem;
  text-align: center;
  margin-top: 1.5rem;
}
.start-drop-card.dragover {
  background: #bae6fd;
  border-color: var(--blue-dark);
}
.start-drop-card h2 {
  margin-bottom: 0.5rem;
  font-size: 1.35rem;
}
"""

HTML = f"""<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Caminata Gemba — Automatizador Offline</title>
  <style>
{STYLES}
{EXTRA_CSS}
  </style>
</head>
<body>
  <header class="topbar">
    <div class="topbar-inner">
      <div class="brand">
        <img src="{LOGO_URI}" alt="Logo empresa" class="brand-logo" width="44" height="36">
        <div>
          <h1>Caminata Gemba</h1>
          <p id="file-summary" class="file-summary">Cargá tu archivo Caminata Gemba.pptx para empezar</p>
        </div>
      </div>

      <div class="topbar-actions">
        <label for="pptx-upload-input" class="btn btn-secondary">
          Abrir .pptx
        </label>
        <input id="pptx-upload-input" type="file" accept=".pptx" class="visually-hidden">
        <button id="btn-download-pptx" type="button" class="btn btn-primary" disabled>Descargar .pptx</button>
      </div>
    </div>
  </header>

  <main class="container">
    <div id="status-banner" class="status-banner" role="status" aria-live="polite" hidden></div>

    <section id="pptx-drop-card" class="start-drop-card">
      <h2>Paso 1: Abrí tu archivo Caminata Gemba.pptx</h2>
      <p class="hint">Arrastrá el archivo <strong>Caminata Gemba.pptx</strong> descargado de Drive hasta acá, o hacé clic en el botón.</p>
      <p class="hint">Funciona 100% en tu navegador, sin instalar programas y sin subir tus fotos a ningún servidor.</p>
      <label for="pptx-upload-input" class="btn btn-primary btn-lg">Elegir Caminata Gemba.pptx</label>
    </section>

    <div id="app-workspace" hidden>
      <nav class="tabs" aria-label="Modos de trabajo">
        <button id="tab-btn-antes" type="button" class="tab-btn active" aria-controls="panel-antes" aria-expanded="true">
          1. Nueva Caminata (Miércoles · ANTES)
        </button>
        <button id="tab-btn-despues" type="button" class="tab-btn" aria-controls="panel-despues" aria-expanded="false">
          2. Cargar Soluciones (DESPUÉS)
        </button>
      </nav>

      <!-- PANEL 1: NUEVA CAMINATA (ANTES) -->
      <section id="panel-antes" class="panel">
        <form id="walk-form" action="#" method="POST" aria-labelledby="heading-antes" novalidate>
          <div class="split-grid">
            <div class="card">
              <h2 id="heading-antes">Datos de la Portada</h2>

              <div class="form-group">
                <label for="walk-date">Fecha de la caminata <span aria-hidden="true">*</span></label>
                <input id="walk-date" name="date" type="date" required>
              </div>

              <div class="form-group">
                <label for="walk-sector">Ubicación / Sector <span aria-hidden="true">*</span></label>
                <input id="walk-sector" name="sector" type="text" list="sectores-list" required placeholder="Ej. MAQUINADO" enterkeyhint="next">
                <datalist id="sectores-list">
                  <option value="Embalaje - Pulido"></option>
                  <option value="LIVING"></option>
                  <option value="MAQUINADO"></option>
                  <option value="ARMADO"></option>
                  <option value="MAQUINADO - ARMADO"></option>
                  <option value="MAQUINADO – ESQUELETOS"></option>
                </datalist>
              </div>

              <div class="form-group">
                <label for="walk-schedule">Turno &amp; Horario <span aria-hidden="true">*</span></label>
                <input id="walk-schedule" name="schedule" type="text" list="horarios-list" value="Mañana · 08:00 - 09:00 AM" required enterkeyhint="next">
                <datalist id="horarios-list">
                  <option value="Mañana · 08:00 - 09:00 AM"></option>
                  <option value="Mañana · 10:35 - 11:35 AM"></option>
                  <option value="Tarde · 14:00 - 15:00 PM"></option>
                </datalist>
              </div>

              <div class="form-group">
                <label for="walk-owner">Encargado del sector <span aria-hidden="true">*</span></label>
                <input id="walk-owner" name="owner" type="text" list="encargados-list" required placeholder="Ej. ARIEL DEHESA" enterkeyhint="done">
                <datalist id="encargados-list">
                  <option value="Tino Varela"></option>
                  <option value="ARIEL CARANDINO"></option>
                  <option value="ARIEL DEHESA"></option>
                  <option value="LUCAS LUQUE"></option>
                  <option value="ARIEL DEHESA – LUCAS LUQUE"></option>
                  <option value="ARIEL DEHESA – ARIEL CARANDINO"></option>
                </datalist>
              </div>
            </div>

            <div class="card preview-card">
              <h2>Vista previa de portada</h2>
              <p class="hint">Así se va a insertar la diapositiva separadora en el PowerPoint.</p>
              <canvas id="cover-canvas" width="1800" height="809" class="cover-preview"></canvas>
            </div>
          </div>

          <div class="card upload-section">
            <div class="section-header">
              <div>
                <h2>Fotos del ANTES</h2>
                <p class="hint">Arrastrá todas las fotos que bajaste de WhatsApp. Se comprimen y rotan solas.</p>
              </div>
              <span id="antes-count-badge" class="badge">0 fotos</span>
            </div>

            <div id="antes-dropzone" class="dropzone">
              <input id="antes-files-input" name="photos" type="file" accept="image/*" multiple class="visually-hidden">
              <label for="antes-files-input" class="dropzone-label">
                <strong>Hacé clic para elegir fotos</strong> o arrastralas y soltalas acá
                <span class="hint">Podés seleccionar 10, 20 o más fotos de una sola vez</span>
              </label>
            </div>

            <div id="antes-grid" class="photo-grid"></div>

            <div class="form-actions">
              <button id="btn-clear-antes" type="button" class="btn btn-secondary">Limpiar fotos</button>
              <button id="btn-submit-walk" type="submit" class="btn btn-primary btn-lg">
                Agregar caminata al PowerPoint
              </button>
            </div>
          </div>
        </form>
      </section>

      <!-- PANEL 2: CARGAR SOLUCIONES (DESPUÉS) -->
      <section id="panel-despues" class="panel" hidden>
        <div class="card">
          <div class="despues-toolbar">
            <div class="form-group inline-group">
              <label for="walk-selector">Caminata</label>
              <select id="walk-selector" name="walk_selector"></select>
            </div>

            <div class="checkbox-inline">
              <input id="filter-pending" type="checkbox" checked>
              <label for="filter-pending">Mostrar solo pendientes (sin DESPUÉS)</label>
            </div>
          </div>

          <div id="walk-cover-strip" class="walk-cover-strip" hidden>
            <img id="walk-cover-img" alt="Portada de la caminata seleccionada">
            <div id="walk-stats" class="walk-stats"></div>
          </div>
        </div>

        <div id="despues-list" class="despues-list"></div>

        <div id="despues-save-bar" class="save-bar" hidden>
          <span id="pending-changes-text">0 soluciones listas para aplicar</span>
          <button id="btn-save-despues" type="button" class="btn btn-primary btn-lg">
            Aplicar soluciones al PowerPoint
          </button>
        </div>
      </section>
    </div>
  </main>

  <script>window.__GEMBA_LOGO__ = "{LOGO_URI}";</script>
  <script>{JSZIP}</script>
  <script>{ENGINE}</script>
  <script>{APP_JS}</script>
</body>
</html>
"""

out_path = BASE / "Caminata_Gemba_App.html"
out_path.write_text(HTML, encoding="utf-8")
print(f"Built {out_path.name}: {out_path.stat().st_size / 1024:.1f} KB")
