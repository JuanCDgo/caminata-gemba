(() => {
  const SECTOR_OWNERS = {
    "Embalaje - Pulido": "Tino Varela",
    "LIVING": "ARIEL CARANDINO",
    "MAQUINADO": "ARIEL DEHESA",
    "ARMADO": "LUCAS LUQUE",
    "MAQUINADO - ARMADO": "ARIEL DEHESA – LUCAS LUQUE",
    "MAQUINADO – ESQUELETOS": "ARIEL DEHESA – ARIEL CARANDINO",
  };

  let appState = {
    status: null,
    antesFiles: [], // { id, file, url }
    despuesDrafts: new Map(), // slideIndex -> { file?, previewUrl?, note }
  };

  // DOM refs
  const fileSummary = document.getElementById("file-summary");
  const statusBanner = document.getElementById("status-banner");
  const pptxUploadInput = document.getElementById("pptx-upload-input");
  const btnOpenFolder = document.getElementById("btn-open-folder");

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
  const btnRefreshWalks = document.getElementById("btn-refresh-walks");
  const walkCoverStrip = document.getElementById("walk-cover-strip");
  const walkCoverImg = document.getElementById("walk-cover-img");
  const walkStats = document.getElementById("walk-stats");
  const despuesList = document.getElementById("despues-list");
  const despuesSaveBar = document.getElementById("despues-save-bar");
  const pendingChangesText = document.getElementById("pending-changes-text");
  const btnSaveDespues = document.getElementById("btn-save-despues");

  // Load company logo for cover canvas
  const logoImg = new Image();
  logoImg.src = "/static/logo.png";
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

    // Logo
    if (logoImg.complete && logoImg.naturalWidth > 0) {
      const targetH = 120;
      const targetW = (logoImg.naturalWidth / logoImg.naturalHeight) * targetH;
      ctx.drawImage(logoImg, (W - targetW) / 2, 118, targetW, targetH);
    }

    // Subtitle
    ctx.fillStyle = "#1D63ED";
    ctx.font = '700 34px "Segoe UI", Calibri, sans-serif';
    ctx.textAlign = "center";
    ctx.fillText("CAMINATA GEMBA · REGISTRO DE FECHA", W / 2, 300);

    // Date
    ctx.fillStyle = "#0F172A";
    ctx.font = '700 132px "Segoe UI", Calibri, sans-serif';
    ctx.fillText(formatDisplayDate(walkDate.value), W / 2, 452);

    // Bottom info pill
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

    // Dividers
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

      // Fit value font size if long
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

  // Tabs
  function switchTab(mode) {
    const isAntes = mode === "antes";
    tabBtnAntes.classList.toggle("active", isAntes);
    tabBtnAntes.setAttribute("aria-expanded", String(isAntes));
    tabBtnDespues.classList.toggle("active", !isAntes);
    tabBtnDespues.setAttribute("aria-expanded", String(!isAntes));
    panelAntes.hidden = !isAntes;
    panelDespues.hidden = isAntes;
    if (!isAntes) {
      renderDespuesPanel();
    }
  }

  tabBtnAntes.addEventListener("click", () => switchTab("antes"));
  tabBtnDespues.addEventListener("click", () => switchTab("despues"));

  // Form events
  walkDate.addEventListener("input", drawCover);
  walkSchedule.addEventListener("input", drawCover);
  walkOwner.addEventListener("input", drawCover);
  walkSector.addEventListener("input", () => {
    const matched = SECTOR_OWNERS[walkSector.value.trim()];
    if (matched) {
      walkOwner.value = matched;
    }
    drawCover();
  });

  // ANTES photo management
  function addAntesFiles(fileList) {
    for (const file of fileList) {
      if (!file.type.startsWith("image/")) continue;
      appState.antesFiles.push({
        id: crypto.randomUUID(),
        file,
        url: URL.createObjectURL(file),
      });
    }
    renderAntesGrid();
  }

  function renderAntesGrid() {
    antesCountBadge.textContent = `${appState.antesFiles.length} ${appState.antesFiles.length === 1 ? "foto" : "fotos"}`;
    antesGrid.innerHTML = "";
    appState.antesFiles.forEach((item, idx) => {
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
        [appState.antesFiles[idx - 1], appState.antesFiles[idx]] = [
          appState.antesFiles[idx],
          appState.antesFiles[idx - 1],
        ];
        renderAntesGrid();
      });

      const btnRight = document.createElement("button");
      btnRight.type = "button";
      btnRight.className = "icon-btn";
      btnRight.title = "Mover a la derecha";
      btnRight.textContent = "→";
      btnRight.disabled = idx === appState.antesFiles.length - 1;
      btnRight.addEventListener("click", () => {
        [appState.antesFiles[idx + 1], appState.antesFiles[idx]] = [
          appState.antesFiles[idx],
          appState.antesFiles[idx + 1],
        ];
        renderAntesGrid();
      });

      const btnDel = document.createElement("button");
      btnDel.type = "button";
      btnDel.className = "icon-btn danger";
      btnDel.title = "Quitar foto";
      btnDel.textContent = "×";
      btnDel.addEventListener("click", () => {
        URL.revokeObjectURL(item.url);
        appState.antesFiles.splice(idx, 1);
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
    appState.antesFiles.forEach((f) => URL.revokeObjectURL(f.url));
    appState.antesFiles = [];
    renderAntesGrid();
  });

  // Submit new walk
  walkForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    clearStatus();

    if (!walkForm.checkValidity()) {
      const firstInvalid = walkForm.querySelector(":invalid");
      if (firstInvalid) firstInvalid.focus();
      showStatus("Completá todos los campos obligatorios de la portada.", true);
      return;
    }

    if (appState.antesFiles.length === 0) {
      showStatus("Agregá al menos una foto del ANTES antes de generar.", true);
      return;
    }

    btnSubmitWalk.disabled = true;
    btnSubmitWalk.textContent = "Generando diapositivas…";

    try {
      drawCover();
      const coverDataUrl = coverCanvas.toDataURL("image/png");
      const formData = new FormData();
      formData.append("cover_png", coverDataUrl);
      for (const item of appState.antesFiles) {
        formData.append("photos", item.file, item.file.name);
      }

      const res = await fetch("/api/add-walk", { method: "POST", body: formData });
      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || "Error al guardar la caminata.");
      }

      appState.antesFiles.forEach((f) => URL.revokeObjectURL(f.url));
      appState.antesFiles = [];
      renderAntesGrid();

      await fetchStatus();
      showStatus(
        `Listo: se agregó la portada + ${data.added_photo_slides} diapositivas al archivo (total ${data.total_slides} diapositivas · ${data.file_size_mb} MB).`
      );
    } catch (err) {
      showStatus(err.message, true);
    } finally {
      btnSubmitWalk.disabled = false;
      btnSubmitWalk.textContent = "Agregar caminata al PowerPoint";
    }
  });

  // Upload another PPTX
  pptxUploadInput.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    clearStatus();
    const formData = new FormData();
    formData.append("pptx", file);
    try {
      const res = await fetch("/api/upload-pptx", { method: "POST", body: formData });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "No se pudo cargar el archivo.");
      appState.status = data;
      appState.despuesDrafts.clear();
      updateFileSummary();
      populateWalkSelector();
      renderDespuesPanel();
      showStatus(`Archivo cargado correctamente (${data.total_slides} diapositivas · ${data.file_size_mb} MB).`);
    } catch (err) {
      showStatus(err.message, true);
    } finally {
      pptxUploadInput.value = "";
    }
  });

  btnOpenFolder.addEventListener("click", () => {
    fetch("/api/open-folder", { method: "POST" });
  });

  // DESPUÉS panel
  function populateWalkSelector() {
    walkSelector.innerHTML = "";
    const walks = appState.status?.walks || [];
    if (walks.length === 0) {
      const opt = document.createElement("option");
      opt.value = "";
      opt.textContent = "No se encontraron caminatas";
      walkSelector.append(opt);
      return;
    }
    // Newest first
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
    for (const draft of appState.despuesDrafts.values()) {
      if (draft.file || (draft.note && draft.note.trim())) count++;
    }
    despuesSaveBar.hidden = count === 0;
    pendingChangesText.textContent = `${count} ${count === 1 ? "solución lista" : "soluciones listas"} para guardar`;
  }

  function renderDespuesPanel() {
    const walks = appState.status?.walks || [];
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

    if (walk.cover_slide_index !== null) {
      walkCoverImg.src = `/api/slide-image/${walk.cover_slide_index}/0?t=${Date.now()}`;
      walkCoverStrip.hidden = false;
    } else {
      walkCoverStrip.hidden = true;
    }
    walkStats.innerHTML = `<strong>${done} de ${total} resueltas</strong> · ${pending} pendientes`;

    const onlyPending = filterPending.checked;
    const slidesToShow = walk.slides.filter((s) => {
      if (!onlyPending) return true;
      return !s.has_despues || appState.despuesDrafts.has(s.slide_index);
    });

    despuesList.innerHTML = "";
    if (slidesToShow.length === 0) {
      despuesList.innerHTML =
        '<div class="empty-state">Todas las diapositivas de esta caminata ya tienen su DESPUÉS cargado. Destildá "Mostrar solo pendientes" si querés reemplazar alguna.</div>';
      updateSaveBar();
      return;
    }

    const cacheBust = Date.now();
    slidesToShow.forEach((s) => {
      const draft = appState.despuesDrafts.get(s.slide_index) || {};
      const card = document.createElement("div");
      card.className = `slide-pair-card ${draft.file || draft.note ? "modified" : ""}`;

      // Left column (ANTES)
      const colLeft = document.createElement("div");
      colLeft.className = "pair-col";
      colLeft.innerHTML = `<h3><span>Diapositiva #${s.slide_number} · ANTES</span></h3>`;
      const antesBox = document.createElement("div");
      antesBox.className = "pair-img-box";
      if (s.antes_pic_index !== null) {
        const img = document.createElement("img");
        img.src = `/api/slide-image/${s.slide_index}/${s.antes_pic_index}?t=${cacheBust}`;
        img.alt = `Antes diapositiva ${s.slide_number}`;
        antesBox.append(img);
      } else {
        antesBox.textContent = "Sin foto ANTES";
      }
      colLeft.append(antesBox);

      // Right column (DESPUÉS)
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
        const current = appState.despuesDrafts.get(s.slide_index) || { note: s.notes || "" };
        if (current.previewUrl) URL.revokeObjectURL(current.previewUrl);
        current.file = file;
        current.previewUrl = URL.createObjectURL(file);
        appState.despuesDrafts.set(s.slide_index, current);
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

      if (draft.previewUrl) {
        const prev = document.createElement("img");
        prev.src = draft.previewUrl;
        prev.alt = "Vista previa DESPUÉS";
        dropLabel.append(fileInput, prev);
      } else if (s.despues_pic_index !== null) {
        const existingImg = document.createElement("img");
        existingImg.src = `/api/slide-image/${s.slide_index}/${s.despues_pic_index}?t=${cacheBust}`;
        existingImg.alt = `Después diapositiva ${s.slide_number}`;
        dropLabel.append(fileInput, existingImg);
      } else {
        const placeholder = document.createElement("div");
        placeholder.innerHTML = "<strong>Arrastrá la foto del DESPUÉS acá</strong><br><span class='hint'>o hacé clic para buscarla</span>";
        dropLabel.append(fileInput, placeholder);
      }

      const noteInput = document.createElement("input");
      noteInput.type = "text";
      noteInput.className = "note-input";
      noteInput.placeholder = "Nota opcional (ej. Se acomodó el cartel)";
      noteInput.value = draft.note !== undefined ? draft.note : s.notes || "";
      noteInput.addEventListener("input", () => {
        const current = appState.despuesDrafts.get(s.slide_index) || {};
        current.note = noteInput.value;
        appState.despuesDrafts.set(s.slide_index, current);
        card.classList.add("modified");
        updateSaveBar();
      });

      colRight.append(h3, dropLabel, noteInput);
      card.append(colLeft, colRight);
      despuesList.append(card);
    });

    updateSaveBar();
  }

  walkSelector.addEventListener("change", renderDespuesPanel);
  filterPending.addEventListener("change", renderDespuesPanel);
  btnRefreshWalks.addEventListener("click", async () => {
    await fetchStatus();
    renderDespuesPanel();
  });

  btnSaveDespues.addEventListener("click", async () => {
    clearStatus();
    btnSaveDespues.disabled = true;
    btnSaveDespues.textContent = "Guardando en PowerPoint…";

    try {
      const formData = new FormData();
      const meta = [];
      for (const [slideIndex, draft] of appState.despuesDrafts.entries()) {
        if (!draft.file && !(draft.note && draft.note.trim())) continue;
        meta.push({ slide_index: slideIndex, note: draft.note || "" });
        if (draft.file) {
          formData.append(`photo_${slideIndex}`, draft.file, draft.file.name);
        }
      }
      formData.append("updates_meta", JSON.stringify(meta));

      const res = await fetch("/api/update-despues", { method: "POST", body: formData });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Error al guardar las soluciones.");

      for (const draft of appState.despuesDrafts.values()) {
        if (draft.previewUrl) URL.revokeObjectURL(draft.previewUrl);
      }
      appState.despuesDrafts.clear();

      await fetchStatus();
      renderDespuesPanel();
      showStatus(`Se guardaron ${data.updated_slides} soluciones en Caminata Gemba.pptx (${data.file_size_mb} MB).`);
    } catch (err) {
      showStatus(err.message, true);
    } finally {
      btnSaveDespues.disabled = false;
      btnSaveDespues.textContent = "Guardar soluciones en PowerPoint";
    }
  });

  function updateFileSummary() {
    if (!appState.status || !appState.status.loaded) {
      fileSummary.textContent = "Sin archivo cargado — subí tu Caminata Gemba.pptx";
      return;
    }
    const { filename, total_slides, file_size_mb, walks } = appState.status;
    fileSummary.textContent = `${filename} · ${walks.length} caminatas · ${total_slides} diapositivas · ${file_size_mb} MB`;
  }

  async function fetchStatus() {
    const res = await fetch("/api/status");
    const data = await res.json();
    appState.status = data;
    updateFileSummary();
    const prevWalk = walkSelector.value;
    populateWalkSelector();
    if (prevWalk && [...walkSelector.options].some((o) => o.value === prevWalk)) {
      walkSelector.value = prevWalk;
    }
  }

  // Init
  setDefaultDate();
  drawCover();
  fetchStatus();
})();
