let state = {
  fields: [],
  rows: [],
  excel_mapping: {},
  email_subject_template: "Carta para {empresa}",
  email_body_template: "<p>Estimados,</p><p>Adjunto la carta de presentación correspondiente.</p><p>Atentamente,</p>"
};

// --- Referencias API de pywebview ---
function api() {
  return window.pywebview.api;
}

// --- Notificaciones Toast ---
function toast(msg, type = "info") {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.className = "show " + type;
  setTimeout(() => { el.className = ""; }, 4500);
}

// --- Consola de Log ---
function logToConsole(message, type = "info") {
  const consoleEl = document.getElementById("consoleLog");
  const timestamp = new Date().toLocaleTimeString();
  const logItem = document.createElement("div");
  logItem.className = `log-${type}`;
  logItem.textContent = `[${timestamp}] ${message}`;
  consoleEl.appendChild(logItem);
  consoleEl.scrollTop = consoleEl.scrollHeight;
}

// --- Cambio de Pestañas ---
function setupTabs() {
  const navItems = document.querySelectorAll(".nav-item");
  navItems.forEach(btn => {
    btn.addEventListener("click", () => {
      // Remover activas
      document.querySelectorAll(".nav-item").forEach(item => item.classList.remove("active"));
      document.querySelectorAll(".tab-view").forEach(view => view.classList.remove("active"));
      
      // Agregar activa al botón
      btn.classList.add("active");
      
      // Mostrar vista destino
      const targetId = btn.getAttribute("data-target");
      document.getElementById(targetId).classList.add("active");
    });
  });
}

// --- Formateo de Texto Enriquecido ---
function format(command, value = null) {
  document.execCommand(command, false, value);
  document.getElementById("emailBodyEditor").focus();
}

// --- Inserción de Placeholders en el Editor/Asunto ---
function insertPlaceholder(placeholderName) {
  const placeholder = `{${placeholderName}}`;
  const subjectInput = document.getElementById("emailSubjectInput");
  const bodyEditor = document.getElementById("emailBodyEditor");

  if (document.activeElement === subjectInput) {
    const start = subjectInput.selectionStart;
    const end = subjectInput.selectionEnd;
    const oldText = subjectInput.value;
    subjectInput.value = oldText.slice(0, start) + placeholder + oldText.slice(end);
    subjectInput.focus();
    subjectInput.setSelectionRange(start + placeholder.length, start + placeholder.length);
  } else {
    bodyEditor.focus();
    const sel = window.getSelection();
    if (sel.getRangeAt && sel.rangeCount) {
      let range = sel.getRangeAt(0);
      range.deleteContents();
      const node = document.createTextNode(placeholder);
      range.insertNode(node);
      range.setStartAfter(node);
      range.setEndAfter(node);
      sel.removeAllRanges();
      sel.addRange(range);
    } else {
      // Si no hay foco ni selección previa
      bodyEditor.innerHTML += placeholder;
    }
  }
  toast(`Variable ${placeholder} insertada.`);
}

// --- Refrescar Vista a partir del Estado ---
async function refreshFromState(newState) {
  state = newState;
  renderTemplateStatus();
  renderTable();
  renderMappingPanel();
  renderEmailTemplate();
  renderVariables();
  renderAddressBook();
  renderSelectionPanel();

  const hasFields = state.fields && state.fields.length > 0;
  document.getElementById("btnExcel").disabled = !hasFields;
  document.getElementById("btnAddRow").disabled = !hasFields;
  
  const selectedCount = (state.rows || []).filter(r => r.selected !== false).length;
  const canSend = Boolean(state.template_path && selectedCount > 0);
  const startBtn = document.getElementById("btnStartSend");
  if (startBtn) {
    startBtn.disabled = !canSend;
  }

  document.getElementById("emptyHint").style.display = hasFields ? "none" : "block";
  document.getElementById("rowCountBadge").textContent = `${state.rows ? state.rows.length : 0} Registros`;
}

async function refreshState() {
  const s = await api().get_state();
  await refreshFromState(s);
}

// --- Renderizar Estado de Archivos ---
function renderTemplateStatus() {
  const tStatus = document.getElementById("templateStatus");
  const tDot = document.getElementById("templateDot");
  if (state.template_path) {
    tStatus.textContent = `Plantilla: ${state.template_path.split(/[\\/]/).pop()}`;
    tDot.classList.add("active");
  } else {
    tStatus.textContent = "Sin plantilla cargada";
    tDot.classList.remove("active");
  }

  const oStatus = document.getElementById("outputStatus");
  const oDot = document.getElementById("outputDot");
  if (state.output_folder) {
    oStatus.textContent = `Salida: ${state.output_folder}`;
    oDot.classList.add("active");
  } else {
    oStatus.textContent = "Sin carpeta de salida";
    oDot.classList.remove("active");
  }
}

// --- Renderizar Panel de Mapeo de Columnas ---
function renderMappingPanel() {
  const panel = document.getElementById("mapping-panel");
  if (!state.excel_path || !state.fields || !state.fields.length) {
    panel.style.display = "none";
    return;
  }
  panel.style.display = "block";

  const container = document.getElementById("mappingTable");
  container.innerHTML = "";
  state.fields.forEach((field) => {
    const mappedCol = state.excel_mapping && state.excel_mapping[field];
    const row = document.createElement("div");
    row.className = "mapping-row" + (mappedCol ? "" : " unmapped");

    const label = document.createElement("span");
    label.textContent = field;
    const value = document.createElement("span");
    value.textContent = mappedCol ? `← columna "${mappedCol}"` : "(sin mapeo automático, edita la tabla a mano)";

    row.appendChild(label);
    row.appendChild(value);
    container.appendChild(row);
  });
}

// --- Renderizar Tabla de Datos ---
function renderTable() {
  const head = document.getElementById("tableHead");
  const body = document.getElementById("tableBody");
  head.innerHTML = "";
  body.innerHTML = "";

  if (!state.fields || !state.fields.length) return;

  // Cabecera de Casilla Master (Seleccionar Todos)
  const masterTh = document.createElement("th");
  masterTh.style.width = "36px";
  masterTh.style.textAlign = "center";
  const masterCheck = document.createElement("input");
  masterCheck.type = "checkbox";
  masterCheck.title = "Seleccionar/Deseleccionar Todos";
  const rowsList = state.rows || [];
  masterCheck.checked = rowsList.length > 0 && rowsList.every(r => r.selected !== false);
  masterCheck.addEventListener("change", async () => {
    const s = await api().toggle_all_selection(masterCheck.checked);
    await refreshFromState(s.state);
  });
  masterTh.appendChild(masterCheck);
  head.appendChild(masterTh);

  // Cabeceras de tabla
  state.fields.forEach((f) => {
    const th = document.createElement("th");
    th.textContent = f;
    head.appendChild(th);
  });
  
  const statusTh = document.createElement("th");
  statusTh.textContent = "Estado";
  head.appendChild(statusTh);
  
  const actionsTh = document.createElement("th");
  actionsTh.textContent = "Acciones";
  head.appendChild(actionsTh);

  // Filas de tabla
  (state.rows || []).forEach((row, index) => {
    const tr = document.createElement("tr");

    // Casilla de Selección por Fila
    const checkTd = document.createElement("td");
    checkTd.style.textAlign = "center";
    const rowCheck = document.createElement("input");
    rowCheck.type = "checkbox";
    rowCheck.checked = row.selected !== false;
    rowCheck.addEventListener("change", async () => {
      const s = await api().toggle_row_selection(index, rowCheck.checked);
      await refreshFromState(s.state);
    });
    checkTd.appendChild(rowCheck);
    tr.appendChild(checkTd);

    state.fields.forEach((field) => {
      const td = document.createElement("td");
      td.contentEditable = "true";
      td.textContent = row.data[field] || "";
      td.addEventListener("blur", () => {
        api().update_cell(index, field, td.textContent).catch(() => {});
      });
      tr.appendChild(td);
    });

    // Columna de Estado
    const statusTd = document.createElement("td");
    const select = document.createElement("select");
    select.className = "status-select";
    ["Pendiente", "Generado", "Enviado", "Respondido"].forEach((opt) => {
      const o = document.createElement("option");
      o.value = opt;
      o.textContent = opt;
      if (row.status === opt) o.selected = true;
      select.appendChild(o);
    });
    select.addEventListener("change", async () => {
      await api().update_status(index, select.value);
      await refreshState();
    });
    statusTd.appendChild(select);
    tr.appendChild(statusTd);

    // Columna de Acción (Ver PDF y Eliminar)
    const actionsTd = document.createElement("td");
    
    const pdfBtn = document.createElement("button");
    pdfBtn.textContent = "📄 Ver PDF";
    pdfBtn.className = "btn-small-info";
    pdfBtn.addEventListener("click", async () => {
      toast("Generando PDF de auditoría...", "info");
      pdfBtn.disabled = true;
      pdfBtn.textContent = "⏳...";
      try {
        const res = await api().generate_single_pdf(index);
        if (res.ok) {
            toast("PDF generado y abierto.", "success");
            logToConsole(`PDF de auditoría abierto para fila ${index + 1}.`, "success");
        } else {
            toast(res.error, "error");
            logToConsole(`Error PDF fila ${index + 1}: ${res.error}`, "error");
        }
      } finally {
        pdfBtn.disabled = false;
        pdfBtn.textContent = "📄 Ver PDF";
      }
    });
    actionsTd.appendChild(pdfBtn);

    const delBtn = document.createElement("button");
    delBtn.textContent = "Eliminar";
    delBtn.className = "danger-small";
    delBtn.addEventListener("click", async () => {
      const s = await api().delete_row(index);
      await refreshFromState(s);
    });
    actionsTd.appendChild(delBtn);

    tr.appendChild(actionsTd);

    body.appendChild(tr);
  });
}

// --- Renderizar Panel de Selección en Pestaña 3 ---
function renderSelectionPanel() {
  const listEl = document.getElementById("selectionList");
  const badgeEl = document.getElementById("selectedCountBadge");
  if (!listEl || !badgeEl) return;
  listEl.innerHTML = "";

  const rows = state.rows || [];
  const selectedCount = rows.filter(r => r.selected !== false).length;
  badgeEl.textContent = `${selectedCount} / ${rows.length} seleccionados`;

  if (rows.length === 0) {
    listEl.innerHTML = "<span style='color: var(--text-muted); font-size: 12px;'>No hay registros cargados. Carga primero un Excel o plantilla.</span>";
    return;
  }

  rows.forEach((r, idx) => {
    const item = document.createElement("label");
    item.style.display = "flex";
    item.style.alignItems = "center";
    item.style.gap = "10px";
    item.style.padding = "6px 8px";
    item.style.borderBottom = "1px solid #e2e8f0";
    item.style.fontSize = "12px";
    item.style.cursor = "pointer";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = r.selected !== false;
    cb.addEventListener("change", async () => {
      const s = await api().toggle_row_selection(idx, cb.checked);
      await refreshFromState(s.state);
    });

    let labelText = `Fila ${idx + 1}`;
    for (const key of ["empresa", "nombre_empresa", "compania", "cliente", "razon_social", "destinatario_correo"]) {
      if (r.data && r.data[key]) {
        labelText += `: ${r.data[key]}`;
        break;
      }
    }

    const span = document.createElement("span");
    span.textContent = labelText;

    item.appendChild(cb);
    item.appendChild(span);
    listEl.appendChild(item);
  });
}

// --- Renderizar Redactor de Correo ---
function renderEmailTemplate() {
  if (state.email_to_template !== undefined) {
    document.getElementById("emailToInput").value = state.email_to_template;
  }
  if (state.email_cc_template !== undefined) {
    document.getElementById("emailCcInput").value = state.email_cc_template;
  }
  if (state.email_subject_template !== undefined) {
    document.getElementById("emailSubjectInput").value = state.email_subject_template;
  }
  if (state.email_body_template !== undefined) {
    document.getElementById("emailBodyEditor").innerHTML = state.email_body_template;
  }
}

// --- Renderizar Lista de Variables ---
function renderVariables() {
  const container = document.getElementById("variablesContainer");
  container.innerHTML = "";
  if (!state.fields || !state.fields.length) {
    const hint = document.createElement("span");
    hint.className = "sidebar-desc";
    hint.textContent = "Carga una plantilla Word para ver las variables.";
    container.appendChild(hint);
    return;
  }

  state.fields.forEach(field => {
    const pill = document.createElement("button");
    pill.className = "var-pill";
    pill.textContent = field;
    pill.addEventListener("click", () => {
      insertPlaceholder(field);
    });
    container.appendChild(pill);
  });
}

// --- Renderizar Agenda de Contactos ---
function renderAddressBook() {
  const container = document.getElementById("addressBookContainer");
  container.innerHTML = "";
  const book = state.address_book || [];
  
  if (book.length === 0) {
    container.innerHTML = '<li style="font-size:12px; color:var(--text-muted);">Sin contactos aún.</li>';
    return;
  }
  
  book.forEach(email => {
    const li = document.createElement("li");
    li.style.display = "flex";
    li.style.justifyContent = "space-between";
    li.style.alignItems = "center";
    li.style.padding = "4px 8px";
    li.style.background = "#f1f5f9";
    li.style.borderRadius = "4px";
    
    const emailSpan = document.createElement("span");
    emailSpan.textContent = email;
    emailSpan.style.fontSize = "12px";
    emailSpan.style.fontWeight = "500";
    emailSpan.style.overflow = "hidden";
    emailSpan.style.textOverflow = "ellipsis";
    emailSpan.style.maxWidth = "110px";
    emailSpan.title = email;
    
    const actions = document.createElement("div");
    actions.style.display = "flex";
    actions.style.gap = "4px";
    
    const btnTo = document.createElement("button");
    btnTo.textContent = "Para";
    btnTo.className = "btn-small";
    btnTo.style.fontSize = "10px";
    btnTo.style.padding = "2px 4px";
    btnTo.onclick = () => {
      const input = document.getElementById("emailToInput");
      input.value = input.value ? input.value + "; " + email : email;
      toast("Agregado a Para.");
    };
    
    const btnCc = document.createElement("button");
    btnCc.textContent = "CC";
    btnCc.className = "btn-small";
    btnCc.style.fontSize = "10px";
    btnCc.style.padding = "2px 4px";
    btnCc.onclick = () => {
      const input = document.getElementById("emailCcInput");
      input.value = input.value ? input.value + "; " + email : email;
      toast("Agregado a CC.");
    };
    
    const btnDel = document.createElement("button");
    btnDel.textContent = "✕";
    btnDel.className = "danger-small";
    btnDel.style.fontSize = "10px";
    btnDel.style.padding = "2px 4px";
    btnDel.onclick = async () => {
      const res = await api().remove_contact(email);
      if (res.ok) await refreshFromState(res.state);
    };
    
    actions.appendChild(btnTo);
    actions.appendChild(btnCc);
    actions.appendChild(btnDel);
    
    li.appendChild(emailSpan);
    li.appendChild(actions);
    container.appendChild(li);
  });
}

// --- Enlazar Botones y Eventos ---
function wireButtons() {
  document.getElementById("btnAddContact").addEventListener("click", async () => {
    const input = document.getElementById("newContactInput");
    const email = input.value.trim();
    if (!email) return;
    const res = await api().add_contact(email);
    if (res.ok) {
      input.value = "";
      toast("Contacto guardado.");
      await refreshFromState(res.state);
    }
  });

  // Cargar Plantilla Word
  document.getElementById("btnTemplate").addEventListener("click", async () => {
    logToConsole("Iniciando selección de plantilla Word...");
    const res = await api().choose_template();
    if (res.ok) {
      toast("Plantilla cargada con éxito.", "success");
      logToConsole(`Plantilla cargada. Campos variables detectados: ${res.state.fields.join(", ")}`, "success");
      await refreshFromState(res.state);
    } else if (res.error) {
      toast(res.error, "error");
      logToConsole(`Error cargando plantilla: ${res.error}`, "error");
    }
  });

  // Cargar Archivo Excel
  document.getElementById("btnExcel").addEventListener("click", async () => {
    logToConsole("Iniciando selección de archivo Excel...");
    const res = await api().choose_excel();
    if (res.ok) {
      toast("Excel cargado correctamente.", "success");
      logToConsole("Excel cargado y columnas mapeadas automáticamente al grid.", "success");
      await refreshFromState(res.state);
    } else if (res.error) {
      toast(res.error, "error");
      logToConsole(`Error cargando Excel: ${res.error}`, "error");
    }
  });

  // Carpeta de Salida
  document.getElementById("btnOutput").addEventListener("click", async () => {
    logToConsole("Seleccionando carpeta de salida...");
    const res = await api().choose_output_folder();
    if (res.ok) {
      toast("Carpeta de salida configurada.", "success");
      logToConsole(`Carpeta de salida establecida en: ${res.state.output_folder}`, "success");
      await refreshFromState(res.state);
    }
  });

  // Agregar Registro Manual
  document.getElementById("btnAddRow").addEventListener("click", async () => {
    logToConsole("Agregando fila manual al grid de datos.");
    const s = await api().add_row();
    await refreshFromState(s);
  });



  // Guardar Cambios de Plantilla de Correo
  document.getElementById("btnSaveTemplate").addEventListener("click", async () => {
    const toVal = document.getElementById("emailToInput").value;
    const ccVal = document.getElementById("emailCcInput").value;
    const subject = document.getElementById("emailSubjectInput").value;
    const body = document.getElementById("emailBodyEditor").innerHTML;
    logToConsole("Guardando cambios en la plantilla del correo...");
    const res = await api().save_email_template(subject, body, toVal, ccVal);
    if (res.ok) {
      toast("Plantilla de correo guardada.", "success");
      logToConsole("Destinatarios, asunto y cuerpo del correo guardados en el estado de la campaña.", "success");
      await refreshFromState(res.state);
    }
  });

  // Enviar / Crear Borradores en Outlook
  document.getElementById("btnStartSend").addEventListener("click", async () => {
    const sendMode = document.querySelector('input[name="sendMode"]:checked').value;
    const isDirectSend = (sendMode === "send");
    
    // Primero, guardar la plantilla actual para estar seguros
    const toVal = document.getElementById("emailToInput").value;
    const ccVal = document.getElementById("emailCcInput").value;
    const subject = document.getElementById("emailSubjectInput").value;
    const body = document.getElementById("emailBodyEditor").innerHTML;
    await api().save_email_template(subject, body, toVal, ccVal);

    logToConsole(`Iniciando conexión con Outlook (Modo: ${isDirectSend ? "Envío directo" : "Crear borradores"})...`);
    
    // UI Progress State
    const progressCard = document.getElementById("progressCard");
    const progressBarFill = document.getElementById("progressBarFill");
    const progressText = document.getElementById("progressText");
    const progressPercent = document.getElementById("progressPercent");
    const startBtn = document.getElementById("btnStartSend");

    progressCard.style.display = "block";
    progressBarFill.style.width = "0%";
    progressPercent.textContent = "0%";
    progressText.textContent = "Estableciendo conexión COM con Microsoft Outlook...";
    startBtn.disabled = true;

    try {
      const res = await api().generate_outlook_drafts(isDirectSend);
      if (res.ok) {
        progressBarFill.style.width = "100%";
        progressPercent.textContent = "100%";
        
        let msg = `Proceso finalizado. Registros procesados con éxito: ${res.processed}.`;
        if (res.errors && res.errors.length) {
          msg += ` Errores detectados: ${res.errors.length}.`;
          res.errors.forEach(err => logToConsole(err, "error"));
        }
        
        progressText.textContent = msg;
        toast(msg, res.errors && res.errors.length > 0 ? "error" : "success");
        logToConsole(msg, "success");
        await refreshFromState(res.state);
      } else {
        progressText.textContent = `Error: ${res.error}`;
        toast(res.error, "error");
        logToConsole(`Error en procesamiento Outlook: ${res.error}`, "error");
      }
    } catch (e) {
      progressText.textContent = `Error del sistema: ${e}`;
      logToConsole(`Error crítico del sistema: ${e}`, "error");
    } finally {
      startBtn.disabled = false;
    }
  });

  // Limpiar Consola de Log
  document.getElementById("btnClearConsole").addEventListener("click", () => {
    document.getElementById("consoleLog").innerHTML = "";
    logToConsole("Consola limpia.");
  });

  // Botones de Selección Rápida en Pestaña 3
  const btnSelectAll = document.getElementById("btnSelectAll");
  if (btnSelectAll) {
    btnSelectAll.addEventListener("click", async () => {
      const s = await api().toggle_all_selection(true);
      await refreshFromState(s.state);
      toast("Todos los registros seleccionados.", "info");
    });
  }

  const btnDeselectAll = document.getElementById("btnDeselectAll");
  if (btnDeselectAll) {
    btnDeselectAll.addEventListener("click", async () => {
      const s = await api().toggle_all_selection(false);
      await refreshFromState(s.state);
      toast("Todos los registros desmarcados.", "info");
    });
  }

  const btnSelectTest = document.getElementById("btnSelectTest");
  if (btnSelectTest) {
    btnSelectTest.addEventListener("click", async () => {
      const s = await api().select_only_first();
      await refreshFromState(s.state);
      toast("Seleccionado únicamente el 1er registro para pruebas.", "success");
      logToConsole("Modo prueba: Marcado 1 solo registro (Fila 1). Listo para enviar.", "info");
    });
  }
}

// --- Inicio al cargar WebView ---
window.addEventListener("pywebviewready", async () => {
  setupTabs();
  wireButtons();
  await refreshState();
  logToConsole("Sistema inicializado y listo para operar.");
});
