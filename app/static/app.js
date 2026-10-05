const state = {
  view: "dashboard",
  entity: "movement",
  tableData: null,
  selectedRow: null,
  meta: null,
  pendingPlan: null,
  searchTimer: null,
  conversationId: null,
  conversations: [],
};

const labels = {
  dashboard: "Resumen", assistant: "Asistente IA", movement: "Movimientos", document: "Documentos",
  third_party: "Terceros", account: "Cuentas", category: "Categorías", rule: "Reglas del asistente", logs: "Registro de actividad",
};
const visibleColumns = {
  movement: ["ID", "Fecha", "Tipo", "Tercero", "Concepto", "Categoria", "Total", "Estado"],
  document: ["ID", "Fecha", "TipoDocumento", "Numero", "Tercero", "Total", "Estado", "TipoArchivo"],
  third_party: ["ID", "Tipo", "NombreRazonSocial", "NIT_ID", "Email", "Telefono", "Estado"],
  account: ["ID", "Cuenta", "Tipo", "Entidad", "NumeroReferencia", "SaldoInicial", "SaldoActual", "Activa"],
  category: ["Codigo", "Tipo", "Categoria", "Subcategoria", "Activa", "Descripcion"],
  rule: ["ID", "Activa", "PalabraClave", "Tipo", "Categoria", "Subcategoria", "Prioridad", "Observaciones"],
};
const fieldLabels = {
  Fecha: "Fecha", Tipo: "Tipo", Documento: "Documento", Numero: "Número", Tercero: "Tercero", NIT_ID: "NIT / identificación",
  Concepto: "Concepto", Categoria: "Categoría", Subcategoria: "Subcategoría", MedioPago: "Medio de pago", Cuenta: "Cuenta",
  Subtotal: "Subtotal / base", Impuestos: "Impuestos", Retenciones: "Retenciones", Estado: "Estado", SoporteID: "ID del soporte",
  Observaciones: "Observaciones", TipoDocumento: "Tipo de documento", Identificacion: "Identificación", Total: "Total del documento",
  CUFE_Referencia: "CUFE / referencia", RutaArchivo: "Ruta del archivo", NombreArchivo: "Nombre del archivo", TipoArchivo: "Tipo de archivo",
  FechaExtraccion: "Fecha de extracción", MetodoExtraccion: "Método de extracción", Confianza: "Confianza", TipoTercero: "Tipo",
  NombreRazonSocial: "Nombre o razón social", Contacto: "Contacto", Email: "Correo", Telefono: "Teléfono", Direccion: "Dirección",
  Ciudad: "Ciudad", SaldoInicial: "Saldo inicial", Activa: "Activa", NumeroReferencia: "Número / referencia", Entidad: "Entidad",
  Categoria: "Categoría", Subcategoria: "Subcategoría", Descripcion: "Descripción", PalabraClave: "Palabra clave",
  TerceroOpcional: "Tercero específico (opcional)", Prioridad: "Prioridad (1–100)",
};
const choiceOptions = {
  Tipo: ["Ingreso", "Gasto", "Compra", "Venta", "Activo", "Otro"],
  "Estado@movement": ["Pendiente", "Pagado", "Anulado"],
  "Estado@document": ["Pendiente", "Validado", "Revision Manual", "Duplicado", "Error", "Anulado"],
  "Estado@third_party": ["Activo", "Inactivo"],
  TipoDocumento: ["Factura electronica", "Documento equivalente", "Documento soporte", "Nota credito", "Nota debito", "Recibo", "Comprobante", "Extracto bancario", "Nomina", "Contrato", "Otro"],
  TipoTercero: ["Cliente", "Proveedor", "Empleado", "Socio", "Entidad publica", "Otro"],
  Activa: ["Si", "No"],
  ActivaRule: ["Sí", "No"],
  TipoCuenta: ["Banco", "Caja", "Efectivo", "Billetera digital", "Tarjeta", "Cuenta por cobrar", "Cuenta por pagar", "Otros"],
};

const $ = (selector, root = document) => root.querySelector(selector);
const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];

function esc(value) {
  return String(value ?? "").replace(/[&<>"']/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]));
}

async function api(url, options = {}) {
  const response = await fetch(url, { headers: { "Content-Type": "application/json" }, ...options });
  const body = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(body.error || `Error ${response.status}`);
  return body;
}

function showToast(message, isError = false) {
  const toast = $("#toast");
  toast.textContent = message;
  toast.classList.toggle("error", isError);
  toast.classList.add("show");
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toast.classList.remove("show"), 4200);
}

function money(value) {
  return new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(Number(value) || 0);
}

function navigate(view) {
  state.view = view;
  $$(".view").forEach(section => section.classList.remove("active"));
  const target = view === "dashboard" ? "#view-dashboard" : view === "assistant" ? "#view-assistant" : view === "logs" ? "#view-logs" : "#view-data";
  $(target).classList.add("active");
  $$(".nav-item").forEach(button => button.classList.toggle("active", button.dataset.view === view));
  $("#breadcrumb").textContent = labels[view] || "Resumen";
  if (["movement", "document", "third_party", "account", "category", "rule"].includes(view)) {
    state.entity = view;
    $("#tableTitle").textContent = labels[view];
    $("#tableDescription").textContent = `Consulta y administra ${labels[view].toLowerCase()} en el libro seleccionado.`;
    $("#tableEyebrow").textContent = `MÓDULO · ${labels[view].toUpperCase()}`;
    $("#tableSearch").value = "";
    state.selectedRow = null;
    $("#editRecord").disabled = true;
    loadTable();
  } else if (view === "dashboard") {
    loadDashboard();
  } else if (view === "logs") {
    loadLogs();
  }
}

function prettyField(key) {
  return fieldLabels[key] || key.replaceAll("_", " ");
}

function renderState(value) {
  const lower = String(value || "").toLowerCase();
  let cls = "";
  if (["pendiente", "revision manual", "revisión manual"].includes(lower)) cls = "pending";
  else if (["pagado", "activo", "validado", "si", "sí"].includes(lower)) cls = "paid";
  else if (["error", "anulado", "duplicado", "inactivo"].includes(lower)) cls = "error";
  return `<span class="state-pill ${cls}">${esc(value || "—")}</span>`;
}

function cellHtml(key, value) {
  if (key === "Total" || key === "Subtotal" || key === "SaldoInicial" || key === "SaldoActual" || key === "Impuestos" || key === "Retenciones") return esc(money(value));
  if (key === "Estado" || key === "Activa") return renderState(value);
  if (key === "Tipo" && ["Ingreso", "Venta", "Gasto", "Compra"].includes(value)) {
    return `<span class="kind-pill ${["Ingreso", "Venta"].includes(value) ? "income" : "expense"}">${esc(value)}</span>`;
  }
  return esc(value || "—");
}

async function loadStatus() {
  try {
    const status = await api("/api/status");
    $("#sidebarBook").textContent = status.workbook;
    const good = Boolean(status.ai.available);
    $("#aiDot").className = `status-dot ${good ? "on" : "off"}`;
    $("#assistantAiDot").className = `status-dot ${good ? "on" : "off"}`;
    $("#aiStatus").textContent = good ? "IA local conectada" : "IA local desconectada";
    $("#assistantAiStatus").textContent = good ? "Conectado y listo" : "Ollama no está disponible";
    $("#aiModel").textContent = status.ai.model || "Inicia Ollama para usar IA";
    $("#assistantAiModel").textContent = status.ai.model || "Sin modelo local activo";
    $("#chatModelLine").textContent = good ? `${status.ai.model} · tus datos no salen de este equipo` : "Puedes consultar el libro; inicia Ollama para usar el asistente IA";
  } catch (error) {
    $("#aiStatus").textContent = "No se encontró el libro";
    $("#sidebarBook").textContent = "Selecciona un libro";
    $("#aiDot").className = "status-dot off";
    $("#assistantAiDot").className = "status-dot off";
    $("#assistantAiStatus").textContent = "No se pudo comprobar la conexión";
    $("#aiModel").textContent = "Sin conexión";
    $("#assistantAiModel").textContent = "Sin conexión";
  }
}

async function loadDashboard() {
  try {
    const data = await api("/api/dashboard");
    $("#todayLabel").textContent = new Intl.DateTimeFormat("es-CO", { dateStyle: "medium" }).format(new Date());
    const period = new Intl.DateTimeFormat("es-CO", { month: "long", year: "numeric" }).format(new Date());
    $("#periodLabel").textContent = `Mes actual · ${period}`;
    $("#metricIncome").textContent = money(data.income);
    $("#metricExpenses").textContent = money(data.expenses);
    $("#metricNet").textContent = money(data.net);
    $("#metricReceivable").textContent = money(data.receivable);
    $("#metricPayableLabel").textContent = `Por pagar: ${money(data.payable)}`;
    $("#pendingDocs").textContent = data.documents_pending;
    $("#errorDocs").textContent = data.documents_error;
    $("#registeredDocs").textContent = data.document_count;
    const rows = data.recent_movements || [];
    const headers = ["Fecha", "Concepto", "Tercero", "Tipo", "Total", "Estado"];
    $("#recentTable thead").innerHTML = `<tr>${headers.map(h => `<th>${esc(prettyField(h))}</th>`).join("")}</tr>`;
    $("#recentTable tbody").innerHTML = rows.length ? rows.map(row => `<tr>${headers.map(h => `<td>${cellHtml(h, row[h])}</td>`).join("")}</tr>`).join("") : `<tr><td class="empty-cell" colspan="${headers.length}">Aún no hay movimientos. Puedes crear uno con el asistente.</td></tr>`;
  } catch (error) {
    showToast(error.message, true);
  }
}

async function loadTable() {
  const body = $("#moduleTable tbody");
  body.innerHTML = `<tr><td class="loading-row"><span class="spinner"></span> Leyendo el libro…</td></tr>`;
  try {
    const params = new URLSearchParams({ entity: state.entity, search: $("#tableSearch").value || "" });
    const data = await api(`/api/table?${params}`);
    state.tableData = data;
    const preferred = visibleColumns[state.entity] || data.headers;
    const columns = preferred.filter(key => data.headers.includes(key));
    $("#moduleTable thead").innerHTML = `<tr>${columns.map(key => `<th>${esc(prettyField(key))}</th>`).join("")}</tr>`;
    $("#moduleTable tbody").innerHTML = data.rows.length ? data.rows.map((row, index) => `<tr data-index="${index}" data-id="${esc(row[data.id_field])}">${columns.map(key => `<td title="${esc(row[key] ?? "")}">${cellHtml(key, row[key])}</td>`).join("")}</tr>`).join("") : `<tr><td class="empty-cell" colspan="${columns.length || 1}">No hay registros para mostrar en este módulo.</td></tr>`;
    $("#rowCount").textContent = `${data.count} ${data.count === 1 ? "registro" : "registros"}`;
    state.selectedRow = null;
    $("#editRecord").disabled = true;
    $$("#moduleTable tbody tr[data-index]").forEach(row => row.addEventListener("click", () => {
      $$("#moduleTable tbody tr").forEach(r => r.classList.remove("selected"));
      row.classList.add("selected");
      state.selectedRow = data.rows[Number(row.dataset.index)];
      $("#editRecord").disabled = false;
    }));
  } catch (error) {
    body.innerHTML = `<tr><td class="empty-cell">${esc(error.message)}</td></tr>`;
    showToast(error.message, true);
  }
}

async function loadLogs() {
  const table = $("#logTable");
  table.querySelector("tbody").innerHTML = `<tr><td class="loading-row"><span class="spinner"></span> Cargando historial…</td></tr>`;
  try {
    const rows = await api("/api/logs");
    const headers = ["Fecha", "Hora", "Usuario", "Accion", "Modulo", "RegistroAfectado", "Resultado"];
    table.querySelector("thead").innerHTML = `<tr>${headers.map(h => `<th>${esc(prettyField(h))}</th>`).join("")}</tr>`;
    table.querySelector("tbody").innerHTML = rows.length ? rows.map(row => `<tr>${headers.map(h => `<td>${cellHtml(h, row[h])}</td>`).join("")}</tr>`).join("") : `<tr><td class="empty-cell" colspan="${headers.length}">Aún no hay acciones registradas desde la app.</td></tr>`;
  } catch (error) { showToast(error.message, true); }
}

function addMessage(role, text, options = {}) {
  const messages = $("#chatMessages");
  const item = document.createElement("div");
  item.className = `message ${role === "user" ? "user-message" : "assistant-message"}${options.loading ? " loading-message" : ""}`;
  const badge = options.badge ? `<span class="message-status">${esc(options.badge)}</span>` : "";
  item.innerHTML = `<div class="message-avatar">${role === "user" ? "A" : "✳"}</div><div class="message-body"><strong>${role === "user" ? "Tú" : "Asistente"}</strong><p>${options.loading ? `<span class="loading-row"><span class="spinner"></span>${esc(text)}</span>` : esc(text)}</p>${badge}</div>`;
  messages.appendChild(item);
  messages.scrollTop = messages.scrollHeight;
  return item;
}

function showWelcomeMessage() {
  $("#chatMessages").innerHTML = `<div class="message assistant-message"><div class="message-avatar">✳</div><div class="message-body"><strong>Asistente</strong><p>Puedo ayudarte con tus datos en lenguaje cotidiano. Guarda conversaciones y preferencias aquí; revisarás cualquier cambio antes de confirmarlo.</p></div></div>`;
}

async function refreshConversationList(selectedId = state.conversationId) {
  const result = await api("/api/assistant/conversations");
  state.conversations = result.conversations || [];
  const select = $("#conversationSelect");
  select.innerHTML = "";
  if (!state.conversations.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "Sin conversaciones";
    select.appendChild(option);
    select.disabled = true;
    $("#deleteConversation").disabled = true;
    return;
  }
  select.disabled = false;
  for (const conversation of state.conversations) {
    const option = document.createElement("option");
    option.value = conversation.id;
    const date = conversation.updated_at ? new Date(conversation.updated_at).toLocaleDateString("es-CO") : "";
    option.textContent = `${conversation.title || "Conversación"}${date ? ` · ${date}` : ""}`;
    select.appendChild(option);
  }
  const selected = state.conversations.some(item => item.id === selectedId) ? selectedId : state.conversations[0].id;
  select.value = selected;
  state.conversationId = selected;
  $("#deleteConversation").disabled = false;
}

async function loadConversation(conversationId) {
  if (!conversationId) return;
  const conversation = await api(`/api/assistant/conversation?id=${encodeURIComponent(conversationId)}`);
  state.conversationId = conversation.id;
  const messages = $("#chatMessages");
  messages.innerHTML = "";
  if (!conversation.messages?.length) showWelcomeMessage();
  for (const item of conversation.messages || []) {
    let text = item.content;
    let badge = "";
    if (item.metadata?.kind === "proposal") {
      const status = item.metadata.status || "pending";
      badge = ({ pending: "Propuesta pendiente · sin guardar", applied: "Confirmada y guardada", cancelled: "Cancelada · sin guardar", expired: "Caducada · sin guardar" })[status] || "Propuesta";
      const summaries = (item.metadata.actions || []).map(action => {
        const values = Object.entries(action.data || {}).map(([key, value]) => `${prettyField(key)}: ${value}`).join(" · ");
        return `${action.title}${action.id ? ` (${action.id})` : ""}${values ? ` — ${values}` : ""}`;
      });
      if (summaries.length) text += `\n\n${summaries.join("\n")}`;
      const memories = item.metadata.memory_updates || [];
      if (memories.length) text += `\n\nMemoria propuesta: ${memories.map(memory => `${memory.key}: ${memory.value}`).join("; ")}`;
    }
    addMessage(item.role, text, { badge });
  }
  await refreshConversationList(conversation.id);
}

async function startNewConversation() {
  if (state.pendingPlan) await cancelPlan();
  const result = await api("/api/assistant/conversations/new", { method: "POST", body: "{}" });
  state.conversationId = result.id;
  showWelcomeMessage();
  await refreshConversationList(result.id);
}

async function deleteCurrentConversation() {
  if (!state.conversationId) return;
  if (!window.confirm("¿Eliminar esta conversación y sus mensajes guardados? Esta acción no se puede deshacer.")) return;
  if (state.pendingPlan) await cancelPlan();
  const result = await api("/api/assistant/conversations/delete", {
    method: "POST", body: JSON.stringify({ conversation_id: state.conversationId }),
  });
  state.conversationId = null;
  await refreshConversationList("");
  if (state.conversations.length) await loadConversation(state.conversations[0].id);
  else await startNewConversation();
  showToast("Conversación eliminada.");
}

async function refreshMemories() {
  const data = await api("/api/assistant/memories");
  const memories = data.memories || [];
  $("#memoryCount").textContent = String(memories.length);
  $("#memoryList").innerHTML = memories.length ? memories.map(memory => {
    const title = String(memory.memory_key || "").replace(/[_-]+/g, " ");
    return `<article class="memory-item"><div class="memory-item-top"><strong>${esc(title)}</strong><button type="button" data-delete-memory="${Number(memory.id)}" title="Olvidar esta memoria" aria-label="Olvidar ${esc(title)}">×</button></div><p>${esc(memory.memory_value)}</p></article>`;
  }).join("") : `<p class="memory-empty">Dile «recuerda que…» y revisa la propuesta antes de guardar una preferencia para este libro.</p>`;
  $$("[data-delete-memory]").forEach(button => button.addEventListener("click", async () => {
    try {
      await api("/api/assistant/memories/delete", { method: "POST", body: JSON.stringify({ id: button.dataset.deleteMemory }) });
      await refreshMemories();
      showToast("Memoria eliminada.");
    } catch (error) { showToast(error.message, true); }
  }));
}

async function sendMessage(text) {
  const message = String(text || "").trim();
  if (!message) return;
  if (document.activeElement === $("#quickInput")) $("#quickInput").value = "";
  navigate("assistant");
  addMessage("user", message);
  const waiting = addMessage("assistant", "", { loading: true });
  $("#chatInput").value = "";
  $("#chatInput").disabled = true;
  $("#quickInput").disabled = true;
  try {
    const plan = await api("/api/assistant/plan", { method: "POST", body: JSON.stringify({ message, conversation_id: state.conversationId }) });
    state.conversationId = plan.conversation_id || state.conversationId;
    waiting.remove();
    addMessage("assistant", plan.reply || "Revisa la propuesta.");
    await refreshConversationList(state.conversationId);
    if (plan.actions?.length || plan.memory_updates?.length) showPlan(plan);
  } catch (error) {
    waiting.remove();
    addMessage("assistant", `No pude preparar el cambio: ${error.message}`);
  } finally {
    $("#chatInput").disabled = false;
    $("#quickInput").disabled = false;
    $("#chatInput").focus();
  }
}

function showPlan(plan) {
  state.pendingPlan = plan;
  $("#planReply").textContent = plan.reply || "Revisa estos cambios antes de guardarlos.";
  $("#planAssumptions").innerHTML = (plan.assumptions || []).map(note => `<div class="assumption-note">${esc(note)}</div>`).join("");
  $("#planActions").innerHTML = plan.actions.map(action => {
    const clearFields = new Set(action.clear_fields || []);
    const fields = Object.entries(action.data || {}).map(([key, value]) => `<div class="action-field"><span>${esc(prettyField(key))}</span><span>${clearFields.has(key) ? "Se vaciará" : esc(value ?? "—")}</span></div>`).join("");
    const target = action.id ? ` · ${esc(action.id)}` : "";
    return `<article class="action-card"><div class="action-card-title">${esc(action.title)}${target}</div><div class="action-fields">${fields || `<div class="action-field"><span>Campos</span><span>Se completan con valores predeterminados</span></div>`}</div></article>`;
  }).join("");
  $("#planMemories").innerHTML = (plan.memory_updates || []).length ? `<section class="memory-proposal"><div class="memory-proposal-title">El asistente recordará para este libro</div>${plan.memory_updates.map(memory => `<p><strong>${esc(String(memory.key).replace(/[_-]+/g, " "))}:</strong> ${esc(memory.value)}</p>`).join("")}</section>` : "";
  $("#planModal").showModal();
}

async function cancelPlan() {
  const planId = state.pendingPlan?.plan_id;
  state.pendingPlan = null;
  if ($("#planModal").open) $("#planModal").close();
  if (planId) {
    try {
      await api("/api/assistant/cancel", { method: "POST", body: JSON.stringify({ plan_id: planId }) });
      addMessage("assistant", "Entendido, cancelé la propuesta; no se guardó ningún cambio.");
      await refreshConversationList(state.conversationId);
    }
    catch (_) { /* Cancelar localmente no depende de que el servidor responda. */ }
  }
}

async function applyPlan() {
  if (!state.pendingPlan?.plan_id) return;
  const button = $("#applyPlan");
  button.disabled = true;
  button.innerHTML = `<span class="spinner"></span> Guardando…`;
  try {
    const result = await api("/api/assistant/apply", { method: "POST", body: JSON.stringify({ plan_id: state.pendingPlan.plan_id }) });
    $("#planModal").close();
    const ids = result.results.map(item => item.id).join(", ");
    if (ids) addMessage("assistant", `Cambios guardados. Registro: ${ids}.${result.backup ? ` Copia de seguridad: ${result.backup}.` : ""}${(result.memory_saved || []).length ? ` También guardé ${result.memory_saved.length} preferencia${result.memory_saved.length === 1 ? "" : "s"} para este libro.` : ""}`);
    else addMessage("assistant", `Preferencia${(result.memory_saved || []).length === 1 ? " guardada" : "s guardadas"} en la memoria local de este libro.`);
    state.pendingPlan = null;
    showToast(ids ? `Cambios guardados · respaldo ${result.backup}` : "Memoria del asistente actualizada.");
    await refreshConversationList(state.conversationId);
    await refreshMemories();
    await refreshCurrentData();
  } catch (error) {
    state.pendingPlan = null;
    if ($("#planModal").open) $("#planModal").close();
    showToast(error.message, true);
  } finally {
    button.disabled = false;
    button.textContent = "Confirmar cambios";
  }
}

async function refreshCurrentData() {
  await loadStatus();
  if (state.view === "dashboard") await loadDashboard();
  if (state.view === "logs") await loadLogs();
  if (["movement", "document", "third_party", "account", "category", "rule"].includes(state.view)) await loadTable();
}

function inputType(field) {
  if (DATE_FIELDS.has(field)) return "date";
  if (field === "Email") return "email";
  return "text";
}
const DATE_FIELDS = new Set(["Fecha", "FechaExtraccion"]);
const LONG_FIELDS = new Set(["Concepto", "Observaciones", "Descripcion"]);
const SELECT_FIELDS = { "Tipo@third_party": "TipoTercero", "Tipo@account": "TipoCuenta", "Activa@rule": "ActivaRule" };

function localDateString() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
}

function formField(entity, field, value, isRequired, mode) {
  const key = SELECT_FIELDS[`${field}@${entity}`] || `${field}@${entity}`;
  const choices = state.meta?.choices || {};
  let configured = [];
  if (["movement", "rule"].includes(entity) && field === "Categoria") configured = [...new Set((choices.categories || []).map(item => item.Categoria).filter(Boolean))];
  if (["movement", "rule"].includes(entity) && field === "Subcategoria") configured = [...new Set((choices.categories || []).map(item => item.Subcategoria).filter(Boolean))];
  if (entity === "movement" && field === "Cuenta") configured = choices.accounts || [];
  if (value && !configured.includes(value)) configured.unshift(String(value));
  const baseOptions = choiceOptions[key] || choiceOptions[field];
  const configuredSelect = (configured.length && ["Categoria", "Subcategoria"].includes(field) && ["movement", "rule"].includes(entity)) || (configured.length && entity === "movement" && field === "Cuenta");
  const options = baseOptions ? [...new Set([...baseOptions, ...configured])] : configuredSelect ? configured : null;
  const suggestions = entity === "movement" && field === "Tercero" ? choices.third_parties || [] : [];
  const listId = `suggest-${entity}-${field}`;
  const full = LONG_FIELDS.has(field) ? " full" : "";
  const required = isRequired ? " required" : "";
  let control;
  if (options) {
    control = `<select name="${esc(field)}"${required}><option value="">Selecciona…</option>${options.map(opt => `<option value="${esc(opt)}" ${String(value ?? "") === opt ? "selected" : ""}>${esc(opt)}</option>`).join("")}</select>`;
  } else if (LONG_FIELDS.has(field)) {
    control = `<textarea name="${esc(field)}" placeholder="Escribe ${esc(prettyField(field).toLowerCase())}…">${esc(value ?? "")}</textarea>`;
  } else {
    const type = inputType(field);
    const numeric = ["Subtotal", "Impuestos", "Retenciones", "Total", "SaldoInicial", "Confianza", "Prioridad"].includes(field);
    const minMax = field === "Prioridad" ? ' min="1" max="100"' : "";
    const listAttr = suggestions.length ? ` list="${listId}"` : "";
    control = `<input name="${esc(field)}" type="${numeric ? "number" : type}" ${numeric ? 'step="any" inputmode="decimal"' : ""}${listAttr}${type === "date" && value ? ` value="${esc(String(value).slice(0, 10))}"` : ` value="${esc(value ?? "")}"`}${minMax}${required} placeholder="${numeric ? "0" : `Escribe ${esc(prettyField(field).toLowerCase())}`}">`;
    if (suggestions.length) control += `<datalist id="${listId}">${suggestions.map(item => `<option value="${esc(item)}"></option>`).join("")}</datalist>`;
  }
  const clearButton = mode === "edit" && !isRequired ? `<button type="button" class="field-clear" data-clear-field="${esc(field)}">Vaciar</button>` : "";
  return `<div class="form-field${full}"><div class="form-field-heading"><label>${esc(prettyField(field))}${isRequired ? " · obligatorio" : ""}</label>${clearButton}</div>${control}</div>`;
}

function openRecordForm(mode, record = null) {
  const entity = state.entity;
  const config = state.meta.entities[entity];
  const title = mode === "edit" ? `Editar ${config.label.toLowerCase()}` : `Nuevo ${config.label.toLowerCase()}`;
  $("#modalTitle").textContent = title;
  $("#recordForm").dataset.mode = mode;
  $("#recordForm").dataset.entity = entity;
  $("#recordForm").dataset.recordId = record?.[config.id_field] || "";
  $("#recordForm").dataset.requestId = globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const fields = config.fields;
  const requiredFields = {
    movement: ["Fecha", "Tipo", "Concepto", "Cuenta", "Subtotal"], document: ["Fecha", "TipoDocumento", "Numero", "Tercero", "Subtotal", "Total"],
    third_party: ["Tipo", "NombreRazonSocial"], account: ["Cuenta", "Tipo"], category: ["Tipo", "Categoria"], rule: ["PalabraClave", "Tipo", "Categoria", "Subcategoria"],
  }[entity] || [];
  $("#recordForm").dataset.clearFields = "[]";
  $("#formFields").innerHTML = `<div class="form-grid">${fields.map(field => formField(entity, field, record?.[field] ?? config.defaults?.[field] ?? "", requiredFields.includes(field), mode)).join("")}</div>`;
  if (mode === "create" && ["movement", "document"].includes(entity)) {
    const dateInput = $("input[name='Fecha']", $("#formFields"));
    if (dateInput && !dateInput.value) dateInput.value = localDateString();
  }
  $("#recordModal").showModal();
}

function toastAndRefresh(message, error = false) {
  showToast(message, error);
  refreshCurrentData();
}

async function saveManualRecord(event) {
  if (event.submitter?.value === "cancel") return;
  event.preventDefault();
  const form = event.currentTarget;
  const data = Object.fromEntries(new FormData(form).entries());
  for (const field of ["Subtotal", "Impuestos", "Retenciones", "SaldoInicial", "Prioridad", "Confianza", "Total"]) {
    if (data[field] !== undefined && data[field] !== "") data[field] = Number(data[field]);
  }
  const payload = { entity: form.dataset.entity, action: form.dataset.mode === "edit" ? "update" : "create", id: form.dataset.recordId, data, clear_fields: JSON.parse(form.dataset.clearFields || "[]"), request_id: form.dataset.requestId };
  const button = $("#saveRecord");
  button.disabled = true;
  button.innerHTML = `<span class="spinner"></span> Guardando…`;
  try {
    const result = await api("/api/records/save", { method: "POST", body: JSON.stringify(payload) });
    $("#recordModal").close();
    const warnings = result.results[0].warnings || [];
    toastAndRefresh(`Registro ${result.results[0].id} guardado.${warnings.length ? ` ${warnings.join(" ")}` : ""} Respaldo: ${result.backup}`);
  } catch (error) {
    showToast(error.message, true);
  } finally {
    button.disabled = false;
    button.textContent = "Guardar cambios";
  }
}

async function chooseWorkbook() {
  try {
    const result = await api("/api/workbook/select", { method: "POST", body: "{}" });
    if (!result.cancelled) {
      await cancelPlan();
      state.meta = await api("/api/meta");
      state.conversationId = null;
      showToast(`Libro activo: ${result.workbook}`);
      await loadStatus();
      await loadDashboard();
      await refreshConversationList("");
      if (state.conversations.length) await loadConversation(state.conversations[0].id);
      else await startNewConversation();
      await refreshMemories();
      if (["movement", "document", "third_party", "account", "category", "rule"].includes(state.view)) await loadTable();
    }
  } catch (error) { showToast(error.message, true); }
}

function wireEvents() {
  $$(".nav-item").forEach(button => button.addEventListener("click", () => navigate(button.dataset.view)));
  $$('[data-go]').forEach(button => button.addEventListener("click", () => navigate(button.dataset.go)));
  $$("[data-prompt]").forEach(button => button.addEventListener("click", () => sendMessage(button.dataset.prompt)));
  $("#quickForm").addEventListener("submit", event => { event.preventDefault(); const text = $("#quickInput").value; $("#quickInput").value = ""; sendMessage(text); });
  $("#chatForm").addEventListener("submit", event => { event.preventDefault(); sendMessage($("#chatInput").value); });
  $("#changeBookButton").addEventListener("click", chooseWorkbook);
  $("#sidebarBookButton").addEventListener("click", chooseWorkbook);
  $("#newRecord").addEventListener("click", () => openRecordForm("create"));
  $("#editRecord").addEventListener("click", () => state.selectedRow && openRecordForm("edit", state.selectedRow));
  $("#recordForm").addEventListener("submit", saveManualRecord);
  $("#applyPlan").addEventListener("click", applyPlan);
  $("#cancelPlan").addEventListener("click", cancelPlan);
  $("#closePlan").addEventListener("click", cancelPlan);
  $("#planModal").addEventListener("cancel", event => { event.preventDefault(); cancelPlan(); });
  $("#formFields").addEventListener("click", event => {
    const button = event.target.closest("[data-clear-field]");
    if (!button) return;
    const field = button.dataset.clearField;
    const control = $("#recordForm").elements.namedItem(field);
    if (control) control.value = "";
    const clearing = new Set(JSON.parse($("#recordForm").dataset.clearFields || "[]"));
    clearing.add(field);
    $("#recordForm").dataset.clearFields = JSON.stringify([...clearing]);
    button.classList.add("is-clearing");
    button.textContent = "Se vaciará";
  });
  const removeClearMark = event => {
    const control = event.target;
    if (!control.name) return;
    const form = $("#recordForm");
    const clearing = new Set(JSON.parse(form.dataset.clearFields || "[]"));
    if (!clearing.delete(control.name)) return;
    form.dataset.clearFields = JSON.stringify([...clearing]);
    const button = $(`[data-clear-field="${CSS.escape(control.name)}"]`, $("#formFields"));
    if (button) { button.classList.remove("is-clearing"); button.textContent = "Vaciar"; }
  };
  $("#formFields").addEventListener("input", removeClearMark);
  $("#formFields").addEventListener("change", removeClearMark);
  $("#clearChat").addEventListener("click", () => startNewConversation().catch(error => showToast(error.message, true)));
  $("#deleteConversation").addEventListener("click", () => deleteCurrentConversation().catch(error => showToast(error.message, true)));
  $("#conversationSelect").addEventListener("change", event => {
    const nextId = event.target.value;
    if (!nextId || nextId === state.conversationId) return;
    (async () => {
      if (state.pendingPlan) await cancelPlan();
      await loadConversation(nextId);
    })().catch(error => showToast(error.message, true));
  });
  $("#refreshTable").addEventListener("click", loadTable);
  $("#refreshLogs").addEventListener("click", loadLogs);
  $("#tableSearch").addEventListener("input", () => {
    clearTimeout(state.searchTimer);
    state.searchTimer = setTimeout(loadTable, 180);
  });
  $("#helpButton").addEventListener("click", () => showToast("Pide al asistente que cree o edite un registro. La app siempre muestra la propuesta antes de guardarla."));
  for (const input of [$("#quickInput"), $("#chatInput")]) {
    input.addEventListener("keydown", event => {
      if (event.key === "Enter" && !event.shiftKey) {
        event.preventDefault();
        input.closest("form").requestSubmit();
      }
    });
  }
}

async function init() {
  $("#todayLabel").textContent = new Intl.DateTimeFormat("es-CO", { dateStyle: "medium" }).format(new Date());
  wireEvents();
  try { state.meta = await api("/api/meta"); } catch (error) { showToast(error.message, true); }
  await Promise.all([loadStatus(), loadDashboard()]);
  try {
    await refreshConversationList();
    if (state.conversations.length) await loadConversation(state.conversationId);
    else await startNewConversation();
    await refreshMemories();
  } catch (error) { showToast(error.message, true); }
}

init();
