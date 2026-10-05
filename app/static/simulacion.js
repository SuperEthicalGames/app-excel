const money = value => new Intl.NumberFormat("es-CO", {style:"currency",currency:"COP",maximumFractionDigits:0}).format(Number(value)||0);
const steps = ["Empresa", "Documentos", "Extracción", "Validación", "Registro", "Reportes", "Excel listo"];
let manifest;
let currentStep = -1;
const $ = selector => document.querySelector(selector);

async function loadManifest() {
  const response = await fetch("/demo/manifest.json", {cache:"no-store"});
  if (!response.ok) throw new Error("No se pudo cargar la simulación.");
  manifest = await response.json();
  render();
}

function documentTable(docs) {
  return `<div class="demo-table-wrap"><table class="demo-table"><thead><tr><th>DOCUMENTO</th><th>TERCERO</th><th>FORMATO</th><th>ESTADO</th></tr></thead><tbody>${docs.map(d=>`<tr><td>${d["Numero"]}</td><td>${d["Tercero"]}</td><td>${d["TipoArchivo"]}</td><td><span class="state ${d["Estado"] === "Revision Manual" ? "pending" : ""}">${d["Estado"] === "Revision Manual" ? "Revisión manual" : d["Estado"]}</span></td></tr>`).join("")}</tbody></table></div>`;
}

function movementTable(rows) {
  return `<div class="demo-table-wrap"><table class="demo-table"><thead><tr><th>FECHA</th><th>TIPO</th><th>CONCEPTO</th><th>CUENTA</th><th>VALOR</th><th>ESTADO</th></tr></thead><tbody>${rows.map(m=>{const [year,month,day]=m.Fecha.split("-").map(Number);return `<tr><td>${new Date(year,month-1,day).toLocaleDateString("es-CO")}</td><td>${m.Tipo}</td><td>${m.Concepto}</td><td>${m.Cuenta}</td><td class="amount">${money(m.Subtotal)}</td><td><span class="state ${m.Estado === "Pendiente" ? "pending" : ""}">${m.Estado}</span></td></tr>`}).join("")}</tbody></table></div>`;
}

function bodyForStep(index) {
  const c = manifest.counts, t = manifest.totals;
  if (index === 0) return `<div class="step-kicker">ETAPA 1 · CONFIGURACIÓN</div><h3>La empresa define sus propios catálogos</h3><p class="step-lead">La plantilla se personaliza desde Excel. La estructura y las fórmulas no contienen porcentajes tributarios universales.</p><div class="data-grid"><div class="info-card"><small>EMPRESA</small><strong>${manifest.company}</strong><p>NIT ficticio · ejercicio de demostración</p></div><div class="info-card"><small>CUENTAS ACTIVAS</small><strong>Caja General · Banco Principal</strong><p>También están preparadas las cuentas por cobrar y pagar.</p></div><div class="info-card"><small>CATÁLOGOS</small><strong>Ingresos, operación, activos</strong><p>Categorías e impuestos permanecen editables.</p></div></div>`;
  if (index === 1) return `<div class="step-kicker">ETAPA 2 · RECEPCIÓN DOCUMENTAL</div><h3>Entran XML y soportes de ejemplo</h3><p class="step-lead">Se muestra una factura XML estructurada y un recibo en imagen. La simulación no lee archivos reales ni activa OCR; presenta resultados ficticios para explicar el flujo.</p>${documentTable(manifest.documents)}<div class="completion-alert">${c.documents} soportes simulados · ${c.validated_documents} listos para relacionar · ${c.manual_review_documents} retenido para revisión humana.</div>`;
  if (index === 2) {
    const d=manifest.documents[0], low=manifest.documents[manifest.documents.length-1];
    return `<div class="step-kicker">ETAPA 3 · EXTRACCIÓN Y NORMALIZACIÓN</div><h3>La fuente estructurada conserva los datos exactos</h3><p class="step-lead">Para el XML se extraen los campos del ejemplo sin OCR. El recibo de imagen ilustra el caso en que la confianza es menor.</p><div class="data-grid"><div class="info-card"><small>FACTURA XML · ${d.Numero}</small><strong>${d.Tercero}</strong><p>Fecha: 08/09/2026 · Identificación: ficticia</p></div><div class="info-card"><small>VALORES EXTRAÍDOS</small><strong>Base ${money(d.Subtotal)} · total ${money(d.Total)}</strong><p>Impuestos y retenciones: $0 en este ejemplo. No es una liquidación fiscal.</p></div><div class="info-card"><small>RESULTADO ESTRUCTURADO</small><span class="status-check"><i></i> XML · confianza 100%</span><p>Referencia, tercero, fecha e importes quedan vinculados al soporte.</p></div></div><div class="ocr-card"><div><strong>Recibo ${low.Numero} · lectura de imagen ilustrativa</strong><p>El sistema detecta baja confianza en fecha y tercero, por eso no lo contabiliza automáticamente.</p></div><div class="ocr-score">64%<small>confianza</small></div></div>`;
  }
  if (index === 3) return `<div class="step-kicker">ETAPA 4 · VALIDACIÓN</div><h3>Los controles separan lo correcto de lo que requiere atención</h3><p class="step-lead">Las revisiones no borran ni contabilizan automáticamente documentos dudosos; explican el resultado y los dejan para revisión.</p><div class="validation-list"><div class="validation-row"><span class="check">✓</span><div><strong>Fecha y número presentes</strong><small>Documentos XML del ejemplo</small></div></div><div class="validation-row"><span class="check">✓</span><div><strong>Subtotal, impuestos y total concilian</strong><small>En los ejemplos XML; impuestos en cero</small></div></div><div class="validation-row"><span class="check">✓</span><div><strong>Posibles duplicados comprobados</strong><small>Sin coincidencias exactas en la muestra</small></div></div><div class="validation-row"><span class="check warn">!</span><div><strong>${manifest.documents.at(-1).Numero}: revisión manual</strong><small>Confianza 64%; revisar antes de registrarlo</small></div></div></div>`;
  if (index === 4) return `<div class="step-kicker">ETAPA 5 · REGISTRO FINANCIERO</div><h3>Movimientos, saldos y soportes quedan relacionados</h3><p class="step-lead">Se crean cinco movimientos de ejemplo. Los estados pagados alimentan el flujo de caja; los pendientes aparecen en cartera.</p>${movementTable(manifest.movements)}<div class="report-strip"><span class="report-tag"><b>${c.third_parties}</b> terceros únicos</span><span class="report-tag"><b>${c.payments}</b> pagos aplicados</span><span class="report-tag"><b>${c.collections}</b> cobro aplicado</span><span class="report-tag">Cada operación conserva su ID de soporte</span></div>`;
  if (index === 5) return `<div class="step-kicker">ETAPA 6 · DASHBOARD Y REPORTES</div><h3>El libro resume los movimientos para septiembre de 2026</h3><p class="step-lead">Los indicadores de la simulación corresponden a las filas insertadas. Al abrir el XLSX en Excel, sus fórmulas se recalculan con el motor de Excel.</p><div class="money-cards"><div class="money-card"><span>Ingresos pagados</span><strong>${money(t.income_paid)}</strong></div><div class="money-card"><span>Gastos pagados</span><strong>${money(t.expenses_paid)}</strong></div><div class="money-card net"><span>Flujo neto</span><strong>${money(t.net_cash_flow)}</strong></div><div class="money-card"><span>Por cobrar</span><strong>${money(t.receivable)}</strong></div><div class="money-card"><span>Por pagar</span><strong>${money(t.payable)}</strong></div></div><div class="report-strip"><span class="report-tag"><b>${c.documents}</b> documentos · ${c.validated_documents} validados · ${c.manual_review_documents} revisión</span><span class="report-tag">Dashboard</span><span class="report-tag">Flujo de caja</span><span class="report-tag">Cuentas por cobrar/pagar</span><span class="report-tag">Reporte mensual y anual</span></div>`;
  return `<div class="step-kicker">ETAPA 7 · ARCHIVO GENERADO</div><h3>El libro de demostración está listo</h3><p class="step-lead">Se generó una copia independiente de la plantilla con sus 28 hojas, fórmulas, catálogos, datos ficticios y auditoría. Ábrela en Excel para recalcular fórmulas y recorrer las pestañas.</p><div class="final-card"><div class="final-icon">▦</div><div class="final-copy"><h3>${manifest.workbook}</h3><p>${manifest.company} · periodo ${manifest.period}. Incluye movimientos, soportes, pagos, cobros y reportes.</p></div><a class="download-button" href="${manifest.download}" download><span>↓</span> Descargar Excel</a></div><div class="file-meta"><span>3 terceros ficticios</span><span>${c.documents} documentos</span><span>${c.movements} movimientos</span><span>2 pagos y 1 cobro</span><span>Datos tributarios configurables</span></div><div class="completion-alert">${manifest.notice}</div>`;
}

function renderNav() {
  $("#stepNav").innerHTML = steps.map((step,index)=>`<button class="step-tab ${index===currentStep?"active":""} ${index<currentStep?"done":""}" data-step="${index}" ${index>Math.max(0,currentStep+1)?"disabled":""}><span class="step-num">${index<currentStep?"✓":index+1}</span><span class="step-name">${step}</span></button>`).join("");
  document.querySelectorAll(".step-tab").forEach(button=>button.addEventListener("click",()=>{const next=Number(button.dataset.step);if(next<=currentStep+1){currentStep=next;render();}}));
}

function render() {
  renderNav();
  const started=currentStep>=0, done=currentStep===steps.length-1;
  $("#stepCounter").textContent=started?`ETAPA ${currentStep+1} DE ${steps.length}`:"LISTO PARA INICIAR";
  $("#progressBar").style.width=started?`${((currentStep+1)/steps.length)*100}%`:"0%";
  $("#stepTitle").textContent=started?steps[currentStep]:"El recorrido, de principio a fin";
  $("#stepContent").innerHTML=started?bodyForStep(currentStep):`<div class="welcome-card"><div class="welcome-icon">✦</div><div><strong>Usaremos una empresa y documentos completamente ficticios.</strong><p>La simulación recorre configuración, lectura de soportes, controles, registro de operaciones y reportes. Al final tendrás un archivo Excel basado en el libro del sistema.</p></div></div><div class="start-summary"><span><b>3</b> terceros</span><span><b>6</b> documentos</span><span><b>5</b> movimientos</span><span><b>1</b> documento enviado a revisión</span></div>`;
  $("#prevButton").disabled=!started||currentStep===0;
  $("#nextButton").disabled=done;
  $("#nextButton").innerHTML=done?"Simulación completada <span>✓</span>":started?`Siguiente etapa <span>→</span>`:`Comenzar <span>→</span>`;
  $("#walkNote").textContent=done?"Libro separado, creado desde la plantilla y listo para abrir en Excel.":started?`Etapa ${currentStep+1}: ${steps[currentStep]}. Avanza para continuar el recorrido.`:"Puedes avanzar etapa por etapa y volver atrás cuando quieras.";
  $("#startButton").textContent=started?"Reiniciar recorrido":"Iniciar simulación";
}

$("#startButton").addEventListener("click",()=>{currentStep=currentStep<0?0:-1;render();document.querySelector(".walkthrough").scrollIntoView({behavior:"smooth",block:"start"});});
$("#prevButton").addEventListener("click",()=>{if(currentStep>0){currentStep--;render();}});
$("#nextButton").addEventListener("click",()=>{if(currentStep<steps.length-1){currentStep++;render();}});
loadManifest().catch(error=>{document.querySelector(".step-content").textContent=error.message;});
