const defaults = {
  fechaSolicitud: "",
  cliente: "",
  responsable: "",
  consecutivo: "",
  fase: "100000000",
  hardware: "",
  obsequios: "",
  margenObjetivo: "",
  iva: 19,
  comision: 15,
  impuestos: 3.5,
  fletes: "",
  imprevistos: 0.02,
  plazoCredito: "",
  baseFinanciacion: 1.1
};

const moneyFields = new Set([
  "hardware",
  "obsequios",
  "totalCostosDirectos",
  "utilidadEsperada",
  "valorVenta",
  "ivaValor",
  "valorFacturar",
  "comisionValor",
  "impuestosValor",
  "fletes",
  "imprevistosValor",
  "financiacionValor",
  "totalCostosInternos",
  "utilidadBruta"
]);

const percentFields = new Set([
  "margenObjetivo",
  "iva",
  "comision",
  "impuestos",
  "imprevistos",
  "baseFinanciacion"
]);

const ratioFields = new Set([
  "financiacionPorcentaje",
  "margenSobreCostos",
  "margenSobreVenta"
]);

const storageKey = "matriz-comercial-px-empty-inputs";
const fields = {};
const textFields = ["fechaSolicitud", "cliente", "responsable", "fase"];
let msalApp = null;
let currentAccount = null;
let matrixEntitySetName = null;
let activeMatrixId = null;

const dataverseConfig = {
  environmentUrl: "https://db-px.crm2.dynamics.com",
  tenantId: "e6805558-f5bb-444c-8af2-5f3a4d6dd3fc",
  clientId: "ffcf61c2-18a2-4be4-a753-c81934026e4d",
  tableLogicalName: "px_matrizcomercial"
};

const estadoValues = {
  borrador: 100000000,
  enRevision: 100000001,
  aprobada: 100000002,
  rechazada: 100000003,
  ganada: 100000004,
  perdida: 100000005,
  cerrada: 100000006
};

const faseValues = {
  solicitudCredito: 100000000,
  oferta: 100000001,
  cierre: 100000006
};

const faseLabels = {
  [faseValues.solicitudCredito]: "Solicitud de credito",
  [faseValues.oferta]: "Oferta",
  [faseValues.cierre]: "Cierre"
};

const faseSteps = [
  { value: String(faseValues.solicitudCredito), label: "Solicitud de credito", icon: "clipboard-list" },
  { value: String(faseValues.oferta), label: "Oferta", icon: "file-check-2" },
  { value: String(faseValues.cierre), label: "Cierre", icon: "handshake" }
];

const currency = new Intl.NumberFormat("es-CO", {
  style: "currency",
  currency: "COP",
  maximumFractionDigits: 0
});

const percent = new Intl.NumberFormat("es-CO", {
  style: "percent",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});

function readState() {
  try {
    return { ...defaults, ...JSON.parse(localStorage.getItem(storageKey)) };
  } catch {
    return { ...defaults };
  }
}

let state = readState();

function asRate(value) {
  return Number(value || 0) / 100;
}

function getAuthRedirectUri() {
  const origin = window.location.origin || "";
  let path = window.location.pathname || "/";
  path = path.split("?")[0].split("#")[0];
  if (path.endsWith("/index.html")) path = path.slice(0, -"/index.html".length) || "/";
  if (!path.endsWith("/")) path += "/";
  return origin + path;
}

function initMsal() {
  if (!window.msal) {
    throw new Error("MSAL no esta disponible");
  }

  if (msalApp) return msalApp;

  msalApp = new window.msal.PublicClientApplication({
    auth: {
      clientId: dataverseConfig.clientId,
      authority: `https://login.microsoftonline.com/${dataverseConfig.tenantId}`,
      redirectUri: getAuthRedirectUri()
    },
    cache: { cacheLocation: "sessionStorage" }
  });

  return msalApp;
}

function loadMsal() {
  if (window.msal) return Promise.resolve();

  const sources = [
    "https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js",
    "https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.3/lib/msal-browser.min.js"
  ];

  return new Promise((resolve, reject) => {
    const loadSource = (index) => {
      if (!sources[index]) {
        reject(new Error("No se pudo cargar la libreria de autenticacion Microsoft"));
        return;
      }

      const script = document.createElement("script");
      script.src = sources[index];
      script.onload = () => {
        if (window.msal) {
          resolve();
          return;
        }
        loadSource(index + 1);
      };
      script.onerror = () => loadSource(index + 1);
      document.head.appendChild(script);
    };

    loadSource(0);
  });
}

function updateConnectionLabel(text) {
  const label = document.querySelector("#loginBtn span");
  if (label) label.textContent = text;
}

function updateAuthStatus(text) {
  const status = document.querySelector("#authStatus");
  if (status) status.textContent = text;
}

function setAuthenticatedUi() {
  document.querySelector("#appShell")?.classList.remove("app-locked");
  document.querySelector("#authGate")?.classList.add("auth-hidden");
  document.body.classList.remove("auth-pending");
  document.body.classList.add("auth-ready");
  document.querySelector("#appShell")?.setAttribute("aria-hidden", "false");
}

async function getDataverseToken({ interactive = true } = {}) {
  await loadMsal();
  const app = initMsal();
  const scopes = [`${dataverseConfig.environmentUrl}/user_impersonation`];
  currentAccount = app.getAllAccounts()[0] || currentAccount;

  if (!currentAccount) {
    if (!interactive) throw new Error("Conecta con Microsoft para sincronizar matrices");
    const login = await app.loginPopup({ scopes });
    currentAccount = login.account;
  }

  try {
    const result = await app.acquireTokenSilent({ scopes, account: currentAccount });
    return result.accessToken;
  } catch {
    if (!interactive) throw new Error("La sesion necesita renovarse. Conecta de nuevo.");
    const result = await app.acquireTokenPopup({ scopes, account: currentAccount });
    currentAccount = result.account;
    return result.accessToken;
  }
}

async function connectMicrosoft() {
  const authButton = document.querySelector("#authConnectBtn");
  if (authButton) authButton.disabled = true;
  updateAuthStatus("Conectando con Microsoft...");
  updateConnectionLabel("Conectando...");
  try {
    await getDataverseToken();
    if (currentAccount) {
      const name = currentAccount.name || currentAccount.username || "Conectado";
      updateConnectionLabel(name);
      updateAuthStatus(`Conectado como ${name}`);
    }
    setAuthenticatedUi();
    await loadMatrices();
  } catch (error) {
    updateAuthStatus(error.message);
    updateConnectionLabel("Conectar");
  } finally {
    if (authButton) authButton.disabled = false;
  }
}

async function dataverseRequest(method, relativeUrl, body, extraHeaders = {}, options = {}) {
  const token = await getDataverseToken(options);
  const url = new URL(relativeUrl, `${dataverseConfig.environmentUrl}/api/data/v9.2/`);
  const response = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      "Content-Type": "application/json; charset=utf-8",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      ...extraHeaders
    },
    body: body ? JSON.stringify(body) : undefined
  });

  if (response.status === 204) return null;

  const text = await response.text();
  const data = text ? JSON.parse(text) : null;

  if (!response.ok) {
    throw new Error(data?.error?.message || "Error consultando Dataverse");
  }

  return data;
}

async function getMatrixEntitySetName(options = {}) {
  if (matrixEntitySetName) return matrixEntitySetName;

  const data = await dataverseRequest(
    "GET",
    `EntityDefinitions(LogicalName='${dataverseConfig.tableLogicalName}')?$select=EntitySetName`,
    null,
    {},
    options
  );
  matrixEntitySetName = data.EntitySetName;
  return matrixEntitySetName;
}

function calculate() {
  const totalCostosDirectos =
    Number(state.hardware) +
    Number(state.obsequios);
  const margenObjetivo = asRate(state.margenObjetivo);
  const valorVenta = margenObjetivo >= 1 ? 0 : totalCostosDirectos / (1 - margenObjetivo);
  const utilidadEsperada = valorVenta - totalCostosDirectos;
  const ivaValor = valorVenta * asRate(state.iva);
  const valorFacturar = valorVenta + ivaValor;
  const comisionValor = asRate(state.comision) * utilidadEsperada;
  const impuestosValor = valorVenta * asRate(state.impuestos);
  const imprevistosValor = valorVenta * asRate(state.imprevistos);
  const financiacionPorcentaje =
    Number(state.plazoCredito) > 60
      ? ((Number(state.plazoCredito) - 60) / 30) * asRate(state.baseFinanciacion)
      : 0;
  const financiacionValor = financiacionPorcentaje * valorVenta;
  const totalCostosInternos =
    financiacionValor + imprevistosValor + Number(state.fletes) + impuestosValor + comisionValor;
  const utilidadBruta = utilidadEsperada - totalCostosInternos;
  const margenSobreCostos = totalCostosDirectos ? utilidadBruta / totalCostosDirectos : 0;
  const margenSobreVenta = valorVenta ? utilidadBruta / valorVenta : 0;

  return {
    ...state,
    totalCostosDirectos,
    utilidadEsperada,
    valorVenta,
    ivaValor,
    valorFacturar,
    comisionValor,
    impuestosValor,
    imprevistosValor,
    financiacionPorcentaje,
    financiacionValor,
    totalCostosInternos,
    utilidadBruta,
    margenSobreCostos,
    margenSobreVenta
  };
}

function formatValue(key, value) {
  if (moneyFields.has(key)) return currency.format(Number(value || 0));
  if (percentFields.has(key)) return percent.format(asRate(value));
  if (ratioFields.has(key)) return percent.format(Number(value || 0));
  return new Intl.NumberFormat("es-CO").format(Number(value || 0));
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#039;"
  }[char]));
}

function saveState() {
  localStorage.setItem(storageKey, JSON.stringify(state));
  const saveStateNode = document.querySelector("#saveState");
  saveStateNode.textContent = "Cambios guardados";
  window.clearTimeout(saveState.timer);
  saveState.timer = window.setTimeout(() => {
    saveStateNode.textContent = "Listo para evaluar";
  }, 1300);
}

function render() {
  const result = calculate();
  document.querySelectorAll("[data-out]").forEach((cell) => {
    const key = cell.dataset.out;
    cell.textContent = formatValue(key, result[key]);
  });

  document.querySelector("#kpiCostosDirectos").textContent = currency.format(result.totalCostosDirectos);
  document.querySelector("#kpiVenta").textContent = currency.format(result.valorVenta);
  document.querySelector("#kpiFacturar").textContent = currency.format(result.valorFacturar);
  document.querySelector("#kpiUtilidad").textContent = currency.format(result.utilidadBruta);
  document.querySelector(".metric.profit").classList.toggle("negative", result.utilidadBruta < 0);
  renderPhaseProgress();
  updateSaveButtonMode();
}

function renderPhaseProgress() {
  const progress = document.querySelector("#phaseProgress");
  if (!progress) return;

  const activeIndex = Math.max(0, faseSteps.findIndex((step) => step.value === String(state.fase)));
  progress.innerHTML = faseSteps.map((step, index) => {
    const status = index < activeIndex ? "done" : index === activeIndex ? "current" : "pending";
    return `
      <li class="phase-step ${status}">
        <span class="phase-icon"><i data-lucide="${step.icon}"></i></span>
        <span class="phase-line" aria-hidden="true"></span>
        <span class="phase-dot">${index + 1}</span>
        <span class="phase-label">${escapeHtml(step.label)}</span>
      </li>
    `;
  }).join("");

  if (window.lucide) window.lucide.createIcons();
}

function updateSaveButtonMode() {
  const label = document.querySelector("#saveMatrixBtn span");
  if (!label) return;
  label.textContent = activeMatrixId ? "Actualizar matriz" : "Guardar matriz";
}

function bindInputs() {
  document.querySelectorAll("[data-field]").forEach((input) => {
    fields[input.dataset.field] = input;
    input.value = state[input.dataset.field];
    input.addEventListener("input", () => {
      state[input.dataset.field] = input.value === "" ? "" : Number(input.value);
      render();
      saveState();
    });
  });

  textFields.forEach((id) => {
    const input = document.querySelector(`#${id}`);
    input.value = state[id] ?? "";
    input.addEventListener("input", () => {
      state[id] = input.value;
      if (id === "fase") renderPhaseProgress();
      saveState();
    });
  });
}

function syncInputsFromState() {
  Object.entries(fields).forEach(([key, input]) => {
    input.value = state[key] ?? "";
  });
  textFields.forEach((id) => {
    const input = document.querySelector(`#${id}`);
    if (input) input.value = state[id] ?? "";
  });
  render();
}

function resetValues() {
  state = { ...defaults };
  activeMatrixId = null;
  syncInputsFromState();
  saveState();
}

function switchView(target) {
  const isEvaluation = target === "evaluationView";
  document.querySelector("#recordsView").classList.toggle("active", !isEvaluation);
  document.querySelector("#evaluationView").classList.toggle("active", isEvaluation);
  document.querySelector("#evaluationProcess").classList.toggle("active", isEvaluation);
  document.querySelector("#evaluationViewContent").classList.toggle("active", isEvaluation);
  document.body.classList.toggle("app-mode-records", !isEvaluation);
  document.body.classList.toggle("app-mode-evaluation", isEvaluation);

  document.querySelectorAll("[data-view-target]").forEach((button) => {
    button.classList.toggle("active", button.dataset.viewTarget === target);
  });
}

function startNewMatrix() {
  resetValues();
  switchView("evaluationView");
}

function loadExcelJs() {
  if (window.ExcelJS) return Promise.resolve();

  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js";
    script.onload = () => window.ExcelJS ? resolve() : reject(new Error("No se pudo cargar ExcelJS"));
    script.onerror = () => reject(new Error("No se pudo cargar ExcelJS"));
    document.head.appendChild(script);
  });
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function safeFileName(value) {
  return String(value || "provexpress")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[\\/:*?"<>|%]+/g, "-")
    .replace(/\s+/g, "-")
    .toLowerCase();
}

function loadExternalScript(src, globalCheck) {
  if (globalCheck()) return Promise.resolve();

  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = src;
    script.onload = () => globalCheck() ? resolve() : reject(new Error(`No se pudo cargar ${src}`));
    script.onerror = () => reject(new Error(`No se pudo cargar ${src}`));
    document.head.appendChild(script);
  });
}

async function loadPdfLibs() {
  await loadExternalScript(
    "https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js",
    () => Boolean(window.jspdf?.jsPDF)
  );
  await loadExternalScript(
    "https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.2/dist/jspdf.plugin.autotable.min.js",
    () => Boolean(window.jspdf?.jsPDF?.API?.autoTable)
  );
}

function pdfMoney(value) {
  return currency.format(Number(value || 0));
}

function pdfPercent(value) {
  return percent.format(Number(value || 0));
}

async function exportPdf() {
  const button = document.querySelector("#printBtn");
  const label = button?.querySelector("span");
  const previousText = label?.textContent || "PDF";
  if (button) button.disabled = true;
  if (label) label.textContent = "Generando...";

  try {
    await loadPdfLibs();
    const result = calculate();
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: "portrait", unit: "pt", format: "letter" });
    const pageWidth = doc.internal.pageSize.getWidth();
    const margin = 28;
    const purple = [112, 48, 160];
    const ink = [23, 32, 51];
    const muted = [106, 114, 131];
    const line = [223, 229, 239];
    const soft = [247, 249, 252];
    const paleBlue = [237, 243, 250];
    const palePurple = [242, 231, 250];

    doc.setProperties({
      title: `Evaluacion ${state.consecutivo || state.cliente || "Provexpress"}`,
      subject: "Matriz Comercial PX",
      author: "Provexpress"
    });

    doc.setFont("helvetica", "bold");
    doc.setTextColor(...purple);
    doc.setFontSize(9.5);
    doc.text("PROVEXPRESS - MATRIZ COMERCIAL PX", pageWidth / 2, 28, { align: "center" });
    doc.setTextColor(...ink);
    doc.setFontSize(19);
    doc.text("Evaluacion de proyectos", pageWidth / 2, 50, { align: "center" });
    doc.setDrawColor(...purple);
    doc.setLineWidth(2);
    doc.line(margin, 62, pageWidth - margin, 62);

    doc.autoTable({
      startY: 72,
      margin: { left: margin, right: margin },
      theme: "grid",
      tableWidth: pageWidth - margin * 2,
      styles: {
        font: "helvetica",
        fontSize: 7.2,
        cellPadding: 3.5,
        lineColor: line,
        lineWidth: 0.5,
        textColor: ink,
        minCellHeight: 12
      },
      body: [
        ["Consecutivo", state.consecutivo || "Sin asignar", "Fecha solicitud", state.fechaSolicitud || ""],
        ["Cliente", state.cliente || "", "Responsable", state.responsable || ""],
        ["Fase del producto", faseLabels[state.fase] || "", "", ""]
      ],
      columnStyles: {
        0: { fontStyle: "bold", fillColor: paleBlue, cellWidth: 95 },
        1: { fontStyle: "bold", cellWidth: 175 },
        2: { fontStyle: "bold", fillColor: paleBlue, cellWidth: 95 },
        3: { fontStyle: "bold", cellWidth: 175 }
      }
    });

    const kpiY = doc.lastAutoTable.finalY + 10;
    const kpiGap = 7;
    const kpiWidth = (pageWidth - margin * 2 - kpiGap * 3) / 4;
    const kpis = [
      ["Total costos directos", pdfMoney(result.totalCostosDirectos)],
      ["Valor venta antes de IVA", pdfMoney(result.valorVenta)],
      ["Valor a facturar", pdfMoney(result.valorFacturar)],
      ["Utilidad bruta", pdfMoney(result.utilidadBruta)]
    ];

    kpis.forEach(([title, value], index) => {
      const x = margin + index * (kpiWidth + kpiGap);
      doc.setFillColor(...soft);
      doc.setDrawColor(...line);
      doc.setLineWidth(0.8);
      doc.roundedRect(x, kpiY, kpiWidth, 38, 2, 2, "FD");
      doc.setFont("helvetica", "bold");
      doc.setFontSize(6.7);
      doc.setTextColor(...purple);
      doc.text(title.toUpperCase(), x + 6, kpiY + 12, { maxWidth: kpiWidth - 12 });
      doc.setFontSize(10.5);
      doc.setTextColor(...ink);
      doc.text(value, x + 6, kpiY + 30, { maxWidth: kpiWidth - 12 });
    });

    const processY = kpiY + 60;
    const processBoxY = processY - 13;
    const processBoxHeight = 58;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(8.2);

    const currentPhaseIndex = Math.max(0, faseSteps.findIndex((step) => step.value === String(state.fase)));
    const startX = margin + 70;
    const endX = pageWidth - margin - 70;
    const stepGap = (endX - startX) / (faseSteps.length - 1);
    const trackY = processY + 13;
    const activeX = startX + currentPhaseIndex * stepGap;

    doc.setDrawColor(...line);
    doc.setFillColor(255, 255, 255);
    doc.roundedRect(margin, processBoxY, pageWidth - margin * 2, processBoxHeight, 3, 3, "S");
    doc.setFillColor(...palePurple);
    doc.rect(margin, processBoxY, pageWidth - margin * 2, 14, "F");
    doc.setDrawColor(...line);
    doc.line(margin, processBoxY + 14, pageWidth - margin, processBoxY + 14);
    doc.setTextColor(...purple);
    doc.text("PROCESO DEL PRODUCTO", margin + 8, processBoxY + 10);

    doc.setDrawColor(...line);
    doc.setLineWidth(1.2);
    doc.line(startX, trackY, endX, trackY);
    doc.setDrawColor(...purple);
    doc.setLineWidth(2);
    doc.line(startX, trackY, activeX, trackY);

    faseSteps.forEach((step, index) => {
      const x = startX + index * stepGap;
      const active = index <= currentPhaseIndex;

      doc.setDrawColor(...(active ? purple : line));
      doc.setLineWidth(active ? 1.6 : 1.1);
      doc.setFillColor(...(active ? purple : [255, 255, 255]));
      doc.circle(x, trackY, 9.5, "FD");

      doc.setFont("helvetica", "bold");
      doc.setFontSize(7.6);
      doc.setTextColor(...(active ? [255, 255, 255] : muted));
      doc.text(String(index + 1), x, trackY + 2.7, { align: "center" });

      doc.setFontSize(6.7);
      doc.setTextColor(...(index === currentPhaseIndex ? purple : ink));
      doc.text(step.label.toUpperCase(), x, trackY + 25, { align: "center", maxWidth: 118 });
    });

    const tableRows = [
      { type: "section", concept: "Costos Directos del Proyecto", value: "" },
      { concept: "Hardware", value: pdfMoney(state.hardware), kind: "money" },
      { concept: "Obsequios", value: pdfMoney(state.obsequios), kind: "money" },
      { type: "total", concept: "Total Costos Directos del Proyecto", value: pdfMoney(result.totalCostosDirectos) },
      { concept: "% Margen Objetivo", value: pdfPercent(asRate(state.margenObjetivo)) },
      { type: "accent", concept: "Utilidad antes de gastos internos", value: pdfMoney(result.utilidadEsperada) },
      { type: "total", concept: "Valor de venta del Proyecto Antes de IVA", value: pdfMoney(result.valorVenta) },
      { concept: "% IVA aplicable", value: pdfPercent(asRate(state.iva)) },
      { concept: "IVA", value: pdfMoney(result.ivaValor) },
      { type: "total", concept: "Valor a Facturar", value: pdfMoney(result.valorFacturar) },
      { type: "section", concept: "Costos Internos", value: "" },
      { concept: "% Comision de Ventas del comercial", value: pdfPercent(asRate(state.comision)) },
      { concept: "Comision estimada del Comercial incluida carga prestacional", value: pdfMoney(result.comisionValor) },
      { concept: "% Impuestos estimados de evaluacion", value: pdfPercent(asRate(state.impuestos)) },
      { concept: "Valor impuestos estimados sobre valor de venta", value: pdfMoney(result.impuestosValor) },
      { concept: "Costo fletes", value: pdfMoney(state.fletes) },
      { concept: "% Imprevistos estimados", value: pdfPercent(asRate(state.imprevistos)) },
      { concept: "Valor Imprevistos estimados", value: pdfMoney(result.imprevistosValor) },
      { concept: "Plazo de Credito al cliente en dias", value: String(Number(state.plazoCredito || 0)) },
      { concept: "% Base de financiacion mensual", value: pdfPercent(asRate(state.baseFinanciacion)) },
      { concept: "% de financiacion ligado al plazo", value: pdfPercent(result.financiacionPorcentaje) },
      { concept: "Costo de financiacion para evaluacion", value: pdfMoney(result.financiacionValor) },
      { type: "total", concept: "Total Costos internos del proyecto", value: pdfMoney(result.totalCostosInternos) },
      { type: "accent", concept: "Utilidad Bruta Provexpress", value: pdfMoney(result.utilidadBruta) },
      { concept: "% Margen Bruto Provexpress sobre costos", value: pdfPercent(result.margenSobreCostos) },
      { type: "accent", concept: "% Margen Bruto Provexpress sobre Valor de Venta", value: pdfPercent(result.margenSobreVenta) }
    ];

    doc.autoTable({
      startY: processY + 50,
      margin: { left: margin, right: margin, bottom: 18 },
      tableWidth: pageWidth - margin * 2,
      theme: "grid",
      columns: [
        { header: "Concepto", dataKey: "concept" },
        { header: "Valor", dataKey: "value" }
      ],
      body: tableRows,
      styles: {
        font: "helvetica",
        fontSize: 6.1,
        cellPadding: 2,
        lineColor: line,
        lineWidth: 0.4,
        textColor: ink,
        minCellHeight: 8.2,
        overflow: "linebreak"
      },
      headStyles: {
        fillColor: purple,
        textColor: [255, 255, 255],
        halign: "center",
        fontStyle: "bold",
        fontSize: 7
      },
      columnStyles: {
        concept: { cellWidth: 390 },
        value: { cellWidth: 150, halign: "right", fontStyle: "bold" }
      },
      didParseCell(data) {
        const row = data.row.raw;
        if (data.section !== "body" || !row) return;
        if (row.type === "section") {
          data.cell.styles.fillColor = paleBlue;
          data.cell.styles.textColor = purple;
          data.cell.styles.fontStyle = "bold";
          data.cell.styles.halign = "center";
          if (data.column.dataKey === "value") data.cell.text = [""];
        }
        if (row.type === "total") {
          data.cell.styles.fillColor = soft;
          data.cell.styles.fontStyle = "bold";
        }
        if (row.type === "accent") {
          data.cell.styles.fillColor = purple;
          data.cell.styles.textColor = [255, 255, 255];
          data.cell.styles.fontStyle = "bold";
        }
      }
    });

    doc.save(`evaluacion-proyecto-${safeFileName(state.consecutivo || state.cliente || "provexpress")}.pdf`);
  } catch (error) {
    document.querySelector("#saveState").textContent = error.message;
  } finally {
    if (button) button.disabled = false;
    if (label) label.textContent = previousText;
  }
}

async function exportTable() {
  if (!window.ExcelJS) await loadExcelJs();

  const result = calculate();
  const currentPhaseIndex = Math.max(0, faseSteps.findIndex((step) => step.value === String(state.fase)));

  const workbook = new window.ExcelJS.Workbook();
  workbook.creator = "Provexpress";
  workbook.created = new Date();
  const sheet = workbook.addWorksheet("Evaluacion", {
    views: [{ showGridLines: false }]
  });

  const colors = {
    ink: "172033",
    muted: "6A7283",
    line: "DFE5EF",
    purple: "7030A0",
    blue: "0070C0",
    palePurple: "F2E7FA",
    paleBlue: "EDF3FA",
    paper: "FFFFFF",
    soft: "F7F9FC",
    success: "16815A",
    danger: "BD2F45"
  };

  sheet.columns = [
    { width: 34 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 18 }
  ];

  const border = {
    top: { style: "thin", color: { argb: colors.line } },
    left: { style: "thin", color: { argb: colors.line } },
    bottom: { style: "thin", color: { argb: colors.line } },
    right: { style: "thin", color: { argb: colors.line } }
  };
  const titleBorder = {
    top: { style: "thin", color: { argb: colors.purple } },
    left: { style: "thin", color: { argb: colors.purple } },
    bottom: { style: "medium", color: { argb: colors.purple } },
    right: { style: "thin", color: { argb: colors.purple } }
  };
  const moneyFormat = '"$"#,##0';
  const percentFormat = "0.00%";

  function mergeRow(rowNumber, from = 1, to = 6) {
    sheet.mergeCells(rowNumber, from, rowNumber, to);
    return sheet.getCell(rowNumber, from);
  }

  function sectionTitle(rowNumber, text) {
    const cell = mergeRow(rowNumber);
    cell.value = text.toUpperCase();
    cell.font = { bold: true, color: { argb: "FFFFFF" }, size: 13 };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.purple } };
    cell.border = titleBorder;
    sheet.getRow(rowNumber).height = 24;
  }

  function styleRange(rowNumber, from = 1, to = 6, fill = colors.paper) {
    for (let col = from; col <= to; col += 1) {
      const cell = sheet.getCell(rowNumber, col);
      cell.border = border;
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fill } };
      cell.alignment = { vertical: "middle", wrapText: true };
    }
  }

  sheet.mergeCells("A1:F1");
  sheet.getCell("A1").value = "PROVEXPRESS - MATRIZ COMERCIAL PX";
  sheet.getCell("A1").font = { bold: true, color: { argb: colors.purple }, size: 12 };
  sheet.getCell("A1").alignment = { horizontal: "center" };
  sheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.palePurple } };
  sheet.getCell("A1").border = titleBorder;
  sheet.getRow(1).height = 20;

  sheet.mergeCells("A2:F2");
  sheet.getCell("A2").value = "Evaluacion de proyectos";
  sheet.getCell("A2").font = { bold: true, color: { argb: colors.ink }, size: 24 };
  sheet.getCell("A2").alignment = { horizontal: "center" };
  sheet.getCell("A2").border = titleBorder;
  sheet.getRow(2).height = 34;

  sectionTitle(4, "Informacion del proyecto");
  [
    ["Consecutivo", state.consecutivo || "Sin asignar"],
    ["Fecha solicitud", state.fechaSolicitud || ""],
    ["Cliente", state.cliente || ""],
    ["Responsable", state.responsable || ""],
    ["Fase del producto", faseLabels[state.fase] || ""]
  ].forEach(([label, value], index) => {
    const row = 5 + index;
    sheet.mergeCells(row, 2, row, 6);
    sheet.getCell(row, 1).value = label;
    sheet.getCell(row, 2).value = value;
    sheet.getCell(row, 1).font = { bold: true, color: { argb: colors.purple }, size: 10 };
    sheet.getCell(row, 2).font = { bold: true, color: { argb: colors.ink } };
    styleRange(row, 1, 6, index % 2 ? colors.paper : colors.soft);
  });

  sectionTitle(11, "Resumen financiero");
  const kpis = [
    ["Total costos directos", result.totalCostosDirectos],
    ["Valor venta antes de IVA", result.valorVenta],
    ["Valor a facturar", result.valorFacturar],
    ["Utilidad bruta", result.utilidadBruta]
  ];
  kpis.forEach(([label, value], index) => {
    const startCol = index === 0 ? 1 : index + 2;
    const cell = sheet.getCell(12, startCol);
    cell.value = label;
    cell.font = { bold: true, color: { argb: colors.purple }, size: 9.5 };
    const valueCell = sheet.getCell(13, startCol);
    valueCell.value = Number(value || 0);
    valueCell.numFmt = moneyFormat;
    valueCell.font = { bold: true, color: { argb: value < 0 ? colors.danger : colors.ink }, size: 13 };
    styleRange(12, startCol, startCol, colors.soft);
    styleRange(13, startCol, startCol, colors.soft);
  });

  sectionTitle(15, "Proceso del producto");
  faseSteps.forEach((step, index) => {
    const startCol = index * 2 + 1;
    sheet.mergeCells(16, startCol, 16, startCol + 1);
    sheet.mergeCells(17, startCol, 17, startCol + 1);
    const isActive = index <= currentPhaseIndex;
    const numberCell = sheet.getCell(16, startCol);
    numberCell.value = index + 1;
    numberCell.font = { bold: true, color: { argb: "FFFFFF" }, size: 12 };
    numberCell.alignment = { horizontal: "center", vertical: "middle" };
    numberCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: isActive ? colors.purple : colors.ink } };
    numberCell.border = border;
    const labelCell = sheet.getCell(17, startCol);
    labelCell.value = step.label;
    labelCell.font = { bold: true, color: { argb: isActive ? colors.purple : colors.ink }, size: 10.5 };
    labelCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    labelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: isActive ? "F7F0FC" : colors.soft } };
    labelCell.border = border;
  });

  sectionTitle(19, "Formato de evaluacion");
  sheet.getRow(20).values = ["Concepto", "Valor", "", "", "", ""];
  sheet.mergeCells("B20:F20");
  styleRange(20, 1, 6, colors.purple);
  sheet.getCell("A20").font = { bold: true, color: { argb: "FFFFFF" }, size: 12 };
  sheet.getCell("B20").font = { bold: true, color: { argb: "FFFFFF" }, size: 12 };
  sheet.getCell("A20").alignment = { horizontal: "center" };
  sheet.getCell("B20").alignment = { horizontal: "center" };
  sheet.getRow(20).height = 24;

  let rowNumber = 21;
  function addSection(text) {
    sheet.mergeCells(rowNumber, 1, rowNumber, 6);
    sheet.getCell(rowNumber, 1).value = text;
    sheet.getCell(rowNumber, 1).font = { bold: true, color: { argb: colors.purple }, size: 11 };
    sheet.getCell(rowNumber, 1).alignment = { horizontal: "center" };
    styleRange(rowNumber, 1, 6, colors.paleBlue);
    sheet.getRow(rowNumber).height = 21;
    rowNumber += 1;
  }

  function addLine(label, value, options = {}) {
    const row = sheet.getRow(rowNumber);
    sheet.mergeCells(rowNumber, 2, rowNumber, 6);
    const labelCell = sheet.getCell(rowNumber, 1);
    const valueCell = sheet.getCell(rowNumber, 2);
    labelCell.value = label;
    valueCell.value = value;
    valueCell.alignment = { horizontal: "right", vertical: "middle" };
    if (options.type === "money") valueCell.numFmt = moneyFormat;
    if (options.type === "percent") valueCell.numFmt = percentFormat;
    if (options.bold) {
      labelCell.font = { bold: true };
      valueCell.font = { bold: true };
    }
    if (options.note) labelCell.note = options.note;
    styleRange(rowNumber, 1, 6, options.fill || colors.paper);
    if (options.hidden) row.hidden = true;
    rowNumber += 1;
  }

  addSection("Costos Directos del Proyecto");
  addLine("Hardware", Number(state.hardware || 0), { type: "money" });
  addLine("Obsequios", Number(state.obsequios || 0), { type: "money" });
  addLine("Total Costos Directos del Proyecto", result.totalCostosDirectos, { type: "money", bold: true, fill: colors.soft });
  addLine("% Margen Objetivo", asRate(state.margenObjetivo), { type: "percent" });
  addLine("Utilidad antes de gastos internos", result.utilidadEsperada, { type: "money", bold: true, fill: "F7F0FC" });
  addLine("Valor de venta del Proyecto Antes de IVA", result.valorVenta, { type: "money", bold: true, fill: colors.soft });
  addLine("% IVA aplicable", asRate(state.iva), { type: "percent", hidden: true });
  addLine("IVA", result.ivaValor, { type: "money" });
  addLine("Valor a Facturar", result.valorFacturar, { type: "money", bold: true, fill: colors.soft });

  addSection("Costos Internos");
  addLine("% Comision de Ventas del comercial", asRate(state.comision), {
    type: "percent",
    hidden: true,
    note: "Porcentaje usado para estimar el costo comercial total sobre la utilidad esperada."
  });
  addLine("Comision estimada del Comercial incluida carga prestacional", result.comisionValor, { type: "money" });
  addLine("% Impuestos estimados de evaluacion", asRate(state.impuestos), {
    type: "percent",
    hidden: true,
    note: "Provision comercial para cubrir impuestos/transacciones como retenciones, ICA, GMF/4x1000 u otros costos tributarios indirectos."
  });
  addLine("Valor impuestos estimados sobre valor de venta", result.impuestosValor, { type: "money" });
  addLine("Costo fletes", Number(state.fletes || 0), { type: "money" });
  addLine("% Imprevistos estimados", asRate(state.imprevistos), {
    type: "percent",
    hidden: true,
    note: "Provision minima para cubrir variaciones menores o costos no previstos."
  });
  addLine("Valor Imprevistos estimados", result.imprevistosValor, { type: "money" });
  addLine("Plazo de Credito al cliente en dias", Number(state.plazoCredito || 0));
  addLine("% Base de financiacion mensual", asRate(state.baseFinanciacion), { type: "percent", hidden: true });
  addLine("% de financiacion ligado al plazo", result.financiacionPorcentaje, { type: "percent" });
  addLine("Costo de financiacion para evaluacion", result.financiacionValor, { type: "money" });
  addLine("Total Costos internos del proyecto", result.totalCostosInternos, { type: "money", bold: true, fill: colors.soft });
  addLine("Utilidad Bruta Provexpress", result.utilidadBruta, { type: "money", bold: true, fill: "F7F0FC" });
  addLine("% Margen Bruto Provexpress sobre costos", result.margenSobreCostos, { type: "percent" });
  addLine("% Margen Bruto Provexpress sobre Valor de Venta", result.margenSobreVenta, { type: "percent", bold: true, fill: "F7F0FC" });

  sheet.eachRow((row) => {
    row.eachCell((cell) => {
      cell.border = cell.border || border;
      cell.alignment = { vertical: "middle", wrapText: true, ...cell.alignment };
    });
  });

  const buffer = await workbook.xlsx.writeBuffer();
  downloadBlob(
    new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    `evaluacion-proyecto-${safeFileName(state.consecutivo || state.cliente || "provexpress")}.xlsx`
  );
}

function renderMatrices(rows) {
  const tbody = document.querySelector("#matricesList");
  if (!tbody) return;

  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="7">No hay matrices creadas</td></tr>`;
    return;
  }

  tbody.innerHTML = rows.map((row) => `
    <tr data-id="${escapeHtml(row.id)}">
      <td>${escapeHtml(row.consecutivo)}</td>
      <td>${escapeHtml(row.cliente)}</td>
      <td>${escapeHtml(row.responsable)}</td>
      <td>
        <select class="phase-select" data-phase-select data-previous-value="${Number(row.faseValue)}" aria-label="Fase del producto">
          ${Object.entries(faseLabels).map(([value, label]) => `
            <option value="${value}" ${Number(row.faseValue) === Number(value) ? "selected" : ""}>${escapeHtml(label)}</option>
          `).join("")}
        </select>
      </td>
      <td>${currency.format(Number(row.valorVenta || 0))}</td>
      <td>${currency.format(Number(row.utilidadBruta || 0))}</td>
      <td>
        <button class="icon-button compact-button" type="button" data-review-matrix="${escapeHtml(row.id)}" title="Revisar matriz">
          <i data-lucide="eye"></i>
          <span>Revisar</span>
        </button>
      </td>
    </tr>
  `).join("");

  if (window.lucide) window.lucide.createIcons();
}

async function loadMatrices(options = {}) {
  const tbody = document.querySelector("#matricesList");
  if (tbody) tbody.innerHTML = `<tr><td colspan="8">Cargando...</td></tr>`;

  try {
    const entitySet = await getMatrixEntitySetName(options);
    const query = [
      "$select=px_matrizcomercialid,px_consecutivo,px_cliente,px_responsable,px_fechasolicitud,px_estado,px_fase,px_hardware,px_obsequios,px_margenobjetivo,px_fletes,px_plazocredito,px_valorventa,px_valorfacturar,px_utilidadbruta",
      "$orderby=createdon desc",
      "$top=25"
    ].join("&");
    const data = await dataverseRequest("GET", `${entitySet}?${query}`, null, {
      Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
    }, options);
    renderMatrices(data.value.map(toMatrixRow));
  } catch (error) {
    if (tbody) tbody.innerHTML = `<tr><td colspan="8">${escapeHtml(error.message)}</td></tr>`;
  }
}

async function syncOnStartup() {
  const tbody = document.querySelector("#matricesList");
  if (tbody) tbody.innerHTML = `<tr><td colspan="7">Sincronizando...</td></tr>`;

  try {
    await loadMsal();
    const app = initMsal();
    currentAccount = app.getAllAccounts()[0] || null;

    if (!currentAccount) {
      if (tbody) tbody.innerHTML = `<tr><td colspan="7">Conecta con Microsoft para sincronizar matrices</td></tr>`;
      updateAuthStatus("Conecta con Microsoft para continuar");
      return;
    }

    updateConnectionLabel(currentAccount.name || currentAccount.username || "Conectado");
    updateAuthStatus("Sincronizando matrices...");
    setAuthenticatedUi();
    await loadMatrices({ interactive: false });
  } catch (error) {
    if (tbody) tbody.innerHTML = `<tr><td colspan="8">${escapeHtml(error.message)}</td></tr>`;
    updateAuthStatus(error.message);
  }
}

function toMatrixRow(row) {
  return {
    id: row.px_matrizcomercialid,
    consecutivo: row.px_consecutivo,
    cliente: row.px_cliente,
    responsable: row.px_responsable,
    fechaSolicitud: row.px_fechasolicitud,
    estado: row["px_estado@OData.Community.Display.V1.FormattedValue"] ?? row.px_estado,
    fase: faseLabels[row.px_fase] ?? row["px_fase@OData.Community.Display.V1.FormattedValue"] ?? row.px_fase,
    faseValue: row.px_fase,
    hardware: row.px_hardware,
    obsequios: row.px_obsequios,
    margenObjetivo: row.px_margenobjetivo,
    fletes: row.px_fletes,
    plazoCredito: row.px_plazocredito,
    valorVenta: row.px_valorventa,
    valorFacturar: row.px_valorfacturar,
    utilidadBruta: row.px_utilidadbruta
  };
}

async function nextConsecutive(entitySet) {
  const year = new Date().getFullYear();
  const prefix = `MX-${year}-`;
  const data = await dataverseRequest(
    "GET",
    `${entitySet}?$select=px_consecutivo&$filter=startswith(px_consecutivo,'${prefix}')&$orderby=px_consecutivo desc&$top=1`
  );
  const last = data.value?.[0]?.px_consecutivo;
  const next = last ? Number(String(last).split("-").at(-1)) + 1 : 1;
  return `${prefix}${String(next).padStart(4, "0")}`;
}

function buildDataversePayload(input) {
  const result = calculate();
  return {
    px_cliente: input.cliente || "",
    px_responsable: input.responsable || currentAccount?.name || currentAccount?.username || "",
    px_fechasolicitud: input.fechaSolicitud || null,
    px_estado: estadoValues.borrador,
    px_fase: Number(input.fase || faseValues.solicitudCredito),
    px_hardware: Number(input.hardware || 0),
    px_obsequios: Number(input.obsequios || 0),
    px_margenobjetivo: Number(input.margenObjetivo || 0),
    px_fletes: Number(input.fletes || 0),
    px_plazocredito: Number(input.plazoCredito || 0),
    px_totalcostosdirectos: result.totalCostosDirectos,
    px_valorventa: result.valorVenta,
    px_ivavalor: result.ivaValor,
    px_valorfacturar: result.valorFacturar,
    px_utilidadbruta: result.utilidadBruta,
    px_margensobrecostos: result.margenSobreCostos,
    px_margensobreventa: result.margenSobreVenta
  };
}

function stateFromMatrixRow(row) {
  return {
    ...defaults,
    consecutivo: row.consecutivo || "",
    fechaSolicitud: row.fechaSolicitud || "",
    cliente: row.cliente || "",
    responsable: row.responsable || "",
    fase: String(row.faseValue || faseValues.solicitudCredito),
    hardware: row.hardware ?? "",
    obsequios: row.obsequios ?? "",
    margenObjetivo: row.margenObjetivo ?? "",
    fletes: row.fletes ?? "",
    plazoCredito: row.plazoCredito ?? ""
  };
}

async function reviewMatrix(recordId) {
  if (!recordId) return;

  const entitySet = await getMatrixEntitySetName();
  const query = "$select=px_matrizcomercialid,px_consecutivo,px_cliente,px_responsable,px_fechasolicitud,px_estado,px_fase,px_hardware,px_obsequios,px_margenobjetivo,px_fletes,px_plazocredito,px_valorventa,px_valorfacturar,px_utilidadbruta";
  const row = await dataverseRequest("GET", `${entitySet}(${recordId})?${query}`, null, {
    Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
  });

  activeMatrixId = recordId;
  state = stateFromMatrixRow(toMatrixRow(row));
  syncInputsFromState();
  document.querySelector("#saveState").textContent = `Revisando ${state.consecutivo || "matriz"}`;
  switchView("evaluationView");
}

async function updateMatrixPhase(recordId, phaseValue) {
  if (!recordId) throw new Error("No se encontro la matriz para actualizar");
  const entitySet = await getMatrixEntitySetName();
  await dataverseRequest("PATCH", `${entitySet}(${recordId})`, {
    px_fase: Number(phaseValue)
  });
}

async function handleReviewClick(event) {
  const button = event.target.closest("[data-review-matrix]");
  if (!button) return;

  const previousText = button.querySelector("span").textContent;
  button.disabled = true;
  button.querySelector("span").textContent = "Abriendo...";

  try {
    await reviewMatrix(button.dataset.reviewMatrix);
  } catch (error) {
    document.querySelector("#saveState").textContent = error.message;
  } finally {
    button.disabled = false;
    button.querySelector("span").textContent = previousText;
  }
}

async function handlePhaseChange(event) {
  const select = event.target.closest("[data-phase-select]");
  if (!select) return;

  const row = select.closest("tr");
  const previousValue = select.dataset.previousValue || select.value;
  select.disabled = true;

  try {
    await updateMatrixPhase(row?.dataset.id, select.value);
    select.dataset.previousValue = select.value;
    document.querySelector("#saveState").textContent = `Fase actualizada: ${faseLabels[select.value]}`;
  } catch (error) {
    select.value = previousValue;
    document.querySelector("#saveState").textContent = error.message;
  } finally {
    select.disabled = false;
  }
}

async function saveMatrix() {
  const button = document.querySelector("#saveMatrixBtn");
  const previousText = button.querySelector("span").textContent;
  button.disabled = true;
  button.querySelector("span").textContent = "Guardando...";

  try {
    const entitySet = await getMatrixEntitySetName();
    if (activeMatrixId) {
      await dataverseRequest("PATCH", `${entitySet}(${activeMatrixId})`, buildDataversePayload(state));
      document.querySelector("#saveState").textContent = `Actualizada ${state.consecutivo || "matriz"}`;
    } else {
      const consecutivo = await nextConsecutive(entitySet);
      state.consecutivo = consecutivo;
      await dataverseRequest("POST", entitySet, {
        px_consecutivo: consecutivo,
        ...buildDataversePayload(state)
      });
      document.querySelector("#saveState").textContent = `Guardada ${consecutivo}`;
    }
    await loadMatrices();
    switchView("recordsView");
  } catch (error) {
    document.querySelector("#saveState").textContent = error.message;
  } finally {
    button.disabled = false;
    button.querySelector("span").textContent = previousText;
  }
}

document.addEventListener("DOMContentLoaded", () => {
  bindInputs();
  render();
  document.querySelector("#resetBtn").addEventListener("click", resetValues);
  document.querySelector("#printBtn").addEventListener("click", exportPdf);
  document.querySelector("#exportBtn").addEventListener("click", exportTable);
  document.querySelector("#saveMatrixBtn").addEventListener("click", saveMatrix);
  document.querySelector("#refreshMatricesBtn").addEventListener("click", loadMatrices);
  document.querySelector("#loginBtn").addEventListener("click", connectMicrosoft);
  document.querySelector("#authConnectBtn").addEventListener("click", connectMicrosoft);
  document.querySelector("#newMatrixBtn").addEventListener("click", startNewMatrix);
  document.querySelector("#matricesList").addEventListener("change", handlePhaseChange);
  document.querySelector("#matricesList").addEventListener("click", handleReviewClick);
  document.querySelectorAll("[data-view-target]").forEach((button) => {
    button.addEventListener("click", () => switchView(button.dataset.viewTarget));
  });
  switchView("recordsView");
  syncOnStartup();
  if (window.lucide) window.lucide.createIcons();
});
