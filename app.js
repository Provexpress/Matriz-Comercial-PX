const defaults = {
  fechaSolicitud: "",
  cliente: "",
  responsable: "",
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

async function getDataverseToken() {
  await loadMsal();
  const app = initMsal();
  const scopes = [`${dataverseConfig.environmentUrl}/user_impersonation`];
  currentAccount = app.getAllAccounts()[0] || currentAccount;

  if (!currentAccount) {
    const login = await app.loginPopup({ scopes });
    currentAccount = login.account;
  }

  try {
    const result = await app.acquireTokenSilent({ scopes, account: currentAccount });
    return result.accessToken;
  } catch {
    const result = await app.acquireTokenPopup({ scopes, account: currentAccount });
    currentAccount = result.account;
    return result.accessToken;
  }
}

async function connectMicrosoft() {
  await getDataverseToken();
  const label = document.querySelector("#loginBtn span");
  if (label && currentAccount) label.textContent = currentAccount.name || currentAccount.username || "Conectado";
  await loadMatrices();
}

async function dataverseRequest(method, relativeUrl, body, extraHeaders = {}) {
  const token = await getDataverseToken();
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

async function getMatrixEntitySetName() {
  if (matrixEntitySetName) return matrixEntitySetName;

  const data = await dataverseRequest(
    "GET",
    `EntityDefinitions(LogicalName='${dataverseConfig.tableLogicalName}')?$select=EntitySetName`
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
      saveState();
    });
  });
}

function resetValues() {
  state = { ...defaults };
  Object.entries(fields).forEach(([key, input]) => {
    input.value = state[key];
  });
  textFields.forEach((id) => {
    document.querySelector(`#${id}`).value = state[id] ?? "";
  });
  render();
  saveState();
}

function switchView(target) {
  const isEvaluation = target === "evaluationView";
  document.querySelector("#recordsView").classList.toggle("active", !isEvaluation);
  document.querySelector("#evaluationView").classList.toggle("active", isEvaluation);
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

function exportTable() {
  const projectRows = `
    <tr><th>Fecha solicitud</th><td>${state.fechaSolicitud || ""}</td></tr>
    <tr><th>Cliente</th><td>${state.cliente || ""}</td></tr>
    <tr><th>Responsable</th><td>${state.responsable || ""}</td></tr>
  `;
  const html = `
    <html><head><meta charset="utf-8"></head><body>
      <table>${projectRows}</table>
      ${document.querySelector("#evaluationTable").outerHTML}
    </body></html>
  `;
  const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `evaluacion-proyecto-${state.cliente || "provexpress"}.xls`;
  link.click();
  URL.revokeObjectURL(url);
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
      <td>${escapeHtml(row.estado)}</td>
      <td>
        <select class="phase-select" data-phase-select data-previous-value="${Number(row.faseValue)}" aria-label="Fase del producto">
          ${Object.entries(faseLabels).map(([value, label]) => `
            <option value="${value}" ${Number(row.faseValue) === Number(value) ? "selected" : ""}>${escapeHtml(label)}</option>
          `).join("")}
        </select>
      </td>
      <td>${currency.format(Number(row.valorVenta || 0))}</td>
      <td>${currency.format(Number(row.utilidadBruta || 0))}</td>
    </tr>
  `).join("");
}

async function loadMatrices() {
  const tbody = document.querySelector("#matricesList");
  if (tbody) tbody.innerHTML = `<tr><td colspan="7">Cargando...</td></tr>`;

  try {
    const entitySet = await getMatrixEntitySetName();
    const query = [
      "$select=px_matrizcomercialid,px_consecutivo,px_cliente,px_responsable,px_fechasolicitud,px_estado,px_fase,px_valorventa,px_valorfacturar,px_utilidadbruta",
      "$orderby=createdon desc",
      "$top=25"
    ].join("&");
    const data = await dataverseRequest("GET", `${entitySet}?${query}`, null, {
      Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
    });
    renderMatrices(data.value.map(toMatrixRow));
  } catch (error) {
    if (tbody) tbody.innerHTML = `<tr><td colspan="7">${escapeHtml(error.message)}</td></tr>`;
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

async function updateMatrixPhase(recordId, phaseValue) {
  if (!recordId) throw new Error("No se encontro la matriz para actualizar");
  const entitySet = await getMatrixEntitySetName();
  await dataverseRequest("PATCH", `${entitySet}(${recordId})`, {
    px_fase: Number(phaseValue)
  });
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
    const consecutivo = await nextConsecutive(entitySet);
    await dataverseRequest("POST", entitySet, {
      px_consecutivo: consecutivo,
      ...buildDataversePayload(state)
    });
    document.querySelector("#saveState").textContent = `Guardada ${consecutivo}`;
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
  document.querySelector("#printBtn").addEventListener("click", () => window.print());
  document.querySelector("#exportBtn").addEventListener("click", exportTable);
  document.querySelector("#saveMatrixBtn").addEventListener("click", saveMatrix);
  document.querySelector("#refreshMatricesBtn").addEventListener("click", loadMatrices);
  document.querySelector("#loginBtn").addEventListener("click", connectMicrosoft);
  document.querySelector("#newMatrixBtn").addEventListener("click", startNewMatrix);
  document.querySelector("#matricesList").addEventListener("change", handlePhaseChange);
  document.querySelectorAll("[data-view-target]").forEach((button) => {
    button.addEventListener("click", () => switchView(button.dataset.viewTarget));
  });
  switchView("recordsView");
  if (window.lucide) window.lucide.createIcons();
});
