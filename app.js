const defaults = {
  fechaSolicitud: "",
  cliente: "",
  responsable: "",
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

  ["fechaSolicitud", "cliente", "responsable"].forEach((id) => {
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
  ["fechaSolicitud", "cliente", "responsable"].forEach((id) => {
    document.querySelector(`#${id}`).value = state[id] ?? "";
  });
  render();
  saveState();
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

  const escapeHtml = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#039;"
  }[char]));

  tbody.innerHTML = rows.map((row) => `
    <tr>
      <td>${escapeHtml(row.consecutivo)}</td>
      <td>${escapeHtml(row.cliente)}</td>
      <td>${escapeHtml(row.responsable)}</td>
      <td>${escapeHtml(row.estado)}</td>
      <td>${escapeHtml(row.fase)}</td>
      <td>${currency.format(Number(row.valorVenta || 0))}</td>
      <td>${currency.format(Number(row.utilidadBruta || 0))}</td>
    </tr>
  `).join("");
}

async function loadMatrices() {
  const tbody = document.querySelector("#matricesList");
  if (tbody) tbody.innerHTML = `<tr><td colspan="7">Cargando...</td></tr>`;

  try {
    const response = await fetch("/api/matrices");
    if (!response.ok) throw new Error("No fue posible cargar matrices");
    renderMatrices(await response.json());
  } catch (error) {
    if (tbody) tbody.innerHTML = `<tr><td colspan="7">${error.message}</td></tr>`;
  }
}

async function saveMatrix() {
  const button = document.querySelector("#saveMatrixBtn");
  const previousText = button.querySelector("span").textContent;
  button.disabled = true;
  button.querySelector("span").textContent = "Guardando...";

  try {
    const response = await fetch("/api/matrices", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(state)
    });

    if (!response.ok) {
      const error = await response.json().catch(() => ({}));
      throw new Error(error.error || "No fue posible guardar la matriz");
    }

    const result = await response.json();
    document.querySelector("#saveState").textContent = `Guardada ${result.consecutivo}`;
    await loadMatrices();
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
  loadMatrices();
  if (window.lucide) window.lucide.createIcons();
});
