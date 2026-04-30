const defaults = {
  fechaSolicitud: new Date().toISOString().slice(0, 10),
  cliente: "",
  responsable: "",
  hardware: 2000000,
  servicios: 0,
  licenciamientos: 0,
  obsequios: 0,
  margenObjetivo: 15,
  iva: 19,
  comision: 15,
  impuestos: 3.5,
  fletes: 0,
  imprevistos: 0.02,
  plazoCredito: 90,
  baseFinanciacion: 1.1
};

const moneyFields = new Set([
  "hardware",
  "servicios",
  "licenciamientos",
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

const storageKey = "matriz-comercial-px";
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
    Number(state.servicios) +
    Number(state.licenciamientos) +
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
      state[input.dataset.field] = Number(input.value || 0);
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

document.addEventListener("DOMContentLoaded", () => {
  bindInputs();
  render();
  document.querySelector("#resetBtn").addEventListener("click", resetValues);
  document.querySelector("#printBtn").addEventListener("click", () => window.print());
  document.querySelector("#exportBtn").addEventListener("click", exportTable);
  if (window.lucide) window.lucide.createIcons();
});
