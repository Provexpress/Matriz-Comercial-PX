import express from "express";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { dataverseRequest, getTableEntitySet, readSchema, choiceValue } from "./dataverse.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, "..");
const app = express();
const port = process.env.PORT || 3000;

app.use(express.json({ limit: "1mb" }));
app.use(express.static(rootDir));

function calculate(input) {
  const asRate = (value) => Number(value || 0) / 100;
  const totalCostosDirectos = Number(input.hardware || 0) + Number(input.obsequios || 0);
  const margenObjetivo = asRate(input.margenObjetivo);
  const valorVenta = margenObjetivo >= 1 ? 0 : totalCostosDirectos / (1 - margenObjetivo);
  const utilidadEsperada = valorVenta - totalCostosDirectos;
  const ivaValor = valorVenta * asRate(input.iva ?? 19);
  const valorFacturar = valorVenta + ivaValor;
  const comisionValor = asRate(input.comision ?? 15) * utilidadEsperada;
  const impuestosValor = valorVenta * asRate(input.impuestos ?? 3.5);
  const imprevistosValor = valorVenta * asRate(input.imprevistos ?? 0.02);
  const financiacionPorcentaje =
    Number(input.plazoCredito || 0) > 60
      ? ((Number(input.plazoCredito || 0) - 60) / 30) * asRate(input.baseFinanciacion ?? 1.1)
      : 0;
  const financiacionValor = financiacionPorcentaje * valorVenta;
  const totalCostosInternos =
    financiacionValor + imprevistosValor + Number(input.fletes || 0) + impuestosValor + comisionValor;
  const utilidadBruta = utilidadEsperada - totalCostosInternos;
  const margenSobreCostos = totalCostosDirectos ? utilidadBruta / totalCostosDirectos : 0;
  const margenSobreVenta = valorVenta ? utilidadBruta / valorVenta : 0;

  return {
    totalCostosDirectos,
    valorVenta,
    ivaValor,
    valorFacturar,
    utilidadBruta,
    margenSobreCostos,
    margenSobreVenta
  };
}

async function getMatrixConfig() {
  const schema = await readSchema();
  const table = schema.tables[0];
  return {
    table,
    entitySet: await getTableEntitySet(table.schemaName)
  };
}

async function nextConsecutive(entitySet) {
  const year = new Date().getFullYear();
  const prefix = `MX-${year}-`;
  const result = await dataverseRequest(
    "GET",
    `${entitySet}?$select=px_consecutivo&$filter=startswith(px_consecutivo,'${prefix}')&$orderby=px_consecutivo desc&$top=1`
  );
  const last = result.value?.[0]?.px_consecutivo;
  const next = last ? Number(last.split("-").at(-1)) + 1 : 1;
  return `${prefix}${String(next).padStart(4, "0")}`;
}

function toApiRecord(row) {
  return {
    id: row.px_matrizcomercialid,
    consecutivo: row.px_consecutivo,
    cliente: row.px_cliente,
    responsable: row.px_responsable,
    fechaSolicitud: row.px_fechasolicitud,
    estado: row["px_estado@OData.Community.Display.V1.FormattedValue"] ?? row.px_estado,
    fase: row["px_fase@OData.Community.Display.V1.FormattedValue"] ?? row.px_fase,
    valorVenta: row.px_valorventa,
    valorFacturar: row.px_valorfacturar,
    utilidadBruta: row.px_utilidadbruta
  };
}

app.get("/api/health", async (_request, response, next) => {
  try {
    const { entitySet } = await getMatrixConfig();
    response.json({ ok: true, entitySet });
  } catch (error) {
    next(error);
  }
});

app.get("/api/matrices", async (_request, response, next) => {
  try {
    const { entitySet } = await getMatrixConfig();
    const query = [
      "$select=px_matrizcomercialid,px_consecutivo,px_cliente,px_responsable,px_fechasolicitud,px_estado,px_fase,px_valorventa,px_valorfacturar,px_utilidadbruta",
      "$orderby=createdon desc",
      "$top=25"
    ].join("&");
    const data = await dataverseRequest("GET", `${entitySet}?${query}`, null, {
      Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
    });
    response.json(data.value.map(toApiRecord));
  } catch (error) {
    next(error);
  }
});

app.post("/api/matrices", async (request, response, next) => {
  try {
    const input = request.body;
    const { entitySet } = await getMatrixConfig();
    const totals = calculate(input);
    const consecutivo = input.consecutivo || await nextConsecutive(entitySet);
    const payload = {
      px_consecutivo: consecutivo,
      px_cliente: input.cliente || "",
      px_responsable: input.responsable || "",
      px_fechasolicitud: input.fechaSolicitud || null,
      px_estado: choiceValue(0),
      px_fase: choiceValue(0),
      px_hardware: Number(input.hardware || 0),
      px_obsequios: Number(input.obsequios || 0),
      px_margenobjetivo: Number(input.margenObjetivo || 0),
      px_fletes: Number(input.fletes || 0),
      px_plazocredito: Number(input.plazoCredito || 0),
      px_totalcostosdirectos: totals.totalCostosDirectos,
      px_valorventa: totals.valorVenta,
      px_ivavalor: totals.ivaValor,
      px_valorfacturar: totals.valorFacturar,
      px_utilidadbruta: totals.utilidadBruta,
      px_margensobrecostos: totals.margenSobreCostos,
      px_margensobreventa: totals.margenSobreVenta
    };

    await dataverseRequest("POST", entitySet, payload);
    response.status(201).json({ consecutivo, ...totals });
  } catch (error) {
    next(error);
  }
});

app.use((error, _request, response, _next) => {
  console.error(error);
  response.status(500).json({ error: error.message });
});

app.listen(port, () => {
  console.log(`Matriz Comercial PX running at http://localhost:${port}`);
});
