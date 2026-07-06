/**
 * Verificación rápida del parser Excel (misma lógica que el Web Worker).
 * Ejecutar: node scripts/verify-excel-parse.mjs
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import * as XLSX from "xlsx";

const EXCEL_HEADER_ROW_INDEX = 7;

const A3_METADATA_CELLS = {
  companyCode: "D2",
  companyName: "E2",
  selection: "D3",
  exportDate: "D5",
};

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fixturePath = path.join(__dirname, "fixtures", "sample.xlsx");

function normalizeCellValue(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "boolean" || typeof value === "number") return value;
  return String(value);
}

function isValidColumnKey(key) {
  return key.length > 0 && !key.startsWith("__EMPTY");
}

function readCellDisplay(worksheet, address) {
  const cell = worksheet[address];
  if (!cell) return null;

  if (cell.w != null && cell.w !== "") {
    return String(cell.w);
  }

  const value = cell.v;
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) {
    return value.toLocaleDateString("es-ES");
  }
  return String(value);
}

function extractA3Metadata(worksheet) {
  return {
    companyCode: readCellDisplay(worksheet, A3_METADATA_CELLS.companyCode),
    companyName: readCellDisplay(worksheet, A3_METADATA_CELLS.companyName),
    selection: readCellDisplay(worksheet, A3_METADATA_CELLS.selection),
    exportDate: readCellDisplay(worksheet, A3_METADATA_CELLS.exportDate),
  };
}

function parseBuffer(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("Sin hojas");

  const worksheet = workbook.Sheets[sheetName];
  const metadata = extractA3Metadata(worksheet);
  const rawRows = XLSX.utils.sheet_to_json(worksheet, {
    range: EXCEL_HEADER_ROW_INDEX,
    defval: null,
    raw: false,
  });

  const rows = rawRows.map((raw) => {
    const normalized = {};
    for (const [key, value] of Object.entries(raw)) {
      if (!isValidColumnKey(key)) continue;
      normalized[String(key)] = normalizeCellValue(value);
    }
    return normalized;
  });

  const columnSet = new Set();
  for (const row of rows) {
    for (const key of Object.keys(row)) {
      if (isValidColumnKey(key)) columnSet.add(key);
    }
  }

  return {
    sheetName,
    columns: Array.from(columnSet),
    rows,
    totalRows: rows.length,
    metadata,
  };
}

const fixturesDir = path.dirname(fixturePath);
if (!fs.existsSync(fixturesDir)) fs.mkdirSync(fixturesDir, { recursive: true });

const ws = XLSX.utils.aoa_to_sheet([
  [],
  ["", "", "EMPRESA:", "COD001", "Mi Empresa SA"],
  ["", "", "SELECCIÓN", "Todos los registros"],
  [],
  ["", "", "FECHA:", "15/01/2025"],
  [],
  [],
  ["Nombre", "Edad", "Activo"],
  ["Ana", 30, true],
  ["Luis", 25, false],
  ["María", 35, true],
]);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "Datos");
XLSX.writeFile(wb, fixturePath);
console.log("Fixture actualizado:", fixturePath);

const buffer = fs.readFileSync(fixturePath);
const start = Date.now();
const result = parseBuffer(buffer);
const elapsed = Date.now() - start;

console.log("Parse OK en", elapsed, "ms");
console.log("Hoja:", result.sheetName);
console.log("Columnas:", result.columns.join(", "));
console.log("Filas:", result.totalRows);
console.log("Metadatos:", JSON.stringify(result.metadata));
console.log("Primera fila:", JSON.stringify(result.rows[0]));

if (result.totalRows !== 3 || result.columns.length !== 3) {
  console.error("FAIL: datos inesperados");
  process.exit(1);
}

if (result.metadata.companyCode !== "COD001") {
  console.error("FAIL: código de empresa incorrecto");
  process.exit(1);
}

if (result.metadata.companyName !== "Mi Empresa SA") {
  console.error("FAIL: nombre de empresa incorrecto");
  process.exit(1);
}

if (result.metadata.selection !== "Todos los registros") {
  console.error("FAIL: selección incorrecta");
  process.exit(1);
}

if (result.metadata.exportDate !== "15/01/2025") {
  console.error("FAIL: fecha incorrecta");
  process.exit(1);
}

console.log("PASS");
