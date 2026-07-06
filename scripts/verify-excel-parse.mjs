/**
 * Verificación rápida del parser Excel (misma lógica que el Web Worker).
 * Ejecutar: node scripts/verify-excel-parse.mjs
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import * as XLSX from "xlsx";

const EXCEL_HEADER_ROW_INDEX = 5;

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fixturePath = path.join(__dirname, "fixtures", "sample.xlsx");

function normalizeCellValue(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "boolean" || typeof value === "number") return value;
  return String(value);
}

function parseBuffer(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("Sin hojas");

  const worksheet = workbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json(worksheet, {
    range: EXCEL_HEADER_ROW_INDEX,
    defval: null,
    raw: false,
  });

  const rows = rawRows.map((raw) => {
    const normalized = {};
    for (const [key, value] of Object.entries(raw)) {
      normalized[String(key)] = normalizeCellValue(value);
    }
    return normalized;
  });

  const columnSet = new Set();
  for (const row of rows) {
    for (const key of Object.keys(row)) columnSet.add(key);
  }

  return {
    sheetName,
    columns: Array.from(columnSet),
    rows,
    totalRows: rows.length,
  };
}

// Crear o actualizar fixture con metadatos en filas 1-5 y encabezados en fila 6
const fixturesDir = path.dirname(fixturePath);
if (!fs.existsSync(fixturesDir)) fs.mkdirSync(fixturesDir, { recursive: true });

const ws = XLSX.utils.aoa_to_sheet([
  ["Metadato 1"],
  ["Metadato 2"],
  ["Metadato 3"],
  ["Metadato 4"],
  ["Metadato 5"],
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
console.log("Primera fila:", JSON.stringify(result.rows[0]));

if (result.totalRows !== 3 || result.columns.length !== 3) {
  console.error("FAIL: datos inesperados");
  process.exit(1);
}

console.log("PASS");
