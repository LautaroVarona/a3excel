/**
 * Diagnóstico YMANT sobre el .XLS original de A3.
 * Ejecutar: npx tsx scripts/debug-ymant-detect.mjs
 */
import fs from "node:fs";
import path from "node:path";
import * as XLSX from "xlsx";

const samplePath = path.join(
  process.cwd(),
  "examples/E99995_060720_161422_15.XLS"
);

if (!fs.existsSync(samplePath)) {
  console.error("Falta el archivo de ejemplo:", samplePath);
  process.exit(1);
}

const input = fs.readFileSync(samplePath);
console.log("Tamaño:", input.length);

const { decryptXls97Buffer } = await import("../src/lib/xls97-decrypt.ts");
const { detectYmantLayout, detectYmantControlCode } = await import(
  "../src/lib/a3-ymant-format.ts"
);
const { parseWorkbookBuffer } = await import(
  "../src/lib/parse-workbook-buffer.ts"
);

const decrypted = decryptXls97Buffer(input, "VelvetSweatshop");
console.log("Descifrado:", decrypted?.length ?? "falló");

const buffer = decrypted ?? input;
const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
const sheet = workbook.Sheets[workbook.SheetNames[0]];

for (const col of [0, 1, 2, 3]) {
  const addr = XLSX.utils.encode_cell({ c: col, r: 7 });
  const cell = sheet[addr];
  console.log(
    addr,
    cell ? { t: cell.t, v: cell.v, w: cell.w } : "(vacío)"
  );
}

const control = detectYmantControlCode(sheet);
console.log("Control detectado:", control);

const layout = detectYmantLayout(sheet);
console.log(
  "Layout:",
  layout
    ? {
        controlCode: layout.controlCode.slice(0, 32),
        headerRow: layout.headerRow1Based,
        columns: Object.keys(layout.columnIndices).slice(0, 5),
      }
    : null
);

const parsed = parseWorkbookBuffer(buffer, "VelvetSweatshop");
console.log("Parse layout:", parsed.layout?.kind);
console.log("Filas:", parsed.totalRows, "Columnas:", parsed.columns.length);
console.log("Primeras columnas:", parsed.columns.slice(0, 5));

if (parsed.layout?.kind !== "ymant") {
  process.exit(1);
}

console.log("OK — YMANT detectado");
