/**
 * Diagnóstico de detección de exports A3.
 * Ejecutar: npx tsx scripts/debug-a3-detect.mjs
 */
import fs from "node:fs";
import path from "node:path";

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

const { parseWorkbookBuffer } = await import("../src/lib/parse-workbook-buffer.ts");

const parsed = parseWorkbookBuffer(input, "VelvetSweatshop");
console.log("Layout:", parsed.layout?.kind);
console.log("Control:", parsed.layout?.controlCode?.slice(0, 40));
console.log("Cabecera fila:", parsed.layout?.headerRow1Based);
console.log("Preamble celdas:", parsed.layout?.preambleCells?.length ?? 0);
console.log("Columnas:", parsed.columns.length, parsed.columns.slice(0, 5));
console.log("Filas:", parsed.totalRows);
console.log("Metadatos:", parsed.metadata);

if (parsed.layout?.kind !== "a3-export") {
  process.exit(1);
}

console.log("OK — export A3 detectado");
