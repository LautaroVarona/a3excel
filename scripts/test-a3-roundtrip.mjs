/**
 * Valida roundtrip de export A3 (sin cambios y con un parche numérico).
 * Ejecutar: npx tsx scripts/test-a3-roundtrip.mjs
 */
import fs from "node:fs";
import path from "node:path";
import * as CFB from "cfb";

const origPath = path.join(process.cwd(), "examples/E99995_060720_161422_15.XLS");

if (!fs.existsSync(origPath)) {
  console.error("Falta:", origPath);
  process.exit(1);
}

const input = fs.readFileSync(origPath);
const { parseWorkbookBuffer } = await import("../src/lib/parse-workbook-buffer.ts");
const { exportYmantPreservingBuffer } = await import("../src/lib/xls-preserving-export.ts");

function workbookSize(buffer) {
  const cfb = CFB.read(buffer, { type: "buffer" });
  const entry = CFB.find(cfb, "Workbook") ?? CFB.find(cfb, "Book");
  if (!entry) return 0;
  const content = entry.content;
  return Buffer.isBuffer(content) ? content.length : Buffer.from(content).length;
}

const parsed = parseWorkbookBuffer(input, "VelvetSweatshop");
if (!parsed.layout) {
  console.error("No se detectó layout A3");
  process.exit(1);
}

const exported = exportYmantPreservingBuffer(input, parsed, "VelvetSweatshop");
console.log("Original OLE:", input.length, "Workbook:", workbookSize(input));
console.log("Export OLE:", exported.length, "Workbook:", workbookSize(exported));
console.log("Sin cambios = original:", exported.equals(input));

const numericCol =
  parsed.columns.find((c) => /^\d{2,4}$/.test(c)) ?? parsed.columns[2];
const edited = {
  ...parsed,
  rows: parsed.rows.map((row, i) =>
    i === 0 ? { ...row, [numericCol]: 999 } : row
  ),
  originalRows: parsed.rows,
};

const patched = exportYmantPreservingBuffer(input, edited, "VelvetSweatshop");
console.log("Con parche OLE:", patched.length, "Workbook:", workbookSize(patched));
console.log(
  "Workbook mismo tamaño:",
  workbookSize(patched) === workbookSize(input)
);

if (!exported.equals(input)) {
  console.error("FAIL: export sin cambios difiere del original");
  process.exit(1);
}

if (workbookSize(patched) !== workbookSize(input)) {
  console.error("FAIL: stream Workbook cambió de tamaño tras parche");
  process.exit(1);
}

console.log("OK");
