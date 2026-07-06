import fs from "fs";
import path from "path";
import { createRequire } from "module";

const require = createRequire(import.meta.url);
const tsNodePaths = [
  path.join(process.cwd(), "node_modules/tsx/dist/cli.mjs"),
];

// Dynamic import compiled logic via tsx
const { register } = await import("tsx/esm/api").catch(() => ({}));

async function loadModule(relativePath) {
  const full = path.join(process.cwd(), relativePath);
  return import(full);
}

const origPath = path.join(process.cwd(), "examples/E99995_060720_161422_15.XLS");
const input = fs.readFileSync(origPath);

console.log("Original size:", input.length);

const { parseWorkbookBuffer } = await import("../src/lib/parse-workbook-buffer.ts");
const { exportYmantPreservingBuffer } = await import("../src/lib/xls-preserving-export.ts");
const { getXlsEncryptionInfo, decryptXls97Buffer } = await import("../src/lib/xls97-decrypt.ts");

console.log("Encryption:", getXlsEncryptionInfo(input));

const decrypted = decryptXls97Buffer(input, "VelvetSweatshop");
console.log("Decrypted size:", decrypted?.length ?? "fail");

const parsed = parseWorkbookBuffer(input, "VelvetSweatshop");
console.log("Layout:", parsed.layout?.kind, parsed.layout?.controlCode?.slice(0, 30));
console.log("Columns:", parsed.columns.length, parsed.columns.slice(0, 3));
console.log("Rows:", parsed.totalRows);

const exported = exportYmantPreservingBuffer(input, parsed, "VelvetSweatshop");
console.log("Export size (no patches):", exported.length);
console.log("Same as original:", exported.equals(input));

if (parsed.layout) {
  const edited = {
    ...parsed,
    rows: parsed.rows.map((row, i) =>
      i === 0
        ? { ...row, [parsed.columns.find((c) => c.includes("093")) ?? parsed.columns[2]]: 999 }
        : row
    ),
    originalRows: parsed.rows,
  };
  const exportedEdit = exportYmantPreservingBuffer(input, edited, "VelvetSweatshop");
  console.log("Export size (with patch):", exportedEdit.length);
}
