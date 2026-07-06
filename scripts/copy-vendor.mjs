import { copyFileSync, mkdirSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const vendorDir = join(root, "public", "vendor");

mkdirSync(vendorDir, { recursive: true });
copyFileSync(
  join(root, "node_modules", "xlsx", "dist", "xlsx.full.min.js"),
  join(vendorDir, "xlsx.full.min.js")
);

console.log("Vendor copiado: public/vendor/xlsx.full.min.js");
