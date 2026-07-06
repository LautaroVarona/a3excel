import { copyFileSync, mkdirSync, readFileSync, writeFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const vendorDir = join(root, "public", "vendor");

mkdirSync(vendorDir, { recursive: true });
copyFileSync(
  join(root, "node_modules", "xlsx", "dist", "xlsx.full.min.js"),
  join(vendorDir, "xlsx.full.min.js")
);

const workersDir = join(root, "public", "workers");
mkdirSync(workersDir, { recursive: true });

const discoverySource = readFileSync(
  join(root, "src", "lib", "a3-layout-discovery.js"),
  "utf8"
);
const workerDiscovery = discoverySource.replace(/\nexport \{[\s\S]*?\};\s*/g, "\n");
writeFileSync(join(workersDir, "a3-layout-discovery.js"), workerDiscovery);

console.log("Vendor copiado: public/vendor/xlsx.full.min.js");
console.log("Worker copiado: public/workers/a3-layout-discovery.js");
