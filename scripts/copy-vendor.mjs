import { copyFileSync, mkdirSync, readFileSync, writeFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const publicVendorDir = join(root, "public", "vendor");
const officecryptoVendorDir = join(root, "lib", "officecrypto");
const officecryptoSource = join(root, "node_modules", "officecrypto-tool", "src");

mkdirSync(publicVendorDir, { recursive: true });
mkdirSync(officecryptoVendorDir, { recursive: true });

copyFileSync(
  join(root, "node_modules", "xlsx", "dist", "xlsx.full.min.js"),
  join(publicVendorDir, "xlsx.full.min.js")
);

const officecryptoFiles = [
  { from: join(officecryptoSource, "util", "xls97.js"), to: "xls97.js", patch: true },
  { from: join(officecryptoSource, "crypto", "xor_obfuscation.js"), to: "xor_obfuscation.js" },
  { from: join(officecryptoSource, "crypto", "rc4.js"), to: "rc4.js" },
  { from: join(officecryptoSource, "crypto", "rc4_cryptoapi.js"), to: "rc4_cryptoapi.js" },
];

for (const file of officecryptoFiles) {
  let content = readFileSync(file.from, "utf8");
  if (file.patch) {
    content = content
      .replace("require('../crypto/rc4')", "require('./rc4')")
      .replace("require('../crypto/rc4_cryptoapi')", "require('./rc4_cryptoapi')")
      .replace("require('../crypto/xor_obfuscation')", "require('./xor_obfuscation')");
  }
  writeFileSync(join(officecryptoVendorDir, file.to), content);
}

console.log("Vendor copiado: public/vendor/xlsx.full.min.js");
console.log("Vendor copiado: lib/officecrypto/*");
