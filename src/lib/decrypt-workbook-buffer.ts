import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";

import * as CFB from "cfb";
import officeCrypto from "officecrypto-tool";

import { parseWorkbookBuffer } from "./parse-workbook-buffer";

/** Claves habituales en exports .xls de a3ERP (XOR con contraseña vacía o un espacio). */
const DEFAULT_PASSWORD_CANDIDATES = ["", " ", "velneo", "VELNEO", "a3", "A3"] as const;

type Xls97Module = {
  decrypt: (
    currCfb: ReturnType<typeof CFB.read>,
    blob: Buffer,
    password: string,
    input: Buffer
  ) => Buffer;
};

function loadXls97Module(): Xls97Module | null {
  const modulePaths = [
    path.join(process.cwd(), "src/lib/vendor/officecrypto/xls97.js"),
    path.join(
      path.dirname(fileURLToPath(import.meta.url)),
      "vendor/officecrypto/xls97.js"
    ),
  ];

  for (const modulePath of modulePaths) {
    try {
      const require = createRequire(modulePath);
      return require("./xls97.js") as Xls97Module;
    } catch {
      // Probamos la siguiente ruta.
    }
  }

  return null;
}

const xls97Module = loadXls97Module();

export function buildDecryptPasswordCandidates(
  userPassword?: string
): string[] {
  const seen = new Set<string>();
  const candidates: string[] = [];
  const add = (value: string) => {
    if (seen.has(value)) return;
    seen.add(value);
    candidates.push(value);
  };

  if (userPassword) add(userPassword);
  for (const candidate of DEFAULT_PASSWORD_CANDIDATES) {
    add(candidate);
  }

  return candidates;
}

export function isOleXls(buffer: Buffer): boolean {
  return (
    buffer.length >= 4 &&
    buffer[0] === 0xd0 &&
    buffer[1] === 0xcf &&
    buffer[2] === 0x11 &&
    buffer[3] === 0xe0
  );
}

function decryptXls97WithPassword(
  input: Buffer,
  password: string
): Buffer | null {
  if (!xls97Module) return null;

  const cfb = CFB.read(input, { type: "buffer" });
  const workbookEntry =
    CFB.find(cfb, "Workbook") ?? CFB.find(cfb, "Book");
  if (!workbookEntry) return null;

  const workbookContent = workbookEntry.content;
  let workbookBlob: Buffer;
  if (!Buffer.isBuffer(workbookContent)) {
    workbookBlob = Buffer.from(workbookContent);
    CFB.utils.prep_blob(workbookBlob, 0);
  } else {
    workbookBlob = workbookContent;
  }

  try {
    const output = xls97Module.decrypt(cfb, workbookBlob, password, input);
    if (!Buffer.isBuffer(output)) {
      return Buffer.from(output);
    }
    return output;
  } catch {
    return null;
  }
}

function canParseWorkbook(buffer: Buffer, password?: string): boolean {
  try {
    parseWorkbookBuffer(buffer, password);
    return true;
  } catch {
    return false;
  }
}

function isUsableDecryptResult(
  input: Buffer,
  output: Buffer,
  password?: string
): boolean {
  if (output.length === 0) return false;
  if (!input.equals(output) && canParseWorkbook(output, password)) {
    return true;
  }
  return canParseWorkbook(output, password);
}

async function decryptWithCandidate(
  input: Buffer,
  password: string
): Promise<Buffer | null> {
  if (isOleXls(input)) {
    const xls97Output = decryptXls97WithPassword(input, password);
    if (xls97Output && isUsableDecryptResult(input, xls97Output, password)) {
      return xls97Output;
    }
  }

  if (password === "") {
    return null;
  }

  try {
    const output = Buffer.from(
      await officeCrypto.decrypt(input, { password })
    );
    if (isUsableDecryptResult(input, output, password)) {
      return output;
    }
  } catch {
    // Probamos la siguiente clave.
  }

  return null;
}

export type DecryptDiagnostics = {
  encrypted: boolean;
  oleXls: boolean;
  xls97ModuleLoaded: boolean;
  candidatesTried: number;
};

/**
 * Intenta descifrar un libro Excel protegido probando contraseña del usuario y
 * claves habituales de a3ERP. Devuelve null si ninguna candidata funciona.
 */
export async function tryDecryptWorkbookBuffer(
  input: Buffer,
  userPassword?: string
): Promise<Buffer | null> {
  const candidates = buildDecryptPasswordCandidates(userPassword);

  for (const password of candidates) {
    const output = await decryptWithCandidate(input, password);
    if (output) {
      return output;
    }
  }

  return null;
}

export function getDecryptDiagnostics(input: Buffer): DecryptDiagnostics {
  return {
    encrypted: officeCrypto.isEncrypted(input),
    oleXls: isOleXls(input),
    xls97ModuleLoaded: Boolean(xls97Module?.decrypt),
    candidatesTried: buildDecryptPasswordCandidates().length,
  };
}
