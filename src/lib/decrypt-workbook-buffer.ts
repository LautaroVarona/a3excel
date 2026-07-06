import { createRequire } from "node:module";

import * as XLSX from "xlsx";

const require = createRequire(import.meta.url);

/** Claves habituales en exports .xls de a3ERP (XOR con contraseña vacía o un espacio). */
const DEFAULT_PASSWORD_CANDIDATES = ["", " "] as const;

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

function decryptXls97WithPassword(
  input: Buffer,
  password: string
): Buffer | null {
  const CFB = require("cfb") as typeof import("cfb");
  const xls97 = require("officecrypto-tool/src/util/xls97") as {
    decrypt: (
      currCfb: ReturnType<typeof CFB.read>,
      blob: Buffer,
      password: string,
      input: Buffer
    ) => Buffer;
  };

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
    return xls97.decrypt(cfb, workbookBlob, password, input);
  } catch {
    return null;
  }
}

function canReadWorkbook(buffer: Buffer): boolean {
  try {
    const workbook = XLSX.read(buffer, {
      type: "buffer",
      cellDates: true,
      sheetRows: 1,
    });
    return workbook.SheetNames.length > 0;
  } catch {
    return false;
  }
}

async function decryptWithCandidate(
  input: Buffer,
  password: string
): Promise<Buffer | null> {
  if (password === "") {
    return decryptXls97WithPassword(input, "");
  }

  const { decrypt } = await import("officecrypto-tool");
  try {
    return Buffer.from(await decrypt(input, { password }));
  } catch {
    return null;
  }
}

/**
 * Intenta descifrar un libro Excel protegido probando contraseña del usuario y
 * claves habituales de a3ERP. Devuelve null si ninguna candidata funciona.
 */
export async function tryDecryptWorkbookBuffer(
  input: Buffer,
  userPassword?: string
): Promise<Buffer | null> {
  const { isEncrypted } = await import("officecrypto-tool");
  const candidates = buildDecryptPasswordCandidates(userPassword);
  const shouldTryDecrypt = isEncrypted(input) || candidates.length > 0;

  if (!shouldTryDecrypt) return null;

  for (const password of candidates) {
    const output = await decryptWithCandidate(input, password);
    if (!output) continue;

    if (canReadWorkbook(output)) {
      return output;
    }
  }

  return null;
}
