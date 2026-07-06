import * as XLSX from "xlsx";
import officeCrypto from "officecrypto-tool";

import { parseWorkbookBuffer } from "./parse-workbook-buffer";
import {
  decryptXls97Buffer,
  findVerifiedRc4Password,
  getXlsEncryptionInfo,
  xls97DecryptAvailable,
} from "./xls97-decrypt";

/**
 * Claves habituales en exports .xls de a3ERP / Office BIFF.
 * VelvetSweatshop: contraseña RC4 por defecto que Excel aplica sin pedirla al usuario.
 */
const DEFAULT_PASSWORD_CANDIDATES = [
  "VelvetSweatshop",
  "",
  " ",
  "velneo",
  "VELNEO",
  "Velneo",
  "a3",
  "A3",
  "a3erp",
  "A3ERP",
] as const;

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

function isUsableDecryptResult(input: Buffer, output: Buffer): boolean {
  if (output.length === 0) return false;
  if (input.equals(output)) return false;
  return canParseWorkbook(output);
}

function canParseWorkbook(buffer: Buffer): boolean {
  try {
    const workbook = XLSX.read(buffer, {
      type: "buffer",
      cellDates: true,
      sheetRows: 2,
    });
    if (workbook.SheetNames.length > 0) return true;
  } catch {
    // Seguimos con el parser completo de la app.
  }

  try {
    parseWorkbookBuffer(buffer);
    return true;
  } catch {
    return false;
  }
}

async function decryptWithCandidate(
  input: Buffer,
  password: string
): Promise<Buffer | null> {
  if (password !== "") {
    try {
      const output = Buffer.from(
        await officeCrypto.decrypt(input, { password })
      );
      if (isUsableDecryptResult(input, output)) {
        return output;
      }
    } catch {
      // officecrypto-tool usa crypto-js internamente; probamos nuestro RC4.
    }
  }

  if (isOleXls(input)) {
    const xls97Output = decryptXls97Buffer(input, password);
    if (xls97Output && isUsableDecryptResult(input, xls97Output)) {
      return xls97Output;
    }
  }

  return null;
}

export type DecryptDiagnostics = {
  encrypted: boolean;
  oleXls: boolean;
  xls97ModuleLoaded: boolean;
  candidatesTried: number;
  encryptionType?: string;
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
  const xlsInfo = isOleXls(input) ? getXlsEncryptionInfo(input) : null;

  if (xlsInfo?.encryptionType === "rc4") {
    const verified = findVerifiedRc4Password(input, candidates);
    if (verified) {
      const output = await decryptWithCandidate(input, verified);
      if (output) return output;
    }
  }

  for (const password of candidates) {
    const output = await decryptWithCandidate(input, password);
    if (output) {
      return output;
    }
  }

  return null;
}

export function getDecryptDiagnostics(input: Buffer): DecryptDiagnostics {
  const xlsInfo = isOleXls(input) ? getXlsEncryptionInfo(input) : null;
  return {
    encrypted: officeCrypto.isEncrypted(input),
    oleXls: isOleXls(input),
    xls97ModuleLoaded: xls97DecryptAvailable,
    candidatesTried: buildDecryptPasswordCandidates().length,
    encryptionType: xlsInfo?.encryptionType,
  };
}
