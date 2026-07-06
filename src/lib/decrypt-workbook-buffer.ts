import officeCrypto from "officecrypto-tool";

import { parseWorkbookBuffer } from "./parse-workbook-buffer";
import {
  decryptXls97Buffer,
  xls97DecryptAvailable,
} from "./xls97-decrypt";

/** Claves habituales en exports .xls de a3ERP (XOR con contraseña vacía o un espacio). */
const DEFAULT_PASSWORD_CANDIDATES = ["", " ", "velneo", "VELNEO", "a3", "A3"] as const;

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
    const xls97Output = decryptXls97Buffer(input, password);
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
    xls97ModuleLoaded: xls97DecryptAvailable,
    candidatesTried: buildDecryptPasswordCandidates().length,
  };
}
