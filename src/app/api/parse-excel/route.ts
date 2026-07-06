import { NextResponse } from "next/server";

import {
  canUseExcelCom,
  convertEncryptedWorkbookViaExcel,
} from "@/lib/excel-com";
import { parseWorkbookBuffer } from "@/lib/parse-workbook-buffer";
import { MAX_FILE_SIZE_BYTES } from "@/lib/excel-types";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const VALID_EXTENSIONS = [".xls", ".xlsx"];

function getExtension(fileName: string): ".xls" | ".xlsx" | null {
  const dotIndex = fileName.lastIndexOf(".");
  if (dotIndex === -1) return null;
  const extension = fileName.slice(dotIndex).toLowerCase();
  if (extension === ".xls") return ".xls";
  if (extension === ".xlsx") return ".xlsx";
  return null;
}

function isEncryptionError(message: string): boolean {
  const text = message.toLowerCase();
  return (
    text.includes("password") ||
    text.includes("encrypt") ||
    text.includes("file is password")
  );
}

/** Clave XOR habitual en exports .xls de a3ERP (protección de escritura). */
const A3ERP_XOR_PASSWORD = " ";

function buildDecryptPasswords(userPassword?: string): string[] {
  const seen = new Set<string>();
  const candidates: string[] = [];
  const add = (value: string) => {
    if (seen.has(value)) return;
    seen.add(value);
    candidates.push(value);
  };
  if (userPassword) add(userPassword);
  add(A3ERP_XOR_PASSWORD);
  return candidates;
}

async function tryOfficeCryptoDecrypt(
  input: Buffer,
  passwords: string[]
): Promise<Buffer | null> {
  const { decrypt } = await import("officecrypto-tool");
  for (const candidate of passwords) {
    try {
      const output = await decrypt(input, { password: candidate });
      return Buffer.from(output);
    } catch {
      // Probamos la siguiente clave.
    }
  }
  return null;
}

async function parseEncryptedBuffer(
  input: Buffer,
  extension: ".xls" | ".xlsx",
  password?: string
): Promise<Buffer> {
  const decrypted = await tryOfficeCryptoDecrypt(
    input,
    buildDecryptPasswords(password)
  );
  if (decrypted) {
    return decrypted;
  }

  if (canUseExcelCom()) {
    return convertEncryptedWorkbookViaExcel(input, extension);
  }

  throw new Error(
    password
      ? "La contraseña no es válida o el cifrado no es compatible con este entorno."
      : "Este archivo está cifrado. Ingresá la contraseña de apertura en el campo inferior, " +
          "o guardá una copia .xlsx sin protección desde Excel."
  );
}

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");
    const passwordField = formData.get("password");
    const password =
      typeof passwordField === "string" && passwordField.trim()
        ? passwordField.trim()
        : undefined;

    if (!(file instanceof File)) {
      return NextResponse.json(
        { error: "No se recibió ningún archivo." },
        { status: 400 }
      );
    }

    if (file.size <= 0) {
      return NextResponse.json(
        { error: "El archivo está vacío." },
        { status: 400 }
      );
    }

    if (file.size > MAX_FILE_SIZE_BYTES) {
      return NextResponse.json(
        { error: "El archivo supera el tamaño máximo permitido." },
        { status: 400 }
      );
    }

    const extension = getExtension(file.name);
    if (!extension || !VALID_EXTENSIONS.includes(extension)) {
      return NextResponse.json(
        { error: "Formato no válido. Solo se admiten .XLS y .XLSX." },
        { status: 400 }
      );
    }

    const input = Buffer.from(await file.arrayBuffer());

    try {
      return NextResponse.json(parseWorkbookBuffer(input, password));
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Error al leer el archivo.";

      if (!isEncryptionError(message)) {
        throw error;
      }

      const decrypted = await parseEncryptedBuffer(input, extension, password);
      return NextResponse.json(parseWorkbookBuffer(decrypted, password));
    }
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "No se pudo procesar el archivo Excel.";

    const status = message.includes("requiere Microsoft Excel") ? 501 : 422;

    return NextResponse.json({ error: message }, { status });
  }
}
