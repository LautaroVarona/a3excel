import { isEncrypted, decrypt } from "officecrypto-tool";
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

function isPasswordError(message: string): boolean {
  const text = message.toLowerCase();
  return text.includes("password") || text.includes("incorrect");
}

async function resolveWorkbookBuffer(
  input: Buffer,
  extension: ".xls" | ".xlsx",
  password?: string
): Promise<Buffer> {
  if (!isEncrypted(input)) {
    return input;
  }

  if (password) {
    try {
      const decrypted = await decrypt(input, { password });
      return Buffer.from(decrypted);
    } catch (error) {
      if (
        error instanceof Error &&
        isPasswordError(error.message) &&
        canUseExcelCom()
      ) {
        return convertEncryptedWorkbookViaExcel(input, extension);
      }
      throw error;
    }
  }

  if (canUseExcelCom()) {
    return convertEncryptedWorkbookViaExcel(input, extension);
  }

  throw new Error(
    "Este archivo está cifrado. En Windows con Excel instalado se abre automáticamente; " +
      "en otros entornos ingresá la contraseña de apertura."
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
    const workbookBuffer = await resolveWorkbookBuffer(
      input,
      extension,
      password
    );
    const data = parseWorkbookBuffer(workbookBuffer);

    return NextResponse.json(data);
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "No se pudo procesar el archivo Excel.";

    const status = message.includes("requiere Microsoft Excel") ? 501 : 422;

    return NextResponse.json({ error: message }, { status });
  }
}
