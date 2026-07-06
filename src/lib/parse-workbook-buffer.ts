import * as XLSX from "xlsx";

import { EXCEL_HEADER_ROW_INDEX } from "./excel-parse-constants";
import { extractA3Metadata } from "./extract-a3-metadata";
import type { ExcelRow, ParsedExcel } from "./excel-types";

const MAX_ROWS = 50_000;

/** Clave XOR habitual en exports .xls de a3ERP (protección de escritura). */
const A3ERP_XOR_PASSWORD = " ";

function isPasswordError(message: string): boolean {
  const text = message.toLowerCase();
  return (
    text.includes("password") ||
    text.includes("encrypt") ||
    text.includes("file is password") ||
    text.includes("unsupported password")
  );
}

function readWorkbook(
  buffer: Buffer | ArrayBuffer,
  password?: string
): XLSX.WorkBook {
  const base = {
    type: "buffer" as const,
    cellDates: true,
    sheetRows: MAX_ROWS,
  };

  const strategies: Array<{ password?: string }> = [];
  if (password) {
    strategies.push({ password });
  }
  strategies.push({});
  strategies.push({ password: A3ERP_XOR_PASSWORD });

  let lastError: unknown = null;

  for (const extra of strategies) {
    try {
      return XLSX.read(buffer, { ...base, ...extra });
    } catch (err) {
      lastError = err;
      const message = err instanceof Error ? err.message : String(err);
      if (!isPasswordError(message)) {
        throw err;
      }
    }
  }

  throw lastError instanceof Error
    ? lastError
    : new Error("No se pudo leer el archivo Excel protegido.");
}

function normalizeCellValue(
  value: unknown
): string | number | boolean | null {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "boolean" || typeof value === "number") return value;
  return String(value);
}

function isValidColumnKey(key: string): boolean {
  return key.length > 0 && !key.startsWith("__EMPTY");
}

export function parseWorkbookBuffer(
  buffer: Buffer | ArrayBuffer,
  password?: string
): ParsedExcel {
  const workbook = readWorkbook(buffer, password);

  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("El archivo no contiene hojas de cálculo.");
  }

  const worksheet = workbook.Sheets[sheetName];
  const metadata = extractA3Metadata(worksheet);
  const rawRows = XLSX.utils.sheet_to_json(worksheet, {
    range: EXCEL_HEADER_ROW_INDEX,
    defval: null,
    raw: false,
    blankrows: false,
  }) as Record<string, unknown>[];

  const totalRows = rawRows.length;
  const rows: ExcelRow[] = rawRows.map((raw) => {
    const normalized: ExcelRow = {};
    for (const [key, value] of Object.entries(raw)) {
      if (!isValidColumnKey(key)) continue;
      normalized[String(key)] = normalizeCellValue(value);
    }
    return normalized;
  });

  const columnSet = new Set<string>();
  for (const row of rows) {
    for (const key of Object.keys(row)) {
      if (isValidColumnKey(key)) columnSet.add(key);
    }
  }

  const columns = Array.from(columnSet);
  const filledRows: ExcelRow[] = rows.map((row) => {
    const filled: ExcelRow = {};
    for (const column of columns) {
      filled[column] = row[column] ?? null;
    }
    return filled;
  });

  return {
    sheetName,
    columns,
    rows: filledRows,
    totalRows,
    metadata,
  };
}
