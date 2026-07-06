import * as XLSX from "xlsx";

import { EXCEL_HEADER_ROW_INDEX } from "./excel-parse-constants";
import type { ExcelRow, ParsedExcel } from "./excel-types";

const MAX_ROWS = 50_000;

function normalizeCellValue(
  value: unknown
): string | number | boolean | null {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "boolean" || typeof value === "number") return value;
  return String(value);
}

export function parseWorkbookBuffer(buffer: Buffer | ArrayBuffer): ParsedExcel {
  const workbook = XLSX.read(buffer, {
    type: "buffer",
    cellDates: true,
    sheetRows: MAX_ROWS,
  });

  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("El archivo no contiene hojas de cálculo.");
  }

  const worksheet = workbook.Sheets[sheetName];
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
      normalized[String(key)] = normalizeCellValue(value);
    }
    return normalized;
  });

  const columnSet = new Set<string>();
  for (const row of rows) {
    for (const key of Object.keys(row)) {
      columnSet.add(key);
    }
  }

  return {
    sheetName,
    columns: Array.from(columnSet),
    rows,
    totalRows,
  };
}
