import * as XLSX from "xlsx";
import type { WorkSheet } from "xlsx";

import type {
  A3YmantLayout,
  ExcelExportMetadata,
  ExcelRow,
} from "./excel-types";

/** Código de control A3NOM formato 77 (YMANT) en celda B8. */
export const A3_YMANT_CONTROL_CELL = "B8";

const CONTROL_CODE_PATTERN = /^077\d+/;

export function readCellDisplay(
  worksheet: WorkSheet,
  address: string
): string | null {
  const cell = worksheet[address];
  if (!cell) return null;

  if (cell.w != null && cell.w !== "") {
    return String(cell.w);
  }

  const value = cell.v;
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) {
    return value.toLocaleDateString("es-ES");
  }
  return String(value);
}

function readCellAt(
  worksheet: WorkSheet,
  colIndex: number,
  rowIndex: number
): string | null {
  const address = XLSX.utils.encode_cell({ c: colIndex, r: rowIndex });
  return readCellDisplay(worksheet, address);
}

export function isYmantControlCode(value: string | null | undefined): boolean {
  if (!value) return false;
  return CONTROL_CODE_PATTERN.test(value.trim());
}

export function detectYmantControlCode(
  worksheet: WorkSheet
): string | null {
  const value = readCellDisplay(worksheet, A3_YMANT_CONTROL_CELL);
  if (!isYmantControlCode(value)) return null;
  return value!;
}

function findYmantHeaderRow(worksheet: WorkSheet): number | null {
  for (let rowIndex = 7; rowIndex <= 30; rowIndex++) {
    const label = readCellAt(worksheet, 1, rowIndex);
    if (label && /c[oó]digo/i.test(label.trim())) {
      return rowIndex;
    }
  }
  return null;
}

function buildColumnIndices(
  worksheet: WorkSheet,
  headerRowIndex: number,
  maxCol: number
): Record<string, number> {
  const indices: Record<string, number> = {};

  for (let colIndex = 1; colIndex <= maxCol; colIndex++) {
    const header = readCellAt(worksheet, colIndex, headerRowIndex);
    if (!header) continue;
    const key = header.trim();
    if (key.length === 0) continue;
    indices[key] = colIndex;
  }

  return indices;
}

export function extractYmantMetadata(
  worksheet: WorkSheet
): ExcelExportMetadata {
  return {
    companyCode: readCellAt(worksheet, 3, 8),
    companyName: readCellAt(worksheet, 4, 8),
    selection: readCellAt(worksheet, 4, 9),
    exportDate: readCellAt(worksheet, 3, 10),
    sectionTitle: readCellAt(worksheet, 1, 11),
  };
}

export function detectYmantLayout(worksheet: WorkSheet): A3YmantLayout | null {
  const controlCode = detectYmantControlCode(worksheet);
  if (!controlCode) return null;

  const headerRowIndex = findYmantHeaderRow(worksheet);
  if (headerRowIndex === null) return null;

  const ref = worksheet["!ref"];
  const maxCol = ref ? parseMaxCol(ref) : 120;

  const columnIndices = buildColumnIndices(
    worksheet,
    headerRowIndex,
    maxCol
  );
  const columns = Object.keys(columnIndices);
  if (columns.length === 0) return null;

  return {
    kind: "ymant",
    controlCode,
    headerRow1Based: headerRowIndex + 1,
    dataStartRow1Based: headerRowIndex + 2,
    columnIndices,
  };
}

function parseMaxCol(ref: string): number {
  const match = ref.match(/:([A-Z]+)(\d+)$/i);
  if (!match) return 120;
  const letters = match[1].toUpperCase();
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col - 1;
}

function normalizeCellValue(
  value: unknown
): string | number | boolean | null {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "boolean" || typeof value === "number") return value;
  return String(value);
}

export function parseYmantWorksheet(
  worksheet: WorkSheet,
  layout: A3YmantLayout,
  sheetName: string,
  maxRows: number
): {
  sheetName: string;
  columns: string[];
  rows: ExcelRow[];
  totalRows: number;
  metadata: ExcelExportMetadata;
  layout: A3YmantLayout;
} {
  const columns = Object.keys(layout.columnIndices);
  const metadata = extractYmantMetadata(worksheet);
  const dataStartIndex = layout.dataStartRow1Based - 1;
  const maxRow = layout.dataStartRow1Based - 1 + maxRows;

  const rows: ExcelRow[] = [];

  for (let rowIndex = dataStartIndex; rowIndex < maxRow; rowIndex++) {
    const row: ExcelRow = {};
    let hasValue = false;

    for (const column of columns) {
      const colIndex = layout.columnIndices[column];
      const raw = readCellAt(worksheet, colIndex, rowIndex);
      const value = normalizeCellValue(raw);
      row[column] = value;
      if (value !== null) hasValue = true;
    }

    if (hasValue) {
      rows.push(row);
    } else if (rows.length > 0) {
      break;
    }
  }

  return {
    sheetName,
    columns,
    rows,
    totalRows: rows.length,
    metadata,
    layout,
  };
}
