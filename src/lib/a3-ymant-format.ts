import * as XLSX from "xlsx";
import type { WorkSheet } from "xlsx";

import type {
  A3YmantLayout,
  ExcelExportMetadata,
  ExcelRow,
} from "./excel-types";

/** Ubicación habitual del código de control YMANT (fila 8, columna B). */
export const A3_YMANT_CONTROL_CELL = "B8";

const CONTROL_CODE_ROW_INDEX = 7;
const CONTROL_CODE_PATTERN = /^077\d+/;
const CONTROL_CODE_LOOSE_PATTERN = /^077\d+.*\bCT\s+1/i;

function normalizeAccents(value: string): string {
  return value.normalize("NFD").replace(/\p{M}/gu, "");
}

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
  const text = value.trim();
  if (CONTROL_CODE_PATTERN.test(text)) return true;
  if (CONTROL_CODE_LOOSE_PATTERN.test(text)) return true;
  return false;
}

function isYmantHeaderLabel(value: string | null | undefined): boolean {
  if (!value) return false;
  const text = normalizeAccents(value.trim());
  return /^codigo\b/i.test(text) || /^cod\b\.?/i.test(text);
}

export function detectYmantControlCode(
  worksheet: WorkSheet
): Pick<A3YmantLayout, "controlCode" | "controlCodeColIndex" | "controlCodeRowIndex"> | null {
  const preferred = readCellDisplay(worksheet, A3_YMANT_CONTROL_CELL);
  if (isYmantControlCode(preferred)) {
    return {
      controlCode: preferred!.trim(),
      controlCodeColIndex: 1,
      controlCodeRowIndex: CONTROL_CODE_ROW_INDEX,
    };
  }

  for (let rowIndex = CONTROL_CODE_ROW_INDEX - 1; rowIndex <= CONTROL_CODE_ROW_INDEX + 1; rowIndex++) {
    for (let colIndex = 0; colIndex <= 4; colIndex++) {
      const value = readCellAt(worksheet, colIndex, rowIndex);
      if (!isYmantControlCode(value)) continue;
      return {
        controlCode: value!.trim(),
        controlCodeColIndex: colIndex,
        controlCodeRowIndex: rowIndex,
      };
    }
  }

  return null;
}

function findYmantHeaderRow(worksheet: WorkSheet): number | null {
  for (let rowIndex = CONTROL_CODE_ROW_INDEX; rowIndex <= 30; rowIndex++) {
    for (let colIndex = 0; colIndex <= 5; colIndex++) {
      const label = readCellAt(worksheet, colIndex, rowIndex);
      if (isYmantHeaderLabel(label)) {
        return rowIndex;
      }
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

  for (let colIndex = 0; colIndex <= maxCol; colIndex++) {
    const header = readCellAt(worksheet, colIndex, headerRowIndex);
    if (!header) continue;
    const key = header.trim();
    if (key.length === 0 || isYmantControlCode(key)) continue;
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
  const control = detectYmantControlCode(worksheet);
  if (!control) return null;

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
    controlCode: control.controlCode,
    controlCodeColIndex: control.controlCodeColIndex,
    controlCodeRowIndex: control.controlCodeRowIndex,
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
