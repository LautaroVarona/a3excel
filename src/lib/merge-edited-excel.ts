import * as XLSX from "xlsx";

import { detectA3ExportLayout } from "./a3-workbook-layout";
import type { A3ExportLayout, ExcelRow } from "./excel-types";

function normalizeCellValue(
  value: unknown
): string | number | boolean | null {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "boolean" || typeof value === "number") return value;
  return String(value);
}

function readCellAt(
  worksheet: XLSX.WorkSheet,
  colIndex: number,
  rowIndex: number
): string | number | boolean | null {
  const address = XLSX.utils.encode_cell({ c: colIndex, r: rowIndex });
  const cell = worksheet[address];
  if (!cell) return null;

  if (cell.w != null && cell.w !== "") {
    const text = String(cell.w).trim();
    const asNumber = Number(text.replace(",", "."));
    if (cell.t === "n") return cell.v as number;
    if (!Number.isNaN(asNumber) && /^-?\d+([.,]\d+)?$/.test(text)) {
      return asNumber;
    }
    return text;
  }

  return normalizeCellValue(cell.v);
}

export function parseEditedA3Rows(
  buffer: ArrayBuffer,
  layout: A3ExportLayout,
  columns: string[]
): ExcelRow[] {
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("El archivo editado no contiene hojas.");
  }

  const worksheet = workbook.Sheets[sheetName];
  const rows: ExcelRow[] = [];
  const dataStartIndex = layout.dataStartRow1Based - 1;
  const maxRow = dataStartIndex + 50_000;

  for (let rowIndex = dataStartIndex; rowIndex < maxRow; rowIndex++) {
    const row: ExcelRow = {};
    let hasValue = false;

    for (const column of columns) {
      const colIndex = layout.columnIndices[column];
      if (colIndex === undefined) continue;
      const value = readCellAt(worksheet, colIndex, rowIndex);
      row[column] = value;
      if (value !== null) hasValue = true;
    }

    if (hasValue) {
      rows.push(row);
    } else if (rows.length > 0) {
      break;
    }
  }

  if (rows.length === 0) {
    throw new Error(
      "No se encontraron filas de datos en el Excel editado. " +
        "Guardá el archivo sin cambiar la estructura (filas/columnas del export)."
    );
  }

  return rows;
}

/** @deprecated Usar parseEditedA3Rows */
export const parseEditedYmantRows = parseEditedA3Rows;

export function resolveLayoutFromEditedFile(
  buffer: ArrayBuffer,
  fallback?: A3ExportLayout
): A3ExportLayout {
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("El archivo editado no contiene hojas.");
  }

  const detected = detectA3ExportLayout(workbook.Sheets[sheetName]);
  if (detected) return detected;

  if (fallback) return fallback;

  throw new Error(
    "No se reconoce la estructura del export A3 en el Excel editado. " +
      "Usá el .xlsx descargado desde esta app."
  );
}
