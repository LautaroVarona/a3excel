import type { WorkSheet } from "xlsx";

import type {
  A3ExportLayout,
  ExcelExportMetadata,
  ExcelRow,
} from "./excel-types";

import {
  detectA3ExportLayout as detectLayout,
  parseA3ExportWorksheet as parseLayoutWorksheet,
  extractA3Metadata as extractMetadata,
  isA3ControlCode,
} from "./a3-layout-discovery.js";

export {
  isA3ControlCode,
  extractMetadata as extractA3MetadataFromWorksheet,
};

export function detectA3ExportLayout(
  worksheet: WorkSheet
): A3ExportLayout | null {
  return detectLayout(worksheet) as A3ExportLayout | null;
}

export function parseA3ExportWorksheet(
  worksheet: WorkSheet,
  layout: A3ExportLayout,
  sheetName: string,
  maxRows: number
): {
  sheetName: string;
  columns: string[];
  rows: ExcelRow[];
  totalRows: number;
  metadata: ExcelExportMetadata;
  layout: A3ExportLayout;
} {
  return parseLayoutWorksheet(
    worksheet,
    layout,
    sheetName,
    maxRows
  ) as ReturnType<typeof parseA3ExportWorksheet>;
}

export function extractA3Metadata(
  worksheet: WorkSheet,
  headerRowIndex?: number
): ExcelExportMetadata {
  return extractMetadata(worksheet, headerRowIndex) as ExcelExportMetadata;
}
