import type { WorkSheet } from "xlsx";

import type { ExcelExportMetadata } from "./excel-types";
import { extractA3Metadata as extractFromLayout } from "./a3-workbook-layout";
import {
  A3_METADATA_CELLS,
  DEFAULT_A3_SECTION_TITLE,
} from "./excel-parse-constants";

function readCellDisplay(worksheet: WorkSheet, address: string): string | null {
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

function extractLegacyMetadata(worksheet: WorkSheet): ExcelExportMetadata {
  return {
    companyCode: readCellDisplay(worksheet, A3_METADATA_CELLS.companyCode),
    companyName: readCellDisplay(worksheet, A3_METADATA_CELLS.companyName),
    selection: readCellDisplay(worksheet, A3_METADATA_CELLS.selection),
    exportDate: readCellDisplay(worksheet, A3_METADATA_CELLS.exportDate),
    sectionTitle:
      readCellDisplay(worksheet, A3_METADATA_CELLS.sectionTitle) ??
      DEFAULT_A3_SECTION_TITLE,
  };
}

export function extractA3Metadata(worksheet: WorkSheet): ExcelExportMetadata {
  const discovered = extractFromLayout(worksheet);
  if (
    discovered.companyCode ||
    discovered.companyName ||
    discovered.selection ||
    discovered.exportDate
  ) {
    return {
      ...discovered,
      sectionTitle: discovered.sectionTitle ?? DEFAULT_A3_SECTION_TITLE,
    };
  }

  return extractLegacyMetadata(worksheet);
}

export function hasA3Metadata(metadata: ExcelExportMetadata): boolean {
  return Boolean(
    metadata.companyCode ||
      metadata.companyName ||
      metadata.selection ||
      metadata.exportDate
  );
}
