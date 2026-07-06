import {
  A3_DATA_FIRST_COL_INDEX,
  A3_DATA_START_ROW_1_BASED,
  A3_HEADER_LABELS,
  A3_HEADER_ROW_1_BASED,
  A3_LABEL_CELLS,
  A3_METADATA_CELLS,
  DEFAULT_A3_SECTION_TITLE,
} from "./excel-parse-constants";
import type { ExcelRow, ParsedExcel } from "./excel-types";

const INVALID_SHEET_CHARS = /[:\\/?*[\]]/g;

type XlsxModule = typeof import("xlsx-js-style");
type WorkSheet = import("xlsx-js-style").WorkSheet;
type CellObject = import("xlsx-js-style").CellObject;

interface CellStyle {
  font: {
    name: string;
    sz: number;
    bold?: boolean;
  };
}

function cellStyle(bold: boolean): CellStyle {
  return {
    font: {
      name: "Arial",
      sz: 10,
      ...(bold ? { bold: true } : {}),
    },
  };
}

function sanitizeSheetName(name: string): string {
  const cleaned = name.replace(INVALID_SHEET_CHARS, " ").trim() || "Datos";
  return cleaned.slice(0, 31);
}

function buildOutputFileName(
  sourceFileName?: string | null,
  sheetName?: string,
  extension: "xls" | "xlsx" = "xlsx"
): string {
  const base = sourceFileName
    ? sourceFileName.replace(/\.(xls|xlsx)$/i, "")
    : sheetName?.trim() || "exportacion-a3";
  return `${base}-editable.${extension}`;
}

function patchCellValue(
  worksheet: WorkSheet,
  address: string,
  value: string | number | boolean | null | undefined
): void {
  if (value === null || value === undefined || value === "") {
    delete worksheet[address];
    return;
  }

  if (typeof value === "number") {
    worksheet[address] = { t: "n", v: value };
    return;
  }

  if (typeof value === "boolean") {
    worksheet[address] = { t: "b", v: value };
    return;
  }

  worksheet[address] = { t: "s", v: String(value) };
}

async function exportYmantFromSourceBuffer(
  data: ParsedExcel,
  sourceBuffer: ArrayBuffer,
  sourceFileName?: string | null
): Promise<void> {
  const layout = data.layout;
  if (!layout || layout.kind !== "ymant") {
    throw new Error("El layout YMANT no está disponible para exportar.");
  }

  const XLSX = await import("xlsx-js-style");
  const workbook = XLSX.read(sourceBuffer, { type: "array", cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("El archivo original no contiene hojas.");
  }

  const worksheet = workbook.Sheets[sheetName];
  patchCellValue(worksheet, "B8", layout.controlCode);

  data.rows.forEach((row: ExcelRow, rowIndex) => {
    const excelRowIndex = layout.dataStartRow1Based - 1 + rowIndex;

    for (const column of data.columns) {
      const colIndex = layout.columnIndices[column];
      if (colIndex === undefined) continue;

      const address = XLSX.utils.encode_cell({
        c: colIndex,
        r: excelRowIndex,
      });
      patchCellValue(worksheet, address, row[column] ?? null);
    }
  });

  const useXls =
    sourceFileName?.toLowerCase().endsWith(".xls") ?? true;
  const outputName = buildOutputFileName(
    sourceFileName,
    data.sheetName,
    useXls ? "xls" : "xlsx"
  );

  XLSX.writeFile(workbook, outputName, {
    bookType: useXls ? "biff8" : "xlsx",
    compression: !useXls,
  });
}

function setCellValue(
  worksheet: WorkSheet,
  address: string,
  value: string | number | boolean | null | undefined,
  bold: boolean
): void {
  if (value === null || value === undefined || value === "") return;

  let cell: CellObject;

  if (typeof value === "number") {
    cell = { t: "n", v: value };
  } else if (typeof value === "boolean") {
    cell = { t: "b", v: value };
  } else {
    cell = { t: "s", v: String(value) };
  }

  cell.s = cellStyle(bold);
  worksheet[address] = cell;
}

function writeA3Header(worksheet: WorkSheet): void {
  setCellValue(worksheet, A3_LABEL_CELLS.company, A3_HEADER_LABELS.company, true);
  setCellValue(worksheet, A3_LABEL_CELLS.selection, A3_HEADER_LABELS.selection, true);
  setCellValue(worksheet, A3_LABEL_CELLS.date, A3_HEADER_LABELS.date, true);
}

function writeA3HeaderValues(worksheet: WorkSheet, data: ParsedExcel): void {
  const { metadata } = data;

  setCellValue(worksheet, A3_METADATA_CELLS.companyCode, metadata.companyCode, true);
  setCellValue(worksheet, A3_METADATA_CELLS.companyName, metadata.companyName, true);
  setCellValue(worksheet, A3_METADATA_CELLS.selection, metadata.selection, true);
  setCellValue(worksheet, A3_METADATA_CELLS.exportDate, metadata.exportDate, true);
  setCellValue(
    worksheet,
    A3_METADATA_CELLS.sectionTitle,
    metadata.sectionTitle ?? DEFAULT_A3_SECTION_TITLE,
    true
  );
}

function writeA3Table(
  worksheet: WorkSheet,
  XLSX: XlsxModule,
  data: ParsedExcel
): { lastRow: number; lastCol: number } {
  const headerRowIndex = A3_HEADER_ROW_1_BASED - 1;
  const firstColIndex = A3_DATA_FIRST_COL_INDEX;
  const lastColIndex = firstColIndex + data.columns.length - 1;

  data.columns.forEach((column, columnIndex) => {
    const address = XLSX.utils.encode_cell({
      c: firstColIndex + columnIndex,
      r: headerRowIndex,
    });
    setCellValue(worksheet, address, column, true);
  });

  data.rows.forEach((row: ExcelRow, rowIndex) => {
    const excelRowIndex = A3_DATA_START_ROW_1_BASED - 1 + rowIndex;

    data.columns.forEach((column, columnIndex) => {
      const address = XLSX.utils.encode_cell({
        c: firstColIndex + columnIndex,
        r: excelRowIndex,
      });
      setCellValue(worksheet, address, row[column] ?? null, false);
    });
  });

  const lastRowIndex = Math.max(
    headerRowIndex,
    data.rows.length > 0
      ? A3_DATA_START_ROW_1_BASED - 1 + data.rows.length - 1
      : headerRowIndex
  );

  return { lastRow: lastRowIndex, lastCol: lastColIndex };
}

function buildA3Worksheet(data: ParsedExcel, XLSX: XlsxModule): WorkSheet {
  const worksheet: WorkSheet = {};
  const firstColIndex = A3_DATA_FIRST_COL_INDEX;
  const headerRowIndex = A3_HEADER_ROW_1_BASED - 1;

  writeA3Header(worksheet);
  writeA3HeaderValues(worksheet, data);
  const { lastRow, lastCol } = writeA3Table(worksheet, XLSX, data);

  worksheet["!ref"] = XLSX.utils.encode_range({
    s: { c: 0, r: 0 },
    e: { c: lastCol, r: lastRow },
  });

  worksheet["!autofilter"] = {
    ref: XLSX.utils.encode_range({
      s: { c: firstColIndex, r: headerRowIndex },
      e: { c: lastCol, r: lastRow },
    }),
  };

  worksheet["!views"] = [
    {
      state: "frozen",
      ySplit: A3_HEADER_ROW_1_BASED,
      topLeftCell: XLSX.utils.encode_cell({
        c: firstColIndex,
        r: A3_DATA_START_ROW_1_BASED - 1,
      }),
      activeCell: XLSX.utils.encode_cell({
        c: firstColIndex,
        r: A3_DATA_START_ROW_1_BASED - 1,
      }),
    },
  ];

  worksheet["!merges"] = [
    {
      s: { c: firstColIndex, r: headerRowIndex - 1 },
      e: { c: firstColIndex + 1, r: headerRowIndex - 1 },
    },
  ];

  worksheet["!cols"] = [
    { wch: 4 },
    ...data.columns.map((column) => ({
      wch: Math.min(48, Math.max(10, column.length + 2)),
    })),
  ];

  return worksheet;
}

export async function exportParsedExcelToFile(
  data: ParsedExcel,
  options?: {
    sourceFileName?: string | null;
    sourceBuffer?: ArrayBuffer | null;
  }
): Promise<void> {
  if (data.layout?.kind === "ymant" && options?.sourceBuffer) {
    await exportYmantFromSourceBuffer(
      data,
      options.sourceBuffer,
      options.sourceFileName
    );
    return;
  }

  const XLSX = await import("xlsx-js-style");
  const worksheet = buildA3Worksheet(data, XLSX);

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    sanitizeSheetName(data.sheetName)
  );

  const outputName = buildOutputFileName(
    options?.sourceFileName,
    data.sheetName
  );
  XLSX.writeFile(workbook, outputName, { bookType: "xlsx", compression: true });
}
