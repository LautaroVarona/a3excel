import type { ParsedExcel } from "./excel-types";
import {
  buildEditableExportName,
} from "./xls-preserving-export";

function downloadBuffer(buffer: ArrayBuffer, fileName: string): void {
  const blob = new Blob([buffer], {
    type: "application/vnd.ms-excel",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

async function exportYmantViaServer(
  data: ParsedExcel,
  sourceBuffer: ArrayBuffer,
  sourceFileName?: string | null,
  password?: string
): Promise<void> {
  const formData = new FormData();
  formData.append(
    "file",
    new Blob([sourceBuffer], { type: "application/vnd.ms-excel" }),
    sourceFileName ?? "export.xls"
  );
  formData.append(
    "payload",
    JSON.stringify({
      rows: data.rows,
      originalRows: data.originalRows,
      layout: data.layout,
      columns: data.columns,
      sheetName: data.sheetName,
      totalRows: data.totalRows,
      metadata: data.metadata,
    })
  );
  if (password) {
    formData.append("password", password);
  }

  const response = await fetch("/api/export-excel", {
    method: "POST",
    body: formData,
  });

  if (!response.ok) {
    const payload = (await response.json()) as { error?: string };
    throw new Error(payload.error ?? "No se pudo exportar el archivo Excel.");
  }

  const output = await response.arrayBuffer();
  downloadBuffer(output, buildEditableExportName(sourceFileName));
}

export async function exportParsedExcelToFile(
  data: ParsedExcel,
  options?: {
    sourceFileName?: string | null;
    sourceBuffer?: ArrayBuffer | null;
    password?: string;
  }
): Promise<void> {
  if (data.layout?.kind === "ymant" && options?.sourceBuffer) {
    await exportYmantViaServer(
      data,
      options.sourceBuffer,
      options.sourceFileName,
      options.password
    );
    return;
  }

  const {
    A3_DATA_FIRST_COL_INDEX,
    A3_DATA_START_ROW_1_BASED,
    A3_HEADER_LABELS,
    A3_HEADER_ROW_1_BASED,
    A3_LABEL_CELLS,
    A3_METADATA_CELLS,
    DEFAULT_A3_SECTION_TITLE,
  } = await import("./excel-parse-constants");

  const XLSX = await import("xlsx-js-style");
  type WorkSheet = import("xlsx-js-style").WorkSheet;
  type CellObject = import("xlsx-js-style").CellObject;

  const worksheet: WorkSheet = {};
  const firstColIndex = A3_DATA_FIRST_COL_INDEX;
  const headerRowIndex = A3_HEADER_ROW_1_BASED - 1;

  const setCellValue = (
    ws: WorkSheet,
    address: string,
    value: string | number | boolean | null | undefined,
    bold: boolean
  ) => {
    if (value === null || value === undefined || value === "") return;

    let cell: CellObject;
    if (typeof value === "number") {
      cell = { t: "n", v: value };
    } else if (typeof value === "boolean") {
      cell = { t: "b", v: value };
    } else {
      cell = { t: "s", v: String(value) };
    }

    cell.s = {
      font: {
        name: "Arial",
        sz: 10,
        ...(bold ? { bold: true } : {}),
      },
    };
    ws[address] = cell;
  };

  setCellValue(worksheet, A3_LABEL_CELLS.company, A3_HEADER_LABELS.company, true);
  setCellValue(worksheet, A3_LABEL_CELLS.selection, A3_HEADER_LABELS.selection, true);
  setCellValue(worksheet, A3_LABEL_CELLS.date, A3_HEADER_LABELS.date, true);
  setCellValue(worksheet, A3_METADATA_CELLS.companyCode, data.metadata.companyCode, true);
  setCellValue(worksheet, A3_METADATA_CELLS.companyName, data.metadata.companyName, true);
  setCellValue(worksheet, A3_METADATA_CELLS.selection, data.metadata.selection, true);
  setCellValue(worksheet, A3_METADATA_CELLS.exportDate, data.metadata.exportDate, true);
  setCellValue(
    worksheet,
    A3_METADATA_CELLS.sectionTitle,
    data.metadata.sectionTitle ?? DEFAULT_A3_SECTION_TITLE,
    true
  );

  const lastColIndex = firstColIndex + data.columns.length - 1;
  data.columns.forEach((column, columnIndex) => {
    const address = XLSX.utils.encode_cell({
      c: firstColIndex + columnIndex,
      r: headerRowIndex,
    });
    setCellValue(worksheet, address, column, true);
  });

  data.rows.forEach((row, rowIndex) => {
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

  worksheet["!ref"] = XLSX.utils.encode_range({
    s: { c: 0, r: 0 },
    e: { c: lastColIndex, r: lastRowIndex },
  });

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    (data.sheetName || "Datos").replace(/[:\\/?*[\]]/g, " ").trim().slice(0, 31)
  );

  const base = options?.sourceFileName
    ? options.sourceFileName.replace(/\.(xls|xlsx)$/i, "")
    : data.sheetName?.trim() || "exportacion-a3";

  XLSX.writeFile(workbook, `${base}-editable.xlsx`, {
    bookType: "xlsx",
    compression: true,
  });
}
