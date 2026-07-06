import type { ParsedExcel } from "./excel-types";
import { sanitizeA3ExportBaseName } from "./a3-export-filename";

type XlsxModule = typeof import("xlsx-js-style");
type WorkSheet = import("xlsx-js-style").WorkSheet;

function buildExcelEditingFileName(sourceFileName?: string | null): string {
  return `${sanitizeA3ExportBaseName(sourceFileName)}_excel.xlsx`;
}

function setCell(
  ws: WorkSheet,
  XLSX: XlsxModule,
  col: number,
  row: number,
  value: string | number | null | undefined,
  bold = false
): void {
  if (value === null || value === undefined || value === "") return;
  const address = XLSX.utils.encode_cell({ c: col, r: row });
  const cell =
    typeof value === "number"
      ? { t: "n" as const, v: value }
      : { t: "s" as const, v: String(value) };
  ws[address] = {
    ...cell,
    s: {
      font: { name: "Arial", sz: 10, ...(bold ? { bold: true } : {}) },
    },
  };
}

function buildYmantWorksheetForExcel(
  data: ParsedExcel,
  XLSX: XlsxModule
): WorkSheet {
  const layout = data.layout;
  if (!layout || layout.kind !== "ymant") {
    throw new Error("Solo se puede exportar para Excel en formato YMANT.");
  }

  const ws: WorkSheet = {};
  const { metadata } = data;

  setCell(
    ws,
    XLSX,
    layout.controlCodeColIndex ?? 1,
    layout.controlCodeRowIndex ?? 7,
    layout.controlCode
  );
  setCell(ws, XLSX, 2, 8, "EMPRESA:", true);
  setCell(ws, XLSX, 3, 8, metadata.companyCode);
  setCell(ws, XLSX, 4, 8, metadata.companyName);
  setCell(ws, XLSX, 2, 9, "SELECCIÓN:", true);
  setCell(ws, XLSX, 4, 9, metadata.selection);
  setCell(ws, XLSX, 2, 10, "FECHA:", true);
  setCell(ws, XLSX, 3, 10, metadata.exportDate);
  setCell(ws, XLSX, 1, 11, metadata.sectionTitle ?? "Trabajadores", true);

  const headerRowIndex = layout.headerRow1Based - 1;
  for (const column of data.columns) {
    const colIndex = layout.columnIndices[column];
    if (colIndex === undefined) continue;
    setCell(ws, XLSX, colIndex, headerRowIndex, column, true);
  }

  data.rows.forEach((row, rowIndex) => {
    const excelRowIndex = layout.dataStartRow1Based - 1 + rowIndex;
    for (const column of data.columns) {
      const colIndex = layout.columnIndices[column];
      if (colIndex === undefined) continue;
      setCell(ws, XLSX, colIndex, excelRowIndex, row[column] ?? null);
    }
  });

  const lastCol = Math.max(...Object.values(layout.columnIndices), 1);
  const lastRow = Math.max(
    layout.dataStartRow1Based - 1 + data.rows.length - 1,
    headerRowIndex
  );

  ws["!ref"] = XLSX.utils.encode_range({
    s: { c: 0, r: 0 },
    e: { c: lastCol, r: lastRow },
  });

  ws["!autofilter"] = {
    ref: XLSX.utils.encode_range({
      s: { c: 1, r: headerRowIndex },
      e: { c: lastCol, r: lastRow },
    }),
  };

  ws["!freeze"] = {
    xSplit: 1,
    ySplit: layout.dataStartRow1Based,
    topLeftCell: XLSX.utils.encode_cell({ c: 1, r: layout.dataStartRow1Based - 1 }),
  };

  ws["!cols"] = [{ wch: 3 }, ...data.columns.map((c) => ({ wch: Math.min(42, c.length + 2) }))];

  return ws;
}

export async function exportForExcelEditing(
  data: ParsedExcel,
  sourceFileName?: string | null
): Promise<void> {
  const XLSX = await import("xlsx-js-style");
  const worksheet = buildYmantWorksheetForExcel(data, XLSX);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    (data.sheetName || "Conceptos").replace(/[:\\/?*[\]]/g, " ").trim().slice(0, 31)
  );

  XLSX.writeFile(workbook, buildExcelEditingFileName(sourceFileName), {
    bookType: "xlsx",
    compression: true,
  });
}

export { buildExcelEditingFileName };
