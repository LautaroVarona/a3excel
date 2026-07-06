import type { ParsedExcel } from "./excel-types";

const INVALID_SHEET_CHARS = /[:\\/?*[\]]/g;

function sanitizeSheetName(name: string): string {
  const cleaned = name.replace(INVALID_SHEET_CHARS, " ").trim() || "Datos";
  return cleaned.slice(0, 31);
}

function buildOutputFileName(sourceFileName?: string | null, sheetName?: string): string {
  const base = sourceFileName
    ? sourceFileName.replace(/\.(xls|xlsx)$/i, "")
    : sheetName?.trim() || "exportacion-a3";
  return `${base}-editable.xlsx`;
}

export async function exportParsedExcelToFile(
  data: ParsedExcel,
  options?: { sourceFileName?: string | null }
): Promise<void> {
  const XLSX = await import("xlsx");

  const worksheet = XLSX.utils.json_to_sheet(data.rows, {
    header: data.columns,
    skipHeader: false,
  });

  if (worksheet["!ref"]) {
    worksheet["!autofilter"] = { ref: worksheet["!ref"] };
    worksheet["!views"] = [
      {
        state: "frozen",
        ySplit: 1,
        topLeftCell: "A2",
        activeCell: "A2",
      },
    ];
  }

  worksheet["!cols"] = data.columns.map((column) => ({
    wch: Math.min(48, Math.max(10, column.length + 2)),
  }));

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    sanitizeSheetName(data.sheetName)
  );

  const outputName = buildOutputFileName(options?.sourceFileName, data.sheetName);
  XLSX.writeFile(workbook, outputName, { bookType: "xlsx", compression: true });
}
