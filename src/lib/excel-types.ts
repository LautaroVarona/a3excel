export type ParsePhase =
  | "reading"
  | "parsing"
  | "converting"
  | "indexing"
  | "complete";

export interface ParseProgress {
  phase: ParsePhase;
  message: string;
  percent: number;
  processed: number;
  total: number;
  elapsedMs: number;
  estimatedRemainingMs: number | null;
}

export type ExcelRow = Record<string, string | number | boolean | null>;

export interface ExcelExportMetadata {
  companyCode: string | null;
  companyName: string | null;
  selection: string | null;
  exportDate: string | null;
}

export interface ParsedExcel {
  sheetName: string;
  columns: string[];
  rows: ExcelRow[];
  totalRows: number;
  metadata: ExcelExportMetadata;
}

export const PHASE_LABELS: Record<ParsePhase, string> = {
  reading: "Lectura del archivo",
  parsing: "Análisis del libro Excel",
  converting: "Conversión de filas",
  indexing: "Indexación de columnas",
  complete: "Completado",
};

export const MAX_FILE_SIZE_BYTES = 10 * 1024 * 1024;
export const PARSE_TIMEOUT_MS = 60_000;
