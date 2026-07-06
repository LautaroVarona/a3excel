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
  sectionTitle: string | null;
}

export interface A3YmantLayout {
  kind: "ymant";
  controlCode: string;
  headerRow1Based: number;
  dataStartRow1Based: number;
  /** Etiqueta de columna → índice de columna 0-based en Excel. */
  columnIndices: Record<string, number>;
}

export interface ParsedExcel {
  sheetName: string;
  columns: string[];
  rows: ExcelRow[];
  totalRows: number;
  metadata: ExcelExportMetadata;
  /** Presente en exports A3NOM formato 77 (YMANT). */
  layout?: A3YmantLayout;
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
