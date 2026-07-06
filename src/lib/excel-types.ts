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

export interface A3PreambleCell {
  col: number;
  row: number;
  text: string;
}

/** Layout descubierto de un export nativo A3NOM (.XLS cifrado u homólogo). */
export interface A3ExportLayout {
  kind: "a3-export";
  controlCode: string;
  controlCodeColIndex: number;
  controlCodeRowIndex: number;
  headerRow1Based: number;
  dataStartRow1Based: number;
  /** Etiqueta de columna → índice de columna 0-based en Excel. */
  columnIndices: Record<string, number>;
  /** Bloque superior del export (metadatos, etiquetas) tal como vino del archivo. */
  preambleCells?: A3PreambleCell[];
}

/** @deprecated Usar A3ExportLayout */
export type A3YmantLayout = A3ExportLayout;

export interface ParsedExcel {
  sheetName: string;
  columns: string[];
  rows: ExcelRow[];
  totalRows: number;
  metadata: ExcelExportMetadata;
  /** Snapshot inicial para detectar cambios al exportar. */
  originalRows?: ExcelRow[];
  /** Presente cuando el archivo coincide con un export nativo de A3. */
  layout?: A3ExportLayout;
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
