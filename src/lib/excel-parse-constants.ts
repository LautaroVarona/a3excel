/** Índice 0-based de la fila de encabezados en exports a3ERP (fila 8 en Excel). */
export const EXCEL_HEADER_ROW_INDEX = 7;

/** Primera columna de datos en exports a3ERP (B = índice 1). */
export const A3_DATA_FIRST_COL_INDEX = 1;

/** Fila 1-based del encabezado de columnas en exports a3ERP. */
export const A3_HEADER_ROW_1_BASED = 8;

/** Fila 1-based donde comienzan los registros en exports a3ERP. */
export const A3_DATA_START_ROW_1_BASED = 9;

/** Etiquetas fijas del encabezado a3ERP. */
export const A3_HEADER_LABELS = {
  company: "EMPRESA:",
  selection: "SELECCIÓN:",
  date: "FECHA:",
} as const;

/** Celdas de metadatos del encabezado a3ERP. */
export const A3_METADATA_CELLS = {
  companyCode: "D2",
  companyName: "E2",
  selection: "D3",
  exportDate: "D5",
  sectionTitle: "B7",
} as const;

/** Celdas de etiquetas del encabezado a3ERP. */
export const A3_LABEL_CELLS = {
  company: "C2",
  selection: "C3",
  date: "C5",
} as const;

export const DEFAULT_A3_SECTION_TITLE = "Trabajadores";
