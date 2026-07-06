/** Índice 0-based de la fila de encabezados en exports a3ERP (fila 8 en Excel). */
export const EXCEL_HEADER_ROW_INDEX = 7;

/** Celdas de metadatos del encabezado a3ERP (valores, no etiquetas). */
export const A3_METADATA_CELLS = {
  companyCode: "D2",
  companyName: "E2",
  selection: "D3",
  exportDate: "D5",
} as const;
