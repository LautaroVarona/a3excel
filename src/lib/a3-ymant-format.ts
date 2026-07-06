/**
 * @deprecated Importar desde `@/lib/a3-workbook-layout`.
 */
export {
  detectA3ExportLayout as detectYmantLayout,
  parseA3ExportWorksheet as parseYmantWorksheet,
  extractA3Metadata as extractYmantMetadata,
  isA3ControlCode as isYmantControlCode,
} from "./a3-workbook-layout";

export { readCellDisplay } from "./a3-layout-discovery.js";

/** @deprecated El código de control se descubre escaneando la hoja. */
export const A3_YMANT_CONTROL_CELL = "B8";
