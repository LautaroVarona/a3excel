/** Caracteres no admitidos por A3NOM en nombres de archivo de importación. */
const INVALID_A3_NAME_CHARS = /[^A-Za-z0-9_]/g;

export function sanitizeA3ExportBaseName(sourceFileName?: string | null): string {
  const raw = sourceFileName
    ? sourceFileName.replace(/\.(xls|xlsx)$/i, "")
    : "exportacion_a3";

  const withoutCopySuffix = raw.replace(/\s*\(\d+\)$/i, "");

  const cleaned = withoutCopySuffix
    .replace(INVALID_A3_NAME_CHARS, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");

  return cleaned || "exportacion_a3";
}

/** Nombre de export compatible con A3: solo .xls, sin espacios, guiones ni paréntesis. */
export function buildEditableExportName(sourceFileName?: string | null): string {
  return `${sanitizeA3ExportBaseName(sourceFileName)}.xls`;
}
