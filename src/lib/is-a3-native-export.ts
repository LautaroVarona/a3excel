import type { A3ExportLayout } from "./excel-types";

/** Export .XLS nativo de A3NOM (detectado por layout, no por nombre de formato). */
export function isA3NativeExportLayout(
  layout?: { kind?: string } | null
): layout is A3ExportLayout {
  return layout?.kind === "a3-export" || layout?.kind === "ymant";
}
