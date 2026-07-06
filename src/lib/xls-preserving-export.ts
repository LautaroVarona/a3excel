import {
  collectChangedNumericPatches,
  patchBiffWorkbookStream,
} from "./biff-patch";
import { buildDecryptPasswordCandidates } from "./decrypt-workbook-buffer";
import type { ParsedExcel } from "./excel-types";
import {
  exportPreservingXlsBuffer,
  resolvePlainWorkbookForExport,
} from "./xls97-decrypt";

export function exportYmantPreservingBuffer(
  input: Buffer,
  data: ParsedExcel,
  password?: string
): Buffer {
  const layout = data.layout;
  if (!layout || layout.kind !== "ymant") {
    throw new Error("Solo se admite exportación preservada para formato YMANT.");
  }

  const candidates = buildDecryptPasswordCandidates(password);
  const { plainWorkbook, workingCopy, password: resolvedPassword } =
    resolvePlainWorkbookForExport(input, candidates);

  const patches = collectChangedNumericPatches(
    layout,
    data.originalRows ?? data.rows,
    data.rows
  );
  if (patches.length === 0) {
    return input;
  }

  const patchedWorkbook = patchBiffWorkbookStream(plainWorkbook, patches);

  if (workingCopy === input) {
    return exportPreservingXlsBuffer(input, patchedWorkbook, resolvedPassword ?? " ");
  }

  return exportPreservingXlsBuffer(
    input,
    patchedWorkbook,
    resolvedPassword ?? candidates[0] ?? "VelvetSweatshop"
  );
}

export function buildEditableExportName(sourceFileName?: string | null): string {
  const base = sourceFileName
    ? sourceFileName.replace(/\.(xls|xlsx)$/i, "")
    : "exportacion-a3";
  return `${base}-editable.xls`;
}
