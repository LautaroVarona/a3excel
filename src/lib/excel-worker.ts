/// <reference lib="webworker" />

import * as XLSX from "xlsx";

import type { ExcelRow, ParseProgress, ParsedExcel } from "./excel";

const ROW_CHUNK_SIZE = 2000;

type WorkerRequest = {
  type: "parse";
  buffer: ArrayBuffer;
  fileName: string;
  fileSize: number;
  startTime: number;
};

type WorkerResponse =
  | { type: "progress"; progress: ParseProgress }
  | { type: "result"; data: ParsedExcel }
  | { type: "error"; message: string };

function normalizeCellValue(value: unknown): string | number | boolean | null {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "boolean" || typeof value === "number") return value;
  return String(value);
}

function estimateRemaining(
  processed: number,
  total: number,
  elapsedMs: number
): number | null {
  if (processed <= 0 || total <= 0 || processed >= total) return null;
  const rate = processed / elapsedMs;
  if (rate <= 0) return null;
  return Math.round((total - processed) / rate);
}

function emitProgress(
  startTime: number,
  partial: Omit<ParseProgress, "elapsedMs" | "estimatedRemainingMs">
) {
  const elapsedMs = Date.now() - startTime;
  const progress: ParseProgress = {
    ...partial,
    elapsedMs,
    estimatedRemainingMs: estimateRemaining(
      partial.processed,
      partial.total,
      elapsedMs
    ),
  };
  postMessage({ type: "progress", progress } satisfies WorkerResponse);
}

function yieldToMain(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

async function parseBuffer(
  buffer: ArrayBuffer,
  fileName: string,
  fileSize: number,
  startTime: number
): Promise<ParsedExcel> {
  emitProgress(startTime, {
    phase: "parsing",
    message: "Decodificando estructura del libro Excel…",
    percent: 30,
    processed: 0,
    total: 1,
  });

  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error("El archivo no contiene hojas de cálculo.");
  }

  await yieldToMain();

  emitProgress(startTime, {
    phase: "parsing",
    message: `Hoja activa: «${sheetName}» — extrayendo celdas…`,
    percent: 40,
    processed: 1,
    total: 1,
  });

  const worksheet = workbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(
    worksheet,
    { defval: null, raw: false }
  );

  const totalRows = rawRows.length;
  const rows: ExcelRow[] = [];

  emitProgress(startTime, {
    phase: "converting",
    message: `Convirtiendo ${totalRows.toLocaleString("es-ES")} filas a formato tabular…`,
    percent: 45,
    processed: 0,
    total: totalRows,
  });

  await yieldToMain();

  for (let i = 0; i < totalRows; i += ROW_CHUNK_SIZE) {
    const end = Math.min(i + ROW_CHUNK_SIZE, totalRows);

    for (let j = i; j < end; j++) {
      const raw = rawRows[j];
      const normalized: ExcelRow = {};
      for (const [key, value] of Object.entries(raw)) {
        normalized[String(key)] = normalizeCellValue(value);
      }
      rows.push(normalized);
    }

    const processed = end;
    const percent = 45 + Math.round((processed / Math.max(totalRows, 1)) * 40);

    emitProgress(startTime, {
      phase: "converting",
      message: `Filas procesadas: ${processed.toLocaleString("es-ES")} / ${totalRows.toLocaleString("es-ES")}`,
      percent,
      processed,
      total: totalRows,
    });

    await yieldToMain();
  }

  emitProgress(startTime, {
    phase: "indexing",
    message: "Detectando columnas…",
    percent: 88,
    processed: 0,
    total: rows.length > 0 ? Object.keys(rows[0]).length : 0,
  });

  await yieldToMain();

  const columnSet = new Set<string>();
  for (const row of rows) {
    for (const key of Object.keys(row)) {
      columnSet.add(key);
    }
  }
  const columns = Array.from(columnSet);

  emitProgress(startTime, {
    phase: "indexing",
    message: `${columns.length} columnas detectadas`,
    percent: 95,
    processed: columns.length,
    total: columns.length,
  });

  await yieldToMain();

  emitProgress(startTime, {
    phase: "complete",
    message: `Listo — ${totalRows.toLocaleString("es-ES")} filas en «${sheetName}»`,
    percent: 100,
    processed: totalRows,
    total: totalRows,
  });

  void fileName;
  void fileSize;

  return {
    sheetName,
    columns,
    rows,
    totalRows,
  };
}

self.onmessage = async (event: MessageEvent<WorkerRequest>) => {
  const message = event.data;
  if (message.type !== "parse") return;

  const startTime = message.startTime;

  try {
    const data = await parseBuffer(
      message.buffer,
      message.fileName,
      message.fileSize,
      startTime
    );
    postMessage({ type: "result", data } satisfies WorkerResponse);
  } catch (err) {
    postMessage({
      type: "error",
      message:
        err instanceof Error
          ? err.message
          : "No se pudo procesar el archivo Excel.",
    } satisfies WorkerResponse);
  }
};
