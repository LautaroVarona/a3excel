import type { ParseProgress, ParsedExcel } from "./excel-types";
import {
  MAX_FILE_SIZE_BYTES,
  PARSE_TIMEOUT_MS,
  PHASE_LABELS,
} from "./excel-types";

const SERVER_PARSE_TIMEOUT_MS = 120_000;

export type { ExcelRow, ParsePhase, ParseProgress, ParsedExcel } from "./excel-types";
export { MAX_FILE_SIZE_BYTES, PARSE_TIMEOUT_MS, PHASE_LABELS };

const VALID_EXTENSIONS = [".xls", ".xlsx"];

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
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
  onProgress: (p: ParseProgress) => void,
  startTime: number,
  partial: Omit<ParseProgress, "elapsedMs" | "estimatedRemainingMs">
) {
  const elapsedMs = Date.now() - startTime;
  onProgress({
    ...partial,
    elapsedMs,
    estimatedRemainingMs: estimateRemaining(
      partial.processed,
      partial.total,
      elapsedMs
    ),
  });
}

export function flushUI(): Promise<void> {
  return new Promise((resolve) => {
    requestAnimationFrame(() => {
      requestAnimationFrame(() => resolve());
    });
  });
}

function isEncryptionErrorMessage(message: string): boolean {
  const text = message.toLowerCase();
  return (
    text.includes("protección") ||
    text.includes("cifrado") ||
    text.includes("password") ||
    text.includes("encrypt")
  );
}

async function parseViaLocalServer(
  file: File,
  onProgress: (p: ParseProgress) => void,
  startTime: number,
  password?: string
): Promise<ParsedExcel> {
  emitProgress(onProgress, startTime, {
    phase: "parsing",
    message:
      "El archivo está protegido — descifrando en el servidor (sin Excel local)…",
    percent: 35,
    processed: 0,
    total: 1,
  });

  const formData = new FormData();
  formData.append("file", file);
  if (password) {
    formData.append("password", password);
  }

  const controller = new AbortController();
  const timer = setTimeout(
    () => controller.abort(),
    SERVER_PARSE_TIMEOUT_MS
  );

  try {
    const response = await fetch("/api/parse-excel", {
      method: "POST",
      body: formData,
      signal: controller.signal,
    });

    const payload = (await response.json()) as ParsedExcel & { error?: string };

    if (!response.ok) {
      throw new Error(
        payload.error ||
          "No se pudo abrir el archivo con Excel en esta computadora."
      );
    }

    emitProgress(onProgress, startTime, {
      phase: "converting",
      message: `Listo — ${payload.totalRows.toLocaleString("es-ES")} filas extraídas`,
      percent: 95,
      processed: payload.totalRows,
      total: payload.totalRows,
    });

    return payload;
  } catch (error) {
    if (error instanceof Error && error.name === "AbortError") {
      throw new Error(
        "Excel tardó demasiado en abrir el archivo. Probá de nuevo o guardá una copia .xlsx."
      );
    }
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

export function validateExcelFile(file: File): string | null {
  if (file.size <= 0) {
    return "El archivo está vacío.";
  }

  if (file.size > MAX_FILE_SIZE_BYTES) {
    return `El archivo supera el límite de ${formatBytes(MAX_FILE_SIZE_BYTES)}.`;
  }

  const dotIndex = file.name.lastIndexOf(".");
  if (dotIndex === -1) {
    return "Formato no válido. Solo se admiten archivos .XLS y .XLSX.";
  }

  const extension = file.name.slice(dotIndex).toLowerCase();
  if (!VALID_EXTENSIONS.includes(extension)) {
    return "Formato no válido. Solo se admiten archivos .XLS y .XLSX.";
  }

  return null;
}

function withTimeout<T>(
  promise: Promise<T>,
  timeoutMs: number,
  message: string
): Promise<T> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(message)), timeoutMs);
    promise
      .then((value) => {
        clearTimeout(timer);
        resolve(value);
      })
      .catch((error: unknown) => {
        clearTimeout(timer);
        reject(error instanceof Error ? error : new Error(message));
      });
  });
}

function readFileWithProgress(
  file: File,
  onProgress: (p: ParseProgress) => void,
  startTime: number
): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onprogress = (event) => {
      if (!event.lengthComputable) return;
      emitProgress(onProgress, startTime, {
        phase: "reading",
        message: `Leyendo ${formatBytes(event.loaded)} de ${formatBytes(event.total)}…`,
        percent: Math.round((event.loaded / event.total) * 25),
        processed: event.loaded,
        total: event.total,
      });
    };

    reader.onload = () => {
      emitProgress(onProgress, startTime, {
        phase: "reading",
        message: "Archivo leído en memoria",
        percent: 25,
        processed: file.size,
        total: file.size,
      });
      resolve(reader.result as ArrayBuffer);
    };

    reader.onerror = () =>
      reject(new Error("No se pudo leer el archivo desde el disco."));
    reader.readAsArrayBuffer(file);
  });
}

export async function parseExcelFileWithProgress(
  file: File,
  onProgress: (progress: ParseProgress) => void,
  parseInWorker: (
    buffer: ArrayBuffer,
    startTime: number,
    onProgress: (progress: ParseProgress) => void,
    password?: string
  ) => Promise<ParsedExcel>,
  password?: string
): Promise<ParsedExcel> {
  const validationError = validateExcelFile(file);
  if (validationError) {
    throw new Error(validationError);
  }

  const startTime = Date.now();

  emitProgress(onProgress, startTime, {
    phase: "reading",
    message: `Preparando lectura de ${file.name} (${formatBytes(file.size)})…`,
    percent: 0,
    processed: 0,
    total: file.size,
  });

  const parseTask = async (): Promise<ParsedExcel> => {
    const buffer = await readFileWithProgress(file, onProgress, startTime);

    try {
      return await parseInWorker(buffer, startTime, onProgress, password);
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Error al procesar el archivo.";

      if (typeof window !== "undefined" && isEncryptionErrorMessage(message)) {
        return parseViaLocalServer(file, onProgress, startTime, password);
      }

      throw error;
    }
  };

  return withTimeout(
    parseTask(),
    PARSE_TIMEOUT_MS,
    "El archivo tardó demasiado. Probá con .xlsx o un archivo más simple."
  );
}

export function formatCellValue(
  value: string | number | boolean | null
): string {
  if (value === null) return "—";
  if (typeof value === "boolean") return value ? "Sí" : "No";
  return String(value);
}

export function formatDuration(ms: number): string {
  if (ms < 1000) return "< 1 s";
  const seconds = Math.ceil(ms / 1000);
  if (seconds < 60) return `~${seconds} s`;
  const minutes = Math.floor(seconds / 60);
  const rem = seconds % 60;
  return rem > 0 ? `~${minutes} min ${rem} s` : `~${minutes} min`;
}
