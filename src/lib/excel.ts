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

export interface ParsedExcel {
  sheetName: string;
  columns: string[];
  rows: ExcelRow[];
  totalRows: number;
}

const PHASE_LABELS: Record<ParsePhase, string> = {
  reading: "Lectura del archivo",
  parsing: "Análisis del libro Excel",
  converting: "Conversión de filas",
  indexing: "Indexación de columnas",
  complete: "Completado",
};

export const MAX_FILE_SIZE_BYTES = 10 * 1024 * 1024;
export const PARSE_TIMEOUT_MS = 60_000;
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
      const loaded = event.loaded;
      const total = event.total;
      emitProgress(onProgress, startTime, {
        phase: "reading",
        message: `Leyendo ${formatBytes(loaded)} de ${formatBytes(total)}…`,
        percent: Math.round((loaded / total) * 25),
        processed: loaded,
        total,
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

function parseBufferInWorker(
  buffer: ArrayBuffer,
  file: File,
  onProgress: (progress: ParseProgress) => void,
  startTime: number
): Promise<ParsedExcel> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL("./excel-worker.ts", import.meta.url));

    const cleanup = () => {
      worker.terminate();
    };

    worker.onmessage = (event: MessageEvent) => {
      const data = event.data as
        | { type: "progress"; progress: ParseProgress }
        | { type: "result"; data: ParsedExcel }
        | { type: "error"; message: string };

      if (data.type === "progress") {
        onProgress(data.progress);
        return;
      }

      cleanup();
      if (data.type === "result") {
        resolve(data.data);
        return;
      }

      reject(new Error(data.message));
    };

    worker.onerror = () => {
      cleanup();
      reject(
        new Error(
          "Error al procesar el archivo. Probá con un archivo .xlsx o uno más simple."
        )
      );
    };

    worker.postMessage({
      type: "parse",
      buffer,
      fileName: file.name,
      fileSize: file.size,
      startTime,
    });
  });
}

export async function parseExcelFileWithProgress(
  file: File,
  onProgress: (progress: ParseProgress) => void
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
    return parseBufferInWorker(buffer, file, onProgress, startTime);
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

export { PHASE_LABELS };
