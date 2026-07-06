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

const ROW_CHUNK_SIZE = 2000;

function normalizeCellValue(value: unknown): string | number | boolean | null {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "boolean" || typeof value === "number") return value;
  return String(value);
}

function yieldToMain(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

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

export async function parseExcelFileWithProgress(
  file: File,
  onProgress: (progress: ParseProgress) => void
): Promise<ParsedExcel> {
  const startTime = Date.now();
  const XLSX = await import("xlsx");

  emitProgress(onProgress, startTime, {
    phase: "reading",
    message: `Preparando lectura de ${file.name} (${formatBytes(file.size)})…`,
    percent: 0,
    processed: 0,
    total: file.size,
  });

  const buffer = await readFileWithProgress(file, onProgress, startTime);
  await yieldToMain();

  emitProgress(onProgress, startTime, {
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

  emitProgress(onProgress, startTime, {
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

  emitProgress(onProgress, startTime, {
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

    emitProgress(onProgress, startTime, {
      phase: "converting",
      message: `Filas procesadas: ${processed.toLocaleString("es-ES")} / ${totalRows.toLocaleString("es-ES")}`,
      percent,
      processed,
      total: totalRows,
    });

    await yieldToMain();
  }

  emitProgress(onProgress, startTime, {
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

  emitProgress(onProgress, startTime, {
    phase: "indexing",
    message: `${columns.length} columnas detectadas`,
    percent: 95,
    processed: columns.length,
    total: columns.length,
  });

  await yieldToMain();

  emitProgress(onProgress, startTime, {
    phase: "complete",
    message: `Listo — ${totalRows.toLocaleString("es-ES")} filas en «${sheetName}»`,
    percent: 100,
    processed: totalRows,
    total: totalRows,
  });

  return {
    sheetName,
    columns,
    rows,
    totalRows,
  };
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
