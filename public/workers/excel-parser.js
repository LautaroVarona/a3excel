/* global XLSX */

const ROW_CHUNK_SIZE = 500;
const MAX_ROWS = 50_000;
/** Índice 0-based de fila de encabezados en exports genéricos (fallback). */
const EXCEL_HEADER_ROW_INDEX = 7;

const DEFAULT_A3_SECTION_TITLE = "Trabajadores";

let layoutDiscoveryReady = false;

function ensureLayoutDiscovery() {
  if (!layoutDiscoveryReady) {
    importScripts("/workers/a3-layout-discovery.js");
    layoutDiscoveryReady = true;
  }
}

function normalizeCellValue(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "boolean" || typeof value === "number") return value;
  return String(value);
}

function isValidColumnKey(key) {
  return key.length > 0 && !key.startsWith("__EMPTY");
}

function estimateRemaining(processed, total, elapsedMs) {
  if (processed <= 0 || total <= 0 || processed >= total) return null;
  const rate = processed / elapsedMs;
  if (rate <= 0) return null;
  return Math.round((total - processed) / rate);
}

function emitProgress(jobId, startTime, partial) {
  const elapsedMs = Date.now() - startTime;
  self.postMessage({
    type: "progress",
    jobId,
    progress: {
      ...partial,
      elapsedMs,
      estimatedRemainingMs: estimateRemaining(
        partial.processed,
        partial.total,
        elapsedMs
      ),
    },
  });
}

function yieldToMain() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

let xlsxReady = false;

function ensureXlsx() {
  if (!xlsxReady) {
    importScripts("/vendor/xlsx.full.min.js");
    xlsxReady = true;
  }
}

function isPasswordError(message) {
  const text = String(message || "").toLowerCase();
  return (
    text.includes("password") ||
    text.includes("encrypt") ||
    text.includes("unsupported password")
  );
}

function mapPasswordError() {
  return new Error(
    "Este archivo tiene protección de apertura o cifrado que no pudimos leer automáticamente. " +
      "En Excel suele abrirse en solo lectura aunque no pida contraseña. " +
      "Probá «Archivo → Guardar una copia» como .xlsx sin protección, " +
      "o ingresá la contraseña de apertura en el campo de abajo."
  );
}

function readWorkbook(buffer, password) {
  const base = { type: "array", cellDates: true, sheetRows: MAX_ROWS };
  const strategies = [];
  const seen = new Set();

  const add = (opts) => {
    const key = opts.password ?? "";
    if (seen.has(key)) return;
    seen.add(key);
    strategies.push({ ...base, ...opts });
  };

  if (password) add({ password });
  add({ password: "VelvetSweatshop" });
  add({ password: " " });
  add({});

  let lastError = null;

  for (const opts of strategies) {
    try {
      return XLSX.read(buffer, opts);
    } catch (err) {
      lastError = err;
      if (!isPasswordError(err.message)) {
        throw err;
      }
    }
  }

  throw mapPasswordError(lastError);
}

async function parseBuffer(buffer, jobId, startTime, password) {
  ensureXlsx();
  ensureLayoutDiscovery();

  emitProgress(jobId, startTime, {
    phase: "parsing",
    message: "Decodificando estructura del libro Excel…",
    percent: 30,
    processed: 0,
    total: 1,
  });

  await yieldToMain();

  const workbook = readWorkbook(buffer, password);
  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error("El archivo no contiene hojas de cálculo.");
  }

  await yieldToMain();

  emitProgress(jobId, startTime, {
    phase: "parsing",
    message: `Hoja activa: «${sheetName}» — extrayendo celdas…`,
    percent: 40,
    processed: 1,
    total: 1,
  });

  await yieldToMain();

  const worksheet = workbook.Sheets[sheetName];
  const a3Layout = A3LayoutDiscovery.detectA3ExportLayout(worksheet);

  if (a3Layout) {
    const a3Data = A3LayoutDiscovery.parseA3ExportWorksheet(
      worksheet,
      a3Layout,
      sheetName,
      MAX_ROWS
    );

    emitProgress(jobId, startTime, {
      phase: "complete",
      message: `Listo — ${a3Data.totalRows.toLocaleString("es-ES")} filas (export A3)`,
      percent: 100,
      processed: a3Data.totalRows,
      total: a3Data.totalRows,
    });

    return a3Data;
  }

  const metadata = A3LayoutDiscovery.extractA3Metadata(worksheet);
  if (!metadata.sectionTitle) {
    metadata.sectionTitle = DEFAULT_A3_SECTION_TITLE;
  }
  const rawRows = XLSX.utils.sheet_to_json(worksheet, {
    range: EXCEL_HEADER_ROW_INDEX,
    defval: null,
    raw: false,
    blankrows: false,
  });

  const totalRows = rawRows.length;
  const truncated = totalRows >= MAX_ROWS;
  const rows = [];

  emitProgress(jobId, startTime, {
    phase: "converting",
    message: `Convirtiendo ${totalRows.toLocaleString("es-ES")} filas…`,
    percent: 45,
    processed: 0,
    total: totalRows,
  });

  await yieldToMain();

  for (let i = 0; i < totalRows; i += ROW_CHUNK_SIZE) {
    const end = Math.min(i + ROW_CHUNK_SIZE, totalRows);

    for (let j = i; j < end; j++) {
      const raw = rawRows[j];
      const normalized = {};
      for (const [key, value] of Object.entries(raw)) {
        if (!isValidColumnKey(key)) continue;
        normalized[String(key)] = normalizeCellValue(value);
      }
      rows.push(normalized);
    }

    emitProgress(jobId, startTime, {
      phase: "converting",
      message: `Filas procesadas: ${end.toLocaleString("es-ES")} / ${totalRows.toLocaleString("es-ES")}`,
      percent: 45 + Math.round((end / Math.max(totalRows, 1)) * 40),
      processed: end,
      total: totalRows,
    });

    await yieldToMain();
  }

  const columnSet = new Set();
  for (const row of rows) {
    for (const key of Object.keys(row)) {
      if (isValidColumnKey(key)) columnSet.add(key);
    }
  }
  const columns = Array.from(columnSet);
  const filledRows = rows.map((row) => {
    const filled = {};
    for (const column of columns) {
      filled[column] = row[column] ?? null;
    }
    return filled;
  });

  emitProgress(jobId, startTime, {
    phase: "complete",
    message: truncated
      ? `Listo — ${totalRows.toLocaleString("es-ES")} filas (límite ${MAX_ROWS.toLocaleString("es-ES")})`
      : `Listo — ${totalRows.toLocaleString("es-ES")} filas en «${sheetName}»`,
    percent: 100,
    processed: totalRows,
    total: totalRows,
  });

  return { sheetName, columns, rows: filledRows, totalRows, metadata };
}

self.onmessage = async (event) => {
  const message = event.data;

  if (message.type === "ping") {
    try {
      ensureXlsx();
      self.postMessage({ type: "pong" });
    } catch {
      self.postMessage({
        type: "error",
        jobId: "init",
        message: "No se pudo cargar el motor de lectura Excel.",
      });
    }
    return;
  }

  if (message.type !== "parse") return;

  const { jobId, buffer, startTime, password } = message;

  try {
    const data = await parseBuffer(buffer, jobId, startTime, password);
    self.postMessage({ type: "result", jobId, data });
  } catch (err) {
    self.postMessage({
      type: "error",
      jobId,
      message:
        err instanceof Error
          ? err.message
          : "No se pudo procesar el archivo Excel.",
    });
  }
};
