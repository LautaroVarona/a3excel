/* global XLSX */

const ROW_CHUNK_SIZE = 500;
const MAX_ROWS = 50_000;
/** Índice 0-based de la fila de encabezados (fila 8 en Excel). */
const EXCEL_HEADER_ROW_INDEX = 7;

const A3_YMANT_CONTROL_CELL = "B8";
const CONTROL_CODE_ROW_INDEX = 7;
const CONTROL_CODE_PATTERN = /^077\d+/;
const CONTROL_CODE_LOOSE_PATTERN = /^077\d+.*\bCT\s+1/i;

function normalizeAccents(value) {
  return value.normalize("NFD").replace(/\p{M}/gu, "");
}

function isYmantControlCode(value) {
  if (!value) return false;
  const text = value.trim();
  if (CONTROL_CODE_PATTERN.test(text)) return true;
  if (CONTROL_CODE_LOOSE_PATTERN.test(text)) return true;
  return false;
}

function isYmantHeaderLabel(value) {
  if (!value) return false;
  const text = normalizeAccents(value.trim());
  return /^codigo\b/i.test(text) || /^cod\b\.?/i.test(text);
}

function readCellDisplay(worksheet, address) {
  const cell = worksheet[address];
  if (!cell) return null;

  if (cell.w != null && cell.w !== "") {
    return String(cell.w);
  }

  const value = cell.v;
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) {
    return value.toLocaleDateString("es-ES");
  }
  return String(value);
}

function readCellAt(worksheet, colIndex, rowIndex) {
  const address = XLSX.utils.encode_cell({ c: colIndex, r: rowIndex });
  return readCellDisplay(worksheet, address);
}

function detectYmantControlCode(worksheet) {
  const preferred = readCellDisplay(worksheet, A3_YMANT_CONTROL_CELL);
  if (isYmantControlCode(preferred)) {
    return {
      controlCode: preferred.trim(),
      controlCodeColIndex: 1,
      controlCodeRowIndex: CONTROL_CODE_ROW_INDEX,
    };
  }

  for (let rowIndex = CONTROL_CODE_ROW_INDEX - 1; rowIndex <= CONTROL_CODE_ROW_INDEX + 1; rowIndex++) {
    for (let colIndex = 0; colIndex <= 4; colIndex++) {
      const value = readCellAt(worksheet, colIndex, rowIndex);
      if (!isYmantControlCode(value)) continue;
      return {
        controlCode: value.trim(),
        controlCodeColIndex: colIndex,
        controlCodeRowIndex: rowIndex,
      };
    }
  }

  return null;
}

function findYmantHeaderRow(worksheet) {
  for (let rowIndex = CONTROL_CODE_ROW_INDEX; rowIndex <= 30; rowIndex++) {
    for (let colIndex = 0; colIndex <= 5; colIndex++) {
      const label = readCellAt(worksheet, colIndex, rowIndex);
      if (isYmantHeaderLabel(label)) {
        return rowIndex;
      }
    }
  }
  return null;
}

function parseMaxCol(ref) {
  const match = ref.match(/:([A-Z]+)(\d+)$/i);
  if (!match) return 120;
  const letters = match[1].toUpperCase();
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col - 1;
}

function detectYmantLayout(worksheet) {
  const control = detectYmantControlCode(worksheet);
  if (!control) return null;

  const headerRowIndex = findYmantHeaderRow(worksheet);
  if (headerRowIndex === null) return null;

  const maxCol = worksheet["!ref"] ? parseMaxCol(worksheet["!ref"]) : 120;
  const columnIndices = {};

  for (let colIndex = 0; colIndex <= maxCol; colIndex++) {
    const header = readCellAt(worksheet, colIndex, headerRowIndex);
    if (!header) continue;
    const key = header.trim();
    if (!key || isYmantControlCode(key)) continue;
    columnIndices[key] = colIndex;
  }

  const columns = Object.keys(columnIndices);
  if (columns.length === 0) return null;

  return {
    kind: "ymant",
    controlCode: control.controlCode,
    controlCodeColIndex: control.controlCodeColIndex,
    controlCodeRowIndex: control.controlCodeRowIndex,
    headerRow1Based: headerRowIndex + 1,
    dataStartRow1Based: headerRowIndex + 2,
    columnIndices,
  };
}

function extractYmantMetadata(worksheet) {
  return {
    companyCode: readCellAt(worksheet, 3, 8),
    companyName: readCellAt(worksheet, 4, 8),
    selection: readCellAt(worksheet, 4, 9),
    exportDate: readCellAt(worksheet, 3, 10),
    sectionTitle: readCellAt(worksheet, 1, 11),
  };
}

function parseYmantWorksheet(worksheet, layout, sheetName) {
  const columns = Object.keys(layout.columnIndices);
  const metadata = extractYmantMetadata(worksheet);
  const dataStartIndex = layout.dataStartRow1Based - 1;
  const maxRow = dataStartIndex + MAX_ROWS;
  const rows = [];

  for (let rowIndex = dataStartIndex; rowIndex < maxRow; rowIndex++) {
    const row = {};
    let hasValue = false;

    for (const column of columns) {
      const colIndex = layout.columnIndices[column];
      const raw = readCellAt(worksheet, colIndex, rowIndex);
      const value = normalizeCellValue(raw);
      row[column] = value;
      if (value !== null) hasValue = true;
    }

    if (hasValue) {
      rows.push(row);
    } else if (rows.length > 0) {
      break;
    }
  }

  return {
    sheetName,
    columns,
    rows,
    totalRows: rows.length,
    metadata,
    layout,
  };
}

const A3_METADATA_CELLS = {
  companyCode: "D2",
  companyName: "E2",
  selection: "D3",
  exportDate: "D5",
  sectionTitle: "B7",
};

const DEFAULT_A3_SECTION_TITLE = "Trabajadores";

function extractA3Metadata(worksheet) {
  return {
    companyCode: readCellDisplay(worksheet, A3_METADATA_CELLS.companyCode),
    companyName: readCellDisplay(worksheet, A3_METADATA_CELLS.companyName),
    selection: readCellDisplay(worksheet, A3_METADATA_CELLS.selection),
    exportDate: readCellDisplay(worksheet, A3_METADATA_CELLS.exportDate),
    sectionTitle:
      readCellDisplay(worksheet, A3_METADATA_CELLS.sectionTitle) ??
      DEFAULT_A3_SECTION_TITLE,
  };
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
  const ymantLayout = detectYmantLayout(worksheet);

  if (ymantLayout) {
    const ymantData = parseYmantWorksheet(worksheet, ymantLayout, sheetName);

    emitProgress(jobId, startTime, {
      phase: "complete",
      message: `Listo — ${ymantData.totalRows.toLocaleString("es-ES")} trabajadores (formato YMANT 77)`,
      percent: 100,
      processed: ymantData.totalRows,
      total: ymantData.totalRows,
    });

    return ymantData;
  }

  const metadata = extractA3Metadata(worksheet);
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
