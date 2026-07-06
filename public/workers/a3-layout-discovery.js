/**
 * Descubrimiento heurístico de exports .XLS/.XLSX nativos de A3NOM.
 * Sin celdas fijas: escanea el libro en busca de código de control, cabecera tabular y metadatos.
 */

const SCAN_MAX_ROW = 120;
const SCAN_MAX_COL = 200;
const MIN_HEADER_COLUMNS = 3;
const MIN_CONTROL_CODE_LENGTH = 18;

function normalizeAccents(value) {
  return value.normalize("NFD").replace(/\p{M}/gu, "");
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

function encodeCellAddress(colIndex, rowIndex) {
  let label = "";
  let n = colIndex + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return `${label}${rowIndex + 1}`;
}

function readCellAt(worksheet, colIndex, rowIndex) {
  const address = encodeCellAddress(colIndex, rowIndex);
  return readCellDisplay(worksheet, address);
}

function parseMaxCol(ref) {
  const match = ref.match(/:([A-Z]+)(\d+)$/i);
  if (!match) return SCAN_MAX_COL;
  const letters = match[1].toUpperCase();
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col - 1;
}

function parseMaxRow(ref) {
  const match = ref.match(/(\d+)$/i);
  if (!match) return SCAN_MAX_ROW;
  return Math.max(0, parseInt(match[1], 10) - 1);
}

function isA3ControlCode(value) {
  if (!value) return false;
  const text = value.trim();
  if (text.length < MIN_CONTROL_CODE_LENGTH) return false;
  if (/^\d{3}\d{12,}/.test(text)) return true;
  if (/^\d{3}.{12,}CT/i.test(text)) return true;
  return false;
}

function isMetadataLabel(value) {
  if (!value) return false;
  const text = normalizeAccents(String(value).trim().replace(/:+\s*$/, ""));
  return /^(empresa|seleccion|fecha)$/i.test(text);
}

function looksLikeColumnHeader(value) {
  if (!value) return false;
  const text = value.trim();
  if (text.length === 0 || text.length > 80) return false;
  if (isA3ControlCode(text)) return false;
  if (isMetadataLabel(text)) return false;
  if (/^\d{1,6}([.,]\d+)?$/.test(text)) return true;
  if (/^\d+$/.test(text) && text.length > 8) return false;
  return true;
}

function findControlCode(worksheet, maxRow, maxCol) {
  const rowLimit = Math.min(maxRow, 40);
  const colLimit = Math.min(maxCol, 20);

  for (let rowIndex = 0; rowIndex <= rowLimit; rowIndex++) {
    for (let colIndex = 0; colIndex <= colLimit; colIndex++) {
      const value = readCellAt(worksheet, colIndex, rowIndex);
      if (!isA3ControlCode(value)) continue;
      return {
        controlCode: value.trim(),
        col: colIndex,
        row: rowIndex,
      };
    }
  }

  return null;
}

function scoreHeaderRow(worksheet, rowIndex, maxCol) {
  let headers = 0;
  let score = 0;

  for (let colIndex = 0; colIndex <= maxCol; colIndex++) {
    const value = readCellAt(worksheet, colIndex, rowIndex);
    if (!looksLikeColumnHeader(value)) continue;

    headers++;
    score += 1;

    const norm = normalizeAccents(value.trim());
    if (/^codigo\b/i.test(norm)) score += 5;
    if (/^nombre\b/i.test(norm)) score += 3;
    if (/^\d{2,4}$/.test(norm)) score += 2;
  }

  if (headers < MIN_HEADER_COLUMNS) return 0;
  return score + headers * 2;
}

function findBestHeaderRow(worksheet, afterRowIndex, maxRow, maxCol) {
  let bestRow = null;
  let bestScore = 0;
  const start = Math.max(afterRowIndex + 1, 0);
  const end = Math.min(maxRow, afterRowIndex + 45);

  for (let rowIndex = start; rowIndex <= end; rowIndex++) {
    const score = scoreHeaderRow(worksheet, rowIndex, maxCol);
    if (score > bestScore) {
      bestScore = score;
      bestRow = rowIndex;
    }
  }

  if (bestRow === null || bestScore < MIN_HEADER_COLUMNS * 2) {
    return null;
  }

  return bestRow;
}

function buildColumnIndices(worksheet, headerRowIndex, maxCol) {
  const indices = {};

  for (let colIndex = 0; colIndex <= maxCol; colIndex++) {
    const header = readCellAt(worksheet, colIndex, headerRowIndex);
    if (!looksLikeColumnHeader(header)) continue;
    const key = header.trim();
    indices[key] = colIndex;
  }

  return indices;
}

function collectPreambleCells(worksheet, headerRowIndex, control) {
  const cells = [];

  for (let rowIndex = 0; rowIndex < headerRowIndex; rowIndex++) {
    for (let colIndex = 0; colIndex <= 25; colIndex++) {
      const text = readCellAt(worksheet, colIndex, rowIndex);
      if (!text) continue;
      if (control && rowIndex === control.row && colIndex === control.col) continue;
      cells.push({ col: colIndex, row: rowIndex, text });
    }
  }

  return cells;
}

function readAdjacentValue(worksheet, colIndex, rowIndex) {
  for (let offset = 1; offset <= 4; offset++) {
    const value = readCellAt(worksheet, colIndex + offset, rowIndex);
    if (!value) continue;
    if (isMetadataLabel(value)) continue;
    return value;
  }
  return null;
}

function extractA3Metadata(worksheet, headerRowIndex) {
  const metadata = {
    companyCode: null,
    companyName: null,
    selection: null,
    exportDate: null,
    sectionTitle: null,
  };

  const scanUntil = headerRowIndex ?? 40;

  for (let rowIndex = 0; rowIndex < scanUntil; rowIndex++) {
    for (let colIndex = 0; colIndex <= 20; colIndex++) {
      const label = readCellAt(worksheet, colIndex, rowIndex);
      if (!label) continue;

      const norm = normalizeAccents(label.trim().replace(/:+\s*$/, ""));

      if (/^empresa$/i.test(norm)) {
        metadata.companyCode =
          readCellAt(worksheet, colIndex + 1, rowIndex) ??
          readAdjacentValue(worksheet, colIndex, rowIndex);
        metadata.companyName =
          readCellAt(worksheet, colIndex + 2, rowIndex) ??
          readCellAt(worksheet, colIndex + 3, rowIndex);
        continue;
      }

      if (/^seleccion$/i.test(norm)) {
        metadata.selection =
          readAdjacentValue(worksheet, colIndex, rowIndex) ??
          readCellAt(worksheet, colIndex + 2, rowIndex);
        continue;
      }

      if (/^fecha$/i.test(norm)) {
        metadata.exportDate = readAdjacentValue(worksheet, colIndex, rowIndex);
      }
    }
  }

  if (headerRowIndex != null && headerRowIndex > 0) {
    for (let colIndex = 0; colIndex <= 4; colIndex++) {
      const title = readCellAt(worksheet, colIndex, headerRowIndex - 1);
      if (!title) continue;
      if (isMetadataLabel(title) || isA3ControlCode(title)) continue;
      if (title.trim().length < 2) continue;
      metadata.sectionTitle = title.trim();
      break;
    }
  }

  return metadata;
}

function detectA3ExportLayout(worksheet) {
  const ref = worksheet["!ref"];
  const maxCol = ref ? parseMaxCol(ref) : SCAN_MAX_COL;
  const maxRow = ref ? parseMaxRow(ref) : SCAN_MAX_ROW;

  const control = findControlCode(worksheet, maxRow, maxCol);
  if (!control) return null;

  const headerRowIndex = findBestHeaderRow(
    worksheet,
    control.row,
    maxRow,
    maxCol
  );
  if (headerRowIndex === null) return null;

  const columnIndices = buildColumnIndices(worksheet, headerRowIndex, maxCol);
  const columns = Object.keys(columnIndices);
  if (columns.length < MIN_HEADER_COLUMNS) return null;

  const preambleCells = collectPreambleCells(worksheet, headerRowIndex, control);

  return {
    kind: "a3-export",
    controlCode: control.controlCode,
    controlCodeColIndex: control.col,
    controlCodeRowIndex: control.row,
    headerRow1Based: headerRowIndex + 1,
    dataStartRow1Based: headerRowIndex + 2,
    columnIndices,
    preambleCells,
  };
}

function normalizeCellValue(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "boolean" || typeof value === "number") return value;
  return String(value);
}

function parseA3ExportWorksheet(worksheet, layout, sheetName, maxRows) {
  const columns = Object.keys(layout.columnIndices);
  const metadata = extractA3Metadata(worksheet, layout.headerRow1Based - 1);
  const dataStartIndex = layout.dataStartRow1Based - 1;
  const maxRow = dataStartIndex + maxRows;
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

self.A3LayoutDiscovery = {
  detectA3ExportLayout,
  parseA3ExportWorksheet,
  extractA3Metadata,
  isA3ControlCode,
  readCellDisplay,
  readCellAt,
};
