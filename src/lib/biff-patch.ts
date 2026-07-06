/** Parcheo mínimo de celdas numéricas en streams BIFF8 (XLS). */

const RECORD_BOF = 0x0209;
const RECORD_NUMBER = 0x0203;
const RECORD_RK = 0x027e;
const SHEET_BOF_TYPE = 0x0010;

export interface BiffCellPatch {
  row: number;
  col: number;
  value: number;
}

function encodeNumberRecord(row: number, col: number, value: number): Buffer {
  const body = Buffer.alloc(14);
  body.writeUInt16LE(row, 0);
  body.writeUInt16LE(col, 2);
  body.writeDoubleLE(value, 4);
  const header = Buffer.alloc(4);
  header.writeUInt16LE(RECORD_NUMBER, 0);
  header.writeUInt16LE(10, 2);
  return Buffer.concat([header, body]);
}

function encodeRkRecord(row: number, col: number, value: number): Buffer | null {
  const rk = numberToRk(value);
  if (rk === null) return null;

  const body = Buffer.alloc(10);
  body.writeUInt16LE(row, 0);
  body.writeUInt16LE(col, 2);
  body.writeInt32LE(rk, 4);
  const header = Buffer.alloc(4);
  header.writeUInt16LE(RECORD_RK, 0);
  header.writeUInt16LE(6, 2);
  return Buffer.concat([header, body]);
}

function numberToRk(value: number): number | null {
  if (!Number.isFinite(value)) return null;

  const rounded = Math.round(value);
  if (Math.abs(value - rounded) > 1e-9) return null;
  if (rounded < -536870912 || rounded > 536870911) return null;

  return rounded << 2;
}

function decodeRk(rk: number): number {
  const isInt = (rk & 2) === 0;
  const isX100 = (rk & 1) !== 0;
  if (isInt) {
    let num = rk >> 2;
    if (isX100) num /= 100;
    return num;
  }

  const buf = Buffer.alloc(8);
  buf.writeInt32LE(rk & 0xfffffffc, 0);
  let num = buf.readDoubleLE(0);
  if (isX100) num /= 100;
  return num;
}

function patchRecordBody(
  sid: number,
  body: Buffer,
  patch: BiffCellPatch
): Buffer | null {
  if (body.length < 10) return null;

  const row = body.readUInt16LE(0);
  const col = body.readUInt16LE(2);
  if (row !== patch.row || col !== patch.col) return null;

  if (sid === RECORD_NUMBER && body.length >= 10) {
    const next = Buffer.from(body);
    next.writeDoubleLE(patch.value, 4);
    return next;
  }

  if (sid === RECORD_RK && body.length >= 6) {
    const rk = numberToRk(patch.value);
    if (rk === null) return null;
    const next = Buffer.from(body);
    next.writeInt32LE(rk, 4);
    return next;
  }

  return null;
}

function findPatch(
  patches: BiffCellPatch[],
  row: number,
  col: number
): BiffCellPatch | undefined {
  return patches.find((patch) => patch.row === row && patch.col === col);
}

export function patchBiffWorkbookStream(
  stream: Buffer,
  patches: BiffCellPatch[]
): Buffer {
  if (patches.length === 0) return stream;

  const parts: Buffer[] = [];
  let offset = 0;
  let inFirstSheet = false;
  let sheetIndex = 0;

  while (offset + 4 <= stream.length) {
    const sid = stream.readUInt16LE(offset);
    const len = stream.readUInt16LE(offset + 2);
    const total = 4 + len;
    if (total <= 4 || offset + total > stream.length) break;

    let record = stream.subarray(offset, offset + total);

    if (sid === RECORD_BOF && len >= 2) {
      const bofType = stream.readUInt16LE(offset + 4);
      if (bofType === SHEET_BOF_TYPE) {
        sheetIndex += 1;
        inFirstSheet = sheetIndex === 1;
      }
    }

    if (inFirstSheet && (sid === RECORD_NUMBER || sid === RECORD_RK)) {
      const row = stream.readUInt16LE(offset + 4);
      const col = stream.readUInt16LE(offset + 6);
      const patch = findPatch(patches, row, col);

      if (patch) {
        const body = stream.subarray(offset + 4, offset + total);
        const patchedBody = patchRecordBody(sid, body, patch);

        if (patchedBody) {
          const header = Buffer.alloc(4);
          header.writeUInt16LE(sid, 0);
          header.writeUInt16LE(patchedBody.length, 2);
          record = Buffer.concat([header, patchedBody]);
        } else if (sid === RECORD_RK) {
          const replacement = encodeNumberRecord(row, col, patch.value);
          if (replacement) record = replacement;
        }
      }
    }

    parts.push(record);
    offset += total;
  }

  if (offset < stream.length) {
    parts.push(stream.subarray(offset));
  }

  return Buffer.concat(parts);
}

export function collectChangedNumericPatches(
  layout: {
    dataStartRow1Based: number;
    columnIndices: Record<string, number>;
  },
  originalRows: Array<Record<string, string | number | boolean | null>>,
  nextRows: Array<Record<string, string | number | boolean | null>>
): BiffCellPatch[] {
  const patches: BiffCellPatch[] = [];

  nextRows.forEach((row, rowIndex) => {
    const previous = originalRows[rowIndex] ?? {};
    const excelRow = layout.dataStartRow1Based - 1 + rowIndex;

    for (const [column, colIndex] of Object.entries(layout.columnIndices)) {
      const nextValue = normalizeNumeric(row[column]);
      if (nextValue === null) continue;

      const previousValue = normalizeNumeric(previous[column]);
      if (previousValue !== null && previousValue === nextValue) continue;

      patches.push({ row: excelRow, col: colIndex, value: nextValue });
    }
  });

  return patches;
}

export function collectNumericPatches(
  layout: {
    dataStartRow1Based: number;
    columnIndices: Record<string, number>;
  },
  rows: Array<Record<string, string | number | boolean | null>>
): BiffCellPatch[] {
  const patches: BiffCellPatch[] = [];

  rows.forEach((row, rowIndex) => {
    const excelRow = layout.dataStartRow1Based - 1 + rowIndex;

    for (const [column, colIndex] of Object.entries(layout.columnIndices)) {
      const value = normalizeNumeric(row[column]);
      if (value === null) continue;
      patches.push({ row: excelRow, col: colIndex, value });
    }
  });

  return patches;
}

function normalizeNumeric(value: string | number | boolean | null): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string" && value.trim() !== "") {
    const parsed = Number(value.trim().replace(",", "."));
    if (Number.isFinite(parsed)) return parsed;
  }
  return null;
}

export function decodeRkValue(rk: number): number {
  return decodeRk(rk);
}
