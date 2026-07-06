/**
 * Descifrado XLS BIFF97 con XOR (exports a3ERP). Portado de officecrypto-tool
 * para empaquetarse en el bundle serverless de Next.js/Vercel.
 */
import * as CFB from "cfb";
import type { CFB$Container } from "cfb";

const BIFF8 = 1536;

const RECORD = {
  FilePass: 47,
  WriteProtect: 134,
  BOF: 2057,
  UsrExcl: 404,
  FileLock: 405,
  InterfaceHdr: 225,
  RRDInfo: 406,
  RRDHead: 312,
  BoundSheet8: 133,
} as const;

const PLAIN_RECORDS = new Set<number>([
  RECORD.BOF,
  RECORD.FilePass,
  RECORD.UsrExcl,
  RECORD.FileLock,
  RECORD.InterfaceHdr,
  RECORD.RRDInfo,
  RECORD.RRDHead,
]);

const padArray = [
  0xbb, 0xff, 0xff, 0xba, 0xff, 0xff, 0xb9, 0x80, 0x00, 0xbe, 0x0f, 0x00, 0xbf,
  0x0f, 0x00,
];

const initialCode = [
  0xe1f0, 0x1d0f, 0xcc9c, 0x84c0, 0x110c, 0x0e10, 0xf1ce, 0x313e, 0x1872, 0xe139,
  0xd40f, 0x84f9, 0x280c, 0xa96a, 0x4ec3,
];

const xorMatrix = [
  0xaefc, 0x4dd9, 0x9bb2, 0x2745, 0x4e8a, 0x9d14, 0x2a09, 0x7b61, 0xf6c2, 0xfda5,
  0xeb6b, 0xc6f7, 0x9dcf, 0x2bbf, 0x4563, 0x8ac6, 0x05ad, 0x0b5a, 0x16b4, 0x2d68,
  0x5ad0, 0x0375, 0x06ea, 0x0dd4, 0x1ba8, 0x3750, 0x6ea0, 0xdd40, 0xd849, 0xa0b3,
  0x5147, 0xa28e, 0x553d, 0xaa7a, 0x44d5, 0x6f45, 0xde8a, 0xad35, 0x4a4b, 0x9496,
  0x390d, 0x721a, 0xeb23, 0xc667, 0x9cef, 0x29ff, 0x53fe, 0xa7fc, 0x5fd9, 0x47d3,
  0x8fa6, 0x0f6d, 0x1eda, 0x3db4, 0x7b68, 0xf6d0, 0xb861, 0x60e3, 0xc1c6, 0x93ad,
  0x377b, 0x6ef6, 0xddec, 0x45a0, 0x8b40, 0x06a1, 0x0d42, 0x1a84, 0x3508, 0x6a10,
  0xaa51, 0x4483, 0x8906, 0x022d, 0x045a, 0x08b4, 0x1168, 0x76b4, 0xed68, 0xcaf1,
  0x85c3, 0x1ba7, 0x374e, 0x6e9c, 0x3730, 0x6e60, 0xdcc0, 0xa9a1, 0x4363, 0x86c6,
  0x1dad, 0x3331, 0x6662, 0xccc4, 0x89a9, 0x0373, 0x06e6, 0x0dcc, 0x1021, 0x2042,
  0x4084, 0x8108, 0x1231, 0x2462, 0x48c4,
];

type CfbBlob = Buffer & { l: number; read_shift: (size: number) => number };

function prepBlob(buffer: Buffer): CfbBlob {
  const blob = Buffer.from(buffer) as CfbBlob;
  CFB.utils.prep_blob(blob, 0);
  return blob;
}

function verifyXorPassword(password: string, verificationBytes: number): boolean {
  let verifier = 0x0000;
  const passwordArray: number[] = [password.length];
  for (const ch of password) {
    passwordArray.push(ch.charCodeAt(0));
  }
  passwordArray.reverse();

  for (const passwordByte of passwordArray) {
    const intermediate1 = (verifier & 0x4000) === 0 ? 0 : 1;
    const intermediate2 = (verifier * 2) & 0x7fff;
    verifier = intermediate1 ^ intermediate2 ^ passwordByte;
  }

  return (verifier ^ 0xce4b) === verificationBytes;
}

function createXorKeyMethod1(password: string): number {
  const codeIndex = Math.max(0, password.length - 1);
  let xorKey = initialCode[codeIndex] ?? initialCode[0];
  let currentElement = 0x68;

  for (let i = password.length - 1; i >= 0; i--) {
    let char = password.charCodeAt(i);
    for (let j = 0; j < 7; j++) {
      if (char & 0x40) xorKey ^= xorMatrix[currentElement];
      char *= 2;
      currentElement--;
    }
  }

  return xorKey;
}

function ror(byte: number): number {
  return ((byte / 2) | (byte * 128)) & 0xff;
}

function xorRor(byte1: number, byte2: number): number {
  return ror(byte1 ^ byte2);
}

function createXorArrayMethod1(password: string): number[] {
  const xorKey = createXorKeyMethod1(password);
  let index = password.length;
  const obfuscationArray = new Array<number>(16).fill(0);

  if (index % 2 === 1) {
    let temp = xorKey >> 8;
    obfuscationArray[index] = xorRor(padArray[0], temp);
    index -= 1;
    temp = xorKey & 0x00ff;
    const passwordLastChar = password.charCodeAt(password.length - 1);
    obfuscationArray[index] = xorRor(passwordLastChar, temp);
  }

  while (index > 0) {
    index -= 1;
    let temp = xorKey >> 8;
    obfuscationArray[index] = xorRor(password.charCodeAt(index), temp);
    index -= 1;
    temp = xorKey & 0x00ff;
    obfuscationArray[index] = xorRor(password.charCodeAt(index), temp);
  }

  index = 15;
  let padIndex = 15 - password.length;
  while (padIndex > 0) {
    let temp = xorKey >> 8;
    obfuscationArray[index] = xorRor(padArray[padIndex], temp);
    index -= 1;
    padIndex -= 1;
    temp = xorKey & 0x00ff;
    obfuscationArray[index] = xorRor(padArray[padIndex], temp);
    index -= 1;
    padIndex -= 1;
  }

  return obfuscationArray;
}

function decryptXorData(
  password: string,
  input: Buffer,
  plaintext: number[]
): Buffer {
  const xorArray = createXorArrayMethod1(password);
  const chunks: Buffer[] = [];
  let encrypted = input;
  let dataIndex = 0;

  while (dataIndex < plaintext.length) {
    let count = 1;

    if (plaintext[dataIndex] === -1 || plaintext[dataIndex] === -2) {
      for (let j = dataIndex + 1; j < plaintext.length; j++) {
        if (plaintext[j] >= 0) break;
        count += 1;
      }

      let xorArrayIndex =
        plaintext[dataIndex] === -2
          ? (dataIndex + count + 4) % 16
          : (dataIndex + count) % 16;

      for (let item = 0; item < count; item++) {
        let tempRes = encrypted[0] ^ xorArray[xorArrayIndex];
        tempRes = ((tempRes >> 5) | (tempRes << 3)) & 0xff;
        chunks.push(Buffer.from([tempRes]));
        encrypted = encrypted.subarray(1);
        xorArrayIndex = (xorArrayIndex + 1) % 16;
      }
    } else {
      chunks.push(encrypted.subarray(0, 1));
      encrypted = encrypted.subarray(1);
    }

    dataIndex += count;
  }

  return Buffer.concat(chunks);
}

type WorkbookRecord = {
  header: Buffer;
  num: number;
  size: number;
  record: Buffer;
};

function iterRecord(blob: CfbBlob): WorkbookRecord[] {
  const dataList: WorkbookRecord[] = [];

  while (blob.l < blob.length) {
    const start = blob.l;
    const h = blob.read_shift(4);
    if (!h) break;

    blob.l = start;
    const header = blob.slice(blob.l, blob.l + 4);
    const num = blob.read_shift(2);
    const size = blob.read_shift(2);
    const record = blob.slice(blob.l, blob.l + size);
    dataList.push({ header, num, size, record });
    blob.l += size;
  }

  return dataList;
}

function rebuildWorkbookCfb(
  currCfb: CFB$Container,
  blob: CfbBlob,
  password: string
): Buffer {
  blob.l = 0;
  const dataList = iterRecord(blob);
  const plainBuf: number[] = [];
  const encryptedParts: Buffer[] = [];

  for (const { header, num, size, record } of dataList) {
    if (num === RECORD.FilePass) {
      plainBuf.push(0, 0, header[2], header[3], ...Array(size).fill(0));
      encryptedParts.push(Buffer.alloc(4 + size));
    } else if (PLAIN_RECORDS.has(num)) {
      plainBuf.push(...header, ...record);
      encryptedParts.push(Buffer.alloc(4 + size));
    } else if (num === RECORD.BoundSheet8) {
      const lbPlyPos = record.subarray(0, 4);
      const restSize = size - 4;
      plainBuf.push(...header, ...lbPlyPos, ...Array(restSize).fill(-2));
      encryptedParts.push(
        Buffer.concat([Buffer.alloc(4), Buffer.alloc(4), record.subarray(4)])
      );
    } else {
      plainBuf.push(...header, ...Array(size).fill(-1));
      encryptedParts.push(Buffer.concat([Buffer.alloc(4), record]));
    }
  }

  const encryptedBuf = Buffer.concat(encryptedParts);
  const dec = decryptXorData(password, encryptedBuf, plainBuf);

  for (let i = 0; i < plainBuf.length; i++) {
    const marker = plainBuf[i];
    if (marker !== -1 && marker !== -2) {
      dec.writeUInt8(marker, i);
    }
  }

  let output = CFB.utils.cfb_new();
  CFB.utils.cfb_add(output, "Workbook", dec);

  const optionalEntries = [
    "ETExtData",
    "\u0001CompObj",
    "\u0005SummaryInformation",
    "\u0005DocumentSummaryInformation",
  ] as const;

  for (const name of optionalEntries) {
    const entry = CFB.find(currCfb, name);
    if (entry) {
      CFB.utils.cfb_add(output, name, entry.content);
    }
  }

  CFB.utils.cfb_del(output, "\u0001Sh33tJ5");
  const written = CFB.write(output);
  return Buffer.isBuffer(written) ? written : Buffer.from(written);
}

function decryptWorkbookStream(
  currCfb: CFB$Container,
  workbookBlob: Buffer,
  password: string,
  input: Buffer
): Buffer | null {
  const blob = prepBlob(workbookBlob);
  blob.read_shift(2);
  const bofSize = blob.read_shift(2);
  const vers = blob.read_shift(2);
  blob.l -= 2;
  blob.l += bofSize;

  let record = blob.read_shift(2);
  let filePass = record;
  if (record === RECORD.WriteProtect) {
    blob.read_shift(2);
    filePass = blob.read_shift(2);
  }

  if (filePass !== RECORD.FilePass) {
    return input;
  }

  blob.read_shift(2);
  const wEncryptionType = vers === BIFF8 ? blob.read_shift(2) : 0;

  if (wEncryptionType !== 0x0000) {
    return null;
  }

  blob.read_shift(2);
  const verificationBytes = blob.read_shift(2);
  if (!verifyXorPassword(password, verificationBytes)) {
    return null;
  }

  return rebuildWorkbookCfb(currCfb, blob, password);
}

export function decryptXls97Buffer(input: Buffer, password: string): Buffer | null {
  const cfb = CFB.read(input, { type: "buffer" });
  const workbookEntry = CFB.find(cfb, "Workbook") ?? CFB.find(cfb, "Book");
  if (!workbookEntry) return null;

  let workbookContent = workbookEntry.content;
  const workbookBlob = Buffer.isBuffer(workbookContent)
    ? workbookContent
    : Buffer.from(workbookContent);

  try {
    return decryptWorkbookStream(cfb, workbookBlob, password, input);
  } catch {
    return null;
  }
}

export const xls97DecryptAvailable = true;
