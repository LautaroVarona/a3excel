/**
 * Descifrado XLS BIFF97 (XOR + RC4) para exports a3ERP.
 */
import * as CFB from "cfb";
import type { CFB$Container } from "cfb";

import {
  decryptRc4Buffer,
  decryptRc4CryptoApiBuffer,
  verifyRc4Password,
} from "./rc4-decrypt";

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

type EncryptionData = {
  type?: "rc4" | "rc4_crypto_api";
  salt?: Buffer;
  keySize?: number;
};

export type XlsEncryptionInfo = {
  encrypted: boolean;
  encryptionType: "none" | "xor" | "rc4" | "rc4_crypto_api" | "unsupported";
};

type WorkbookRecord = {
  header: Buffer;
  num: number;
  size: number;
  record: Buffer;
};

function prepBlob(buffer: Buffer): CfbBlob {
  const blob = Buffer.from(buffer) as CfbBlob;
  CFB.utils.prep_blob(blob, 0);
  return blob;
}

function createXorKeyMethod1(password: string): number {
  // Igual que officecrypto-tool: con contraseña vacía usa initialCode[-1] → undefined.
  let xorKey = initialCode[password.length - 1] as number;
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
    obfuscationArray[index] = xorRor(
      password.charCodeAt(password.length - 1),
      temp
    );
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

function parseHeaderRC4(blob: CfbBlob) {
  const salt = blob.slice(blob.l, blob.l + 16);
  blob.l += 16;
  const encryptedVerifier = blob.slice(blob.l, blob.l + 16);
  blob.l += 16;
  const encryptedVerifierHash = blob.slice(blob.l, blob.l + 16);
  blob.l += 16;
  return { salt, encryptedVerifier, encryptedVerifierHash };
}

function parseHeaderRC4CryptoAPI(blob: CfbBlob, headerSize: number) {
  const length = blob.l + headerSize;
  blob.read_shift(4);
  blob.read_shift(4);
  blob.read_shift(4);
  blob.read_shift(4);
  const keySize = blob.read_shift(4);
  blob.l = length;
  return { keySize };
}

function parseRc4CryptoApiEncryptionVerifier(blob: CfbBlob) {
  blob.read_shift(4);
  const salt = blob.slice(blob.l, blob.l + 16);
  blob.l += 16;
  const encryptedVerifier = blob.slice(blob.l, blob.l + 16);
  blob.l += 16;
  const verifierHashSize = blob.read_shift(4);
  const encryptedVerifierHash = blob.slice(blob.l, blob.l + verifierHashSize);
  blob.l += verifierHashSize;
  return { salt, encryptedVerifier, encryptedVerifierHash };
}

function rebuildWorkbookCfb(
  currCfb: CFB$Container,
  blob: CfbBlob,
  password: string,
  encryption: EncryptionData
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
      plainBuf.push(
        ...header,
        ...record.subarray(0, 4),
        ...Array(size - 4).fill(-2)
      );
      encryptedParts.push(
        Buffer.concat([Buffer.alloc(4), Buffer.alloc(4), record.subarray(4)])
      );
    } else {
      plainBuf.push(...header, ...Array(size).fill(-1));
      encryptedParts.push(Buffer.concat([Buffer.alloc(4), record]));
    }
  }

  const encryptedBuf = Buffer.concat(encryptedParts);
  let dec: Buffer;

  if (encryption.type === "rc4" && encryption.salt) {
    dec = decryptRc4Buffer(password, encryption.salt, encryptedBuf, 1024);
  } else if (
    encryption.type === "rc4_crypto_api" &&
    encryption.salt &&
    encryption.keySize
  ) {
    dec = decryptRc4CryptoApiBuffer(
      password,
      encryption.salt,
      encryption.keySize,
      encryptedBuf,
      1024
    );
  } else {
    dec = decryptXorData(password, encryptedBuf, plainBuf);
  }

  for (let i = 0; i < plainBuf.length; i++) {
    const marker = plainBuf[i];
    if (marker !== -1 && marker !== -2) {
      dec.writeUInt8(marker, i);
    }
  }

  const output = CFB.utils.cfb_new();
  CFB.utils.cfb_add(output, "Workbook", dec);

  for (const name of [
    "ETExtData",
    "\u0001CompObj",
    "\u0005SummaryInformation",
    "\u0005DocumentSummaryInformation",
  ]) {
    const entry = CFB.find(currCfb, name);
    if (entry) CFB.utils.cfb_add(output, name, entry.content);
  }

  CFB.utils.cfb_del(output, "\u0001Sh33tJ5");
  const written = CFB.write(output);
  return Buffer.isBuffer(written) ? written : Buffer.from(written);
}

function parseEncryptionHeader(
  blob: CfbBlob,
  vers: number
): { info: XlsEncryptionInfo; encryption: EncryptionData } | null {
  const record = blob.read_shift(2);
  let filePass = record;
  if (record === RECORD.WriteProtect) {
    blob.read_shift(2);
    filePass = blob.read_shift(2);
  }

  if (filePass !== RECORD.FilePass) {
    return {
      info: { encrypted: false, encryptionType: "none" },
      encryption: {},
    };
  }

  blob.read_shift(2);
  const wEncryptionType = vers === BIFF8 ? blob.read_shift(2) : 0;

  if (wEncryptionType === 0x0000) {
    blob.read_shift(2);
    const verificationBytes = blob.read_shift(2);
    return {
      info: { encrypted: true, encryptionType: "xor" },
      encryption: { verificationBytes } as EncryptionData & {
        verificationBytes: number;
      },
    };
  }

  if (wEncryptionType !== 0x0001) {
    return {
      info: { encrypted: true, encryptionType: "unsupported" },
      encryption: {},
    };
  }

  const vMajor = blob.read_shift(2);
  const vMinor = blob.read_shift(2);

  if (vMajor === 0x0001 && vMinor === 0x0001) {
    const { salt, encryptedVerifier, encryptedVerifierHash } = parseHeaderRC4(blob);
    return {
      info: { encrypted: true, encryptionType: "rc4" },
      encryption: {
        type: "rc4",
        salt,
        encryptedVerifier,
        encryptedVerifierHash,
      } as EncryptionData & {
        encryptedVerifier: Buffer;
        encryptedVerifierHash: Buffer;
      },
    };
  }

  if ([0x0002, 0x0003, 0x0004].includes(vMajor) && vMinor === 0x0002) {
    blob.read_shift(4);
    const headerSize = blob.read_shift(4);
    const { keySize } = parseHeaderRC4CryptoAPI(blob, headerSize);
    const { salt, encryptedVerifier, encryptedVerifierHash } =
      parseRc4CryptoApiEncryptionVerifier(blob);
    return {
      info: { encrypted: true, encryptionType: "rc4_crypto_api" },
      encryption: {
        type: "rc4_crypto_api",
        keySize,
        salt,
        encryptedVerifier,
        encryptedVerifierHash,
      } as EncryptionData & {
        encryptedVerifier: Buffer;
        encryptedVerifierHash: Buffer;
      },
    };
  }

  return {
    info: { encrypted: true, encryptionType: "unsupported" },
    encryption: {},
  };
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

  const parsed = parseEncryptionHeader(blob, vers);
  if (!parsed) return null;

  const { info, encryption } = parsed;
  if (!info.encrypted) return input;
  if (info.encryptionType === "unsupported") return null;

  try {
    return rebuildWorkbookCfb(currCfb, blob, password, encryption);
  } catch {
    return null;
  }
}

type Rc4EncryptionParams = {
  salt: Buffer;
  encryptedVerifier: Buffer;
  encryptedVerifierHash: Buffer;
};

function getRc4EncryptionParams(input: Buffer): Rc4EncryptionParams | null {
  try {
    const cfb = CFB.read(input, { type: "buffer" });
    const workbookEntry = CFB.find(cfb, "Workbook") ?? CFB.find(cfb, "Book");
    if (!workbookEntry) return null;

    const workbookBlob = Buffer.isBuffer(workbookEntry.content)
      ? workbookEntry.content
      : Buffer.from(workbookEntry.content);
    const blob = prepBlob(workbookBlob);
    blob.read_shift(2);
    const bofSize = blob.read_shift(2);
    const vers = blob.read_shift(2);
    blob.l -= 2;
    blob.l += bofSize;

    const parsed = parseEncryptionHeader(blob, vers);
    if (!parsed || parsed.info.encryptionType !== "rc4") return null;

    const encryption = parsed.encryption as EncryptionData & {
      salt?: Buffer;
      encryptedVerifier?: Buffer;
      encryptedVerifierHash?: Buffer;
    };

    if (
      !encryption.salt ||
      !encryption.encryptedVerifier ||
      !encryption.encryptedVerifierHash
    ) {
      return null;
    }

    return {
      salt: encryption.salt,
      encryptedVerifier: encryption.encryptedVerifier,
      encryptedVerifierHash: encryption.encryptedVerifierHash,
    };
  } catch {
    return null;
  }
}

export function findVerifiedRc4Password(
  input: Buffer,
  candidates: string[]
): string | null {
  const params = getRc4EncryptionParams(input);
  if (!params) return null;

  for (const password of candidates) {
    if (
      verifyRc4Password(
        password,
        params.salt,
        params.encryptedVerifier,
        params.encryptedVerifierHash
      )
    ) {
      return password;
    }
  }

  return null;
}

export function getXlsEncryptionInfo(input: Buffer): XlsEncryptionInfo {
  try {
    const cfb = CFB.read(input, { type: "buffer" });
    const workbookEntry = CFB.find(cfb, "Workbook") ?? CFB.find(cfb, "Book");
    if (!workbookEntry) return { encrypted: false, encryptionType: "none" };

    const workbookBlob = Buffer.isBuffer(workbookEntry.content)
      ? workbookEntry.content
      : Buffer.from(workbookEntry.content);
    const blob = prepBlob(workbookBlob);
    blob.read_shift(2);
    const bofSize = blob.read_shift(2);
    const vers = blob.read_shift(2);
    blob.l -= 2;
    blob.l += bofSize;
    return parseEncryptionHeader(blob, vers)?.info ?? {
      encrypted: false,
      encryptionType: "none",
    };
  } catch {
    return { encrypted: false, encryptionType: "none" };
  }
}

export function decryptXls97Buffer(input: Buffer, password: string): Buffer | null {
  const cfb = CFB.read(input, { type: "buffer" });
  const workbookEntry = CFB.find(cfb, "Workbook") ?? CFB.find(cfb, "Book");
  if (!workbookEntry) return null;

  const workbookBlob = Buffer.isBuffer(workbookEntry.content)
    ? workbookEntry.content
    : Buffer.from(workbookEntry.content);

  try {
    return decryptWorkbookStream(cfb, workbookBlob, password, input);
  } catch {
    return null;
  }
}

function buildEncryptedBlobFromPlain(
  plainWorkbook: Buffer,
  password: string,
  encryption: EncryptionData
): Buffer {
  const plainRecords = iterRecord(prepBlob(plainWorkbook));
  const plainBuf: number[] = [];
  const encryptedParts: Buffer[] = [];

  for (const { header, num, size, record } of plainRecords) {
    if (num === RECORD.FilePass) {
      plainBuf.push(0, 0, header[2], header[3], ...Array(size).fill(0));
      encryptedParts.push(Buffer.alloc(4 + size));
    } else if (PLAIN_RECORDS.has(num)) {
      plainBuf.push(...header, ...record);
      encryptedParts.push(Buffer.alloc(4 + size));
    } else if (num === RECORD.BoundSheet8) {
      plainBuf.push(
        ...header,
        ...record.subarray(0, 4),
        ...Array(size - 4).fill(-2)
      );
      encryptedParts.push(
        Buffer.concat([Buffer.alloc(4), record.subarray(4)])
      );
    } else {
      plainBuf.push(...header, ...Array(size).fill(-1));
      encryptedParts.push(Buffer.concat([Buffer.alloc(4), record]));
    }
  }

  const encryptedPayload = Buffer.concat(encryptedParts);

  let encryptedStream: Buffer;
  if (encryption.type === "rc4" && encryption.salt) {
    encryptedStream = decryptRc4Buffer(
      password,
      encryption.salt,
      encryptedPayload,
      1024
    );
  } else if (
    encryption.type === "rc4_crypto_api" &&
    encryption.salt &&
    encryption.keySize
  ) {
    encryptedStream = decryptRc4CryptoApiBuffer(
      password,
      encryption.salt,
      encryption.keySize,
      encryptedPayload,
      1024
    );
  } else {
    encryptedStream = decryptXorData(password, encryptedPayload, plainBuf);
  }

  const output = Buffer.alloc(plainBuf.length);
  let encryptedOffset = 0;

  for (let i = 0; i < plainBuf.length; i++) {
    const marker = plainBuf[i];
    if (marker === -1 || marker === -2) {
      output[i] = encryptedStream[encryptedOffset];
      encryptedOffset += 1;
    } else {
      output[i] = marker;
    }
  }

  return output;
}

function getWorkbookBlob(input: Buffer): Buffer | null {
  const cfb = CFB.read(input, { type: "buffer" });
  const workbookEntry = CFB.find(cfb, "Workbook") ?? CFB.find(cfb, "Book");
  if (!workbookEntry) return null;

  return Buffer.isBuffer(workbookEntry.content)
    ? workbookEntry.content
    : Buffer.from(workbookEntry.content);
}

function getPlainWorkbookStream(input: Buffer): Buffer {
  const workbookBlob = getWorkbookBlob(input);
  if (!workbookBlob) {
    throw new Error("No se encontró el stream Workbook en el archivo.");
  }

  const blob = prepBlob(workbookBlob);
  blob.read_shift(2);
  const bofSize = blob.read_shift(2);
  const vers = blob.read_shift(2);
  blob.l -= 2;
  blob.l += bofSize;

  const parsed = parseEncryptionHeader(blob, vers);
  if (!parsed || !parsed.info.encrypted) {
    return workbookBlob;
  }

  throw new Error("El stream Workbook sigue cifrado.");
}

export function exportPreservingXlsBuffer(
  input: Buffer,
  plainWorkbook: Buffer,
  password: string
): Buffer {
  const workbookBlob = getWorkbookBlob(input);
  if (!workbookBlob) {
    throw new Error("No se encontró el stream Workbook en el archivo.");
  }

  const blob = prepBlob(workbookBlob);
  blob.read_shift(2);
  const bofSize = blob.read_shift(2);
  const vers = blob.read_shift(2);
  blob.l -= 2;
  blob.l += bofSize;

  const parsed = parseEncryptionHeader(blob, vers);
  const outputCfb = CFB.read(input, { type: "buffer" });

  if (!parsed || !parsed.info.encrypted) {
    CFB.utils.cfb_add(outputCfb, "Workbook", plainWorkbook);
    const written = CFB.write(outputCfb);
    return Buffer.isBuffer(written) ? written : Buffer.from(written);
  }

  const encryptedWorkbook = buildEncryptedBlobFromPlain(
    plainWorkbook,
    password,
    parsed.encryption
  );
  CFB.utils.cfb_add(outputCfb, "Workbook", encryptedWorkbook);
  const written = CFB.write(outputCfb);
  return Buffer.isBuffer(written) ? written : Buffer.from(written);
}

export function resolvePlainWorkbookForExport(
  input: Buffer,
  passwordCandidates: string[]
): { plainWorkbook: Buffer; workingCopy: Buffer; password: string | null } {
  for (const password of passwordCandidates) {
    const decrypted = decryptXls97Buffer(input, password);
    if (!decrypted) continue;

    try {
      return {
        plainWorkbook: getPlainWorkbookStream(decrypted),
        workingCopy: decrypted,
        password,
      };
    } catch {
      continue;
    }
  }

  try {
    return {
      plainWorkbook: getPlainWorkbookStream(input),
      workingCopy: input,
      password: null,
    };
  } catch {
    throw new Error(
      "No se pudo descifrar el .XLS original para exportar en formato A3."
    );
  }
}

export const xls97DecryptAvailable = true;
