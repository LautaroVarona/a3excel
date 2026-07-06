/**
 * RC4 / RC4-CryptoAPI para XLS BIFF97 (portado de officecrypto-tool).
 */
import crypto from "node:crypto";

import CryptoJS from "crypto-js";

function bufferToWordArray(buf: Buffer) {
  return CryptoJS.lib.WordArray.create(Array.from(buf));
}

function wordArrayToBuffer(wordArray: CryptoJS.lib.WordArray): Buffer {
  return Buffer.from(wordArray.toString(CryptoJS.enc.Hex), "hex");
}

function convertRc4PasswordToKey(
  password: string,
  salt: Buffer,
  block: number
): Buffer {
  const passwordBuffer = Buffer.from(password, "utf16le");
  const h0 = crypto.createHash("md5").update(passwordBuffer).digest();
  let truncatedHash = h0.subarray(0, 5);
  let intermediateBuffer = Buffer.concat([truncatedHash, salt]);
  intermediateBuffer = Buffer.concat(Array(16).fill(intermediateBuffer));
  intermediateBuffer = crypto.createHash("md5").update(intermediateBuffer).digest();

  truncatedHash = intermediateBuffer.subarray(0, 5);
  const blockBytes = Buffer.alloc(4);
  blockBytes.writeInt32LE(block, 0);
  const finalBuffer = Buffer.concat([truncatedHash, blockBytes]);
  return crypto.createHash("md5").update(finalBuffer).digest().subarray(0, 16);
}

function convertRc4CryptoApiPasswordToKey(
  password: string,
  salt: Buffer,
  keyLength: number,
  block: number
): Buffer {
  const passwordBuffer = Buffer.from(password, "utf16le");
  const h0 = crypto
    .createHash("sha1")
    .update(Buffer.concat([salt, passwordBuffer]))
    .digest();
  const blockBytes = Buffer.alloc(4);
  blockBytes.writeUInt32LE(block, 0);
  const hFinal = crypto
    .createHash("sha1")
    .update(Buffer.concat([h0, blockBytes]))
    .digest();

  if (keyLength === 40) {
    return Buffer.concat([hFinal.subarray(0, 5), Buffer.alloc(11)]);
  }
  return hFinal.subarray(0, keyLength / 8);
}

export function verifyRc4Password(
  password: string,
  salt: Buffer,
  encryptedVerifier: Buffer,
  encryptedVerifierHash: Buffer
): boolean {
  const key = convertRc4PasswordToKey(password, salt, 0);
  const cipher = CryptoJS.algo.RC4.createDecryptor(bufferToWordArray(key));
  const verifier = cipher.finalize(bufferToWordArray(encryptedVerifier));
  const hash = CryptoJS.MD5(verifier);
  const verifierHash = cipher.finalize(bufferToWordArray(encryptedVerifierHash));
  return verifierHash.toString(CryptoJS.enc.Hex) === hash.toString(CryptoJS.enc.Hex);
}

export function verifyRc4CryptoApiPassword(
  password: string,
  salt: Buffer,
  keySize: number,
  encryptedVerifier: Buffer,
  encryptedVerifierHash: Buffer
): boolean {
  const key = convertRc4CryptoApiPasswordToKey(password, salt, keySize, 0);
  const cipher = CryptoJS.algo.RC4.createDecryptor(bufferToWordArray(key));
  const verifier = cipher.finalize(bufferToWordArray(encryptedVerifier));
  const verifierHash = cipher.finalize(bufferToWordArray(encryptedVerifierHash));
  const hash = CryptoJS.SHA1(verifier);
  return verifierHash.toString(CryptoJS.enc.Hex) === hash.toString(CryptoJS.enc.Hex);
}

export function decryptRc4Buffer(
  password: string,
  salt: Buffer,
  input: Buffer,
  blocksize = 0x200
): Buffer {
  const outputChunks: Buffer[] = [];
  let block = 0;
  let start = 0;

  while (start < input.length) {
    const end = Math.min(start + blocksize, input.length);
    const inputChunk = input.subarray(start, end);
    const key = convertRc4PasswordToKey(password, salt, block);
    const cipher = CryptoJS.algo.RC4.createDecryptor(bufferToWordArray(key));
    const outputChunk = wordArrayToBuffer(
      cipher.finalize(bufferToWordArray(inputChunk))
    );
    outputChunks.push(outputChunk);
    block += 1;
    start = end;
  }

  return Buffer.concat(outputChunks);
}

export function decryptRc4CryptoApiBuffer(
  password: string,
  salt: Buffer,
  keySize: number,
  input: Buffer,
  blocksize = 0x200,
  block = 0
): Buffer {
  const outputChunks: Buffer[] = [];
  let start = 0;

  while (start < input.length) {
    const end = Math.min(start + blocksize, input.length);
    const inputChunk = input.subarray(start, end);
    const key = convertRc4CryptoApiPasswordToKey(
      password,
      salt,
      keySize,
      block
    );
    const cipher = CryptoJS.algo.RC4.createDecryptor(bufferToWordArray(key));
    const outputChunk = wordArrayToBuffer(
      cipher.finalize(bufferToWordArray(inputChunk))
    );
    outputChunks.push(outputChunk);
    block += 1;
    start = end;
  }

  return Buffer.concat(outputChunks);
}
