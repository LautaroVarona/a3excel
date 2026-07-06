/**
 * RC4 / RC4-CryptoAPI para XLS BIFF97.
 * RC4 en JS puro: OpenSSL 3 (Node 20+ en Vercel) no expone RC4 vía crypto.createDecipheriv.
 */
import crypto from "node:crypto";

/** RC4 simétrico (encrypt = decrypt). */
function rc4(key: Buffer, data: Buffer): Buffer {
  const state = new Uint8Array(256);
  for (let i = 0; i < 256; i++) state[i] = i;

  let j = 0;
  for (let i = 0; i < 256; i++) {
    j = (j + state[i] + key[i % key.length]) & 255;
    [state[i], state[j]] = [state[j], state[i]];
  }

  const output = Buffer.alloc(data.length);
  let i = 0;
  j = 0;
  for (let n = 0; n < data.length; n++) {
    i = (i + 1) & 255;
    j = (j + state[i]) & 255;
    [state[i], state[j]] = [state[j], state[i]];
    output[n] = data[n] ^ state[(state[i] + state[j]) & 255];
  }

  return output;
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
  const stream = rc4(key, Buffer.concat([encryptedVerifier, encryptedVerifierHash]));
  const verifier = stream.subarray(0, encryptedVerifier.length);
  const hash = crypto.createHash("md5").update(verifier).digest();
  const verifierHash = stream.subarray(encryptedVerifier.length);
  return verifierHash.equals(hash);
}

export function verifyRc4CryptoApiPassword(
  password: string,
  salt: Buffer,
  keySize: number,
  encryptedVerifier: Buffer,
  encryptedVerifierHash: Buffer
): boolean {
  const key = convertRc4CryptoApiPasswordToKey(password, salt, keySize, 0);
  const stream = rc4(key, Buffer.concat([encryptedVerifier, encryptedVerifierHash]));
  const verifier = stream.subarray(0, encryptedVerifier.length);
  const hash = crypto.createHash("sha1").update(verifier).digest();
  const verifierHash = stream.subarray(encryptedVerifier.length);
  return verifierHash.equals(hash);
}

export function decryptRc4Buffer(
  password: string,
  salt: Buffer,
  input: Buffer,
  blocksize = 1024
): Buffer {
  const outputChunks: Buffer[] = [];
  let block = 0;
  let start = 0;

  while (start < input.length) {
    const end = Math.min(start + blocksize, input.length);
    const inputChunk = input.subarray(start, end);
    const key = convertRc4PasswordToKey(password, salt, block);
    outputChunks.push(rc4(key, inputChunk));
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
  blocksize = 1024,
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
    outputChunks.push(rc4(key, inputChunk));
    block += 1;
    start = end;
  }

  return Buffer.concat(outputChunks);
}
