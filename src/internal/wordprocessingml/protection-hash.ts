import { sha512 } from "../crypto/sha512.js";

/**
 * The password hash of `<w:documentProtection>` / `<w:writeProtection>`
 * (ECMA-376 Part 1 §17.15.1.29, attributes `w:cryptProviderType` …
 * `w:hash` / `w:salt`), as Word 2013 and later write it.
 *
 * The password first goes through Word's legacy 32-bit key derivation (the
 * algorithm the standard gives in its description of `w:hash`); the key's four
 * bytes, low byte first, as an upper-case hex string in UTF-16LE, are then
 * hashed with the salt and iterated: H0 = H(salt ‖ key), Hn = H(Hn-1 ‖ n as
 * a 32-bit little-endian integer).
 */

/** `w:cryptAlgorithmSid` 14 = SHA-512 (ST_AlgType / the CryptoAPI ALG_ID table). */
export const SHA512_ALGORITHM_SID = 14;
/** The iteration count Word uses. */
export const DEFAULT_SPIN_COUNT = 100_000;

/** Word truncates passwords to 15 characters before deriving the legacy key. */
const MAX_PASSWORD_LENGTH = 15;

// prettier-ignore
const INITIAL_CODE = [
  0xe1f0, 0x1d0f, 0xcc9c, 0x84c0, 0x110c, 0x0e10, 0xf1ce, 0x313e,
  0x1872, 0xe139, 0xd40f, 0x84f9, 0x280c, 0xa96a, 0x4ec3,
];

// prettier-ignore
const ENCRYPTION_MATRIX = [
  [0xaefc, 0x4dd9, 0x9bb2, 0x2745, 0x4e8a, 0x9d14, 0x2a09],
  [0x7b61, 0xf6c2, 0xfda5, 0xeb6b, 0xc6f7, 0x9dcf, 0x2bbf],
  [0x4563, 0x8ac6, 0x05ad, 0x0b5a, 0x16b4, 0x2d68, 0x5ad0],
  [0x0375, 0x06ea, 0x0dd4, 0x1ba8, 0x3750, 0x6ea0, 0xdd40],
  [0xd849, 0xa0b3, 0x5147, 0xa28e, 0x553d, 0xaa7a, 0x44d5],
  [0x6f45, 0xde8a, 0xad35, 0x4a4b, 0x9496, 0x390d, 0x721a],
  [0xeb23, 0xc667, 0x9cef, 0x29ff, 0x53fe, 0xa7fc, 0x5fd9],
  [0x47d3, 0x8fa6, 0x0f6d, 0x1eda, 0x3db4, 0x7b68, 0xf6d0],
  [0xb861, 0x60e3, 0xc1c6, 0x93ad, 0x377b, 0x6ef6, 0xddec],
  [0x45a0, 0x8b40, 0x06a1, 0x0d42, 0x1a84, 0x3508, 0x6a10],
  [0xaa51, 0x4483, 0x8906, 0x022d, 0x045a, 0x08b4, 0x1168],
  [0x76b4, 0xed68, 0xcaf1, 0x85c3, 0x1ba7, 0x374e, 0x6e9c],
  [0x3730, 0x6e60, 0xdcc0, 0xa9a1, 0x4363, 0x86c6, 0x1dad],
  [0x3331, 0x6662, 0xccc4, 0x89a9, 0x0373, 0x06e6, 0x0dcc],
  [0x1021, 0x2042, 0x4084, 0x0108, 0x0210, 0x0420, 0x0840],
];

const BITS_PER_PASSWORD_BYTE = 7;
const LOW_WORD_XOR = 0xce4b;

/** Rotate a 15-bit value left by one. */
function rotateLeft15(value: number): number {
  return ((value << 1) & 0x7fff) | ((value >> 14) & 1);
}

/**
 * The legacy low-order verifier word on its own — also the sheet-protection
 * hash of SpreadsheetML, which has published test values.
 */
export function legacyVerifier(bytes: readonly number[]): number {
  let verifier = 0;
  for (const byte of bytes.toReversed()) verifier = rotateLeft15(verifier) ^ byte;
  return rotateLeft15(verifier) ^ bytes.length ^ LOW_WORD_XOR;
}

/** One byte per character: the low byte, or the high byte when the low one is zero. */
export function legacyPasswordBytes(password: string): number[] {
  // UTF-16 code units, as Word processes them.
  const bytes: number[] = [];
  for (let i = 0; i < Math.min(password.length, MAX_PASSWORD_LENGTH); i++) {
    const code = password.charCodeAt(i);
    const low = code & 0xff;
    bytes.push(low === 0 ? (code >> 8) & 0xff : low);
  }
  return bytes;
}

/** Word's legacy 32-bit password key (high word from the matrix, low word the verifier). */
export function legacyPasswordKey(password: string): number {
  const bytes = legacyPasswordBytes(password);
  if (bytes.length === 0) return 0;
  let high = INITIAL_CODE[bytes.length - 1] ?? 0;
  for (const [i, byte] of bytes.entries()) {
    const row = ENCRYPTION_MATRIX[MAX_PASSWORD_LENGTH - bytes.length + i] ?? [];
    for (let bit = 0; bit < BITS_PER_PASSWORD_BYTE; bit++) {
      if (byte & (1 << bit)) high ^= row[bit] ?? 0;
    }
  }
  return ((high << 16) | legacyVerifier(bytes)) >>> 0;
}

function utf16le(text: string): Uint8Array {
  const out = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    out[i * 2] = code & 0xff;
    out[i * 2 + 1] = code >> 8;
  }
  return out;
}

function concat(a: Uint8Array, b: Uint8Array): Uint8Array {
  const out = new Uint8Array(a.length + b.length);
  out.set(a);
  out.set(b, a.length);
  return out;
}

/** The iterated SHA-512 password hash for `w:hash`. */
export function protectionHash(password: string, salt: Uint8Array, spinCount: number): Uint8Array {
  const key = legacyPasswordKey(password);
  // The key's bytes low byte first, each as two upper-case hex digits.
  const keyHex = [0, 8, 16, 24]
    .map((shift) => ((key >>> shift) & 0xff).toString(16).toUpperCase().padStart(2, "0"))
    .join("");
  let hash = sha512(concat(salt, utf16le(keyHex)));
  const iterator = new Uint8Array(4);
  const iteratorView = new DataView(iterator.buffer);
  for (let i = 0; i < spinCount; i++) {
    iteratorView.setUint32(0, i, true);
    hash = sha512(concat(hash, iterator));
  }
  return hash;
}

export function toBase64(bytes: Uint8Array): string {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

export function fromBase64(text: string): Uint8Array {
  const binary = atob(text);
  return Uint8Array.from(binary, (ch) => ch.charCodeAt(0));
}
