/**
 * ExcelForge — OOXML Agile Encryption Module (optional, tree-shakeable).
 *
 * Implements ECMA-376 Agile Encryption (EncryptedPackage) for password-protecting
 * .xlsx files. Uses standard Web Crypto API — works in Node.js 18+, browsers,
 * Deno, and Bun.
 *
 * Reference: [MS-OFFCRYPTO] — Office Document Cryptography Structure
 * Reference: ECMA-376 Part 2: Open Packaging Conventions
 *
 * Encryption flow:
 * 1. Generate random salt, key salt, password verifier salt
 * 2. Derive encryption key from password via iterated SHA-512
 * 3. Encrypt the XLSX package bytes with AES-256-CBC
 * 4. Generate HMAC for integrity verification
 * 5. Encrypt the key, verifier, and HMAC
 * 6. Wrap everything in a CFB container with EncryptionInfo + EncryptedPackage streams
 *
 * Decryption flow (reverse):
 * 1. Read CFB container, extract EncryptionInfo + EncryptedPackage
 * 2. Parse EncryptionInfo XML to get salts, IVs, cipher params
 * 3. Derive key from password via iterated SHA-512
 * 4. Decrypt key verifier to check password correctness
 * 5. Decrypt the package to get original XLSX bytes
 */

// ── CFB helpers (inline minimal version for encryption container) ────────────

const CFB_SIG = new Uint8Array([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
const ENDOFCHAIN = 0xFFFFFFFE;
const FREESECT   = 0xFFFFFFFF;
const FATSECT    = 0xFFFFFFFD;
const SECTOR_SZ  = 512;
const MINI_SZ    = 64;
const MINI_CUT   = 0x1000;
const DIR_SZ     = 128;

interface CfbEntry { name: string; data: Uint8Array; }

function setU16(buf: Uint8Array, off: number, v: number): void {
  buf[off] = v & 0xFF; buf[off + 1] = (v >> 8) & 0xFF;
}
function setU32(buf: Uint8Array, off: number, v: number): void {
  buf[off] = v & 0xFF; buf[off + 1] = (v >> 8) & 0xFF;
  buf[off + 2] = (v >> 16) & 0xFF; buf[off + 3] = (v >> 24) & 0xFF;
}
function u16(buf: Uint8Array, off: number): number {
  return buf[off] | (buf[off + 1] << 8);
}
function u32(buf: Uint8Array, off: number): number {
  return (buf[off] | (buf[off + 1] << 8) | (buf[off + 2] << 16) | (buf[off + 3] << 24)) >>> 0;
}

/** UTF-16LE encode name with null terminator */
function encUtf16(name: string): { bytes: Uint8Array; size: number } {
  const bytes = new Uint8Array(64);
  const len = Math.min(name.length, 31);
  for (let i = 0; i < len; i++) { setU16(bytes, i * 2, name.charCodeAt(i)); }
  setU16(bytes, len * 2, 0); // null
  return { bytes, size: (len + 1) * 2 };
}

/** Build a CFB file containing the given named streams */
function buildCfb(entries: CfbEntry[]): Uint8Array {
  // Single-level directory: Root Entry + streams
  const dirCount = 1 + entries.length;
  const allStreams = entries.map(e => e.data);

  // Use mini-stream for everything < 4096, regular for larger
  let miniData = new Uint8Array(0);
  const miniBlocks: { offset: number; size: number }[] = [];
  const regularBlocks: { data: Uint8Array; startSector: number }[] = [];

  // Lay out FAT, directory, mini-stream, and regular streams
  // For simplicity: FAT sector=0, directory sector=1, mini-stream sectors, regular sectors

  // Compute mini-stream
  let miniOffset = 0;
  for (const stream of allStreams) {
    if (stream.length < MINI_CUT) {
      const padLen = Math.ceil(stream.length / MINI_SZ) * MINI_SZ;
      const newMini = new Uint8Array(miniData.length + padLen);
      newMini.set(miniData);
      newMini.set(stream, miniData.length);
      miniBlocks.push({ offset: miniOffset, size: stream.length });
      miniOffset += padLen;
      miniData = newMini;
    } else {
      miniBlocks.push({ offset: -1, size: stream.length }); // marker for regular
    }
  }

  // How many sectors for mini-stream data?
  const miniStreamSectors = Math.ceil(miniData.length / SECTOR_SZ);
  // FAT sector: 0, Dir sectors: depends on dirCount
  const dirSectors = Math.ceil((dirCount * DIR_SZ) / SECTOR_SZ);

  // Layout: [FAT sector][dir sectors][mini-stream sectors][regular stream sectors]
  let nextSector = 1 + dirSectors + miniStreamSectors;
  const regularStartSectors: number[] = [];
  for (let i = 0; i < entries.length; i++) {
    if (miniBlocks[i].offset === -1) {
      regularStartSectors.push(nextSector);
      const sects = Math.ceil(allStreams[i].length / SECTOR_SZ);
      regularBlocks.push({ data: allStreams[i], startSector: nextSector });
      nextSector += sects;
    } else {
      regularStartSectors.push(-1); // in mini-stream
    }
  }

  const totalSectors = nextSector;
  // Build FAT
  const fatEntries = new Uint32Array(Math.max(totalSectors, 128));
  fatEntries.fill(FREESECT);
  fatEntries[0] = FATSECT; // FAT sector self-reference
  // Directory chain
  for (let i = 1; i < 1 + dirSectors; i++) {
    fatEntries[i] = (i < dirSectors) ? i + 1 : ENDOFCHAIN;
  }
  // Mini-stream chain (Root Entry data)
  for (let i = 1 + dirSectors; i < 1 + dirSectors + miniStreamSectors; i++) {
    fatEntries[i] = (i < dirSectors + miniStreamSectors) ? i + 1 : ENDOFCHAIN;
  }
  // Regular stream chains
  for (const rb of regularBlocks) {
    const sects = Math.ceil(rb.data.length / SECTOR_SZ);
    for (let j = 0; j < sects; j++) {
      fatEntries[rb.startSector + j] = j < sects - 1 ? rb.startSector + j + 1 : ENDOFCHAIN;
    }
  }

  // Mini-FAT
  const miniSectorCount = Math.ceil(miniData.length / MINI_SZ);
  const miniFat = new Uint32Array(Math.max(miniSectorCount, 128));
  miniFat.fill(FREESECT);
  let mIdx = 0;
  for (let i = 0; i < entries.length; i++) {
    if (miniBlocks[i].offset !== -1) {
      const msectors = Math.ceil(allStreams[i].length / MINI_SZ);
      const startMiniSector = miniBlocks[i].offset / MINI_SZ;
      for (let j = 0; j < msectors; j++) {
        miniFat[startMiniSector + j] = j < msectors - 1 ? startMiniSector + j + 1 : ENDOFCHAIN;
      }
    }
  }
  // Mini-FAT needs its own sector(s) — add after regular streams
  const miniFatSectorCount = Math.ceil(miniFat.length * 4 / SECTOR_SZ);
  const miniFatStartSector = nextSector;
  for (let i = 0; i < miniFatSectorCount; i++) {
    fatEntries[miniFatStartSector + i] = i < miniFatSectorCount - 1 ? miniFatStartSector + i + 1 : ENDOFCHAIN;
  }
  const actualTotalSectors = miniFatStartSector + miniFatSectorCount;

  // Build output
  const fileSize = (1 + actualTotalSectors) * SECTOR_SZ; // header + sectors
  const out = new Uint8Array(fileSize);

  // ── Header ──
  out.set(CFB_SIG, 0);
  // Minor version = 0x003E, Major = 0x0003 (v3)
  setU16(out, 0x18, 0x003E);
  setU16(out, 0x1A, 0x0003);
  setU16(out, 0x1C, 0xFFFE); // byte order (little-endian)
  setU16(out, 0x1E, 9);       // sector size power (2^9 = 512)
  setU16(out, 0x20, 6);       // mini sector size power (2^6 = 64)
  setU32(out, 0x28, dirSectors); // directory sectors (v3 = 0 usually, but we set it)
  setU32(out, 0x2C, 1);       // FAT sectors = 1
  setU32(out, 0x30, 1);       // first directory sector
  setU32(out, 0x38, MINI_CUT); // mini-stream cutoff
  setU32(out, 0x3C, miniFatStartSector); // first mini-FAT sector
  setU32(out, 0x40, miniFatSectorCount); // mini-FAT sector count
  setU32(out, 0x44, ENDOFCHAIN); // first DIFAT sector (none)
  setU32(out, 0x48, 0);       // DIFAT sectors count = 0
  // DIFAT array at 0x4C: first entry = sector 0 (FAT)
  for (let i = 0; i < 109; i++) setU32(out, 0x4C + i * 4, FREESECT);
  setU32(out, 0x4C, 0); // FAT is in sector 0

  // ── Sector 0: FAT ──
  const fatOff = SECTOR_SZ;
  for (let i = 0; i < 128; i++) { // 128 entries fill 512 bytes
    setU32(out, fatOff + i * 4, fatEntries[i]);
  }

  // ── Directory sectors ──
  const dirOff = SECTOR_SZ + SECTOR_SZ; // after header + FAT sector
  // Root Entry
  const root = encUtf16('Root Entry');
  out.set(root.bytes, dirOff);
  setU16(out, dirOff + 0x40, root.size);
  out[dirOff + 0x42] = 5; // root storage
  out[dirOff + 0x43] = 1; // red
  setU32(out, dirOff + 0x44, FREESECT); // left child
  setU32(out, dirOff + 0x48, FREESECT); // right child
  // Child ID: first stream
  setU32(out, dirOff + 0x4C, entries.length > 0 ? 1 : FREESECT);
  // Root entry start sector = mini-stream start
  setU32(out, dirOff + 0x74, miniStreamSectors > 0 ? 1 + dirSectors : ENDOFCHAIN);
  setU32(out, dirOff + 0x78, miniData.length);

  // Stream entries (using simple left-linear tree)
  for (let i = 0; i < entries.length; i++) {
    const eOff = dirOff + (i + 1) * DIR_SZ;
    const eName = encUtf16(entries[i].name);
    out.set(eName.bytes, eOff);
    setU16(out, eOff + 0x40, eName.size);
    out[eOff + 0x42] = 2; // stream
    out[eOff + 0x43] = 1; // red
    setU32(out, eOff + 0x44, FREESECT); // left sibling (none)
    setU32(out, eOff + 0x48, i + 2 < entries.length + 1 ? i + 2 : FREESECT); // right sibling
    setU32(out, eOff + 0x4C, FREESECT); // child (none for streams)
    // Start sector and size
    if (miniBlocks[i].offset !== -1) {
      // Mini-stream
      setU32(out, eOff + 0x74, miniBlocks[i].offset / MINI_SZ);
      setU32(out, eOff + 0x78, allStreams[i].length);
    } else {
      // Regular
      setU32(out, eOff + 0x74, regularStartSectors[i]);
      setU32(out, eOff + 0x78, allStreams[i].length);
    }
  }

  // ── Mini-stream data sectors ──
  const miniDataOff = SECTOR_SZ + SECTOR_SZ + dirSectors * SECTOR_SZ;
  out.set(miniData.subarray(0, Math.min(miniData.length, miniStreamSectors * SECTOR_SZ)), miniDataOff);

  // ── Regular stream data ──
  for (const rb of regularBlocks) {
    const off = SECTOR_SZ + rb.startSector * SECTOR_SZ;
    out.set(rb.data, off);
  }

  // ── Mini-FAT sector(s) ──
  const mfOff = SECTOR_SZ + miniFatStartSector * SECTOR_SZ;
  for (let i = 0; i < miniFat.length && i * 4 < miniFatSectorCount * SECTOR_SZ; i++) {
    setU32(out, mfOff + i * 4, miniFat[i]);
  }

  return out;
}

/** Read a CFB file and return named streams */
function readCfb(data: Uint8Array): CfbEntry[] {
  // Verify signature
  for (let i = 0; i < 8; i++) {
    if (data[i] !== CFB_SIG[i]) throw new Error('Not a CFB file');
  }
  const sectorPow = u16(data, 0x1E);
  const sectorSize = 1 << sectorPow;
  const miniPow = u16(data, 0x20);
  const miniSize = 1 << miniPow;
  const miniCutoff = u32(data, 0x38);
  const fatSectors = u32(data, 0x2C);
  const dirStart = u32(data, 0x30);
  const miniFatStart = u32(data, 0x3C);

  const sectorOff = (s: number) => sectorSize + s * sectorSize;

  // Read FAT
  const fatSectorList: number[] = [];
  for (let i = 0; i < 109; i++) {
    const s = u32(data, 0x4C + i * 4);
    if (s === FREESECT || s === ENDOFCHAIN) break;
    fatSectorList.push(s);
  }
  const fat: number[] = [];
  for (const s of fatSectorList) {
    const off = sectorOff(s);
    for (let i = 0; i < sectorSize / 4; i++) {
      fat.push(u32(data, off + i * 4));
    }
  }

  // Follow a chain through FAT
  const followChain = (start: number): number[] => {
    const chain: number[] = [];
    let s = start;
    while (s !== ENDOFCHAIN && s !== FREESECT && chain.length < 10000) {
      chain.push(s);
      s = fat[s] ?? ENDOFCHAIN;
    }
    return chain;
  };

  // Read stream data from chain
  const readStream = (start: number, size: number): Uint8Array => {
    const chain = followChain(start);
    const buf = new Uint8Array(size);
    let pos = 0;
    for (const s of chain) {
      const off = sectorOff(s);
      const len = Math.min(sectorSize, size - pos);
      buf.set(data.subarray(off, off + len), pos);
      pos += len;
    }
    return buf;
  };

  // Read directory
  const dirChain = followChain(dirStart);
  const dirData = new Uint8Array(dirChain.length * sectorSize);
  dirChain.forEach((s, i) => dirData.set(data.subarray(sectorOff(s), sectorOff(s) + sectorSize), i * sectorSize));
  const dirEntries: Array<{ name: string; type: number; start: number; size: number; child: number }> = [];
  for (let i = 0; i < dirData.length / DIR_SZ; i++) {
    const off = i * DIR_SZ;
    const nameLen = u16(dirData, off + 0x40);
    if (nameLen === 0) continue;
    let name = '';
    for (let j = 0; j < (nameLen - 2) / 2; j++) {
      name += String.fromCharCode(u16(dirData, off + j * 2));
    }
    dirEntries.push({
      name,
      type: dirData[off + 0x42],
      start: u32(dirData, off + 0x74),
      size: u32(dirData, off + 0x78),
      child: u32(dirData, off + 0x4C),
    });
  }

  // Read mini-stream (from Root Entry's data)
  const root = dirEntries[0];
  let miniStreamData: Uint8Array = new Uint8Array(0);
  if (root && root.start !== ENDOFCHAIN) {
    miniStreamData = new Uint8Array(readStream(root.start, root.size));
  }

  // Read mini-FAT
  const miniFat: number[] = [];
  if (miniFatStart !== ENDOFCHAIN) {
    const mfChain = followChain(miniFatStart);
    for (const s of mfChain) {
      const off = sectorOff(s);
      for (let i = 0; i < sectorSize / 4; i++) {
        miniFat.push(u32(data, off + i * 4));
      }
    }
  }

  const readMiniStream = (start: number, size: number): Uint8Array => {
    const buf = new Uint8Array(size);
    let s = start, pos = 0;
    while (s !== ENDOFCHAIN && pos < size) {
      const off = s * miniSize;
      const len = Math.min(miniSize, size - pos);
      buf.set(miniStreamData.subarray(off, off + len), pos);
      pos += len;
      s = miniFat[s] ?? ENDOFCHAIN;
    }
    return buf;
  };

  // Extract streams (skip root)
  const result: CfbEntry[] = [];
  for (let i = 1; i < dirEntries.length; i++) {
    const e = dirEntries[i];
    if (e.type !== 2) continue; // only user streams
    let streamData: Uint8Array;
    if (e.size < miniCutoff) {
      streamData = readMiniStream(e.start, e.size);
    } else {
      streamData = readStream(e.start, e.size);
    }
    result.push({ name: e.name, data: streamData });
  }
  return result;
}

// ── Crypto helpers using Web Crypto API ──────────────────────────────────────

function getCrypto(): Crypto {
  if (typeof globalThis.crypto !== 'undefined') return globalThis.crypto;
  throw new Error('Web Crypto API not available. Requires Node.js 18+ or a modern browser.');
}

function randomBytes(n: number): Uint8Array {
  return getCrypto().getRandomValues(new Uint8Array(n));
}

async function sha512(data: Uint8Array): Promise<Uint8Array> {
  const buf = await getCrypto().subtle.digest('SHA-512', data as any);
  return new Uint8Array(buf);
}

async function sha1(data: Uint8Array): Promise<Uint8Array> {
  const buf = await getCrypto().subtle.digest('SHA-1', data as any);
  return new Uint8Array(buf);
}

async function hmacSha512(key: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const cryptoKey = await getCrypto().subtle.importKey('raw', key as any, { name: 'HMAC', hash: 'SHA-512' }, false, ['sign']);
  const sig = await getCrypto().subtle.sign('HMAC', cryptoKey, data as any);
  return new Uint8Array(sig);
}

/** Derive encryption key from password using iterated SHA-512 (MS-OFFCRYPTO §2.3.6.2) */
async function deriveKey(password: string, salt: Uint8Array, spinCount: number, keyBits: number, blockKey: Uint8Array): Promise<Uint8Array> {
  // Convert password to UTF-16LE
  const pwBuf = new Uint8Array(password.length * 2);
  for (let i = 0; i < password.length; i++) {
    pwBuf[i * 2] = password.charCodeAt(i) & 0xFF;
    pwBuf[i * 2 + 1] = (password.charCodeAt(i) >> 8) & 0xFF;
  }

  // H0 = SHA-512(salt + password)
  const h0Input = new Uint8Array(salt.length + pwBuf.length);
  h0Input.set(salt);
  h0Input.set(pwBuf, salt.length);
  let hash = await sha512(h0Input);

  // Iterate: Hn = SHA-512( iterator(LE32) + H(n-1) )
  for (let i = 0; i < spinCount; i++) {
    const iterBuf = new Uint8Array(4 + hash.length);
    iterBuf[0] = i & 0xFF;
    iterBuf[1] = (i >> 8) & 0xFF;
    iterBuf[2] = (i >> 16) & 0xFF;
    iterBuf[3] = (i >> 24) & 0xFF;
    iterBuf.set(hash, 4);
    hash = await sha512(iterBuf);
  }

  // Derive final key: SHA-512(Hlast + blockKey)
  const final = new Uint8Array(hash.length + blockKey.length);
  final.set(hash);
  final.set(blockKey, hash.length);
  const derived = await sha512(final);

  // Truncate/pad to keyBits/8 bytes
  const keyLen = keyBits / 8;
  if (derived.length >= keyLen) return derived.subarray(0, keyLen);
  // Pad with 0x36
  const padded = new Uint8Array(keyLen);
  padded.set(derived);
  padded.fill(0x36, derived.length);
  return padded;
}

/** AES-CBC encrypt (no PKCS#7 — data MUST be block-aligned, per MS-OFFCRYPTO) */
async function aesCbcEncrypt(key: Uint8Array, iv: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const cryptoKey = await getCrypto().subtle.importKey('raw', key as any, { name: 'AES-CBC' }, false, ['encrypt']);
  const buf = await getCrypto().subtle.encrypt({ name: 'AES-CBC', iv: iv as any } as any, cryptoKey, data as any);
  // Web Crypto adds PKCS#7 padding (extra 16 bytes). Strip it — MS-OFFCRYPTO uses zero-padding.
  return new Uint8Array(buf).subarray(0, data.length);
}

/** AES-CBC decrypt (no PKCS#7 — data MUST be block-aligned, per MS-OFFCRYPTO) */
async function aesCbcDecrypt(key: Uint8Array, iv: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const subtle = getCrypto().subtle;
  // Web Crypto expects PKCS#7 padding. Synthesize a valid padding block so it can decrypt.
  const lastBlock = data.subarray(data.length - 16);
  const pkcs7 = new Uint8Array(16).fill(16); // 0x10 = full-block PKCS#7
  const xored = new Uint8Array(16);
  for (let i = 0; i < 16; i++) xored[i] = pkcs7[i] ^ lastBlock[i];
  // ECB-encrypt one block (CBC with zero IV) to create a ciphertext block
  // that decrypts to valid PKCS#7 padding in the CBC chain
  const encKey = await subtle.importKey('raw', key as any, { name: 'AES-CBC' }, false, ['encrypt']);
  const fakeEnc = await subtle.encrypt({ name: 'AES-CBC', iv: new Uint8Array(16) as any } as any, encKey, xored as any);
  const fakeBlock = new Uint8Array(fakeEnc).subarray(0, 16);
  const padded = concat(data, fakeBlock);
  const decKey = await subtle.importKey('raw', key as any, { name: 'AES-CBC' }, false, ['decrypt']);
  const buf = await subtle.decrypt({ name: 'AES-CBC', iv: iv as any } as any, decKey, padded as any);
  return new Uint8Array(buf);
}

/** Pad data to blockSize using PKCS7-like padding (pad with 0x00 for OOXML) */
function padToBlockSize(data: Uint8Array, blockSize: number): Uint8Array {
  const rem = data.length % blockSize;
  if (rem === 0) return data;
  const padded = new Uint8Array(data.length + (blockSize - rem));
  padded.set(data);
  return padded;
}

/** Concatenate Uint8Arrays */
function concat(...arrays: Uint8Array[]): Uint8Array {
  let len = 0;
  for (const a of arrays) len += a.length;
  const out = new Uint8Array(len);
  let pos = 0;
  for (const a of arrays) { out.set(a, pos); pos += a.length; }
  return out;
}

// ── Block key constants (MS-OFFCRYPTO §2.3.6.2) ─────────────────────────────

const BLOCK_KEY_VERIFIER_INPUT  = new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]);
const BLOCK_KEY_VERIFIER_VALUE  = new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]);
const BLOCK_KEY_ENCRYPTED_KEY   = new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]);
const BLOCK_KEY_DATA_INTEGRITY1 = new Uint8Array([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]);
const BLOCK_KEY_DATA_INTEGRITY2 = new Uint8Array([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]);

// ── Encryption XML builder ───────────────────────────────────────────────────

function buildEncryptionInfoXml(
  keySalt: Uint8Array,
  encKeyValue: Uint8Array,
  encVerifierInput: Uint8Array,
  encVerifierHash: Uint8Array,
  encKeyValueHmac: Uint8Array,
  encKeyValueHmac2: Uint8Array,
  passwordSalt: Uint8Array,
  spinCount: number,
): string {
  const b64 = (buf: Uint8Array) => {
    let s = '';
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
    for (let i = 0; i < buf.length; i += 3) {
      const b0 = buf[i], b1 = buf[i + 1] ?? 0, b2 = buf[i + 2] ?? 0;
      const n = (b0 << 16) | (b1 << 8) | b2;
      s += chars[(n >> 18) & 63] + chars[(n >> 12) & 63];
      s += i + 1 < buf.length ? chars[(n >> 6) & 63] : '=';
      s += i + 2 < buf.length ? chars[n & 63] : '=';
    }
    return s;
  };

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
  xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64"
    cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
    hashAlgorithm="SHA512"
    saltValue="${b64(keySalt)}"/>
  <dataIntegrity encryptedHmacKey="${b64(encKeyValueHmac)}"
    encryptedHmacValue="${b64(encKeyValueHmac2)}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="${spinCount}" saltSize="16" blockSize="16"
        keyBits="256" hashSize="64"
        cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
        hashAlgorithm="SHA512"
        saltValue="${b64(passwordSalt)}"
        encryptedVerifierHashInput="${b64(encVerifierInput)}"
        encryptedVerifierHashValue="${b64(encVerifierHash)}"
        encryptedKeyValue="${b64(encKeyValue)}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>`;
}

// ── Parse EncryptionInfo XML ─────────────────────────────────────────────────

function b64Decode(s: string): Uint8Array {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  const lookup = new Uint8Array(128);
  for (let i = 0; i < chars.length; i++) lookup[chars.charCodeAt(i)] = i;
  const clean = s.replace(/[^A-Za-z0-9+/]/g, '');
  const pad = s.endsWith('==') ? 2 : s.endsWith('=') ? 1 : 0;
  const rawLen = clean.length + pad; // original base64 length including '='
  const len = (rawLen / 4) * 3 - pad;
  const out = new Uint8Array(len);
  let pos = 0;
  for (let i = 0; i < clean.length; i += 4) {
    const a = lookup[clean.charCodeAt(i)];
    const b = lookup[clean.charCodeAt(i + 1)];
    const c = lookup[clean.charCodeAt(i + 2)];
    const d = lookup[clean.charCodeAt(i + 3)];
    const n = (a << 18) | (b << 12) | (c << 6) | d;
    if (pos < len) out[pos++] = (n >> 16) & 0xFF;
    if (pos < len) out[pos++] = (n >> 8) & 0xFF;
    if (pos < len) out[pos++] = n & 0xFF;
  }
  return out;
}

function getAttr(xml: string, name: string): string {
  const re = new RegExp(`${name}="([^"]*)"`, 'i');
  const m = xml.match(re);
  return m ? m[1] : '';
}

interface EncryptionParams {
  keySalt: Uint8Array;
  passwordSalt: Uint8Array;
  spinCount: number;
  keyBits: number;
  encryptedVerifierInput: Uint8Array;
  encryptedVerifierHash: Uint8Array;
  encryptedKeyValue: Uint8Array;
  encryptedHmacKey: Uint8Array;
  encryptedHmacValue: Uint8Array;
}

function parseEncryptionInfo(xml: string): EncryptionParams {
  return {
    keySalt: b64Decode(getAttr(xml.split('keyData')[1] ?? xml, 'saltValue')),
    passwordSalt: b64Decode(getAttr(xml.split('encryptedKey')[1] ?? xml, 'saltValue')),
    spinCount: parseInt(getAttr(xml, 'spinCount'), 10) || 100000,
    keyBits: parseInt(getAttr(xml.split('encryptedKey')[1] ?? xml, 'keyBits'), 10) || 256,
    encryptedVerifierInput: b64Decode(getAttr(xml, 'encryptedVerifierHashInput')),
    encryptedVerifierHash: b64Decode(getAttr(xml, 'encryptedVerifierHashValue')),
    encryptedKeyValue: b64Decode(getAttr(xml, 'encryptedKeyValue')),
    encryptedHmacKey: b64Decode(getAttr(xml, 'encryptedHmacKey')),
    encryptedHmacValue: b64Decode(getAttr(xml, 'encryptedHmacValue')),
  };
}

// ── Public API ───────────────────────────────────────────────────────────────

export interface EncryptionOptions {
  /** Number of hash iterations (default 100000) */
  spinCount?: number;
}

// ── DataSpaces binary stream builders (MS-OFFCRYPTO §2.3.1) ──────────────────

/** Build a length-prefixed UTF-16LE string (uint32 byte-count + string data + padding to 4 bytes) */
function buildPrefixedUtf16(s: string): Uint8Array {
  const byteLen = s.length * 2;
  const padLen = (4 - (byteLen % 4)) % 4;
  const buf = new Uint8Array(4 + byteLen + padLen);
  setU32(buf, 0, byteLen);
  for (let i = 0; i < s.length; i++) setU16(buf, 4 + i * 2, s.charCodeAt(i));
  return buf;
}

function buildDataSpacesVersion(): Uint8Array {
  const feature = 'Microsoft.Container.DataSpaces';
  const strBytes = feature.length * 2; // 60
  const buf = new Uint8Array(4 + strBytes + 12); // length prefix + string + 3 version pairs
  setU32(buf, 0, strBytes);
  for (let i = 0; i < feature.length; i++) setU16(buf, 4 + i * 2, feature.charCodeAt(i));
  const off = 4 + strBytes;
  setU16(buf, off, 1); setU16(buf, off + 2, 0);     // reader 1.0
  setU16(buf, off + 4, 1); setU16(buf, off + 6, 0);  // updater 1.0
  setU16(buf, off + 8, 1); setU16(buf, off + 10, 0);  // writer 1.0
  return buf;
}

function buildDataSpaceMap(): Uint8Array {
  const refName = buildPrefixedUtf16('EncryptedPackage');
  const dsName = buildPrefixedUtf16('StrongEncryptionDataSpace');
  const entryLen = 4 /* refCompCount */ + 4 /* type */ + refName.length + dsName.length;
  const buf = new Uint8Array(8 + 4 + entryLen);
  setU32(buf, 0, 8);  // header length
  setU32(buf, 4, 1);  // entry count
  setU32(buf, 8, entryLen); // entry length
  let off = 12;
  setU32(buf, off, 1); off += 4; // reference component count
  setU32(buf, off, 0); off += 4; // type = stream
  buf.set(refName, off); off += refName.length;
  buf.set(dsName, off);
  return buf;
}

function buildStrongEncryptionDataSpace(): Uint8Array {
  const refName = buildPrefixedUtf16('StrongEncryptionTransform');
  const buf = new Uint8Array(8 + refName.length);
  setU32(buf, 0, 8);  // header length
  setU32(buf, 4, 1);  // transform reference count
  buf.set(refName, 8);
  return buf;
}

function buildPrimaryTransform(): Uint8Array {
  const id = buildPrefixedUtf16('{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}');
  const name = buildPrefixedUtf16('Microsoft.Container.EncryptionTransform');
  const contentLen = 4 /* type */ + id.length + name.length + 12 /* 3 version pairs */;
  const buf = new Uint8Array(4 + contentLen);
  setU32(buf, 0, contentLen); // transform length
  let off = 4;
  setU32(buf, off, 1); off += 4; // type = password
  buf.set(id, off); off += id.length;
  buf.set(name, off); off += name.length;
  // Versions: reader, updater, writer — each major.minor uint16 pairs
  setU16(buf, off, 1); setU16(buf, off + 2, 0); off += 4;
  setU16(buf, off, 1); setU16(buf, off + 2, 0); off += 4;
  setU16(buf, off, 1); setU16(buf, off + 2, 0);
  return buf;
}

/**
 * Build a CFB file with the DataSpaces structure required by MS-OFFCRYPTO.
 * Directory layout:
 *   Root → \x06DataSpaces (storage), EncryptionInfo (stream), EncryptedPackage (stream)
 *   \x06DataSpaces → Version, DataSpaceMap, DataSpaceInfo (storage), TransformInfo (storage)
 *   DataSpaceInfo → StrongEncryptionDataSpace
 *   TransformInfo → StrongEncryptionTransform (storage)
 *   StrongEncryptionTransform → \x06Primary
 */
function buildEncryptionCfb(encryptionInfo: Uint8Array, encryptedPackage: Uint8Array): Uint8Array {
  // Directory entries: [0]=Root, [1]=\x06DataSpaces, [2]=EncryptionInfo,
  // [3]=EncryptedPackage, [4]=Version, [5]=DataSpaceMap,
  // [6]=DataSpaceInfo, [7]=TransformInfo, [8]=StrongEncryptionDataSpace,
  // [9]=StrongEncryptionTransform, [10]=\x06Primary
  const versionData = buildDataSpacesVersion();
  const dsmData = buildDataSpaceMap();
  const sedsData = buildStrongEncryptionDataSpace();
  const primaryData = buildPrimaryTransform();

  interface DirEntry {
    name: string; type: number; /* 1=storage, 2=stream, 5=root */
    childId: number; leftId: number; rightId: number;
    data?: Uint8Array; startSector?: number; size: number;
  }

  const dirs: DirEntry[] = [
    { name: 'Root Entry', type: 5, childId: 2, leftId: -1, rightId: -1, size: 0 },
    // Children of root sorted by CFB name order (shorter < longer):
    // \x06DataSpaces (12) < EncryptionInfo (14) < EncryptedPackage (16)
    // BST root = EncryptionInfo
    { name: '\x06DataSpaces', type: 1, childId: 5, leftId: -1, rightId: -1, size: 0 },
    { name: 'EncryptionInfo', type: 2, childId: -1, leftId: 1, rightId: 3, data: encryptionInfo, size: encryptionInfo.length },
    { name: 'EncryptedPackage', type: 2, childId: -1, leftId: -1, rightId: -1, data: encryptedPackage, size: encryptedPackage.length },
    // Children of \x06DataSpaces sorted by CFB name order (shorter < longer):
    // Version (7) < DataSpaceMap (12) < DataSpaceInfo (13) < TransformInfo (13)
    // BST root = DataSpaceMap
    { name: 'Version', type: 2, childId: -1, leftId: -1, rightId: -1, data: versionData, size: versionData.length },
    { name: 'DataSpaceMap', type: 2, childId: -1, leftId: 4, rightId: 6, data: dsmData, size: dsmData.length },
    { name: 'DataSpaceInfo', type: 1, childId: 8, leftId: -1, rightId: 7, size: 0 },
    { name: 'TransformInfo', type: 1, childId: 9, leftId: -1, rightId: -1, size: 0 },
    // Child of DataSpaceInfo:
    { name: 'StrongEncryptionDataSpace', type: 2, childId: -1, leftId: -1, rightId: -1, data: sedsData, size: sedsData.length },
    // Child of TransformInfo:
    { name: 'StrongEncryptionTransform', type: 1, childId: 10, leftId: -1, rightId: -1, size: 0 },
    // Child of StrongEncryptionTransform:
    { name: '\x06Primary', type: 2, childId: -1, leftId: -1, rightId: -1, data: primaryData, size: primaryData.length },
  ];

  // Build mini-stream for small streams (< 4096 bytes), regular for large
  let miniData = new Uint8Array(0);
  let miniOffset = 0;
  type StreamInfo = { inMini: boolean; miniStart: number; regularStart: number };
  const streamInfo: StreamInfo[] = [];
  const regularStreams: { data: Uint8Array; startSector: number }[] = [];

  for (const d of dirs) {
    if (!d.data) { streamInfo.push({ inMini: false, miniStart: 0, regularStart: 0 }); continue; }
    if (d.data.length < MINI_CUT) {
      const padLen = Math.ceil(d.data.length / MINI_SZ) * MINI_SZ;
      const newMini = new Uint8Array(miniData.length + padLen);
      newMini.set(miniData); newMini.set(d.data, miniData.length);
      streamInfo.push({ inMini: true, miniStart: miniOffset / MINI_SZ, regularStart: 0 });
      miniOffset += padLen; miniData = newMini;
    } else {
      streamInfo.push({ inMini: false, miniStart: 0, regularStart: 0 });
    }
  }

  const dirCount = dirs.length;
  const dirSectors = Math.ceil((dirCount * DIR_SZ) / SECTOR_SZ);
  const miniStreamSectors = Math.ceil(miniData.length / SECTOR_SZ);
  let nextSector = 1 + dirSectors + miniStreamSectors; // FAT + dir + mini-stream

  for (let i = 0; i < dirs.length; i++) {
    if (dirs[i].data && !streamInfo[i].inMini) {
      streamInfo[i].regularStart = nextSector;
      regularStreams.push({ data: dirs[i].data!, startSector: nextSector });
      nextSector += Math.ceil(dirs[i].data!.length / SECTOR_SZ);
    }
  }

  // Build FAT
  const totalSectors = nextSector + 1; // +1 for mini-FAT sector
  const fatEntries = new Uint32Array(Math.max(totalSectors + 8, 128));
  fatEntries.fill(FREESECT);
  fatEntries[0] = FATSECT;
  for (let i = 1; i < 1 + dirSectors; i++)
    fatEntries[i] = i < dirSectors ? i + 1 : ENDOFCHAIN;
  for (let i = 1 + dirSectors; i < 1 + dirSectors + miniStreamSectors; i++)
    fatEntries[i] = i < dirSectors + miniStreamSectors ? i + 1 : ENDOFCHAIN;
  for (const rb of regularStreams) {
    const sects = Math.ceil(rb.data.length / SECTOR_SZ);
    for (let j = 0; j < sects; j++)
      fatEntries[rb.startSector + j] = j < sects - 1 ? rb.startSector + j + 1 : ENDOFCHAIN;
  }

  // Mini-FAT
  const miniSectorCount = Math.max(Math.ceil(miniData.length / MINI_SZ), 1);
  const miniFat = new Uint32Array(Math.max(miniSectorCount, 128));
  miniFat.fill(FREESECT);
  for (let i = 0; i < dirs.length; i++) {
    if (dirs[i].data && streamInfo[i].inMini) {
      const msects = Math.ceil(dirs[i].data!.length / MINI_SZ);
      const start = streamInfo[i].miniStart;
      for (let j = 0; j < msects; j++)
        miniFat[start + j] = j < msects - 1 ? start + j + 1 : ENDOFCHAIN;
    }
  }
  const miniFatSector = nextSector;
  fatEntries[miniFatSector] = ENDOFCHAIN;

  const fileSize = (1 + miniFatSector + 1) * SECTOR_SZ;
  const out = new Uint8Array(fileSize);

  // Header
  out.set(CFB_SIG, 0);
  setU16(out, 0x18, 0x003E); setU16(out, 0x1A, 0x0003);
  setU16(out, 0x1C, 0xFFFE); setU16(out, 0x1E, 9); setU16(out, 0x20, 6);
  setU32(out, 0x2C, 1); // FAT sectors
  setU32(out, 0x30, 1); // first dir sector
  setU32(out, 0x38, MINI_CUT);
  setU32(out, 0x3C, miniFatSector); setU32(out, 0x40, 1);
  setU32(out, 0x44, ENDOFCHAIN); setU32(out, 0x48, 0);
  for (let i = 0; i < 109; i++) setU32(out, 0x4C + i * 4, FREESECT);
  setU32(out, 0x4C, 0);

  // FAT sector
  const fatOff = SECTOR_SZ;
  for (let i = 0; i < 128; i++) setU32(out, fatOff + i * 4, fatEntries[i]);

  // Directory sectors
  const dirOff = SECTOR_SZ * 2;
  for (let i = 0; i < dirs.length; i++) {
    const d = dirs[i];
    const eOff = dirOff + i * DIR_SZ;
    const eName = encUtf16(d.name);
    out.set(eName.bytes, eOff);
    setU16(out, eOff + 0x40, eName.size);
    out[eOff + 0x42] = d.type;
    out[eOff + 0x43] = 1; // red
    setU32(out, eOff + 0x44, d.leftId >= 0 ? d.leftId : FREESECT);
    setU32(out, eOff + 0x48, d.rightId >= 0 ? d.rightId : FREESECT);
    setU32(out, eOff + 0x4C, d.childId >= 0 ? d.childId : FREESECT);

    if (d.type === 5) {
      // Root: mini-stream start sector + size
      setU32(out, eOff + 0x74, miniStreamSectors > 0 ? 1 + dirSectors : ENDOFCHAIN);
      setU32(out, eOff + 0x78, miniData.length);
    } else if (d.data) {
      if (streamInfo[i].inMini) {
        setU32(out, eOff + 0x74, streamInfo[i].miniStart);
      } else {
        setU32(out, eOff + 0x74, streamInfo[i].regularStart);
      }
      setU32(out, eOff + 0x78, d.data.length);
    } else {
      setU32(out, eOff + 0x74, ENDOFCHAIN);
      setU32(out, eOff + 0x78, 0);
    }
  }

  // Mini-stream data
  const miniOff = SECTOR_SZ * (1 + 1 + dirSectors);
  out.set(miniData, miniOff);

  // Regular streams
  for (const rb of regularStreams)
    out.set(rb.data, SECTOR_SZ * (1 + rb.startSector));

  // Mini-FAT sector
  const mfOff = SECTOR_SZ * (1 + miniFatSector);
  for (let i = 0; i < 128; i++) setU32(out, mfOff + i * 4, miniFat[i]);

  return out;
}

/**
 * Encrypt an XLSX/XLSM file with a password using OOXML Agile Encryption.
 * Returns a CFB binary container (.xlsx extension still works in Excel).
 *
 * @param xlsxData - The unencrypted XLSX file bytes
 * @param password - The password to encrypt with
 * @param options - Optional encryption parameters
 * @returns Encrypted file bytes (CFB container)
 */
export async function encryptWorkbook(xlsxData: Uint8Array, password: string, options?: EncryptionOptions): Promise<Uint8Array> {
  const spinCount = options?.spinCount ?? 100000;

  // Generate random salts
  const keySalt = randomBytes(16);
  const passwordSalt = randomBytes(16);
  const verifierInput = randomBytes(16);

  // 1. Derive the encryption key from password
  const keyDerived = await deriveKey(password, passwordSalt, spinCount, 256, BLOCK_KEY_ENCRYPTED_KEY);

  // Generate actual data encryption key
  const dataKey = randomBytes(32); // AES-256 key

  // 2. Encrypt the data key with the password-derived key
  // Per [MS-OFFCRYPTO] §2.3.6.2: IV for password key encryptor = raw passwordSalt
  const encKeyValue = await aesCbcEncrypt(keyDerived, passwordSalt, padToBlockSize(dataKey, 16));

  // 3. Create and encrypt verifier (IV = raw passwordSalt for all password key encryptor ops)
  const verifierKey = await deriveKey(password, passwordSalt, spinCount, 256, BLOCK_KEY_VERIFIER_INPUT);
  const encVerifierInput = await aesCbcEncrypt(verifierKey, passwordSalt, padToBlockSize(verifierInput, 16));

  const verifierHash = await sha512(verifierInput);
  const verifierHashKey = await deriveKey(password, passwordSalt, spinCount, 256, BLOCK_KEY_VERIFIER_VALUE);
  const encVerifierHash = await aesCbcEncrypt(verifierHashKey, passwordSalt, padToBlockSize(verifierHash, 16));

  // 4. Encrypt the package data
  // Per [MS-OFFCRYPTO] §2.3.6.1: EncryptedPackage = StreamSize(8) + EncryptedData
  // StreamSize is unencrypted; only the XLSX data is encrypted in 4096-byte segments
  const packageData = xlsxData;

  // Encrypt in 4096-byte segments
  const segmentSize = 4096;
  const encryptedSegments: Uint8Array[] = [];
  for (let offset = 0; offset < packageData.length; offset += segmentSize) {
    const segment = packageData.subarray(offset, Math.min(offset + segmentSize, packageData.length));
    const paddedSegment = padToBlockSize(segment, 16);

    // IV for each segment: SHA-512(keySalt + segmentIndex(LE32)), truncated to 16 bytes
    const segIdx = offset / segmentSize;
    const segIdxBuf = new Uint8Array(4);
    segIdxBuf[0] = segIdx & 0xFF; segIdxBuf[1] = (segIdx >> 8) & 0xFF;
    segIdxBuf[2] = (segIdx >> 16) & 0xFF; segIdxBuf[3] = (segIdx >> 24) & 0xFF;
    const segIvHash = await sha512(concat(keySalt, segIdxBuf));
    const segIv = segIvHash.subarray(0, 16);

    const encSegment = await aesCbcEncrypt(dataKey, segIv, paddedSegment);
    encryptedSegments.push(encSegment);
  }
  const encryptedData = concat(...encryptedSegments);

  // Prepend unencrypted StreamSize (8 bytes LE) per [MS-OFFCRYPTO]
  const streamSizeHeader = new Uint8Array(8);
  setU32(streamSizeHeader, 0, xlsxData.length);
  // High 32 bits at offset 4 stay 0 (files < 4GB)
  const encryptedPackage = concat(streamSizeHeader, encryptedData);

  // 5. HMAC for data integrity (encrypted with dataKey, IVs derived from keySalt)
  const hmacKey = randomBytes(64);
  const hmacKeyIv = (await sha512(concat(keySalt, BLOCK_KEY_DATA_INTEGRITY1))).subarray(0, 16);
  const encHmacKey = await aesCbcEncrypt(dataKey, hmacKeyIv, padToBlockSize(hmacKey, 16));

  // HMAC over the full EncryptedPackage stream (StreamSize + encrypted data)
  const hmacValue = await hmacSha512(hmacKey, encryptedPackage);
  const hmacValueIv = (await sha512(concat(keySalt, BLOCK_KEY_DATA_INTEGRITY2))).subarray(0, 16);
  const encHmacValue = await aesCbcEncrypt(dataKey, hmacValueIv, padToBlockSize(hmacValue, 16));

  // 6. Build EncryptionInfo stream
  const xmlStr = buildEncryptionInfoXml(
    keySalt, encKeyValue, encVerifierInput,
    encVerifierHash,
    encHmacKey, encHmacValue,
    passwordSalt, spinCount,
  );
  const xmlBytes = new TextEncoder().encode(xmlStr);
  // EncryptionInfo header: version (4,4) + reserved (0x40) + XML length
  const infoHeader = new Uint8Array(8);
  setU16(infoHeader, 0, 4); // major version
  setU16(infoHeader, 2, 4); // minor version
  setU32(infoHeader, 4, 0x00040); // flags = agile
  const encryptionInfo = concat(infoHeader, xmlBytes);

  // 7. Build CFB container with DataSpaces
  return buildEncryptionCfb(encryptionInfo, encryptedPackage);
}

/**
 * Decrypt a password-protected .xlsx file (OOXML Agile Encryption).
 *
 * @param encryptedData - The encrypted CFB file bytes
 * @param password - The password to decrypt with
 * @returns Decrypted XLSX file bytes
 * @throws Error if password is incorrect or file is not encrypted
 */
export async function decryptWorkbook(encryptedData: Uint8Array, password: string): Promise<Uint8Array> {
  // 1. Read CFB container
  const streams = readCfb(encryptedData);
  const infoStream = streams.find(s => s.name === 'EncryptionInfo');
  const pkgStream = streams.find(s => s.name === 'EncryptedPackage');
  if (!infoStream || !pkgStream) throw new Error('Not an encrypted Office file');

  // 2. Parse EncryptionInfo
  // Skip 8-byte header (version + flags)
  const xmlBytes = infoStream.data.subarray(8);
  const xmlStr = new TextDecoder().decode(xmlBytes);
  const params = parseEncryptionInfo(xmlStr);

  // 3. Derive key and verify password
  // Per [MS-OFFCRYPTO] §2.3.6.2: IV for password key encryptor = raw passwordSalt
  const keyDerived = await deriveKey(password, params.passwordSalt, params.spinCount, params.keyBits, BLOCK_KEY_ENCRYPTED_KEY);

  // Decrypt the actual data key
  let dataKey: Uint8Array;
  try {
    dataKey = await aesCbcDecrypt(keyDerived, params.passwordSalt, params.encryptedKeyValue);
    dataKey = dataKey.subarray(0, params.keyBits / 8);
  } catch {
    throw new Error('Incorrect password');
  }

  // Verify: decrypt verifier and check hash
  const verifierKey = await deriveKey(password, params.passwordSalt, params.spinCount, params.keyBits, BLOCK_KEY_VERIFIER_INPUT);
  let verifierInput: Uint8Array;
  try {
    verifierInput = await aesCbcDecrypt(verifierKey, params.passwordSalt, params.encryptedVerifierInput);
    verifierInput = verifierInput.subarray(0, 16);
  } catch {
    throw new Error('Incorrect password');
  }

  const verifierHashKey = await deriveKey(password, params.passwordSalt, params.spinCount, params.keyBits, BLOCK_KEY_VERIFIER_VALUE);
  try {
    const decVerifierHash = await aesCbcDecrypt(verifierHashKey, params.passwordSalt, params.encryptedVerifierHash);
    const expectedHash = await sha512(verifierInput);
    // Compare first 64 bytes
    for (let i = 0; i < 64; i++) {
      if (decVerifierHash[i] !== expectedHash[i]) throw new Error('Incorrect password');
    }
  } catch (e: any) {
    if (e.message === 'Incorrect password') throw e;
    throw new Error('Incorrect password');
  }

  // 4. Decrypt the package (keySalt from keyData)
  const keySalt = params.keySalt;

  // Per [MS-OFFCRYPTO]: first 8 bytes are unencrypted StreamSize
  const streamSize = u32(pkgStream.data, 0); // original XLSX size
  const encryptedPackage = pkgStream.data.subarray(8);

  const segmentSize = 4096;
  const decryptedSegments: Uint8Array[] = [];
  for (let offset = 0; offset < encryptedPackage.length; offset += segmentSize) {
    // Find segment end — encrypted segments are padded to 16-byte boundaries
    let segEnd = offset + segmentSize;
    // For the last segment, use remaining data
    if (segEnd > encryptedPackage.length) segEnd = encryptedPackage.length;
    // Round up to 16-byte boundary
    const encSegLen = Math.ceil((segEnd - offset) / 16) * 16;
    const encSegment = encryptedPackage.subarray(offset, offset + encSegLen);

    const segIdx = offset / segmentSize;
    const segIdxBuf = new Uint8Array(4);
    segIdxBuf[0] = segIdx & 0xFF; segIdxBuf[1] = (segIdx >> 8) & 0xFF;
    segIdxBuf[2] = (segIdx >> 16) & 0xFF; segIdxBuf[3] = (segIdx >> 24) & 0xFF;
    const segIvHash = await sha512(concat(keySalt, segIdxBuf));
    const segIv = segIvHash.subarray(0, 16);

    try {
      const decSegment = await aesCbcDecrypt(dataKey, segIv, encSegment);
      decryptedSegments.push(decSegment);
    } catch {
      throw new Error('Decryption failed — data may be corrupted');
    }
  }
  const decryptedData = concat(...decryptedSegments);

  // 5. Truncate to original stream size (removes padding from last segment)
  return decryptedData.subarray(0, streamSize);
}

/**
 * Check if a file is an encrypted Office document.
 */
export function isEncrypted(data: Uint8Array): boolean {
  // Check for CFB signature
  if (data.length < 8) return false;
  for (let i = 0; i < 8; i++) {
    if (data[i] !== CFB_SIG[i]) return false;
  }
  // Quick check: try to find EncryptionInfo in the CFB
  try {
    const streams = readCfb(data);
    return streams.some(s => s.name === 'EncryptionInfo');
  } catch {
    return false;
  }
}
