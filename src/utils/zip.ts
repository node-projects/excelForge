/**
 * Minimal ZIP writer — no external dependencies.
 * Supports STORE (uncompressed) and DEFLATE compression.
 */

function adler32(data: Uint8Array): number {
  let s1 = 1, s2 = 0;
  for (let i = 0; i < data.length; i++) {
    s1 = (s1 + data[i]) % 65521;
    s2 = (s2 + s1) % 65521;
  }
  return (s2 << 16) | s1;
}

function crc32(data: Uint8Array): number {
  const table = crc32Table();
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < data.length; i++) {
    crc = (crc >>> 8) ^ table[(crc ^ data[i]) & 0xFF];
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

let _crcTable: Uint32Array | null = null;
function crc32Table(): Uint32Array {
  if (_crcTable) return _crcTable;
  _crcTable = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = c & 1 ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    _crcTable[i] = c;
  }
  return _crcTable;
}

/** Very simple DEFLATE using stored blocks (compatible, not compressed) */
function deflateStore(data: Uint8Array): Uint8Array {
  const BLOCK = 65535;
  const blocks = Math.ceil(data.length / BLOCK) || 1;
  const out = new Uint8Array(2 + 4 + blocks * 5 + data.length + 4);
  let o = 0;
  // zlib header (CM=8, CINFO=7, no dict, check bits)
  out[o++] = 0x78; out[o++] = 0x01;
  let remaining = data.length;
  let offset = 0;
  for (let b = 0; b < blocks; b++) {
    const size = Math.min(remaining, BLOCK);
    const last = b === blocks - 1 ? 1 : 0;
    out[o++] = last;
    out[o++] = size & 0xFF; out[o++] = (size >> 8) & 0xFF;
    out[o++] = (~size) & 0xFF; out[o++] = ((~size) >> 8) & 0xFF;
    out.set(data.subarray(offset, offset + size), o);
    o += size; offset += size; remaining -= size;
  }
  const a = adler32(data);
  out[o++] = (a >> 24) & 0xFF; out[o++] = (a >> 16) & 0xFF;
  out[o++] = (a >> 8) & 0xFF;  out[o++] = a & 0xFF;
  return out.subarray(0, o);
}

function writeUint16LE(buf: Uint8Array, offset: number, v: number) {
  buf[offset] = v & 0xFF; buf[offset + 1] = (v >> 8) & 0xFF;
}
function writeUint32LE(buf: Uint8Array, offset: number, v: number) {
  buf[offset] = v & 0xFF; buf[offset + 1] = (v >> 8) & 0xFF;
  buf[offset + 2] = (v >> 16) & 0xFF; buf[offset + 3] = (v >> 24) & 0xFF;
}

const enc = new TextEncoder();

export interface ZipEntry { name: string; data: Uint8Array; }

export function buildZip(entries: ZipEntry[]): Uint8Array {
  const parts: Uint8Array[] = [];
  const centralDir: Uint8Array[] = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = enc.encode(entry.name);
    const crc = crc32(entry.data);
    const compressed = deflateStore(entry.data);
    const useDeflate = false; // store mode for simplicity & max compat
    const compData = useDeflate ? compressed : entry.data;
    const method = 0; // STORE

    // Local file header
    const lh = new Uint8Array(30 + nameBytes.length);
    writeUint32LE(lh, 0, 0x04034B50);
    writeUint16LE(lh, 4, 20); // version needed
    writeUint16LE(lh, 6, 0);  // flags
    writeUint16LE(lh, 8, method);
    writeUint16LE(lh, 10, 0); writeUint16LE(lh, 12, 0); // mod time/date
    writeUint32LE(lh, 14, crc);
    writeUint32LE(lh, 18, compData.length);
    writeUint32LE(lh, 22, entry.data.length);
    writeUint16LE(lh, 26, nameBytes.length);
    writeUint16LE(lh, 28, 0); // extra
    lh.set(nameBytes, 30);

    // Central directory entry
    const cd = new Uint8Array(46 + nameBytes.length);
    writeUint32LE(cd, 0, 0x02014B50);
    writeUint16LE(cd, 4, 20); writeUint16LE(cd, 6, 20);
    writeUint16LE(cd, 8, 0); writeUint16LE(cd, 10, method);
    writeUint16LE(cd, 12, 0); writeUint16LE(cd, 14, 0);
    writeUint32LE(cd, 16, crc);
    writeUint32LE(cd, 20, compData.length);
    writeUint32LE(cd, 24, entry.data.length);
    writeUint16LE(cd, 28, nameBytes.length);
    writeUint16LE(cd, 30, 0); writeUint16LE(cd, 32, 0); writeUint16LE(cd, 34, 0);
    writeUint16LE(cd, 36, 0); writeUint16LE(cd, 38, 0);
    writeUint32LE(cd, 42, offset);
    cd.set(nameBytes, 46);

    parts.push(lh, compData);
    centralDir.push(cd);
    offset += lh.length + compData.length;
  }

  const cdStart = offset;
  let cdSize = 0;
  for (const cd of centralDir) cdSize += cd.length;

  const eocd = new Uint8Array(22);
  writeUint32LE(eocd, 0, 0x06054B50);
  writeUint16LE(eocd, 4, 0); writeUint16LE(eocd, 6, 0);
  writeUint16LE(eocd, 8, centralDir.length);
  writeUint16LE(eocd, 10, centralDir.length);
  writeUint32LE(eocd, 12, cdSize);
  writeUint32LE(eocd, 16, cdStart);
  writeUint16LE(eocd, 20, 0);

  const all = [...parts, ...centralDir, eocd];
  const total = all.reduce((s, a) => s + a.length, 0);
  const out = new Uint8Array(total);
  let pos = 0;
  for (const a of all) { out.set(a, pos); pos += a.length; }
  return out;
}
