/**
 * Minimal ZIP reader — reads ZIP archives produced by Excel and other tools.
 * Handles STORE (method=0) and DEFLATE (method=8) compressed entries.
 */

export interface ZipReadEntry {
  name: string;
  data: Uint8Array;
}

function readUint16LE(buf: Uint8Array, o: number): number {
  return buf[o] | (buf[o + 1] << 8);
}
function readUint32LE(buf: Uint8Array, o: number): number {
  return (buf[o] | (buf[o+1] << 8) | (buf[o+2] << 16) | (buf[o+3] << 24)) >>> 0;
}

/** Inflate a raw DEFLATE stream (no zlib header) */
function inflate(data: Uint8Array): Uint8Array {
  // We use DecompressionStream if available (modern browsers + Node 18+)
  // For environments without it we fall back to a pure-JS implementation.
  // Since this runs synchronously we use a sync approach via shared buffer trick.
  throw new Error('inflate: DecompressionStream not available synchronously; use inflateAsync');
}

async function inflateAsync(data: Uint8Array): Promise<Uint8Array> {
  // DecompressionStream expects raw deflate
  if (typeof DecompressionStream !== 'undefined') {
    const ds = new (DecompressionStream as any)('deflate-raw');
    const writer = ds.writable.getWriter();
    const reader = ds.readable.getReader();
    writer.write(data);
    writer.close();
    const chunks: Uint8Array[] = [];
    let done = false;
    while (!done) {
      const { value, done: d } = await reader.read();
      if (value) chunks.push(value);
      done = d;
    }
    const total = chunks.reduce((s, c) => s + c.length, 0);
    const out = new Uint8Array(total);
    let pos = 0;
    for (const c of chunks) { out.set(c, pos); pos += c.length; }
    return out;
  }
  // Node.js fallback via zlib
  try {
    // @ts-ignore
    const { inflateRawSync } = await import('zlib');
    return inflateRawSync(data) as unknown as Uint8Array;
  } catch {
    throw new Error('No deflate implementation available. Use Node.js 18+ or a modern browser.');
  }
}

export async function readZip(data: Uint8Array): Promise<Map<string, ZipReadEntry>> {
  const dec = new TextDecoder('utf-8');
  const entries = new Map<string, ZipReadEntry>();

  // Find End of Central Directory (search from end)
  let eocdPos = -1;
  for (let i = data.length - 22; i >= 0; i--) {
    if (data[i] === 0x50 && data[i+1] === 0x4B && data[i+2] === 0x05 && data[i+3] === 0x06) {
      eocdPos = i;
      break;
    }
  }
  if (eocdPos < 0) throw new Error('Not a valid ZIP file (EOCD not found)');

  const cdOffset = readUint32LE(data, eocdPos + 16);
  const cdCount  = readUint16LE(data, eocdPos + 8);

  let pos = cdOffset;
  for (let i = 0; i < cdCount; i++) {
    if (readUint32LE(data, pos) !== 0x02014B50) throw new Error('Invalid central directory entry');
    const method      = readUint16LE(data, pos + 10);
    const compSize    = readUint32LE(data, pos + 20);
    const uncompSize  = readUint32LE(data, pos + 24);
    const nameLen     = readUint16LE(data, pos + 28);
    const extraLen    = readUint16LE(data, pos + 30);
    const commentLen  = readUint16LE(data, pos + 32);
    const localOffset = readUint32LE(data, pos + 42);
    const nameBytes   = data.subarray(pos + 46, pos + 46 + nameLen);
    const name        = dec.decode(nameBytes);
    pos += 46 + nameLen + extraLen + commentLen;

    // Read local file header
    const lhBase      = localOffset;
    if (readUint32LE(data, lhBase) !== 0x04034B50) throw new Error('Invalid local file header');
    const lhNameLen   = readUint16LE(data, lhBase + 26);
    const lhExtraLen  = readUint16LE(data, lhBase + 28);
    const dataStart   = lhBase + 30 + lhNameLen + lhExtraLen;
    const compData    = data.subarray(dataStart, dataStart + compSize);

    let fileData: Uint8Array;
    if (method === 0) {
      // STORE — no compression
      fileData = compData.slice();
    } else if (method === 8) {
      // DEFLATE
      fileData = await inflateAsync(compData);
    } else {
      // Unknown method — store raw, best effort
      fileData = compData.slice();
    }

    // Skip directory entries
    if (!name.endsWith('/')) {
      entries.set(name, { name, data: fileData });
    }
  }

  return entries;
}

/** Decode a ZIP entry as UTF-8 string */
export function entryText(entry: ZipReadEntry): string {
  return new TextDecoder('utf-8').decode(entry.data);
}
