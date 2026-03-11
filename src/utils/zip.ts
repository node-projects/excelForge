/**
 * ZIP writer with real DEFLATE compression — zero external dependencies.
 *
 * Compression pipeline:
 *   LZ77 (lazy matching, hash chains) → Huffman coding (dynamic trees)
 *
 * Compression levels:
 *   0  = STORE   (no compression, fastest)
 *   1  = FAST    (LZ77 + fixed Huffman, good speed)
 *   6  = DEFAULT (LZ77 lazy + dynamic Huffman, balanced — default)
 *   9  = BEST    (maximum LZ77 effort + dynamic Huffman, smallest output)
 *
 * The ZIP format uses raw DEFLATE (method 8) inside local file entries.
 * XML entries (the bulk of an XLSX) compress to ~10–20% of original size.
 */

// ─── CRC-32 ───────────────────────────────────────────────────────────────────

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

function crc32(data: Uint8Array): number {
  const t = crc32Table();
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < data.length; i++) crc = (crc >>> 8) ^ t[(crc ^ data[i]) & 0xFF];
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

// ─── Bit writer ───────────────────────────────────────────────────────────────

class BitWriter {
  buf: number[] = [];
  private bits = 0;
  private bitLen = 0;

  writeBits(val: number, n: number): void {
    this.bits |= (val & ((1 << n) - 1)) << this.bitLen;
    this.bitLen += n;
    while (this.bitLen >= 8) {
      this.buf.push(this.bits & 0xFF);
      this.bits >>>= 8;
      this.bitLen -= 8;
    }
  }

  flush(): void {
    if (this.bitLen > 0) { this.buf.push(this.bits & 0xFF); this.bits = 0; this.bitLen = 0; }
  }

  toBytes(): Uint8Array { return new Uint8Array(this.buf); }
}

// ─── Huffman coding ───────────────────────────────────────────────────────────

/** Build canonical Huffman code lengths from symbol frequencies */
function buildCodeLengths(freq: Uint32Array, maxBits: number): Uint8Array {
  const n = freq.length;
  const lengths = new Uint8Array(n);
  const nonZero = [];
  for (let i = 0; i < n; i++) if (freq[i] > 0) nonZero.push(i);
  if (nonZero.length === 0) return lengths;
  if (nonZero.length === 1) { lengths[nonZero[0]] = 1; return lengths; }

  // Package-merge algorithm for length-limited Huffman
  // Simplified: use standard Huffman + clamp to maxBits
  type Node = { freq: number; sym: number; left?: Node; right?: Node };
  let heap: Node[] = nonZero.map(s => ({ freq: freq[s], sym: s }));

  // Build min-heap by insertion sort for small sizes, or full heap for large
  const merge = (): void => {
    heap.sort((a, b) => a.freq - b.freq);
    while (heap.length > 1) {
      const a = heap.shift()!;
      const b = heap.shift()!;
      const node: Node = { freq: a.freq + b.freq, sym: -1, left: a, right: b };
      // Insert in sorted position
      let idx = 0;
      while (idx < heap.length && heap[idx].freq <= node.freq) idx++;
      heap.splice(idx, 0, node);
    }
  };
  merge();

  // Assign depths
  const setDepth = (node: Node | undefined, depth: number): void => {
    if (!node) return;
    if (node.sym >= 0) { lengths[node.sym] = Math.min(depth, maxBits); return; }
    setDepth(node.left, depth + 1);
    setDepth(node.right, depth + 1);
  };
  setDepth(heap[0], 0);

  // If any length exceeds maxBits, clamp and redistribute (simple approach)
  let excess = 0;
  for (let i = 0; i < n; i++) {
    if (lengths[i] > maxBits) { excess += (1 << (lengths[i] - maxBits)) - 1; lengths[i] = maxBits; }
  }
  // Redistribute excess by incrementing shortest codes (rare edge case)
  if (excess > 0) {
    for (let bl = maxBits - 1; bl >= 1 && excess > 0; bl--) {
      for (let i = 0; i < n && excess > 0; i++) {
        if (lengths[i] === bl) { lengths[i]++; excess--; }
      }
    }
  }

  return lengths;
}

/** Build canonical codes from lengths */
function buildCodes(lengths: Uint8Array): Uint32Array {
  const maxLen = Math.max(...lengths);
  const blCount = new Uint32Array(maxLen + 1);
  for (const l of lengths) if (l > 0) blCount[l]++;

  const nextCode = new Uint32Array(maxLen + 1);
  let code = 0;
  for (let bits = 1; bits <= maxLen; bits++) {
    code = (code + blCount[bits - 1]) << 1;
    nextCode[bits] = code;
  }

  const codes = new Uint32Array(lengths.length);
  for (let i = 0; i < lengths.length; i++) {
    if (lengths[i] > 0) codes[i] = nextCode[lengths[i]]++;
  }
  return codes;
}

/** Reverse bits (for DEFLATE bit ordering) */
function reverseBits(val: number, len: number): number {
  let r = 0;
  for (let i = 0; i < len; i++) { r = (r << 1) | (val & 1); val >>= 1; }
  return r;
}

// ─── Fixed Huffman tables (DEFLATE spec RFC 1951) ─────────────────────────────

function fixedLitLengths(): Uint8Array {
  const l = new Uint8Array(288);
  for (let i =   0; i <= 143; i++) l[i] = 8;
  for (let i = 144; i <= 255; i++) l[i] = 9;
  for (let i = 256; i <= 279; i++) l[i] = 7;
  for (let i = 280; i <= 287; i++) l[i] = 8;
  return l;
}

function fixedDistLengths(): Uint8Array {
  const l = new Uint8Array(32);
  l.fill(5);
  return l;
}

// ─── DEFLATE length/distance tables ───────────────────────────────────────────

// Length: extra bits and base values for symbols 257–285
const LENGTH_EXTRA = [0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0,0,0];
const LENGTH_BASE  = [3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258,0,0];

// Distance: extra bits and base values for codes 0–29
const DIST_EXTRA = [0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13];
const DIST_BASE  = [1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577];

function lenCode(len: number): [number, number, number] {
  // Returns [symbol, extraBits, extraVal]
  for (let i = 0; i < 29; i++) {
    if (len <= LENGTH_BASE[i] + (1 << LENGTH_EXTRA[i]) - 1 && len >= LENGTH_BASE[i]) {
      return [257 + i, LENGTH_EXTRA[i], len - LENGTH_BASE[i]];
    }
  }
  return [285, 0, 0];
}

function distCode(dist: number): [number, number, number] {
  for (let i = 0; i < 30; i++) {
    const max = DIST_BASE[i] + (1 << DIST_EXTRA[i]) - 1;
    if (dist >= DIST_BASE[i] && dist <= max) {
      return [i, DIST_EXTRA[i], dist - DIST_BASE[i]];
    }
  }
  return [29, 13, 0];
}

// ─── LZ77 ─────────────────────────────────────────────────────────────────────

const WSIZE     = 32768;   // window size
const MAX_MATCH = 258;
const MIN_MATCH = 3;
const CHAIN_LEN_FAST = 8;
const CHAIN_LEN_DEFAULT = 32;
const CHAIN_LEN_BEST = 128;

type Token = { lit: number } | { len: number; dist: number };

function lz77(data: Uint8Array, effort: number): Token[] {
  const chainLen = effort <= 1 ? CHAIN_LEN_FAST : effort <= 6 ? CHAIN_LEN_DEFAULT : CHAIN_LEN_BEST;
  const tokens: Token[] = [];
  const n = data.length;

  // Hash table: 3-byte hash → most recent position
  const HSIZE = 65536;
  const head  = new Int32Array(HSIZE).fill(-1);
  const prev  = new Int32Array(Math.min(n, WSIZE)).fill(-1);

  const hash3 = (pos: number) =>
    ((data[pos] * 0x1021 ^ data[pos+1] * 0x9B ^ data[pos+2]) & (HSIZE - 1)) >>> 0;

  let i = 0;
  let pendingLit = -1;

  const emitLit = (b: number) => { tokens.push({ lit: b }); };
  const emitMatch = (len: number, dist: number) => {
    if (pendingLit >= 0) { emitLit(pendingLit); pendingLit = -1; }
    tokens.push({ len, dist });
  };

  while (i < n) {
    if (i + 2 >= n) { emitLit(data[i++]); continue; }

    const h = hash3(i);
    let matchLen = MIN_MATCH - 1;
    let matchDist = 0;
    let chain = head[h];
    let steps = 0;

    while (chain >= 0 && steps < chainLen) {
      const dist = i - chain;
      if (dist > WSIZE) break;
      // Compare
      let mLen = 0;
      const limit = Math.min(MAX_MATCH, n - i);
      while (mLen < limit && data[chain + mLen] === data[i + mLen]) mLen++;
      if (mLen > matchLen) { matchLen = mLen; matchDist = dist; }
      if (matchLen === MAX_MATCH) break;
      chain = prev[chain & (WSIZE - 1)];
      steps++;
    }

    // Update hash chain
    prev[i & (WSIZE - 1)] = head[h];
    head[h] = i;

    if (matchLen >= MIN_MATCH) {
      // Lazy matching: check if next byte produces a longer match
      if (effort >= 6 && i + 3 < n) {
        const h2 = hash3(i + 1);
        let chain2 = head[h2];
        let lazyLen = 0, lazyDist = 0;
        let steps2 = 0;
        while (chain2 >= 0 && steps2 < chainLen) {
          const d2 = (i + 1) - chain2;
          if (d2 > WSIZE) break;
          let ml = 0;
          const lim = Math.min(MAX_MATCH, n - i - 1);
          while (ml < lim && data[chain2 + ml] === data[i + 1 + ml]) ml++;
          if (ml > lazyLen) { lazyLen = ml; lazyDist = d2; }
          chain2 = prev[chain2 & (WSIZE - 1)];
          steps2++;
        }
        if (lazyLen > matchLen + 1) {
          // Emit current byte as literal, use lazy match next iteration
          emitLit(data[i]);
          i++;
          prev[i & (WSIZE - 1)] = head[h2];
          head[h2] = i;
          emitMatch(lazyLen, lazyDist);
          // Advance past the match, keeping hash updated
          for (let k = 1; k < lazyLen; k++) {
            i++;
            if (i + 2 < n) {
              const hk = hash3(i);
              prev[i & (WSIZE - 1)] = head[hk];
              head[hk] = i;
            }
          }
          i++;
          continue;
        }
      }

      emitMatch(matchLen, matchDist);
      for (let k = 0; k < matchLen; k++) {
        if (i + k + 2 < n) {
          const hk = hash3(i + k);
          prev[(i + k) & (WSIZE - 1)] = head[hk];
          head[hk] = i + k;
        }
      }
      i += matchLen;
    } else {
      emitLit(data[i++]);
    }
  }

  return tokens;
}

// ─── DEFLATE block encoder ─────────────────────────────────────────────────────

/** Encode code-length tree for dynamic Huffman (DEFLATE §3.2.7) */
function encodeCodeLengths(lengths: number[], bw: BitWriter,
  clCodes: Uint32Array, clLens: Uint8Array): void {
  let i = 0;
  while (i < lengths.length) {
    const l = lengths[i];
    if (l === 0) {
      // Count zeros
      let run = 0;
      while (i + run < lengths.length && lengths[i + run] === 0 && run < 138) run++;
      if (run < 3) {
        bw.writeBits(reverseBits(clCodes[0], clLens[0]), clLens[0]); i++;
      } else if (run <= 10) {
        bw.writeBits(reverseBits(clCodes[17], clLens[17]), clLens[17]);
        bw.writeBits(run - 3, 3); i += run;
      } else {
        bw.writeBits(reverseBits(clCodes[18], clLens[18]), clLens[18]);
        bw.writeBits(run - 11, 7); i += run;
      }
    } else {
      bw.writeBits(reverseBits(clCodes[l], clLens[l]), clLens[l]);
      i++;
      // Check for repeat
      let run = 0;
      while (i + run < lengths.length && lengths[i + run] === l && run < 6) run++;
      if (run >= 3) {
        bw.writeBits(reverseBits(clCodes[16], clLens[16]), clLens[16]);
        bw.writeBits(run - 3, 2); i += run;
      }
    }
  }
}

function deflateBlock(
  tokens: Token[],
  bw: BitWriter,
  isLast: boolean,
  useDynamic: boolean,
): void {
  // Count frequencies
  const litFreq  = new Uint32Array(286);
  const distFreq = new Uint32Array(30);
  litFreq[256] = 1; // EOB always present

  for (const tok of tokens) {
    if ('lit' in tok) {
      litFreq[tok.lit]++;
    } else {
      const [lSym] = lenCode(tok.len);
      litFreq[lSym]++;
      const [dSym] = distCode(tok.dist);
      distFreq[dSym]++;
    }
  }

  let litLens: Uint8Array, distLens: Uint8Array;
  let litCodes: Uint32Array, distCodes: Uint32Array;

  if (useDynamic) {
    // Dynamic Huffman
    litLens  = buildCodeLengths(litFreq,  15);
    distLens = buildCodeLengths(distFreq, 15);
    // Ensure dist tree is non-empty
    if (distLens.every(l => l === 0)) { distLens[0] = 1; }
    litCodes  = buildCodes(litLens);
    distCodes = buildCodes(distLens);

    // Build code-length alphabet
    const hlit  = findMaxUsed(litLens,  257, 286) + 1;
    const hdist = findMaxUsed(distLens, 1,   30)  + 1;
    const allLens = [...litLens.subarray(0, hlit), ...distLens.subarray(0, hdist)];

    const clFreq = new Uint32Array(19);
    // Simulate RLE to count frequencies
    simulateCL(allLens, clFreq);
    const CL_ORDER = [16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];
    const clLens_  = buildCodeLengths(clFreq, 7);
    const clCodes  = buildCodes(clLens_);
    const hclen   = findMaxUsedArr(CL_ORDER.map(i => clLens_[i]), 4, 19) + 1;

    bw.writeBits(isLast ? 1 : 0, 1);
    bw.writeBits(2, 2); // dynamic Huffman
    bw.writeBits(hlit - 257, 5);
    bw.writeBits(hdist - 1,  5);
    bw.writeBits(hclen - 4,  4);
    for (let i = 0; i < hclen; i++) bw.writeBits(clLens_[CL_ORDER[i]], 3);
    encodeCodeLengths(allLens, bw, clCodes, clLens_);
  } else {
    // Fixed Huffman
    litLens  = fixedLitLengths();
    distLens = fixedDistLengths();
    litCodes  = buildCodes(litLens);
    distCodes = buildCodes(distLens);
    bw.writeBits(isLast ? 1 : 0, 1);
    bw.writeBits(1, 2); // fixed Huffman
  }

  // Emit tokens
  for (const tok of tokens) {
    if ('lit' in tok) {
      const l = litLens[tok.lit];
      bw.writeBits(reverseBits(litCodes[tok.lit], l), l);
    } else {
      const [lSym, lExtra, lVal] = lenCode(tok.len);
      const ll = litLens[lSym];
      bw.writeBits(reverseBits(litCodes[lSym], ll), ll);
      if (lExtra > 0) bw.writeBits(lVal, lExtra);

      const [dSym, dExtra, dVal] = distCode(tok.dist);
      const dl = distLens[dSym];
      bw.writeBits(reverseBits(distCodes[dSym], dl), dl);
      if (dExtra > 0) bw.writeBits(dVal, dExtra);
    }
  }

  // End of block
  const eobLen = litLens[256];
  bw.writeBits(reverseBits(litCodes[256], eobLen), eobLen);
}

function findMaxUsed(arr: Uint8Array, min: number, max: number): number {
  let r = min - 1;
  for (let i = min; i < Math.min(arr.length, max); i++) if (arr[i] > 0) r = i;
  return Math.max(r, min - 1);
}

function findMaxUsedArr(arr: number[], min: number, max: number): number {
  let r = min - 1;
  for (let i = min; i < Math.min(arr.length, max); i++) if (arr[i] > 0) r = i;
  return Math.max(r, min - 1);
}

function simulateCL(lengths: number[], freq: Uint32Array): void {
  let i = 0;
  while (i < lengths.length) {
    const l = lengths[i];
    if (l === 0) {
      let run = 0;
      while (i + run < lengths.length && lengths[i + run] === 0 && run < 138) run++;
      if (run < 3) { freq[0]++; i++; }
      else if (run <= 10) { freq[17]++; i += run; }
      else { freq[18]++; i += run; }
    } else {
      freq[l]++; i++;
      let run = 0;
      while (i + run < lengths.length && lengths[i + run] === l && run < 6) run++;
      if (run >= 3) { freq[16]++; i += run; }
    }
  }
}

// ─── Main deflate function ─────────────────────────────────────────────────────

/** BLOCK_SIZE: split input into blocks of this size for better streaming */
const BLOCK_SIZE = 65536;

/**
 * Compress data using DEFLATE (raw, no zlib wrapper).
 * @param data  Input bytes
 * @param level 0=store, 1=fast, 6=default, 9=best
 */
export function deflateRaw(data: Uint8Array, level = 6): Uint8Array {
  if (level === 0) {
    // STORE blocks
    const bw = new BitWriter();
    let offset = 0;
    while (offset < data.length || data.length === 0) {
      const size = Math.min(BLOCK_SIZE, data.length - offset);
      const isLast = offset + size >= data.length;
      bw.flush();
      bw.writeBits(isLast ? 1 : 0, 1);
      bw.writeBits(0, 2); // BTYPE = no compression
      bw.flush(); // align to byte
      // Write len and ~len
      const len = size;
      bw.buf.push(len & 0xFF, (len >> 8) & 0xFF, (~len) & 0xFF, ((~len) >> 8) & 0xFF);
      for (let i = 0; i < size; i++) bw.buf.push(data[offset + i]);
      offset += size;
      if (data.length === 0) break;
    }
    return bw.toBytes();
  }

  const effort = Math.max(1, Math.min(9, level));
  const useDynamic = effort >= 2;
  const bw = new BitWriter();

  // Split into blocks and compress each
  let offset = 0;
  while (offset < data.length || data.length === 0) {
    const chunk = data.subarray(offset, offset + BLOCK_SIZE);
    const isLast = offset + BLOCK_SIZE >= data.length;
    const tokens = lz77(chunk, effort);
    deflateBlock(tokens, bw, isLast, useDynamic);
    offset += chunk.length;
    if (data.length === 0) { deflateBlock([], bw, true, useDynamic); break; }
  }

  bw.flush();
  return bw.toBytes();
}

// ─── ZIP writer ───────────────────────────────────────────────────────────────

function writeUint16LE(buf: Uint8Array, o: number, v: number) {
  buf[o] = v & 0xFF; buf[o+1] = (v >> 8) & 0xFF;
}
function writeUint32LE(buf: Uint8Array, o: number, v: number) {
  buf[o] = v & 0xFF; buf[o+1] = (v >> 8) & 0xFF;
  buf[o+2] = (v >> 16) & 0xFF; buf[o+3] = (v >> 24) & 0xFF;
}

const textEnc = new TextEncoder();

export interface ZipEntry {
  name: string;
  data: Uint8Array;
  /** Override per-entry compression level (0–9). Default: uses buildZip's level */
  level?: number;
}

export interface ZipOptions {
  /**
   * Compression level for all entries (unless overridden per-entry).
   * 0 = STORE (no compression)
   * 1 = fastest
   * 6 = default (good balance — recommended for XLSX)
   * 9 = maximum compression (slower)
   * Default: 6
   */
  level?: number;
  /**
   * File extensions that should always be stored uncompressed
   * (e.g. already-compressed images).
   * Default: ['png', 'jpg', 'jpeg', 'gif', 'tiff', 'emf', 'wmf']
   */
  noCompress?: string[];
}

const DEFAULT_NO_COMPRESS = new Set(['png','jpg','jpeg','gif','tiff','emf','wmf','bmp','webp']);

export function buildZip(entries: ZipEntry[], opts: ZipOptions = {}): Uint8Array {
  const globalLevel = opts.level ?? 6;
  const noCompress  = opts.noCompress
    ? new Set(opts.noCompress.map(e => e.toLowerCase()))
    : DEFAULT_NO_COMPRESS;

  const localParts: Uint8Array[] = [];
  const centralDir: Uint8Array[] = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = textEnc.encode(entry.name);
    const ext = entry.name.split('.').pop()?.toLowerCase() ?? '';

    // Decide compression level for this entry
    const entryLevel = entry.level ?? (noCompress.has(ext) ? 0 : globalLevel);
    const useDeflate = entryLevel > 0 && entry.data.length > 0;

    const rawCrc = crc32(entry.data);
    let compData: Uint8Array;
    let method: number;

    if (useDeflate) {
      compData = deflateRaw(entry.data, entryLevel);
      // Only use compressed data if it's actually smaller
      if (compData.length >= entry.data.length) {
        compData = entry.data;
        method   = 0; // STORE
      } else {
        method   = 8; // DEFLATE
      }
    } else {
      compData = entry.data;
      method   = 0;
    }

    // ── Local file header (30 + name) ───────────────────────────────────────
    const lh = new Uint8Array(30 + nameBytes.length);
    writeUint32LE(lh, 0,  0x04034B50);       // signature
    writeUint16LE(lh, 4,  20);               // version needed (2.0)
    writeUint16LE(lh, 6,  0);                // flags
    writeUint16LE(lh, 8,  method);
    writeUint16LE(lh, 10, 0);                // mod time
    writeUint16LE(lh, 12, 0);                // mod date
    writeUint32LE(lh, 14, rawCrc);
    writeUint32LE(lh, 18, compData.length);  // compressed size
    writeUint32LE(lh, 22, entry.data.length);// uncompressed size
    writeUint16LE(lh, 26, nameBytes.length);
    writeUint16LE(lh, 28, 0);               // extra field length
    lh.set(nameBytes, 30);

    // ── Central directory entry (46 + name) ─────────────────────────────────
    const cd = new Uint8Array(46 + nameBytes.length);
    writeUint32LE(cd, 0,  0x02014B50);
    writeUint16LE(cd, 4,  20);              // version made by
    writeUint16LE(cd, 6,  20);              // version needed
    writeUint16LE(cd, 8,  0);              // flags
    writeUint16LE(cd, 10, method);
    writeUint16LE(cd, 12, 0);              // mod time
    writeUint16LE(cd, 14, 0);              // mod date
    writeUint32LE(cd, 16, rawCrc);
    writeUint32LE(cd, 20, compData.length);
    writeUint32LE(cd, 24, entry.data.length);
    writeUint16LE(cd, 28, nameBytes.length);
    writeUint16LE(cd, 30, 0);              // extra
    writeUint16LE(cd, 32, 0);              // comment
    writeUint16LE(cd, 34, 0);              // disk start
    writeUint16LE(cd, 36, 0);              // internal attr
    writeUint32LE(cd, 38, 0);              // external attr
    writeUint32LE(cd, 42, offset);         // local header offset
    cd.set(nameBytes, 46);

    localParts.push(lh, compData);
    centralDir.push(cd);
    offset += lh.length + compData.length;
  }

  // ── End of central directory ─────────────────────────────────────────────
  const cdStart = offset;
  const cdSize  = centralDir.reduce((s, c) => s + c.length, 0);

  const eocd = new Uint8Array(22);
  writeUint32LE(eocd, 0,  0x06054B50);
  writeUint16LE(eocd, 4,  0);
  writeUint16LE(eocd, 6,  0);
  writeUint16LE(eocd, 8,  centralDir.length);
  writeUint16LE(eocd, 10, centralDir.length);
  writeUint32LE(eocd, 12, cdSize);
  writeUint32LE(eocd, 16, cdStart);
  writeUint16LE(eocd, 20, 0);

  const all   = [...localParts, ...centralDir, eocd];
  const total = all.reduce((s, a) => s + a.length, 0);
  const out   = new Uint8Array(total);
  let pos = 0;
  for (const a of all) { out.set(a, pos); pos += a.length; }
  return out;
}
