/**
 * MS-OVBA Compression / Decompression
 * Reference: [MS-OVBA] §2.4.1
 *
 * VBA source code inside vbaProject.bin is stored using this custom
 * LZ77-variant compression.  Each compressed container starts with a
 * 0x01 signature byte followed by one or more 4 096-byte chunks.
 */

// ── helpers ──────────────────────────────────────────────────────────────────

/** Number of *offset* bits in a CopyToken at a given decompressed position */
function bitCount(pos: number): number {
  if (pos <= 16)   return 4;
  if (pos <= 32)   return 5;
  if (pos <= 64)   return 6;
  if (pos <= 128)  return 7;
  if (pos <= 256)  return 8;
  if (pos <= 512)  return 9;
  if (pos <= 1024) return 10;
  if (pos <= 2048) return 11;
  return 12;
}

// ── compress ─────────────────────────────────────────────────────────────────

export function compressOvba(src: Uint8Array): Uint8Array {
  const out: number[] = [0x01]; // signature
  let srcPos = 0;

  while (srcPos < src.length) {
    const chunkStart = srcPos;
    const chunkEnd   = Math.min(srcPos + 4096, src.length);

    // Try compressing
    const compressed = compressChunk(src, chunkStart, chunkEnd);

    if (compressed.length < 4096) {
      // compressed form is smaller
      const totalChunkSize = compressed.length + 2; // +2 header
      const header = 0xB000 | ((totalChunkSize - 3) & 0x0FFF);
      out.push(header & 0xFF, (header >> 8) & 0xFF);
      for (const b of compressed) out.push(b);
    } else {
      // store raw (pad to 4096 with 0x00)
      const header = 0x3FFF; // flag=0, size=4098-3=4095=0x0FFF
      out.push(header & 0xFF, (header >> 8) & 0xFF);
      for (let i = chunkStart; i < chunkEnd; i++) out.push(src[i]);
      for (let i = chunkEnd - chunkStart; i < 4096; i++) out.push(0);
    }
    srcPos = chunkEnd;
  }
  return new Uint8Array(out);
}

function compressChunk(src: Uint8Array, start: number, end: number): number[] {
  const out: number[] = [];
  let dp = 0; // decompressed position within chunk

  while (start + dp < end) {
    let flagByte = 0;
    const flagIdx = out.length;
    out.push(0); // placeholder

    for (let bit = 0; bit < 8 && start + dp < end; bit++) {
      if (dp === 0) {
        out.push(src[start]);
        dp++;
        continue;
      }

      const bc   = bitCount(dp);
      const maxOff = 1 << bc;
      const lenBits = 16 - bc;
      const maxLen  = ((1 << lenBits) - 1) + 3;

      // Find longest match
      const searchStart = Math.max(0, dp - maxOff);
      let bestOff = 0, bestLen = 0;
      for (let s = searchStart; s < dp; s++) {
        let len = 0;
        while (dp + len < end - start && len < maxLen &&
               src[start + s + len] === src[start + dp + len]) len++;
        if (len > bestLen) { bestLen = len; bestOff = dp - s; }
      }

      if (bestLen >= 3) {
        flagByte |= 1 << bit;
        const token = ((bestOff - 1) << lenBits) | (bestLen - 3);
        out.push(token & 0xFF, (token >> 8) & 0xFF);
        dp += bestLen;
      } else {
        out.push(src[start + dp]);
        dp++;
      }
    }
    out[flagIdx] = flagByte;
  }
  return out;
}

// ── decompress ───────────────────────────────────────────────────────────────

export function decompressOvba(data: Uint8Array): Uint8Array {
  if (data.length === 0 || data[0] !== 0x01)
    throw new Error('Invalid OVBA compressed container (bad signature)');

  const out: number[] = [];
  let pos = 1;

  while (pos < data.length) {
    if (pos + 1 >= data.length) break;
    const header   = data[pos] | (data[pos + 1] << 8);
    pos += 2;
    const chunkSize     = (header & 0x0FFF) + 3;      // total incl. header
    const isCompressed  = (header >> 15) & 1;
    const chunkDataSize = chunkSize - 2;
    const chunkEnd      = pos + chunkDataSize;

    if (!isCompressed) {
      for (let i = 0; i < 4096 && pos < data.length; i++) out.push(data[pos++]);
      pos = Math.max(pos, chunkEnd); // skip padding
    } else {
      const chunkDecompStart = out.length;
      while (pos < chunkEnd && pos < data.length) {
        const flagByte = data[pos++];
        for (let bit = 0; bit < 8 && pos < chunkEnd && pos < data.length; bit++) {
          if ((flagByte & (1 << bit)) === 0) {
            out.push(data[pos++]);
          } else {
            if (pos + 1 >= data.length) break;
            const token = data[pos] | (data[pos + 1] << 8);
            pos += 2;
            const dp  = out.length - chunkDecompStart;
            const bc  = bitCount(dp);
            const lb  = 16 - bc;
            const off = (token >> lb) + 1;
            const len = (token & ((1 << lb) - 1)) + 3;
            const copyStart = out.length - off;
            for (let i = 0; i < len; i++) out.push(out[copyStart + i]);
          }
        }
      }
    }
  }
  return new Uint8Array(out);
}
