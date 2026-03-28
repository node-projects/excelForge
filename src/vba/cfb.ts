/**
 * Minimal CFB (Compound Binary File / OLE2) reader & writer.
 * Used exclusively for vbaProject.bin.
 *
 * Reference: [MS-CFB] — Microsoft Compound File Binary File Format
 *
 * Implementation constraints (by design):
 *  – v3 only (512-byte sectors, 64-byte mini-sectors)
 *  – All streams < 4096 bytes → mini-stream only
 *  – Single FAT sector (≤ 128 total sectors)
 */

// ── constants ────────────────────────────────────────────────────────────────

const SIGNATURE     = new Uint8Array([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
const ENDOFCHAIN    = 0xFFFFFFFE;
const FREESECT      = 0xFFFFFFFF;
const FATSECT       = 0xFFFFFFFD;
const NOSTREAM      = 0xFFFFFFFF;
const SECTOR_SIZE   = 512;
const MINI_SIZE     = 64;
const MINI_CUTOFF   = 0x1000; // 4096
const DIR_ENTRY_SZ  = 128;

// ── public types ─────────────────────────────────────────────────────────────

export interface CfbStream {
  name:    string;
  data:    Uint8Array;
  storage?: string;          // parent storage name (e.g. "VBA")
}

// ── helpers ──────────────────────────────────────────────────────────────────

function u16(buf: Uint8Array, off: number): number {
  return buf[off] | (buf[off + 1] << 8);
}
function u32(buf: Uint8Array, off: number): number {
  return (buf[off] | (buf[off + 1] << 8) | (buf[off + 2] << 16) | (buf[off + 3] << 24)) >>> 0;
}
function setU16(buf: Uint8Array, off: number, v: number): void {
  buf[off] = v & 0xFF; buf[off + 1] = (v >> 8) & 0xFF;
}
function setU32(buf: Uint8Array, off: number, v: number): void {
  buf[off] = v & 0xFF; buf[off + 1] = (v >> 8) & 0xFF;
  buf[off + 2] = (v >> 16) & 0xFF; buf[off + 3] = (v >> 24) & 0xFF;
}

/** UTF-16LE encode a name (null-terminated) into 64-byte buffer, return byte size including null. */
function encodeName(name: string): { bytes: Uint8Array; size: number } {
  const bytes = new Uint8Array(64);
  let i = 0;
  for (const ch of name) {
    const code = ch.charCodeAt(0);
    bytes[i++] = code & 0xFF;
    bytes[i++] = (code >> 8) & 0xFF;
  }
  // null terminator
  bytes[i++] = 0; bytes[i++] = 0;
  return { bytes, size: i };
}

/** Decode UTF-16LE name from directory entry. */
function decodeName(buf: Uint8Array, off: number, nameSize: number): string {
  let s = '';
  const len = Math.max(0, nameSize - 2); // exclude null terminator
  for (let i = 0; i < len; i += 2) {
    const code = buf[off + i] | (buf[off + i + 1] << 8);
    if (code === 0) break;
    s += String.fromCharCode(code);
  }
  return s;
}

function followChain(fat: Uint32Array, start: number): number[] {
  const chain: number[] = [];
  let sid = start;
  while (sid !== ENDOFCHAIN && sid < fat.length) {
    chain.push(sid);
    sid = fat[sid];
  }
  return chain;
}

function ceilDiv(a: number, b: number): number { return Math.ceil(a / b); }

// ── READ ─────────────────────────────────────────────────────────────────────

export function readCfb(data: Uint8Array): CfbStream[] {
  // ── header ──────────────────────────────────────────────────────────────
  for (let i = 0; i < 8; i++) if (data[i] !== SIGNATURE[i]) throw new Error('Not a CFB file');
  const sectorPow  = u16(data, 30);
  const sectorSize = 1 << sectorPow;
  const miniPow    = u16(data, 32);
  const miniSize   = 1 << miniPow;
  const fatSectors = u32(data, 44);
  const dirStart   = u32(data, 48);
  const miniFatStart = u32(data, 60);
  const miniFatCount = u32(data, 64);

  const sectorOff = (sid: number) => 512 + sid * sectorSize;

  // ── build FAT ───────────────────────────────────────────────────────────
  const difat: number[] = [];
  for (let i = 0; i < 109 && i < fatSectors; i++) {
    const sid = u32(data, 76 + i * 4);
    if (sid !== FREESECT) difat.push(sid);
  }
  const totalSectors = Math.floor((data.length - 512) / sectorSize);
  const fat = new Uint32Array(totalSectors);
  fat.fill(FREESECT);
  for (let fi = 0; fi < difat.length; fi++) {
    const off = sectorOff(difat[fi]);
    const entriesPerSector = sectorSize / 4;
    for (let j = 0; j < entriesPerSector && fi * entriesPerSector + j < totalSectors; j++) {
      fat[fi * entriesPerSector + j] = u32(data, off + j * 4);
    }
  }

  // ── read directory ──────────────────────────────────────────────────────
  const dirChain = followChain(fat, dirStart);
  const dirBuf   = new Uint8Array(dirChain.length * sectorSize);
  for (let i = 0; i < dirChain.length; i++) {
    dirBuf.set(data.subarray(sectorOff(dirChain[i]), sectorOff(dirChain[i]) + sectorSize), i * sectorSize);
  }
  const numEntries = dirBuf.length / DIR_ENTRY_SZ;

  interface DirEntry {
    name: string; type: number; child: number; left: number; right: number;
    startSector: number; size: number; index: number;
  }
  const entries: DirEntry[] = [];
  for (let i = 0; i < numEntries; i++) {
    const off      = i * DIR_ENTRY_SZ;
    const nameSize = u16(dirBuf, off + 64);
    if (nameSize === 0) continue;
    entries.push({
      name: decodeName(dirBuf, off, nameSize),
      type: dirBuf[off + 66],
      child: u32(dirBuf, off + 76),
      left: u32(dirBuf, off + 68),
      right: u32(dirBuf, off + 72),
      startSector: u32(dirBuf, off + 116),
      size: u32(dirBuf, off + 120),
      index: i,
    });
  }

  // root entry
  const root = entries.find(e => e.type === 5);
  if (!root) throw new Error('No root entry in CFB');

  // ── mini-stream container ───────────────────────────────────────────────
  const miniContainerChain = followChain(fat, root.startSector);
  const miniContainer = new Uint8Array(miniContainerChain.length * sectorSize);
  for (let i = 0; i < miniContainerChain.length; i++) {
    miniContainer.set(data.subarray(sectorOff(miniContainerChain[i]), sectorOff(miniContainerChain[i]) + sectorSize), i * sectorSize);
  }

  // ── mini-FAT ────────────────────────────────────────────────────────────
  const miniFatChain = followChain(fat, miniFatStart);
  const numMiniEntries = miniFatChain.length * (sectorSize / 4);
  const miniFat = new Uint32Array(numMiniEntries);
  miniFat.fill(FREESECT);
  for (let i = 0; i < miniFatChain.length; i++) {
    const off = sectorOff(miniFatChain[i]);
    const count = sectorSize / 4;
    for (let j = 0; j < count; j++) {
      miniFat[i * count + j] = u32(data, off + j * 4);
    }
  }

  // ── resolve parent storages ─────────────────────────────────────────────
  const parentMap = new Map<number, string>();
  function collectChildren(parentIdx: number, parentName: string) {
    const parent = entries.find(e => e.index === parentIdx);
    if (!parent || parent.child === NOSTREAM) return;
    const visit = (idx: number) => {
      if (idx === NOSTREAM || idx >= numEntries) return;
      const e = entries.find(en => en.index === idx);
      if (!e) return;
      parentMap.set(idx, parentName);
      if (e.type === 1) collectChildren(idx, e.name); // recurse into sub-storages
      visit(e.left);
      visit(e.right);
    };
    visit(parent.child);
  }
  collectChildren(root.index, '');

  // ── extract streams ─────────────────────────────────────────────────────
  const streams: CfbStream[] = [];
  for (const entry of entries) {
    if (entry.type !== 2) continue; // only streams
    let streamData: Uint8Array;
    if (entry.size < MINI_CUTOFF) {
      // read from mini-stream
      const chain = followChain(miniFat, entry.startSector);
      const buf = new Uint8Array(chain.length * miniSize);
      for (let i = 0; i < chain.length; i++) {
        const off = chain[i] * miniSize;
        buf.set(miniContainer.subarray(off, off + miniSize), i * miniSize);
      }
      streamData = buf.subarray(0, entry.size);
    } else {
      // read from regular sectors
      const chain = followChain(fat, entry.startSector);
      const buf = new Uint8Array(chain.length * sectorSize);
      for (let i = 0; i < chain.length; i++) {
        buf.set(data.subarray(sectorOff(chain[i]), sectorOff(chain[i]) + sectorSize), i * sectorSize);
      }
      streamData = buf.subarray(0, entry.size);
    }
    streams.push({
      name: entry.name,
      data: streamData,
      storage: parentMap.get(entry.index) || undefined,
    });
  }
  return streams;
}

// ── WRITE ────────────────────────────────────────────────────────────────────

interface DirNode {
  name: string;
  type: 1 | 2 | 5;          // 1=storage, 2=stream, 5=root
  children?: DirNode[];
  data?: Uint8Array;
}

export function buildCfb(streams: CfbStream[]): Uint8Array {
  // ── build tree ──────────────────────────────────────────────────────────
  const root: DirNode = { name: 'Root Entry', type: 5, children: [] };
  const storages = new Map<string, DirNode>();
  storages.set('', root);

  for (const s of streams) {
    const parentName = s.storage ?? '';
    let parent = storages.get(parentName);
    if (!parent) {
      parent = { name: parentName, type: 1, children: [] };
      storages.set(parentName, parent);
      root.children!.push(parent);
    }
    const node: DirNode = { name: s.name, type: 2, data: s.data };
    parent.children!.push(node);
  }

  // Flatten into directory entries (BFS)
  interface FlatEntry {
    node: DirNode;
    parentIdx: number;
    childIdx: number;  // will be set
    leftIdx: number;   // will be set
    rightIdx: number;  // will be set
    miniStart: number; // will be set
    size: number;
  }
  const flat: FlatEntry[] = [];
  const queue: DirNode[] = [root];
  while (queue.length) {
    const n = queue.shift()!;
    flat.push({
      node: n, parentIdx: -1, childIdx: NOSTREAM,
      leftIdx: NOSTREAM, rightIdx: NOSTREAM,
      miniStart: 0, size: n.data?.length ?? 0,
    });
    if (n.children) {
      for (const c of n.children) queue.push(c);
    }
  }

  // Set child / sibling links.
  // For each storage/root node, set its child to the first child entry,
  // then chain all children as right-siblings.
  for (let i = 0; i < flat.length; i++) {
    const n = flat[i].node;
    if (!n.children?.length) continue;
    const childIdxs = flat
      .map((e, idx) => ({ e, idx }))
      .filter(({ e }) => n.children!.includes(e.node))
      .map(({ idx }) => idx);
    if (childIdxs.length === 0) continue;
    // Build a balanced-ish binary tree for the children
    // MS-CFB §2.6.4: sort by name length first, then by uppercased content
    const sorted = childIdxs.sort((a, b) => {
      const na = flat[a].node.name, nb = flat[b].node.name;
      if (na.length !== nb.length) return na.length - nb.length;
      return na.toUpperCase() < nb.toUpperCase() ? -1 : na.toUpperCase() > nb.toUpperCase() ? 1 : 0;
    });
    const buildTree = (arr: number[]): number => {
      if (arr.length === 0) return NOSTREAM;
      const mid = arr.length >> 1;
      flat[arr[mid]].leftIdx = buildTree(arr.slice(0, mid));
      flat[arr[mid]].rightIdx = buildTree(arr.slice(mid + 1));
      return arr[mid];
    };
    flat[i].childIdx = buildTree(sorted);
  }

  // ── pack mini-stream ────────────────────────────────────────────────────
  const miniSectorData: number[] = [];
  const miniChains: Map<number, number[]> = new Map(); // flatIdx → mini-sector chain
  let nextMiniSector = 0;

  for (let i = 0; i < flat.length; i++) {
    const d = flat[i].node.data;
    if (!d || d.length === 0) continue;
    const numMiniSectors = ceilDiv(d.length, MINI_SIZE);
    const chain: number[] = [];
    for (let m = 0; m < numMiniSectors; m++) {
      chain.push(nextMiniSector++);
      const chunk = d.subarray(m * MINI_SIZE, Math.min((m + 1) * MINI_SIZE, d.length));
      for (const b of chunk) miniSectorData.push(b);
      // pad mini-sector to 64 bytes
      for (let p = chunk.length; p < MINI_SIZE; p++) miniSectorData.push(0);
    }
    flat[i].miniStart = chain[0];
    miniChains.set(i, chain);
  }

  // ── build mini-FAT ──────────────────────────────────────────────────────
  const totalMiniSectors = nextMiniSector;
  const miniFatEntries   = Math.max(totalMiniSectors, 1);
  const miniFatSectors   = ceilDiv(miniFatEntries * 4, SECTOR_SIZE);
  const miniFat          = new Uint32Array(miniFatSectors * (SECTOR_SIZE / 4));
  miniFat.fill(FREESECT);
  for (const chain of miniChains.values()) {
    for (let j = 0; j < chain.length; j++) {
      miniFat[chain[j]] = j + 1 < chain.length ? chain[j + 1] : ENDOFCHAIN;
    }
  }

  // ── container stream ────────────────────────────────────────────────────
  const containerData = new Uint8Array(miniSectorData);
  const containerSectors = ceilDiv(containerData.length || 1, SECTOR_SIZE);

  // ── directory ───────────────────────────────────────────────────────────
  const numDirEntries = flat.length;
  const dirSectors    = ceilDiv(numDirEntries * DIR_ENTRY_SZ, SECTOR_SIZE);

  // ── sector layout ──────────────────────────────────────────────────────
  // Sector 0        : FAT
  // Sectors 1..D    : Directory (D sectors)
  // Sectors D+1..MF : Mini-FAT (MF sectors)
  // Sectors MF+1..C : Mini-stream container (C sectors)
  const fatSectorIdx   = 0;
  const dirStartSector = 1;
  const miniFatStartSector = dirStartSector + dirSectors;
  const containerStartSector = miniFatStartSector + miniFatSectors;
  const totalSectors = 1 + dirSectors + miniFatSectors + containerSectors;

  // ── build FAT ───────────────────────────────────────────────────────────
  const fat = new Uint32Array(SECTOR_SIZE / 4);
  fat.fill(FREESECT);
  fat[fatSectorIdx] = FATSECT;
  // directory chain
  for (let i = 0; i < dirSectors; i++) {
    fat[dirStartSector + i] = i + 1 < dirSectors ? dirStartSector + i + 1 : ENDOFCHAIN;
  }
  // mini-FAT chain
  for (let i = 0; i < miniFatSectors; i++) {
    fat[miniFatStartSector + i] = i + 1 < miniFatSectors ? miniFatStartSector + i + 1 : ENDOFCHAIN;
  }
  // container chain
  for (let i = 0; i < containerSectors; i++) {
    fat[containerStartSector + i] = i + 1 < containerSectors ? containerStartSector + i + 1 : ENDOFCHAIN;
  }

  // ── root entry adjustments ──────────────────────────────────────────────
  flat[0].miniStart = containerStartSector;
  flat[0].size      = containerData.length;

  // ── build directory sectors ─────────────────────────────────────────────
  const dirBuf = new Uint8Array(dirSectors * SECTOR_SIZE);
  for (let i = 0; i < flat.length; i++) {
    const off = i * DIR_ENTRY_SZ;
    const e   = flat[i];
    const { bytes, size: nameSize } = encodeName(e.node.name);
    dirBuf.set(bytes, off);
    setU16(dirBuf, off + 64, nameSize);
    dirBuf[off + 66] = e.node.type;
    dirBuf[off + 67] = 1; // Black
    setU32(dirBuf, off + 68, e.leftIdx);
    setU32(dirBuf, off + 72, e.rightIdx);
    setU32(dirBuf, off + 76, e.childIdx);
    // CLSID, state, times = 0 (already zeroed)
    if (e.node.type === 5) {
      // root: start = first container sector, size = container data length
      setU32(dirBuf, off + 116, containerStartSector);
      setU32(dirBuf, off + 120, containerData.length);
    } else if (e.node.type === 2 && e.node.data && e.node.data.length > 0) {
      setU32(dirBuf, off + 116, e.miniStart);
      setU32(dirBuf, off + 120, e.node.data.length);
    } else if (e.node.type === 1) {
      // Storage entry: startSector=0, size=0
      setU32(dirBuf, off + 116, 0);
      setU32(dirBuf, off + 120, 0);
    } else {
      setU32(dirBuf, off + 116, ENDOFCHAIN);
      setU32(dirBuf, off + 120, 0);
    }
  }

  // ── build header ────────────────────────────────────────────────────────
  const header = new Uint8Array(512);
  header.set(SIGNATURE, 0);
  // minor version
  setU16(header, 24, 0x003E);
  // major version (3)
  setU16(header, 26, 0x0003);
  // byte order (little-endian)
  setU16(header, 28, 0xFFFE);
  // sector size power (9 → 512)
  setU16(header, 30, 9);
  // mini-sector size power (6 → 64)
  setU16(header, 32, 6);
  // total FAT sectors
  setU32(header, 44, 1);
  // first directory sector
  setU32(header, 48, dirStartSector);
  // mini-stream cutoff
  setU32(header, 56, MINI_CUTOFF);
  // first mini-FAT sector
  setU32(header, 60, miniFatStartSector);
  // total mini-FAT sectors
  setU32(header, 64, miniFatSectors);
  // first DIFAT sector (none needed)
  setU32(header, 68, ENDOFCHAIN);
  // total DIFAT sectors
  setU32(header, 72, 0);
  // DIFAT array: first entry = FAT sector 0
  setU32(header, 76, fatSectorIdx);
  for (let i = 1; i < 109; i++) setU32(header, 76 + i * 4, FREESECT);

  // ── assemble file ───────────────────────────────────────────────────────
  const fileSize = 512 + totalSectors * SECTOR_SIZE;
  const file = new Uint8Array(fileSize);
  file.set(header, 0);

  // FAT sector
  const fatBuf = new Uint8Array(SECTOR_SIZE);
  for (let i = 0; i < fat.length; i++) setU32(fatBuf, i * 4, fat[i]);
  file.set(fatBuf, 512 + fatSectorIdx * SECTOR_SIZE);

  // directory sectors
  file.set(dirBuf, 512 + dirStartSector * SECTOR_SIZE);

  // mini-FAT sectors
  const miniFatBuf = new Uint8Array(miniFatSectors * SECTOR_SIZE);
  for (let i = 0; i < miniFat.length; i++) setU32(miniFatBuf, i * 4, miniFat[i]);
  file.set(miniFatBuf, 512 + miniFatStartSector * SECTOR_SIZE);

  // container sectors
  const containerBuf = new Uint8Array(containerSectors * SECTOR_SIZE);
  containerBuf.set(containerData);
  file.set(containerBuf, 512 + containerStartSector * SECTOR_SIZE);

  return file;
}
