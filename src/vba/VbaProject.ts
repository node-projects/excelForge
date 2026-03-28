/**
 * VBA Project — read, create, and serialise VBA macros inside an XLSX / XLSM.
 *
 * Public API:
 *   const vba = new VbaProject();
 *   vba.addModule({ name: 'Module1', type: 'standard', code: '...' });
 *   const bin = vba.build();               // → Uint8Array (vbaProject.bin)
 *
 *   const vba2 = VbaProject.fromBytes(bin); // parse existing
 *   for (const m of vba2.modules) console.log(m.name, m.code);
 */

import { compressOvba, decompressOvba } from './ovba.js';
import { buildCfb, readCfb, type CfbStream } from './cfb.js';

// ── types ────────────────────────────────────────────────────────────────────

export type VbaModuleType = 'document' | 'standard' | 'class';

export interface VbaModule {
  name: string;
  type: VbaModuleType;
  code: string;
}

// ── helpers ──────────────────────────────────────────────────────────────────

const enc  = new TextEncoder();
const dec  = new TextDecoder('utf-8');

function u16le(v: number): number[] { return [v & 0xFF, (v >> 8) & 0xFF]; }
function u32le(v: number): number[] { return [v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF, (v >> 24) & 0xFF]; }

function readU16(buf: Uint8Array, off: number): number { return buf[off] | (buf[off + 1] << 8); }
function readU32(buf: Uint8Array, off: number): number { return (buf[off] | (buf[off + 1] << 8) | (buf[off + 2] << 16) | (buf[off + 3] << 24)) >>> 0; }

/** Build a dir-stream record: id (2) + size (4) + data */
function rec(id: number, data: number[] | Uint8Array): number[] {
  const d = data instanceof Uint8Array ? Array.from(data) : data;
  return [...u16le(id), ...u32le(d.length), ...d];
}

function toUtf16le(s: string): number[] {
  const out: number[] = [];
  for (const ch of s) { const c = ch.charCodeAt(0); out.push(c & 0xFF, (c >> 8) & 0xFF); }
  return out;
}

// ── module source attributes ─────────────────────────────────────────────────

function moduleAttributes(m: VbaModule): string {
  const lines: string[] = [
    `Attribute VB_Name = "${m.name}"`,
  ];
  if (m.type === 'document') {
    // 00020819 = Workbook class, 00020820 = Worksheet class
    const baseClsid = m.name === 'ThisWorkbook'
      ? '0{00020819-0000-0000-C000-000000000046}'
      : '0{00020820-0000-0000-C000-000000000046}';
    lines.push(
      `Attribute VB_Base = "${baseClsid}"`,
      'Attribute VB_GlobalNameSpace = False',
      'Attribute VB_Creatable = False',
      'Attribute VB_PredeclaredId = True',
      'Attribute VB_Exposed = True',
      'Attribute VB_TemplateDerived = False',
      'Attribute VB_Customizable = True',
    );
  } else if (m.type === 'class') {
    lines.push(
      'Attribute VB_GlobalNameSpace = False',
      'Attribute VB_Creatable = False',
      'Attribute VB_PredeclaredId = False',
      'Attribute VB_Exposed = False',
    );
  }
  return lines.join('\r\n') + '\r\n';
}

// ── VbaProject class ─────────────────────────────────────────────────────────

export class VbaProject {
  modules: VbaModule[] = [];
  /** Original raw bytes — kept for lossless round-trip when unmodified. */
  private _raw?: Uint8Array;
  private _dirty = false;

  addModule(m: VbaModule): this {
    this.modules.push(m);
    this._dirty = true;
    return this;
  }

  removeModule(name: string): this {
    this.modules = this.modules.filter(m => m.name !== name);
    this._dirty = true;
    return this;
  }

  getModule(name: string): VbaModule | undefined {
    return this.modules.find(m => m.name === name);
  }

  // ── build vbaProject.bin ──────────────────────────────────────────────────

  build(): Uint8Array {
    if (this._raw && !this._dirty) return this._raw;

    // Ensure we always have a ThisWorkbook document module
    const hasThisWorkbook = this.modules.some(
      m => m.name === 'ThisWorkbook' && m.type === 'document');
    const unsorted: VbaModule[] = hasThisWorkbook
      ? this.modules
      : [{ name: 'ThisWorkbook', type: 'document', code: '' }, ...this.modules];
    // Sort: document modules first (ThisWorkbook first), then others
    const allModules = [
      ...unsorted.filter(m => m.type === 'document'),
      ...unsorted.filter(m => m.type !== 'document'),
    ];

    // ── compress each module stream ───────────────────────────────────────
    const moduleStreams: CfbStream[] = [];
    for (const m of allModules) {
      const src = moduleAttributes(m) + m.code;
      const compressed = compressOvba(enc.encode(src));
      moduleStreams.push({ name: m.name, data: compressed, storage: 'VBA' });
    }

    // ── _VBA_PROJECT stream ───────────────────────────────────────────────
    const vbaProjectStream: CfbStream = {
      name: '_VBA_PROJECT',
      data: new Uint8Array([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00]),
      storage: 'VBA',
    };

    // ── dir stream (compressed) ───────────────────────────────────────────
    const dirBytes = this._buildDirStream(allModules);
    const dirCompressed = compressOvba(dirBytes);
    const dirStream: CfbStream = { name: 'dir', data: dirCompressed, storage: 'VBA' };

    // ── PROJECT stream (plaintext) ────────────────────────────────────────
    const projectText = this._buildProjectText(allModules);
    const projectStream: CfbStream = { name: 'PROJECT', data: enc.encode(projectText) };

    // ── PROJECTwm stream ──────────────────────────────────────────────────
    const wmData = this._buildProjectWm(allModules);
    const wmStream: CfbStream = { name: 'PROJECTwm', data: new Uint8Array(wmData) };

    // ── assemble CFB ──────────────────────────────────────────────────────
    const allStreams: CfbStream[] = [
      vbaProjectStream, dirStream, ...moduleStreams,
      projectStream, wmStream,
    ];
    return buildCfb(allStreams);
  }

  private _buildDirStream(modules: VbaModule[]): Uint8Array {
    const b: number[] = [];
    const projectName = 'VBAProject';
    const nameBytes = enc.encode(projectName);

    // PROJECTSYSKIND
    b.push(...rec(0x0001, u32le(0x00000001))); // Win32
    // PROJECTLCID
    b.push(...rec(0x0002, u32le(0x0409)));
    // PROJECTLCIDINVOKE
    b.push(...rec(0x0014, u32le(0x0409)));
    // PROJECTCODEPAGE
    b.push(...rec(0x0003, u16le(0x04E4))); // Windows-1252
    // PROJECTNAME
    b.push(...rec(0x0004, Array.from(nameBytes)));
    // PROJECTDOCSTRING + Unicode
    b.push(...rec(0x0005, []));
    b.push(...rec(0x0040, []));
    // PROJECTHELPFILEPATH + PROJECTHELPFILEPATH2
    b.push(...rec(0x0006, []));
    b.push(...rec(0x003D, []));
    // PROJECTHELPCONTEXT
    b.push(...rec(0x0007, u32le(0)));
    // PROJECTLIBFLAGS
    b.push(...rec(0x0008, u32le(0)));
    // PROJECTVERSION: special record — Reserved(4)=0x04, MajorVersion(4), MinorVersion(2)
    b.push(...u16le(0x0009), ...u32le(0x00000004), ...u32le(0x05A3), ...u16le(0x0002));
    // PROJECTCONSTANTS + Unicode
    b.push(...rec(0x000C, []));
    b.push(...rec(0x003C, []));

    // PROJECTMODULES
    b.push(...rec(0x000F, u16le(modules.length)));
    // PROJECTCOOKIE
    b.push(...rec(0x0013, u16le(0xFFFF)));

    // ── per-module records ────────────────────────────────────────────────
    for (const m of modules) {
      const nameB = enc.encode(m.name);
      const nameU = toUtf16le(m.name);

      b.push(...rec(0x0019, Array.from(nameB)));          // MODULENAME
      b.push(...rec(0x0047, nameU));                       // MODULENAMEUNICODE
      b.push(...rec(0x001A, Array.from(nameB)));          // MODULESTREAMNAME
      b.push(...rec(0x0032, nameU));                       // MODULESTREAMNAMEUNICODE
      b.push(...rec(0x001C, []));                          // MODULEDOCSTRING
      b.push(...rec(0x0048, []));                          // MODULEDOCSTRINGUNICODE
      b.push(...rec(0x0031, u32le(0)));                   // MODULEOFFSET (0 = no p-code)
      b.push(...rec(0x001E, u32le(0)));                   // MODULEHELPCONTEXT
      b.push(...rec(0x002C, u16le(0xFFFF)));              // MODULECOOKIE
      // MODULETYPE
      if (m.type === 'document') {
        b.push(...rec(0x0022, []));                        // MODULETYPEDOCUMENT
      } else {
        b.push(...rec(0x0021, []));                        // MODULETYPEPROCEDURAL
      }
      b.push(...rec(0x002B, []));                          // MODULEEOF
    }

    // Terminator
    b.push(...rec(0x0010, []));

    return new Uint8Array(b);
  }

  private _buildProjectText(modules: VbaModule[]): string {
    const lines: string[] = [
      'ID="{00000000-0000-0000-0000-000000000000}"',
    ];
    // Document modules first, then others (matches EPPlus ordering)
    const docs = modules.filter(m => m.type === 'document');
    const others = modules.filter(m => m.type !== 'document');
    for (const m of docs) {
      lines.push(`Document=${m.name}/&H00000000`);
    }
    for (const m of others) {
      if (m.type === 'class') {
        lines.push(`Class=${m.name}`);
      } else {
        lines.push(`Module=${m.name}`);
      }
    }
    lines.push(
      'Name="VBAProject"',
      'HelpContextID=0',
      'VersionCompatible32="393222000"',
      '',
      '[Host Extender Info]',
      '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000',
      '',
      '[Workspace]',
    );
    // Workspace entries: documents first, then others, with trailing space
    for (const m of [...docs, ...others]) {
      lines.push(`${m.name}=0, 0, 0, 0, C `);
    }
    return lines.join('\r\n') + '\r\n';
  }

  private _buildProjectWm(modules: VbaModule[]): number[] {
    const out: number[] = [];
    for (const m of modules) {
      // ASCII null-terminated
      for (const ch of m.name) out.push(ch.charCodeAt(0));
      out.push(0);
      // UTF-16LE null-terminated
      for (const ch of m.name) { const c = ch.charCodeAt(0); out.push(c & 0xFF, (c >> 8) & 0xFF); }
      out.push(0, 0);
    }
    out.push(0); // terminator
    return out;
  }

  // ── parse existing vbaProject.bin ─────────────────────────────────────────

  static fromBytes(data: Uint8Array): VbaProject {
    const vba = new VbaProject();
    vba._raw = data;

    const streams = readCfb(data);
    const streamMap = new Map<string, Uint8Array>();
    for (const s of streams) {
      const key = s.storage ? `${s.storage}/${s.name}` : s.name;
      streamMap.set(key, s.data);
    }

    // parse dir stream
    const dirCompressed = streamMap.get('VBA/dir');
    if (!dirCompressed) return vba;

    let dirData: Uint8Array;
    try { dirData = decompressOvba(dirCompressed); } catch { return vba; }

    // parse PROJECT stream to detect class modules (dir stream doesn't distinguish)
    const classNames = new Set<string>();
    const projectData = streamMap.get('PROJECT');
    if (projectData) {
      const projectText = dec.decode(projectData);
      for (const line of projectText.split(/\r?\n/)) {
        const m = line.match(/^Class=(.+)$/);
        if (m) classNames.add(m[1]);
      }
    }

    // parse module info from dir
    const moduleInfos = parseDirStream(dirData);

    for (const info of moduleInfos) {
      const streamKey = `VBA/${info.name}`;
      const streamData = streamMap.get(streamKey);
      if (!streamData) continue;

      let code = '';
      try {
        const decompressed = decompressOvba(streamData);
        // skip p-code bytes up to offset
        const sourceBytes = decompressed.subarray(info.offset);
        code = dec.decode(sourceBytes);
      } catch {
        // if decompression fails, try raw
        code = dec.decode(streamData.subarray(info.offset));
      }

      // Strip attribute lines — expose only the user's code
      const stripped = stripAttributes(code);

      const type = classNames.has(info.name) ? 'class' : info.type;
      vba.modules.push({
        name: info.name,
        type,
        code: stripped,
      });
    }

    vba._dirty = false;
    return vba;
  }
}

// ── dir stream parser ────────────────────────────────────────────────────────

interface ModuleInfo {
  name:   string;
  type:   VbaModuleType;
  offset: number;
}

function parseDirStream(data: Uint8Array): ModuleInfo[] {
  const modules: ModuleInfo[] = [];
  let pos = 0;
  let currentModule: Partial<ModuleInfo> | null = null;

  while (pos + 6 <= data.length) {
    const id   = readU16(data, pos);
    const size = readU32(data, pos + 2);
    pos += 6;

    // PROJECTVERSION (0x0009) is special: Reserved=4, then MajorVersion(4)+MinorVersion(2)
    if (id === 0x0009) { pos += 6; continue; }

    if (pos + size > data.length) break; // malformed record
    const body = data.subarray(pos, pos + size);
    pos += size;

    switch (id) {
      case 0x0019: // MODULENAME
        currentModule = {
          name: dec.decode(body),
          type: 'standard',
          offset: 0,
        };
        break;
      case 0x0031: // MODULEOFFSET
        if (currentModule && body.length >= 4) {
          currentModule.offset = readU32(body, 0);
        }
        break;
      case 0x0021: // MODULETYPEPROCEDURAL
        if (currentModule) currentModule.type = 'standard';
        break;
      case 0x0022: // MODULETYPEDOCUMENT
        if (currentModule) currentModule.type = 'document';
        break;
      case 0x002B: // MODULEEOF
        if (currentModule?.name) {
          modules.push(currentModule as ModuleInfo);
        }
        currentModule = null;
        break;
    }
  }
  return modules;
}

// ── strip VB_Attribute lines from source ─────────────────────────────────────

function stripAttributes(code: string): string {
  const lines = code.split(/\r?\n/);
  const start = lines.findIndex(l => !l.startsWith('Attribute '));
  if (start < 0) return code;
  const trimmed = lines.slice(start).join('\n');
  return trimmed;
}
