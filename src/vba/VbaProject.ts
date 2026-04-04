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

export type VbaModuleType = 'document' | 'standard' | 'class' | 'userform';

export interface VbaModule {
  name: string;
  type: VbaModuleType;
  code: string;
  /** Designer binary data for UserForms (f stream). Auto-generated if omitted for 'userform' type. */
  designerData?: Uint8Array;
  /** For UserForms: list of controls to place on the form (simplified API) */
  controls?: VbaFormControl[];
  /** Preserved form streams for lossless round-trip (f, o, CompObj, VBFrame) */
  _formStreams?: Map<string, Uint8Array>;
}

/** Simplified control definition for VBA UserForms */
export interface VbaFormControl {
  type: 'CommandButton' | 'TextBox' | 'Label' | 'ComboBox' | 'ListBox' | 'CheckBox' | 'OptionButton' | 'Frame' | 'Image' | 'SpinButton' | 'ScrollBar' | 'ToggleButton';
  name: string;
  caption?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
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
  } else if (m.type === 'userform') {
    lines.push(
      `Attribute VB_Base = "0{F8A47041-B2A6-11CE-8027-00AA00611080}"`,
      'Attribute VB_GlobalNameSpace = False',
      'Attribute VB_Creatable = False',
      'Attribute VB_PredeclaredId = True',
      'Attribute VB_Exposed = False',
      'Attribute VB_TemplateDerived = False',
      'Attribute VB_Customizable = False',
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

    // ── UserForm designer streams ─────────────────────────────────────────
    for (const m of allModules) {
      if (m.type === 'userform') {
        if (m._formStreams && m._formStreams.size > 0) {
          // Round-trip: use preserved streams verbatim
          for (const [streamName, data] of m._formStreams) {
            allStreams.push({ name: streamName, data, storage: m.name });
          }
        } else {
          // Fresh creation: generate proper designer streams
          allStreams.push({ name: '\x01CompObj', data: buildCompObj(), storage: m.name });
          allStreams.push({ name: '\x03VBFrame', data: enc.encode(buildVBFrameText(m)), storage: m.name });
          allStreams.push({ name: 'f', data: buildFormControlBinary(m), storage: m.name });
          allStreams.push({ name: 'o', data: buildOleObjectBlob(m), storage: m.name });
        }
      }
    }

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
      if (m.type === 'document' || m.type === 'userform') {
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
      } else if (m.type === 'userform') {
        // Package= line is required before BaseClass for UserForms
        lines.push(`Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F}`);
        lines.push(`BaseClass=${m.name}`);
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

    // parse PROJECT stream to detect class modules and userforms (dir stream doesn't distinguish)
    const classNames = new Set<string>();
    const userFormNames = new Set<string>();
    const projectData = streamMap.get('PROJECT');
    if (projectData) {
      const projectText = dec.decode(projectData);
      for (const line of projectText.split(/\r?\n/)) {
        const m = line.match(/^Class=(.+)$/);
        if (m) classNames.add(m[1]);
        const uf = line.match(/^BaseClass=(.+)$/);
        if (uf) userFormNames.add(uf[1]);
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
        // Compressed source starts at offset (bytes before are p-code)
        const compressedSource = streamData.subarray(info.offset);
        const decompressed = decompressOvba(compressedSource);
        code = dec.decode(decompressed);
      } catch {
        // if decompression fails, try raw from offset
        code = dec.decode(streamData.subarray(info.offset));
      }

      // Strip attribute lines — expose only the user's code
      const stripped = stripAttributes(code);

      const type = userFormNames.has(info.name) ? 'userform' : classNames.has(info.name) ? 'class' : info.type;
      const mod: VbaModule = {
        name: info.name,
        type,
        code: stripped,
      };
      // Preserve all designer streams for userforms
      if (type === 'userform') {
        const formStreams = new Map<string, Uint8Array>();
        for (const [key, data] of streamMap) {
          if (key.startsWith(info.name + '/')) {
            const streamName = key.substring(info.name.length + 1);
            formStreams.set(streamName, data);
          }
        }
        if (formStreams.size > 0) mod._formStreams = formStreams;
        // Also keep designerData for backward compat
        const designerKey = `${info.name}/f`;
        const designerData = streamMap.get(designerKey);
        if (designerData) mod.designerData = designerData;
      }
      vba.modules.push(mod);
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

// ── UserForm designer helpers ────────────────────────────────────────────────

/** MSForms UserForm CLSID {C62A69F0-16DC-11CE-9E98-00AA00574A4F} */
const USERFORM_CLSID_STR = '{C62A69F0-16DC-11CE-9E98-00AA00574A4F}';

/**
 * Predefined ClsidCacheIndex values for well-known control types.
 * Values verified against real Excel reference file (allelements.xlsm).
 */
const CLSID_CACHE_INDEX: Record<string, number> = {
  Image: 12, Frame: 14, SpinButton: 16, CommandButton: 17,
  Label: 21, TextBox: 23, ListBox: 24, ComboBox: 25,
  CheckBox: 26, OptionButton: 27, ToggleButton: 28, ScrollBar: 47,
};

/** Build the \x01CompObj stream for a UserForm storage */
function buildCompObj(): Uint8Array {
  // Standard CompObj header for "Microsoft Forms 2.0 Form"
  return new Uint8Array([
    0x01, 0x00, 0xFE, 0xFF, 0x03, 0x0A, 0x00, 0x00,
    0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x19, 0x00, 0x00, 0x00,
    // "Microsoft Forms 2.0 Form\0"
    0x4D, 0x69, 0x63, 0x72, 0x6F, 0x73, 0x6F, 0x66,
    0x74, 0x20, 0x46, 0x6F, 0x72, 0x6D, 0x73, 0x20,
    0x32, 0x2E, 0x30, 0x20, 0x46, 0x6F, 0x72, 0x6D,
    0x00, 0x10, 0x00, 0x00, 0x00,
    // "Embedded Object\0"
    0x45, 0x6D, 0x62,
    0x65, 0x64, 0x64, 0x65, 0x64, 0x20, 0x4F, 0x62,
    0x6A, 0x65, 0x63, 0x74, 0x00, 0x00, 0x00, 0x00,
    0x00, 0xF4, 0x39, 0xB2, 0x71, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00,
  ]);
}

/** Build the \x03VBFrame text stream for a UserForm */
function buildVBFrameText(m: VbaModule): string {
  const lines = [
    'VERSION 5.00',
    // Trailing space after name is required per Excel convention
    `Begin ${USERFORM_CLSID_STR} ${m.name} `,
    `   Caption         =   "${m.name}"`,
    '   ClientHeight    =   3015',
    '   ClientLeft      =   120',
    '   ClientTop       =   465',
    '   ClientWidth     =   4560',
    '   StartUpPosition =   1  \'CenterOwner',
    '   TypeInfoVer     =   2',
    'End',
  ];
  return lines.join('\r\n') + '\r\n';
}

/**
 * Build a valid MS-OFORMS FormControl binary for the "f" stream.
 * Follows [MS-OFORMS] 2.2.10 — FormControl + FormStreamData.
 *
 * Structure:
 *   FormControl header:  MinorVersion(1) + MajorVersion(1) + cbForm(2)
 *   FormControl body:    PropMask(4) + DataBlock + ExtraDataBlock  (cbForm bytes)
 *   FormStreamData:      ClassTable + FormObjectDepthTypeCount + SiteData
 *
 * FormPropMask bits (from [MS-OFORMS] 2.2.10.2):
 *   bit 0: Unused1, bit 1: fBackColor, bit 2: fForeColor, bit 3: fNextAvailableID,
 *   bit 4-5: Unused2, bit 6: fBooleanProperties, bit 7: Unused3,
 *   bit 8: fBorderStyle, bit 9: fMousePointer, bit 10: fScrollBars,
 *   bit 10: fDisplayedSize(ExtraData), bit 11: fLogicalSize(ExtraData),
 *   bit 12: fScrollPosition(ExtraData), bit 13: fGroupCnt, ...
 *   bit 26: fShapeCookie, bit 27: fDrawBuffer
 */
function buildFormControlBinary(m: VbaModule): Uint8Array {
  const controls = m.controls ?? [];
  const numControls = controls.length;

  // Pre-build the 'o' stream for each control so we know fObjectStreamSize per control
  const oBlobs = controls.map(ctrl => buildSingleControlOBlob(ctrl));

  const b: number[] = [];

  // ══ FormControl header ══════════════════════════════════════════════════
  b.push(0x00);           // MinorVersion
  b.push(0x04);           // MajorVersion
  const cbPos = b.length;
  b.push(0x00, 0x00);     // cbForm placeholder (= PropMask + DataBlock + ExtraDataBlock)

  // PropMask (4 bytes) — [MS-OFORMS] 2.2.10.2 FormPropMask
  // bit 3: fNextAvailableID, bit 10: fDisplayedSize(ExtraData),
  // bit 11: fLogicalSize(ExtraData), bit 26: fShapeCookie, bit 27: fDrawBuffer
  const PROP_fNextAvailableID = 1 << 3;
  const PROP_fDisplayedSize   = 1 << 10;
  const PROP_fLogicalSize     = 1 << 11;
  const PROP_fShapeCookie     = 1 << 26;
  const PROP_fDrawBuffer      = 1 << 27;
  const propMask = PROP_fNextAvailableID | PROP_fDisplayedSize | PROP_fLogicalSize
                 | PROP_fShapeCookie | PROP_fDrawBuffer;
  b.push(propMask & 0xFF, (propMask >> 8) & 0xFF, (propMask >> 16) & 0xFF, (propMask >> 24) & 0xFF);

  // DataBlock (properties in bit order, only for non-ExtraData bits):
  // bit 3: fNextAvailableID (DWORD) — next control ID to assign
  b.push(...u32le(numControls + 1));

  // bit 26: fShapeCookie (DWORD)
  b.push(...u32le(0));

  // bit 27: fDrawBuffer (DWORD) — standard value 32000
  b.push(...u32le(32000));

  // ExtraDataBlock (properties in bit order, only for ExtraData bits):
  // bit 10: fDisplayedSize — Width(4) + Height(4) in HIMETRIC
  // Compute form dimensions from controls bounding box + margin
  let maxR = 200, maxB = 120; // minimum form size in points
  for (const c of controls) {
    const r = (c.left ?? 0) + (c.width ?? 72);
    const bt = (c.top ?? 0) + (c.height ?? 24);
    if (r > maxR) maxR = r;
    if (bt > maxB) maxB = bt;
  }
  const formWidthHimetric = (maxR + 20) * 26;   // add margin, convert to HIMETRIC
  const formHeightHimetric = (maxB + 20) * 26;
  b.push(...u32le(formWidthHimetric));
  b.push(...u32le(formHeightHimetric));

  // bit 11: fLogicalSize — Width(4) + Height(4) in HIMETRIC (0,0 = same as displayed)
  b.push(...u32le(0));
  b.push(...u32le(0));

  // Back-patch cbForm (everything from PropMask to end of ExtraDataBlock)
  const cbForm = b.length - 4; // offset 4 = start of body (after header)
  b[cbPos] = cbForm & 0xFF;
  b[cbPos + 1] = (cbForm >> 8) & 0xFF;

  // ══ FormStreamData [MS-OFORMS] 2.2.10.5 ════════════════════════════════
  // No fMouseIcon, fFont, fPicture → skip GuidAndPicture/GuidAndFont

  // ClassTable: Since we did NOT set fBooleanProperties (bit 6),
  // FORM_FLAG_DONTSAVECLASSTABLE defaults to 0, so ClassTable IS written.
  // CountOfSiteClassInfo = 0 (we use predefined ClsidCacheIndex, no custom CLSIDs)
  b.push(...u16le(0));  // CountOfSiteClassInfo (USHORT)

  // ══ FormObjectDepthTypeCount [MS-OFORMS] 2.2.10.7 ══════════════════════
  b.push(...u32le(numControls));   // CountOfSites

  // CountOfBytes placeholder (DepthTypeCombo + SiteData)
  const cbBytesPos = b.length;
  b.push(0x00, 0x00, 0x00, 0x00);

  const countedStart = b.length;

  // DepthTypeCombo: [MS-OFORMS] 2.2.10.7 — per-site encoding
  // Simple format: depth(1 byte) + 0x01(1 byte) = 1 site at given depth
  // All controls are top-level (depth=0)
  for (let i = 0; i < numControls; i++) {
    b.push(0x00);  // depth = 0
    b.push(0x01);  // type=1 (individual site)
  }
  // Pad DepthTypeCombo to 4-byte boundary from countedStart
  while ((b.length - countedStart) % 4 !== 0) b.push(0x00);

  // ══ SiteData (OleSiteConcrete per control) [MS-OFORMS] 2.2.10.12 ══════
  for (let i = 0; i < numControls; i++) {
    const ctrl = controls[i];
    const nameBytes = enc.encode(ctrl.name);
    const cacheIdx = CLSID_CACHE_INDEX[ctrl.type] ?? CLSID_CACHE_INDEX['CommandButton'];

    // OleSiteConcrete header
    b.push(0x00, 0x00);          // Version = 0x0000

    const cbSitePos = b.length;
    b.push(0x00, 0x00);          // cbSite placeholder

    // SitePropMask (4 bytes) [MS-OFORMS] 2.2.10.12.2:
    // bit 0: fName, bit 2: fID, bit 5: fObjectStreamSize,
    // bit 6: fTabIndex, bit 7: fClsidCacheIndex, bit 8: fPosition
    const siteMask = (1 << 0) | (1 << 2) | (1 << 5) | (1 << 6) | (1 << 7) | (1 << 8);
    b.push(siteMask & 0xFF, (siteMask >> 8) & 0xFF, (siteMask >> 16) & 0xFF, (siteMask >> 24) & 0xFF);

    // SiteDataBlock (properties in bit order, with padded alignment):
    // bit 0: fName — CountOfBytesWithCompressionFlag (4 bytes, bit 31 = ANSI)
    b.push(...u32le(nameBytes.length | 0x80000000));

    // bit 2: fID (DWORD) — control ID (1-based)
    b.push(...u32le(i + 1));

    // bit 5: fObjectStreamSize (DWORD) — size of this control's data in 'o' stream
    b.push(...u32le(oBlobs[i].length));

    // bit 6: fTabIndex (SHORT) — tab stop order
    b.push(i & 0xFF, (i >> 8) & 0xFF);

    // bit 7: fClsidCacheIndex (SHORT) — predefined control type index
    b.push(cacheIdx & 0xFF, (cacheIdx >> 8) & 0xFF);

    // SiteExtraDataBlock [MS-OFORMS] 2.2.10.12.4:
    // fName string data
    b.push(...nameBytes);
    while ((b.length - countedStart) % 4 !== 0) b.push(0x00);

    // bit 8: fPosition — left(4) + top(4) in HIMETRIC
    const left = (ctrl.left ?? (10 + i * 80)) * 26;
    const top = (ctrl.top ?? (10 + i * 30)) * 26;
    b.push(...u32le(left));
    b.push(...u32le(top));

    // Back-patch cbSite (SitePropMask + SiteDataBlock + SiteExtraDataBlock)
    const cbSite = b.length - cbSitePos - 2;
    b[cbSitePos] = cbSite & 0xFF;
    b[cbSitePos + 1] = (cbSite >> 8) & 0xFF;
  }

  // Back-patch CountOfBytes (DepthTypeCombo + pad + SiteData)
  const cbBytes = b.length - countedStart;
  b[cbBytesPos]     = cbBytes & 0xFF;
  b[cbBytesPos + 1] = (cbBytes >> 8) & 0xFF;
  b[cbBytesPos + 2] = (cbBytes >> 16) & 0xFF;
  b[cbBytesPos + 3] = (cbBytes >> 24) & 0xFF;

  return new Uint8Array(b);
}

/**
 * Build control blob for a single control in the 'o' stream.
 * Dispatches to the appropriate builder based on control type.
 */
function buildSingleControlOBlob(ctrl: { type: string; name: string; caption?: string; width?: number; height?: number }): Uint8Array {
  const cacheIdx = CLSID_CACHE_INDEX[ctrl.type] ?? CLSID_CACHE_INDEX['CommandButton'];
  // ClsidCacheIndex determines the binary format:
  //   17 → CommandButtonControl
  //   21 → LabelControl
  //   23-28 → MorphDataControl (TextBox, CheckBox, etc.)
  //   12 → ImageControl
  //   16 → SpinButtonControl
  //   47 → ScrollBarControl
  if (cacheIdx === 17) return buildCommandButtonOBlob(ctrl);
  if (cacheIdx === 21) return buildLabelOBlob(ctrl);
  // All MorphData indices: 23, 24, 25, 26, 27, 28
  return buildMorphDataOBlob(ctrl);
}

/** Append standard TextProps block (Tahoma 8pt) to a byte array */
function appendTextProps(b: number[]): void {
  b.push(0x00, 0x02); // Version: minor=0, major=2
  const cbTPPos = b.length;
  b.push(0x00, 0x00); // cbTextProps placeholder

  // TextPropsPropMask [MS-OFORMS] 2.3.2:
  //   bit 0: fFontName, bit 1: fFontEffects, bit 2: fFontHeight,
  //   bit 3: unused, bit 4: fFontCharSet, bit 5: fFontPitchAndFamily,
  //   bit 6: fParagraphAlign, bit 7: fFontWeight
  // We set: fFontName(0) + fFontHeight(2) = 0x05
  b.push(0x05, 0x00, 0x00, 0x00);

  // DataBlock (padded alignment):
  // bit 0: fFontName CCH (4 bytes with ANSI flag)
  const fontNameBytes = enc.encode('Tahoma');
  b.push(...u32le(fontNameBytes.length | 0x80000000));
  // bit 2: fFontHeight (4 bytes) — height in twips (1pt = 20 twips, 8pt = 160)
  b.push(0xA0, 0x00, 0x00, 0x00); // 160 = 0xA0

  // ExtraDataBlock: font name string
  b.push(...fontNameBytes);
  while (b.length % 4 !== 0) b.push(0x00);

  const cbTP = b.length - cbTPPos - 2;
  b[cbTPPos] = cbTP & 0xFF;
  b[cbTPPos + 1] = (cbTP >> 8) & 0xFF;
}

/**
 * Build CommandButtonControl binary [MS-OFORMS] 2.2.1.1
 * Minimal format matching real Excel output: only fCaption + fSize.
 */
function buildCommandButtonOBlob(ctrl: { type: string; caption?: string; width?: number; height?: number }): Uint8Array {
  const b: number[] = [];
  const hasCaption = !!ctrl.caption;

  b.push(0x00, 0x02); // Version
  const cbPos = b.length;
  b.push(0x00, 0x00); // cb placeholder

  // PropMask (4 bytes): fCaption(3) + fSize(5) = 0x28
  let mask = (1 << 5); // fSize
  if (hasCaption) mask |= (1 << 3); // fCaption
  b.push(mask & 0xFF, (mask >> 8) & 0xFF, (mask >> 16) & 0xFF, (mask >> 24) & 0xFF);

  // DataBlock (padded alignment):
  // bit 3: fCaption (CCH with ANSI flag)
  if (hasCaption) {
    const capBytes = enc.encode(ctrl.caption!);
    b.push(...u32le(capBytes.length | 0x80000000));
  }

  // ExtraDataBlock:
  // fCaption string
  if (hasCaption) {
    const capBytes = enc.encode(ctrl.caption!);
    b.push(...capBytes);
    while (b.length % 4 !== 0) b.push(0x00);
  }
  // bit 5: fSize — width(4) + height(4) in HIMETRIC
  const width = (ctrl.width ?? 72) * 26;
  const height = (ctrl.height ?? 24) * 26;
  b.push(...u32le(width));
  b.push(...u32le(height));

  // Back-patch cb
  const cb = b.length - cbPos - 2;
  b[cbPos] = cb & 0xFF;
  b[cbPos + 1] = (cb >> 8) & 0xFF;

  // StreamData: TextProps (no Picture, no MouseIcon)
  appendTextProps(b);

  return new Uint8Array(b);
}

/**
 * Build LabelControl binary [MS-OFORMS] 2.2.4.1
 * Minimal format matching real Excel output: only fCaption + fSize.
 */
function buildLabelOBlob(ctrl: { type: string; caption?: string; width?: number; height?: number }): Uint8Array {
  const b: number[] = [];
  const hasCaption = !!ctrl.caption;

  b.push(0x00, 0x02); // Version
  const cbPos = b.length;
  b.push(0x00, 0x00); // cb placeholder

  // PropMask: fCaption(3) + fSize(5) = 0x28
  let mask = (1 << 5); // fSize
  if (hasCaption) mask |= (1 << 3); // fCaption
  b.push(mask & 0xFF, (mask >> 8) & 0xFF, (mask >> 16) & 0xFF, (mask >> 24) & 0xFF);

  // DataBlock (padded alignment):
  // bit 3: fCaption (CCH with ANSI flag)
  if (hasCaption) {
    const capBytes = enc.encode(ctrl.caption!);
    b.push(...u32le(capBytes.length | 0x80000000));
  }

  // ExtraDataBlock:
  // fCaption string
  if (hasCaption) {
    const capBytes = enc.encode(ctrl.caption!);
    b.push(...capBytes);
    while (b.length % 4 !== 0) b.push(0x00);
  }
  // bit 5: fSize — width(4) + height(4)
  const width = (ctrl.width ?? 80) * 26;
  const height = (ctrl.height ?? 18) * 26;
  b.push(...u32le(width));
  b.push(...u32le(height));

  // Back-patch cb
  const cb = b.length - cbPos - 2;
  b[cbPos] = cb & 0xFF;
  b[cbPos + 1] = (cb >> 8) & 0xFF;

  // StreamData: TextProps (no Picture, no MouseIcon)
  appendTextProps(b);

  return new Uint8Array(b);
}

/**
 * Build MorphDataControl binary [MS-OFORMS] 2.2.5.1
 * Used for TextBox, CheckBox, OptionButton, ComboBox, ListBox, ToggleButton.
 * Minimal format matching real Excel: fVariousPropertyBits + fSize, plus fCaption when needed.
 * MorphDataPropMask is 8 bytes (part1 + part2).
 */
function buildMorphDataOBlob(ctrl: { type: string; caption?: string; width?: number; height?: number }): Uint8Array {
  const b: number[] = [];
  const hasCaption = !!ctrl.caption;

  b.push(0x00, 0x02); // Version
  const cbPos = b.length;
  b.push(0x00, 0x00); // cb placeholder

  // MorphDataPropMask (8 bytes)
  // bit 0: fVariousPropertyBits, bit 8: fSize(ExtraData), bit 23: fCaption
  let part1 = (1 << 0) | (1 << 8); // fVariousPropertyBits + fSize
  if (hasCaption) part1 |= (1 << 23); // fCaption
  // Set bit 31 (Reserved) to match Excel's convention
  part1 |= (1 << 31);
  b.push(part1 & 0xFF, (part1 >> 8) & 0xFF, (part1 >> 16) & 0xFF, (part1 >>> 24) & 0xFF);
  b.push(0x00, 0x00, 0x00, 0x00); // Part2 = 0

  // DataBlock (padded alignment):
  // bit 0: fVariousPropertyBits (4 bytes) — default flags
  b.push(0x1B, 0x48, 0x80, 0x2C);

  // bit 23: fCaption (CCH with ANSI flag)
  if (hasCaption) {
    const capBytes = enc.encode(ctrl.caption!);
    b.push(...u32le(capBytes.length | 0x80000000));
  }

  // ExtraDataBlock:
  // bit 8: fSize — width(4) + height(4) in HIMETRIC
  const width = (ctrl.width ?? 72) * 26;
  const height = (ctrl.height ?? 24) * 26;
  b.push(...u32le(width));
  b.push(...u32le(height));

  // Caption string data (if hasCaption)
  if (hasCaption) {
    const capBytes = enc.encode(ctrl.caption!);
    b.push(...capBytes);
    while (b.length % 4 !== 0) b.push(0x00);
  }

  // Back-patch cbMorphData
  const cb = b.length - cbPos - 2;
  b[cbPos] = cb & 0xFF;
  b[cbPos + 1] = (cb >> 8) & 0xFF;

  // StreamData: TextProps (no MouseIcon, no Picture)
  appendTextProps(b);

  return new Uint8Array(b);
}

/**
 * Build the "o" stream (OLE Object Blob) containing per-control embedded data.
 * Each control gets a type-specific control blob.
 */
function buildOleObjectBlob(m: VbaModule): Uint8Array {
  const controls = m.controls ?? [];
  if (controls.length === 0) return new Uint8Array(0);

  const parts: Uint8Array[] = [];
  for (const ctrl of controls) {
    parts.push(buildSingleControlOBlob(ctrl));
  }

  // Concatenate all control blobs
  const total = parts.reduce((s, p) => s + p.length, 0);
  const result = new Uint8Array(total);
  let off = 0;
  for (const p of parts) { result.set(p, off); off += p.length; }
  return result;
}
