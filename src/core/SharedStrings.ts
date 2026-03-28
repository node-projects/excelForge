import { escapeXml } from '../utils/helpers.js';
import type { RichTextRun } from '../core/types.js';

export class SharedStrings {
  private table: Map<string, number> = new Map();
  private strings: string[] = [];
  private _count = 0;

  get count(): number { return this._count; }
  get uniqueCount(): number { return this.strings.length; }

  intern(s: string): number {
    this._count++;
    const existing = this.table.get(s);
    if (existing !== undefined) return existing;
    const idx = this.strings.length;
    this.strings.push(s);
    this.table.set(s, idx);
    return idx;
  }

  internRichText(runs: RichTextRun[]): number {
    const key = JSON.stringify(runs);
    this._count++;
    const existing = this.table.get(key);
    if (existing !== undefined) return existing;

    const xml = runs.map(r => {
      const rPr = r.font ? richFontXml(r.font) : '';
      return `<r>${rPr}<t xml:space="preserve">${escapeXml(r.text)}</t></r>`;
    }).join('');

    const idx = this.strings.length;
    // store the raw XML for rich text (prefixed so we can detect it)
    this.strings.push('\x00RICH\x00' + xml);
    this.table.set(key, idx);
    return idx;
  }

  toXml(): string {
    const items = this.strings.map(s => {
      if (s.startsWith('\x00RICH\x00')) {
        return `<si>${s.slice(6)}</si>`;
      }
      // Preserve leading/trailing spaces with xml:space
      const needsSpace = s !== s.trim() || s.includes('\n');
      return `<si><t${needsSpace ? ' xml:space="preserve"' : ''}>${escapeXml(s)}</t></si>`;
    }).join('');

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${this._count}" uniqueCount="${this.strings.length}">
${items}
</sst>`;
  }
}

function richFontXml(f: import('../core/types.js').Font): string {
  const parts: string[] = [];
  if (f.bold)   parts.push('<b/>');
  if (f.italic) parts.push('<i/>');
  if (f.strike) parts.push('<strike/>');
  if (f.underline && f.underline !== 'none') parts.push(`<u val="${f.underline}"/>`);
  if (f.size)   parts.push(`<sz val="${f.size}"/>`);
  if (f.color) {
    if (f.color.startsWith('theme:')) parts.push(`<color theme="${f.color.slice(6)}"/>`);
    else parts.push(`<color rgb="${f.color.startsWith('#') ? 'FF'+f.color.slice(1) : f.color}"/>`);
  }
  if (f.name)   parts.push(`<name val="${escapeXml(f.name)}"/>`);
  return parts.length ? `<rPr>${parts.join('')}</rPr>` : '';
}
