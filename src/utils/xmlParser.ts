/**
 * Lightweight XML parser.
 *
 * Produces a simple node tree that:
 *  - Preserves ALL original content (including unknown elements and attributes)
 *  - Supports fast attribute and child lookups
 *  - Can be serialised back to XML with minimal changes (roundtrip-safe)
 */

export interface XmlNode {
  tag:        string;
  attrs:      Record<string, string>;
  children:   XmlNode[];
  text?:      string;
  /** The original raw XML string for this subtree (set on parse, cleared on mutation) */
  _raw?:      string;
  /** Namespace prefix map inherited from ancestors */
  _ns?:       Record<string, string>;
}

// ─── Serialise ────────────────────────────────────────────────────────────────

export function nodeToXml(node: XmlNode): string {
  const attrs = Object.entries(node.attrs)
    .map(([k, v]) => `${k}="${escAttr(v)}"`)
    .join(' ');
  const open = attrs ? `<${node.tag} ${attrs}` : `<${node.tag}`;

  const inner = (node.text ?? '') +
    node.children.map(nodeToXml).join('');

  if (!inner) return `${open}/>`;
  return `${open}>${inner}</${node.tag}>`;
}

function escAttr(s: string): string {
  return s.replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// ─── Parse ────────────────────────────────────────────────────────────────────

/** Parse an XML string into a node tree */
export function parseXml(xml: string): XmlNode {
  const root: XmlNode[] = [];
  const stack: XmlNode[] = [];
  let i = 0;

  // Strip XML declaration
  if (xml.startsWith('<?')) {
    const end = xml.indexOf('?>');
    i = end + 2;
  }

  while (i < xml.length) {
    if (xml[i] !== '<') {
      // text node
      const end = xml.indexOf('<', i);
      const text = end < 0 ? xml.slice(i) : xml.slice(i, end);
      if (stack.length && text.trim()) {
        const top = stack[stack.length - 1];
        top.text = (top.text ?? '') + decodeEntities(text);
      }
      i = end < 0 ? xml.length : end;
      continue;
    }

    // Comment or CDATA
    if (xml.startsWith('<!--', i)) {
      const end = xml.indexOf('-->', i);
      if (stack.length) {
        const top = stack[stack.length - 1];
        top.text = (top.text ?? '') + xml.slice(i, end + 3);
      }
      i = end + 3;
      continue;
    }
    if (xml.startsWith('<![CDATA[', i)) {
      const end = xml.indexOf(']]>', i);
      const cdata = xml.slice(i + 9, end);
      if (stack.length) {
        const top = stack[stack.length - 1];
        top.text = (top.text ?? '') + cdata;
      }
      i = end + 3;
      continue;
    }

    // Closing tag
    if (xml[i + 1] === '/') {
      const end = xml.indexOf('>', i);
      const node = stack.pop()!;
      if (stack.length) {
        stack[stack.length - 1].children.push(node);
      } else {
        root.push(node);
      }
      i = end + 1;
      continue;
    }

    // Processing instruction
    if (xml[i + 1] === '?') {
      const end = xml.indexOf('?>', i);
      i = end + 2;
      continue;
    }

    // Opening tag
    const end = findTagEnd(xml, i);
    const raw = xml.slice(i + 1, end);
    const selfClose = raw.endsWith('/');
    const tagContent = selfClose ? raw.slice(0, -1).trim() : raw.trim();
    const { tag, attrs } = parseTag(tagContent);
    const node: XmlNode = { tag, attrs, children: [] };

    if (selfClose) {
      if (stack.length) {
        stack[stack.length - 1].children.push(node);
      } else {
        root.push(node);
      }
    } else {
      stack.push(node);
    }
    i = end + 1;
  }

  // Drain remaining stack (malformed XML tolerance)
  while (stack.length > 1) {
    const node = stack.pop()!;
    stack[stack.length - 1].children.push(node);
  }
  if (stack.length === 1) root.push(stack[0]);

  if (root.length === 0) throw new Error('Empty XML document');
  return root[0];
}

function findTagEnd(xml: string, start: number): number {
  let inQuote: string | null = null;
  for (let i = start + 1; i < xml.length; i++) {
    const c = xml[i];
    if (inQuote) { if (c === inQuote) inQuote = null; }
    else if (c === '"' || c === "'") inQuote = c;
    else if (c === '>') return i;
  }
  return xml.length - 1;
}

function parseTag(s: string): { tag: string; attrs: Record<string, string> } {
  const attrs: Record<string, string> = {};
  // tag name is up to first whitespace
  const spaceIdx = s.search(/\s/);
  const tag = spaceIdx < 0 ? s : s.slice(0, spaceIdx);
  if (spaceIdx < 0) return { tag, attrs };

  const rest = s.slice(spaceIdx);
  const re = /(\S+?)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(rest)) !== null) {
    attrs[m[1]] = decodeEntities(m[2] ?? m[3] ?? '');
  }
  return { tag, attrs };
}

function decodeEntities(s: string): string {
  return s
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(parseInt(n, 10)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
}

// ─── Query helpers ────────────────────────────────────────────────────────────

/** Find first child with given local tag name (ignores namespace prefix) */
export function child(node: XmlNode, localTag: string): XmlNode | undefined {
  return node.children.find(c => localName(c.tag) === localTag);
}

/** Find all children with given local tag name */
export function children(node: XmlNode, localTag: string): XmlNode[] {
  return node.children.filter(c => localName(c.tag) === localTag);
}

/** Get attribute value, checking both prefixed and unprefixed */
export function attr(node: XmlNode, name: string): string | undefined {
  if (node.attrs[name] !== undefined) return node.attrs[name];
  // Try without namespace prefix
  const local = localName(name);
  for (const [k, v] of Object.entries(node.attrs)) {
    if (localName(k) === local) return v;
  }
  return undefined;
}

/** Strip namespace prefix */
export function localName(tag: string): string {
  const colon = tag.indexOf(':');
  return colon < 0 ? tag : tag.slice(colon + 1);
}

/** Walk all descendants depth-first */
export function walk(node: XmlNode, fn: (n: XmlNode) => void): void {
  fn(node);
  for (const c of node.children) walk(c, fn);
}

/** Find first descendant matching predicate */
export function find(node: XmlNode, pred: (n: XmlNode) => boolean): XmlNode | undefined {
  if (pred(node)) return node;
  for (const c of node.children) {
    const r = find(c, pred);
    if (r) return r;
  }
  return undefined;
}

/** Find all descendants matching predicate */
export function findAll(node: XmlNode, pred: (n: XmlNode) => boolean): XmlNode[] {
  const results: XmlNode[] = [];
  walk(node, n => { if (pred(n)) results.push(n); });
  return results;
}
