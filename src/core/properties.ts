/**
 * Extended & custom workbook properties.
 *
 * Sources:
 *  - OOXML §22.2  — Core properties  (docProps/core.xml)
 *  - OOXML §15.2.12 — Extended properties (docProps/app.xml)
 *  - OOXML §22.3  — Custom properties (docProps/custom.xml)
 */

// ─── Core properties (dc / cp namespace) ─────────────────────────────────────

export interface CoreProperties {
  title?:          string;
  subject?:        string;
  creator?:        string;   // dc:creator
  keywords?:       string;
  description?:    string;
  lastModifiedBy?: string;
  revision?:       string;
  created?:        Date;
  modified?:       Date;
  category?:       string;
  contentStatus?:  string;
  language?:       string;
  identifier?:     string;
  version?:        string;
}

// ─── Extended properties (vt namespace, app.xml) ─────────────────────────────

export interface ExtendedProperties {
  /** Application that created the file, e.g. "Microsoft Excel" */
  application?:         string;
  /** Application version string, e.g. "16.0300" */
  appVersion?:          string;
  /** Company name */
  company?:             string;
  /** Manager name */
  manager?:             string;
  /** Doc security level (0 = none) */
  docSecurity?:         number;
  scaleCrop?:           boolean;
  linksUpToDate?:       boolean;
  sharedDoc?:           boolean;
  hyperlinksChanged?:   boolean;
  /** Names of the sheets in display order */
  headingPairs?:        HeadingPair[];
  /** Flat list of part titles (sheet names) */
  titlesOfParts?:       string[];
  /** Number of characters (not always set) */
  characters?:          number;
  charactersByWord?:     number;
  /** Word count */
  words?:               number;
  /** Line count */
  lines?:               number;
  /** Paragraph count */
  paragraphs?:          number;
  /** Page count */
  pages?:               number;
  /** Slide count */
  slides?:              number;
  /** Note count */
  notes?:               number;
  /** Hidden slide count */
  hiddenSlides?:        number;
  /** Multimedia clip count */
  mmClips?:             number;
  /** Template used */
  template?:            string;
  /** Presentation format */
  presentationFormat?:  string;
  /** Total editing time (100-nanosecond intervals) */
  totalTime?:           number;
  /** Digital signature */
  digitalSignature?:    boolean;
  /** Hyperlink base URL */
  hyperlinkBase?:       string;
}

export interface HeadingPair {
  name:  string;
  count: number;
}

// ─── Custom properties (docProps/custom.xml) ─────────────────────────────────

export type CustomPropValue =
  | { type: 'string';   value: string   }
  | { type: 'int';      value: number   }
  | { type: 'decimal';  value: number   }
  | { type: 'bool';     value: boolean  }
  | { type: 'date';     value: Date     }
  | { type: 'r8';       value: number   }  // 8-byte real
  | { type: 'i8';       value: bigint   }  // 8-byte int
  | { type: 'error';    value: string   };

export interface CustomProperty {
  name:  string;
  value: CustomPropValue;
}

// ─── Combined "all properties" container ─────────────────────────────────────

export interface WorkbookAllProperties {
  core?:     CoreProperties;
  extended?: ExtendedProperties;
  custom?:   CustomProperty[];
}

// ─── Serialisers ─────────────────────────────────────────────────────────────

import { escapeXml } from '../utils/helpers.js';

function isoDate(d: Date): string { return d.toISOString(); }

export function buildCoreXml(p: CoreProperties): string {
  const t = (tag: string, val: string | undefined) =>
    val !== undefined ? `<${tag}>${escapeXml(val)}</${tag}>` : '';
  const dt = (tag: string, val: Date | undefined) =>
    val ? `<${tag} xsi:type="dcterms:W3CDTF">${isoDate(val)}</${tag}>` : '';

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
${t('dc:title',           p.title)}
${t('dc:subject',         p.subject)}
${t('dc:creator',         p.creator)}
${t('cp:keywords',        p.keywords)}
${t('dc:description',     p.description)}
${t('cp:lastModifiedBy',  p.lastModifiedBy)}
${t('cp:revision',        p.revision ?? '1')}
${t('dc:language',        p.language)}
${t('dc:identifier',      p.identifier)}
${t('cp:version',         p.version)}
${t('cp:category',        p.category)}
${t('cp:contentStatus',   p.contentStatus)}
${dt('dcterms:created',   p.created  ?? new Date())}
${dt('dcterms:modified',  p.modified ?? new Date())}
</cp:coreProperties>`;
}

export function buildAppXml(p: ExtendedProperties, extraRaw?: string): string {
  const t = (tag: string, val: string | number | boolean | undefined) =>
    val !== undefined ? `<${tag}>${escapeXml(String(val))}</${tag}>` : '';

  let headingPairsXml = '';
  let titlesXml = '';

  if (p.headingPairs?.length) {
    const pairs = p.headingPairs.map(hp =>
      `<vt:variant><vt:lpstr>${escapeXml(hp.name)}</vt:lpstr></vt:variant>` +
      `<vt:variant><vt:i4>${hp.count}</vt:i4></vt:variant>`
    ).join('');
    headingPairsXml = `<HeadingPairs><vt:vector size="${p.headingPairs.length * 2}" baseType="variant">${pairs}</vt:vector></HeadingPairs>`;
  }

  if (p.titlesOfParts?.length) {
    const titles = p.titlesOfParts.map(s =>
      `<vt:lpstr>${escapeXml(s)}</vt:lpstr>`
    ).join('');
    titlesXml = `<TitlesOfParts><vt:vector size="${p.titlesOfParts.length}" baseType="lpstr">${titles}</vt:vector></TitlesOfParts>`;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
${t('Application',        p.application ?? 'ExcelForge')}
${t('AppVersion',         p.appVersion)}
${t('Company',            p.company)}
${t('Manager',            p.manager)}
${t('DocSecurity',        p.docSecurity ?? 0)}
${t('ScaleCrop',          p.scaleCrop   ?? false)}
${t('LinksUpToDate',      p.linksUpToDate ?? false)}
${t('SharedDoc',          p.sharedDoc ?? false)}
${t('HyperlinksChanged',  p.hyperlinksChanged ?? false)}
${t('Characters',         p.characters)}
${t('CharactersWithSpaces', p.charactersByWord)}
${t('Words',              p.words)}
${t('Lines',              p.lines)}
${t('Paragraphs',         p.paragraphs)}
${t('Pages',              p.pages)}
${t('Slides',             p.slides)}
${t('Notes',              p.notes)}
${t('HiddenSlides',       p.hiddenSlides)}
${t('MMClips',            p.mmClips)}
${t('Template',           p.template)}
${t('PresentationFormat', p.presentationFormat)}
${t('TotalTime',          p.totalTime)}
${t('HyperlinkBase',      p.hyperlinkBase)}
${headingPairsXml}
${titlesXml}
${extraRaw ?? ''}
</Properties>`;
}

export function buildCustomXml(props: CustomProperty[]): string {
  let pid = 2; // PIDs start at 2 in OOXML
  const items = props.map(p => {
    let valXml: string;
    const v = p.value;
    switch (v.type) {
      case 'string':  valXml = `<vt:lpwstr>${escapeXml(v.value)}</vt:lpwstr>`; break;
      case 'int':     valXml = `<vt:i4>${v.value}</vt:i4>`; break;
      case 'decimal': valXml = `<vt:decimal>${v.value}</vt:decimal>`; break;
      case 'bool':    valXml = `<vt:bool>${v.value}</vt:bool>`; break;
      case 'date':    valXml = `<vt:filetime>${v.value.toISOString()}</vt:filetime>`; break;
      case 'r8':      valXml = `<vt:r8>${v.value}</vt:r8>`; break;
      case 'i8':      valXml = `<vt:i8>${v.value}</vt:i8>`; break;
      case 'error':   valXml = `<vt:error>${escapeXml(v.value)}</vt:error>`; break;
      default:        valXml = `<vt:lpwstr></vt:lpwstr>`; break;
    }
    return `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${pid++}" name="${escapeXml(p.name)}">${valXml}</property>`;
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
${items.join('\n')}
</Properties>`;
}

// ─── Parsers ──────────────────────────────────────────────────────────────────

import { parseXml, child, children, attr, localName, type XmlNode } from '../utils/xmlParser.js';

export function parseCoreXml(xml: string): CoreProperties {
  const root = parseXml(xml);
  const g = (tag: string) => {
    const n = root.children.find(c => localName(c.tag) === tag);
    return n?.text ?? n?.children[0]?.text;
  };
  const gDate = (tag: string) => {
    const s = g(tag); return s ? new Date(s) : undefined;
  };
  return {
    title:          g('title'),
    subject:        g('subject'),
    creator:        g('creator'),
    keywords:       g('keywords'),
    description:    g('description'),
    lastModifiedBy: g('lastModifiedBy'),
    revision:       g('revision'),
    created:        gDate('created'),
    modified:       gDate('modified'),
    category:       g('category'),
    contentStatus:  g('contentStatus'),
    language:       g('language'),
    identifier:     g('identifier'),
    version:        g('version'),
  };
}

export function parseAppXml(xml: string): { props: ExtendedProperties; unknownRaw: string } {
  const root = parseXml(xml);
  const g = (tag: string) => root.children.find(c => localName(c.tag) === tag)?.text;
  const gNum = (tag: string) => { const v = g(tag); return v !== undefined ? Number(v) : undefined; };
  const gBool = (tag: string) => { const v = g(tag); return v !== undefined ? v === 'true' || v === '1' : undefined; };

  // Parse heading pairs
  let headingPairs: HeadingPair[] | undefined;
  const hpNode = root.children.find(c => localName(c.tag) === 'HeadingPairs');
  if (hpNode) {
    const vec = hpNode.children[0];
    if (vec) {
      const variants = vec.children.filter(c => localName(c.tag) === 'variant');
      headingPairs = [];
      for (let i = 0; i < variants.length - 1; i += 2) {
        const nameNode  = variants[i].children[0];
        const countNode = variants[i + 1].children[0];
        headingPairs.push({
          name:  nameNode?.text ?? '',
          count: parseInt(countNode?.text ?? '0', 10),
        });
      }
    }
  }

  // Parse titles of parts
  let titlesOfParts: string[] | undefined;
  const topNode = root.children.find(c => localName(c.tag) === 'TitlesOfParts');
  if (topNode) {
    const vec = topNode.children[0];
    if (vec) {
      titlesOfParts = vec.children.map(c => c.text ?? '');
    }
  }

  // Collect unknown tags for roundtrip preservation
  const knownTags = new Set([
    'Application','AppVersion','Company','Manager','DocSecurity','ScaleCrop',
    'LinksUpToDate','SharedDoc','HyperlinksChanged','Characters','CharactersWithSpaces',
    'Words','Lines','Paragraphs','Pages','Slides','Notes','HiddenSlides','MMClips',
    'Template','PresentationFormat','TotalTime','HyperlinkBase','HeadingPairs','TitlesOfParts',
  ]);
  const unknownNodes = root.children.filter(c => !knownTags.has(localName(c.tag)));
  // unknown raw preserved as-is
  // We serialise unknowns as raw XML to preserve them
  const unknownRaw = unknownNodes.map((n: XmlNode) => {
    // Rebuild simple tag
    const attrs = Object.entries(n.attrs).map(([k,v]) => ` ${k}="${v}"`).join('');
    const inner = n.text ?? '';
    if (!inner && !n.children.length) return `<${n.tag}${attrs}/>`;
    return `<${n.tag}${attrs}>${inner}</${n.tag}>`;
  }).join('\n');

  return {
    props: {
      application:        g('Application'),
      appVersion:         g('AppVersion'),
      company:            g('Company'),
      manager:            g('Manager'),
      docSecurity:        gNum('DocSecurity'),
      scaleCrop:          gBool('ScaleCrop'),
      linksUpToDate:      gBool('LinksUpToDate'),
      sharedDoc:          gBool('SharedDoc'),
      hyperlinksChanged:  gBool('HyperlinksChanged'),
      characters:         gNum('Characters'),
      charactersByWord:   gNum('CharactersWithSpaces'),
      words:              gNum('Words'),
      lines:              gNum('Lines'),
      paragraphs:         gNum('Paragraphs'),
      pages:              gNum('Pages'),
      slides:             gNum('Slides'),
      notes:              gNum('Notes'),
      hiddenSlides:       gNum('HiddenSlides'),
      mmClips:            gNum('MMClips'),
      template:           g('Template'),
      presentationFormat: g('PresentationFormat'),
      totalTime:          gNum('TotalTime'),
      hyperlinkBase:      g('HyperlinkBase'),
      headingPairs,
      titlesOfParts,
    },
    unknownRaw,
  };
}

export function parseCustomXml(xml: string): CustomProperty[] {
  const root = parseXml(xml);
  return root.children
    .filter(c => localName(c.tag) === 'property')
    .map(p => {
      const name = p.attrs['name'] ?? '';
      const valNode = p.children[0];
      if (!valNode) return null;
      const vTag = localName(valNode.tag);
      const text = valNode.text ?? '';
      let value: CustomPropValue;
      switch (vTag) {
        case 'lpwstr': case 'lpstr': case 'bstr':
          value = { type: 'string', value: text }; break;
        case 'i4': case 'int':
          value = { type: 'int', value: parseInt(text, 10) }; break;
        case 'decimal':
          value = { type: 'decimal', value: parseFloat(text) }; break;
        case 'bool':
          value = { type: 'bool', value: text === 'true' || text === '1' }; break;
        case 'filetime':
          value = { type: 'date', value: new Date(text) }; break;
        case 'r8':
          value = { type: 'r8', value: parseFloat(text) }; break;
        case 'i8':
          value = { type: 'i8', value: BigInt(text) }; break;
        case 'error':
          value = { type: 'error', value: text }; break;
        default:
          value = { type: 'string', value: text }; break;
      }
      return { name, value } as CustomProperty;
    })
    .filter(Boolean) as CustomProperty[];
}
