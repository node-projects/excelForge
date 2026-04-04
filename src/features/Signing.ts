/**
 * Digital Signature support for OOXML packages and VBA projects.
 *
 * - OOXML signatures: XML-DSig (ECMA-376 Part 2) stored in _xmlsignatures/
 * - VBA signatures: [MS-OVBA] §2.4.2 DigitalSignatureEx (V3, SHA-256)
 *
 * Uses Web Crypto API (SubtleCrypto) for all cryptographic operations.
 */

// ── ASN.1 / DER helpers ────────────────────────────────────────────────────

function derLen(len: number): Uint8Array {
  if (len < 0x80) return new Uint8Array([len]);
  if (len < 0x100) return new Uint8Array([0x81, len]);
  if (len < 0x10000) return new Uint8Array([0x82, len >> 8, len & 0xFF]);
  return new Uint8Array([0x83, (len >> 16) & 0xFF, (len >> 8) & 0xFF, len & 0xFF]);
}

function derTag(tag: number, content: Uint8Array): Uint8Array {
  const len = derLen(content.length);
  const out = new Uint8Array(1 + len.length + content.length);
  out[0] = tag;
  out.set(len, 1);
  out.set(content, 1 + len.length);
  return out;
}

function derSeq(...items: Uint8Array[]): Uint8Array {
  let total = 0;
  for (const i of items) total += i.length;
  const body = new Uint8Array(total);
  let off = 0;
  for (const i of items) { body.set(i, off); off += i.length; }
  return derTag(0x30, body);
}

function derSet(...items: Uint8Array[]): Uint8Array {
  let total = 0;
  for (const i of items) total += i.length;
  const body = new Uint8Array(total);
  let off = 0;
  for (const i of items) { body.set(i, off); off += i.length; }
  return derTag(0x31, body);
}

function derOid(oid: number[]): Uint8Array {
  const bytes: number[] = [40 * oid[0] + oid[1]];
  for (let i = 2; i < oid.length; i++) {
    let v = oid[i];
    if (v < 128) { bytes.push(v); }
    else {
      const parts: number[] = [];
      while (v > 0) { parts.unshift(v & 0x7F); v >>= 7; }
      for (let j = 0; j < parts.length - 1; j++) parts[j] |= 0x80;
      bytes.push(...parts);
    }
  }
  return derTag(0x06, new Uint8Array(bytes));
}

function derInt(n: number): Uint8Array {
  if (n === 0) return derTag(0x02, new Uint8Array([0]));
  const bytes: number[] = [];
  let v = n;
  while (v > 0) { bytes.unshift(v & 0xFF); v >>= 8; }
  if (bytes[0] & 0x80) bytes.unshift(0);
  return derTag(0x02, new Uint8Array(bytes));
}

function derOctetStr(data: Uint8Array): Uint8Array {
  return derTag(0x04, data);
}

function derUtf8Str(s: string): Uint8Array {
  return derTag(0x0C, new TextEncoder().encode(s));
}

function derExplicit(tag: number, content: Uint8Array): Uint8Array {
  return derTag(0xA0 | tag, content);
}

function derBool(val: boolean): Uint8Array {
  return derTag(0x01, new Uint8Array([val ? 0xFF : 0x00]));
}

// ── OID constants ────────────────────────────────────────────────────────────
const OID_SHA256 = [2, 16, 840, 1, 101, 3, 4, 2, 1];
const OID_RSA_ENCRYPTION = [1, 2, 840, 113549, 1, 1, 1];
const OID_SHA256_WITH_RSA = [1, 2, 840, 113549, 1, 1, 11];
const OID_SIGNED_DATA = [1, 2, 840, 113549, 1, 7, 2];
const OID_DATA = [1, 2, 840, 113549, 1, 7, 1];
const OID_CONTENT_TYPE = [1, 2, 840, 113549, 1, 9, 3];
const OID_MESSAGE_DIGEST = [1, 2, 840, 113549, 1, 9, 4];
const OID_SIGNING_TIME = [1, 2, 840, 113549, 1, 9, 5];
const OID_SPC_INDIRECT_DATA = [1, 3, 6, 1, 4, 1, 311, 2, 1, 4];

// ── PEM parsing ──────────────────────────────────────────────────────────────

function parsePem(pem: string): Uint8Array {
  const b64 = pem.replace(/-----[A-Z\s]+-----/g, '').replace(/\s+/g, '');
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

function base64Encode(data: Uint8Array): string {
  let s = '';
  for (let i = 0; i < data.length; i++) s += String.fromCharCode(data[i]);
  return btoa(s);
}

function toHex(data: Uint8Array): string {
  return Array.from(data).map(b => b.toString(16).padStart(2, '0')).join('');
}

// ── DER parsing helpers ──────────────────────────────────────────────────────

function readDerTLV(data: Uint8Array, offset: number): { tag: number; content: Uint8Array; next: number } {
  const tag = data[offset];
  let off = offset + 1;
  let len = data[off++];
  if (len & 0x80) {
    const numBytes = len & 0x7F;
    len = 0;
    for (let i = 0; i < numBytes; i++) len = (len << 8) | data[off++];
  }
  return { tag, content: data.slice(off, off + len), next: off + len };
}

interface CertInfo {
  issuer: Uint8Array;
  serial: Uint8Array;
  raw: Uint8Array;
}

function extractCertInfo(certDer: Uint8Array): CertInfo {
  // Certificate SEQUENCE → tbsCertificate SEQUENCE
  const cert = readDerTLV(certDer, 0);
  const tbs = readDerTLV(cert.content, 0);
  let off = 0;
  // version [0] EXPLICIT
  let tlv = readDerTLV(tbs.content, off);
  if (tlv.tag === 0xA0) { off = tlv.next; tlv = readDerTLV(tbs.content, off); }
  // serial INTEGER
  const serial = tbs.content.slice(off, tlv.next);
  off = tlv.next;
  // algorithm
  tlv = readDerTLV(tbs.content, off); off = tlv.next;
  // issuer SEQUENCE
  const issuer = tbs.content.slice(off);
  tlv = readDerTLV(tbs.content, off);
  const issuerPart = tbs.content.slice(off, tlv.next);

  return { issuer: issuerPart, serial: readDerTLV(tbs.content, serial === tbs.content.slice(off) ? 0 : (readDerTLV(tbs.content, 0).tag === 0xA0 ? readDerTLV(tbs.content, 0).next : 0)).content, raw: certDer };
}

// ── SubtleCrypto helpers ─────────────────────────────────────────────────────

async function sha256(data: Uint8Array): Promise<Uint8Array> {
  const buf = await crypto.subtle.digest('SHA-256', data as Uint8Array<ArrayBuffer>);
  return new Uint8Array(buf);
}

async function importPrivateKey(pkcs8Der: Uint8Array): Promise<CryptoKey> {
  return crypto.subtle.importKey('pkcs8', pkcs8Der as Uint8Array<ArrayBuffer>, { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['sign']);
}

async function rsaSign(key: CryptoKey, data: Uint8Array): Promise<Uint8Array> {
  const sig = await crypto.subtle.sign('RSASSA-PKCS1-v1_5', key, data as Uint8Array<ArrayBuffer>);
  return new Uint8Array(sig);
}

// ── PKCS#7 / CMS SignedData builder ──────────────────────────────────────────

async function buildCmsSignature(
  contentOid: number[],
  contentData: Uint8Array,
  certDer: Uint8Array,
  privateKey: CryptoKey,
): Promise<Uint8Array> {
  const digest = await sha256(contentData);

  // Parse cert info for issuer + serial
  const certSeq = readDerTLV(certDer, 0);
  const tbsSeq = readDerTLV(certSeq.content, 0);
  let off = 0;
  let tlv = readDerTLV(tbsSeq.content, off);
  if (tlv.tag === 0xA0) { off = tlv.next; tlv = readDerTLV(tbsSeq.content, off); }
  const serialContent = tlv.content;
  off = tlv.next;
  // skip algorithm
  tlv = readDerTLV(tbsSeq.content, off); off = tlv.next;
  // issuer
  tlv = readDerTLV(tbsSeq.content, off);
  const issuerBytes = tbsSeq.content.slice(off, tlv.next);

  // Build authenticated attributes
  const now = new Date();
  const utcTime = now.toISOString().replace(/[-:T]/g, '').slice(0, 14) + 'Z';
  const utcTimeBytes = new TextEncoder().encode(utcTime);

  const algId = derSeq(derOid(OID_SHA256), derTag(0x05, new Uint8Array(0)));

  const authAttrs = [
    derSeq(derOid(OID_CONTENT_TYPE), derSet(derOid(contentOid))),
    derSeq(derOid(OID_SIGNING_TIME), derSet(derTag(0x17, utcTimeBytes))),
    derSeq(derOid(OID_MESSAGE_DIGEST), derSet(derOctetStr(digest))),
  ];

  let authAttrsBody = new Uint8Array(0);
  for (const a of authAttrs) {
    const tmp = new Uint8Array(authAttrsBody.length + a.length);
    tmp.set(authAttrsBody); tmp.set(a, authAttrsBody.length);
    authAttrsBody = tmp;
  }
  const authAttrsSet = derTag(0x31, authAttrsBody);

  // Sign the SET OF authenticated attributes (but change tag to 0x31 for DER encoding)
  const authAttrsDer = derTag(0xA0, authAttrsBody);
  // For signing, use SET (0x31) encoding
  const sigInput = authAttrsSet;
  const signature = await rsaSign(privateKey, sigInput);

  // Build SignerInfo
  const signerInfo = derSeq(
    derInt(1), // version
    derSeq(issuerBytes, derTag(0x02, serialContent)), // issuerAndSerialNumber
    algId, // digestAlgorithm
    authAttrsDer, // authenticatedAttributes [0] IMPLICIT
    derSeq(derOid(OID_RSA_ENCRYPTION), derTag(0x05, new Uint8Array(0))), // signatureAlgorithm
    derOctetStr(signature), // signature
  );

  // Build SignedData
  const encapContentInfo = derSeq(
    derOid(contentOid),
    derExplicit(0, derOctetStr(contentData)),
  );

  const signedData = derSeq(
    derInt(1), // version
    derSet(algId), // digestAlgorithms
    encapContentInfo,
    derExplicit(0, certDer), // certificates [0] IMPLICIT
    derSet(signerInfo),
  );

  // ContentInfo wrapper
  return derSeq(
    derOid(OID_SIGNED_DATA),
    derExplicit(0, signedData),
  );
}

// ── Self-signed test certificate generator ──────────────────────────────────

const OID_COMMON_NAME = [2, 5, 4, 3];

function derBitStr(data: Uint8Array): Uint8Array {
  const inner = new Uint8Array(1 + data.length);
  inner[0] = 0; // no unused bits
  inner.set(data, 1);
  return derTag(0x03, inner);
}

function derBigInt(bytes: Uint8Array): Uint8Array {
  // Ensure leading zero if high bit set (positive)
  if (bytes[0] & 0x80) {
    const padded = new Uint8Array(bytes.length + 1);
    padded.set(bytes, 1);
    return derTag(0x02, padded);
  }
  return derTag(0x02, bytes);
}

/**
 * Generate a self-signed X.509 certificate for testing.
 * Returns PEM-encoded certificate string.
 *
 * @param subject  Common Name for the cert (e.g. "ExcelForge Test")
 * @param privateKeyPem PEM-encoded PKCS#8 private key
 * @param publicKeySpkiDer DER-encoded SubjectPublicKeyInfo
 */
export async function generateTestCertificate(
  subject: string,
  privateKeyPem: string,
  publicKeySpkiDer: Uint8Array,
): Promise<string> {
  const keyDer = parsePem(privateKeyPem);
  const key = await importPrivateKey(keyDer);

  const algId = derSeq(derOid(OID_SHA256_WITH_RSA), derTag(0x05, new Uint8Array(0)));
  const name = derSeq(derSet(derSeq(derOid(OID_COMMON_NAME), derUtf8Str(subject))));

  // Validity: now to +10 years
  const now = new Date();
  const later = new Date(now.getFullYear() + 10, now.getMonth(), now.getDate());
  const utcFmt = (d: Date) => {
    const s = d.toISOString().replace(/[-:T]/g, '').slice(2, 14) + 'Z';
    return derTag(0x17, new TextEncoder().encode(s)); // UTCTime
  };
  const validity = derSeq(utcFmt(now), utcFmt(later));

  // TBSCertificate
  const serial = derTag(0x02, new Uint8Array([0x01])); // serial = 1
  const version = derExplicit(0, derInt(2)); // v3
  const tbs = derSeq(version, serial, algId, name, validity, name, derTag(0x30, publicKeySpkiDer.slice(publicKeySpkiDer.indexOf(0x30, 1) >= 0 ? 0 : 0)));

  // Actually we need to wrap the raw SPKI properly. SPKI is already a SEQUENCE, use as-is.
  const tbs2 = derSeq(version, serial, algId, name, validity, name, publicKeySpkiDer);

  // Sign the TBS
  const tbsDer = tbs2;
  const tbsHash = await sha256(tbsDer);
  // DigestInfo for PKCS#1 v1.5: SEQUENCE { SEQUENCE { OID sha256, NULL }, OCTET STRING hash }
  // But RSASSA-PKCS1-v1_5 in WebCrypto does the DigestInfo wrapping internally, so we sign raw data
  const signature = await crypto.subtle.sign('RSASSA-PKCS1-v1_5', key, tbsDer as Uint8Array<ArrayBuffer>);
  const sigBytes = new Uint8Array(signature);

  // Certificate
  const cert = derSeq(tbsDer, algId, derBitStr(sigBytes));

  // Convert to PEM
  const b64 = base64Encode(cert);
  const lines = b64.match(/.{1,64}/g)!.join('\n');
  return `-----BEGIN CERTIFICATE-----\n${lines}\n-----END CERTIFICATE-----`;
}

// ── Signing Options ──────────────────────────────────────────────────────────

export interface SigningOptions {
  /** PEM-encoded X.509 certificate */
  certificate: string;
  /** PEM-encoded PKCS#8 private key */
  privateKey: string;
}

// ── OOXML Package Digital Signature ──────────────────────────────────────────

/**
 * Add an XML-DSig digital signature to an OOXML package.
 * Returns entries to add to the ZIP: _xmlsignatures/sig1.xml, origin.sigs,
 * and relationships.
 */
export async function signPackage(
  partsToSign: Map<string, Uint8Array>,
  options: SigningOptions,
): Promise<Map<string, Uint8Array>> {
  const certDer = parsePem(options.certificate);
  const keyDer = parsePem(options.privateKey);
  const privateKey = await importPrivateKey(keyDer);
  const enc = new TextEncoder();
  const result = new Map<string, Uint8Array>();

  // Compute digests of all parts
  const refXmls: string[] = [];
  for (const [uri, data] of partsToSign) {
    const digest = await sha256(data);
    const b64Digest = base64Encode(digest);
    refXmls.push(`<Reference URI="/${uri}?ContentType=${getContentType(uri)}">
  <DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>
  <DigestValue>${b64Digest}</DigestValue>
</Reference>`);
  }

  // Build SignedInfo
  const signedInfoXml = `<SignedInfo xmlns="http://www.w3.org/2000/09/xmldsig#">
<CanonicalizationMethod Algorithm="http://www.w3.org/TR/2001/REC-xml-c14n-20010315"/>
<SignatureMethod Algorithm="http://www.w3.org/2001/04/xmldsig-more#rsa-sha256"/>
${refXmls.join('\n')}
</SignedInfo>`;

  // Sign the SignedInfo (rsaSign uses RSASSA-PKCS1-v1_5 which hashes internally)
  const signedInfoBytes = enc.encode(signedInfoXml);
  const signature = await rsaSign(privateKey, signedInfoBytes);
  const b64Sig = base64Encode(signature);
  const b64Cert = base64Encode(certDer);

  // Build the full signature XML
  const sigXml = `<?xml version="1.0" encoding="UTF-8"?>
<Signature xmlns="http://www.w3.org/2000/09/xmldsig#" Id="idPackageSignature">
${signedInfoXml}
<SignatureValue>${b64Sig}</SignatureValue>
<KeyInfo>
  <X509Data>
    <X509Certificate>${b64Cert}</X509Certificate>
  </X509Data>
</KeyInfo>
<Object>
  <SignatureProperties>
    <SignatureProperty Id="idSignatureTime" Target="#idPackageSignature">
      <mdssi:SignatureTime xmlns:mdssi="http://schemas.openxmlformats.org/package/2006/digital-signature">
        <mdssi:Format>YYYY-MM-DDThh:mm:ssTZD</mdssi:Format>
        <mdssi:Value>${new Date().toISOString()}</mdssi:Value>
      </mdssi:SignatureTime>
    </SignatureProperty>
  </SignatureProperties>
</Object>
</Signature>`;

  result.set('_xmlsignatures/sig1.xml', enc.encode(sigXml));
  result.set('_xmlsignatures/origin.sigs', new Uint8Array(0));
  result.set('_xmlsignatures/_rels/origin.sigs.rels', enc.encode(
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="sig1.xml"/>
</Relationships>`
  ));

  // Update root _rels/.rels to include digital signature origin relationship
  const existingRels = partsToSign.get('_rels/.rels');
  if (existingRels) {
    let relsXml = new TextDecoder().decode(existingRels);
    if (!relsXml.includes('digital-signature/origin')) {
      relsXml = relsXml.replace('</Relationships>',
        '  <Relationship Id="rIdSig" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="_xmlsignatures/origin.sigs"/>\n</Relationships>');
      result.set('_rels/.rels', enc.encode(relsXml));
    }
  }

  return result;
}

function getContentType(uri: string): string {
  if (uri.endsWith('.xml')) return 'application/xml';
  if (uri.endsWith('.rels')) return 'application/vnd.openxmlformats-package.relationships+xml';
  if (uri.endsWith('.bin')) return 'application/vnd.ms-office.vbaProject';
  return 'application/octet-stream';
}

// ── VBA Code Signing ─────────────────────────────────────────────────────────

/**
 * Sign a VBA project. Returns DigitalSignatureEx stream data
 * to embed in the VBA CFB container.
 *
 * Format: [MS-OVBA] §2.4.2 — SpcIndirectDataContent + PKCS#7 SignedData
 */
export async function signVbaProject(
  vbaProjectBin: Uint8Array,
  options: SigningOptions,
): Promise<Uint8Array> {
  const certDer = parsePem(options.certificate);
  const keyDer = parsePem(options.privateKey);
  const privateKey = await importPrivateKey(keyDer);

  // Hash the VBA project content (the dir stream + module sources)
  const projectHash = await sha256(vbaProjectBin);

  // Build SpcIndirectDataContent
  const spcData = derSeq(
    derSeq(
      derOid(OID_SPC_INDIRECT_DATA),
      derSeq(derOid(OID_SHA256), derTag(0x05, new Uint8Array(0))),
    ),
    derSeq(
      derSeq(derOid(OID_SHA256), derTag(0x05, new Uint8Array(0))),
      derOctetStr(projectHash),
    ),
  );

  // Build CMS SignedData wrapping the SpcIndirectDataContent
  const cms = await buildCmsSignature(OID_SPC_INDIRECT_DATA, spcData, certDer, privateKey);

  // DigitalSignatureEx format: version(4) + signatureSize(4) + signature
  const result = new Uint8Array(8 + cms.length);
  const dv = new DataView(result.buffer);
  dv.setUint32(0, 3, true);       // version = 3 (V3 = SHA-256)
  dv.setUint32(4, cms.length, true);
  result.set(cms, 8);

  return result;
}

// ── Convenience: sign both package and VBA ───────────────────────────────────

export interface SignResult {
  /** Entries to add to the XLSX ZIP for the package signature */
  packageSignatureEntries: Map<string, Uint8Array>;
  /** DigitalSignatureEx stream bytes (if VBA project was provided) */
  vbaSignature?: Uint8Array;
}

/**
 * Sign an OOXML workbook package and optionally its VBA project.
 *
 * @param parts - Map of part names to part data from the XLSX ZIP
 * @param options - Signing certificate and private key (PEM)
 * @param vbaProjectBin - Optional VBA project binary to sign
 */
export async function signWorkbook(
  parts: Map<string, Uint8Array>,
  options: SigningOptions,
  vbaProjectBin?: Uint8Array,
): Promise<SignResult> {
  const packageSignatureEntries = await signPackage(parts, options);
  let vbaSignature: Uint8Array | undefined;
  if (vbaProjectBin) {
    vbaSignature = await signVbaProject(vbaProjectBin, options);
  }
  return { packageSignatureEntries, vbaSignature };
}
