// Dependency-free .docx generation. A .docx is a zip (stored, no compression)
// of a handful of XML parts; we emit the minimal set Word/Google Docs accept.
// Pure module — testable in Node.

// ---- tiny zip writer (store method) ----
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}
const enc = new TextEncoder();
function u16(v) { return [v & 0xFF, (v >> 8) & 0xFF]; }
function u32(v) { return [v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF, (v >> 24) & 0xFF]; }

export function zipStore(files) {   // files: [{ name, text }]
  const chunks = [], central = [];
  let offset = 0;
  for (const f of files) {
    const name = enc.encode(f.name);
    const data = enc.encode(f.text);
    const crc = crc32(data);
    const local = new Uint8Array([
      0x50, 0x4B, 0x03, 0x04, ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0),
      ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(name.length), ...u16(0),
    ]);
    chunks.push(local, name, data);
    central.push(new Uint8Array([
      0x50, 0x4B, 0x01, 0x02, ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0),
      ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(name.length),
      ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(offset),
    ]), name);
    offset += local.length + name.length + data.length;
  }
  let cdSize = 0;
  for (const c of central) cdSize += c.length;
  const end = new Uint8Array([
    0x50, 0x4B, 0x05, 0x06, ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length),
    ...u32(cdSize), ...u32(offset), ...u16(0),
  ]);
  const total = offset + cdSize + end.length;
  const out = new Uint8Array(total);
  let p = 0;
  for (const c of [...chunks, ...central, end]) { out.set(c, p); p += c.length; }
  return out;
}

// ---- markdown-ish inline emphasis -> runs ----
export function toRuns(text) {
  const out = [];
  const re = /(\*\*[^*\n]+\*\*|\*[^*\n]+\*|_[^_\n]+_)/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push({ t: text.slice(last, m.index) });
    const s = m[0];
    if (s.startsWith('**')) out.push({ t: s.slice(2, -2), b: true });
    else out.push({ t: s.slice(1, -1), i: true });
    last = m.index + s.length;
  }
  if (last < text.length) out.push({ t: text.slice(last) });
  return out;
}

const xml = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function runXml(r) {
  const props = (r.i ? '<w:i/>' : '') + (r.b ? '<w:b/>' : '');
  return `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${xml(r.t)}</w:t></w:r>`;
}
// p: { runs?|text?, center, pageBreak, heading (1..n), noIndent }
function paraXml(p, smf) {
  const props = [];
  if (p.pageBreak) props.push('<w:pageBreakBefore/>');
  if (p.center) props.push('<w:jc w:val="center"/>');
  if (smf) {
    props.push('<w:spacing w:line="480" w:lineRule="auto"/>');
    if (!p.center && !p.heading && !p.noIndent) props.push('<w:ind w:firstLine="720"/>');
  } else if (!p.heading) {
    props.push('<w:spacing w:after="160"/>');
  }
  const runs = (p.runs || (p.text !== undefined ? toRuns(p.text) : []))
    .map(r => p.heading && !smf ? { ...r, b: true } : r);
  const sizeRpr = p.heading && !smf ? `<w:pPr>${props.join('')}</w:pPr>` : `<w:pPr>${props.join('')}</w:pPr>`;
  return `<w:p>${sizeRpr}${runs.map(runXml).join('')}</w:p>`;
}

// blocks: compiled markdown blocks ("# Heading", "* * *", paragraphs)
// opts: { title, titlePage: {contact[], by, approx}|null, smf, author }
export function makeDocx(blocks, opts) {
  const smf = !!opts.smf;
  const paras = [];

  if (opts.titlePage) {
    const tp = opts.titlePage;
    for (const line of tp.contact) paras.push({ text: line, noIndent: true });
    paras.push({ runs: [{ t: `approx. ${tp.approx} words` }], center: false, noIndent: true, rightish: true });
    for (let i = 0; i < 8; i++) paras.push({ runs: [], noIndent: true });
    paras.push({ runs: [{ t: opts.title }], center: true });
    if (tp.by) paras.push({ runs: [{ t: 'by ' + tp.by }], center: true });
  }

  let first = true;
  for (const b of blocks) {
    const h = b.match(/^(#{1,6}) (.*)$/s);
    if (h) {
      const level = h[1].length;
      paras.push({
        text: h[2].replace(/\n/g, ' '),
        heading: level,
        center: smf || level <= 2,
        pageBreak: level <= 2 && (!first || !!opts.titlePage),
        noIndent: true,
      });
      first = false;
      continue;
    }
    if (b.trim() === '* * *' || b.trim() === '#') {
      paras.push({ runs: [{ t: smf ? '#' : '* * *' }], center: true, noIndent: true });
      continue;
    }
    for (const para of b.split(/\n+/)) if (para.trim()) paras.push({ text: para });
    first = false;
  }

  // the word-count line sits right-aligned opposite the contact block
  const bodyXml = paras.map(p => {
    if (p.rightish) return `<w:p><w:pPr><w:jc w:val="right"/>${smf ? '<w:spacing w:line="480" w:lineRule="auto"/>' : ''}</w:pPr>${(p.runs || []).map(runXml).join('')}</w:p>`;
    return paraXml(p, smf);
  }).join('');

  const headerRef = smf ? '<w:headerReference w:type="default" r:id="rId9"/>' : '';
  const titlePg = opts.titlePage ? '<w:titlePg/>' : '';
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${bodyXml}<w:sectPr>${headerRef}${titlePg}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`;

  const font = smf ? 'Times New Roman' : 'Georgia';
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`;

  // running header: "Surname / TITLE / page"
  const surname = (opts.author || '').trim().split(/\s+/).pop() || '';
  const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t xml:space="preserve">${xml(surname ? surname + ' / ' : '')}${xml(opts.title.toUpperCase())} / </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>`;

  const files = [
    { name: '[Content_Types].xml', text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${smf ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : ''}</Types>` },
    { name: '_rels/.rels', text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>` },
    { name: 'word/_rels/document.xml.rels', text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${smf ? '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' : ''}</Relationships>` },
    { name: 'word/document.xml', text: documentXml },
    { name: 'word/styles.xml', text: stylesXml },
  ];
  if (smf) files.push({ name: 'word/header1.xml', text: headerXml });
  return zipStore(files);
}
