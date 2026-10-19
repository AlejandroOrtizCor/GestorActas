import PizZip from 'pizzip';
import { addFieldsFromText, placeholderRegex, formatValue, splitId } from './placeholders';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const parts = (zip) =>
  Object.keys(zip.files).filter((n) => /^word\/(document|header\d*|footer\d*)\.xml$/.test(n));
const parse = (xml) => new DOMParser().parseFromString(xml, 'application/xml');
const textNodes = (p) => [...p.getElementsByTagNameNS(W, 't')];

const loadDocs = (zip) => parts(zip).map((name) => ({ name, doc: parse(zip.file(name).asText()) }));
const paragraphText = (p) => textNodes(p).map((t) => t.textContent).join('');

// Si la fila tiene variables "Grupo.Columna" devuelve { group, ids, label }
// label = primer texto fijo de la fila que no sea solo un número (ej. "Casco de seguridad")
function rowInfo(tr) {
  let group = null;
  const ids = [];
  for (const p of tr.getElementsByTagNameNS(W, 'p')) {
    for (const m of paragraphText(p).matchAll(placeholderRegex())) {
      const id = m[2].trim();
      const g = splitId(id);
      if (!g) continue;
      group ??= g.group;
      if (g.group === group && !ids.includes(id)) ids.push(id);
    }
  }
  if (!group) return null;
  let label = '';
  for (const tc of tr.childNodes) {
    if (tc.localName !== 'tc') continue;
    const text = [...tc.getElementsByTagNameNS(W, 'p')].map(paragraphText).join(' ')
      .replace(placeholderRegex(), '').trim();
    if (text && !/^\d+$/.test(text)) { label = text; break; }
  }
  return { group, ids, label };
}

function scanRows(docs) {
  const rows = [];
  for (const { name, doc } of docs) {
    for (const tr of doc.getElementsByTagNameNS(W, 'tr')) {
      const info = rowInfo(tr);
      if (info) rows.push({ ...info, name, doc, tr });
    }
  }
  return rows;
}

// Devuelve { fields, fixed }
//  fields: todos los campos únicos
//  fixed:  { Grupo: [ { label, ids }, ... ] } para tablas con las variables en 2+ filas (filas fijas)
export function readDocxFields(buffer) {
  const zip = new PizZip(buffer);
  const docs = loadDocs(zip);
  const map = new Map();
  for (const { doc } of docs) {
    for (const p of doc.getElementsByTagNameNS(W, 'p')) addFieldsFromText(map, paragraphText(p));
  }
  const byGroup = {};
  for (const r of scanRows(docs)) (byGroup[r.group] ??= []).push({ label: r.label, ids: r.ids });
  const fixed = Object.fromEntries(Object.entries(byGroup).filter(([, rows]) => rows.length > 1));
  return { fields: [...map.values()], fixed };
}

// Word parte el texto en varios "runs" (<w:t>). Reemplazamos sobre el texto
// concatenado del párrafo y repartimos el cambio entre los runs afectados.
function replaceRange(ts, start, end, replacement) {
  let pos = 0, first = null;
  for (const t of ts) {
    const s = t.textContent, a = pos, b = pos + s.length;
    pos = b;
    if (!(a < end && b > start)) continue;
    const from = Math.max(start, a) - a, to = Math.min(end, b) - a;
    t.textContent = s.slice(0, from) + (first ? '' : replacement) + s.slice(to);
    t.setAttribute('xml:space', 'preserve');
    first ??= t;
  }
  return first;
}

const bytes = (dataUrl) => Uint8Array.from(atob(dataUrl.split(',')[1]), (c) => c.charCodeAt(0));

function registerImage(zip, part, img, n) {
  const media = `fill_${n}.png`;
  zip.file(`word/media/${media}`, bytes(img.dataUrl));
  const relsPath = part.replace(/([^/]+)$/, '_rels/$1.rels');
  let rels = zip.file(relsPath)?.asText() ??
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const rid = `rIdFill${n}`;
  rels = rels.replace('</Relationships>',
    `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${media}"/></Relationships>`);
  zip.file(relsPath, rels);
  let ct = zip.file('[Content_Types].xml').asText();
  if (!/Extension="png"/i.test(ct)) {
    ct = ct.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>');
    zip.file('[Content_Types].xml', ct);
  }
  return rid;
}

function imageRun(rid, img, n) {
  const scale = Math.min(1, 450 / img.w);
  const cx = Math.round(img.w * scale * 9525), cy = Math.round(img.h * scale * 9525);
  return `<w:r xmlns:w="${W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${1000 + n}" name="img${n}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="img${n}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

// Llena todos los párrafos dentro de `root` (documento o fila clonada)
function fillParagraphs(root, doc, zip, part, values, counter) {
  for (const p of [...root.getElementsByTagNameNS(W, 'p')]) {
    const ts = textNodes(p);
    if (!ts.length) continue;
    const text = ts.map((t) => t.textContent).join('');
    for (const m of [...text.matchAll(placeholderRegex())].reverse()) {
      const [full, type, rawId] = m;
      const v = values[rawId.trim()];
      const end = m.index + full.length;
      if (type === 'i') {
        const first = replaceRange(ts, m.index, end, '');
        if (v && first) {
          const n = ++counter.n;
          const rid = registerImage(zip, part, v, n);
          const run = doc.importNode(parse(imageRun(rid, v, n)).documentElement, true);
          const r = first.parentNode;
          r.parentNode.insertBefore(run, r.nextSibling);
        }
      } else {
        replaceRange(ts, m.index, end, formatValue(type, v));
      }
    }
  }
}

// Evita ids duplicados al clonar filas
function stripIds(el) {
  for (const node of el.getElementsByTagName('*')) {
    for (const a of [...node.attributes]) {
      if (a.localName === 'paraId' || a.localName === 'textId') node.removeAttribute(a.name);
    }
  }
}

// values: campos sueltos { id: valor }
// tables: { Grupo: [ { 'Grupo.Columna': valor, ... }, ... ] }
//  - variables en UNA fila  -> esa fila se repite por cada fila de datos
//  - variables en 2+ filas  -> filas fijas: la fila k recibe los datos de la fila k
export function fillDocx(buffer, values, tables = {}) {
  const zip = new PizZip(buffer);
  const counter = { n: 0 };
  const docs = loadDocs(zip);

  const rows = scanRows(docs);
  const total = {};
  for (const r of rows) total[r.group] = (total[r.group] ?? 0) + 1;
  const seen = {};

  for (const r of rows) {
    const k = (seen[r.group] = (seen[r.group] ?? -1) + 1);
    if (!r.doc.documentElement.contains(r.tr)) continue;
    const data = tables[r.group] ?? [];
    if (total[r.group] > 1) {
      fillParagraphs(r.tr, r.doc, zip, r.name, { ...values, ...(data[k] ?? {}) }, counter);
    } else {
      for (const row of data.length ? data : [{}]) {
        const clone = r.tr.cloneNode(true);
        stripIds(clone);
        fillParagraphs(clone, r.doc, zip, r.name, { ...values, ...row }, counter);
        r.tr.parentNode.insertBefore(clone, r.tr);
      }
      r.tr.parentNode.removeChild(r.tr);
    }
  }

  for (const { name, doc } of docs) {
    fillParagraphs(doc, doc, zip, name, values, counter);
    zip.file(name, new XMLSerializer().serializeToString(doc));
  }
  return zip.generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}