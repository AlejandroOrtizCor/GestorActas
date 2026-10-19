import ExcelJS from 'exceljs';
import { addFieldsFromText, placeholderRegex, fillText, splitId } from './placeholders';

const cellText = (c) =>
  typeof c.value === 'string' ? c.value
  : c.value?.richText ? c.value.richText.map((r) => r.text).join('')
  : null;

const eachCell = (wb, fn) =>
  wb.eachSheet((ws) => ws.eachRow({ includeEmpty: false }, (row) =>
    row.eachCell({ includeEmpty: false }, (cell) => {
      const text = cellText(cell);
      if (text) fn(ws, cell, text);
    })));

const load = async (buffer) => { const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buffer); return wb; };

// Filas con variables "Grupo.Columna": { ws, row, group, ids, label }
function scanTemplates(wb) {
  const list = [];
  wb.eachSheet((ws) => ws.eachRow({ includeEmpty: false }, (row) => {
    let group = null, label = '';
    const ids = [];
    row.eachCell({ includeEmpty: false }, (cell) => {
      const t = cellText(cell);
      if (!t) return;
      for (const m of t.matchAll(placeholderRegex())) {
        const id = m[2].trim();
        const g = splitId(id);
        if (!g) continue;
        group ??= g.group;
        if (g.group === group && !ids.includes(id)) ids.push(id);
      }
      const fixedText = t.replace(placeholderRegex(), '').trim();
      if (!label && fixedText && !/^\d+$/.test(fixedText)) label = fixedText;
    });
    if (group) list.push({ ws, row: row.number, group, ids, label });
  }));
  return list;
}

// Devuelve { fields, fixed } (ver readDocxFields)
export async function readXlsxFields(buffer) {
  const wb = await load(buffer);
  const map = new Map();
  eachCell(wb, (_, __, text) => addFieldsFromText(map, text));
  const byGroup = {};
  for (const t of scanTemplates(wb)) (byGroup[t.group] ??= []).push({ label: t.label, ids: t.ids });
  const fixed = Object.fromEntries(Object.entries(byGroup).filter(([, rows]) => rows.length > 1));
  return { fields: [...map.values()], fixed };
}

function fillCell(wb, ws, cell, text, values) {
  const matches = [...text.matchAll(placeholderRegex())];
  if (!matches.length) return;
  const only = matches.length === 1 && matches[0][0] === text.trim() ? matches[0] : null;

  if (only && only[1] === 'i') {            // imagen en la celda
    const img = values[only[2].trim()];
    cell.value = '';
    if (img) {
      const id = wb.addImage({ base64: img.dataUrl, extension: 'png' });
      const scale = Math.min(1, 300 / img.w);
      ws.addImage(id, {
        tl: { col: cell.col - 1, row: cell.row - 1 },
        ext: { width: img.w * scale, height: img.h * scale },
      });
    }
    return;
  }
  if (only && only[1] === 'n') {            // número real, no texto
    const num = Number(values[only[2].trim()]);
    if (values[only[2].trim()] && !Number.isNaN(num)) { cell.value = num; return; }
  }
  cell.value = fillText(text, values);
}

// values: campos sueltos { id: valor }
// tables: { Grupo: [ { 'Grupo.Columna': valor, ... }, ... ] }
//  - variables en UNA fila -> se duplica por cada fila de datos
//  - variables en 2+ filas -> filas fijas: la fila k recibe los datos de la fila k
export async function fillXlsx(buffer, values, tables = {}) {
  const wb = await load(buffer);

  const templates = scanTemplates(wb);
  const total = {}, seen = {};
  for (const t of templates) total[t.group] = (total[t.group] ?? 0) + 1;
  for (const t of templates) {
    t.k = seen[t.group] = (seen[t.group] ?? -1) + 1;
    t.fixed = total[t.group] > 1;
  }

  // De abajo hacia arriba para que insertar filas no mueva las plantillas pendientes
  for (const t of [...templates].reverse()) {
    const data = tables[t.group] ?? [];
    const list = t.fixed ? [data[t.k] ?? {}] : (data.length ? data : [{}]);
    if (list.length > 1) t.ws.duplicateRow(t.row, list.length - 1, true); // copia valores y estilos
    list.forEach((d, k) => {
      t.ws.getRow(t.row + k).eachCell({ includeEmpty: false }, (cell) => {
        const text = cellText(cell);
        if (text) fillCell(wb, t.ws, cell, text, { ...values, ...d });
      });
    });
  }

  // 2) Resto de celdas
  eachCell(wb, (ws, cell, text) => fillCell(wb, ws, cell, text, values));

  const out = await wb.xlsx.writeBuffer();
  return new Blob([out], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}