// Tipos soportados
export const TYPES = {
  t: 'Texto', n: 'Número entero positivo', nt: 'Teléfono (10 dígitos, empieza por 3)',
  c: 'Correo electrónico', f: 'Fecha', dma: 'Fecha en letras', fh: 'Fecha y hora', h: 'Hora', i: 'Imagen',
};

const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio',
  'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

// t:Nombre del cliente:t        -> texto libre
// t:Nombre del cliente:t:50     -> texto de máximo 50 caracteres
// n:Valor:n:8                   -> número de máximo 8 dígitos
// El tipo se repite al final (:\1). Los tipos largos (fh, nt, dma) van primero para que no se confundan con t, n o f.
const SOURCE = '(?<![A-Za-z0-9])(fh|nt|dma|t|n|c|f|h|i):([^:<>]+?):\\1(?::(\\d+))?(?![A-Za-z0-9])';
export const placeholderRegex = () => new RegExp(SOURCE, 'g');

// "Productos.Nombre" -> tabla "Productos", columna "Nombre".
// El punto debe ir pegado a la columna (así "No. documento" sigue siendo un campo normal).
export function splitId(id) {
  const m = /^(.+?)\.(\S.*)$/.exec(id);
  return m ? { group: m[1].trim(), label: m[2].trim() } : null;
}

// Devuelve los campos únicos (por id) encontrados en un texto
export function addFieldsFromText(map, text) {
  for (const m of text.matchAll(placeholderRegex())) {
    const [, type, rawId, max] = m;
    const id = rawId.trim();
    if (map.has(id)) continue;
    const g = splitId(id);
    map.set(id, { id, type, max: max ? Number(max) : null, group: g?.group ?? null, label: g?.label ?? id });
  }
}

const pad = (s) => String(s).padStart(2, '0');

// Cómo se escribe cada valor dentro del documento
export function formatValue(type, value) {
  if (value == null || value === '') return '';
  if (type === 'dma') {
    const [y, m, d] = value.split('-');
    return `${Number(d)} de ${MESES[Number(m) - 1]} del ${y}`;
  }
  if (type === 'f') { const [y, m, d] = value.split('-'); return `${d}/${m}/${y}`; }
  if (type === 'fh') {
    const [date, time] = value.split('T'); const [y, m, d] = date.split('-');
    return `${d}/${m}/${y} ${time}`;
  }
  return String(value);
}

export const fillText = (text, values) =>
  text.replace(placeholderRegex(), (_, type, id) => formatValue(type, values[id.trim()]));