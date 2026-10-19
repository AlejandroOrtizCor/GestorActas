import { useEffect, useRef, useState } from 'react';
import { readDocxFields, fillDocx } from '../lib/docx';
import { readXlsxFields, fillXlsx } from '../lib/xlsx';
import { docxToPdf } from '../lib/pdf';

async function fileToPng(file) {
  const bmp = await createImageBitmap(file);
  const k = Math.min(1, 1200 / bmp.width);
  const c = document.createElement('canvas');
  c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  return { dataUrl: c.toDataURL('image/png'), w: c.width, h: c.height };
}

function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

// Un control de formulario según el tipo de campo (sirve para campos sueltos y celdas de tabla)
function renderInput(f, value, onChange, id) {
  const common = { id, required: true, value: value ?? '', 'aria-label': f.label };
  switch (f.type) {
    case 't':
      return f.max && f.max > 120
        ? <textarea {...common} maxLength={f.max} rows={4} onChange={(e) => onChange(e.target.value)} />
        : <input {...common} maxLength={f.max ?? undefined} onChange={(e) => onChange(e.target.value)} />;
    case 'n': // entero positivo
      return <input {...common} inputMode="numeric" pattern="[1-9][0-9]*"
        title="Solo números enteros positivos" maxLength={f.max ?? undefined}
        onChange={(e) => onChange(e.target.value.replace(/\D/g, ''))} />;
    case 'nt': // máx. 10 dígitos, empieza por 3
      return <input {...common} inputMode="numeric" pattern="3[0-9]{0,9}"
        title="Máximo 10 dígitos y debe empezar por 3" maxLength={10}
        onChange={(e) => onChange(e.target.value.replace(/\D/g, ''))} />;
    case 'c':
      return <input {...common} type="email" inputMode="email" autoComplete="email"
        onChange={(e) => onChange(e.target.value.trim())} />;
    case 'f':
    case 'dma': return <input {...common} type="date" onChange={(e) => onChange(e.target.value)} />;
    case 'fh': return <input {...common} type="datetime-local" onChange={(e) => onChange(e.target.value)} />;
    case 'h': return <input {...common} type="time" onChange={(e) => onChange(e.target.value)} />;
    case 'i':
      return (
        <>
          <input id={id} type="file" accept="image/*" aria-label={f.label} required={!value}
            onChange={async (e) => e.target.files[0] && onChange(await fileToPng(e.target.files[0]))} />
          {value && <span title="Imagen cargada"> ✓</span>}
        </>
      );
    default: return null;
  }
}

export default function FillForm({ file, onBack }) {
  const [buffer, setBuffer] = useState(null);
  const [fields, setFields] = useState(null);
  const [values, setValues] = useState({});   // campos sueltos
  const [tables, setTables] = useState({});   // { Grupo: [ { _k, 'Grupo.Col': valor } ] }
  const [fixed, setFixed] = useState({});     // { Grupo: [ { label, ids } ] } tablas de filas fijas
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const form = useRef();
  const keys = useRef(0);
  const isDocx = file.ext === 'docx';
  const base = file.name.replace(/\.[^.]+$/, '');

  useEffect(() => {
    (async () => {
      try {
        const buf = await file.getBuffer();
        setBuffer(buf);
        const { fields: fs, fixed: fx } = isDocx ? readDocxFields(buf) : await readXlsxFields(buf);
        setFields(fs);
        setFixed(fx);
        const groups = [...new Set(fs.filter((f) => f.group).map((f) => f.group))];
        // tabla fija: una fila de datos por cada fila del documento; si no, empieza con una
        setTables(Object.fromEntries(groups.map((g) => [g, (fx[g] ?? [null]).map(() => ({ _k: ++keys.current }))])));
      } catch (e) { setError('No se pudo leer el archivo: ' + e.message); }
    })();
  }, [file]);

  const singles = (fields ?? []).filter((f) => !f.group);
  const groups = Object.keys(tables);
  const columns = (g) => fields.filter((f) => f.group === g);

  const setCell = (g, i, id, v) =>
    setTables((s) => ({ ...s, [g]: s[g].map((r, j) => (j === i ? { ...r, [id]: v } : r)) }));
  const addRow = (g) => setTables((s) => ({ ...s, [g]: [...s[g], { _k: ++keys.current }] }));
  const removeRow = (g, i) => setTables((s) => ({ ...s, [g]: s[g].filter((_, j) => j !== i) }));

  async function run(kind) {
    if (!form.current.reportValidity()) return;
    const imgs = fields.filter((f) => f.type === 'i');
    const rowHas = (g, i, id) => !fixed[g] || fixed[g][i].ids.includes(id);
    const missing = imgs.find((f) => (f.group
      ? tables[f.group].some((r, i) => rowHas(f.group, i, f.id) && !r[f.id])
      : !values[f.id]));
    if (missing) return setError(`Falta una imagen en "${missing.label}"`);
    setError(''); setBusy(true);
    try {
      if (isDocx) {
        const blob = fillDocx(buffer, values, tables);
        if (kind === 'pdf') download(await docxToPdf(blob), base + '.pdf');
        else download(blob, base + '.docx');
      } else {
        download(await fillXlsx(buffer, values, tables), base + '.xlsx');
      }
    } catch (e) { setError('Error generando el archivo: ' + e.message); }
    setBusy(false);
  }

  return (
    <div>
      <button onClick={onBack}>← Volver</button>
      <h2>{file.name}</h2>
      {!fields && !error && <p>Leyendo documento…</p>}
      {fields && !fields.length && <p>Este archivo no tiene campos para llenar.</p>}
      {fields && (
        <form ref={form} className="form" onSubmit={(e) => e.preventDefault()}>
          {singles.map((f) => (
            <label key={f.id} htmlFor={f.id}>
              <span>{f.id}{f.max ? ` (máx. ${f.max})` : ''}</span>
              {renderInput(f, values[f.id], (v) => setValues((s) => ({ ...s, [f.id]: v })), f.id)}
            </label>
          ))}

          {groups.map((g) => {
            const fx = fixed[g]; // tabla de filas fijas (no se agregan ni quitan filas)
            return (
              <fieldset key={g} className="group">
                <legend>{g}</legend>
                <div className="scroll">
                  <table>
                    <thead>
                      <tr>
                        {fx && <th />}
                        {columns(g).map((c) => <th key={c.id}>{c.label}{c.max ? ` (máx. ${c.max})` : ''}</th>)}
                        {!fx && <th />}
                      </tr>
                    </thead>
                    <tbody>
                      {tables[g].map((row, i) => (
                        <tr key={row._k}>
                          {fx && <th scope="row" className="rowlabel">{fx[i].label || `Fila ${i + 1}`}</th>}
                          {columns(g).map((c) => (
                            <td key={c.id}>
                              {(!fx || fx[i].ids.includes(c.id)) &&
                                renderInput(c, row[c.id], (v) => setCell(g, i, c.id, v))}
                            </td>
                          ))}
                          {!fx && (
                            <td>
                              <button type="button" title="Quitar fila" disabled={tables[g].length === 1}
                                onClick={() => removeRow(g, i)}>✕</button>
                            </td>
                          )}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {!fx && <button type="button" onClick={() => addRow(g)}>+ Agregar fila</button>}
              </fieldset>
            );
          })}

          <div className="actions">
            <button disabled={busy} onClick={() => run('file')}>
              Descargar {isDocx ? 'Word' : 'Excel'}
            </button>
            {isDocx && <button disabled={busy} onClick={() => run('pdf')}>Descargar PDF</button>}
          </div>
        </form>
      )}
      {error && <p className="error">{error}</p>}
    </div>
  );
}