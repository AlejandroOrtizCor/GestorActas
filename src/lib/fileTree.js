import JSZip from 'jszip';

// Lee un .zip y arma el árbol de carpetas con los .docx y .xlsx que contiene.
// Los archivos no se descomprimen hasta que se abren (getBuffer).
export async function buildTreeFromZip(zipFile) {
  const zip = await JSZip.loadAsync(zipFile);

  const entries = [];
  zip.forEach((path, entry) => {
    if (entry.dir) return;
    const parts = path.split('/');
    const name = parts[parts.length - 1];
    if (parts[0] === '__MACOSX' || name.startsWith('._') || name.startsWith('~$')) return;
    const ext = name.includes('.') ? name.split('.').pop().toLowerCase() : '';
    if (ext !== 'docx' && ext !== 'xlsx') return;
    entries.push({ parts, entry, name, ext });
  });

  // Si todo el zip viene dentro de una única carpeta raíz (ej. "files/"), se omite
  if (entries.length && entries.every((e) => e.parts.length > 1 && e.parts[0] === entries[0].parts[0])) {
    entries.forEach((e) => e.parts.shift());
  }

  const root = { name: 'files', folders: {}, files: [] };
  for (const { parts, entry, name, ext } of entries) {
    let node = root;
    for (const p of parts.slice(0, -1)) node = node.folders[p] ??= { name: p, folders: {}, files: [] };
    node.files.push({ name, ext, path: parts.join('/'), getBuffer: () => entry.async('arraybuffer') });
  }
  return { root, count: entries.length };
}