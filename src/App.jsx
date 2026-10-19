import { useState } from 'react';
import { buildTreeFromZip } from './lib/fileTree';
import FileBrowser from './components/FileBrowser';
import FillForm from './components/FillForm';
import './App.css';

export default function App() {
  const [tree, setTree] = useState(null);
  const [zipName, setZipName] = useState('');
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  async function loadZip(f) {
    if (!f) return;
    setError(''); setLoading(true);
    try {
      const { root, count } = await buildTreeFromZip(f);
      if (!count) throw new Error('El zip no contiene archivos .docx ni .xlsx');
      setTree(root); setZipName(f.name); setFile(null);
    } catch (e) {
      setError('No se pudo leer el zip: ' + e.message);
    }
    setLoading(false);
  }

  return (
    <main>
      <h1>Documentos</h1>

      {!tree && (
        <label
          className="drop"
          onDragOver={(e) => e.preventDefault()}
          onDrop={(e) => { e.preventDefault(); loadZip(e.dataTransfer.files[0]); }}
        >
          <strong>{loading ? 'Leyendo zip…' : 'Selecciona o arrastra el archivo .zip'}</strong>
          <span>Con la estructura de carpetas y los Word / Excel</span>
          <input type="file" accept=".zip,application/zip" hidden disabled={loading}
            onChange={(e) => loadZip(e.target.files[0])} />
        </label>
      )}
      {error && <p className="error">{error}</p>}

      {tree && !file && (
        <>
          <p className="zipbar">
            📦 {zipName}{' '}
            <label className="link">
              Cambiar zip
              <input type="file" accept=".zip,application/zip" hidden
                onChange={(e) => loadZip(e.target.files[0])} />
            </label>
          </p>
          <FileBrowser tree={tree} onOpen={setFile} />
        </>
      )}
      {tree && file && <FillForm key={file.path} file={file} onBack={() => setFile(null)} />}
    </main>
  );
}