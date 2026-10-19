import { useState } from 'react';

export default function FileBrowser({ tree, onOpen }) {
  const [path, setPath] = useState([]); // ['Contratos', '2024']
  const node = path.reduce((n, p) => n.folders[p], tree);

  return (
    <div>
      <nav className="crumbs">
        <button onClick={() => setPath([])}>files</button>
        {path.map((p, i) => (
          <span key={i}> / <button onClick={() => setPath(path.slice(0, i + 1))}>{p}</button></span>
        ))}
      </nav>
      <ul className="list">
        {Object.keys(node.folders).sort().map((f) => (
          <li key={f}><button onClick={() => setPath([...path, f])}>📁 {f}</button></li>
        ))}
        {node.files.sort((a, b) => a.name.localeCompare(b.name)).map((f) => (
          <li key={f.name}>
            <button onClick={() => onOpen(f)}>{f.ext === 'docx' ? '📄' : '📊'} {f.name}</button>
          </li>
        ))}
        {!Object.keys(node.folders).length && !node.files.length && <li>Carpeta vacía</li>}
      </ul>
    </div>
  );
}
