import { renderAsync } from 'docx-preview';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';

// Word -> HTML (docx-preview) -> imagen por página -> PDF (jsPDF)
export async function docxToPdf(blob) {
  const host = document.createElement('div');
  host.style.cssText = 'position:fixed;left:-10000px;top:0;';
  document.body.appendChild(host);
  try {
    await renderAsync(blob, host, undefined, { inWrapper: false, breakPages: true });
    const pages = [...host.querySelectorAll('section.docx')];
    let pdf;
    for (const page of pages) {
      const canvas = await html2canvas(page, { scale: 2, useCORS: true });
      const w = page.offsetWidth, h = page.offsetHeight;
      const orient = w > h ? 'l' : 'p';
      if (!pdf) pdf = new jsPDF({ unit: 'px', format: [w, h], orientation: orient, hotfixes: ['px_scaling'] });
      else pdf.addPage([w, h], orient);
      pdf.addImage(canvas.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, w, h);
    }
    return pdf.output('blob');
  } finally {
    host.remove();
  }
}
