import PDFDocument from 'pdfkit';

export function renderReportPdf({ subject, text }) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: 'LETTER', margin: 54,
      info: { Title: String(subject || 'London report'), Author: 'London Assistant' } });
    const chunks = [];
    doc.on('data', chunk => chunks.push(chunk));
    doc.on('error', reject);
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.fillColor('#18364a').font('Helvetica-Bold').fontSize(20).text(String(subject || 'London report'));
    doc.moveDown(0.5).font('Helvetica').fontSize(9).fillColor('#526573').text('LONDON | MINACO');
    doc.moveDown(1).fillColor('#202b33');
    for (const line of String(text || '').split(/\r?\n/)) {
      const heading = /^#{1,6}\s|^\*\*[^*]+\*\*\s*$/.test(line);
      const clean = line.replace(/^#{1,6}\s+/, '').replace(/\*\*([^*]+)\*\*/g, '$1').replace(/`([^`]+)`/g, '$1').replace(/[\u2010-\u2015]/g, '-');
      if (!clean.trim()) { doc.moveDown(0.45); continue; }
      doc.font(heading ? 'Helvetica-Bold' : 'Helvetica').fontSize(heading ? 12 : 10.5)
        .text(clean, { width: 504, lineGap: 3 });
    }
    doc.end();
  });
}
