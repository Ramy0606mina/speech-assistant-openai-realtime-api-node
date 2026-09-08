import { deflateRawSync } from 'node:zlib';
import { parseReport, reportSubject, escapeHtml, isMoney } from './report-format.js';

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) { crc ^= byte; for (let i = 0; i < 8; i++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0); }
  return (crc ^ 0xffffffff) >>> 0;
}
// Minimal standards-compliant ZIP container; no installed Word application needed.
function zip(files) {
  const local = [], central = []; let offset = 0;
  for (const [path, text] of Object.entries(files)) {
    const name = Buffer.from(path), bytes = Buffer.from(text), packed = deflateRawSync(bytes), crc = crc32(bytes);
    const header = Buffer.alloc(30); header.writeUInt32LE(0x04034b50); header.writeUInt16LE(20,4); header.writeUInt16LE(8,8);
    header.writeUInt32LE(crc,14); header.writeUInt32LE(packed.length,18); header.writeUInt32LE(bytes.length,22); header.writeUInt16LE(name.length,26);
    local.push(header,name,packed);
    const entry=Buffer.alloc(46); entry.writeUInt32LE(0x02014b50); entry.writeUInt16LE(20,4); entry.writeUInt16LE(20,6); entry.writeUInt16LE(8,10);
    entry.writeUInt32LE(crc,16); entry.writeUInt32LE(packed.length,20); entry.writeUInt32LE(bytes.length,24); entry.writeUInt16LE(name.length,28); entry.writeUInt32LE(offset,42);
    central.push(entry,name); offset += header.length + name.length + packed.length;
  }
  const directory=Buffer.concat(central), end=Buffer.alloc(22), count=Object.keys(files).length;
  end.writeUInt32LE(0x06054b50); end.writeUInt16LE(count,8); end.writeUInt16LE(count,10); end.writeUInt32LE(directory.length,12); end.writeUInt32LE(offset,16);
  return Buffer.concat([...local,directory,end]);
}

function p(text, {bold=false, size=22, align='left', keep=false}={}) {
  return `<w:p><w:pPr><w:spacing w:after="100"/>${keep?'<w:keepNext/>':''}<w:jc w:val="${align}"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="${size}"/>${bold?'<w:b/>':''}</w:rPr><w:t xml:space="preserve">${escapeHtml(text).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g,'')}</w:t></w:r></w:p>`;
}

export function renderReportDocx({subject,text}) {
  const body = [p(reportSubject(subject),{bold:true,size:36,keep:true})];
  for (const b of parseReport(text)) {
    if(b.type!=='table'){body.push(p(b.text,{bold:b.type==='heading',size:b.type==='heading'?28:22,keep:b.type==='heading'}));continue;}
    const width=Math.floor(10080/b.columns.length);
    const rows=[b.columns,...b.rows].map((row,index)=>`<w:tr><w:trPr><w:cantSplit/>${index===0?'<w:tblHeader/>':''}</w:trPr>${row.map(value=>`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:shd w:fill="${index===0?'E5EAF0':index%2?'FFFFFF':'F5F7F9'}"/></w:tcPr>${p(value,{bold:index===0,size:21,align:isMoney(value)?'right':'left'})}</w:tc>`).join('')}</w:tr>`).join('');
    body.push(`<w:tbl><w:tblPr><w:tblW w:w="10080" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>${['top','left','bottom','right','insideH','insideV'].map(edge=>`<w:${edge} w:val="single" w:sz="4" w:color="CBD5E1"/>`).join('')}</w:tblBorders><w:tblCellMar><w:top w:w="90" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:left w:w="110" w:type="dxa"/><w:right w:w="110" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${b.columns.map(()=>`<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>${p('')}`);
  }
  return zip({
    '[Content_Types].xml':'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    '_rels/.rels':'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/document.xml':`<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body.join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/></w:sectPr></w:body></w:document>`
  });
}
