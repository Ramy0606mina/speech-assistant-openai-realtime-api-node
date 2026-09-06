import PDFDocument from 'pdfkit';

const BRIEF_SECTIONS=[
  ['priorities','Top priorities',['priority','item','hint','status','due']],
  ['calendar','Today\'s calendar',['time','item','hint']],
  ['emails','Emails to answer',['from','item','hint','status']],
  ['tasks','Tasks and follow-ups',['item','hint','status','due']],
  ['risks','Risks and blockers',['item','hint','status']],
  ['completed','Completed since yesterday',['item','hint','status']],
];

function statusColors(value){const s=String(value||'').toLowerCase();if(s.includes('overdue'))return ['#fde8e7','#a3261d'];if(s.includes('complete')||s.includes('done'))return ['#e3f4e8','#216e39'];if(s.includes('upcoming'))return ['#e7f0fb','#1f5c99'];return ['#fff2cc','#7a5500'];}

function ensureRoom(doc,height){if(doc.y+height>doc.page.height-doc.page.margins.bottom)doc.addPage();}

function drawBrief(doc,reportData){
  doc.rect(0,0,doc.page.width,84).fill('#18364a');
  doc.fillColor('#ffffff').font('Helvetica-Bold').fontSize(19).text('London — Morning Executive Report',54,28,{width:504});
  doc.fillColor('#d8e6ef').font('Helvetica').fontSize(9).text(String(reportData.date||''),54,56);
  doc.y=104;
  for(const [key,label,columns] of BRIEF_SECTIONS){
    const rows=reportData.sections?.[key]||[];if(!rows.length)continue;
    const widths=columns.map((_,index)=>index===columns.indexOf('hint')?190:Math.floor((504-190)/(columns.length-1||1)));
    const tableHeight=28+rows.length*32+26;ensureRoom(doc,tableHeight);
    doc.fillColor('#18364a').font('Helvetica-Bold').fontSize(12).text(label,54,doc.y);doc.moveDown(.55);
    const top=doc.y;let x=54;
    columns.forEach((column,index)=>{doc.rect(x,top,widths[index],24).fill('#eaf0f4');doc.fillColor('#18364a').font('Helvetica-Bold').fontSize(7.5).text(column==='item'?'ITEM':column.toUpperCase(),x+5,top+8,{width:widths[index]-10,height:10});x+=widths[index];});
    let y=top+24;
    for(const [rowIndex,row] of rows.entries()){
      x=54;doc.rect(54,y,504,32).fill(rowIndex%2?'#f8fafb':'#ffffff');
      columns.forEach((column,index)=>{const value=String(row[column]||'—');if(column==='status'&&value!=='—'){const [bg,fg]=statusColors(value);doc.roundedRect(x+4,y+7,widths[index]-8,18,8).fill(bg);doc.fillColor(fg).font('Helvetica-Bold').fontSize(7).text(value,x+7,y+12,{width:widths[index]-14,height:8,align:'center'});}else doc.fillColor('#202b33').font(column==='item'?'Helvetica-Bold':'Helvetica').fontSize(7.5).text(value,x+5,y+6,{width:widths[index]-10,height:22,ellipsis:true});x+=widths[index];});
      doc.moveTo(54,y+32).lineTo(558,y+32).strokeColor('#e2e8ec').lineWidth(.5).stroke();y+=32;
    }
    doc.y=y+14;
  }
  ensureRoom(doc,30);doc.fillColor('#647681').font('Helvetica').fontSize(8).text('Reply to the report or tell London by phone to update an action: done, waiting, deferred, cancelled, or still pending.',54,doc.y,{width:504});
}

export function renderReportPdf({ subject, text, reportData }) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: 'LETTER', margin: 54,
      info: { Title: String(subject || 'London report'), Author: 'London Assistant' } });
    const chunks = [];
    doc.on('data', chunk => chunks.push(chunk));
    doc.on('error', reject);
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    if(reportData){drawBrief(doc,reportData);doc.end();return;}
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

