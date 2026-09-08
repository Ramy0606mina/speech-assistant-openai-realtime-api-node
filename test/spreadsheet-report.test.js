import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { extractSpreadsheet, renderReportXlsx } from '../src/spreadsheet-report.js';
import { DropboxClient } from '../src/dropbox-client.js';
import { MicrosoftGraphClient } from '../src/microsoft-graph.js';

async function fixture(){
 const book=new ExcelJS.Workbook();const s=book.addWorksheet('Costs');
 s.addRow(['Item','Amount']);s.addRow(['First',100]);s.addRow(['Second',200]);s.getCell('B4').value={formula:'SUM(B2:B3)',result:300};
 s.getCell('B5').value={formula:'B2-B2',result:0};s.getCell('B6').value={formula:'1/0',result:{error:'#DIV/0!'}};
 s.getCell('B7').value={formula:'B2*2'};s.getCell('B2').numFmt='$0.00';s.getRow(3).hidden=true;
 const h=book.addWorksheet('Assumptions',{state:'hidden'});h.getCell('A1').value='Tax';h.getCell('B1').value=0.15;
 return Buffer.from(await book.xlsx.writeBuffer());
}
const data=part=>JSON.parse(part.text.slice(part.text.indexOf('\n')+1));

test('all workbook sheets, hidden data, formulas, cached zero and errors reach analysis',async()=>{
 const result=data(await extractSpreadsheet({filename:'costs.xlsx',bytes:await fixture()}));
 assert.equal(result.complete,true);assert.equal(result.sheets.length,2);assert.equal(result.sheets[1].state,'hidden');
 const cells=result.sheets[0].cells;assert.equal(cells.find(c=>c.address==='B4').formula,'SUM(B2:B3)');
 assert.equal(cells.find(c=>c.address==='B4').cachedResult,300);assert.equal(cells.find(c=>c.address==='B5').cachedResult,0);
 assert.deepEqual(cells.find(c=>c.address==='B6').cachedResult,{error:'#DIV/0!'});
 assert.match(cells.find(c=>c.address==='B7').warning,/No cached/);assert.equal(cells.find(c=>c.address==='B3').hiddenRow,true);
 assert.match(result.calculation,/not executed/);
});
test('limits explicitly report omitted cells while retaining the sheet inventory',async()=>{
 const result=data(await extractSpreadsheet({filename:'costs.xlsx',bytes:await fixture(),maxCells:2}));
 assert.equal(result.complete,false);assert.equal(result.sheets.length,2);assert.ok(result.sheets[0].omittedCells>0);assert.ok(result.sheets[1].omittedCells>0);
});
test('CSV quoting and TSV are parsed and legacy XLS never claims successful analysis',async()=>{
 const csv=data(await extractSpreadsheet({filename:'source.csv',bytes:Buffer.from('Item,Value\n"Line, with comma",0012\n')}));
 assert.equal(csv.sheets[0].cells.find(c=>c.address==='A2').value,'Line, with comma');assert.equal(csv.sheets[0].cells.find(c=>c.address==='B2').value,'0012');
 const tsv=data(await extractSpreadsheet({filename:'source.tsv',bytes:Buffer.from('Item\tValue\nTest\t2')}));assert.equal(tsv.sheets[0].columns,2);
 await assert.rejects(extractSpreadsheet({filename:'old.xls',bytes:Buffer.from('old')}),/not supported/);
});
test('Excel analysis is editable, preserves prose and numeric types, and never executes source strings',async()=>{
 const bytes=await renderReportXlsx({subject:'Cost analysis',text:'# Findings\nReview the assumptions.\n\n| Item | Value |\n|---|---|\n| Cost | $1,250.50 |\n| Tax | 15% |\n| Text | =HYPERLINK("bad") |'});
 const book=new ExcelJS.Workbook();await book.xlsx.load(bytes);const s=book.worksheets[0];
 const values=[];s.eachRow(row=>row.eachCell(cell=>values.push(cell.value)));
 assert.ok(values.includes(1250.5));assert.ok(values.includes(.15));assert.ok(values.includes('=HYPERLINK("bad")'));assert.ok(values.includes('Review the assumptions.'));
 assert.equal(s.views[0].state,'frozen');assert.equal(s.getColumn(1).width,32);
});
test('Dropbox workbook reader exposes structured source and requested analysis saves a real XLSX',async()=>{
 const bytes=await fixture();const uploads=[];
 const dbx=new DropboxClient({accessToken:'test',saveReports:true,fetchImpl:async(url,opts)=>{
  if(String(url).endsWith('/download'))return new Response(bytes);
  if(String(url).endsWith('create_folder_v2'))return Response.json({id:'folder'});
  const path=JSON.parse(opts.headers['Dropbox-API-Arg']).path;uploads.push({path,bytes:opts.body});return Response.json({id:'saved',path_display:path});
 }});
 const source=await dbx.readFile('costs.xlsx');assert.equal(data(source.part).sheets.length,2);
 const report=await dbx.saveReport({taskKey:'excel',subject:'Analyze costs',text:'Analysis complete.\n\n| Item | Cost |\n|---|---|\n| First | $100.00 |',includeXlsx:true});
 assert.equal(report.attachments.length,3);assert.ok(report.xlsxPath.endsWith('.xlsx'));
 const workbook=new ExcelJS.Workbook();await workbook.xlsx.load(uploads.find(u=>u.path.endsWith('.xlsx')).bytes);assert.equal(workbook.worksheets[0].name,'Analysis');
});
test('email workbook attachments use the same sheet and formula extraction',async()=>{
 const bytes=await fixture();
 const graph=new MicrosoftGraphClient({readTenantId:'t',readClientId:'c',readClientSecret:'test',ramyMailbox:'owner@example.com',londonMailbox:'london@example.com',fetchImpl:async url=>String(url).includes('login.microsoftonline.com')?Response.json({access_token:'test',expires_in:3600}):Response.json({value:[{'@odata.type':'#microsoft.graph.fileAttachment',name:'costs.xlsx',contentBytes:bytes.toString('base64')}]})});
 const parts=await graph.getLondonAttachments('message');assert.equal(parts[0].type,'input_text');assert.equal(data(parts[0]).sheets.length,2);
});
test('invalid and over-expanded workbook archives are rejected before loading',async()=>{
 await assert.rejects(extractSpreadsheet({filename:'bad.xlsx',bytes:Buffer.from('invalid')}),/valid XLSX/);
 const bytes=await fixture();const offset=bytes.indexOf(Buffer.from([0x50,0x4b,0x01,0x02]));bytes.writeUInt32LE(121*1024*1024,offset+24);
 await assert.rejects(extractSpreadsheet({filename:'large.xlsx',bytes}),/expanded size/);
});
