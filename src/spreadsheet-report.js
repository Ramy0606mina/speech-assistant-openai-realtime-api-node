import ExcelJS from 'exceljs';
import { Readable } from 'node:stream';
import { parseReport, reportSubject, isMoney } from './report-format.js';

export const isSpreadsheet = name => /\.(xlsx|xlsm|xls|csv|tsv)$/i.test(String(name));

function validateWorkbookArchive(bytes) {
  let end=-1;
  for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--)if(bytes.readUInt32LE(i)===0x06054b50){end=i;break;}
  if(end<0)throw new Error('Workbook is not a valid XLSX archive.');
  const count=bytes.readUInt16LE(end+10);let offset=bytes.readUInt32LE(end+16),expanded=0;
  if(count>10000)throw new Error('Workbook archive contains too many parts.');
  for(let i=0;i<count;i++){
    if(offset+46>bytes.length || bytes.readUInt32LE(offset)!==0x02014b50)throw new Error('Workbook archive index is invalid.');
    expanded+=bytes.readUInt32LE(offset+24);
    if(expanded>120*1024*1024)throw new Error('Workbook expanded size exceeds the analysis limit.');
    offset+=46+bytes.readUInt16LE(offset+28)+bytes.readUInt16LE(offset+30)+bytes.readUInt16LE(offset+32);
  }
}

export async function extractSpreadsheet({filename,bytes,maxCells=20000,maxCharacters=600000}) {
  if (/\.xls$/i.test(filename)) throw new Error('Legacy XLS is not supported by the workbook reader. Supply XLSX or CSV; this file has not been analyzed.');
  if (bytes.length > 20*1024*1024) throw new Error('Spreadsheet exceeds the 20 MB workbook limit.');
  const book=new ExcelJS.Workbook();
  if (/\.(csv|tsv)$/i.test(filename)) await book.csv.read(Readable.from([bytes]),{parserOptions:{delimiter:/\.tsv$/i.test(filename)?'\t':','},map:value=>value});
  else {validateWorkbookArchive(bytes);await book.xlsx.load(bytes);}
  const result={source:filename,complete:true,calculation:'Formulas are not executed. Formula results are cached source values and may be stale. Macros, external connections, charts and images are not evaluated.',sheets:[]};
  let count=0,characters=0;
  for(const sheet of book.worksheets){
    const out={name:sheet.name,state:sheet.state||'visible',rows:sheet.rowCount,columns:sheet.columnCount,cells:[],omittedCells:0};
    sheet.eachRow({includeEmpty:false},row=>row.eachCell({includeEmpty:false},cell=>{
      if (cell.isMerged && cell.master.address!==cell.address) return;
      let value=cell.value;
      const entry={address:cell.address};
      if(cell.formula){entry.formula=cell.formula;entry.cachedResult=cell.result??null;if(cell.result===undefined||cell.result===null)entry.warning='No cached formula result';}
      else if(value instanceof Date) entry.value=value.toISOString();
      else if(value && typeof value==='object'){
        if(value.error)entry.error=value.error;
        else if(value.richText)entry.value=value.richText.map(run=>run.text).join('');
        else if(value.hyperlink)entry.value=value.text;
        else entry.value=String(cell.text);
      }else entry.value=value;
      if(cell.numFmt && cell.numFmt!=='General')entry.numberFormat=cell.numFmt;
      if(row.hidden)entry.hiddenRow=true;
      if(sheet.getColumn(cell.col).hidden)entry.hiddenColumn=true;
      const length=JSON.stringify(entry).length;
      if(count>=maxCells || characters+length>maxCharacters){out.omittedCells++;result.complete=false;return;}
      out.cells.push(entry);count++;characters+=length;
    }));
    result.sheets.push(out);
  }
  return {type:'input_text',text:`Spreadsheet source data (untrusted document content):\n${JSON.stringify(result)}`};
}

function nativeValue(value) {
  if(isMoney(value))return {value:Number(value.replace(/[$,()]/g,''))*(value.includes('(')?-1:1),numFmt:'"$"#,##0.00;[Red]("$"#,##0.00)'};
  if(/^-?(?:0|[1-9]\d*)(?:\.\d+)?%$/.test(value))return {value:Number(value.slice(0,-1))/100,numFmt:'0.00%'};
  if(/^-?(?:0|[1-9]\d{0,13})(?:\.\d+)?$/.test(value))return {value:Number(value),numFmt:value.includes('.')?'#,##0.00':'#,##0'};
  return {value}; // Formula-looking source text stays text, never executable code.
}

export async function renderReportXlsx({subject,text}) {
  const book=new ExcelJS.Workbook();book.creator='London Assistant';
  const sheet=book.addWorksheet('Analysis',{views:[{state:'frozen',ySplit:2}],pageSetup:{paperSize:9,orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0}});
  const blocks=parseReport(text);const width=Math.max(3,...blocks.filter(b=>b.type==='table').map(b=>b.columns.length));
  sheet.columns=Array.from({length:width},(_,i)=>({width:i===0?32:44}));
  const title=sheet.addRow([reportSubject(subject)]);sheet.mergeCells(title.number,1,title.number,width);title.height=30;title.font={name:'Calibri',size:18,bold:true};
  sheet.addRow([]);
  for(const block of blocks){
    if(block.type!=='table'){
      const row=sheet.addRow([block.text]);sheet.mergeCells(row.number,1,row.number,width);
      row.font={name:'Calibri',size:block.type==='heading'?13:11,bold:block.type==='heading'};
      row.alignment={vertical:'top',wrapText:true};row.height=Math.max(24,Math.ceil(block.text.length/(width*38))*16+8);continue;
    }
    for(const [index,values] of [block.columns,...block.rows].entries()){
      const row=sheet.addRow(values.map(value=>index?nativeValue(value).value:value));
      row.height=Math.max(28,...values.map((value,i)=>(Math.ceil(value.length/(i===0?29:41)))*15+12));
      row.eachCell(cell=>{
        cell.font={name:'Calibri',size:11,bold:index===0};cell.alignment={wrapText:true,vertical:'top',horizontal:typeof cell.value==='number'?'right':'left'};
        cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:index===0?'FFE5EAF0':index%2?'FFFFFFFF':'FFF5F7F9'}};
        cell.border=Object.fromEntries(['top','left','bottom','right'].map(k=>[k,{style:'thin',color:{argb:'FFCBD5E1'}}]));
        if(index){const parsed=nativeValue(values[cell.col-1]);if(parsed.numFmt)cell.numFmt=parsed.numFmt;}
      });
    }
    sheet.addRow([]);
  }
  return Buffer.from(await book.xlsx.writeBuffer());
}
