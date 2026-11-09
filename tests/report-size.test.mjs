import test from 'node:test';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import {readReport} from '../core.mjs';
const require=createRequire(import.meta.url);
const JSZip=require('jszip');
const escape=s=>String(s).replaceAll('&','&amp;').replaceAll('<','&lt;');
const cell=(ref,value)=>`<c r="${ref}" t="inlineStr"><is><t>${escape(value)}</t></is></c>`;
function workbook(sheet){const z=new JSZip();z.file('xl/workbook.xml','<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Rapor" sheetId="1" r:id="r1"/></sheets></workbook>');z.file('xl/_rels/workbook.xml.rels','<Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/></Relationships>');z.file('xl/worksheets/sheet1.xml',sheet);return z;}
const columns=['Kayıt No','Kayıt Tarihi','Departman','Plaka','ISO Tank No','Sürücü Adı','Sürücü Soyadı','Açıklama'];
const header='<row r="1">'+columns.map((name,i)=>cell(String.fromCharCode(65+i)+'1',name)).join('')+'</row>';
let big;
async function monthly(){if(big)return big;big=(async()=>{
 const rows=[header],links=[],rels=[];const noise='A'.repeat(1900);
 for(let i=0;i<31000;i++){const r=i+2,day=Math.floor(i/1000)+1,terminal=i%10===0;
 rows.push(`<row r="${r}">`+[String(i+1),`${String(day).padStart(2,'0')}.10.2026 10:30:15`,terminal?'Terminal':'Diğer','34 TEST 01','ISO1234567','ALİ','ŞEN',noise].map((v,c)=>cell(String.fromCharCode(65+c)+r,v)).join('')+'</row>');
 links.push(`<hyperlink ref="D${r}" r:id="h${r}"/>`);rels.push(`<Relationship Id="h${r}" Target="https://reports.example/ReportServer?/Form&amp;TransactionID=${i}" TargetMode="External"/>`);
 }
 const sheet='<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>'+rows.join('')+'</sheetData><hyperlinks>'+links.join('')+'</hyperlinks></worksheet>';
 assert.ok(Buffer.byteLength(sheet)>50*1024*1024);
 const z=workbook(sheet);z.file('xl/worksheets/_rels/sheet1.xml.rels','<Relationships>'+rels.join('')+'</Relationships>');
 return z.generateAsync({type:'nodebuffer',compression:'DEFLATE',compressionOptions:{level:1}});
})();return big;}
test('31-day worksheet above former 50 MB XML ceiling loads every day without dropping Terminal rows',async()=>{
 const progress=[];const started=Date.now();const rows=await readReport(await monthly(),JSZip,{onProgress:p=>progress.push(p)});
 assert.equal(rows.length,3100);assert.equal(new Set(rows.map(r=>r.timestamp.slice(0,10))).size,31);
 assert.equal(rows[0].ad,'ALİ');assert.equal(rows[0].soyad,'ŞEN');assert.equal(rows.at(-1).timestamp,'2026-10-31T10:30:15');
 assert.equal(rows[0].url,'https://reports.example/ReportServer?/Form&TransactionID=0');assert.equal(new Set(rows.map(r=>r.key)).size,3100);
 assert.ok(progress.some(p=>p.rows>=31000));console.log('Monthly workbook parsed in',Date.now()-started,'ms');
});
test('streaming reader respects actual worksheet byte limit and reports useful errors',async()=>{
 const sheet='<worksheet><sheetData>'+header+'<row r="2">'+cell('H2','X'.repeat(3000))+'</row></sheetData></worksheet>';
 const bytes=await workbook(sheet).generateAsync({type:'nodebuffer'});
 await assert.rejects(()=>readReport(bytes,JSZip,{limits:{worksheetBytes:1024}}),/sayfa.*sınır|sınır.*sayfa/i);
});
test('malformed XML is rejected rather than reporting a partial success',async()=>{
 const bytes=await workbook('<worksheet><sheetData>'+header+'</wrong></worksheet>').generateAsync({type:'nodebuffer'});
 await assert.rejects(()=>readReport(bytes,JSZip),/Excel.*okunamadı|XML/i);
});
test('shared strings, rich text, numeric 1904 dates and hyperlink identity survive streaming',async()=>{
 const z=workbook('<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>'+header+'<row r="2"><c r="B2"><v>0.5</v></c><c r="C2" t="s"><v>0</v></c><c r="D2" t="s"><v>1</v></c><c r="F2" t="s"><v>2</v></c>'+cell('G2','ŞEN')+'</row></sheetData><hyperlinks><hyperlink ref="D2" r:id="link1"/></hyperlinks></worksheet>');
 z.file('xl/workbook.xml',(await z.file('xl/workbook.xml').async('string')).replace('<bookViews>','<workbookPr date1904="1"/><bookViews>'));
 z.file('xl/sharedStrings.xml','<sst><si><t>Terminal</t></si><si><t>34 ABC 01</t></si><si><r><t>ALİ </t></r><r><t>VELİ</t></r></si></sst>');
 z.file('xl/worksheets/_rels/sheet1.xml.rels','<Relationships><Relationship Id="link1" Target="https://reports.example/ReportServer?/Form&amp;TransactionID=1"/></Relationships>');
 const rows=await readReport(await z.generateAsync({type:'nodebuffer'}),JSZip);
 assert.equal(rows.length,1);assert.equal(rows[0].ad,'ALİ VELİ');assert.equal(rows[0].timestamp,'1904-01-01T12:00:00');assert.ok(rows[0].key.includes('34ABC01'));
});
