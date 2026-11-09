import sax from './vendor/sax.mjs';
import {norm,parseDate,identity} from './core.mjs?v=20261109-large1';
import {REPORT_LIMITS} from './report-config.mjs';
const local=name=>name.split(':').at(-1);
const rid=attributes=>attributes['r:id']||Object.entries(attributes).find(([key])=>key.endsWith(':id'))?.[1]||'';
function zipPath(base,target){const out=[];for(const p of (target.startsWith('/')?target.slice(1):base+'/'+target).split('/')){if(p==='..')out.pop();else if(p&&p!=='.')out.push(p);}return out.join('/');}
function column(ref){const m=/^([A-Z]{1,3})[1-9]\d{0,6}$/.exec(ref||'');if(!m)throw new Error('Excel hücre adresi okunamadı.');let n=0;for(const c of m[1])n=n*26+c.charCodeAt(0)-64;if(n>16384)throw new Error('Excel sütun sınırı aşıldı.');return n-1;}
async function streamXml(zip,path,limit,handlers={},required=true,onBytes=()=>{}){
 const file=zip.file(path);if(!file){if(required)throw new Error('Geçerli bir .xlsx raporu seçin.');return false;}
 const oversized=()=>new Error('Excel sayfası / bileşeni için okuma sınırı aşıldı. Bu rapor desteklenen kapasiteden büyük.');
 if(file._data?.uncompressedSize>limit)throw oversized();
 const parser=sax.parser(true,{position:false,strictEntities:true});let depth=0,bytes=0,finished=false;
 parser.onopentag=node=>{if(++depth>64)throw new Error('Excel XML yapısı okunamadı.');handlers.open?.(local(node.name),node.attributes);};
 parser.onclosetag=name=>{handlers.close?.(local(name));depth--;};
 parser.ontext=text=>handlers.text?.(text);parser.oncdata=parser.ontext;
 parser.ondoctype=()=>{throw new Error('Excel XML içeriğinde desteklenmeyen belge tanımı var.');};
 const decoder=new TextDecoder('utf-8',{fatal:true});
 return new Promise((resolve,reject)=>{
  const stream=file.internalStream('uint8array');
  const fail=error=>{if(finished)return;finished=true;stream.pause();reject(error);};
  stream.on('data',chunk=>{if(finished)return;try{bytes+=chunk.byteLength;if(bytes>limit)throw oversized();parser.write(decoder.decode(chunk,{stream:true}));onBytes(bytes);}catch(error){fail(error.message.startsWith('Excel')?error:new Error('Excel XML içeriği okunamadı. Dosyanın eksiksiz bir .xlsx raporu olduğunu kontrol edin.'));}});
  stream.on('error',()=>fail(new Error('Excel sıkıştırılmış içeriği okunamadı. Dosyayı yeniden dışa aktarın.')));
  stream.on('end',()=>{if(finished)return;try{parser.write(decoder.decode()).close();finished=true;resolve(true);}catch{fail(new Error('Excel XML içeriği okunamadı. Dosyayı yeniden dışa aktarın.'));}});
  stream.resume();
 });
}
function columns(header){
 const find=(candidates,fallback)=>{for(const candidate of candidates){const i=header.findIndex(c=>c&&norm(c.value)===norm(candidate));if(i>=0)return i;}for(const candidate of candidates){const i=header.findIndex(c=>c&&norm(c.value).includes(norm(candidate)));if(i>=0)return i;}return fallback;};
 return {tarih:find(['Kayıt Tarihi','Tarih'],1),departman:find(['Departman'],5),plaka:find(['Plaka','Çekici Plaka'],6),dorse:find(['Dorse Plaka','Dorse'],8),konteyner:find(['ISO Tank No','ISO Tank','Isotank No','Isotank','Konteyner No','Konteyner'],9),tip:find(['Araç Tipi','Araç Tip'],12),ad:find(['Sürücü Adı','Sürücü Ad'],14),soyad:find(['Sürücü Soyadı','Soyadı'],15),nakliyeci:find(['Nakliyeci'],16)};
}
export async function readReportLarge(data,JSZip,{limits:overrides={},onProgress=()=>{}}={}){
 const limits={...REPORT_LIMITS,...overrides};
 if((data.byteLength??data.size??0)>limits.fileBytes)throw new Error('Rapor dosyası en fazla 512 MB olabilir.');
 const zip=await JSZip.loadAsync(data);if(Object.keys(zip.files).length>limits.entries)throw new Error('Excel dosyasında çok fazla bileşen var.');
 let date1904=false,active=0,sawView=false;const sheets=[];
 onProgress({phase:'metadata',rows:0,text:'Excel çalışma kitabı açılıyor…'});
 await streamXml(zip,'xl/workbook.xml',limits.metadataBytes,{open:(name,a)=>{if(name==='workbookPr')date1904=['1','true'].includes(a.date1904);if(name==='workbookView'&&!sawView){active=Number(a.activeTab)||0;sawView=true;}if(name==='sheet')sheets.push({rid:rid(a)});}});
 const sheet=sheets[active]||sheets[0];if(!sheet)throw new Error('Excel sayfası bulunamadı.');let target='';
 await streamXml(zip,'xl/_rels/workbook.xml.rels',limits.metadataBytes,{open:(name,a)=>{if(name==='Relationship'&&a.Id===sheet.rid)target=a.Target;}});
 if(!target)throw new Error('Excel sayfası bulunamadı.');
 const path=zipPath('xl',target),shared=[];let inString=false,inText=false,string='';
 onProgress({phase:'strings',rows:0,text:'Excel metinleri okunuyor…'});
 await streamXml(zip,'xl/sharedStrings.xml',limits.sharedBytes,{open:name=>{if(name==='si'){inString=true;string='';}if(name==='t'&&inString)inText=true;},text:text=>{if(inText){string+=text;if(string.length>65536)throw new Error('Excel hücresindeki metin sınırı aşıldı.');}},close:name=>{if(name==='t')inText=false;if(name==='si'){shared.push(string);inString=false;}}},false);
 const records=[],refs=new Map(),relationRecords=new Map();let values=[],rowNumber=0,cell=null,cols=null,needed=null,rows=0,inValue=false,inInline=false,lastProgress=0;
 const progress=bytes=>{const now=Date.now();if(now-lastProgress>180){lastProgress=now;onProgress({phase:'rows',rows,matches:records.length,bytes,totalBytes:zip.file(path)?._data?.uncompressedSize||0,text:`${rows.toLocaleString('tr-TR')} satır okundu · ${records.length.toLocaleString('tr-TR')} Terminal aracı`});}};
 await streamXml(zip,path,limits.worksheetBytes,{
  open:(name,a)=>{
   if(name==='row'){rowNumber=a.r?Number(a.r):rowNumber+1;if(!Number.isSafeInteger(rowNumber)||rowNumber<1||rowNumber>limits.rows)throw new Error('Excel satır sınırı aşıldı.');values=[];}
   if(name==='c'){const index=column(a.r);cell=(!needed||needed.has(index))?{index,ref:a.r,type:a.t||'',raw:'',inline:''}:null;}
   if(name==='v')inValue=true;if(name==='t')inInline=true;
   if(name==='hyperlink'){
    const rec=refs.get(a.ref),id=rid(a);if(rec&&id){const matches=relationRecords.get(id)||[];matches.push(rec);relationRecords.set(id,matches);}
   }
  },
  text:text=>{if(!cell)return;if(inValue)cell.raw+=text;else if(inInline&&cell.type==='inlineStr')cell.inline+=text;if(cell.raw.length+cell.inline.length>65536)throw new Error('Excel hücresindeki metin sınırı aşıldı.');},
  close:name=>{
   if(name==='v')inValue=false;if(name==='t')inInline=false;
   if(name==='c'&&cell){values[cell.index]={value:cell.type==='s'?shared[Number(cell.raw)]||'':cell.type==='inlineStr'?cell.inline:cell.raw,ref:cell.ref,numeric:!cell.type||cell.type==='n'};cell=null;}
   if(name==='row'){
    rows++;
    if(!cols){if(rowNumber<=20&&values.some(c=>c&&['kayıt no','plaka'].includes(norm(c.value)))){cols=columns(values);needed=new Set(Object.values(cols));}return;}
    const get=key=>String(values[cols[key]]?.value??'').trim();if(!norm(get('departman')).includes('terminal')||!get('plaka'))return;
    const date=values[cols.tarih],parsed=parseDate(date?.value,date?.numeric,date1904),rec=Object.fromEntries(Object.keys(cols).map(key=>[key,get(key)]));
    rec.tarih=parsed.display;rec.timestamp=parsed.timestamp;rec.url='';rec.key='';rec.row=rowNumber;records.push(rec);const plate=values[cols.plaka]?.ref;if(plate)refs.set(plate,rec);
   }
  }
 },true,progress);
 if(!cols)throw new Error('Başlık bulunamadı. Kayıt No veya Plaka sütunu gerekli.');
 onProgress({phase:'links',rows,matches:records.length,text:`${rows.toLocaleString('tr-TR')} satır okundu; kontrol formu bağlantıları eşleştiriliyor…`});
 const parts=path.split('/'),name=parts.pop();
 await streamXml(zip,parts.join('/')+'/_rels/'+name+'.rels',limits.worksheetBytes,{open:(name,a)=>{if(name!=='Relationship')return;for(const rec of relationRecords.get(a.Id)||[]){try{rec.url=a.Target||'';rec.key=rec.url?identity(rec):'';}catch{rec.url='';rec.key='';}}}},false);
 onProgress({phase:'complete',rows,matches:records.length,text:`${rows.toLocaleString('tr-TR')} satır tamamlandı · ${records.length.toLocaleString('tr-TR')} Terminal aracı`});return records;
}
