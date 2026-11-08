const MONTHS=['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran','Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık'];
export const norm = value => String(value ?? '').replace(/\s+/g,' ').trim().toLocaleLowerCase('tr-TR');
export function transformUrl(raw) {
  let decoded=String(raw || '').trim();
  // SSRS encodes report paths and parameter names in its hyperlink.
  decoded=decoded.replace(/%2f/ig,'/').replace(/%3a/ig,':');
  decoded=decoded.replace(/\/ReportServer\/Pages\/ReportViewer\.aspx\?/i,'/ReportServer?');
  const url=new URL(decoded);
  if(!['http:','https:'].includes(url.protocol)||url.username||url.password) throw new Error('Geçersiz rapor bağlantısı.');
  const query=url.search.slice(1).split('&').filter(part=>part&&!/^(rs:|rc:)/i.test(decodeURIComponent(part.split('=')[0])));
  query.push('rs:Format=EXCELOPENXML');
  url.search='?'+query.join('&');
  return url.href;
}
export function identity(record) {
  const url=new URL(transformUrl(record.url));
  const params=[...url.searchParams].filter(([k])=>!(/^(rs:|rc:)/i.test(k))).sort(([a,av],[b,bv])=>a.localeCompare(b)||av.localeCompare(bv));
  return JSON.stringify([url.origin.toLowerCase(),url.pathname,params,String(record.plaka).replace(/\s/g,'').toUpperCase(),record.timestamp || record.tarih || '']);
}
export function clean(value,spaces=false) {
  let text=String(value || '').toLocaleUpperCase('tr-TR').replace(/[ĞÜŞİÖÇ]/g,c=>({'Ğ':'G','Ü':'U','Ş':'S','İ':'I','Ö':'O','Ç':'C'}[c]));
  text=text.replace(/[\\/*?:"<>|\x00-\x1f]/g,'');
  return (spaces?text.replace(/\s+/g,' '):text.replace(/\s/g,'')).replace(/^[. ]+|[. ]+$/g,'');
}
export function filename(r) {
  const p=clean(r.plaka), d=clean(r.dorse);
  return [p,d!==p?d:'',clean(r.konteyner),clean(r.ad,true),clean(r.soyad,true)].filter(Boolean).join(' ').slice(0,175)+'.xlsx';
}
export function folder(r) {
  if(!r.timestamp) return 'Kontrol Formları/Tarihsiz';
  const [y,m,d]=r.timestamp.slice(0,10).split('-').map(Number);
  return `Kontrol Formları/${y}/${MONTHS[m-1]}/${d}`;
}
export function parseDate(value,numeric=false,date1904=false) {
  let parts;
  if(numeric && Number.isFinite(Number(value))) {
    const date=new Date(Math.round((Number(value)-(date1904?24107:25569))*86400000));
    parts=[date.getUTCFullYear(),date.getUTCMonth()+1,date.getUTCDate(),date.getUTCHours(),date.getUTCMinutes(),date.getUTCSeconds()];
  } else {
    const text=String(value ?? '').trim();
    let m=text.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if(m) parts=[+m[3],+m[2],+m[1],+(m[4]||0),+(m[5]||0),+(m[6]||0)];
    else { m=text.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?$/); if(m) parts=m.slice(1).map(v=>+(v||0)); }
  }
  if(!parts) return {timestamp:'',display:String(value ?? '')};
  const [y,m,d,h,min,s]=parts;
  const date=new Date(Date.UTC(y,m-1,d,h,min,s));
  if(date.getUTCFullYear()!==y||date.getUTCMonth()!==m-1||date.getUTCDate()!==d||h>23||min>59||s>59) return {timestamp:'',display:String(value ?? '')};
  const pad=n=>String(n).padStart(2,'0');
  return {timestamp:`${y}-${pad(m)}-${pad(d)}T${pad(h)}:${pad(min)}:${pad(s)}`,display:`${pad(d)}.${pad(m)}.${y} ${pad(h)}:${pad(min)}:${pad(s)}`};
}
const elements=(node,name)=>Array.from(node.getElementsByTagNameNS('*',name));
function xml(text) {
  const doc=new DOMParser().parseFromString(text,'application/xml');
  if(elements(doc,'parsererror').length) throw new Error('Excel içeriği okunamadı.');
  return doc;
}
function zipPath(base,target) {
  const parts=(target.startsWith('/')?target.slice(1):base+'/'+target).split('/'), result=[];
  for(const p of parts) {if(p==='..') result.pop();else if(p&&p!=='.')result.push(p);}
  return result.join('/');
}
async function entry(zip,path,required=true) {
  const file=zip.file(path);
  if(!file) {if(required)throw new Error('Geçerli bir .xlsx raporu seçin.'); return null;}
  // Limit decompressed XML size to avoid accidental oversized reports.
  if(file._data?.uncompressedSize>50*1024*1024)throw new Error('Excel sayfası çok büyük. Daha küçük bir rapor yükleyin.');
  return xml(await file.async('string'));
}
export async function readReport(data,JSZip) {
  const zip=await JSZip.loadAsync(data);
  const workbook=await entry(zip,'xl/workbook.xml');
  const rels=await entry(zip,'xl/_rels/workbook.xml.rels');
  const sheets=elements(workbook,'sheet');
  const view=elements(workbook,'workbookView')[0];
  const sheet=sheets[Number(view?.getAttribute('activeTab')||0)]||sheets[0];
  if(!sheet)throw new Error('Excel sayfası bulunamadı.');
  const rel=elements(rels,'Relationship').find(r=>r.getAttribute('Id')===sheet.getAttribute('r:id'));
  if(!rel)throw new Error('Excel sayfası bulunamadı.');
  const path=zipPath('xl',rel.getAttribute('Target'));
  const ws=await entry(zip,path);
  const strings=await entry(zip,'xl/sharedStrings.xml',false);
  const shared=strings?elements(strings,'si').map(si=>elements(si,'t').map(t=>t.textContent).join('')):[];
  const parts=path.split('/');const name=parts.pop();const base=parts.join('/');
  const linksDoc=await entry(zip,base+'/_rels/'+name+'.rels',false);
  const linkMap=new Map(linksDoc?elements(linksDoc,'Relationship').map(r=>[r.getAttribute('Id'),r.getAttribute('Target')]):[]);
  const links=new Map(elements(ws,'hyperlink').map(l=>[l.getAttribute('ref'),linkMap.get(l.getAttribute('r:id'))||'']));
  const rows=elements(ws,'row').map(row=>{
    const values=[];
    for(const c of elements(row,'c')) {
      const ref=c.getAttribute('r');const letters=ref.replace(/\d/g,'');let index=0;
      for(const l of letters)index=index*26+l.charCodeAt(0)-64;
      const type=c.getAttribute('t'),raw=elements(c,'v')[0]?.textContent||'';
      const value=type==='s'?shared[Number(raw)]||'':type==='inlineStr'?elements(c,'t').map(t=>t.textContent).join(''):raw;
      values[index-1]={value,ref,numeric:!type||type==='n',link:links.get(ref)||''};
    }
    return {number:Number(row.getAttribute('r')),values};
  });
  const headerIndex=rows.findIndex(row=>row.number<=20&&row.values.some(c=>c&&['kayıt no','plaka'].includes(norm(c.value))));
  if(headerIndex<0)throw new Error('Başlık bulunamadı. Kayıt No veya Plaka sütunu gerekli.');
  const header=rows[headerIndex].values;
  function find(candidates,fallback) {
    for(const candidate of candidates) {const exact=header.findIndex(c=>c&&norm(c.value)===norm(candidate));if(exact>=0)return exact;}
    for(const candidate of candidates) {const partial=header.findIndex(c=>c&&norm(c.value).includes(norm(candidate)));if(partial>=0)return partial;}
    return fallback;
  }
  const cols={tarih:find(['Kayıt Tarihi','Tarih'],1),departman:find(['Departman'],5),plaka:find(['Plaka','Çekici Plaka'],6),dorse:find(['Dorse Plaka','Dorse'],8),konteyner:find(['ISO Tank No','ISO Tank','Isotank No','Isotank','Konteyner No','Konteyner'],9),tip:find(['Araç Tipi','Araç Tip'],12),ad:find(['Sürücü Adı','Sürücü Ad'],14),soyad:find(['Sürücü Soyadı','Soyadı'],15),nakliyeci:find(['Nakliyeci'],16)};
  const date1904=['1','true'].includes(elements(workbook,'workbookPr')[0]?.getAttribute('date1904'));
  const records=[];
  for(const row of rows.slice(headerIndex+1)) {
    const get=key=>String(row.values[cols[key]]?.value??'').trim();
    if(!norm(get('departman')).includes('terminal')||!get('plaka'))continue;
    const dateCell=row.values[cols.tarih];const parsed=parseDate(dateCell?.value,dateCell?.numeric,date1904);
    const rec=Object.fromEntries(Object.keys(cols).map(key=>[key,get(key)]));
    rec.tarih=parsed.display;rec.timestamp=parsed.timestamp;rec.url=row.values[cols.plaka]?.link||'';
    try{rec.key=rec.url?identity(rec):'';}catch{rec.url='';rec.key='';}
    rec.row=row.number;records.push(rec);
  }
  return records;
}
export async function validateExcel(bytes,JSZip) {
  const zip=await JSZip.loadAsync(bytes);
  if(!zip.file('xl/workbook.xml'))throw new Error('Sunucu Excel yerine farklı bir yanıt döndürdü. Şirket oturumunuzu kontrol edin.');
}
