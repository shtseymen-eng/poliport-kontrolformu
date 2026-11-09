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
export async function readReport(data,JSZip,options={}) {
  const {readReportLarge}=await import('./large-report.mjs?v=20261109-large1');
  return readReportLarge(data,JSZip,options);
}
export async function validateExcel(bytes,JSZip) {
  const zip=await JSZip.loadAsync(bytes);
  if(!zip.file('xl/workbook.xml'))throw new Error('Sunucu Excel yerine farklı bir yanıt döndürdü. Şirket oturumunuzu kontrol edin.');
}
