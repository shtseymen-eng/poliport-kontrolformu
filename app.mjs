import {transformUrl,identity,filename,folder,validateExcel,norm} from './core.mjs?v=20261109-large1';
import {readReportOffThread} from './report-reader.mjs?v=20261109-large1';
import {REPORT_LIMITS,MAX_FORM_BYTES,REPORT_FILE_LABEL} from './report-config.mjs';
import {openArchive,getForm,putForm,listForms} from './history.mjs';
import {keyHash,folderRecord,rememberFolder,knownFolderKeys} from './folder.mjs?v=20261109-large1';
import {writeExcelToFolder} from './downloads.mjs?v=20261108-excel2';
const $=id=>document.getElementById(id);
let folderKnown=new Set();
let records=[],selected=new Set(),history=new Map(),folderHandle=null,page=0,busy=false,stop=false,controller=null;
const PAGE_SIZE=100;
let lastSource='',currentReportFile=null;
function notice(message,error=false){$('notice').textContent=message;$('notice').hidden=false;$('notice').classList.toggle('error',error);}
function log(message,type=''){const p=document.createElement('p');p.textContent=new Date().toLocaleTimeString('tr-TR')+'  '+message;p.className=type;$('log').append(p);while($('log').children.length>100)$('log').firstChild.remove();$('log').scrollTop=$('log').scrollHeight;}
function known(key){return folderHandle?folderKnown.has(key):history.has(key);}
function status(record){return record.status||(known(record.key)?'Arşivde':!record.url?'Link yok':!record.timestamp?'Saat eksik':'Bekliyor');}
function controls(){for(const id of ['selectAll','selectNone','testOne','download'])$(id).disabled=busy||!records.length;for(const id of ['reportFile','browseReport','loadFile','loadUrl','pickFolder','hosts','reportUrl','fileTab','urlTab'])$(id).disabled=busy;$('stop').hidden=!busy;}
function stats(){const unique=new Map(records.filter(r=>r.key&&r.timestamp).map(r=>[r.key,r]));$('total').textContent=records.length;$('newCount').textContent=[...unique.keys()].filter(k=>!known(k)).length;$('duplicateCount').textContent=records.filter(r=>r.key&&known(r.key)).length;$('selectedCount').textContent=selected.size;$('historyCount').textContent=history.size;}
function render(){
 const query=norm($('search').value);const visible=records.filter(r=>(!$('onlyDownloaded').checked||known(r.key))&&(!query||norm([r.plaka,r.dorse,r.konteyner,r.ad,r.soyad,r.nakliyeci].join(' ')).includes(query)));
 const pages=Math.max(1,Math.ceil(visible.length/PAGE_SIZE));page=Math.min(page,pages-1);
 $('vehicles').replaceChildren();$('empty').hidden=visible.length>0;$('pagination').hidden=visible.length<=PAGE_SIZE;
 $('pageInfo').textContent=`${page+1} / ${pages} · ${visible.length} araç`;$('prev').disabled=page===0;$('next').disabled=page===pages-1;
 for(const rec of visible.slice(page*PAGE_SIZE,(page+1)*PAGE_SIZE)){
  const tr=document.createElement('tr');tr.className=known(rec.key)||['Kaydedildi','Arşivlendi'].includes(rec.status)?'downloaded':'pending';const td=document.createElement('td'),checkbox=document.createElement('input');checkbox.type='checkbox';checkbox.checked=selected.has(rec.row);checkbox.disabled=busy||!rec.url||!rec.timestamp||known(rec.key);checkbox.setAttribute('aria-label',rec.plaka+' aracını seç');checkbox.addEventListener('change',()=>{if(checkbox.checked)selected.add(rec.row);else selected.delete(rec.row);stats();});td.append(checkbox);tr.append(td);
  for(const value of [rec.tarih,rec.plaka,rec.dorse,rec.konteyner,rec.tip,rec.ad,rec.soyad,rec.nakliyeci]){const cell=document.createElement('td');cell.textContent=value||'—';tr.append(cell);}
  const state=status(rec),cell=document.createElement('td'),badge=document.createElement('span');badge.textContent=state;badge.className='badge'+(['Arşivlendi','Kaydedildi','Arşivde'].includes(state)?' ok':state==='Atlandı'?' skip':state==='Hata'?' error':'');cell.append(badge);tr.append(cell);
  const action=document.createElement('span');action.className='form-action';
  if(known(rec.key)){const btn=document.createElement('button');btn.textContent='Kopyayı indir';btn.className='text-button small';btn.disabled=busy;btn.addEventListener('click',()=>saveStored(rec.key));action.append(btn);}
  else if(rec.url){const a=document.createElement('a');a.textContent='Formu aç ↗';a.className='link';a.href=transformUrl(rec.url);a.target='_blank';a.rel='noopener noreferrer';action.append(a);}
  cell.append(action);$('vehicles').append(tr);
 }
 stats();controls();
}
function allowedHosts(){return new Set($('hosts').value.split(',').map(s=>s.trim().toLowerCase()).filter(Boolean));}
async function fetchExcel(raw,{report=false}={}){
 const byteLimit=report?REPORT_LIMITS.fileBytes:MAX_FORM_BYTES,sizeLabel=report?REPORT_FILE_LABEL:'25 MB';
 const url=new URL(raw);
 if(!['http:','https:'].includes(url.protocol)||url.username||url.password||!allowedHosts().has(url.host.toLowerCase()))throw new Error('Bağlantı sunucusu izinli değil. Kayıt noktası bölümündeki şirket rapor sunucularını kontrol edin.');
 controller=new AbortController();const timer=setTimeout(()=>controller?.abort(),report?300000:90000);
 try{
  const response=await fetch(url,{credentials:'include',redirect:'error',cache:'no-store',signal:controller.signal,targetAddressSpace:['127.0.0.1','localhost','[::1]'].includes(url.hostname)?'loopback':'local'});
  if(response.status===401||response.status===403)throw new Error('Şirket oturumu doğrulanamadı. Önce Formu aç bağlantısından şirket oturumunuzu açın.');
  if(!response.ok)throw new Error('Rapor sunucusu HTTP '+response.status+' yanıtı verdi.');
  if(Number(response.headers.get('Content-Length'))>byteLimit)throw new Error('Excel dosyası '+sizeLabel+' sınırını aştı.');
  const reader=response.body.getReader();const chunks=[];let size=0;
  while(true){const {done,value}=await reader.read();if(done)break;size+=value.length;if(size>byteLimit){await reader.cancel();throw new Error('Excel dosyası '+sizeLabel+' sınırını aştı.');}chunks.push(value);}
  return new Blob(chunks,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
 }catch(error){
  if(error.name==='TypeError')throw new Error('Şirket sunucusuna ağ erişimi engellendi. Şirket ağı / VPN bağlantısını kontrol edin. BT ekibinin bu sayfa için CORS ve yerel ağ erişimi tanımlaması gerekebilir.');
  if(error.name==='AbortError')throw new Error(stop?'İşlem durduruldu.':(report?'Rapor sunucusu 5 dakikada yanıt vermedi.':'Sunucu 90 saniyede yanıt vermedi.')+' Tekrar deneyebilirsiniz.');
  throw error;
 }finally{clearTimeout(timer);controller=null;}
}
async function loadReport(data,label){
 if(busy)return;
 busy=true;stop=false;controls();const readController=new AbortController();controller=readController;$('progress').removeAttribute('value');$('progressText').textContent='Rapor arka planda okunuyor…';
 try{
  const parsed=await readReportOffThread(data,{signal:readController.signal,onProgress:p=>{$('progressText').textContent=p.text;if(p.phase==='rows'&&p.totalBytes)$('progress').value=Math.min(90,Math.round(p.bytes/p.totalBytes*90));if(p.phase==='links')$('progress').value=95;if(p.phase==='complete')$('progress').value=100;}});
  const checked=folderHandle?await knownFolderKeys(folderHandle,parsed,{signal:readController.signal,onProgress:p=>{$('progressText').textContent=`Arşiv kontrol ediliyor · ${p.completed.toLocaleString('tr-TR')} / ${p.total.toLocaleString('tr-TR')} form`;}}):null;if(readController.signal.aborted)throw new DOMException('Rapor okuma durduruldu. Önceki liste korundu.','AbortError');
  records=parsed;if(checked)folderKnown=checked;if(!$('hosts').value.trim()){$('hosts').value=[...new Set(records.filter(r=>r.url).map(r=>new URL(r.url).host))].join(', ');}selected=new Set();page=0;lastSource=label;$('onlyDownloaded').checked=false;
  for(const rec of records)if(rec.key&&rec.timestamp&&!known(rec.key))selected.add(rec.row);
  $('reportSummary').textContent=label+' · '+records.length+' Terminal aracı';
  $('fileName').value=label;
  if(!records.length)notice('Raporda Terminal departmanına ait araç bulunamadı.');
  else notice(`${records.length} araç yüklendi. ${records.filter(r=>r.url&&!r.timestamp).length} kayıtta geçerli tarih/saat yok; bu kayıtlar otomatik indirilmeyecek.`);
  $('progressText').textContent=records.length.toLocaleString('tr-TR')+' Terminal aracı yüklendi.';log(label+' yüklendi.');return true;
 }catch(error){notice(error.message||'Rapor okunamadı.',error.name!=='AbortError');$('progressText').textContent=error.name==='AbortError'?'Rapor okuma durduruldu.':'Rapor yüklenemedi.';log('Rapor yüklenemedi: '+error.message,'error');return false;}
 finally{busy=false;if(controller===readController)controller=null;if(!$('progress').hasAttribute('value'))$('progress').value=0;render();}
}
async function fileChosen(file){if(!file)return;if(!/\.xlsx$/i.test(file.name)||file.size>REPORT_LIMITS.fileBytes){notice('En fazla 512 MB büyüklüğünde bir .xlsx raporu seçin.',true);return;}if(await loadReport(await file.arrayBuffer(),file.name))currentReportFile=file;}
function saveBlob(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);}
async function saveStored(key){try{const stored=await getForm(key);if(!stored?.blob)throw new Error('Saklanan belge bulunamadı.');saveBlob(stored.blob,records.find(r=>r.key===key)?filename(records.find(r=>r.key===key)):stored.name);}catch(error){notice(error.message,true);}}
async function chooseFolder(){
 if(!window.showDirectoryPicker)throw new Error('Excel dosyalarını doğrudan klasöre kaydetmek için güncel Chrome veya Edge kullanın.');
 const chosen=await window.showDirectoryPicker({mode:'readwrite',id:'poliport-arsiv'});
 folderHandle=chosen;folderKnown.clear();
 for(const rec of records){if(rec.key){const saved=await folderRecord(folderHandle,rec.key);rec.status='';if(saved){folderKnown.add(rec.key);history.set(rec.key,saved);selected.delete(rec.row);}else if(rec.timestamp)selected.add(rec.row);}}
 $('folderLabel').value=folderHandle.name;
 $('folderHint').textContent='Her kontrol formu ayrı .xlsx dosyası olarak yıl / ay / gün klasörüne kaydedilir.';
 render();
}
async function download(test=false){
 if(busy)return;
 const chosen=records.filter(r=>selected.has(r.row)&&r.url&&r.timestamp);if(test)chosen.splice(1);
 if(!chosen.length){notice('İndirilecek yeni form yok. Tümünü seç düğmesini kullanabilirsiniz.');return;}
 if(!navigator.locks){notice('Tekrar kayıt koruması için güncel Chrome veya Edge kullanın.',true);return;}
 try{if(!folderHandle)await chooseFolder();if(await folderHandle.requestPermission({mode:'readwrite'})!=='granted')throw new Error('Kayıt klasörüne yazma izni verilmedi.');}catch(error){if(error.name!=='AbortError')notice(error.message,true);return;}
 busy=true;stop=false;$('onlyDownloaded').checked=false;render();
 let success=0,skipped=0,failed=0;
 for(let i=0;i<chosen.length;i++){
  if(stop)break;const rec=chosen[i];
  $('progressText').textContent=`${i+1} / ${chosen.length} · ${rec.plaka}`;rec.status='İndiriliyor';render();
  try{
   const hash=await keyHash(rec.key);
   await navigator.locks.request('poliport-'+hash,async()=>{
    const existing=await folderRecord(folderHandle,rec.key);
    if(existing?.blob){history.set(rec.key,existing);if(folderHandle)folderKnown.add(rec.key);rec.status='Atlandı';selected.delete(rec.row);skipped++;log(rec.plaka+' · Aynı saat ve form daha önce arşivlendi.');return;}
    const cached=await getForm(rec.key);
    const blob=cached?.blob||await fetchExcel(transformUrl(rec.url));await validateExcel(await blob.arrayBuffer(),window.JSZip);
    if(stop)throw new Error('İşlem durduruldu.');
    const {name,path}=await writeExcelToFolder(folderHandle,rec,blob);
    const stored={key:rec.key,plaka:rec.plaka,timestamp:rec.timestamp,name,path,blob,savedAt:new Date().toISOString()};
    if(folderHandle)await rememberFolder(folderHandle,stored);
    await putForm(stored);history.set(rec.key,stored);if(folderHandle)folderKnown.add(rec.key);selected.delete(rec.row);
    success++;rec.status='Kaydedildi';log(rec.plaka+' · '+rec.status+' · '+rec.tarih,'ok');
   });
  }catch(error){failed++;rec.status='Hata';log(rec.plaka+' · '+error.message,'error');notice(error.message,true);}
  $('progress').value=Math.round((i+1)/chosen.length*100);render();
 }
 $('progressText').textContent=`${success} Excel kaydedildi · ${skipped} tekrar atlandı · ${failed} hata${stop?' · İşlem durdu':''}`;
 log($('progressText').textContent);if(success||skipped)$('onlyDownloaded').checked=true;busy=false;render();
}
$('fileTab').addEventListener('click',()=>switchTab(false));$('urlTab').addEventListener('click',()=>switchTab(true));
function switchTab(url){$('fileTab').setAttribute('aria-selected',String(!url));$('urlTab').setAttribute('aria-selected',String(url));$('filePanel').hidden=url;$('urlPanel').hidden=!url;}
$('browseReport').addEventListener('click',()=>$('reportFile').click());
$('loadFile').addEventListener('click',()=>{if(currentReportFile)fileChosen(currentReportFile);else if(lastSource==='Bağlantıdan alınan rapor'&&$('reportUrl').value)$('loadUrl').click();else $('reportFile').click();});
$('reportFile').addEventListener('change',event=>fileChosen(event.target.files[0]));
$('dropzone').addEventListener('dragover',event=>{event.preventDefault();$('dropzone').classList.add('dragging');});
$('dropzone').addEventListener('dragleave',()=>$('dropzone').classList.remove('dragging'));
$('dropzone').addEventListener('drop',event=>{event.preventDefault();$('dropzone').classList.remove('dragging');if(!busy)fileChosen(event.dataTransfer.files[0]);});
$('loadUrl').addEventListener('click',async()=>{
 if(busy)return;
 try{busy=true;stop=false;controls();$('progressText').textContent='Rapor bağlantısından dosya alınıyor…';const blob=await fetchExcel(transformUrl($('reportUrl').value),{report:true});busy=false;if(await loadReport(await blob.arrayBuffer(),'Bağlantıdan alınan rapor'))currentReportFile=null;}
 catch(error){notice(error.message,true);busy=false;controls();}
});
$('pickFolder').addEventListener('click',async()=>{try{await chooseFolder();}catch(error){if(error.name!=='AbortError')notice(error.message,true);}});
$('selectAll').addEventListener('click',()=>{$('onlyDownloaded').checked=false;selected=new Set(records.filter(r=>r.key&&r.timestamp&&!known(r.key)).map(r=>r.row));render();});
$('selectNone').addEventListener('click',()=>{selected.clear();render();});
$('download').addEventListener('click',()=>download());$('testOne').addEventListener('click',()=>download(true));
$('stop').addEventListener('click',()=>{stop=true;controller?.abort();});
$('search').addEventListener('input',()=>{page=0;render();});
$('prev').addEventListener('click',()=>{page--;render();});$('next').addEventListener('click',()=>{page++;render();});
$('onlyDownloaded').addEventListener('change',()=>{page=0;render();});
$('historyButton').addEventListener('click',async()=>{
 try{const forms=await listForms();$('historyList').replaceChildren();
  if(!forms.length){const p=document.createElement('p');p.textContent='Henüz arşivlenmiş belge yok.';$('historyList').append(p);}
  for(const form of forms.sort((a,b)=>b.savedAt.localeCompare(a.savedAt))){const row=document.createElement('div');row.className='history-row';const info=document.createElement('div'),strong=document.createElement('strong'),p=document.createElement('p'),btn=document.createElement('button');strong.textContent=form.plaka+' · '+form.timestamp.replace('T',' ');p.textContent=form.name;info.append(strong,p);btn.textContent='Belgeyi indir';btn.className='secondary small';btn.addEventListener('click',()=>saveStored(form.key));row.append(info,btn);$('historyList').append(row);}
  $('historyDialog').showModal();
 }catch(error){notice(error.message,true);}
});
$('closeHistory').addEventListener('click',()=>$('historyDialog').close());
$('hosts').addEventListener('change',()=>{try{localStorage.setItem('poliport-hosts',$('hosts').value);}catch{}});
try{const saved=localStorage.getItem('poliport-hosts');if(saved)$('hosts').value=saved;await openArchive();for(const form of await listForms())history.set(form.key,form);}catch(error){notice(error.message,true);}
render();
