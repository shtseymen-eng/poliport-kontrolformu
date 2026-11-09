import {REPORT_LIMITS} from './report-config.mjs';
export function readReportOffThread(data,{signal,onProgress=()=>{},createWorker=()=>new Worker(new URL('./report-worker.js?v=20261109-large1',import.meta.url))}={}){
 return new Promise((resolve,reject)=>{
  if(signal?.aborted){reject(new DOMException('Rapor okuma durduruldu.','AbortError'));return;}
  if(!(data instanceof ArrayBuffer)){reject(new Error('Rapor dosyası okunamadı.'));return;}
  if(data.byteLength>REPORT_LIMITS.fileBytes){reject(new Error('Rapor dosyası en fazla 512 MB olabilir.'));return;}
  let worker,settled=false;
  const cleanup=()=>{signal?.removeEventListener('abort',abort);worker?.terminate();};
  const finish=(error,value)=>{if(settled)return;settled=true;cleanup();if(error)reject(error);else resolve(value);};
  const abort=()=>finish(new DOMException('Rapor okuma durduruldu. Önceki liste korundu.','AbortError'));
  try{
   worker=createWorker();worker.onmessage=event=>{
    const data=event.data;if(data.type==='progress'){onProgress(data.progress);return;}
    if(data.type==='complete')finish(null,data.records);
    if(data.type==='error')finish(new Error(data.message||'Rapor okunamadı.'));
   };
   worker.onerror=event=>{event.preventDefault?.();finish(new Error('Rapor okunamadı. Tarayıcının bellek sınırına ulaşılmış olabilir. Gereksiz sekmeleri kapatıp yeniden deneyin.'));};
   worker.onmessageerror=()=>finish(new Error('Rapor sonucu tarayıcıya aktarılamadı.'));
   signal?.addEventListener('abort',abort,{once:true});worker.postMessage({buffer:data},[data]);
  }catch{finish(new Error('Rapor okuyucusu başlatılamadı. Güncel Chrome veya Edge ile yeniden deneyin.'));}
 });
}
