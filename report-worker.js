// Dedicated, per-import worker: decompresses XML incrementally without blocking the page.
importScripts('./jszip.min.js');
self.onmessage=async event=>{
 try{
  const {readReport}=await import('./core.mjs?v=20261109-large1');
  const records=await readReport(event.data.buffer,self.JSZip,{onProgress:progress=>self.postMessage({type:'progress',progress})});
  self.postMessage({type:'complete',records});
 }catch(error){self.postMessage({type:'error',message:error.message||'Rapor okunamadı.'});}
};
