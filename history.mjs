let opening;
export function openArchive() {
  if(opening)return opening;
  opening=new Promise((resolve,reject)=>{
    const request=indexedDB.open('poliport-kontrol-arsivi',1);
    request.onupgradeneeded=()=>request.result.createObjectStore('forms',{keyPath:'key'});
    request.onerror=()=>reject(new Error('Tarayıcı arşivi açılamadı. Normal bir Chrome veya Edge penceresinde deneyin.'));
    request.onsuccess=()=>resolve(request.result);
  });
  return opening;
}
async function store(mode,action) {
  const db=await openArchive();
  return new Promise((resolve,reject)=>{
    const tx=db.transaction('forms',mode);let result;
    const req=action(tx.objectStore('forms'));
    req.onsuccess=()=>{result=req.result;};
    tx.oncomplete=()=>resolve(result);
    tx.onabort=tx.onerror=()=>reject(new Error('Form tarayıcı arşivine yazılamadı. Boş disk alanını kontrol edin; kayıt yeniden denenebilir.'));
  });
}
export const getForm=key=>store('readonly',s=>s.get(key));
export const putForm=form=>store('readwrite',s=>s.put(form));
export const listForms=()=>store('readonly',s=>s.getAll());
