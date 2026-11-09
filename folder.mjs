export async function keyHash(key){return [...new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(key)))].map(n=>n.toString(16).padStart(2,'0')).join('');}
export async function folderRecord(root,key){
 try{
  const index=await root.getDirectoryHandle('.poliport-kayitlar');
  const handle=await index.getFileHandle(await keyHash(key)+'.json');
  const metadata=JSON.parse(await (await handle.getFile()).text());
  if(metadata.key!==key||typeof metadata.path!=='string')return null;
  const parts=metadata.path.split('/');
  if(parts.some(p=>!p||p==='.'||p==='..'))return null;
  const name=parts.pop();let dir=root;
  for(const part of parts)dir=await dir.getDirectoryHandle(part);
  const blob=await (await dir.getFileHandle(name)).getFile();
  if(!blob.size)return null;
  return {...metadata,blob};
 }catch(error){if(error.name==='NotFoundError'||error instanceof SyntaxError)return null;throw error;}
}
export async function rememberFolder(root,form){
 const index=await root.getDirectoryHandle('.poliport-kayitlar',{create:true});
 const handle=await index.getFileHandle(await keyHash(form.key)+'.json',{create:true});
 const writer=await handle.createWritable();
 const {blob,...metadata}=form;
 try{await writer.write(JSON.stringify(metadata));await writer.close();}
 catch(error){try{await writer.abort();}catch{}throw error;}
}
// Check only distinct forms, without serial filesystem reads for every monthly-report row.
export async function knownFolderKeys(root,records,{signal,onProgress=()=>{},read=folderRecord}={}){
 const abort=()=>{if(signal?.aborted)throw new DOMException('Rapor okuma durduruldu. Önceki liste korundu.','AbortError');};
 abort();try{await root.getDirectoryHandle('.poliport-kayitlar');}catch(error){if(error.name==='NotFoundError')return new Set();throw error;}abort();
 const keys=[...new Set(records.map(r=>r.key).filter(Boolean))],found=new Set();let cursor=0,completed=0,failed=false;
 async function scan(){try{while(cursor<keys.length&&!failed){abort();const key=keys[cursor++],saved=await read(root,key);abort();if(failed)return;if(saved)found.add(key);completed++;if(completed%100===0||completed===keys.length)onProgress({completed,total:keys.length});}}catch(error){failed=true;throw error;}}
 await Promise.all(Array.from({length:Math.min(8,keys.length)},scan));abort();return found;
}
