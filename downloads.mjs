import {filename,folder} from './core.mjs?v=20261108-excel2';

// Write each verified form directly to its own Excel file.
export async function writeExcelToFolder(root,record,blob){
  const name=filename(record),parts=folder(record).split('/');let directory=root;
  for(const part of parts)directory=await directory.getDirectoryHandle(part,{create:true});
  let candidate=name;
  for(let number=0;number<10000;number++){
    try{await directory.getFileHandle(candidate);candidate=name.slice(0,-5)+` (${number+1}).xlsx`;}
    catch(error){if(error.name==='NotFoundError')break;throw error;}
    if(number===9999)throw new Error('Aynı isimli çok fazla belge var.');
  }
  const handle=await directory.getFileHandle(candidate,{create:true});
  const writable=await handle.createWritable();
  try{await writable.write(blob);await writable.close();}
  catch(error){try{await writable.abort();}catch{}throw error;}
  return {name:candidate,path:parts.join('/')+'/'+candidate};
}
