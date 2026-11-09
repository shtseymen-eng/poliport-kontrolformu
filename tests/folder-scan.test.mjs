import test from 'node:test';
import assert from 'node:assert/strict';
import {knownFolderKeys} from '../folder.mjs';
test('large fresh destination skips per-record file checks when no archive index exists',async()=>{
 let calls=0;const root={async getDirectoryHandle(){calls++;throw new DOMException('missing','NotFoundError');}};
 const result=await knownFolderKeys(root,Array.from({length:10000},(_,i)=>({key:String(i)})));
 assert.equal(result.size,0);assert.equal(calls,1);
});
test('archive scan deduplicates keys, bounds concurrency and returns only verified forms',async()=>{
 let active=0,peak=0,calls=0;const root={async getDirectoryHandle(){return {};}};
 const found=await knownFolderKeys(root,Array.from({length:40},(_,i)=>({key:String(i%20)})),{read:async(_root,key)=>{calls++;peak=Math.max(peak,++active);await new Promise(r=>setTimeout(r,1));active--;return key==='5'?{}:null;}});
 assert.equal(calls,20);assert.ok(peak<=8);assert.deepEqual([...found],['5']);
});
test('stopping during folder verification rejects without a partial result',async()=>{
 const controller=new AbortController(),root={async getDirectoryHandle(){controller.abort();return {};}};
 await assert.rejects(()=>knownFolderKeys(root,[{key:'a'}],{signal:controller.signal}),e=>e.name==='AbortError');
});
