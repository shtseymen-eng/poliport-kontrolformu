import test from 'node:test';
import assert from 'node:assert/strict';
import {readReportOffThread} from '../report-reader.mjs';
import {REPORT_LIMITS} from '../report-config.mjs';
function fake(){return {terminated:0,terminate(){this.terminated++;},postMessage(message,transfer){this.message=message;this.transfer=transfer;}};}
test('report file capacity is 512 MiB and parsing transfers the buffer instead of copying',async()=>{
 assert.equal(REPORT_LIMITS.fileBytes,512*1024*1024);const worker=fake(),data=new ArrayBuffer(8),progress=[];
 const result=readReportOffThread(data,{createWorker:()=>worker,onProgress:p=>progress.push(p)});
 assert.equal(worker.transfer[0],data);worker.onmessage({data:{type:'progress',progress:{rows:10000}}});worker.onmessage({data:{type:'complete',records:[{row:2}]}});
 assert.deepEqual(await result,[{row:2}]);assert.equal(progress[0].rows,10000);assert.equal(worker.terminated,1);
});
test('stopping a large read terminates its worker and rejects without partial records',async()=>{
 const worker=fake(),controller=new AbortController();const pending=readReportOffThread(new ArrayBuffer(8),{createWorker:()=>worker,signal:controller.signal});
 controller.abort();await assert.rejects(pending,e=>e.name==='AbortError');assert.equal(worker.terminated,1);
});
test('worker failure is visible and always cleans up',async()=>{
 const worker=fake(),pending=readReportOffThread(new ArrayBuffer(8),{createWorker:()=>worker});
 worker.onerror({preventDefault(){}});await assert.rejects(pending,/Rapor.*okunamadı/);assert.equal(worker.terminated,1);
});
