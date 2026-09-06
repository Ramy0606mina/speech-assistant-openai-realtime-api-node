import test from 'node:test';
import assert from 'node:assert/strict';
import { DeliveryGuard } from '../src/delivery-guard.js';
import { LondonCore } from '../src/london-core.js';
import { StateStore } from '../src/state-store.js';
import { DropboxClient } from '../src/dropbox-client.js';

function store() {
  const records = new Map();
  return { records, async readDeliveryRecord(k) { return records.get(k) || null; },
    async createDeliveryRecord(k,v) { if (records.has(k)) return false; records.set(k,v); return true; } };
}
function core(guard, sends, { failSend=false, failAnalysis=false }={}) {
  return new LondonCore({ state: new StateStore(), deliveryGuard: guard,
    graph: { principalMailbox:'owner@example.com', readMailbox:'london@example.com',
      getLondonMessage: async () => ({ subject:'Kitchen',receivedDateTime:new Date(guard.notBefore+1000).toISOString(),from:{emailAddress:{address:'owner@example.com'}} }),
      sendMail:async () => { sends.push('send');if(failSend)throw new Error('ambiguous timeout'); } },
    openai:{analyzeDelegatedEmail:async()=>{if(failAnalysis)throw new Error('analysis unavailable');return {text:'Report'};}} });
}
test('durable claim prevents replay after local state loss and across concurrent workers',async()=>{
 const s=store(),a=new DeliveryGuard(s,'london@example.com'),b=new DeliveryGuard(s,'london@example.com');
 await Promise.all([a.initialize(),b.initialize()]);const sends=[];
 const msg={id:'one',internetMessageId:'one'};
 await Promise.all([core(a,sends).processMessage(msg),core(b,sends).processMessage(msg)]);
 await core(newGuard(),sends).processMessage(msg);
 function newGuard(){const g=new DeliveryGuard(s,'london@example.com');g.notBefore=a.notBefore;return g;}
 assert.equal(sends.length,1);
});
test('ambiguous dispatch never sends again after restart',async()=>{
 const s=store(),g=new DeliveryGuard(s,'london@example.com');await g.initialize();const sends=[];const msg={id:'one'};
 await assert.rejects(core(g,sends,{failSend:true}).processMessage(msg),/timeout/);
 await core(g,sends).processMessage(msg);assert.equal(sends.length,1);
});
test('analysis failure before claiming remains retryable',async()=>{
 const s=store(),g=new DeliveryGuard(s,'london@example.com');await g.initialize();const sends=[];const msg={id:'one'};
 await assert.rejects(core(g,sends,{failAnalysis:true}).processMessage(msg),/analysis/);
 await core(g,sends).processMessage(msg);assert.equal(sends.length,1);
});
test('migration boundary survives restart and holds old or undated requests',async()=>{
 const s=store(),g=new DeliveryGuard(s,'london@example.com');await g.initialize();const boundary=g.notBefore;
 const again=new DeliveryGuard(s,'london@example.com');await again.initialize();assert.equal(again.notBefore,boundary);
 assert.equal(await again.check('old',new Date(boundary-1).toISOString()),'historical-review');
 assert.equal(await again.check('undated',null),'historical-review');
 assert.equal(await again.check('new',new Date(boundary+1).toISOString()),null);
});
test('ledger failure blocks delivery instead of falling back to volatile state',async()=>{
 const g=new DeliveryGuard({readDeliveryRecord:async()=>{throw new Error('offline');}},'london@example.com');
 await assert.rejects(g.initialize(),/offline/);await assert.rejects(g.check('x',new Date().toISOString()),/initialized/);
});
test('Dropbox claim uses strict conflict and treats only file conflict as a lost claim',async()=>{
 const args=[];
 const db=new DropboxClient({accessToken:'test',fetchImpl:async(url,o)=>{
   if(String(url).endsWith('create_folder_v2'))return Response.json({id:'folder'});
   args.push(JSON.parse(o.headers['Dropbox-API-Arg']));
   return Response.json({error_summary:'path/conflict/file/'},{status:409});
 }});
 assert.equal(await db.createDeliveryRecord('safe-key',{result:'pending'}),false);
 assert.equal(args[0].strict_conflict,true);assert.equal(args[0].mode,'add');assert.equal(args[0].autorename,false);
 await assert.rejects(db.readDeliveryRecord('../escape'),/Invalid/);
});
