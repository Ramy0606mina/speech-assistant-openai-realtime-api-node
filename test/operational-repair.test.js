import test from 'node:test';
import assert from 'node:assert/strict';
import { MailboxWorker } from '../src/mailbox-worker.js';
import { OpenAIClient } from '../src/openai-client.js';
import { LondonCore } from '../src/london-core.js';
import { StateStore } from '../src/state-store.js';
import { DeliveryGuard } from '../src/delivery-guard.js';
import { reportSubject } from '../src/report-format.js';

test('ledger outage allows inbox visibility but no execution, then recovers with original historical boundary', async () => {
  let outage = true, analyses = 0, reads = 0;
  const records = new Map();
  const guard = new DeliveryGuard({
    readDeliveryRecord: async key => { if (outage) throw Error('offline'); return records.get(key); },
    createDeliveryRecord: async (key, value) => { if (records.has(key)) return false; records.set(key, value); return true; },
  }, 'london@example.com');
  records.set(`${guard.prefix}-cutover`, { notBefore: '2026-09-08T00:00:00Z' });
  const summary = { id: 'old', internetMessageId: 'old' };
  const state = new StateStore();
  const graph = { principalMailbox: 'owner@example.com', readMailbox: 'london@example.com',
    listLondonInbox: async () => { reads++; return [summary]; },
    getLondonMessage: async () => ({ ...summary, from: { emailAddress: { address: 'owner@example.com' } }, receivedDateTime: '2026-09-07T00:00:00Z' }),
    sendMail: async () => assert.fail('historical mail must never replay'),
  };
  const london = new LondonCore({ graph, state, deliveryGuard: guard, openai: { analyzeDelegatedEmail: async () => { analyses++; } } });
  const worker = new MailboxWorker({ graph, london, guard });
  assert.equal((await worker.poll()).reason, 'delivery-ledger-unavailable');
  assert.equal(reads, 1);
  assert.equal(state.hasMessage('old'), false);
  assert.equal(analyses, 0);
  outage = false;
  assert.equal((await worker.poll()).ok, true);
  assert.equal(analyses, 0);
  assert.equal(state.state.processedMessages.old.result, 'historical-review');
});

test('quota exhaustion is exposed and prevents repeated network attempts until cooldown; recovery clears it', async () => {
  let calls = 0;
  const client = new OpenAIClient({ apiKey: 'test', fetchImpl: async () => {
    calls++;
    return calls === 1 ? Response.json({ error: { type: 'insufficient_quota', code: 'credit_balance_exhausted', message: 'No credits' } }, { status: 429 }) : Response.json({ output_text: 'Recovered' });
  } });
  await assert.rejects(client.respond({ input: 'test' }));
  await assert.rejects(client.respond({ input: 'test' }), /quota/);
  assert.equal(calls, 1);
  assert.equal(client.status.reason, 'openai-quota-exhausted');
  client.retryAfter = Date.now() - 1;
  assert.equal((await client.respond({ input: 'test' })).text, 'Recovered');
  assert.equal(client.status.state, 'available');
});

test('a successful inbox read with failed tasks is degraded, not healthy', async () => {
  const worker = new MailboxWorker({ guard: { initialize: async () => {} }, london: {
    pollOnce: async () => ({ checked: 2, results: [{ processed: false }, { skipped: true }] }),
  } });
  assert.deepEqual(await worker.poll(), { ok: false, state: 'degraded', checked: 2, failed: 1, heldForReview: 0, newFailures: 1, lastCheckedAt: worker.status.lastCheckedAt });
});

test('optional report failure still delivers the result once and never repeats the verified rename', async () => {
  let renames = 0, sent = 0;
  const records = new Map();
  const guard = new DeliveryGuard({ readDeliveryRecord: async key => records.get(key), createDeliveryRecord: async (key, value) => {
    if (records.has(key)) return false; records.set(key, value); return true;
  } }, 'london@example.com');
  records.set(`${guard.prefix}-cutover`, { notBefore: '2026-09-08T00:00:00Z' });
  await guard.initialize();
  const dependencies = {
    graph: { principalMailbox: 'owner@example.com', readMailbox: 'london@example.com',
      getLondonMessage: async () => ({ from: { emailAddress: { address: 'owner@example.com' } }, receivedDateTime: '2026-09-10T00:00:00Z', body: { content: 'Rename old.pdf to new.pdf' } }),
      sendMail: async mail => { sent++; assert.match(mail.subject, /Needs Attention/); assert.match(mail.body, /old.pdf → new.pdf/); assert.match(mail.body, /saving was not confirmed/); },
    },
    openai: { analyzeDelegatedEmail: async () => ({ text: 'Prepared', dropboxRenames: [{ sourcePath: 'old.pdf', destinationName: 'new.pdf' }] }) },
    dropbox: { saveReports: true, saveReport: async () => { throw Error('offline'); }, renameFile: async () => { renames++; return { from: 'old.pdf', path: 'new.pdf' }; } },
    deliveryGuard: guard, logger: { error() {} },
  };
  await new LondonCore({ ...dependencies, state: new StateStore() }).processMessage({ id: 'one' });
  // A process restart loses ephemeral state, but the durable claim remains.
  await new LondonCore({ ...dependencies, state: new StateStore() }).processMessage({ id: 'one' });
  assert.equal(sent, 1); assert.equal(renames, 1);
});

test('new response subjects do not accumulate across reply chains', () => {
  assert.equal(reportSubject('RE: LONDON — Task Response | RE: LONDON — Task Needs Attention | Rename files'), 'Rename files');
});


test('health distinguishes held analysis from new failures without exposing message data',async()=>{
 const worker=new MailboxWorker({guard:{initialize:async()=>{}},london:{pollOnce:async()=>({checked:3,results:[{processed:false,reason:'analysis-needs-review',key:'private'},{processed:false,error:'secret'},{skipped:true}]})}});
 const status=await worker.poll();assert.equal(status.heldForReview,1);assert.equal(status.newFailures,1);assert.equal(status.ok,false);assert.doesNotMatch(JSON.stringify(status),/private|secret/);
});


test('inbound classification is durable across restart and holds failed analysis',async()=>{
 const records=new Map();const guard=new DeliveryGuard({readDeliveryRecord:async k=>records.get(k),createDeliveryRecord:async(k,v)=>{if(records.has(k))return false;records.set(k,v);return true;}},'inbound');
 await guard.initialize();let calls=0,fail=false;
 const deps={deliveryGuard:guard,graph:{principalMailbox:'owner@example.com',readMailbox:'london@example.com',getLondonMessage:async()=>({from:{emailAddress:{address:'vendor@example.com'}},receivedDateTime:new Date(Date.now()+1000).toISOString()})},openai:{classifyInboundEmail:async()=>{calls++;if(fail)throw Error('offline');return{text:'ACTION'};}}};
 const make=()=>new LondonCore({...deps,state:new StateStore()});
 await make().processMessage({id:'ok'});await make().processMessage({id:'ok'});assert.equal(calls,1);
 fail=true;await assert.rejects(make().processMessage({id:'failed'}));assert.equal((await make().processMessage({id:'failed'})).reason,'analysis-needs-review');assert.equal(calls,2);
});
