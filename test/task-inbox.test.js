import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { LondonCore } from '../src/london-core.js';
import { StateStore } from '../src/state-store.js';
import { MicrosoftGraphClient } from '../src/microsoft-graph.js';
import { OpenAIClient } from '../src/openai-client.js';

const owner = 'ramy.mina@minaco.ca';
const summary = { id: 'graph-1', internetMessageId: '<test-1>' };
const email = { ...summary, from: { emailAddress: { address: owner } },
  subject: 'London Test 1 - Basic Task',
  body: { content: 'Confirm receipt, echo the subject, and reply with the exact phrase London Task Inbox is operational.' } };
const replyText = 'Receipt confirmed.\nSubject: London Test 1 - Basic Task\nLondon Task Inbox is operational';
function fixture(overrides = {}) {
  const sent = [];
  const graph = { principalMailbox: owner, readMailbox: 'london@minaco.ca',
    getLondonMessage: async () => structuredClone(email),
    listLondonInbox: async () => [summary],
    sendMail: async message => sent.push(message), ...overrides };
  const state = new StateStore();
  const core = new LondonCore({ graph, state, openai: {
    analyzeDelegatedEmail: async () => ({ text: replyText }),
    classifyInboundEmail: async () => ({ text: 'ACTION' }),
  }, logger: { error() {} } });
  return { core, sent, state };
}
test('basic task sends exact requested phrase to owner once, even with overlapping calls', async () => {
  const { core, sent } = fixture();
  const results = await Promise.all([core.processMessage(summary), core.processMessage(summary)]);
  await core.processMessage(summary);
  assert.equal(sent.length, 1);
  assert.equal(sent[0].to, owner);
  assert.equal(sent[0].body, replyText);
  assert.ok(sent[0].subject.includes(email.subject));
  assert.ok(results.some(result => result.reason === 'in-flight'));
});
test('external, absent and mismatched sender identities never trigger a completion send', async () => {
  for (const address of ['vendor@example.com', '', 'ramy.mina@minaco.ca.evil.example']) {
    const { core, sent } = fixture({ getLondonMessage: async () => ({ ...email, from: { emailAddress: { address } } }) });
    await core.processMessage(summary);
    assert.equal(sent.length, 0);
  }
  const { core, sent } = fixture({ getLondonMessage: async () => ({ ...email, sender: { emailAddress: { address: 'vendor@example.com' } } }) });
  await core.processMessage(summary);
  assert.equal(sent.length, 0);
});
test('delivery uncertainty is retained across a restart and never blindly resent', async () => {
  const directory = mkdtempSync(join(tmpdir(), 'london-test-'));
  try {
    const file = join(directory, 'state.json');
    const { core } = fixture({ sendMail: async () => { throw new Error('connection lost after submission'); } });
    core.state = new StateStore(file);
    await assert.rejects(core.processMessage(summary), /connection lost/);
    const restarted = fixture();
    restarted.core.state = new StateStore(file);
    assert.equal(restarted.core.state.state.processedMessages[summary.internetMessageId].result, 'delivery-pending-review');
    assert.equal((await restarted.core.processMessage(summary)).reason, 'duplicate');
    assert.equal(restarted.sent.length, 0);
  } finally { rmSync(directory, { recursive: true }); }
});
test('analysis failure remains retryable and attachment bytes reach the analysis', async () => {
  const parts = [{ type: 'input_file', filename: 'test.pdf', file_data: 'data:application/pdf;base64,JVBERg==' }];
  const { core, sent, state } = fixture({
    getLondonMessage: async () => ({ ...email, hasAttachments: true }),
    getLondonAttachments: async () => parts,
  });
  core.openai.analyzeDelegatedEmail = async () => { throw new Error('model unavailable'); };
  await assert.rejects(core.processMessage(summary), /model unavailable/);
  assert.equal(state.hasMessage(summary.internetMessageId), false);
  core.openai.analyzeDelegatedEmail = async (full, attachments) => {
    assert.equal(full.subject, email.subject); assert.deepEqual(attachments, parts);
    return { text: 'PDF analyzed.' };
  };
  await core.processMessage(summary);
  assert.equal(sent[0].body, 'PDF analyzed.');
});
function graphClient(fetchImpl) {
  return new MicrosoftGraphClient({ londonMailbox: 'london@minaco.ca', ramyMailbox: owner,
    readTenantId: 'tenant', readClientId: 'client', readClientSecret: 'test-only', fetchImpl });
}
test('Graph refuses third-party To and Cc before making any network request', async () => {
  let calls = 0;
  const graph = graphClient(async () => { calls++; throw new Error('unexpected'); });
  for (const message of [{ to: 'third@example.com' }, { to: [owner, 'third@example.com'] }, { to: owner, cc: ['third@example.com'] }, { to: [] }]) {
    await assert.rejects(graph.sendMail(message), /draft-only/);
  }
  assert.equal(calls, 0);
});
test('Graph fetches paginated documents, ignores inline signatures, and labels unsupported files', async () => {
  let calls = 0;
  const graph = graphClient(async url => {
    calls++;
    if (String(url).includes('/token')) return Response.json({ access_token: 'test', expires_in: 3600 });
    if (String(url).includes('page=2')) return Response.json({ value: [
      { '@odata.type': '#microsoft.graph.fileAttachment', name: 'archive.zip', contentBytes: 'AA==' },
    ] });
    return Response.json({ value: [
      { isInline: true, name: 'signature.png' },
      { '@odata.type': '#microsoft.graph.fileAttachment', name: 'test.pdf', contentType: 'application/pdf', contentBytes: 'JVBERg==' },
    ], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/attachments?page=2' });
  });
  const parts = await graph.getLondonAttachments('id');
  assert.equal(calls, 3);
  assert.equal(parts.length, 2);
  assert.equal(parts[0].file_data, 'data:application/pdf;base64,JVBERg==');
  assert.match(parts[1].text, /unsupported.*archive.zip/);
});
test('missing attachment content fails instead of claiming it was read', async () => {
  const graph = graphClient(async url => String(url).includes('/token')
    ? Response.json({ access_token: 'test' })
    : Response.json({ value: [{ '@odata.type': '#microsoft.graph.fileAttachment', name: 'test.pdf' }] }));
  await assert.rejects(graph.getLondonAttachments('id'), /content unavailable/);
});
test('real OpenAI request carries document bytes and asks for a deliverable, not a brief', async () => {
  let request;
  const client = new OpenAIClient({ apiKey: 'test-only', fetchImpl: async (url, options) => {
    request = JSON.parse(options.body);
    return Response.json({ output_text: replyText });
  } });
  const attachment = { type: 'input_file', filename: 'test.pdf', file_data: 'data:application/pdf;base64,JVBERg==' };
  assert.equal((await client.analyzeDelegatedEmail(email, [attachment])).text, replyText);
  assert.deepEqual(request.input[0].content[1], attachment);
  assert.match(request.input[0].content[0].text, /London Task Inbox is operational/);
  assert.match(request.instructions, /actual reply/);
  assert.match(request.instructions, /drafts/);
  assert.doesNotMatch(request.instructions, /Return a concise execution brief/);
});
test('state keys do not inherit prototype properties', () => {
  const state = new StateStore();
  assert.equal(state.hasMessage('toString'), false);
  state.markMessage('__proto__', { result: 'done' });
  assert.equal(state.hasMessage('__proto__'), true);
});
