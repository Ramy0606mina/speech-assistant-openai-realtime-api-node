import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { readFileSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { DurableTaskState } from '../task-state.js';
const source = readFileSync(new URL('../index.js', import.meta.url), 'utf8');
const owner = 'ramy.mina@minaco.ca';
const phrase = 'London Task Inbox is operational';
function section(start, end) {
  const a = source.indexOf(start), b = source.indexOf(end, a);
  assert.ok(a >= 0 && b > a);
  return source.slice(a, b);
}
function jobContext(overrides = {}) {
  const sent = [], state = new DurableTaskState();
  state.jobs.set('j1', { jobId: 'j1', status: 'queued' });
  const context = vm.createContext({
    taskInboxJobs: state.jobs, taskState: state, RAMY_MINACO_EMAIL: owner,
    LONDON_MINACO_EMAIL: 'london@minaco.ca', TASK_INBOX_ANALYSIS_MODEL: 'test',
    taskSubjectWithoutPrefixes: x => x, normalizeTaskSender: x => String(x).toLowerCase(),
    findLondonTaskInboxEmail: async () => ({ id: 'm1', from: { emailAddress: { address: owner } },
      subject: 'London Test 1 - Basic Task', body: { content: phrase }, hasAttachments: false }),
    listMailboxEmailAttachments: async () => [],
    buildTaskInboxFileInputs: async () => ({ content: [], reviewedFiles: [], skippedFiles: [] }),
    callOpenAIResponses: async request => {
      assert.match(request.instructions, /Replies and results to Ramy himself are authorized/);
      assert.match(request.instructions, /OTHER people/);
      return { text: JSON.stringify({ executiveConclusion: phrase }) };
    },
    extractOpenAIResponseText: x => x.text, stripJsonCodeFence: x => x,
    normalizeTaskReport: x => x, renderTaskInboxReportHtml: ({ report }) => report.executiveConclusion,
    sendEmailFromLondon: async msg => sent.push(msg), pruneTaskInboxState: () => state.save(),
    console: { log() {}, error() {} }, ...overrides,
  });
  vm.runInContext(section('const runTaskInboxJob', 'const queueTaskInboxJob') + '\nthis.run = runTaskInboxJob;', context);
  return { context, state, sent };
}
test('actual production job delivers owner receipt once with no action-register calls', async () => {
  const { context, sent, state } = jobContext();
  await context.run({ jobId: 'j1', messageId: 'm1', subject: 'London Test 1 - Basic Task' });
  assert.equal(sent.length, 1);
  assert.equal(sent[0].to, owner);
  assert.equal(sent[0].body, phrase);
  assert.ok(sent[0].subject.endsWith('London Test 1 - Basic Task'));
  assert.equal(state.jobs.get('j1').status, 'completed');
});
test('uncertain completion send is held for review without failure-email duplication', async () => {
  let sends = 0;
  const { context, state } = jobContext({ sendEmailFromLondon: async () => { sends++; throw Error('timeout'); } });
  await context.run({ jobId: 'j1', messageId: 'm1', subject: 'Test' });
  assert.equal(sends, 1);
  assert.equal(state.jobs.get('j1').status, 'delivery-pending-review');
});
test('Graph sender is verified before analysis and no third-party recipient is contacted', async () => {
  let analyses = 0;
  const { context, sent } = jobContext({
    findLondonTaskInboxEmail: async () => ({ from: { emailAddress: { address: 'stranger@example.com' } } }),
    callOpenAIResponses: async () => { analyses++; },
  });
  await context.run({ jobId: 'j1', messageId: 'm1', subject: 'Test' });
  assert.equal(analyses, 0);
  assert.ok(sent.every(x => x.to === owner));
});
function intake() {
  const routes = new Map(), state = new DurableTaskState(), deferred = [];
  const context = vm.createContext({
    fastify: { post: (path, fn) => routes.set(path, fn), get() {} },
    TASK_INBOX_SECRET: 'test-secret', RAMY_MINACO_EMAIL: owner,
    normalizeTaskSender: x => String(x).trim().toLowerCase(),
    taskSubjectWithoutPrefixes: x => x, taskState: state,
    taskInboxJobs: state.jobs, processedTaskInboxKeys: state.keys,
    setImmediate: fn => deferred.push(fn), runTaskInboxJob: async () => {},
    console: { error() {} },
  });
  vm.runInContext(section('const queueTaskInboxJob', '// Fastify / London voice configuration'), context);
  vm.runInContext(section("fastify.post('/task-inbox'", "fastify.all('/incoming-sms'"), context);
  const send = async (body, secret = 'test-secret') => {
    const reply = { status: 200, code(n) { this.status = n; return this; }, send(x) { this.body = x; return this; } };
    await routes.get('/task-inbox')({ body, headers: { 'x-london-task-secret': secret } }, reply);
    return reply;
  };
  return { send, state, deferred };
}
test('endpoint rejects invalid secret and unauthorized sender', async () => {
  const { send, deferred } = intake();
  assert.equal((await send({ from: owner, id: 'm1' }, 'bad')).status, 401);
  assert.equal((await send({ from: 'stranger@example.com', id: 'm1' })).status, 403);
  assert.equal(deferred.length, 0);
});
test('endpoint requires exact message identity and suppresses repeated deliveries', async () => {
  const { send, deferred } = intake();
  assert.equal((await send({ from: owner, subject: 'Test' })).status, 400);
  assert.equal((await send({ from: owner, id: 'm1' })).status, 202);
  assert.equal((await send({ from: owner, id: 'm1', internetMessageId: 'different' })).body.duplicate, true);
  assert.equal(deferred.length, 1);
});
test('state reload retains queued jobs, duplicate keys and uncertain deliveries', () => {
  const dir = mkdtempSync(join(tmpdir(), 'london-live-'));
  try {
    const file = join(dir, 'state.json'), state = new DurableTaskState(file);
    state.jobs.set('q', { status: 'queued' });
    state.jobs.set('r', { status: 'running' });
    state.jobs.set('s', { status: 'delivery-pending-review' });
    state.keys.set('message', { jobId: 's' }); state.save();
    const restored = new DurableTaskState(file);
    assert.equal(restored.jobs.get('q').status, 'queued');
    assert.equal(restored.jobs.get('r').status, 'interrupted-review');
    assert.equal(restored.jobs.get('s').status, 'delivery-pending-review');
    assert.equal(restored.keys.get('message').jobId, 's');
  } finally { rmSync(dir, { recursive: true }); }
});
test('exact message lookup failure cannot silently select another email', async () => {
  const context = vm.createContext({
    LONDON_MINACO_EMAIL: 'london@minaco.ca', RAMY_MINACO_EMAIL: owner,
    normalizeTaskSender: x => x, getFullMailboxEmail: async () => { throw Error('not found'); },
    getRecentInboxEmails: async () => { throw Error('must not fall back'); },
  });
  vm.runInContext(section('const findLondonTaskInboxEmail', 'const buildTaskInboxFileInputs') + '\nthis.find = findLondonTaskInboxEmail;', context);
  await assert.rejects(context.find({ messageId: 'missing', subject: 'Test' }), /not found/);
});
test('actual attachment builder sends PDF bytes and excludes inline signatures', async () => {
  const context = vm.createContext({
    MAX_TASK_ATTACHMENTS: 10, MAX_ATTACHMENT_BYTES: 45000000, MAX_TASK_TOTAL_BYTES: 45000000,
    LONDON_MINACO_EMAIL: 'london@minaco.ca',
    getMailboxEmailAttachment: async () => ({ name: 'test.pdf', contentType: 'application/pdf', contentBytes: 'JVBERg==', size: 4 }),
  });
  vm.runInContext(section('const buildTaskInboxFileInputs', 'const normalizeTaskReport') + '\nthis.build = buildTaskInboxFileInputs;', context);
  const result = await context.build({ messageId: 'm1', attachments: [{ attachmentId: 'a', size: 4 }, { attachmentId: 'b', isInline: true }] });
  assert.equal(result.content.length, 1);
  assert.equal(result.content[0].file_data, 'data:application/pdf;base64,JVBERg==');
});
test('messaging routes are removed or inert, voice routes remain', () => {
  assert.doesNotMatch(source, /fastify\.all\('\/incoming-whatsapp'/);
  assert.match(source, /fastify\.all\('\/incoming-sms', async/);
  assert.match(source, /const sendTwilioChannelMessage = async \(\) => \{ throw new Error/);
  assert.match(source, /\/incoming-call/);
  assert.match(source, /\/media-stream/);
});

