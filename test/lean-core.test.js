import test from 'node:test';
import assert from 'node:assert/strict';
import { extractResponseText } from '../src/openai-client.js';
import { DropboxClient } from '../src/dropbox-client.js';
import { LondonCore } from '../src/london-core.js';
import { loadConfig, configurationStatus } from '../src/config.js';

class MemoryState {
  constructor() { this.seen = new Set(); this.polls = 0; }
  hasMessage(key) { return this.seen.has(key); }
  markMessage(key) { this.seen.add(key); }
  markPoll() { this.polls += 1; }
}

test('OpenAI text parser tolerates reasoning item before message item', () => {
  const text = extractResponseText({ output: [
    { type: 'reasoning', content: [] },
    { type: 'message', content: [{ type: 'output_text', text: 'URGENT' }] },
  ] });
  assert.equal(text, 'URGENT');
});

test('OpenAI text parser prefers top-level output_text', () => {
  assert.equal(extractResponseText({ output_text: 'ACTION', output: [] }), 'ACTION');
});

test('Dropbox client keeps relative paths under configured London root', () => {
  const dbx = new DropboxClient({ accessToken: 'x', rootPath: '/LONDON - ACCESS', fetchImpl: async () => {} });
  assert.equal(dbx.resolvePath('Projects/Laval'), '/LONDON - ACCESS/Projects/Laval');
  assert.equal(dbx.resolvePath('/LONDON - ACCESS/Projects'), '/LONDON - ACCESS/Projects');
});

test('configuration status does not expose secret values', () => {
  const config = loadConfig({
    OPENAI_API_KEY: 'secret-openai', MS_TENANT_ID: 't', MS_CLIENT_ID: 'c', MS_CLIENT_SECRET: 's',
    ACTIONS_MS_TENANT_ID: 't2', ACTIONS_MS_CLIENT_ID: 'c2', ACTIONS_MS_CLIENT_SECRET: 's2',
    LONDON_MINACO_EMAIL: 'london@example.com', DROPBOX_ACCESS_TOKEN: 'dbx', DROPBOX_ROOT_PATH: '/LONDON - ACCESS',
  });
  const status = configurationStatus(config);
  assert.equal(status.openaiConfigured, true);
  assert.equal(JSON.stringify(status).includes('secret-openai'), false);
  assert.equal(JSON.stringify(status).includes('dbx'), false);
});

test('London processes Ramy email as delegated task and sends completion through Graph', async () => {
  const state = new MemoryState();
  const sent = [];
  const graph = {
    principalMailbox: 'ramy@minaco.ca', readMailbox: 'london@minaco.ca',
    async listLondonInbox() { return [{ id: '1', internetMessageId: 'm1' }]; },
    async getLondonMessage() { return { id: '1', internetMessageId: 'm1', from: { emailAddress: { address: 'ramy@minaco.ca' } }, subject: 'Review', body: { content: 'Review this.' } }; },
    async sendMail(message) { sent.push(message); return { sent: true }; },
  };
  const openai = {
    async analyzeDelegatedEmail() { return { text: 'Objective: review.' }; },
    async classifyInboundEmail() { throw new Error('wrong classifier'); },
  };
  const core = new LondonCore({ graph, openai, dropbox: {}, state, logger: { error() {} } });
  const result = await core.pollOnce(10);
  assert.equal(result.results[0].type, 'delegated-task');
  assert.equal(result.results[0].analysis, 'Objective: review.');
  assert.equal(result.results[0].completionSent, true);
  assert.equal(sent.length, 1);
  assert.equal(sent[0].to, 'ramy@minaco.ca');
  assert.equal(sent[0].subject, 'LONDON — Task Complete | Review');
});

test('London classifies external email without executing it as Ramy task', async () => {
  const state = new MemoryState();
  const graph = {
    principalMailbox: 'ramy@minaco.ca', readMailbox: 'london@minaco.ca',
    async listLondonInbox() { return [{ id: '2', internetMessageId: 'm2' }]; },
    async getLondonMessage() { return { id: '2', internetMessageId: 'm2', from: { emailAddress: { address: 'vendor@example.com' } }, subject: 'Invoice', body: { content: 'Invoice attached.' } }; },
  };
  const openai = {
    async analyzeDelegatedEmail() { throw new Error('must not execute external sender as delegated task'); },
    async classifyInboundEmail() { return { text: 'ACTION' }; },
  };
  const core = new LondonCore({ graph, openai, dropbox: {}, state, logger: { error() {} } });
  const result = await core.pollOnce(10);
  assert.equal(result.results[0].type, 'inbound-email');
  assert.equal(result.results[0].classification, 'ACTION');
});

test('London deduplicates already processed messages', async () => {
  const state = new MemoryState();
  state.markMessage('m3');
  let fullReads = 0;
  const graph = {
    principalMailbox: 'ramy@minaco.ca', readMailbox: 'london@minaco.ca',
    async listLondonInbox() { return [{ id: '3', internetMessageId: 'm3' }]; },
    async getLondonMessage() { fullReads += 1; return {}; },
  };
  const core = new LondonCore({ graph, openai: {}, dropbox: {}, state, logger: { error() {} } });
  const result = await core.pollOnce(10);
  assert.equal(result.results[0].reason, 'duplicate');
  assert.equal(fullReads, 0);
});

test('London ignores messages sent by its own mailbox', async () => {
  const state = new MemoryState();
  const graph = {
    principalMailbox: 'ramy@minaco.ca', readMailbox: 'london@minaco.ca',
    async listLondonInbox() { return [{ id: '4', internetMessageId: 'm4' }]; },
    async getLondonMessage() { return { from: { emailAddress: { address: 'London@Minaco.ca' } } }; },
  };
  const core = new LondonCore({ graph, openai: {}, dropbox: {}, state, logger: { error() {} } });
  const result = await core.pollOnce(10);
  assert.equal(result.results[0].reason, 'self-message');
});
