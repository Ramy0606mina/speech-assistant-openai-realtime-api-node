import test from 'node:test';
import assert from 'node:assert/strict';
import { OpenAIClient } from '../src/openai-client.js';
import { DropboxClient } from '../src/dropbox-client.js';
import { LondonCore } from '../src/london-core.js';
import { StateStore } from '../src/state-store.js';

test('owner email searches Dropbox, downloads PDF bytes and delivers their analysis once', async () => {
  const path = '/LONDON - ACCESS/Kitchen.pdf';
  const calls = [];
  const dbx = new DropboxClient({ accessToken: 'test', fetchImpl: async (url, options) => {
    calls.push(String(url));
    if (String(url).endsWith('search_v2')) return Response.json({ matches: [{ metadata: { metadata: { name: 'Kitchen.pdf', path_display: path, '.tag': 'file' } } }] });
    assert.equal(JSON.parse(options.headers['Dropbox-API-Arg']).path, path);
    return new Response('%PDF-1.7\nunique-test-kitchen-plan');
  } });
  let step = 0;
  const ai = new OpenAIClient({ apiKey: 'test', fetchImpl: async (url, options) => {
    const request = JSON.parse(options.body);
    step++;
    if (step === 1) return Response.json({ output: [{ type: 'function_call', call_id: 'search', name: 'search_dropbox', arguments: '{"query":"kitchen"}' }] });
    if (step === 2) {
      assert.match(JSON.stringify(request.input), /Kitchen.pdf/);
      return Response.json({ output: [{ type: 'function_call', call_id: 'read', name: 'read_dropbox_file', arguments: JSON.stringify({ path }) }] });
    }
    const doc = request.input.flatMap(item => item.content || []).find(item => item.type === 'input_file');
    assert.match(Buffer.from(doc.file_data.split(',')[1], 'base64').toString(), /unique-test-kitchen-plan/);
    return Response.json({ output_text: 'Read Kitchen.pdf from the shared folder.' });
  } });
  const sent = [];
  const core = new LondonCore({ dropbox: dbx, openai: ai, state: new StateStore(), graph: {
    principalMailbox: 'owner@example.com', readMailbox: 'london@example.com',
    getLondonMessage: async () => ({ from: { emailAddress: { address: 'owner@example.com' } }, subject: 'Kitchen drawings', body: { content: 'Find and review my kitchen PDF in Dropbox.' } }),
    sendMail: async value => sent.push(value),
  } });
  await core.processMessage({ id: 'one', internetMessageId: 'one' });
  await core.processMessage({ id: 'one', internetMessageId: 'one' });
  assert.equal(sent.length, 1);
  assert.equal(sent[0].to, 'owner@example.com');
  assert.equal(calls.length, 2);
});

test('Dropbox refuses outside paths, traversal and unsupported files before network access', async () => {
  const dbx = new DropboxClient({ accessToken: 'test', fetchImpl: async () => { throw new Error('unexpected network'); } });
  for (const path of ['/Other/private.pdf', '../private.pdf', '/LONDON - ACCESS/../private.pdf', '/LONDON - ACCESS/a/./b.pdf']) {
    await assert.rejects(dbx.readFile(path), /root/);
  }
  await assert.rejects(dbx.readFile('run.exe'), /Unsupported/);
});

test('failed Dropbox read is reported to model without invented document bytes', async () => {
  let step = 0;
  const client = new OpenAIClient({ apiKey: 'test', fetchImpl: async (url, options) => {
    if (++step === 1) return Response.json({ output: [{ type: 'function_call', call_id: 'read', name: 'read_dropbox_file', arguments: '{"path":"missing.pdf"}' }] });
    const request = JSON.parse(options.body);
    assert.match(JSON.stringify(request.input), /HTTP 401/);
    assert.equal(request.input.flatMap(item => item.content || []).some(part => part.type === 'input_file'), false);
    return Response.json({ output_text: 'Dropbox authorization failed; I could not read the PDF.' });
  } });
  const response = await client.analyzeDelegatedEmail({ subject: 'Read PDF' }, [], { dropbox: {
    readFile: async () => { throw Object.assign(new Error('private server detail'), { status: 401 }); },
  } });
  assert.match(response.text, /authorization failed/);
});

test('Dropbox stream enforces byte limit and rejects failed downloads', async () => {
  const dbx = new DropboxClient({ accessToken: 'test', fetchImpl: async () => new Response('too many bytes') });
  await assert.rejects(dbx.readFile('test.pdf', 3), /limit/);
  dbx.fetchImpl = async () => new Response('', { status: 401 });
  await assert.rejects(dbx.readFile('test.pdf'), /HTTP 401/);
});
