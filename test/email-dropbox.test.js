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

test('Dropbox renews credentials once for concurrent reads and retries rejected token once', async () => {
  let renewals = 0;
  let reads = 0;
  const dbx = new DropboxClient({ refreshToken: 'refresh-test', appKey: 'app-test', appSecret: 'secret-test', fetchImpl: async (url, options) => {
    if (String(url).endsWith('/oauth2/token')) {
      renewals++;
      const form = new URLSearchParams(options.body);
      assert.equal(form.get('grant_type'), 'refresh_token');
      assert.equal(form.get('refresh_token'), 'refresh-test');
      return Response.json({ access_token: `renewed-${renewals}`, expires_in: 14400 });
    }
    reads++;
    assert.match(options.headers.Authorization, /^Bearer renewed-/);
    return Response.json({ entries: [] });
  } });
  await Promise.all([dbx.listFolder(), dbx.listFolder()]);
  assert.equal(renewals, 1);
  assert.equal(reads, 2);
  let rejected = false;
  const previous = dbx.fetchImpl;
  dbx.fetchImpl = async (url, options) => {
    if (!String(url).endsWith('/oauth2/token') && !rejected) { rejected = true; return Response.json({ error: 'expired' }, { status: 401 }); }
    return previous(url, options);
  };
  await dbx.listFolder();
  assert.equal(renewals, 2);
});

test('kitchen-sized PDF is accepted by the document limit', async () => {
  const bytes = new Uint8Array(29147040);
  const dbx = new DropboxClient({ accessToken: 'test', fetchImpl: async () => new Response(bytes) });
  const file = await dbx.readFile('kitchen.pdf');
  assert.equal(file.size, 29147040);
});

test('report is saved before owner email and save failure does not claim completion', async () => {
  const events = [];
  const core = new LondonCore({ state: new StateStore(), graph: {
    principalMailbox: 'owner@example.com', readMailbox: 'london@example.com',
    getLondonMessage: async () => ({ from: { emailAddress: { address: 'owner@example.com' } }, subject: 'Compare plans' }),
    sendMail: async mail => { events.push('email'); assert.match(mail.body, /Saved in Dropbox:/); },
  }, openai: { analyzeDelegatedEmail: async () => ({ text: 'Comparison findings.' }) }, dropbox: {
    saveReports: true,
    saveReport: async report => { events.push('save'); assert.equal(report.text, 'Comparison findings.'); return { path: '/LONDON - ACCESS/London Work/report.md' }; },
  } });
  await core.processMessage({ id: 'save-test' });
  assert.deepEqual(events, ['save', 'email']);
  core.dropbox.saveReport = async () => { throw new Error('Storage unavailable'); };
  await assert.rejects(core.processMessage({ id: 'save-failed' }), /Storage unavailable/);
  assert.equal(core.state.hasMessage('save-failed'), false);
  assert.deepEqual(events, ['save', 'email']);
});

test('generated reports stay in London Work, retain changed versions and retry identical output safely', async () => {
  const uploaded = [];
  const dbx = new DropboxClient({ accessToken: 'test', saveReports: true, fetchImpl: async (url, options) => {
    if (String(url).endsWith('create_folder_v2')) return Response.json({ error_summary: 'path/conflict/folder/...' }, { status: 409 });
    const arg = JSON.parse(options.headers['Dropbox-API-Arg']);
    assert.ok(arg.path.startsWith('/LONDON - ACCESS/London Work/London - '));
    assert.equal(arg.autorename, false);
    assert.match(options.body.toString(), /Findings/);
    uploaded.push(arg.path);
    return Response.json({ id: 'report-id', path_display: arg.path });
  } });
  const report = { taskKey: 'one', subject: '../../source.pdf', text: 'Findings' };
  await dbx.saveReport(report);
  await dbx.saveReport(report);
  await dbx.saveReport({ ...report, text: 'Findings revised' });
  assert.equal(uploaded[0], uploaded[1]);
  assert.notEqual(uploaded[0], uploaded[2]);
  assert.ok(uploaded.every(path => !path.includes('..')));
});
