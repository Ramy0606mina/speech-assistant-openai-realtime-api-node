import test from 'node:test';
import assert from 'node:assert/strict';
import { normalizePhone, isAuthorizedCaller, buildIncomingCallTwiML } from '../src/voice-gateway.js';

test('normalizes North American phone formatting', () => {
  assert.equal(normalizePhone('(514) 814-3333'), '+15148143333');
  assert.equal(normalizePhone('+1 514 814 3333'), '+15148143333');
});

test('authorizes only the configured principal caller', () => {
  assert.equal(isAuthorizedCaller('+1 514 814 3333', '+15148143333'), true);
  assert.equal(isAuthorizedCaller('+1 514 555 0000', '+15148143333'), false);
});

test('incoming call TwiML connects only to the generated tokenized media route', () => {
  const xml = buildIncomingCallTwiML({ host: 'london-ai-pr-1.onrender.com', streamToken: 'abc-123' });
  assert.match(xml, /wss:\/\/london-ai-pr-1\.onrender\.com\/media-stream\/abc-123/);
  assert.match(xml, /<Connect>/);
});
