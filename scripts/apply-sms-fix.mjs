import fs from 'node:fs';

const indexPath = 'index.js';
const testPath = 'test/task-inbox.test.js';
let source = fs.readFileSync(indexPath, 'utf8');

const disabledSender = "const sendTwilioChannelMessage = async () => { throw new Error('SMS is paused and WhatsApp has been removed.'); };";
const activeSender = [
  "const sendTwilioChannelMessage = async ({ to, from, body }) => {",
  "  if (!TWILIO_ACCOUNT_SID || !TWILIO_AUTH_TOKEN) {",
  "    throw new Error('Missing Twilio messaging credentials.');",
  "  }",
  "  if (!to || !from) throw new Error('Twilio To and From addresses are required.');",
  "",
  "  const auth = Buffer.from(",
  "    `${TWILIO_ACCOUNT_SID}:${TWILIO_AUTH_TOKEN}`",
  "  ).toString('base64');",
  "",
  "  const cleanBody = String(body || '').trim().slice(0, MAX_MESSAGING_REPLY_CHARS);",
  "  const form = new URLSearchParams({",
  "    To: String(to),",
  "    From: String(from),",
  "    Body: cleanBody || 'Updated.',",
  "  });",
  "",
  "  const response = await fetchWithTimeout(",
  "    `https://api.twilio.com/2010-04-01/Accounts/${TWILIO_ACCOUNT_SID}/Messages.json`,",
  "    {",
  "      method: 'POST',",
  "      headers: {",
  "        Authorization: `Basic ${auth}` ,",
  "        'Content-Type': 'application/x-www-form-urlencoded',",
  "      },",
  "      body: form.toString(),",
  "    },",
  "    12000",
  "  );",
  "",
  "  const data = await response.json();",
  "  if (!response.ok) {",
  "    throw new Error(`Twilio messaging send failed: ${data.message || response.status}`);",
  "  }",
  "  return data;",
  "};",
].join('\n');

if (source.includes(disabledSender)) source = source.replace(disabledSender, activeSender);

const inertRoute = "fastify.all('/incoming-sms', async (_request, reply) => reply.type('text/xml').send('<Response/>'));";
const activeRoute = [
  "const handleIncomingSms = async (request, reply) => {",
  "  const body = request.body && typeof request.body === 'object' ? request.body : {};",
  "  const from = String(body.From || request.query?.From || '').trim();",
  "  const to = String(body.To || request.query?.To || '').trim();",
  "  const messageBody = String(body.Body || request.query?.Body || '').trim();",
  "  const messageSid = String(body.MessageSid || body.SmsMessageSid || '').trim();",
  "  const numMedia = Number(body.NumMedia || 0);",
  "",
  "  const validSignature = validateTwilioFormWebhook(request);",
  "  const authorizedSender = isAuthorizedRamyMessagingSender(from);",
  "  const emptyTwiml = '<?xml version=\"1.0\" encoding=\"UTF-8\"?><Response></Response>';",
  "",
  "  if (!validSignature || !authorizedSender) {",
  "    return reply.type('text/xml').code(403).send(emptyTwiml);",
  "  }",
  "",
  "  pruneMessagingState();",
  "  if (messageSid && processedMessagingSids.has(messageSid)) {",
  "    return reply.type('text/xml').send(emptyTwiml);",
  "  }",
  "  if (messageSid) processedMessagingSids.set(messageSid, Date.now());",
  "  reply.type('text/xml').send(emptyTwiml);",
  "",
  "  setImmediate(async () => {",
  "    try {",
  "      let result;",
  "      if (!messageBody && numMedia > 0) {",
  "        result = { success: false, reply: 'I received the attachment. For document or spreadsheet analysis, email it to london@minaco.ca so I can process the full file safely.' };",
  "      } else {",
  "        result = await processExecutiveMessagingInstruction({ text: messageBody, channel: 'sms', sender: from });",
  "      }",
  "      await sendTwilioChannelMessage({",
  "        to: normalizePhoneIdentity(from),",
  "        from: normalizePhoneIdentity(to || TWILIO_PHONE_NUMBER),",
  "        body: result.reply,",
  "      });",
  "    } catch (error) {",
  "      console.error('Inbound SMS failure:', error);",
  "      try {",
  "        await sendTwilioChannelMessage({",
  "          to: normalizePhoneIdentity(from),",
  "          from: normalizePhoneIdentity(to || TWILIO_PHONE_NUMBER),",
  "          body: 'I could not complete that SMS request. Please try again.',",
  "        });",
  "      } catch (sendError) {",
  "        console.error('Inbound SMS error reply failed:', sendError);",
  "      }",
  "    }",
  "  });",
  "};",
  "",
  "fastify.all('/incoming-sms', handleIncomingSms);",
].join('\n');

if (source.includes(inertRoute)) source = source.replace(inertRoute, activeRoute);

source = source.replace(
  'SMS is paused and WhatsApp is removed. Do not offer these channels.',
  'Ramy may send natural-language executive instructions by SMS. SMS uses the same verified Action Register and Microsoft Graph-backed action logic as voice. WhatsApp remains removed. Do not require special command syntax.'
);

if (!source.includes("fastify.all('/incoming-sms', handleIncomingSms);")) throw new Error('SMS route patch did not apply.');
if (!source.includes('const sendTwilioChannelMessage = async ({ to, from, body })')) throw new Error('SMS sender patch did not apply.');
if (source.includes("SMS is paused and WhatsApp has been removed.")) throw new Error('Disabled SMS sender remains in index.js.');

fs.writeFileSync(indexPath, source);

if (fs.existsSync(testPath)) {
  let tests = fs.readFileSync(testPath, 'utf8');
  tests = tests.replace(
    /test\('messaging routes are removed or inert, voice routes remain',[\s\S]*?\n\}\);/,
    `test('SMS route is active, WhatsApp remains removed, voice routes remain', () => {
  assert.doesNotMatch(source, /fastify\\.all\\('\\/incoming-whatsapp'/);
  assert.match(source, /const handleIncomingSms = async/);
  assert.match(source, /fastify\\.all\\('\\/incoming-sms', handleIncomingSms\\)/);
  assert.match(source, /const sendTwilioChannelMessage = async \\(\\{ to, from, body \\}\\)/);
  assert.match(source, /\\/incoming-call/);
  assert.match(source, /\\/media-stream/);
});`
  );
  fs.writeFileSync(testPath, tests);
}

console.log('SMS patch applied.');
