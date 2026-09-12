import Fastify from 'fastify';
import fastifyFormBody from '@fastify/formbody';
import fastifyWs from '@fastify/websocket';
import dotenv from 'dotenv';
import { loadConfig, configurationStatus } from './src/config.js';
import { OpenAIClient } from './src/openai-client.js';
import { MicrosoftGraphClient } from './src/microsoft-graph.js';
import { DropboxClient } from './src/dropbox-client.js';
import { StateStore } from './src/state-store.js';
import { DeliveryGuard } from './src/delivery-guard.js';
import { MailboxWorker } from './src/mailbox-worker.js';
import { LondonCore } from './src/london-core.js';
import { registerVoiceRoutes } from './src/voice-gateway.js';
import { MorningBrief } from './src/morning-brief.js';
import { SmsClient } from './src/sms-client.js';
import { TextMeetings } from './src/text-meetings.js';
import { TextReminders } from './src/text-reminders.js';
import { UrgentAlerts } from './src/urgent-alerts.js';
import { SmsConversation, registerSmsWebhook } from './src/sms-conversation.js';

dotenv.config();
const config = loadConfig();
const app = Fastify({ logger: true });

await app.register(fastifyFormBody);
await app.register(fastifyWs);

const graph = new MicrosoftGraphClient({ ...config.microsoft });
const openai = new OpenAIClient({ apiKey: config.openai.apiKey, model: config.openai.taskModel });
const dropbox = new DropboxClient(config.dropbox);
const state = new StateStore(config.runtime.stateFile);
const deliveryGuard = new DeliveryGuard(dropbox, graph.readMailbox);
const sms = new SmsClient({accountSid:process.env.TWILIO_ACCOUNT_SID,authToken:process.env.TWILIO_AUTH_TOKEN,from:process.env.TWILIO_PHONE_NUMBER,to:config.voice.principalPhone});
const meetings = new TextMeetings({graph,dropbox,openai});
const reminders = new TextReminders({graph,dropbox,openai});
const london = new LondonCore({ graph, openai, dropbox, state, deliveryGuard, sms, meetings, logger: app.log });
const urgentAlerts=new UrgentAlerts({graph,openai,sms,guard:new DeliveryGuard(dropbox,`${graph.principalMailbox}:urgent-alerts`)});
const smsConversation = new SmsConversation({sms,openai,graph,dropbox,state,meetings,reminders,guard:new DeliveryGuard(dropbox,`${graph.principalMailbox}:incoming-sms`),logger:app.log});
// Do not activate owner texts in PR previews that share production credentials.
const smsConversationEnabled = process.env.LONDON_SMS_CONVERSATION_ENABLED === undefined
  ? process.env.RENDER_EXTERNAL_URL === 'https://london-ai-pr-1.onrender.com'
  : process.env.LONDON_SMS_CONVERSATION_ENABLED === 'true';
let smsWebhookStatus = 'not-initialized';
async function safeSmsConversation(){if(!smsConversationEnabled)return;try{await smsConversation.tick();}catch(error){app.log.error({status:error.status||null},'Owner SMS conversation failed');}}
async function initializeSmsWebhook(){
  if(!smsConversationEnabled || !sms.configured)return;
  try{smsWebhookStatus=await sms.configureInbound(process.env.RENDER_EXTERNAL_URL);}
  catch(error){smsWebhookStatus='configuration-needs-review';app.log.error({status:error.status||null},'SMS webhook configuration failed; inbox polling remains available');}
}
registerSmsWebhook(app,{sms,publicUrl:process.env.RENDER_EXTERNAL_URL,onMessage:safeSmsConversation});
async function safeUrgentAlerts(){try{await urgentAlerts.tick();}catch(error){app.log.error({err:error},'Urgent email alert failed');}}
const mailboxWorker = new MailboxWorker({ graph, london, guard: deliveryGuard, limit: config.runtime.pollBatchSize, logger: app.log });
const morningBrief = new MorningBrief({graph,openai,dropbox,guard:deliveryGuard});
const morningBriefEnabled = process.env.LONDON_MORNING_BRIEF_ENABLED !== 'false';
async function safeMorningBrief() {
  if (!morningBriefEnabled || !mailboxWorker.guardReady) return;
  try { await morningBrief.tick(); } catch(error) { app.log.error({err:error},'Morning executive report failed'); }
}

registerVoiceRoutes(app, {
  openAiApiKey: config.openai.apiKey,
  principalPhone: config.voice.principalPhone,
  twilioAuthToken: process.env.TWILIO_AUTH_TOKEN,
  publicUrl: process.env.RENDER_EXTERNAL_URL,
  model: config.voice.model,
  voice: config.voice.voice,
  graph,
  dropbox,
  logger: app.log,
});

async function safePoll() {
  const result = await mailboxWorker.poll();
  if (!result.busy) app.log[result.ok ? 'info' : 'error'](result, 'London mailbox poll outcome');
  return result;
}

function operational() {
  const checkedAt = Date.parse(mailboxWorker.status.lastCheckedAt);
  return mailboxWorker.status.state === 'ready' && openai.status.state === 'available'
    && Date.now() - checkedAt < Math.max(config.runtime.pollIntervalMs * 3, 180000);
}

app.get('/health', async () => ({
  smsReminders: {enabled:smsConversationEnabled,provider:'Microsoft Graph',calendar:'primary',lastOutcome:reminders.lastOutcome},
  ok: operational(),
  live: true,
  mailbox: mailboxWorker.status,
  openai: openai.status,
  service: 'London Assistant',
  architecture: 'lean-single-backend',
  powerAutomateRequired: false,
  revision: process.env.RENDER_GIT_COMMIT || null,
  sms: sms.configured ? 'owner-requested' : 'not-configured',
  smsConversation: {enabled:smsConversationEnabled,configured:sms.configured,ready:smsConversation.ready,mode:'two-way-owner-only',transport:'webhook-with-inbox-polling',webhook:smsWebhookStatus,intervalSeconds:15,lastCheckedAt:smsConversation.lastCheckedAt,lastOutcome:smsConversation.lastOutcome,lastReplyStatus:smsConversation.lastReplyStatus},
  urgentEmailAlerts: {configured:sms.configured,ready:urgentAlerts.ready,newMessagesOnly:true,heldForReview:urgentAlerts.heldForReview||0},
  textMeetings: {email:true,sms:smsConversationEnabled,confirmation:'explicit-proposal-code',provider:'Microsoft Teams'},
  whatsapp: 'removed',
  pendingDeliveryReview: Object.values(state.state.processedMessages).filter(item => item.result === 'delivery-pending-review').length,
  durableDeliveryGuard: mailboxWorker.guardReady,
  historicalRequestsHeld: Object.values(state.state.processedMessages).filter(item => item.result === 'historical-review').length,
  morningBrief: {enabled:morningBriefEnabled,time:'07:30',timeZone:'America/Toronto',cadence:'weekdays',catchUpUntil:'12:00',lastOutcome:morningBrief.lastOutcome},
  ...configurationStatus(config),
  lastPollAt: state.state.lastPollAt,
  time: new Date().toISOString(),
}));

// Keep /health HTTP 200 for Render liveness: a provider outage must not cause
// restart loops. Readiness reflects actual worker outcomes, not key presence.
app.get('/ready', async (_request, reply) => reply.code(operational() ? 200 : 503).send({
  ok: operational(), mailbox: mailboxWorker.status, openai: openai.status,
}));

app.get('/health/deep', async (request, reply) => {
  if (!config.runtime.healthSecret || request.headers['x-london-health-secret'] !== config.runtime.healthSecret) {
    return reply.code(401).send({ ok: false, error: 'Unauthorized.' });
  }
  return { ok: operational(), mailbox: mailboxWorker.status, openai: openai.status, architecture: 'lean-single-backend', powerAutomateRequired: false, ...configurationStatus(config) };
});

app.post('/internal/poll-once', async (request, reply) => {
  if (!config.runtime.healthSecret || request.headers['x-london-health-secret'] !== config.runtime.healthSecret) {
    return reply.code(401).send({ ok: false, error: 'Unauthorized.' });
  }
  if (mailboxWorker.running) return reply.code(409).send({ ok: false, error: 'Mailbox poll already running.' });
  const result = await safePoll();
  return reply.code(result.ok ? 200 : 503).send(result);
});

const port = Number(process.env.PORT || 3000);
await app.listen({ port, host: '0.0.0.0' });

const pollTimer = setInterval(safePoll, config.runtime.pollIntervalMs);
pollTimer.unref?.();
setTimeout(safePoll, 1500).unref?.();
setInterval(safeMorningBrief,60000).unref?.();
setInterval(safeUrgentAlerts,60000).unref?.();
setTimeout(safeUrgentAlerts,2500).unref?.();
setInterval(safeSmsConversation,15000).unref?.();
setTimeout(safeSmsConversation,3000).unref?.();
setTimeout(initializeSmsWebhook,3500).unref?.();
