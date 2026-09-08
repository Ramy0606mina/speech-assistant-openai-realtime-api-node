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
import { LondonCore } from './src/london-core.js';
import { registerVoiceRoutes } from './src/voice-gateway.js';
import { MorningBrief } from './src/morning-brief.js';
import { SmsClient } from './src/sms-client.js';
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
const london = new LondonCore({ graph, openai, dropbox, state, deliveryGuard, sms, logger: app.log });
const urgentAlerts=new UrgentAlerts({graph,openai,sms,guard:new DeliveryGuard(dropbox,`${graph.principalMailbox}:urgent-alerts`)});
const smsConversation = new SmsConversation({sms,openai,graph,dropbox,state,guard:new DeliveryGuard(dropbox,`${graph.principalMailbox}:incoming-sms`),logger:app.log});
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
let deliveryGuardReady = false;
const morningBrief = new MorningBrief({graph,openai,dropbox,guard:deliveryGuard});
const morningBriefEnabled = process.env.LONDON_MORNING_BRIEF_ENABLED !== 'false';
async function safeMorningBrief() {
  if (!morningBriefEnabled || !deliveryGuardReady) return;
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

let pollInFlight = false;
async function safePoll() {
  if (pollInFlight) return;
  pollInFlight = true;
  try {
    if (!deliveryGuardReady) { await deliveryGuard.initialize(); deliveryGuardReady = true; }
    const result = await london.pollOnce(config.runtime.pollBatchSize);
    app.log.info({ checked: result.checked }, 'London mailbox poll complete');
  } catch (error) {
    app.log.error({ err: error }, 'London mailbox poll failed');
  } finally {
    pollInFlight = false;
  }
}

app.get('/health', async () => ({
  ok: true,
  service: 'London Assistant',
  architecture: 'lean-single-backend',
  powerAutomateRequired: false,
  revision: process.env.RENDER_GIT_COMMIT || null,
  sms: sms.configured ? 'owner-requested' : 'not-configured',
  smsConversation: {enabled:smsConversationEnabled,configured:sms.configured,ready:smsConversation.ready,mode:'two-way-owner-only',transport:'webhook-with-inbox-polling',webhook:smsWebhookStatus,intervalSeconds:15,lastCheckedAt:smsConversation.lastCheckedAt,lastOutcome:smsConversation.lastOutcome,lastReplyStatus:smsConversation.lastReplyStatus},
  urgentEmailAlerts: {configured:sms.configured,ready:urgentAlerts.ready,newMessagesOnly:true},
  whatsapp: 'removed',
  pendingDeliveryReview: Object.values(state.state.processedMessages).filter(item => item.result === 'delivery-pending-review').length,
  durableDeliveryGuard: deliveryGuardReady,
  historicalRequestsHeld: Object.values(state.state.processedMessages).filter(item => item.result === 'historical-review').length,
  morningBrief: {enabled:morningBriefEnabled,time:'07:30',timeZone:'America/Toronto',cadence:'daily',catchUpUntil:'12:00',lastOutcome:morningBrief.lastOutcome},
  ...configurationStatus(config),
  lastPollAt: state.state.lastPollAt,
  time: new Date().toISOString(),
}));

app.get('/health/deep', async (request, reply) => {
  if (!config.runtime.healthSecret || request.headers['x-london-health-secret'] !== config.runtime.healthSecret) {
    return reply.code(401).send({ ok: false, error: 'Unauthorized.' });
  }
  return { ok: true, architecture: 'lean-single-backend', powerAutomateRequired: false, ...configurationStatus(config) };
});

app.post('/internal/poll-once', async (request, reply) => {
  if (!config.runtime.healthSecret || request.headers['x-london-health-secret'] !== config.runtime.healthSecret) {
    return reply.code(401).send({ ok: false, error: 'Unauthorized.' });
  }
  if (!deliveryGuardReady) return reply.code(503).send({ ok: false, error: 'Delivery ledger is not ready.' });
  if (pollInFlight) return reply.code(409).send({ ok: false, error: 'Mailbox poll already running.' });
  await safePoll();
  return { ok: true };
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
