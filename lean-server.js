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
const london = new LondonCore({ graph, openai, dropbox, state, deliveryGuard, logger: app.log });
let deliveryGuardReady = false;

registerVoiceRoutes(app, {
  openAiApiKey: config.openai.apiKey,
  principalPhone: config.voice.principalPhone,
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
  sms: 'paused',
  whatsapp: 'removed',
  pendingDeliveryReview: Object.values(state.state.processedMessages).filter(item => item.result === 'delivery-pending-review').length,
  durableDeliveryGuard: deliveryGuardReady,
  historicalRequestsHeld: Object.values(state.state.processedMessages).filter(item => item.result === 'historical-review').length,
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
