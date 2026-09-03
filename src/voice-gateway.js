import WebSocket from 'ws';
import { randomUUID } from 'node:crypto';

export function normalizePhone(value) {
  const raw = String(value || '').trim();
  const digits = raw.replace(/\D/g, '');
  if (!digits) return '';
  if (digits.length === 10) return `+1${digits}`;
  return `+${digits}`;
}

export function isAuthorizedCaller(caller, principalPhone) {
  const a = normalizePhone(caller);
  const b = normalizePhone(principalPhone);
  return Boolean(a && b && a === b);
}

export function buildIncomingCallTwiML({ host, streamToken }) {
  const safeHost = String(host || '').trim();
  const safeToken = encodeURIComponent(String(streamToken || '').trim());
  return `<?xml version="1.0" encoding="UTF-8"?>
<Response>
  <Say voice="Google.en-US-Chirp3-HD-Aoede">Please wait while I connect you to London Assistant.</Say>
  <Pause length="1"/>
  <Connect>
    <Stream url="wss://${safeHost}/media-stream/${safeToken}" />
  </Connect>
</Response>`;
}

function realtimeInstructions() {
  return [
    'You are London Assistant, executive assistant to Ramy Mina for Minaco.',
    'Ramy is speaking to you by phone.',
    'Speak in polished British English with a calm, mature, discreet executive-assistant manner.',
    'Keep answers concise unless Ramy asks for detail.',
    'Ramy may pause briefly while forming a sentence; do not interrupt unnecessarily.',
    'If he interrupts you, stop promptly and listen.',
    'Never invent current email, calendar, Dropbox, financial, tenant, project, or business facts.',
    'This restored voice gateway is conversational only while London tools are being reconnected to the lean core.',
    'If Ramy asks you to perform a live action that is not available in this call, say briefly that the live action is not yet connected rather than pretending it was completed.',
    'When asked who you are, say: I am London Assistant, your executive assistant for Minaco.',
    'Ramy is spelled R-A-M-Y.',
  ].join(' ');
}

export function registerVoiceRoutes(app, {
  openAiApiKey,
  principalPhone,
  model = 'gpt-realtime',
  voice = 'marin',
  logger = console,
} = {}) {
  const authorizedStreamTokens = new Map();

  app.all('/incoming-call', async (request, reply) => {
    const caller = request.body?.From || request.query?.From || '';
    const authorized = isAuthorizedCaller(caller, principalPhone);
    logger.info?.({ caller: normalizePhone(caller), authorized }, 'London call security check');

    if (!authorized) {
      return reply.type('text/xml').send(
        '<?xml version="1.0" encoding="UTF-8"?><Response><Say>Sorry, this line is private.</Say><Hangup/></Response>'
      );
    }

    if (!openAiApiKey) {
      return reply.type('text/xml').send(
        '<?xml version="1.0" encoding="UTF-8"?><Response><Say>London is temporarily unavailable.</Say><Hangup/></Response>'
      );
    }

    const streamToken = randomUUID();
    authorizedStreamTokens.set(streamToken, Date.now() + 2 * 60 * 1000);
    const timer = setTimeout(() => authorizedStreamTokens.delete(streamToken), 2 * 60 * 1000);
    timer.unref?.();

    const host = String(request.headers['x-forwarded-host'] || request.headers.host || '')
      .split(',')[0]
      .trim();
    return reply.type('text/xml').send(buildIncomingCallTwiML({ host, streamToken }));
  });

  app.register(async (instance) => {
    instance.get('/media-stream/:streamToken', { websocket: true }, (connection, req) => {
      const token = String(req.params?.streamToken || '');
      const expiry = authorizedStreamTokens.get(token);
      if (!expiry || expiry < Date.now()) {
        authorizedStreamTokens.delete(token);
        try { connection.close(1008, 'Unauthorized'); } catch { connection.close(); }
        return;
      }
      authorizedStreamTokens.delete(token);

      let streamSid = '';
      let closed = false;
      let openAiReady = false;
      const pendingAudio = [];

      const openAiWs = new WebSocket(
        `wss://api.openai.com/v1/realtime?model=${encodeURIComponent(model)}`,
        {
          headers: {
            Authorization: `Bearer ${openAiApiKey}`,
            'OpenAI-Beta': 'realtime=v1',
          },
        }
      );

      const sendOpenAi = (event) => {
        if (openAiWs.readyState !== WebSocket.OPEN) return false;
        openAiWs.send(JSON.stringify(event));
        return true;
      };

      const sendTwilio = (event) => {
        if (connection.readyState !== WebSocket.OPEN) return false;
        connection.send(JSON.stringify(event));
        return true;
      };

      const flushAudio = () => {
        while (pendingAudio.length && openAiWs.readyState === WebSocket.OPEN) {
          sendOpenAi({ type: 'input_audio_buffer.append', audio: pendingAudio.shift() });
        }
      };

      openAiWs.on('open', () => {
        openAiReady = true;
        sendOpenAi({
          type: 'session.update',
          session: {
            type: 'realtime',
            model,
            output_modalities: ['audio'],
            audio: {
              input: {
                format: { type: 'audio/pcmu' },
                transcription: { model: 'gpt-4o-mini-transcribe' },
                turn_detection: {
                  type: 'semantic_vad',
                  eagerness: 'high',
                  create_response: true,
                  interrupt_response: true,
                },
              },
              output: {
                format: { type: 'audio/pcmu' },
                voice,
              },
            },
            instructions: realtimeInstructions(),
          },
        });
        flushAudio();
        sendOpenAi({
          type: 'response.create',
          response: {
            instructions: 'Greet Ramy briefly as London and ask how you can help. One short sentence.',
          },
        });
      });

      openAiWs.on('message', (data) => {
        try {
          const event = JSON.parse(String(data));
          if (event.type === 'response.output_audio.delta' && event.delta && streamSid) {
            sendTwilio({ event: 'media', streamSid, media: { payload: event.delta } });
          }

          if (event.type === 'input_audio_buffer.speech_started' && streamSid) {
            sendTwilio({ event: 'clear', streamSid });
          }

          if (event.type === 'error') {
            logger.error?.({ error: event.error || event }, 'OpenAI realtime voice error');
          }
        } catch (error) {
          logger.error?.({ err: error }, 'London voice event parse error');
        }
      });

      openAiWs.on('error', (error) => {
        logger.error?.({ err: error }, 'OpenAI realtime websocket error');
      });

      openAiWs.on('close', () => {
        openAiReady = false;
        if (!closed) {
          try { connection.close(); } catch {}
        }
      });

      connection.on('message', (data) => {
        try {
          const event = JSON.parse(String(data));
          if (event.event === 'start') {
            streamSid = String(event.start?.streamSid || event.streamSid || '');
            return;
          }
          if (event.event === 'media' && event.media?.payload) {
            if (openAiReady && openAiWs.readyState === WebSocket.OPEN) {
              sendOpenAi({ type: 'input_audio_buffer.append', audio: event.media.payload });
            } else if (pendingAudio.length < 150) {
              pendingAudio.push(event.media.payload);
            }
            return;
          }
          if (event.event === 'stop') {
            closed = true;
            if (openAiWs.readyState === WebSocket.OPEN || openAiWs.readyState === WebSocket.CONNECTING) {
              openAiWs.close();
            }
          }
        } catch (error) {
          logger.error?.({ err: error }, 'Twilio media event parse error');
        }
      });

      connection.on('close', () => {
        closed = true;
        if (openAiWs.readyState === WebSocket.OPEN || openAiWs.readyState === WebSocket.CONNECTING) {
          openAiWs.close();
        }
      });
    });
  });
}
