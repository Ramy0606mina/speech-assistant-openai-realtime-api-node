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

function currentMontrealContext() {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: 'America/Toronto',
    dateStyle: 'full',
    timeStyle: 'short',
  }).format(new Date());
}

function realtimeInstructions() {
  return [
    'You are London Assistant, executive assistant to Ramy Mina for Minaco.',
    `The current Montreal date and time is ${currentMontrealContext()}.`,
    'Ramy is speaking to you by phone.',
    'Speak in polished British English with a calm, mature, discreet executive-assistant manner.',
    'Keep answers concise unless Ramy asks for detail.',
    'Ramy may pause briefly while forming a sentence; do not interrupt unnecessarily.',
    'If he interrupts you, stop promptly and listen.',
    'Never invent current email, calendar, Dropbox, financial, tenant, project, or business facts.',
    'Use the live tools whenever Ramy asks about current email, his calendar, or Dropbox.',
    'The currently connected tools are read-only. Never claim an email was sent or a calendar event was changed from this call.',
    'If Ramy asks for a live action that is not connected, say briefly that the action is not yet connected rather than pretending it was completed.',
    'When asked who you are, say: I am London Assistant, your executive assistant for Minaco.',
    'Ramy is spelled R-A-M-Y.',
  ].join(' ');
}

function voiceTools() {
  return [
    {
      type: 'function',
      name: 'check_email',
      description: 'Read Ramy Mina’s latest live Minaco inbox messages. Use for current inbox or latest email questions.',
      parameters: {
        type: 'object',
        properties: {
          limit: { type: 'integer', minimum: 1, maximum: 10, description: 'Number of latest messages. Default 5.' },
        },
        additionalProperties: false,
      },
    },
    {
      type: 'function',
      name: 'check_calendar',
      description: 'Read Ramy’s live Minaco calendar for a precise date/time window.',
      parameters: {
        type: 'object',
        properties: {
          start_iso: { type: 'string', description: 'Window start as ISO 8601 datetime with timezone offset or Z.' },
          end_iso: { type: 'string', description: 'Window end as ISO 8601 datetime with timezone offset or Z.' },
        },
        required: ['start_iso', 'end_iso'],
        additionalProperties: false,
      },
    },
    {
      type: 'function',
      name: 'search_dropbox',
      description: 'Search London’s controlled LONDON - ACCESS Dropbox workspace for a file or folder by name/topic.',
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Filename, folder name, project, or search term.' },
        },
        required: ['query'],
        additionalProperties: false,
      },
    },
    {
      type: 'function',
      name: 'list_dropbox',
      description: 'List one folder inside London’s controlled LONDON - ACCESS Dropbox workspace.',
      parameters: {
        type: 'object',
        properties: {
          path: { type: 'string', description: 'Path inside LONDON - ACCESS. Use empty string for the root.' },
        },
        additionalProperties: false,
      },
    },
  ];
}

function simplifyEmail(message) {
  return {
    id: message?.id || '',
    subject: message?.subject || '(no subject)',
    from: message?.from?.emailAddress?.address || '',
    senderName: message?.from?.emailAddress?.name || '',
    receivedDateTime: message?.receivedDateTime || '',
    isRead: Boolean(message?.isRead),
    hasAttachments: Boolean(message?.hasAttachments),
    preview: String(message?.bodyPreview || '').slice(0, 700),
  };
}

function simplifyCalendarEvent(event) {
  return {
    id: event?.id || '',
    subject: event?.subject || '(no subject)',
    start: event?.start || null,
    end: event?.end || null,
    location: event?.location?.displayName || '',
    organizer: event?.organizer?.emailAddress?.address || '',
    isAllDay: Boolean(event?.isAllDay),
    isCancelled: Boolean(event?.isCancelled),
  };
}

function simplifyDropboxEntry(entry) {
  const meta = entry?.metadata?.metadata || entry?.metadata || entry || {};
  return {
    type: meta['.tag'] || meta.type || '',
    name: meta.name || '',
    path: meta.path_display || meta.path_lower || meta.path || '',
    size: Number(meta.size || 0),
    modified: meta.server_modified || meta.client_modified || '',
  };
}

async function runVoiceTool(name, args, { graph, dropbox }) {
  if (name === 'check_email') {
    if (!graph) throw new Error('Microsoft Graph is not connected to the voice gateway.');
    const messages = await graph.listPrincipalInbox(args.limit || 5);
    return { success: true, messages: messages.map(simplifyEmail) };
  }

  if (name === 'check_calendar') {
    if (!graph) throw new Error('Microsoft Graph is not connected to the voice gateway.');
    const events = await graph.listPrincipalCalendar({ startIso: args.start_iso, endIso: args.end_iso });
    return { success: true, events: events.map(simplifyCalendarEvent) };
  }

  if (name === 'search_dropbox') {
    if (!dropbox) throw new Error('Dropbox is not connected to the voice gateway.');
    const matches = await dropbox.search(args.query || '');
    return { success: true, matches: matches.slice(0, 20).map(simplifyDropboxEntry) };
  }

  if (name === 'list_dropbox') {
    if (!dropbox) throw new Error('Dropbox is not connected to the voice gateway.');
    const entries = await dropbox.listFolder(args.path || '');
    return { success: true, entries: entries.slice(0, 50).map(simplifyDropboxEntry) };
  }

  throw new Error(`Unsupported voice tool: ${name}`);
}

export function registerVoiceRoutes(app, {
  openAiApiKey,
  principalPhone,
  model = 'gpt-realtime',
  voice = 'marin',
  graph,
  dropbox,
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
      let greetingSent = false;
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

      const sendToolOutput = (callId, output) => {
        sendOpenAi({
          type: 'conversation.item.create',
          item: {
            type: 'function_call_output',
            call_id: callId,
            output: JSON.stringify(output),
          },
        });
        sendOpenAi({
          type: 'response.create',
          response: {
            instructions: 'Answer Ramy concisely using only the verified live tool output. If the tool returned an error, state it plainly.',
          },
        });
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
            tools: voiceTools(),
            tool_choice: 'auto',
          },
        });
        flushAudio();
      });

      openAiWs.on('message', async (data) => {
        try {
          const event = JSON.parse(String(data));

          if (event.type === 'session.updated' && !greetingSent) {
            greetingSent = true;
            sendOpenAi({
              type: 'response.create',
              response: {
                instructions: 'Greet Ramy briefly as London and ask how you can help. One short sentence.',
              },
            });
          }

          if (event.type === 'response.output_audio.delta' && event.delta && streamSid) {
            sendTwilio({ event: 'media', streamSid, media: { payload: event.delta } });
          }

          if (event.type === 'input_audio_buffer.speech_started' && streamSid) {
            sendTwilio({ event: 'clear', streamSid });
          }

          if (event.type === 'response.function_call_arguments.done') {
            try {
              const args = JSON.parse(event.arguments || '{}');
              const output = await runVoiceTool(event.name, args, { graph, dropbox });
              sendToolOutput(event.call_id, output);
            } catch (error) {
              logger.error?.({ err: error, tool: event.name }, 'London voice tool failed');
              sendToolOutput(event.call_id, { success: false, error: error.message });
            }
            return;
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
