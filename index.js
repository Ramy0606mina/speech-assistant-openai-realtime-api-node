import Fastify from 'fastify';
import WebSocket from 'ws';
import dotenv from 'dotenv';
import fastifyFormBody from '@fastify/formbody';
import fastifyWs from '@fastify/websocket';

// Load environment variables from .env file
dotenv.config();

// Retrieve the OpenAI API key from environment variables.
const {
  OPENAI_API_KEY,
  TWILIO_ACCOUNT_SID,
  TWILIO_AUTH_TOKEN,
  TWILIO_PHONE_NUMBER,
  RAMY_PHONE_NUMBER,
  MS_TENANT_ID,
  MS_CLIENT_ID,
  MS_CLIENT_SECRET,
  RAMY_MINACO_EMAIL,
  ACTIONS_MS_TENANT_ID,
  ACTIONS_MS_CLIENT_ID,
  ACTIONS_MS_CLIENT_SECRET,
  LONDON_MINACO_EMAIL,
} = process.env;

const sendSmsToRamy = async (message) => {
  if (
    !TWILIO_ACCOUNT_SID ||
    !TWILIO_AUTH_TOKEN ||
    !TWILIO_PHONE_NUMBER ||
    !RAMY_PHONE_NUMBER
  ) {
    throw new Error('Missing Twilio SMS environment variables.');
  }

  const auth = Buffer.from(
    `${TWILIO_ACCOUNT_SID}:${TWILIO_AUTH_TOKEN}`
  ).toString('base64');

  const form = new URLSearchParams({
    To: RAMY_PHONE_NUMBER,
    From: TWILIO_PHONE_NUMBER,
    Body: message
  });

  const response = await fetch(
    `https://api.twilio.com/2010-04-01/Accounts/${TWILIO_ACCOUNT_SID}/Messages.json`,
    {
      method: 'POST',
      headers: {
        Authorization: `Basic ${auth}`,
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      body: form.toString()
    }
  );

  const data = await response.json();

  if (!response.ok) {
    throw new Error(`Twilio SMS failed: ${data.message || response.status}`);
  }

  return data;
};
const getMicrosoftGraphToken = async () => {
  if (!MS_TENANT_ID || !MS_CLIENT_ID || !MS_CLIENT_SECRET) {
    throw new Error('Missing Microsoft Graph environment variables.');
  }

  const form = new URLSearchParams({
    client_id: MS_CLIENT_ID,
    client_secret: MS_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });

  const response = await fetch(
    `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      body: form.toString()
    }
  );

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Microsoft authentication failed: ${data.error_description || data.error || response.status}`
    );
  }

  return data.access_token;
};
const getMicrosoftGraphActionsToken = async () => {
  if (
    !ACTIONS_MS_TENANT_ID ||
    !ACTIONS_MS_CLIENT_ID ||
    !ACTIONS_MS_CLIENT_SECRET
  ) {
    throw new Error('Missing Microsoft Graph Actions environment variables.');
  }

  const form = new URLSearchParams({
    client_id: ACTIONS_MS_CLIENT_ID,
    client_secret: ACTIONS_MS_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });

  const response = await fetch(
    `https://login.microsoftonline.com/${ACTIONS_MS_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      body: form.toString()
    }
  );

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Microsoft Actions authentication failed: ${
        data.error_description || data.error || response.status
      }`
    );
  }

  return data.access_token;
};
const getRecentMinacoEmails = async (limit = 5) => {
  if (!RAMY_MINACO_EMAIL) {
    throw new Error('RAMY_MINACO_EMAIL is not configured.');
  }

  const token = await getMicrosoftGraphToken();

  const url = new URL(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      RAMY_MINACO_EMAIL
    )}/mailFolders/inbox/messages`
  );

  url.searchParams.set('$top', String(limit));
  url.searchParams.set(
    '$select',
    'subject,from,receivedDateTime,bodyPreview,isRead'
  );
  url.searchParams.set('$orderby', 'receivedDateTime desc');

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph email lookup failed: ${data.error?.message || response.status}`
    );
  }

  return data.value || [];
};
// Initialize Fastify
const fastify = Fastify();
fastify.register(fastifyFormBody);
fastify.register(fastifyWs);

// Constants
const SYSTEM_MESSAGE = `
You are London Assistant, the executive assistant to Ramy Mina.

IDENTITY AND BUSINESS CONTEXT

You operate primarily for Minaco / Mina Group.

Mina Capital was previously used as a technical test environment for your AI systems. Do not assume Mina Capital is the company Ramy is referring to unless he specifically says Mina Capital.

Verified Minaco identities:
- Ramy Mina: principal executive
- Ramy's primary Minaco email: ramy.mina@minaco.ca
- London Assistant email: london@minaco.ca
- Minaco accounting email: accounting@minaco.ca
- London Assistant phone number: +1 438-255-9099

When Ramy asks "what is my email address?" without specifying another company, answer:
ramy.mina@minaco.ca

When Ramy asks who you are, say:
"I am London Assistant, your executive assistant for Minaco."

ACCURACY IS YOUR HIGHEST PRIORITY

Never invent or guess facts about:
- Ramy
- Minaco
- Mina Capital
- Mina Group
- employees
- consultants
- contractors
- lawyers or legal counsel
- partners
- investors
- lenders
- tenants
- projects
- properties
- emails
- meetings
- financial information
- contracts
- deadlines
- business relationships

If information is not explicitly provided in your verified context or retrieved from an authorized live system, say clearly that you do not have verified information.

Examples:
"I don't have a verified legal counsel recorded for Minaco."
"I don't have live access to that information through this voice connection yet."
"I would need to check the connected Minaco system before answering that."

Never make up a plausible answer just to be helpful.

LIVE DATA RULE

Current emails, calendar events, messages, tasks, project status, invoices, and other changing business information must come from connected systems.

Do not pretend you checked email, calendar, SMS, accounting, or another system unless you actually used a tool that retrieved that information.

When the check_email tool is available, use it whenever Ramy asks about current emails, his inbox, whether someone emailed him, or what requires his attention. Never answer a current email question from memory. Only say live email access is unavailable if the tool is unavailable or returns an error.

Do not answer a current-state question from general model knowledge.

TOOLS

You may use the send_sms tool only when Ramy explicitly asks you to send him a text message.

Do not claim that an SMS was sent unless the tool confirms success.

Do not claim access to tools that are not actually available in the current session.

EXECUTIVE ASSISTANT BEHAVIOR

Your role is to reduce Ramy's workload.

Be:
- concise
- practical
- commercially aware
- organized
- calm
- professional
- proactive when appropriate

Prioritize:
1. decisions requiring Ramy's attention
2. deadlines and risks
3. financial or contractual consequences
4. commitments owed to Minaco
5. follow-ups
6. routine information

Do not overwhelm Ramy with unnecessary detail.

If a question is ambiguous and the answer would materially differ depending on which company, property, project, or person he means, ask a short clarification question instead of guessing.

AUTHORITY

You may provide information, summarize, organize, remind, and send an SMS to Ramy when explicitly requested.

Do not claim to have approved payments, signed contracts, committed Minaco to pricing, settled disputes, or made legal or financial decisions unless an authorized system actually performed that action.

PHONE CONVERSATION STYLE

You are speaking with Ramy by phone.

Speak naturally.
Keep most answers short unless Ramy asks for detail.
Do not recite long disclaimers.
If you do not know something, say so simply and accurately.
Accuracy is more important than sounding helpful.
`;
const VOICE = 'marin';
const TEMPERATURE = 0.3; // Controls the randomness of the AI's responses
const PORT = process.env.PORT || 5050; // Allow dynamic port assignment

// List of Event Types to log to the console. See the OpenAI Realtime API Documentation: https://platform.openai.com/docs/api-reference/realtime
const LOG_EVENT_TYPES = [
    'error',
    'response.content.done',
    'rate_limits.updated',
    'response.done',
    'input_audio_buffer.committed',
    'input_audio_buffer.speech_stopped',
    'input_audio_buffer.speech_started',
    'session.created',
    'session.updated'
];

// Show AI response elapsed timing calculations
const SHOW_TIMING_MATH = false;

// Root Route
fastify.get('/', async (request, reply) => {
    reply.send({ message: 'Twilio Media Stream Server is running!' });
});

// Route for Twilio to handle incoming calls
// <Say> punctuation to improve text-to-speech translation
fastify.all('/incoming-call', async (request, reply) => {
  const caller = request.body?.From || request.query?.From;
console.log('CALL SECURITY CHECK:', { caller, authorized: caller === RAMY_PHONE_NUMBER });
if (!caller || caller !== RAMY_PHONE_NUMBER) {
  const deniedResponse = `<?xml version="1.0" encoding="UTF-8"?>
<Response>
  <Say>Sorry, this line is private.</Say>
  <Hangup/>
</Response>`;

  return reply.type('text/xml').send(deniedResponse);
}  
  const twimlResponse = `<?xml version="1.0" encoding="UTF-8"?>
                          <Response>
                              <Say voice="Google.en-US-Chirp3-HD-Aoede">Please wait while we connect your call to the A. I. voice assistant, powered by Twilio and the Open A I Realtime API</Say>
                              <Pause length="1"/>
                              <Say voice="Google.en-US-Chirp3-HD-Aoede">O.K. you can start talking!</Say>
                              <Connect>
                                  <Stream url="wss://${request.headers.host}/media-stream" />
                              </Connect>
                          </Response>`;

    reply.type('text/xml').send(twimlResponse);
});

// WebSocket route for media-stream
fastify.register(async (fastify) => {
    fastify.get('/media-stream', { websocket: true }, (connection, req) => {
        console.log('Client connected');

        // Connection-specific state
        let streamSid = null;
        let latestMediaTimestamp = 0;
        let lastAssistantItem = null;
        let markQueue = [];
        let responseStartTimestampTwilio = null;

        const openAiWs = new WebSocket(`wss://api.openai.com/v1/realtime?model=gpt-realtime&temperature=${TEMPERATURE}`, {
            headers: {
                Authorization: `Bearer ${OPENAI_API_KEY}`,
            }
        });

        // Control initial session with OpenAI
        const initializeSession = () => {
            const sessionUpdate = {
                type: 'session.update',
                session: {
                    type: 'realtime',
                    model: "gpt-realtime",
                    output_modalities: ["audio"],
                    audio: {
                        input: { format: { type: 'audio/pcmu' }, turn_detection: { type: "server_vad" } },
                        output: { format: { type: 'audio/pcmu' }, voice: VOICE },
                    },
                    instructions: SYSTEM_MESSAGE,
                  tools: [
  {
    type: 'function',
    name: 'send_sms',
    description: 'Send an SMS message to Ramy when Ramy explicitly asks London to text him.',
    parameters: {
      type: 'object',
      properties: {
        message: {
          type: 'string',
          description: 'The exact SMS message to send to Ramy.'
        }
      },
      required: ['message'],
      additionalProperties: false
    }
  },
  {
    type: 'function',
    name: 'check_email',
    description: 'Read Ramy Mina’s live Minaco inbox. Use this whenever Ramy asks about current emails, latest emails, unread emails, or whether someone emailed him.',
    parameters: {
      type: 'object',
      properties: {
        limit: {
          type: 'integer',
          minimum: 1,
          maximum: 10,
          description: 'Number of recent emails to retrieve. Use 5 unless Ramy asks otherwise.'
        }
      },
      additionalProperties: false
    }
  }
],
tool_choice: 'auto',
  
                },
            };

            console.log('Sending session update:', JSON.stringify(sessionUpdate));
            openAiWs.send(JSON.stringify(sessionUpdate));

            // Uncomment the following line to have AI speak first:
            // sendInitialConversationItem();
        };

        // Send initial conversation item if AI talks first
        const sendInitialConversationItem = () => {
            const initialConversationItem = {
                type: 'conversation.item.create',
                item: {
                    type: 'message',
                    role: 'user',
                    content: [
                        {
                            type: 'input_text',
                            text: 'Greet the user with "Hello there! I am an AI voice assistant powered by Twilio and the OpenAI Realtime API. You can ask me for facts, jokes, or anything you can imagine. How can I help you?"'
                        }
                    ]
                }
            };

            if (SHOW_TIMING_MATH) console.log('Sending initial conversation item:', JSON.stringify(initialConversationItem));
            openAiWs.send(JSON.stringify(initialConversationItem));
            openAiWs.send(JSON.stringify({ type: 'response.create' }));
        };

        // Handle interruption when the caller's speech starts
        const handleSpeechStartedEvent = () => {
            if (markQueue.length > 0 && responseStartTimestampTwilio != null) {
                const elapsedTime = latestMediaTimestamp - responseStartTimestampTwilio;
                if (SHOW_TIMING_MATH) console.log(`Calculating elapsed time for truncation: ${latestMediaTimestamp} - ${responseStartTimestampTwilio} = ${elapsedTime}ms`);

                if (lastAssistantItem) {
                    const truncateEvent = {
                        type: 'conversation.item.truncate',
                        item_id: lastAssistantItem,
                        content_index: 0,
                        audio_end_ms: elapsedTime
                    };
                    if (SHOW_TIMING_MATH) console.log('Sending truncation event:', JSON.stringify(truncateEvent));
                    openAiWs.send(JSON.stringify(truncateEvent));
                }

                connection.send(JSON.stringify({
                    event: 'clear',
                    streamSid: streamSid
                }));

                // Reset
                markQueue = [];
                lastAssistantItem = null;
                responseStartTimestampTwilio = null;
            }
        };

        // Send mark messages to Media Streams so we know if and when AI response playback is finished
        const sendMark = (connection, streamSid) => {
            if (streamSid) {
                const markEvent = {
                    event: 'mark',
                    streamSid: streamSid,
                    mark: { name: 'responsePart' }
                };
                connection.send(JSON.stringify(markEvent));
                markQueue.push('responsePart');
            }
        };

        // Open event for OpenAI WebSocket
        openAiWs.on('open', () => {
            console.log('Connected to the OpenAI Realtime API');
            setTimeout(initializeSession, 100);
        });

        // Listen for messages from the OpenAI WebSocket (and send to Twilio if necessary)
        openAiWs.on('message', async (data) => {
            try {
                const response = JSON.parse(data);
             if (
  response.type === 'response.function_call_arguments.done' &&
  response.name === 'check_email'
) {
  let toolResult;

  try {
    const args = JSON.parse(response.arguments || '{}');

    const requestedLimit = Number(args.limit);
    const limit =
      Number.isInteger(requestedLimit) && requestedLimit >= 1
        ? Math.min(requestedLimit, 10)
        : 5;

    const emails = await getRecentMinacoEmails(limit);

    const simplifiedEmails = emails.map((email) => ({
      from:
        email.from?.emailAddress?.name ||
        email.from?.emailAddress?.address ||
        'Unknown sender',
      fromEmail: email.from?.emailAddress?.address || '',
      subject: email.subject || '(No subject)',
      receivedTimeMontreal: new Intl.DateTimeFormat('en-US', {
  timeZone: 'America/Toronto',
  dateStyle: 'medium',
  timeStyle: 'short'
}).format(new Date(email.receivedDateTime)),
      isRead: Boolean(email.isRead),
      preview: email.bodyPreview || ''
    }));

    toolResult = JSON.stringify({
      success: true,
      mailbox: RAMY_MINACO_EMAIL,
      emails: simplifiedEmails
    });
  } catch (error) {
    console.error('Email lookup error:', error);

    toolResult = JSON.stringify({
      success: false,
      error: error.message
    });
  }

  openAiWs.send(JSON.stringify({
    type: 'conversation.item.create',
    item: {
      type: 'function_call_output',
      call_id: response.call_id,
      output: toolResult
    }
  }));

  openAiWs.send(JSON.stringify({
    type: 'response.create',
    response: {
      instructions:
       
  'Answer Ramy using only the live Minaco email results returned by the tool. All email times are already converted to Montreal local time in receivedTimeMontreal. Use that time when mentioning when an email arrived. Be concise and practical. Mention sender, subject, and what matters. Do not invent anything. If the lookup failed, tell Ramy the live email lookup failed.'
    }
  }));

  return;
}
              if (
  response.type === 'response.function_call_arguments.done' &&
  response.name === 'send_sms'
) {
  let toolResult;

  try {
    const args = JSON.parse(response.arguments || '{}');

    if (!args.message || typeof args.message !== 'string') {
      throw new Error('No SMS message was provided.');
    }

    await sendSmsToRamy(args.message);

    toolResult = JSON.stringify({
      success: true,
      message: 'SMS sent successfully to Ramy.'
    });
  } catch (error) {
    console.error('SMS error:', error);

    toolResult = JSON.stringify({
      success: false,
      error: error.message
    });
  }

  openAiWs.send(JSON.stringify({
    type: 'conversation.item.create',
    item: {
      type: 'function_call_output',
      call_id: response.call_id,
      output: toolResult
    }
  }));

  openAiWs.send(JSON.stringify({
    type: 'response.create',
    response: {
      instructions: 'Briefly tell Ramy whether the SMS was sent successfully.'
    }
  }));

  return;
}

                if (LOG_EVENT_TYPES.includes(response.type)) {
                    console.log(`Received event: ${response.type}`, response);
                }

                if (response.type === 'response.output_audio.delta' && response.delta) {
                    const audioDelta = {
                        event: 'media',
                        streamSid: streamSid,
                        media: { payload: response.delta }
                    };
                    connection.send(JSON.stringify(audioDelta));

                    // First delta from a new response starts the elapsed time counter
                    if (!responseStartTimestampTwilio) {
                        responseStartTimestampTwilio = latestMediaTimestamp;
                        if (SHOW_TIMING_MATH) console.log(`Setting start timestamp for new response: ${responseStartTimestampTwilio}ms`);
                    }

                    if (response.item_id) {
                        lastAssistantItem = response.item_id;
                    }
                    
                    sendMark(connection, streamSid);
                }

                if (response.type === 'input_audio_buffer.speech_started') {
                    handleSpeechStartedEvent();
                }
            } catch (error) {
                console.error('Error processing OpenAI message:', error, 'Raw message:', data);
            }
        });

        // Handle incoming messages from Twilio
        connection.on('message', (message) => {
            try {
                const data = JSON.parse(message);

                switch (data.event) {
                    case 'media':
                        latestMediaTimestamp = data.media.timestamp;
                        if (SHOW_TIMING_MATH) console.log(`Received media message with timestamp: ${latestMediaTimestamp}ms`);
                        if (openAiWs.readyState === WebSocket.OPEN) {
                            const audioAppend = {
                                type: 'input_audio_buffer.append',
                                audio: data.media.payload
                            };
                            openAiWs.send(JSON.stringify(audioAppend));
                        }
                        break;
                    case 'start':
                        streamSid = data.start.streamSid;
                        console.log('Incoming stream has started', streamSid);

                        // Reset start and media timestamp on a new stream
                        responseStartTimestampTwilio = null; 
                        latestMediaTimestamp = 0;
                        break;
                    case 'mark':
                        if (markQueue.length > 0) {
                            markQueue.shift();
                        }
                        break;
                    default:
                        console.log('Received non-media event:', data.event);
                        break;
                }
            } catch (error) {
                console.error('Error parsing message:', error, 'Message:', message);
            }
        });

        // Handle connection close
        connection.on('close', () => {
            if (openAiWs.readyState === WebSocket.OPEN) openAiWs.close();
            console.log('Client disconnected.');
        });

        // Handle WebSocket close and errors
        openAiWs.on('close', () => {
            console.log('Disconnected from the OpenAI Realtime API');
        });

        openAiWs.on('error', (error) => {
            console.error('Error in the OpenAI WebSocket:', error);
        });
    });
});

fastify.listen({ port: PORT, host: '0.0.0.0' }, (err) => {
    if (err) {
        console.error(err);
        process.exit(1);
    }
    console.log(`Server is listening on port ${PORT}`);
});
