import Fastify from 'fastify';
import WebSocket from 'ws';
import dotenv from 'dotenv';
import fastifyFormBody from '@fastify/formbody';
import fastifyWs from '@fastify/websocket';
import { randomUUID } from 'node:crypto';

// Load environment variables from .env file
dotenv.config();

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

const MONTREAL_IANA_TIME_ZONE = 'America/Toronto';
const MICROSOFT_EASTERN_TIME_ZONE = 'Eastern Standard Time';

// -----------------------------------------------------------------------------
// Shared helpers
// -----------------------------------------------------------------------------

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
    Body: message,
  });

  const response = await fetch(
    `https://api.twilio.com/2010-04-01/Accounts/${TWILIO_ACCOUNT_SID}/Messages.json`,
    {
      method: 'POST',
      headers: {
        Authorization: `Basic ${auth}`,
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: form.toString(),
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
    grant_type: 'client_credentials',
  });

  const response = await fetch(
    `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: form.toString(),
    }
  );

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Microsoft authentication failed: ${
        data.error_description || data.error || response.status
      }`
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
    grant_type: 'client_credentials',
  });

  const response = await fetch(
    `https://login.microsoftonline.com/${ACTIONS_MS_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: form.toString(),
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
      Authorization: `Bearer ${token}`,
    },
  });

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph email lookup failed: ${
        data.error?.message || response.status
      }`
    );
  }

  return data.value || [];
};

const sendEmailFromLondon = async ({ to, subject, body }) => {
  if (!LONDON_MINACO_EMAIL) {
    throw new Error('LONDON_MINACO_EMAIL is not configured.');
  }

  if (!to || !subject || !body) {
    throw new Error('Email requires recipient, subject, and body.');
  }

  const token = await getMicrosoftGraphActionsToken();

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      LONDON_MINACO_EMAIL
    )}/sendMail`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        message: {
          subject,
          body: {
            contentType: 'Text',
            content: body,
          },
          toRecipients: [
            {
              emailAddress: {
                address: to,
              },
            },
          ],
        },
        saveToSentItems: true,
      }),
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(
      `Microsoft Graph email send failed: ${errorText || response.status}`
    );
  }

  return {
    success: true,
    from: LONDON_MINACO_EMAIL,
    to,
    subject,
  };
};

const formatMontrealDateTime = (date) =>
  new Intl.DateTimeFormat('en-US', {
    timeZone: MONTREAL_IANA_TIME_ZONE,
    weekday: 'short',
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
    timeZoneName: 'short',
  }).format(date);

const formatGraphEasternDateTime = (dateTime) => {
  if (!dateTime) return '';

  // Graph returns the date/time already converted to Eastern time because we
  // send Prefer: outlook.timezone="Eastern Standard Time". Treat the returned
  // components as display components rather than converting them a second time.
  const match = String(dateTime).match(
    /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?/
  );

  if (!match) return String(dateTime);

  const [, year, month, day, hour, minute, second = '00'] = match;
  const pseudoUtc = new Date(
    Date.UTC(
      Number(year),
      Number(month) - 1,
      Number(day),
      Number(hour),
      Number(minute),
      Number(second)
    )
  );

  return new Intl.DateTimeFormat('en-US', {
    timeZone: 'UTC',
    weekday: 'short',
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
  }).format(pseudoUtc);
};

const getTimeZoneOffsetMs = (date, timeZone) => {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hourCycle: 'h23',
  }).formatToParts(date);

  const values = {};
  for (const part of parts) {
    if (part.type !== 'literal') values[part.type] = part.value;
  }

  const asUtc = Date.UTC(
    Number(values.year),
    Number(values.month) - 1,
    Number(values.day),
    Number(values.hour),
    Number(values.minute),
    Number(values.second)
  );

  return asUtc - date.getTime();
};

const montrealLocalToUtcIso = (localDateTime) => {
  const match = String(localDateTime || '').match(
    /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?$/
  );

  if (!match) {
    throw new Error(
      `Invalid Montreal local date/time "${localDateTime}". Use YYYY-MM-DDTHH:mm:ss.`
    );
  }

  const [, year, month, day, hour, minute, second = '00'] = match;

  const localAsUtc = new Date(
    Date.UTC(
      Number(year),
      Number(month) - 1,
      Number(day),
      Number(hour),
      Number(minute),
      Number(second)
    )
  );

  let offset = getTimeZoneOffsetMs(localAsUtc, MONTREAL_IANA_TIME_ZONE);
  let actualUtc = new Date(localAsUtc.getTime() - offset);

  const correctedOffset = getTimeZoneOffsetMs(
    actualUtc,
    MONTREAL_IANA_TIME_ZONE
  );

  if (correctedOffset !== offset) {
    offset = correctedOffset;
    actualUtc = new Date(localAsUtc.getTime() - offset);
  }

  return actualUtc.toISOString();
};

const ensureLocalDateTime = (value, label) => {
  if (
    typeof value !== 'string' ||
    !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2})?$/.test(value)
  ) {
    throw new Error(`${label} must use YYYY-MM-DDTHH:mm:ss Montreal local time.`);
  }
  return value.length === 16 ? `${value}:00` : value;
};

const normalizeAttendeeEmails = (attendees = []) => {
  if (!Array.isArray(attendees)) return [];

  return [...new Set(
    attendees
      .filter((email) => typeof email === 'string')
      .map((email) => email.trim())
      .filter(Boolean)
  )];
};

const simplifyCalendarEvent = (event) => ({
  id: event.id,
  subject: event.subject || '(No subject)',
  startLocal: event.start?.dateTime || '',
  endLocal: event.end?.dateTime || '',
  startMontreal: formatGraphEasternDateTime(event.start?.dateTime),
  endMontreal: formatGraphEasternDateTime(event.end?.dateTime),
  location: event.location?.displayName || '',
  organizer:
    event.organizer?.emailAddress?.name ||
    event.organizer?.emailAddress?.address ||
    '',
  organizerEmail: event.organizer?.emailAddress?.address || '',
  attendees: (event.attendees || []).map((attendee) => ({
    name: attendee.emailAddress?.name || '',
    email: attendee.emailAddress?.address || '',
    type: attendee.type || 'required',
    response: attendee.status?.response || '',
  })),
  showAs: event.showAs || '',
  isCancelled: Boolean(event.isCancelled),
  preview: event.bodyPreview || '',
});

const fetchCalendarView = async ({ startUtc, endUtc, top = 50 }) => {
  if (!RAMY_MINACO_EMAIL) {
    throw new Error('RAMY_MINACO_EMAIL is not configured.');
  }

  const token = await getMicrosoftGraphActionsToken();

  const url = new URL(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      RAMY_MINACO_EMAIL
    )}/calendarView`
  );

  url.searchParams.set('startDateTime', startUtc);
  url.searchParams.set('endDateTime', endUtc);
  url.searchParams.set('$top', String(Math.min(Math.max(top, 1), 100)));
  url.searchParams.set(
    '$select',
    'id,subject,start,end,location,organizer,attendees,bodyPreview,isCancelled,showAs'
  );
  url.searchParams.set('$orderby', 'start/dateTime');

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Prefer: `outlook.timezone="${MICROSOFT_EASTERN_TIME_ZONE}"`,
    },
  });

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph calendar lookup failed: ${
        data.error?.message || response.status
      }`
    );
  }

  return data.value || [];
};

const getCalendarEvents = async ({ startLocal, endLocal, top = 50 } = {}) => {
  let startUtc;
  let endUtc;

  if (startLocal && endLocal) {
    const normalizedStart = ensureLocalDateTime(startLocal, 'start_local');
    const normalizedEnd = ensureLocalDateTime(endLocal, 'end_local');
    startUtc = montrealLocalToUtcIso(normalizedStart);
    endUtc = montrealLocalToUtcIso(normalizedEnd);
  } else if (!startLocal && !endLocal) {
    const start = new Date();
    const end = new Date(start.getTime() + 7 * 24 * 60 * 60 * 1000);
    startUtc = start.toISOString();
    endUtc = end.toISOString();
  } else {
    throw new Error('Provide both start_local and end_local, or neither.');
  }

  const events = await fetchCalendarView({ startUtc, endUtc, top });
  return events.map(simplifyCalendarEvent);
};

const checkCalendarAvailability = async ({ startLocal, endLocal }) => {
  const normalizedStart = ensureLocalDateTime(startLocal, 'start_local');
  const normalizedEnd = ensureLocalDateTime(endLocal, 'end_local');

  const startUtc = montrealLocalToUtcIso(normalizedStart);
  const endUtc = montrealLocalToUtcIso(normalizedEnd);

  if (new Date(endUtc) <= new Date(startUtc)) {
    throw new Error('End time must be after start time.');
  }

  const events = (await fetchCalendarView({ startUtc, endUtc, top: 50 })).map(
    simplifyCalendarEvent
  );

  const conflicts = events.filter(
    (event) => !event.isCancelled && event.showAs !== 'free'
  );

  return {
    startLocal: normalizedStart,
    endLocal: normalizedEnd,
    available: conflicts.length === 0,
    conflicts,
  };
};

const createCalendarEvent = async ({
  subject,
  startLocal,
  endLocal,
  attendees = [],
  location = '',
  body = '',
}) => {
  if (!RAMY_MINACO_EMAIL) {
    throw new Error('RAMY_MINACO_EMAIL is not configured.');
  }

  if (!subject) {
    throw new Error('Calendar event requires a subject.');
  }

  const normalizedStart = ensureLocalDateTime(startLocal, 'start_local');
  const normalizedEnd = ensureLocalDateTime(endLocal, 'end_local');

  if (
    new Date(montrealLocalToUtcIso(normalizedEnd)) <=
    new Date(montrealLocalToUtcIso(normalizedStart))
  ) {
    throw new Error('End time must be after start time.');
  }

  const token = await getMicrosoftGraphActionsToken();
  const attendeeEmails = normalizeAttendeeEmails(attendees);

  const eventPayload = {
    subject,
    body: {
      contentType: 'Text',
      content: body || '',
    },
    start: {
      dateTime: normalizedStart,
      timeZone: MICROSOFT_EASTERN_TIME_ZONE,
    },
    end: {
      dateTime: normalizedEnd,
      timeZone: MICROSOFT_EASTERN_TIME_ZONE,
    },
    attendees: attendeeEmails.map((email) => ({
      emailAddress: { address: email },
      type: 'required',
    })),
  };

  if (location) {
    eventPayload.location = { displayName: location };
  }

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      RAMY_MINACO_EMAIL
    )}/calendar/events`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        Prefer: `outlook.timezone="${MICROSOFT_EASTERN_TIME_ZONE}"`,
      },
      body: JSON.stringify(eventPayload),
    }
  );

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph calendar create failed: ${
        data.error?.message || response.status
      }`
    );
  }

  return simplifyCalendarEvent(data);
};

const updateCalendarEvent = async ({
  eventId,
  subject,
  startLocal,
  endLocal,
  attendees,
  location,
  body,
}) => {
  if (!RAMY_MINACO_EMAIL) {
    throw new Error('RAMY_MINACO_EMAIL is not configured.');
  }

  if (!eventId) {
    throw new Error('Calendar update requires an event id.');
  }

  const patch = {};

  if (typeof subject === 'string' && subject.trim()) {
    patch.subject = subject.trim();
  }

  if (startLocal || endLocal) {
    if (!startLocal || !endLocal) {
      throw new Error('A reschedule requires both start_local and end_local.');
    }

    const normalizedStart = ensureLocalDateTime(startLocal, 'start_local');
    const normalizedEnd = ensureLocalDateTime(endLocal, 'end_local');

    if (
      new Date(montrealLocalToUtcIso(normalizedEnd)) <=
      new Date(montrealLocalToUtcIso(normalizedStart))
    ) {
      throw new Error('End time must be after start time.');
    }

    patch.start = {
      dateTime: normalizedStart,
      timeZone: MICROSOFT_EASTERN_TIME_ZONE,
    };
    patch.end = {
      dateTime: normalizedEnd,
      timeZone: MICROSOFT_EASTERN_TIME_ZONE,
    };
  }

  if (typeof location === 'string') {
    patch.location = { displayName: location };
  }

  if (typeof body === 'string') {
    patch.body = {
      contentType: 'Text',
      content: body,
    };
  }

  if (Array.isArray(attendees)) {
    patch.attendees = normalizeAttendeeEmails(attendees).map((email) => ({
      emailAddress: { address: email },
      type: 'required',
    }));
  }

  if (Object.keys(patch).length === 0) {
    throw new Error('No calendar changes were provided.');
  }

  const token = await getMicrosoftGraphActionsToken();

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      RAMY_MINACO_EMAIL
    )}/events/${encodeURIComponent(eventId)}`,
    {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        Prefer: `outlook.timezone="${MICROSOFT_EASTERN_TIME_ZONE}"`,
      },
      body: JSON.stringify(patch),
    }
  );

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph calendar update failed: ${
        data.error?.message || response.status
      }`
    );
  }

  return simplifyCalendarEvent(data);
};

const deleteCalendarEvent = async (eventId) => {
  if (!RAMY_MINACO_EMAIL) {
    throw new Error('RAMY_MINACO_EMAIL is not configured.');
  }

  if (!eventId) {
    throw new Error('Calendar cancellation requires an event id.');
  }

  const token = await getMicrosoftGraphActionsToken();

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      RAMY_MINACO_EMAIL
    )}/events/${encodeURIComponent(eventId)}`,
    {
      method: 'DELETE',
      headers: {
        Authorization: `Bearer ${token}`,
      },
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(
      `Microsoft Graph calendar cancellation failed: ${
        errorText || response.status
      }`
    );
  }

  return { success: true, eventId };
};

// -----------------------------------------------------------------------------
// Fastify / London voice configuration
// -----------------------------------------------------------------------------

const fastify = Fastify();
fastify.register(fastifyFormBody);
fastify.register(fastifyWs);

const SYSTEM_MESSAGE = `
You are London Assistant, the executive assistant to Ramy Mina.

IDENTITY AND BUSINESS CONTEXT

You operate primarily for Minaco / Mina Group.
Mina Capital was previously used as a technical test environment. Do not assume Mina Capital unless Ramy specifically says Mina Capital.

Verified Minaco identities:
- Ramy Mina: principal executive
- Ramy's primary Minaco email: ramy.mina@minaco.ca
- London Assistant email: london@minaco.ca
- Minaco accounting email: accounting@minaco.ca
- London Assistant phone number: +1 438-255-9099

When Ramy asks what his email address is without specifying another company, answer ramy.mina@minaco.ca.
When Ramy asks who you are, say: "I am London Assistant, your executive assistant for Minaco."

ACCURACY IS YOUR HIGHEST PRIORITY

Never invent or guess facts about Ramy, Minaco, Mina Capital, Mina Group, employees, consultants, contractors, lawyers, partners, investors, lenders, tenants, projects, properties, emails, meetings, financial information, contracts, deadlines, or business relationships.

If information is not explicitly provided in verified context or retrieved from an authorized live system, say that you do not have verified information.
Never make up a plausible answer just to be helpful.

LIVE EMAIL RULES

Use check_email whenever Ramy asks about current emails or his inbox. Never answer a current email question from memory.
Use prepare_email when Ramy asks you to draft, write, prepare, or send an email.
Preparing an email NEVER sends it.
After prepare_email succeeds, read back the recipient, subject, and message and clearly say it has not been sent. Ask Ramy to confirm.
Use send_confirmed_email ONLY after Ramy explicitly confirms the pending email with "Send it" or an unmistakable equivalent in response to your confirmation request.
Never claim an email was sent unless the send tool confirms success.
Never invent a recipient email address. If Ramy gives only a person's name and you do not have a verified email address, ask for the email address.

LIVE CALENDAR RULES

Use check_calendar whenever Ramy asks what is on his calendar, what meetings he has, or asks about a specific date or period.
Use check_availability whenever Ramy asks whether he is free at a particular time.
All spoken calendar times are Montreal local time unless Ramy explicitly specifies another time zone.

For a NEW meeting:
1. Gather subject, Montreal start time, Montreal end time or duration, and attendee email addresses if attendees are required.
2. Use prepare_calendar_event. This does NOT create the event.
3. Read back the subject, date, start/end time, attendees, and location if any. Clearly say it is not booked yet.
4. Ask Ramy to confirm by saying "Book it."
5. Use confirm_calendar_action only after that explicit confirmation.

For RESCHEDULING or other calendar changes:
1. Use check_calendar first to identify the exact event and its event id.
2. Use prepare_calendar_update with the exact event id and proposed changes.
3. Read back what will change and clearly say it has not been changed yet.
4. Ask Ramy to confirm by saying "Reschedule it" or "Make the change."
5. Use confirm_calendar_action only after explicit confirmation.

For CANCELLING:
1. Use check_calendar first to identify the exact event and its event id.
2. Use prepare_calendar_cancel.
3. Read back the event being cancelled and clearly say it has not been cancelled yet.
4. Ask Ramy to confirm by saying "Cancel it."
5. Use confirm_calendar_action only after explicit confirmation.

If more than one calendar event could match Ramy's request, do not guess. Tell him the matching events and ask which one he means.
Never call confirm_calendar_action unless there is a pending prepared action and Ramy has explicitly confirmed it.
Never claim a meeting was booked, changed, or cancelled unless Microsoft Graph confirms success.

SMS

Use send_sms only when Ramy explicitly asks you to text him. Never claim an SMS was sent unless the tool confirms success.

EXECUTIVE ASSISTANT BEHAVIOR

Your role is to reduce Ramy's workload. Be concise, practical, commercially aware, organized, calm, professional, and proactive when appropriate.
Prioritize decisions, deadlines and risks, financial or contractual consequences, commitments, follow-ups, then routine information.
Do not overwhelm Ramy with unnecessary detail.
If a question is materially ambiguous, ask one short clarification instead of guessing.

AUTHORITY

You may retrieve information, summarize, organize, prepare actions, send SMS when explicitly requested, send a confirmed email, and execute a confirmed calendar action through the authorized tools.
Do not claim to have approved payments, signed contracts, committed Minaco to pricing, settled disputes, or made legal or financial decisions unless an authorized system actually performed that action.

PHONE STYLE

You are speaking with Ramy by phone. Speak naturally and keep most answers short unless Ramy asks for detail.
Accuracy is more important than sounding helpful.
`;

const VOICE = 'marin';
const TEMPERATURE = 0.3;
const PORT = process.env.PORT || 5050;

const LOG_EVENT_TYPES = [
  'error',
  'response.content.done',
  'rate_limits.updated',
  'response.done',
  'input_audio_buffer.committed',
  'input_audio_buffer.speech_stopped',
  'input_audio_buffer.speech_started',
  'session.created',
  'session.updated',
];

const SHOW_TIMING_MATH = false;

// A short-lived, one-time token prevents someone from bypassing the Twilio
// caller-number check by opening the public WebSocket URL directly.
const authorizedStreamTokens = new Map();

fastify.get('/', async (request, reply) => {
  reply.send({ message: 'London Assistant Twilio Media Stream Server is running!' });
});

fastify.all('/incoming-call', async (request, reply) => {
  const caller = request.body?.From || request.query?.From;
  const authorized = Boolean(caller && caller === RAMY_PHONE_NUMBER);

  console.log('CALL SECURITY CHECK:', { caller, authorized });

  if (!authorized) {
    const deniedResponse = `<?xml version="1.0" encoding="UTF-8"?>
<Response>
  <Say>Sorry, this line is private.</Say>
  <Hangup/>
</Response>`;

    return reply.type('text/xml').send(deniedResponse);
  }

  const streamToken = randomUUID();
  authorizedStreamTokens.set(streamToken, Date.now() + 2 * 60 * 1000);

  const cleanupTimer = setTimeout(() => {
    authorizedStreamTokens.delete(streamToken);
  }, 2 * 60 * 1000);
  if (typeof cleanupTimer.unref === 'function') cleanupTimer.unref();

  const twimlResponse = `<?xml version="1.0" encoding="UTF-8"?>
<Response>
  <Say voice="Google.en-US-Chirp3-HD-Aoede">Please wait while I connect you to London Assistant.</Say>
  <Pause length="1"/>
  <Say voice="Google.en-US-Chirp3-HD-Aoede">You can start speaking.</Say>
  <Connect>
    <Stream url="wss://${request.headers.host}/media-stream/${streamToken}" />
  </Connect>
</Response>`;

  return reply.type('text/xml').send(twimlResponse);
});

fastify.register(async (fastifyInstance) => {
  fastifyInstance.get(
    '/media-stream/:streamToken',
    { websocket: true },
    (connection, req) => {
      const streamToken = req.params?.streamToken;
      const expiry = authorizedStreamTokens.get(streamToken);

      if (!expiry || expiry < Date.now()) {
        console.warn('Rejected unauthorized or expired media stream.');
        try {
          connection.close(1008, 'Unauthorized');
        } catch {
          connection.close();
        }
        return;
      }

      authorizedStreamTokens.delete(streamToken);
      console.log('Authorized client connected');

      let streamSid = null;
      let latestMediaTimestamp = 0;
      let lastAssistantItem = null;
      let markQueue = [];
      let responseStartTimestampTwilio = null;
      let pendingEmailDraft = null;
      let pendingCalendarAction = null;

      const openAiWs = new WebSocket(
        `wss://api.openai.com/v1/realtime?model=gpt-realtime&temperature=${TEMPERATURE}`,
        {
          headers: {
            Authorization: `Bearer ${OPENAI_API_KEY}`,
          },
        }
      );

      const respondToToolCall = (callId, output, instructions) => {
        openAiWs.send(
          JSON.stringify({
            type: 'conversation.item.create',
            item: {
              type: 'function_call_output',
              call_id: callId,
              output:
                typeof output === 'string' ? output : JSON.stringify(output),
            },
          })
        );

        openAiWs.send(
          JSON.stringify({
            type: 'response.create',
            response: {
              instructions,
            },
          })
        );
      };

      const initializeSession = () => {
        const currentMontrealContext = formatMontrealDateTime(new Date());

        const sessionUpdate = {
          type: 'session.update',
          session: {
            type: 'realtime',
            model: 'gpt-realtime',
            output_modalities: ['audio'],
            audio: {
              input: {
                format: { type: 'audio/pcmu' },
                turn_detection: { type: 'server_vad' },
              },
              output: {
                format: { type: 'audio/pcmu' },
                voice: VOICE,
              },
            },
            instructions: `${SYSTEM_MESSAGE}\n\nCURRENT MONTREAL DATE AND TIME AT CALL START: ${currentMontrealContext}. Use this to resolve words such as today, tomorrow, this afternoon, and next week.`,
            tools: [
              {
                type: 'function',
                name: 'send_sms',
                description:
                  'Send an SMS message to Ramy when Ramy explicitly asks London to text him.',
                parameters: {
                  type: 'object',
                  properties: {
                    message: {
                      type: 'string',
                      description: 'The exact SMS message to send to Ramy.',
                    },
                  },
                  required: ['message'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'check_email',
                description:
                  'Read Ramy Mina’s latest live Minaco inbox emails. Use for current inbox questions. This retrieves the most recent messages; it is not a full historical mailbox search.',
                parameters: {
                  type: 'object',
                  properties: {
                    limit: {
                      type: 'integer',
                      minimum: 1,
                      maximum: 10,
                      description:
                        'Number of recent emails to retrieve. Use 5 unless Ramy asks otherwise.',
                    },
                  },
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'prepare_email',
                description:
                  'Prepare an email draft for Ramy. Never sends the email. Use when Ramy asks to write, draft, prepare, or send an email. After preparing, read it back and ask for explicit confirmation.',
                parameters: {
                  type: 'object',
                  properties: {
                    to: {
                      type: 'string',
                      description: 'Verified recipient email address.',
                    },
                    subject: {
                      type: 'string',
                      description: 'Email subject.',
                    },
                    body: {
                      type: 'string',
                      description: 'Complete email body.',
                    },
                  },
                  required: ['to', 'subject', 'body'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'send_confirmed_email',
                description:
                  'Send the already-prepared email draft. Use ONLY after Ramy explicitly confirms the pending email by saying Send it or an unmistakable equivalent.',
                parameters: {
                  type: 'object',
                  properties: {},
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'check_calendar',
                description:
                  'Read Ramy’s live Minaco calendar. Use when he asks what is on his calendar, about meetings on a date, or needs an event identified before changing or cancelling it. Times are Montreal local time. If start_local and end_local are omitted, returns the next 7 days.',
                parameters: {
                  type: 'object',
                  properties: {
                    start_local: {
                      type: 'string',
                      description:
                        'Optional Montreal local start in YYYY-MM-DDTHH:mm:ss.',
                    },
                    end_local: {
                      type: 'string',
                      description:
                        'Optional Montreal local end in YYYY-MM-DDTHH:mm:ss.',
                    },
                    limit: {
                      type: 'integer',
                      minimum: 1,
                      maximum: 50,
                      description: 'Maximum events to return. Default 20.',
                    },
                  },
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'check_availability',
                description:
                  'Check whether Ramy is free during a specific Montreal-local time range. Returns any conflicting calendar events.',
                parameters: {
                  type: 'object',
                  properties: {
                    start_local: {
                      type: 'string',
                      description:
                        'Montreal local start in YYYY-MM-DDTHH:mm:ss.',
                    },
                    end_local: {
                      type: 'string',
                      description:
                        'Montreal local end in YYYY-MM-DDTHH:mm:ss.',
                    },
                  },
                  required: ['start_local', 'end_local'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'prepare_calendar_event',
                description:
                  'Prepare a new event on Ramy’s Minaco calendar. This NEVER books the event. After preparing it, read back the details and ask Ramy to say Book it.',
                parameters: {
                  type: 'object',
                  properties: {
                    subject: {
                      type: 'string',
                      description: 'Meeting or event subject.',
                    },
                    start_local: {
                      type: 'string',
                      description:
                        'Montreal local start in YYYY-MM-DDTHH:mm:ss.',
                    },
                    end_local: {
                      type: 'string',
                      description:
                        'Montreal local end in YYYY-MM-DDTHH:mm:ss.',
                    },
                    attendees: {
                      type: 'array',
                      items: { type: 'string' },
                      description:
                        'Required attendee email addresses. Use an empty array when there are no invitees.',
                    },
                    location: {
                      type: 'string',
                      description: 'Optional meeting location.',
                    },
                    body: {
                      type: 'string',
                      description: 'Optional meeting notes or invitation message.',
                    },
                  },
                  required: ['subject', 'start_local', 'end_local'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'prepare_calendar_update',
                description:
                  'Prepare changes to an existing calendar event. This NEVER changes the calendar. Use the exact event_id retrieved from check_calendar, then read back the proposed change and ask for explicit confirmation.',
                parameters: {
                  type: 'object',
                  properties: {
                    event_id: {
                      type: 'string',
                      description:
                        'Exact Microsoft Graph event id returned by check_calendar.',
                    },
                    event_subject: {
                      type: 'string',
                      description:
                        'Current event subject for a human-friendly confirmation.',
                    },
                    subject: {
                      type: 'string',
                      description: 'Optional new subject.',
                    },
                    start_local: {
                      type: 'string',
                      description:
                        'Optional new Montreal local start in YYYY-MM-DDTHH:mm:ss. If supplied, end_local is also required.',
                    },
                    end_local: {
                      type: 'string',
                      description:
                        'Optional new Montreal local end in YYYY-MM-DDTHH:mm:ss. If supplied, start_local is also required.',
                    },
                    attendees: {
                      type: 'array',
                      items: { type: 'string' },
                      description:
                        'Optional replacement attendee email list. Omit unless Ramy explicitly wants attendee changes.',
                    },
                    location: {
                      type: 'string',
                      description:
                        'Optional new location. Omit unless Ramy explicitly wants it changed.',
                    },
                    body: {
                      type: 'string',
                      description:
                        'Optional new meeting notes. Omit unless Ramy explicitly wants them changed.',
                    },
                  },
                  required: ['event_id'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'prepare_calendar_cancel',
                description:
                  'Prepare cancellation of an existing calendar event. This NEVER cancels the event. Use the exact event_id retrieved from check_calendar and ask Ramy to say Cancel it.',
                parameters: {
                  type: 'object',
                  properties: {
                    event_id: {
                      type: 'string',
                      description:
                        'Exact Microsoft Graph event id returned by check_calendar.',
                    },
                    event_subject: {
                      type: 'string',
                      description:
                        'Event subject for a human-friendly confirmation.',
                    },
                    start_local: {
                      type: 'string',
                      description:
                        'Optional event start time for a human-friendly confirmation.',
                    },
                  },
                  required: ['event_id'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'confirm_calendar_action',
                description:
                  'Execute the pending prepared calendar create, update, or cancellation. Use ONLY after Ramy explicitly confirms the pending action with Book it, Reschedule it, Make the change, Cancel it, or an unmistakable equivalent in direct response to the confirmation request.',
                parameters: {
                  type: 'object',
                  properties: {},
                  additionalProperties: false,
                },
              },
            ],
            tool_choice: 'auto',
          },
        };

        console.log('Sending session update');
        openAiWs.send(JSON.stringify(sessionUpdate));
      };

      const handleSpeechStartedEvent = () => {
        if (markQueue.length > 0 && responseStartTimestampTwilio != null) {
          const elapsedTime =
            latestMediaTimestamp - responseStartTimestampTwilio;

          if (SHOW_TIMING_MATH) {
            console.log(
              `Calculating elapsed time for truncation: ${latestMediaTimestamp} - ${responseStartTimestampTwilio} = ${elapsedTime}ms`
            );
          }

          if (lastAssistantItem) {
            const truncateEvent = {
              type: 'conversation.item.truncate',
              item_id: lastAssistantItem,
              content_index: 0,
              audio_end_ms: elapsedTime,
            };
            openAiWs.send(JSON.stringify(truncateEvent));
          }

          connection.send(
            JSON.stringify({
              event: 'clear',
              streamSid,
            })
          );

          markQueue = [];
          lastAssistantItem = null;
          responseStartTimestampTwilio = null;
        }
      };

      const sendMark = (streamConnection, activeStreamSid) => {
        if (!activeStreamSid) return;

        const markEvent = {
          event: 'mark',
          streamSid: activeStreamSid,
          mark: { name: 'responsePart' },
        };

        streamConnection.send(JSON.stringify(markEvent));
        markQueue.push('responsePart');
      };

      openAiWs.on('open', () => {
        console.log('Connected to the OpenAI Realtime API');
        setTimeout(initializeSession, 100);
      });

      openAiWs.on('message', async (data) => {
        try {
          const response = JSON.parse(data);

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'send_confirmed_email'
          ) {
            try {
              if (!pendingEmailDraft) {
                throw new Error('There is no pending email draft to send.');
              }

              const emailToSend = { ...pendingEmailDraft };
              const result = await sendEmailFromLondon(emailToSend);
              pendingEmailDraft = null;

              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  sent: true,
                  from: result.from,
                  to: result.to,
                  subject: result.subject,
                },
                'Confirm briefly that the email was successfully sent. Mention the recipient and subject. Do not invent anything.'
              );
            } catch (error) {
              console.error('Confirmed email send error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, sent: false, error: error.message },
                'Tell Ramy the email was NOT sent and state the returned error concisely. Do not invent anything.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'prepare_email'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');

              if (!args.to || !args.subject || !args.body) {
                throw new Error('Recipient, subject, and body are required.');
              }

              pendingEmailDraft = {
                to: args.to,
                subject: args.subject,
                body: args.body,
              };

              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  draft: pendingEmailDraft,
                  sent: false,
                },
                'Read back the recipient, subject, and email message to Ramy. Clearly say the email has NOT been sent. Ask Ramy to confirm by saying Send it.'
              );
            } catch (error) {
              console.error('Prepare email error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the email draft could not be prepared and state the error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'check_email'
          ) {
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
                  timeZone: MONTREAL_IANA_TIME_ZONE,
                  dateStyle: 'medium',
                  timeStyle: 'short',
                }).format(new Date(email.receivedDateTime)),
                isRead: Boolean(email.isRead),
                preview: email.bodyPreview || '',
              }));

              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  mailbox: RAMY_MINACO_EMAIL,
                  emails: simplifiedEmails,
                },
                'Answer using only the live Minaco email results. Email times are already Montreal local time. Be concise and practical. Mention sender, subject, and what matters. Do not invent anything.'
              );
            } catch (error) {
              console.error('Email lookup error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the live email lookup failed and state the returned error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'check_calendar'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const requestedLimit = Number(args.limit);
              const limit =
                Number.isInteger(requestedLimit) && requestedLimit >= 1
                  ? Math.min(requestedLimit, 50)
                  : 20;

              const events = await getCalendarEvents({
                startLocal: args.start_local,
                endLocal: args.end_local,
                top: limit,
              });

              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  mailbox: RAMY_MINACO_EMAIL,
                  timezone: 'Montreal local time',
                  events,
                },
                'Answer using only the live calendar results. All displayed times are Montreal local time. Be concise. If the tool returned no events, say the calendar is clear for the requested period. Never invent an event.'
              );
            } catch (error) {
              console.error('Calendar lookup error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the live calendar lookup failed and state the returned error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'check_availability'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const result = await checkCalendarAvailability({
                startLocal: args.start_local,
                endLocal: args.end_local,
              });

              respondToToolCall(
                response.call_id,
                { success: true, timezone: 'Montreal local time', ...result },
                'Tell Ramy whether he is free during the requested Montreal-local time. If there are conflicts, mention the conflicting event subjects and times. Use only the tool results.'
              );
            } catch (error) {
              console.error('Availability check error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the availability check failed and state the returned error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'prepare_calendar_event'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const startLocal = ensureLocalDateTime(
                args.start_local,
                'start_local'
              );
              const endLocal = ensureLocalDateTime(args.end_local, 'end_local');

              if (
                new Date(montrealLocalToUtcIso(endLocal)) <=
                new Date(montrealLocalToUtcIso(startLocal))
              ) {
                throw new Error('End time must be after start time.');
              }

              if (!args.subject) {
                throw new Error('A calendar event subject is required.');
              }

              pendingCalendarAction = {
                type: 'create',
                payload: {
                  subject: args.subject,
                  startLocal,
                  endLocal,
                  attendees: normalizeAttendeeEmails(args.attendees || []),
                  location: args.location || '',
                  body: args.body || '',
                },
              };

              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  booked: false,
                  pendingAction: pendingCalendarAction,
                },
                'Read back the proposed meeting subject, Montreal date, start and end time, attendee email addresses, and location if any. Clearly say it is NOT booked yet. Ask Ramy to confirm by saying Book it.'
              );
            } catch (error) {
              console.error('Prepare calendar event error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the meeting could not be prepared and state the returned error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'prepare_calendar_update'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');

              if (!args.event_id) {
                throw new Error('An exact event id is required.');
              }

              const payload = {
                eventId: args.event_id,
              };

              if (typeof args.subject === 'string') payload.subject = args.subject;
              if (typeof args.start_local === 'string') {
                payload.startLocal = ensureLocalDateTime(
                  args.start_local,
                  'start_local'
                );
              }
              if (typeof args.end_local === 'string') {
                payload.endLocal = ensureLocalDateTime(
                  args.end_local,
                  'end_local'
                );
              }
              if (Array.isArray(args.attendees)) {
                payload.attendees = normalizeAttendeeEmails(args.attendees);
              }
              if (typeof args.location === 'string') {
                payload.location = args.location;
              }
              if (typeof args.body === 'string') payload.body = args.body;

              if (payload.startLocal || payload.endLocal) {
                if (!payload.startLocal || !payload.endLocal) {
                  throw new Error(
                    'A reschedule requires both start_local and end_local.'
                  );
                }
              }

              if (Object.keys(payload).length === 1) {
                throw new Error('No calendar changes were provided.');
              }

              pendingCalendarAction = {
                type: 'update',
                eventSubject: args.event_subject || '',
                payload,
              };

              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  changed: false,
                  pendingAction: pendingCalendarAction,
                },
                'Read back the event and the proposed changes in Montreal local time. Clearly say the calendar has NOT been changed yet. Ask Ramy to confirm by saying Reschedule it or Make the change.'
              );
            } catch (error) {
              console.error('Prepare calendar update error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the calendar change could not be prepared and state the returned error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'prepare_calendar_cancel'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');

              if (!args.event_id) {
                throw new Error('An exact event id is required.');
              }

              pendingCalendarAction = {
                type: 'cancel',
                eventSubject: args.event_subject || '',
                startLocal: args.start_local || '',
                payload: {
                  eventId: args.event_id,
                },
              };

              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  cancelled: false,
                  pendingAction: pendingCalendarAction,
                },
                'Read back exactly which event will be cancelled, including the Montreal-local time if available. Clearly say it has NOT been cancelled yet. Ask Ramy to confirm by saying Cancel it.'
              );
            } catch (error) {
              console.error('Prepare calendar cancellation error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the cancellation could not be prepared and state the returned error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'confirm_calendar_action'
          ) {
            try {
              if (!pendingCalendarAction) {
                throw new Error('There is no pending calendar action to confirm.');
              }

              const action = { ...pendingCalendarAction };
              let result;

              if (action.type === 'create') {
                result = await createCalendarEvent(action.payload);
              } else if (action.type === 'update') {
                result = await updateCalendarEvent(action.payload);
              } else if (action.type === 'cancel') {
                result = await deleteCalendarEvent(action.payload.eventId);
              } else {
                throw new Error('Unknown pending calendar action.');
              }

              pendingCalendarAction = null;

              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  action: action.type,
                  result,
                  timezone: 'Montreal local time',
                },
                'Confirm briefly that the requested calendar action succeeded. For a new or updated event, mention the subject and Montreal-local time returned by the tool. For a cancellation, confirm the event was cancelled. Do not invent anything.'
              );
            } catch (error) {
              console.error('Confirmed calendar action error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the calendar action was NOT completed and state the returned error concisely. Do not claim success.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'send_sms'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');

              if (!args.message || typeof args.message !== 'string') {
                throw new Error('No SMS message was provided.');
              }

              await sendSmsToRamy(args.message);

              respondToToolCall(
                response.call_id,
                { success: true, message: 'SMS sent successfully to Ramy.' },
                'Briefly tell Ramy the SMS was sent successfully.'
              );
            } catch (error) {
              console.error('SMS error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the SMS was not sent and state the returned error concisely.'
              );
            }
            return;
          }

          if (LOG_EVENT_TYPES.includes(response.type)) {
            console.log(`Received event: ${response.type}`, response);
          }

          if (
            response.type === 'response.output_audio.delta' &&
            response.delta
          ) {
            const audioDelta = {
              event: 'media',
              streamSid,
              media: { payload: response.delta },
            };

            connection.send(JSON.stringify(audioDelta));

            if (!responseStartTimestampTwilio) {
              responseStartTimestampTwilio = latestMediaTimestamp;
              if (SHOW_TIMING_MATH) {
                console.log(
                  `Setting start timestamp for new response: ${responseStartTimestampTwilio}ms`
                );
              }
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
          console.error(
            'Error processing OpenAI message:',
            error,
            'Raw message:',
            data
          );
        }
      });

      connection.on('message', (message) => {
        try {
          const data = JSON.parse(message);

          switch (data.event) {
            case 'media':
              latestMediaTimestamp = data.media.timestamp;
              if (SHOW_TIMING_MATH) {
                console.log(
                  `Received media message with timestamp: ${latestMediaTimestamp}ms`
                );
              }
              if (openAiWs.readyState === WebSocket.OPEN) {
                openAiWs.send(
                  JSON.stringify({
                    type: 'input_audio_buffer.append',
                    audio: data.media.payload,
                  })
                );
              }
              break;

            case 'start':
              streamSid = data.start.streamSid;
              console.log('Incoming stream has started', streamSid);
              responseStartTimestampTwilio = null;
              latestMediaTimestamp = 0;
              break;

            case 'mark':
              if (markQueue.length > 0) markQueue.shift();
              break;

            default:
              console.log('Received non-media event:', data.event);
              break;
          }
        } catch (error) {
          console.error('Error parsing message:', error, 'Message:', message);
        }
      });

      connection.on('close', () => {
        if (openAiWs.readyState === WebSocket.OPEN) openAiWs.close();
        console.log('Client disconnected.');
      });

      openAiWs.on('close', () => {
        console.log('Disconnected from the OpenAI Realtime API');
      });

      openAiWs.on('error', (error) => {
        console.error('Error in the OpenAI WebSocket:', error);
      });
    }
  );
});

fastify.listen({ port: PORT, host: '0.0.0.0' }, (err) => {
  if (err) {
    console.error(err);
    process.exit(1);
  }
  console.log(`Server is listening on port ${PORT}`);
});
