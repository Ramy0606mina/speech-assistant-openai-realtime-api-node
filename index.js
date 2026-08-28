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
  ACCOUNTING_MINACO_EMAIL,
  DAILY_BRIEF_SECRET,
  OPENAI_DOCUMENT_MODEL,
  EXECUTIVE_BRIEF_MODEL,
} = process.env;

const MONTREAL_IANA_TIME_ZONE = 'America/Toronto';
const MICROSOFT_EASTERN_TIME_ZONE = 'Eastern Standard Time';
const ACCOUNTING_MAILBOX = ACCOUNTING_MINACO_EMAIL || 'accounting@minaco.ca';
const ACTION_REGISTER_CALENDAR_NAME = 'London Action Register';
const DOCUMENT_ANALYSIS_MODEL = OPENAI_DOCUMENT_MODEL || 'gpt-5.6';
const DAILY_BRIEF_MODEL = EXECUTIVE_BRIEF_MODEL || 'gpt-5.6-luna';
const MAX_ATTACHMENT_BYTES = 45 * 1024 * 1024;

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
      signal: AbortSignal.timeout(12000),
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
    'id,subject,from,receivedDateTime,bodyPreview,isRead'
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

const getFullMinacoEmail = async (messageId) => {
  if (!RAMY_MINACO_EMAIL) {
    throw new Error('RAMY_MINACO_EMAIL is not configured.');
  }

  if (!messageId) {
    throw new Error('A message id is required.');
  }

  const token = await getMicrosoftGraphToken();

  const url = new URL(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      RAMY_MINACO_EMAIL
    )}/messages/${encodeURIComponent(messageId)}`
  );

  url.searchParams.set(
    '$select',
    'id,conversationId,internetMessageId,subject,from,replyTo,toRecipients,ccRecipients,receivedDateTime,body,isRead'
  );

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Prefer: 'outlook.body-content-type="text"',
    },
  });

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph full email lookup failed: ${
        data.error?.message || response.status
      }`
    );
  }

  return data;
};


// -----------------------------------------------------------------------------
// Advanced mailbox search, attachments, and contact intelligence
// -----------------------------------------------------------------------------

const mailboxAddressFromKey = (mailbox = 'ramy') => {
  const key = String(mailbox || 'ramy').trim().toLowerCase();
  if (key === 'ramy' || key === 'executive' || key === 'inbox') {
    return RAMY_MINACO_EMAIL;
  }
  if (key === 'accounting' || key === 'accounts' || key === 'finance') {
    return ACCOUNTING_MAILBOX;
  }
  if (key === 'london' || key === 'assistant') {
    return LONDON_MINACO_EMAIL;
  }
  if (key.includes('@')) return key;
  throw new Error(`Unknown mailbox: ${mailbox}`);
};

const getRecentMailboxEmails = async (mailboxAddress, limit = 5) => {
  if (!mailboxAddress) throw new Error('Mailbox address is required.');

  const token = await getMicrosoftGraphToken();
  const url = new URL(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      mailboxAddress
    )}/mailFolders/inbox/messages`
  );

  url.searchParams.set('$top', String(Math.min(Math.max(Number(limit) || 5, 1), 25)));
  url.searchParams.set(
    '$select',
    'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments'
  );
  url.searchParams.set('$orderby', 'receivedDateTime desc');

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph mailbox lookup failed for ${mailboxAddress}: ${
        data.error?.message || response.status
      }`
    );
  }

  return data.value || [];
};

const getFullMailboxEmail = async (mailboxAddress, messageId) => {
  if (!mailboxAddress || !messageId) {
    throw new Error('Mailbox address and message id are required.');
  }

  const token = await getMicrosoftGraphToken();
  const url = new URL(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      mailboxAddress
    )}/messages/${encodeURIComponent(messageId)}`
  );
  url.searchParams.set(
    '$select',
    'id,conversationId,internetMessageId,subject,from,replyTo,toRecipients,ccRecipients,receivedDateTime,body,isRead,hasAttachments'
  );

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Prefer: 'outlook.body-content-type="text"',
    },
  });
  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph full email lookup failed for ${mailboxAddress}: ${
        data.error?.message || response.status
      }`
    );
  }

  return data;
};

const searchMailboxEmails = async ({
  mailboxAddress,
  query,
  limit = 10,
  startDate,
  endDate,
}) => {
  if (!mailboxAddress) throw new Error('Mailbox address is required.');
  const searchTerm = String(query || '').trim();
  if (!searchTerm) throw new Error('An email search term is required.');

  const token = await getMicrosoftGraphToken();
  const requestedLimit = Math.min(Math.max(Number(limit) || 10, 1), 25);
  const fetchLimit = Math.min(Math.max(requestedLimit * 4, 25), 100);

  const url = new URL(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      mailboxAddress
    )}/messages`
  );
  const escaped = searchTerm.replace(/"/g, '\\"');
  url.searchParams.set('$search', `"${escaped}"`);
  url.searchParams.set('$top', String(fetchLimit));
  url.searchParams.set(
    '$select',
    'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments'
  );

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph email search failed for ${mailboxAddress}: ${
        data.error?.message || response.status
      }`
    );
  }

  let results = data.value || [];

  if (startDate) {
    const start = new Date(`${String(startDate).slice(0, 10)}T00:00:00Z`);
    if (!Number.isNaN(start.getTime())) {
      results = results.filter(
        (email) => new Date(email.receivedDateTime).getTime() >= start.getTime()
      );
    }
  }

  if (endDate) {
    const end = new Date(`${String(endDate).slice(0, 10)}T23:59:59Z`);
    if (!Number.isNaN(end.getTime())) {
      results = results.filter(
        (email) => new Date(email.receivedDateTime).getTime() <= end.getTime()
      );
    }
  }

  return results.slice(0, requestedLimit);
};

const listMailboxEmailAttachments = async (mailboxAddress, messageId) => {
  if (!mailboxAddress || !messageId) {
    throw new Error('Mailbox address and message id are required.');
  }

  const token = await getMicrosoftGraphToken();
  const url = new URL(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      mailboxAddress
    )}/messages/${encodeURIComponent(messageId)}/attachments`
  );
  url.searchParams.set('$select', 'id,name,contentType,size,isInline');

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph attachment list failed: ${
        data.error?.message || response.status
      }`
    );
  }

  return (data.value || []).map((attachment, index) => ({
    number: index + 1,
    attachmentId: attachment.id,
    name: attachment.name || 'Unnamed attachment',
    contentType: attachment.contentType || '',
    size: Number(attachment.size) || 0,
    isInline: Boolean(attachment.isInline),
    odataType: attachment['@odata.type'] || '',
  }));
};

const getMailboxEmailAttachment = async (
  mailboxAddress,
  messageId,
  attachmentId
) => {
  if (!mailboxAddress || !messageId || !attachmentId) {
    throw new Error('Mailbox address, message id, and attachment id are required.');
  }

  const token = await getMicrosoftGraphToken();
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      mailboxAddress
    )}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(
      attachmentId
    )}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph attachment retrieval failed: ${
        data.error?.message || response.status
      }`
    );
  }

  return data;
};

const extractOpenAIResponseText = (data) => {
  if (typeof data?.output_text === 'string' && data.output_text.trim()) {
    return data.output_text.trim();
  }

  const chunks = [];
  for (const item of Array.isArray(data?.output) ? data.output : []) {
    for (const content of Array.isArray(item?.content) ? item.content : []) {
      if (content?.type === 'output_text' && typeof content.text === 'string') {
        chunks.push(content.text);
      }
    }
  }
  return chunks.join('\n').trim();
};

const callOpenAIResponses = async (payload) => {
  if (!OPENAI_API_KEY) throw new Error('OPENAI_API_KEY is not configured.');

  const response = await fetch('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${OPENAI_API_KEY}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload),
  });

  const data = await response.json();
  if (!response.ok) {
    throw new Error(
      `OpenAI Responses API failed: ${data.error?.message || response.status}`
    );
  }
  return data;
};

const analyzeMailboxAttachment = async ({
  mailboxAddress,
  messageId,
  attachmentId,
  instruction,
}) => {
  const attachment = await getMailboxEmailAttachment(
    mailboxAddress,
    messageId,
    attachmentId
  );

  const size = Number(attachment.size) || 0;
  if (size > MAX_ATTACHMENT_BYTES) {
    throw new Error(
      `The attachment is too large for voice analysis (${Math.round(
        size / 1024 / 1024
      )} MB). The current limit is 45 MB.`
    );
  }

  if (!attachment.contentBytes) {
    throw new Error(
      'This attachment type does not expose file bytes for direct analysis. File attachments such as PDF, Word, Excel, PowerPoint, text, and images are supported.'
    );
  }

  const filename = attachment.name || 'attachment';
  const contentType = attachment.contentType || 'application/octet-stream';
  const dataUri = `data:${contentType};base64,${attachment.contentBytes}`;
  const requestInstruction =
    String(instruction || '').trim() ||
    'Summarize this attachment accurately and identify any decisions, deadlines, amounts, risks, or actions that matter to Ramy.';

  const content = [];
  if (contentType.toLowerCase().startsWith('image/')) {
    content.push({
      type: 'input_image',
      image_url: dataUri,
      detail: 'auto',
    });
  } else {
    content.push({
      type: 'input_file',
      filename,
      file_data: dataUri,
      detail: 'auto',
    });
  }
  content.push({ type: 'input_text', text: requestInstruction });

  const data = await callOpenAIResponses({
    model: DOCUMENT_ANALYSIS_MODEL,
    instructions:
      'You are London Assistant analyzing a live Minaco email attachment for Ramy. Follow the user instruction exactly. Do not invent missing content. Preserve names, dates, amounts, legal terms, and material qualifiers. If asked to translate, translate the complete available document faithfully.',
    input: [{ role: 'user', content }],
  });

  const result = extractOpenAIResponseText(data);
  if (!result) throw new Error('The attachment analysis returned no text.');

  return {
    filename,
    contentType,
    size,
    result,
  };
};

const addAddressCandidate = (map, candidate, query) => {
  const address = candidate?.emailAddress?.address || candidate?.address || '';
  const name = candidate?.emailAddress?.name || candidate?.name || '';
  if (!address) return;

  const normalizedAddress = address.toLowerCase();
  const q = String(query || '').trim().toLowerCase();
  const normalizedName = name.toLowerCase();
  const localPart = normalizedAddress.split('@')[0] || '';

  let score = 0;
  if (q === normalizedAddress || q === normalizedName) score = 100;
  else if (normalizedName.startsWith(q) && q) score = 90;
  else if (normalizedName.includes(q) && q) score = 80;
  else if (localPart.includes(q.replace(/\s+/g, '.')) && q) score = 70;
  else if (normalizedAddress.includes(q.replace(/\s+/g, '')) && q) score = 60;
  else if (q && normalizedAddress.includes(q.split(/\s+/)[0])) score = 50;

  if (score <= 0) return;

  const existing = map.get(normalizedAddress);
  if (!existing || score > existing.score) {
    map.set(normalizedAddress, { name, address, score });
  }
};

const resolvePersonFromMailHistory = async (query) => {
  const q = String(query || '').trim();
  if (!q) throw new Error('A person name or email is required.');

  const candidates = new Map();

  const verified = [
    { name: 'Ramy Mina', address: RAMY_MINACO_EMAIL },
    { name: 'London Assistant', address: LONDON_MINACO_EMAIL },
    { name: 'Minaco Accounting', address: ACCOUNTING_MAILBOX },
  ].filter((item) => item.address);

  for (const item of verified) addAddressCandidate(candidates, item, q);

  let messages = [];
  try {
    const [searched, recent] = await Promise.all([
      searchMailboxEmails({
        mailboxAddress: RAMY_MINACO_EMAIL,
        query: q,
        limit: 25,
      }),
      getRecentMailboxEmails(RAMY_MINACO_EMAIL, 25),
    ]);
    messages = [...searched, ...recent];
  } catch (error) {
    console.warn('Contact resolution mailbox search warning:', error.message);
  }

  for (const message of messages) {
    addAddressCandidate(candidates, message.from, q);
    for (const recipient of message.toRecipients || []) {
      addAddressCandidate(candidates, recipient, q);
    }
    for (const recipient of message.ccRecipients || []) {
      addAddressCandidate(candidates, recipient, q);
    }
  }

  return [...candidates.values()]
    .sort((a, b) => b.score - a.score || a.address.localeCompare(b.address))
    .slice(0, 8);
};

const simplifyEmailRecipients = (recipients = []) =>
  (Array.isArray(recipients) ? recipients : [])
    .map((recipient) => ({
      name: recipient?.emailAddress?.name || '',
      address: recipient?.emailAddress?.address || '',
    }))
    .filter((recipient) => recipient.address);

const getReplyRecipientSummary = (email, mode) => {
  const ownAddress = String(RAMY_MINACO_EMAIL || '').toLowerCase();
  const firstRecipients =
    Array.isArray(email?.replyTo) && email.replyTo.length > 0
      ? email.replyTo
      : email?.from
        ? [email.from]
        : [];

  const sourceRecipients =
    mode === 'all'
      ? [
          ...firstRecipients,
          ...(Array.isArray(email?.toRecipients) ? email.toRecipients : []),
          ...(Array.isArray(email?.ccRecipients) ? email.ccRecipients : []),
        ]
      : firstRecipients;

  const seen = new Set();
  const result = [];

  for (const recipient of simplifyEmailRecipients(sourceRecipients)) {
    const normalized = recipient.address.toLowerCase();
    if (!normalized || normalized === ownAddress || seen.has(normalized)) continue;
    seen.add(normalized);
    result.push(recipient);
  }

  return result;
};

const replyToMinacoEmail = async ({ messageId, mode, body }) => {
  if (!RAMY_MINACO_EMAIL) {
    throw new Error('RAMY_MINACO_EMAIL is not configured.');
  }

  if (!messageId || !body) {
    throw new Error('Reply requires a message id and reply body.');
  }

  if (!['sender', 'all'].includes(mode)) {
    throw new Error('Reply mode must be sender or all.');
  }

  const token = await getMicrosoftGraphActionsToken();
  const action = mode === 'all' ? 'replyAll' : 'reply';

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      RAMY_MINACO_EMAIL
    )}/messages/${encodeURIComponent(messageId)}/${action}`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ comment: body }),
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(
      `Microsoft Graph email ${action} failed: ${errorText || response.status}`
    );
  }

  return {
    success: true,
    from: RAMY_MINACO_EMAIL,
    mode,
    messageId,
  };
};

const sendEmailFromLondon = async ({ to, subject, body, contentType = 'Text' }) => {
  if (!LONDON_MINACO_EMAIL) {
    throw new Error('LONDON_MINACO_EMAIL is not configured.');
  }

  if (!to || !subject || !body) {
    throw new Error('Email requires recipient, subject, and body.');
  }

  const normalizedContentType =
    String(contentType || '').toLowerCase() === 'html' ? 'HTML' : 'Text';

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
            contentType: normalizedContentType,
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
    contentType: normalizedContentType,
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
    signal: AbortSignal.timeout(12000),
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

const parseLocalDateTimePseudoMs = (value) => {
  const normalized = ensureLocalDateTime(value, 'local_date_time');
  const match = normalized.match(
    /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/
  );
  if (!match) throw new Error('Invalid local date/time.');
  const [, year, month, day, hour, minute, second] = match;
  return Date.UTC(
    Number(year),
    Number(month) - 1,
    Number(day),
    Number(hour),
    Number(minute),
    Number(second)
  );
};

const addDaysToLocalDate = (dateString, days) => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateString || ''))) {
    throw new Error('Date must use YYYY-MM-DD.');
  }
  const [year, month, day] = dateString.split('-').map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  date.setUTCDate(date.getUTCDate() + Number(days || 0));
  return date.toISOString().slice(0, 10);
};

const validateClockTime = (value, label) => {
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(String(value || ''))) {
    throw new Error(`${label} must use HH:mm in 24-hour Montreal local time.`);
  }
  return value;
};

const findCalendarAvailability = async ({
  startDate,
  endDate,
  durationMinutes = 30,
  dayStart = '09:00',
  dayEnd = '17:00',
  maxSlots = 5,
  includeWeekends = false,
}) => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(startDate || ''))) {
    throw new Error('start_date must use YYYY-MM-DD.');
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(endDate || ''))) {
    throw new Error('end_date must use YYYY-MM-DD.');
  }

  validateClockTime(dayStart, 'day_start');
  validateClockTime(dayEnd, 'day_end');

  const duration = Math.min(Math.max(Number(durationMinutes) || 30, 15), 240);
  const limit = Math.min(Math.max(Number(maxSlots) || 5, 1), 12);

  const rangeStartLocal = `${startDate}T00:00:00`;
  const dayAfterEnd = addDaysToLocalDate(endDate, 1);
  const rangeEndLocal = `${dayAfterEnd}T00:00:00`;

  const rangeStartMs = parseLocalDateTimePseudoMs(rangeStartLocal);
  const rangeEndMs = parseLocalDateTimePseudoMs(rangeEndLocal);
  if (rangeEndMs <= rangeStartMs) {
    throw new Error('end_date must be on or after start_date.');
  }
  if (rangeEndMs - rangeStartMs > 31 * 24 * 60 * 60 * 1000) {
    throw new Error('Availability searches are limited to 31 days at a time.');
  }

  const events = (
    await fetchCalendarView({
      startUtc: montrealLocalToUtcIso(rangeStartLocal),
      endUtc: montrealLocalToUtcIso(rangeEndLocal),
      top: 100,
    })
  )
    .map(simplifyCalendarEvent)
    .filter((event) => !event.isCancelled && event.showAs !== 'free')
    .map((event) => ({
      ...event,
      startMs: parseLocalDateTimePseudoMs(event.startLocal),
      endMs: parseLocalDateTimePseudoMs(event.endLocal),
    }));

  const slots = [];
  let currentDate = startDate;
  while (currentDate <= endDate && slots.length < limit) {
    const [year, month, day] = currentDate.split('-').map(Number);
    const weekday = new Date(Date.UTC(year, month - 1, day)).getUTCDay();
    const isWeekend = weekday === 0 || weekday === 6;

    if (includeWeekends || !isWeekend) {
      const windowStart = `${currentDate}T${dayStart}:00`;
      const windowEnd = `${currentDate}T${dayEnd}:00`;
      let cursorMs = parseLocalDateTimePseudoMs(windowStart);
      const windowEndMs = parseLocalDateTimePseudoMs(windowEnd);
      const durationMs = duration * 60 * 1000;
      const stepMs = 30 * 60 * 1000;

      while (cursorMs + durationMs <= windowEndMs && slots.length < limit) {
        const candidateEndMs = cursorMs + durationMs;
        const conflict = events.some(
          (event) => event.startMs < candidateEndMs && event.endMs > cursorMs
        );

        if (!conflict) {
          const startIso = new Date(cursorMs).toISOString().slice(0, 19);
          const endIso = new Date(candidateEndMs).toISOString().slice(0, 19);
          slots.push({
            startLocal: startIso,
            endLocal: endIso,
            startMontreal: formatGraphEasternDateTime(startIso),
            endMontreal: formatGraphEasternDateTime(endIso),
            durationMinutes: duration,
          });
        }
        cursorMs += stepMs;
      }
    }

    currentDate = addDaysToLocalDate(currentDate, 1);
  }

  return {
    startDate,
    endDate,
    durationMinutes: duration,
    dayStart,
    dayEnd,
    includeWeekends: Boolean(includeWeekends),
    slots,
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
// Persistent Master Action Register (stored in a dedicated Outlook calendar)
// -----------------------------------------------------------------------------

const ACTION_REGISTER_MARKER = 'LONDON_ACTION_V1';
const ACTION_STATUSES = [
  'NEW',
  'ACTIVE',
  'WAITING - EXTERNAL',
  'WAITING - INTERNAL',
  'WAITING - RAMY',
  'BLOCKED',
  'OVERDUE',
  'COMPLETED',
  'CANCELLED',
];
const ACTION_PRIORITIES = ['CRITICAL', 'HIGH', 'NORMAL', 'LOW'];

const montrealDateParts = (date = new Date()) => {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: MONTREAL_IANA_TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(date);
  const values = {};
  for (const part of parts) {
    if (part.type !== 'literal') values[part.type] = part.value;
  }
  return `${values.year}-${values.month}-${values.day}`;
};

const addDaysToDateOnly = (dateString, days) => {
  const match = String(dateString || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return montrealDateParts();
  const date = new Date(
    Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3]) + days)
  );
  return date.toISOString().slice(0, 10);
};

const normalizeActionStatus = (status) => {
  if (!status) return 'NEW';
  const normalized = String(status)
    .trim()
    .toUpperCase()
    .replace(/[–—]/g, '-')
    .replace(/\s+/g, ' ');
  const aliases = {
    WAITING: 'WAITING - EXTERNAL',
    'WAITING EXTERNAL': 'WAITING - EXTERNAL',
    'WAITING INTERNAL': 'WAITING - INTERNAL',
    'WAITING RAMY': 'WAITING - RAMY',
    DONE: 'COMPLETED',
    COMPLETE: 'COMPLETED',
    CANCELED: 'CANCELLED',
  };
  const mapped = aliases[normalized] || normalized;
  return ACTION_STATUSES.includes(mapped) ? mapped : 'ACTIVE';
};

const normalizeActionPriority = (priority) => {
  const normalized = String(priority || 'NORMAL').trim().toUpperCase();
  return ACTION_PRIORITIES.includes(normalized) ? normalized : 'NORMAL';
};

let cachedActionCalendarId = null;

const getOrCreateActionRegisterCalendar = async () => {
  if (cachedActionCalendarId) return cachedActionCalendarId;
  if (!RAMY_MINACO_EMAIL) throw new Error('RAMY_MINACO_EMAIL is not configured.');

  const token = await getMicrosoftGraphActionsToken();
  const listResponse = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      RAMY_MINACO_EMAIL
    )}/calendars?$select=id,name`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const listData = await listResponse.json();
  if (!listResponse.ok) {
    throw new Error(
      `Could not list calendars for the action register: ${
        listData.error?.message || listResponse.status
      }`
    );
  }

  const existing = (listData.value || []).find(
    (calendar) =>
      String(calendar.name || '').toLowerCase() ===
      ACTION_REGISTER_CALENDAR_NAME.toLowerCase()
  );
  if (existing?.id) {
    cachedActionCalendarId = existing.id;
    return existing.id;
  }

  const createResponse = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      RAMY_MINACO_EMAIL
    )}/calendars`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ name: ACTION_REGISTER_CALENDAR_NAME }),
    }
  );
  const createData = await createResponse.json();
  if (!createResponse.ok) {
    throw new Error(
      `Could not create the action register calendar: ${
        createData.error?.message || createResponse.status
      }`
    );
  }

  cachedActionCalendarId = createData.id;
  return createData.id;
};

const actionAnchorDate = (action) => {
  for (const candidate of [
    action.nextFollowUp,
    action.hardDeadline,
    action.promisedDate,
    action.dateOpened,
  ]) {
    const match = String(candidate || '').match(/^(\d{4}-\d{2}-\d{2})/);
    if (match) return match[1];
  }
  return montrealDateParts();
};

const serializeActionBody = (action) =>
  `${ACTION_REGISTER_MARKER}\n${JSON.stringify(action, null, 2)}`;

const parseActionBody = (content) => {
  const raw = String(content || '');
  const markerIndex = raw.indexOf(ACTION_REGISTER_MARKER);
  if (markerIndex < 0) return null;
  const jsonStart = raw.indexOf('{', markerIndex + ACTION_REGISTER_MARKER.length);
  if (jsonStart < 0) return null;
  try {
    return JSON.parse(raw.slice(jsonStart));
  } catch {
    return null;
  }
};

const actionEventPayload = (action) => {
  const anchor = actionAnchorDate(action);
  const nextDay = addDaysToDateOnly(anchor, 1);
  return {
    subject: `[${normalizeActionPriority(action.priority)}] [${normalizeActionStatus(
      action.status
    )}] ${action.title}`,
    body: {
      contentType: 'Text',
      content: serializeActionBody(action),
    },
    isAllDay: true,
    showAs: 'free',
    isReminderOn: false,
    sensitivity: 'private',
    start: {
      dateTime: `${anchor}T00:00:00`,
      timeZone: MICROSOFT_EASTERN_TIME_ZONE,
    },
    end: {
      dateTime: `${nextDay}T00:00:00`,
      timeZone: MICROSOFT_EASTERN_TIME_ZONE,
    },
  };
};

const createActionItem = async (input) => {
  if (!input?.title) throw new Error('Action title is required.');
  const calendarId = await getOrCreateActionRegisterCalendar();
  const token = await getMicrosoftGraphActionsToken();
  const now = new Date().toISOString();

  const action = {
    actionId: randomUUID(),
    title: String(input.title).trim(),
    project: String(input.project || '').trim(),
    category: String(input.category || '').trim(),
    owner: String(input.owner || 'London').trim(),
    dateOpened: String(input.dateOpened || montrealDateParts()).slice(0, 10),
    promisedDate: String(input.promisedDate || '').trim(),
    hardDeadline: String(input.hardDeadline || '').trim(),
    nextFollowUp: String(input.nextFollowUp || '').trim(),
    status: normalizeActionStatus(input.status),
    priority: normalizeActionPriority(input.priority),
    lastContact: String(input.lastContact || '').trim(),
    nextAction: String(input.nextAction || '').trim(),
    waitingOn: String(input.waitingOn || '').trim(),
    ramyRequired: Boolean(input.ramyRequired),
    ramyDecisionBy: String(input.ramyDecisionBy || '').trim(),
    riskIfDelayed: String(input.riskIfDelayed || '').trim(),
    source: String(input.source || 'Voice').trim(),
    notes: String(input.notes || '').trim(),
    createdAt: now,
    updatedAt: now,
  };

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      RAMY_MINACO_EMAIL
    )}/calendars/${encodeURIComponent(calendarId)}/events`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        Prefer: `outlook.timezone="${MICROSOFT_EASTERN_TIME_ZONE}"`,
      },
      body: JSON.stringify(actionEventPayload(action)),
    }
  );
  const data = await response.json();
  if (!response.ok) {
    throw new Error(
      `Could not create action item: ${data.error?.message || response.status}`
    );
  }

  return { eventId: data.id, ...action };
};

const listActionItems = async ({
  status,
  priority,
  project,
  owner,
  onlyOverdue = false,
  includeClosed = false,
  limit = 100,
} = {}) => {
  const calendarId = await getOrCreateActionRegisterCalendar();
  const token = await getMicrosoftGraphActionsToken();
  let url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
    RAMY_MINACO_EMAIL
  )}/calendars/${encodeURIComponent(
    calendarId
  )}/events?$top=100&$select=id,subject,body,start,end,lastModifiedDateTime`;

  const events = [];
  while (url && events.length < 500) {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Prefer: 'outlook.body-content-type="text"',
      },
    });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(
        `Could not read action register: ${data.error?.message || response.status}`
      );
    }
    events.push(...(data.value || []));
    url = data['@odata.nextLink'] || null;
  }

  const today = montrealDateParts();
  let actions = events
    .map((event) => {
      const meta = parseActionBody(event.body?.content);
      if (!meta) return null;
      const normalizedStatus = normalizeActionStatus(meta.status);
      const isClosed = ['COMPLETED', 'CANCELLED'].includes(normalizedStatus);
      const deadlineOverdue =
        meta.hardDeadline && String(meta.hardDeadline).slice(0, 10) < today;
      const followUpOverdue =
        meta.nextFollowUp && String(meta.nextFollowUp).slice(0, 10) < today;
      return {
        eventId: event.id,
        ...meta,
        status: normalizedStatus,
        priority: normalizeActionPriority(meta.priority),
        overdue: !isClosed && Boolean(deadlineOverdue || followUpOverdue),
        lastModifiedDateTime: event.lastModifiedDateTime,
      };
    })
    .filter(Boolean);

  if (!includeClosed) {
    actions = actions.filter(
      (action) => !['COMPLETED', 'CANCELLED'].includes(action.status)
    );
  }
  if (status) {
    const wanted = normalizeActionStatus(status);
    actions = actions.filter((action) => action.status === wanted);
  }
  if (priority) {
    const wanted = normalizeActionPriority(priority);
    actions = actions.filter((action) => action.priority === wanted);
  }
  if (project) {
    const q = String(project).toLowerCase();
    actions = actions.filter((action) =>
      String(action.project || '').toLowerCase().includes(q)
    );
  }
  if (owner) {
    const q = String(owner).toLowerCase();
    actions = actions.filter((action) =>
      String(action.owner || '').toLowerCase().includes(q)
    );
  }
  if (onlyOverdue) actions = actions.filter((action) => action.overdue);

  const priorityRank = { CRITICAL: 0, HIGH: 1, NORMAL: 2, LOW: 3 };
  actions.sort((a, b) => {
    if (a.overdue !== b.overdue) return a.overdue ? -1 : 1;
    const pr = priorityRank[a.priority] - priorityRank[b.priority];
    if (pr !== 0) return pr;
    return String(a.hardDeadline || a.nextFollowUp || '9999-12-31').localeCompare(
      String(b.hardDeadline || b.nextFollowUp || '9999-12-31')
    );
  });

  return actions.slice(0, Math.min(Math.max(Number(limit) || 100, 1), 200));
};

const updateActionItem = async (eventId, changes = {}) => {
  if (!eventId) throw new Error('An action event id is required.');
  const calendarId = await getOrCreateActionRegisterCalendar();
  const token = await getMicrosoftGraphActionsToken();

  const getResponse = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      RAMY_MINACO_EMAIL
    )}/calendars/${encodeURIComponent(calendarId)}/events/${encodeURIComponent(
      eventId
    )}?$select=id,subject,body,start,end`,
    {
      headers: {
        Authorization: `Bearer ${token}`,
        Prefer: 'outlook.body-content-type="text"',
      },
    }
  );
  const current = await getResponse.json();
  if (!getResponse.ok) {
    throw new Error(
      `Could not retrieve action item: ${current.error?.message || getResponse.status}`
    );
  }

  const existing = parseActionBody(current.body?.content);
  if (!existing) throw new Error('The selected event is not a London action item.');

  const allowedFields = [
    'title',
    'project',
    'category',
    'owner',
    'promisedDate',
    'hardDeadline',
    'nextFollowUp',
    'status',
    'priority',
    'lastContact',
    'nextAction',
    'waitingOn',
    'ramyRequired',
    'ramyDecisionBy',
    'riskIfDelayed',
    'source',
    'notes',
  ];

  const merged = { ...existing };
  for (const field of allowedFields) {
    if (Object.prototype.hasOwnProperty.call(changes, field)) {
      merged[field] = changes[field];
    }
  }
  merged.status = normalizeActionStatus(merged.status);
  merged.priority = normalizeActionPriority(merged.priority);
  merged.updatedAt = new Date().toISOString();

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      RAMY_MINACO_EMAIL
    )}/calendars/${encodeURIComponent(calendarId)}/events/${encodeURIComponent(
      eventId
    )}`,
    {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(actionEventPayload(merged)),
    }
  );
  const data = await response.json();
  if (!response.ok) {
    throw new Error(
      `Could not update action item: ${data.error?.message || response.status}`
    );
  }

  return { eventId: data.id || eventId, ...merged };
};

// -----------------------------------------------------------------------------
// Daily executive brief
// -----------------------------------------------------------------------------

const compactEmailForBrief = (email) => ({
  id: email.id,
  from:
    email.from?.emailAddress?.name || email.from?.emailAddress?.address || '',
  fromEmail: email.from?.emailAddress?.address || '',
  subject: email.subject || '(No subject)',
  received: email.receivedDateTime || '',
  unread: !email.isRead,
  hasAttachments: Boolean(email.hasAttachments),
  preview: email.bodyPreview || '',
});

const escapeHtml = (value) =>
  String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');

const stripJsonCodeFence = (value) =>
  String(value || '')
    .trim()
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/\s*```$/i, '')
    .trim();

const formatBriefDisplayDate = (date = new Date()) =>
  new Intl.DateTimeFormat('en-US', {
    timeZone: MONTREAL_IANA_TIME_ZONE,
    weekday: 'long',
    month: 'short',
    day: 'numeric',
    year: 'numeric',
  }).format(date);

const formatBriefPreparedTime = (date = new Date()) =>
  new Intl.DateTimeFormat('en-US', {
    timeZone: MONTREAL_IANA_TIME_ZONE,
    hour: 'numeric',
    minute: '2-digit',
    timeZoneName: 'short',
  }).format(date);

const normalizeBriefItem = (item = {}) => ({
  title: String(item.title || '').trim(),
  badge: String(item.badge || '').trim(),
  fact: String(item.fact || '').trim(),
  nextAction: String(item.nextAction || '').trim(),
  note: String(item.note || '').trim(),
});

const normalizeBriefData = (data = {}) => {
  const array = (value) =>
    Array.isArray(value) ? value.map(normalizeBriefItem).filter((item) => item.title) : [];

  const schedule = Array.isArray(data.schedule)
    ? data.schedule
        .map((item = {}) => ({
          time: String(item.time || '').trim(),
          title: String(item.title || '').trim(),
          detail: String(item.detail || '').trim(),
        }))
        .filter((item) => item.time || item.title || item.detail)
    : [];

  const priorities = Array.isArray(data.londonPriorities)
    ? data.londonPriorities.map((item) => String(item || '').trim()).filter(Boolean)
    : [];

  return {
    attentionSummary: String(data.attentionSummary || '').trim(),
    urgentDecisions: array(data.urgentDecisions),
    deadlinesRisks: array(data.deadlinesRisks),
    schedule,
    scheduleSummary: String(data.scheduleSummary || '').trim(),
    followUps: array(data.followUps),
    accounting: array(data.accounting),
    londonPriorities: priorities,
    closing: String(data.closing || '').trim(),
  };
};

const briefItemCardHtml = (item, accentColor) => {
  const badge = item.badge
    ? `<span style="display:inline-block;background:${accentColor};color:#ffffff;font-size:10px;line-height:14px;font-weight:700;letter-spacing:.4px;text-transform:uppercase;padding:3px 8px;border-radius:10px;white-space:nowrap;">${escapeHtml(item.badge)}</span>`
    : '';

  const fact = item.fact
    ? `<div style="margin-top:7px;color:#475467;font-size:13px;line-height:19px;"><span style="font-weight:700;color:#344054;">Fact:</span> ${escapeHtml(item.fact)}</div>`
    : '';

  const nextAction = item.nextAction
    ? `<div style="margin-top:5px;color:#175cd3;font-size:13px;line-height:19px;"><span style="font-weight:700;">Next action:</span> ${escapeHtml(item.nextAction)}</div>`
    : '';

  const note = item.note
    ? `<div style="margin-top:5px;color:#667085;font-size:12px;line-height:18px;">${escapeHtml(item.note)}</div>`
    : '';

  return `
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:0;margin:0 0 10px 0;background:#ffffff;border:1px solid #eaecf0;border-left:4px solid ${accentColor};border-radius:7px;">
      <tr>
        <td style="padding:12px 14px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
            <tr>
              <td style="font-size:14px;line-height:20px;font-weight:700;color:#101828;padding-right:10px;">${escapeHtml(item.title)}</td>
              <td align="right" valign="top" style="width:1%;">${badge}</td>
            </tr>
          </table>
          ${fact}${nextAction}${note}
        </td>
      </tr>
    </table>`;
};

const briefSectionHtml = ({ title, subtitle, color, items, emptyText }) => {
  const content = items.length
    ? items.map((item) => briefItemCardHtml(item, color)).join('')
    : `<div style="background:#ffffff;border:1px solid #eaecf0;border-radius:7px;padding:12px 14px;color:#667085;font-size:13px;line-height:19px;">${escapeHtml(emptyText)}</div>`;

  return `
    <tr>
      <td style="padding:0 24px 18px 24px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
          <tr>
            <td style="padding:0 0 9px 0;border-bottom:2px solid ${color};">
              <span style="font-size:13px;line-height:18px;font-weight:800;letter-spacing:.45px;color:${color};text-transform:uppercase;">${escapeHtml(title)}</span>
              ${subtitle ? `<span style="display:block;margin-top:2px;color:#667085;font-size:11px;line-height:16px;">${escapeHtml(subtitle)}</span>` : ''}
            </td>
          </tr>
          <tr><td style="padding-top:10px;">${content}</td></tr>
        </table>
      </td>
    </tr>`;
};

const scheduleSectionHtml = (brief) => {
  const color = '#175CD3';
  let content = '';

  if (brief.schedule.length) {
    content = brief.schedule
      .map((item) => `
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:0;margin:0 0 8px 0;background:#ffffff;border:1px solid #eaecf0;border-radius:7px;">
          <tr>
            <td valign="top" style="width:92px;padding:11px 12px;color:${color};font-size:13px;line-height:19px;font-weight:800;white-space:nowrap;">${escapeHtml(item.time)}</td>
            <td style="padding:11px 12px 11px 0;color:#101828;font-size:13px;line-height:19px;"><span style="font-weight:700;">${escapeHtml(item.title)}</span>${item.detail ? `<div style="margin-top:2px;color:#667085;font-size:12px;line-height:18px;">${escapeHtml(item.detail)}</div>` : ''}</td>
          </tr>
        </table>`)
      .join('');
  } else {
    content = `<div style="background:#ffffff;border:1px solid #eaecf0;border-radius:7px;padding:12px 14px;color:#667085;font-size:13px;line-height:19px;">${escapeHtml(brief.scheduleSummary || 'No calendar meetings scheduled today.')}</div>`;
  }

  return `
    <tr>
      <td style="padding:0 24px 18px 24px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
          <tr><td style="padding:0 0 9px 0;border-bottom:2px solid ${color};"><span style="font-size:13px;line-height:18px;font-weight:800;letter-spacing:.45px;color:${color};text-transform:uppercase;">Today’s Schedule</span></td></tr>
          <tr><td style="padding-top:10px;">${content}</td></tr>
        </table>
      </td>
    </tr>`;
};

const prioritiesSectionHtml = (brief) => {
  const priorities = brief.londonPriorities.length
    ? brief.londonPriorities
    : ['Continue monitoring Minaco email, Accounting, calendar, and tracked follow-ups.'];

  return `
    <tr>
      <td style="padding:0 24px 18px 24px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
          <tr><td style="padding:0 0 9px 0;border-bottom:2px solid #344054;"><span style="font-size:13px;line-height:18px;font-weight:800;letter-spacing:.45px;color:#344054;text-transform:uppercase;">London’s Priorities Today</span></td></tr>
          <tr>
            <td style="padding-top:10px;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:0;background:#ffffff;border:1px solid #eaecf0;border-radius:7px;">
                ${priorities
                  .map((priority, index) => `<tr><td valign="top" style="width:34px;padding:10px 0 10px 13px;color:#344054;font-size:13px;font-weight:800;">${index + 1}.</td><td style="padding:10px 13px 10px 4px;color:#344054;font-size:13px;line-height:19px;${index ? 'border-top:1px solid #f2f4f7;' : ''}">${escapeHtml(priority)}</td></tr>`)
                  .join('')}
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>`;
};

const renderDailyBriefHtml = (brief, generatedAt = new Date()) => {
  const displayDate = formatBriefDisplayDate(generatedAt);
  const preparedTime = formatBriefPreparedTime(generatedAt);
  const summary = brief.attentionSummary || 'Minaco is being monitored. Key items are summarized below.';
  const closing = brief.closing || 'Everything else is being tracked by London.';

  const sections = [
    briefSectionHtml({
      title: 'Needs Your Decision',
      subtitle: 'Items that specifically require Ramy',
      color: '#B42318',
      items: brief.urgentDecisions,
      emptyText: 'No immediate decisions requiring your attention.',
    }),
    briefSectionHtml({
      title: 'Deadlines & Financial Risks',
      subtitle: 'Time-sensitive, financial, contractual, or project exposure',
      color: '#B54708',
      items: brief.deadlinesRisks,
      emptyText: 'No material deadline or financial risk identified in the live data.',
    }),
    scheduleSectionHtml(brief),
    briefSectionHtml({
      title: 'Follow-Ups / Waiting',
      subtitle: 'Commitments London is tracking',
      color: '#6941C6',
      items: brief.followUps,
      emptyText: 'No tracked follow-up currently needs attention.',
    }),
    briefSectionHtml({
      title: 'Accounting',
      subtitle: 'Invoices, payments, statements, taxes, and finance items',
      color: '#027A48',
      items: brief.accounting,
      emptyText: 'No accounting item currently requires attention.',
    }),
    prioritiesSectionHtml(brief),
  ].join('');

  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>London Executive Brief</title>
</head>
<body style="margin:0;padding:0;background:#f2f4f7;font-family:Arial,Helvetica,sans-serif;color:#101828;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;background:#f2f4f7;">
    <tr>
      <td align="center" style="padding:24px 10px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:760px;border-collapse:separate;border-spacing:0;background:#f8fafc;border:1px solid #e4e7ec;border-radius:12px;overflow:hidden;">
          <tr>
            <td style="background:#102A43;padding:23px 24px 20px 24px;border-bottom:4px solid #C8A96B;">
              <div style="font-size:11px;line-height:16px;font-weight:700;letter-spacing:1.8px;color:#C8A96B;text-transform:uppercase;">London Assistant · Minaco</div>
              <div style="margin-top:5px;font-size:24px;line-height:31px;font-weight:800;color:#ffffff;">Daily Executive Brief</div>
              <div style="margin-top:5px;font-size:12px;line-height:18px;color:#d0d5dd;">${escapeHtml(displayDate)} · Prepared ${escapeHtml(preparedTime)} Montreal</div>
            </td>
          </tr>
          <tr>
            <td style="padding:18px 24px 20px 24px;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:0;background:#ffffff;border:1px solid #d0d5dd;border-radius:8px;">
                <tr>
                  <td style="padding:14px 16px;">
                    <div style="font-size:11px;line-height:16px;font-weight:800;letter-spacing:.55px;color:#667085;text-transform:uppercase;">Executive Summary</div>
                    <div style="margin-top:3px;font-size:15px;line-height:22px;font-weight:700;color:#101828;">${escapeHtml(summary)}</div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          ${sections}
          <tr>
            <td style="padding:0 24px 24px 24px;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:0;background:#102A43;border-radius:8px;">
                <tr>
                  <td style="padding:14px 16px;color:#ffffff;font-size:13px;line-height:19px;font-weight:700;">${escapeHtml(closing)}</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td align="center" style="padding:0 24px 18px 24px;color:#98a2b3;font-size:10px;line-height:15px;">Generated from live Minaco email, Accounting, calendar, and London Action Register data.</td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;
};

const buildDailyBriefVoiceText = (brief) => {
  const lines = [];
  if (brief.attentionSummary) lines.push(brief.attentionSummary);

  const addItems = (heading, items) => {
    if (!items.length) return;
    lines.push(heading);
    for (const item of items) {
      const parts = [item.title];
      if (item.fact) parts.push(`Fact: ${item.fact}`);
      if (item.nextAction) parts.push(`Next action: ${item.nextAction}`);
      lines.push(parts.join('. '));
    }
  };

  addItems('Needs your decision.', brief.urgentDecisions);
  addItems('Deadlines and financial risks.', brief.deadlinesRisks);

  if (brief.schedule.length) {
    lines.push('Today’s schedule.');
    for (const event of brief.schedule) {
      lines.push(`${event.time}${event.title ? `, ${event.title}` : ''}${event.detail ? `. ${event.detail}` : ''}`);
    }
  } else if (brief.scheduleSummary) {
    lines.push(`Today’s schedule. ${brief.scheduleSummary}`);
  }

  addItems('Follow-ups.', brief.followUps);
  addItems('Accounting.', brief.accounting);

  if (brief.londonPriorities.length) {
    lines.push('London’s priorities today.');
    brief.londonPriorities.forEach((priority, index) => {
      lines.push(`${index + 1}. ${priority}`);
    });
  }

  if (brief.closing) lines.push(brief.closing);
  return lines.join('\n');
};

const generateDailyExecutiveBrief = async () => {
  const today = montrealDateParts();
  const tomorrow = addDaysToDateOnly(today, 1);
  const dayAfter = addDaysToDateOnly(today, 2);

  const [calendarEvents, ramyEmails, accountingEmails, actions] = await Promise.all([
    getCalendarEvents({
      startLocal: `${today}T00:00:00`,
      endLocal: `${dayAfter}T00:00:00`,
      top: 30,
    }),
    getRecentMailboxEmails(RAMY_MINACO_EMAIL, 20),
    getRecentMailboxEmails(ACCOUNTING_MAILBOX, 15).catch((error) => {
      console.warn('Accounting brief lookup warning:', error.message);
      return [];
    }),
    listActionItems({ includeClosed: false, limit: 100 }).catch((error) => {
      console.warn('Action brief lookup warning:', error.message);
      return [];
    }),
  ]);

  const generatedAt = new Date();
  const context = {
    generatedAtMontreal: formatMontrealDateTime(generatedAt),
    today,
    tomorrow,
    calendar: calendarEvents,
    executiveInbox: ramyEmails.map(compactEmailForBrief),
    accountingInbox: accountingEmails.map(compactEmailForBrief),
    actions,
  };

  const data = await callOpenAIResponses({
    model: DAILY_BRIEF_MODEL,
    instructions: `You are London Assistant preparing Ramy Mina's Minaco Daily Executive Brief.
Use ONLY the supplied live data. Never invent names, amounts, dates, deadlines, obligations, or status.
Prioritize: Ramy decisions, urgent deadlines, financial/contractual/project risk, today's meetings, overdue or promised follow-ups, people waiting on Ramy, and material Accounting items.
Do not turn every email into an action. Separate verified facts from recommended next actions.
Return ONLY valid JSON with no Markdown, no code fences, and exactly this shape:
{
  "attentionSummary": "one short sentence stating how many/what kind of items need Ramy's attention",
  "urgentDecisions": [{"title":"...","badge":"ACTION or short label","fact":"...","nextAction":"...","note":"optional"}],
  "deadlinesRisks": [{"title":"...","badge":"DUE SEP 1 / OVERDUE / FINANCIAL etc","fact":"...","nextAction":"...","note":"optional"}],
  "schedule": [{"time":"9:30 AM","title":"...","detail":"optional short detail"}],
  "scheduleSummary": "use only when there are no meetings",
  "followUps": [{"title":"...","badge":"OVERDUE / SEP 3 / WAITING etc","fact":"...","nextAction":"...","note":"optional"}],
  "accounting": [{"title":"...","badge":"PAYMENT / TAX / INVOICE etc","fact":"...","nextAction":"...","note":"optional"}],
  "londonPriorities": ["what London should actively drive today", "..."],
  "closing": "one reassuring sentence distinguishing Ramy's attention from what London is tracking"
}
Keep each fact and next action concise. Put an item in urgentDecisions only if Ramy personally needs to decide, approve, answer, sign, or provide availability. Put financial/tax/invoice/payment risk under deadlinesRisks and/or accounting as appropriate, without duplicating the same wording unnecessarily.`,
    input: `LIVE MINACO DATA:\n${JSON.stringify(context)}`,
  });

  const raw = extractOpenAIResponseText(data);
  if (!raw) throw new Error('Daily executive brief returned no text.');

  let parsed;
  try {
    parsed = JSON.parse(stripJsonCodeFence(raw));
  } catch (error) {
    console.error('Daily executive brief JSON parse error:', error, 'Raw:', raw);
    // Safe fallback: preserve the generated content rather than failing the scheduled brief.
    parsed = {
      attentionSummary: 'London generated the live brief, but its structured formatting could not be parsed.',
      urgentDecisions: [],
      deadlinesRisks: [],
      schedule: [],
      scheduleSummary: 'See the detail below.',
      followUps: [],
      accounting: [],
      londonPriorities: [],
      closing: raw,
    };
  }

  const brief = normalizeBriefData(parsed);
  const voiceText = buildDailyBriefVoiceText(brief);
  const html = renderDailyBriefHtml(brief, generatedAt);

  return {
    data: brief,
    voiceText,
    html,
    subject: `LONDON — Executive Brief | ${formatBriefDisplayDate(generatedAt)}`,
  };
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
check_email returns a recent-email list and message ids. It does NOT contain the entire body.
Use read_email whenever Ramy asks you to read, translate, summarize, analyze, or respond based on the complete contents of a specific email. read_email retrieves the full live message body in text form.
If the target email has not yet been identified, use check_email first, identify the exact email, then use read_email with its message id. Never translate or analyze an email from a preview when the full body is available.

Use prepare_email when Ramy asks you to draft, write, prepare, or send a NEW standalone email from London Assistant. Preparing an email NEVER sends it.
After prepare_email succeeds, read back the recipient, subject, and message and clearly say it has not been sent. Ask Ramy to confirm.

For a REPLY to an email in Ramy's Minaco inbox:
1. Identify the exact email and message id with check_email and, when the reply depends on its contents, read_email.
2. If Ramy says reply, respond, or answer without specifying the recipient scope, ask one short question: "Reply to sender only or reply all?"
3. Use prepare_email_reply with reply_mode "sender" for sender-only or "all" for reply-all.
4. Replies are sent from ramy.mina@minaco.ca and remain in the original Outlook conversation thread.
5. Read back whether it is sender-only or reply-all, the recipients, subject, and complete reply text. Clearly say it has NOT been sent yet.
6. Ask Ramy to confirm by saying "Send it."

Use send_confirmed_email ONLY after Ramy explicitly confirms the currently pending NEW email or REPLY with "Send it" or an unmistakable equivalent in direct response to your confirmation request.
Never claim an email or reply was sent unless the send tool confirms success.
Never invent a recipient email address. If Ramy gives only a person's name for a new email and you do not have a verified email address, ask for the email address.


ADVANCED EMAIL, ATTACHMENTS, AND CONTACTS

Use search_email when Ramy asks you to find an older email, search by person/company/topic, or search a specific period. Search can target Ramy's mailbox or Accounting.
Use check_accounting when Ramy asks about current invoices, statements, payment reminders, deposits, supplier credits, overdue notices, insurance, taxes, financing charges, or other accounting-mailbox items.
Use read_accounting_email to retrieve the complete Accounting email body before translating, analyzing, or making conclusions from it.

When an email has attachments, use list_email_attachments to identify the exact attachment. Use analyze_email_attachment when Ramy asks to open, summarize, translate, analyze, extract amounts/deadlines, or answer questions about an attachment. The attachment tool analyzes the actual file; never pretend you opened an attachment when you only saw its filename.

Use resolve_person when Ramy gives a person's name but not an email address. It searches verified identities and Minaco email history. If one clear candidate is returned, use that verified address. If multiple plausible candidates are returned, ask Ramy which one he means. Never invent an address.

MASTER ACTION & FOLLOW-UP REGISTER

Use create_action when Ramy explicitly asks you to track, remember, follow up, add an action, or when he clearly instructs you that a business outcome must be monitored. Do not create action items merely because an email contains information.
Use list_actions when Ramy asks what is overdue, what he is waiting for, what needs his decision, what London must follow up on, or for a project/action status.
Use update_action when Ramy tells you an action changed, someone replied, a follow-up occurred, a deadline changed, an item is completed/cancelled, or the next action changes.
The register fields are outcome, project, owner, dates, status, priority, next action, waiting on, Ramy requirement, risk, source, and notes. Preserve the distinction between promised date, hard deadline, and next follow-up.

DAILY EXECUTIVE BRIEF

Use daily_executive_brief when Ramy asks for his morning brief, daily brief, what needs his attention today, or an executive overview. It combines live calendar, executive email, Accounting email, and the action register. Never invent missing items.

LIVE CALENDAR RULES

Use check_calendar whenever Ramy asks what is on his calendar, what meetings he has, or asks about a specific date or period.
Use check_availability whenever Ramy asks whether he is free at one specific time.
Use find_availability whenever Ramy asks for several possible times, asks for his availability over a day/week/date range, or asks you to reply to someone with his availability. Do NOT ask Ramy to tell you his availability when his live calendar can answer it.
If no meeting duration is stated for an availability request, default to 30 minutes. Unless Ramy or the email specifies otherwise, search normal business hours from 9:00 AM to 5:00 PM Montreal time and offer 3 to 5 useful slots. Respect any duration, day, or time constraints stated in the email or by Ramy.
All spoken calendar times are Montreal local time unless Ramy explicitly specifies another time zone.
If Ramy asks whether you have calendar access, answer directly that you have live access to his Minaco calendar through the authorized calendar tools. Do not call a tool merely to explain that access.

CALENDAR-AWARE EMAIL TASKS

If Ramy asks you to respond to someone and provide his availability, complete the workflow yourself using live systems:
1. Identify the exact email. If he names the sender, use that name to identify/search the live email rather than asking Ramy for the email address if the message can be found.
2. Read the full email when its contents or requested meeting constraints matter.
3. Use find_availability for the requested period and derive suitable open slots from Ramy's live calendar.
4. Prepare the reply using those actual available times.
5. Read back the reply and wait for Ramy's explicit "Send it" confirmation before sending.
If Ramy says "respond to Francis", "reply to Joannie", or otherwise names the person he wants to answer, treat that as reply-to-sender unless he explicitly asks for reply-all. If he only says "reply to this email" or "respond to this message" and the intended scope is unclear, ask sender-only or reply-all.

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
When Ramy gives you a task that can be resolved from connected systems, use the tools yourself instead of asking him for information those systems can provide. Ask Ramy only for a genuinely missing required fact that cannot be retrieved or safely inferred.
For a multi-step task, acknowledge it in one short natural phrase such as "I'll check that and prepare it," then proceed. Do not leave Ramy wondering whether you understood the task.

AUTHORITY

You may retrieve information, summarize, translate, analyze attachments, organize and maintain the Master Action Register, prepare actions, send SMS when explicitly requested, send a confirmed email, and execute a confirmed calendar action through the authorized tools.
Do not claim to have approved payments, signed contracts, committed Minaco to pricing, settled disputes, or made legal or financial decisions unless an authorized system actually performed that action.

PHONE STYLE

You are speaking with Ramy by phone.
Speak naturally at a calm executive-assistant pace, roughly 10 to 15 percent slower than a fast conversational assistant. Use short sentences, clear punctuation, and small natural pauses between ideas. Slow down slightly when reading names, dates, dollar amounts, email addresses, meeting times, and action items.
Do not rush to answer while Ramy is still forming a sentence. Allow normal hesitations and short pauses. Listen for the full thought before responding.
If Ramy begins speaking while you are talking, stop and listen rather than talking over him.
Keep most answers short unless Ramy asks for detail.
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
  reply.send({
    message: 'London Assistant Twilio Media Stream Server is running!',
    features: [
      'voice',
      'full email',
      'email search',
      'attachments',
      'contact resolution',
      'calendar',
      'action register',
      'accounting mailbox',
      'daily executive brief',
    ],
  });
});

// Optional scheduled endpoint. Power Automate can POST here each morning.
// Set DAILY_BRIEF_SECRET in Render and pass it as x-london-brief-secret.
fastify.post('/daily-brief', async (request, reply) => {
  try {
    if (!DAILY_BRIEF_SECRET) {
      return reply.code(503).send({
        success: false,
        error: 'DAILY_BRIEF_SECRET is not configured, so scheduled briefs are disabled.',
      });
    }

    const suppliedSecret = request.headers['x-london-brief-secret'];
    if (!suppliedSecret || suppliedSecret !== DAILY_BRIEF_SECRET) {
      return reply.code(401).send({ success: false, error: 'Unauthorized.' });
    }

    const brief = await generateDailyExecutiveBrief();
    await sendEmailFromLondon({
      to: RAMY_MINACO_EMAIL,
      subject: brief.subject,
      body: brief.html,
      contentType: 'HTML',
    });

    return reply.send({ success: true, sentTo: RAMY_MINACO_EMAIL });
  } catch (error) {
    console.error('Scheduled daily brief error:', error);
    return reply.code(500).send({ success: false, error: error.message });
  }
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
      let pendingEmailReply = null;
      let pendingEmailActionType = null;
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
                turn_detection: {
                  type: 'semantic_vad',
                  eagerness: 'high',
                  create_response: true,
                  interrupt_response: true,
                },
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
                  'List Ramy Mina’s latest live Minaco inbox emails with sender, subject, preview, time, and message id. Use for current inbox questions or to identify an email before reading or replying. This is not a full historical mailbox search.',
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
                name: 'read_email',
                description:
                  'Retrieve the complete live body and recipient details of one specific email in Ramy’s Minaco mailbox. Use whenever Ramy asks to read, translate, summarize, analyze, or respond based on the full email rather than a preview.',
                parameters: {
                  type: 'object',
                  properties: {
                    message_id: {
                      type: 'string',
                      description:
                        'Exact Microsoft Graph message id returned by check_email.',
                    },
                  },
                  required: ['message_id'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'prepare_email_reply',
                description:
                  'Prepare a reply to a specific live email in Ramy’s Minaco mailbox. This NEVER sends. reply_mode must be sender for sender-only or all for reply-all. Use only after the exact message id is known.',
                parameters: {
                  type: 'object',
                  properties: {
                    message_id: {
                      type: 'string',
                      description:
                        'Exact Microsoft Graph message id returned by check_email or read_email.',
                    },
                    reply_mode: {
                      type: 'string',
                      enum: ['sender', 'all'],
                      description:
                        'sender = reply to sender/reply-to only. all = reply all to the original conversation recipients.',
                    },
                    body: {
                      type: 'string',
                      description: 'Complete reply body to send after confirmation.',
                    },
                  },
                  required: ['message_id', 'reply_mode', 'body'],
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
                  'Send the currently pending prepared email action, either a new London email or a reply from Ramy’s mailbox. Use ONLY after Ramy explicitly confirms by saying Send it or an unmistakable equivalent.',
                parameters: {
                  type: 'object',
                  properties: {},
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'search_email',
                description:
                  'Search historical live email by person, company, subject, or keywords. Can search Ramy or Accounting. Use when the desired email may not be in the latest inbox list.',
                parameters: {
                  type: 'object',
                  properties: {
                    mailbox: {
                      type: 'string',
                      enum: ['ramy', 'accounting'],
                      description: 'Mailbox to search. Default ramy.',
                    },
                    query: {
                      type: 'string',
                      description: 'Person, company, subject, or keywords to search.',
                    },
                    start_date: {
                      type: 'string',
                      description: 'Optional start date YYYY-MM-DD.',
                    },
                    end_date: {
                      type: 'string',
                      description: 'Optional end date YYYY-MM-DD.',
                    },
                    limit: {
                      type: 'integer',
                      minimum: 1,
                      maximum: 25,
                      description: 'Maximum results. Default 10.',
                    },
                  },
                  required: ['query'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'check_accounting',
                description:
                  'List current live emails in Minaco Accounting. Use for invoices, statements, payment reminders, deposits, supplier credits, overdue notices, insurance, taxes, financing charges, or other finance inbox questions.',
                parameters: {
                  type: 'object',
                  properties: {
                    limit: {
                      type: 'integer',
                      minimum: 1,
                      maximum: 20,
                      description: 'Number of recent Accounting emails. Default 10.',
                    },
                  },
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'read_accounting_email',
                description:
                  'Retrieve the complete body and recipient details of a specific Minaco Accounting email.',
                parameters: {
                  type: 'object',
                  properties: {
                    message_id: {
                      type: 'string',
                      description: 'Exact message id returned by check_accounting or search_email.',
                    },
                  },
                  required: ['message_id'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'list_email_attachments',
                description:
                  'List the actual attachments on a specific live email. Use before opening or analyzing an attachment.',
                parameters: {
                  type: 'object',
                  properties: {
                    mailbox: {
                      type: 'string',
                      enum: ['ramy', 'accounting'],
                      description: 'Mailbox containing the email.',
                    },
                    message_id: {
                      type: 'string',
                      description: 'Exact email message id.',
                    },
                  },
                  required: ['message_id'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'analyze_email_attachment',
                description:
                  'Open and analyze a real email file attachment. Use for PDF, Word, Excel, PowerPoint, text/code files, and images. Can summarize, translate, extract facts/amounts/deadlines, or answer questions about the attachment.',
                parameters: {
                  type: 'object',
                  properties: {
                    mailbox: {
                      type: 'string',
                      enum: ['ramy', 'accounting'],
                      description: 'Mailbox containing the email.',
                    },
                    message_id: {
                      type: 'string',
                      description: 'Exact email message id.',
                    },
                    attachment_id: {
                      type: 'string',
                      description: 'Exact attachment id returned by list_email_attachments.',
                    },
                    instruction: {
                      type: 'string',
                      description: 'What Ramy wants done with the attachment, e.g. summarize it, translate it into French, or extract all deadlines and amounts.',
                    },
                  },
                  required: ['message_id', 'attachment_id', 'instruction'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'resolve_person',
                description:
                  'Resolve a person name to verified email candidates using Minaco identities and Ramy’s live email history. Use before sending email or inviting someone when Ramy gives only a name.',
                parameters: {
                  type: 'object',
                  properties: {
                    query: {
                      type: 'string',
                      description: 'Person name or partial email address.',
                    },
                  },
                  required: ['query'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'create_action',
                description:
                  'Create a persistent item in London’s Master Action & Follow-Up Register. Use when Ramy explicitly asks to track, remember, follow up, or monitor an outcome.',
                parameters: {
                  type: 'object',
                  properties: {
                    title: { type: 'string', description: 'Action or outcome to track.' },
                    project: { type: 'string', description: 'Project/property if applicable.' },
                    category: { type: 'string', description: 'Category such as project, tenant, finance, legal, vendor, or admin.' },
                    owner: { type: 'string', description: 'Person responsible. Default London.' },
                    promised_date: { type: 'string', description: 'Promised date YYYY-MM-DD if someone committed to a date.' },
                    hard_deadline: { type: 'string', description: 'True deadline YYYY-MM-DD if one exists.' },
                    next_follow_up: { type: 'string', description: 'Next follow-up date YYYY-MM-DD.' },
                    status: { type: 'string', description: 'NEW, ACTIVE, WAITING - EXTERNAL, WAITING - INTERNAL, WAITING - RAMY, BLOCKED, OVERDUE, COMPLETED, or CANCELLED.' },
                    priority: { type: 'string', description: 'CRITICAL, HIGH, NORMAL, or LOW.' },
                    next_action: { type: 'string', description: 'Concrete next action.' },
                    waiting_on: { type: 'string', description: 'Who or what the action is waiting on.' },
                    ramy_required: { type: 'boolean', description: 'Whether Ramy personally must act or decide.' },
                    ramy_decision_by: { type: 'string', description: 'Date Ramy must decide by, YYYY-MM-DD.' },
                    risk_if_delayed: { type: 'string', description: 'Consequence if delayed.' },
                    source: { type: 'string', description: 'Source, e.g. voice, email subject, meeting.' },
                    notes: { type: 'string', description: 'Concise history or context.' },
                  },
                  required: ['title'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'list_actions',
                description:
                  'Read London’s persistent Master Action & Follow-Up Register. Use for overdue items, waiting items, Ramy decisions, project follow-ups, or action status.',
                parameters: {
                  type: 'object',
                  properties: {
                    status: { type: 'string', description: 'Optional status filter.' },
                    priority: { type: 'string', description: 'Optional priority filter.' },
                    project: { type: 'string', description: 'Optional project/property filter.' },
                    owner: { type: 'string', description: 'Optional owner filter.' },
                    only_overdue: { type: 'boolean', description: 'Return only overdue items.' },
                    include_closed: { type: 'boolean', description: 'Include completed/cancelled items. Default false.' },
                    limit: { type: 'integer', minimum: 1, maximum: 100, description: 'Maximum items. Default 30.' },
                  },
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'update_action',
                description:
                  'Update an existing Master Action Register item after a reply, follow-up, decision, deadline change, completion, cancellation, or change in next action.',
                parameters: {
                  type: 'object',
                  properties: {
                    event_id: { type: 'string', description: 'Exact action event id returned by list_actions.' },
                    title: { type: 'string' },
                    project: { type: 'string' },
                    category: { type: 'string' },
                    owner: { type: 'string' },
                    promised_date: { type: 'string' },
                    hard_deadline: { type: 'string' },
                    next_follow_up: { type: 'string' },
                    status: { type: 'string' },
                    priority: { type: 'string' },
                    last_contact: { type: 'string' },
                    next_action: { type: 'string' },
                    waiting_on: { type: 'string' },
                    ramy_required: { type: 'boolean' },
                    ramy_decision_by: { type: 'string' },
                    risk_if_delayed: { type: 'string' },
                    source: { type: 'string' },
                    notes: { type: 'string' },
                  },
                  required: ['event_id'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'daily_executive_brief',
                description:
                  'Generate Ramy’s live Minaco executive brief using calendar, Ramy inbox, Accounting inbox, and the Master Action Register. Use when he asks what needs attention today, for his morning brief, or daily executive overview.',
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
                name: 'find_availability',
                description:
                  'Find several real open time slots from Ramy’s live Minaco calendar across a Montreal-local date range. Use for questions such as what times am I available next week, give someone my availability, or find a few meeting options. Do not ask Ramy to state availability when this tool can calculate it.',
                parameters: {
                  type: 'object',
                  properties: {
                    start_date: {
                      type: 'string',
                      description: 'First Montreal local date in YYYY-MM-DD.',
                    },
                    end_date: {
                      type: 'string',
                      description: 'Last Montreal local date in YYYY-MM-DD, inclusive.',
                    },
                    duration_minutes: {
                      type: 'integer',
                      minimum: 15,
                      maximum: 240,
                      description: 'Required meeting duration. Default 30 minutes if unspecified.',
                    },
                    day_start: {
                      type: 'string',
                      description: 'Earliest acceptable local time in HH:mm. Default 09:00.',
                    },
                    day_end: {
                      type: 'string',
                      description: 'Latest acceptable local ending boundary in HH:mm. Default 17:00.',
                    },
                    max_slots: {
                      type: 'integer',
                      minimum: 1,
                      maximum: 12,
                      description: 'Maximum open slots to return. Default 5.',
                    },
                    include_weekends: {
                      type: 'boolean',
                      description: 'Whether Saturday and Sunday may be offered. Default false.',
                    },
                  },
                  required: ['start_date', 'end_date'],
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
              if (pendingEmailActionType === 'reply') {
                if (!pendingEmailReply) {
                  throw new Error('There is no pending email reply to send.');
                }

                const replyToSend = { ...pendingEmailReply };
                const result = await replyToMinacoEmail(replyToSend);
                pendingEmailReply = null;
                pendingEmailDraft = null;
                pendingEmailActionType = null;

                respondToToolCall(
                  response.call_id,
                  {
                    success: true,
                    sent: true,
                    action: 'reply',
                    from: result.from,
                    replyMode: result.mode,
                    subject: replyToSend.subject,
                    recipients: replyToSend.recipients,
                  },
                  'Confirm briefly that the reply was successfully sent. Say whether it was reply-to-sender or reply-all, mention the subject, and do not invent anything.'
                );
              } else if (pendingEmailActionType === 'new') {
                if (!pendingEmailDraft) {
                  throw new Error('There is no pending new email draft to send.');
                }

                const emailToSend = { ...pendingEmailDraft };
                const result = await sendEmailFromLondon(emailToSend);
                pendingEmailDraft = null;
                pendingEmailReply = null;
                pendingEmailActionType = null;

                respondToToolCall(
                  response.call_id,
                  {
                    success: true,
                    sent: true,
                    action: 'new_email',
                    from: result.from,
                    to: result.to,
                    subject: result.subject,
                  },
                  'Confirm briefly that the new email was successfully sent. Mention the recipient and subject. Do not invent anything.'
                );
              } else {
                throw new Error('There is no pending email action to send.');
              }
            } catch (error) {
              console.error('Confirmed email send error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, sent: false, error: error.message },
                'Tell Ramy the email action was NOT sent and state the returned error concisely. Do not invent anything.'
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
              pendingEmailReply = null;
              pendingEmailActionType = 'new';

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
              const simplifiedEmails = emails.map((email, index) => ({
                number: index + 1,
                messageId: email.id,
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
            response.name === 'read_email'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const email = await getFullMinacoEmail(args.message_id);

              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  messageId: email.id,
                  conversationId: email.conversationId || '',
                  subject: email.subject || '(No subject)',
                  from: {
                    name: email.from?.emailAddress?.name || '',
                    address: email.from?.emailAddress?.address || '',
                  },
                  replyTo: simplifyEmailRecipients(email.replyTo),
                  toRecipients: simplifyEmailRecipients(email.toRecipients),
                  ccRecipients: simplifyEmailRecipients(email.ccRecipients),
                  receivedTimeMontreal: email.receivedDateTime
                    ? new Intl.DateTimeFormat('en-US', {
                        timeZone: MONTREAL_IANA_TIME_ZONE,
                        dateStyle: 'medium',
                        timeStyle: 'short',
                      }).format(new Date(email.receivedDateTime))
                    : '',
                  isRead: Boolean(email.isRead),
                  bodyContentType: email.body?.contentType || 'text',
                  fullBody: email.body?.content || '',
                },
                'Use the complete live email body returned by the tool. If Ramy asked to translate it, translate the full body, not the preview. If he asked to summarize or analyze it, use the full body. Do not invent omitted attachments or content.'
              );
            } catch (error) {
              console.error('Full email lookup error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the full email could not be retrieved and state the returned error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'prepare_email_reply'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const mode = args.reply_mode;

              if (!args.message_id || !args.body) {
                throw new Error('Message id and reply body are required.');
              }

              if (!['sender', 'all'].includes(mode)) {
                throw new Error('Reply mode must be sender or all.');
              }

              const email = await getFullMinacoEmail(args.message_id);
              const recipients = getReplyRecipientSummary(email, mode);

              if (recipients.length === 0) {
                throw new Error('No valid reply recipient could be determined.');
              }

              pendingEmailReply = {
                messageId: email.id,
                mode,
                body: args.body,
                subject: email.subject || '(No subject)',
                recipients,
              };
              pendingEmailDraft = null;
              pendingEmailActionType = 'reply';

              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  sent: false,
                  from: RAMY_MINACO_EMAIL,
                  replyMode: mode,
                  subject: pendingEmailReply.subject,
                  recipients,
                  body: args.body,
                },
                'Read back whether this is reply-to-sender or reply-all, the recipient names or addresses, the subject, and the complete reply message. Clearly say the reply has NOT been sent. Ask Ramy to confirm by saying Send it.'
              );
            } catch (error) {
              console.error('Prepare email reply error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the reply could not be prepared and state the returned error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'search_email'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const mailboxKey = args.mailbox || 'ramy';
              const mailboxAddress = mailboxAddressFromKey(mailboxKey);
              const emails = await searchMailboxEmails({
                mailboxAddress,
                query: args.query,
                limit: args.limit || 10,
                startDate: args.start_date,
                endDate: args.end_date,
              });

              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  mailbox: mailboxAddress,
                  results: emails.map((email, index) => ({
                    number: index + 1,
                    messageId: email.id,
                    from:
                      email.from?.emailAddress?.name ||
                      email.from?.emailAddress?.address ||
                      'Unknown sender',
                    fromEmail: email.from?.emailAddress?.address || '',
                    subject: email.subject || '(No subject)',
                    receivedTimeMontreal: email.receivedDateTime
                      ? new Intl.DateTimeFormat('en-US', {
                          timeZone: MONTREAL_IANA_TIME_ZONE,
                          dateStyle: 'medium',
                          timeStyle: 'short',
                        }).format(new Date(email.receivedDateTime))
                      : '',
                    isRead: Boolean(email.isRead),
                    hasAttachments: Boolean(email.hasAttachments),
                    preview: email.bodyPreview || '',
                  })),
                },
                'Use only these live search results. Identify the best matching email(s) by sender, subject, date, and preview. If multiple results are plausible, ask Ramy which one he means instead of guessing.'
              );
            } catch (error) {
              console.error('Historical email search error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the email search failed and state the returned error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'check_accounting'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const limit = Math.min(Math.max(Number(args.limit) || 10, 1), 20);
              const emails = await getRecentMailboxEmails(ACCOUNTING_MAILBOX, limit);
              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  mailbox: ACCOUNTING_MAILBOX,
                  emails: emails.map((email, index) => ({
                    number: index + 1,
                    messageId: email.id,
                    from:
                      email.from?.emailAddress?.name ||
                      email.from?.emailAddress?.address ||
                      'Unknown sender',
                    fromEmail: email.from?.emailAddress?.address || '',
                    subject: email.subject || '(No subject)',
                    receivedTimeMontreal: email.receivedDateTime
                      ? new Intl.DateTimeFormat('en-US', {
                          timeZone: MONTREAL_IANA_TIME_ZONE,
                          dateStyle: 'medium',
                          timeStyle: 'short',
                        }).format(new Date(email.receivedDateTime))
                      : '',
                    isRead: Boolean(email.isRead),
                    hasAttachments: Boolean(email.hasAttachments),
                    preview: email.bodyPreview || '',
                  })),
                },
                'Summarize only the live Accounting inbox results. Prioritize invoices, statements, overdue notices, payment deadlines, deposits, credits, insurance, taxes, financing charges, unusual amounts, or anything requiring Ramy’s attention. Do not authorize payment.'
              );
            } catch (error) {
              console.error('Accounting mailbox lookup error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the Accounting mailbox lookup failed and state the returned error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'read_accounting_email'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const email = await getFullMailboxEmail(
                ACCOUNTING_MAILBOX,
                args.message_id
              );
              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  mailbox: ACCOUNTING_MAILBOX,
                  messageId: email.id,
                  subject: email.subject || '(No subject)',
                  from: {
                    name: email.from?.emailAddress?.name || '',
                    address: email.from?.emailAddress?.address || '',
                  },
                  toRecipients: simplifyEmailRecipients(email.toRecipients),
                  ccRecipients: simplifyEmailRecipients(email.ccRecipients),
                  receivedTimeMontreal: email.receivedDateTime
                    ? new Intl.DateTimeFormat('en-US', {
                        timeZone: MONTREAL_IANA_TIME_ZONE,
                        dateStyle: 'medium',
                        timeStyle: 'short',
                      }).format(new Date(email.receivedDateTime))
                    : '',
                  hasAttachments: Boolean(email.hasAttachments),
                  fullBody: email.body?.content || '',
                },
                'Use the complete live Accounting email body. Preserve amounts, dates, invoice numbers, deadlines, and material conditions. Do not infer that a payment is approved.'
              );
            } catch (error) {
              console.error('Accounting full email lookup error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the Accounting email could not be retrieved and state the error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'list_email_attachments'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const mailboxAddress = mailboxAddressFromKey(args.mailbox || 'ramy');
              const attachments = await listMailboxEmailAttachments(
                mailboxAddress,
                args.message_id
              );
              respondToToolCall(
                response.call_id,
                { success: true, mailbox: mailboxAddress, attachments },
                'Tell Ramy which real attachments are available, with filename and approximate size. Ignore inline signature images unless he asks for them. If there are no attachments, say so.'
              );
            } catch (error) {
              console.error('Attachment list error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the attachment list could not be retrieved and state the error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'analyze_email_attachment'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const mailboxAddress = mailboxAddressFromKey(args.mailbox || 'ramy');
              const analysis = await analyzeMailboxAttachment({
                mailboxAddress,
                messageId: args.message_id,
                attachmentId: args.attachment_id,
                instruction: args.instruction,
              });
              respondToToolCall(
                response.call_id,
                { success: true, mailbox: mailboxAddress, ...analysis },
                'Answer Ramy using the actual attachment analysis returned by the tool. If he requested translation, provide that translation. If the result is long, start with the requested answer and offer to continue reading rather than inventing or omitting material facts.'
              );
            } catch (error) {
              console.error('Attachment analysis error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the attachment could not be analyzed and state the returned error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'resolve_person'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const candidates = await resolvePersonFromMailHistory(args.query);
              respondToToolCall(
                response.call_id,
                { success: true, query: args.query, candidates },
                'Use these verified email candidates only. If there is one clearly matching candidate, state the name and address and continue the requested workflow. If there are multiple plausible candidates or no candidate, ask one short clarification. Never invent an email address.'
              );
            } catch (error) {
              console.error('Person resolution error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the contact could not be resolved from verified Minaco data and ask for the email address if needed.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'create_action'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const action = await createActionItem({
                title: args.title,
                project: args.project,
                category: args.category,
                owner: args.owner,
                promisedDate: args.promised_date,
                hardDeadline: args.hard_deadline,
                nextFollowUp: args.next_follow_up,
                status: args.status,
                priority: args.priority,
                nextAction: args.next_action,
                waitingOn: args.waiting_on,
                ramyRequired: args.ramy_required,
                ramyDecisionBy: args.ramy_decision_by,
                riskIfDelayed: args.risk_if_delayed,
                source: args.source,
                notes: args.notes,
              });
              respondToToolCall(
                response.call_id,
                { success: true, action },
                'Confirm briefly that London is now tracking the action. Mention the outcome, owner, next follow-up or deadline if present, and whether Ramy is required. Do not over-explain.'
              );
            } catch (error) {
              console.error('Create action error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the action could not be added to the register and state the error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'list_actions'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const actions = await listActionItems({
                status: args.status,
                priority: args.priority,
                project: args.project,
                owner: args.owner,
                onlyOverdue: Boolean(args.only_overdue),
                includeClosed: Boolean(args.include_closed),
                limit: Math.min(Math.max(Number(args.limit) || 30, 1), 100),
              });
              respondToToolCall(
                response.call_id,
                { success: true, actions },
                'Summarize the live action register based on Ramy’s question. Prioritize overdue, critical/high, WAITING - RAMY, hard deadlines, and next follow-ups. Mention eventId only if a follow-up tool action needs it; do not read technical ids aloud unnecessarily.'
              );
            } catch (error) {
              console.error('List actions error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the action register could not be read and state the error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'update_action'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const changes = {};
              const mapping = {
                title: 'title',
                project: 'project',
                category: 'category',
                owner: 'owner',
                promised_date: 'promisedDate',
                hard_deadline: 'hardDeadline',
                next_follow_up: 'nextFollowUp',
                status: 'status',
                priority: 'priority',
                last_contact: 'lastContact',
                next_action: 'nextAction',
                waiting_on: 'waitingOn',
                ramy_required: 'ramyRequired',
                ramy_decision_by: 'ramyDecisionBy',
                risk_if_delayed: 'riskIfDelayed',
                source: 'source',
                notes: 'notes',
              };
              for (const [inputKey, targetKey] of Object.entries(mapping)) {
                if (Object.prototype.hasOwnProperty.call(args, inputKey)) {
                  changes[targetKey] = args[inputKey];
                }
              }
              const action = await updateActionItem(args.event_id, changes);
              respondToToolCall(
                response.call_id,
                { success: true, action },
                'Confirm the action register was updated. Mention the new status, next action, follow-up/deadline, or waiting party that changed. Be brief.'
              );
            } catch (error) {
              console.error('Update action error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the action could not be updated and state the error concisely.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'daily_executive_brief'
          ) {
            try {
              const brief = await generateDailyExecutiveBrief();
              respondToToolCall(
                response.call_id,
                { success: true, brief: brief.voiceText },
                'Read the executive brief naturally and concisely. Lead with anything requiring Ramy’s immediate decision or carrying deadline/financial/legal/project risk. Do not add facts not in the brief.'
              );
            } catch (error) {
              console.error('Daily brief error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the daily executive brief could not be generated and state the error concisely.'
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
            response.name === 'find_availability'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const result = await findCalendarAvailability({
                startDate: args.start_date,
                endDate: args.end_date,
                durationMinutes: args.duration_minutes ?? 30,
                dayStart: args.day_start || '09:00',
                dayEnd: args.day_end || '17:00',
                maxSlots: args.max_slots ?? 5,
                includeWeekends: Boolean(args.include_weekends),
              });

              respondToToolCall(
                response.call_id,
                { success: true, timezone: 'Montreal local time', ...result },
                'Use only these live calendar results. Give Ramy the best 3 to 5 available slots unless he asked for a different number. If this tool is being used to prepare an email reply, continue the reply workflow using these exact available times instead of asking Ramy what his availability is. If there are no slots, say so and ask whether to widen the hours or dates.'
              );
            } catch (error) {
              console.error('Availability range search error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the live calendar availability search failed and state the returned error concisely. Do not freeze or pretend availability was found.'
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
