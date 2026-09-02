import Fastify from 'fastify';
import WebSocket from 'ws';
import dotenv from 'dotenv';
import fastifyFormBody from '@fastify/formbody';
import fastifyWs from '@fastify/websocket';
import { randomUUID, createHmac, timingSafeEqual } from 'node:crypto';
import { readFileSync, writeFileSync, renameSync, mkdirSync } from 'node:fs';
import { dirname } from 'node:path';
import { resolveFileContentType } from './file-content-type.js';

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
  TASK_INBOX_SECRET,
  TASK_INBOX_MODEL,
  OPENAI_DOCUMENT_MODEL,
  EXECUTIVE_BRIEF_MODEL,
  DROPBOX_ACCESS_TOKEN,
  DROPBOX_ROOT_PATH,
  HEALTH_SECRET,
  LONDON_STATE_FILE,
} = process.env;

const MONTREAL_IANA_TIME_ZONE = 'America/Toronto';
const MICROSOFT_EASTERN_TIME_ZONE = 'Eastern Standard Time';
const ACCOUNTING_MAILBOX = ACCOUNTING_MINACO_EMAIL || 'accounting@minaco.ca';
const ACTION_REGISTER_CALENDAR_NAME = 'London Action Register';
const DOCUMENT_ANALYSIS_MODEL = OPENAI_DOCUMENT_MODEL || 'gpt-5.6';
const DAILY_BRIEF_MODEL = EXECUTIVE_BRIEF_MODEL || 'gpt-5.6-luna';
const TASK_INBOX_ANALYSIS_MODEL = TASK_INBOX_MODEL || DOCUMENT_ANALYSIS_MODEL;
const ACTION_MESSAGE_MODEL = process.env.ACTION_MESSAGE_MODEL || 'gpt-5.6-luna';
const MAX_MESSAGING_REPLY_CHARS = 1450;
const MAX_ATTACHMENT_BYTES = 45 * 1024 * 1024;
const MAX_TASK_ATTACHMENTS = 10;
const MAX_TASK_TOTAL_BYTES = 45 * 1024 * 1024;
const LONDON_DROPBOX_ROOT = String(DROPBOX_ROOT_PATH || '/LONDON - ACCESS').trim() || '/LONDON - ACCESS';
const STATE_FILE = String(LONDON_STATE_FILE || '/tmp/london-state.json').trim();
const MAX_DROPBOX_TEXT_BYTES = 8 * 1024 * 1024;
const MAX_DROPBOX_ANALYSIS_BYTES = 45 * 1024 * 1024;

// -----------------------------------------------------------------------------
// Network resilience
// -----------------------------------------------------------------------------

const DEFAULT_NETWORK_TIMEOUT_MS = 10000;

const fetchWithTimeout = async (url, options = {}, timeoutMs = DEFAULT_NETWORK_TIMEOUT_MS) => {
  const signal =
    options.signal ||
    (typeof AbortSignal !== 'undefined' && typeof AbortSignal.timeout === 'function'
      ? AbortSignal.timeout(timeoutMs)
      : undefined);

  return fetch(url, {
    ...options,
    ...(signal ? { signal } : {}),
  });
};

let graphReadTokenCache = { token: '', expiresAt: 0 };
let graphActionsTokenCache = { token: '', expiresAt: 0 };

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

  const response = await fetchWithTimeout(
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


const normalizePhoneIdentity = (value) => {
  const raw = String(value || '').trim().toLowerCase();
  const withoutChannel = raw.startsWith('whatsapp:') ? raw.slice('whatsapp:'.length) : raw;
  const digits = withoutChannel.replace(/[^0-9+]/g, '');
  if (!digits) return '';
  if (digits.startsWith('+')) return `+${digits.slice(1).replace(/\D/g, '')}`;
  return `+${digits.replace(/\D/g, '')}`;
};

const isAuthorizedRamyMessagingSender = (value) => {
  const sender = normalizePhoneIdentity(value);
  const ramy = normalizePhoneIdentity(RAMY_PHONE_NUMBER);
  return Boolean(sender && ramy && sender === ramy);
};

const twilioPublicWebhookUrl = (request) => {
  const forwardedProto = String(request.headers['x-forwarded-proto'] || '').split(',')[0].trim();
  const forwardedHost = String(request.headers['x-forwarded-host'] || '').split(',')[0].trim();
  const protocol = forwardedProto || 'https';
  const host = forwardedHost || request.headers.host;
  return `${protocol}://${host}${request.raw.url}`;
};

const validateTwilioFormWebhook = (request) => {
  if (!TWILIO_AUTH_TOKEN) return false;
  const signature = String(request.headers['x-twilio-signature'] || '').trim();
  if (!signature) return false;

  const params = request.body && typeof request.body === 'object' ? request.body : {};
  let payload = twilioPublicWebhookUrl(request);
  for (const key of Object.keys(params).sort()) {
    const value = params[key];
    if (Array.isArray(value)) {
      for (const item of [...value].map(String).sort()) payload += `${key}${item}`;
    } else if (value != null) {
      payload += `${key}${String(value)}`;
    }
  }

  const expected = createHmac('sha1', TWILIO_AUTH_TOKEN)
    .update(payload, 'utf8')
    .digest('base64');

  const a = Buffer.from(expected);
  const b = Buffer.from(signature);
  return a.length === b.length && timingSafeEqual(a, b);
};

const sendTwilioChannelMessage = async ({ to, from, body }) => {
  if (!TWILIO_ACCOUNT_SID || !TWILIO_AUTH_TOKEN) {
    throw new Error('Missing Twilio messaging credentials.');
  }
  if (!to || !from) throw new Error('Twilio To and From addresses are required.');

  const auth = Buffer.from(
    `${TWILIO_ACCOUNT_SID}:${TWILIO_AUTH_TOKEN}`
  ).toString('base64');

  const cleanBody = String(body || '').trim().slice(0, MAX_MESSAGING_REPLY_CHARS);
  const form = new URLSearchParams({
    To: String(to),
    From: String(from),
    Body: cleanBody || 'Updated.',
  });

  const response = await fetchWithTimeout(
    `https://api.twilio.com/2010-04-01/Accounts/${TWILIO_ACCOUNT_SID}/Messages.json`,
    {
      method: 'POST',
      headers: {
        Authorization: `Basic ${auth}`,
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: form.toString(),
    },
    12000
  );

  const data = await response.json();
  if (!response.ok) {
    throw new Error(`Twilio messaging send failed: ${data.message || response.status}`);
  }
  return data;
};

const getMicrosoftGraphToken = async () => {
  if (!MS_TENANT_ID || !MS_CLIENT_ID || !MS_CLIENT_SECRET) {
    throw new Error('Missing Microsoft Graph environment variables.');
  }

  if (graphReadTokenCache.token && graphReadTokenCache.expiresAt > Date.now() + 60000) {
    return graphReadTokenCache.token;
  }

  const form = new URLSearchParams({
    client_id: MS_CLIENT_ID,
    client_secret: MS_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });

  const response = await fetchWithTimeout(
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

  graphReadTokenCache = {
    token: data.access_token,
    expiresAt: Date.now() + Math.max(Number(data.expires_in || 3600) - 120, 60) * 1000,
  };

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

  if (graphActionsTokenCache.token && graphActionsTokenCache.expiresAt > Date.now() + 60000) {
    return graphActionsTokenCache.token;
  }

  const form = new URLSearchParams({
    client_id: ACTIONS_MS_CLIENT_ID,
    client_secret: ACTIONS_MS_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });

  const response = await fetchWithTimeout(
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

  graphActionsTokenCache = {
    token: data.access_token,
    expiresAt: Date.now() + Math.max(Number(data.expires_in || 3600) - 120, 60) * 1000,
  };

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

  const response = await fetchWithTimeout(url, {
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

  const response = await fetchWithTimeout(url, {
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

  const response = await fetchWithTimeout(url, {
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

  const response = await fetchWithTimeout(url, {
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

  const response = await fetchWithTimeout(url, {
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

  results.sort(
    (a, b) =>
      new Date(b.receivedDateTime || 0).getTime() -
      new Date(a.receivedDateTime || 0).getTime()
  );

  return results.slice(0, requestedLimit);
};

const normalizePersonSearchText = (value) =>
  String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^\p{L}\p{N}@._+-]+/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();

const getRecentInboxEmails = async ({
  mailboxAddress,
  limit = 100,
} = {}) => {
  if (!mailboxAddress) throw new Error('Mailbox address is required.');

  const token = await getMicrosoftGraphToken();
  const requestedLimit = Math.min(Math.max(Number(limit) || 100, 1), 100);
  const url = new URL(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      mailboxAddress
    )}/mailFolders/inbox/messages`
  );

  url.searchParams.set('$top', String(requestedLimit));
  url.searchParams.set(
    '$select',
    'id,conversationId,subject,from,replyTo,toRecipients,ccRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments'
  );
  url.searchParams.set('$orderby', 'receivedDateTime desc');

  const response = await fetchWithTimeout(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph inbox lookup failed for ${mailboxAddress}: ${
        data.error?.message || response.status
      }`
    );
  }

  return data.value || [];
};

const scoreSenderMatch = (email, personQuery) => {
  const query = normalizePersonSearchText(personQuery);
  if (!query) return 0;

  const senderName = normalizePersonSearchText(
    email.from?.emailAddress?.name || ''
  );
  const senderAddress = normalizePersonSearchText(
    email.from?.emailAddress?.address || ''
  );
  const senderText = `${senderName} ${senderAddress}`.trim();

  if (!senderText) return 0;
  if (senderAddress === query) return 120;
  if (senderName === query) return 115;
  if (senderText.includes(query)) return 105;

  const tokens = query.split(' ').filter(Boolean);
  if (tokens.length && tokens.every((token) => senderText.includes(token))) {
    return 95;
  }
  if (tokens.length === 1 && senderName.split(' ').includes(tokens[0])) {
    return 90;
  }
  return 0;
};

const findLatestInboundEmailBySender = async ({
  mailboxAddress,
  personQuery,
} = {}) => {
  const query = String(personQuery || '').trim();
  if (!query) throw new Error('A sender name or email is required.');

  // First search the actual Inbox and compare only the From field.
  // This avoids selecting Ramy's own sent message merely because it mentions
  // the person's name in the subject/body.
  let candidates = await getRecentInboxEmails({
    mailboxAddress,
    limit: 100,
  });

  let scored = candidates
    .map((email) => ({
      email,
      score: scoreSenderMatch(email, query),
    }))
    .filter((entry) => entry.score > 0)
    .sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      return (
        new Date(b.email.receivedDateTime || 0).getTime() -
        new Date(a.email.receivedDateTime || 0).getTime()
      );
    });

  // Fallback to broader mailbox search only if no inbound sender match was
  // found in the recent Inbox. Still require the From field itself to match.
  if (scored.length === 0) {
    const searched = await searchMailboxEmails({
      mailboxAddress,
      query,
      limit: 25,
    });
    scored = searched
      .map((email) => ({
        email,
        score: scoreSenderMatch(email, query),
      }))
      .filter((entry) => entry.score > 0)
      .sort((a, b) => {
        if (b.score !== a.score) return b.score - a.score;
        return (
          new Date(b.email.receivedDateTime || 0).getTime() -
          new Date(a.email.receivedDateTime || 0).getTime()
        );
      });
  }

  if (scored.length === 0) {
    return {
      match: null,
      candidates: [],
      reason: `No inbound email from a sender matching ${query} was found.`,
    };
  }

  const bestScore = scored[0].score;
  const strongMatches = scored.filter((entry) => entry.score >= bestScore - 5);

  const uniqueByAddress = new Map();
  for (const entry of strongMatches) {
    const address = String(
      entry.email.from?.emailAddress?.address || ''
    ).toLowerCase();
    if (address && !uniqueByAddress.has(address)) {
      uniqueByAddress.set(address, entry.email);
    }
  }

  if (uniqueByAddress.size > 1) {
    return {
      match: null,
      reason: `More than one inbound sender matching ${query} was found.`,
      candidates: [...uniqueByAddress.values()].slice(0, 5),
    };
  }

  return {
    match: scored[0].email,
    candidates: [],
    reason: '',
  };
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

  const response = await fetchWithTimeout(url, {
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
  const response = await fetchWithTimeout(
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

const callOpenAIResponses = async (payload, timeoutMs = 45000) => {
  if (!OPENAI_API_KEY) throw new Error('OPENAI_API_KEY is not configured.');

  const response = await fetchWithTimeout('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${OPENAI_API_KEY}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload),
  }, timeoutMs);

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

  const response = await fetchWithTimeout(
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


// -----------------------------------------------------------------------------
// Delegated email-thread review jobs
// -----------------------------------------------------------------------------

const delegatedThreadReviewJobs = new Map();

const pruneDelegatedThreadReviewJobs = () => {
  if (delegatedThreadReviewJobs.size <= 50) return;
  const entries = [...delegatedThreadReviewJobs.entries()].sort(
    (a, b) => Number(a[1]?.createdAt || 0) - Number(b[1]?.createdAt || 0)
  );
  for (const [jobId] of entries.slice(0, delegatedThreadReviewJobs.size - 50)) {
    delegatedThreadReviewJobs.delete(jobId);
  }
};

const getMailboxConversationThread = async (
  mailboxAddress,
  conversationId,
  limit = 50
) => {
  if (!mailboxAddress || !conversationId) {
    throw new Error('Mailbox address and conversation id are required.');
  }

  const token = await getMicrosoftGraphToken();
  const escapedConversationId = String(conversationId).replace(/'/g, "''");
  const url = new URL(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
      mailboxAddress
    )}/messages`
  );

  url.searchParams.set('$filter', `conversationId eq '${escapedConversationId}'`);
  url.searchParams.set('$top', String(Math.min(Math.max(Number(limit) || 50, 1), 50)));
  url.searchParams.set(
    '$select',
    'id,conversationId,subject,from,replyTo,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments'
  );

  const response = await fetchWithTimeout(
    url,
    {
      headers: {
        Authorization: `Bearer ${token}`,
        Prefer: 'outlook.body-content-type="text"',
      },
    },
    15000
  );

  const data = await response.json();
  if (!response.ok) {
    throw new Error(
      `Microsoft Graph conversation lookup failed: ${
        data.error?.message || response.status
      }`
    );
  }

  return (data.value || []).sort(
    (a, b) =>
      new Date(a.receivedDateTime || 0).getTime() -
      new Date(b.receivedDateTime || 0).getTime()
  );
};

const compactThreadForAnalysis = (messages = []) => {
  const maxTotalChars = 140000;
  const maxBodyCharsPerMessage = 24000;
  let totalChars = 0;
  const chunks = [];

  for (let index = 0; index < messages.length; index += 1) {
    const message = messages[index] || {};
    const from = message.from?.emailAddress || {};
    const to = simplifyEmailRecipients(message.toRecipients)
      .map((item) => item.name ? `${item.name} <${item.address}>` : item.address)
      .join(', ');
    const cc = simplifyEmailRecipients(message.ccRecipients)
      .map((item) => item.name ? `${item.name} <${item.address}>` : item.address)
      .join(', ');
    let body = String(message.body?.content || '').trim();
    if (body.length > maxBodyCharsPerMessage) {
      body = `${body.slice(0, maxBodyCharsPerMessage)}\n[Body truncated for analysis size]`;
    }

    const chunk = [
      `MESSAGE ${index + 1}`,
      `Received: ${message.receivedDateTime || ''}`,
      `From: ${from.name || ''}${from.address ? ` <${from.address}>` : ''}`,
      `To: ${to}`,
      `CC: ${cc}`,
      `Subject: ${message.subject || ''}`,
      `Has attachments: ${Boolean(message.hasAttachments)}`,
      'Body:',
      body,
    ].join('\n');

    if (totalChars + chunk.length > maxTotalChars) {
      chunks.push('[Earlier/later thread content omitted because the conversation exceeded the analysis size limit.]');
      break;
    }

    chunks.push(chunk);
    totalChars += chunk.length;
  }

  return chunks.join('\n\n---\n\n');
};

const escapeDelegatedHtml = (value) =>
  String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

const normalizeStringArray = (value) =>
  (Array.isArray(value) ? value : [])
    .map((item) => String(item || '').trim())
    .filter(Boolean);

const renderDelegatedThreadReviewHtml = ({
  review,
  seedSubject,
  lookupQuery,
  focus,
}) => {
  const summary = String(review.executiveSummary || '').trim();
  const supplierPoints = normalizeStringArray(review.supplierPosition);
  const positivePoints = normalizeStringArray(review.positivePoints);
  const questions = normalizeStringArray(review.questionsBeforeSigning);
  const limitations = String(review.limitations || '').trim();
  const recommendation = String(review.recommendation || '').trim();
  const risks = Array.isArray(review.risks) ? review.risks : [];

  const list = (items) =>
    items.length
      ? `<ul style="margin:7px 0 0 18px;padding:0;">${items
          .map(
            (item) =>
              `<li style="margin:0 0 7px 0;line-height:20px;">${escapeDelegatedHtml(item)}</li>`
          )
          .join('')}</ul>`
      : '<div style="margin-top:6px;color:#667085;">None identified from the available thread.</div>';

  const riskHtml = risks.length
    ? risks
        .map((risk) => {
          const level = String(risk?.level || 'Review').toUpperCase();
          const color = level === 'HIGH' ? '#B42318' : level === 'MEDIUM' ? '#B54708' : '#175CD3';
          return `<div style="border:1px solid #eaecf0;border-left:4px solid ${color};border-radius:7px;padding:11px 13px;margin:8px 0;background:#fff;">
            <div style="font-weight:700;color:#101828;">${escapeDelegatedHtml(risk?.title || 'Risk')}</div>
            <div style="margin-top:4px;color:#475467;line-height:20px;">${escapeDelegatedHtml(risk?.detail || '')}</div>
          </div>`;
        })
        .join('')
    : '<div style="margin-top:6px;color:#667085;">No specific risk was identified from the available thread.</div>';

  const section = (title, color, body) => `
    <tr><td style="padding:0 24px 18px;">
      <div style="padding-bottom:8px;border-bottom:2px solid ${color};font-size:13px;font-weight:800;letter-spacing:.45px;color:${color};">${escapeDelegatedHtml(title)}</div>
      <div style="padding-top:8px;font-size:13px;line-height:20px;color:#344054;">${body}</div>
    </td></tr>`;

  return `<!doctype html><html><body style="margin:0;background:#f2f4f7;font-family:Arial,Helvetica,sans-serif;color:#101828;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td align="center" style="padding:24px 10px;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:760px;background:#f8fafc;border:1px solid #e4e7ec;border-radius:12px;overflow:hidden;">
        <tr><td style="background:#102A43;padding:22px 24px;border-bottom:4px solid #C8A96B;">
          <div style="font-size:11px;font-weight:700;letter-spacing:1.8px;color:#C8A96B;">LONDON ASSISTANT · MINACO</div>
          <div style="margin-top:5px;font-size:23px;font-weight:800;color:#fff;">Email Thread Review</div>
          <div style="margin-top:5px;font-size:12px;color:#d0d5dd;">${escapeDelegatedHtml(seedSubject || lookupQuery || 'Supplier correspondence')}</div>
          ${focus ? `<div style="margin-top:4px;font-size:12px;color:#d0d5dd;">Focus: ${escapeDelegatedHtml(focus)}</div>` : ''}
        </td></tr>
        <tr><td style="padding:18px 24px;">
          <div style="background:#fff;border:1px solid #d0d5dd;border-radius:8px;padding:14px 16px;">
            <div style="font-size:11px;font-weight:800;letter-spacing:.5px;color:#667085;">EXECUTIVE SUMMARY</div>
            <div style="margin-top:4px;font-size:14px;line-height:21px;font-weight:700;color:#101828;">${escapeDelegatedHtml(summary || 'See the detailed review below.')}</div>
          </div>
        </td></tr>
        ${section('WHAT THE SUPPLIER / THREAD SAYS', '#175CD3', list(supplierPoints))}
        ${section('RISKS / POINTS TO PROTECT BEFORE SIGNING', '#B42318', riskHtml)}
        ${section('POSITIVE POINTS', '#027A48', list(positivePoints))}
        ${section('QUESTIONS TO CLOSE BEFORE SIGNING', '#B54708', list(questions))}
        ${section('LONDON\'S VIEW', '#6941C6', `<div>${escapeDelegatedHtml(recommendation || 'No recommendation could be formed from the available thread.')}</div>`)}
        ${limitations ? section('LIMITATIONS', '#667085', `<div>${escapeDelegatedHtml(limitations)}</div>`) : ''}
      </table>
    </td></tr></table>
  </body></html>`;
};

const findThreadSeedEmail = async ({ mailboxAddress, lookupQuery }) => {
  const query = String(lookupQuery || '').trim();
  if (!query) throw new Error('An email sender, company, or thread search term is required.');

  try {
    const senderResult = await findLatestInboundEmailBySender({
      mailboxAddress,
      personQuery: query,
    });
    if (senderResult.match) return senderResult.match;
  } catch (error) {
    console.warn('Delegated thread sender lookup warning:', error.message);
  }

  const searched = await searchMailboxEmails({
    mailboxAddress,
    query,
    limit: 10,
  });
  if (!searched.length) {
    throw new Error(`No email thread matching ${query} was found.`);
  }
  return searched[0];
};

const runDelegatedThreadReviewJob = async ({
  jobId,
  lookupQuery,
  focus,
  instruction,
  outputLanguage,
}) => {
  const job = delegatedThreadReviewJobs.get(jobId);
  if (job) {
    job.status = 'running';
    job.startedAt = Date.now();
  }

  let seedSubject = lookupQuery;

  try {
    const seed = await findThreadSeedEmail({
      mailboxAddress: RAMY_MINACO_EMAIL,
      lookupQuery,
    });
    const fullSeed = await getFullMailboxEmail(RAMY_MINACO_EMAIL, seed.id);
    seedSubject = fullSeed.subject || seed.subject || lookupQuery;

    if (!fullSeed.conversationId) {
      throw new Error('The selected email does not expose a conversation id for thread retrieval.');
    }

    const thread = await getMailboxConversationThread(
      RAMY_MINACO_EMAIL,
      fullSeed.conversationId,
      50
    );
    if (!thread.length) {
      throw new Error('The email conversation thread could not be retrieved.');
    }

    const threadText = compactThreadForAnalysis(thread);
    const userInstruction =
      String(instruction || '').trim() ||
      'Summarize the complete thread and identify any material commercial, operational, service, warranty, support, responsibility, pricing, schedule, or contractual risks Ramy should understand before signing.';
    const language = String(outputLanguage || 'English').trim() || 'English';

    const data = await callOpenAIResponses({
      model: DOCUMENT_ANALYSIS_MODEL,
      instructions: `You are London Assistant performing a delegated executive review of a live Minaco email conversation for Ramy Mina.
Use ONLY the supplied thread. Do not invent missing terms or facts.
The requested output language is ${language}.
Distinguish clearly between facts stated in the thread and your commercial inference.
If the actual agreement/contract is not included, say that the email thread alone cannot establish all signing risks.
Return ONLY valid JSON with exactly this shape:
{
  "executiveSummary": "short executive summary",
  "supplierPosition": ["fact from the thread", "..."],
  "risks": [{"level":"High|Medium|Low","title":"short title","detail":"why it matters and what to protect"}],
  "positivePoints": ["positive or reassuring point", "..."],
  "questionsBeforeSigning": ["specific question or confirmation to obtain", "..."],
  "recommendation": "London's concise commercial view and recommended next step",
  "limitations": "what cannot be concluded from this thread alone"
}`,
      input: `RAMY'S INSTRUCTION:\n${userInstruction}\n\nFOCUS PERSON / SUPPLIER (if any):\n${focus || 'Not specified'}\n\nLIVE EMAIL THREAD:\n${threadText}`,
    });

    const raw = extractOpenAIResponseText(data);
    if (!raw) throw new Error('The delegated email-thread analysis returned no text.');

    let review;
    try {
      review = JSON.parse(stripJsonCodeFence(raw));
    } catch (error) {
      throw new Error('The delegated email-thread analysis could not be parsed into a structured review.');
    }

    const html = renderDelegatedThreadReviewHtml({
      review,
      seedSubject,
      lookupQuery,
      focus,
    });

    await sendEmailFromLondon({
      to: RAMY_MINACO_EMAIL,
      subject: `LONDON — Email Thread Review | ${seedSubject || lookupQuery}`,
      body: html,
      contentType: 'HTML',
    });

    if (job) {
      job.status = 'completed';
      job.completedAt = Date.now();
      job.subject = seedSubject;
    }

    console.log('DELEGATED THREAD REVIEW COMPLETED:', {
      jobId,
      lookupQuery,
      subject: seedSubject,
    });
  } catch (error) {
    console.error('DELEGATED THREAD REVIEW FAILED:', {
      jobId,
      lookupQuery,
      error: error.message,
    });

    if (job) {
      job.status = 'failed';
      job.completedAt = Date.now();
      job.error = error.message;
    }

    try {
      await sendEmailFromLondon({
        to: RAMY_MINACO_EMAIL,
        subject: `LONDON — Delegated Email Review Could Not Complete | ${seedSubject || lookupQuery}`,
        body: `London could not complete the delegated email-thread review.\n\nTechnical error: ${error.message}\n\nNo external email was sent.`,
        contentType: 'Text',
      });
    } catch (notifyError) {
      console.error('Delegated review failure notification email also failed:', notifyError);
    }
  } finally {
    pruneDelegatedThreadReviewJobs();
  }
};

const queueDelegatedThreadReviewJob = ({
  lookupQuery,
  focus = '',
  instruction = '',
  outputLanguage = 'English',
}) => {
  const query = String(lookupQuery || '').trim();
  if (!query) throw new Error('An email sender, company, or thread search term is required.');
  if (!RAMY_MINACO_EMAIL) throw new Error('RAMY_MINACO_EMAIL is not configured.');

  const jobId = `thread-review-${Date.now()}-${Math.random()
    .toString(36)
    .slice(2, 8)}`;

  delegatedThreadReviewJobs.set(jobId, {
    jobId,
    status: 'queued',
    createdAt: Date.now(),
    lookupQuery: query,
    focus: String(focus || '').trim(),
  });

  setImmediate(() => {
    runDelegatedThreadReviewJob({
      jobId,
      lookupQuery: query,
      focus,
      instruction,
      outputLanguage,
    }).catch((error) => {
      console.error('Unexpected delegated thread review job failure:', error);
    });
  });

  return {
    success: true,
    queued: true,
    jobId,
    lookupQuery: query,
    emailDestination: RAMY_MINACO_EMAIL,
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

  const response = await fetchWithTimeout(
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


const sendEmailFromMailbox = async ({ from, to, subject, body, contentType = 'Text' }) => {
  const sender = String(from || '').trim().toLowerCase();
  const allowed = new Set(
    [RAMY_MINACO_EMAIL, LONDON_MINACO_EMAIL]
      .filter(Boolean)
      .map((value) => String(value).trim().toLowerCase())
  );
  if (!allowed.has(sender)) {
    throw new Error('The requested sender mailbox is not authorized for London.');
  }
  if (!to || !subject || !body) throw new Error('Email requires recipient, subject, and body.');
  const normalizedContentType = String(contentType || '').toLowerCase() === 'html' ? 'HTML' : 'Text';
  const token = await getMicrosoftGraphActionsToken();
  const response = await fetchWithTimeout(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(sender)}/sendMail`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message: {
          subject,
          body: { contentType: normalizedContentType, content: body },
          toRecipients: [{ emailAddress: { address: to } }],
        },
        saveToSentItems: true,
      }),
    },
    15000
  );
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Microsoft Graph email send failed from ${sender}: ${errorText || response.status}`);
  }
  return { success: true, from: sender, to, subject, contentType: normalizedContentType };
};

const normalizeDropboxPath = (value = '') => {
  const raw = String(value || '').replace(/\\/g, '/').trim();
  const absolute = raw.startsWith('/') ? raw : `${LONDON_DROPBOX_ROOT}/${raw}`;
  const parts = [];
  for (const part of absolute.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') throw new Error('Dropbox path traversal is not allowed.');
    parts.push(part);
  }
  const normalized = `/${parts.join('/')}`;
  const root = LONDON_DROPBOX_ROOT.replace(/\/$/, '');
  if (normalized !== root && !normalized.startsWith(`${root}/`)) {
    throw new Error(`Dropbox access is restricted to ${root}.`);
  }
  return normalized;
};

const dropboxApi = async (endpoint, body, { download = false } = {}) => {
  if (!DROPBOX_ACCESS_TOKEN) throw new Error('DROPBOX_ACCESS_TOKEN is not configured.');
  const url = `${download ? 'https://content.dropboxapi.com/2' : 'https://api.dropboxapi.com/2'}/${endpoint}`;
  const response = await fetchWithTimeout(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${DROPBOX_ACCESS_TOKEN}`,
      ...(download ? { 'Dropbox-API-Arg': JSON.stringify(body).replace(/[\u007F-\uFFFF]/g, (c) => `\\u${c.charCodeAt(0).toString(16).padStart(4, '0')}`) } : { 'Content-Type': 'application/json' }),
    },
    ...(download ? {} : { body: JSON.stringify(body) }),
  }, 20000);
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Dropbox API ${endpoint} failed: ${errorText || response.status}`);
  }
  return response;
};

const listLondonDropbox = async (path = LONDON_DROPBOX_ROOT, recursive = false) => {
  const safePath = normalizeDropboxPath(path);
  const response = await dropboxApi('files/list_folder', { path: safePath, recursive: Boolean(recursive), limit: 200 });
  const data = await response.json();
  const entries = [...(data.entries || [])];
  let cursor = data.cursor;
  let hasMore = Boolean(data.has_more);
  while (hasMore && entries.length < 1000) {
    const next = await dropboxApi('files/list_folder/continue', { cursor });
    const nextData = await next.json();
    entries.push(...(nextData.entries || []));
    cursor = nextData.cursor;
    hasMore = Boolean(nextData.has_more);
  }
  return entries.slice(0, 1000).map((entry) => ({
    type: entry['.tag'],
    name: entry.name,
    path: entry.path_display || entry.path_lower || '',
    size: Number(entry.size || 0),
    modified: entry.server_modified || '',
  }));
};

const searchLondonDropbox = async (query) => {
  const q = String(query || '').trim();
  if (!q) throw new Error('A Dropbox search term is required.');
  const response = await dropboxApi('files/search_v2', {
    query: q,
    options: { path: normalizeDropboxPath(LONDON_DROPBOX_ROOT), max_results: 100, filename_only: false },
  });
  const data = await response.json();
  return (data.matches || []).map((match) => {
    const metadata = match.metadata?.metadata || match.metadata || {};
    return {
      type: metadata['.tag'] || '',
      name: metadata.name || '',
      path: metadata.path_display || metadata.path_lower || '',
      size: Number(metadata.size || 0),
      modified: metadata.server_modified || '',
    };
  }).filter((item) => item.path && normalizeDropboxPath(item.path));
};

const readLondonDropboxText = async (path) => {
  const safePath = normalizeDropboxPath(path);
  const metaResponse = await dropboxApi('files/get_metadata', { path: safePath });
  const meta = await metaResponse.json();
  const size = Number(meta.size || 0);
  if (size > MAX_DROPBOX_TEXT_BYTES) throw new Error('Dropbox file is too large for direct text reading.');
  const downloadResponse = await dropboxApi('files/download', { path: safePath }, { download: true });
  const contentType = String(downloadResponse.headers.get('content-type') || '').toLowerCase();
  const buffer = Buffer.from(await downloadResponse.arrayBuffer());
  const ext = safePath.split('.').pop()?.toLowerCase() || '';
  const textLike = contentType.startsWith('text/') || ['txt','md','csv','json','js','ts','html','xml','yaml','yml','log'].includes(ext);
  if (!textLike) throw new Error('This Dropbox file is not plain text. Use the Task Inbox for PDF, Word, Excel, PowerPoint, or image analysis.');
  return { path: safePath, size: buffer.length, text: buffer.toString('utf8') };
};

const downloadLondonDropboxFile = async (path) => {
  const safePath = normalizeDropboxPath(path);
  const metaResponse = await dropboxApi('files/get_metadata', { path: safePath });
  const meta = await metaResponse.json();
  const size = Number(meta.size || 0);
  if (size > MAX_DROPBOX_ANALYSIS_BYTES) {
    throw new Error(
      `Dropbox file is too large for direct analysis (${Math.round(size / 1024 / 1024)} MB). The current limit is 45 MB.`
    );
  }
  const downloadResponse = await dropboxApi('files/download', { path: safePath }, { download: true });
  const buffer = Buffer.from(await downloadResponse.arrayBuffer());
  const filename = meta.name || safePath.split('/').at(-1) || 'dropbox-file';
  const contentType = resolveFileContentType({
    filename,
    reportedContentType:
      downloadResponse.headers.get('content-type') || meta.mime_type || '',
    buffer,
  });
  return {
    path: safePath,
    filename,
    contentType,
    size: buffer.length,
    buffer,
  };
};

const analyzeLondonDropboxFile = async ({ path, instruction }) => {
  const file = await downloadLondonDropboxFile(path);
  const dataUri = `data:${file.contentType};base64,${file.buffer.toString('base64')}`;
  const requestInstruction =
    String(instruction || '').trim() ||
    'Analyze this file accurately. Summarize the important content and identify decisions, deadlines, amounts, risks, discrepancies, and recommended next actions that matter to Ramy.';

  const content = [];
  if (file.contentType.startsWith('image/')) {
    content.push({ type: 'input_image', image_url: dataUri, detail: 'auto' });
  } else {
    content.push({
      type: 'input_file',
      filename: file.filename,
      file_data: dataUri,
      detail: 'auto',
    });
  }
  content.push({ type: 'input_text', text: requestInstruction });

  const data = await callOpenAIResponses(
    {
      model: DOCUMENT_ANALYSIS_MODEL,
      instructions:
        'You are London Assistant analyzing a real file retrieved directly from Ramy Mina’s controlled Dropbox workspace. Use the actual file only. Do not invent missing content. Preserve names, dates, amounts, drawing references, revision numbers, legal/commercial terms, and material qualifiers. If the file is a drawing set or technical document, focus on the user’s requested review and clearly distinguish observed facts from your interpretation.',
      input: [{ role: 'user', content }],
    },
    120000
  );

  const result = extractOpenAIResponseText(data);
  if (!result) throw new Error('The Dropbox file analysis returned no text.');
  return {
    path: file.path,
    filename: file.filename,
    contentType: file.contentType,
    size: file.size,
    result,
  };
};

const delegatedDropboxAnalysisJobs = new Map();

const pruneDelegatedDropboxAnalysisJobs = () => {
  if (delegatedDropboxAnalysisJobs.size <= 50) return;
  const entries = [...delegatedDropboxAnalysisJobs.entries()].sort(
    (a, b) => Number(a[1]?.createdAt || 0) - Number(b[1]?.createdAt || 0)
  );
  for (const [jobId] of entries.slice(0, delegatedDropboxAnalysisJobs.size - 50)) {
    delegatedDropboxAnalysisJobs.delete(jobId);
  }
};

const runDelegatedDropboxAnalysisJob = async ({ jobId, path, instruction }) => {
  const job = delegatedDropboxAnalysisJobs.get(jobId);
  if (job) {
    job.status = 'running';
    job.startedAt = Date.now();
  }

  try {
    const analysis = await analyzeLondonDropboxFile({ path, instruction });
    const body = `London Dropbox Analysis\n\nFile: ${analysis.filename}\nPath: ${analysis.path}\n\n${analysis.result}`;
    await sendEmailFromLondon({
      to: RAMY_MINACO_EMAIL,
      subject: `LONDON — Dropbox Analysis | ${analysis.filename}`,
      body,
      contentType: 'Text',
    });

    if (job) {
      job.status = 'completed';
      job.completedAt = Date.now();
      job.filename = analysis.filename;
    }
    console.log('DELEGATED DROPBOX ANALYSIS COMPLETED:', {
      jobId,
      path: analysis.path,
      filename: analysis.filename,
    });
  } catch (error) {
    console.error('DELEGATED DROPBOX ANALYSIS FAILED:', {
      jobId,
      path,
      error: error.message,
    });
    if (job) {
      job.status = 'failed';
      job.completedAt = Date.now();
      job.error = error.message;
    }
    try {
      await sendEmailFromLondon({
        to: RAMY_MINACO_EMAIL,
        subject: 'LONDON — Dropbox Analysis Could Not Complete',
        body: `London could not complete the Dropbox file analysis.\n\nFile: ${path}\nTechnical issue: ${error.message}\n\nNo external email or commitment was sent.`,
        contentType: 'Text',
      });
    } catch (notifyError) {
      console.error('Dropbox analysis failure notification also failed:', notifyError);
    }
  } finally {
    pruneDelegatedDropboxAnalysisJobs();
  }
};

const queueDelegatedDropboxAnalysisJob = ({ path, instruction = '' }) => {
  const safePath = normalizeDropboxPath(path);
  if (!RAMY_MINACO_EMAIL) throw new Error('RAMY_MINACO_EMAIL is not configured.');
  const jobId = `dropbox-analysis-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  delegatedDropboxAnalysisJobs.set(jobId, {
    jobId,
    status: 'queued',
    createdAt: Date.now(),
    path: safePath,
  });
  setImmediate(() => {
    runDelegatedDropboxAnalysisJob({ jobId, path: safePath, instruction }).catch((error) => {
      console.error('Unexpected delegated Dropbox analysis failure:', error);
    });
  });
  return {
    success: true,
    queued: true,
    jobId,
    path: safePath,
    emailDestination: RAMY_MINACO_EMAIL,
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

const normalizeGraphLocalDateTime = (value) => {
  const text = String(value || '').trim();
  const match = text.match(
    /^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2})(?::(\d{2}))?(?:\.\d+)?$/
  );
  if (!match) return text;
  return `${match[1]}:${match[2] || '00'}`;
};

const ensureLocalDateTime = (value, label) => {
  if (typeof value !== 'string') {
    throw new Error(`${label} must use YYYY-MM-DDTHH:mm:ss Montreal local time.`);
  }

  // Microsoft Graph commonly returns Outlook dateTime values with fractional
  // seconds (for example 2026-09-02T13:30:00.0000000). They still represent
  // Montreal-local wall time because calendar requests use the Outlook
  // Eastern time-zone preference. Normalize those values before validating.
  const normalized = normalizeGraphLocalDateTime(value);

  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(normalized)) {
    throw new Error(`${label} must use YYYY-MM-DDTHH:mm:ss Montreal local time.`);
  }

  return normalized;
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
  // Normalize Graph's fractional-second Outlook timestamps once at the edge
  // so every downstream availability calculation receives a stable local
  // YYYY-MM-DDTHH:mm:ss value.
  startLocal: normalizeGraphLocalDateTime(event.start?.dateTime || ''),
  endLocal: normalizeGraphLocalDateTime(event.end?.dateTime || ''),
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

  const response = await fetchWithTimeout(url, {
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

const normalizeDateOnly = (value) => {
  const match = String(value || '').trim().match(/^(\d{4}-\d{2}-\d{2})(?:T.*)?$/);
  return match ? match[1] : '';
};

const weekdayForDateOnly = (dateString) => {
  const normalized = normalizeDateOnly(dateString);
  if (!normalized) throw new Error('Invalid date.');
  const [year, month, day] = normalized.split('-').map(Number);
  return new Date(Date.UTC(year, month - 1, day)).getUTCDay();
};

const resolveAvailabilityDateRange = ({
  startDate,
  endDate,
  dateRangeText,
} = {}) => {
  const rawStart = String(startDate || '').trim();
  const rawEnd = String(endDate || '').trim();
  const normalizedStart = normalizeDateOnly(rawStart);
  const normalizedEnd = normalizeDateOnly(rawEnd);

  // If the model already supplied real dates, trust them after normalization.
  if (normalizedStart && normalizedEnd) {
    if (normalizedEnd < normalizedStart) {
      throw new Error('The availability end date is before the start date.');
    }
    return {
      startDate: normalizedStart,
      endDate: normalizedEnd,
      source: 'explicit_dates',
    };
  }

  if (normalizedStart && !rawEnd) {
    return {
      startDate: normalizedStart,
      endDate: normalizedStart,
      source: 'single_explicit_date',
    };
  }

  // If a relative phrase accidentally arrived in start_date/end_date, recover it.
  const phrase = String(
    dateRangeText ||
      (!normalizedStart && rawStart ? rawStart : '') ||
      (!normalizedEnd && rawEnd ? rawEnd : '') ||
      ''
  )
    .trim()
    .toLowerCase()
    .replace(/[–—]/g, '-')
    .replace(/\s+/g, ' ');

  const today = montrealDateParts();
  const todayWeekday = weekdayForDateOnly(today); // Sun=0, Mon=1 ... Sat=6

  if (!phrase) {
    throw new Error(
      'No usable availability date range was provided. Use YYYY-MM-DD dates or a phrase such as "next week".'
    );
  }

  if (phrase === 'today') {
    return { startDate: today, endDate: today, source: 'today' };
  }

  if (phrase === 'tomorrow') {
    const tomorrow = addDaysToLocalDate(today, 1);
    return { startDate: tomorrow, endDate: tomorrow, source: 'tomorrow' };
  }

  if (
    phrase === 'next week' ||
    phrase === 'next business week' ||
    phrase === 'the next week' ||
    phrase === 'following week' ||
    phrase === 'the following week'
  ) {
    // Always resolve to the next Monday-Friday in Montreal.
    let daysUntilNextMonday = (8 - todayWeekday) % 7;
    if (daysUntilNextMonday === 0) daysUntilNextMonday = 7;
    const monday = addDaysToLocalDate(today, daysUntilNextMonday);
    const friday = addDaysToLocalDate(monday, 4);
    return { startDate: monday, endDate: friday, source: 'next_week' };
  }

  if (phrase === 'this week' || phrase === 'the rest of this week') {
    if (todayWeekday >= 1 && todayWeekday <= 5) {
      const friday = addDaysToLocalDate(today, 5 - todayWeekday);
      return { startDate: today, endDate: friday, source: 'this_week' };
    }
    // On a weekend, use the immediately upcoming Monday-Friday.
    const daysUntilMonday = todayWeekday === 6 ? 2 : 1;
    const monday = addDaysToLocalDate(today, daysUntilMonday);
    const friday = addDaysToLocalDate(monday, 4);
    return { startDate: monday, endDate: friday, source: 'this_week_weekend' };
  }

  const nextDaysMatch = phrase.match(/^next\s+(\d{1,2})\s+days?$/);
  if (nextDaysMatch) {
    const count = Math.min(Math.max(Number(nextDaysMatch[1]), 1), 31);
    return {
      startDate: today,
      endDate: addDaysToLocalDate(today, count - 1),
      source: 'next_n_days',
    };
  }

  // Also accept an explicit date range supplied as text.
  const explicitTextRange = phrase.match(
    /^(\d{4}-\d{2}-\d{2})\s*(?:to|through|thru|-)\s*(\d{4}-\d{2}-\d{2})$/
  );
  if (explicitTextRange) {
    const [, first, last] = explicitTextRange;
    if (last < first) throw new Error('The availability end date is before the start date.');
    return { startDate: first, endDate: last, source: 'text_date_range' };
  }

  throw new Error(
    `Could not resolve availability date range "${phrase}". Say "next week", "this week", "today", "tomorrow", "next N days", or provide YYYY-MM-DD dates.`
  );
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
  const limit = Math.min(Math.max(Number(maxSlots) || 5, 1), 200);

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


const selectRepresentativeAvailabilitySlots = (slots, maxSlots = 3) => {
  const limit = Math.min(Math.max(Number(maxSlots) || 3, 1), 5);
  if (!Array.isArray(slots) || slots.length === 0) return [];

  const byDate = new Map();
  for (const slot of slots) {
    const date = String(slot.startLocal || '').slice(0, 10);
    if (!date) continue;
    if (!byDate.has(date)) byDate.set(date, []);
    byDate.get(date).push(slot);
  }

  const targetMinutes = [10 * 60, 13 * 60 + 30, 15 * 60, 9 * 60 + 30, 14 * 60 + 30];
  const chosen = [];
  const usedKeys = new Set();

  const slotMinutes = (slot) => {
    const match = String(slot.startLocal || '').match(/T(\d{2}):(\d{2})/);
    return match ? Number(match[1]) * 60 + Number(match[2]) : 0;
  };

  // Prefer one useful option on different days before offering multiple
  // choices on the same day.
  let dayIndex = 0;
  for (const [, daySlots] of [...byDate.entries()].sort(([a], [b]) =>
    a.localeCompare(b)
  )) {
    if (chosen.length >= limit) break;
    const target = targetMinutes[dayIndex % targetMinutes.length];
    const sorted = [...daySlots].sort(
      (a, b) =>
        Math.abs(slotMinutes(a) - target) -
        Math.abs(slotMinutes(b) - target)
    );
    if (sorted[0]) {
      const key = `${sorted[0].startLocal}|${sorted[0].endLocal}`;
      chosen.push(sorted[0]);
      usedKeys.add(key);
    }
    dayIndex += 1;
  }

  if (chosen.length < limit) {
    for (const slot of slots) {
      if (chosen.length >= limit) break;
      const key = `${slot.startLocal}|${slot.endLocal}`;
      if (usedKeys.has(key)) continue;

      // Avoid stacking nearly identical half-hour options when possible.
      const tooClose = chosen.some((existing) => {
        const sameDate =
          String(existing.startLocal || '').slice(0, 10) ===
          String(slot.startLocal || '').slice(0, 10);
        if (!sameDate) return false;
        return Math.abs(slotMinutes(existing) - slotMinutes(slot)) < 90;
      });
      if (tooClose) continue;

      chosen.push(slot);
      usedKeys.add(key);
    }
  }

  return chosen.slice(0, limit);
};

const looksFrench = (value) => {
  const text = ` ${String(value || '').toLowerCase()} `;
  const markers = [
    ' bonjour ',
    ' merci ',
    ' votre ',
    ' vous ',
    ' pour ',
    ' concernant ',
    ' disponibilité ',
    ' disponibilite ',
    ' cordialement ',
    ' veuillez ',
  ];
  return markers.filter((marker) => text.includes(marker)).length >= 2;
};

const formatLocalSlotForReply = (startLocal, locale = 'en-CA') => {
  const normalized = ensureLocalDateTime(startLocal, 'slot_start');
  const pseudoMs = parseLocalDateTimePseudoMs(normalized);
  return new Intl.DateTimeFormat(locale, {
    timeZone: 'UTC',
    weekday: 'long',
    month: 'long',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
    hour12: locale.startsWith('en'),
  }).format(new Date(pseudoMs));
};

const buildAvailabilityReplyDraft = async ({
  personQuery,
  startDate,
  endDate,
  dateRangeText,
  durationMinutes = 30,
  dayStart = '09:00',
  dayEnd = '17:00',
  maxSlots = 3,
  includeWeekends = false,
  replyMode = 'sender',
}) => {
  const query = String(personQuery || '').trim();
  if (!query) throw new Error('A sender name or email is required.');
  if (!['sender', 'all'].includes(replyMode)) {
    throw new Error('Reply mode must be sender or all.');
  }

  const resolvedRange = resolveAvailabilityDateRange({
    startDate,
    endDate,
    dateRangeText,
  });

  // Resolve the actual inbound sender and calendar openings in parallel.
  // The sender lookup checks Inbox From metadata rather than a broad message
  // keyword result, so it cannot accidentally choose Ramy's own sent message.
  const [senderLookup, availability] = await Promise.all([
    findLatestInboundEmailBySender({
      mailboxAddress: RAMY_MINACO_EMAIL,
      personQuery: query,
    }),
    findCalendarAvailability({
      startDate: resolvedRange.startDate,
      endDate: resolvedRange.endDate,
      durationMinutes,
      dayStart,
      dayEnd,
      // Collect enough candidates to spread recommendations across days.
      maxSlots: 100,
      includeWeekends,
    }),
  ]);

  if (!senderLookup.match) {
    return {
      needsClarification: true,
      reason: senderLookup.reason || `No inbound email from ${query} was found.`,
      candidates: (senderLookup.candidates || []).slice(0, 5).map((email) => ({
        messageId: email.id,
        from: email.from?.emailAddress?.name || '',
        fromEmail: email.from?.emailAddress?.address || '',
        subject: email.subject || '(No subject)',
        receivedDateTime: email.receivedDateTime || '',
      })),
    };
  }

  const email = senderLookup.match;
  const slots = selectRepresentativeAvailabilitySlots(
    availability.slots,
    maxSlots
  );

  if (slots.length === 0) {
    return {
      needsClarification: true,
      reason: 'No open calendar slots were found in the requested period.',
      candidates: [],
    };
  }

  const senderName = String(email.from?.emailAddress?.name || '').trim();
  const firstName = senderName.split(/\s+/)[0] || 'there';
  const recipients = getReplyRecipientSummary(email, replyMode);

  if (recipients.length === 0) {
    throw new Error('No valid reply recipient could be determined.');
  }

  const sourceText = `${email.subject || ''}\n${email.bodyPreview || ''}`;
  const useFrench = looksFrench(sourceText);

  let body;
  if (useFrench) {
    const slotLines = slots
      .map(
        (slot, index) =>
          `${index + 1}. ${formatLocalSlotForReply(slot.startLocal, 'fr-CA')}`
      )
      .join('\n');

    body = `Bonjour ${firstName},\n\nMerci pour votre courriel. Voici trois plages de disponibilité pour la semaine prochaine, heure de Montréal :\n\n${slotLines}\n\nMerci de me confirmer la plage qui vous convient le mieux.\n\nCordialement,\nRamy`;
  } else {
    const slotLines = slots
      .map(
        (slot, index) =>
          `${index + 1}. ${formatLocalSlotForReply(slot.startLocal, 'en-CA')}`
      )
      .join('\n');

    body = `Hi ${firstName},\n\nThank you for your email. Here are three available times for next week, Montreal time:\n\n${slotLines}\n\nPlease let me know which option works best for you.\n\nBest,\nRamy`;
  }

  return {
    needsClarification: false,
    email: {
      messageId: email.id,
      subject: email.subject || '(No subject)',
      from: senderName || email.from?.emailAddress?.address || '',
      fromEmail: email.from?.emailAddress?.address || '',
      receivedDateTime: email.receivedDateTime || '',
      preview: email.bodyPreview || '',
    },
    replyMode,
    recipients,
    availability: {
      startDate: resolvedRange.startDate,
      endDate: resolvedRange.endDate,
      rangeSource: resolvedRange.source,
      durationMinutes: availability.durationMinutes,
      slots,
    },
    body,
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

  const response = await fetchWithTimeout(
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

  const response = await fetchWithTimeout(
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

  const response = await fetchWithTimeout(
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
  const listResponse = await fetchWithTimeout(
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

  const createResponse = await fetchWithTimeout(
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

  const response = await fetchWithTimeout(
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
    const response = await fetchWithTimeout(url, {
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

  const getResponse = await fetchWithTimeout(
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

  const response = await fetchWithTimeout(
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
// Unified quick Action Register updates for SMS, WhatsApp, and Voice
// -----------------------------------------------------------------------------

const messagingState = new Map();
const processedMessagingSids = new Map();
const MESSAGING_STATE_TTL_MS = 6 * 60 * 60 * 1000;
const MESSAGING_DEDUPE_TTL_MS = 24 * 60 * 60 * 1000;


const loadPersistentMessagingState = () => {
  try {
    const parsed = JSON.parse(readFileSync(STATE_FILE, 'utf8'));
    for (const [key, value] of Object.entries(parsed.messagingState || {})) {
      if (value && typeof value === 'object') messagingState.set(key, value);
    }
    for (const [key, value] of Object.entries(parsed.processedMessagingSids || {})) {
      processedMessagingSids.set(key, Number(value));
    }
  } catch (error) {
    if (error?.code !== 'ENOENT') console.warn('STATE load warning:', error.message);
  }
};

let stateSaveTimer = null;
const persistMessagingState = () => {
  if (stateSaveTimer) return;
  stateSaveTimer = setTimeout(() => {
    stateSaveTimer = null;
    try {
      mkdirSync(dirname(STATE_FILE), { recursive: true });
      const tmp = `${STATE_FILE}.tmp`;
      writeFileSync(tmp, JSON.stringify({
        savedAt: new Date().toISOString(),
        messagingState: Object.fromEntries(messagingState.entries()),
        processedMessagingSids: Object.fromEntries(processedMessagingSids.entries()),
      }));
      renameSync(tmp, STATE_FILE);
    } catch (error) {
      console.warn('STATE save warning:', error.message);
    }
  }, 250);
  if (typeof stateSaveTimer.unref === 'function') stateSaveTimer.unref();
};

loadPersistentMessagingState();

const pruneMessagingState = () => {
  const now = Date.now();
  for (const [key, value] of messagingState.entries()) {
    if (!value?.updatedAt || now - value.updatedAt > MESSAGING_STATE_TTL_MS) {
      messagingState.delete(key);
    }
  }
  for (const [key, timestamp] of processedMessagingSids.entries()) {
    if (now - timestamp > MESSAGING_DEDUPE_TTL_MS) processedMessagingSids.delete(key);
  }
};

const messagingKey = (channel, sender) =>
  `${String(channel || 'sms').toLowerCase()}:${normalizePhoneIdentity(sender)}`;

const actionCompactForParser = (action, index) => ({
  index: index + 1,
  title: action.title,
  project: action.project,
  status: action.status,
  priority: action.priority,
  owner: action.owner,
  waitingOn: action.waitingOn,
  nextFollowUp: action.nextFollowUp,
  hardDeadline: action.hardDeadline,
  nextAction: action.nextAction,
  ramyRequired: action.ramyRequired,
});

const messagingActionLine = (action, index) => {
  const detail = [];
  if (action.status) detail.push(action.status);
  if (action.waitingOn) detail.push(`waiting on ${action.waitingOn}`);
  if (action.nextFollowUp) detail.push(`follow-up ${action.nextFollowUp}`);
  if (action.hardDeadline) detail.push(`deadline ${action.hardDeadline}`);
  return `${index + 1}. ${action.title}${detail.length ? ` — ${detail.join('; ')}` : ''}`;
};

const formatActionListForMessaging = (actions, heading = 'Open actions') => {
  const visible = actions.slice(0, 8);
  if (!visible.length) return `${heading}: none.`;
  return `${heading}:\n${visible.map(messagingActionLine).join('\n')}\nReply naturally, e.g. “1 done”, “2 waiting on Anass until Friday”, or “add task: call Makar tomorrow”.`;
};

const parseDateOnlySafe = (value) => {
  const match = String(value || '').trim().match(/^(\d{4}-\d{2}-\d{2})$/);
  return match ? match[1] : '';
};

const nextWeekdayFromMontrealToday = (weekday, forceNext = false) => {
  const today = montrealDateParts();
  const current = weekdayForDateOnly(today);
  let delta = (weekday - current + 7) % 7;
  if (delta === 0 && forceNext) delta = 7;
  return addDaysToLocalDate(today, delta);
};

const resolveSimpleActionDateText = (value) => {
  const raw = String(value || '').trim();
  const exact = parseDateOnlySafe(raw);
  if (exact) return exact;
  const phrase = raw.toLowerCase().replace(/[.,]/g, '').replace(/\s+/g, ' ').trim();
  if (!phrase) return '';
  if (phrase === 'today') return montrealDateParts();
  if (phrase === 'tomorrow') return addDaysToLocalDate(montrealDateParts(), 1);
  if (phrase === 'next week') {
    return resolveAvailabilityDateRange({ dateRangeText: 'next week' }).startDate;
  }
  const names = {
    sunday: 0,
    monday: 1,
    tuesday: 2,
    wednesday: 3,
    thursday: 4,
    friday: 5,
    saturday: 6,
  };
  const weekdayMatch = phrase.match(/^(next\s+)?(sunday|monday|tuesday|wednesday|thursday|friday|saturday)$/);
  if (weekdayMatch) {
    return nextWeekdayFromMontrealToday(names[weekdayMatch[2]], Boolean(weekdayMatch[1]));
  }
  return '';
};

const normalizedMessagingChanges = (changes = {}) => {
  const out = {};
  const copy = (input, output = input) => {
    if (Object.prototype.hasOwnProperty.call(changes, input) && changes[input] !== null) {
      out[output] = changes[input];
    }
  };
  copy('status');
  copy('priority');
  copy('owner');
  copy('next_action', 'nextAction');
  copy('waiting_on', 'waitingOn');
  copy('ramy_required', 'ramyRequired');
  copy('risk_if_delayed', 'riskIfDelayed');
  copy('notes');

  for (const [input, output] of [
    ['next_follow_up', 'nextFollowUp'],
    ['hard_deadline', 'hardDeadline'],
    ['promised_date', 'promisedDate'],
    ['ramy_decision_by', 'ramyDecisionBy'],
  ]) {
    if (!Object.prototype.hasOwnProperty.call(changes, input) || changes[input] == null) continue;
    const raw = String(changes[input]).trim();
    const resolved = resolveSimpleActionDateText(raw) || parseDateOnlySafe(raw);
    if (resolved) out[output] = resolved;
  }
  return out;
};

const parseMessagingActionCommand = async ({ text, actions, recentList, recentMessages = [] }) => {
  const context = {
    currentMontrealDate: montrealDateParts(),
    currentMontrealDateTime: formatMontrealDateTime(new Date()),
    openActions: actions.map(actionCompactForParser),
    recentNumberedList: Array.isArray(recentList)
      ? recentList.map((action, index) => ({ index: index + 1, title: action.title }))
      : [],
    recentMessages: Array.isArray(recentMessages)
      ? recentMessages.slice(-6).map((item) => String(item || '').trim()).filter(Boolean)
      : [],
    message: String(text || '').trim(),
  };

  const response = await callOpenAIResponses(
    {
      model: ACTION_MESSAGE_MODEL,
      instructions: `You are London's Action Register command parser for Ramy Mina. Interpret a short SMS, WhatsApp, or voice status update and return ONLY valid JSON, no Markdown.
The current Montreal date/time and current open actions are supplied.
Ramy should be able to write naturally: “Joannie done”, “waiting on Anass until Friday”, “follow up with Franco next Tuesday”, “add task: call Makar tomorrow”, “cancel the EV item”, “what's overdue?”, or several updates in one message.
IMPORTANT: SMS and WhatsApp are conversational. Use recentMessages to resolve short follow-up fragments such as “tomorrow”, “yes”, “that one”, “set it as a reminder”, “put this in tomorrow's to-do list”, or a task description sent in the next message. Treat them as continuations of the immediately preceding conversation when that connection is clear.
If Ramy says "this", "that", or "it", infer the referent from the most recent substantive message whenever there is only one sensible candidate. Do not ask him to repeat information already present in recentMessages.
Example: if a prior message says a WhatsApp template category changed from Utility to Marketing and Ramy then says “put this in the things to do tomorrow”, create a task about reviewing/following up on that specific template-category issue for tomorrow.
Example: if a prior message describes an item, then Ramy says “set as a reminder”, then “tomorrow”, combine those messages into one reminder/action instead of asking him to restart in one sentence.
Do not guess which action Ramy means when two actions are genuinely plausible. In that case return a clarify operation.
Only mark COMPLETED when Ramy clearly says done/completed/finished/resolved/closed. Only mark CANCELLED when he clearly cancels/drops it.
“Waiting on NAME” normally means WAITING - EXTERNAL, waitingOn NAME, ramyRequired false. “Waiting on me/Ramy/my decision” means WAITING - RAMY and ramyRequired true.
“Follow up DATE” sets next_follow_up to a date. Return dates as YYYY-MM-DD whenever possible using the supplied Montreal date. If the user uses a simple phrase such as tomorrow, Friday, next Friday, or next week, you may return that phrase and the server will resolve it.
If Ramy says he personally needs to decide/approve/review, use WAITING - RAMY and ramy_required true.
For new tasks, title must be a concise outcome/action, not the entire message. Do not invent deadlines or people.
If Ramy asks a status question, return a list operation rather than changing anything.
Use target_index to refer ONLY to the numbered openActions list. Never invent an index.
Return this exact JSON shape:
{
  "operations": [
    {
      "type": "update|create|list|clarify|noop",
      "target_index": 1,
      "changes": {
        "status": null,
        "priority": null,
        "owner": null,
        "next_follow_up": null,
        "hard_deadline": null,
        "promised_date": null,
        "next_action": null,
        "waiting_on": null,
        "ramy_required": null,
        "ramy_decision_by": null,
        "risk_if_delayed": null,
        "notes": null
      },
      "create": {
        "title": null,
        "project": null,
        "category": null,
        "owner": "London",
        "promised_date": null,
        "hard_deadline": null,
        "next_follow_up": null,
        "status": "ACTIVE",
        "priority": "NORMAL",
        "next_action": null,
        "waiting_on": null,
        "ramy_required": false,
        "ramy_decision_by": null,
        "risk_if_delayed": null,
        "notes": null
      },
      "list_filter": "open|overdue|waiting_ramy|waiting_external|critical|all",
      "clarification": null
    }
  ]
}
Maximum 6 operations. If the message is just politeness or unrelated, return one noop operation.`,
      input: JSON.stringify(context),
    },
    22000
  );

  const raw = extractOpenAIResponseText(response);
  if (!raw) throw new Error('London could not interpret the update.');
  const parsed = JSON.parse(stripJsonCodeFence(raw));
  return Array.isArray(parsed.operations) ? parsed.operations.slice(0, 6) : [];
};

const actionListByFilter = (actions, filter) => {
  const normalized = String(filter || 'open').toLowerCase();
  if (normalized === 'overdue') return actions.filter((a) => a.overdue);
  if (normalized === 'waiting_ramy') return actions.filter((a) => a.status === 'WAITING - RAMY');
  if (normalized === 'waiting_external') return actions.filter((a) => a.status === 'WAITING - EXTERNAL');
  if (normalized === 'critical') return actions.filter((a) => ['CRITICAL', 'HIGH'].includes(a.priority));
  return actions;
};

const processQuickActionInstruction = async ({ text, channel = 'voice', sender = RAMY_PHONE_NUMBER }) => {
  const cleanText = String(text || '').trim();
  if (!cleanText) return { success: false, reply: 'I did not receive an update to process.' };

  pruneMessagingState();
  const key = messagingKey(channel, sender);
  const state = messagingState.get(key) || { recentList: [], recentMessages: [], lastEventId: '', updatedAt: Date.now() };
  state.recentMessages = Array.isArray(state.recentMessages) ? state.recentMessages : [];
  const priorMessages = state.recentMessages.slice(-5);
  state.recentMessages.push(cleanText);
  state.recentMessages = state.recentMessages.slice(-6);
  state.updatedAt = Date.now();
  messagingState.set(key, state);
    persistMessagingState();
  const lower = cleanText.toLowerCase().trim();

  // Ordinary conversation should not be treated as an Action Register update.
  if (/^(hi|hello|hey|hey london|hi london|hello london)[.!?\s]*$/i.test(cleanText)) {
    return { success: true, reply: 'Hi Ramy. I’m here — what do you need?' };
  }
  if (/^(thanks|thank you|thx|perfect|great)[.!?\s]*$/i.test(cleanText)) {
    return { success: true, reply: 'You’re welcome.' };
  }
  if (/^(are you there|you there|london|london\?)[.!?\s]*$/i.test(cleanText)) {
    return { success: true, reply: 'Yes, I’m here.' };
  }

  const allOpen = await listActionItems({ includeClosed: false, limit: 100 });

  // Very fast common list commands do not need an AI interpretation round-trip.
  let directListFilter = '';
  if (/^(open|open tasks|tasks|status|what'?s open|what is open)$/i.test(cleanText)) directListFilter = 'open';
  if (/overdue/i.test(lower) && /^(what|show|list|overdue)/i.test(lower)) directListFilter = 'overdue';
  if (/waiting on me|waiting on ramy|my decisions|needs my decision/i.test(lower)) directListFilter = 'waiting_ramy';
  if (directListFilter) {
    const selected = actionListByFilter(allOpen, directListFilter).slice(0, 8);
    state.recentList = selected;
    state.updatedAt = Date.now();
    messagingState.set(key, state);
    persistMessagingState();
    return { success: true, reply: formatActionListForMessaging(selected, directListFilter === 'overdue' ? 'Overdue actions' : 'Open actions'), actions: selected };
  }

  // Numbered shortcuts after London has listed items: "1 done", "2 cancel".
  const numbered = cleanText.match(/^\s*(\d{1,2})\s+(done|completed?|finished|cancel(?:led)?|drop(?:ped)?)\s*[.!]?\s*$/i);
  if (numbered && state.recentList?.length) {
    const idx = Number(numbered[1]) - 1;
    const target = state.recentList[idx];
    if (!target) return { success: false, reply: `I don't have item ${numbered[1]} in the last list. Ask me for open tasks again.` };
    const status = /cancel|drop/i.test(numbered[2]) ? 'CANCELLED' : 'COMPLETED';
    const updated = await updateActionItem(target.eventId, {
      status,
      lastContact: montrealDateParts(),
      nextAction: status === 'COMPLETED' ? 'Completed.' : 'Cancelled by Ramy.',
      waitingOn: '',
      ramyRequired: false,
    });
    state.lastEventId = updated.eventId;
    state.updatedAt = Date.now();
    messagingState.set(key, state);
    persistMessagingState();
    return { success: true, reply: `Updated: ${updated.title} — ${status}.`, action: updated };
  }

  const operations = await parseMessagingActionCommand({
    text: cleanText,
    actions: allOpen,
    recentList: state.recentList,
    recentMessages: priorMessages,
  });

  if (!operations.length) return { success: false, reply: 'I could not identify an Action Register update in that message.' };

  const confirmations = [];
  for (const operation of operations) {
    const type = String(operation?.type || '').toLowerCase();

    if (type === 'noop') continue;

    if (type === 'clarify') {
      confirmations.push(String(operation.clarification || 'Which action do you mean?'));
      continue;
    }

    if (type === 'list') {
      const selected = actionListByFilter(allOpen, operation.list_filter).slice(0, 8);
      state.recentList = selected;
      confirmations.push(formatActionListForMessaging(selected, 'Action Register'));
      continue;
    }

    if (type === 'update') {
      const index = Number(operation.target_index);
      const target = Number.isInteger(index) && index >= 1 ? allOpen[index - 1] : null;
      if (!target) {
        confirmations.push('I could not safely match one of those updates to an open action. Please name the item more specifically.');
        continue;
      }
      const changes = normalizedMessagingChanges(operation.changes || {});
      changes.lastContact = montrealDateParts();
      const updated = await updateActionItem(target.eventId, changes);
      state.lastEventId = updated.eventId;
      const summary = [`Updated: ${updated.title}`];
      if (changes.status) summary.push(updated.status);
      if (changes.waitingOn) summary.push(`waiting on ${updated.waitingOn}`);
      if (changes.nextFollowUp) summary.push(`follow-up ${updated.nextFollowUp}`);
      if (changes.hardDeadline) summary.push(`deadline ${updated.hardDeadline}`);
      if (changes.nextAction && !changes.status) summary.push(updated.nextAction);
      confirmations.push(`${summary.join(' — ')}.`);
      continue;
    }

    if (type === 'create') {
      const create = operation.create || {};
      if (!String(create.title || '').trim()) {
        confirmations.push('I understood that you want a new task, but I need the task itself stated more clearly.');
        continue;
      }
      const action = await createActionItem({
        title: create.title,
        project: create.project,
        category: create.category,
        owner: create.owner || 'London',
        promisedDate: resolveSimpleActionDateText(create.promised_date) || create.promised_date || '',
        hardDeadline: resolveSimpleActionDateText(create.hard_deadline) || create.hard_deadline || '',
        nextFollowUp: resolveSimpleActionDateText(create.next_follow_up) || create.next_follow_up || '',
        status: create.status || 'ACTIVE',
        priority: create.priority || 'NORMAL',
        nextAction: create.next_action,
        waitingOn: create.waiting_on,
        ramyRequired: Boolean(create.ramy_required),
        ramyDecisionBy: resolveSimpleActionDateText(create.ramy_decision_by) || create.ramy_decision_by || '',
        riskIfDelayed: create.risk_if_delayed,
        source: String(channel || 'message').toUpperCase(),
        notes: create.notes || `Created from ${channel}: ${cleanText}`,
      });
      state.lastEventId = action.eventId;
      confirmations.push(`Added: ${action.title}${action.nextFollowUp ? ` — follow-up ${action.nextFollowUp}` : ''}.`);
    }
  }

  state.updatedAt = Date.now();
  messagingState.set(key, state);
    persistMessagingState();

  const reply = confirmations.filter(Boolean).join('\n').slice(0, MAX_MESSAGING_REPLY_CHARS);
  return {
    success: Boolean(reply),
    reply: reply || 'No Action Register change was needed.',
  };
};


// -----------------------------------------------------------------------------
// General SMS / WhatsApp executive command router
// -----------------------------------------------------------------------------

const formatMessagingEmailList = (emails = []) => {
  if (!emails.length) return 'I found no recent emails in your Minaco inbox.';
  return emails.slice(0, 5).map((email, index) => {
    const sender =
      email.from?.emailAddress?.name ||
      email.from?.emailAddress?.address ||
      'Unknown sender';
    const subject = email.subject || '(No subject)';
    return `${index + 1}. ${sender} — ${subject}`;
  }).join('\n');
};

const formatMessagingCalendarList = (events = []) => {
  const active = events.filter((event) => !event.isCancelled).slice(0, 8);
  if (!active.length) return 'I found no upcoming meetings on your Minaco calendar.';
  return active.map((event, index) => {
    const when = event.startLocal || event.start?.dateTime || '';
    return `${index + 1}. ${when} — ${event.subject || '(No subject)'}`;
  }).join('\n');
};

const compactTaskTitleFromMessage = (text) => {
  const clean = String(text || '').replace(/\s+/g, ' ').trim();
  return clean.length > 140 ? `${clean.slice(0, 137)}...` : clean;
};

const mostRecentSubstantiveMessagingContext = (messages = []) => {
  const ignored = /^(hi|hello|hey|thanks|thank you|okay|ok|yes|no|tomorrow|today|perfect|great|set as a reminder|remind me|put this in the things to do tomorrow|put that in the things to do tomorrow)[.!?\s]*$/i;

  const candidates = (Array.isArray(messages) ? messages : [])
    .map((item) => String(item || '').replace(/\s+/g, ' ').trim())
    .filter(Boolean)
    .filter((item) => !ignored.test(item));

  return candidates.at(-1) || '';
};

const inferTaskFromReferentialMessagingCommand = async ({
  cleanText,
  channel,
  recentMessages,
}) => {
  const referential =
    /\b(this|that|it)\b/i.test(cleanText) &&
    /\b(remind|reminder|to[- ]?do|things? to do|task|follow up|put|add|schedule)\b/i.test(cleanText);

  if (!referential) return null;

  const prior = mostRecentSubstantiveMessagingContext(recentMessages);
  if (!prior) return null;

  const lower = cleanText.toLowerCase();
  const tomorrow = /\btomorrow\b/.test(lower);
  const today = /\btoday\b/.test(lower);

  let nextFollowUp = '';
  if (tomorrow) nextFollowUp = addDaysToLocalDate(montrealDateParts(), 1);
  if (today) nextFollowUp = montrealDateParts();

  const titlePrompt = `Create a concise executive task title from this prior WhatsApp/SMS message. Preserve the actual subject and outcome. Do not invent facts. Return only the task title, maximum 110 characters.\n\nPRIOR MESSAGE:\n${prior}`;

  let title = '';
  try {
    const response = await callOpenAIResponses(
      {
        model: ACTION_MESSAGE_MODEL,
        instructions:
          'You turn an existing business message into a concise Action Register task title. Preserve the specific subject and required follow-up. Return only the title with no quotes or punctuation wrapper.',
        input: titlePrompt,
      },
      12000
    );
    title = extractOpenAIResponseText(response).replace(/^["']|["']$/g, '').trim();
  } catch (error) {
    console.warn('Referential task title inference warning:', error.message);
  }

  if (!title) title = compactTaskTitleFromMessage(prior);

  const action = await createActionItem({
    title,
    owner: 'London',
    status: 'ACTIVE',
    priority: 'NORMAL',
    nextFollowUp,
    nextAction: prior,
    source: String(channel || 'message').toUpperCase(),
    notes: `Created from conversational context. Ramy said: "${cleanText}". Prior message: "${prior}"`,
  });

  return {
    success: true,
    action,
    reply: `Done. I added “${action.title}”${action.nextFollowUp ? ` for ${action.nextFollowUp}` : ''}.`,
  };
};

const processExecutiveMessagingInstruction = async ({
  text,
  channel = 'sms',
  sender = RAMY_PHONE_NUMBER,
}) => {
  const cleanText = String(text || '').trim();
  const lower = cleanText.toLowerCase();
  if (!cleanText) return { success: false, reply: 'I did not receive a command.' };

  // Preserve short conversational context across SMS/WhatsApp turns so Ramy can
  // give a task naturally over several messages rather than restating it.
  pruneMessagingState();
  const executiveKey = messagingKey(channel, sender);
  const executiveState = messagingState.get(executiveKey) || {
    recentList: [],
    recentMessages: [],
    lastEventId: '',
    updatedAt: Date.now(),
  };
  executiveState.recentMessages = Array.isArray(executiveState.recentMessages)
    ? executiveState.recentMessages
    : [];
  if (executiveState.recentMessages.at(-1) !== cleanText) {
    executiveState.recentMessages.push(cleanText);
    executiveState.recentMessages = executiveState.recentMessages.slice(-6);
  }
  executiveState.updatedAt = Date.now();
  messagingState.set(executiveKey, executiveState);
  persistMessagingState();

  // Natural conversation.
  if (/^(hi|hello|hey|hey london|hi london|hello london)[.!?\s]*$/i.test(cleanText)) {
    return { success: true, reply: 'Hi Ramy. I’m here — what do you need?' };
  }
  if (/^(thanks|thank you|thx|perfect|great)[.!?\s]*$/i.test(cleanText)) {
    return { success: true, reply: 'You’re welcome.' };
  }
  if (/^(are you there|you there|london|london\?)[.!?\s]*$/i.test(cleanText)) {
    return { success: true, reply: 'Yes, I’m here.' };
  }

  // Read-only email commands.
  if (
    /(check|show|read|latest|recent|what).*\b(email|emails|inbox|mail)\b/i.test(cleanText) ||
    /\b(email|emails|inbox)\b.*\b(latest|recent|new|today)\b/i.test(cleanText)
  ) {
    try {
      const emails = await getRecentMinacoEmails(5);
      return {
        success: true,
        reply: `Latest Minaco emails:\n${formatMessagingEmailList(emails)}`,
      };
    } catch (error) {
      return { success: false, reply: `I could not read your inbox: ${error.message}` };
    }
  }

  // Read-only calendar commands.
  if (
    /\b(calendar|schedule|meetings|appointments)\b/i.test(cleanText) &&
    /\b(check|show|what|today|tomorrow|week|upcoming|next)\b/i.test(cleanText)
  ) {
    try {
      const events = await getCalendarEvents();
      return {
        success: true,
        reply: `Upcoming Minaco calendar:\n${formatMessagingCalendarList(events)}`,
      };
    } catch (error) {
      return { success: false, reply: `I could not read your calendar: ${error.message}` };
    }
  }

  // Referential task commands such as "put this in tomorrow's to-do list"
  // should resolve the subject from the recent conversation automatically.
  const inferredTask = await inferTaskFromReferentialMessagingCommand({
    cleanText,
    channel,
    recentMessages: executiveState.recentMessages.slice(0, -1),
  });
  if (inferredTask?.success) return inferredTask;

  // Action Register commands remain the primary lightweight update path.
  const actionResult = await processQuickActionInstruction({
    text: cleanText,
    channel,
    sender,
  });
  if (actionResult?.success) return actionResult;

  // For an ordinary executive instruction that is not safely executable through
  // messaging yet, capture it as an Action Register task instead of rejecting it.
  // External sends and calendar writes still require a separate confirmation path.
  const looksLikeInstruction =
    /^(please\s+)?(email|reply|respond|send|call|follow|review|analyze|analyse|check|book|schedule|cancel|reschedule|prepare|draft|contact|ask|remind|find|compare|update)\b/i.test(cleanText);

  if (looksLikeInstruction) {
    try {
      const action = await createActionItem({
        title: compactTaskTitleFromMessage(cleanText),
        owner: 'London',
        status: 'ACTIVE',
        priority: 'NORMAL',
        nextAction: cleanText,
        source: String(channel || 'message').toUpperCase(),
        notes: `Command received from Ramy via ${channel}: ${cleanText}`,
      });
      return {
        success: true,
        reply: `Got it. I registered this task: ${action.title}. I’ll keep it in the Action Register. External sends or calendar changes still require confirmation before I execute them.`,
      };
    } catch (error) {
      return { success: false, reply: `I received the command but could not register it: ${error.message}` };
    }
  }

  return {
    success: false,
    reply: 'I received that. I may be missing part of the instruction from the conversation. Tell me the missing detail and I’ll continue from what you already sent.',
  };
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
// Automatic London Task Inbox
// -----------------------------------------------------------------------------

const taskInboxJobs = new Map();
const processedTaskInboxKeys = new Map();

const normalizeTaskSender = (value) =>
  String(value || '')
    .trim()
    .replace(/^mailto:/i, '')
    .toLowerCase();

const pruneTaskInboxState = () => {
  const cutoff = Date.now() - 7 * 24 * 60 * 60 * 1000;

  for (const [key, value] of processedTaskInboxKeys.entries()) {
    if (Number(value?.createdAt || 0) < cutoff) {
      processedTaskInboxKeys.delete(key);
    }
  }

  if (taskInboxJobs.size > 100) {
    const entries = [...taskInboxJobs.entries()].sort(
      (a, b) => Number(a[1]?.createdAt || 0) - Number(b[1]?.createdAt || 0)
    );
    for (const [jobId] of entries.slice(0, taskInboxJobs.size - 100)) {
      taskInboxJobs.delete(jobId);
    }
  }
};

const taskSubjectWithoutPrefixes = (subject) =>
  String(subject || 'Untitled task')
    .replace(/^(re|fw|fwd)\s*:\s*/i, '')
    .trim() || 'Untitled task';

const findLondonTaskInboxEmail = async ({
  messageId,
  subject,
}) => {
  if (!LONDON_MINACO_EMAIL) {
    throw new Error('LONDON_MINACO_EMAIL is not configured.');
  }
  if (!RAMY_MINACO_EMAIL) {
    throw new Error('RAMY_MINACO_EMAIL is not configured.');
  }

  const expectedSender = normalizeTaskSender(RAMY_MINACO_EMAIL);
  const wantedSubject = String(subject || '').trim().toLowerCase();

  if (messageId) {
    try {
      const direct = await getFullMailboxEmail(LONDON_MINACO_EMAIL, messageId);
      const directSender = normalizeTaskSender(
        direct.from?.emailAddress?.address
      );
      if (directSender !== expectedSender) {
        throw new Error(
          `The task email sender is not authorized: ${
            direct.from?.emailAddress?.address || 'unknown sender'
          }.`
        );
      }
      return direct;
    } catch (error) {
      console.warn(
        'TASK INBOX direct message-id lookup warning:',
        error.message
      );
    }
  }

  const recent = await getRecentInboxEmails({
    mailboxAddress: LONDON_MINACO_EMAIL,
    limit: 100,
  });

  const authorized = recent.filter(
    (email) =>
      normalizeTaskSender(email.from?.emailAddress?.address) === expectedSender
  );

  const exactSubjectMatches = wantedSubject
    ? authorized.filter(
        (email) =>
          String(email.subject || '').trim().toLowerCase() === wantedSubject
      )
    : [];

  const candidate =
    exactSubjectMatches[0] ||
    authorized[0];

  if (!candidate?.id) {
    throw new Error(
      'London could not locate a recent task email from Ramy in the London inbox.'
    );
  }

  return getFullMailboxEmail(LONDON_MINACO_EMAIL, candidate.id);
};

const buildTaskInboxFileInputs = async ({
  messageId,
  attachments,
}) => {
  const content = [];
  const reviewedFiles = [];
  const skippedFiles = [];
  let totalBytes = 0;

  const usableAttachments = (Array.isArray(attachments) ? attachments : [])
    .filter((attachment) => !attachment.isInline)
    .slice(0, MAX_TASK_ATTACHMENTS);

  for (const meta of usableAttachments) {
    if (!meta?.attachmentId) continue;

    const size = Number(meta.size) || 0;
    if (size > MAX_ATTACHMENT_BYTES) {
      skippedFiles.push(
        `${meta.name || 'Unnamed attachment'} — larger than the 45 MB single-file limit`
      );
      continue;
    }
    if (totalBytes + size > MAX_TASK_TOTAL_BYTES) {
      skippedFiles.push(
        `${meta.name || 'Unnamed attachment'} — skipped because the task attachments exceeded the 45 MB combined analysis limit`
      );
      continue;
    }

    const attachment = await getMailboxEmailAttachment(
      LONDON_MINACO_EMAIL,
      messageId,
      meta.attachmentId
    );

    if (!attachment?.contentBytes) {
      skippedFiles.push(
        `${attachment?.name || meta.name || 'Unnamed attachment'} — file bytes were not available`
      );
      continue;
    }

    const filename = attachment.name || meta.name || 'attachment';
    const contentType =
      attachment.contentType || meta.contentType || 'application/octet-stream';
    const dataUri = `data:${contentType};base64,${attachment.contentBytes}`;

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

    totalBytes += Number(attachment.size) || size || 0;
    reviewedFiles.push(filename);
  }

  return {
    content,
    reviewedFiles,
    skippedFiles,
    totalBytes,
  };
};

const normalizeTaskReport = (value = {}) => {
  const array = (candidate) => (Array.isArray(candidate) ? candidate : []);
  const object = (candidate) =>
    candidate && typeof candidate === 'object' && !Array.isArray(candidate)
      ? candidate
      : {};

  const table = object(value.comparisonTable);
  const draft = object(value.draftResponse);

  return {
    taskTitle: String(value.taskTitle || '').trim(),
    executiveConclusion: String(value.executiveConclusion || '').trim(),
    status: String(value.status || 'Complete').trim(),
    keyFindings: array(value.keyFindings),
    comparisonTable: {
      columns: array(table.columns).map((item) => String(item || '').trim()),
      rows: array(table.rows).map((row) =>
        array(row).map((cell) => String(cell ?? '').trim())
      ),
    },
    risks: array(value.risks),
    missingInformation: array(value.missingInformation).map((item) =>
      String(item || '').trim()
    ),
    recommendedNextActions: array(value.recommendedNextActions).map((item) =>
      String(item || '').trim()
    ),
    draftResponse: {
      included: Boolean(draft.included),
      recipient: String(draft.recipient || '').trim(),
      subject: String(draft.subject || '').trim(),
      body: String(draft.body || '').trim(),
    },
    ramyActionRequired: Boolean(value.ramyActionRequired),
    ramyNextAction: String(value.ramyNextAction || '').trim(),
    limitations: String(value.limitations || '').trim(),
  };
};

const renderTaskInboxReportHtml = ({
  report,
  taskSubject,
  reviewedFiles,
  skippedFiles,
  receivedAt,
}) => {
  const section = (title, color, body) => `
    <tr><td style="padding:0 24px 18px;">
      <div style="padding-bottom:8px;border-bottom:2px solid ${color};font-size:13px;font-weight:800;letter-spacing:.45px;color:${color};">${escapeDelegatedHtml(title)}</div>
      <div style="padding-top:9px;font-size:13px;line-height:20px;color:#344054;">${body}</div>
    </td></tr>`;

  const list = (items, emptyText = 'None identified.') => {
    const normalized = (Array.isArray(items) ? items : [])
      .map((item) => String(item || '').trim())
      .filter(Boolean);
    if (!normalized.length) {
      return `<div style="color:#667085;">${escapeDelegatedHtml(emptyText)}</div>`;
    }
    return `<ul style="margin:4px 0 0 18px;padding:0;">${normalized
      .map(
        (item) =>
          `<li style="margin:0 0 7px 0;line-height:20px;">${escapeDelegatedHtml(item)}</li>`
      )
      .join('')}</ul>`;
  };

  const findings = (Array.isArray(report.keyFindings)
    ? report.keyFindings
    : []
  )
    .map((item) => {
      if (typeof item === 'string') {
        return `<div style="border:1px solid #eaecf0;border-radius:7px;padding:11px 13px;margin:8px 0;background:#fff;">${escapeDelegatedHtml(item)}</div>`;
      }
      const importance = String(item?.importance || 'Normal').toUpperCase();
      const color =
        importance === 'HIGH'
          ? '#B42318'
          : importance === 'MEDIUM'
          ? '#B54708'
          : '#175CD3';
      return `<div style="border:1px solid #eaecf0;border-left:4px solid ${color};border-radius:7px;padding:11px 13px;margin:8px 0;background:#fff;">
        <div style="font-weight:700;color:#101828;">${escapeDelegatedHtml(
          item?.title || 'Finding'
        )}</div>
        <div style="margin-top:4px;color:#475467;line-height:20px;">${escapeDelegatedHtml(
          item?.detail || ''
        )}</div>
      </div>`;
    })
    .join('');

  const risks = (Array.isArray(report.risks) ? report.risks : [])
    .map((risk) => {
      const level = String(risk?.level || 'Review').toUpperCase();
      const color =
        level === 'HIGH'
          ? '#B42318'
          : level === 'MEDIUM'
          ? '#B54708'
          : '#175CD3';
      return `<div style="border:1px solid #eaecf0;border-left:4px solid ${color};border-radius:7px;padding:11px 13px;margin:8px 0;background:#fff;">
        <div style="font-weight:700;color:#101828;">${escapeDelegatedHtml(
          risk?.title || 'Risk'
        )}</div>
        <div style="margin-top:4px;color:#475467;line-height:20px;">${escapeDelegatedHtml(
          risk?.detail || ''
        )}</div>
      </div>`;
    })
    .join('');

  let tableHtml = '';
  const columns = report.comparisonTable?.columns || [];
  const rows = report.comparisonTable?.rows || [];
  if (columns.length && rows.length) {
    const header = columns
      .map(
        (column) =>
          `<th align="left" style="padding:8px 9px;background:#102A43;color:#fff;font-size:11px;line-height:16px;border:1px solid #344054;">${escapeDelegatedHtml(
            column
          )}</th>`
      )
      .join('');
    const body = rows
      .map(
        (row) =>
          `<tr>${columns
            .map(
              (_, index) =>
                `<td valign="top" style="padding:8px 9px;font-size:11px;line-height:16px;color:#344054;border:1px solid #eaecf0;background:#fff;">${escapeDelegatedHtml(
                  row[index] ?? ''
                )}</td>`
            )
            .join('')}</tr>`
      )
      .join('');
    tableHtml = `<div style="overflow-x:auto;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;"><thead><tr>${header}</tr></thead><tbody>${body}</tbody></table></div>`;
  }

  const draftHtml = report.draftResponse?.included
    ? `<div style="background:#fff;border:1px solid #d0d5dd;border-radius:7px;padding:12px 14px;">
        ${
          report.draftResponse.recipient
            ? `<div><b>Proposed recipient:</b> ${escapeDelegatedHtml(
                report.draftResponse.recipient
              )}</div>`
            : ''
        }
        ${
          report.draftResponse.subject
            ? `<div style="margin-top:4px;"><b>Subject:</b> ${escapeDelegatedHtml(
                report.draftResponse.subject
              )}</div>`
            : ''
        }
        <div style="margin-top:9px;white-space:pre-wrap;">${escapeDelegatedHtml(
          report.draftResponse.body
        )}</div>
        <div style="margin-top:9px;color:#B42318;font-weight:700;">Draft only — nothing has been sent externally.</div>
      </div>`
    : '';

  const fileLines = [
    ...(reviewedFiles || []).map((name) => `Reviewed: ${name}`),
    ...(skippedFiles || []).map((name) => `Not reviewed: ${name}`),
  ];

  return `<!doctype html><html><body style="margin:0;background:#f2f4f7;font-family:Arial,Helvetica,sans-serif;color:#101828;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td align="center" style="padding:24px 10px;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:900px;background:#f8fafc;border:1px solid #e4e7ec;border-radius:12px;overflow:hidden;">
        <tr><td style="background:#102A43;padding:22px 24px;border-bottom:4px solid #C8A96B;">
          <div style="font-size:11px;font-weight:700;letter-spacing:1.8px;color:#C8A96B;">LONDON ASSISTANT · MINACO</div>
          <div style="margin-top:5px;font-size:23px;font-weight:800;color:#fff;">Task Analysis Complete</div>
          <div style="margin-top:5px;font-size:13px;color:#d0d5dd;">${escapeDelegatedHtml(
            report.taskTitle || taskSubject || 'Task'
          )}</div>
          ${
            receivedAt
              ? `<div style="margin-top:3px;font-size:11px;color:#98A2B3;">Task received: ${escapeDelegatedHtml(
                  formatMontrealDateTime(new Date(receivedAt))
                )}</div>`
              : ''
          }
        </td></tr>
        <tr><td style="padding:18px 24px;">
          <div style="background:#fff;border:1px solid #d0d5dd;border-radius:8px;padding:14px 16px;">
            <div style="font-size:11px;font-weight:800;letter-spacing:.5px;color:#667085;">EXECUTIVE CONCLUSION</div>
            <div style="margin-top:4px;font-size:14px;line-height:21px;font-weight:700;color:#101828;">${escapeDelegatedHtml(
              report.executiveConclusion || 'See detailed analysis below.'
            )}</div>
          </div>
        </td></tr>
        ${section(
          'KEY FINDINGS',
          '#175CD3',
          findings || '<div style="color:#667085;">No material finding was identified.</div>'
        )}
        ${
          tableHtml
            ? section('DETAILED COMPARISON / ANALYSIS', '#344054', tableHtml)
            : ''
        }
        ${section(
          'RISKS / ISSUES TO PROTECT',
          '#B42318',
          risks || '<div style="color:#667085;">No specific risk was identified from the supplied material.</div>'
        )}
        ${section(
          'MISSING / UNVERIFIED INFORMATION',
          '#B54708',
          list(report.missingInformation, 'No material information gap was identified.')
        )}
        ${section(
          'RECOMMENDED NEXT ACTIONS',
          '#027A48',
          list(report.recommendedNextActions, 'No further action is recommended from the supplied material.')
        )}
        ${
          report.ramyActionRequired
            ? section(
                'RAMY ACTION REQUIRED',
                '#6941C6',
                `<div style="font-weight:700;">${escapeDelegatedHtml(
                  report.ramyNextAction || 'Review the analysis and decide the next step.'
                )}</div>`
              )
            : ''
        }
        ${
          draftHtml
            ? section('DRAFT RESPONSE — FOR RAMY REVIEW ONLY', '#6941C6', draftHtml)
            : ''
        }
        ${
          report.limitations
            ? section(
                'LIMITATIONS',
                '#667085',
                `<div>${escapeDelegatedHtml(report.limitations)}</div>`
              )
            : ''
        }
        ${section(
          'SOURCE FILES',
          '#667085',
          list(fileLines, 'No file attachment was included with this task.')
        )}
      </table>
    </td></tr></table>
  </body></html>`;
};

const sendTaskInboxAcknowledgement = async ({
  taskSubject,
  attachmentNames,
  jobId,
}) => {
  const files = (attachmentNames || []).filter(Boolean);
  const fileText = files.length
    ? `<div style="margin-top:8px;color:#475467;font-size:13px;"><b>Files received:</b> ${files
        .map(escapeDelegatedHtml)
        .join(', ')}</div>`
    : '<div style="margin-top:8px;color:#667085;font-size:13px;">No file attachment was detected.</div>';

  const html = `<!doctype html><html><body style="margin:0;background:#f2f4f7;font-family:Arial,Helvetica,sans-serif;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td align="center" style="padding:22px 10px;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:720px;background:#fff;border:1px solid #e4e7ec;border-radius:10px;overflow:hidden;">
      <tr><td style="background:#102A43;padding:18px 20px;border-bottom:4px solid #C8A96B;">
        <div style="font-size:11px;font-weight:700;letter-spacing:1.6px;color:#C8A96B;">LONDON ASSISTANT · MINACO</div>
        <div style="margin-top:4px;font-size:20px;font-weight:800;color:#fff;">Task Received</div>
      </td></tr>
      <tr><td style="padding:18px 20px;color:#344054;font-size:13px;line-height:20px;">
        <div><b>${escapeDelegatedHtml(taskSubject || 'Task')}</b></div>
        <div style="margin-top:7px;">I received your task and have started the analysis. I will email you the completed review when it is finished.</div>
        ${fileText}
        <div style="margin-top:10px;color:#667085;">No external email or commitment will be sent from this task unless you separately authorize it.</div>
      </td></tr>
    </table>
  </td></tr></table>
  </body></html>`;

  await sendEmailFromLondon({
    to: RAMY_MINACO_EMAIL,
    subject: `LONDON — Task Received | ${taskSubjectWithoutPrefixes(taskSubject)}`,
    body: html,
    contentType: 'HTML',
  });

  console.log('TASK INBOX ACKNOWLEDGED:', { jobId, taskSubject });
};

const runTaskInboxJob = async ({
  jobId,
  messageId,
  subject,
}) => {
  const job = taskInboxJobs.get(jobId);
  let actionEventId = null;
  let actualSubject = taskSubjectWithoutPrefixes(subject);

  try {
    if (job) {
      job.status = 'running';
      job.startedAt = Date.now();
    }

    const taskEmail = await findLondonTaskInboxEmail({
      messageId,
      subject,
    });

    const actualSender = normalizeTaskSender(
      taskEmail.from?.emailAddress?.address
    );
    if (actualSender !== normalizeTaskSender(RAMY_MINACO_EMAIL)) {
      throw new Error('The task email was not sent by the authorized Ramy Minaco mailbox.');
    }

    actualSubject = taskSubjectWithoutPrefixes(taskEmail.subject || subject);
    const taskInstruction = String(taskEmail.body?.content || '').trim();

    const attachments = taskEmail.hasAttachments
      ? await listMailboxEmailAttachments(LONDON_MINACO_EMAIL, taskEmail.id)
      : [];
    const visibleAttachments = attachments.filter((item) => !item.isInline);

    try {
      const action = await createActionItem({
        title: `Task Inbox — ${actualSubject}`,
        project: '',
        category: 'Executive Assistant Task',
        owner: 'London',
        status: 'ACTIVE',
        priority: 'NORMAL',
        lastContact: montrealDateParts(),
        nextAction: 'Analyze the task email and attachments, then report back to Ramy.',
        waitingOn: 'London',
        ramyRequired: false,
        source: `London Task Inbox | ${taskEmail.internetMessageId || taskEmail.id}`,
        notes: `Automatically triggered from an email sent by Ramy to ${LONDON_MINACO_EMAIL}.`,
      });
      actionEventId = action.eventId;
      if (job) job.actionEventId = actionEventId;
    } catch (error) {
      console.warn('TASK INBOX action-register warning:', error.message);
    }

    await sendTaskInboxAcknowledgement({
      taskSubject: actualSubject,
      attachmentNames: visibleAttachments.map((item) => item.name),
      jobId,
    });

    const fileInputs = await buildTaskInboxFileInputs({
      messageId: taskEmail.id,
      attachments: visibleAttachments,
    });

    const userContent = [...fileInputs.content];
    userContent.push({
      type: 'input_text',
      text: `RAMY'S TASK EMAIL
Subject: ${actualSubject}
From: ${taskEmail.from?.emailAddress?.name || ''} <${taskEmail.from?.emailAddress?.address || ''}>
Received: ${taskEmail.receivedDateTime || ''}

TASK INSTRUCTIONS / EMAIL BODY:
${taskInstruction || '[No task instruction text was found in the email body.]'}

FILES PROVIDED:
${fileInputs.reviewedFiles.length ? fileInputs.reviewedFiles.join('\n') : '[No analyzable file attachments were included.]'}

FILES NOT ANALYZED:
${fileInputs.skippedFiles.length ? fileInputs.skippedFiles.join('\n') : '[None]'}`,
    });

    const data = await callOpenAIResponses(
      {
        model: TASK_INBOX_ANALYSIS_MODEL,
        instructions: `You are London Assistant performing an autonomous internal task assigned directly by Ramy Mina through the verified London Task Inbox.

Use ONLY the task email and supplied attachments. Analyze the actual files, including every relevant Excel worksheet/tab when the task asks for spreadsheet review. Do not silently skip worksheets, rows, units, calculations, assumptions, dates, or material discrepancies requested by Ramy.

Be objective. Do not manufacture a challenge or conclusion. If the source material supports the counterparty, say so. If the evidence is incomplete, identify exactly what is missing rather than guessing.

Important authority rules:
- This is an INTERNAL analysis for Ramy.
- Never contact an external party.
- Never send a reply to a lender, supplier, tenant, consultant, or other third party from this task.
- If Ramy requested a response, prepare it only as a DRAFT for his review.
- Never approve payments, sign, commit pricing, settle disputes, or make legal/financial commitments.

For comparison or compliance work, independently reconstruct the calculations from source files before comparing them with another party's findings.
For Excel tasks, use sheet/tab names and source values where material.
If Ramy requests a line-by-line, unit-by-unit, or tab-by-tab review, include all requested material rows in the comparison table rather than only examples.

Return ONLY valid JSON, with no Markdown or code fences, using exactly this shape:
{
  "taskTitle": "short descriptive task title",
  "executiveConclusion": "clear overall conclusion",
  "status": "Complete or Needs information",
  "keyFindings": [
    {"importance":"High|Medium|Low","title":"short finding","detail":"supported detail"}
  ],
  "comparisonTable": {
    "columns": ["column 1","column 2","..."],
    "rows": [["cell","cell","..."]]
  },
  "risks": [
    {"level":"High|Medium|Low","title":"risk title","detail":"why it matters"}
  ],
  "missingInformation": ["specific missing or unverifiable item"],
  "recommendedNextActions": ["specific recommended action"],
  "draftResponse": {
    "included": false,
    "recipient": "",
    "subject": "",
    "body": ""
  },
  "ramyActionRequired": false,
  "ramyNextAction": "",
  "limitations": "what cannot be concluded from the supplied material"
}`,
        input: [
          {
            role: 'user',
            content: userContent,
          },
        ],
      },
      120000
    );

    const raw = extractOpenAIResponseText(data);
    if (!raw) {
      throw new Error('The Task Inbox analysis returned no text.');
    }

    let parsed;
    try {
      parsed = JSON.parse(stripJsonCodeFence(raw));
    } catch (error) {
      throw new Error(
        'The Task Inbox analysis could not be parsed into a structured report.'
      );
    }

    const report = normalizeTaskReport(parsed);
    const html = renderTaskInboxReportHtml({
      report,
      taskSubject: actualSubject,
      reviewedFiles: fileInputs.reviewedFiles,
      skippedFiles: fileInputs.skippedFiles,
      receivedAt: taskEmail.receivedDateTime,
    });

    await sendEmailFromLondon({
      to: RAMY_MINACO_EMAIL,
      subject: `LONDON — Task Complete | ${actualSubject}`,
      body: html,
      contentType: 'HTML',
    });

    if (actionEventId) {
      try {
        await updateActionItem(actionEventId, {
          status: report.ramyActionRequired ? 'WAITING - RAMY' : 'COMPLETED',
          lastContact: montrealDateParts(),
          nextAction: report.ramyActionRequired
            ? report.ramyNextAction || 'Review London’s completed analysis and decide the next step.'
            : 'Completed analysis emailed to Ramy.',
          waitingOn: report.ramyActionRequired ? 'Ramy' : '',
          ramyRequired: Boolean(report.ramyActionRequired),
          notes: `Task Inbox analysis completed. Files reviewed: ${
            fileInputs.reviewedFiles.join(', ') || 'none'
          }. ${
            fileInputs.skippedFiles.length
              ? `Files not analyzed: ${fileInputs.skippedFiles.join(', ')}.`
              : ''
          }`,
        });
      } catch (error) {
        console.warn('TASK INBOX completion action-register warning:', error.message);
      }
    }

    if (job) {
      job.status = 'completed';
      job.completedAt = Date.now();
      job.taskSubject = actualSubject;
      job.reportSentTo = RAMY_MINACO_EMAIL;
    }

    console.log('TASK INBOX COMPLETED:', {
      jobId,
      taskSubject: actualSubject,
      files: fileInputs.reviewedFiles,
    });
  } catch (error) {
    console.error('TASK INBOX FAILED:', {
      jobId,
      taskSubject: actualSubject,
      error: error.message,
    });

    if (actionEventId) {
      try {
        await updateActionItem(actionEventId, {
          status: 'BLOCKED',
          lastContact: montrealDateParts(),
          nextAction: 'Review the Task Inbox failure and provide missing information or retry.',
          waitingOn: 'Ramy',
          ramyRequired: true,
          riskIfDelayed: 'The delegated task did not complete.',
          notes: `Task Inbox failure: ${error.message}`,
        });
      } catch (updateError) {
        console.error(
          'TASK INBOX failure action-register update also failed:',
          updateError
        );
      }
    }

    try {
      await sendEmailFromLondon({
        to: RAMY_MINACO_EMAIL,
        subject: `LONDON — Task Could Not Complete | ${actualSubject}`,
        body: `London received the task but could not complete it.\n\nTechnical issue: ${error.message}\n\nNo external email or commitment was sent.`,
        contentType: 'Text',
      });
    } catch (notifyError) {
      console.error('TASK INBOX failure notification also failed:', notifyError);
    }

    if (job) {
      job.status = 'failed';
      job.completedAt = Date.now();
      job.error = error.message;
    }
  } finally {
    pruneTaskInboxState();
  }
};

const queueTaskInboxJob = ({
  messageId,
  subject,
  dedupeKey,
}) => {
  const jobId = `task-inbox-${Date.now()}-${Math.random()
    .toString(36)
    .slice(2, 8)}`;

  taskInboxJobs.set(jobId, {
    jobId,
    status: 'queued',
    createdAt: Date.now(),
    messageId: String(messageId || '').trim(),
    taskSubject: taskSubjectWithoutPrefixes(subject),
  });

  if (dedupeKey) {
    processedTaskInboxKeys.set(dedupeKey, {
      jobId,
      createdAt: Date.now(),
    });
  }

  setImmediate(() => {
    runTaskInboxJob({
      jobId,
      messageId,
      subject,
    }).catch((error) => {
      console.error('Unexpected Task Inbox background failure:', error);
    });
  });

  return {
    success: true,
    queued: true,
    jobId,
    taskSubject: taskSubjectWithoutPrefixes(subject),
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

VOICE ALIGNMENT

Use Vale's voice identity. Keep Vale recognizably Vale at all times.
Speak in polished British English with a natural England accent.
Make Vale sound warmer, fuller, deeper-feeling, calmer, and more mature. Favor the lowest comfortable natural register available to Vale, with richer chest-like resonance and less brightness, while keeping the voice clearly feminine and recognizably Vale. Do not sound artificial, theatrical, breathy, sleepy, or masculine.
Reduce brightness and sharpness substantially. Avoid a high, perky, overly cheerful, thin, or sales-like tone.
Use a composed executive-assistant presence: grounded, reassuring, discreet, confident, and professional.
Respond promptly once Ramy's thought is clearly complete. Do not add unnecessary pre-response silence.
Speak at a natural, calm pace—only slightly slower than normal conversation—with short natural pauses rather than long pauses between thoughts.
Never rush names, dates, numbers, email addresses, or meeting times.
Keep answers concise and avoid long monologues. Prefer one or two short thoughts at a time.
Ramy may pause briefly while forming a sentence. Allow a short hesitation, but once his thought is clearly complete, respond quickly. Do not leave a long silent gap before answering.
If Ramy begins speaking while you are talking, stop promptly and listen. Do not continue over him.
When guiding him through a technical setup, give one short step at a time and wait for his reply before giving the next step.
Ramy's first name is spelled RAMY with a Y. When writing his name or signing an email, always use "Ramy", never "Rami".

ACCURACY IS YOUR HIGHEST PRIORITY

Never invent or guess facts about Ramy, Minaco, Mina Capital, Mina Group, employees, consultants, contractors, lawyers, partners, investors, lenders, tenants, projects, properties, emails, meetings, financial information, contracts, deadlines, or business relationships.

If information is not explicitly provided in verified context or retrieved from an authorized live system, say that you do not have verified information.
Never make up a plausible answer just to be helpful.

LIVE EMAIL RULES

Use check_email whenever Ramy asks about current emails or his inbox. Never answer a current email question from memory.
check_email returns a recent-email list and message ids. It does NOT contain the entire body.
Use read_email whenever Ramy asks you to read, translate, summarize, analyze, or respond based on the complete contents of a specific email. read_email retrieves the full live message body in text form.
If the target email has not yet been identified, use check_email first, identify the exact email, then use read_email with its message id. Never translate or analyze an email from a preview when the full body is available.

For NEW standalone email, respect sender identity exactly. If Ramy says "from my email", "send it from me", or otherwise explicitly identifies his mailbox, prepare it from ramy.mina@minaco.ca. If he says "from London" or explicitly identifies London, prepare it from london@minaco.ca. Never substitute one sender for the other. If no sender identity is specified for a new external email, default to Ramy's mailbox because London is acting as his executive assistant. Preparing an email NEVER sends it.
After prepare_email succeeds, read back the recipient, subject, and message and clearly say it has not been sent. Ask Ramy to confirm.

For a REPLY to an email in Ramy's Minaco inbox:
1. Identify the exact email and message id with check_email and, when the reply depends on its contents, read_email.
2. If Ramy says reply, respond, or answer without specifying the recipient scope, ask one short question: "Reply to sender only or reply all?"
3. Use prepare_email_reply with reply_mode "sender" for sender-only or "all" for reply-all.
4. Replies are sent from ramy.mina@minaco.ca and remain in the original Outlook conversation thread.
5. Read back whether it is sender-only or reply-all, the recipients, subject, and complete reply text. Clearly say it has NOT been sent yet.
6. Ask Ramy to confirm by saying "Send it."

Use send_confirmed_email ONLY after Ramy explicitly confirms the currently pending NEW email or REPLY with a direct send command such as "Send it", "Yes, send it", "Send the email", or "Go ahead and send it" in a NEW user turn after you read the draft back.
IMPORTANT: "Thank you", "thanks", "okay", "perfect", "sounds good", "yes", silence, or moving to another topic are NOT permission to send. Never interpret politeness or acknowledgment as authorization. The server independently validates the user's spoken confirmation and will refuse the send tool without an explicit send phrase.
Never claim an email or reply was sent unless the send tool confirms success.
Never invent a recipient email address. If Ramy gives only a person's name for a new email and you do not have a verified email address, ask for the email address.


ADVANCED EMAIL, ATTACHMENTS, AND CONTACTS

LONDON TASK INBOX

Emails that Ramy sends directly to london@minaco.ca can be treated as task assignments by the automatic Task Inbox workflow. That background workflow is separate from the live phone call: Power Automate triggers the server, the server verifies the sender is ramy.mina@minaco.ca, acknowledges the task, analyzes the complete task email and supported attachments, records the work in the Action Register, and emails the completed internal analysis back to Ramy.
If Ramy asks whether a task email reached London, use search_email with mailbox "london" to verify the live London inbox. Do not claim a task has started or completed unless the live mailbox or task workflow confirms it.
Task Inbox work may prepare a draft response for Ramy's review, but it must never send externally or make a commitment on its own.


DELEGATED EMAIL REVIEW / BACKGROUND WORK

For a multi-step request where Ramy asks you to read or review an email thread, summarize/analyze/translate it, give your view or identify risks, AND email the result to Ramy, use delegate_email_thread_review immediately instead of chaining several synchronous voice tools.
This delegated tool may send the finished analysis ONLY to Ramy's own verified Minaco email address. It can never email an external recipient. Ramy's explicit request in the same turn to "send me an email", "email me the summary", or equivalent is authorization for this self-email only; do not ask for a second Send it confirmation for the delegated self-email.
Once delegate_email_thread_review confirms the job is queued, tell Ramy briefly that the task has been handed off and he may hang up; the completed review will be emailed to him. Do NOT keep him waiting on the call and do NOT continue a synchronous search/read/analyze chain for the same task.
If no delegated job has actually been queued, never claim you will continue working after the call ends.
External replies or emails to anyone other than Ramy still require the normal prepared-draft and explicit Send it confirmation workflow.

DELAY / RESPONSIVENESS

Never leave Ramy in unexplained silence while you are retrieving or processing information.
Before starting any lookup, tool call, or multi-step operation that may take more than about two seconds, immediately say one short progress phrase such as:
- "Okay, hang on — I'm checking that now."
- "I'm gathering that for you now."
- "Give me a moment — I'm checking the live system."
Keep the progress phrase brief, then perform the lookup. Do not give a long explanation before working.
If the work continues for an unusually long time, give Ramy a brief status update rather than remaining silent, but do not repeatedly interrupt him with unnecessary updates.
Once the result is ready, give the answer directly without repeating the progress message.
If Ramy asks why you were slow to respond, never invent a reason such as a flight delay or another event. Say only that there was a connection or processing delay unless a live tool or system error provides a verified cause.
When a delegated job has been queued, remain responsive to Ramy immediately; the long analysis is no longer part of the live voice turn.

Use search_email when Ramy asks you to find an older email, search by person/company/topic, or search a specific period. Search can target Ramy's mailbox or Accounting.
Use check_accounting when Ramy asks about current invoices, statements, payment reminders, deposits, supplier credits, overdue notices, insurance, taxes, financing charges, or other accounting-mailbox items.
Use read_accounting_email to retrieve the complete Accounting email body before translating, analyzing, or making conclusions from it.

When an email has attachments, use list_email_attachments to identify the exact attachment. Use analyze_email_attachment when Ramy asks to open, summarize, translate, analyze, extract amounts/deadlines, or answer questions about an attachment. The attachment tool analyzes the actual file; never pretend you opened an attachment when you only saw its filename.

DROPBOX DOCUMENT WORKSPACE

London's controlled Dropbox workspace is LONDON - ACCESS. Use list_dropbox, search_dropbox, read_dropbox_text, and delegate_dropbox_file_analysis only for content inside that root. Never access or claim access to anything outside that folder. Dropbox is read-only through these tools.
When Ramy asks what is inside a folder, inspect that folder yourself. If he asks for subfolders, sub-subfolders, deeper structure, everything under a folder, or says "find it on your own", use list_dropbox with recursive=true and inspect the returned descendants. Do not ask him which child folder to inspect when his request is to discover the deeper structure autonomously.
When Ramy names a folder or file approximately, use search_dropbox to resolve the real path rather than asking him to restate the exact filename.
For PDF, Word, Excel, PowerPoint, image, or other supported binary documents already in Dropbox, do NOT tell Ramy to email the file to the Task Inbox. Use delegate_dropbox_file_analysis. If he asks for analysis and says to email him the result, queue the delegated Dropbox analysis immediately; it may email the completed analysis only to Ramy's verified Minaco email. Tell him briefly that the analysis is queued and he may continue with another task or hang up.

Use resolve_person when Ramy gives a person's name but not an email address. It searches verified identities and Minaco email history. If one clear candidate is returned, use that verified address. If multiple plausible candidates are returned, ask Ramy which one he means. Never invent an address.

MASTER ACTION & FOLLOW-UP REGISTER

Use create_action when Ramy explicitly asks you to track, remember, follow up, add an action, or when he clearly instructs you that a business outcome must be monitored. Do not create action items merely because an email contains information.
Use list_actions when Ramy asks what is overdue, what he is waiting for, what needs his decision, what London must follow up on, or for a project/action status.
Use quick_action_update as the FIRST choice for simple natural-language status updates such as “Joannie done”, “waiting on Anass until Friday”, “follow up with Franco next Tuesday”, “add task: call Makar tomorrow”, “cancel the EV follow-up”, or several quick updates in one sentence. The server matches the live Action Register and makes the update in one step, which is faster than chaining list_actions + update_action.
Use update_action when you already have the exact event id or when a precise field edit is needed after listing actions.
The register fields are outcome, project, owner, dates, status, priority, next action, waiting on, Ramy requirement, risk, source, and notes. Preserve the distinction between promised date, hard deadline, and next follow-up.
Ramy may also send the same natural-language updates by SMS or WhatsApp. Those channels feed the same Action Register and acknowledge the resulting update briefly. Do not require special command syntax.

DAILY EXECUTIVE BRIEF

Use daily_executive_brief when Ramy asks for his morning brief, daily brief, what needs his attention today, or an executive overview. It combines live calendar, executive email, Accounting email, and the action register. Never invent missing items.

LIVE CALENDAR RULES

Use check_calendar whenever Ramy asks what is on his calendar, what meetings he has, or asks about a specific date or period.
Use check_availability whenever Ramy asks whether he is free at one specific time.
Use find_availability whenever Ramy asks for several possible times, asks for his availability over a day/week/date range, or asks you to reply to someone with his availability. Do NOT ask Ramy to tell you his availability when his live calendar can answer it.
For relative date phrases such as "next week" or "this week", pass the phrase itself in date_range. Do NOT manually calculate start_date/end_date unless Ramy gave explicit calendar dates. The server resolves relative ranges using the current Montreal date.
If no meeting duration is stated for an availability request, default to 30 minutes. Unless Ramy or the email specifies otherwise, search normal business hours from 9:00 AM to 5:00 PM Montreal time and offer 3 to 5 useful slots. Respect any duration, day, or time constraints stated in the email or by Ramy.

AVAILABILITY TRUTH RULE: Never say you checked Ramy's calendar unless the current calendar tool call returned success. Never invent, recycle, or infer calendar openings after a failed tool call. If a calendar or availability tool fails, say it failed. When proposing availability, use the exact returned slots only. Do not turn separate exact slots into a broad range with contradictory exceptions.
All spoken calendar times are Montreal local time unless Ramy explicitly specifies another time zone.
If Ramy asks whether you have calendar access, answer directly that you have live access to his Minaco calendar through the authorized calendar tools. Do not call a tool merely to explain that access.

CALENDAR-AWARE EMAIL TASKS

FAST PATH FOR NEXT WEEK: If Ramy says something like "respond to Francis and give him my availability next week", use prepare_next_week_availability_reply as the FIRST choice. This dedicated tool calculates the next Monday-Friday on the server in Montreal time, finds the latest actual inbound email from that sender, checks the live calendar, chooses three useful openings across different days when possible, and prepares the reply. It does not depend on you calculating or formatting dates.

For availability periods other than next week, use prepare_availability_reply. Do not chain search_email + read_email + find_availability + prepare_email_reply unless the full email imposes special meeting constraints that the fast path cannot handle.

If the fast path cannot identify the sender or find open slots, ask one concise clarification question. Do not ask Ramy to state his availability when the live calendar can answer it.

For a more complex availability reply where the full email body contains important constraints, complete the workflow yourself using live systems:
1. Identify the exact email. If he names the sender, use that name to identify/search the live email rather than asking Ramy for the email address if the message can be found.
2. Read the full email only when its contents or requested meeting constraints materially matter.
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
Never call confirm_calendar_action unless there is a pending prepared action and Ramy has explicitly confirmed it in a NEW user turn after you read the proposed calendar action back. The server independently validates the spoken confirmation. Polite acknowledgments such as "thank you", "okay", "perfect", or "yes" alone are NOT authorization.
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
When you need to check a live system, tell Ramy immediately that you are checking before any silence begins.
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
  'conversation.item.input_audio_transcription.completed',
  'conversation.item.input_audio_transcription.failed',
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
      'automatic task inbox',
      'dropbox controlled workspace',
      'health monitoring',
    ],
  });
});


fastify.get('/health', async (request, reply) => {
  return reply.send({
    ok: true,
    service: 'London Assistant',
    time: new Date().toISOString(),
    voiceModel: 'gpt-realtime',
    dropboxConfigured: Boolean(DROPBOX_ACCESS_TOKEN),
    stateFileConfigured: Boolean(STATE_FILE),
  });
});

fastify.get('/health/deep', async (request, reply) => {
  if (!HEALTH_SECRET || request.headers['x-london-health-secret'] !== HEALTH_SECRET) {
    return reply.code(401).send({ ok: false, error: 'Unauthorized.' });
  }
  const checks = {};
  const run = async (name, fn) => {
    const started = Date.now();
    try { await fn(); checks[name] = { ok: true, ms: Date.now() - started }; }
    catch (error) { checks[name] = { ok: false, ms: Date.now() - started, error: String(error.message || error).slice(0, 300) }; }
  };
  await Promise.all([
    run('microsoftRead', async () => { await getMicrosoftGraphToken(); }),
    run('microsoftActions', async () => { await getMicrosoftGraphActionsToken(); }),
    run('dropbox', async () => { if (!DROPBOX_ACCESS_TOKEN) throw new Error('Not configured'); await listLondonDropbox(LONDON_DROPBOX_ROOT, false); }),
    run('twilio', async () => {
      if (!TWILIO_ACCOUNT_SID || !TWILIO_AUTH_TOKEN) throw new Error('Not configured');
      const auth = Buffer.from(`${TWILIO_ACCOUNT_SID}:${TWILIO_AUTH_TOKEN}`).toString('base64');
      const response = await fetchWithTimeout(`https://api.twilio.com/2010-04-01/Accounts/${TWILIO_ACCOUNT_SID}.json`, { headers: { Authorization: `Basic ${auth}` } }, 12000);
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
    }),
  ]);
  const ok = Object.values(checks).every((item) => item.ok);
  return reply.code(ok ? 200 : 503).send({ ok, time: new Date().toISOString(), checks });
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


// Automatic task endpoint. Power Automate should call this when a new email
// arrives in london@minaco.ca. Only task emails from Ramy's verified Minaco
// address are accepted. The endpoint returns quickly while the analysis runs
// as a delegated server job.
fastify.post('/task-inbox', async (request, reply) => {
  try {
    if (!TASK_INBOX_SECRET) {
      return reply.code(503).send({
        success: false,
        error: 'TASK_INBOX_SECRET is not configured, so automatic task intake is disabled.',
      });
    }

    const suppliedSecret = request.headers['x-london-task-secret'];
    if (!suppliedSecret || suppliedSecret !== TASK_INBOX_SECRET) {
      return reply.code(401).send({ success: false, error: 'Unauthorized.' });
    }

    const payload = request.body || {};
    const fromAddress = normalizeTaskSender(
      payload.from_address || payload.from || payload.sender || ''
    );

    if (!fromAddress || fromAddress !== normalizeTaskSender(RAMY_MINACO_EMAIL)) {
      return reply.code(403).send({
        success: false,
        error: 'This message is not an authorized task instruction from Ramy.',
      });
    }

    const messageId = String(
      payload.message_id || payload.messageId || payload.id || ''
    ).trim();
    const internetMessageId = String(
      payload.internet_message_id || payload.internetMessageId || ''
    ).trim();
    const subject = String(payload.subject || 'Untitled task').trim();
    const received = String(
      payload.received_date_time || payload.receivedDateTime || ''
    ).trim();

    const dedupeKey =
      internetMessageId ||
      messageId ||
      `${fromAddress}|${subject.toLowerCase()}|${received}`;

    const alreadyProcessed = processedTaskInboxKeys.get(dedupeKey);
    if (alreadyProcessed?.jobId) {
      return reply.code(202).send({
        success: true,
        queued: true,
        duplicate: true,
        jobId: alreadyProcessed.jobId,
      });
    }

    const queued = queueTaskInboxJob({
      messageId,
      subject,
      dedupeKey,
    });

    return reply.code(202).send(queued);
  } catch (error) {
    console.error('TASK INBOX endpoint error:', error);
    return reply.code(500).send({ success: false, error: error.message });
  }
});

fastify.get('/task-inbox/status/:jobId', async (request, reply) => {
  const suppliedSecret = request.headers['x-london-task-secret'];
  if (!TASK_INBOX_SECRET || suppliedSecret !== TASK_INBOX_SECRET) {
    return reply.code(401).send({ success: false, error: 'Unauthorized.' });
  }

  const job = taskInboxJobs.get(request.params.jobId);
  if (!job) {
    return reply.code(404).send({ success: false, error: 'Task job not found.' });
  }

  return reply.send({ success: true, job });
});


// Unified inbound SMS + WhatsApp command hub.
// Configure Twilio SMS and WhatsApp "A message comes in" webhooks to POST here.
const handleIncomingExecutiveMessage = async (request, reply) => {
  const body = request.body && typeof request.body === 'object' ? request.body : {};
  const from = String(body.From || request.query?.From || '').trim();
  const to = String(body.To || request.query?.To || '').trim();
  const messageBody = String(body.Body || request.query?.Body || '').trim();
  const messageSid = String(body.MessageSid || body.SmsMessageSid || '').trim();
  const numMedia = Number(body.NumMedia || 0);
  const waId = String(body.WaId || request.query?.WaId || '').trim();
  const profileName = String(body.ProfileName || request.query?.ProfileName || '').trim();

  // Twilio normally prefixes WhatsApp addresses with "whatsapp:", but WaId/ProfileName
  // provide a second reliable signal. This prevents a WhatsApp message from being
  // accidentally classified and answered as SMS.
  const isWhatsApp =
    Boolean(waId) ||
    Boolean(profileName) ||
    from.toLowerCase().startsWith('whatsapp:') ||
    to.toLowerCase().startsWith('whatsapp:');

  const channel = isWhatsApp ? 'whatsapp' : 'sms';

  const validSignature = validateTwilioFormWebhook(request);
  const authorizedSender = isAuthorizedRamyMessagingSender(from || waId);
  console.log('MESSAGING SECURITY CHECK:', {
    channel,
    sender: normalizePhoneIdentity(from || waId),
    waId: waId || '',
    profileName: profileName || '',
    rawFrom: from,
    rawTo: to,
    signatureValid: validSignature,
    authorizedSender,
    messageSid,
  });

  // Always return valid TwiML quickly; do not leave Twilio waiting while Graph/OpenAI runs.
  const emptyTwiml = '<?xml version="1.0" encoding="UTF-8"?><Response></Response>';

  if (!validSignature || !authorizedSender) {
    return reply.type('text/xml').code(403).send(emptyTwiml);
  }

  pruneMessagingState();
  if (messageSid && processedMessagingSids.has(messageSid)) {
    return reply.type('text/xml').send(emptyTwiml);
  }
  if (messageSid) { processedMessagingSids.set(messageSid, Date.now()); persistMessagingState(); }

  reply.type('text/xml').send(emptyTwiml);

  setImmediate(async () => {
    try {
      let result;
      if (!messageBody && numMedia > 0) {
        result = {
          success: false,
          reply: 'I received the attachment. For document or spreadsheet analysis, email it to london@minaco.ca so I can process the full file safely.',
        };
      } else {
        result = await processExecutiveMessagingInstruction({
          text: messageBody,
          channel,
          sender: from || waId,
        });
      }

      const replyTo =
        channel === 'whatsapp'
          ? `whatsapp:${normalizePhoneIdentity(from || waId)}`
          : normalizePhoneIdentity(from);

      const replyFrom =
        channel === 'whatsapp'
          ? `whatsapp:${normalizePhoneIdentity(to || TWILIO_PHONE_NUMBER)}`
          : normalizePhoneIdentity(to || TWILIO_PHONE_NUMBER);

      await sendTwilioChannelMessage({
        to: replyTo,
        from: replyFrom,
        body: result.reply,
      });
    } catch (error) {
      console.error('Inbound executive messaging failure:', error);
      try {
        const errorReplyTo =
          channel === 'whatsapp'
            ? `whatsapp:${normalizePhoneIdentity(from || waId)}`
            : normalizePhoneIdentity(from);

        const errorReplyFrom =
          channel === 'whatsapp'
            ? `whatsapp:${normalizePhoneIdentity(to || TWILIO_PHONE_NUMBER)}`
            : normalizePhoneIdentity(to || TWILIO_PHONE_NUMBER);

        await sendTwilioChannelMessage({
          to: errorReplyTo,
          from: errorReplyFrom,
          body: `I received your ${channel === 'whatsapp' ? 'WhatsApp' : 'text'} command but could not complete it: ${error.message}`,
        });
      } catch (notifyError) {
        console.error('Inbound executive messaging failure notification failed:', notifyError);
      }
    }
  });
};

fastify.all('/incoming-message', handleIncomingExecutiveMessage);
fastify.all('/incoming-sms', handleIncomingExecutiveMessage);
fastify.all('/incoming-whatsapp', handleIncomingExecutiveMessage);

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
      let latestUserSpeechStartedAt = 0;
      let lastUserTranscript = '';
      let lastUserTranscriptSpeechStartedAt = 0;
      let emailConfirmationArmedAt = 0;
      let calendarConfirmationArmedAt = 0;
      let activeToolCallId = null;
      let activeToolName = null;
      let activeToolStartedAt = 0;
      const cancelledToolCalls = new Set();
      const cancellableToolNames = new Set([
        'check_email',
        'read_email',
        'search_email',
        'check_accounting',
        'read_accounting_email',
        'list_email_attachments',
        'analyze_email_attachment',
        'resolve_person',
        'check_calendar',
        'check_availability',
        'find_availability',
        'prepare_availability_reply',
        'prepare_next_week_availability_reply',
        'daily_executive_brief',
        'list_actions',
      ]);

      let openAiWs = null;
      let openAiGeneration = 0;
      let reconnectAttempts = 0;
      let reconnectTimer = null;
      let heartbeatTimer = null;
      let watchdogTimer = null;
      let lastOpenAiEventAt = Date.now();
      let lastOpenAiPongAt = Date.now();
      let lastAssistantAudioAt = 0;
      let awaitingResponseSince = 0;
      let callClosed = false;
      let sessionReady = false;
      const pendingAudioFrames = [];

      const OPENAI_HEARTBEAT_MS = 10000;
      const OPENAI_PONG_TIMEOUT_MS = 25000;
      const VOICE_TOOL_TIMEOUT_MS = 18000;
      const RESPONSE_STALL_TIMEOUT_MS = 20000;
      const MAX_OPENAI_RECONNECT_ATTEMPTS = 4;
      const MAX_BUFFERED_AUDIO_FRAMES = 150;

      const safeOpenAiSend = (payload) => {
        if (!openAiWs || openAiWs.readyState !== WebSocket.OPEN) return false;
        const data = typeof payload === 'string' ? payload : JSON.stringify(payload);
        try {
          openAiWs.send(data);
          return true;
        } catch (error) {
          console.error('OPENAI_SOCKET send failed:', error.message);
          return false;
        }
      };

      const respondToToolCall = (callId, output, instructions) => {
        if (cancelledToolCalls.has(callId)) {
          cancelledToolCalls.delete(callId);
          return;
        }

        if (callId === activeToolCallId) {
          activeToolCallId = null;
          activeToolName = null;
          activeToolStartedAt = 0;
        }

        safeOpenAiSend(
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

        safeOpenAiSend(
          JSON.stringify({
            type: 'response.create',
            response: {
              instructions,
            },
          })
        );
      };

      const cancelActiveReadOnlyTool = () => {
        if (!activeToolCallId || !activeToolName) return false;
        if (!cancellableToolNames.has(activeToolName)) return false;
        if (Date.now() - activeToolStartedAt < 500) return false;

        const callId = activeToolCallId;
        const toolName = activeToolName;
        cancelledToolCalls.add(callId);
        activeToolCallId = null;
        activeToolName = null;
        activeToolStartedAt = 0;

        safeOpenAiSend(
          JSON.stringify({
            type: 'conversation.item.create',
            item: {
              type: 'function_call_output',
              call_id: callId,
              output: JSON.stringify({
                success: false,
                cancelled: true,
                error: `The ${toolName} lookup was interrupted by Ramy.`,
              }),
            },
          })
        );

        console.log('Cancelled active read-only voice tool:', toolName);
        return true;
      };

      const normalizeConfirmationTranscript = (value) =>
        String(value || '')
          .toLowerCase()
          .replace(/[^a-z0-9\s]/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();

      const isExplicitEmailSendConfirmation = (value) => {
        const text = normalizeConfirmationTranscript(value);
        return /^(?:yes\s+)?(?:please\s+)?send\s+(?:it|the\s+email|the\s+reply)(?:\s+please)?$/.test(text) ||
          /^go\s+ahead(?:\s+and)?\s+send\s+(?:it|the\s+email|the\s+reply)$/.test(text);
      };

      const isExplicitCalendarConfirmation = (value, actionType) => {
        const text = normalizeConfirmationTranscript(value);
        if (actionType === 'create') {
          return /^(?:yes\s+)?(?:please\s+)?book\s+it(?:\s+please)?$/.test(text) ||
            /^go\s+ahead(?:\s+and)?\s+book\s+it$/.test(text);
        }
        if (actionType === 'update') {
          return /^(?:yes\s+)?(?:please\s+)?(?:reschedule\s+it|make\s+the\s+change)(?:\s+please)?$/.test(text) ||
            /^go\s+ahead(?:\s+and)?\s+(?:reschedule\s+it|make\s+the\s+change)$/.test(text);
        }
        if (actionType === 'cancel') {
          return /^(?:yes\s+)?(?:please\s+)?cancel\s+it(?:\s+please)?$/.test(text) ||
            /^go\s+ahead(?:\s+and)?\s+cancel\s+it$/.test(text);
        }
        return false;
      };

      const waitForUserConfirmationTranscript = async (armedAt, timeoutMs = 2200) => {
        const deadline = Date.now() + timeoutMs;
        while (Date.now() < deadline) {
          if (
            lastUserTranscript &&
            lastUserTranscriptSpeechStartedAt > armedAt
          ) {
            return lastUserTranscript;
          }
          await new Promise((resolve) => setTimeout(resolve, 50));
        }
        return '';
      };

      const armEmailConfirmation = () => {
        emailConfirmationArmedAt = Date.now();
      };

      const armCalendarConfirmation = () => {
        calendarConfirmationArmedAt = Date.now();
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
                transcription: {
                  model: 'gpt-4o-mini-transcribe',
                  prompt: 'Ramy Mina, Minaco, London Assistant, Francis Deslauriers, Curé-Labelle, Laval, Joannie, Addenda 01',
                },
                turn_detection: {
                  type: 'semantic_vad',
                  // Medium eagerness balances waiting through Ramy's brief pauses with responsiveness.
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
                name: 'delegate_email_thread_review',
                description:
                  'Queue a background review of a live Minaco email conversation and email the completed summary/analysis to Ramy himself. Use for multi-step requests such as read the thread from Lias, review Hussein supplier response, summarize it in English, identify risks before signing, and email me your analysis. This tool returns immediately and sends only to Ramy, never to an external recipient.',
                parameters: {
                  type: 'object',
                  properties: {
                    lookup_query: {
                      type: 'string',
                      description:
                        'Sender, company, email address, subject, or thread search term used to identify the live email conversation, for example Lias.',
                    },
                    focus: {
                      type: 'string',
                      description:
                        'Optional person, supplier, or issue within the thread to focus on, for example Hussein or heat-pump service and warranty.',
                    },
                    instruction: {
                      type: 'string',
                      description:
                        'What Ramy wants done with the thread, including summary, translation, commercial assessment, risks, questions, or recommendation.',
                    },
                    output_language: {
                      type: 'string',
                      description:
                        'Language for the emailed review. Default English unless Ramy asks otherwise.',
                    },
                  },
                  required: ['lookup_query', 'instruction'],
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
                    sender_identity: {
                      type: 'string',
                      enum: ['ramy', 'london'],
                      description: 'Mailbox identity for this new email. Use ramy when Ramy says from my email/from me or does not specify; use london only when he explicitly says from London.',
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
                  'Send the currently pending prepared email action. Use ONLY after Ramy explicitly says Send it, Yes send it, Send the email/reply, or Go ahead and send it in a new user turn after the draft was read back. The server validates the actual user transcript; thanks, okay, perfect, yes alone, or other acknowledgments are never sufficient.',
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
                  'Search historical live email by person, company, subject, or keywords. Can search Ramy, Accounting, or London Task Inbox. Use when the desired email may not be in the latest inbox list.',
                parameters: {
                  type: 'object',
                  properties: {
                    mailbox: {
                      type: 'string',
                      enum: ['ramy', 'accounting', 'london'],
                      description: 'Mailbox to search: ramy, accounting, or london. Default ramy.',
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
                      enum: ['ramy', 'accounting', 'london'],
                      description: 'Mailbox containing the email: ramy, accounting, or london.',
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
                      enum: ['ramy', 'accounting', 'london'],
                      description: 'Mailbox containing the email: ramy, accounting, or london.',
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
                name: 'list_dropbox',
                description: 'List files and folders inside London\'s controlled Dropbox workspace only. Read-only.',
                parameters: {
                  type: 'object',
                  properties: {
                    path: { type: 'string', description: 'Path inside LONDON - ACCESS. Default root.' },
                    recursive: { type: 'boolean', description: 'Whether to include descendants. Use true when Ramy asks for subfolders, sub-subfolders, deeper structure, everything under a folder, or tells London to find it on her own. Default false.' },
                  },
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'search_dropbox',
                description: 'Search files inside London\'s LONDON - ACCESS Dropbox workspace. Read-only.',
                parameters: {
                  type: 'object',
                  properties: { query: { type: 'string', description: 'Filename or content search term.' } },
                  required: ['query'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'read_dropbox_text',
                description: 'Read a plain-text file inside LONDON - ACCESS. Read-only.',
                parameters: {
                  type: 'object',
                  properties: { path: { type: 'string', description: 'Exact path inside LONDON - ACCESS.' } },
                  required: ['path'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'delegate_dropbox_file_analysis',
                description: 'Queue background analysis of a real PDF, Word, Excel, PowerPoint, image, or other supported file already inside LONDON - ACCESS and email the completed analysis to Ramy. Use when Ramy asks to analyze/review a Dropbox file, especially when he says email me the analysis. This is read-only and never emails an external party.',
                parameters: {
                  type: 'object',
                  properties: {
                    path: { type: 'string', description: 'Exact Dropbox path inside LONDON - ACCESS. Resolve approximate names with search_dropbox first.' },
                    instruction: { type: 'string', description: 'What Ramy wants analyzed, checked, compared, summarized, or extracted from the file.' },
                  },
                  required: ['path', 'instruction'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'quick_action_update',
                description:
                  'Fast natural-language Action Register update. Use as the first choice for simple phrases like Joannie done, waiting on Anass until Friday, follow up next Tuesday, add task call Makar tomorrow, cancel this item, or several quick updates in one instruction. The server matches and updates the live register in one step.',
                parameters: {
                  type: 'object',
                  properties: {
                    instruction: {
                      type: 'string',
                      description: 'Ramy’s natural-language action/status update exactly as intended.',
                    },
                  },
                  required: ['instruction'],
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
                    date_range: {
                      type: 'string',
                      description: 'Relative or explicit period such as next week, this week, tomorrow, next 7 days, or 2026-08-31 to 2026-09-04. Prefer this for relative requests.',
                    },
                    start_date: {
                      type: 'string',
                      description: 'Optional first Montreal local date in YYYY-MM-DD when Ramy gives explicit dates.',
                    },
                    end_date: {
                      type: 'string',
                      description: 'Optional last Montreal local date in YYYY-MM-DD, inclusive, when Ramy gives explicit dates.',
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
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'prepare_next_week_availability_reply',
                description:
                  'Dedicated fast path for requests such as respond to Francis and give him three availabilities for next week. The server itself resolves next Monday-Friday in Montreal time, finds the latest inbound email FROM that sender, checks Ramy’s live calendar, and prepares a reply. It NEVER sends. Use this instead of prepare_availability_reply whenever the user explicitly says next week.',
                parameters: {
                  type: 'object',
                  properties: {
                    person_query: {
                      type: 'string',
                      description:
                        'Sender name or email exactly as Ramy identifies the person, for example Francis.',
                    },
                    duration_minutes: {
                      type: 'integer',
                      minimum: 15,
                      maximum: 240,
                      description:
                        'Meeting duration. Default 30 minutes if Ramy does not specify one.',
                    },
                    max_slots: {
                      type: 'integer',
                      minimum: 1,
                      maximum: 5,
                      description:
                        'Number of availability options. Default 3.',
                    },
                    reply_mode: {
                      type: 'string',
                      enum: ['sender', 'all'],
                      description:
                        'sender for reply-to-sender only; all for reply-all. Default sender.',
                    },
                  },
                  required: ['person_query'],
                  additionalProperties: false,
                },
              },
              {
                type: 'function',
                name: 'prepare_availability_reply',
                description:
                  'Fast one-step workflow for requests such as respond to Francis and give him my availability next week. It searches the named sender in Ramy’s live mailbox, checks his live calendar for real open slots, and prepares a reply draft. It NEVER sends. Prefer this over chaining several separate tools for a simple availability reply.',
                parameters: {
                  type: 'object',
                  properties: {
                    person_query: {
                      type: 'string',
                      description: 'Sender name or email, for example Francis.',
                    },
                    date_range: {
                      type: 'string',
                      description: 'Relative or explicit period such as next week, this week, tomorrow, next 7 days, or 2026-08-31 to 2026-09-04. Prefer this for relative requests.',
                    },
                    start_date: {
                      type: 'string',
                      description: 'Optional first Montreal local date in YYYY-MM-DD when Ramy gives explicit dates.',
                    },
                    end_date: {
                      type: 'string',
                      description: 'Optional last Montreal local date in YYYY-MM-DD, inclusive, when Ramy gives explicit dates.',
                    },
                    duration_minutes: {
                      type: 'integer',
                      minimum: 15,
                      maximum: 240,
                      description: 'Meeting duration. Default 30 minutes.',
                    },
                    day_start: {
                      type: 'string',
                      description: 'Earliest Montreal local time in HH:mm. Default 09:00.',
                    },
                    day_end: {
                      type: 'string',
                      description: 'Latest Montreal local ending boundary in HH:mm. Default 17:00.',
                    },
                    max_slots: {
                      type: 'integer',
                      minimum: 1,
                      maximum: 5,
                      description: 'Number of availability options. Default 3.',
                    },
                    include_weekends: {
                      type: 'boolean',
                      description: 'Whether weekends may be offered. Default false.',
                    },
                    reply_mode: {
                      type: 'string',
                      enum: ['sender', 'all'],
                      description: 'sender for reply-to-sender only; all for reply-all. Default sender when Ramy names the person directly.',
                    },
                  },
                  required: ['person_query'],
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
        safeOpenAiSend(JSON.stringify(sessionUpdate));
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
            safeOpenAiSend(JSON.stringify(truncateEvent));
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

      const handleOpenAiMessage = async (data) => {
        try {
          const response = JSON.parse(data);
          lastOpenAiEventAt = Date.now();

          if (response.type === 'input_audio_buffer.speech_stopped') {
            awaitingResponseSince = Date.now();
          }

          if (response.type === 'response.output_audio.delta' && response.delta) {
            lastAssistantAudioAt = Date.now();
            awaitingResponseSince = 0;
          }

          if (response.type === 'response.done') {
            awaitingResponseSince = 0;
          }

          if (response.type === 'conversation.item.input_audio_transcription.completed') {
            lastUserTranscript = String(response.transcript || '').trim();
            lastUserTranscriptSpeechStartedAt = latestUserSpeechStartedAt;
            console.log('USER TRANSCRIPT:', lastUserTranscript);

            const normalizedStop = lastUserTranscript
              .toLowerCase()
              .replace(/[^a-z0-9\s]/g, ' ')
              .replace(/\s+/g, ' ')
              .trim();

            if (/^(stop|stop london|cancel|never mind|nevermind)$/.test(normalizedStop)) {
              console.log('VOICE_STOP received from Ramy');
              cancelActiveReadOnlyTool();
              safeOpenAiSend({ type: 'response.cancel' });
              if (streamSid && connection.readyState === WebSocket.OPEN) {
                try {
                  connection.send(JSON.stringify({ event: 'clear', streamSid }));
                } catch (error) {
                  console.warn('TWILIO clear on stop failed:', error.message);
                }
              }
              markQueue = [];
              lastAssistantItem = null;
              responseStartTimestampTwilio = null;
              awaitingResponseSince = 0;
              setTimeout(() => {
                safeOpenAiSend({
                  type: 'response.create',
                  response: {
                    instructions: 'Ramy said stop. Say only: Stopped. Then wait for his next instruction.',
                  },
                });
              }, 100);
            }
          }

          if (response.type === 'conversation.item.input_audio_transcription.failed') {
            console.error('User transcription failed:', response.error || response);
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            cancellableToolNames.has(response.name)
          ) {
            activeToolCallId = response.call_id;
            activeToolName = response.name;
            activeToolStartedAt = Date.now();
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'delegate_email_thread_review'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const queued = queueDelegatedThreadReviewJob({
                lookupQuery: args.lookup_query,
                focus: args.focus || '',
                instruction: args.instruction || '',
                outputLanguage: args.output_language || 'English',
              });

              respondToToolCall(
                response.call_id,
                queued,
                `Tell Ramy the delegated email-thread review has been queued successfully and the completed review will be emailed to ${RAMY_MINACO_EMAIL}. Tell him he may hang up now if he wants. Keep this to one or two short sentences. Do not say the review is already complete.`
              );
            } catch (error) {
              console.error('Delegate email thread review queue error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, queued: false, error: error.message },
                'Tell Ramy the delegated review could not be queued and state the returned error concisely. Do not claim any background work will continue.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'send_confirmed_email'
          ) {
            try {
              if (!pendingEmailActionType || (!pendingEmailReply && !pendingEmailDraft)) {
                throw new Error('There is no pending email action to send.');
              }

              const confirmationTranscript = await waitForUserConfirmationTranscript(
                emailConfirmationArmedAt
              );

              if (!isExplicitEmailSendConfirmation(confirmationTranscript)) {
                respondToToolCall(
                  response.call_id,
                  {
                    success: false,
                    sent: false,
                    confirmationRequired: true,
                    heard: confirmationTranscript || null,
                  },
                  'Do NOT call another tool. Tell Ramy the email was NOT sent because an explicit send command was not heard. Ask him to say exactly Send it if he wants to send. Thank you, okay, perfect, sounds good, and yes alone do not count.'
                );
                return;
              }

              if (pendingEmailActionType === 'reply') {
                if (!pendingEmailReply) {
                  throw new Error('There is no pending email reply to send.');
                }

                const replyToSend = { ...pendingEmailReply };
                const result = await replyToMinacoEmail(replyToSend);
                pendingEmailReply = null;
                pendingEmailDraft = null;
                pendingEmailActionType = null;
                emailConfirmationArmedAt = 0;

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
                const result = await sendEmailFromMailbox(emailToSend);
                pendingEmailDraft = null;
                pendingEmailReply = null;
                pendingEmailActionType = null;
                emailConfirmationArmedAt = 0;

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

              const senderIdentity = String(args.sender_identity || 'ramy').toLowerCase();
              const senderMailbox = senderIdentity === 'london' ? LONDON_MINACO_EMAIL : RAMY_MINACO_EMAIL;
              if (!senderMailbox) throw new Error('The selected sender mailbox is not configured.');
              pendingEmailDraft = {
                from: senderMailbox,
                to: args.to,
                subject: args.subject,
                body: args.body,
              };
              pendingEmailReply = null;
              pendingEmailActionType = 'new';
              armEmailConfirmation();

              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  draft: pendingEmailDraft,
                  sent: false,
                },
                'Read back the sender mailbox, recipient, subject, and email message to Ramy. Clearly say the email has NOT been sent. Ask Ramy to confirm by saying Send it.'
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
              armEmailConfirmation();

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
            response.name === 'list_dropbox'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const entries = await listLondonDropbox(args.path || LONDON_DROPBOX_ROOT, Boolean(args.recursive));
              respondToToolCall(response.call_id, { success: true, root: LONDON_DROPBOX_ROOT, entries }, 'Summarize the Dropbox listing concisely. If recursive descendants were returned, organize them by folder hierarchy and answer the requested depth without asking Ramy to choose a child folder. Use only the returned entries and never imply access outside the controlled root.');
            } catch (error) {
              respondToToolCall(response.call_id, { success: false, error: error.message }, 'Tell Ramy the Dropbox listing failed and state the error concisely.');
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'search_dropbox'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const results = await searchLondonDropbox(args.query);
              respondToToolCall(response.call_id, { success: true, root: LONDON_DROPBOX_ROOT, results }, 'Summarize the Dropbox search results concisely. Use only the returned results.');
            } catch (error) {
              respondToToolCall(response.call_id, { success: false, error: error.message }, 'Tell Ramy the Dropbox search failed and state the error concisely.');
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'read_dropbox_text'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const result = await readLondonDropboxText(args.path);
              respondToToolCall(response.call_id, { success: true, ...result }, 'Use only the returned Dropbox file text. Answer Ramy directly and concisely.');
            } catch (error) {
              respondToToolCall(response.call_id, { success: false, error: error.message }, 'Tell Ramy the Dropbox file could not be read and state the error concisely.');
            }
            return;
          }


          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'delegate_dropbox_file_analysis'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const queued = queueDelegatedDropboxAnalysisJob({
                path: args.path,
                instruction: args.instruction || '',
              });
              respondToToolCall(
                response.call_id,
                queued,
                `Tell Ramy the Dropbox file analysis has been queued and the completed review will be emailed to ${RAMY_MINACO_EMAIL}. Keep it to one or two short sentences. Do not say the analysis is already complete.`
              );
            } catch (error) {
              respondToToolCall(
                response.call_id,
                { success: false, queued: false, error: error.message },
                'Tell Ramy the Dropbox analysis could not be queued and state the returned error concisely. Do not tell him to re-upload or email the file unless the returned error specifically requires it.'
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'quick_action_update'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const result = await processQuickActionInstruction({
                text: args.instruction,
                channel: 'voice',
                sender: RAMY_PHONE_NUMBER,
              });
              respondToToolCall(
                response.call_id,
                result,
                'Tell Ramy exactly what the Action Register update result says, briefly. If it asks a clarification question, ask only that question. Do not add new facts.'
              );
            } catch (error) {
              console.error('Quick action update error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the quick Action Register update failed and state the error concisely.'
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
            response.name === 'prepare_next_week_availability_reply'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const result = await buildAvailabilityReplyDraft({
                personQuery: args.person_query,
                dateRangeText: 'next week',
                durationMinutes: args.duration_minutes ?? 30,
                dayStart: '09:00',
                dayEnd: '17:00',
                maxSlots: args.max_slots ?? 3,
                includeWeekends: false,
                replyMode: args.reply_mode || 'sender',
              });

              if (result.needsClarification) {
                respondToToolCall(
                  response.call_id,
                  { success: false, ...result },
                  'The next-week availability workflow needs one clarification about the sender or recipient. Ask ONE short question based only on the returned reason or candidates. Do not ask Ramy for dates or for his availability.'
                );
                return;
              }

              pendingEmailReply = {
                messageId: result.email.messageId,
                mode: result.replyMode,
                body: result.body,
                subject: result.email.subject,
                recipients: result.recipients,
              };
              pendingEmailDraft = null;
              pendingEmailActionType = 'reply';
              armEmailConfirmation();

              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  sent: false,
                  from: RAMY_MINACO_EMAIL,
                  sourceEmail: result.email,
                  replyMode: result.replyMode,
                  recipients: result.recipients,
                  availability: result.availability,
                  body: result.body,
                },
                'This is the completed next-week availability reply. Read back the three exact Montreal-time options and the prepared reply. Clearly say it has NOT been sent and ask Ramy to say Send it. Do not run another availability or email search unless Ramy changes the request.'
              );
            } catch (error) {
              console.error('Next-week availability reply error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                "Tell Ramy the next-week reply workflow failed and state the exact technical error in one sentence. Do not ask him to provide calendar availability, dates, or Francis's email address. Stay ready for another command."
              );
            }
            return;
          }

          if (
            response.type === 'response.function_call_arguments.done' &&
            response.name === 'prepare_availability_reply'
          ) {
            try {
              const args = JSON.parse(response.arguments || '{}');
              const result = await buildAvailabilityReplyDraft({
                personQuery: args.person_query,
                startDate: args.start_date,
                endDate: args.end_date,
                dateRangeText: args.date_range,
                durationMinutes: args.duration_minutes ?? 30,
                dayStart: args.day_start || '09:00',
                dayEnd: args.day_end || '17:00',
                maxSlots: args.max_slots ?? 3,
                includeWeekends: Boolean(args.include_weekends),
                replyMode: args.reply_mode || 'sender',
              });

              if (result.needsClarification) {
                respondToToolCall(
                  response.call_id,
                  { success: false, ...result },
                  'The fast availability-reply workflow needs clarification. Ask ONE short question based only on the returned reason or candidates. Do not ask Ramy to state his availability.'
                );
                return;
              }

              pendingEmailReply = {
                messageId: result.email.messageId,
                mode: result.replyMode,
                body: result.body,
                subject: result.email.subject,
                recipients: result.recipients,
              };
              pendingEmailDraft = null;
              pendingEmailActionType = 'reply';
              armEmailConfirmation();

              respondToToolCall(
                response.call_id,
                {
                  success: true,
                  sent: false,
                  from: RAMY_MINACO_EMAIL,
                  sourceEmail: result.email,
                  replyMode: result.replyMode,
                  recipients: result.recipients,
                  availability: result.availability,
                  body: result.body,
                },
                'Read back the actual available Montreal-time options and the prepared reply. Clearly say it has NOT been sent. Ask Ramy to confirm by saying Send it. Do not run more tools unless Ramy changes the request.'
              );
            } catch (error) {
              console.error('Fast availability reply error:', error);
              respondToToolCall(
                response.call_id,
                { success: false, error: error.message },
                'Tell Ramy the availability lookup did not complete. If the error is about a date range, say the date-range resolver failed and that he can simply say next week again; do not ask him to calculate dates. Stay responsive and ready for his next instruction.'
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
              const resolvedRange = resolveAvailabilityDateRange({
                startDate: args.start_date,
                endDate: args.end_date,
                dateRangeText: args.date_range,
              });
              const result = await findCalendarAvailability({
                startDate: resolvedRange.startDate,
                endDate: resolvedRange.endDate,
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
              armCalendarConfirmation();

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
              armCalendarConfirmation();

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
              armCalendarConfirmation();

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
              const confirmationTranscript = await waitForUserConfirmationTranscript(
                calendarConfirmationArmedAt
              );

              if (!isExplicitCalendarConfirmation(confirmationTranscript, action.type)) {
                const expected =
                  action.type === 'create'
                    ? 'Book it'
                    : action.type === 'update'
                      ? 'Reschedule it or Make the change'
                      : 'Cancel it';
                respondToToolCall(
                  response.call_id,
                  {
                    success: false,
                    completed: false,
                    confirmationRequired: true,
                    heard: confirmationTranscript || null,
                  },
                  `Do NOT call another tool. Tell Ramy the calendar was NOT changed because an explicit confirmation was not heard. Ask him to say ${expected}. Polite acknowledgments do not count.`
                );
                return;
              }

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
              calendarConfirmationArmedAt = 0;

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
            latestUserSpeechStartedAt = Date.now();

            // Barge-in should stop London's spoken response, but it must NOT
            // cancel an email/calendar/brief lookup merely because Ramy begins
            // speaking. Read-only tools continue in the background unless Ramy
            // explicitly says stop/cancel/never mind, which is handled by the
            // transcript-based stop command logic.
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
      };

      const stopOpenAiTimers = () => {
        if (heartbeatTimer) {
          clearInterval(heartbeatTimer);
          heartbeatTimer = null;
        }
        if (reconnectTimer) {
          clearTimeout(reconnectTimer);
          reconnectTimer = null;
        }
      };

      const scheduleOpenAiReconnect = (reason) => {
        if (callClosed || reconnectTimer) return;

        if (reconnectAttempts >= MAX_OPENAI_RECONNECT_ATTEMPTS) {
          console.error('RECONNECT exhausted:', {
            attempts: reconnectAttempts,
            reason,
          });
          try {
            if (connection.readyState === WebSocket.OPEN) {
              connection.close(1011, 'London voice connection unavailable');
            }
          } catch (error) {
            console.error('TWILIO close after reconnect exhaustion failed:', error.message);
          }
          return;
        }

        reconnectAttempts += 1;
        const delay = Math.min(500 * 2 ** (reconnectAttempts - 1), 4000);
        console.warn('RECONNECT scheduled:', {
          attempt: reconnectAttempts,
          delayMs: delay,
          reason,
        });

        reconnectTimer = setTimeout(() => {
          reconnectTimer = null;
          connectOpenAi(`reconnect:${reason}`);
        }, delay);
      };

      const attachOpenAiSocketHandlers = (ws, generation, connectionReason) => {
        ws.on('open', () => {
          if (generation !== openAiGeneration || callClosed) {
            try { ws.close(); } catch {}
            return;
          }

          console.log('OPENAI_SOCKET open:', {
            generation,
            reason: connectionReason,
            reconnectAttempts,
          });
          sessionReady = false;
          lastOpenAiEventAt = Date.now();
          lastOpenAiPongAt = Date.now();

          setTimeout(() => {
            if (generation !== openAiGeneration || ws.readyState !== WebSocket.OPEN) return;
            initializeSession();
            sessionReady = true;

            setTimeout(() => {
              if (generation !== openAiGeneration || ws.readyState !== WebSocket.OPEN) return;
              while (pendingAudioFrames.length) {
                const payload = pendingAudioFrames.shift();
                safeOpenAiSend({
                  type: 'input_audio_buffer.append',
                  audio: payload,
                });
              }

              if (connectionReason !== 'initial') {
                safeOpenAiSend({
                  type: 'response.create',
                  response: {
                    instructions:
                      'The realtime voice connection briefly dropped and has now recovered. Tell Ramy in one short sentence that you are back and ask him to repeat the last instruction. Do not claim any interrupted tool or external action completed.',
                  },
                });
              }
            }, 350);
          }, 100);

          // Reset the reconnect counter only after the replacement socket stays stable.
          setTimeout(() => {
            if (
              generation === openAiGeneration &&
              openAiWs === ws &&
              ws.readyState === WebSocket.OPEN
            ) {
              reconnectAttempts = 0;
            }
          }, 15000);

          if (heartbeatTimer) clearInterval(heartbeatTimer);
          heartbeatTimer = setInterval(() => {
            if (generation !== openAiGeneration || callClosed) return;
            if (ws.readyState !== WebSocket.OPEN) return;

            const now = Date.now();
            if (now - lastOpenAiPongAt > OPENAI_PONG_TIMEOUT_MS) {
              console.error('WATCHDOG OpenAI pong timeout:', {
                generation,
                msSincePong: now - lastOpenAiPongAt,
              });
              try { ws.terminate(); } catch {}
              return;
            }

            try {
              ws.ping();
            } catch (error) {
              console.error('OPENAI_SOCKET ping failed:', error.message);
              try { ws.terminate(); } catch {}
            }
          }, OPENAI_HEARTBEAT_MS);
          if (typeof heartbeatTimer.unref === 'function') heartbeatTimer.unref();
        });

        ws.on('pong', () => {
          if (generation === openAiGeneration) {
            lastOpenAiPongAt = Date.now();
          }
        });

        ws.on('message', handleOpenAiMessage);

        ws.on('close', (code, reasonBuffer) => {
          if (generation !== openAiGeneration) return;
          sessionReady = false;
          if (heartbeatTimer) {
            clearInterval(heartbeatTimer);
            heartbeatTimer = null;
          }

          const reason = Buffer.isBuffer(reasonBuffer)
            ? reasonBuffer.toString('utf8')
            : String(reasonBuffer || '');
          console.warn('OPENAI_SOCKET closed:', { generation, code, reason });

          if (activeToolCallId && cancellableToolNames.has(activeToolName)) {
            cancelledToolCalls.add(activeToolCallId);
            activeToolCallId = null;
            activeToolName = null;
            activeToolStartedAt = 0;
          }

          if (!callClosed) scheduleOpenAiReconnect(`close:${code}`);
        });

        ws.on('error', (error) => {
          if (generation !== openAiGeneration) return;
          console.error('OPENAI_SOCKET error:', {
            generation,
            message: error.message,
          });
          // ws will normally emit close after error. Force termination so the
          // reconnect path is deterministic if the socket remains half-open.
          try {
            if (ws.readyState !== WebSocket.CLOSED) ws.terminate();
          } catch {}
        });
      };

      const connectOpenAi = (reason = 'initial') => {
        if (callClosed) return;

        openAiGeneration += 1;
        const generation = openAiGeneration;
        sessionReady = false;

        console.log('OPENAI_SOCKET connecting:', { generation, reason });
        const ws = new WebSocket(
          `wss://api.openai.com/v1/realtime?model=gpt-realtime&temperature=${TEMPERATURE}`,
          {
            headers: {
              Authorization: `Bearer ${OPENAI_API_KEY}`,
            },
            handshakeTimeout: 12000,
          }
        );
        openAiWs = ws;
        attachOpenAiSocketHandlers(ws, generation, reason);
      };

      watchdogTimer = setInterval(() => {
        if (callClosed) return;
        const now = Date.now();

        if (
          activeToolCallId &&
          activeToolName &&
          cancellableToolNames.has(activeToolName) &&
          activeToolStartedAt &&
          now - activeToolStartedAt > VOICE_TOOL_TIMEOUT_MS
        ) {
          const timedOutTool = activeToolName;
          console.error('WATCHDOG tool timeout:', {
            tool: timedOutTool,
            ms: now - activeToolStartedAt,
          });
          const cancelled = cancelActiveReadOnlyTool();
          if (cancelled) {
            safeOpenAiSend({
              type: 'response.create',
              response: {
                instructions: `The live ${timedOutTool} lookup timed out. Tell Ramy concisely that the lookup did not complete and that you are still responsive. Do not invent a result.`,
              },
            });
          }
        }

        if (
          awaitingResponseSince &&
          now - awaitingResponseSince > RESPONSE_STALL_TIMEOUT_MS &&
          now - lastOpenAiEventAt > 8000 &&
          openAiWs &&
          openAiWs.readyState === WebSocket.OPEN
        ) {
          console.error('WATCHDOG response stall:', {
            awaitingMs: now - awaitingResponseSince,
            msSinceOpenAiEvent: now - lastOpenAiEventAt,
            msSinceAssistantAudio: lastAssistantAudioAt
              ? now - lastAssistantAudioAt
              : null,
          });
          awaitingResponseSince = 0;
          try { openAiWs.terminate(); } catch {}
        }
      }, 2000);
      if (typeof watchdogTimer.unref === 'function') watchdogTimer.unref();

      connectOpenAi('initial');

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
              if (openAiWs && openAiWs.readyState === WebSocket.OPEN && sessionReady) {
                safeOpenAiSend({
                  type: 'input_audio_buffer.append',
                  audio: data.media.payload,
                });
              } else {
                pendingAudioFrames.push(data.media.payload);
                if (pendingAudioFrames.length > MAX_BUFFERED_AUDIO_FRAMES) {
                  pendingAudioFrames.shift();
                }
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
        callClosed = true;
        stopOpenAiTimers();
        if (watchdogTimer) {
          clearInterval(watchdogTimer);
          watchdogTimer = null;
        }
        if (openAiWs && openAiWs.readyState !== WebSocket.CLOSED) {
          try { openAiWs.close(1000, 'Twilio client disconnected'); } catch {}
        }
        console.log('TWILIO client disconnected.');
      });

      connection.on('error', (error) => {
        console.error('TWILIO socket error:', error.message || error);
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
