const read = (env, key, fallback = '') => String(env[key] ?? fallback).trim();

const normalizeVoiceModel = (value) => {
  const requested = String(value || '').trim();
  if (!requested || requested === 'gpt-realtime') return 'gpt-realtime-1.5';
  return requested;
};

export function loadConfig(env = process.env) {
  const config = {
    openai: {
      apiKey: read(env, 'OPENAI_API_KEY'),
      taskModel: read(env, 'TASK_INBOX_MODEL', read(env, 'OPENAI_DOCUMENT_MODEL', 'gpt-5.6')),
    },
    microsoft: {
      readTenantId: read(env, 'MS_TENANT_ID'),
      readClientId: read(env, 'MS_CLIENT_ID'),
      readClientSecret: read(env, 'MS_CLIENT_SECRET'),
      actionTenantId: read(env, 'ACTIONS_MS_TENANT_ID', read(env, 'MS_TENANT_ID')),
      actionClientId: read(env, 'ACTIONS_MS_CLIENT_ID', read(env, 'MS_CLIENT_ID')),
      actionClientSecret: read(env, 'ACTIONS_MS_CLIENT_SECRET', read(env, 'MS_CLIENT_SECRET')),
      londonMailbox: read(env, 'LONDON_MINACO_EMAIL', 'london@minaco.ca'),
      ramyMailbox: read(env, 'RAMY_MINACO_EMAIL'),
    },
    dropbox: {
      accessToken: read(env, 'DROPBOX_ACCESS_TOKEN'),
      rootPath: read(env, 'DROPBOX_ROOT_PATH', '/LONDON - ACCESS') || '/LONDON - ACCESS',
    },
    voice: {
      principalPhone: read(env, 'RAMY_PHONE_NUMBER'),
      model: normalizeVoiceModel(read(env, 'VOICE_MODEL', 'gpt-realtime-1.5')),
      voice: read(env, 'VOICE_NAME', 'marin'),
    },
    runtime: {
      healthSecret: read(env, 'HEALTH_SECRET'),
      stateFile: read(env, 'LONDON_STATE_FILE', '/tmp/london-lean-state.json'),
      pollIntervalMs: Math.max(15000, Number(read(env, 'MAIL_POLL_INTERVAL_MS', '60000')) || 60000),
      pollBatchSize: Math.min(25, Math.max(1, Number(read(env, 'MAIL_POLL_BATCH_SIZE', '10')) || 10)),
    },
  };

  return config;
}

export function configurationStatus(config) {
  return {
    openaiConfigured: Boolean(config.openai.apiKey),
    microsoftReadConfigured: Boolean(
      config.microsoft.readTenantId &&
      config.microsoft.readClientId &&
      config.microsoft.readClientSecret &&
      config.microsoft.londonMailbox
    ),
    microsoftActionsConfigured: Boolean(
      config.microsoft.actionTenantId &&
      config.microsoft.actionClientId &&
      config.microsoft.actionClientSecret
    ),
    dropboxConfigured: Boolean(config.dropbox.accessToken && config.dropbox.rootPath),
    voiceConfigured: Boolean(config.openai.apiKey && config.voice.principalPhone),
    voiceModel: config.voice.model,
    voiceName: config.voice.voice,
    stateFileConfigured: Boolean(config.runtime.stateFile),
  };
}
