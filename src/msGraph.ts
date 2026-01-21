import { CalendarEvent, UserOAuthAccount } from './types';
import { updateUserOAuthAccountTokens } from './db';

const graphConfig = {
  tenantId: process.env.MS_TENANT_ID,
  clientId: process.env.MS_CLIENT_ID,
  clientSecret: process.env.MS_CLIENT_SECRET,
};

type TokenCache = { accessToken: string; expiresAt: number } | null;
let tokenCache: TokenCache = null;

async function getGraphToken(logger?: any): Promise<string> {
  if (!graphConfig.clientId || !graphConfig.clientSecret || !graphConfig.tenantId) {
    throw new Error('Microsoft Graph credentials are missing');
  }

  const now = Date.now();
  if (tokenCache && tokenCache.expiresAt > now + 60_000) {
    return tokenCache.accessToken;
  }

  const params = new URLSearchParams({
    client_id: graphConfig.clientId,
    client_secret: graphConfig.clientSecret,
    grant_type: 'client_credentials',
    scope: 'https://graph.microsoft.com/.default',
  });

  const res = await fetch(
    `https://login.microsoftonline.com/${graphConfig.tenantId}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params,
    },
  );

  const data = (await res.json()) as { access_token?: string; expires_in?: number; error?: string };
  if (!res.ok || !data.access_token) {
    logger?.error({ data }, 'graph-token-failed');
    throw new Error(`Failed to fetch Microsoft Graph token: ${data.error ?? res.statusText}`);
  }

  tokenCache = {
    accessToken: data.access_token,
    expiresAt: Date.now() + (data.expires_in ?? 3500) * 1000,
  };
  return tokenCache.accessToken;
}

function ensureDate(input: string): Date {
  const dt = new Date(input);
  if (Number.isNaN(dt.getTime())) {
    throw new Error('Invalid date range');
  }
  return dt;
}

function startOfWeekUtc(date: Date): Date {
  const copy = new Date(date);
  const day = copy.getUTCDay();
  const diff = copy.getUTCDate() - day;
  copy.setUTCDate(diff);
  copy.setUTCHours(0, 0, 0, 0);
  return copy;
}

function toIso(date: Date, dayOffset: number, hour: number, minute: number, durationMinutes: number) {
  const start = new Date(date);
  start.setUTCDate(start.getUTCDate() + dayOffset);
  start.setUTCHours(hour, minute, 0, 0);
  const end = new Date(start);
  end.setUTCMinutes(end.getUTCMinutes() + durationMinutes);
  return { start: start.toISOString(), end: end.toISOString() };
}

function buildSampleEvents(rangeStart: string): CalendarEvent[] {
  const base = startOfWeekUtc(ensureDate(rangeStart));
  const template: Array<{
    day: number;
    hour: number;
    minute: number;
    duration: number;
    title: string;
    location?: string;
  }> = [
    { day: 1, hour: 15, minute: 0, duration: 60, title: 'Interview - CRM - Senior', location: 'Boardroom A' },
    { day: 1, hour: 19, minute: 0, duration: 30, title: 'Interview', location: 'Teams call' },
    { day: 2, hour: 14, minute: 0, duration: 45, title: 'Interview with Data', location: 'Virtual room' },
    { day: 3, hour: 10, minute: 30, duration: 60, title: 'Interview - QA & Senior', location: 'HQ - Blue' },
    { day: 3, hour: 15, minute: 0, duration: 60, title: 'Interview', location: 'Teams call' },
    { day: 4, hour: 12, minute: 0, duration: 60, title: 'Interview with Product', location: 'Zoom' },
    { day: 4, hour: 15, minute: 30, duration: 60, title: 'Interview with Data', location: 'Zoom' },
    { day: 4, hour: 17, minute: 30, duration: 60, title: 'Interview with Executive', location: 'Office' },
    { day: 5, hour: 14, minute: 30, duration: 30, title: 'Interview', location: 'Office' },
    { day: 6, hour: 16, minute: 0, duration: 45, title: 'Interview with Brandon', location: 'Virtual' },
  ];

  return template.map((item, idx) => {
    const times = toIso(base, item.day, item.hour, item.minute, item.duration);
    return {
      id: `sample-${idx}`,
      title: item.title,
      start: times.start,
      end: times.end,
      location: item.location,
      organizer: 'Scheduling bot',
    };
  });
}

export async function refreshAzureADToken(
  account: UserOAuthAccount,
): Promise<UserOAuthAccount> {
  if (!account.refreshToken) {
    throw new Error('No refresh token available');
  }
  if (!graphConfig.clientId || !graphConfig.clientSecret || !graphConfig.tenantId) {
    throw new Error('Microsoft Graph credentials are missing');
  }

  const tenantId = graphConfig.tenantId || 'common';
  const baseScope = 'openid profile email offline_access Calendars.Read Mail.Read Mail.Send User.Read';
  const includeSharedCalendars =
    process.env.MS_GRAPH_SHARED_CALENDARS === 'true' ||
    (tenantId !== 'common' && tenantId !== 'consumers');
  const scope = includeSharedCalendars
    ? `${baseScope} Calendars.Read.Shared Mail.Read.Shared Mail.Send.Shared`
    : baseScope;

  const params = new URLSearchParams({
    client_id: graphConfig.clientId,
    client_secret: graphConfig.clientSecret,
    grant_type: 'refresh_token',
    refresh_token: account.refreshToken,
    scope,
  });

  const res = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params,
    },
  );

  const data = (await res.json()) as {
    access_token?: string;
    refresh_token?: string;
    expires_in?: number;
    error_description?: string;
  };

  if (!res.ok || !data.access_token) {
    throw new Error(
      data?.error_description || 'Failed to refresh access token',
    );
  }

  const expiresAt = data.expires_in
    ? Math.floor(Date.now() / 1000) + data.expires_in
    : account.expiresAt ?? undefined;

  const updated = await updateUserOAuthAccountTokens(account.id, {
    accessToken: data.access_token,
    refreshToken: data.refresh_token ?? account.refreshToken ?? null,
    expiresAt: expiresAt ?? null,
  });

  if (!updated) {
    throw new Error('Failed to update tokens in database');
  }

  return updated;
}

export async function ensureFreshToken(
  account: UserOAuthAccount,
): Promise<UserOAuthAccount> {
  const expiresAt = account.expiresAt ? account.expiresAt * 1000 : null;
  if (account.accessToken && expiresAt && Date.now() < expiresAt - 60_000) {
    return account;
  }
  if (!account.refreshToken) {
    return account;
  }
  try {
    return await refreshAzureADToken(account);
  } catch (err) {
    console.error('Failed to refresh token:', err);
    return account;
  }
}

export async function loadOutlookEvents(params: {
  email: string;
  rangeStart: string;
  rangeEnd: string;
  timezone?: string | null;
  logger?: any;
  account?: UserOAuthAccount | null;
}): Promise<{ events: CalendarEvent[]; source: 'graph' | 'sample'; warning?: string }> {
  const { email, rangeStart, rangeEnd, timezone, logger, account } = params;
  const tz = timezone || 'UTC';

  try {
    let token: string;
    let useMeEndpoint = false;
    if (account) {
      const freshAccount = await ensureFreshToken(account);
      if (!freshAccount.accessToken) {
        throw new Error('No access token available');
      }
      token = freshAccount.accessToken;
      // Use /me endpoint if the account email matches the requested email (token belongs to this user)
      if (account.email.toLowerCase() === email.toLowerCase()) {
        useMeEndpoint = true;
      }
    } else {
      token = await getGraphToken(logger);
    }

    // Use /me/calendarView for personal accounts, /users/{email}/calendarView for shared/other accounts
    const endpoint = useMeEndpoint
      ? 'https://graph.microsoft.com/v1.0/me/calendarView'
      : `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(email)}/calendarView`;
    
    const url = new URL(endpoint);
    url.searchParams.set('startDateTime', rangeStart);
    url.searchParams.set('endDateTime', rangeEnd);
    url.searchParams.set('$select', 'subject,start,end,location,isAllDay,organizer');

    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Prefer: `outlook.timezone="${tz}"`,
      },
    });

    const data = (await res.json()) as any;
    if (!res.ok) {
      logger?.error({ status: res.status, data }, 'graph-events-failed');
      throw new Error(data?.error?.message || 'Failed to fetch events from Graph');
    }

    const events: CalendarEvent[] = Array.isArray(data?.value)
      ? data.value
          .map((ev: any) => ({
            id: ev.id as string,
            title: (ev.subject as string) || 'Busy',
            start: ev.start?.dateTime as string,
            end: ev.end?.dateTime as string,
            isAllDay: Boolean(ev.isAllDay),
            organizer:
              ev.organizer?.emailAddress?.name ||
              ev.organizer?.emailAddress?.address ||
              undefined,
            location: ev.location?.displayName as string | undefined,
          }))
          .filter((ev: CalendarEvent) => Boolean(ev.start) && Boolean(ev.end))
      : [];

    return { events, source: 'graph' };
  } catch (err) {
    logger?.warn({ err }, 'graph-events-fallback');
    return {
      events: buildSampleEvents(rangeStart),
      source: 'sample',
      warning:
        err instanceof Error ? err.message : 'Falling back to sample events due to an unknown error',
    };
  }
}

export async function loadOutlookMessages(params: {
  mailbox: string;
  account: UserOAuthAccount;
  maxItems?: number;
  since?: string | null;
  folder?: 'inbox' | 'sent';
  logger?: any;
}): Promise<{
  messages: Array<{
    id: string;
    internetMessageId?: string | null;
    subject?: string | null;
    from?: { emailAddress?: { address?: string; name?: string } } | null;
    toRecipients?: Array<{ emailAddress?: { address?: string; name?: string } }> | null;
    isRead?: boolean | null;
    receivedDateTime?: string | null;
    bodyPreview?: string | null;
    webLink?: string | null;
    body?: { content?: string | null; contentType?: string | null } | null;
  }>;
  source: 'graph' | 'empty';
  warning?: string;
}> {
  const { mailbox, account, maxItems = 50, since, folder = 'inbox', logger } = params;
  const doFetch = async (accessToken: string) => {
    const useMe = account.email.toLowerCase() === mailbox.toLowerCase();
    const basePath = folder === 'sent' ? 'sentItems' : 'messages';
    const endpoint = useMe
      ? `https://graph.microsoft.com/v1.0/me/${basePath}`
      : `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/${basePath}`;

    const url = new URL(endpoint);
    url.searchParams.set('$top', String(maxItems));
    url.searchParams.set(
      '$select',
      'id,subject,from,toRecipients,receivedDateTime,isRead,bodyPreview,webLink,internetMessageId,body',
    );
    url.searchParams.set('$orderby', 'receivedDateTime DESC');
    if (since) {
      url.searchParams.set('$filter', `receivedDateTime ge ${since}`);
    }

    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Prefer: 'outlook.body-content-type="text"',
      },
    });
    return res;
  };

  try {
    let freshAccount = await ensureFreshToken(account);
    if (!freshAccount.accessToken) {
      throw new Error('No access token available');
    }
    let res = await doFetch(freshAccount.accessToken);

    // Retry once on invalid/expired token
    if (!res.ok && res.status === 401) {
      try {
        freshAccount = await refreshAzureADToken(freshAccount);
        if (freshAccount.accessToken) {
          res = await doFetch(freshAccount.accessToken);
        }
      } catch (refreshErr) {
        logger?.warn({ refreshErr }, 'graph-mail-refresh-failed');
      }
    }

    const data = (await res.json()) as any;
    if (!res.ok) {
      logger?.error({ status: res.status, data }, 'graph-mail-failed');
      throw new Error(data?.error?.message || 'Failed to fetch mail from Graph');
    }

    const messages: any[] = Array.isArray(data?.value) ? data.value : [];
    return { messages, source: 'graph' };
  } catch (err) {
    logger?.warn({ err }, 'graph-mail-fallback');
    return {
      messages: [],
      source: 'empty',
      warning:
        err instanceof Error ? err.message : 'Unable to fetch mail from Graph',
    };
  }
}

export async function sendOutlookMail(params: {
  account: UserOAuthAccount;
  mailbox: string;
  to: Array<{ email: string; name?: string | null }>;
  subject?: string | null;
  body?: string | null;
  bodyContentType?: 'Text' | 'HTML';
  logger?: any;
}) {
  const { account, mailbox, to, subject, body, bodyContentType = 'HTML', logger } = params;
  const freshAccount = await ensureFreshToken(account);
  if (!freshAccount.accessToken) {
    throw new Error('No access token available');
  }
  const useMe = freshAccount.email.toLowerCase() === mailbox.toLowerCase();
  const endpoint = useMe
    ? 'https://graph.microsoft.com/v1.0/me/sendMail'
    : `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/sendMail`;

  const payload = {
    message: {
      subject: subject || '',
      body: {
        contentType: bodyContentType,
        content: body || '',
      },
      toRecipients: to.map((rec) => ({
        emailAddress: {
          address: rec.email,
          name: rec.name || undefined,
        },
      })),
    },
    saveToSentItems: true,
  };
  const doSend = async (accessToken: string) => {
    return fetch(endpoint, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload),
    });
  };

  let res = await doSend(freshAccount.accessToken);
  if (!res.ok && res.status === 401) {
    try {
      const refreshed = await refreshAzureADToken(freshAccount);
      if (refreshed.accessToken) {
        res = await doSend(refreshed.accessToken);
      }
    } catch (refreshErr) {
      logger?.warn({ refreshErr }, 'graph-send-refresh-failed');
    }
  }

  if (!res.ok) {
    const text = await res.text();
    logger?.error({ status: res.status, text }, 'graph-send-mail-failed');
    throw new Error(text || 'Failed to send mail');
  }
}
