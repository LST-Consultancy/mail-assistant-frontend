import type { Address, DigestThread, Email, GraphMessage } from "../types";
import { getGraphToken } from "./auth";
import { env } from "./env";
import { FIXTURE_MESSAGES, FIXTURE_SENT } from "./fixtures";

const GRAPH = "https://graph.microsoft.com/v1.0";
const SELECT =
  "id,subject,from,toRecipients,ccRecipients,receivedDateTime,conversationId,uniqueBody,bodyPreview,isRead,importance,webLink";
const TEXT_BODY = { Prefer: 'outlook.body-content-type="text"' };

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** fetch with Graph throttling handling: honors Retry-After on 429/503/504, max 3 retries. */
export async function graphFetch(
  url: string,
  init: RequestInit = {},
  deps: {
    token?: () => Promise<string>;
    fetch?: typeof fetch;
    sleep?: (ms: number) => Promise<void>;
  } = {}
): Promise<Response> {
  const f = deps.fetch ?? fetch;
  const wait = deps.sleep ?? sleep;
  const token = await (deps.token ?? getGraphToken)();
  for (let attempt = 0; ; attempt++) {
    const res = await f(url.startsWith("http") ? url : GRAPH + url, {
      ...init,
      headers: { ...init.headers, Authorization: `Bearer ${token}` },
    });
    if (![429, 503, 504].includes(res.status) || attempt >= 3) {
      if (!res.ok) throw new Error(`Microsoft Graph error ${res.status}`);
      return res;
    }
    const retryAfter = Number(res.headers.get("Retry-After"));
    await wait(retryAfter > 0 ? retryAfter * 1000 : 2 ** attempt * 1000);
  }
}

/** Follows @odata.nextLink until `max` items or the last page. */
export async function fetchPaged<T>(
  firstUrl: string,
  max: number,
  init: RequestInit,
  fetchPage = graphFetch
): Promise<T[]> {
  const out: T[] = [];
  let url: string | undefined = firstUrl;
  while (url && out.length < max) {
    const page = (await (await fetchPage(url, init)).json()) as {
      value: T[];
      "@odata.nextLink"?: string;
    };
    out.push(...page.value);
    url = page["@odata.nextLink"];
  }
  return out.slice(0, max);
}

export async function getMe(): Promise<{
  displayName: string;
  mail: string | null;
  userPrincipalName: string;
}> {
  if (env.mock)
    return { displayName: "Sam Me", mail: "me@contoso.com", userPrincipalName: "me@contoso.com" };
  return (await graphFetch("/me?$select=displayName,mail,userPrincipalName")).json();
}

/** All messages of a conversation, oldest first. Sorted client-side: $orderby with this $filter hits InefficientFilter. */
export async function getThread(conversationId: string): Promise<GraphMessage[]> {
  if (env.mock)
    return FIXTURE_MESSAGES.filter((m) => m.conversationId === conversationId).sort(byDateAsc);
  const filter = encodeURIComponent(`conversationId eq '${conversationId.replace(/'/g, "''")}'`);
  const list = await fetchPaged<GraphMessage>(
    `/me/messages?$filter=${filter}&$select=${SELECT}&$top=50`,
    50,
    {
      headers: TEXT_BODY,
    }
  );
  return list.sort(byDateAsc);
}

const byDateAsc = (a: GraphMessage, b: GraphMessage) =>
  a.receivedDateTime.localeCompare(b.receivedDateTime);

export async function listInbox(days: number, max: number): Promise<GraphMessage[]> {
  if (env.mock) return FIXTURE_MESSAGES;
  const since = new Date(Date.now() - days * 86_400_000).toISOString();
  const url =
    `/me/mailFolders/inbox/messages?$filter=receivedDateTime ge ${since}` +
    `&$orderby=receivedDateTime desc&$select=${SELECT}&$top=50`;
  return fetchPaged<GraphMessage>(url, max, { headers: TEXT_BODY });
}

/** Sent items since a date, only what's needed to detect "already replied". */
export async function listSent(days: number): Promise<GraphMessage[]> {
  if (env.mock) return FIXTURE_SENT;
  const since = new Date(Date.now() - days * 86_400_000).toISOString();
  const url =
    `/me/mailFolders/sentitems/messages?$filter=sentDateTime ge ${since}` +
    `&$orderby=sentDateTime desc&$select=conversationId,sentDateTime,receivedDateTime&$top=100`;
  return fetchPaged<GraphMessage>(url, 1000, {});
}

/** Saves a reply draft in Drafts. Never sends. */
export async function createReplyDraft(messageId: string, comment: string): Promise<void> {
  if (env.mock) return void (await sleep(300));
  await graphFetch(`/me/messages/${encodeURIComponent(messageId)}/createReply`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ comment }),
  });
}

// ---- Pure helpers (unit-tested) ----

const addr = (r?: { emailAddress: Address }): Address => ({
  name: r?.emailAddress?.name ?? "",
  address: r?.emailAddress?.address ?? "",
});

export function toEmail(
  m: GraphMessage,
  body = m.uniqueBody?.content ?? m.bodyPreview ?? ""
): Email {
  return {
    id: m.id,
    subject: m.subject ?? "",
    from: addr(m.from),
    to: (m.toRecipients ?? []).map(addr),
    cc: (m.ccRecipients ?? []).map(addr),
    receivedDateTime: m.receivedDateTime,
    body: body.trim().slice(0, env.maxBodyChars),
    importance: m.importance,
  };
}

/**
 * Groups inbox messages by conversation and keeps one entry per thread, keyed on the latest
 * message not sent by the user. The body is the thread's uniqueBody parts (no quoted
 * duplicates), oldest first, trimmed from the start so the latest message always fits.
 * Threads the user already answered after the latest incoming message are dropped.
 */
export function groupThreads(
  inbox: GraphMessage[],
  me: string,
  sent: GraphMessage[] = [],
  maxBodyChars = env.maxBodyChars
): DigestThread[] {
  const meLc = me.toLowerCase();
  const lastSent = new Map<string, string>();
  for (const s of sent) {
    const t = s.sentDateTime ?? s.receivedDateTime;
    if (t > (lastSent.get(s.conversationId) ?? "")) lastSent.set(s.conversationId, t);
  }

  const groups = new Map<string, GraphMessage[]>();
  for (const m of inbox) {
    if (m.from?.emailAddress?.address?.toLowerCase() === meLc) continue;
    const g = groups.get(m.conversationId);
    if (g) g.push(m);
    else groups.set(m.conversationId, [m]);
  }

  const out: DigestThread[] = [];
  for (const [conversationId, msgs] of groups) {
    msgs.sort(byDateAsc);
    const latest = msgs[msgs.length - 1];
    if ((lastSent.get(conversationId) ?? "") > latest.receivedDateTime) continue;
    const parts = msgs.map(
      (m) =>
        `--- ${addr(m.from).name || addr(m.from).address} (${m.receivedDateTime}):\n${(m.uniqueBody?.content ?? m.bodyPreview ?? "").trim()}`
    );
    const joined = parts.join("\n\n");
    const body = joined.length > maxBodyChars ? "…" + joined.slice(-maxBodyChars + 1) : joined;
    out.push({
      conversationId,
      latest,
      messageCount: msgs.length,
      email: { ...toEmail(latest), body },
    });
  }
  return out.sort((a, b) => b.latest.receivedDateTime.localeCompare(a.latest.receivedDateTime));
}
