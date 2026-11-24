import type { Address, DigestThread, Email, GraphMessage } from "../types";
import { getGraphToken } from "./auth";
import { env } from "./env";
import { FIXTURE_MESSAGES, FIXTURE_SENT } from "./fixtures";

const GRAPH = "https://graph.microsoft.com/v1.0";
const SELECT =
  "id,subject,from,toRecipients,ccRecipients,receivedDateTime,conversationId,uniqueBody,bodyPreview,isRead,importance,webLink,categories";
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

/** Sent items since a date: used for "already replied" detection and sent-mail follow-ups. */
export async function listSent(days: number, max: number): Promise<GraphMessage[]> {
  if (env.mock) return FIXTURE_SENT;
  const since = new Date(Date.now() - days * 86_400_000).toISOString();
  const url =
    `/me/mailFolders/sentitems/messages?$filter=sentDateTime ge ${since}` +
    `&$orderby=sentDateTime desc&$select=${SELECT},sentDateTime&$top=50`;
  return fetchPaged<GraphMessage>(url, max, { headers: TEXT_BODY });
}

/**
 * Saves a reply draft in Drafts. Never sends. Pass `to` for a follow-up on the user's own sent
 * email, so the draft goes to its original recipients instead of back to the user.
 */
export async function createReplyDraft(
  messageId: string,
  comment: string,
  to?: Pick<GraphMessage, "toRecipients" | "ccRecipients">
): Promise<void> {
  if (env.mock) return void (await sleep(300));
  const message = to && {
    toRecipients: to.toRecipients ?? [],
    ccRecipients: to.ccRecipients ?? [],
  };
  await graphFetch(`/me/messages/${encodeURIComponent(messageId)}/createReply`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(message ? { comment, message } : { comment }),
  });
}

/** Outlook category that marks a conversation as ignored. It lives in the mailbox, so it is permanent and per user. */
export const IGNORED_CATEGORY = "Mail Assistant: Ignored";

/**
 * Tags (or untags) a message with IGNORED_CATEGORY; the digest then skips its whole conversation.
 * shortcut: only the latest message is tagged, so a conversation that revives after that message
 * falls out of the digest range shows up again; tag every message in the thread if users hit this.
 */
export async function setIgnored(m: GraphMessage, ignored: boolean): Promise<void> {
  const rest = (m.categories ?? []).filter((c) => c !== IGNORED_CATEGORY);
  const categories = ignored ? [...rest, IGNORED_CATEGORY] : rest;
  if (!env.mock)
    await graphFetch(`/me/messages/${encodeURIComponent(m.id)}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ categories }),
    });
  m.categories = categories;
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

const isIgnored = (m: GraphMessage) => !!m.categories?.includes(IGNORED_CATEGORY);
const ignoredConversations = (msgs: GraphMessage[]) =>
  new Set(msgs.filter(isIgnored).map((m) => m.conversationId));

const REPLY_PREFIX = /^\s*(re|fw|fwd|aw|wg|sv)\s*:/i;
const OTP =
  /\b(otp|one[- ]time (pass(word|code)?|code|pin)|verification code|security code|login code|sign[- ]?in code|passcode|2fa code|authentication code|confirmation code)\b/i;
const AUTO =
  /^\s*(automatic reply|auto[- ]?reply|out of (the )?office|undeliverable|delivery status notification|mail delivery (failed|subsystem)|read:|not read:|recall:)/i;

/**
 * True for mail the assistant should never show: OTP/verification codes, auto-replies, bounces,
 * read receipts, and calendar items (meeting requests/responses are not emails to act on).
 * Spam itself lives in Junk, which the digest never reads; the AI tags what slips through.
 */
export function isNoise(m: GraphMessage): boolean {
  if (m["@odata.type"]?.includes("eventMessage")) return true;
  const subject = m.subject ?? "";
  if (AUTO.test(subject)) return true;
  if (REPLY_PREFIX.test(subject)) return false; // real OTP mails are never replies
  const text = `${subject}\n${m.bodyPreview ?? ""}`;
  return OTP.test(text) && /\b\d{4,8}\b/.test(text); // keyword alone is too broad; real OTP mails carry the code
}

/** Thread body from uniqueBody parts (no quoted duplicates), oldest first, trimmed from the start. */
function threadBody(msgs: GraphMessage[], maxBodyChars: number): string {
  const joined = msgs
    .map(
      (m) =>
        `--- ${addr(m.from).name || addr(m.from).address} (${m.sentDateTime ?? m.receivedDateTime}):\n${(m.uniqueBody?.content ?? m.bodyPreview ?? "").trim()}`
    )
    .join("\n\n");
  return joined.length > maxBodyChars ? "…" + joined.slice(-maxBodyChars + 1) : joined;
}

const sentAt = (m: GraphMessage) => m.sentDateTime ?? m.receivedDateTime;

/** Latest time per conversation. */
function latestBy(msgs: GraphMessage[], time: (m: GraphMessage) => string): Map<string, string> {
  const out = new Map<string, string>();
  for (const m of msgs)
    if (time(m) > (out.get(m.conversationId) ?? "")) out.set(m.conversationId, time(m));
  return out;
}

/**
 * Groups inbox messages by conversation and keeps one entry per thread, keyed on the latest
 * message not sent by the user. The body is the thread's uniqueBody parts (no quoted
 * duplicates), oldest first, trimmed from the start so the latest message always fits.
 * Threads the user already answered after the latest incoming message are dropped, and so are
 * ignored conversations and noise (see isNoise).
 */
export function groupThreads(
  inbox: GraphMessage[],
  me: string,
  sent: GraphMessage[] = [],
  maxBodyChars = env.maxBodyChars
): DigestThread[] {
  const meLc = me.toLowerCase();
  const lastSent = latestBy(sent, sentAt);
  const ignored = ignoredConversations([...inbox, ...sent]);

  const groups = new Map<string, GraphMessage[]>();
  for (const m of inbox) {
    if (m.from?.emailAddress?.address?.toLowerCase() === meLc) continue;
    if (ignored.has(m.conversationId) || isNoise(m)) continue;
    const g = groups.get(m.conversationId);
    if (g) g.push(m);
    else groups.set(m.conversationId, [m]);
  }

  const out: DigestThread[] = [];
  for (const [conversationId, msgs] of groups) {
    msgs.sort(byDateAsc);
    const latest = msgs[msgs.length - 1];
    if ((lastSent.get(conversationId) ?? "") > latest.receivedDateTime) continue;
    const body = threadBody(msgs, maxBodyChars);
    out.push({
      kind: "inbox",
      conversationId,
      latest,
      messageCount: msgs.length,
      email: { ...toEmail(latest), body },
    });
  }
  return out.sort((a, b) => b.latest.receivedDateTime.localeCompare(a.latest.receivedDateTime));
}

/**
 * Sent-mail follow-ups: one entry per conversation where the user's latest sent email has had no
 * incoming answer since. Whether it actually needs a follow-up is left to the AI.
 */
export function followUpThreads(
  sent: GraphMessage[],
  inbox: GraphMessage[],
  me: string,
  maxBodyChars = env.maxBodyChars
): DigestThread[] {
  const meLc = me.toLowerCase();
  const lastIncoming = latestBy(
    inbox.filter((m) => m.from?.emailAddress?.address?.toLowerCase() !== meLc),
    (m) => m.receivedDateTime
  );
  const ignored = ignoredConversations([...inbox, ...sent]);
  const toOthers = (m: GraphMessage) =>
    [...(m.toRecipients ?? []), ...(m.ccRecipients ?? [])].some(
      (r) => r.emailAddress?.address && r.emailAddress.address.toLowerCase() !== meLc
    );

  const groups = new Map<string, GraphMessage[]>();
  for (const m of sent) {
    if (ignored.has(m.conversationId) || isNoise(m)) continue;
    const g = groups.get(m.conversationId);
    if (g) g.push(m);
    else groups.set(m.conversationId, [m]);
  }

  const out: DigestThread[] = [];
  for (const [conversationId, msgs] of groups) {
    msgs.sort((a, b) => sentAt(a).localeCompare(sentAt(b)));
    const latest = msgs[msgs.length - 1];
    if (!toOthers(latest) || (lastIncoming.get(conversationId) ?? "") >= sentAt(latest)) continue;
    out.push({
      kind: "sent",
      conversationId,
      latest,
      messageCount: msgs.length,
      email: { ...toEmail(latest), body: threadBody(msgs, maxBodyChars) },
    });
  }
  return out.sort((a, b) => sentAt(b.latest).localeCompare(sentAt(a.latest)));
}
