/* global Office */
import type { Address, Email, Settings } from "../types";
import { env } from "./env";
import { FIXTURE_MESSAGES, fixtureToEmail } from "./fixtures";

type AnyItem = Office.MessageRead & Office.MessageCompose;

function call<T>(fn: (cb: (r: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise((resolve, reject) =>
    fn((r) =>
      r.status === Office.AsyncResultStatus.Succeeded
        ? resolve(r.value)
        : reject(new Error(r.error.message))
    )
  );
}

const item = () => Office.context.mailbox?.item as AnyItem | null | undefined;

export type Mode = "read" | "compose" | "none";

export function getMode(): Mode {
  if (env.mock)
    return new URLSearchParams(location.search).get("tab") === "write" ? "compose" : "read";
  const it = item();
  if (!it) return "none";
  return typeof it.displayReplyForm === "function" ? "read" : "compose";
}

const addr = (d?: Office.EmailAddressDetails): Address => ({
  name: d?.displayName ?? "",
  address: d?.emailAddress ?? "",
});

export function myAddress(): string {
  return env.mock ? "me@contoso.com" : (Office.context.mailbox?.userProfile.emailAddress ?? "");
}

export interface CurrentEmail {
  email: Email;
  conversationId: string;
}

/** The message open in read mode. email.id is the Graph (REST) id. */
export async function readCurrentEmail(): Promise<CurrentEmail> {
  if (env.mock) {
    const m = FIXTURE_MESSAGES[0];
    return { email: fixtureToEmail(m), conversationId: m.conversationId };
  }
  const it = item();
  if (!it || getMode() !== "read") throw new Error("Open or select an email to use this tab.");
  const body = await call<string>((cb) => it.body.getAsync(Office.CoercionType.Text, cb));
  const restId = Office.context.mailbox.convertToRestId(
    it.itemId,
    Office.MailboxEnums.RestVersion.v2_0
  );
  return {
    conversationId: it.conversationId,
    email: {
      id: restId,
      subject: it.subject ?? "",
      from: addr(it.from),
      to: (it.to ?? []).map(addr),
      cc: (it.cc ?? []).map(addr),
      receivedDateTime: it.dateTimeCreated?.toISOString() ?? "",
      body: body.slice(0, env.maxBodyChars),
    },
  };
}

/** Fires when the user selects another message while the pane is pinned. */
export function onItemChanged(handler: () => void): void {
  if (env.mock || !Office.context.mailbox) return;
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, handler);
}

export function textToHtml(text: string): string {
  const esc = text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  return esc
    .split(/\n{2,}/)
    .map((p) => `<p>${p.replace(/\n/g, "<br>")}</p>`)
    .join("");
}

/** Opens Outlook's native reply form prefilled; the user edits and sends it. */
export function openReplyForm(text: string, replyAll: boolean): void {
  const html = textToHtml(text).slice(0, 32_000); // displayReplyForm limit is 32 KB
  if (env.mock) return void alert(`[mock] ${replyAll ? "Reply all" : "Reply"} form:\n\n${text}`);
  const it = item();
  if (!it) throw new Error("No email selected.");
  if (replyAll) it.displayReplyAllForm(html);
  else it.displayReplyForm(html);
}

export function openMessage(restId: string, webLink?: string): void {
  if (env.mock) return void alert(`[mock] open message ${restId}`);
  try {
    const ewsId = Office.context.mailbox.convertToEwsId(
      restId,
      Office.MailboxEnums.RestVersion.v2_0
    );
    Office.context.mailbox.displayMessageForm(ewsId);
  } catch {
    if (webLink) window.open(webLink, "_blank", "noopener");
    else throw new Error("Couldn't open this email.");
  }
}

// ---- Compose ----

let mockBody =
  "Hi Ann,\n\nthanks for the update. i think we can do tuesday but not sure about the budget, lets talk\n\nSam";

export async function getComposeBody(): Promise<string> {
  if (env.mock) return mockBody;
  const it = item();
  if (!it) throw new Error("No draft open.");
  // HostConfig respects the user's setting and leaves out the quoted thread where supported (Mailbox 1.14).
  return call<string>((cb) =>
    it.body.getAsync(
      Office.CoercionType.Text,
      { bodyMode: Office.MailboxEnums.BodyMode?.HostConfig } as Office.AsyncContextOptions,
      cb
    )
  );
}

async function bodyCoercion(it: AnyItem): Promise<Office.CoercionType> {
  return call<Office.CoercionType>((cb) => it.body.getTypeAsync(cb));
}

export async function insertIntoBody(text: string, replace: boolean): Promise<void> {
  if (env.mock) {
    mockBody = replace ? text : `${mockBody}\n${text}`;
    return void alert(`[mock] ${replace ? "Body replaced" : "Inserted"}:\n\n${text}`);
  }
  const it = item();
  if (!it) throw new Error("No draft open.");
  const type = await bodyCoercion(it);
  const data = type === Office.CoercionType.Html ? textToHtml(text) : text;
  await call<void>((cb) =>
    replace
      ? it.body.setAsync(data, { coercionType: type }, cb)
      : it.body.setSelectedDataAsync(data, { coercionType: type }, cb)
  );
}

/** conversationId of the draft, if it's a reply (used to fetch the original as context). */
export function composeConversationId(): string | undefined {
  if (env.mock) return FIXTURE_MESSAGES[0].conversationId;
  return item()?.conversationId || undefined;
}

// ---- Settings (roaming, per mailbox) ----

export const DEFAULT_SETTINGS: Settings = {
  signatureName: "",
  defaultTone: "professional",
  language: "match",
  digestDays: 30,
  maxEmails: env.maxEmails,
  privacyAcknowledged: false,
};

const KEY = "settings";

export function loadSettings(): Settings {
  try {
    const raw = env.mock ? localStorage.getItem(KEY) : Office.context.roamingSettings.get(KEY);
    const saved = typeof raw === "string" ? JSON.parse(raw) : raw;
    return { ...DEFAULT_SETTINGS, ...(saved ?? {}) };
  } catch {
    return DEFAULT_SETTINGS;
  }
}

export async function saveSettings(s: Settings): Promise<void> {
  if (env.mock) {
    try {
      localStorage.setItem(KEY, JSON.stringify(s));
    } catch {
      // storage blocked: settings last for this session only
    }
    return;
  }
  Office.context.roamingSettings.set(KEY, s);
  await call<void>((cb) => Office.context.roamingSettings.saveAsync(cb));
}

/** Office theme first, OS preference as fallback. */
export function isDarkTheme(): boolean {
  const bg = (Office.context as { officeTheme?: Office.OfficeTheme })?.officeTheme
    ?.bodyBackgroundColor;
  if (bg && /^#?[0-9a-f]{6}$/i.test(bg)) {
    const n = parseInt(bg.replace("#", ""), 16);
    const lum = 0.299 * (n >> 16) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255);
    return lum < 128;
  }
  return window.matchMedia?.("(prefers-color-scheme: dark)").matches ?? false;
}
