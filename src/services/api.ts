import type { Analysis, DigestResult, Email, Settings, Tone } from "../types";
import { getApiToken } from "./auth";
import { env } from "./env";

type AiSettings = { signatureName?: string; language: string };
export const aiSettings = (s: Settings): AiSettings => ({
  signatureName: s.signatureName || undefined,
  language: s.language,
});

async function request(path: string, body: unknown, signal?: AbortSignal): Promise<Response> {
  const token = await getApiToken();
  const res = await fetch(`${env.apiBaseUrl}/api${path}`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
    body: JSON.stringify(body),
    signal,
  }).catch((e: unknown) => {
    if (e instanceof DOMException && e.name === "AbortError") throw e;
    throw new Error(
      "Couldn't reach the Mail Assistant service. Check your connection and try again."
    );
  });
  if (!res.ok) {
    const msg = await res
      .json()
      .then((j: { error?: string }) => j.error)
      .catch(() => undefined);
    throw new Error(msg ?? `Request failed (${res.status})`);
  }
  return res;
}

const postJson = async <T>(path: string, body: unknown, signal?: AbortSignal): Promise<T> =>
  (await request(path, body, signal)).json() as Promise<T>;

/** Reads `data: {"text"}` chunks until `event: done`; throws on `event: error`. */
export async function readSse(res: Response, onText: (t: string) => void): Promise<void> {
  const reader = res.body!.pipeThrough(new TextDecoderStream()).getReader();
  let buf = "";
  for (;;) {
    const { value, done } = await reader.read();
    if (done) break;
    buf += value;
    let i: number;
    while ((i = buf.indexOf("\n\n")) >= 0) {
      const frame = buf.slice(0, i);
      buf = buf.slice(i + 2);
      const event = frame.match(/^event: (.+)$/m)?.[1] ?? "message";
      const data = frame.match(/^data: (.*)$/m)?.[1];
      if (!data) continue;
      const parsed = JSON.parse(data) as { text?: string; error?: string };
      if (event === "error") throw new Error(parsed.error ?? "Generation failed.");
      if (event === "done") return;
      if (parsed.text) onText(parsed.text);
    }
  }
  throw new Error("The connection closed before the reply finished.");
}

export const analyzeEmail = (
  email: Email,
  thread: Email[] | undefined,
  s: Settings,
  signal?: AbortSignal
) => postJson<Analysis>("/analyze-email", { email, thread, settings: aiSettings(s) }, signal);

export async function streamReply(
  opts: {
    email: Email;
    thread?: Email[];
    tone: Tone;
    replyType: "reply" | "replyAll";
    instructions?: string;
  },
  s: Settings,
  onText: (t: string) => void,
  signal?: AbortSignal
): Promise<void> {
  return readSse(await request("/reply", { ...opts, settings: aiSettings(s) }, signal), onText);
}

export async function streamCompose(
  opts: {
    mode: "draft" | "improve";
    text?: string;
    instructions?: string;
    context?: Email;
    tone: Tone;
  },
  s: Settings,
  onText: (t: string) => void,
  signal?: AbortSignal
): Promise<void> {
  return readSse(await request("/compose", { ...opts, settings: aiSettings(s) }, signal), onText);
}

export type BatchResult = DigestResult | { id: string; error: string };

export const analyzeBatch = (threads: Email[], s: Settings, signal?: AbortSignal) =>
  postJson<BatchResult[]>("/digest/analyze-batch", { threads, settings: aiSettings(s) }, signal);

export const digestOverview = (
  items: {
    subject: string;
    from: string;
    summary: string;
    priority: string;
    category: string;
    needsReply: boolean;
  }[],
  signal?: AbortSignal
) => postJson<{ markdown: string }>("/digest/overview", { items }, signal);
