import { describe, expect, it, vi } from "vitest";
import type { GraphMessage } from "../types";
import { readSse } from "./api";
import { FIXTURE_MESSAGES, FIXTURE_SENT } from "./fixtures";
import {
  fetchPaged,
  followUpThreads,
  graphFetch,
  groupThreads,
  IGNORED_CATEGORY,
  isNoise,
} from "./graph";

const json = (body: unknown, status = 200, headers: Record<string, string> = {}) =>
  new Response(JSON.stringify(body), { status, headers });

describe("fetchPaged", () => {
  it("follows nextLink and stops at max", async () => {
    const pages: Record<string, unknown> = {
      first: { value: [1, 2], "@odata.nextLink": "second" },
      second: { value: [3, 4], "@odata.nextLink": "third" },
      third: { value: [5] },
    };
    const seen: string[] = [];
    const page = async (url: string) => (seen.push(url), json(pages[url]));
    expect(await fetchPaged<number>("first", 100, {}, page)).toEqual([1, 2, 3, 4, 5]);
    seen.length = 0;
    expect(await fetchPaged<number>("first", 3, {}, page)).toEqual([1, 2, 3]);
    expect(seen).toEqual(["first", "second"]);
  });
});

describe("graphFetch", () => {
  const token = async () => "t";
  it("honors Retry-After on 429 then succeeds", async () => {
    const fetch = vi
      .fn()
      .mockResolvedValueOnce(json({}, 429, { "Retry-After": "2" }))
      .mockResolvedValueOnce(json({ ok: 1 }));
    const sleep = vi.fn(async () => {});
    const res = await graphFetch("/me", {}, { token, fetch, sleep });
    expect(await res.json()).toEqual({ ok: 1 });
    expect(sleep).toHaveBeenCalledWith(2000);
    expect(fetch.mock.calls[0][0]).toBe("https://graph.microsoft.com/v1.0/me");
    expect(fetch.mock.calls[0][1].headers.Authorization).toBe("Bearer t");
  });
  it("gives up after 3 retries with exponential backoff", async () => {
    const fetch = vi.fn(async () => json({}, 503));
    const sleep = vi.fn(async () => {});
    await expect(graphFetch("/me", {}, { token, fetch, sleep })).rejects.toThrow("503");
    expect(fetch).toHaveBeenCalledTimes(4);
    expect(sleep.mock.calls.map((c) => (c as unknown[])[0])).toEqual([1000, 2000, 4000]);
  });
  it("does not retry other errors", async () => {
    const fetch = vi.fn(async () => json({}, 403));
    await expect(graphFetch("/me", {}, { token, fetch, sleep: async () => {} })).rejects.toThrow(
      "403"
    );
    expect(fetch).toHaveBeenCalledTimes(1);
  });
});

describe("groupThreads", () => {
  const threads = groupThreads(FIXTURE_MESSAGES, "ME@contoso.com", FIXTURE_SENT);
  it("keeps one entry per conversation, keyed on the latest message", () => {
    const c1 = threads.find((t) => t.conversationId === "c1")!;
    expect(c1.latest.id).toBe("m1");
    expect(c1.messageCount).toBe(2);
    expect(new Set(threads.map((t) => t.conversationId)).size).toBe(threads.length);
  });
  it("builds the body from uniqueBody parts, oldest first, without quoted duplicates", () => {
    const body = threads.find((t) => t.conversationId === "c1")!.email.body;
    expect(body.indexOf("attaching the draft")).toBeLessThan(body.indexOf("Following up"));
    expect(body.match(/Following up/g)).toHaveLength(1);
  });
  it("drops threads the user already answered after the latest incoming message", () => {
    expect(threads.some((t) => t.conversationId === "c6")).toBe(false);
  });
  it("keeps a thread when the user's reply is older than the latest incoming", () => {
    const sent = [
      { ...FIXTURE_SENT[0], conversationId: "c1", sentDateTime: "2000-01-01T00:00:00Z" },
    ];
    expect(
      groupThreads(FIXTURE_MESSAGES, "me@contoso.com", sent).some((t) => t.conversationId === "c1")
    ).toBe(true);
  });
  it("skips messages sent by the user and sorts newest first", () => {
    const mine: GraphMessage = {
      ...FIXTURE_MESSAGES[0],
      id: "x",
      conversationId: "cx",
      from: { emailAddress: { name: "Me", address: "me@contoso.com" } },
    };
    const out = groupThreads([mine, ...FIXTURE_MESSAGES], "me@contoso.com");
    expect(out.some((t) => t.conversationId === "cx")).toBe(false);
    const dates = out.map((t) => t.latest.receivedDateTime);
    expect([...dates].sort().reverse()).toEqual(dates);
  });
  it("trims long threads from the start so the latest message survives", () => {
    const long = FIXTURE_MESSAGES.map((m) =>
      m.conversationId === "c1" ? { ...m, uniqueBody: { content: m.id + "x".repeat(500) } } : m
    );
    const body = groupThreads(long, "me@contoso.com", [], 300).find(
      (t) => t.conversationId === "c1"
    )!.email.body;
    expect(body.length).toBe(300);
    expect(body.startsWith("…")).toBe(true);
    expect(body).not.toContain("m5x");
  });
});

describe("noise, ignore and follow-ups", () => {
  const base = FIXTURE_MESSAGES[0];
  const m = (subject: string, extra: Partial<GraphMessage> = {}): GraphMessage => ({
    ...base,
    subject,
    ...extra,
  });
  it("isNoise drops OTPs, auto-replies and calendar items but keeps real mail", () => {
    expect(isNoise(m("Your one-time password", { bodyPreview: "Use 482913 to sign in." }))).toBe(
      true
    );
    expect(isNoise(m("Your one-time password"))).toBe(false); // no code: left to the AI
    expect(isNoise(m("Sign in", { bodyPreview: "Your verification code is 123456" }))).toBe(true);
    expect(isNoise(m("Automatic reply: Q4 plan"))).toBe(true);
    expect(isNoise(m("Undeliverable: hello"))).toBe(true);
    expect(isNoise(m("Team sync", { "@odata.type": "#microsoft.graph.eventMessageRequest" }))).toBe(
      true
    );
    expect(isNoise(m("Re: OTP feature design"))).toBe(false);
    expect(
      isNoise(
        m("Need the verification code process documented", { bodyPreview: "Can you write it up?" })
      )
    ).toBe(false);
    expect(isNoise(base)).toBe(false);
  });
  it("the OTP fixture never reaches the digest", () => {
    expect(
      groupThreads(FIXTURE_MESSAGES, "me@contoso.com").some((t) => t.conversationId === "c8")
    ).toBe(false);
  });
  it("an ignored message hides its whole conversation from inbox and follow-ups", () => {
    const tagged = FIXTURE_MESSAGES.map((x) =>
      x.id === "m5" ? { ...x, categories: [IGNORED_CATEGORY] } : x
    );
    expect(groupThreads(tagged, "me@contoso.com").some((t) => t.conversationId === "c1")).toBe(
      false
    );
    const sent = FIXTURE_SENT.map((x) => ({ ...x, categories: [IGNORED_CATEGORY] }));
    expect(followUpThreads(sent, FIXTURE_MESSAGES, "me@contoso.com")).toEqual([]);
  });
  it("follow-ups are sent mails with no answer since, addressed to someone else", () => {
    const out = followUpThreads(FIXTURE_SENT, FIXTURE_MESSAGES, "me@contoso.com");
    expect(out.map((t) => [t.conversationId, t.kind])).toEqual([
      ["c7", "sent"],
      ["c6", "sent"],
    ]);
    const answered = [
      ...FIXTURE_MESSAGES,
      { ...base, id: "a", conversationId: "c7", receivedDateTime: new Date().toISOString() },
    ];
    expect(
      followUpThreads(FIXTURE_SENT, answered, "me@contoso.com").some(
        (t) => t.conversationId === "c7"
      )
    ).toBe(false);
    const toSelf = FIXTURE_SENT.map((x) => ({
      ...x,
      toRecipients: [{ emailAddress: { name: "", address: "ME@contoso.com" } }],
    }));
    expect(followUpThreads(toSelf, FIXTURE_MESSAGES, "me@contoso.com")).toEqual([]);
  });
  it("a conversation is either an inbox thread or a follow-up, never both", () => {
    const inbox = groupThreads(FIXTURE_MESSAGES, "me@contoso.com", FIXTURE_SENT).map(
      (t) => t.conversationId
    );
    const sent = followUpThreads(FIXTURE_SENT, FIXTURE_MESSAGES, "me@contoso.com").map(
      (t) => t.conversationId
    );
    expect(inbox.filter((c) => sent.includes(c))).toEqual([]);
  });
});

describe("readSse", () => {
  const sse = (chunks: string[]) =>
    new Response(
      new ReadableStream({
        start(c) {
          for (const ch of chunks) c.enqueue(new TextEncoder().encode(ch));
          c.close();
        },
      })
    );
  it("joins text chunks split across reads", async () => {
    let out = "";
    await readSse(
      sse([
        'data: {"text":"Hel',
        'lo"}\n\ndata: {"text":" there"}\n\n',
        "event: done\ndata: {}\n\n",
      ]),
      (t) => (out += t)
    );
    expect(out).toBe("Hello there");
  });
  it("throws on an error event and on a cut-off stream", async () => {
    await expect(
      readSse(sse(['event: error\ndata: {"error":"busy"}\n\n']), () => {})
    ).rejects.toThrow("busy");
    await expect(readSse(sse(['data: {"text":"a"}\n\n']), () => {})).rejects.toThrow("closed");
  });
});
