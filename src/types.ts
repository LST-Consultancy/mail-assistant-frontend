export interface Address {
  name: string;
  address: string;
}

/** Shape the backend accepts (see backend/src/llm/schemas.ts). */
export interface Email {
  id: string;
  subject: string;
  from: Address;
  to: Address[];
  cc: Address[];
  receivedDateTime: string;
  body: string;
  importance?: string;
}

export type Tone = "professional" | "friendly" | "short" | "formal";
export type Priority = "high" | "medium" | "low";
export const CATEGORIES = [
  "action_required",
  "question",
  "fyi",
  "meeting",
  "newsletter_or_promo",
  "notification",
  "personal",
  "other",
] as const;
export type Category = (typeof CATEGORIES)[number];

export interface Analysis {
  summary: string;
  keyPoints: string[];
  actionItems: { text: string; deadline: string | null }[];
  category: Category;
  priority: Priority;
  needsReply: boolean;
}

export interface DigestResult {
  id: string;
  summary: string;
  category: Category;
  priority: Priority;
  needsReply: boolean;
  replyDraft: string;
}

export interface Settings {
  signatureName: string;
  defaultTone: Tone;
  /** "match" = sender's language, otherwise a language name like "German". */
  language: string;
  digestDays: number;
  maxEmails: number;
  privacyAcknowledged: boolean;
}

/** Subset of a Graph message resource we $select. */
export interface GraphMessage {
  id: string;
  subject: string | null;
  from?: { emailAddress: Address };
  toRecipients?: { emailAddress: Address }[];
  ccRecipients?: { emailAddress: Address }[];
  receivedDateTime: string;
  sentDateTime?: string;
  conversationId: string;
  uniqueBody?: { content: string };
  bodyPreview?: string;
  isRead?: boolean;
  importance?: string;
  webLink?: string;
}

/** One digest row: the latest incoming message of a conversation plus thread context. */
export interface DigestThread {
  conversationId: string;
  latest: GraphMessage;
  email: Email;
  messageCount: number;
}
