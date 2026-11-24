import {
  Body1,
  Button,
  Caption1,
  Card,
  Checkbox,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Dropdown,
  Field,
  makeStyles,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  Option,
  ProgressBar,
  Select,
  Switch,
  Textarea,
  tokens,
} from "@fluentui/react-components";
import { DocumentEditRegular, EyeOffRegular, OpenRegular } from "@fluentui/react-icons";
import { useMemo, useRef, useState } from "react";
import { analyzeBatch, digestOverview, type BatchResult } from "../../services/api";
import {
  createReplyDraft,
  followUpThreads,
  groupThreads,
  listInbox,
  listSent,
  setIgnored,
} from "../../services/graph";
import { myAddress, openMessage } from "../../services/office";
import {
  CATEGORIES,
  type Category,
  type DigestResult,
  type DigestThread,
  type Priority,
  type Settings,
} from "../../types";
import {
  categoryLabel,
  Empty,
  ErrorBar,
  Loading,
  Markdown,
  PriorityBadge,
  Section,
  useStyles,
} from "../components/common";

const BATCH = 10;
const RANK: Record<Priority, number> = { high: 0, medium: 1, low: 2 };
const PRIORITIES: Priority[] = ["high", "medium", "low"];
/** Hidden by default and never fed into the overview. */
const NOISE: Category[] = ["newsletter_or_promo", "notification", "spam"];
const KINDS: [DigestThread["kind"], string][] = [
  ["inbox", "Received"],
  ["sent", "Sent, awaiting reply"],
];
const errMsg = (e: unknown) => (e instanceof Error ? e.message : String(e));
const modeOf = (t: DigestThread) => (t.kind === "sent" ? "followup" : "inbox");

// Cards flow into columns when the user widens the task pane.
const useGrid = makeStyles({
  cards: {
    display: "grid",
    gap: tokens.spacingVerticalM,
    gridTemplateColumns: "repeat(auto-fill, minmax(min(100%, 360px), 1fr))",
    alignItems: "start",
  },
});

type Row = {
  thread: DigestThread;
  result?: BatchResult;
  draft?: "saving" | "saved" | string;
  ignored?: boolean;
  /** "saving" or an error message. */
  ignoreStatus?: string;
};
const ok = (r?: BatchResult): r is DigestResult => !!r && !("error" in r);

export function Digest({ settings }: { settings: Settings }) {
  const s = useStyles();
  const grid = useGrid();
  const [days, setDays] = useState(settings.digestDays);
  const [phase, setPhase] = useState<"idle" | "fetching" | "analyzing" | "done">("idle");
  const [error, setError] = useState<string>();
  const [rows, setRows] = useState<Row[]>([]);
  const [progress, setProgress] = useState(0);
  const [overview, setOverview] = useState<{
    markdown?: string;
    error?: string;
    loading?: boolean;
  }>({});
  const [edits, setEdits] = useState<Record<string, string>>({});
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [confirm, setConfirm] = useState(false);
  const [filters, setFilters] = useState({
    needsReply: true,
    unread: false,
    priorities: PRIORITIES as string[],
    categories: CATEGORIES.filter((c) => !NOISE.includes(c)) as string[],
    kinds: KINDS.map(([k]) => k) as string[],
  });
  const ctrl = useRef<AbortController>();

  const patch = (id: string, p: Partial<Row>) =>
    setRows((rs) => rs.map((r) => (r.thread.conversationId === id ? { ...r, ...p } : r)));

  async function run() {
    ctrl.current?.abort();
    const ac = (ctrl.current = new AbortController());
    setError(undefined);
    setRows([]);
    setOverview({});
    setEdits({});
    setSelected(new Set());
    setProgress(0);
    setPhase("fetching");
    let threads: DigestThread[];
    try {
      const [inbox, sent] = await Promise.all([
        listInbox(days, settings.maxEmails),
        listSent(days, settings.maxEmails).catch(() => []), // reply detection and follow-ups are best-effort
      ]);
      const me = myAddress();
      threads = [...groupThreads(inbox, me, sent), ...followUpThreads(sent, inbox, me)];
    } catch (e) {
      setError(errMsg(e));
      return setPhase("idle");
    }
    setRows(threads.map((thread) => ({ thread })));
    setPhase("analyzing");

    // Each batch holds one kind only: the backend uses a different prompt for sent follow-ups.
    const chunks: DigestThread[][] = [];
    for (const [kind] of KINDS) {
      const list = threads.filter((t) => t.kind === kind);
      for (let i = 0; i < list.length; i += BATCH) chunks.push(list.slice(i, i + BATCH));
    }
    const results: DigestResult[] = [];
    let done = 0;
    for (const chunk of chunks) {
      if (ac.signal.aborted) break;
      let out: BatchResult[];
      try {
        out = await analyzeBatch(
          chunk.map((t) => t.email),
          settings,
          ac.signal,
          modeOf(chunk[0])
        );
      } catch (e) {
        if (ac.signal.aborted) return;
        out = chunk.map((t) => ({ id: t.email.id, error: errMsg(e) }));
      }
      const byId = new Map(out.map((r) => [r.id, r]));
      setRows((rs) =>
        rs.map((r) =>
          byId.has(r.thread.email.id) ? { ...r, result: byId.get(r.thread.email.id) } : r
        )
      );
      results.push(...out.filter(ok));
      setProgress((done += chunk.length));
    }
    if (ac.signal.aborted) return;
    setPhase("done");
    const relevant = results.filter((r) => !NOISE.includes(r.category));
    if (relevant.length) void loadOverview(threads, relevant, ac.signal);
  }

  async function loadOverview(
    threads: DigestThread[],
    results: DigestResult[],
    signal?: AbortSignal
  ) {
    setOverview({ loading: true });
    const byId = new Map(threads.map((t) => [t.email.id, t]));
    try {
      const { markdown } = await digestOverview(
        results.map((r) => {
          const t = byId.get(r.id)!;
          return {
            subject: t.email.subject,
            from: t.email.from.name || t.email.from.address,
            summary: r.summary,
            priority: r.priority,
            category: r.category,
            needsReply: r.needsReply,
            followUp: t.kind === "sent",
          };
        }),
        signal
      );
      setOverview({ markdown });
    } catch (e) {
      setOverview({ error: errMsg(e) });
    }
  }

  async function retryOne(row: Row) {
    patch(row.thread.conversationId, { result: undefined });
    try {
      const [r] = await analyzeBatch([row.thread.email], settings, undefined, modeOf(row.thread));
      patch(row.thread.conversationId, { result: r });
    } catch (e) {
      patch(row.thread.conversationId, { result: { id: row.thread.email.id, error: errMsg(e) } });
    }
  }

  const replyText = (r: Row) =>
    edits[r.thread.conversationId] ?? (ok(r.result) ? r.result.replyDraft : "");

  const visible = useMemo(() => {
    return rows
      .filter((r) => {
        if (r.ignored) return true; // stays in place so Undo is reachable
        const res = ok(r.result) ? r.result : undefined;
        if (!filters.kinds.includes(r.thread.kind)) return false;
        if (filters.unread && r.thread.latest.isRead) return false;
        const all =
          filters.priorities.length === PRIORITIES.length &&
          filters.categories.length === CATEGORIES.length;
        if (!res) return !filters.needsReply && all;
        if (filters.needsReply && !res.needsReply) return false;
        if (!filters.priorities.includes(res.priority)) return false;
        if (!filters.categories.includes(res.category)) return false;
        return true;
      })
      .sort((a, b) => {
        const pa = ok(a.result) ? RANK[a.result.priority] : 3;
        const pb = ok(b.result) ? RANK[b.result.priority] : 3;
        return (
          pa - pb ||
          b.thread.latest.receivedDateTime.localeCompare(a.thread.latest.receivedDateTime)
        );
      });
  }, [rows, filters]);

  const selectedRows = rows.filter(
    (r) => selected.has(r.thread.conversationId) && !r.ignored && replyText(r).trim()
  );

  async function createDrafts() {
    setConfirm(false);
    for (const r of selectedRows) {
      const id = r.thread.conversationId;
      patch(id, { draft: "saving" });
      try {
        await createReplyDraft(
          r.thread.latest.id,
          replyText(r),
          r.thread.kind === "sent" ? r.thread.latest : undefined
        );
        patch(id, { draft: "saved" });
      } catch (e) {
        patch(id, { draft: `Failed: ${errMsg(e)}` });
      }
    }
    setSelected(new Set());
  }

  async function ignore(row: Row, on: boolean) {
    const id = row.thread.conversationId;
    patch(id, { ignoreStatus: "saving" });
    try {
      await setIgnored(row.thread.latest, on);
      patch(id, { ignored: on, ignoreStatus: undefined });
      if (on) toggle(id, false);
    } catch (e) {
      patch(id, { ignoreStatus: `Failed: ${errMsg(e)}` });
    }
  }

  const toggle = (id: string, on: boolean) =>
    setSelected((sel) => {
      const n = new Set(sel);
      if (on) n.add(id);
      else n.delete(id);
      return n;
    });

  const busy = phase === "fetching" || phase === "analyzing";
  const needsReplyCount = rows.filter(
    (r) => !r.ignored && ok(r.result) && r.result.needsReply
  ).length;

  return (
    <div className={s.stack}>
      <div className={s.row}>
        <Field label="Range">
          <Select
            value={String(days)}
            onChange={(_, d) => setDays(Number(d.value))}
            disabled={busy}
          >
            <option value="7">Last 7 days</option>
            <option value="14">Last 14 days</option>
            <option value="30">Last 30 days</option>
          </Select>
        </Field>
        <Button
          appearance="primary"
          onClick={run}
          disabled={busy}
          style={{ alignSelf: "flex-end" }}
        >
          {rows.length ? "Regenerate digest" : "Generate digest"}
        </Button>
        {busy && (
          <Button
            onClick={() => (ctrl.current?.abort(), setPhase("done"))}
            style={{ alignSelf: "flex-end" }}
          >
            Stop
          </Button>
        )}
      </div>

      {error && <ErrorBar error={error} onRetry={run} />}
      {phase === "fetching" && <Loading label="Reading your inbox…" />}
      {phase === "analyzing" && (
        <Field label={`Analyzing ${progress} / ${rows.length}`} validationState="none">
          <ProgressBar value={rows.length ? progress / rows.length : 0} />
        </Field>
      )}
      {phase === "done" && rows.length === 0 && (
        <Empty title="Nothing to review">No threads in this range that still need you.</Empty>
      )}

      {(overview.loading || overview.markdown || overview.error) && (
        <Card>
          <Section title="Overview">
            {overview.loading && <Loading label="Writing overview…" />}
            {overview.error && <ErrorBar error={overview.error} />}
            {overview.markdown && <Markdown text={overview.markdown} />}
          </Section>
        </Card>
      )}

      {rows.length > 0 && (
        <>
          <div className={s.row}>
            <Switch
              checked={filters.needsReply}
              onChange={(_, d) => setFilters((f) => ({ ...f, needsReply: d.checked }))}
              label={`Needs reply or follow-up only (${needsReplyCount})`}
            />
            <Checkbox
              checked={filters.unread}
              onChange={(_, d) => setFilters((f) => ({ ...f, unread: !!d.checked }))}
              label="Unread only"
            />
          </div>
          <div className={s.row}>
            <MultiSelect
              label="Emails"
              options={KINDS}
              value={filters.kinds}
              onChange={(kinds) => setFilters((f) => ({ ...f, kinds }))}
            />
            <MultiSelect
              label="Priorities"
              options={PRIORITIES.map((p) => [p, p[0].toUpperCase() + p.slice(1)])}
              value={filters.priorities}
              onChange={(priorities) => setFilters((f) => ({ ...f, priorities }))}
            />
            <MultiSelect
              label="Categories"
              options={CATEGORIES.map((c) => [c, categoryLabel(c)])}
              value={filters.categories}
              onChange={(categories) => setFilters((f) => ({ ...f, categories }))}
            />
          </div>
          <Button
            icon={<DocumentEditRegular />}
            disabled={!selectedRows.length}
            onClick={() => setConfirm(true)}
          >
            Create drafts for selected ({selectedRows.length})
          </Button>
          <Caption1 className={s.muted}>
            Showing {visible.length} of {rows.length} threads
          </Caption1>
        </>
      )}

      <div className={grid.cards}>
        {visible.map((r) => (
          <DigestCard
            key={r.thread.conversationId}
            row={r}
            reply={replyText(r)}
            onReplyChange={(v) => setEdits((e) => ({ ...e, [r.thread.conversationId]: v }))}
            selected={selected.has(r.thread.conversationId)}
            onSelect={(on) => toggle(r.thread.conversationId, on)}
            onRetry={() => retryOne(r)}
            onIgnore={(on) => ignore(r, on)}
            pending={busy && !r.result}
          />
        ))}
      </div>

      <Dialog open={confirm} onOpenChange={(_, d) => setConfirm(d.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Create reply drafts</DialogTitle>
            <DialogContent>
              Create {selectedRows.length} reply draft{selectedRows.length === 1 ? "" : "s"} in your
              Drafts folder? Nothing will be sent.
            </DialogContent>
            <DialogActions>
              <Button onClick={() => setConfirm(false)}>Cancel</Button>
              <Button appearance="primary" onClick={createDrafts}>
                Create drafts
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
}

function MultiSelect(props: {
  label: string;
  options: [string, string][];
  value: string[];
  onChange: (v: string[]) => void;
}) {
  const { label, options, value } = props;
  const text =
    value.length === options.length
      ? `All ${label.toLowerCase()}`
      : value.length === 0
        ? `No ${label.toLowerCase()}`
        : options
            .filter(([v]) => value.includes(v))
            .map(([, l]) => l)
            .join(", ");
  return (
    <Dropdown
      aria-label={`${label} filter`}
      multiselect
      selectedOptions={value}
      value={text}
      onOptionSelect={(_, d) => props.onChange(d.selectedOptions)}
    >
      {options.map(([v, l]) => (
        <Option key={v} value={v}>
          {l}
        </Option>
      ))}
    </Dropdown>
  );
}

function DigestCard(props: {
  row: Row;
  reply: string;
  onReplyChange: (v: string) => void;
  selected: boolean;
  onSelect: (on: boolean) => void;
  onRetry: () => void;
  onIgnore: (on: boolean) => void;
  pending: boolean;
}) {
  const s = useStyles();
  const [open, setOpen] = useState(false);
  const { thread, result, draft, ignored, ignoreStatus } = props.row;
  const m = thread.latest;
  const res = ok(result) ? result : undefined;
  const sent = thread.kind === "sent";
  const sender = sent
    ? "To: " + (thread.email.to.map((a) => a.name || a.address).join(", ") || "(no recipients)")
    : thread.email.from.name || thread.email.from.address;
  const date = new Date(m.sentDateTime ?? m.receivedDateTime).toLocaleDateString(undefined, {
    month: "short",
    day: "numeric",
  });
  const busyIgnoring = ignoreStatus === "saving";
  const ignoreError =
    ignoreStatus && !busyIgnoring ? <Caption1 role="alert">{ignoreStatus}</Caption1> : null;

  if (ignored)
    return (
      <MessageBar intent="info" aria-label={`Ignored: ${thread.email.subject}`}>
        <MessageBarBody>
          Ignored “{thread.email.subject || "(no subject)"}”. It won’t appear in the assistant
          again. {ignoreError}
        </MessageBarBody>
        <MessageBarActions>
          <Button size="small" disabled={busyIgnoring} onClick={() => props.onIgnore(false)}>
            Undo
          </Button>
        </MessageBarActions>
      </MessageBar>
    );

  return (
    <Card aria-label={`${sender}: ${thread.email.subject}`}>
      <div className={s.row}>
        <Checkbox
          aria-label="Select for bulk drafts"
          checked={props.selected}
          disabled={!props.reply || draft === "saved"}
          onChange={(_, d) => props.onSelect(!!d.checked)}
        />
        <div className={s.grow}>
          <Body1 as="p" block style={{ margin: 0 }}>
            <strong>{sender}</strong>
            {!m.isRead && " ●"}
          </Body1>
          <Caption1>
            {thread.email.subject || "(no subject)"} · {date}
            {thread.messageCount > 1 ? ` · ${thread.messageCount} msgs` : ""}
          </Caption1>
        </div>
        {res && <PriorityBadge priority={res.priority} />}
      </div>

      {props.pending && <Loading label="Analyzing…" />}
      {result && "error" in result && (
        <ErrorBar error={`Couldn't analyze: ${result.error}`} onRetry={props.onRetry} />
      )}
      {res && (
        <>
          <Body1>{res.summary}</Body1>
          <Caption1 className={s.muted}>
            {sent ? "Sent · awaiting reply · " : ""}
            {categoryLabel(res.category)}
            {res.needsReply ? (sent ? " · needs follow-up" : " · needs reply") : ""}
          </Caption1>
        </>
      )}

      <div className={s.row}>
        <Button size="small" icon={<OpenRegular />} onClick={() => openMessage(m.id, m.webLink)}>
          Open email
        </Button>
        {res?.needsReply && (
          <Button
            size="small"
            appearance="subtle"
            onClick={() => setOpen(!open)}
            aria-expanded={open}
          >
            {open ? "Hide" : sent ? "Suggested follow-up" : "Suggested reply"}
          </Button>
        )}
        <Button
          size="small"
          appearance="subtle"
          icon={<EyeOffRegular />}
          disabled={busyIgnoring}
          onClick={() => props.onIgnore(true)}
          title="Hide this conversation from the assistant for good"
        >
          {busyIgnoring ? "Ignoring…" : "Ignore"}
        </Button>
        {ignoreError}
        {draft && (
          <Caption1 role="status">
            {draft === "saving" ? "Saving draft…" : draft === "saved" ? "Draft saved ✓" : draft}
          </Caption1>
        )}
      </div>
      {open && (
        <Textarea
          aria-label={sent ? "Suggested follow-up" : "Suggested reply"}
          value={props.reply}
          onChange={(_, d) => props.onReplyChange(d.value)}
          rows={6}
          resize="vertical"
        />
      )}
    </Card>
  );
}
