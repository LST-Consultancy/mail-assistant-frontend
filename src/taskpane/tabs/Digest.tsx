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
  Field,
  ProgressBar,
  Select,
  Switch,
  Textarea,
} from "@fluentui/react-components";
import { DocumentEditRegular, OpenRegular } from "@fluentui/react-icons";
import { useMemo, useRef, useState } from "react";
import { analyzeBatch, digestOverview, type BatchResult } from "../../services/api";
import { createReplyDraft, groupThreads, listInbox, listSent } from "../../services/graph";
import { myAddress, openMessage } from "../../services/office";
import { CATEGORIES, type DigestResult, type DigestThread, type Priority, type Settings } from "../../types";
import { categoryLabel, Empty, ErrorBar, Loading, Markdown, PriorityBadge, Section, useStyles } from "../components/common";

const BATCH = 10;
const RANK: Record<Priority, number> = { high: 0, medium: 1, low: 2 };
const errMsg = (e: unknown) => (e instanceof Error ? e.message : String(e));

type Row = { thread: DigestThread; result?: BatchResult; draft?: "saving" | "saved" | string };
const ok = (r?: BatchResult): r is DigestResult => !!r && !("error" in r);

export function Digest({ settings }: { settings: Settings }) {
  const s = useStyles();
  const [days, setDays] = useState(settings.digestDays);
  const [phase, setPhase] = useState<"idle" | "fetching" | "analyzing" | "done">("idle");
  const [error, setError] = useState<string>();
  const [rows, setRows] = useState<Row[]>([]);
  const [progress, setProgress] = useState(0);
  const [overview, setOverview] = useState<{ markdown?: string; error?: string; loading?: boolean }>({});
  const [edits, setEdits] = useState<Record<string, string>>({});
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [confirm, setConfirm] = useState(false);
  const [filters, setFilters] = useState({ needsReply: true, priority: "all", category: "all", unread: false });
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
        listSent(days).catch(() => []), // "already replied" detection is best-effort
      ]);
      threads = groupThreads(inbox, myAddress(), sent);
    } catch (e) {
      setError(errMsg(e));
      return setPhase("idle");
    }
    setRows(threads.map((thread) => ({ thread })));
    setPhase("analyzing");

    const results: DigestResult[] = [];
    for (let i = 0; i < threads.length && !ac.signal.aborted; i += BATCH) {
      const chunk = threads.slice(i, i + BATCH);
      let out: BatchResult[];
      try {
        out = await analyzeBatch(chunk.map((t) => t.email), settings, ac.signal);
      } catch (e) {
        if (ac.signal.aborted) return;
        out = chunk.map((t) => ({ id: t.email.id, error: errMsg(e) }));
      }
      const byId = new Map(out.map((r) => [r.id, r]));
      setRows((rs) =>
        rs.map((r) => (byId.has(r.thread.email.id) ? { ...r, result: byId.get(r.thread.email.id) } : r)),
      );
      results.push(...out.filter(ok));
      setProgress(Math.min(i + BATCH, threads.length));
    }
    if (ac.signal.aborted) return;
    setPhase("done");
    if (results.length) void loadOverview(threads, results, ac.signal);
  }

  async function loadOverview(threads: DigestThread[], results: DigestResult[], signal?: AbortSignal) {
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
          };
        }),
        signal,
      );
      setOverview({ markdown });
    } catch (e) {
      setOverview({ error: errMsg(e) });
    }
  }

  async function retryOne(row: Row) {
    patch(row.thread.conversationId, { result: undefined });
    try {
      const [r] = await analyzeBatch([row.thread.email], settings);
      patch(row.thread.conversationId, { result: r });
    } catch (e) {
      patch(row.thread.conversationId, { result: { id: row.thread.email.id, error: errMsg(e) } });
    }
  }

  const replyText = (r: Row) => edits[r.thread.conversationId] ?? (ok(r.result) ? r.result.replyDraft : "");

  const visible = useMemo(() => {
    return rows
      .filter((r) => {
        const res = ok(r.result) ? r.result : undefined;
        if (filters.unread && r.thread.latest.isRead) return false;
        if (!res) return !filters.needsReply && filters.priority === "all" && filters.category === "all";
        if (filters.needsReply && !res.needsReply) return false;
        if (filters.priority !== "all" && res.priority !== filters.priority) return false;
        if (filters.category !== "all" && res.category !== filters.category) return false;
        return true;
      })
      .sort((a, b) => {
        const pa = ok(a.result) ? RANK[a.result.priority] : 3;
        const pb = ok(b.result) ? RANK[b.result.priority] : 3;
        return pa - pb || b.thread.latest.receivedDateTime.localeCompare(a.thread.latest.receivedDateTime);
      });
  }, [rows, filters]);

  const selectedRows = rows.filter((r) => selected.has(r.thread.conversationId) && replyText(r).trim());

  async function createDrafts() {
    setConfirm(false);
    for (const r of selectedRows) {
      const id = r.thread.conversationId;
      patch(id, { draft: "saving" });
      try {
        await createReplyDraft(r.thread.latest.id, replyText(r));
        patch(id, { draft: "saved" });
      } catch (e) {
        patch(id, { draft: `Failed: ${errMsg(e)}` });
      }
    }
    setSelected(new Set());
  }

  const toggle = (id: string, on: boolean) =>
    setSelected((sel) => {
      const n = new Set(sel);
      if (on) n.add(id);
      else n.delete(id);
      return n;
    });

  const busy = phase === "fetching" || phase === "analyzing";
  const needsReplyCount = rows.filter((r) => ok(r.result) && r.result.needsReply).length;

  return (
    <div className={s.stack}>
      <div className={s.row}>
        <Field label="Range">
          <Select value={String(days)} onChange={(_, d) => setDays(Number(d.value))} disabled={busy}>
            <option value="7">Last 7 days</option>
            <option value="14">Last 14 days</option>
            <option value="30">Last 30 days</option>
          </Select>
        </Field>
        <Button appearance="primary" onClick={run} disabled={busy} style={{ alignSelf: "flex-end" }}>
          {rows.length ? "Regenerate digest" : "Generate digest"}
        </Button>
        {busy && (
          <Button onClick={() => (ctrl.current?.abort(), setPhase("done"))} style={{ alignSelf: "flex-end" }}>
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
        <Empty title="Nothing to review">No incoming threads in this range that still need you.</Empty>
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
              label={`Needs reply only (${needsReplyCount})`}
            />
            <Checkbox
              checked={filters.unread}
              onChange={(_, d) => setFilters((f) => ({ ...f, unread: !!d.checked }))}
              label="Unread only"
            />
          </div>
          <div className={s.row}>
            <Select
              aria-label="Priority filter"
              value={filters.priority}
              onChange={(_, d) => setFilters((f) => ({ ...f, priority: d.value }))}
            >
              <option value="all">All priorities</option>
              <option value="high">High</option>
              <option value="medium">Medium</option>
              <option value="low">Low</option>
            </Select>
            <Select
              aria-label="Category filter"
              value={filters.category}
              onChange={(_, d) => setFilters((f) => ({ ...f, category: d.value }))}
            >
              <option value="all">All categories</option>
              {CATEGORIES.map((c) => (
                <option key={c} value={c}>
                  {categoryLabel(c)}
                </option>
              ))}
            </Select>
          </div>
          <Button icon={<DocumentEditRegular />} disabled={!selectedRows.length} onClick={() => setConfirm(true)}>
            Create drafts for selected ({selectedRows.length})
          </Button>
          <Caption1 className={s.muted}>
            Showing {visible.length} of {rows.length} threads
          </Caption1>
        </>
      )}

      {visible.map((r) => (
        <DigestCard
          key={r.thread.conversationId}
          row={r}
          reply={replyText(r)}
          onReplyChange={(v) => setEdits((e) => ({ ...e, [r.thread.conversationId]: v }))}
          selected={selected.has(r.thread.conversationId)}
          onSelect={(on) => toggle(r.thread.conversationId, on)}
          onRetry={() => retryOne(r)}
          pending={busy && !r.result}
        />
      ))}

      <Dialog open={confirm} onOpenChange={(_, d) => setConfirm(d.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Create reply drafts</DialogTitle>
            <DialogContent>
              Create {selectedRows.length} reply draft{selectedRows.length === 1 ? "" : "s"} in your Drafts folder?
              Nothing will be sent.
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

function DigestCard(props: {
  row: Row;
  reply: string;
  onReplyChange: (v: string) => void;
  selected: boolean;
  onSelect: (on: boolean) => void;
  onRetry: () => void;
  pending: boolean;
}) {
  const s = useStyles();
  const [open, setOpen] = useState(false);
  const { thread, result, draft } = props.row;
  const m = thread.latest;
  const res = ok(result) ? result : undefined;
  const sender = thread.email.from.name || thread.email.from.address;
  const date = new Date(m.receivedDateTime).toLocaleDateString(undefined, { month: "short", day: "numeric" });

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
            {categoryLabel(res.category)}
            {res.needsReply ? " · needs reply" : ""}
          </Caption1>
        </>
      )}

      <div className={s.row}>
        <Button size="small" icon={<OpenRegular />} onClick={() => openMessage(m.id, m.webLink)}>
          Open email
        </Button>
        {res?.needsReply && (
          <Button size="small" appearance="subtle" onClick={() => setOpen(!open)} aria-expanded={open}>
            {open ? "Hide reply" : "Suggested reply"}
          </Button>
        )}
        {draft && <Caption1 role="status">{draft === "saving" ? "Saving draft…" : draft === "saved" ? "Draft saved ✓" : draft}</Caption1>}
      </div>
      {open && (
        <Textarea
          aria-label="Suggested reply"
          value={props.reply}
          onChange={(_, d) => props.onReplyChange(d.value)}
          rows={6}
          resize="vertical"
        />
      )}
    </Card>
  );
}
