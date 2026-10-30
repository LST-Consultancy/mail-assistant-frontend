import { Body1, Button, Caption1, Field, Radio, RadioGroup, Switch, Textarea } from "@fluentui/react-components";
import { ArrowReplyRegular, SparkleRegular } from "@fluentui/react-icons";
import { useCallback, useEffect, useRef, useState } from "react";
import { analyzeEmail, streamReply } from "../../services/api";
import { getThread, toEmail } from "../../services/graph";
import { onItemChanged, openReplyForm, readCurrentEmail, type CurrentEmail } from "../../services/office";
import type { Analysis, Email, Settings, Tone } from "../../types";
import {
  categoryLabel,
  CopyButton,
  DraftBox,
  Empty,
  ErrorBar,
  Loading,
  PriorityBadge,
  Section,
  ToneSelect,
  useStyles,
} from "../components/common";

const errMsg = (e: unknown) => (e instanceof Error ? e.message : String(e));
const isAbort = (e: unknown) => e instanceof DOMException && e.name === "AbortError";

export function ThisEmail({ settings }: { settings: Settings }) {
  const s = useStyles();
  const [current, setCurrent] = useState<CurrentEmail>();
  const [thread, setThread] = useState<Email[]>();
  const [threadNote, setThreadNote] = useState<string>();
  const [includeThread, setIncludeThread] = useState(true);
  const [analysis, setAnalysis] = useState<Analysis>();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>();

  const [tone, setTone] = useState<Tone>(settings.defaultTone);
  const [replyType, setReplyType] = useState<"reply" | "replyAll">("reply");
  const [instructions, setInstructions] = useState("");
  const [reply, setReply] = useState("");
  const [streaming, setStreaming] = useState(false);
  const [replyError, setReplyError] = useState<string>();

  const ctrl = useRef<AbortController>();
  const replyCtrl = useRef<AbortController>();

  const load = useCallback(async () => {
    ctrl.current?.abort();
    replyCtrl.current?.abort();
    const ac = (ctrl.current = new AbortController());
    setAnalysis(undefined);
    setReply("");
    setError(undefined);
    setReplyError(undefined);
    setThreadNote(undefined);
    setLoading(true);
    try {
      const cur = await readCurrentEmail();
      if (ac.signal.aborted) return;
      setCurrent(cur);
      let th: Email[] | undefined;
      if (includeThread && cur.conversationId) {
        try {
          const msgs = await getThread(cur.conversationId);
          th = msgs.length > 1 ? msgs.map((m) => toEmail(m)) : undefined;
        } catch (e) {
          setThreadNote(`Couldn't load the whole thread (${errMsg(e)}). Using this email only.`);
        }
      }
      if (ac.signal.aborted) return;
      setThread(th);
      setAnalysis(await analyzeEmail(cur.email, th, settings, ac.signal));
    } catch (e) {
      if (!isAbort(e) && !ac.signal.aborted) setError(errMsg(e));
    } finally {
      if (!ac.signal.aborted) setLoading(false);
    }
  }, [includeThread, settings]);

  useEffect(() => {
    void load();
  }, [load]);

  // Pinned pane: re-run for the newly selected email.
  const loadRef = useRef(load);
  loadRef.current = load;
  useEffect(() => onItemChanged(() => void loadRef.current()), []);

  useEffect(() => setTone(settings.defaultTone), [settings.defaultTone]);

  async function generate() {
    if (!current) return;
    replyCtrl.current?.abort();
    const ac = (replyCtrl.current = new AbortController());
    setReply("");
    setReplyError(undefined);
    setStreaming(true);
    try {
      await streamReply(
        { email: current.email, thread, tone, replyType, instructions: instructions || undefined },
        settings,
        (t) => setReply((r) => r + t),
        ac.signal,
      );
    } catch (e) {
      if (!isAbort(e)) setReplyError(errMsg(e));
    } finally {
      if (replyCtrl.current === ac) setStreaming(false);
    }
  }

  function openAsReply() {
    try {
      openReplyForm(reply, replyType === "replyAll");
    } catch (e) {
      setReplyError(errMsg(e));
    }
  }

  if (error && !current) return <ErrorBar error={error} onRetry={load} />;

  return (
    <div className={s.stack}>
      {current && (
        <div>
          <Body1 as="p" block style={{ margin: 0 }}>
            <strong>{current.email.subject || "(no subject)"}</strong>
          </Body1>
          <Caption1 className={s.muted}>
            {current.email.from.name || current.email.from.address}
            {thread ? ` · ${thread.length} messages in thread` : ""}
          </Caption1>
        </div>
      )}
      <Switch
        checked={includeThread}
        onChange={(_, d) => setIncludeThread(d.checked)}
        label="Include whole thread"
      />
      {threadNote && <Caption1 className={s.muted}>{threadNote}</Caption1>}

      {loading && <Loading label="Summarizing…" />}
      {error && <ErrorBar error={error} onRetry={load} />}

      {analysis && (
        <>
          <div className={s.row}>
            <PriorityBadge priority={analysis.priority} />
            <Caption1>{categoryLabel(analysis.category)}</Caption1>
            {analysis.needsReply && <Caption1>· needs reply</Caption1>}
          </div>
          <Section title="Summary">
            <Body1>{analysis.summary}</Body1>
          </Section>
          {analysis.keyPoints.length > 0 && (
            <Section title="Key points">
              <ul className={s.list}>
                {analysis.keyPoints.map((k, i) => (
                  <li key={i}>{k}</li>
                ))}
              </ul>
            </Section>
          )}
          <Section title="Action items for me">
            {analysis.actionItems.length ? (
              <ul className={s.list}>
                {analysis.actionItems.map((a, i) => (
                  <li key={i}>
                    {a.text}
                    {a.deadline && <strong> (due {a.deadline})</strong>}
                  </li>
                ))}
              </ul>
            ) : (
              <Caption1 className={s.muted}>None</Caption1>
            )}
          </Section>
        </>
      )}

      {current && (
        <Section title="Suggested reply">
          <ToneSelect value={tone} onChange={setTone} />
          <RadioGroup
            layout="horizontal"
            value={replyType}
            onChange={(_, d) => setReplyType(d.value as "reply" | "replyAll")}
            aria-label="Reply type"
          >
            <Radio value="reply" label="Reply" />
            <Radio value="replyAll" label="Reply all" />
          </RadioGroup>
          <Field label="Instructions (optional)">
            <Textarea
              value={instructions}
              onChange={(_, d) => setInstructions(d.value)}
              placeholder="e.g. decline politely, suggest next week"
            />
          </Field>
          <div className={s.row}>
            <Button appearance="primary" icon={<SparkleRegular />} onClick={generate} disabled={streaming}>
              {reply ? "Regenerate" : "Generate reply"}
            </Button>
            {streaming && <Button onClick={() => replyCtrl.current?.abort()}>Stop</Button>}
          </div>
          {replyError && <ErrorBar error={replyError} onRetry={generate} />}
          {(reply || streaming) && (
            <>
              <DraftBox label="Reply (editable)" value={reply} onChange={setReply} streaming={streaming} />
              <div className={s.row}>
                <Button icon={<ArrowReplyRegular />} onClick={openAsReply} disabled={!reply || streaming}>
                  Open as reply
                </Button>
                <CopyButton text={reply} />
              </div>
            </>
          )}
        </Section>
      )}
      {!current && !loading && !error && <Empty title="No email selected">Select an email to summarize it.</Empty>}
    </div>
  );
}
