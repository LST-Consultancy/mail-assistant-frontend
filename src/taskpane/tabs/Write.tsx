import { Button, Caption1, Field, Tab, TabList, Textarea } from "@fluentui/react-components";
import { SparkleRegular } from "@fluentui/react-icons";
import { useEffect, useRef, useState } from "react";
import { streamCompose } from "../../services/api";
import { getThread, toEmail } from "../../services/graph";
import { composeConversationId, getComposeBody, insertIntoBody } from "../../services/office";
import type { Email, Settings, Tone } from "../../types";
import { CopyButton, DraftBox, ErrorBar, ToneSelect, useStyles } from "../components/common";

const errMsg = (e: unknown) => (e instanceof Error ? e.message : String(e));

const IMPROVE = [
  { label: "Clearer", instructions: "Make it clearer and better structured." },
  { label: "Shorter", instructions: "Make it shorter and more concise." },
  { label: "More polite", instructions: "Make it more polite and courteous." },
  { label: "Fix grammar", instructions: "Only fix grammar, spelling and punctuation. Change nothing else." },
];

export function Write({ settings }: { settings: Settings }) {
  const s = useStyles();
  const [mode, setMode] = useState<"draft" | "improve">("draft");
  const [tone, setTone] = useState<Tone>(settings.defaultTone);
  const [instructions, setInstructions] = useState("");
  const [output, setOutput] = useState("");
  const [streaming, setStreaming] = useState(false);
  const [error, setError] = useState<string>();
  const [status, setStatus] = useState<string>();
  const [context, setContext] = useState<Email>();
  const ctrl = useRef<AbortController>();

  // When composing a reply, use the latest message of the conversation as context (best-effort).
  useEffect(() => {
    const cid = composeConversationId();
    if (!cid) return;
    getThread(cid)
      .then((msgs) => msgs.length && setContext(toEmail(msgs[msgs.length - 1])))
      .catch(() => undefined);
  }, []);

  useEffect(() => setTone(settings.defaultTone), [settings.defaultTone]);

  async function generate(extra?: string) {
    ctrl.current?.abort();
    const ac = (ctrl.current = new AbortController());
    setOutput("");
    setError(undefined);
    setStatus(undefined);
    setStreaming(true);
    try {
      const text = mode === "improve" ? await getComposeBody() : undefined;
      if (mode === "improve" && !text?.trim()) throw new Error("Your draft is empty. Write something first.");
      const instr = [extra, instructions].filter(Boolean).join(" ") || undefined;
      if (mode === "draft" && !instr) throw new Error("Describe what the email should say.");
      await streamCompose({ mode, text, instructions: instr, context, tone }, settings, (t) => setOutput((o) => o + t), ac.signal);
    } catch (e) {
      if (!(e instanceof DOMException && e.name === "AbortError")) setError(errMsg(e));
    } finally {
      if (ctrl.current === ac) setStreaming(false);
    }
  }

  async function insert(replace: boolean) {
    setError(undefined);
    try {
      await insertIntoBody(output, replace);
      setStatus(replace ? "Body replaced." : "Inserted at cursor.");
    } catch (e) {
      setError(errMsg(e));
    }
  }

  return (
    <div className={s.stack}>
      <TabList size="small" selectedValue={mode} onTabSelect={(_, d) => setMode(d.value as "draft" | "improve")}>
        <Tab value="draft">Draft from instructions</Tab>
        <Tab value="improve">Improve my draft</Tab>
      </TabList>
      {context && <Caption1 className={s.muted}>Replying to: {context.subject}</Caption1>}
      <ToneSelect value={tone} onChange={setTone} />
      <Field label={mode === "draft" ? "What should the email say?" : "Extra instructions (optional)"}>
        <Textarea
          value={instructions}
          onChange={(_, d) => setInstructions(d.value)}
          placeholder={mode === "draft" ? "e.g. thank Ann, confirm Nov 3 pilot, ask for the DPA" : "e.g. keep it under 100 words"}
        />
      </Field>
      {mode === "draft" ? (
        <Button appearance="primary" icon={<SparkleRegular />} onClick={() => generate()} disabled={streaming}>
          Write draft
        </Button>
      ) : (
        <div className={s.row}>
          {IMPROVE.map((o) => (
            <Button key={o.label} onClick={() => generate(o.instructions)} disabled={streaming}>
              {o.label}
            </Button>
          ))}
        </div>
      )}
      {streaming && <Button onClick={() => ctrl.current?.abort()}>Stop</Button>}
      {error && <ErrorBar error={error} />}
      {(output || streaming) && (
        <>
          <DraftBox label="Result (editable)" value={output} onChange={setOutput} streaming={streaming} />
          <div className={s.row}>
            <Button appearance="primary" onClick={() => insert(false)} disabled={!output || streaming}>
              Insert
            </Button>
            <Button onClick={() => insert(true)} disabled={!output || streaming}>
              Replace body
            </Button>
            <CopyButton text={output} />
          </div>
          {status && <Caption1 role="status">{status}</Caption1>}
        </>
      )}
    </div>
  );
}
