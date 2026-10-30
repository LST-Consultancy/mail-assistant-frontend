import {
  Badge,
  Body1,
  Button,
  Field,
  makeStyles,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  MessageBarTitle,
  Select,
  Spinner,
  Subtitle2,
  Textarea,
  tokens,
} from "@fluentui/react-components";
import { CopyRegular } from "@fluentui/react-icons";
import { Fragment, useState, type ReactNode } from "react";
import type { Priority, Tone } from "../../types";

export const useStyles = makeStyles({
  stack: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  row: { display: "flex", gap: tokens.spacingHorizontalS, flexWrap: "wrap", alignItems: "center" },
  section: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS },
  muted: { color: tokens.colorNeutralForeground3 },
  list: { margin: 0, paddingLeft: "20px" },
  grow: { flexGrow: 1, minWidth: 0 },
});

export function Loading({ label }: { label: string }) {
  return <Spinner size="small" label={label} labelPosition="after" />;
}

export function ErrorBar({ error, onRetry }: { error: string; onRetry?: () => void }) {
  return (
    <MessageBar intent="error" role="alert">
      <MessageBarBody>{error}</MessageBarBody>
      {onRetry && (
        <MessageBarActions>
          <Button size="small" onClick={onRetry}>
            Retry
          </Button>
        </MessageBarActions>
      )}
    </MessageBar>
  );
}

export function Empty({ title, children }: { title: string; children?: ReactNode }) {
  return (
    <MessageBar intent="info">
      <MessageBarBody>
        <MessageBarTitle>{title}</MessageBarTitle>
        {children}
      </MessageBarBody>
    </MessageBar>
  );
}

export function Section({ title, children }: { title: string; children: ReactNode }) {
  const s = useStyles();
  return (
    <section className={s.section} aria-label={title}>
      <Subtitle2 as="h3">{title}</Subtitle2>
      {children}
    </section>
  );
}

const PRIORITY_COLOR = { high: "danger", medium: "warning", low: "informative" } as const;
export function PriorityBadge({ priority }: { priority: Priority }) {
  return (
    <Badge appearance="tint" color={PRIORITY_COLOR[priority]} aria-label={`Priority ${priority}`}>
      {priority}
    </Badge>
  );
}

export const categoryLabel = (c: string) => c.replace(/_/g, " ");

export const TONES: { value: Tone; label: string }[] = [
  { value: "professional", label: "Professional" },
  { value: "friendly", label: "Friendly" },
  { value: "short", label: "Short" },
  { value: "formal", label: "Formal" },
];

export function ToneSelect({ value, onChange }: { value: Tone; onChange: (t: Tone) => void }) {
  return (
    <Field label="Tone">
      <Select value={value} onChange={(_, d) => onChange(d.value as Tone)}>
        {TONES.map((t) => (
          <option key={t.value} value={t.value}>
            {t.label}
          </option>
        ))}
      </Select>
    </Field>
  );
}

export async function copyText(text: string): Promise<void> {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    // Clipboard API is blocked in some Office webviews.
    const ta = document.createElement("textarea");
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    ta.remove();
  }
}

export function CopyButton({ text }: { text: string }) {
  const [done, setDone] = useState(false);
  return (
    <Button
      icon={<CopyRegular />}
      disabled={!text}
      onClick={async () => {
        await copyText(text);
        setDone(true);
        setTimeout(() => setDone(false), 1500);
      }}
    >
      {done ? "Copied" : "Copy"}
    </Button>
  );
}

export function DraftBox({
  label,
  value,
  onChange,
  streaming,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  streaming?: boolean;
}) {
  return (
    <Field label={label} hint={streaming ? "Writing…" : undefined}>
      <Textarea value={value} onChange={(_, d) => onChange(d.value)} resize="vertical" rows={10} aria-busy={streaming} />
    </Field>
  );
}

/** Minimal Markdown (## headings, - bullets, **bold**) rendered as React elements, no HTML injection. */
export function Markdown({ text }: { text: string }) {
  const s = useStyles();
  const inline = (line: string) =>
    line.split(/(\*\*[^*]+\*\*)/g).map((part, i) =>
      part.startsWith("**") && part.endsWith("**") ? <strong key={i}>{part.slice(2, -2)}</strong> : <Fragment key={i}>{part}</Fragment>,
    );
  const blocks: ReactNode[] = [];
  let bullets: string[] = [];
  const flush = () => {
    if (bullets.length)
      blocks.push(
        <ul className={s.list} key={blocks.length}>
          {bullets.map((b, i) => (
            <li key={i}>{inline(b)}</li>
          ))}
        </ul>,
      );
    bullets = [];
  };
  for (const raw of text.split("\n")) {
    const line = raw.trim();
    if (/^[-*] /.test(line)) {
      bullets.push(line.slice(2));
      continue;
    }
    flush();
    if (!line) continue;
    if (line.startsWith("#")) blocks.push(<Subtitle2 as="h4" key={blocks.length}>{inline(line.replace(/^#+\s*/, ""))}</Subtitle2>);
    else blocks.push(<Body1 as="p" key={blocks.length}>{inline(line)}</Body1>);
  }
  flush();
  return <div className={s.section}>{blocks}</div>;
}
