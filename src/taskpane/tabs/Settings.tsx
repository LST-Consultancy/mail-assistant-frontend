import { Button, Field, Input, Select, SpinButton } from "@fluentui/react-components";
import { useState } from "react";
import { saveSettings } from "../../services/office";
import type { Settings as S, Tone } from "../../types";
import { ErrorBar, TONES, useStyles } from "../components/common";

const LANGUAGES = ["English", "German", "French", "Spanish", "Italian", "Dutch", "Portuguese", "Hindi", "Japanese", "Chinese"];

export function Settings({ settings, onSaved }: { settings: S; onSaved: (s: S) => void }) {
  const st = useStyles();
  const [draft, setDraft] = useState(settings);
  const [error, setError] = useState<string>();
  const [saving, setSaving] = useState(false);
  const set = <K extends keyof S>(k: K, v: S[K]) => setDraft((d) => ({ ...d, [k]: v }));

  async function save() {
    setSaving(true);
    setError(undefined);
    try {
      await saveSettings(draft);
      onSaved(draft);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Couldn't save settings.");
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className={st.stack}>
      <Field label="Signature name used in drafts">
        <Input value={draft.signatureName} onChange={(_, d) => set("signatureName", d.value)} placeholder="e.g. Sam" />
      </Field>
      <Field label="Default tone">
        <Select value={draft.defaultTone} onChange={(_, d) => set("defaultTone", d.value as Tone)}>
          {TONES.map((t) => (
            <option key={t.value} value={t.value}>
              {t.label}
            </option>
          ))}
        </Select>
      </Field>
      <Field label="Reply language">
        <Select value={draft.language} onChange={(_, d) => set("language", d.value)}>
          <option value="match">Match sender</option>
          {LANGUAGES.map((l) => (
            <option key={l} value={l}>
              {l}
            </option>
          ))}
        </Select>
      </Field>
      <Field label="Digest default range">
        <Select value={String(draft.digestDays)} onChange={(_, d) => set("digestDays", Number(d.value))}>
          <option value="7">7 days</option>
          <option value="14">14 days</option>
          <option value="30">30 days</option>
        </Select>
      </Field>
      <Field label="Max emails per digest" hint="Between 20 and 1000">
        <SpinButton
          value={draft.maxEmails}
          min={20}
          max={1000}
          step={50}
          onChange={(_, d) => {
            const v = d.value ?? Number(d.displayValue);
            if (Number.isFinite(v)) set("maxEmails", Math.min(1000, Math.max(20, Math.round(v))));
          }}
        />
      </Field>
      {error && <ErrorBar error={error} />}
      <Button appearance="primary" onClick={save} disabled={saving}>
        {saving ? "Saving…" : "Save"}
      </Button>
    </div>
  );
}
