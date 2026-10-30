import { Body1, Button, makeStyles, MessageBar, MessageBarActions, MessageBarBody, Tab, TabList, tokens } from "@fluentui/react-components";
import { DismissRegular, SettingsRegular } from "@fluentui/react-icons";
import { useState } from "react";
import { getMode, loadSettings, saveSettings } from "../services/office";
import type { Settings as SettingsT } from "../types";
import { Digest } from "./tabs/Digest";
import { Settings } from "./tabs/Settings";
import { ThisEmail } from "./tabs/ThisEmail";
import { Write } from "./tabs/Write";

const useStyles = makeStyles({
  root: { padding: tokens.spacingHorizontalM, display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM, minWidth: "300px" },
  header: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalXS },
  tabs: { flexGrow: 1, minWidth: 0 },
});

type TabId = "read" | "digest" | "write";

export const PRIVACY_NOTICE =
  "Email content you analyze is sent to OpenAI's API for processing. Nothing is stored by this add-in's server.";

export function App() {
  const s = useStyles();
  const mode = getMode();
  const available: TabId[] = mode === "read" ? ["read", "digest"] : mode === "compose" ? ["write", "digest"] : ["digest"];
  const requested = new URLSearchParams(location.search).get("tab") as TabId | null;
  const [tab, setTab] = useState<TabId>(requested && available.includes(requested) ? requested : available[0]);
  const [settings, setSettings] = useState<SettingsT>(loadSettings);
  const [showSettings, setShowSettings] = useState(false);

  const label: Record<TabId, string> = { read: "This email", digest: "Last 30 days", write: "Write" };

  return (
    <main className={s.root}>
      <div className={s.header}>
        <TabList className={s.tabs} selectedValue={tab} onTabSelect={(_, d) => (setTab(d.value as TabId), setShowSettings(false))}>
          {available.map((t) => (
            <Tab key={t} value={t}>
              {label[t]}
            </Tab>
          ))}
        </TabList>
        <Button
          appearance="subtle"
          icon={showSettings ? <DismissRegular /> : <SettingsRegular />}
          aria-label={showSettings ? "Close settings" : "Settings"}
          aria-pressed={showSettings}
          onClick={() => setShowSettings(!showSettings)}
        />
      </div>

      {!settings.privacyAcknowledged && (
        <MessageBar intent="warning" layout="multiline">
          <MessageBarBody>{PRIVACY_NOTICE}</MessageBarBody>
          <MessageBarActions>
            <Button
              size="small"
              onClick={() => {
                const next = { ...settings, privacyAcknowledged: true };
                setSettings(next);
                void saveSettings(next).catch(() => undefined);
              }}
            >
              I understand
            </Button>
          </MessageBarActions>
        </MessageBar>
      )}

      {showSettings && (
        <Settings
          settings={settings}
          onSaved={(n) => {
            setSettings(n);
            setShowSettings(false);
          }}
        />
      )}
      {/* Tabs stay mounted so results (e.g. the digest) survive tab switches for the session. */}
      <div hidden={showSettings}>
        {available.includes("read") && (
          <div hidden={tab !== "read"}>
            <ThisEmail settings={settings} />
          </div>
        )}
        {available.includes("write") && (
          <div hidden={tab !== "write"}>
            <Write settings={settings} />
          </div>
        )}
        <div hidden={tab !== "digest"}>
          <Digest settings={settings} />
        </div>
      </div>
      {mode === "none" && <Body1>Tip: select an email to summarize it or get a suggested reply.</Body1>}
    </main>
  );
}
