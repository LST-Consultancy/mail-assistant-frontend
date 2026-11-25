import { Body1, Button, Caption1, Card } from "@fluentui/react-components";
import { ArrowUndoRegular, OpenRegular } from "@fluentui/react-icons";
import { useCallback, useEffect, useState } from "react";
import { listIgnored, setIgnored } from "../../services/graph";
import { myAddress, openMessage } from "../../services/office";
import type { GraphMessage } from "../../types";
import { Empty, ErrorBar, Loading, useStyles } from "../components/common";

const errMsg = (e: unknown) => (e instanceof Error ? e.message : String(e));

/** Conversations the user ignored. Reloads each time the tab is shown so it reflects ignores made in the digest. */
export function Ignored({ active }: { active: boolean }) {
  const s = useStyles();
  const [groups, setGroups] = useState<GraphMessage[][]>();
  const [error, setError] = useState<string>();
  const [status, setStatus] = useState<Record<string, string>>({}); // conversationId -> "saving" | error

  const load = useCallback(async () => {
    setError(undefined);
    try {
      setGroups(await listIgnored());
    } catch (e) {
      setError(errMsg(e));
    }
  }, []);

  useEffect(() => {
    if (active) void load();
  }, [active, load]);

  async function unignore(group: GraphMessage[]) {
    const id = group[0].conversationId;
    setStatus((st) => ({ ...st, [id]: "saving" }));
    try {
      // Every tagged message in the conversation, or it would stay hidden from the digest.
      await Promise.all(group.map((m) => setIgnored(m, false)));
      setGroups((gs) => gs?.filter((g) => g !== group));
    } catch (e) {
      setStatus((st) => ({ ...st, [id]: `Failed: ${errMsg(e)}` }));
    }
  }

  if (error) return <ErrorBar error={error} onRetry={load} />;
  if (!groups) return <Loading label="Loading ignored emails…" />;
  if (!groups.length)
    return (
      <Empty title="No ignored emails">
        Use Ignore on a digest card to hide a conversation from the assistant.
      </Empty>
    );

  const me = myAddress().toLowerCase();
  return (
    <div className={s.stack}>
      <Caption1 className={s.muted}>
        {groups.length} ignored conversation{groups.length === 1 ? "" : "s"}. Unignored emails are
        included the next time you generate the digest.
      </Caption1>
      {groups.map((group) => {
        const m = group[0];
        const fromMe = m.from?.emailAddress?.address?.toLowerCase() === me;
        const who = fromMe
          ? "To: " +
            ((m.toRecipients ?? [])
              .map((r) => r.emailAddress.name || r.emailAddress.address)
              .join(", ") || "(no recipients)")
          : m.from?.emailAddress?.name || m.from?.emailAddress?.address || "(unknown sender)";
        const date = new Date(m.sentDateTime ?? m.receivedDateTime).toLocaleDateString(undefined, {
          month: "short",
          day: "numeric",
        });
        const st = status[m.conversationId];
        return (
          <Card key={m.conversationId} aria-label={`Ignored: ${m.subject ?? ""}`}>
            <div>
              <Body1 as="p" block style={{ margin: 0 }}>
                <strong>{who}</strong>
              </Body1>
              <Caption1>
                {m.subject || "(no subject)"} · {date}
              </Caption1>
            </div>
            <div className={s.row}>
              <Button
                size="small"
                icon={<OpenRegular />}
                onClick={() => openMessage(m.id, m.webLink)}
              >
                Open email
              </Button>
              <Button
                size="small"
                appearance="primary"
                icon={<ArrowUndoRegular />}
                disabled={st === "saving"}
                onClick={() => unignore(group)}
              >
                {st === "saving" ? "Unignoring…" : "Unignore"}
              </Button>
              {st && st !== "saving" && <Caption1 role="alert">{st}</Caption1>}
            </div>
          </Card>
        );
      })}
    </div>
  );
}
