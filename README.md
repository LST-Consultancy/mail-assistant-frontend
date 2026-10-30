# Mail Assistant: Outlook add-in (AI summaries and replies)

Mail Assistant adds a ribbon button to Outlook that opens a side panel. The panel can summarize the open email or thread, suggest a reply in a chosen tone, build a digest of the last 7, 14 or 30 days of the inbox, and help draft text while composing. It uses an AI model through a small backend, which lives in its own repo (`mail-assistant-backend`).

> **AI provider:** the paired backend uses OpenAI. The API key stays on the backend; never put it in this add-in's `.env` (its values are inlined into the bundle).

**It never sends email.** Every mailbox action is an explicit click by the user:
- **Open as reply** opens Outlook's own reply window.
- **Create drafts** saves drafts to the Drafts folder.
- **Insert** and **Replace body** edit the draft that is open.

The code contains no Graph `send` call.

`REQUIREMENTS.md` is the spec and the source of truth. Its section 16 lists the decisions taken and where the build differs from the spec.

Stack: React 18 + TypeScript + Fluent UI v9, webpack (official yo office template), XML manifest.

## How it works

```
Outlook task pane ── MSAL (NAA, or Office dialog fallback) ──► Entra ID
     │  Graph token ──► Microsoft Graph   (read mail, thread, createReply drafts)
     │  API token   ──► Backend /api/*    (trimmed email text only, never Graph tokens)
```

## 1. Entra ID app registration

The add-in and the backend share one app registration. The full setup steps are in the backend repo's README, section 1. For this repo you need:
- The **Application (client) ID**, as `VITE_ENTRA_CLIENT_ID` in `.env`.
- These SPA redirect URIs on the registration: `brk-multihub://localhost:3000` and `https://localhost:3000/dialog.html` for dev, plus the same two with your production add-in domain.

## 2. Local development

Prerequisites:
- Node 20 or later (22 recommended).
- The backend running on `https://localhost:8080`. See the backend repo.
- A Microsoft 365 or Outlook.com mailbox. Not needed for mock mode.

```bash
cp .env.example .env                       # fill in the values
npm install
npx office-addin-dev-certs install         # trusted localhost cert, also reused by the backend
npm start                                  # webpack dev server on https://localhost:3000 + sideloads manifest.xml
```

`.env` is read at build time and inlined into the bundle, so it must never hold secrets. Restart `npm start` after changing it.

**Sideloading:**
- **Classic Outlook (Windows):** `npm start` sideloads automatically.
- **Outlook on the web and new Outlook:** open <https://aka.ms/olksideload> → *My add-ins* → *Custom add-ins* → *Add from file*, then choose `manifest.xml`.
- **Then:** open any email. The **Mail Assistant** button appears on the ribbon in read mode and in compose mode.

### Mock mode (no mailbox, no sign-in)

```bash
npm run mock                       # VITE_MOCK=1; run the backend with DEV_AUTH_BYPASS=1
```

Open these pages in a browser:
- <https://localhost:3000/taskpane.html?tab=read> for the read tabs.
- <https://localhost:3000/taskpane.html?tab=write> for the compose tab.

Fixture emails stand in for Outlook and Graph, and the AI calls are real.

### Tests and checks

```bash
npm test && npm run typecheck && npm run lint && npm run validate
```

The tests cover Graph paging, `Retry-After` and backoff, thread grouping and dedupe, the "already replied" filter, and SSE parsing.

## 3. Deployment

```bash
ADDIN_PROD_URL=https://addin.example.com/ VITE_API_BASE_URL=https://api.example.com npm run build
```

Upload `dist/` to any static HTTPS host, such as Azure Static Web Apps or a Storage static website. The built `dist/manifest.xml` already points at `ADDIN_PROD_URL`. Add the production redirect URIs to the app registration, and set the backend's `ADDIN_ORIGIN` to this origin.

**Org-wide rollout:**
1. Go to **Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps**.
2. Choose *Office Add-in* and upload `dist/manifest.xml`.
3. Assign it to users or groups. It can take up to 24 hours to appear in Outlook.

## 4. Client support notes

- **Manifest format:** add-in only XML. The unified JSON manifest doesn't support new Outlook for Mac yet. XML works on all four target clients: new and classic Outlook for Windows, Outlook on the web, and new Outlook for Mac.
- **NAA** works in Outlook on the web, Outlook for Windows (Version 2409 or later) and Outlook for Mac (16.89 or later). **It isn't supported in Outlook.com mailboxes.** Those clients, and any older ones, use the Office dialog sign-in, which shows a popup once per session.
- **Pinning** isn't available for Outlook.com accounts. There, the pane closes when you switch emails.
- **"Open email"** uses `convertToEwsId` and `displayMessageForm`, and falls back to the message's `webLink`.
- **Outlook mobile** isn't supported in v1.

## 5. Configuration

See [.env.example](.env.example).
