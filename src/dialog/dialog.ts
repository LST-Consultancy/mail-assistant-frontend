/* global Office */
import { PublicClientApplication } from "@azure/msal-browser";
import { apiScopes, AUTHORITY, GRAPH_SCOPES } from "../services/auth";
import { env } from "../services/env";

// Runs inside the Office dialog (fallback auth for clients without NAA). Redirect flow:
// first load -> loginRedirect; back from Entra -> handleRedirectPromise -> post tokens to the task pane.
async function run() {
  await Office.onReady();
  const pca = new PublicClientApplication({
    auth: {
      clientId: env.clientId,
      authority: AUTHORITY,
      redirectUri: `${location.origin}/dialog.html`,
    },
    cache: { cacheLocation: "sessionStorage" },
  });
  await pca.initialize();
  const result = await pca.handleRedirectPromise();
  if (!result) {
    await pca.loginRedirect({ scopes: GRAPH_SCOPES, extraScopesToConsent: apiScopes() });
    return;
  }
  const api = await pca.acquireTokenSilent({ scopes: apiScopes(), account: result.account });
  const pack = (r: { accessToken: string; expiresOn: Date | null }) => ({
    token: r.accessToken,
    expiresOn: r.expiresOn?.getTime() ?? Date.now() + 3_000_000,
  });
  Office.context.ui.messageParent(JSON.stringify({ graph: pack(result), api: pack(api) }));
}

run().catch((err: Error) => {
  try {
    Office.context.ui.messageParent(JSON.stringify({ error: err.message || "Sign-in failed." }));
  } catch {
    document.getElementById("status")!.textContent =
      "Sign-in failed. Close this window and try again.";
  }
});
