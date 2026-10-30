/* global Office */
import {
  createNestablePublicClientApplication,
  InteractionRequiredAuthError,
  type IPublicClientApplication,
} from "@azure/msal-browser";
import { env } from "./env";

export const GRAPH_SCOPES = ["User.Read", "Mail.ReadWrite"];
export const apiScopes = () => [`api://${env.clientId}/access_as_user`];
/** "common": work/school accounts of the allowed tenant plus personal Microsoft accounts (backend enforces the tenant list). */
export const AUTHORITY = "https://login.microsoftonline.com/common";

type Resource = "graph" | "api";
interface CachedToken {
  token: string;
  expiresOn: number;
}

let naa: Promise<IPublicClientApplication> | undefined;
const cache: Partial<Record<Resource, CachedToken>> = {};
let dialogInFlight: Promise<void> | undefined;

const supportsNaa = () =>
  Office.context.requirements?.isSetSupported("NestedAppAuth", "1.1") ?? false;

function getNaa(): Promise<IPublicClientApplication> {
  naa ??= createNestablePublicClientApplication({
    auth: { clientId: env.clientId, authority: AUTHORITY },
    cache: { cacheLocation: "localStorage" },
  });
  return naa;
}

async function viaNaa(scopes: string[]): Promise<string> {
  const pca = await getNaa();
  try {
    return (await pca.acquireTokenSilent({ scopes })).accessToken;
  } catch (err) {
    if (!(err instanceof InteractionRequiredAuthError)) throw err;
    return (await pca.acquireTokenPopup({ scopes })).accessToken;
  }
}

/**
 * Fallback for clients without NAA (e.g. Outlook.com mailboxes): dialog.html runs a standard MSAL
 * redirect sign-in and posts both tokens back with messageParent. Tokens live in memory only.
 */
function viaDialog(): Promise<void> {
  dialogInFlight ??= new Promise<void>((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      `${location.origin}/dialog.html`,
      { height: 60, width: 30 },
      (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded)
          return reject(new Error(res.error.message));
        const dialog = res.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          dialog.close();
          try {
            const msg = JSON.parse((arg as { message: string }).message) as
              { graph: CachedToken; api: CachedToken } | { error: string };
            if ("error" in msg) return reject(new Error(msg.error));
            cache.graph = msg.graph;
            cache.api = msg.api;
            resolve();
          } catch {
            reject(new Error("Sign-in failed."));
          }
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () =>
          reject(new Error("Sign-in was cancelled."))
        );
      }
    );
  }).finally(() => (dialogInFlight = undefined));
  return dialogInFlight;
}

async function getToken(resource: Resource): Promise<string> {
  if (env.mock) return "dev";
  if (!env.clientId) throw new Error("VITE_ENTRA_CLIENT_ID is not configured.");
  const scopes = resource === "graph" ? GRAPH_SCOPES : apiScopes();
  if (supportsNaa()) return viaNaa(scopes);
  const hit = cache[resource];
  if (hit && hit.expiresOn - Date.now() > 60_000) return hit.token;
  await viaDialog();
  return cache[resource]!.token;
}

export const getGraphToken = () => getToken("graph");
export const getApiToken = () => getToken("api");
