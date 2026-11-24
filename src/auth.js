import { createNestablePublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";
import { GRAPH_SCOPES } from "./config.js";

let client = null;
let clientKey = "";

// Uses Nested App Authentication (single sign-on through Outlook). MSAL falls
// back to a regular popup sign-in on hosts that don't support NAA.
function getClient(clientId, tenantId) {
  const key = `${clientId}/${tenantId}`;
  if (key !== clientKey) {
    const promise = createNestablePublicClientApplication({
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId || "organizations"}`,
        // Outlook appends ?_host_Info=... to the task pane URL; the registered
        // single-page-application redirect URI has to match without it.
        redirectUri: new URL("taskpane.html", window.location.href).href,
      },
    });
    client = promise;
    clientKey = key;
    // Don't keep a failed initialisation (for example a network error) cached.
    promise.catch(() => {
      if (client === promise) {
        client = null;
        clientKey = "";
      }
    });
  }
  return client;
}

export async function getGraphToken({ clientId, tenantId }) {
  const msal = await getClient(clientId, tenantId);
  // Without an account, the non-NAA fallback can't find a cached sign-in and
  // would prompt on every run.
  const account = msal.getActiveAccount() ?? msal.getAllAccounts().find((a) => !tenantId || a.tenantId === tenantId) ?? msal.getAllAccounts()[0];
  const request = account ? { scopes: GRAPH_SCOPES, account } : { scopes: GRAPH_SCOPES };
  try {
    const result = await msal.acquireTokenSilent(request);
    return result.accessToken;
  } catch (error) {
    // Only a missing sign-in needs the popup; real failures (network, configuration) surface as they are.
    if (!(error instanceof InteractionRequiredAuthError) && error?.errorCode !== "no_account_error") throw error;
    const result = await msal.acquireTokenPopup({ scopes: GRAPH_SCOPES });
    if (result.account) msal.setActiveAccount(result.account);
    return result.accessToken;
  }
}
