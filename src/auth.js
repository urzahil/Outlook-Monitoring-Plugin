import { createNestablePublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";
import { GRAPH_SCOPES } from "./config.js";

let client = null;
let clientKey = "";

// Uses Nested App Authentication (single sign-on through Outlook). MSAL falls
// back to a regular popup sign-in on hosts that don't support NAA.
function getClient(clientId, tenantId) {
  const key = `${clientId}/${tenantId}`;
  if (key !== clientKey) {
    client = createNestablePublicClientApplication({
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId || "organizations"}`,
      },
    });
    clientKey = key;
  }
  return client;
}

export async function getGraphToken({ clientId, tenantId }) {
  const msal = await getClient(clientId, tenantId);
  const request = { scopes: GRAPH_SCOPES };
  try {
    const result = await msal.acquireTokenSilent(request);
    return result.accessToken;
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError) && error?.errorCode !== "no_account_error") {
      console.warn("Silent token request failed, trying interactive sign-in.", error);
    }
    const result = await msal.acquireTokenPopup(request);
    return result.accessToken;
  }
}
