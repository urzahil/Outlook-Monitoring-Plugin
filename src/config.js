export const GRAPH_SCOPES = ["Mail.Read.Shared"];

// Defaults for the settings panel; each user can change them in the task pane.
// The app (client) ID and tenant ID come from the Entra app registration (see README).
export const DEFAULT_SETTINGS = {
  clientId: "",
  tenantId: "",
  mailbox: "",
  teamDomain: "",
  excludedAddresses: "",
};
