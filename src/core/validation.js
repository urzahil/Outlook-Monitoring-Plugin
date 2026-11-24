/** Input checks for the settings panel. Each returns an error message, or "" when the value is fine. */

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const DOMAIN = /^(?=.{1,253}$)([a-z0-9]([a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,}$/i;
const EMAIL = /^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/;

export function validateClientId(value) {
  return GUID.test(value) ? "" : "The app (client) ID must be a GUID such as 00000000-0000-0000-0000-000000000000.";
}

/** Optional. A GUID, a domain, or one of the special values Entra accepts. */
export function validateTenantId(value) {
  if (!value || GUID.test(value) || DOMAIN.test(value) || ["organizations", "common", "consumers"].includes(value)) return "";
  return "The tenant ID must be a GUID or a domain such as contoso.onmicrosoft.com.";
}

export function validateMailbox(value) {
  return EMAIL.test(value) ? "" : "The shared mailbox must be an email address such as support@yourcompany.com.";
}

/** A bare domain such as "yourcompany.com" (a leading "@" is tolerated). Full addresses and URLs are rejected. */
export function validateTeamDomain(value) {
  const domain = value.replace(/^@/, "");
  if (domain.includes("@")) return 'Enter only the domain (for example "yourcompany.com"), not a full email address.';
  if (/[/:?#\s]/.test(domain)) return 'Enter only the domain (for example "yourcompany.com"), not a URL.';
  return DOMAIN.test(domain) ? "" : 'The team domain must look like "yourcompany.com".';
}

/** Comma-separated list of email addresses (may be empty). */
export function validateAddressList(value, label) {
  const bad = splitList(value).find((entry) => !EMAIL.test(entry));
  return bad ? `${label}: "${bad}" is not an email address.` : "";
}

export function splitList(value) {
  return value.split(",").map((entry) => entry.trim()).filter(Boolean);
}

/** Validates all settings; returns the first problem as { field, message } or null. */
export function validateSettings({ clientId, tenantId, mailbox, teamDomain, excludedAddresses, allowedAddresses }) {
  const checks = [
    ["clientId", validateClientId(clientId)],
    ["tenantId", validateTenantId(tenantId)],
    ["mailbox", validateMailbox(mailbox)],
    ["teamDomain", validateTeamDomain(teamDomain)],
    ["excludedAddresses", validateAddressList(excludedAddresses, "Addresses to ignore")],
    ["allowedAddresses", validateAddressList(allowedAddresses, "Only include these team addresses")],
  ];
  const failed = checks.find(([, message]) => message);
  return failed ? { field: failed[0], message: failed[1] } : null;
}
