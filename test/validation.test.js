import { test } from "node:test";
import assert from "node:assert/strict";
import { validateClientId, validateMailbox, validateSettings, validateTeamDomain } from "../src/core/validation.js";

test("client ID must be a GUID", () => {
  assert.equal(validateClientId("6f1d2c7e-3b4a-4e8f-9a51-2d7c0b9e4f13"), "");
  assert.notEqual(validateClientId("not-a-guid"), "");
});

test("mailbox must be an email address", () => {
  assert.equal(validateMailbox("support@contoso.com"), "");
  assert.notEqual(validateMailbox("support"), "");
});

test("team domain rejects emails and URLs", () => {
  assert.equal(validateTeamDomain("contoso.com"), "");
  assert.equal(validateTeamDomain("@contoso.com"), "");
  assert.notEqual(validateTeamDomain("user@contoso.com"), "");
  assert.notEqual(validateTeamDomain("https://contoso.com"), "");
});

test("validateSettings reports the first problem and its field", () => {
  const ok = { clientId: "6f1d2c7e-3b4a-4e8f-9a51-2d7c0b9e4f13", tenantId: "", mailbox: "s@contoso.com", teamDomain: "contoso.com", excludedAddresses: "", allowedAddresses: "a@contoso.com, b@contoso.com" };
  assert.equal(validateSettings(ok), null);
  assert.equal(validateSettings({ ...ok, teamDomain: "a@b.com" }).field, "teamDomain");
  assert.equal(validateSettings({ ...ok, allowedAddresses: "nope" }).field, "allowedAddresses");
});
