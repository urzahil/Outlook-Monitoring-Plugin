import { test } from "node:test";
import assert from "node:assert/strict";
import { countCases, identifyAgent, toCsv } from "../src/core/counting.js";

const options = { teamDomain: "contoso.com", excludedAddresses: ["support@contoso.com"] };

const addr = (address, name = address) => ({ emailAddress: { address, name } });
const msg = (conversationId, fromAddress, extra = {}) => ({
  conversationId,
  subject: `Subject ${conversationId}`,
  from: addr(fromAddress),
  sender: addr(fromAddress),
  sentDateTime: "2026-10-01T10:00:00Z",
  isDraft: false,
  ...extra,
});

test("several replies by the same person on one thread count as one case", () => {
  const result = countCases([msg("A", "alice@contoso.com"), msg("A", "alice@contoso.com"), msg("A", "alice@contoso.com")], options);
  assert.equal(result.people.length, 1);
  assert.equal(result.people[0].cases, 1);
  assert.equal(result.people[0].messages, 3);
  assert.equal(result.people[0].threads[0].replies, 3);
});

test("each person working on a shared thread gets one case", () => {
  const result = countCases([msg("A", "alice@contoso.com"), msg("A", "bob@contoso.com"), msg("B", "alice@contoso.com")], options);
  const byAddress = Object.fromEntries(result.people.map((p) => [p.address, p.cases]));
  assert.deepEqual(byAddress, { "alice@contoso.com": 2, "bob@contoso.com": 1 });
  assert.equal(result.uniqueCases, 2);
  assert.equal(result.people[0].address, "alice@contoso.com", "sorted by cases, highest first");
});

test("customer messages and drafts are ignored", () => {
  const result = countCases(
    [msg("A", "customer@example.com"), msg("B", "alice@contoso.com", { isDraft: true })],
    options
  );
  assert.equal(result.people.length, 0);
  assert.equal(result.uniqueCases, 0);
});

test("send-on-behalf replies are credited to the person in `sender`", () => {
  const message = msg("A", "support@contoso.com", { sender: addr("Carol@Contoso.com", "Carol") });
  assert.deepEqual(identifyAgent(message, options), { address: "carol@contoso.com", name: "Carol" });
});

test("send-as replies from the shared mailbox are reported as unattributed", () => {
  const result = countCases([msg("A", "support@contoso.com")], options);
  assert.equal(result.people.length, 0);
  assert.equal(result.unattributedMessages, 1);
});

test("CSV export lists one row per person", () => {
  const csv = toCsv(countCases([msg("A", "alice@contoso.com")], options));
  assert.equal(csv, '"Name","Email","Cases","Messages sent"\r\n"alice@contoso.com","alice@contoso.com","1","1"');
});

test("nobody is counted when no team domain is set", () => {
  const result = countCases([msg("A", "alice@contoso.com")], { teamDomain: "" });
  assert.equal(result.people.length, 0);
  assert.equal(result.unattributedMessages, 0);
});
