import { test } from "node:test";
import assert from "node:assert/strict";
import { agentConversationIds, countCases, identifyAgent, toCsv } from "../src/core/counting.js";

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

test("duplicate copies of a message (same internetMessageId) are counted once", () => {
  const a = msg("A", "alice@contoso.com", { internetMessageId: "<1@x>" });
  const copy = msg("A", "alice@contoso.com", { internetMessageId: "<1@x>" });
  const result = countCases([a, copy], options);
  assert.equal(result.people[0].messages, 1);
  assert.equal(result.duplicates, 1);
});

test("allowlist limits who is credited and does not flag other colleagues as unattributed", () => {
  const withList = { ...options, allowedAddresses: ["alice@contoso.com"] };
  const result = countCases([msg("A", "alice@contoso.com"), msg("B", "dave@contoso.com")], withList);
  assert.deepEqual(result.people.map((p) => p.address), ["alice@contoso.com"]);
  assert.equal(result.unattributedMessages, 0);
});

test("a later display name replaces a bare address", () => {
  const result = countCases(
    [msg("A", "alice@contoso.com", { sender: addr("alice@contoso.com", "") }), msg("B", "alice@contoso.com", { sender: addr("alice@contoso.com", "Alice Smith") })],
    options
  );
  assert.equal(result.people[0].name, "Alice Smith");
});

test("thread filters skip old threads and threads started by the team", () => {
  const messages = [msg("OLD", "alice@contoso.com"), msg("OWN", "alice@contoso.com"), msg("NEW", "alice@contoso.com")];
  const threadStarts = new Map([
    ["OLD", { sentDateTime: "2026-01-01T00:00:00Z", from: addr("customer@example.com") }],
    ["OWN", { sentDateTime: "2026-02-10T00:00:00Z", from: addr("bob@contoso.com") }],
    ["NEW", { sentDateTime: "2026-02-10T00:00:00Z", from: addr("customer@example.com") }],
  ]);
  const result = countCases(messages, { ...options, onlyNewThreads: true, skipOwnThreads: true, periodStart: new Date("2026-02-01T00:00:00Z"), threadStarts });
  assert.equal(result.uniqueCases, 1);
  assert.deepEqual(result.skippedThreads, { old: 1, own: 1 });
  assert.deepEqual([...agentConversationIds(messages, options)].sort(), ["NEW", "OLD", "OWN"]);
});

test("CSV cells that look like formulas are neutralised", () => {
  const csv = toCsv({ people: [{ name: "=HYPERLINK(\"x\")", address: "+a@contoso.com", cases: 1, messages: 1 }] });
  assert.match(csv, /"'=HYPERLINK\(""x""\)","'\+a@contoso.com"/);
});
