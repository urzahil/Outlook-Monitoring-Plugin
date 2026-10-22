/**
 * Case-counting rules.
 *
 * A "case" is an email thread (Graph `conversationId`). A team member handles a
 * case when they send at least one message on that thread in the period:
 *   - several replies by the same person on one thread count once;
 *   - if several people reply on the same thread, each of them gets one case.
 */

const normalize = (address) => (address || "").trim().toLowerCase();

/** "@domain" for a configured team domain such as "example.com" or "@example.com", or null if none is set. */
function domainSuffix(teamDomain) {
  const domain = normalize(teamDomain).replace(/^@/, "");
  return domain ? "@" + domain : null;
}

/**
 * Works out which team member sent a message, or returns null.
 *
 * `sender` is checked before `from` so "Send on behalf of" replies (where
 * `from` is the shared mailbox and `sender` is the person) are attributed to
 * the person.
 */
export function identifyAgent(message, { teamDomain, excludedAddresses = [] }) {
  const suffix = domainSuffix(teamDomain);
  if (!suffix) return null;
  const excluded = new Set(excludedAddresses.map(normalize));

  for (const field of [message.sender, message.from]) {
    const address = normalize(field?.emailAddress?.address);
    if (address.endsWith(suffix) && !excluded.has(address)) {
      return { address, name: field.emailAddress.name || address };
    }
  }
  return null;
}

/** True when the message was sent by the team domain but no person can be identified (e.g. "Send As" the shared mailbox). */
function isUnattributedTeamMessage(message, teamDomain) {
  const suffix = domainSuffix(teamDomain);
  if (!suffix) return false;
  return [message.sender, message.from].some((field) =>
    normalize(field?.emailAddress?.address).endsWith(suffix)
  );
}

/**
 * @param {Array<object>} messages Graph message resources with from, sender,
 *   conversationId, subject, sentDateTime and isDraft selected.
 * @param {{teamDomain: string, excludedAddresses?: string[]}} options
 */
export function countCases(messages, options) {
  const agents = new Map();
  const allCases = new Set();
  let unattributedMessages = 0;

  for (const message of messages) {
    if (message.isDraft || !message.conversationId) continue;

    const agent = identifyAgent(message, options);
    if (!agent) {
      if (isUnattributedTeamMessage(message, options.teamDomain)) unattributedMessages++;
      continue;
    }

    let entry = agents.get(agent.address);
    if (!entry) {
      entry = { address: agent.address, name: agent.name, messages: 0, threads: new Map() };
      agents.set(agent.address, entry);
    }
    entry.messages++;

    let thread = entry.threads.get(message.conversationId);
    if (!thread) {
      thread = { conversationId: message.conversationId, subject: message.subject || "(no subject)", replies: 0, lastSent: null };
      entry.threads.set(message.conversationId, thread);
    }
    thread.replies++;
    if (!thread.lastSent || message.sentDateTime > thread.lastSent) thread.lastSent = message.sentDateTime;

    allCases.add(message.conversationId);
  }

  const people = [...agents.values()]
    .map((entry) => ({
      address: entry.address,
      name: entry.name,
      cases: entry.threads.size,
      messages: entry.messages,
      threads: [...entry.threads.values()].sort((a, b) => (b.lastSent || "").localeCompare(a.lastSent || "")),
    }))
    .sort((a, b) => b.cases - a.cases || a.name.localeCompare(b.name));

  return { people, uniqueCases: allCases.size, unattributedMessages };
}

/** Builds a CSV summary (one row per person). */
export function toCsv(result) {
  const escape = (value) => `"${String(value).replace(/"/g, '""')}"`;
  const rows = [["Name", "Email", "Cases", "Messages sent"]];
  for (const person of result.people) rows.push([person.name, person.address, person.cases, person.messages]);
  return rows.map((row) => row.map(escape).join(",")).join("\r\n");
}
