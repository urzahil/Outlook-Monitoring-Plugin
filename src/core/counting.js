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
export function identifyAgent(message, { teamDomain, excludedAddresses = [], allowedAddresses = [] }) {
  const suffix = domainSuffix(teamDomain);
  if (!suffix) return null;
  const excluded = new Set(excludedAddresses.map(normalize));
  const allowed = new Set(allowedAddresses.map(normalize).filter(Boolean));

  for (const field of [message.sender, message.from]) {
    const address = normalize(field?.emailAddress?.address);
    if (address.endsWith(suffix) && !excluded.has(address) && (allowed.size === 0 || allowed.has(address))) {
      return { address, name: field.emailAddress.name || address };
    }
  }
  return null;
}

/** True when the message was sent "as" an excluded team address (typically the shared mailbox), so no person can be credited. */
function isUnattributedTeamMessage(message, { teamDomain, excludedAddresses = [] }) {
  const suffix = domainSuffix(teamDomain);
  if (!suffix) return false;
  const excluded = new Set(excludedAddresses.map(normalize));
  return [message.sender, message.from].some((field) => {
    const address = normalize(field?.emailAddress?.address);
    return address.endsWith(suffix) && excluded.has(address);
  });
}

/** True when either `sender` or `from` of the message belongs to the team domain. */
function isFromTeam(message, teamDomain) {
  const suffix = domainSuffix(teamDomain);
  if (!suffix) return false;
  return [message.sender, message.from].some((field) => normalize(field?.emailAddress?.address).endsWith(suffix));
}

/**
 * Conversation IDs that would be credited to somebody. Use it to decide which
 * thread starts need to be looked up when the thread filters are on.
 */
export function agentConversationIds(messages, options) {
  const ids = new Set();
  for (const message of messages) {
    if (message.isDraft || !message.conversationId) continue;
    if (identifyAgent(message, options)) ids.add(message.conversationId);
  }
  return ids;
}

/**
 * Decides whether a thread is skipped, given its first message:
 * "old" = started before the period, "own" = started by the team.
 * Returns null when the thread is counted (or its start is unknown).
 */
function threadSkipReason(start, { teamDomain, periodStart, onlyNewThreads, skipOwnThreads }) {
  if (!start) return null;
  if (onlyNewThreads && periodStart && start.sentDateTime && new Date(start.sentDateTime) < periodStart) return "old";
  if (skipOwnThreads && isFromTeam(start, teamDomain)) return "own";
  return null;
}

/**
 * @param {Array<object>} messages Graph message resources with from, sender,
 *   conversationId, subject, sentDateTime and isDraft selected.
 *   internetMessageId is used to ignore copies of a message filed in several folders.
 * @param {object} options
 * @param {string} options.teamDomain
 * @param {string[]} [options.excludedAddresses] team addresses that are not people
 * @param {string[]} [options.allowedAddresses] when not empty, only these addresses are credited
 * @param {boolean} [options.onlyNewThreads] skip threads that started before `periodStart`
 * @param {boolean} [options.skipOwnThreads] skip threads that were started by the team
 * @param {Date} [options.periodStart]
 * @param {Map<string, object>} [options.threadStarts] conversationId -> first message of the thread
 */
export function countCases(messages, options) {
  const agents = new Map();
  const allCases = new Set();
  const seenMessageIds = new Set();
  const skippedThreads = { old: 0, own: 0 };
  const verdicts = new Map();
  let unattributedMessages = 0;
  let duplicates = 0;

  for (const message of messages) {
    if (message.isDraft || !message.conversationId) continue;

    // The same message can be filed in several folders (rules, CC to the shared mailbox).
    if (message.internetMessageId) {
      if (seenMessageIds.has(message.internetMessageId)) {
        duplicates++;
        continue;
      }
      seenMessageIds.add(message.internetMessageId);
    }

    const agent = identifyAgent(message, options);
    if (!agent) {
      if (isUnattributedTeamMessage(message, options)) unattributedMessages++;
      continue;
    }

    if (!verdicts.has(message.conversationId)) {
      const reason = threadSkipReason(options.threadStarts?.get(message.conversationId), options);
      verdicts.set(message.conversationId, reason);
      if (reason) skippedThreads[reason]++;
    }
    if (verdicts.get(message.conversationId)) continue;

    let entry = agents.get(agent.address);
    if (!entry) {
      entry = { address: agent.address, name: agent.name, messages: 0, threads: new Map() };
      agents.set(agent.address, entry);
    } else if (entry.name === entry.address && agent.name !== agent.address) {
      // The first message had no display name; use the first real one.
      entry.name = agent.name;
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

  return { people, uniqueCases: allCases.size, unattributedMessages, duplicates, skippedThreads };
}

/**
 * Quotes a CSV text cell. Cells that a spreadsheet could read as a formula
 * (display names come from email headers) get a leading apostrophe.
 */
export function csvText(value) {
  let text = String(value);
  if (/^[=+\-@\t\r]/.test(text)) text = "'" + text;
  return `"${text.replace(/"/g, '""')}"`;
}

/** Builds a CSV summary (one row per person). */
export function toCsv(result) {
  const rows = [['"Name"', '"Email"', '"Cases"', '"Messages sent"']];
  for (const person of result.people) {
    rows.push([csvText(person.name), csvText(person.address), `"${Number(person.cases)}"`, `"${Number(person.messages)}"`]);
  }
  return rows.map((row) => row.join(",")).join("\r\n");
}
