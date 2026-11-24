const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const MAX_RETRIES = 5;
const MAX_WAIT_SECONDS = 60;
// Messages are filtered by sent time after download, so the server-side
// received-time window is widened to include mail that was in transit.
const TRANSIT_MARGIN_MS = 60 * 60 * 1000;
// Folders whose mail must not be counted: deleted mistakes, spoofed mail in
// Junk, and unsent or failed items.
const EXCLUDED_FOLDERS = ["deleteditems", "junkemail", "outbox", "drafts", "syncissues"];
const MESSAGE_FIELDS = "conversationId,internetMessageId,parentFolderId,subject,from,sender,sentDateTime,receivedDateTime,isDraft";

/** Seconds to wait from a Retry-After header, which is either a number of seconds or an HTTP date. */
export function parseRetryAfter(value, fallbackSeconds, now = Date.now()) {
  let seconds = NaN;
  if (value && /^\s*\d+(\.\d+)?\s*$/.test(value)) seconds = Number(value);
  else if (value) {
    const date = Date.parse(value);
    if (!Number.isNaN(date)) seconds = (date - now) / 1000;
  }
  if (!Number.isFinite(seconds)) seconds = fallbackSeconds;
  return Math.min(Math.max(seconds, 0), MAX_WAIT_SECONDS);
}

function sleep(ms, signal) {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) return reject(signal.reason ?? new DOMException("Cancelled", "AbortError"));
    const timer = setTimeout(() => {
      signal?.removeEventListener("abort", onAbort);
      resolve();
    }, ms);
    const onAbort = () => {
      clearTimeout(timer);
      reject(signal.reason ?? new DOMException("Cancelled", "AbortError"));
    };
    signal?.addEventListener("abort", onAbort, { once: true });
  });
}

async function graphGet(url, token, signal) {
  for (let attempt = 0; ; attempt++) {
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
      signal,
    });
    if (response.ok) return response.json();

    // Throttled or temporarily unavailable: wait as instructed and retry.
    if ([429, 503, 504].includes(response.status) && attempt < MAX_RETRIES) {
      await sleep(parseRetryAfter(response.headers.get("Retry-After"), 2 ** attempt) * 1000, signal);
      continue;
    }

    let detail = "";
    try {
      detail = (await response.json())?.error?.message || "";
    } catch {
      // Non-JSON error body.
    }
    const error =
      response.status === 403 || response.status === 404
        ? new Error(`Cannot read the mailbox (${response.status}). Check the address and that you have Full Access to it. ${detail}`)
        : new Error(`Microsoft Graph request failed (${response.status}). ${detail}`);
    error.status = response.status;
    throw error;
  }
}

const userUrl = (mailbox) => `${GRAPH_BASE}/users/${encodeURIComponent(mailbox)}`;

/** IDs of the well-known folders that must not be counted. Folders that don't exist are skipped. */
async function excludedFolderIds(mailbox, token, signal) {
  const ids = await Promise.all(
    EXCLUDED_FOLDERS.map(async (name) => {
      try {
        return (await graphGet(`${userUrl(mailbox)}/mailFolders/${name}?$select=id`, token, signal)).id;
      } catch (error) {
        // A mailbox without that folder is fine; anything else would silently let
        // Deleted Items or Junk be counted, so report it.
        if (error?.status === 404) return null;
        throw error;
      }
    })
  );
  return new Set(ids.filter(Boolean));
}

/**
 * Fetches the messages of `mailbox` sent in [start, end), from every folder
 * except Deleted Items, Junk Email, Outbox, Drafts and Sync Issues.
 * Only the fields needed for counting are requested.
 *
 * @returns {Promise<{messages: object[], excludedFolders: Set<string>}>}
 */
export async function fetchMessages({ mailbox, start, end, token, onProgress, signal }) {
  const excluded = await excludedFolderIds(mailbox, token, signal);
  const params = new URLSearchParams({
    $select: MESSAGE_FIELDS,
    $filter:
      `receivedDateTime ge ${new Date(start.getTime() - TRANSIT_MARGIN_MS).toISOString()} and ` +
      `receivedDateTime lt ${new Date(end.getTime() + TRANSIT_MARGIN_MS).toISOString()}`,
    $top: "500",
  });
  let url = `${userUrl(mailbox)}/messages?${params}`;

  const messages = [];
  let fetched = 0;
  while (url) {
    const page = await graphGet(url, token, signal);
    fetched += page.value.length;
    for (const message of page.value) {
      if (excluded.has(message.parentFolderId)) continue;
      // Count by sent time, like the desktop macro does.
      if (message.sentDateTime) {
        const sent = Date.parse(message.sentDateTime);
        if (sent < start.getTime() || sent >= end.getTime()) continue;
      }
      messages.push(message);
    }
    onProgress?.(fetched);
    url = page["@odata.nextLink"];
  }
  return { messages, excludedFolders: excluded };
}

/**
 * Finds the first message of each conversation, even when it is older than
 * the reporting period. Used by the "only threads started in this period" and
 * "ignore threads started by your team" filters.
 *
 * @returns {Promise<Map<string, object>>} conversationId -> earliest message
 */
export async function fetchThreadStarts({ mailbox, conversationIds, excludedFolders, token, onProgress, signal }) {
  const starts = new Map();
  const queue = [...conversationIds];
  let done = 0;

  async function worker() {
    while (queue.length) {
      const id = queue.shift();
      const params = new URLSearchParams({
        $select: MESSAGE_FIELDS,
        $filter: `conversationId eq '${id.replace(/'/g, "''")}'`,
        $top: "100",
      });
      let first = null;
      for (let url = `${userUrl(mailbox)}/messages?${params}`; url; ) {
        const page = await graphGet(url, token, signal);
        for (const message of page.value) {
          if (message.isDraft || !message.sentDateTime || excludedFolders?.has(message.parentFolderId)) continue;
          if (!first || message.sentDateTime < first.sentDateTime) first = message;
        }
        url = page["@odata.nextLink"];
      }
      if (first) starts.set(id, first);
      onProgress?.(++done, conversationIds.size);
    }
  }

  try {
    await Promise.all(Array.from({ length: 4 }, worker));
  } catch (error) {
    queue.length = 0; // stop the other workers
    throw error;
  }
  return starts;
}
