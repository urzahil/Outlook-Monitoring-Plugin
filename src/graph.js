const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const MAX_RETRIES = 5;

async function graphGet(url, token) {
  for (let attempt = 0; ; attempt++) {
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (response.ok) return response.json();

    // Throttled or temporarily unavailable: wait as instructed and retry.
    if ((response.status === 429 || response.status === 503) && attempt < MAX_RETRIES) {
      const retryAfter = Number(response.headers.get("Retry-After")) || 2 ** attempt;
      await new Promise((resolve) => setTimeout(resolve, retryAfter * 1000));
      continue;
    }

    let detail = "";
    try {
      detail = (await response.json())?.error?.message || "";
    } catch {
      // Non-JSON error body.
    }
    if (response.status === 403 || response.status === 404) {
      throw new Error(`Cannot read the mailbox (${response.status}). Check the address and that you have Full Access to it. ${detail}`);
    }
    throw new Error(`Microsoft Graph request failed (${response.status}). ${detail}`);
  }
}

/**
 * Fetches every message in all folders of `mailbox` received in [start, end).
 * Only the fields needed for counting are requested.
 */
export async function fetchMessages({ mailbox, start, end, token, onProgress }) {
  const params = new URLSearchParams({
    $select: "conversationId,subject,from,sender,sentDateTime,receivedDateTime,isDraft",
    $filter: `receivedDateTime ge ${start.toISOString()} and receivedDateTime lt ${end.toISOString()}`,
    $top: "500",
  });
  let url = `${GRAPH_BASE}/users/${encodeURIComponent(mailbox)}/messages?${params}`;

  const messages = [];
  while (url) {
    const page = await graphGet(url, token);
    messages.push(...page.value);
    onProgress?.(messages.length);
    url = page["@odata.nextLink"];
  }
  return messages;
}
