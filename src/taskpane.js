import { getGraphToken } from "./auth.js";
import { DEFAULT_SETTINGS } from "./config.js";
import { countCases, toCsv } from "./core/counting.js";
import { getDateRange, toMonthValue } from "./core/dateRanges.js";
import { fetchMessages } from "./graph.js";

const $ = (id) => document.getElementById(id);
const SETTING_KEYS = Object.keys(DEFAULT_SETTINGS);

let lastResult = null;
let lastLabel = "";

// Settings roam with the user's Exchange mailbox, so they follow them across devices.
function loadSettings() {
  const roaming = Office.context.roamingSettings;
  for (const key of SETTING_KEYS) $(key).value = roaming.get(key) ?? DEFAULT_SETTINGS[key];
}

function saveSettings() {
  const roaming = Office.context.roamingSettings;
  for (const key of SETTING_KEYS) roaming.set(key, $(key).value.trim());
  roaming.saveAsync((result) => {
    setStatus(result.status === Office.AsyncResultStatus.Succeeded ? "Settings saved." : "Could not save settings.");
  });
}

function setStatus(text, isError = false) {
  $("status").textContent = text;
  $("status").classList.toggle("error", isError);
}

function selectedPeriod() {
  return document.querySelector('input[name="period"]:checked').value;
}

async function run() {
  const clientId = $("clientId").value.trim();
  const tenantId = $("tenantId").value.trim();
  const mailbox = $("mailbox").value.trim();
  const teamDomain = $("teamDomain").value.trim();
  if (!clientId || !mailbox || !teamDomain) {
    $("settings").open = true;
    setStatus("Enter the app ID, the shared mailbox address and the team email domain in Settings.", true);
    return;
  }

  $("run").disabled = true;
  $("results").hidden = true;
  try {
    const range = getDateRange(selectedPeriod(), { month: $("month").value });
    setStatus("Signing in…");
    const token = await getGraphToken({ clientId, tenantId });

    setStatus("Reading messages…");
    const messages = await fetchMessages({
      mailbox,
      start: range.start,
      end: range.end,
      token,
      onProgress: (count) => setStatus(`Reading messages… ${count}`),
    });

    const excludedAddresses = [mailbox, ...$("excludedAddresses").value.split(",")].map((a) => a.trim()).filter(Boolean);
    lastResult = countCases(messages, { teamDomain, excludedAddresses });
    lastLabel = range.label;
    render(lastResult, range, messages.length);
    setStatus("");
  } catch (error) {
    console.error(error);
    setStatus(error.message || String(error), true);
  } finally {
    $("run").disabled = false;
  }
}

function render(result, range, scanned) {
  const dateFormat = { day: "numeric", month: "short", year: "numeric" };
  $("resultsTitle").textContent = range.label;
  $("summary").textContent =
    `${range.start.toLocaleDateString(undefined, dateFormat)} – ${range.end.toLocaleDateString(undefined, dateFormat)} · ` +
    `${result.uniqueCases} distinct cases handled · ${scanned} messages scanned`;

  $("warning").hidden = result.unattributedMessages === 0;
  $("warning").textContent =
    `${result.unattributedMessages} team ${result.unattributedMessages === 1 ? "reply was" : "replies were"} sent "as" the shared mailbox, ` +
    `so the sender can't be identified and ${result.unattributedMessages === 1 ? "it is" : "they are"} not counted. Use "Send on behalf" or reply from personal addresses to track them.`;

  const rows = $("rows");
  rows.replaceChildren();
  if (result.people.length === 0) {
    const tr = rows.insertRow();
    const td = tr.insertCell();
    td.colSpan = 3;
    td.textContent = "No replies from the team in this period.";
  }
  for (const person of result.people) {
    const tr = rows.insertRow();

    const who = tr.insertCell();
    const details = document.createElement("details");
    const summary = document.createElement("summary");
    summary.textContent = person.name;
    const address = document.createElement("div");
    address.className = "address";
    address.textContent = person.address;
    const list = document.createElement("ul");
    for (const thread of person.threads) {
      const li = document.createElement("li");
      li.textContent = thread.replies > 1 ? `${thread.subject} (${thread.replies} replies)` : thread.subject;
      list.append(li);
    }
    details.append(summary, address, list);
    who.append(details);

    const cases = tr.insertCell();
    cases.className = "num";
    cases.textContent = person.cases;
    const sent = tr.insertCell();
    sent.className = "num";
    sent.textContent = person.messages;
  }
  $("results").hidden = false;
}

function exportCsv() {
  if (!lastResult) return;
  const blob = new Blob([toCsv(lastResult)], { type: "text/csv" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `team-cases-${lastLabel.replace(/\s+/g, "-").toLowerCase()}.csv`;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

Office.onReady(() => {
  loadSettings();
  if (!$("clientId").value || !$("mailbox").value || !$("teamDomain").value) $("settings").open = true;

  $("month").value = toMonthValue(new Date());
  $("month").max = toMonthValue(new Date());
  for (const radio of document.querySelectorAll('input[name="period"]')) {
    radio.addEventListener("change", () => ($("month").disabled = selectedPeriod() !== "month"));
  }

  $("saveSettings").addEventListener("click", saveSettings);
  $("run").addEventListener("click", run);
  $("exportCsv").addEventListener("click", exportCsv);
});
