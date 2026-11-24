import { getGraphToken } from "./auth.js";
import { DEFAULT_SETTINGS } from "./config.js";
import { agentConversationIds, countCases, toCsv } from "./core/counting.js";
import { getDateRange, toMonthValue } from "./core/dateRanges.js";
import { splitList, validateSettings } from "./core/validation.js";
import { fetchMessages, fetchThreadStarts } from "./graph.js";

const $ = (id) => document.getElementById(id);
const SETTING_KEYS = Object.keys(DEFAULT_SETTINGS);

let lastResult = null;
let lastLabel = "";
let abortController = null;

const isCheckbox = (key) => typeof DEFAULT_SETTINGS[key] === "boolean";
const readSetting = (key) => (isCheckbox(key) ? $(key).checked : $(key).value.trim());

// Settings roam with the user's Exchange mailbox, so they follow them across devices.
function loadSettings() {
  const roaming = Office.context.roamingSettings;
  for (const key of SETTING_KEYS) {
    const value = roaming.get(key) ?? DEFAULT_SETTINGS[key];
    if (isCheckbox(key)) $(key).checked = value === true;
    else $(key).value = value;
  }
}

function saveSettings({ quiet = false } = {}) {
  const roaming = Office.context.roamingSettings;
  for (const key of SETTING_KEYS) roaming.set(key, readSetting(key));
  roaming.saveAsync((result) => {
    const saved = result.status === Office.AsyncResultStatus.Succeeded;
    if (!quiet) setStatus(saved ? "Settings saved." : "Could not save settings.", !saved);
  });
}

function validateAndReport(settings) {
  const problem = validateSettings(settings);
  if (!problem) return true;
  $("settings").open = true;
  $(problem.field).focus();
  setStatus(problem.message, true);
  return false;
}

function setStatus(text, isError = false) {
  $("status").textContent = text;
  $("status").classList.toggle("error", isError);
}

function selectedPeriod() {
  return document.querySelector('input[name="period"]:checked').value;
}

async function run() {
  const settings = Object.fromEntries(SETTING_KEYS.map((key) => [key, readSetting(key)]));
  const { clientId, tenantId, mailbox, teamDomain } = settings;
  if (!clientId || !mailbox || !teamDomain) {
    $("settings").open = true;
    setStatus("Enter the app ID, the shared mailbox address and the team email domain in Settings.", true);
    return;
  }
  if (!validateAndReport(settings)) return;

  $("run").disabled = true;
  $("cancel").hidden = false;
  $("results").hidden = true;
  abortController = new AbortController();
  const { signal } = abortController;
  try {
    const range = getDateRange(selectedPeriod(), { month: $("month").value });
    setStatus("Signing in…");
    const token = await getGraphToken({ clientId, tenantId });

    setStatus("Reading messages…");
    const { messages, excludedFolders } = await fetchMessages({
      mailbox,
      start: range.start,
      end: range.end,
      token,
      signal,
      onProgress: (count) => setStatus(`Reading messages… ${count}`),
    });

    const options = {
      teamDomain,
      excludedAddresses: [mailbox, ...splitList(settings.excludedAddresses)],
      allowedAddresses: splitList(settings.allowedAddresses),
      onlyNewThreads: settings.onlyNewThreads,
      skipOwnThreads: settings.skipOwnThreads,
      periodStart: range.start,
    };
    if (options.onlyNewThreads || options.skipOwnThreads) {
      const conversationIds = agentConversationIds(messages, options);
      options.threadStarts = await fetchThreadStarts({
        mailbox,
        conversationIds,
        excludedFolders,
        token,
        signal,
        onProgress: (done, total) => setStatus(`Checking when threads started… ${done}/${total}`),
      });
    }

    lastResult = countCases(messages, options);
    lastLabel = range.label;
    render(lastResult, range, messages.length, options);
    setStatus("");
    // Keep what was typed, so a first run doesn't lose the settings when Outlook closes.
    saveSettings({ quiet: true });
  } catch (error) {
    if (error?.name === "AbortError") {
      setStatus("Cancelled.");
    } else {
      console.error(error);
      setStatus(error.message || String(error), true);
    }
  } finally {
    abortController = null;
    $("run").disabled = false;
    $("cancel").hidden = true;
  }
}

function render(result, range, scanned, options) {
  const dateFormat = { day: "numeric", month: "short", year: "numeric" };
  $("resultsTitle").textContent = range.label;
  $("summary").textContent =
    `${range.start.toLocaleDateString(undefined, dateFormat)} – ${range.displayEnd.toLocaleDateString(undefined, dateFormat)} · ` +
    `${result.uniqueCases} distinct cases handled · ${scanned} messages scanned`;

  const notes = [];
  if (options.onlyNewThreads) notes.push(`${result.skippedThreads.old} older ${result.skippedThreads.old === 1 ? "thread" : "threads"} skipped`);
  if (options.skipOwnThreads) notes.push(`${result.skippedThreads.own} ${result.skippedThreads.own === 1 ? "thread" : "threads"} started by the team skipped`);
  if (result.duplicates) notes.push(`${result.duplicates} duplicate ${result.duplicates === 1 ? "copy" : "copies"} ignored`);
  $("notes").hidden = notes.length === 0;
  $("notes").textContent = notes.join(" · ");

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
  const blob = new Blob(["\uFEFF", toCsv(lastResult)], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `team-cases-${lastLabel.replace(/\s+/g, "-").toLowerCase()}.csv`;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

Office.onReady(() => {
  // MSAL's sign-in popup loads this page again outside Outlook, where there is
  // no roaming settings object; there's nothing to show in that case.
  if (!Office.context?.roamingSettings) return;
  loadSettings();
  if (!$("clientId").value || !$("mailbox").value || !$("teamDomain").value) $("settings").open = true;

  $("month").value = toMonthValue(new Date());
  $("month").max = toMonthValue(new Date());
  for (const radio of document.querySelectorAll('input[name="period"]')) {
    radio.addEventListener("change", () => ($("month").disabled = selectedPeriod() !== "month"));
  }

  $("saveSettings").addEventListener("click", () => {
    const settings = Object.fromEntries(SETTING_KEYS.map((key) => [key, readSetting(key)]));
    if (validateAndReport(settings)) saveSettings();
  });
  $("cancel").addEventListener("click", () => abortController?.abort());
  $("run").addEventListener("click", run);
  $("exportCsv").addEventListener("click", exportCsv);
});
