const DAY_MS = 24 * 60 * 60 * 1000;

/**
 * Returns the reporting window as { start, end, displayEnd, label } with `end`
 * exclusive. `displayEnd` is the last moment inside the window, for showing to
 * people (a full February ends on 28 Feb, not 1 Mar).
 *
 * @param {"last7" | "last30" | "month"} period
 * @param {{ now?: Date, month?: string }} options `month` is "YYYY-MM" and is
 *   interpreted in the user's local time zone.
 */
export function getDateRange(period, { now = new Date(), month } = {}) {
  switch (period) {
    case "last7":
      return { start: new Date(now.getTime() - 7 * DAY_MS), end: now, displayEnd: now, label: "Last 7 days" };
    case "last30":
      return { start: new Date(now.getTime() - 30 * DAY_MS), end: now, displayEnd: now, label: "Last 30 days" };
    case "month": {
      const match = /^(\d{4})-(\d{2})$/.exec(month || "");
      if (!match) throw new Error("Pick a calendar month.");
      const year = Number(match[1]);
      const monthIndex = Number(match[2]) - 1;
      const start = new Date(year, monthIndex, 1);
      const end = new Date(year, monthIndex + 1, 1);
      const label = start.toLocaleString("en-GB", { month: "long", year: "numeric" });
      const complete = end <= now;
      return { start, end: complete ? end : now, displayEnd: complete ? new Date(end.getTime() - 1) : now, label };
    }
    default:
      throw new Error(`Unknown period: ${period}`);
  }
}

/** "YYYY-MM" for the month containing `date` (local time). */
export function toMonthValue(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}
