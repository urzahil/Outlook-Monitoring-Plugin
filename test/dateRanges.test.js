import { test } from "node:test";
import assert from "node:assert/strict";
import { getDateRange, toMonthValue } from "../src/core/dateRanges.js";

const now = new Date(2026, 9, 22, 15, 30); // 22 Oct 2026, local time

test("last 7 and 30 days end now", () => {
  const week = getDateRange("last7", { now });
  assert.equal(week.end, now);
  assert.equal(now - week.start, 7 * 24 * 3600 * 1000);
  assert.equal(now - getDateRange("last30", { now }).start, 30 * 24 * 3600 * 1000);
});

test("calendar month covers the first to the end of the month", () => {
  const range = getDateRange("month", { now, month: "2026-02" });
  assert.deepEqual(range.start, new Date(2026, 1, 1));
  assert.deepEqual(range.end, new Date(2026, 2, 1));
  assert.equal(range.label, "February 2026");
});

test("the current month stops at now", () => {
  assert.equal(getDateRange("month", { now, month: "2026-10" }).end, now);
});

test("a month must be chosen", () => {
  assert.throws(() => getDateRange("month", { now, month: "" }));
});

test("toMonthValue pads the month", () => {
  assert.equal(toMonthValue(new Date(2026, 0, 5)), "2026-01");
});
