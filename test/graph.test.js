import { test } from "node:test";
import assert from "node:assert/strict";
import { parseRetryAfter } from "../src/graph.js";

test("Retry-After accepts seconds, HTTP dates and garbage", () => {
  const now = Date.parse("2026-11-24T10:00:00Z");
  assert.equal(parseRetryAfter("7", 1, now), 7);
  assert.equal(parseRetryAfter("Tue, 24 Nov 2026 10:00:05 GMT", 1, now), 5);
  assert.equal(parseRetryAfter("nonsense", 3, now), 3);
  assert.equal(parseRetryAfter(null, 4, now), 4);
  assert.equal(parseRetryAfter("9999", 1, now), 60);
});
