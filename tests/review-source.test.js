const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const { normalizeSource } = require("../api/_supabase");
const appMain = fs.readFileSync(path.join(__dirname, "..", "app-main.js"), "utf8");

function appFunctionSource(name) {
  const functionStart = appMain.indexOf(`function ${name}(`);
  assert.notEqual(functionStart, -1, `missing function ${name}`);
  const bodyStart = appMain.indexOf("{", appMain.indexOf(")", functionStart));
  let depth = 0;
  for (let index = bodyStart; index < appMain.length; index += 1) {
    if (appMain[index] === "{") depth += 1;
    if (appMain[index] === "}") depth -= 1;
    if (depth === 0) return appMain.slice(functionStart, index + 1);
  }
  throw new Error(`could not extract function ${name}`);
}

const normalizeReviewSourceKey = new Function(
  `${appFunctionSource("clean")}\n${appFunctionSource("normalizeReviewSourceKey")}\nreturn normalizeReviewSourceKey;`,
)();

test("review source normalization maps Travelocity to Expedia", () => {
  assert.equal(normalizeSource("Travelocity"), "expedia");
});

test("review import parsing maps Travelocity to Expedia", () => {
  assert.equal(normalizeReviewSourceKey("travelocity"), "expedia");
});

test("Travelocity migration converts stored review sources to Expedia", () => {
  const migrationPath = path.join(__dirname, "..", "supabase-migrations", "2026-09-11-reviews-travelocity-to-expedia.sql");
  assert.equal(fs.existsSync(migrationPath), true);
  const migration = fs.readFileSync(migrationPath, "utf8");
  assert.match(migration, /update public\.reviews\s+set source = 'expedia'/i);
  assert.match(migration, /lower\(btrim\(source\)\) = 'travelocity'/i);
});
