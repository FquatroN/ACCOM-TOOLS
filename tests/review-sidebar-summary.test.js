const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const handlerPath = require.resolve("../api/review-sidebar-summary");
const supabasePath = require.resolve("../api/_supabase");
const appMain = fs.readFileSync(path.join(root, "app-main.js"), "utf8");

function responseRecorder() {
  return {
    statusCode: 0,
    body: null,
    status(code) {
      this.statusCode = code;
      return this;
    },
    json(body) {
      this.body = body;
      return this;
    },
  };
}

async function withHandler(restQuery, run) {
  const previousHandler = require.cache[handlerPath];
  const previousSupabase = require.cache[supabasePath];
  delete require.cache[handlerPath];
  require.cache[supabasePath] = {
    id: supabasePath,
    filename: supabasePath,
    loaded: true,
    exports: {
      cleanText: (value) => String(value ?? "").trim(),
      requireFeature: async () => ({ user: { id: "user-1" } }),
      restQuery,
      sendError(res, error) {
        res.status(Number(error?.statusCode || 500)).json({ error: error?.message || "Unexpected server error." });
      },
    },
  };
  try {
    await run(require(handlerPath));
  } finally {
    delete require.cache[handlerPath];
    if (previousHandler) require.cache[handlerPath] = previousHandler;
    if (previousSupabase) require.cache[supabasePath] = previousSupabase;
    else delete require.cache[supabasePath];
  }
}

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

test("review snapshot returns source averages for all, Hostel, and Cruz", async () => {
  const currentMonth = new Date().toISOString().slice(0, 7);
  const response = responseRecorder();
  await withHandler(async () => [
    { review_date: `${currentMonth}-10`, rating_normalized_100: 90, source: "agoda", properties: { name: "Lisboa Central Hostel" } },
    { review_date: `${currentMonth}-11`, rating_normalized_100: 70, source: "agoda", properties: { name: "Cruz Apartments" } },
    { review_date: `${currentMonth}-12`, rating_normalized_100: 80, source: "booking", properties: { name: "Cruz Apartments" } },
  ], async (handler) => handler({ method: "GET" }, response));

  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body.summary.sourceSummaries.all, [
    { source: "agoda", last2MonthsAverage: 80, last6MonthsAverage: 80, last2MonthsCount: 2, last6MonthsCount: 2 },
    { source: "booking", last2MonthsAverage: 80, last6MonthsAverage: 80, last2MonthsCount: 1, last6MonthsCount: 1 },
  ]);
  assert.deepEqual(response.body.summary.sourceSummaries.hostel, [
    { source: "agoda", last2MonthsAverage: 90, last6MonthsAverage: 90, last2MonthsCount: 1, last6MonthsCount: 1 },
  ]);
  assert.deepEqual(response.body.summary.sourceSummaries.cruz, [
    { source: "agoda", last2MonthsAverage: 70, last6MonthsAverage: 70, last2MonthsCount: 1, last6MonthsCount: 1 },
    { source: "booking", last2MonthsAverage: 80, last6MonthsAverage: 80, last2MonthsCount: 1, last6MonthsCount: 1 },
  ]);
});

test("Review Snapshot source selector defaults to all and returns the requested property scope", () => {
  const sourceSummaryForProperty = new Function(`${appFunctionSource("clean")}\n${appFunctionSource("sourceSummaryForProperty")}\nreturn sourceSummaryForProperty;`)();
  const summary = {
    sources: [{ source: "agoda" }],
    sourceSummaries: {
      all: [{ source: "agoda" }],
      hostel: [{ source: "booking" }],
      cruz: [{ source: "expedia" }],
    },
  };

  assert.deepEqual(sourceSummaryForProperty(summary), [{ source: "agoda" }]);
  assert.deepEqual(sourceSummaryForProperty(summary, "hostel"), [{ source: "booking" }]);
  assert.deepEqual(sourceSummaryForProperty(summary, "cruz"), [{ source: "expedia" }]);
  assert.deepEqual(sourceSummaryForProperty(summary, "unknown"), [{ source: "agoda" }]);
});
