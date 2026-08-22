const test = require("node:test");
const assert = require("node:assert/strict");

const handlerPath = require.resolve("../api/reconciliation-history");
const supabasePath = require.resolve("../api/_supabase");

function responseRecorder() {
  return {
    statusCode: 0,
    body: null,
    headers: {},
    status(code) { this.statusCode = code; return this; },
    json(body) { this.body = body; return this; },
    setHeader(name, value) { this.headers[name] = value; },
  };
}

async function withHandler(supabase, run) {
  const previousHandler = require.cache[handlerPath];
  const previousSupabase = require.cache[supabasePath];
  delete require.cache[handlerPath];
  require.cache[supabasePath] = {
    id: supabasePath,
    filename: supabasePath,
    loaded: true,
    exports: supabase,
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

test("history GET authorizes app access and calls the paginated RPC", async () => {
  const calls = [];
  const response = responseRecorder();
  await withHandler({
    requireFeature: async (_req, area, feature) => calls.push({ area, feature }),
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      return { rows: [], page: 2, pageSize: 25, total: 0 };
    },
    sendError: (res, error) => res.status(error.statusCode || 500).json({ error: error.message }),
  }, async (handler) => {
    await handler({
      method: "GET",
      query: {
        created_from: "2026-01-01",
        created_to: "2026-08-22",
        origin: "automatic",
        status: "complete",
        difference_from: "-10.50",
        difference_to: "5",
        page: "2",
        page_size: "25",
      },
    }, response);
  });

  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, { rows: [], page: 2, pageSize: 25, total: 0 });
  assert.deepEqual(calls, [
    { area: "app", feature: "financial-reconciliation" },
    {
      resource: "rpc/get_financial_reconciliation_history",
      options: {
        method: "POST",
        body: {
          p_created_from: "2026-01-01",
          p_created_to: "2026-08-22",
          p_origin: "automatic",
          p_status: "complete",
          p_difference_from: -10.5,
          p_difference_to: 5,
          p_page: 2,
          p_page_size: 25,
        },
      },
    },
  ]);
});

test("history endpoint rejects unsupported methods", async () => {
  const response = responseRecorder();
  await withHandler({
    requireFeature: async () => ({}),
    restQuery: async () => { throw new Error("must not be called"); },
    sendError: (res, error) => res.status(error.statusCode || 500).json({ error: error.message }),
  }, async (handler) => handler({ method: "POST", query: {} }, response));
  assert.equal(response.statusCode, 405);
  assert.equal(response.headers.Allow, "GET");
});
