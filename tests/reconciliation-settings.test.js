const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const settingsPath = require.resolve("../api/reconciliation-settings");
const supabasePath = require.resolve("../api/_supabase");
const migration = fs.readFileSync(path.join(root, "supabase-migrations", "2026-08-11-financial-reconciliation-source-rules.sql"), "utf8");
const workspaceFilterFixPath = path.join(root, "supabase-migrations", "2026-08-11-financial-reconciliation-source-rules-workspace-filter-fix.sql");
const actionOverloadFixPath = path.join(root, "supabase-migrations", "2026-08-12-financial-reconciliation-action-overload-fix.sql");
const safeDeleteFixPath = path.join(root, "supabase-migrations", "2026-08-21-financial-reconciliation-source-rules-safe-delete.sql");

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

async function withSettingsHandler(supabase, run) {
  const previousSettings = require.cache[settingsPath];
  const previousSupabase = require.cache[supabasePath];
  delete require.cache[settingsPath];
  require.cache[supabasePath] = {
    id: supabasePath,
    filename: supabasePath,
    loaded: true,
    exports: supabase,
  };

  try {
    await run(require(settingsPath));
  } finally {
    delete require.cache[settingsPath];
    if (previousSettings) require.cache[settingsPath] = previousSettings;
    if (previousSupabase) require.cache[supabasePath] = previousSupabase;
    else delete require.cache[supabasePath];
  }
}

test("PUT validates rules then calls one replacement RPC without direct table mutation", async () => {
  const calls = [];
  const response = responseRecorder();

  await withSettingsHandler({
    parseBody: async (request) => request.body,
    requireFeature: async () => ({}),
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      return null;
    },
    sendError: (res, error) => res.status(error.statusCode || 500).json({ error: error.message }),
  }, async (handler) => {
    await handler({
      method: "PUT",
      body: {
        rules: [
          {
            baseSourceType: "financial_documents",
            matchingSourceType: "import_cgd_extrato_ordem",
            operator: "+",
          },
          {
            baseSourceType: "financial_documents",
            matchingSourceType: "import_cgd_cartao_credito",
            operator: "+",
          },
          {
            baseSourceType: "import_cgd_extrato_ordem",
            matchingSourceType: "import_fdm_accounts",
            operator: "-",
          },
          {
            baseSourceType: "import_fdm_accounts",
            matchingSourceType: "import_cgd_extrato_ordem",
            operator: "-",
          },
        ],
      },
    }, response);
  });

  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, {
    rules: [
      {
        baseSourceType: "financial_documents",
        matchingSourceType: "import_cgd_extrato_ordem",
        operator: "+",
      },
      {
        baseSourceType: "financial_documents",
        matchingSourceType: "import_cgd_cartao_credito",
        operator: "+",
      },
      {
        baseSourceType: "import_cgd_extrato_ordem",
        matchingSourceType: "import_fdm_accounts",
        operator: "-",
      },
      {
        baseSourceType: "import_fdm_accounts",
        matchingSourceType: "import_cgd_extrato_ordem",
        operator: "-",
      },
    ],
  });
  assert.deepEqual(calls, [{
    resource: "rpc/replace_financial_reconciliation_source_rules",
    options: {
      method: "POST",
      body: {
        p_rules: [
          {
            base_source_type: "financial_documents",
            matching_source_type: "import_cgd_extrato_ordem",
            operator: "+",
          },
          {
            base_source_type: "financial_documents",
            matching_source_type: "import_cgd_cartao_credito",
            operator: "+",
          },
          {
            base_source_type: "import_cgd_extrato_ordem",
            matching_source_type: "import_fdm_accounts",
            operator: "-",
          },
          {
            base_source_type: "import_fdm_accounts",
            matching_source_type: "import_cgd_extrato_ordem",
            operator: "-",
          },
        ],
      },
    },
  }]);
});

test("PUT rejects changing or removing the managed Credit Card source rule before RPC", async () => {
  const invalidRules = [
    [
      {
        baseSourceType: "financial_documents",
        matchingSourceType: "import_cgd_extrato_ordem",
        operator: "+",
      },
      {
        baseSourceType: "financial_documents",
        matchingSourceType: "import_cgd_cartao_credito",
        operator: "-",
      },
      {
        baseSourceType: "import_cgd_extrato_ordem",
        matchingSourceType: "import_fdm_accounts",
        operator: "-",
      },
    ],
    [
      {
        baseSourceType: "financial_documents",
        matchingSourceType: "import_cgd_extrato_ordem",
        operator: "+",
      },
      {
        baseSourceType: "import_cgd_extrato_ordem",
        matchingSourceType: "import_fdm_accounts",
        operator: "-",
      },
    ],
  ];

  for (const rules of invalidRules) {
    const calls = [];
    const response = responseRecorder();
    await withSettingsHandler({
      parseBody: async (request) => request.body,
      requireFeature: async () => ({}),
      restQuery: async (resource, options) => {
        calls.push({ resource, options });
        return null;
      },
      sendError: (res, error) => res.status(error.statusCode || 500).json({ error: error.message }),
    }, async (handler) => {
      await handler({ method: "PUT", body: { rules } }, response);
    });

    assert.equal(response.statusCode, 400);
    assert.deepEqual(response.body, {
      error: "The managed Credit Card source rule must remain enabled with operator +.",
    });
    assert.deepEqual(calls, []);
  }
});

test("PUT rejects changing or removing the managed Bank Statement source rule before RPC", async () => {
  const invalidRules = [
    [
      {
        baseSourceType: "financial_documents",
        matchingSourceType: "import_cgd_extrato_ordem",
        operator: "-",
      },
      {
        baseSourceType: "financial_documents",
        matchingSourceType: "import_cgd_cartao_credito",
        operator: "+",
      },
      {
        baseSourceType: "import_cgd_extrato_ordem",
        matchingSourceType: "import_fdm_accounts",
        operator: "-",
      },
    ],
    [
      {
        baseSourceType: "financial_documents",
        matchingSourceType: "import_cgd_cartao_credito",
        operator: "+",
      },
      {
        baseSourceType: "import_cgd_extrato_ordem",
        matchingSourceType: "import_fdm_accounts",
        operator: "-",
      },
    ],
  ];

  for (const rules of invalidRules) {
    const calls = [];
    const response = responseRecorder();
    await withSettingsHandler({
      parseBody: async (request) => request.body,
      requireFeature: async () => ({}),
      restQuery: async (resource, options) => {
        calls.push({ resource, options });
        return null;
      },
      sendError: (res, error) => res.status(error.statusCode || 500).json({ error: error.message }),
    }, async (handler) => {
      await handler({ method: "PUT", body: { rules } }, response);
    });

    assert.equal(response.statusCode, 400);
    assert.deepEqual(response.body, {
      error: "The managed Bank Statement source rule must remain enabled with operator +.",
    });
    assert.deepEqual(calls, []);
  }
});

test("PUT rejects changing or removing the managed POS income source rule before RPC", async () => {
  const invalidRules = [
    [
      {
        baseSourceType: "financial_documents",
        matchingSourceType: "import_cgd_extrato_ordem",
        operator: "+",
      },
      {
        baseSourceType: "financial_documents",
        matchingSourceType: "import_cgd_cartao_credito",
        operator: "+",
      },
      {
        baseSourceType: "import_cgd_extrato_ordem",
        matchingSourceType: "import_fdm_accounts",
        operator: "+",
      },
    ],
    [
      {
        baseSourceType: "financial_documents",
        matchingSourceType: "import_cgd_extrato_ordem",
        operator: "+",
      },
      {
        baseSourceType: "financial_documents",
        matchingSourceType: "import_cgd_cartao_credito",
        operator: "+",
      },
    ],
  ];

  for (const rules of invalidRules) {
    const rpcCalls = [];
    const response = responseRecorder();
    await withSettingsHandler({
      parseBody: async (request) => request.body,
      requireFeature: async () => ({}),
      restQuery: async (resource, options) => {
        rpcCalls.push({ resource, options });
        return null;
      },
      sendError: (res, error) => res.status(error.statusCode || 500).json({ error: error.message }),
    }, async (handler) => {
      await handler({ method: "PUT", body: { rules } }, response);
    });

    assert.equal(response.statusCode, 400);
    assert.match(response.body.error, /managed POS income source rule must remain enabled with operator -/i);
    assert.equal(rpcCalls.length, 0);
  }
});

test("PUT rejects changing or removing the managed Bank Reservation source rule before RPC", async () => {
  const requiredRules = [
    { baseSourceType: "financial_documents", matchingSourceType: "import_cgd_extrato_ordem", operator: "+" },
    { baseSourceType: "financial_documents", matchingSourceType: "import_cgd_cartao_credito", operator: "+" },
    { baseSourceType: "import_cgd_extrato_ordem", matchingSourceType: "import_fdm_accounts", operator: "-" },
    { baseSourceType: "import_fdm_accounts", matchingSourceType: "import_cgd_extrato_ordem", operator: "-" },
  ];
  const invalidRules = [
    requiredRules.map((rule) => rule.baseSourceType === "import_fdm_accounts"
      && rule.matchingSourceType === "import_cgd_extrato_ordem"
      ? { ...rule, operator: "+" }
      : rule),
    requiredRules.filter((rule) => rule.baseSourceType !== "import_fdm_accounts"
      || rule.matchingSourceType !== "import_cgd_extrato_ordem"),
  ];

  for (const rules of invalidRules) {
    const rpcCalls = [];
    const response = responseRecorder();
    await withSettingsHandler({
      parseBody: async (request) => request.body,
      requireFeature: async () => ({}),
      restQuery: async (resource, options) => {
        rpcCalls.push({ resource, options });
        return null;
      },
      sendError: (res, error) => res.status(error.statusCode || 500).json({ error: error.message }),
    }, async (handler) => {
      await handler({ method: "PUT", body: { rules } }, response);
    });

    assert.equal(response.statusCode, 400);
    assert.deepEqual(response.body, {
      error: "The managed Bank Reservation source rule must remain enabled with operator -.",
    });
    assert.deepEqual(rpcCalls, []);
  }
});

test("PUT preserves the managed Adyen source rule while allowing unrelated source-rule edits", async () => {
  const requiredRules = [
    { baseSourceType: "financial_documents", matchingSourceType: "import_cgd_extrato_ordem", operator: "+" },
    { baseSourceType: "financial_documents", matchingSourceType: "import_cgd_cartao_credito", operator: "+" },
    { baseSourceType: "import_cgd_extrato_ordem", matchingSourceType: "import_fdm_accounts", operator: "-" },
    { baseSourceType: "import_fdm_accounts", matchingSourceType: "import_cgd_extrato_ordem", operator: "-" },
  ];
  const invalidRules = [
    requiredRules.map((rule) => rule.baseSourceType === "import_cgd_extrato_ordem"
      && rule.matchingSourceType === "import_fdm_accounts"
      ? { ...rule, operator: "+" }
      : rule),
    requiredRules.filter((rule) => rule.baseSourceType !== "import_cgd_extrato_ordem"
      || rule.matchingSourceType !== "import_fdm_accounts"),
  ];

  for (const rules of invalidRules) {
    const rpcCalls = [];
    const response = responseRecorder();
    await withSettingsHandler({
      parseBody: async (request) => request.body,
      requireFeature: async () => ({}),
      restQuery: async (resource, options) => {
        rpcCalls.push({ resource, options });
        return null;
      },
      sendError: (res, error) => res.status(error.statusCode || 500).json({ error: error.message }),
    }, async (handler) => {
      await handler({ method: "PUT", body: { rules } }, response);
    });

    assert.equal(response.statusCode, 400);
    assert.match(response.body.error, /managed POS income source rule must remain enabled with operator -/i);
    assert.deepEqual(rpcCalls, []);
  }

  const response = responseRecorder();
  const rpcCalls = [];
  const rules = [...requiredRules, {
    baseSourceType: "import_cgd_cartao_credito",
    matchingSourceType: "financial_documents",
    operator: "-",
  }];
  await withSettingsHandler({
    parseBody: async (request) => request.body,
    requireFeature: async () => ({}),
    restQuery: async (resource, options) => {
      rpcCalls.push({ resource, options });
      return null;
    },
    sendError: (res, error) => res.status(error.statusCode || 500).json({ error: error.message }),
  }, async (handler) => {
    await handler({ method: "PUT", body: { rules } }, response);
  });

  assert.equal(response.statusCode, 200);
  assert.equal(rpcCalls.length, 1);
  assert.equal(rpcCalls[0].resource, "rpc/replace_financial_reconciliation_source_rules");
});

test("migration restricts source rules to the service role and defines an atomic validating RPC", () => {
  assert.match(migration, /alter table public\.financial_reconciliation_source_rules enable row level security;/);
  assert.match(migration, /revoke all on table public\.financial_reconciliation_source_rules from public, anon, authenticated;/);
  assert.match(migration, /grant select, insert, update, delete on table public\.financial_reconciliation_source_rules to service_role;/);
  assert.match(migration, /create or replace function public\.replace_financial_reconciliation_source_rules\(p_rules jsonb\)/);
  assert.match(migration, /returns jsonb language plpgsql security definer/);
  assert.match(migration, /if jsonb_typeof\(p_rules\) <> 'array' then/);
  assert.match(migration, /Rule source type is invalid\./);
  assert.match(migration, /Rule sources must be different\./);
  assert.match(migration, /raise exception 'Rule operator must be ''\+'' or ''-''\.';/);
  assert.match(migration, /Duplicate reconciliation rule\./);
  assert.match(migration, /revoke all on function public\.replace_financial_reconciliation_source_rules\(jsonb\) from public, anon, authenticated;/);
  assert.match(migration, /grant execute on function public\.replace_financial_reconciliation_source_rules\(jsonb\) to service_role;/);
});

test("migration preserves the legacy difference-function parameter name when replacing it", () => {
  assert.match(
    migration,
    /create or replace function public\.financial_reconciliation_difference\(p_base text, p_matching jsonb, p_reconciliation_id uuid\)/,
  );
  assert.doesNotMatch(
    migration,
    /create or replace function public\.financial_reconciliation_difference\(p_base text, p_rules jsonb, p_reconciliation_id uuid\)/,
  );
});

test("source-rules workspace filters extract text before building ilike patterns", () => {
  assert.match(migration, /s\.description ilike '%' \|\| \(p_filters->>'description'\) \|\| '%'/);
  assert.match(migration, /s\.supplier ilike '%' \|\| \(p_filters->>'supplier'\) \|\| '%'/);
  assert.ok(fs.existsSync(workspaceFilterFixPath), "a post-deployment workspace filter fix migration should exist");
  const workspaceFilterFix = fs.readFileSync(workspaceFilterFixPath, "utf8");
  assert.match(workspaceFilterFix, /get_financial_reconciliation_workspace\(uuid,text,jsonb,integer,integer\)/);
  assert.match(workspaceFilterFix, /s\.description ilike '%' \|\| \(p_filters->>'description'\) \|\| '%'/);
  assert.match(workspaceFilterFix, /s\.supplier ilike '%' \|\| \(p_filters->>'supplier'\) \|\| '%'/);
});

test("source-rules migration removes the obsolete action overload", () => {
  const oldSignature = "financial_reconciliation_action(text,text,uuid,text,text[],text,uuid,text)";
  assert.match(migration, new RegExp(`drop function if exists public\\.${oldSignature.replace(/[()[\].+*?^$\\|]/g, "\\$&")};`));
  assert.ok(fs.existsSync(actionOverloadFixPath), "a post-deployment action-overload fix migration should exist");
  const actionOverloadFix = fs.readFileSync(actionOverloadFixPath, "utf8");
  assert.match(actionOverloadFix, new RegExp(`drop function if exists public\\.${oldSignature.replace(/[()[\].+*?^$\\|]/g, "\\$&")};`));
  assert.match(actionOverloadFix, /notify pgrst, 'reload schema';/);
});

test("source-rule replacement migration uses a scoped delete accepted by database safety guards", () => {
  assert.ok(fs.existsSync(safeDeleteFixPath), "a forward migration should repair the deployed replacement RPC");
  const safeDeleteFix = fs.readFileSync(safeDeleteFixPath, "utf8");
  assert.match(
    safeDeleteFix,
    /delete from public\.financial_reconciliation_source_rules\s+where base_source_type in \(/,
  );
  assert.doesNotMatch(
    safeDeleteFix,
    /delete from public\.financial_reconciliation_source_rules\s*;/,
  );
  assert.match(
    safeDeleteFix,
    /grant execute on function public\.replace_financial_reconciliation_source_rules\(jsonb\) to service_role;/,
  );
});
