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
            operator: "-",
          },
          {
            baseSourceType: "financial_documents",
            matchingSourceType: "import_cgd_cartao_credito",
            operator: "+",
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
        operator: "-",
      },
      {
        baseSourceType: "financial_documents",
        matchingSourceType: "import_cgd_cartao_credito",
        operator: "+",
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
            operator: "-",
          },
          {
            base_source_type: "financial_documents",
            matching_source_type: "import_cgd_cartao_credito",
            operator: "+",
          },
        ],
      },
    },
  }]);
});

test("PUT rejects changing or removing the managed Credit Card source rule before RPC", async () => {
  const invalidRules = [
    [{
      baseSourceType: "financial_documents",
      matchingSourceType: "import_cgd_cartao_credito",
      operator: "-",
    }],
    [{
      baseSourceType: "financial_documents",
      matchingSourceType: "import_cgd_extrato_ordem",
      operator: "+",
    }],
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
