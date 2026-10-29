const assert = require("node:assert/strict");
const { readFileSync } = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const SUPABASE_MODULE_PATH = require.resolve("../api/_supabase");
const COMMUNICATIONS_HANDLER_PATH = require.resolve("../api/communications");
const COMMUNICATIONS_AUDIT_HANDLER_PATH = path.join(__dirname, "..", "api", "communications-audit.js");
const COMMUNICATION_ID = "00000000-0000-0000-0000-000000000123";
const ACTOR_ID = "00000000-0000-0000-0000-000000000456";

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

async function withMockedHandler(handlerPath, supabase, run) {
  const previousHandler = require.cache[handlerPath];
  const previousSupabase = require.cache[SUPABASE_MODULE_PATH];
  delete require.cache[handlerPath];
  require.cache[SUPABASE_MODULE_PATH] = {
    id: SUPABASE_MODULE_PATH,
    filename: SUPABASE_MODULE_PATH,
    loaded: true,
    exports: supabase,
  };

  try {
    await run(require(handlerPath));
  } finally {
    delete require.cache[handlerPath];
    if (previousHandler) require.cache[handlerPath] = previousHandler;
    else delete require.cache[handlerPath];
    if (previousSupabase) require.cache[SUPABASE_MODULE_PATH] = previousSupabase;
    else delete require.cache[SUPABASE_MODULE_PATH];
  }
}

function mockedSupabase(overrides = {}) {
  return {
    cleanText: (value) => String(value ?? "").trim(),
    parseBody: async (request) => request.body || {},
    requireFeature: async () => ({
      user: { id: ACTOR_ID, email: "Auditor@Example.com" },
      access: {},
    }),
    restQuery: async () => [],
    sanitizeEntry: (entry) => entry,
    sendError: (response, error) => response.status(error.statusCode || 500).json({ error: error.message }),
    ...overrides,
  };
}

function extractFunction(source, name) {
  const asyncStart = source.indexOf(`async function ${name}(`);
  const start = asyncStart === -1 ? source.indexOf(`function ${name}(`) : asyncStart;
  if (start === -1) throw new Error(`Could not find ${name}.`);
  const bodyStart = source.indexOf("{", start);
  let depth = 0;
  for (let index = bodyStart; index < source.length; index += 1) {
    if (source[index] === "{") depth += 1;
    if (source[index] === "}") depth -= 1;
    if (depth === 0) return source.slice(start, index + 1);
  }
  throw new Error(`Could not parse ${name}.`);
}

function communicationsAuditUiHarness(api) {
  const source = readFileSync("app-main.js", "utf8");
  const noOpClassList = { toggle() {} };
  const context = {
    state: {
      communicationsSettingsScreen: "configuration",
      communicationsAuditRows: [],
      communicationsAuditLoaded: false,
      communicationsAuditLoading: false,
      communicationsAuditError: "",
    },
    els: {
      settingsCommunicationsConfigTab: { classList: noOpClassList },
      settingsCommunicationsAuditTab: { classList: noOpClassList },
      settingsCommunicationsConfigPanel: { hidden: false },
      settingsCommunicationsAuditPanel: { hidden: true },
      settingsCommunicationsAuditBody: { innerHTML: "" },
      settingsCommunicationsAuditStatus: { textContent: "" },
    },
    api,
    clean: (value) => String(value ?? "").trim(),
    escape: (value) => String(value ?? ""),
  };
  const script = [
    extractFunction(source, "setCommunicationsSettingsScreen"),
    extractFunction(source, "renderCommunicationsSettingsTabs"),
    extractFunction(source, "renderCommunicationsAudit"),
    extractFunction(source, "loadCommunicationsAudit"),
  ].join("\n");
  vm.runInNewContext(script, context);
  return context;
}

test("communications deletion audit migration protects and atomically snapshots deleted rows", () => {
  const sql = readFileSync("supabase-migrations/2026-10-29-communications-delete-audit.sql", "utf8");

  assert.match(sql, /create table if not exists public\.communications_audit_events/i);
  assert.match(sql, /communication_id\s+uuid\s+not null/i);
  assert.match(sql, /actor_user_id\s+uuid\s+not null/i);
  assert.match(sql, /actor_email\s+text\s+not null/i);
  assert.match(sql, /record_snapshot\s+jsonb\s+not null/i);
  assert.match(sql, /alter table public\.communications_audit_events enable row level security/i);
  assert.match(sql, /revoke all on table public\.communications_audit_events from public, anon, authenticated/i);
  assert.match(sql, /revoke delete on table public\.communications from public, anon, authenticated/i);
  assert.match(sql, /tablename = 'communications'[\s\S]*cmd = 'DELETE'[\s\S]*drop policy if exists/i);
  assert.match(sql, /create or replace function public\.delete_communication_with_audit/i);
  assert.match(sql, /security definer/i);
  assert.match(sql, /for update/i);
  assert.match(sql, /insert into public\.communications_audit_events[\s\S]*to_jsonb\(v_communication\)[\s\S]*delete from public\.communications/i);
  assert.match(sql, /revoke all on function public\.delete_communication_with_audit[\s\S]*from public, anon, authenticated/i);
  assert.match(sql, /grant execute on function public\.delete_communication_with_audit[\s\S]*to service_role/i);
});

test("communications deletion invokes the atomic audit RPC with the verified user", async () => {
  const calls = [];
  const response = responseRecorder();

  await withMockedHandler(COMMUNICATIONS_HANDLER_PATH, mockedSupabase({
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      return [{ id: COMMUNICATION_ID, person: "Guest" }];
    },
  }), async (handler) => {
    await handler({
      method: "DELETE",
      headers: {},
      query: { id: COMMUNICATION_ID },
      body: { actor_user_id: "forged-user", actor_email: "forged@example.com" },
    }, response);
  });

  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, { rows: [{ id: COMMUNICATION_ID, person: "Guest" }] });
  assert.deepEqual(calls, [{
    resource: "rpc/delete_communication_with_audit",
    options: {
      method: "POST",
      body: {
        p_communication_id: COMMUNICATION_ID,
        p_actor_user_id: ACTOR_ID,
        p_actor_email: "auditor@example.com",
      },
    },
  }]);
});

test("communications deletion rejects a missing ID before calling the audit RPC", async () => {
  let rpcCalled = false;
  const response = responseRecorder();

  await withMockedHandler(COMMUNICATIONS_HANDLER_PATH, mockedSupabase({
    restQuery: async () => {
      rpcCalled = true;
      return [];
    },
  }), async (handler) => handler({ method: "DELETE", headers: {}, query: {} }, response));

  assert.equal(response.statusCode, 400);
  assert.equal(rpcCalled, false);
});

test("communications deletion returns no rows when the atomic RPC cannot find the record", async () => {
  const response = responseRecorder();

  await withMockedHandler(COMMUNICATIONS_HANDLER_PATH, mockedSupabase({
    restQuery: async () => [],
  }), async (handler) => handler({
    method: "DELETE",
    headers: {},
    query: { id: COMMUNICATION_ID },
  }, response));

  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, { rows: [] });
});

test("communications audit reads are limited to Communications Settings users and safe fields", async () => {
  const calls = [];
  const response = responseRecorder();
  const auditRow = {
    id: "event-1",
    action: "deleted",
    communication_id: COMMUNICATION_ID,
    actor_user_id: ACTOR_ID,
    actor_email: "auditor@example.com",
    occurred_at: "2026-10-29T19:30:00.000Z",
    record_snapshot: { person: "Guest", message: "Stored" },
    private_extra: "must not leave the API",
  };

  await withMockedHandler(COMMUNICATIONS_AUDIT_HANDLER_PATH, mockedSupabase({
    requireFeature: async (request, area, feature) => {
      assert.equal(area, "settings");
      assert.equal(feature, "communications");
      return { user: { id: ACTOR_ID, email: "auditor@example.com" }, access: {} };
    },
    restQuery: async (resource, options) => {
      calls.push({ resource, options });
      return [auditRow];
    },
  }), async (handler) => handler({ method: "GET", headers: {} }, response));

  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, {
    rows: [{
      id: "event-1",
      action: "deleted",
      communicationId: COMMUNICATION_ID,
      actorUserId: ACTOR_ID,
      actorEmail: "auditor@example.com",
      occurredAt: "2026-10-29T19:30:00.000Z",
      recordSnapshot: { person: "Guest", message: "Stored" },
    }],
  });
  assert.deepEqual(calls, [{
    resource: "communications_audit_events?select=id,action,communication_id,actor_user_id,actor_email,occurred_at,record_snapshot&order=occurred_at.desc&limit=200",
    options: { method: "GET" },
  }]);
});

test("communications audit denies unauthorized and unsupported requests", async () => {
  const permissionError = Object.assign(new Error("No Communications Settings access."), { statusCode: 403 });
  const deniedResponse = responseRecorder();
  let restCalled = false;

  await withMockedHandler(COMMUNICATIONS_AUDIT_HANDLER_PATH, mockedSupabase({
    requireFeature: async () => { throw permissionError; },
    restQuery: async () => {
      restCalled = true;
      return [];
    },
  }), async (handler) => handler({ method: "GET", headers: {} }, deniedResponse));

  assert.equal(deniedResponse.statusCode, 403);
  assert.deepEqual(deniedResponse.body, { error: "No Communications Settings access." });
  assert.equal(restCalled, false);

  const methodResponse = responseRecorder();
  await withMockedHandler(COMMUNICATIONS_AUDIT_HANDLER_PATH, mockedSupabase(), async (handler) => handler({ method: "POST", headers: {} }, methodResponse));
  assert.equal(methodResponse.statusCode, 405);
  assert.deepEqual(methodResponse.body, { error: "Method not allowed." });
});

test("Communications Settings exposes a read-only, escaped audit tab", () => {
  const html = readFileSync("index.html", "utf8");
  const script = readFileSync("app-main.js", "utf8");

  assert.match(html, /id="settings-communications-config-tab"/);
  assert.match(html, /id="settings-communications-audit-tab"/);
  assert.match(html, /id="settings-communications-config-panel"/);
  assert.match(html, /id="settings-communications-audit-panel"/);
  assert.match(html, /id="settings-communications-audit-body"/);
  assert.match(script, /function setCommunicationsSettingsScreen\(/);
  assert.match(script, /function loadCommunicationsAudit\(/);
  assert.match(script, /api\("\/api\/communications-audit"\)/);
  assert.match(script, /function renderCommunicationsAudit\([\s\S]*escape\(clean\(snapshot\.message\) \|\| "-"\)/);
  assert.doesNotMatch(html, /settings-communications-audit[^\n]*data-[^\n]*(delete|restore|edit)/i);
});

test("reopening Communications Audit refreshes the deletion history", async () => {
  let requests = 0;
  const context = communicationsAuditUiHarness(async () => {
    requests += 1;
    return { rows: [] };
  });

  context.setCommunicationsSettingsScreen("audit");
  await new Promise((resolve) => setImmediate(resolve));
  context.setCommunicationsSettingsScreen("configuration");
  context.setCommunicationsSettingsScreen("audit");
  await new Promise((resolve) => setImmediate(resolve));

  assert.equal(requests, 2);
});

test("re-entering Communications Settings refreshes an already-selected Audit tab", async () => {
  const source = readFileSync("app-main.js", "utf8");
  let communicationsLoads = 0;
  let auditLoads = 0;
  const context = {
    state: { settingsSection: "communications", communicationsSettingsScreen: "audit" },
    ensureCommunicationsData: async () => { communicationsLoads += 1; },
    loadCommunicationsAudit: async () => { auditLoads += 1; },
  };
  vm.runInNewContext(extractFunction(source, "ensureSettingsSectionData"), context);

  await context.ensureSettingsSectionData();

  assert.equal(communicationsLoads, 1);
  assert.equal(auditLoads, 1);
});
