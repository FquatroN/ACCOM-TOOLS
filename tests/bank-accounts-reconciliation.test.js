const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const appMain = fs.readFileSync(path.join(root, "app-main.js"), "utf8");
const bankAccountsHandlerPath = require.resolve("../api/bank-accounts");
const supabaseModulePath = require.resolve("../api/_supabase");
const financialDocsModulePath = require.resolve("../api/_financial-docs");
const financialDocsServicePath = require.resolve("../api/_financial-docs-service");

function appFunctionSource(name) {
  const functionStart = appMain.indexOf(`function ${name}(`);
  assert.notEqual(functionStart, -1, `${name} should be defined in app-main.js`);
  const start = appMain.slice(Math.max(0, functionStart - 6), functionStart) === "async " ? functionStart - 6 : functionStart;
  const bodyStart = appMain.indexOf("{", appMain.indexOf(")", functionStart));
  let depth = 0;
  for (let index = bodyStart; index < appMain.length; index += 1) {
    if (appMain[index] === "{") depth += 1;
    if (appMain[index] === "}") depth -= 1;
    if (depth === 0) return appMain.slice(start, index + 1);
  }
  throw new Error(`Could not extract ${name} from app-main.js`);
}

function responseRecorder() {
  return {
    statusCode: 0,
    body: null,
    headers: {},
    status(code) {
      this.statusCode = code;
      return this;
    },
    json(body) {
      this.body = body;
      return this;
    },
    setHeader(name, value) {
      this.headers[name] = value;
      return this;
    },
  };
}

async function withBankAccountsHandler(restQuery, run) {
  const previousHandler = require.cache[bankAccountsHandlerPath];
  const previousSupabase = require.cache[supabaseModulePath];
  delete require.cache[bankAccountsHandlerPath];
  require.cache[supabaseModulePath] = {
    id: supabaseModulePath,
    filename: supabaseModulePath,
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
    await run(require(bankAccountsHandlerPath));
  } finally {
    delete require.cache[bankAccountsHandlerPath];
    if (previousHandler) require.cache[bankAccountsHandlerPath] = previousHandler;
    if (previousSupabase) require.cache[supabaseModulePath] = previousSupabase;
    else delete require.cache[supabaseModulePath];
  }
}

async function withFinancialDocsService(restQuery, run) {
  const cached = new Map([
    [supabaseModulePath, require.cache[supabaseModulePath]],
    [financialDocsModulePath, require.cache[financialDocsModulePath]],
    [financialDocsServicePath, require.cache[financialDocsServicePath]],
  ]);
  delete require.cache[financialDocsModulePath];
  delete require.cache[financialDocsServicePath];
  require.cache[supabaseModulePath] = {
    id: supabaseModulePath,
    filename: supabaseModulePath,
    loaded: true,
    exports: {
      cleanText: (value) => String(value ?? "").trim(),
      normalizeDate: (value) => String(value ?? "").trim(),
      normalizeNumeric: (value) => Number(value),
      restQuery,
    },
  };
  try {
    await run(require(financialDocsServicePath));
  } finally {
    for (const modulePath of [financialDocsServicePath, financialDocsModulePath, supabaseModulePath]) {
      delete require.cache[modulePath];
      if (cached.get(modulePath)) require.cache[modulePath] = cached.get(modulePath);
    }
  }
}

test("Bank Statement rows receive reconciliation metadata in bulk", async () => {
  const calls = [];
  const recordOne = "00000000-0000-0000-0000-000000000101";
  const recordTwo = "00000000-0000-0000-0000-000000000102";
  const reconciliationId = "00000000-0000-0000-0000-000000000201";
  await withBankAccountsHandler(async (query) => {
    calls.push(query);
    if (query.startsWith("import_cgd_extrato_ordem?")) {
      return [
        { id: recordOne, data: "2026-08-01", descritivo: "First", montante: -25 },
        { id: recordTwo, data: "2026-08-02", descritivo: "Second", montante: -30 },
      ];
    }
    if (query.startsWith("financial_reconciliation_items?")) {
      return [{ source_id: recordOne, reconciliation_id: reconciliationId }];
    }
    if (query.startsWith("financial_reconciliations?")) {
      return [{ id: reconciliationId, status: "complete", difference_amount: "0.00" }];
    }
    throw new Error(`Unexpected query: ${query}`);
  }, async (handler) => {
    const res = responseRecorder();
    await handler({ method: "GET", query: { source: "cgd-extrato" } }, res);
    assert.equal(res.statusCode, 200);
    assert.deepEqual(res.body.rows, [
      {
        id: recordOne,
        data: "2026-08-01",
        descritivo: "First",
        montante: -25,
        reconciliationId,
        reconciliationStatus: "complete",
        reconciliationDifferenceAmount: "0.00",
      },
      { id: recordTwo, data: "2026-08-02", descritivo: "Second", montante: -30 },
    ]);
  });

  assert.equal(calls.length, 3, "two source rows should use one item lookup and one reconciliation lookup");
  assert.match(calls[1], /source_type=eq\.import_cgd_extrato_ordem/);
  assert.match(calls[1], new RegExp(recordOne));
  assert.match(calls[1], new RegExp(recordTwo));
  assert.match(calls[2], /select=id,status,difference_amount/);
});

test("Credit Card rows use their source type and retain a non-zero difference", async () => {
  const calls = [];
  const recordId = "00000000-0000-0000-0000-000000000111";
  const reconciliationId = "00000000-0000-0000-0000-000000000211";
  await withBankAccountsHandler(async (query) => {
    calls.push(query);
    if (query.startsWith("import_cgd_cartao_credito?")) {
      return [{ id: recordId, data: "2026-08-03", descricao: "Card", valor: -40 }];
    }
    if (query.startsWith("financial_reconciliation_items?")) {
      return [{ source_id: recordId, reconciliation_id: reconciliationId }];
    }
    if (query.startsWith("financial_reconciliations?")) {
      return [{ id: reconciliationId, status: "complete", difference_amount: "5.25" }];
    }
    throw new Error(`Unexpected query: ${query}`);
  }, async (handler) => {
    const res = responseRecorder();
    await handler({ method: "GET", query: { source: "cartao-credito" } }, res);
    assert.equal(res.statusCode, 200);
    assert.equal(res.body.rows[0].reconciliationDifferenceAmount, "5.25");
  });

  assert.match(calls[1], /source_type=eq\.import_cgd_cartao_credito/);
});

test("Financial Documents list rows expose the authoritative reconciliation difference", async () => {
  const documentId = "00000000-0000-0000-0000-000000000121";
  const reconciliationId = "00000000-0000-0000-0000-000000000221";
  await withFinancialDocsService(async (query) => {
    if (query.startsWith("financial_documents?")) {
      return [{ id: documentId, document_date: "2026-08-01", amount: 20 }];
    }
    if (query.startsWith("financial_document_history?")) return [];
    if (query.startsWith("financial_reconciliation_items?")) {
      return [{ source_id: documentId, reconciliation_id: reconciliationId }];
    }
    if (query.startsWith("financial_reconciliations?")) {
      assert.match(query, /select=id,status,difference_amount/);
      return [{ id: reconciliationId, status: "complete", difference_amount: "2.50" }];
    }
    throw new Error(`Unexpected query: ${query}`);
  }, async ({ listFinancialDocuments }) => {
    const rows = await listFinancialDocuments({});
    assert.equal(rows[0].reconciliationId, reconciliationId);
    assert.equal(rows[0].reconciliationStatus, "complete");
    assert.equal(rows[0].reconciliationDifferenceAmount, 2.5);
  });
});

test("reconciliation list buttons distinguish balanced, different, and unreconciled rows", () => {
  const button = new Function(
    "clean", "escape",
    `${appFunctionSource("reconciliationListButton")}\nreturn reconciliationListButton;`,
  )(
    (value) => String(value ?? "").trim(),
    (value) => String(value).replaceAll("&", "&amp;").replaceAll('"', "&quot;"),
  );

  const balanced = button({
    id: "record-1",
    reconciliationId: "reconciliation-1",
    reconciliationStatus: "complete",
    reconciliationDifferenceAmount: "0.00",
  });
  assert.match(balanced, /reconciliation-list-button--balanced/);
  assert.doesNotMatch(balanced, /class="[^"]*\bghost\b/);
  assert.match(balanced, /data-reconciliation-id="reconciliation-1"/);
  assert.match(balanced, /Open balanced completed reconciliation/);

  const different = button({
    id: "record-2",
    reconciliationId: "reconciliation-2",
    reconciliationStatus: "started",
    reconciliationDifferenceAmount: "-0.01",
  });
  assert.match(different, /reconciliation-list-button--difference/);
  assert.doesNotMatch(different, /class="[^"]*\bghost\b/);
  assert.match(different, /Open reconciliation with a non-zero difference/);

  assert.equal(button({ id: "record-3" }), "");
});

test("Financial Documents use the same balanced and difference button colors", () => {
  const button = new Function(
    "clean", "escape",
    `${appFunctionSource("reconciliationListButton")}\nreturn reconciliationListButton;`,
  )(
    (value) => String(value ?? "").trim(),
    (value) => String(value),
  );
  const financialButton = new Function(
    "reconciliationListButton",
    `${appFunctionSource("financialDocReconciliationButton")}\nreturn financialDocReconciliationButton;`,
  )(button);

  assert.match(financialButton({
    id: "document-1",
    reconciliationId: "reconciliation-1",
    reconciliationDifferenceAmount: 0,
  }), /financial-doc-reconciliation-button[^\"]*reconciliation-list-button--balanced|reconciliation-list-button--balanced[^\"]*financial-doc-reconciliation-button/);
  assert.match(financialButton({
    id: "document-2",
    reconciliationId: "reconciliation-2",
    reconciliationDifferenceAmount: 1,
  }), /financial-doc-reconciliation-button[^\"]*reconciliation-list-button--difference|reconciliation-list-button--difference[^\"]*financial-doc-reconciliation-button/);
});

test("Financial Documents preserve the loaded difference through normalization and rendering", () => {
  const normalize = new Function(
    "emptyFinancialDocDraft", "clean", "normalizeDateInput", "normalizeNumber",
    "latestFinancialDocDuplicateWarningMessage",
    `${appFunctionSource("normalizeFinancialDocRowClient")}\nreturn normalizeFinancialDocRowClient;`,
  )(
    () => ({}),
    (value) => String(value ?? "").trim(),
    (value) => String(value ?? ""),
    (value) => {
      const number = Number(value);
      return Number.isFinite(number) ? number : null;
    },
    () => "",
  );
  const button = new Function(
    "clean", "escape",
    `${appFunctionSource("reconciliationListButton")}\nreturn reconciliationListButton;`,
  )(
    (value) => String(value ?? "").trim(),
    (value) => String(value),
  );
  const financialButton = new Function(
    "reconciliationListButton",
    `${appFunctionSource("financialDocReconciliationButton")}\nreturn financialDocReconciliationButton;`,
  )(button);

  const normalized = normalize({
    id: "document-3",
    reconciliationId: "reconciliation-3",
    reconciliationStatus: "complete",
    reconciliationDifferenceAmount: 0,
  });

  assert.equal(normalized.reconciliationDifferenceAmount, 0);
  assert.match(financialButton(normalized), /reconciliation-list-button--balanced/);
});

test("Bank Accounts tables append one Reconciliation column for both sources", () => {
  const button = new Function(
    "clean", "escape",
    `${appFunctionSource("reconciliationListButton")}\nreturn reconciliationListButton;`,
  )(
    (value) => String(value ?? "").trim(),
    (value) => String(value),
  );
  const render = new Function(
    "state", "els", "bankAccountsSourceConfig", "clean", "canAppBankAccounts", "escape",
    "renderBankAccountsCell", "reconciliationListButton",
    `${appFunctionSource("renderBankAccounts")}\nreturn renderBankAccounts;`,
  );
  for (const fixture of [
    {
      source: "cgd-extrato",
      config: { label: "CGD Extrato", reconciliationSourceType: "import_cgd_extrato_ordem", columns: [{ key: "descritivo", label: "Description" }] },
      row: { id: "bank-1", descritivo: "Bank row", reconciliationId: "reconciliation-1", reconciliationDifferenceAmount: "0.00" },
      colorClass: "reconciliation-list-button--balanced",
    },
    {
      source: "cartao-credito",
      config: { label: "Cartao Credito", reconciliationSourceType: "import_cgd_cartao_credito", columns: [{ key: "descricao", label: "Description" }] },
      row: { id: "card-1", descricao: "Card row", reconciliationId: "reconciliation-2", reconciliationDifferenceAmount: "3.00" },
      colorClass: "reconciliation-list-button--difference",
    },
  ]) {
    const state = {
      bankAccountsSource: fixture.source,
      bankAccountsFilters: { dateFrom: "", dateTo: "", description: "" },
      bankAccountsDateSort: "desc",
      bankAccountsRows: [fixture.row],
      bankAccountsLoading: false,
      bankAccountsTruncated: false,
    };
    const els = {
      bankAccountsSource: {}, bankAccountsDateFrom: {}, bankAccountsDateTo: {}, bankAccountsDescription: {},
      bankAccountsRecordsTitle: {}, bankAccountsCount: {}, bankAccountsHead: {}, bankAccountsRows: {},
    };
    const run = render(
      state,
      els,
      () => fixture.config,
      (value) => String(value ?? "").trim(),
      () => true,
      (value) => String(value),
      (row, column) => String(row[column.key]),
      button,
    );
    run();
    assert.match(els.bankAccountsHead.innerHTML, /<th>Reconciliation<\/th>/);
    assert.match(els.bankAccountsRows.innerHTML, new RegExp(fixture.colorClass));
    assert.match(els.bankAccountsRows.innerHTML, new RegExp(`data-reconciliation-source-type="${fixture.config.reconciliationSourceType}"`));
  }
});

test("clicking a Bank Accounts check opens the selected reconciliation from its source", () => {
  class FakeHTMLElement {
    constructor(dataset) {
      this.dataset = dataset;
    }

    closest(selector) {
      return selector === "button[data-bank-accounts-reconciliation]" ? this : null;
    }
  }

  const calls = [];
  const onClick = new Function(
    "HTMLElement", "clean", "openListReconciliation",
    `${appFunctionSource("onBankAccountsRowsClick")}\nreturn onBankAccountsRowsClick;`,
  )(
    FakeHTMLElement,
    (value) => String(value ?? "").trim(),
    (...args) => calls.push(args),
  );

  onClick({ target: new FakeHTMLElement({
    reconciliationId: "reconciliation-8",
    reconciliationSourceType: "import_cgd_cartao_credito",
  }) });

  assert.deepEqual(calls, [["reconciliation-8", "import_cgd_cartao_credito"]]);
});

test("list reconciliation navigation opens Manual on the selected source", async () => {
  const current = {
    activeTab: "automatic",
    candidateSourceType: "financial_documents",
    page: 4,
    selectedReconciliationId: "",
    workspace: {},
    loaded: true,
  };
  const views = [];
  const open = new Function(
    "clean", "canAppFinancialReconciliation", "showToast", "FINANCIAL_RECONCILIATION_SOURCES",
    "financialReconciliationState", "normalizeFinancialReconciliationWorkspace", "setView",
    `${appFunctionSource("openListReconciliation")}\nreturn openListReconciliation;`,
  )(
    (value) => String(value ?? "").trim(),
    () => true,
    () => {},
    {
      financial_documents: "Financial Documents",
      import_cgd_cartao_credito: "CGD Credit Card",
      import_cgd_extrato_ordem: "CGD Bank Statement",
    },
    () => current,
    (value) => value,
    async (...args) => views.push(args),
  );

  await open("reconciliation-9", "import_cgd_cartao_credito");

  assert.equal(current.activeTab, "manual");
  assert.equal(current.candidateSourceType, "import_cgd_cartao_credito");
  assert.equal(current.page, 1);
  assert.equal(current.selectedReconciliationId, "reconciliation-9");
  assert.deepEqual(current.workspace, {
    reconciliation: { id: "reconciliation-9", base_source_type: "import_cgd_cartao_credito" },
  });
  assert.equal(current.loaded, false);
  assert.deepEqual(views, [["financial-reconciliation", { financialReconciliationTab: "manual" }]]);
});
