const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const appMain = fs.readFileSync(path.join(root, "app-main.js"), "utf8");

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

function sectionById(id) {
  const start = html.indexOf(`<section id="${id}"`);
  assert.notEqual(start, -1, `${id} should exist`);
  const tags = /<\/?section\b[^>]*>/gi;
  tags.lastIndex = start;
  let depth = 0;
  let match;
  while ((match = tags.exec(html))) {
    depth += match[0].startsWith("</") ? -1 : 1;
    if (depth === 0) return html.slice(start, tags.lastIndex);
  }
  throw new Error(`Could not find the closing section for ${id}`);
}

function renderCandidates({ sourceType = "financial_documents", canOpenFinancialDocs = true } = {}) {
  const current = {
    candidateSourceType: sourceType,
    workspace: {
      candidates: [{
        id: 'doc-1" onclick="unsafe',
        source_date: "2026-08-20",
        description: "Invoice <script>alert(1)</script>",
        amount: 12.5,
      }],
      counts: { notStarted: 1, started: 0, complete: 0 },
      sourceConfig: { sourceType, columns: ["description"] },
    },
  };
  const els = {
    financialReconciliationTableHead: { innerHTML: "" },
    financialReconciliationRows: { innerHTML: "" },
    financialReconciliationCount: { innerHTML: "" },
    financialReconciliationStart: { hidden: false, disabled: false, textContent: "" },
  };
  const render = new Function(
    "financialReconciliationState", "normalizeFinancialReconciliationWorkspace", "financialReconciliationActiveRecord",
    "reconciliationRulesFor", "escape", "formatDateOnly", "clean", "formatMoney",
    "financialReconciliationStatusMarkup", "financialReconciliationSourceLabel", "canAppFinancialDocs", "els",
    `${appFunctionSource("renderFinancialReconciliationCandidates")}\nreturn renderFinancialReconciliationCandidates;`,
  )(
    () => current,
    (value) => value,
    () => null,
    () => [{ sourceType: "import_cgd_extrato_ordem" }],
    new Function(`${appFunctionSource("escape")}\nreturn escape;`)(),
    (value) => value,
    (value) => String(value ?? "").trim(),
    (value) => `${Number(value).toFixed(2)} EUR`,
    (status) => `<span>${status}</span>`,
    (value) => value,
    () => canOpenFinancialDocs,
    els,
  );
  render();
  return els.financialReconciliationRows.innerHTML;
}

test("Financial Document candidates expose an escaped description link only to authorized users", () => {
  const authorized = renderCandidates();
  assert.match(authorized, /data-financial-reconciliation-open-financial-doc="doc-1&quot; onclick=&quot;unsafe"/);
  assert.match(authorized, />Invoice &lt;script&gt;alert\(1\)&lt;\/script&gt;<\/button>/);
  assert.doesNotMatch(authorized, /<script>/);

  const unauthorized = renderCandidates({ canOpenFinancialDocs: false });
  assert.doesNotMatch(unauthorized, /data-financial-reconciliation-open-financial-doc/);
  assert.match(unauthorized, />Invoice &lt;script&gt;alert\(1\)&lt;\/script&gt;<\/td>/);

  const bank = renderCandidates({ sourceType: "import_cgd_extrato_ordem" });
  assert.doesNotMatch(bank, /data-financial-reconciliation-open-financial-doc/);
});

test("clicking a Financial Document description opens the existing editor for the rendered record", () => {
  class FakeHTMLElement {
    constructor(dataset) {
      this.dataset = dataset;
      this.disabled = false;
    }

    closest(selector) {
      if (selector === "[data-financial-reconciliation-open-financial-doc]") return this;
      return null;
    }
  }

  const opened = [];
  const actions = [];
  let canOpenFinancialDocs = true;
  const onRowsClick = new Function(
    "HTMLElement", "clean", "canAppFinancialDocs", "openFinancialDocModal", "runFinancialReconciliationAction", "financialReconciliationActiveRecord",
    `${appFunctionSource("onFinancialReconciliationRowsClick")}\nreturn onFinancialReconciliationRowsClick;`,
  )(
    FakeHTMLElement,
    (value) => String(value ?? "").trim(),
    () => canOpenFinancialDocs,
    (id, options) => opened.push({ id, options }),
    (payload) => actions.push(payload),
    () => null,
  );

  onRowsClick({ target: new FakeHTMLElement({ financialReconciliationOpenFinancialDoc: "document-42" }) });
  canOpenFinancialDocs = false;
  onRowsClick({ target: new FakeHTMLElement({ financialReconciliationOpenFinancialDoc: "document-99" }) });

  assert.deepEqual(opened, [{ id: "document-42", options: { origin: "financial-reconciliation" } }]);
  assert.deepEqual(actions, []);
});

test("the shared Financial Document editor records and clears its Reconciliation origin", async () => {
  class FakeHTMLElement {
    constructor() {
      this.isConnected = true;
      this.focused = false;
    }

    focus() {
      this.focused = true;
    }
  }
  const returnFocus = new FakeHTMLElement();
  const dialog = new FakeHTMLElement();
  let resolvePreview;
  let signalPreviewStarted;
  const previewStarted = new Promise((resolve) => { signalPreviewStarted = resolve; });
  const previewFinished = new Promise((resolve) => { resolvePreview = resolve; });
  const state = { currentView: "financial-reconciliation", financialDocsModalRequestToken: 0 };
  const els = {
    financialDocsModal: { hidden: true },
    financialDocsModalDialog: dialog,
    financialDocsAttachmentInput: { value: "" },
    financialDocsParseInput: { value: "" },
  };
  const open = new Function(
    "state", "els", "clean", "setFinancialDocsModalStatus", "setFinancialDocsDuplicateWarning",
    "loadFinancialDocDetail", "normalizeFinancialDocRowClient", "emptyFinancialDocDraft",
    "syncFinancialDocModalBodyState", "renderFinancialDocEditor", "renderFinancialDocPreview", "showToast", "document", "HTMLElement",
    `${appFunctionSource("openFinancialDocModal")}\nreturn openFinancialDocModal;`,
  )(
    state,
    els,
    (value) => String(value ?? "").trim(),
    () => {},
    () => {},
    async (id) => ({ id, description: "Editable" }),
    (value) => value,
    () => ({}),
    () => {},
    () => {},
    async () => { signalPreviewStarted(); return previewFinished; },
    () => {},
    { activeElement: returnFocus },
    FakeHTMLElement,
  );
  const close = new Function(
    "state", "els", "setFinancialDocsDuplicateWarning", "URL", "revokeFinancialDocPreviewUrl", "syncFinancialDocModalBodyState",
    `${appFunctionSource("closeFinancialDocModal")}\nreturn closeFinancialDocModal;`,
  )(
    state,
    els,
    () => {},
    { revokeObjectURL() {} },
    () => {},
    () => {},
  );

  const opening = open("document-42", { origin: "financial-reconciliation" });
  await previewStarted;
  assert.equal(dialog.focused, true, "focus should move before a slow attachment preview finishes");
  resolvePreview();
  await opening;
  assert.equal(state.financialDocsModalOpen, true);
  assert.equal(state.financialDocsModalOrigin, "financial-reconciliation");
  assert.equal(els.financialDocsModal.hidden, false);
  assert.equal(dialog.focused, true);

  close();
  assert.equal(state.financialDocsModalOpen, false);
  assert.equal(state.financialDocsModalOrigin, "");
  assert.equal(els.financialDocsModal.hidden, true);
  assert.equal(returnFocus.focused, true);
});

test("a superseded attachment preview cannot overwrite the active Financial Document", async () => {
  const state = {
    financialDocsModalOpen: true,
    financialDocsModalRequestToken: 1,
    financialDocsDraft: { id: "document-1", driveFileId: "file-1", mimeType: "image/png" },
    financialDocsAttachment: null,
    financialDocsPreviewUrl: "",
  };
  const els = { financialDocsPreview: { innerHTML: "" } };
  const pending = new Map();
  const preview = new Function(
    "state", "els", "revokeFinancialDocPreviewUrl", "emptyFinancialDocDraft", "clean", "apiBlob", "URL", "escape",
    `${appFunctionSource("renderFinancialDocPreview")}\nreturn renderFinancialDocPreview;`,
  )(
    state,
    els,
    () => { state.financialDocsPreviewUrl = ""; },
    () => ({}),
    (value) => String(value ?? "").trim(),
    (url) => new Promise((resolve) => pending.set(url, resolve)),
    { createObjectURL: (blob) => `blob:${blob.name}` },
    (value) => String(value),
  );

  const first = preview({ requestToken: 1, documentId: "document-1" });
  state.financialDocsModalRequestToken = 2;
  state.financialDocsDraft = { id: "document-2", driveFileId: "file-2", mimeType: "image/png" };
  const second = preview({ requestToken: 2, documentId: "document-2" });
  pending.get("/api/financial-docs-file?id=document-1")({ name: "document-1", type: "image/png" });
  assert.equal(await first, false);
  assert.doesNotMatch(els.financialDocsPreview.innerHTML, /document-1/);

  pending.get("/api/financial-docs-file?id=document-2")({ name: "document-2", type: "image/png" });
  assert.equal(await second, true);
  assert.match(els.financialDocsPreview.innerHTML, /blob:document-2/);

  state.financialDocsDraft = { id: "document-3", driveFileId: "file-3", mimeType: "image/png" };
  const savedAttachmentPreview = preview({ requestToken: 2, documentId: "document-3" });
  state.financialDocsAttachment = { previewUrl: "blob:local-upload", upload: { mimeType: "image/png" } };
  assert.equal(await preview({ requestToken: 2, documentId: "document-3" }), true);
  pending.get("/api/financial-docs-file?id=document-3")({ name: "old-saved-file", type: "image/png" });
  assert.equal(await savedAttachmentPreview, false);
  assert.match(els.financialDocsPreview.innerHTML, /blob:local-upload/);
  assert.doesNotMatch(els.financialDocsPreview.innerHTML, /old-saved-file/);
});

test("late and overlapping detail loads cannot reveal a stale Financial Document", async () => {
  class FakeHTMLElement {}
  const state = { currentView: "financial-reconciliation", financialDocsModalRequestToken: 0 };
  const els = { financialDocsModal: { hidden: true }, financialDocsModalDialog: { focus() {} } };
  const pending = new Map();
  const loadFinancialDocDetail = (id) => new Promise((resolve) => pending.set(id, resolve));
  const open = new Function(
    "state", "els", "clean", "setFinancialDocsModalStatus", "setFinancialDocsDuplicateWarning",
    "loadFinancialDocDetail", "normalizeFinancialDocRowClient", "emptyFinancialDocDraft",
    "syncFinancialDocModalBodyState", "renderFinancialDocEditor", "renderFinancialDocPreview", "showToast", "document", "HTMLElement",
    `${appFunctionSource("openFinancialDocModal")}\nreturn openFinancialDocModal;`,
  )(
    state,
    els,
    (value) => String(value ?? "").trim(),
    () => {},
    () => {},
    loadFinancialDocDetail,
    (value) => value,
    () => ({}),
    () => {},
    () => {},
    async () => {},
    () => {},
    { activeElement: null },
    FakeHTMLElement,
  );

  const first = open("document-1", { origin: "financial-reconciliation" });
  const second = open("document-2", { origin: "financial-reconciliation" });
  pending.get("document-2")({ id: "document-2" });
  assert.equal(await second, true);
  pending.get("document-1")({ id: "document-1" });
  assert.equal(await first, false);
  assert.equal(state.financialDocsDraft.id, "document-2");

  els.financialDocsModal.hidden = true;
  state.financialDocsModalOpen = false;
  const third = open("document-3", { origin: "financial-reconciliation" });
  state.currentView = "guests";
  pending.get("document-3")({ id: "document-3" });
  assert.equal(await third, false);
  assert.equal(els.financialDocsModal.hidden, true);
  assert.equal(state.financialDocsDraft.id, "document-2");
});

test("successful popup mutations refresh Reconciliation candidates only for that origin", async () => {
  const state = { financialDocsModalOrigin: "financial-reconciliation" };
  const calls = [];
  const refresh = new Function(
    "state", "clean", "loadFinancialReconciliationWorkspace",
    `${appFunctionSource("refreshFinancialReconciliationAfterFinancialDocMutation")}\nreturn refreshFinancialReconciliationAfterFinancialDocMutation;`,
  )(
    state,
    (value) => String(value ?? "").trim(),
    async (options) => { calls.push(options); return true; },
  );

  assert.equal(await refresh(), true);
  assert.deepEqual(calls, [{ silent: true }]);

  state.financialDocsModalOrigin = "";
  assert.equal(await refresh(), false);
  assert.deepEqual(calls, [{ silent: true }]);
});

test("saving and deleting through the shared popup invoke the Reconciliation refresh", async () => {
  const state = {
    financialDocsModalOrigin: "financial-reconciliation",
    financialDocsRows: [{ id: "document-42", description: "Before" }],
    financialDocsDraft: { id: "document-42", description: "Before" },
    financialDocsAttachment: null,
    financialDocsEditingId: "",
    financialDocsLastOpenedId: "document-42",
  };
  const refreshedOrigins = [];
  let resolveSave;
  const saveRequest = new Promise((resolve) => { resolveSave = resolve; });
  const save = new Function(
    "state", "financialDocDraftFromInputs", "clean", "applyFinancialDocRuleToDraft", "findFinancialDocRuleByEntity",
    "setFinancialDocsDuplicateWarning", "setFinancialDocsModalStatus", "saveFinancialDocRequest", "window",
    "FINANCIAL_DOCS_DUPLICATE_CONFIRM_TEXT", "normalizeFinancialDocRowClient", "sortFinancialDocRows",
    "refreshFinancialDocEntitiesAfterSave", "renderFinancialDocs", "renderFinancialDocEditor", "renderFinancialDocPreview",
    "setFinancialDocsStatus", "showToast", "refreshFinancialReconciliationAfterFinancialDocMutation",
    `${appFunctionSource("saveFinancialDoc")}\nreturn saveFinancialDoc;`,
  )(
    state,
    () => ({ id: "document-42", description: "After", supplierName: "Supplier" }),
    (value) => String(value ?? "").trim(),
    (value) => value,
    () => null,
    () => {},
    () => {},
    async () => saveRequest,
    { confirm: () => true },
    "Confirm duplicate",
    (value) => value,
    (rows) => rows,
    async () => {},
    () => {},
    () => {},
    async () => {},
    () => {},
    () => {},
    async (origin) => { refreshedOrigins.push(origin); return true; },
  );
  const saving = save();
  await Promise.resolve();
  state.financialDocsModalOrigin = "";
  resolveSave({ row: { id: "document-42", description: "After" } });
  await saving;

  const remove = new Function(
    "state", "clean", "window", "setFinancialDocsStatus", "setFinancialDocsModalStatus", "api",
    "closeFinancialDocModal", "renderFinancialDocs", "showToast", "refreshFinancialReconciliationAfterFinancialDocMutation",
    `${appFunctionSource("deleteFinancialDoc")}\nreturn deleteFinancialDoc;`,
  )(
    state,
    (value) => String(value ?? "").trim(),
    { confirm: () => true },
    () => {},
    () => {},
    async () => ({}),
    () => { state.financialDocsModalOrigin = ""; },
    () => {},
    () => {},
    async (origin) => { refreshedOrigins.push(origin); return true; },
  );
  state.financialDocsModalOrigin = "financial-reconciliation";
  await remove("document-42", { fromModal: true });

  assert.deepEqual(refreshedOrigins, ["financial-reconciliation", "financial-reconciliation"]);
});

test("the editable Financial Document modal is not trapped inside the hidden feature view", () => {
  const financialDocsView = sectionById("view-financial-docs");
  const modal = sectionById("financial-docs-modal");

  assert.doesNotMatch(financialDocsView, /id="financial-docs-modal"/);
  assert.match(modal, /id="financial-docs-modal-dialog"[^>]*role="dialog"[^>]*aria-modal="true"[^>]*aria-labelledby="financial-docs-modal-title"/);
  assert.match(modal, /id="financial-docs-save"/);
  assert.match(modal, /id="financial-docs-delete"/);
  assert.match(modal, /id="financial-docs-description-field"[^>]*type="text"/);
  assert.match(modal, /id="financial-docs-supplier-name-field"[^>]*type="text"/);
  assert.doesNotMatch(modal, /id="financial-docs-description-field"[^>]*readonly/);
});
