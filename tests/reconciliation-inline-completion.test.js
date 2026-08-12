const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const appMain = fs.readFileSync(path.join(root, "app-main.js"), "utf8");

function appFunctionSource(name) {
  const start = appMain.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `${name} should be defined in app-main.js`);
  const bodyStart = appMain.indexOf("{", appMain.indexOf(")", start));
  let depth = 0;
  for (let index = bodyStart; index < appMain.length; index += 1) {
    if (appMain[index] === "{") depth += 1;
    if (appMain[index] === "}") depth -= 1;
    if (depth === 0) return appMain.slice(start, index + 1);
  }
  throw new Error(`Could not extract ${name}`);
}

const presentation = new Function("clean", `
${appFunctionSource("financialReconciliationCompletionPresentation")}
return financialReconciliationCompletionPresentation;
`)((value) => String(value || "").trim());

test("zero difference allows an optional empty completion comment", () => {
  assert.deepEqual(presentation(0, 1, ""), {
    action: "complete",
    label: "Complete reconciliation",
    required: false,
    disabled: false,
  });
});

test("non-zero difference requires a non-whitespace force-completion comment", () => {
  assert.deepEqual(presentation(10.25, 1, "  "), {
    action: "force_complete",
    label: "Force complete",
    required: true,
    disabled: true,
  });
  assert.equal(presentation(10.25, 1, "Reviewed variance").disabled, false);
});

test("completion remains unavailable without locked items", () => {
  assert.equal(presentation(0, 0, "optional note").disabled, true);
  assert.equal(presentation(12, 0, "variance approved").disabled, true);
});

function completionDraftHarness() {
  const current = { completionCommentDraft: { reconciliationId: "", value: "" } };
  const helpers = new Function("clean", "financialReconciliationState", `
${appFunctionSource("financialReconciliationCompletionDraft")}
${appFunctionSource("updateFinancialReconciliationCompletionDraft")}
${appFunctionSource("clearFinancialReconciliationCompletionDraft")}
return { financialReconciliationCompletionDraft, updateFinancialReconciliationCompletionDraft, clearFinancialReconciliationCompletionDraft };
  `)((value) => String(value || "").trim(), () => current);
  return { current, ...helpers };
}

test("completion draft survives renders for the same reconciliation", () => {
  const draft = completionDraftHarness();
  draft.updateFinancialReconciliationCompletionDraft("rec-1", "  keep my spacing  ");
  assert.equal(draft.financialReconciliationCompletionDraft("rec-1"), "  keep my spacing  ");
});

test("opening another reconciliation clears the previous draft", () => {
  const draft = completionDraftHarness();
  draft.updateFinancialReconciliationCompletionDraft("rec-1", "old note");
  assert.equal(draft.financialReconciliationCompletionDraft("rec-2"), "");
  assert.deepEqual(draft.current.completionCommentDraft, { reconciliationId: "rec-2", value: "" });
});

test("completion draft can be cleared after a successful lifecycle action", () => {
  const draft = completionDraftHarness();
  draft.updateFinancialReconciliationCompletionDraft("rec-1", "approved variance");
  draft.clearFinancialReconciliationCompletionDraft();
  assert.deepEqual(draft.current.completionCommentDraft, { reconciliationId: "", value: "" });
});

test("completion uses inline controls and removes the modal contract", () => {
  assert.doesNotMatch(html, /financial-reconciliation-complete-modal/);
  assert.doesNotMatch(appMain, /renderFinancialReconciliationCompletionModal|openFinancialReconciliationCompletionModal|closeFinancialReconciliationCompletionModal|confirmFinancialReconciliationCompletion/);
  assert.doesNotMatch(appMain, /financialReconciliationCompleteModal|financialReconciliationForceComment|financialReconciliationConfirmComplete|financialReconciliationConfirmForce/);
  assert.match(appMain, /data-financial-reconciliation-completion-comment/);
  assert.match(appMain, /Completion comment/);
  assert.match(appMain, /Comment is required because the difference is not zero/);
  assert.match(appMain, /data-financial-reconciliation-complete/);
  assert.match(appMain, /const completionControls = complete \?/);
});

function completionActionHarness({ difference, items, comment }) {
  const calls = [];
  const reconciliation = { id: "rec-1", difference_amount: difference };
  const complete = new Function(
    "financialReconciliationActiveRecord",
    "financialReconciliationState",
    "financialReconciliationDifference",
    "financialReconciliationCompletionDraft",
    "financialReconciliationCompletionPresentation",
    "clean",
    "runFinancialReconciliationAction",
    `${appFunctionSource("completeFinancialReconciliation")}\nreturn completeFinancialReconciliation;`,
  )(
    () => reconciliation,
    () => ({ workspace: { items } }),
    (value) => Number(value.difference_amount),
    () => comment,
    presentation,
    (value) => String(value || "").trim(),
    (payload) => calls.push(payload),
  );
  complete();
  return calls;
}

test("zero difference completes immediately with an optional comment", () => {
  assert.deepEqual(completionActionHarness({ difference: 0, items: [{}], comment: "optional note" }), [
    { action: "complete", reconciliationId: "rec-1", comment: "optional note" },
  ]);
});

test("non-zero difference submits only a valid mandatory comment", () => {
  assert.deepEqual(completionActionHarness({ difference: 8, items: [{}], comment: "   " }), []);
  assert.deepEqual(completionActionHarness({ difference: 8, items: [{}], comment: "variance approved" }), [
    { action: "force_complete", reconciliationId: "rec-1", comment: "variance approved" },
  ]);
});

async function completionCleanupCount(api, action = "force_complete") {
  const current = {
    pendingAction: "",
    workspace: { sourceConfig: {}, items: [{}] },
    selectedReconciliationId: "rec-1",
  };
  let clears = 0;
  const source = appFunctionSource("runFinancialReconciliationAction").replace(/^function /, "async function ");
  const runAction = new Function(
    "financialReconciliationState", "api", "normalizeFinancialReconciliationWorkspace", "clean",
    "loadFinancialReconciliationWorkspace", "showToast", "setFinancialReconciliationStatus",
    "renderFinancialReconciliation", "clearFinancialReconciliationCompletionDraft",
    `${source}\nreturn runFinancialReconciliationAction;`,
  )(
    () => current,
    api,
    (value) => ({ items: [], audit: [], history: [], ...value }),
    (value) => String(value || "").trim(),
    async () => {},
    () => {},
    () => {},
    () => {},
    () => { clears += 1; },
  );
  await runAction({ action, reconciliationId: "rec-1", comment: "approved" });
  return clears;
}

test("successful completion clears its draft but a failed completion keeps it", async () => {
  assert.equal(await completionCleanupCount(async () => ({ reconciliation: { id: "rec-1", status: "complete" } })), 1);
  assert.equal(await completionCleanupCount(async () => { throw new Error("offline"); }), 0);
});

test("successful deletion clears its draft", async () => {
  assert.equal(await completionCleanupCount(async () => ({ deleted: true }), "delete"), 1);
});

test("starting a new reconciliation clears its draft", async () => {
  const current = {
    selectedReconciliationId: "rec-1",
    workspace: { reconciliation: { id: "rec-1", status: "complete" }, items: [{}], audit: [{}] },
    loaded: true,
  };
  let clears = 0;
  const source = appFunctionSource("startNewFinancialReconciliation").replace(/^function /, "async function ");
  const startNew = new Function(
    "financialReconciliationState", "financialReconciliationActiveRecord", "clean",
    "clearFinancialReconciliationCompletionDraft", "loadFinancialReconciliationWorkspace",
    "setFinancialReconciliationStatus",
    `${source}\nreturn startNewFinancialReconciliation;`,
  )(
    () => current,
    () => current.workspace.reconciliation,
    (value) => String(value || "").trim(),
    () => { clears += 1; },
    async () => {},
    () => {},
  );
  await startNew();
  assert.equal(clears, 1);
  assert.equal(current.selectedReconciliationId, "");
});
