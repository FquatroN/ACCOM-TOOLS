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
