# Reconciliation Inline Completion Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the reconciliation completion popup with an always-visible inline comment field whose requirement follows the current difference, and make locked-record details more compact.

**Architecture:** Keep the feature inside the existing reconciliation front end. Add small pure helpers for completion presentation and reconciliation-scoped draft state, render the controls inside `renderFinancialReconciliationCurrent`, and dispatch through the existing `runFinancialReconciliationAction` API path. Remove the obsolete modal wiring and scope the density change to locked-record CSS.

**Tech Stack:** Browser JavaScript, server-rendered HTML strings, CSS, Node.js built-in test runner.

## Global Constraints

- The completion comment textarea is always visible for an active reconciliation that is not complete.
- A zero difference makes the comment optional and uses `complete`.
- A non-zero difference makes the comment mandatory, rejects whitespace-only input, and uses `force_complete`.
- No completion or force-completion popup remains.
- Draft text survives Current reconciliation re-renders and clears only after successful completion, deletion, opening another reconciliation, or starting a new reconciliation.
- Existing API and database validation remains unchanged.
- Compact spacing applies only to locked records in the Current reconciliation panel.
- Do not modify reconciliation arithmetic, source rules, database migrations, or other lifecycle confirmation dialogs.

---

## File structure

- Modify `app-main.js` — own the inline completion state, rendering, input/click behavior, lifecycle cleanup, and removal of modal-specific code.
- Modify `index.html` — remove the obsolete completion modal markup.
- Modify `styles.css` — style the inline comment controls and tighten locked-record detail spacing.
- Create `tests/reconciliation-inline-completion.test.js` — executable behavior tests for presentation, draft state, and action payloads, plus the no-modal DOM contract.
- Modify `tests/reconciliation-density.test.js` — retain the existing detail-content assertions and add the exact compact CSS contract.

### Task 1: Completion presentation and reconciliation-scoped draft state

**Files:**
- Create: `tests/reconciliation-inline-completion.test.js`
- Modify: `app-main.js:1292-1302`
- Modify: `app-main.js:21315-21330`

**Interfaces:**
- Produces: `financialReconciliationCompletionPresentation(difference: number, itemCount: number, comment: string): { action: "complete" | "force_complete", label: string, required: boolean, disabled: boolean }`.
- Produces: `financialReconciliationCompletionDraft(reconciliationId: string): string`.
- Produces: `updateFinancialReconciliationCompletionDraft(reconciliationId: string, value: string): void`.
- Produces: `clearFinancialReconciliationCompletionDraft(): void`.
- State shape: `state.financialReconciliation.completionCommentDraft = { reconciliationId: string, value: string }`.

- [ ] **Step 1: Write the failing presentation tests**

Create `tests/reconciliation-inline-completion.test.js` with the standard source extractor used by `tests/reconciliation-density.test.js`, then add these cases:

```js
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
```

- [ ] **Step 2: Run the focused test and confirm the red state**

Run: `node --test --test-isolation=none tests/reconciliation-inline-completion.test.js`

Expected: FAIL because `financialReconciliationCompletionPresentation` is not defined.

- [ ] **Step 3: Add the presentation helper**

Add this near `financialReconciliationDifference` in `app-main.js`:

```js
function financialReconciliationCompletionPresentation(difference, itemCount, comment) {
  const required = Number(difference) !== 0;
  return {
    action: required ? "force_complete" : "complete",
    label: required ? "Force complete" : "Complete reconciliation",
    required,
    disabled: Number(itemCount) <= 0 || (required && !clean(comment)),
  };
}
```

- [ ] **Step 4: Write the failing draft-state tests**

Append tests that compile the three draft helpers against a mutable local state:

```js
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
```

- [ ] **Step 5: Run the focused test and confirm the second red state**

Run: `node --test --test-isolation=none tests/reconciliation-inline-completion.test.js`

Expected: the presentation tests PASS and the draft tests FAIL because the draft helpers are undefined.

- [ ] **Step 6: Add the draft state and helpers**

Initialize the state:

```js
completionCommentDraft: { reconciliationId: "", value: "" },
```

Add these functions near `financialReconciliationState`:

```js
function financialReconciliationCompletionDraft(reconciliationId) {
  const current = financialReconciliationState();
  const normalizedId = clean(reconciliationId);
  if (clean(current.completionCommentDraft?.reconciliationId) !== normalizedId) {
    current.completionCommentDraft = { reconciliationId: normalizedId, value: "" };
  }
  return String(current.completionCommentDraft?.value ?? "");
}

function updateFinancialReconciliationCompletionDraft(reconciliationId, value) {
  financialReconciliationState().completionCommentDraft = {
    reconciliationId: clean(reconciliationId),
    value: String(value ?? ""),
  };
}

function clearFinancialReconciliationCompletionDraft() {
  financialReconciliationState().completionCommentDraft = { reconciliationId: "", value: "" };
}
```

- [ ] **Step 7: Run the Task 1 tests**

Run: `node --test --test-isolation=none tests/reconciliation-inline-completion.test.js`

Expected: all Task 1 tests PASS.

- [ ] **Step 8: Commit Task 1**

```bash
git add app-main.js tests/reconciliation-inline-completion.test.js
git commit -m "test: define inline reconciliation completion state"
```

### Task 2: Inline controls, immediate submission, and modal removal

**Files:**
- Modify: `app-main.js:1940-1948`
- Modify: `app-main.js:2599-2611`
- Modify: `app-main.js:21528-21724`
- Modify: `index.html:3880-3887`
- Modify: `tests/reconciliation-inline-completion.test.js`

**Interfaces:**
- Consumes: all Task 1 presentation and draft helpers.
- Produces: `onFinancialReconciliationCurrentInput(event: Event): void`.
- Produces: `completeFinancialReconciliation(): void`.
- Continues to call `runFinancialReconciliationAction({ action, reconciliationId, comment })`.

- [ ] **Step 1: Write failing no-modal and inline-render contract tests**

Append:

```js
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
```

- [ ] **Step 2: Write failing completion-dispatch behavior tests**

Compile `completeFinancialReconciliation` with stubs and cover both paths:

```js
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
```

- [ ] **Step 3: Run the tests and confirm the red state**

Run: `node --test --test-isolation=none tests/reconciliation-inline-completion.test.js`

Expected: FAIL because the modal still exists and `completeFinancialReconciliation` is undefined.

- [ ] **Step 4: Render the inline textarea and button**

In `renderFinancialReconciliationCurrent`, derive:

```js
const completionDraft = financialReconciliationCompletionDraft(reconciliation.id);
const completion = financialReconciliationCompletionPresentation(difference, workspace.items.length, completionDraft);
```

Build `completionControls` with the existing completed summary as the `complete` branch and this scoped control block as the non-complete branch, then render it immediately after the locked-record list. This explicit branch ensures completed reconciliations never render editable controls:

```js
<div class="financial-reconciliation-completion">
  <label>Completion comment <span class="field-hint">${completion.required ? "Comment is required because the difference is not zero" : "Comment is optional because the difference is zero"}</span>
    <textarea data-financial-reconciliation-completion-comment rows="3" ${completion.required ? 'required aria-required="true"' : ""} placeholder="Add a completion comment.">${escape(completionDraft)}</textarea>
  </label>
  <button type="button" ${completion.required ? 'class="danger"' : ""} data-financial-reconciliation-complete ${completion.disabled ? "disabled" : ""}>${completion.label}</button>
</div>
```

Do not render the controls for a completed reconciliation. Preserve the existing completed summary.

- [ ] **Step 5: Wire draft input without re-rendering the textarea**

Register a delegated `input` listener on `els.financialReconciliationCurrent`. Implement `onFinancialReconciliationCurrentInput` to:

1. match `[data-financial-reconciliation-completion-comment]`;
2. save the raw textarea value with `updateFinancialReconciliationCompletionDraft`;
3. recompute presentation from the current difference and locked-item count;
4. update only the inline completion button's `disabled` property.

This avoids losing focus or cursor position while still making the force-complete button responsive.

- [ ] **Step 6: Implement direct completion dispatch**

Add:

```js
function completeFinancialReconciliation() {
  const reconciliation = financialReconciliationActiveRecord();
  if (!reconciliation) return;
  const current = financialReconciliationState();
  const comment = financialReconciliationCompletionDraft(reconciliation.id);
  const presentation = financialReconciliationCompletionPresentation(
    financialReconciliationDifference(reconciliation),
    current.workspace?.items?.length || 0,
    comment,
  );
  if (presentation.disabled) return;
  runFinancialReconciliationAction({
    action: presentation.action,
    reconciliationId: reconciliation.id,
    comment: clean(comment),
  });
}
```

Update `onFinancialReconciliationCurrentClick` to call this function when the inline button is clicked.

- [ ] **Step 7: Clear drafts only after successful lifecycle transitions**

Inside the successful `try` path of `runFinancialReconciliationAction`, after `api` resolves, clear the draft when `payload.action` is `complete`, `force_complete`, or `delete`. Do not clear it in `catch`, so failed completion attempts can be retried.

Call `clearFinancialReconciliationCompletionDraft()` in `startNewFinancialReconciliation` before loading the empty workspace. Opening a different history record is covered by `financialReconciliationCompletionDraft(record.id)`, which resets a draft whose reconciliation ID differs.

Before implementing the cleanup, add this executable regression test and verify that it fails because `runFinancialReconciliationAction` does not yet call the clear helper:

```js
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
```

Run: `node --test --test-isolation=none tests/reconciliation-inline-completion.test.js`

Expected before cleanup implementation: the success assertion FAILS with `0 !== 1`. Expected after implementation: both assertions PASS.

- [ ] **Step 8: Remove all modal code**

Remove from `app-main.js`:

- the six modal element references;
- the confirm/input/backdrop event listeners;
- `renderFinancialReconciliationCompletionModal` and its call from `renderFinancialReconciliation`;
- `openFinancialReconciliationCompletionModal`;
- `closeFinancialReconciliationCompletionModal`;
- `confirmFinancialReconciliationCompletion`.

Remove the entire `#financial-reconciliation-complete-modal` section from `index.html`. Keep the delete and reopen confirmation dialogs unchanged.

- [ ] **Step 9: Run the focused completion tests**

Run: `node --test --test-isolation=none tests/reconciliation-inline-completion.test.js`

Expected: all completion tests PASS.

- [ ] **Step 10: Commit Task 2**

```bash
git add app-main.js index.html tests/reconciliation-inline-completion.test.js
git commit -m "feat: complete reconciliations inline"
```

### Task 3: Compact locked-record styling and full verification

**Files:**
- Modify: `styles.css:6395-6421`
- Modify: `tests/reconciliation-density.test.js:255-269`

**Interfaces:**
- Consumes: `.financial-reconciliation-completion`, `.financial-reconciliation-items li`, and `.financial-reconciliation-item-details` markup from Task 2.
- Produces: scoped visual density rules only; no JavaScript interface.

- [ ] **Step 1: Write failing compact-spacing assertions**

Add to the reconciliation density test:

```js
assert.match(css, /\.financial-reconciliation-completion\s*\{[\s\S]*display:\s*grid;[\s\S]*gap:\s*\.35rem;/);
assert.match(css, /\.financial-reconciliation-completion textarea\s*\{[\s\S]*min-height:\s*4\.5rem;/);
assert.match(css, /\.financial-reconciliation-items li\s*\{[\s\S]*column-gap:\s*\.45rem;[\s\S]*row-gap:\s*\.12rem;/);
assert.match(css, /\.financial-reconciliation-item-details\s*\{[\s\S]*line-height:\s*1\.15;[\s\S]*margin-top:\s*0;/);
```

- [ ] **Step 2: Run the density test and confirm the red state**

Run: `node --test --test-isolation=none tests/reconciliation-density.test.js`

Expected: FAIL because the new inline-control and compact-spacing declarations do not exist.

- [ ] **Step 3: Add scoped inline-control styling**

Add near the existing Current reconciliation rules:

```css
.financial-reconciliation-completion {
  display: grid;
  gap: .35rem;
  margin-top: .65rem;
}

.financial-reconciliation-completion label {
  display: grid;
  gap: .25rem;
  font-size: .74rem;
}

.financial-reconciliation-completion textarea {
  min-height: 4.5rem;
  resize: vertical;
  font: inherit;
}
```

Keep the existing global button styling and danger treatment; do not introduce a second modal-like container.

- [ ] **Step 4: Tighten only the locked-record source/detail gap**

Replace the single `.45rem` grid gap on `.financial-reconciliation-items li` with:

```css
column-gap: .45rem;
row-gap: .12rem;
```

Update `.financial-reconciliation-item-details` with:

```css
line-height: 1.15;
margin-top: 0;
```

Do not change `.financial-reconciliation-audit` spacing.

- [ ] **Step 5: Run focused UI tests**

Run: `node --test --test-isolation=none tests/reconciliation-inline-completion.test.js tests/reconciliation-density.test.js`

Expected: all focused tests PASS.

- [ ] **Step 6: Run the full automated suite and syntax checks**

Run:

```bash
node --check app-main.js
node --test --test-isolation=none tests/*.test.js
git diff --check
```

Expected: JavaScript syntax is valid, all tests PASS, and `git diff --check` produces no output.

- [ ] **Step 7: Verify both flows in the local browser**

With the existing local app running:

1. Open a started reconciliation with a zero difference.
2. Confirm the textarea is visible, marked optional, and completion works without a comment or popup.
3. Open or create a started reconciliation with a non-zero difference.
4. Confirm Force complete is disabled for empty and whitespace-only text.
5. Enter a comment and confirm Force complete submits immediately without a popup.
6. Before submitting, add or remove a record and confirm the draft remains.
7. Confirm source and detail text in each locked record are visibly closer without overlap.
8. Confirm a failed completion request leaves the draft available for retry.

- [ ] **Step 8: Commit Task 3**

```bash
git add styles.css tests/reconciliation-density.test.js
git commit -m "style: compact reconciliation completion details"
```

- [ ] **Step 9: Final scope check**

Run: `git status --short` and `git diff main...HEAD -- app-main.js index.html styles.css tests/reconciliation-inline-completion.test.js tests/reconciliation-density.test.js`

Expected: only the five planned implementation files differ from the execution base; unrelated pre-existing untracked files remain untouched.
