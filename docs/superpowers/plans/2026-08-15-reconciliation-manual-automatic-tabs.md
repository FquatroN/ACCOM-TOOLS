# Reconciliation Manual and Automatic Tabs Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Separate the Reconciliation application into Manual reconciliation and Automatic reconciliation tabs while keeping one shared history visible from both.

**Architecture:** Keep the existing `financial-reconciliation` application route and introduce a small client-side tab controller backed by `state.financialReconciliation.activeTab`. The Manual and Automatic DOM sections become semantic tab panels; the existing history remains outside them. Manual workspace data loads on entry, automatic rules load only when the Automatic tab is first activated, and the Settings batch-analysis handoff explicitly requests the Automatic entry tab.

**Tech Stack:** Static HTML, vanilla JavaScript, CSS, Node.js built-in test runner, existing Vercel/Supabase APIs unchanged.

## Global Constraints

- Normal navigation into Reconciliation always selects **Manual reconciliation**.
- Settings **Run batch now** is the only entry flow that opens **Automatic reconciliation** directly.
- Shared history remains visible below both tab panels and continues to combine user and automatic reconciliations.
- Switching tabs preserves Manual and Automatic state and does not reload data that is already loaded.
- Automatic configuration and scheduling remain under **Settings → Reconciliation → Automatic reconciliation**.
- The inactive panel uses `hidden`; tabs expose correct ARIA state and support Left and Right Arrow activation.
- No database table, migration, RPC, API, permission, matching, locking, audit, or scheduled-execution change is in scope.
- Preserve all unrelated tracked and untracked workspace files.

## File Map

- Modify `index.html`: add the application tab list, wrap the existing manual and automatic sections in tab panels, and keep history outside both panels.
- Modify `app-main.js`: add tab DOM references, state, rendering, keyboard/click behavior, lazy automatic loading, and the Settings handoff entry option.
- Modify `styles.css`: style the application tabs, focus state, panel spacing, and narrow-screen behavior without changing the existing workbench/proposal layouts.
- Modify `tests/reconciliation-density.test.js`: pin semantic structure, shared-history placement, controller behavior, and responsive CSS.
- Modify `tests/reconciliation-automation-ui.test.js`: pin lazy loading, normal-entry selection, Settings handoff, and retained automatic-run behavior.

---

### Task 1: Add the accessible application tab shell and controller

**Files:**
- Modify: `index.html:3863-3933`
- Modify: `app-main.js:1301-1319`
- Modify: `app-main.js:1938-1970`
- Modify: `app-main.js:2624-2642`
- Modify: `app-main.js:21781-21917`
- Modify: `app-main.js:22457-22465`
- Test: `tests/reconciliation-density.test.js:224-263`

**Interfaces:**
- Consumes: existing `financialReconciliationState()`, `renderFinancialReconciliationAutomation()`, and all existing Manual render functions.
- Produces: `normalizeFinancialReconciliationTab(tab) -> "manual" | "automatic"`, `renderFinancialReconciliationTabs() -> void`, `setFinancialReconciliationTab(tab, options?) -> Promise<void>`, and `onFinancialReconciliationTabKeydown(event) -> void`.

- [ ] **Step 1: Write failing semantic-structure tests**

Replace the old “automatic proposal review sits after filters” structural assertion with tests that require two panels and history outside both:

```js
test("reconciliation separates manual and automatic work into accessible tabs with shared history", () => {
  assert.match(html, /class="[^"]*financial-reconciliation-view-tabs[^"]*"[^>]*role="tablist"/);
  assert.match(html, /id="financial-reconciliation-manual-tab"[^>]*role="tab"[^>]*aria-controls="financial-reconciliation-manual-panel"/);
  assert.match(html, /id="financial-reconciliation-automatic-tab"[^>]*role="tab"[^>]*aria-controls="financial-reconciliation-automatic-panel"/);
  assert.match(html, /id="financial-reconciliation-manual-panel"[^>]*role="tabpanel"[^>]*aria-labelledby="financial-reconciliation-manual-tab"/);
  assert.match(html, /id="financial-reconciliation-automatic-panel"[^>]*role="tabpanel"[^>]*aria-labelledby="financial-reconciliation-automatic-tab"[^>]*hidden/);

  const manualPanel = html.indexOf('id="financial-reconciliation-manual-panel"');
  const automaticPanel = html.indexOf('id="financial-reconciliation-automatic-panel"');
  const history = html.indexOf('class="card financial-reconciliation-history-card"');
  assert.ok(manualPanel >= 0 && automaticPanel > manualPanel && history > automaticPanel);
  assert.match(html.slice(manualPanel, automaticPanel), /id="financial-reconciliation-status"[\s\S]*id="financial-reconciliation-filters"[\s\S]*id="financial-reconciliation-current"/);
  assert.match(html.slice(automaticPanel, history), /id="financial-reconciliation-workbench-automation-rules"[\s\S]*id="financial-reconciliation-workbench-automation-proposals"/);
});
```

Add a controller behavior test that compiles the actual functions from `app-main.js` with fake elements. It must assert Manual defaults, ARIA/tabindex values, `hidden` panels, state preservation, and keyboard movement:

```js
test("reconciliation tab controller defaults to Manual and supports arrow activation", async () => {
  const current = {
    activeTab: "manual",
    filters: { description: "retained" },
    automation: { loaded: true, run: { runId: "retained-run" } },
  };
  const calls = [];
  const els = reconciliationTabElements();
  const controller = compileReconciliationTabController({ current, els, calls });

  controller.renderFinancialReconciliationTabs();
  assert.equal(els.financialReconciliationManualPanel.hidden, false);
  assert.equal(els.financialReconciliationAutomaticPanel.hidden, true);
  assert.equal(els.financialReconciliationManualTab["aria-selected"], "true");

  await controller.setFinancialReconciliationTab("automatic", { focus: true });
  assert.equal(current.activeTab, "automatic");
  assert.equal(current.filters.description, "retained");
  assert.equal(current.automation.run.runId, "retained-run");
  assert.equal(els.financialReconciliationAutomaticPanel.hidden, false);
  assert.equal(els.financialReconciliationAutomaticTab.focused, true);

  controller.onFinancialReconciliationTabKeydown({ key: "ArrowLeft", preventDefault() { calls.push("prevent"); } });
  await Promise.resolve();
  assert.equal(current.activeTab, "manual");
  assert.ok(calls.includes("prevent"));
});
```

The test helpers must use the existing `appFunctionSource()` pattern and real production function bodies. Add these helpers in the test file:

```js
function reconciliationTabClassList() {
  const values = new Set();
  return {
    toggle(name, force) {
      if (force) values.add(name);
      else values.delete(name);
    },
    contains: (name) => values.has(name),
  };
}

function reconciliationTabElements() {
  const tab = () => ({
    classList: reconciliationTabClassList(),
    focused: false,
    setAttribute(name, value) { this[name] = value; },
    focus() { this.focused = true; },
  });
  return {
    financialReconciliationManualTab: tab(),
    financialReconciliationAutomaticTab: tab(),
    financialReconciliationManualPanel: { hidden: false },
    financialReconciliationAutomaticPanel: { hidden: true },
  };
}

function compileReconciliationTabController({ current, els, calls }) {
  let controller;
  controller = new Function(
    "clean",
    "financialReconciliationState",
    "renderFinancialReconciliation",
    "loadFinancialReconciliationAutomationRules",
    "els",
    `${appFunctionSource("normalizeFinancialReconciliationTab")}
     ${appFunctionSource("renderFinancialReconciliationTabs")}
     ${appFunctionSource("setFinancialReconciliationTab")}
     ${appFunctionSource("onFinancialReconciliationTabKeydown")}
     return { normalizeFinancialReconciliationTab, renderFinancialReconciliationTabs, setFinancialReconciliationTab, onFinancialReconciliationTabKeydown };`,
  )(
    (value) => String(value ?? "").trim(),
    () => current,
    () => { calls.push(`render:${current.activeTab}`); controller.renderFinancialReconciliationTabs(); },
    async () => { calls.push("load-rules"); current.automation.loaded = true; },
    els,
  );
  return controller;
}
```

- [ ] **Step 2: Run the focused test and verify RED**

Run:

```powershell
node --test --test-isolation=none --test-name-pattern="reconciliation separates|reconciliation tab controller" tests/reconciliation-density.test.js
```

Expected: FAIL because the application tab markup and controller functions do not exist.

- [ ] **Step 3: Add the semantic tab markup**

In `index.html`, add this tab list after the feature title:

```html
<div class="actions financial-reconciliation-view-tabs" role="tablist" aria-label="Reconciliation mode" aria-orientation="horizontal">
  <button id="financial-reconciliation-manual-tab" type="button" class="active-tab" role="tab" tabindex="0" aria-selected="true" aria-controls="financial-reconciliation-manual-panel">Manual reconciliation</button>
  <button id="financial-reconciliation-automatic-tab" type="button" class="ghost" role="tab" tabindex="-1" aria-selected="false" aria-controls="financial-reconciliation-automatic-panel">Automatic reconciliation</button>
</div>
```

Insert this opening tag immediately before `<section class="card financial-reconciliation-workbench-card">`, then insert its closing `</div>` immediately after the closing tag of `.financial-reconciliation-workbench`:

```html
<div id="financial-reconciliation-manual-panel" role="tabpanel" aria-labelledby="financial-reconciliation-manual-tab">
```

Insert this opening tag immediately before `#financial-reconciliation-workbench-automation`, then insert its closing `</div>` immediately after that section:

```html
<div id="financial-reconciliation-automatic-panel" role="tabpanel" aria-labelledby="financial-reconciliation-automatic-tab" tabindex="0" hidden>
```

Leave `.financial-reconciliation-history-card` after both closing panel tags. Do not duplicate the history table or its tbody ID.

Move the existing `<p id="financial-reconciliation-status" class="auth-status"></p>` from the page title into the beginning of `#financial-reconciliation-manual-panel`, before the workbench card, and add `role="status" aria-live="polite"`. The existing automatic status remains inside the Automatic panel.

- [ ] **Step 4: Add state and DOM references**

Add `activeTab: "manual"` beside the other `state.financialReconciliation` fields. Add these entries to `els`:

```js
financialReconciliationManualTab: document.getElementById("financial-reconciliation-manual-tab"),
financialReconciliationAutomaticTab: document.getElementById("financial-reconciliation-automatic-tab"),
financialReconciliationManualPanel: document.getElementById("financial-reconciliation-manual-panel"),
financialReconciliationAutomaticPanel: document.getElementById("financial-reconciliation-automatic-panel"),
```

- [ ] **Step 5: Implement the minimal tab controller**

Place the controller beside `financialReconciliationState()`:

```js
function normalizeFinancialReconciliationTab(tab) {
  return clean(tab) === "automatic" ? "automatic" : "manual";
}

function renderFinancialReconciliationTabs() {
  const automatic = normalizeFinancialReconciliationTab(financialReconciliationState().activeTab) === "automatic";
  financialReconciliationState().activeTab = automatic ? "automatic" : "manual";
  els.financialReconciliationManualTab?.classList.toggle("active-tab", !automatic);
  els.financialReconciliationManualTab?.classList.toggle("ghost", automatic);
  els.financialReconciliationAutomaticTab?.classList.toggle("active-tab", automatic);
  els.financialReconciliationAutomaticTab?.classList.toggle("ghost", !automatic);
  els.financialReconciliationManualTab?.setAttribute("aria-selected", String(!automatic));
  els.financialReconciliationAutomaticTab?.setAttribute("aria-selected", String(automatic));
  els.financialReconciliationManualTab?.setAttribute("tabindex", automatic ? "-1" : "0");
  els.financialReconciliationAutomaticTab?.setAttribute("tabindex", automatic ? "0" : "-1");
  if (els.financialReconciliationManualPanel) els.financialReconciliationManualPanel.hidden = automatic;
  if (els.financialReconciliationAutomaticPanel) els.financialReconciliationAutomaticPanel.hidden = !automatic;
}

async function setFinancialReconciliationTab(tab, { focus = false } = {}) {
  const next = normalizeFinancialReconciliationTab(tab);
  financialReconciliationState().activeTab = next;
  renderFinancialReconciliation();
  if (next === "automatic" && !financialReconciliationState().automation.loaded) {
    await loadFinancialReconciliationAutomationRules();
  }
  if (focus) (next === "automatic" ? els.financialReconciliationAutomaticTab : els.financialReconciliationManualTab)?.focus();
}

function onFinancialReconciliationTabKeydown(event) {
  if (!event || !["ArrowLeft", "ArrowRight"].includes(event.key)) return;
  event.preventDefault();
  const next = financialReconciliationState().activeTab === "automatic" ? "manual" : "automatic";
  void setFinancialReconciliationTab(next, { focus: true });
}
```

Update `renderFinancialReconciliation()` so it always renders the tab state and shared history, but renders only the active panel’s content:

```js
function renderFinancialReconciliation() {
  if (!canAppFinancialReconciliation()) return;
  renderFinancialReconciliationTabs();
  if (financialReconciliationState().activeTab === "automatic") {
    renderFinancialReconciliationAutomation();
  } else {
    renderFinancialReconciliationSourceControls();
    renderFinancialReconciliationFilters();
    renderFinancialReconciliationCandidates();
    renderFinancialReconciliationCurrent();
  }
  renderFinancialReconciliationHistory();
}
```

- [ ] **Step 6: Bind click and keyboard events**

Add these bindings in `bindEvents()`:

```js
els.financialReconciliationManualTab?.addEventListener("click", () => setFinancialReconciliationTab("manual"));
els.financialReconciliationAutomaticTab?.addEventListener("click", () => setFinancialReconciliationTab("automatic"));
els.financialReconciliationManualTab?.addEventListener("keydown", onFinancialReconciliationTabKeydown);
els.financialReconciliationAutomaticTab?.addEventListener("keydown", onFinancialReconciliationTabKeydown);
```

- [ ] **Step 7: Run focused tests and verify GREEN**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-density.test.js
```

Expected: all reconciliation-density tests PASS, including semantic structure and real controller behavior.

- [ ] **Step 8: Commit Task 1**

```powershell
git add -- index.html app-main.js tests/reconciliation-density.test.js
git commit -m "feat: split reconciliation into manual and automatic tabs"
```

---

### Task 2: Make loading tab-aware and hand Settings runs to Automatic

**Files:**
- Modify: `app-main.js:3600-3730`
- Modify: `app-main.js:21717-21777`
- Modify: `app-main.js:21913-21942`
- Test: `tests/reconciliation-automation-ui.test.js:402-505`
- Test: `tests/reconciliation-automation-ui.test.js:637-740`

**Interfaces:**
- Consumes: Task 1 `normalizeFinancialReconciliationTab()`, `setFinancialReconciliationTab()`, and `state.financialReconciliation.activeTab`.
- Produces: `financialReconciliationEntryTab(options) -> "manual" | "automatic"`; `setView(view, options = {})` recognizes `{ financialReconciliationTab: "automatic" }`; `ensureFinancialReconciliationData()` loads only data required by the active tab.

- [ ] **Step 1: Write failing entry and lazy-loading tests**

Add a pure entry-option test:

```js
test("Reconciliation entry defaults to Manual and accepts only an explicit Automatic handoff", () => {
  const entryTab = new Function(
    `${appFunctionSource("clean")}\n${appFunctionSource("financialReconciliationEntryTab")}\nreturn financialReconciliationEntryTab;`,
  )();
  assert.equal(entryTab(), "manual");
  assert.equal(entryTab({}), "manual");
  assert.equal(entryTab({ financialReconciliationTab: "manual" }), "manual");
  assert.equal(entryTab({ financialReconciliationTab: "automatic" }), "automatic");
  assert.equal(entryTab({ financialReconciliationTab: "AUTOMATIC" }), "manual");
});
```

Add an executable `ensureFinancialReconciliationData()` test using injected loaders:

```js
test("Reconciliation loads Manual first and lazily loads Automatic rules once", async () => {
  const calls = [];
  const current = { loaded: false, activeTab: "manual", automation: { loaded: false } };
  const ensure = compileEnsureFinancialReconciliationData({
    current,
    loadWorkspace: async () => { calls.push("workspace"); current.loaded = true; },
    loadRules: async () => { calls.push("rules"); current.automation.loaded = true; },
    render: () => calls.push("render"),
  });

  await ensure();
  assert.deepEqual(calls, ["workspace", "render"]);

  calls.length = 0;
  current.activeTab = "automatic";
  await ensure();
  assert.deepEqual(calls, ["rules", "render"]);

  calls.length = 0;
  await ensure();
  assert.deepEqual(calls, ["render"]);
});
```

Add this compiler immediately above that test so it executes the production function rather than a copy:

```js
function compileEnsureFinancialReconciliationData({ current, loadWorkspace, loadRules, render }) {
  return new Function(
    "financialReconciliationState",
    "loadFinancialReconciliationWorkspace",
    "loadFinancialReconciliationAutomationRules",
    "renderFinancialReconciliation",
    `${appFunctionSource("ensureFinancialReconciliationData")}
     return ensureFinancialReconciliationData;`,
  )(
    () => current,
    loadWorkspace,
    loadRules,
    render,
  );
}
```

Update the existing **Run batch now** test so its `setView` stub records both arguments and expects:

```js
assert.deepEqual(sequence, [
  "api",
  "view:financial-reconciliation:automatic",
]);
assert.equal(current.automation.run, run);
assert.deepEqual([...current.automation.selectedProposalIds], ["proposal-1"]);
```

The stub should record `options?.financialReconciliationTab || "manual"`. Remove the old expectation that `runReconciliationAutomationBatchNow()` directly calls `renderFinancialReconciliation()` after `setView()`; `setView()` owns rendering and data loading.

- [ ] **Step 2: Run focused tests and verify RED**

Run:

```powershell
node --test --test-isolation=none --test-name-pattern="entry defaults|loads Manual first|Run batch now stores" tests/reconciliation-automation-ui.test.js
```

Expected: FAIL because entry options do not exist, Manual entry still loads automatic rules, and Settings does not request the Automatic tab.

- [ ] **Step 3: Implement the explicit entry option**

Add beside the tab normalizer:

```js
function financialReconciliationEntryTab(options = {}) {
  return options && options.financialReconciliationTab === "automatic" ? "automatic" : "manual";
}
```

Change the view signature from `async function setView(view)` to `async function setView(view, options = {})`. Immediately after the current financial-reconciliation authorization guard, insert:

```js
if (view === "financial-reconciliation") {
  state.financialReconciliation.activeTab = financialReconciliationEntryTab(options);
}
```

Every existing call that passes only `view` therefore opens Manual. Do not persist this option in the URL or local storage.

Immediately after the existing `await ensureCurrentViewData();` at the end of `setView()`, focus the Automatic tab only for the explicit handoff:

```js
if (view === "financial-reconciliation" && financialReconciliationEntryTab(options) === "automatic") {
  els.financialReconciliationAutomaticTab?.focus();
}
```

- [ ] **Step 4: Make initial data loading tab-aware**

Replace the current eager automation load in `ensureFinancialReconciliationData()`:

```js
async function ensureFinancialReconciliationData() {
  const current = financialReconciliationState();
  if (!current.loaded) await loadFinancialReconciliationWorkspace({ silent: true });
  if (current.activeTab === "automatic" && !current.automation.loaded) {
    await loadFinancialReconciliationAutomationRules();
  }
  renderFinancialReconciliation();
}
```

The existing guards in `loadFinancialReconciliationAutomationRules()` remain authoritative for avoiding duplicate requests.

- [ ] **Step 5: Update Settings Run batch now handoff**

After storing the returned run and default selected proposal IDs, navigate with the explicit option:

```js
await setView("financial-reconciliation", { financialReconciliationTab: "automatic" });
```

Remove the following direct `renderFinancialReconciliation()` call. Keep `current.automation.loaded = false` so the authoritative rules catalog is refreshed when the Automatic panel opens; do not clear `current.automation.run` or `selectedProposalIds` during that catalog request.

- [ ] **Step 6: Keep failures scoped without clearing the other panel**

Pin the existing behavior in the lazy-loading test harness and production code:

```js
} catch (error) {
  automation.rules = [];
  automation.loaded = false;
  setFinancialReconciliationAutomationStatus(`Failed to load automatic rules: ${error.message}`, "error");
} finally {
  automation.pendingAction = "";
  renderFinancialReconciliationAutomation();
}
```

Do not modify `current.filters`, `current.workspace`, `current.completionCommentDraft`, `automation.run`, or `automation.selectedProposalIds` in this failure path. Add assertions to the failed-rule-reload test confirming its retained `run` and selection Set are unchanged.

Use object identity so the test detects replacement as well as data loss:

```js
const retainedRun = current.automation.run;
const retainedSelections = current.automation.selectedProposalIds;
await loadRules();
assert.strictEqual(current.automation.run, retainedRun);
assert.strictEqual(current.automation.selectedProposalIds, retainedSelections);
```

- [ ] **Step 7: Run focused tests and verify GREEN**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js
node --test --test-isolation=none tests/reconciliation-density.test.js
```

Expected: both focused files PASS; the Settings handoff opens Automatic and no existing execution/proposal behavior regresses.

- [ ] **Step 8: Commit Task 2**

```powershell
git add -- app-main.js tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
git commit -m "feat: load reconciliation tabs on demand"
```

---

### Task 3: Add responsive tab styling and complete regression verification

**Files:**
- Modify: `styles.css:6430-6460`
- Modify: `styles.css:6870-6905`
- Test: `tests/reconciliation-density.test.js:241-263`
- Test: `tests/reconciliation-automation-ui.test.js`

**Interfaces:**
- Consumes: Task 1 `.financial-reconciliation-view-tabs` markup and existing `.actions .active-tab` theme.
- Produces: responsive, focus-visible application tabs that preserve existing Manual and Automatic internal layouts.

- [ ] **Step 1: Write failing responsive and placement tests**

Add CSS source-contract assertions:

```js
test("reconciliation application tabs remain visible and usable on narrow screens", () => {
  assert.match(css, /\.financial-reconciliation-view-tabs\s*\{[\s\S]*flex-wrap:\s*wrap;[\s\S]*margin-bottom:\s*1rem;/);
  assert.match(css, /\.financial-reconciliation-view-tabs\s*>\s*button:focus-visible\s*\{[\s\S]*outline:/);
  assert.match(css, /\.financial-reconciliation-manual-panel:not\(\[hidden\]\),[\s\S]*\.financial-reconciliation-automatic-panel:not\(\[hidden\]\)\s*\{[\s\S]*display:\s*grid;/);
  assert.match(css, /@media \(max-width:\s*768px\)[\s\S]*\.financial-reconciliation-view-tabs\s*>\s*button\s*\{[\s\S]*flex:\s*1 1 12rem;/);
});
```

Add a static guard that the shared history ID appears exactly once and neither tab panel is nested inside it:

```js
assert.equal((html.match(/id="financial-reconciliation-history-rows"/g) || []).length, 1);
assert.ok(html.indexOf('id="financial-reconciliation-history-rows"') > html.indexOf('id="financial-reconciliation-automatic-panel"'));
```

- [ ] **Step 2: Run the focused test and verify RED**

Run:

```powershell
node --test --test-isolation=none --test-name-pattern="application tabs remain|separates manual and automatic" tests/reconciliation-density.test.js
```

Expected: FAIL because the new tab-specific responsive CSS is absent.

- [ ] **Step 3: Add minimal desktop and narrow-screen CSS**

Place these rules before the existing reconciliation workbench rules:

```css
.financial-reconciliation-view-tabs {
  display: flex;
  flex-wrap: wrap;
  gap: .5rem;
  margin-bottom: 1rem;
}

.financial-reconciliation-view-tabs > button {
  min-width: 12rem;
}

.financial-reconciliation-view-tabs > button:focus-visible {
  outline: 3px solid color-mix(in srgb, var(--brand) 35%, transparent);
  outline-offset: 2px;
}

.financial-reconciliation-manual-panel:not([hidden]),
.financial-reconciliation-automatic-panel:not([hidden]) {
  display: grid;
  gap: 1rem;
}
```

Add the narrow-screen rule inside the existing `@media (max-width: 768px)` reconciliation block:

```css
.financial-reconciliation-view-tabs > button {
  flex: 1 1 12rem;
}
```

Add `class="financial-reconciliation-manual-panel"` and `class="financial-reconciliation-automatic-panel"` to the corresponding panel containers. Keep `:not([hidden])` on the layout selector so author CSS never overrides the browser’s hidden-panel behavior.

- [ ] **Step 4: Run syntax and focused verification**

Run:

```powershell
node --check app-main.js
node --test --test-isolation=none tests/reconciliation-density.test.js
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js
```

Expected: JavaScript syntax PASS; both focused test files PASS with zero failures.

- [ ] **Step 5: Run the complete automated suite and diff checks**

Run:

```powershell
node --test --test-isolation=none tests/*.test.js
git diff --check
git status --short
```

Expected: all tests PASS; `git diff --check` exits 0; status lists only the five scoped product/test files plus pre-existing unrelated untracked files.

- [ ] **Step 6: Perform browser verification**

Start the configured local application runtime:

```powershell
vercel dev --listen 53841
```

Using an authenticated test session, verify at desktop and at a viewport no wider than 768 px:

1. Normal Reconciliation navigation opens Manual.
2. Manual filters and current reconciliation remain unchanged after Automatic → Manual.
3. The first Automatic activation loads rules; subsequent switching does not issue another rules request.
4. Left and Right Arrow keys move between the two tabs and preserve visible focus.
5. History stays visible and usable under both panels.
6. Settings **Run batch now** opens Automatic with the returned proposals and selections.
7. Automatic analysis/execution errors remain in Automatic; Manual work remains intact.
8. Tabs wrap without horizontal clipping on the narrow viewport.

If `vercel` or an authenticated fixture is unavailable, record that exact external limitation in the implementation report and do not claim browser verification passed.

- [ ] **Step 7: Commit Task 3**

```powershell
git add -- index.html app-main.js styles.css tests/reconciliation-density.test.js tests/reconciliation-automation-ui.test.js
git commit -m "style: finish reconciliation tab layout"
```

---

## Final Review Checklist

- [ ] Re-read `docs/superpowers/specs/2026-08-15-reconciliation-manual-automatic-tabs-design.md` and map every acceptance criterion to a passing test or completed browser check.
- [ ] In the isolated implementation worktree, run `$baseCommit = git merge-base HEAD main`. Confirm `git diff --name-only "$baseCommit..HEAD"` contains only `index.html`, `app-main.js`, `styles.css`, `tests/reconciliation-density.test.js`, and `tests/reconciliation-automation-ui.test.js`.
- [ ] Confirm no Supabase migration, API handler, deployment configuration, or unrelated user file changed.
- [ ] Run `node --test --test-isolation=none tests/*.test.js` once more from the final committed tree and report the exact pass/fail counts.
- [ ] Run `git diff --check "$baseCommit..HEAD"` and confirm exit code 0.
- [ ] Request an independent code/spec review before merging or publishing.
