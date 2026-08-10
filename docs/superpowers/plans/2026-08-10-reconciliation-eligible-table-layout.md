# Reconciliation Eligible Records Table Layout Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the Reconciliation Eligible records table wrap long text cleanly, prevent Start/Add actions from overlapping the Date column, and compact the Start actions and row status pills.

**Architecture:** Update the reconciliation renderer to emit semantic classes for the action, date, description, optional-detail, amount, and status cells. Use those hooks for table-column sizing and controlled wrapping in `styles.css`, then protect the rendered markup and CSS contract with the existing Node source-contract test.

**Tech Stack:** Vanilla JavaScript, static HTML/CSS, Node.js built-in test runner.

## Global Constraints

- Apply only to the Eligible records table and its header action in the Reconciliation view.
- Do not change reconciliation data, filtering, selection, locking, or completion behaviour.
- Do not alter the Current reconciliation card or reconciliation history table.
- Give the first action column a fixed `4.5rem` width; keep the Start/Add button inside it.
- Let description and optional source-detail cells wrap, including long unbroken strings; keep Date and Amount on one line.
- Give Date a `6rem` width, Amount a `5.8rem` width, and Status a `6.8rem` width.
- Remove the existing `1%` first-column sizing rule.
- Set the header “Choose Start…” button to `0.84rem`; row Start/Add actions to `0.72rem` on desktop and `0.70rem` at `768px` and below.
- Set Eligible records status pills to `0.70rem` with smaller padding; keep their labels on one line.
- Preserve the existing mobile-safe `16px` filter input/select text size.

---

### Task 1: Render semantic cells and scope the Eligible records layout

**Files:**
- Modify: `app-main.js:21320-21353`
- Modify: `styles.css:6237-6320`
- Modify: `tests/reconciliation-density.test.js:1-23`

**Interfaces:**
- Consumes: `workspace.candidates`, `extraColumns`, `financialReconciliationStatusMarkup(status)`, and the existing `#financial-reconciliation-start` header action.
- Produces: `.financial-reconciliation-action`, `.financial-reconciliation-date`, `.financial-reconciliation-detail`, `.financial-reconciliation-amount`, and `.financial-reconciliation-status-cell` hooks in generated Eligible records table rows and headers.

- [ ] **Step 1: Extend the existing source-contract test with failing layout assertions**

Add these assertions to `tests/reconciliation-density.test.js` after the existing HTML/CSS reads, adding an `appMain` string first:

```js
const appMain = fs.readFileSync(path.join(root, "app-main.js"), "utf8");

assert.match(appMain, /<th class="financial-reconciliation-action"><\/th><th class="financial-reconciliation-date">Date<\/th>/);
assert.match(appMain, /<td class="financial-reconciliation-action"><button/);
assert.match(appMain, /<td class="financial-reconciliation-date">\$\{escape\(formatDateOnly\(row\.source_date\) \|\| "-"\)\}<\/td>/);
assert.match(appMain, /<td class="financial-reconciliation-detail">/);
assert.match(appMain, /<td class="financial-reconciliation-amount">/);
assert.match(appMain, /<td class="financial-reconciliation-status-cell">/);
assert.match(css, /\.financial-reconciliation-action\s*\{\s*width:\s*4\.5rem;/);
assert.match(css, /\.financial-reconciliation-date\s*\{\s*width:\s*6rem;/);
assert.match(css, /\.financial-reconciliation-amount\s*\{\s*width:\s*5\.8rem;/);
assert.match(css, /\.financial-reconciliation-status-cell\s*\{\s*width:\s*6\.8rem;/);
assert.match(css, /overflow-wrap:\s*anywhere;/);
assert.match(css, /#financial-reconciliation-start\s*\{\s*font-size:\s*\.84rem;/);
assert.match(css, /\.financial-reconciliation-table button\s*\{\s*font-size:\s*\.72rem;/);
assert.match(css, /\.financial-reconciliation-table \.financial-reconciliation-status\s*\{\s*font-size:\s*\.70rem;/);
```

- [ ] **Step 2: Run the focused test to verify it fails**

Run: `node --test tests/reconciliation-density.test.js`

Expected: FAIL because the renderer does not yet emit the semantic cell classes and the layout rules do not yet exist.

- [ ] **Step 3: Render explicit Eligible records headers and cells**

In `renderFinancialReconciliationCandidates`, replace the generic `columns` header rendering with semantic headers and retain dynamic optional columns:

```js
const optionalHeaders = extraColumns.map((column) => `<th class="financial-reconciliation-detail">${escape(column.label)}</th>`).join("");
els.financialReconciliationTableHead.innerHTML = `<tr><th class="financial-reconciliation-action"></th><th class="financial-reconciliation-date">Date</th><th class="financial-reconciliation-description">Description</th>${optionalHeaders}<th class="financial-reconciliation-amount">Amount</th><th class="financial-reconciliation-status-cell">Status</th></tr>`;
```

Render the matching row cells using the same classes:

```js
const optional = extraColumns.map(({ key }) => `<td class="financial-reconciliation-detail">${escape(clean(row[key]) || "-")}</td>`);
return `<tr><td class="financial-reconciliation-action"><button type="button" class="ghost" data-financial-reconciliation-row-action="${reconciliation ? "add" : "start"}" data-source-id="${escape(row.id)}" ${disabled}>${actionLabel}</button></td><td class="financial-reconciliation-date">${escape(formatDateOnly(row.source_date) || "-")}</td><td class="financial-reconciliation-description">${escape(clean(row.description) || "-")}</td>${optional.join("")}<td class="financial-reconciliation-amount">${escape(formatMoney(Number(row.amount || 0)))}</td><td class="financial-reconciliation-status-cell">${financialReconciliationStatusMarkup("not-started")}</td></tr>`;
```

Set the empty-row `colspan` to `extraColumns.length + 5` so it still spans action, date, description, optional columns, amount, and status.

- [ ] **Step 4: Add the scoped table rules**

Replace the `1%` first-column rule with these rules in the reconciliation CSS block:

```css
.financial-reconciliation-action { width: 4.5rem; white-space: nowrap; }
.financial-reconciliation-date { width: 6rem; white-space: nowrap; }
.financial-reconciliation-amount { width: 5.8rem; white-space: nowrap; }
.financial-reconciliation-status-cell { width: 6.8rem; white-space: nowrap; }

.financial-reconciliation-table td.financial-reconciliation-description,
.financial-reconciliation-table td.financial-reconciliation-detail {
  min-width: 0;
  white-space: normal;
  overflow-wrap: anywhere;
}

.financial-reconciliation-eligible-card #financial-reconciliation-start {
  font-size: .84rem;
  padding: .42rem .58rem;
}

.financial-reconciliation-table button { font-size: .72rem; }

.financial-reconciliation-table .financial-reconciliation-status {
  font-size: .70rem;
  padding: .32rem .48rem;
  white-space: nowrap;
}

@media (max-width: 768px) {
  .financial-reconciliation-table button { font-size: .70rem; }
}
```

Retain the existing `16px` mobile filter rule. Do not apply the status-pill rule outside `.financial-reconciliation-table`.

- [ ] **Step 5: Run the focused and complete test suite**

Run: `node --test tests/*.test.js`

Expected: all tests PASS.

- [ ] **Step 6: Inspect whitespace and commit the task**

Run: `git diff --check`

Expected: no whitespace errors.

```bash
git add app-main.js styles.css tests/reconciliation-density.test.js
git commit -m "style: wrap reconciliation eligible records"
```
