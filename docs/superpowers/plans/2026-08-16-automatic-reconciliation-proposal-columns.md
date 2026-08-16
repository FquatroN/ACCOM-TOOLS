# Compact Automatic Reconciliation Proposal Columns Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Render each automatic reconciliation proposal as one compact proposal/base/destination row with proposal metadata in the first column and no separate header or footer bands.

**Architecture:** Keep the existing proposal objects, visibility rules, selection state, and record renderer. Refactor only `financialReconciliationAutomationProposalMarkup` to emit a proposal-information column beside the existing record columns, then update the scoped CSS grid and responsive breakpoint. Tests exercise the real extracted renderer and source CSS contracts before production changes.

**Tech Stack:** Vanilla JavaScript, semantic HTML, CSS Grid/Flexbox, Node.js built-in test runner.

## Global Constraints

- This is presentation-only: do not change matching logic, APIs, persistence, reconciliation calculations, or visible-status selection.
- Render a checkbox only for executable proposals whose normalized status is exactly `proposed`.
- Render no checkbox for `ambiguous`, `completed`, `stale`, or `failed` proposals.
- Preserve proposal focus behavior, executable checkbox data attributes and accessible names, evidence, reason and execution-attempt text, record IDs, difference, allowed difference, rule name, and rule version.
- Below 700 px, render proposal metadata as a compact top strip and stack record content vertically.

---

### Task 1: Replace Proposal Bands with a Leading Metadata Column

**Files:**
- Modify: `app-main.js:22112-22149`
- Modify: `styles.css:6804-6970, 7067-7085`
- Test: `tests/reconciliation-automation-ui.test.js:1010-1085`
- Test: `tests/reconciliation-density.test.js:431-434`

**Interfaces:**
- Consumes: `financialReconciliationAutomationProposalMarkup(proposal, run, rules, selectedProposalIds, pending)` and `financialReconciliationAutomationItemMarkup(item, label, operator)`.
- Produces: proposal markup containing `.financial-reconciliation-automation-proposal-meta` followed by `.financial-reconciliation-automation-proposal-records`; executable proposal inputs retain `data-financial-reconciliation-automation-proposal-id`.

- [ ] **Step 1: Write failing behavior tests for the metadata column and conditional selection control**

Update the existing proposal-markup test in `tests/reconciliation-automation-ui.test.js` so the proposed case requires the new leading column and rejects the old bands:

```js
assert.match(proposedMarkup, /financial-reconciliation-automation-proposal-meta/);
assert.match(proposedMarkup, /type="checkbox"[^>]*checked/);
assert.match(proposedMarkup, /financial-reconciliation-automation-proposal-status[^>]*>proposed</);
assert.match(proposedMarkup, /Difference 1\.00/);
assert.match(proposedMarkup, /Allowed 1\.00/);
assert.match(proposedMarkup, /Manual enabled[\s\S]*version 3/i);
assert.doesNotMatch(proposedMarkup, /<header>/);
assert.doesNotMatch(proposedMarkup, /<footer>/);
```

Replace the ambiguous disabled-checkbox assertion with:

```js
assert.doesNotMatch(ambiguousMarkup, /type="checkbox"/);
assert.match(ambiguousMarkup, /financial-reconciliation-automation-proposal-status[^>]*>ambiguous</);
```

Add lifecycle cases and assert none renders a checkbox:

```js
for (const status of ["completed", "stale", "failed"]) {
  const lifecycleMarkup = proposalMarkup({ ...proposed, status }, run, workbenchRules(), new Set(), false);
  assert.doesNotMatch(lifecycleMarkup, /type="checkbox"/);
  assert.match(lifecycleMarkup, new RegExp(`financial-reconciliation-automation-proposal-status[^>]*>${status}<`));
}
```

- [ ] **Step 2: Write failing CSS structure tests**

Replace the existing pair-only density test in `tests/reconciliation-density.test.js` with:

```js
test("automatic proposals use a metadata-first desktop row and narrow stack", () => {
  assert.match(css, /\.financial-reconciliation-automation-proposal\s*\{[\s\S]*grid-template-columns:\s*minmax\([^;]+\)\s+minmax\(0,\s*1fr\)\s+minmax\(0,\s*1fr\)/);
  assert.match(css, /\.financial-reconciliation-automation-proposal-meta\s*\{/);
  assert.match(css, /@media\s*\(max-width:\s*700px\)[\s\S]*\.financial-reconciliation-automation-proposal\s*\{[\s\S]*grid-template-columns:\s*minmax\(0,\s*1fr\)/);
});
```

- [ ] **Step 3: Run the focused tests and capture RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
```

Expected: FAIL because the renderer still emits `<header>` and `<footer>`, ambiguous/lifecycle proposals still contain disabled checkboxes, and the proposal itself is not a three-column grid.

- [ ] **Step 4: Implement the leading proposal metadata column**

In `financialReconciliationAutomationProposalMarkup`, construct the checkbox only when `executable` is true:

```js
const selectionMarkup = !executable ? "" : `<label class="financial-reconciliation-automation-proposal-selection"><input type="checkbox" aria-label="${escape(accessibleName)}" data-financial-reconciliation-automation-proposal-id="${escape(clean(value.id))}" ${selected ? "checked" : ""} ${pending ? "disabled" : ""} /><span>Execute proposal</span></label>`;
```

Replace the proposal `<header>` and `<footer>` with one leading column:

```js
<aside class="financial-reconciliation-automation-proposal-meta">
  ${selectionMarkup}
  <span class="financial-reconciliation-automation-proposal-status">${escape(status || "unknown")}</span>
  ${reason}${executionOutcomeMarkup}
  <strong>Difference ${escape(formatMoney(Number(value.calculatedDifference || 0)))}</strong>
  <span>Allowed ${escape(formatMoney(Number(value.allowedDifference ?? definition.differenceAllowed ?? 0)))}</span>
  <span>${escape(clean(rule.displayName) || clean(value.ruleKey))} &middot; version ${escape(Number(value.ruleVersion) || 1)}</span>
</aside>
<div class="financial-reconciliation-automation-proposal-records">...</div>
```

Keep the proposal article status class and `tabindex="-1"`. Keep base, destination, candidate-group, evidence, and record-ID markup unchanged.

- [ ] **Step 5: Implement the desktop three-column and narrow stacked CSS**

Make `.financial-reconciliation-automation-proposal` the outer grid:

```css
.financial-reconciliation-automation-proposal {
  display: grid;
  grid-template-columns: minmax(8.5rem, .55fr) minmax(0, 1fr) minmax(0, 1fr);
  gap: 0;
  padding: 0;
}

.financial-reconciliation-automation-proposal-meta {
  display: grid;
  align-content: start;
  gap: .3rem;
  min-width: 0;
  padding: .55rem;
  border-right: 1px solid var(--border);
  background: var(--surface-soft, rgba(255,255,255,.55));
  font-size: .7rem;
}

.financial-reconciliation-automation-proposal-records {
  display: grid;
  grid-column: 2 / -1;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 0;
}
```

Remove obsolete direct-child proposal header/footer rules. Keep `.financial-reconciliation-automation-item` compact, but remove its outer border/radius so the two record columns read as cells; add a left border between adjacent record cells. Because the records wrapper spans outer columns 2 and 3, additional destinations flow inside its two-column subgrid without entering the proposal metadata column. Candidate groups retain `grid-column: 1 / -1` inside that wrapper.

At the existing `max-width: 700px` breakpoint, install:

```css
.financial-reconciliation-automation-proposal {
  grid-template-columns: minmax(0, 1fr);
}

.financial-reconciliation-automation-proposal-meta {
  grid-template-columns: repeat(2, minmax(0, 1fr));
  border-right: 0;
  border-bottom: 1px solid var(--border);
}

.financial-reconciliation-automation-proposal-records {
  grid-column: 1;
  grid-template-columns: minmax(0, 1fr);
}
```

- [ ] **Step 6: Run focused GREEN verification**

Run:

```powershell
node --check app-main.js
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
```

Expected: JavaScript syntax passes and all focused tests pass with zero failures.

- [ ] **Step 7: Run the complete regression suite**

Run:

```powershell
node --test --test-isolation=none tests/*.test.js
git diff --check
```

Expected: all Node tests pass, and the diff check reports no whitespace errors.

- [ ] **Step 8: Commit the implementation**

```powershell
git add -- app-main.js styles.css tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js docs/superpowers/plans/2026-08-16-automatic-reconciliation-proposal-columns.md
git commit -m "style: compact reconciliation proposal columns"
```

## Manual verification after deployment

1. Open **Reconciliation → Automatic reconciliation** and analyze an enabled rule.
2. Confirm each proposed match is one metadata/base/destination row and only proposed matches show checkboxes.
3. Confirm ambiguous proposals show no checkbox and retain their reason and candidates.
4. Open a finished run and confirm completed, stale, and failed results show no checkbox.
5. Confirm difference, allowed difference, rule/version, evidence, descriptions, and record IDs remain available.
6. At a viewport below 700 px, confirm metadata appears first and the record cells stack without horizontal overflow.
