# Automatic Reconciliation Proposal Dividers Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Render automatic-reconciliation proposals as compact open rows with clear horizontal and vertical separators while retaining the three-column desktop grid and stacked narrow layout.

**Architecture:** Keep the existing semantic proposal markup and data flow unchanged. Implement the approved Option C entirely in the scoped proposal CSS, with executable source-contract tests pinning the desktop grid, open-row borders, column dividers, and mobile section separators.

**Tech Stack:** Vanilla JavaScript application, CSS Grid, Node.js built-in test runner.

## Global Constraints

- Preserve the desktop column order: proposal status/action, base record, destination records.
- Preserve all proposal selection, analysis, execution, audit, keyboard, and lifecycle behavior.
- Preserve semantic proposal `<article>` elements, checkbox labels, candidate groups, detail ordering, and amount alignment.
- At the existing `700px` breakpoint, retain the single-column stack and convert desktop vertical dividers into horizontal section separators.
- Do not modify APIs, reconciliation rules, matching logic, database migrations, history, or manual reconciliation.

---

### Task 1: Render proposals as divided open rows

**Files:**
- Modify: `styles.css:6793-6885,7086-7114`
- Modify: `tests/reconciliation-density.test.js:437-441`
- Verify: `tests/reconciliation-automation-ui.test.js`

**Interfaces:**
- Consumes: existing `.financial-reconciliation-workbench-automation-proposals`, `.financial-reconciliation-automation-proposal`, `.financial-reconciliation-automation-proposal-meta`, `.financial-reconciliation-automation-proposal-records`, and `.financial-reconciliation-automation-item` markup classes.
- Produces: presentation-only CSS contracts; no JavaScript or data interface changes.

- [ ] **Step 1: Strengthen the failing layout contract test**

Replace the existing density test with assertions that require the approved open-row treatment while preserving the existing grid:

```js
test("automatic proposals use three divided open columns and a narrow stack", () => {
  assert.match(css, /\.financial-reconciliation-workbench-automation-proposals\s*\{[\s\S]*gap:\s*0;[\s\S]*border-top:\s*1px solid var\(--border\)/);
  assert.match(css, /\.financial-reconciliation-automation-proposal\s*\{[\s\S]*grid-template-columns:\s*minmax\([^;]+\)\s+minmax\(0,\s*1fr\)\s+minmax\(0,\s*1fr\)[\s\S]*border:\s*0;[\s\S]*border-bottom:\s*1px solid var\(--border\)[\s\S]*border-left:\s*4px solid var\(--brand\)[\s\S]*border-radius:\s*0/);
  assert.match(css, /\.financial-reconciliation-automation-proposal-meta\s*\{[\s\S]*border-right:\s*1px solid var\(--border\)[\s\S]*background:\s*transparent/);
  assert.match(css, /\.financial-reconciliation-automation-proposal-records\s*>\s*\.financial-reconciliation-automation-item\s*\+\s*\.financial-reconciliation-automation-item\s*\{[\s\S]*border-left:\s*1px solid var\(--border\)/);
  assert.match(css, /@media\s*\(max-width:\s*700px\)[\s\S]*\.financial-reconciliation-automation-proposal\s*\{[\s\S]*grid-template-columns:\s*minmax\(0,\s*1fr\)/);
  assert.match(css, /@media\s*\(max-width:\s*700px\)[\s\S]*\.financial-reconciliation-automation-proposal-meta\s*\{[\s\S]*border-right:\s*0[\s\S]*border-bottom:\s*1px solid var\(--border\)/);
  assert.match(css, /@media\s*\(max-width:\s*700px\)[\s\S]*\.financial-reconciliation-automation-proposal-records\s*>\s*\.financial-reconciliation-automation-item\s*\+\s*\.financial-reconciliation-automation-item\s*\{[\s\S]*border-left:\s*0[\s\S]*border-top:\s*1px solid var\(--border\)/);
});
```

- [ ] **Step 2: Run the focused density test to verify RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-density.test.js
```

Expected: FAIL because the proposal list still has a gap, proposal rows still have a rounded enclosing border, and the metadata/items still use card backgrounds.

- [ ] **Step 3: Implement the minimal Option C CSS**

Update the desktop proposal styles without changing markup:

```css
.financial-reconciliation-workbench-automation-proposals {
  display: grid;
  gap: 0;
  border-top: 1px solid var(--border);
}

.financial-reconciliation-automation-proposal {
  display: grid;
  grid-template-columns: minmax(8.5rem, .55fr) minmax(0, 1fr) minmax(0, 1fr);
  gap: 0;
  padding: 0;
  border: 0;
  border-bottom: 1px solid var(--border);
  border-left: 4px solid var(--brand);
  border-radius: 0;
  overflow-wrap: anywhere;
  background: transparent;
}

.financial-reconciliation-automation-proposal-meta {
  display: grid;
  align-content: start;
  gap: .3rem;
  min-width: 0;
  padding: .55rem;
  border-right: 1px solid var(--border);
  background: transparent;
  font-size: .7rem;
}

.financial-reconciliation-automation-item {
  display: grid;
  gap: .25rem;
  min-width: 0;
  padding: .45rem;
  border: 0;
  border-radius: 0;
  line-height: 1.25;
  background: transparent;
}
```

Retain the existing item-to-item `border-left` rule and the existing `700px` overrides that remove `border-right`/`border-left` and add `border-bottom`/`border-top` separators.

- [ ] **Step 4: Run focused tests to verify GREEN**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-density.test.js tests/reconciliation-automation-ui.test.js
```

Expected: PASS with the new divider contract and all existing proposal behavior tests green.

- [ ] **Step 5: Run the full regression suite and CSS diff check**

Run:

```powershell
node --test --test-isolation=none
git diff --check
```

Expected: all tests pass and `git diff --check` reports no whitespace errors.

- [ ] **Step 6: Perform visual verification when possible**

Open Automatic reconciliation on desktop and a viewport narrower than `700px` using an authenticated session. Verify three separated desktop columns, one horizontal separator per proposal, stacked narrow sections, visible focus, long-text wrapping, ambiguous candidate groups, and multiple destinations. If no authenticated session exists, record the limitation and rely on executable CSS/markup contracts.

- [ ] **Step 7: Commit the layout change**

```powershell
git add -- styles.css tests/reconciliation-density.test.js
git commit -m "style: divide automatic reconciliation proposals"
```
