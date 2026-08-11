# Reconciliation Source Rules Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the three reconciliation workbench selectors with one Source selector and let administrators configure independent directional source rules with `+` or `-` operators.

**Architecture:** Persist one rule per base-source/target-source direction in Supabase, and snapshot the applicable rules onto each reconciliation when it starts. The reconciliation API and SQL RPCs derive permitted candidates and difference calculations from the snapshot, while a new Settings section manages future rules. The workbench consumes the returned rules to provide a single source-driven candidate list with no pairing dialog.

**Tech Stack:** Static HTML, vanilla browser JavaScript, CSS, Node.js Vercel API handlers, Supabase/PostgREST RPCs, PostgreSQL PL/pgSQL, Node built-in test runner.

## Global Constraints

- Eligible records must retain the existing `2026-01-01` minimum date and existing source-specific eligibility rules.
- Rules are directional: saving `A -> B` must never create, remove, or alter `B -> A`.
- Each rule has exactly one operator, either `+` or `-`; source and target must be distinct supported source types.
- New reconciliations snapshot their rules at Start; later Settings edits cannot change existing started or completed reconciliations, totals, locks, or audit history.
- The database and API must independently reject missing, invalid, or non-snapshotted rules; the browser must not be the authorization boundary.
- Preserve existing add/remove, completion, force-completion comment, reopen, delete, record-locking, history, and current-details behavior.

---

## File structure

- `supabase-migrations/2026-08-11-financial-reconciliation-source-rules.sql` — creates and seeds directional rules, adds rule snapshots, migrates existing reconciliations, and replaces the reconciliation RPC functions.
- `api/_reconciliation.js` — source-rule validation, snapshot validation, workspace/mutation payload validation, and operator-driven difference calculation.
- `api/reconciliation.js` — passes the simplified workspace and mutation payloads to Supabase.
- `api/reconciliation-settings.js` — authorized GET/PUT endpoint for persisted directional rules.
- `api/_supabase.js` — exposes the new `financial-reconciliation` Settings permission.
- `index.html` — Settings navigation/editor and the one-selector workbench markup.
- `app-main.js` — settings state/editor events, source-driven workbench state, and API calls.
- `styles.css` — compact settings-rule editor and read-only workbench compatibility hint.
- `tests/reconciliation.test.js` — executable server-side validation and calculation coverage.
- `tests/reconciliation-density.test.js` — static UI contract for the removed selectors, replacement selector, Settings editor, and compatibility hint.
- `tests/reconciliation-rpc.smoke.sql` — database contract checks for rule direction, snapshots, and RPC enforcement.

## Directional rule contract

Use these browser/API objects consistently:

```js
// Persisted Settings row and API response item.
{ baseSourceType: "financial_documents", matchingSourceType: "import_cgd_extrato_ordem", operator: "+" }

// Snapshot stored in financial_reconciliations.matching_source_rules.
{ sourceType: "import_cgd_extrato_ordem", operator: "+" }
```

The initial migration seeds the current approved directions, preserving today’s calculations:

```text
financial_documents -> import_fdm_accounts (+)
financial_documents -> import_cgd_cartao_credito (+)
financial_documents -> import_cgd_extrato_ordem (+)
import_fdm_accounts -> import_cgd_extrato_ordem (-)
import_cgd_cartao_credito -> financial_documents (+)
import_cgd_cartao_credito -> import_cgd_extrato_ordem (+)
import_cgd_extrato_ordem -> financial_documents (+)
import_cgd_extrato_ordem -> import_fdm_accounts (-)
import_cgd_extrato_ordem -> import_cgd_cartao_credito (+)
```

### Task 1: Add directional-rule persistence and snapshot-safe RPCs

**Files:**
- Create: `supabase-migrations/2026-08-11-financial-reconciliation-source-rules.sql`
- Modify: `api/_reconciliation.js`
- Modify: `api/reconciliation.js`
- Modify: `tests/reconciliation.test.js`
- Modify: `tests/reconciliation-rpc.smoke.sql`

**Interfaces:**
- Produces `normalizeReconciliationRules(value) -> Array<{baseSourceType, matchingSourceType, operator}>`.
- Produces `normalizeRuleSnapshot(baseSourceType, value) -> Array<{sourceType, operator}>`.
- Changes `calculateDifference(baseSourceType, matchingSourceRules, items) -> number`.
- Changes `validateWorkspaceQuery(query)` to return `{ reconciliationId, sourceType, page, pageSize, filters }` with no client-supplied matching types.
- Changes `validateMutation("start", payload)` to return `{ action, sourceType, sourceId }`; the RPC derives base rules from persisted settings.
- Replaces RPC calls with `get_financial_reconciliation_workspace(uuid, text, jsonb, integer, integer)` and `financial_reconciliation_action(text, text, uuid, text, uuid, text)`.

- [ ] **Step 1: Write failing Node tests for directional rules and operator calculations**

  Add the following tests to `tests/reconciliation.test.js` and update the import list to include `normalizeReconciliationRules` and `normalizeRuleSnapshot`:

  ```js
  test("rules preserve independent directions and operators", () => {
    assert.deepEqual(normalizeReconciliationRules([
      { baseSourceType: "financial_documents", matchingSourceType: "import_cgd_extrato_ordem", operator: "-" },
      { baseSourceType: "import_cgd_extrato_ordem", matchingSourceType: "financial_documents", operator: "+" },
    ]), [
      { baseSourceType: "financial_documents", matchingSourceType: "import_cgd_extrato_ordem", operator: "-" },
      { baseSourceType: "import_cgd_extrato_ordem", matchingSourceType: "financial_documents", operator: "+" },
    ]);
  });

  test("rules reject self-pairs, duplicates, and unknown operators", () => {
    assert.throws(() => normalizeReconciliationRules([{ baseSourceType: "financial_documents", matchingSourceType: "financial_documents", operator: "+" }]), /different/i);
    assert.throws(() => normalizeReconciliationRules([
      { baseSourceType: "financial_documents", matchingSourceType: "import_fdm_accounts", operator: "+" },
      { baseSourceType: "financial_documents", matchingSourceType: "import_fdm_accounts", operator: "-" },
    ]), /duplicate/i);
    assert.throws(() => normalizeReconciliationRules([{ baseSourceType: "financial_documents", matchingSourceType: "import_fdm_accounts", operator: "*" }]), /operator/i);
  });

  test("difference uses the base amount and each directional snapshot operator", () => {
    const rules = normalizeRuleSnapshot("financial_documents", [
      { sourceType: "import_cgd_extrato_ordem", operator: "-" },
      { sourceType: "import_cgd_cartao_credito", operator: "+" },
    ]);
    assert.equal(calculateDifference("financial_documents", rules, [
      { sourceType: "financial_documents", amountSnapshot: 100 },
      { sourceType: "import_cgd_extrato_ordem", amountSnapshot: 60 },
      { sourceType: "import_cgd_cartao_credito", amountSnapshot: 40 },
    ]), 80);
  });

  test("a Start action contains no client-selected matching sources", () => {
    assert.deepEqual(validateMutation("start", {
      sourceType: "financial_documents", sourceId: "record-1",
    }), { action: "start", reconciliationId: "", sourceType: "financial_documents", sourceId: "record-1", comment: "" });
  });
  ```

- [ ] **Step 2: Run the focused test file and confirm the new assertions fail**

  Run: `node --test tests/reconciliation.test.js`

  Expected: FAIL because the rule normalizers do not exist and `calculateDifference` still accepts a string-array matching mode.

- [ ] **Step 3: Implement pure server validation and calculation helpers**

  In `api/_reconciliation.js`, retain the four `SOURCE_TYPES`, remove `MATCHING_SOURCE_TYPES`, and add explicit normalizers:

  ```js
  function normalizeOperator(value) {
    if (value !== "+" && value !== "-") throw inputError("Rule operator must be '+' or '-'.");
    return value;
  }

  function normalizeReconciliationRules(value) {
    if (!Array.isArray(value)) throw inputError("Reconciliation rules must be an array.");
    const seen = new Set();
    return value.map((rule) => {
      const baseSourceType = normalizeSourceType(rule?.baseSourceType || rule?.base_source_type);
      const matchingSourceType = normalizeSourceType(rule?.matchingSourceType || rule?.matching_source_type);
      if (baseSourceType === matchingSourceType) throw inputError("Rule sources must be different.");
      const key = `${baseSourceType}:${matchingSourceType}`;
      if (seen.has(key)) throw inputError("Duplicate reconciliation rule.");
      seen.add(key);
      return { baseSourceType, matchingSourceType, operator: normalizeOperator(rule?.operator) };
    });
  }

  function normalizeRuleSnapshot(baseSourceType, value) {
    const base = normalizeSourceType(baseSourceType);
    const rows = Array.isArray(value) ? value : [];
    const seen = new Set();
    return rows.map((rule) => {
      const sourceType = normalizeSourceType(rule?.sourceType || rule?.source_type);
      if (sourceType === base) throw inputError("Snapshot source must differ from the base source.");
      if (seen.has(sourceType)) throw inputError("Duplicate reconciliation rule snapshot.");
      seen.add(sourceType);
      return { sourceType, operator: normalizeOperator(rule?.operator) };
    });
  }

  function calculateDifference(baseSourceType, matchingSourceRules, items) {
    const base = normalizeSourceType(baseSourceType);
    const rules = normalizeRuleSnapshot(base, matchingSourceRules);
    const operators = new Map(rules.map((rule) => [rule.sourceType, rule.operator]));
    return roundMoney((Array.isArray(items) ? items : []).reduce((total, item) => {
      const sourceType = normalizeSourceType(item?.sourceType);
      const sign = sourceType === base ? 1 : operators.get(sourceType) === "-" ? -1 : operators.get(sourceType) === "+" ? 1 : null;
      if (sign === null) throw inputError("Item source type is not allowed for this reconciliation.");
      return total + (sign * amountFor(item));
    }, 0));
  }
  ```

  Make `validateWorkspaceQuery` reject/ignore no matching-source input and make `validateMutation("start")` derive no pairings from the request. Export all new helper functions.

- [ ] **Step 4: Run the focused Node tests and confirm they pass**

  Run: `node --test tests/reconciliation.test.js`

  Expected: PASS, including existing validation/lock/error tests updated to use rule snapshots.

- [ ] **Step 5: Write the migration and SQL smoke assertions before changing the RPC implementation**

  Create `supabase-migrations/2026-08-11-financial-reconciliation-source-rules.sql` with these structural statements before the function definitions:

  ```sql
  create table if not exists public.financial_reconciliation_source_rules (
    base_source_type text not null check (base_source_type in ('financial_documents','import_fdm_accounts','import_cgd_cartao_credito','import_cgd_extrato_ordem')),
    matching_source_type text not null check (matching_source_type in ('financial_documents','import_fdm_accounts','import_cgd_cartao_credito','import_cgd_extrato_ordem')),
    operator text not null check (operator in ('+','-')),
    created_at timestamptz not null default timezone('utc', now()),
    updated_at timestamptz not null default timezone('utc', now()),
    primary key (base_source_type, matching_source_type),
    check (base_source_type <> matching_source_type)
  );

  alter table public.financial_reconciliations
    add column if not exists matching_source_rules jsonb not null default '[]'::jsonb;
  ```

  Seed the nine directions listed in the plan’s Directional rule contract using `insert ... values ... on conflict (...) do nothing`. Backfill blank historical `matching_source_rules` to the equivalent `{sourceType, operator}` entries based on each existing base source and `matching_source_types` order. Add smoke assertions proving both `financial_documents -> import_cgd_extrato_ordem (-)` and `import_cgd_extrato_ordem -> financial_documents (+)` can coexist and that existing rows have a non-empty snapshot when their historical matching array was non-empty.

- [ ] **Step 6: Replace the SQL rules, action, difference, and workspace functions**

  In the same migration, replace the prior hard-coded mode checks with these rules:

  ```sql
  -- On start, fetch and snapshot all configured outbound rules for the base source.
  select coalesce(jsonb_agg(jsonb_build_object('sourceType', matching_source_type, 'operator', operator)
                            order by matching_source_type), '[]'::jsonb)
    into v_rules
    from public.financial_reconciliation_source_rules
   where base_source_type = p_source_type;
  if v_rules = '[]'::jsonb then
    raise exception 'No reconciliation rules are configured for this source.';
  end if;
  ```

  The new difference function must add every base item and then add/subtract each matching item using `matching_source_rules`; it must raise when a non-base item has no snapshot entry. The action function must:

  - insert the fetched rule snapshot and its source-type list on `start`;
  - use only `r.matching_source_rules` for `add_item` eligibility and recalculation;
  - retain source eligibility, lock conflict, completion, reopen, delete, and audit logic;
  - include the rule snapshot in the `start` audit metadata.

  The workspace function must return `rules` for an unstarted source and `matchingSourceRules` from the selected reconciliation snapshot when it is open. It must validate the selected candidate source against the base plus this rule set, not against request-supplied pairings. Revoke/grant the new RPC signatures for `service_role`, and remove grants for obsolete signatures.

- [ ] **Step 7: Update the API handler to use the simplified RPC contract**

  In `api/reconciliation.js`, stop passing `p_matching_source_types`. Send exactly:

  ```js
  body: {
    p_reconciliation_id: query.reconciliationId || null,
    p_source_type: query.sourceType,
    p_filters: query.filters,
    p_page: query.page,
    p_page_size: query.pageSize,
  }
  ```

  For mutations, pass `p_action`, `p_actor`, `p_reconciliation_id`, `p_source_type`, `p_source_id`, and `p_comment`; do not accept or send `baseSourceType` or `matchingSourceTypes`.

- [ ] **Step 8: Run automated and database-contract verification**

  Run: `node --test tests/reconciliation.test.js`

  Expected: PASS.

  After applying the migration in Supabase, run: `psql "$SUPABASE_DB_URL" -f tests/reconciliation-rpc.smoke.sql`

  Expected: all assertions pass, including independent reverse-direction operators, a snapshot-stable started reconciliation, and rejection of a source absent from its snapshot.

- [ ] **Step 9: Commit the database and API contract**

  ```bash
  git add supabase-migrations/2026-08-11-financial-reconciliation-source-rules.sql api/_reconciliation.js api/reconciliation.js tests/reconciliation.test.js tests/reconciliation-rpc.smoke.sql
  git commit -m "feat: add directional reconciliation rules"
  ```

### Task 2: Expose rule management under Settings → Reconciliation

**Files:**
- Create: `api/reconciliation-settings.js`
- Modify: `api/_supabase.js`
- Modify: `app-main.js`
- Modify: `index.html`
- Modify: `styles.css`
- Modify: `tests/reconciliation.test.js`
- Modify: `tests/reconciliation-density.test.js`

**Interfaces:**
- `GET /api/reconciliation-settings` returns `{ rules: Array<DirectionalRule> }`.
- `PUT /api/reconciliation-settings` accepts `{ rules: Array<DirectionalRule> }` and returns the validated saved list.
- Browser state adds `reconciliationRules`, `reconciliationRulesLoaded`, and `reconciliationRuleBaseSource`.
- Browser functions: `loadReconciliationSettings()`, `renderReconciliationSettings()`, `saveReconciliationSettings()`, and `reconciliationRulesFor(baseSourceType)`.

- [ ] **Step 1: Write failing tests for Settings access and payload validation**

  Add tests confirming the shared rule normalizer accepts different reverse operators and rejects a self-pair. Add a source-level UI contract test:

  ```js
  test("settings exposes a reconciliation rule editor", () => {
    assert.match(html, /id="settings-menu-financial-reconciliation"/);
    assert.match(html, /id="settings-view-financial-reconciliation"/);
    assert.match(html, /id="financial-reconciliation-settings-base-source"/);
    assert.match(html, /id="financial-reconciliation-settings-rules-body"/);
    assert.match(html, /id="financial-reconciliation-settings-save"/);
  });
  ```

- [ ] **Step 2: Run the focused tests and confirm the Settings UI assertion fails**

  Run: `node --test tests/reconciliation.test.js tests/reconciliation-density.test.js`

  Expected: FAIL because neither the Settings feature nor editor markup exists.

- [ ] **Step 3: Create the authorized Settings API route and permission**

  Add `financial-reconciliation` to `SETTINGS_FEATURES` in `api/_supabase.js`. Implement `api/reconciliation-settings.js` using `requireFeature(req, "settings", "financial-reconciliation")`:

  ```js
  const toRule = (row) => ({
    baseSourceType: row.base_source_type,
    matchingSourceType: row.matching_source_type,
    operator: row.operator,
  });
  const toRow = (rule) => ({
    base_source_type: rule.baseSourceType,
    matching_source_type: rule.matchingSourceType,
    operator: rule.operator,
  });

  if (req.method === "GET") {
    const rules = await restQuery("financial_reconciliation_source_rules?select=base_source_type,matching_source_type,operator&order=base_source_type.asc,matching_source_type.asc", { method: "GET" });
    return res.status(200).json({ rules: rules.map(toRule) });
  }

  if (req.method === "PUT") {
    const input = normalizeReconciliationRules((await parseBody(req))?.rules);
    await restQuery("financial_reconciliation_source_rules", { method: "DELETE" });
    if (input.length) await restQuery("financial_reconciliation_source_rules", { method: "POST", body: input.map(toRow) });
    return res.status(200).json({ rules: input });
  }
  ```

  Keep the delete-and-replace operation server-side and only after `normalizeReconciliationRules` has validated the full submitted list. Return `405` for unsupported methods and use `sendError` for failures.

- [ ] **Step 4: Add Settings navigation, markup, and browser editor state**

  Add a Reconciliation button after Import Data in the Settings navigation, matching the existing `settings-menu-*` pattern. Add an editor panel with:

  ```html
  <div id="settings-view-financial-reconciliation" hidden>
    <div class="actions">
      <button id="close-settings-financial-reconciliation" type="button" class="ghost">Back To App</button>
      <button id="financial-reconciliation-settings-save" type="button" class="ghost">Save Configuration</button>
    </div>
    <section class="settings-block">
      <div class="bar"><h3>Reconciliation rules</h3></div>
      <label>Source<select id="financial-reconciliation-settings-base-source"></select></label>
      <p class="field-hint">Each row is one direction. Configure the reverse direction separately.</p>
      <table class="settings-table"><thead><tr><th>Compatible source</th><th>Operator</th></tr></thead><tbody id="financial-reconciliation-settings-rules-body"></tbody></table>
      <p id="financial-reconciliation-settings-status" class="auth-status"></p>
    </section>
  </div>
  ```

  For the selected base source, render one row for every other supported source with a checkbox and a disabled/enabled `+`/`-` select. The checkbox changes only that directional rule. Selecting a different base source must not mutate the previous source’s rules.

- [ ] **Step 5: Wire lifecycle, access, and save behavior in `app-main.js`**

  Add Settings DOM references, event listeners, visibility and active-tab handling, `setSettingsSection("financial-reconciliation")` authorization, and fallback section selection. Load the editor only for users with the new Settings permission. Implement save as:

  ```js
  async function saveReconciliationSettings() {
    const rules = collectReconciliationSettingsRules();
    const result = await api("/api/reconciliation-settings", { method: "PUT", body: { rules } });
    state.reconciliationRules = result.rules;
    state.reconciliationRulesLoaded = true;
    els.financialReconciliationSettingsStatus.textContent = "Configuration saved.";
    financialReconciliationState().loaded = false;
  }
  ```

  Include `financial-reconciliation` in the browser-side Settings feature options used by admin-user editing. Use concise styling that matches existing Settings tables and keeps the operator control compact.

- [ ] **Step 6: Run Settings/UI regression tests**

  Run: `node --test tests/reconciliation.test.js tests/reconciliation-density.test.js`

  Expected: PASS, including Settings markup, access contract, and directional validation assertions.

- [ ] **Step 7: Commit the Settings editor**

  ```bash
  git add api/_supabase.js api/reconciliation-settings.js app-main.js index.html styles.css tests/reconciliation.test.js tests/reconciliation-density.test.js
  git commit -m "feat: manage reconciliation rules in settings"
  ```

### Task 3: Simplify the workbench to a source-driven candidate list

**Files:**
- Modify: `index.html`
- Modify: `app-main.js`
- Modify: `styles.css`
- Modify: `tests/reconciliation-density.test.js`
- Modify: `tests/reconciliation.test.js`

**Interfaces:**
- Replaces `baseSourceType`, `matchingSourceTypes`, and `candidateSourceType` UI controls with `candidateSourceType` as the sole Source selection state.
- `reconciliationRulesFor(baseSourceType)` returns `Array<{sourceType, operator}>` from the unstarted workspace or selected reconciliation snapshot.
- `onFinancialReconciliationSourceChange()` sets the source, resets the page, and reloads the workspace.
- `start` payload is `{ action: "start", sourceType, sourceId }`.

- [ ] **Step 1: Write failing workbench contract tests**

  Replace the existing selector assumptions in `tests/reconciliation-density.test.js` with:

  ```js
  test("workbench uses one source selector and displays saved rule hints", () => {
    assert.match(html, /<label>Source<select id="financial-reconciliation-source"><\/select><\/label>/);
    assert.match(html, /id="financial-reconciliation-rule-hint"/);
    assert.doesNotMatch(html, /id="financial-reconciliation-base-source"/);
    assert.doesNotMatch(html, /id="financial-reconciliation-matching-sources"/);
    assert.doesNotMatch(html, /id="financial-reconciliation-candidate-source"/);
    assert.match(appMain, /function onFinancialReconciliationSourceChange\(\)/);
    assert.match(appMain, /action: "start", sourceType, sourceId/);
  });
  ```

  Add a server test confirming a workspace query needs only source, page, and filters:

  ```js
  test("workspace query no longer accepts a matching-source mode", () => {
    assert.deepEqual(validateWorkspaceQuery({ source_type: "import_cgd_extrato_ordem", page: "1" }), {
      reconciliationId: "", sourceType: "import_cgd_extrato_ordem", page: 1, pageSize: 50, filters: {},
    });
  });
  ```

- [ ] **Step 2: Run focused tests and confirm they fail**

  Run: `node --test tests/reconciliation.test.js tests/reconciliation-density.test.js`

  Expected: FAIL because the three current selector IDs and mode-change handler still exist.

- [ ] **Step 3: Replace workbench markup and DOM references**

  Replace the three labels in `index.html` with:

  ```html
  <label>Source<select id="financial-reconciliation-source"></select></label>
  <p id="financial-reconciliation-rule-hint" class="field-hint financial-reconciliation-rule-hint"></p>
  ```

  Remove the three prior DOM references/listeners in `app-main.js` and add `financialReconciliationSource` plus `financialReconciliationRuleHint`. Keep dynamic filters in the same card, so source changes still render the correct source-specific filter set.

- [ ] **Step 4: Implement source options, hint, and automatic Start behavior**

  Replace `renderFinancialReconciliationModeControls` with a function that derives its options as follows. Define `normalizeFinancialReconciliationRuleSnapshot(value)` in `app-main.js` to accept only array entries whose `sourceType` is one of `FINANCIAL_RECONCILIATION_SOURCES` and whose operator is `+` or `-`; return `{ sourceType, operator }` values with duplicates removed by source type.

  ```js
  const reconciliation = financialReconciliationActiveRecord();
  const base = clean(reconciliation?.base_source_type);
  const rules = reconciliation
    ? normalizeFinancialReconciliationRuleSnapshot(reconciliation.matching_source_rules)
    : normalizeFinancialReconciliationRuleSnapshot(workspace.rules);
  const allowedSources = reconciliation
    ? [base, ...rules.map((rule) => rule.sourceType)]
    : Object.keys(FINANCIAL_RECONCILIATION_SOURCES);
  ```

  Keep the selected source when it remains allowed; otherwise select the base for an open reconciliation or the first supported source before one starts. Render the hint from outbound rules when unstarted, and from the active reconciliation snapshot when started:

  ```js
  els.financialReconciliationRuleHint.textContent = rules.length
    ? `Matches: ${rules.map((rule) => `${financialReconciliationSourceLabel(rule.sourceType)} (${rule.operator})`).join(", ")}`
    : "No reconciliation rules are configured for this source.";
  ```

  Set `Start` disabled for a source with no outbound rules, and preserve the existing row-level Start/Add action. `onFinancialReconciliationSourceChange` must reset `page` to `1`, mark the workspace unloaded, and call `loadFinancialReconciliationWorkspace({ silent: true })`. Starting from a row sends only its selected source and ID; adding sends only the selected source, ID, and active reconciliation ID.

- [ ] **Step 5: Preserve safe reload and active-reconciliation behavior**

  Update `loadFinancialReconciliationWorkspace` so it retains the previous active workspace and previous selected source when a request fails. After a successful action, use the returned reconciliation’s stored snapshot rather than refetching Settings rules to determine allowable candidate sources. Continue to hide Start once a reconciliation exists, and continue to use `Add` for every allowed source in that active reconciliation.

- [ ] **Step 6: Style the simplified controls and hint**

  In `styles.css`, remove layout rules that reserve space for three source controls. Add scoped rules so the Source selector and compatibility hint align with compact workbench typography:

  ```css
  .financial-reconciliation-rule-hint { align-self: end; margin: 0; font-size: .68rem; }
  .financial-reconciliation-workbench-card .financial-reconciliation-rule-hint { grid-column: span 2; }
  ```

  At the mobile breakpoint, make the hint span the full control grid. Do not change the approved eligible-record or current-reconciliation density rules.

- [ ] **Step 7: Run the complete automated regression suite**

  Run: `node --test tests/*.test.js`

  Expected: PASS. The suite must prove the one-selector markup, removal of the old controls, directional validation, signed calculations, empty-rule behavior, and existing current-details/density behavior.

- [ ] **Step 8: Perform a manual browser verification**

  Verify these flows against a Supabase environment that has run the migration:

  1. Open Settings → Reconciliation; configure `Financial Documents -> CGD Bank Statement (-)` and save.
  2. Configure `CGD Bank Statement -> Financial Documents (+)` separately and confirm the first rule’s operator is unchanged.
  3. Open Reconciliation; select each source and confirm Eligible records and dynamic filters refresh.
  4. Start from Financial Documents without any dialog; confirm the current panel shows the captured rule sources and adding a bank item recalculates using `-`.
  5. Edit the saved Settings rule; reopen the started reconciliation and confirm its prior difference and permitted sources use the earlier snapshot.
  6. Confirm an unconfigured source shows the explicit no-rules message and cannot start.

- [ ] **Step 9: Commit the source-driven workbench**

  ```bash
  git add index.html app-main.js styles.css tests/reconciliation-density.test.js tests/reconciliation.test.js
  git commit -m "feat: simplify reconciliation source selection"
  ```

## Final verification and handoff

- [ ] Run `git diff --check` and `node --test tests/*.test.js` from the feature worktree.
- [ ] Apply `supabase-migrations/2026-08-11-financial-reconciliation-source-rules.sql` once in the target Supabase project, then run `psql "$SUPABASE_DB_URL" -f tests/reconciliation-rpc.smoke.sql`.
- [ ] Review the final diff for unrelated files; stage only the files listed in this plan.
- [ ] Confirm the Settings feature permission is granted to the intended administrator profile before expecting the new menu entry to appear.
