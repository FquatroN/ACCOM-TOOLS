# Automatic Reconciliation Amount-Only Rules Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add two disabled-by-default managed automatic rules that reconcile one exact-`Banco` or exact-`Visa` Financial Document to exactly one Bank Statement or Credit Card item using only signed amount equality and an inclusive configurable date window.

**Architecture:** Extend the existing explicit managed-rule registry from two to four key/version pairs. A single forward-only Supabase migration adds immutable amount-only definitions, fixed-zero configuration guards, allowlisted one-to-one candidate adapters, atomic execution revalidation, and four-rule scheduled snapshots. Existing Node endpoints remain generic but validate the expanded registry, while Settings renders the fixed zero tolerance read-only for the two new rules.

**Tech Stack:** PostgreSQL/Supabase RPC and PostgREST, Vercel Node functions, browser JavaScript, HTML/CSS, Node's built-in test runner, transactional PostgreSQL smoke tests.

## Global Constraints

- Preserve Bank Statement identity rule `financial_documents_cgd_bank_statement` version `2` and Credit Card identity rule `financial_documents_cgd_credit_card` version `1` byte-for-byte in behavior.
- Add `financial_documents_cgd_bank_statement_amount_only` version `1` and `financial_documents_cgd_credit_card_amount_only` version `1`; both start disabled for manual and scheduled execution.
- Bank amount-only bases require exact, case- and whitespace-sensitive `payment = 'Banco'`; Credit Card amount-only bases require exact `payment = 'Visa'`.
- Every base also requires `fat = 'S'`, non-null amount/date, date on or after `2026-01-01`, and no active reconciliation lock.
- Bank destinations use `import_cgd_extrato_ordem.data` and `montante`; Credit Card destinations use `import_cgd_cartao_credito.data` and `valor`. Never use `data_valor` for Credit Card reconciliation dates.
- A candidate qualifies only when base cents plus destination cents equals exactly zero and the destination date falls inclusively within base date plus/minus the snapshotted `0..90` day setting.
- Each proposal contains exactly one destination. Do not enumerate or accept multi-item combinations.
- Zero candidates are skipped without a review row; one candidate is proposed; multiple candidates, candidate-limit results, and cross-base reuse of one destination are non-executable ambiguity.
- Do not read invoice number, description, supplier, supplier NIF, or similarity scores when admitting or rejecting amount-only candidates.
- Difference allowed is permanently `0.00` for both amount-only rules. The UI displays it read-only; API normalization and database replacement reject any nonzero value.
- Keep the `financial_documents -> import_cgd_extrato_ordem (+)` and `financial_documents -> import_cgd_cartao_credito (+)` directional source rules immutable.
- Preserve existing administrator priority order. Append the new Bank amount-only rule and then the new Credit Card amount-only rule at the next two available priorities; the standard installation becomes priorities `3` and `4`.
- Manual analysis processes exactly one selected rule and never advances automatically. Scheduled batches snapshot enabled rules in priority/rule-key order and finish one child before starting the next.
- Unknown keys, unsupported versions, invalid snapshots, nonzero tolerance, changed payment/date/amount/operator, or source-lock conflicts fail closed.
- New SQL helpers use `SECURITY DEFINER SET search_path = public, pg_temp`; revoke access from `public`, `anon`, and `authenticated`, and grant only required `service_role` execution.
- Reuse existing safe extension wrappers. No new amount-only adapter may call `extensions.unaccent`, `digest`, `similarity`, or `word_similarity`.
- Use integer cents for comparisons, retain immutable run/proposal/config evidence, sanitize public errors, and preserve idempotency under retries and concurrency.
- Implement every production change with strict RED/GREEN TDD. Do not claim PostgreSQL smoke success unless it actually runs against a database.

---

### Task 1: Four-rule application registry and immutable-zero validation

**Files:**
- Modify: `api/_reconciliation-automation.js`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Consumes: current explicit `AUTOMATIC_RULE_VERSIONS`, managed-settings normalization, RPC-settings normalization, and manual `analyze_rule` payload validation.
- Produces:
  - `BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY = "financial_documents_cgd_bank_statement_amount_only"` and version `1`;
  - `CREDIT_CARD_AMOUNT_ONLY_RULE_KEY = "financial_documents_cgd_credit_card_amount_only"` and version `1`;
  - four-entry immutable `AUTOMATIC_RULE_VERSIONS`;
  - immutable `AMOUNT_ONLY_RULE_KEYS` and `isAmountOnlyRuleKey(ruleKey)`;
  - nonzero amount-only tolerance rejection in both browser and RPC-shaped settings paths.

- [ ] **Step 1: Add failing registry and tamper tests**

In `tests/reconciliation-automation.test.js`, import the two new constants and helper. Add fixtures with priorities `3` and `4`, disabled/manual false/scheduled false, `differenceAllowed: "0.00"`, and `maxDifferenceDays: 1`.

Add executable assertions equivalent to:

```js
test("managed automation accepts the four explicit key/version pairs", () => {
  const normalized = normalizeAutomationSettingsPayload(fourRuleSettings());
  assert.deepEqual(normalized.rules.map(({ ruleKey, ruleVersion }) => [ruleKey, ruleVersion]), [
    [BANK_STATEMENT_RULE_KEY, 2],
    [CREDIT_CARD_RULE_KEY, 1],
    [BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY, 1],
    [CREDIT_CARD_AMOUNT_ONLY_RULE_KEY, 1],
  ]);
  assert.equal(isAmountOnlyRuleKey(BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY), true);
  assert.equal(isAmountOnlyRuleKey(CREDIT_CARD_RULE_KEY), false);
});

test("amount-only tolerance is fixed at zero in both settings shapes", () => {
  assert.throws(() => normalizeAutomationSettingsPayload(fourRuleSettings({
    amountOnlyDifferenceAllowed: "0.01",
  })), /amount-only.*zero/i);
  assert.throws(() => normalizeRpcSettings(fourRuleRpcSettings({
    amountOnlyDifferenceAllowedCents: 1,
  })), /amount-only.*zero/i);
});
```

Also prove unsupported amount-only version `2`, near-name injection, duplicate keys, and a two-key manual analysis request are rejected before any RPC.

- [ ] **Step 2: Run focused tests and capture RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL only on the missing amount-only constants/registry entries and immutable-zero validation.

- [ ] **Step 3: Implement the minimal explicit registry extension**

In `api/_reconciliation-automation.js`, add:

```js
const BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY = "financial_documents_cgd_bank_statement_amount_only";
const BANK_STATEMENT_AMOUNT_ONLY_RULE_VERSION = 1;
const CREDIT_CARD_AMOUNT_ONLY_RULE_KEY = "financial_documents_cgd_credit_card_amount_only";
const CREDIT_CARD_AMOUNT_ONLY_RULE_VERSION = 1;

const AUTOMATIC_RULE_VERSIONS = Object.freeze({
  [BANK_STATEMENT_RULE_KEY]: BANK_STATEMENT_RULE_VERSION,
  [CREDIT_CARD_RULE_KEY]: CREDIT_CARD_RULE_VERSION,
  [BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY]: BANK_STATEMENT_AMOUNT_ONLY_RULE_VERSION,
  [CREDIT_CARD_AMOUNT_ONLY_RULE_KEY]: CREDIT_CARD_AMOUNT_ONLY_RULE_VERSION,
});

const AMOUNT_ONLY_RULE_KEYS = Object.freeze(new Set([
  BANK_STATEMENT_AMOUNT_ONLY_RULE_KEY,
  CREDIT_CARD_AMOUNT_ONLY_RULE_KEY,
]));

function isAmountOnlyRuleKey(ruleKey) {
  return AMOUNT_ONLY_RULE_KEYS.has(ruleKey);
}
```

After parsing the amount in `normalizeManagedRule` and `normalizeRpcSettings`, reject `differenceAllowedCents !== 0` for `isAmountOnlyRuleKey(ruleKey)`. Export all new constants and the predicate. Do not require all four rules at this application boundary: a pre-migration two-rule database response must remain valid during rollout.

- [ ] **Step 4: Verify GREEN and commit**

Run:

```powershell
node --check api/_reconciliation-automation.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none
git diff --check
```

Expected: all commands exit `0`.

Commit:

```powershell
git add -- api/_reconciliation-automation.js tests/reconciliation-automation.test.js
git commit -m "feat: register amount-only reconciliation rules"
```

---

### Task 2: Managed SQL definitions, configuration, and fixed-zero settings contract

**Files:**
- Create: `supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql`
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- Consumes: existing automatic rule definition/config tables, fixed `+` source rules, `replace_financial_reconciliation_automation_settings`, `get_financial_reconciliation_automation_settings`, and current two identity-rule contracts.
- Produces:
  - immutable version-1 rows for both amount-only definitions;
  - disabled config rows with zero tolerance, one-day window, and deterministic appended priorities;
  - four-entry database allowlist with fixed-zero guards;
  - supporting indexes on Bank and Credit Card destination `(amount, date, id)` lookup order;
  - reapply-safe migration and schema-cache notification.

- [ ] **Step 1: Write migration source-contract and smoke assertions first**

Add `AMOUNT_ONLY_MIGRATION_PATH` to `tests/reconciliation-automation.test.js`. Assert that the migration includes both exact keys, versions, display names, exact `Banco`/`Visa`, fixed `+`, `maxDestinationRecords: 1`, fixed zero tolerance, `0..90` day validation, service-role ACLs, and `pgrst` schema reload. Assert it does not contain dynamic SQL derived from configuration or any similarity function call in an amount-only adapter.

At the top of the new section in `tests/reconciliation-automation-rpc.smoke.sql`, apply the migration twice:

```sql
\ir ../supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql
\ir ../supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql
```

Add assertions that:

- both immutable definitions exactly equal the approved JSON/logic description;
- both configs are disabled, manual false, scheduled false, tolerance `0`, days `1`;
- Bank amount-only precedes Credit Card amount-only and both follow every pre-existing config without rewriting relative existing priorities;
- a standard priorities `1,2` installation produces `3,4`;
- source operators remain `+`;
- the Settings getter returns all four rules once;
- Settings replacement accepts edited days/priority but rejects `0.01` for either amount-only rule and leaves every row unchanged;
- deleting either amount-only rule from the post-migration complete payload is rejected;
- reapplying the migration preserves administrator enablement, days, and priority.

- [ ] **Step 2: Capture RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL because the dated migration and its four-rule database contract do not exist.

- [ ] **Step 3: Create the re-runnable definition/config migration**

In `supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql`:

1. Insert each definition with `ON CONFLICT DO NOTHING`, then compare every immutable column and JSON property in a `DO` block and raise if an installed definition differs.
2. Verify the two directional source rules exist with `operator = '+'`.
3. Lock rule configs, calculate `max(priority)`, and insert only missing amount-only configs in Bank-then-Card order. Use deferred unique-priority handling so reapplication never shifts existing rows.
4. Add/verify indexes:

```sql
create index if not exists import_cgd_extrato_ordem_reconciliation_amount_date_id_idx
  on public.import_cgd_extrato_ordem (montante, data, id)
  where montante is not null and data is not null;

create index if not exists import_cgd_cartao_credito_reconciliation_amount_date_id_idx
  on public.import_cgd_cartao_credito (valor, data, id)
  where valor is not null and data is not null;
```

5. Replace the Settings getter/replacer with the current signatures. The replacer must discover the installed managed definition set, require each installed key exactly once, validate the explicit key/version contract, reject nonzero tolerance for the two amount-only keys, and preserve atomic replacement.
6. Reapply fixed search paths, RLS/privileges, and `NOTIFY pgrst, 'reload schema'`.

- [ ] **Step 4: Verify Task 2 and commit**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none
git diff --check
```

If `psql` and `SUPABASE_DB_URL` are available, also run:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Expected: Node tests pass. SQL smoke must pass if executed; otherwise record it as an external gate.

Commit:

```powershell
git add -- supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql tests/reconciliation-automation.test.js tests/reconciliation-automation-rpc.smoke.sql
git commit -m "feat: define amount-only reconciliation rules"
```

---

### Task 3: One-to-one amount/date analysis adapters and ambiguity semantics

**Files:**
- Modify: `supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql`
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- Consumes: `financial_reconciliation_automatic_rule_contract(text,integer)`, base page/count, candidate dispatcher/page, proposal persistence, overlap classification, and run continuation.
- Produces:
  - four-entry rule contract dispatcher;
  - `financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])`;
  - `financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])`;
  - generic candidate dispatcher routes for all four adapters;
  - exact one-to-one candidate JSON and audit evidence.

- [ ] **Step 1: Add failing behavior-bearing SQL fixtures**

Create isolated Bank and Credit Card fixtures proving, for both rules:

- exact payment qualifies; wrong case, leading/trailing space, blank, null, `fat <> 'S'`, pre-2026 date, null date/amount, and locked base do not;
- destination uses the approved date/amount fields and excludes pre-2026, null date/amount, and locked rows;
- configured day `-1`, `0`, and `+1` qualify at default; `-2` and `+2` do not;
- a `0` day configuration admits same-day only, and `90` admits the boundary;
- `100.00 + -100.00 = 0` qualifies while a one-cent mismatch and same-sign amount do not;
- deliberately matching and nonmatching descriptions/suppliers/invoice numbers do not alter qualification;
- no destination yields skipped/no review row; one yields one proposed row with one item; two yields ambiguous; more than the evidence limit yields `candidate_limit`;
- one destination qualifying for two bases makes both ambiguous;
- two destinations whose sum would balance never produce a proposal;
- each returned candidate item includes stable source type/id/date/amount and evidence for exact signed amount/date distance, with no similarity evidence.

Add source tests that require indexed predicates using destination amount equality, inclusive date range, and stable ID ordering.

- [ ] **Step 2: Capture RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL because the rule contract and candidate dispatcher do not recognize the two new adapters.

- [ ] **Step 3: Implement explicit amount-only adapter functions**

Extend `financial_reconciliation_automatic_rule_contract` with immutable objects containing at least:

```sql
jsonb_build_object(
  'payment', 'Banco', -- or Visa
  'destinationSourceType', 'import_cgd_extrato_ordem', -- or card
  'matchingMode', 'amount_only_one_to_one',
  'maxDestinationRecords', 1,
  'maxCandidates', <existing safe evidence limit>,
  'fixedDifferenceAllowed', 0
)
```

Each amount-only candidate function must:

- require its exact key/version and `p_difference_allowed = 0`;
- materialize only requested eligible base IDs;
- convert amounts through the existing integer-cent helper/path;
- use `destination_amount_cents = -base_amount_cents` plus inclusive date range;
- exclude active locks;
- order by destination date and ID;
- return a bounded evidence array and an unbounded/limit-aware count sufficient for safe ambiguity;
- never call the identity candidate builder or combination enumerator.

Extend `financial_reconciliation_automatic_candidates_for_base_ids`, base page/count, candidate page, single-base compatibility function, and continuation dispatch without changing the two existing branches. Keep page size `25`, cursor order `(document_date,id)`, post-page cross-base overlap detection, and atomic page advancement.

- [ ] **Step 4: Verify analysis regression coverage and commit**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none
git diff --check
```

Run the transactional SQL smoke when the database gate is available.

Commit:

```powershell
git add -- supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql tests/reconciliation-automation.test.js tests/reconciliation-automation-rpc.smoke.sql
git commit -m "feat: analyze one-to-one amount matches"
```

---

### Task 4: Atomic amount-only execution, stale detection, and immutable evidence

**Files:**
- Modify: `supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql`
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- Consumes: `execute_financial_reconciliation_automatic_proposal`, destination lock dispatcher, reconciliation action RPC, proposal/run snapshots, and existing nested rollback/failure lifecycle.
- Produces:
  - allowlisted Bank/Card amount-only destination lock branches;
  - exact one-item execution revalidation;
  - stale reasons for rule/config/operator/payment/date/amount/item/lock drift;
  - automatic reconciliation audit evidence preserving amount-only rule/version/config.

- [ ] **Step 1: Add failing execution and concurrency fixtures**

For each amount-only rule, add transactional scenarios proving:

- executing a selected unique proposal creates exactly one reconciliation containing one Financial Document and one destination;
- the computed difference is `0`, origin is automatic/manual or automatic/scheduled as appropriate, and rule/config/operator/source snapshots are immutable;
- repeat execution is idempotent and cannot duplicate reconciliation/items;
- changing base payment, base amount/date, destination amount/date, configured window, rule definition/version, fixed operator, or proposal item count makes the proposal stale and creates nothing;
- setting amount-only tolerance nonzero in a tampered snapshot fails closed;
- a destination locked after analysis becomes stale;
- two concurrent/competing proposals cannot consume the same destination;
- an ambiguous or candidate-limit proposal cannot execute;
- a one-item identity-rule proposal and multi-item identity-rule proposal still execute under their existing behavior.

- [ ] **Step 2: Capture RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL on missing amount-only lock/revalidation dispatch contracts.

- [ ] **Step 3: Extend locking and revalidation without dynamic SQL**

Update the current destination-lock dispatcher to select `FOR UPDATE` from `import_cgd_extrato_ordem` or `import_cgd_cartao_credito` only through literal allowlisted branches. In `execute_financial_reconciliation_automatic_proposal`:

1. lock run, proposal, base, and exactly one destination in the existing global order;
2. verify the rule contract and immutable definition/config/source-rule snapshots;
3. require snapshot tolerance zero and one item for amount-only keys;
4. repeat exact payment, date floor, unlocked state, signed integer-cent equality, and inclusive date-window checks from analysis;
5. compare current source snapshots to persisted proposal evidence;
6. mark stale and return a sanitized result on deterministic drift;
7. use the existing nested rollback boundary for post-write failure;
8. verify the reconciliation snapshot/difference after `financial_reconciliation_action` before committing proposal completion.

Do not weaken existing identity-rule limits (`1..4`) or evidence checks.

- [ ] **Step 4: Verify and commit**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none
git diff --check
```

Run SQL smoke when available.

Commit:

```powershell
git add -- supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql tests/reconciliation-automation.test.js tests/reconciliation-automation-rpc.smoke.sql
git commit -m "feat: execute amount-only reconciliations atomically"
```

---

### Task 5: Four-rule manual API and sequential scheduled batches

**Files:**
- Modify: `api/reconciliation-automation.js`
- Modify: `api/reconciliation-automation-settings.js`
- Modify: `api/reconciliation-automation-cron.js`
- Modify: `supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql`
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- Consumes: generic manual analysis endpoint, settings endpoint, cron claim/continue/finalize handler, batch snapshot RPC, and one-child-at-a-time scheduled lifecycle.
- Produces:
  - manual analysis for either enabled amount-only rule using one selected key;
  - post-migration four-rule Settings catalog;
  - deterministic scheduled snapshots and sequential children for any enabled subset of the four rules;
  - compatibility with the pre-migration two-rule catalog.

- [ ] **Step 1: Add failing endpoint and scheduled-order tests**

In `tests/reconciliation-automation.test.js`, add handler tests proving:

- manual Analyze accepts each new key/version and sends exactly one key to `create_financial_reconciliation_automatic_analysis`;
- disabled/manual-false rules are not offered by the settings response used by the workbench;
- an unknown fifth key is rejected before RPC;
- Settings GET/PUT accepts two installed rules before migration and four after migration, but nonzero amount-only tolerance is rejected before RPC;
- cron accepts a four-rule batch response, resumes the current child, and does not claim the next child before the current child is terminal;
- standard order is identity Bank, identity Card, amount-only Bank, amount-only Card;
- configured priority changes determine future batch order;
- a failed amount-only child records failure and the next heartbeat advances to the following rule;
- retry/cross-midnight behavior does not duplicate a batch, child run, proposal, or reconciliation.

In SQL smoke, enable all four rules in a controlled fixture, assert the immutable batch snapshot order, then terminalize children one by one and verify no later child exists early.

- [ ] **Step 2: Run focused tests and capture RED where generic assumptions remain**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: any remaining two-rule cardinality or explicit-key assumptions fail. If the generic API files already pass unchanged, retain the tests and do not make gratuitous production edits.

- [ ] **Step 3: Generalize only the remaining cardinality boundaries**

Use `AUTOMATIC_RULE_VERSIONS` everywhere the Node handlers validate keys/versions. Keep manual `ruleKeys.length === 1`. In the SQL claim/get-settings functions:

- derive installed/scheduled rule count from authoritative configs/definitions;
- snapshot rules ordered by `priority, rule_key`;
- require each definition to resolve through the literal contract dispatcher;
- preserve one rule per child run and current resume behavior;
- accept failed children as terminal and continue later children;
- do not auto-enable either amount-only rule.

- [ ] **Step 4: Verify APIs, cron, scheduling, and commit**

Run:

```powershell
node --check api/reconciliation-automation.js
node --check api/reconciliation-automation-settings.js
node --check api/reconciliation-automation-cron.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none
git diff --check
```

Run SQL smoke when available.

Commit only files that actually changed:

```powershell
git add -- api/reconciliation-automation.js api/reconciliation-automation-settings.js api/reconciliation-automation-cron.js supabase-migrations/2026-08-17-financial-reconciliation-automation-amount-only-rules.sql tests/reconciliation-automation.test.js tests/reconciliation-automation-rpc.smoke.sql
git commit -m "feat: schedule four reconciliation rules sequentially"
```

---

### Task 6: Settings read-only tolerance and manual rule selector integration

**Files:**
- Modify: `app-main.js`
- Modify: `styles.css`
- Modify: `tests/reconciliation-automation-ui.test.js`
- Modify: `tests/reconciliation-density.test.js`

**Interfaces:**
- Consumes: existing dynamic automatic-rule cards, Settings draft serializer, rule LOV, active-run restoration/locking, proposal layout, and shared history.
- Produces:
  - amount-only predicate in browser code using the two exact keys;
  - read-only localized `0.00 €` tolerance display for amount-only rules;
  - editable days/priority/enabled/manual/scheduled controls;
  - enabled amount-only entries in the current manual rule LOV;
  - unchanged active-run lock and proposal/history behavior.

- [ ] **Step 1: Write failing browser behavior and source-density tests**

Extend UI fixtures to contain all four rules. Add tests proving:

- Settings renders a number input with `data-reconciliation-automation-rule-field="differenceAllowed"` for each identity rule;
- Settings renders a non-editable semantic value `0.00 €` and no difference input for each amount-only rule;
- days, priority, enabled, manual, and scheduled remain editable for amount-only rules;
- changing other fields and saving retains authoritative `differenceAllowed: "0.00"` for amount-only rules;
- a synthetic change event targeting amount-only `differenceAllowed` is ignored;
- enabled/manual amount-only rules appear escaped in the workbench selector;
- Analyze sends only the selected amount-only key;
- selector stays locked during create/restore/unfinished run and unlocks only after terminal state;
- proposal three-column layout and shared history render unchanged.

Add a density/source assertion for a `.financial-reconciliation-automation-fixed-value` style aligned with other controls at desktop and narrow widths.

- [ ] **Step 2: Capture RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
```

Expected: FAIL because every rule currently renders an editable difference input.

- [ ] **Step 3: Implement the fixed-value UI**

In `app-main.js`, define the exact two-key predicate near the automation Settings renderer. For amount-only cards, replace the editable input with:

```html
<label>Difference allowed
  <output class="financial-reconciliation-automation-fixed-value"
    aria-label="Difference allowed, fixed">0.00 €</output>
</label>
```

Keep the normalized draft's `differenceAllowed` value from the authoritative API response. In `validateReconciliationAutomationSettingsDraft`, require it to normalize to zero for amount-only keys. In field synchronization and change handling, never accept a difference edit for those keys. Do not disable the entire card or disturb focus by rerendering on every input.

In `styles.css`, style the output to match input height/border/background while visually communicating read-only state and preserving narrow-layout flow.

- [ ] **Step 4: Verify UI and commit**

Run:

```powershell
node --check app-main.js
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
node --test --test-isolation=none
git diff --check
```

If an authenticated local browser session is available, verify Settings desktop/narrow layouts, rule selection, Analyze lock, proposal review, and history. Otherwise document the exact authentication limitation.

Commit:

```powershell
git add -- app-main.js styles.css tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
git commit -m "feat: configure amount-only reconciliation rules"
```

---

### Task 7: Release documentation and complete cross-layer verification

**Files:**
- Modify: `README.md`
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- Consumes: completed Tasks 1-6 and current migration/rollout documentation.
- Produces: exact migration order, reapply instructions, controlled enablement sequence, and final local/external verification evidence.

- [ ] **Step 1: Add a failing release-contract test**

Require `README.md` to list the new migration after `2026-08-16-financial-reconciliation-automation-credit-card-rule.sql` and state:

- deploy compatibility-tolerant application code before applying the migration;
- apply and reapply the new migration;
- both amount-only rules start disabled;
- validate manual Bank then manual Card before enabling scheduled execution;
- SQL smoke and protected browser scenarios are mandatory rollout gates.

Run the focused test and confirm RED because README lacks this step.

- [ ] **Step 2: Update migration and rollout documentation**

Add the exact command/order and verification notes to `README.md`. Do not imply Vercel applies Supabase migrations automatically. Explain that Settings will show two rules before migration and four afterward.

- [ ] **Step 3: Run all local verification gates**

Run:

```powershell
node --check api/_reconciliation-automation.js
node --check api/reconciliation-automation.js
node --check api/reconciliation-automation-settings.js
node --check api/reconciliation-automation-cron.js
node --check app-main.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
node --test --test-isolation=none
node -e "JSON.parse(require('fs').readFileSync('vercel.json','utf8')); console.log('vercel.json valid')"
git diff --check
git status --short
```

Expected: every syntax/test/JSON/diff command exits `0`; status contains only intended changes plus any pre-existing unrelated user files.

- [ ] **Step 4: Run mandatory external gates when available**

Database:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Protected browser scenarios:

1. pre-migration app loads two-rule Settings safely;
2. post-migration Settings shows four rules and both new rules disabled;
3. both amount-only tolerances display read-only `0.00 €`;
4. nonzero tampered Settings payload is rejected with no partial write;
5. Bank unique, duplicate, cross-base overlap, and no-match fixtures classify correctly;
6. Credit Card equivalents classify correctly;
7. manual selector analyzes only the selected rule and locks during the run;
8. selected unique proposal executes to zero-difference history/audit;
9. changed source becomes stale and creates no reconciliation;
10. scheduled four-rule order and failed-child continuation behave correctly.

If database credentials or authenticated browser state are unavailable, record these as explicit rollout gates rather than claiming success.

- [ ] **Step 5: Request final review and commit release documentation**

Use `superpowers:requesting-code-review` for a full review against:

`docs/superpowers/specs/2026-08-17-automatic-reconciliation-amount-only-rules-design.md`

Resolve every Critical/Important finding with RED/GREEN tests and re-run all gates. Then commit:

```powershell
git add -- README.md tests/reconciliation-automation.test.js tests/reconciliation-automation-rpc.smoke.sql
git commit -m "docs: release amount-only reconciliation rules"
```

Before merge/publish, use `superpowers:verification-before-completion` and `superpowers:finishing-a-development-branch`. Publish the compatibility application first, apply the new Supabase migration manually second, run SQL smoke/reapply third, and enable the new rules only after controlled manual verification.
