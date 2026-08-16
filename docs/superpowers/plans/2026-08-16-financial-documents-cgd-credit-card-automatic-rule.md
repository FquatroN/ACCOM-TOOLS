# Financial Documents to CGD Credit Card Automatic Rule Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a managed automatic rule that reconciles exact-`Visa` Financial Documents with unique one-to-four-record CGD Credit Card combinations, and generalize manual and scheduled automation so future managed rules use the same audited lifecycle.

**Architecture:** A forward-only Supabase migration adds the immutable Credit Card rule, a synchronized indexed card-search projection, explicit allowlisted SQL rule adapters, generalized paged analysis/execution, and deterministic parent batches with one child run per scheduled rule. Node APIs validate the same two-rule registry, while the Automatic reconciliation UI changes from per-rule cards to a single rule LOV and Settings replaces manual batch execution with navigation to that workbench.

**Tech Stack:** PostgreSQL/Supabase RPC and PostgREST, `pg_trgm`, Vercel Node functions, browser JavaScript, HTML/CSS, Node's built-in test runner, transactional PostgreSQL smoke tests.

## Global Constraints

- Keep Bank Statement rule key `financial_documents_cgd_bank_statement` at immutable version `2` with exact existing matching behavior.
- Add Credit Card rule key `financial_documents_cgd_credit_card` at immutable version `1`; seed it disabled for manual and scheduled use.
- Credit Card bases require exact case- and whitespace-sensitive `financial_documents.payment = 'Visa'`, `fat = 'S'`, an unlocked source, and a date on or after `2026-01-01`.
- Credit Card destinations use `import_cgd_cartao_credito.data`, `descricao`, and generated `valor`; never use `data_valor` for reconciliation matching.
- Credit Card identity is OR logic: symmetric compact invoice-number containment with at least four characters, description similarity `>= 0.55`, or supplier word similarity `>= 0.60`. Every destination item must independently satisfy at least one signal.
- Credit Card combinations contain one through four records, admit no more than 12 identity candidates per base, and calculate `financial_documents.amount + sum(import_cgd_cartao_credito.valor)` using integer-cent tolerance.
- The configurable difference allowance defaults to `0.00`; the inclusive date window defaults to 10 days and remains constrained to `0` through `90`.
- Each run contains exactly one immutable rule/config/operator snapshot. Manual analysis never queues or starts another rule.
- The daily schedule snapshots enabled rules once, processes child runs sequentially by ascending priority and stable rule-key tie-breaker, and continues after an individual rule failure.
- Analysis page size stays server-owned at 25. Partial analyses cannot execute proposals.
- Managed definitions, adapter keys, tables, SQL, functions, operators, thresholds, and combination limits are never accepted from browser payloads.
- New internal SQL functions are `SECURITY DEFINER SET search_path = public, pg_temp` plus the catalog-discovered `pg_trgm` schema where required. Revoke access from `public`, `anon`, and `authenticated`; grant only the required `service_role` execution.
- Use the existing extension wrapper functions. Do not call `extensions.unaccent`, `digest`, `similarity`, or `word_similarity` directly from new fixed-search-path functions.
- Preserve completed reconciliations, runs, proposals, locks, provenance, and audit evidence. Public errors stay sanitized.
- Implement strict RED/GREEN TDD, finish each task with a focused review and commit, and never claim the PostgreSQL smoke passed unless it actually ran against a database.

---

### Task 1: Two-rule public contract and removal of manual batch analysis

**Files:**
- Modify: `api/_reconciliation-automation.js`
- Modify: `api/reconciliation-automation.js`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Consumes: existing Bank Statement constants, settings normalization, snake-to-camel mapper, and `analyze_rule` handler.
- Produces:
  - `BANK_STATEMENT_RULE_KEY = "financial_documents_cgd_bank_statement"` and version `2`;
  - `CREDIT_CARD_RULE_KEY = "financial_documents_cgd_credit_card"` and version `1`;
  - immutable `AUTOMATIC_RULE_VERSIONS` lookup;
  - `normalizeRuleKey(value)` accepting exactly those two keys;
  - `normalizeRuleVersion(value, ruleKey)` validating the key/version pair;
  - manual-only `normalizeAnalyzePayload(...)` accepting exactly one selected rule;
  - public mappings for `batch_id`, `batch_rule_key`, `batch_rule_position`, `batch_rule_count`, and `last_scheduled_batch`;
  - no public `analyze_batch` action or handler.

- [ ] **Step 1: Write failing two-rule contract tests**

Add literal fixtures and assertions to `tests/reconciliation-automation.test.js`:

```js
const creditCardRule = {
  ruleKey: CREDIT_CARD_RULE_KEY,
  ruleVersion: CREDIT_CARD_RULE_VERSION,
  enabled: false,
  allowManualExecution: false,
  includeInScheduledBatch: false,
  differenceAllowed: "0.00",
  maxDifferenceDays: 10,
  priority: 2,
};

test("managed settings accept the two explicit rule/version pairs", () => {
  const input = managedSettings({
    rules: [managedSettings().rules[0], creditCardRule],
  });
  assert.deepEqual(
    normalizeAutomationSettingsPayload(input).rules.map(({ ruleKey, ruleVersion }) => ({ ruleKey, ruleVersion })),
    [
      { ruleKey: BANK_STATEMENT_RULE_KEY, ruleVersion: 2 },
      { ruleKey: CREDIT_CARD_RULE_KEY, ruleVersion: 1 },
    ],
  );
  assert.throws(() => normalizeAutomationSettingsPayload({
    ...input,
    rules: [{ ...creditCardRule, ruleVersion: 2 }],
  }), /rule version/i);
});

test("manual analysis accepts exactly one allowlisted rule and has no batch action", () => {
  assert.deepEqual(normalizeAnalyzePayload({
    action: "analyze_rule",
    ruleKeys: [CREDIT_CARD_RULE_KEY],
    clientRequestId: REQUEST_ID,
  }).ruleKeys, [CREDIT_CARD_RULE_KEY]);
  assert.throws(() => normalizeAnalyzePayload({
    action: "analyze_rule",
    ruleKeys: [BANK_STATEMENT_RULE_KEY, CREDIT_CARD_RULE_KEY],
    clientRequestId: REQUEST_ID,
  }), /exactly one/i);
  assert.throws(() => normalizeAutomationAction("analyze_batch"), /automation action/i);
});

test("batch lifecycle keys map without leaking diagnostic keys", () => {
  assert.deepEqual(toAutomationPublicResult({
    batch_id: RUN_ID,
    batch_rule_key: CREDIT_CARD_RULE_KEY,
    batch_rule_position: 2,
    batch_rule_count: 3,
    last_scheduled_batch: { error_detail: "hidden", status: "partial" },
  }), {
    batchId: RUN_ID,
    batchRuleKey: CREDIT_CARD_RULE_KEY,
    batchRulePosition: 2,
    batchRuleCount: 3,
    lastScheduledBatch: { status: "partial" },
  });
});
```

Update handler tests to prove POST `analyze_batch` returns 400 before any RPC and `analyze_rule` sends exactly one selected Credit Card key to `create_financial_reconciliation_automatic_analysis`.

- [ ] **Step 2: Run the focused test and verify RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL because the helper accepts only the Bank Statement key, still exposes `analyze_batch`, and lacks batch field mappings.

- [ ] **Step 3: Implement the explicit registry and manual-only action**

Replace the single-key validator in `api/_reconciliation-automation.js` with:

```js
const BANK_STATEMENT_RULE_KEY = "financial_documents_cgd_bank_statement";
const BANK_STATEMENT_RULE_VERSION = 2;
const CREDIT_CARD_RULE_KEY = "financial_documents_cgd_credit_card";
const CREDIT_CARD_RULE_VERSION = 1;
const AUTOMATIC_RULE_VERSIONS = Object.freeze({
  [BANK_STATEMENT_RULE_KEY]: BANK_STATEMENT_RULE_VERSION,
  [CREDIT_CARD_RULE_KEY]: CREDIT_CARD_RULE_VERSION,
});
const AUTOMATION_ACTIONS = new Set(["analyze_rule", "continue_analysis", "execute_selected"]);

function normalizeRuleKey(value) {
  if (!Object.hasOwn(AUTOMATIC_RULE_VERSIONS, value)) throw inputError("Rule key is invalid.");
  return value;
}

function normalizeRuleVersion(value, ruleKey) {
  if (value !== AUTOMATIC_RULE_VERSIONS[ruleKey]) throw inputError("Rule version is invalid.");
  return value;
}
```

Normalize `ruleKey` before `ruleVersion` in both managed settings paths. Require `normalizeRuleKeys(input.ruleKeys).length === 1` and require action `analyze_rule`. Keep `AUTOMATIC_RULE_KEY` and `AUTOMATIC_RULE_VERSION` exported as Bank Statement aliases only if existing non-production tests still consume them; new production validation must use the registry.

Add to `PUBLIC_KEY_MAP`:

```js
batch_id: "batchId",
batch_rule_key: "batchRuleKey",
batch_rule_position: "batchRulePosition",
batch_rule_count: "batchRuleCount",
last_scheduled_batch: "lastScheduledBatch",
```

Delete `requireBatchFields`, `analyzeBatch`, the settings authorization branch for batch analysis, and the `analyze_batch` POST dispatch from `api/reconciliation-automation.js`.

- [ ] **Step 4: Run focused and full tests to verify GREEN**

Run:

```powershell
node --check api/_reconciliation-automation.js
node --check api/reconciliation-automation.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none
```

Expected: all commands exit 0.

- [ ] **Step 5: Commit Task 1**

```powershell
git add -- api/_reconciliation-automation.js api/reconciliation-automation.js tests/reconciliation-automation.test.js
git commit -m "feat: allow explicit automatic reconciliation rules"
```

---

### Task 2: Credit Card definition, projection, and SQL candidate adapter

**Files:**
- Create: `supabase-migrations/2026-08-16-financial-reconciliation-automation-credit-card-rule.sql`
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- Consumes: normalization/hash wrappers, existing Bank Statement indexed search, combination builder, source rules, and the latest 90-day migration.
- Produces:
  - immutable Credit Card definition/config row;
  - `financial_reconciliation_cgd_credit_card_match_search` plus synchronization trigger and three indexes;
  - `financial_reconciliation_automatic_rule_contract(text,integer) -> jsonb`;
  - `financial_reconciliation_automatic_bank_candidates_for_base_ids(text,integer,numeric,integer,uuid[])`;
  - `financial_reconciliation_automatic_credit_card_candidates_for_base_ids(text,integer,numeric,integer,uuid[])`;
  - generalized `financial_reconciliation_automatic_candidates_for_base_ids(text,integer,numeric,integer,uuid[])` dispatcher;
  - `financial_reconciliation_automatic_base_page(text,integer,date,uuid,integer)`;
  - `financial_reconciliation_automatic_base_count(text,integer) -> bigint`;
  - existing candidate-page, single-base, and compatibility signatures backed by the dispatcher.

- [ ] **Step 1: Add failing migration-contract and SQL fixtures**

In `tests/reconciliation-automation.test.js`, add `CREDIT_CARD_MIGRATION_PATH` and require the smoke script to include the new migration after `2026-08-16-financial-reconciliation-automation-90-day-performance.sql`. Assert literal managed values:

```js
assert.match(sql, /financial_documents_cgd_credit_card/);
assert.match(sql, /payment\s*=\s*'Visa'/);
assert.match(sql, /description_score\s*>=\s*0\.55/);
assert.match(sql, /supplier_score\s*>=\s*0\.60/);
assert.match(sql, /import_cgd_cartao_credito/);
assert.doesNotMatch(sql, /extensions\.unaccent|extensions\.digest/);
```

Extend `tests/reconciliation-automation-rpc.smoke.sql` with fixed Visa fixtures that prove:

- the config is inserted at priority 2, disabled/manual false/scheduled false, tolerance 0, days 10;
- the source rule is `financial_documents -> import_cgd_cartao_credito (+)`;
- projection INSERT/UPDATE/ID-change/DELETE synchronizes `data`, `valor`, and `descricao`;
- `data_valor` changes alone do not change the projected reconciliation date;
- exact `Visa` qualifies and `VISA`, `visa`, padded, null, pre-2026, and locked bases produce no candidate row;
- day 10 qualifies and day 11 does not;
- compact document-number matches require four characters and work in both containment directions;
- measured description scores immediately below/at `0.55` and supplier scores immediately below/at `0.60` enforce exact boundaries;
- one candidate can qualify through each identity branch independently;
- Bank Statement v2 fixture IDs/evidence remain byte-for-byte unchanged through the dispatcher;
- reapplying the migration does not duplicate definitions, configs, projection rows, triggers, or indexes.

- [ ] **Step 2: Run Node RED and record the PostgreSQL gate**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL because the migration and smoke include do not exist.

If both `psql` and `SUPABASE_DB_URL` exist, also run:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Expected before implementation: FAIL at the new rule/projection assertions. Otherwise record this exact command as an external gate.

- [ ] **Step 3: Seed the immutable rule and synchronized projection**

Start the migration with an immutable definition matching the approved design:

```sql
insert into public.financial_reconciliation_automatic_rule_definitions (
  rule_key, version, display_name, base_source_type,
  destination_source_types, logic_description, definition
) values (
  'financial_documents_cgd_credit_card', 1,
  'Financial Documents to CGD Credit Card', 'financial_documents',
  '["import_cgd_cartao_credito"]'::jsonb,
  'Payment must equal exactly Visa. Each credit-card candidate must satisfy invoice containment, description similarity, or supplier word similarity. Exactly one one-to-four-record amount combination is executable.',
  '{
    "baseEligibility":{"payment":{"operator":"exact_text_equal","value":"Visa","caseSensitive":true,"trim":false}},
    "identityBranches":{"document_number":{"algorithm":"symmetric_compact_containment"},"description_similarity":{"algorithm":"similarity"},"supplier_similarity":{"algorithm":"word_similarity"}},
    "documentNumberMinimumCompactLength":4,
    "descriptionSimilarityThreshold":0.55,
    "supplierWordSimilarityThreshold":0.60,
    "maxDestinationRecords":4,
    "maxIdentityCandidatesPerBase":12
  }'::jsonb
) on conflict (rule_key, version) do nothing;
```

Compare the complete stored row with those literals and raise if an immutable row differs. Verify the existing directional source rule has operator `+`. On the first insert only, normalize the current Bank-Statement-only priority to 1 and insert the card config at priority 2 with version 1, every boolean false, tolerance 0, and days 10. Once the card config exists, reapplication must not overwrite any administrator setting or priority.

Create `financial_reconciliation_cgd_credit_card_match_search(source_id, source_date, amount, description, normalized_description, compact_description, updated_at)`. Backfill from `import_cgd_cartao_credito(id,data,valor,descricao)`, install an `AFTER INSERT OR UPDATE OR DELETE` synchronizer, and add:

```sql
create index if not exists financial_reconciliation_cgd_credit_card_match_search_date_id_idx
  on public.financial_reconciliation_cgd_credit_card_match_search (source_date, source_id);

execute format(
  'create index if not exists financial_reconciliation_cgd_credit_card_match_search_normalized_trgm_idx
     on public.financial_reconciliation_cgd_credit_card_match_search
     using gin (normalized_description %I.gin_trgm_ops)',
  v_trgm_schema
);
execute format(
  'create index if not exists financial_reconciliation_cgd_credit_card_match_search_compact_trgm_idx
     on public.financial_reconciliation_cgd_credit_card_match_search
     using gin (compact_description %I.gin_trgm_ops)',
  v_trgm_schema
);
```

Handle an updated source ID by deleting `old.id` before upserting `new.id`. Enable RLS, revoke all table access, and grant only `service_role` SELECT.

- [ ] **Step 4: Install explicit rule adapters and candidate dispatch**

Create an allowlisted contract function whose only successful branches are:

```sql
case
  when p_rule_key = 'financial_documents_cgd_bank_statement' and p_rule_version = 2 then
    jsonb_build_object('payment','Banco','destinationSourceType','import_cgd_extrato_ordem',
      'descriptionThreshold',0.60,'supplierThreshold',0.70,'maxDestinationRecords',4,'maxCandidates',12)
  when p_rule_key = 'financial_documents_cgd_credit_card' and p_rule_version = 1 then
    jsonb_build_object('payment','Visa','destinationSourceType','import_cgd_cartao_credito',
      'descriptionThreshold',0.55,'supplierThreshold',0.60,'maxDestinationRecords',4,'maxCandidates',12)
  else null
end
```

Raise `Automatic reconciliation rule is unsupported.` when the contract is null at an RPC boundary.

Move the current Bank Statement base-ID query unchanged into its named helper. Implement the Credit Card helper against the new projection. Its exact final predicate is:

```sql
document_number_matched
or description_score >= 0.55
or supplier_score >= 0.60
```

`document_number_matched` is true only when the compact invoice has at least four characters and either compact value contains the other. Build destination snapshots with source type `import_cgd_cartao_credito`, date `data`, amount `valor`, description `descricao`, and literal score/threshold evidence.

Make `financial_reconciliation_automatic_candidates_for_base_ids(text,integer,numeric,integer,uuid[])` dispatch only to the two named helpers. Generalize base paging/count by contract Payment and keep `(document_date,id)` ordering. Redefine the existing candidate-page, single-base, and full compatibility functions to call the dispatcher without changing their signatures or return columns.

Revoke public/anon/authenticated execution on every helper, grant the required functions to `service_role`, and finish with `notify pgrst, 'reload schema';`.

- [ ] **Step 5: Run focused/full verification and commit Task 2**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none
git diff --check
```

Run the SQL smoke when available. Expected: Node tests and diff check pass; SQL smoke passes or remains explicitly unexecuted.

```powershell
git add -- supabase-migrations/2026-08-16-financial-reconciliation-automation-credit-card-rule.sql tests/reconciliation-automation.test.js tests/reconciliation-automation-rpc.smoke.sql
git commit -m "feat: add credit card reconciliation candidates"
```

---

### Task 3: Generalized one-rule resumable analysis

**Files:**
- Modify: `supabase-migrations/2026-08-16-financial-reconciliation-automation-credit-card-rule.sql`
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- Consumes: Task 2 rule contract, base count/page, candidate dispatcher, existing proposal schema, and combination builder.
- Produces generalized replacements for:
  - `create_financial_reconciliation_automatic_analysis(text[],text,text,uuid)`;
  - `continue_financial_reconciliation_automatic_analysis(uuid,text)`;
  - `continue_financial_reconciliation_automatic_oldest_analysis(text)`;
  - `financial_reconciliation_finalize_automatic_analysis(uuid)`;
  - `get_financial_reconciliation_automatic_active_run(text)`;
  - one-open-manual-run-per-actor enforcement.

- [ ] **Step 1: Add failing lifecycle smoke tests**

Add transactional assertions for both rule keys:

1. manual creation accepts mode `manual_rule`, exactly one enabled/manual key, and snapshots exactly one definition;
2. snapshot contains `destinationSourceType` and the current directional operator;
3. mode `manual_batch`, two keys, an unknown key, and a disabled/manual-false rule fail before inserting a run;
4. a second open run for the same actor and different request ID fails with a safe current-run conflict;
5. Credit Card analysis pages 25 ordered bases, advances cursor once, and never duplicates proposals on retry;
6. one-, two-, three-, and four-card exact-zero combinations become proposed;
7. a solution requiring five cards becomes skipped;
8. two valid combinations become ambiguous; 13 identity candidates become `candidate_limit`;
9. no-match/skipped rows affect counts but produce no visible executable row contract;
10. analysis with no proposed rows becomes terminal automatically; an analysis with a proposal remains `ready` until execution;
11. Bank Statement paging/counts remain unchanged.

- [ ] **Step 2: Run focused RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL on source-contract assertions for hard-coded `Banco`, Bank Statement destination, manual-batch mode, and the absence of single-open-run enforcement.

- [ ] **Step 3: Generalize creation, paging, and proposal construction**

In the migration, redefine creation so only this mode is accepted:

```sql
if p_mode <> 'manual_rule' or cardinality(p_rule_keys) <> 1 then
  raise exception 'Manual automatic analysis requires exactly one selected rule.';
end if;
```

Build the snapshot by joining the definition's single destination type to `financial_reconciliation_source_rules`, and store:

```sql
jsonb_build_object(
  'ruleKey', config.rule_key,
  'ruleVersion', config.rule_version,
  'displayName', definition.display_name,
  'priority', config.priority,
  'differenceAllowed', config.difference_allowed,
  'maxDifferenceDays', config.max_difference_days,
  'destinationSourceType', destination.source_type,
  'definition', definition.definition,
  'operator', source_rule.operator
)
```

Count eligible bases through `financial_reconciliation_automatic_base_count(ruleKey,ruleVersion)`. Before inserting, lock the actor's manual-run namespace and reject a different unfinished run. Add a partial unique index on manual actor where `finished_at is null`, after deterministically marking all but the newest preexisting duplicate open rows failed.

In continuation, require `jsonb_array_length(definition_config_snapshot) = 1`; fetch the contract, destination source, maximum candidates, and maximum destination records. Replace hard-coded construction with:

```sql
public.financial_reconciliation_automatic_build_combinations(
  v_base.base_snapshot,
  v_base.candidates,
  jsonb_build_object(v_destination_source_type, v_rule->>'operator'),
  (v_rule->>'differenceAllowed')::numeric,
  v_max_destination_records
)
```

Use the contract candidate limit rather than a second literal. Use base count/page helpers in creation, continuation, and oldest-analysis continuation. Preserve page size 25 and atomic cursor/proposal updates.

- [ ] **Step 4: Make zero-executable analyses terminal**

In `financial_reconciliation_finalize_automatic_analysis`, run overlap resolution first, recalculate counts from persisted proposals, and set:

```sql
status = case when exists (
  select 1 from public.financial_reconciliation_automatic_proposals
  where run_id = p_run_id and status = 'proposed'
) then 'ready' else 'completed' end,
finished_at = case when exists (
  select 1 from public.financial_reconciliation_automatic_proposals
  where run_id = p_run_id and status = 'proposed'
) then null else now() end,
analysis_completed_at = now()
```

Keep ambiguous/skipped evidence and counts. The active-run RPC returns only the actor's unfinished manual run. A terminal no-proposal run therefore releases the rule selector without deleting audit evidence.

- [ ] **Step 5: Verify and commit Task 3**

Run focused/full Node tests, SQL smoke when available, and `git diff --check`. Then:

```powershell
git add -- supabase-migrations/2026-08-16-financial-reconciliation-automation-credit-card-rule.sql tests/reconciliation-automation.test.js tests/reconciliation-automation-rpc.smoke.sql
git commit -m "feat: analyze one managed reconciliation rule"
```

---

### Task 4: Generic atomic proposal execution and audit evidence

**Files:**
- Modify: `supabase-migrations/2026-08-16-financial-reconciliation-automation-credit-card-rule.sql`
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- Consumes: Task 2 single-base adapter, Task 3 one-rule snapshot, existing reconciliation action RPC, source locks, and proposal lifecycle.
- Produces:
  - `financial_reconciliation_automatic_lock_destination_items(text,jsonb) -> integer`;
  - generalized `execute_financial_reconciliation_automatic_proposal(uuid,text) -> jsonb` supporting exactly the two rule/version pairs;
  - dynamic destination/operator audit snapshots with unchanged public outcome shape.

- [ ] **Step 1: Add failing execution and stale-path fixtures**

Extend SQL smoke coverage to prove:

- a unique Credit Card proposal creates one completed reconciliation with one Financial Document and one through four card items;
- the persisted difference is zero for default tolerance and matching-source rule is `import_cgd_cartao_credito (+)`;
- origin, trigger, rule key/version, run ID, proposal ID, source snapshots, operator snapshot, identity evidence, and proposal signature are retained;
- exact Visa Payment changing after analysis makes the proposal stale;
- changes to card `data`, `valor`, `descricao`, selected item IDs, rule version, definition, operator, tolerance, or evidence make it stale and create no reconciliation;
- a locked/deleted card makes it stale;
- two concurrent or repeated execution attempts create at most one reconciliation;
- Bank Statement execution and historical evidence remain unchanged;
- incomplete analysis and ambiguous/skipped/deselected proposals cannot execute.

- [ ] **Step 2: Run RED**

Run the focused Node source-contract test and SQL smoke when available. Expected: the current execution function still hard-codes Bank Statement key/version, table, destination type, operator map, and completion comment.

- [ ] **Step 3: Add a table-allowlisted destination lock helper**

Implement only explicit branches:

```sql
if p_source_type = 'import_cgd_extrato_ordem' then
  perform bank.id from jsonb_array_elements(p_items) item(value)
  join public.import_cgd_extrato_ordem bank on bank.id = (item.value->>'sourceId')::uuid
  order by bank.data, bank.id for update of bank;
elsif p_source_type = 'import_cgd_cartao_credito' then
  perform card.id from jsonb_array_elements(p_items) item(value)
  join public.import_cgd_cartao_credito card on card.id = (item.value->>'sourceId')::uuid
  order by card.data, card.id for update of card;
else
  raise exception 'Automatic reconciliation destination source is unsupported.';
end if;
get diagnostics v_count = row_count;
return v_count;
```

Do not use client-controlled dynamic SQL.

- [ ] **Step 4: Replace the execution RPC with adapter-driven revalidation**

Copy the latest hardened execution lifecycle into the new migration and replace every Bank Statement literal with values derived from the immutable contract/snapshot. Required guards include:

```sql
v_contract := public.financial_reconciliation_automatic_rule_contract(
  v_proposal.rule_key, v_proposal.rule_version
);
v_destination_source_type := v_contract->>'destinationSourceType';
if v_proposal.base_source_type <> 'financial_documents'
   or jsonb_array_length(v_run.definition_config_snapshot) <> 1 then
  -- persist stale rule_snapshot_changed
end if;
```

Validate every proposal item uses `v_destination_source_type`, lock through the helper, re-fetch the base through `financial_reconciliation_automatic_single_base_candidates`, and rebuild combinations with the dynamic operator map and contract maximum. Compare the full signature, items, evidence, snapshots, and calculated difference before setting `executing`.

Use the existing `financial_reconciliation_action` lifecycle with the explicit actions `start`, `add_item`, `complete`, and `force_complete`. Validate the created reconciliation's matching rule using the dynamic destination type. Build force-complete text from snapshot `displayName` and version, and persist:

```sql
'operatorSnapshot', jsonb_build_object(v_destination_source_type, v_rule_snapshot->>'operator')
```

Retain the existing nested exception/rollback behavior so a failed post-write validation leaves no reconciliation and stores only a sanitized failed outcome.

- [ ] **Step 5: Verify and commit Task 4**

Run focused/full Node tests, SQL smoke when available, and diff checks. Then:

```powershell
git add -- supabase-migrations/2026-08-16-financial-reconciliation-automation-credit-card-rule.sql tests/reconciliation-automation.test.js tests/reconciliation-automation-rpc.smoke.sql
git commit -m "feat: execute managed reconciliation adapters"
```

---

### Task 5: Deterministic daily batch and sequential child-run RPCs

**Files:**
- Modify: `supabase-migrations/2026-08-16-financial-reconciliation-automation-credit-card-rule.sql`
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- Consumes: shared schedule, rule configs, generalized analysis/execution, existing scheduled heartbeat identity.
- Produces:
  - table `financial_reconciliation_automatic_batches`;
  - run columns `batch_id`, `batch_rule_position`, `batch_rule_count`;
  - `financial_reconciliation_refresh_automatic_batch(uuid) -> jsonb`;
  - generalized `claim_financial_reconciliation_automatic_schedule(timestamptz,text) -> jsonb`;
  - generalized `get_financial_reconciliation_automatic_run(uuid)` and `financial_reconciliation_automatic_progress_or_run(uuid)` including child batch metadata;
  - `get_financial_reconciliation_automation_settings()` returning `lastScheduledBatch` data;
  - one scheduled child run per snapshotted rule in strict order.

- [ ] **Step 1: Add failing scheduled-batch smoke coverage**

Create fixtures with both rules scheduled and assert:

1. one due heartbeat creates one batch snapshot ordered Bank Statement priority 1, Credit Card priority 2;
2. each snapshot entry includes rule key/version, destination source, operator, config, definition, and priority;
3. the first claim returns a scheduled `scope = 'rule'` child containing only Bank Statement;
4. a second claim before terminal state resumes the same child and cannot create another;
5. after the first child finishes, the next claim creates only the Credit Card child;
6. changing configuration or priority after batch creation does not alter that batch but changes tomorrow's batch;
7. a failed first child is terminal and the second is still claimed;
8. equal priorities, if encountered in a catalog-corruption fixture, are ordered by rule key or rejected safely; normal Settings still enforces unique priorities;
9. retries and next-day/cross-midnight heartbeats do not duplicate batches, child runs, proposals, or reconciliations;
10. after every child is terminal, claim returns `batch_complete` and aggregate counts/status are stable;
11. historical scheduled runs remain readable and are not re-executed.

- [ ] **Step 2: Run RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL because current scheduled-slot uniqueness permits only one run, current claim rejects more than one enabled rule, and no batch entity exists.

- [ ] **Step 3: Add batch schema and safe historical migration**

Create:

```sql
create table if not exists public.financial_reconciliation_automatic_batches (
  id uuid primary key default gen_random_uuid(),
  scheduled_slot text not null check (scheduled_slot ~ '^\d{4}-\d{2}-\d{2}$'),
  actor text not null,
  status text not null check (status in ('pending','running','completed','partial','failed')),
  rule_snapshot jsonb not null check (jsonb_typeof(rule_snapshot) = 'array'),
  counts jsonb not null default '{}'::jsonb check (jsonb_typeof(counts) = 'object'),
  started_at timestamptz not null default now(),
  finished_at timestamptz,
  updated_at timestamptz not null default now(),
  unique (scheduled_slot)
);
```

Add these nullable batch columns to runs:

```sql
batch_id uuid references public.financial_reconciliation_automatic_batches(id),
batch_rule_key text,
batch_rule_position integer check (batch_rule_position is null or batch_rule_position > 0),
batch_rule_count integer check (batch_rule_count is null or batch_rule_count > 0)
```

Drop the current unique scheduled-slot index, retain a legacy uniqueness index where `batch_id is null`, and add a unique `(batch_id,batch_rule_position)` index plus a unique `(batch_id,batch_rule_key)` index. Extend the run trigger/scope check to permit new scheduled `scope='rule'` children while preserving historical scheduled `scope='batch'` rows.

Backfill one terminal legacy batch per existing scheduled slot, link its run, and copy its immutable snapshot/counts without rewriting the run or reconciliation. Deterministically fail only genuinely unfinished legacy scheduled rows using the existing upgrade/restart safe reason.

Enable RLS and revoke table access. Add fixed-search-path service-role RPC grants only.

- [ ] **Step 4: Implement batch snapshot, claim, progression, and aggregation**

On the first due claim for a Lisbon-local scheduled date, snapshot every config satisfying `enabled and include_in_scheduled_batch`, ordered by `priority,rule_key`. Each entry uses the same one-rule snapshot shape from Task 3. Store the complete array once.

Claim logic must:

```sql
-- under batch row lock
select the unfinished child first;
otherwise select the first snapshot element for which no child exists;
insert one scheduled scope='rule' run with jsonb_build_array(selected_rule);
initialize analysis_total through financial_reconciliation_automatic_base_count(
  selected_rule_key, selected_rule_version
);
return claimed/resumed, batchId, batchRulePosition, batchRuleCount, and run;
```

Never overwrite the batch snapshot from current Settings. When no snapshot entry remains, call `financial_reconciliation_refresh_automatic_batch`, set the terminal aggregate status (`completed`, `partial`, or `failed`), and return `{claimed:false, reason:'batch_complete', batchId}`.

Redefine both run serializers so every scheduled child returns `batchId`, `batchRuleKey`, `batchRulePosition`, and `batchRuleCount`; manual runs return null for those fields. Keep definitions/proposals authoritative and preserve the existing public run fields.

Refresh the parent after analysis failure and run finalization. A child failure is terminal and never selected as unfinished work. Generalize `continue_financial_reconciliation_automatic_oldest_analysis` to use the run's one-rule snapshot and adapter base count instead of `Banco`.

Update `get_financial_reconciliation_automation_settings()` to return the latest batch summary as `last_scheduled_batch`; include child count/outcome aggregates but no internal error detail.

- [ ] **Step 5: Verify and commit Task 5**

Run focused/full Node tests, SQL smoke when available, reapply assertions, and diff checks. Then:

```powershell
git add -- supabase-migrations/2026-08-16-financial-reconciliation-automation-credit-card-rule.sql tests/reconciliation-automation.test.js tests/reconciliation-automation-rpc.smoke.sql
git commit -m "feat: schedule reconciliation rules sequentially"
```

---

### Task 6: Manual API and cron batch integration

**Files:**
- Modify: `api/reconciliation-automation.js`
- Modify: `api/reconciliation-automation-cron.js`
- Modify: `api/reconciliation-automation-settings.js`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Consumes: Tasks 1 and 5 helper/RPC contracts.
- Produces:
  - app-authorized manual analysis for one selected rule;
  - cron validation for one-rule scheduled children and parent batch metadata;
  - safe sequential heartbeat response `{ batchId, ruleKey, rulePosition, ruleCount, hasMore }`;
  - no manual batch endpoint.

- [ ] **Step 1: Add failing API/cron behavior tests**

Add mocked handler tests that prove:

- `GET ?view=rules` returns both RPC-provided enabled/manual rules without inventing keys;
- POST `analyze_rule` accepts the Credit Card key and sends `p_mode:'manual_rule'` with exactly one key;
- POST `analyze_batch` is rejected before `requireFeature`/RPC execution;
- a scheduled child must have trigger `scheduled`, scope `rule`, exactly one definition, matching `batchId`, valid position/count, and proposals with that same rule key;
- a heartbeat resumes the current child until terminal, then a later heartbeat claims the next child;
- a failed first child produces HTTP 200 and permits the next child;
- `batch_complete`, schedule disabled, before time, and no enabled rules return safe HTTP 200 reasons;
- malformed batch/run/proposal responses fail closed with HTTP 500 and never execute a proposal;
- at most 25 stable proposed outcomes execute per heartbeat.

- [ ] **Step 2: Run focused RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL because cron requires `scope='batch'`, accepts multi-definition runs, and does not validate parent batch metadata.

- [ ] **Step 3: Harden the manual endpoint around one selected rule**

Keep the existing app authorization/actor binding. `analyzeRule` must call:

```js
createAnalysis({
  action: "analyze_rule",
  ruleKeys: [selectedRuleKey],
  clientRequestId,
}, actor, "manual_rule");
```

No Settings endpoint or action may create an analysis run. Keep settings GET/PUT administrator-only and map `lastScheduledBatch` through the shared public mapper.

- [ ] **Step 4: Validate and progress one scheduled child per heartbeat**

Update `requireScheduledRun` to require:

```js
run.trigger === "scheduled";
run.scope === "rule";
UUID_PATTERN.test(run.batchId);
Number.isSafeInteger(run.batchRulePosition) && run.batchRulePosition >= 1;
Number.isSafeInteger(run.batchRuleCount) && run.batchRulePosition <= run.batchRuleCount;
run.definitions.length === 1;
run.definitions[0].ruleKey === everyProposal.ruleKey;
```

Extend allowed claim reasons with `batch_complete`. Return only sanitized batch/run progress. Preserve the order: continue oldest analysis, claim/resume child, continue one page, execute up to 25 stable proposals, finalize child. Do not loop into the next rule in the same HTTP request; the next heartbeat claims it.

- [ ] **Step 5: Verify and commit Task 6**

Run:

```powershell
node --check api/reconciliation-automation.js
node --check api/reconciliation-automation-settings.js
node --check api/reconciliation-automation-cron.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none
git diff --check
```

Expected: all pass.

```powershell
git add -- api/reconciliation-automation.js api/reconciliation-automation-cron.js api/reconciliation-automation-settings.js tests/reconciliation-automation.test.js
git commit -m "feat: orchestrate one reconciliation rule per run"
```

---

### Task 7: Rule LOV, locked active workflow, and Settings navigation

**Files:**
- Modify: `index.html`
- Modify: `app-main.js`
- Modify: `styles.css`
- Modify: `tests/reconciliation-automation-ui.test.js`
- Modify: `tests/reconciliation-density.test.js`

**Interfaces:**
- Consumes: manual-rules GET, existing active-run restoration/continuation, selected-proposal execution, Settings rules, and shared reconciliation history.
- Produces:
  - `<select id="financial-reconciliation-workbench-automation-rule">`;
  - `<button id="financial-reconciliation-workbench-automation-analyze">Analyze</button>`;
  - `financialReconciliationAutomationOpenRun(run) -> boolean`;
  - `financialReconciliationAutomationRuleOptions(rules,selectedRuleKey) -> string`;
  - `openFinancialReconciliationAutomation() -> Promise<void>`;
  - no per-rule Analyze cards and no **Run batch now** UI.

- [ ] **Step 1: Write failing DOM and behavior tests**

Replace old run-batch/per-rule-card assertions with tests that require:

```js
assert.match(html, /id="financial-reconciliation-workbench-automation-rule"/);
assert.match(html, /id="financial-reconciliation-workbench-automation-analyze"/);
assert.match(html, /id="financial-reconciliation-automation-open-workbench"[^>]*>Open automatic reconciliation</);
assert.doesNotMatch(html, /financial-reconciliation-automation-run-batch-now|>Run batch now</);
```

Executable helper tests must prove:

- the LOV contains only rules with `enabled === true && allowManualExecution === true`, sorted by priority/key, with escaped labels and literal keys as values;
- the first eligible rule becomes the default only when no valid selection exists;
- a user selection survives rerenders and authoritative rule reloads while still eligible;
- Analyze POSTs exactly the selected key;
- the selector and Analyze button are disabled during `analyzing`, `ready`, or `running` unfinished runs;
- a terminal no-proposal run unlocks the selector;
- active-run restoration selects the run's rule and cannot silently switch it;
- opening from Settings performs only `setView('financial-reconciliation',{financialReconciliationTab:'automatic'})`, makes no API POST, and does not execute/select proposals;
- the Open button requires Reconciliation app access but is not blocked by an unsaved Settings draft; its hint explains that only saved rules appear;
- keyboard focus, visible labels, narrow-screen wrapping, and error/status live regions remain accessible;
- shared history stays outside both tab panels and remains visible in each tab.
- opening an automatic history record shows the persisted Credit Card rule key/version, trigger, configuration snapshot, and identity evidence.

- [ ] **Step 2: Run UI RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
```

Expected: FAIL because current markup renders one Analyze button per rule and Settings still dispatches `analyze_batch`.

- [ ] **Step 3: Replace the workbench rule cards with a labeled selector**

Use this semantic structure in `index.html`:

```html
<div class="financial-reconciliation-workbench-automation-rule-picker">
  <label for="financial-reconciliation-workbench-automation-rule">Rule</label>
  <select id="financial-reconciliation-workbench-automation-rule" aria-describedby="financial-reconciliation-workbench-automation-status"></select>
  <button id="financial-reconciliation-workbench-automation-analyze" type="button">Analyze</button>
</div>
```

Add `selectedRuleKey` to `state.financialReconciliation.automation`. Populate options from authoritative manual rules and escape both label and value. Define an open run as an existing run with no `finishedAt` and lifecycle `analyzing`, `ready`, or `running`. While open, lock the selector to `run.definitions[0].ruleKey`; otherwise permit a new selection.

Replace delegated per-card Analyze handling with a select `change` listener and one button listener. `analyzeFinancialReconciliationAutomationRule` still verifies the key exists in the enabled/manual rule array immediately before POST. Keep continuation, proposal selection, result rendering, and execution behavior unchanged.

- [ ] **Step 4: Replace Settings batch execution with navigation**

Change the Settings button ID/text to `financial-reconciliation-automation-open-workbench` / **Open automatic reconciliation** and its hint to `Opens the saved manual-enabled rules in Reconciliation. Unsaved changes are not applied.`

Implement:

```js
async function openFinancialReconciliationAutomation() {
  if (!canAppFinancialReconciliation()) {
    setReconciliationAutomationSettingsStatus("Reconciliation app access is required.", true);
    return;
  }
  await setView("financial-reconciliation", { financialReconciliationTab: "automatic" });
}
```

Delete `runReconciliationAutomationBatchNow`, its event listener, pending-analysis draft handoff, and all `analyze_batch` UI code. Fix the Settings max-days input to `max="90"`.

Rename Settings state/rendering from `lastScheduledRun` to `lastScheduledBatch`. Render the batch's aggregate status, completion time, and safe counts; do not select or display one child as if it represented the whole daily result.

Add compact responsive CSS for the picker without changing the established proposal-column density. Preserve visible focus and full-width controls on narrow screens.

- [ ] **Step 5: Run focused/full verification and commit Task 7**

Run:

```powershell
node --check app-main.js
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
node --test --test-isolation=none
git diff --check
```

Expected: all pass.

```powershell
git add -- index.html app-main.js styles.css tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
git commit -m "feat: select one automatic reconciliation rule"
```

---

### Task 8: Migration order, end-to-end verification, and rollout evidence

**Files:**
- Modify: `README.md`
- Modify: `tests/reconciliation-automation.test.js`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`
- Verify: `vercel.json`
- Verify: all files changed in Tasks 1-7

**Interfaces:**
- Consumes: the complete implementation.
- Produces: documented migration order, reapply-safe rollout instructions, complete automated verification, and an explicit list of any live external gates.

- [ ] **Step 1: Add final cross-layer contract assertions**

Require `README.md` to list the new migration as step 8, after the 90-day migration:

```text
8. supabase-migrations/2026-08-16-financial-reconciliation-automation-credit-card-rule.sql
```

Add final source-contract checks that every new RPC is included in revoke/grant assertions, the smoke includes the migration once in normal order plus one explicit reapply, and no production file contains `analyze_batch`, `Run batch now`, direct extension-schema calls, or a user-controlled SQL/table/function dispatch.

- [ ] **Step 2: Run every static and Node verification command**

Run:

```powershell
node --check api/_reconciliation-automation.js
node --check api/reconciliation-automation.js
node --check api/reconciliation-automation-settings.js
node --check api/reconciliation-automation-cron.js
node --check app-main.js
node --test --test-isolation=none tests/reconciliation-automation.test.js tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
node --test --test-isolation=none
Get-Content -Raw vercel.json | ConvertFrom-Json | Out-Null
git diff --check
```

Expected: every command exits 0 with no failed, skipped, cancelled, or todo tests.

- [ ] **Step 3: Run the transactional PostgreSQL smoke or record the exact gate**

If available:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Expected: transaction completes all assertions and rolls back cleanly. Re-run once to prove migration idempotency. If `psql` or the URL is absent, report both facts and do not mark database verification complete.

- [ ] **Step 4: Perform authenticated browser verification**

Against a database with the migration applied, verify desktop and narrow layouts:

1. Manual reconciliation remains the default tab.
2. Automatic reconciliation LOV lists only saved enabled/manual rules.
3. Selecting Credit Card analyzes only exact-`Visa` bases.
4. Progress resumes after tab change/reload.
5. Proposed, ambiguous, completed, stale, and failed rows render with correct card source details/evidence; no-match rows stay hidden while counts remain accurate.
6. The selector stays locked until the run is terminal.
7. Settings displays immutable Credit Card logic, editable tolerance/days/flags/priority, and defaults disabled/0/10/priority 2.
8. **Open automatic reconciliation** navigates without starting analysis.
9. A test daily batch runs Bank Statement then Credit Card as separate history entries.
10. A deliberately failed first child does not block the second.

Record console/network errors and exact screenshots if any scenario fails. If no authenticated fixture exists, list these ten scenarios as the mandatory live gate.

- [ ] **Step 5: Update README and commit Task 8**

Document migration order, SQL smoke command, disabled-by-default rollout, manual-first enablement, and scheduled two-rule validation. Then:

```powershell
git add -- README.md tests/reconciliation-automation.test.js tests/reconciliation-automation-rpc.smoke.sql
git commit -m "docs: verify credit card reconciliation rollout"
```

- [ ] **Step 6: Request final review before integration**

Resolve the implementation base once and inspect the complete range:

```powershell
$reconciliationImplementationBase = & 'C:\Program Files\Git\cmd\git.exe' merge-base HEAD main
& 'C:\Program Files\Git\cmd\git.exe' status --short
& 'C:\Program Files\Git\cmd\git.exe' log --oneline "$reconciliationImplementationBase..HEAD"
& 'C:\Program Files\Git\cmd\git.exe' diff --check "$reconciliationImplementationBase..HEAD"
```

Request independent spec-compliance and code/security reviews of that range. Resolve every Critical or Important finding under RED/GREEN TDD, rerun all verification, and only then use `superpowers:finishing-a-development-branch` to offer local merge/publish choices.
