# Automatic Reconciliation Proposal Details Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Enrich amount-only automatic reconciliation proposals so Financial Documents show document, supplier, supplier NIF, and description details while Bank Statement and Credit Card destinations show their descriptions, including an audit-safe backfill of unfinished runs.

**Architecture:** Keep the existing three-column proposal renderer and extend its Financial Documents metadata behavior. Add one forward-only Supabase migration that replaces only the two amount-only candidate adapters with richer JSON snapshots and atomically enriches stored proposals belonging to unfinished runs; completed proposal and reconciliation audit snapshots stay unchanged.

**Tech Stack:** PostgreSQL/Supabase RPC and JSONB, browser JavaScript, Node's built-in test runner, transactional PostgreSQL smoke tests, Git.

## Global Constraints

- Apply this change only to `financial_documents_cgd_bank_statement_amount_only` version `1` and `financial_documents_cgd_credit_card_amount_only` version `1`.
- Preserve both identity-based automatic reconciliation rules unchanged.
- Financial Documents base snapshots must add `docNumber`, `description`, `supplierName`, and `supplierNif` using the same camelCase JSON contract as existing identity rules.
- CGD Bank Statement destination and candidate snapshots must add `description` from `import_cgd_extrato_ordem.descritivo`.
- CGD Credit Card destination and candidate snapshots must add `description` from `import_cgd_cartao_credito.descricao`.
- Show supplier name and supplier NIF only for Financial Documents. Never infer or display a supplier for Bank Statement or Credit Card records.
- Backfill only proposals whose parent automatic run has `finished_at is null`.
- Backfill `base_snapshot`, normal destination `items`, flat candidate-limit evidence, and nested ambiguous `candidate_groups` without changing JSON array order.
- Preserve proposal IDs, run IDs, lifecycle statuses, amounts, evidence, signatures, configuration snapshots, reconciliation links, and timestamps except for the proposal `updated_at` written by the enrichment update.
- Do not modify completed-run proposal snapshots or completed reconciliation audit metadata.
- If a referenced source row no longer exists, leave that individual JSON snapshot unchanged; do not invent data.
- Add one dated forward migration after the 2026-08-17 amount-only migration. Do not edit prior migrations.
- Preserve existing SQL function signatures, literal rule dispatch, `SECURITY DEFINER SET search_path = public, pg_temp`, ownership, revokes, and `service_role` grants.
- The migration must be transactional, idempotent, and safe to apply twice.
- Do not add Node tests that read SQL source and search for declarations or prose. PostgreSQL behavior belongs in the transactional SQL smoke.
- Implement production behavior test-first. Do not claim PostgreSQL smoke success unless it actually runs against a database.

---

### Task 1: Render Financial Document supplier details only on the base record

**Files:**
- Modify: `app-main.js:22239-22259`
- Modify: `tests/reconciliation-automation-ui.test.js:1728-1850`

**Interfaces:**
- Consumes: `financialReconciliationAutomationItemMarkup(item, label, operator)` and the existing `baseSnapshot`, `items`, and `candidateGroups` proposal contract.
- Produces: a Financial Documents-only metadata contract that renders `Supplier <name>` and `Supplier NIF <nif>`; non-Financial-Document records ignore supplier-shaped fields while retaining date, description, amount, evidence, and record ID.

- [ ] **Step 1: Add a failing amount-only proposal rendering test**

In `tests/reconciliation-automation-ui.test.js`, add a focused test using the actual extracted production renderer. Use an amount-only proposal shaped like:

```js
const proposal = {
  id: WORKBENCH_PROPOSAL_1,
  ruleKey: "financial_documents_cgd_bank_statement_amount_only",
  ruleVersion: 1,
  status: "proposed",
  baseSnapshot: {
    sourceType: "financial_documents",
    sourceId: "document-details",
    sourceDate: "2026-01-13",
    docNumber: "FT <2026/17>",
    description: "Base <description>",
    supplierName: "Supplier & Sons",
    supplierNif: "PT<500123456>",
    amount: 12.5,
  },
  items: [{
    sourceType: "import_cgd_extrato_ordem",
    sourceId: "bank-details",
    sourceDate: "2026-01-13",
    description: "Bank <description>",
    supplierName: "Must not display",
    supplierNif: "Must not display NIF",
    amount: -12.5,
    evidence: {},
  }],
  candidateGroups: [],
  calculatedDifference: 0,
  allowedDifference: 0,
};
```

Assert that the rendered markup:

```js
assert.match(markup, /Document FT &lt;2026\/17&gt;/);
assert.match(markup, /Supplier Supplier &amp; Sons/);
assert.match(markup, /Supplier NIF PT&lt;500123456&gt;/);
assert.match(markup, /Base &lt;description&gt;/);
assert.match(markup, /Bank &lt;description&gt;/);
assert.doesNotMatch(markup, /Must not display/);
assert.doesNotMatch(markup, /<description>/);
```

Also render an ambiguous proposal with two nested candidate groups and assert every destination description stays inside the destination-column markup.

- [ ] **Step 2: Run the focused test and capture RED**

Run:

```powershell
node --test --test-isolation=none --test-name-pattern="amount-only proposal details" tests/reconciliation-automation-ui.test.js
```

Expected: FAIL because the renderer does not show `supplierNif` and currently accepts supplier-shaped fields from any source type.

- [ ] **Step 3: Implement the Financial Documents-only metadata guard**

In `financialReconciliationAutomationItemMarkup`, normalize the source before the metadata array:

```js
const sourceType = clean(value.sourceType);
const isFinancialDocument = sourceType === "financial_documents";
const supplier = isFinancialDocument ? clean(value.supplierName ?? value.supplier) : "";
const supplierNif = isFinancialDocument ? clean(value.supplierNif ?? value.supplier_nif) : "";
const documentNumber = isFinancialDocument ? clean(value.docNumber ?? value.documentNumber) : "";
```

Build metadata in this exact order:

```js
const meta = [
  formatDateOnly(value.sourceDate) || "-",
  documentNumber ? `Document ${documentNumber}` : "",
  supplier ? `Supplier ${supplier}` : "",
  supplierNif ? `Supplier NIF ${supplierNif}` : "",
].filter(Boolean);
```

Use `sourceType` when resolving the friendly source label. Keep the existing description, evidence, amount, and record-ID markup unchanged.

- [ ] **Step 4: Verify the UI task and commit**

Run:

```powershell
node --check app-main.js
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
node --test --test-isolation=none
git diff --check
```

Expected: every command exits `0`; the focused proposal test proves the new field, source guard, escaping, and destination-column behavior.

Commit:

```powershell
git add -- app-main.js tests/reconciliation-automation-ui.test.js
git commit -m "feat: show automatic proposal record details"
```

---

### Task 2: Enrich future amount-only snapshots and backfill unfinished runs

**Files:**
- Create: `supabase-migrations/2026-08-18-financial-reconciliation-automation-proposal-details.sql`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`

**Interfaces:**
- Consumes:
  - `financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])`;
  - `financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])`;
  - `financial_reconciliation_automatic_runs`, `financial_reconciliation_automatic_proposals`, and their existing JSONB snapshot contract.
- Produces:
  - the same two function signatures returning richer `base_snapshot` and `candidates` JSON;
  - an idempotent unfinished-run backfill over `base_snapshot`, `items`, and `candidate_groups`;
  - unchanged completed proposals and audit rows.

- [ ] **Step 1: Add transactional smoke fixtures before creating the migration**

In `tests/reconciliation-automation-rpc.smoke.sql`, add a section after migration 9 is installed but before migration 10 is applied. Create source fixtures containing unmistakable escaped-text values:

```sql
-- Financial Document fields
doc_number = 'FT <DETAIL/1>'
description = 'Base & detail'
supplier_name = 'Supplier <Detail>'
supplier_nif = 'PT500000001'

-- Destination descriptions
descritivo = 'Bank <detail>'
descricao = 'Card & detail'
```

Insert:

1. one unfinished amount-only run with a proposed Bank item using the legacy minimal snapshot;
2. one unfinished amount-only run with a flat `candidate_limit` `candidate_groups` array;
3. one unfinished amount-only run with nested `multiple_combinations` candidate groups;
4. one unfinished Credit Card proposal;
5. one completed run and proposal using the same legacy minimal shape;
6. one completed reconciliation audit row whose metadata contains the completed proposal snapshots;
7. one unfinished proposal whose base or destination source ID no longer exists.

Save the completed proposal JSON and audit metadata into a temporary comparison table before applying migration 10.

Add `\ir` for the new migration twice, then assert with `DO` blocks that:

```sql
-- Unfinished base snapshot
base_snapshot->>'docNumber' = 'FT <DETAIL/1>'
base_snapshot->>'description' = 'Base & detail'
base_snapshot->>'supplierName' = 'Supplier <Detail>'
base_snapshot->>'supplierNif' = 'PT500000001'

-- Destination snapshots
items->0->>'description' = 'Bank <detail>'
candidate_groups->0->>'description' = 'Bank <detail>'              -- flat
candidate_groups->0->0->>'description' = 'Bank <detail>'           -- nested
items->0->>'description' = 'Card & detail'                          -- card fixture
```

Assert all unrelated JSON keys and evidence still equal the saved pre-migration values. Assert the completed proposal JSON and audit metadata are `IS NOT DISTINCT FROM` their saved copies. Assert the missing-source snapshot is unchanged.

Finally, call both amount-only candidate functions directly and prove newly generated snapshots contain the same detail fields. Execute the backfilled unique Bank and Credit Card proposals through `execute_financial_reconciliation_automatic_proposal` and assert they complete rather than becoming stale.

- [ ] **Step 2: Run PostgreSQL smoke and capture RED when the database gate is available**

Run:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Expected before implementation: FAIL at the missing migration include or at the first enriched-snapshot assertion. If `psql` or `SUPABASE_DB_URL` is unavailable, record this exact command as an external RED/GREEN gate; do not replace it with a Node test that scans SQL text.

- [ ] **Step 3: Create the forward-only migration and enrich future snapshots**

Create `supabase-migrations/2026-08-18-financial-reconciliation-automation-proposal-details.sql` with `BEGIN`/`COMMIT` and no edits to migration 9.

Replace the Bank amount-only adapter using its existing signature and eligibility predicates. Extend its `bases` CTE with:

```sql
document.doc_number,
document.description,
document.supplier_name,
document.supplier_nif
```

Build the base snapshot as:

```sql
jsonb_build_object(
  'sourceType', 'financial_documents',
  'sourceId', base.id,
  'sourceDate', base.document_date,
  'amount', base.amount,
  'docNumber', base.doc_number,
  'description', base.description,
  'supplierName', base.supplier_name,
  'supplierNif', base.supplier_nif
)
```

Select `bank.descritivo AS description` in the destination lateral query and add `'description', description` to each Bank candidate JSON object without changing evidence, ordering, the 12-item display bound, or unbounded candidate count.

Apply the same base snapshot contract to the Credit Card adapter. Select `card.descricao AS description` and add it to every Credit Card candidate JSON object. Preserve exact payment, amount/date, lock, and rule/version predicates byte-for-byte in behavior.

- [ ] **Step 4: Backfill only unfinished amount-only proposals**

Within the same transaction, update only proposals satisfying:

```sql
proposal.rule_key in (
  'financial_documents_cgd_bank_statement_amount_only',
  'financial_documents_cgd_credit_card_amount_only'
)
and proposal.rule_version = 1
and exists (
  select 1
  from public.financial_reconciliation_automatic_runs run
  where run.id = proposal.run_id
    and run.finished_at is null
)
```

Use ordered `jsonb_array_elements(... ) WITH ORDINALITY` reconstruction for `items` and `candidate_groups`. For each item:

- merge only `description` from the matching Bank or Credit Card source row;
- return the original JSON object unchanged when the source type is unsupported, the source ID is malformed, or the source row is absent;
- preserve `sourceType`, `sourceId`, date, amount, evidence, and every unknown future key.

For `candidate_groups`, preserve both supported structures:

```text
candidate_limit:        [item, item, ...]
multiple_combinations:  [[item, ...], [item, ...], ...]
```

Enrich `base_snapshot` only when `financial_documents.id::text = base_snapshot->>'sourceId'` resolves. Merge the four display keys without rebuilding or dropping any existing key. Do not update rows from completed runs. Set `proposal.updated_at = now()` only when at least one JSON document is actually different, making reapplication a no-op.

- [ ] **Step 5: Restore security and schema-cache contracts**

Reapply the existing adapter ACLs:

```sql
revoke all on function public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  from public, anon, authenticated;
revoke all on function public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  from public, anon, authenticated;
grant execute on function public.financial_reconciliation_automatic_bank_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  to service_role;
grant execute on function public.financial_reconciliation_automatic_credit_card_amount_only_candidates_for_base_ids(text,integer,numeric,integer,uuid[])
  to service_role;
```

Retain `SECURITY DEFINER SET search_path = public, pg_temp` on both adapters and finish with:

```sql
notify pgrst, 'reload schema';
```

- [ ] **Step 6: Verify the database task and commit**

Run the transactional smoke if the database gate is available, then always run the local regression suite:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
node --test --test-isolation=none
git diff --check
```

Expected: all available commands exit `0`. The SQL smoke must pass twice before production rollout; unavailable database access remains an explicit external gate.

Commit:

```powershell
git add -- supabase-migrations/2026-08-18-financial-reconciliation-automation-proposal-details.sql tests/reconciliation-automation-rpc.smoke.sql
git commit -m "feat: enrich automatic proposal snapshots"
```

---

### Task 3: Document migration order and complete cross-layer verification

**Files:**
- Modify: `README.md:42-100`

**Interfaces:**
- Consumes: the completed UI and migration tasks.
- Produces: migration 10 rollout instructions, explicit unfinished-only backfill behavior, and final verification evidence.

- [ ] **Step 1: Update the automatic reconciliation migration order**

Append this exact entry after migration 9:

```markdown
10. `supabase-migrations/2026-08-18-financial-reconciliation-automation-proposal-details.sql`
```

State that installations current through migration 9 apply only migration 10. Add the manual apply/reapply commands:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f supabase-migrations/2026-08-18-financial-reconciliation-automation-proposal-details.sql
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f supabase-migrations/2026-08-18-financial-reconciliation-automation-proposal-details.sql
```

Document that migration 10 enriches unfinished amount-only proposals immediately, leaves completed history/audit snapshots unchanged, and does not infer supplier data for Bank Statement or Credit Card records. Do not imply that Vercel applies the database migration automatically.

- [ ] **Step 2: Run complete local verification**

Run:

```powershell
node --check app-main.js
node --check api/_reconciliation-automation.js
node --check api/reconciliation-automation.js
node --check api/reconciliation-automation-settings.js
node --check api/reconciliation-automation-cron.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
node --test --test-isolation=none
node -e "JSON.parse(require('fs').readFileSync('vercel.json','utf8')); console.log('vercel.json valid')"
git diff --check
git status --short
```

Expected: syntax, tests, JSON, and diff checks exit `0`; status contains only intended feature changes plus pre-existing unrelated user files.

- [ ] **Step 3: Run mandatory rollout gates**

Database:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Authenticated browser scenarios:

1. analyze the Bank amount-only rule and confirm base document number, supplier name, supplier NIF, and base description display in column two;
2. confirm every Bank destination and ambiguous candidate description displays in column three with no supplier label;
3. repeat the same checks for the Credit Card amount-only rule;
4. open an unfinished pre-migration run and confirm it is enriched immediately after migration 10;
5. open completed history and confirm its prior audit snapshot is unchanged;
6. execute a backfilled unfinished unique proposal and confirm it completes rather than becoming stale;
7. verify desktop three-column and narrow stacked layouts with escaped punctuation in every detail field.

If database credentials or authenticated browser state are unavailable, report those exact checks as external rollout gates rather than claiming success.

- [ ] **Step 4: Request final review and commit the release documentation**

Use `superpowers:requesting-code-review` for a full review against:

`docs/superpowers/specs/2026-08-18-automatic-reconciliation-proposal-details-design.md`

Resolve every Critical or Important finding with test-first changes and rerun all affected gates. Then commit:

```powershell
git add -- README.md
git commit -m "docs: release automatic proposal details"
```

Before merge or publication, use `superpowers:verification-before-completion` and `superpowers:finishing-a-development-branch`. Publish compatible application code first, apply migration 10 manually second, run the SQL smoke/reapply third, and verify the unfinished-run backfill before continuing automatic execution.
