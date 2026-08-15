# Automatic Reconciliation Analysis Performance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the Financial Documents to CGD Bank Statement analysis finish within Supabase's statement timeout without changing any matching rule or public RPC contract.

**Architecture:** Add one forward-only, idempotent migration that supplies usable date indexes and replaces `financial_reconciliation_automatic_rule_candidates` with materialized normalization and scoring stages. Keep the API/UI untouched and pin the migration contract with Node source tests plus the existing transactional SQL smoke suite.

**Tech Stack:** PostgreSQL 14/Supabase, PL/pgSQL/SQL functions, PostgREST RPCs, Node.js `node:test`.

## Global Constraints

- Preserve the Financial Documents to CGD Bank Statement direction.
- Preserve configurable maximum date difference, currently seven days.
- Preserve configurable allowed amount difference, currently zero euros.
- Preserve the document-number, description-similarity `0.60`, and supplier-similarity `0.70` identity signals.
- Preserve record locks and the `2026-01-01` eligibility floor.
- Preserve deterministic proposal ordering, evidence JSON, function signature, security-definer search path, and service-role-only execution.
- Do not increase `statement_timeout`.
- Do not modify UI, API, rule configuration, or execution behavior.

---

### Task 1: Add the optimized candidate migration under TDD

**Files:**
- Create: `supabase-migrations/2026-08-15-financial-reconciliation-automation-analysis-performance.sql`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Consumes: `public.financial_reconciliation_match_compact(text)`, `public.financial_reconciliation_match_normalize(text)`, `public.financial_reconciliation_extension_similarity(text,text)`, and `public.financial_reconciliation_extension_word_similarity(text,text)` from `2026-08-14-financial-reconciliation-automation-analysis.sql`.
- Produces: the unchanged RPC helper signature `public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)` with columns `base_source_id uuid`, `base_source_date date`, `base_snapshot jsonb`, `candidates jsonb`, and `candidate_count integer`.

- [ ] **Step 1: Add the failing source-contract test**

Add `ANALYSIS_PERFORMANCE_MIGRATION_PATH` beside the other migration constants and a test named `automation performance migration materializes matching work without changing rule semantics`. Assert that the new file:

```js
assert.match(migration, /create index[^;]+financial_documents[^;]+\(document_date\)/i);
assert.match(migration, /create index[^;]+import_cgd_extrato_ordem[^;]+\(data\)/i);
assert.match(migration, /create or replace function public\.financial_reconciliation_automatic_rule_candidates\(/);
for (const stage of ["bases", "bank_rows", "qualified", "scored"]) {
  assert.match(migration, new RegExp(`${stage}\\s+as\\s+materialized`, "i"));
}
assert.match(migration, /b\.data between d\.document_date - p_max_difference_days and d\.document_date \+ p_max_difference_days/);
assert.match(migration, /d\.document_date >= date '2026-01-01'/);
assert.match(migration, /b\.data >= date '2026-01-01'/);
assert.match(migration, /description_score >= 0\.60/);
assert.match(migration, /supplier_score >= 0\.70/);
assert.match(migration, /order by base_date, base_id/);
assert.match(migration, /security definer set search_path = public, pg_temp/);
assert.match(migration, /revoke all on function public\.financial_reconciliation_automatic_rule_candidates\(text,integer,numeric,integer\) from public, anon, authenticated;/);
assert.match(migration, /grant execute on function public\.financial_reconciliation_automatic_rule_candidates\(text,integer,numeric,integer\) to service_role;/);
assert.match(migration, /notify pgrst, 'reload schema';/);
assert.doesNotMatch(migration, /statement_timeout/i);
```

- [ ] **Step 2: Run the focused test and capture RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL because `2026-08-15-financial-reconciliation-automation-analysis-performance.sql` does not exist.

- [ ] **Step 3: Create the minimal idempotent migration**

Create date indexes only when a valid, ready, non-partial B-tree index does not already lead with the required date column. Then replace the function using the existing return contract and these explicit stages:

```sql
with
bases as materialized (
  select d.id, d.document_date, d.doc_number, d.description, d.supplier_name, d.amount,
         public.financial_reconciliation_match_compact(d.doc_number) as compact_document_number,
         public.financial_reconciliation_match_normalize(d.description) as normalized_document_description,
         public.financial_reconciliation_match_normalize(d.supplier_name) as normalized_supplier_name
  from public.financial_documents d
  where p_rule_key = 'financial_documents_cgd_bank_statement'
    and p_rule_version = 1
    and d.fat = 'S'
    and d.document_date >= date '2026-01-01'
    and not exists (
      select 1 from public.financial_reconciliation_items i
      where i.source_type = 'financial_documents' and i.source_id = d.id
    )
),
bank_rows as materialized (
  select b.id, b.data, b.montante, b.descritivo,
         public.financial_reconciliation_match_normalize(b.descritivo) as normalized_bank_description
  from public.import_cgd_extrato_ordem b
  where b.data >= date '2026-01-01'
    and b.montante is not null
    and not exists (
      select 1 from public.financial_reconciliation_items i
      where i.source_type = 'import_cgd_extrato_ordem' and i.source_id = b.id
    )
),
qualified as materialized (
  select
    d.id as base_id,
    d.document_date as base_date,
    jsonb_build_object(
      'sourceType', 'financial_documents', 'sourceId', d.id,
      'sourceDate', d.document_date, 'amount', d.amount,
      'docNumber', d.doc_number, 'description', d.description,
      'supplierName', d.supplier_name
    ) as base_snapshot,
    b.id as source_id, b.data as source_date, b.montante as amount,
    b.descritivo as description, b.normalized_bank_description,
    d.compact_document_number, d.normalized_document_description,
    d.normalized_supplier_name
  from bases d
  left join bank_rows b
    on b.data between d.document_date - p_max_difference_days
                  and d.document_date + p_max_difference_days
),
scored as materialized (
  select q.*,
    coalesce(
      char_length(q.compact_document_number) >= 4
      and q.source_id is not null
      and position(q.compact_document_number in public.financial_reconciliation_match_compact(q.description)) > 0,
      false
    ) as document_number_matched,
    case
      when nullif(q.normalized_document_description, '') is null
        or nullif(q.normalized_bank_description, '') is null then 0::real
      else public.financial_reconciliation_extension_similarity(
        q.normalized_document_description, q.normalized_bank_description
      )
    end as description_score,
    case
      when nullif(q.normalized_supplier_name, '') is null
        or nullif(q.normalized_bank_description, '') is null then 0::real
      else public.financial_reconciliation_extension_word_similarity(
        q.normalized_supplier_name, q.normalized_bank_description
      )
    end as supplier_score
  from qualified q
),
identity_candidates as materialized (
  select *,
    document_number_matched
      or description_score >= 0.60
      or supplier_score >= 0.70 as identity_matched
  from scored
),
grouped as (
  select
    base_id, base_date, base_snapshot,
    coalesce(jsonb_agg(jsonb_build_object(
      'sourceType', 'import_cgd_extrato_ordem',
      'sourceId', source_id,
      'sourceDate', source_date,
      'amount', amount,
      'description', description,
      'evidence', jsonb_build_object(
        'documentNumber', jsonb_build_object(
          'matched', document_number_matched,
          'normalized', compact_document_number
        ),
        'description', jsonb_build_object(
          'matched', description_score >= 0.60,
          'score', description_score,
          'threshold', 0.60
        ),
        'supplier', jsonb_build_object(
          'matched', supplier_score >= 0.70,
          'score', supplier_score,
          'threshold', 0.70
        )
      )
    ) order by source_date, source_id) filter (where identity_matched), '[]'::jsonb) as candidates,
    count(*) filter (where identity_matched)::integer as candidate_count
  from identity_candidates
  group by base_id, base_date, base_snapshot
)
select base_id, base_date, base_snapshot, candidates, candidate_count
from grouped
order by base_date, base_id
```

End the migration with the exact revoke/grant statements asserted in Step 1 and `notify pgrst, 'reload schema';`.

- [ ] **Step 4: Run focused GREEN verification**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: all reconciliation automation tests pass.

- [ ] **Step 5: Commit Task 1**

```powershell
git add -- tests/reconciliation-automation.test.js supabase-migrations/2026-08-15-financial-reconciliation-automation-analysis-performance.sql
git commit -m "fix: optimize automatic reconciliation analysis"
```

---

### Task 2: Extend the transactional smoke contract and verify the branch

**Files:**
- Modify: `tests/reconciliation-automation-rpc.smoke.sql`
- Modify: `tests/reconciliation-automation.test.js`

**Interfaces:**
- Consumes: the Task 1 migration and the existing automation schema/analysis/execution migration chain.
- Produces: a reapply-safe SQL smoke path that verifies the optimized function definition and privileges.

- [ ] **Step 1: Add failing smoke-source assertions**

Extend the Node migration test to require the SQL smoke file to include the new migration after the original analysis migration and to assert the installed definition contains each materialized stage:

```js
assert.match(smokeSql, /2026-08-14-financial-reconciliation-automation-analysis\.sql[\s\S]*2026-08-15-financial-reconciliation-automation-analysis-performance\.sql/);
assert.match(smokeSql, /pg_get_functiondef\('public\.financial_reconciliation_automatic_rule_candidates\(text,integer,numeric,integer\)'::regprocedure\)/);
assert.match(smokeSql, /bases\s+as\s+materialized/i);
assert.match(smokeSql, /bank_rows\s+as\s+materialized/i);
assert.match(smokeSql, /qualified\s+as\s+materialized/i);
assert.match(smokeSql, /scored\s+as\s+materialized/i);
```

- [ ] **Step 2: Run the focused test and capture RED**

Run the same focused Node command. Expected: FAIL because the smoke file does not apply or inspect the performance migration.

- [ ] **Step 3: Update the SQL smoke transaction**

Add:

```sql
\ir ../supabase-migrations/2026-08-15-financial-reconciliation-automation-analysis-performance.sql
```

immediately after every analysis-migration include. Add a `do` block that loads `pg_get_functiondef('public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)'::regprocedure)` and raises unless all four `as materialized` stages are present, service role retains execute, and `anon`/`authenticated` do not.

- [ ] **Step 4: Run all local verification**

Run:

```powershell
node --check api/reconciliation-automation.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
node --test --test-isolation=none tests/*.test.js
git diff --check
```

Expected: syntax passes, focused and full suites have zero failures, and the diff check is clean. If `psql` and `SUPABASE_DB_URL` are available, also run `psql "$env:SUPABASE_DB_URL" -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql`; otherwise record that as an external gate.

- [ ] **Step 5: Commit Task 2**

```powershell
git add -- tests/reconciliation-automation.test.js tests/reconciliation-automation-rpc.smoke.sql
git commit -m "test: cover reconciliation analysis performance migration"
```

---

### Task 3: Merge, publish, and validate production

**Files:**
- No new repository files.

**Interfaces:**
- Consumes: the verified migration branch.
- Produces: published repository changes plus an explicit Supabase application/verification handoff.

- [ ] **Step 1: Merge the feature branch into local `main`**

Use a fast-forward-only merge after confirming both worktrees are clean. Preserve all unrelated untracked files.

- [ ] **Step 2: Verify the merged tree**

Run the full Node suite and `node --check api/reconciliation-automation.js` from `main`. Stop without pushing if either fails.

- [ ] **Step 3: Push `main`**

```powershell
git push origin main
```

- [ ] **Step 4: Apply the migration in Supabase**

Run `supabase-migrations/2026-08-15-financial-reconciliation-automation-analysis-performance.sql` once in the Supabase SQL editor. This is the only database script required for this fix.

- [ ] **Step 5: Verify the original production symptom**

Call `financial_reconciliation_automatic_rule_candidates` read-only with rule version `1`, difference allowed `0`, and maximum difference days `7`. Require HTTP 200 and elapsed time below the active statement timeout. Then press **Analyze** and require a rendered run containing proposed, ambiguous, or skipped records rather than `Unexpected server error.`

- [ ] **Step 6: Roll back if needed**

If production verification regresses matching results, re-run the `create or replace function public.financial_reconciliation_automatic_rule_candidates(...)` definition from `2026-08-14-financial-reconciliation-automation-analysis.sql`; the indexes may remain because they do not change results.
