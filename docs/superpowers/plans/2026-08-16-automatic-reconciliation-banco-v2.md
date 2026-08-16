# Automatic Reconciliation Banco Rule Version 2 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Restrict the managed Financial Documents to CGD Bank Statement automatic rule to exact `Payment = 'Banco'`, migrate it to immutable version 2, and present only relevant proposals in a compact paired-card review layout.

**Architecture:** Add one idempotent Supabase migration that seeds rule version 2, preserves the administrator-owned configuration, patches the latest index-driven candidate implementation, and updates execution revalidation without changing RPC signatures. Keep the existing API boundary and persisted run model, and add one pure browser-side selector that controls which proposal rows are rendered while summary counts continue to use the complete run. Refactor only proposal markup and CSS for the approved compact paired-card presentation.

**Tech Stack:** PostgreSQL/Supabase migrations and PL/pgSQL, CommonJS Node.js API helpers, vanilla JavaScript, HTML/CSS, Node's built-in test runner, SQL smoke tests.

## Global Constraints

- The stable managed rule key remains exactly `financial_documents_cgd_bank_statement`.
- Rule version `1` definitions, completed reconciliations, snapshots, evidence, and audit history remain unchanged.
- Rule version `2` retains the existing date floor, `fat = 'S'`, locking, identity thresholds, ambiguity rules, amount/operator behavior, candidate cap, and indexed bank-date lookup.
- The additional base predicate is exact PostgreSQL text equality: `d.payment = 'Banco'`; do not trim or normalize case.
- `BANCO`, ` banco `, blank, `NULL`, and every other Payment value are completely excluded before candidate processing and create no proposal or skipped row.
- Existing administrator values for enabled state, manual execution, scheduled batch, difference allowance, maximum date difference, and priority move to version 2 unchanged.
- Pending version-1 proposals become `stale` with reason `rule_version_changed`; they never create a reconciliation.
- Active runs show only `proposed` and `ambiguous` rows; finished runs show only `completed`, `stale`, and `failed` rows.
- Unchecked `proposed` rows remain visible and reselectable before execution.
- Hidden `skipped` and `deselected` proposals remain persisted and included in aggregate counts.
- Keep the existing API and RPC signatures, authorization boundaries, RLS, security-definer search paths, service-role grants, and error mapping.
- Do not add direct browser database access, a configurable Payment predicate, a generic expression engine, another rule key, or another API endpoint.
- All source-derived markup must be escaped; long values must wrap; keyboard focus and checkbox accessible names must remain intact.
- Use strict RED/GREEN TDD for every production change and commit each independently reviewable task.
- Treat `docs/superpowers/specs/2026-08-16-automatic-reconciliation-banco-v2-design.md` as the authoritative behavior and rollout specification.

---

## File Map

- Create `supabase-migrations/2026-08-16-financial-reconciliation-automation-banco-v2.sql`: immutable version-2 catalog seed, configuration migration, latest candidate-function installation, guarded execution-function upgrade, privileges, and schema reload.
- Modify `api/_reconciliation-automation.js`: make version 2 the only editable/current managed rule version while leaving recursive historical result mapping unchanged.
- Modify `tests/reconciliation-automation.test.js`: prove version-2 settings behavior through the real API normalization boundary.
- Modify `tests/reconciliation-automation-rpc.smoke.sql`: executable idempotency, exact-Payment eligibility, no-row exclusion, stale-version, Payment-drift, security, and legacy-regression coverage.
- Modify `app-main.js`: pure proposal-visibility selector, compact proposal/item markup, and rendering integration.
- Modify `styles.css`: compact paired-card density, safe wrapping, focus behavior, and narrow-screen stacking.
- Modify `tests/reconciliation-automation-ui.test.js`: executable selector/rendering/escaping behavior tests.
- Modify `tests/reconciliation-density.test.js`: CSS density and responsive-layout source contracts.

## Task 1: Install and Validate Managed Rule Version 2

**Files:**
- Create: `supabase-migrations/2026-08-16-financial-reconciliation-automation-banco-v2.sql`
- Modify: `api/_reconciliation-automation.js:1-5`
- Modify: `tests/reconciliation-automation.test.js:1-65, 175-225, 1820-2070`
- Modify: `tests/reconciliation-automation-rpc.smoke.sql:after the current finalization checks and before rollback`

**Interfaces:**
- Consumes: `public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)`, `public.execute_financial_reconciliation_automatic_proposal(uuid,text)`, the existing rule/config tables, and the existing service-role RPC boundary.
- Produces: current managed catalog entry `(rule_key = 'financial_documents_cgd_bank_statement', version = 2)` and `AUTOMATIC_RULE_VERSION = 2` for settings validation.
- Preserves: all existing function signatures and the camelCase API response structure.

- [ ] **Step 1: Add a failing API behavior test**

Update the managed-settings expectation to `ruleVersion: 2`, change the invalid-version fixture to version `1`, and add this behavior test:

```js
test("automation accepts current version 2 settings and rejects legacy version 1 edits", () => {
  assert.equal(normalizeAutomationSettingsPayload(managedSettings()).rules[0].ruleVersion, 2);
  assert.throws(
    () => normalizeAutomationSettingsPayload({
      ...managedSettings(),
      rules: [{ ...managedSettings().rules[0], ruleVersion: 1 }],
    }),
    /rule version/i,
  );
});
```

Do not add tests that grep SQL migration text. Exact catalog, predicate, idempotency, stale behavior, indexed lookup, and privilege behavior belong to the executable PostgreSQL smoke contract in Step 7.

- [ ] **Step 2: Run the focused Node test and capture RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: FAIL because the editable API normalization boundary still accepts only version `1`.

- [ ] **Step 3: Change only the current editable API rule version**

In `api/_reconciliation-automation.js`, change:

```js
const AUTOMATIC_RULE_VERSION = 2;
```

Do not add version validation to `toAutomationPublicResult()`. Its recursive key mapper must continue returning historical version-1 run and reconciliation payloads unchanged. `normalizeRuleVersion()` remains strict so administrators cannot submit stale version-1 settings after rollout.

- [ ] **Step 4: Create the immutable version-2 definition and preserve configuration**

Build the expected definition once, insert it with `on conflict do nothing`, then verify the installed row is byte-for-byte equivalent to the expected managed metadata. Keep the definition insert and configuration switch inside the same `DO` statement so they succeed or fail atomically. Do not add an explicit `BEGIN`/`COMMIT` wrapper because the repository smoke suite includes migrations inside its own rollback transaction. Use this exact catalog payload:

```sql
do $migration$
declare
  v_definition jsonb := $json$
  {
    "baseSourceType": "financial_documents",
    "destinationSourceTypes": ["import_cgd_extrato_ordem"],
    "baseEligibility": {
      "payment": {
        "operator": "exact_text_equal",
        "value": "Banco",
        "caseSensitive": true,
        "trim": false
      }
    },
    "identityBranches": {
      "document_number": {"algorithm": "compact_containment"},
      "description_similarity": {"algorithm": "similarity"},
      "supplier_similarity": {"algorithm": "word_similarity"}
    },
    "documentNumberMinimumCompactLength": 4,
    "descriptionSimilarityThreshold": 0.60,
    "supplierWordSimilarityThreshold": 0.70,
    "maxDestinationRecords": 4,
    "maxIdentityCandidatesPerBase": 12
  }
  $json$::jsonb;
  v_logic text := 'Payment must equal exactly Banco. A bank candidate must match at least one of three OR identity branches: compact document-number containment, document-description similarity, or supplier-to-bank-description word similarity. A base record is executable only when exactly one complete destination combination is valid; multiple combinations are reported as ambiguous and are never selected automatically.';
begin
  insert into public.financial_reconciliation_automatic_rule_definitions (
    rule_key, version, display_name, base_source_type,
    destination_source_types, logic_description, definition
  ) values (
    'financial_documents_cgd_bank_statement',
    2,
    'Financial Documents to CGD Bank Statement',
    'financial_documents',
    '["import_cgd_extrato_ordem"]'::jsonb,
    v_logic,
    v_definition
  ) on conflict (rule_key, version) do nothing;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_definitions definition
    where definition.rule_key = 'financial_documents_cgd_bank_statement'
      and definition.version = 2
      and definition.display_name = 'Financial Documents to CGD Bank Statement'
      and definition.base_source_type = 'financial_documents'
      and definition.destination_source_types = '["import_cgd_extrato_ordem"]'::jsonb
      and definition.logic_description = v_logic
      and definition.definition = v_definition
  ) then
    raise exception 'Managed automatic reconciliation rule version 2 differs from the expected immutable definition.';
  end if;

  update public.financial_reconciliation_automatic_rule_configs
  set rule_version = 2,
      updated_at = now()
  where rule_key = 'financial_documents_cgd_bank_statement'
    and rule_version = 1;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_rule_configs
    where rule_key = 'financial_documents_cgd_bank_statement'
      and rule_version = 2
  ) then
    raise exception 'Managed automatic reconciliation configuration could not be moved to version 2.';
  end if;
end
$migration$;
```

This updates only `rule_version` and `updated_at`; it must not rewrite `enabled`, `allow_manual_execution`, `include_in_scheduled_batch`, `difference_allowed`, `max_difference_days`, `priority`, or `updated_by`.

- [ ] **Step 5: Install the latest index-driven candidate function with the exact Banco predicate**

Copy the complete `public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)` definition from `supabase-migrations/2026-08-15-financial-reconciliation-automation-candidate-index-lookup.sql` into the new migration. Preserve every CTE, score, ordering rule, lateral bank lookup, function attribute, and return column. The resulting `bases` predicate must be exactly:

```sql
    from public.financial_documents d
    where p_rule_key = 'financial_documents_cgd_bank_statement'
      and p_rule_version = 2
      and d.fat = 'S'
      and d.payment = 'Banco'
      and d.document_date >= date '2026-01-01'
      and not exists (
        select 1
        from public.financial_reconciliation_items i
        where i.source_type = 'financial_documents'
          and i.source_id = d.id
      )
```

Keep the indexed destination lookup in the lateral subquery:

```sql
    left join lateral (
      select
        bank.id,
        bank.data,
        bank.montante,
        bank.descritivo,
        public.financial_reconciliation_match_normalize(bank.descritivo) as normalized_bank_description
      from public.import_cgd_extrato_ordem bank
      where bank.data between d.document_date - p_max_difference_days and d.document_date + p_max_difference_days
        and bank.data >= date '2026-01-01'
        and bank.montante is not null
        and not exists (
          select 1
          from public.financial_reconciliation_items i
          where i.source_type = 'import_cgd_extrato_ordem'
            and i.source_id = bank.id
        )
    ) b on true
```

Reapply the current grants after replacing the function:

```sql
revoke all on function public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)
  from public, anon, authenticated;
grant execute on function public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)
  to service_role;
```

- [ ] **Step 6: Guardedly upgrade execution to version 2 without copying an older function**

Patch the installed function returned by `pg_get_functiondef`, so later execution hardening is retained. Require exactly the known old or new fragments and fail closed on any unexpected definition:

```sql
do $migration$
declare
  v_definition text;
  v_old_version text := 'or v_proposal.rule_version <> 1';
  v_new_version text := 'or v_proposal.rule_version <> 2';
  v_old_comment text := 'v_comment := ''Automatically completed by rule Financial Documents to CGD Bank Statement v1; difference ''';
  v_new_comment text := 'v_comment := ''Automatically completed by rule Financial Documents to CGD Bank Statement v'' || v_proposal.rule_version::text || ''; difference ''';
begin
  select pg_get_functiondef(
    'public.execute_financial_reconciliation_automatic_proposal(uuid,text)'::regprocedure
  ) into strict v_definition;

  if strpos(v_definition, v_old_version) > 0 then
    v_definition := replace(v_definition, v_old_version, v_new_version);
  elsif strpos(v_definition, v_new_version) = 0 then
    raise exception 'Unexpected automatic proposal execution version guard.';
  end if;

  if strpos(v_definition, v_old_comment) > 0 then
    v_definition := replace(v_definition, v_old_comment, v_new_comment);
  elsif strpos(v_definition, v_new_comment) = 0 then
    raise exception 'Unexpected automatic proposal execution comment definition.';
  end if;

  execute v_definition;
end
$migration$;

revoke all on function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
  from public, anon, authenticated;
grant execute on function public.execute_financial_reconciliation_automatic_proposal(uuid,text)
  to service_role;

notify pgrst, 'reload schema';
```

The unchanged execution ordering makes a version-1 proposal stale before snapshot/candidate processing. A version-2 proposal whose Payment changes away from `Banco` fails candidate reproduction and uses the existing `source_snapshot_changed` stale path.

- [ ] **Step 7: Add executable SQL smoke coverage after the legacy version-1 suite**

Immediately before the current final `rollback;`, preserve distinctive administrator values, apply the new migration twice, and assert that only the version reference changed:

```sql
update public.financial_reconciliation_automatic_rule_configs
set enabled = true,
    allow_manual_execution = true,
    include_in_scheduled_batch = false,
    difference_allowed = 4.56,
    max_difference_days = 11,
    priority = 1,
    updated_by = 'smoke:banco-v2'
where rule_key = 'financial_documents_cgd_bank_statement';

\ir ../supabase-migrations/2026-08-16-financial-reconciliation-automation-banco-v2.sql
\ir ../supabase-migrations/2026-08-16-financial-reconciliation-automation-banco-v2.sql

do $$
begin
  if not exists (
    select 1 from public.financial_reconciliation_automatic_rule_definitions
    where rule_key = 'financial_documents_cgd_bank_statement' and version = 1
  ) or not exists (
    select 1 from public.financial_reconciliation_automatic_rule_definitions
    where rule_key = 'financial_documents_cgd_bank_statement'
      and version = 2
      and definition#>>'{baseEligibility,payment,value}' = 'Banco'
      and (definition#>>'{baseEligibility,payment,caseSensitive}')::boolean
      and not (definition#>>'{baseEligibility,payment,trim}')::boolean
  ) then
    raise exception 'Managed Banco rule versions are invalid.';
  end if;

  if not exists (
    select 1 from public.financial_reconciliation_automatic_rule_configs
    where rule_key = 'financial_documents_cgd_bank_statement'
      and rule_version = 2
      and enabled
      and allow_manual_execution
      and not include_in_scheduled_batch
      and difference_allowed = 4.56
      and max_difference_days = 11
      and priority = 1
      and updated_by = 'smoke:banco-v2'
  ) then
    raise exception 'Version 2 migration changed administrator configuration.';
  end if;
end $$;
```

Add five Financial Documents on an isolated date with Payment values `Banco`, `BANCO`, ` banco `, empty string, and `NULL`, plus one matching bank row for each document number. Call the candidate function with version 2 and assert that the returned base-ID set contains exactly the `Banco` document. Then run the authoritative analysis RPC and prove that no proposal row of any status exists for the four excluded IDs:

```sql
do $$
declare
  v_exact_id uuid := '41000000-0000-0000-0000-000000000201';
  v_excluded_ids uuid[] := array[
    '41000000-0000-0000-0000-000000000202'::uuid,
    '41000000-0000-0000-0000-000000000203'::uuid,
    '41000000-0000-0000-0000-000000000204'::uuid,
    '41000000-0000-0000-0000-000000000205'::uuid
  ];
  v_candidate_ids uuid[];
  v_run jsonb;
  v_run_id uuid;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name,
    payment, amount, fat
  ) values
    (v_exact_id, date '2027-01-20', 'BANCO-V2-201', '', '', 'Banco', 101.00, 'S'),
    (v_excluded_ids[1], date '2027-01-20', 'BANCO-V2-202', '', '', 'BANCO', 102.00, 'S'),
    (v_excluded_ids[2], date '2027-01-20', 'BANCO-V2-203', '', '', ' banco ', 103.00, 'S'),
    (v_excluded_ids[3], date '2027-01-20', 'BANCO-V2-204', '', '', '', 104.00, 'S'),
    (v_excluded_ids[4], date '2027-01-20', 'BANCO-V2-205', '', '', null, 105.00, 'S');

  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values
    ('42000000-0000-0000-0000-000000000201', 'smoke-banco-v2', 'banco-v2-201', date '2027-01-20', 'Payment BANCOV2201', -101.00),
    ('42000000-0000-0000-0000-000000000202', 'smoke-banco-v2', 'banco-v2-202', date '2027-01-20', 'Payment BANCOV2202', -102.00),
    ('42000000-0000-0000-0000-000000000203', 'smoke-banco-v2', 'banco-v2-203', date '2027-01-20', 'Payment BANCOV2203', -103.00),
    ('42000000-0000-0000-0000-000000000204', 'smoke-banco-v2', 'banco-v2-204', date '2027-01-20', 'Payment BANCOV2204', -104.00),
    ('42000000-0000-0000-0000-000000000205', 'smoke-banco-v2', 'banco-v2-205', date '2027-01-20', 'Payment BANCOV2205', -105.00);

  select coalesce(array_agg(candidate.base_source_id order by candidate.base_source_id), '{}'::uuid[])
  into v_candidate_ids
  from public.financial_reconciliation_automatic_rule_candidates(
    'financial_documents_cgd_bank_statement', 2, 4.56, 11
  ) candidate
  where candidate.base_source_id = v_exact_id
     or candidate.base_source_id = any(v_excluded_ids);

  if v_candidate_ids is distinct from array[v_exact_id] then
    raise exception 'Exact Banco eligibility returned an unexpected base set: %', v_candidate_ids;
  end if;

  v_run := public.create_financial_reconciliation_automatic_analysis(
    array['financial_documents_cgd_bank_statement'],
    'manual_rule',
    'smoke:banco-v2',
    '43000000-0000-0000-0000-000000000201'
  );
  v_run_id := (v_run->>'runId')::uuid;

  if not exists (
    select 1
    from public.financial_reconciliation_automatic_proposals
    where run_id = v_run_id
      and base_source_id = v_exact_id
      and status = 'proposed'
  ) or exists (
    select 1
    from public.financial_reconciliation_automatic_proposals
    where run_id = v_run_id
      and base_source_id = any(v_excluded_ids)
  ) then
    raise exception 'Non-Banco bases created proposal or skipped rows.';
  end if;
end $$;
```

Use a separate exact-`Banco` proposal to prove Payment drift with an executable fixture:

```sql
do $$
declare
  v_drift_document_id uuid := '41000000-0000-0000-0000-000000000206';
  v_run jsonb;
  v_run_id uuid;
  v_drift_proposal_id uuid;
  v_result jsonb;
begin
  insert into public.financial_documents (
    id, document_date, doc_number, description, supplier_name,
    payment, amount, fat
  ) values (
    v_drift_document_id, date '2027-02-20', 'BANCO-V2-206', '', '',
    'Banco', 106.00, 'S'
  );
  insert into public.import_cgd_extrato_ordem (
    id, import_batch, row_key, data, descritivo, montante
  ) values (
    '42000000-0000-0000-0000-000000000206',
    'smoke-banco-v2', 'banco-v2-206', date '2027-02-20',
    'Payment BANCOV2206', -106.00
  );

  v_run := public.create_financial_reconciliation_automatic_analysis(
    array['financial_documents_cgd_bank_statement'],
    'manual_rule',
    'smoke:banco-drift',
    '43000000-0000-0000-0000-000000000206'
  );
  v_run_id := (v_run->>'runId')::uuid;
  select id into strict v_drift_proposal_id
  from public.financial_reconciliation_automatic_proposals
  where run_id = v_run_id
    and base_source_id = v_drift_document_id
    and status = 'proposed';

  update public.financial_documents
  set payment = 'BANCO'
  where id = v_drift_document_id;

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_drift_proposal_id,
    'smoke:banco-drift'
  );

  if v_result->>'status' <> 'stale'
    or v_result->>'reason' <> 'source_snapshot_changed'
    or exists (
      select 1 from public.financial_reconciliation_automatic_proposals
      where id = v_drift_proposal_id and reconciliation_id is not null
    ) then
    raise exception 'Payment drift created an automatic reconciliation.';
  end if;
end $$;
```

Create an unfinished version-1 run/proposal row and execute it after the v2 migration:

```sql
do $$
declare
  v_version_one_run_id uuid;
  v_version_one_proposal_id uuid;
  v_result jsonb;
begin
  insert into public.financial_reconciliation_automatic_runs (
    trigger, scope, actor, client_request_id
  ) values (
    'manual', 'rule', 'smoke:legacy-version',
    '43000000-0000-0000-0000-000000000207'
  ) returning id into v_version_one_run_id;

  insert into public.financial_reconciliation_automatic_proposals (
    run_id, rule_key, rule_version, base_source_type,
    base_source_id, base_source_date, allowed_difference, status, signature
  ) values (
    v_version_one_run_id,
    'financial_documents_cgd_bank_statement',
    1,
    'financial_documents',
    '41000000-0000-0000-0000-000000000201',
    date '2027-01-20',
    4.56,
    'proposed',
    'smoke-banco-v1-pending'
  ) returning id into v_version_one_proposal_id;

  v_result := public.execute_financial_reconciliation_automatic_proposal(
    v_version_one_proposal_id,
    'smoke:legacy-version'
  );

  if v_result->>'status' <> 'stale'
    or v_result->>'reason' <> 'rule_version_changed' then
    raise exception 'Pending version 1 proposal did not become stale.';
  end if;
end $$;
```

Finally assert candidate/execution functions remain `SECURITY DEFINER`, their `proconfig` contains `search_path=public, pg_temp`, `anon` and `authenticated` lack execute, `service_role` has execute, and `pg_get_functiondef` still contains `left join lateral` without a materialized `bank_rows` CTE:

```sql
do $$
declare
  v_candidate_definition text;
  v_signature text;
begin
  select pg_get_functiondef(
    'public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)'::regprocedure
  ) into strict v_candidate_definition;

  if v_candidate_definition !~* 'left join lateral\s+\('
    or v_candidate_definition ~* 'bank_rows\s+as\s+materialized'
    or v_candidate_definition !~* 'd\.payment\s*=\s*''Banco'''
    or v_candidate_definition !~* 'p_rule_version\s*=\s*2' then
    raise exception 'Version 2 candidate function lost indexed Banco semantics.';
  end if;

  foreach v_signature in array array[
    'public.financial_reconciliation_automatic_rule_candidates(text,integer,numeric,integer)',
    'public.execute_financial_reconciliation_automatic_proposal(uuid,text)'
  ] loop
    if not (
      select procedure.prosecdef
        and coalesce(procedure.proconfig, '{}'::text[]) @> array['search_path=public, pg_temp']
      from pg_proc procedure
      where procedure.oid = v_signature::regprocedure
    )
      or has_function_privilege('anon', v_signature, 'EXECUTE')
      or has_function_privilege('authenticated', v_signature, 'EXECUTE')
      or not has_function_privilege('service_role', v_signature, 'EXECUTE') then
      raise exception 'Automatic reconciliation function security changed for %.', v_signature;
    end if;
  end loop;
end $$;
```

- [ ] **Step 8: Run focused tests and the database smoke**

Run:

```powershell
node --check api/_reconciliation-automation.js
node --test --test-isolation=none tests/reconciliation-automation.test.js
```

Expected: both PASS. The Node suite proves the API boundary; it does not claim PostgreSQL behavior.

When PostgreSQL/Supabase credentials are available, run:

```powershell
psql $env:SUPABASE_DB_URL -v ON_ERROR_STOP=1 -f tests/reconciliation-automation-rpc.smoke.sql
```

Expected: exit code `0`, no raised smoke exception, and the outer transaction rolls back all fixtures. If `psql` or `SUPABASE_DB_URL` is unavailable, record this exact command as a mandatory external gate; do not claim the SQL behavior passed.

- [ ] **Step 9: Commit Task 1**

```powershell
git add api/_reconciliation-automation.js tests/reconciliation-automation.test.js tests/reconciliation-automation-rpc.smoke.sql supabase-migrations/2026-08-16-financial-reconciliation-automation-banco-v2.sql
git commit -m "feat: restrict automatic reconciliation to Banco payments"
```

## Task 2: Filter Proposal Rows by Run Lifecycle Without Changing Counts

**Files:**
- Modify: `app-main.js:22138-22190`
- Modify: `tests/reconciliation-automation-ui.test.js:540-620, 700-1260`

**Interfaces:**
- Consumes: a public automation run object `{ finishedAt: string|null, proposals: Array<{status:string}> }` and the existing selected-proposal `Set`.
- Produces: `financialReconciliationAutomationVisibleProposals(run): Array<object>` and `financialReconciliationAutomationEmptyMessage(run): string`.
- Preserves: `financialReconciliationAutomationOutcomeCounts(run)` continues to consume all `run.proposals` and remains the only source for summary counters.

- [ ] **Step 1: Add failing executable selector tests**

Add a compiler that executes the production helper:

```js
function compileVisibleAutomationProposals() {
  return new Function(
    "clean",
    `${appFunctionSource("financialReconciliationAutomationVisibleProposals")}
     return financialReconciliationAutomationVisibleProposals;`,
  )((value) => String(value ?? "").trim());
}
```

Add lifecycle and immutability coverage:

```js
test("active automation runs show proposed and ambiguous rows only", () => {
  const visible = compileVisibleAutomationProposals();
  const proposals = [
    { id: "checked", status: "proposed" },
    { id: "unchecked", status: "proposed" },
    { id: "ambiguous", status: "ambiguous" },
    { id: "skipped", status: "skipped" },
    { id: "completed", status: "completed" },
  ];
  const run = Object.freeze({ finishedAt: null, proposals: Object.freeze(proposals) });

  assert.deepEqual(visible(run).map((proposal) => proposal.id), ["checked", "unchecked", "ambiguous"]);
  assert.equal(run.proposals, proposals);
});

test("finished automation runs show selected persisted outcomes only", () => {
  const visible = compileVisibleAutomationProposals();
  const run = {
    finishedAt: "2026-08-16T10:00:00.000Z",
    proposals: [
      { id: "completed", status: "completed" },
      { id: "stale", status: "stale" },
      { id: "failed", status: "failed" },
      { id: "ambiguous", status: "ambiguous" },
      { id: "skipped", status: "skipped" },
      { id: "deselected", status: "deselected" },
    ],
  };

  assert.deepEqual(visible(run).map((proposal) => proposal.id), ["completed", "stale", "failed"]);
});
```

Add a render harness test with one checked and one unchecked proposed row and assert both proposal IDs are present before execution. Add a finished-run render test that asserts completed/stale/failed IDs are present and ambiguous/skipped/deselected IDs are absent. In both tests, call `financialReconciliationAutomationOutcomeCounts(run)` and assert hidden rows remain counted.

Every test harness that extracts `renderFinancialReconciliationAutomation` with `new Function` must include the actual selector and empty-message helpers before the render function:

```js
`${appFunctionSource("financialReconciliationAutomationVisibleProposals")}
 ${appFunctionSource("financialReconciliationAutomationEmptyMessage")}
 ${appFunctionSource("renderFinancialReconciliationAutomation")}
 return renderFinancialReconciliationAutomation;`
```

- [ ] **Step 2: Run the focused UI test and capture RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js
```

Expected: FAIL because `financialReconciliationAutomationVisibleProposals` does not exist and rendering still maps every persisted proposal.

- [ ] **Step 3: Implement the pure lifecycle selector**

Add immediately before `financialReconciliationAutomationOutcomeCounts`:

```js
function financialReconciliationAutomationVisibleProposals(run) {
  const proposals = Array.isArray(run?.proposals) ? run.proposals : [];
  const finished = Boolean(clean(run?.finishedAt));
  const visibleStatuses = finished
    ? new Set(["completed", "stale", "failed"])
    : new Set(["proposed", "ambiguous"]);
  return proposals.filter((proposal) => visibleStatuses.has(clean(proposal?.status).toLowerCase()));
}

function financialReconciliationAutomationEmptyMessage(run) {
  return clean(run?.finishedAt)
    ? "No selected execution outcomes to show."
    : "No proposed or ambiguous matches to review.";
}
```

This selector must not inspect `selectedProposalIds`; unchecked proposals remain visible because their persisted status is still `proposed`.

- [ ] **Step 4: Integrate visible rows while retaining authoritative counts and controls**

At the start of `renderFinancialReconciliationAutomation`, separate the complete and visible collections:

```js
  const proposals = Array.isArray(automation.run?.proposals) ? automation.run.proposals : [];
  const visibleProposals = financialReconciliationAutomationVisibleProposals(automation.run);
  const executable = proposals.filter((proposal) => clean(proposal?.status).toLowerCase() === "proposed");
```

Render `visibleProposals`, not `proposals`:

```js
    els.financialReconciliationWorkbenchAutomationProposals.innerHTML = visibleProposals.length
      ? visibleProposals.map((proposal) => financialReconciliationAutomationProposalMarkup(
        proposal,
        automation.run,
        automation.rules,
        automation.selectedProposalIds,
        Boolean(pending),
      )).join("")
      : `<p class="empty">${escape(financialReconciliationAutomationEmptyMessage(automation.run))}</p>`;
```

Leave `financialReconciliationAutomationOutcomeCounts(automation.run)` unchanged. Keep select-all, clear, and execute controls derived from the complete `proposed` collection so presentation filtering cannot alter lifecycle behavior.

- [ ] **Step 5: Run focused UI tests and commit Task 2**

Run:

```powershell
node --check app-main.js
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js
```

Expected: PASS.

Commit:

```powershell
git add app-main.js tests/reconciliation-automation-ui.test.js
git commit -m "feat: focus automatic reconciliation proposal results"
```

## Task 3: Implement the Compact Paired-Card Proposal Layout

**Files:**
- Modify: `app-main.js:22078-22137`
- Modify: `styles.css:6799-6925, responsive section after 7008`
- Modify: `tests/reconciliation-automation-ui.test.js:580-690, proposal markup tests`
- Modify: `tests/reconciliation-density.test.js:automation density assertions`

**Interfaces:**
- Consumes: the visible proposals produced by `financialReconciliationAutomationVisibleProposals(run)` and the existing escaped base/destination snapshots.
- Produces: compact semantic markup classes `financial-reconciliation-automation-item-meta`, `financial-reconciliation-automation-item-description`, `financial-reconciliation-automation-item-operator`, and `financial-reconciliation-automation-item-id`.
- Preserves: checkbox data attributes, accessible names, proposal status classes, reason/attempt markup, evidence semantics, and rule/version footer content.

- [ ] **Step 1: Add failing compact-markup and escaping tests**

Extend the existing proposal-markup test with hostile source content and assert the compact semantic structure:

```js
const hostileBase = {
  sourceType: "financial_documents",
  sourceId: "document-<one>",
  sourceDate: "2026-08-16",
  docNumber: "FT <42>",
  supplierName: "Supplier & Sons",
  description: "Invoice <script>alert(1)</script>",
  amount: 100,
};

const markup = proposalMarkup(
  {
    id: WORKBENCH_PROPOSAL_1,
    ruleKey: "manual-enabled",
    ruleVersion: 3,
    status: "proposed",
    baseSnapshot: hostileBase,
    items: [{
      sourceType: "import_cgd_extrato_ordem",
      sourceId: "bank-<one>",
      sourceDate: "2026-08-16",
      description: "Bank <match>",
      amount: -100,
      evidence: { documentNumber: { matched: true, normalized: "FT42" } },
    }],
    calculatedDifference: 0,
    allowedDifference: 1,
  },
  workbenchRun([]),
  workbenchRules(),
  new Set([WORKBENCH_PROPOSAL_1]),
  false,
);

assert.match(markup, /financial-reconciliation-automation-item-meta/);
assert.match(markup, /financial-reconciliation-automation-item-description/);
assert.match(markup, /financial-reconciliation-automation-item-operator/);
assert.match(markup, /financial-reconciliation-automation-item-id/);
assert.match(markup, /Supplier &amp; Sons/);
assert.match(markup, /FT &lt;42&gt;/);
assert.match(markup, /Invoice &lt;script&gt;alert\(1\)&lt;\/script&gt;/);
assert.doesNotMatch(markup, /<script>/);
assert.match(markup, /Document number matched: FT42/);
assert.match(markup, /Difference[\s\S]*Allowed[\s\S]*version 3/);
```

Also assert the existing checkbox retains its proposal ID, escaped accessible label, checked state, and focusable proposal article.

- [ ] **Step 2: Add minimal failing CSS structure assertions**

Read `styles.css` in `tests/reconciliation-density.test.js` and assert:

```js
assert.match(styles, /\.financial-reconciliation-automation-proposal-records\s*\{[\s\S]*grid-template-columns:\s*repeat\(2,\s*minmax\(0,\s*1fr\)\)/);
assert.match(styles, /@media\s*\(max-width:\s*700px\)[\s\S]*\.financial-reconciliation-automation-proposal-records\s*\{[\s\S]*grid-template-columns:\s*minmax\(0,\s*1fr\)/);
```

These two assertions protect only the desktop-pair/narrow-stack structural contract. Verify density, line height, wrapping, and focus as browser behavior in Step 6 instead of pinning cosmetic CSS declarations in source-text tests.

- [ ] **Step 3: Run the focused tests and capture RED**

Run:

```powershell
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
```

Expected: FAIL because the compact classes and two-column fixed pair layout are not installed.

- [ ] **Step 4: Replace the verbose item definition list with compact semantic markup**

Refactor `financialReconciliationAutomationItemMarkup` to build optional metadata before returning markup:

```js
function financialReconciliationAutomationItemMarkup(item, label, operator = "") {
  const value = item && typeof item === "object" ? item : {};
  const supplier = clean(value.supplierName ?? value.supplier);
  const documentNumber = clean(value.docNumber ?? value.documentNumber);
  const sourceId = clean(value.sourceId) || "-";
  const meta = [
    formatDateOnly(value.sourceDate) || "-",
    documentNumber ? `Document ${documentNumber}` : "",
    supplier ? `Supplier ${supplier}` : "",
  ].filter(Boolean);
  const amount = formatMoney(Number(value.amount || 0));

  return `<article class="financial-reconciliation-automation-item">
    <div class="financial-reconciliation-automation-item-head">
      <span><strong>${escape(label)}</strong><small>${escape(financialReconciliationSourceLabel(clean(value.sourceType)) || "Unknown source")}</small></span>
      <strong class="financial-reconciliation-automation-item-amount">${operator ? `<span class="financial-reconciliation-automation-item-operator">${escape(operator)}</span> ` : ""}${escape(amount)}</strong>
    </div>
    <p class="financial-reconciliation-automation-item-meta">${meta.map((entry) => `<span>${escape(entry)}</span>`).join("")}</p>
    <p class="financial-reconciliation-automation-item-description">${escape(clean(value.description) || "-")}</p>
    ${value.evidence ? financialReconciliationAutomationIdentityEvidenceMarkup(value.evidence) : ""}
    <details class="financial-reconciliation-automation-item-id"><summary>Record ID</summary><code>${escape(sourceId)}</code></details>
  </article>`;
}
```

Do not truncate descriptions or IDs. The closed `<details>` keeps the immutable ID reachable without consuming normal row height. Keep the existing proposal header, reason, execution-attempt, evidence, and footer fields.

- [ ] **Step 5: Apply the approved compact paired-card CSS**

Replace the current proposal/item density values with:

```css
.financial-reconciliation-workbench-automation-proposals {
  display: grid;
  gap: .5rem;
}

.financial-reconciliation-automation-proposal {
  display: grid;
  gap: .4rem;
  padding: .55rem;
  border: 1px solid var(--border);
  border-left: 4px solid var(--brand);
  border-radius: 10px;
  overflow-wrap: anywhere;
  background: var(--surface-soft, rgba(255,255,255,.55));
}

.financial-reconciliation-automation-proposal-records {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: .35rem;
}

.financial-reconciliation-automation-item {
  display: grid;
  gap: .25rem;
  min-width: 0;
  padding: .45rem;
  border: 1px solid var(--border);
  border-radius: 8px;
  line-height: 1.25;
  background: white;
}

.financial-reconciliation-automation-item-head > span {
  display: flex;
  flex-wrap: wrap;
  align-items: baseline;
  gap: .2rem .45rem;
  min-width: 0;
}

.financial-reconciliation-automation-item-head small,
.financial-reconciliation-automation-item-meta,
.financial-reconciliation-automation-item-id {
  color: var(--muted);
  font-size: .68rem;
}

.financial-reconciliation-automation-item-meta {
  display: flex;
  flex-wrap: wrap;
  gap: .15rem .55rem;
  margin: 0;
}

.financial-reconciliation-automation-item-description {
  margin: 0;
  font-size: .72rem;
  line-height: 1.25;
  overflow-wrap: anywhere;
}

.financial-reconciliation-automation-item-amount {
  white-space: nowrap;
  font-size: .78rem;
}

.financial-reconciliation-automation-item-operator {
  color: var(--muted);
}

.financial-reconciliation-automation-item-id summary {
  width: fit-content;
  cursor: pointer;
}

.financial-reconciliation-automation-item-id code {
  display: block;
  margin-top: .15rem;
  white-space: normal;
  overflow-wrap: anywhere;
}
```

Retain the current `:focus-within` outline. Keep candidate groups full-width and give their nested cards the same two-column grid:

```css
.financial-reconciliation-automation-candidate-group {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: .35rem;
  grid-column: 1 / -1;
  min-width: 0;
  padding: .45rem;
  border: 1px solid var(--border);
  border-radius: 8px;
  background: white;
}

.financial-reconciliation-automation-candidate-group h4 {
  grid-column: 1 / -1;
  margin: 0;
  font-size: .72rem;
}
```

Add the narrow rule:

```css
@media (max-width: 700px) {
  .financial-reconciliation-automation-proposal-records,
  .financial-reconciliation-automation-candidate-group {
    grid-template-columns: minmax(0, 1fr);
  }

  .financial-reconciliation-automation-proposal > header,
  .financial-reconciliation-automation-proposal > footer,
  .financial-reconciliation-automation-item-head {
    align-items: flex-start;
  }
}
```

- [ ] **Step 6: Run focused tests and perform authenticated visual verification**

Run:

```powershell
node --check app-main.js
node --test --test-isolation=none tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
```

Expected: PASS.

In an authenticated local app session:

1. Open **Reconciliation → Automatic reconciliation**.
2. Analyze the version-2 rule and confirm paired base/destination cards fit side-by-side on desktop.
3. Confirm unchecked proposals remain visible after clearing selection.
4. Confirm ambiguous rows show compact candidate groups.
5. Execute a subset and confirm the finished view shows only completed/stale/failed selected outcomes.
6. Confirm hidden skipped/deselected rows remain represented in the counters.
7. Set viewport width below `700px`; confirm cards stack, descriptions and IDs wrap, controls remain reachable, and keyboard focus remains visible.

If no authenticated fixture/session is available, record these seven checks as a mandatory external gate and do not claim visual verification passed.

- [ ] **Step 7: Commit Task 3**

```powershell
git add app-main.js styles.css tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
git commit -m "style: compact automatic reconciliation proposals"
```

## Task 4: Full Regression, Migration Reapply, and Rollout Evidence

**Files:**
- Verify only; modify a scoped file only if a failing test identifies a defect owned by Tasks 1-3.

**Interfaces:**
- Consumes: the version-2 migration, API helper, lifecycle selector, proposal markup, CSS, and all existing reconciliation behavior.
- Produces: a clean, reviewable branch with explicit database and browser gate status.

- [ ] **Step 1: Run all syntax and focused checks from the repository root**

```powershell
node --check api/_reconciliation-automation.js
node --check api/reconciliation-automation.js
node --check api/reconciliation-automation-settings.js
node --check api/reconciliation-automation-cron.js
node --check app-main.js
node --test --test-isolation=none tests/reconciliation-automation.test.js tests/reconciliation-automation-ui.test.js tests/reconciliation-density.test.js
```

Expected: all commands exit `0` with no failed tests.

- [ ] **Step 2: Run the complete Node regression suite**

```powershell
node --test --test-isolation=none tests/*.test.js
```

Expected: every test passes with zero failures, skips, cancellations, or unfinished tests.

- [ ] **Step 3: Verify migration and worktree integrity**

```powershell
git diff --check main...HEAD
git status --short
git diff --name-only main...HEAD
```

Expected: diff check exits `0`; only the eight files listed in the File Map plus the approved design and this plan appear; no unrelated user files are changed.

- [ ] **Step 4: Apply and reapply the migration in a development Supabase environment**

Apply migrations in this order:

1. All already-deployed migrations through `2026-08-15-financial-reconciliation-automation-candidate-index-lookup.sql`.
2. `2026-08-16-financial-reconciliation-automation-banco-v2.sql`.

Then re-run only the new migration once to prove idempotency. Query the catalog and configuration:

```sql
select rule_key, version, logic_description, definition
from public.financial_reconciliation_automatic_rule_definitions
where rule_key = 'financial_documents_cgd_bank_statement'
order by version;

select rule_key, rule_version, enabled, allow_manual_execution,
       include_in_scheduled_batch, difference_allowed,
       max_difference_days, priority, updated_by
from public.financial_reconciliation_automatic_rule_configs
where rule_key = 'financial_documents_cgd_bank_statement';
```

Expected: versions 1 and 2 both exist, configuration points to 2, and administrator-owned values are unchanged.

- [ ] **Step 5: Run final authenticated rollout checks before scheduled reliance**

1. Publish the application and migration to a protected non-production environment.
2. Run one manual version-2 analysis.
3. Verify all proposed Financial Documents have stored `payment = 'Banco'`.
4. Verify a deliberately non-`Banco` document produces no persisted proposal row.
5. Verify an old pending version-1 proposal becomes stale with `rule_version_changed`.
6. Verify Payment drift after analysis becomes stale and creates no reconciliation.
7. Verify active/finished proposal visibility and compact desktop/narrow layout.
8. Enable or rely on the shared scheduled batch only after the preceding checks pass.

- [ ] **Step 6: Request independent review and record the final commit**

Review the complete range for rule-version immutability, exact text semantics, SQL privilege preservation, lifecycle visibility, escaping, accessibility, and unrelated-file scope. If review finds an owned defect, add a failing regression test before changing production code, rerun focused/full verification, and commit the fix separately.

```powershell
git status --short
git log --oneline main..HEAD
```

Expected: clean worktree and a readable sequence of task commits.

## Plan Self-Review Checklist

- Spec coverage: Task 1 covers immutable v2 seeding, exact Banco eligibility, configuration preservation, v1 stale handling, Payment drift, latest indexed lookup, privileges, RPC compatibility, idempotency, and SQL rollout. Task 2 covers active/finished visibility, unchecked proposal re-selection, and authoritative hidden-row counts. Task 3 covers the approved compact paired-card UI, audit IDs, escaping, focus, wrapping, and narrow layouts. Task 4 covers complete regression and rollout gates.
- Placeholder scan: no deferred implementation markers, unspecified error-handling instructions, or unnamed tests remain.
- Type consistency: `AUTOMATIC_RULE_VERSION` is numeric `2`; SQL `rule_version` is integer `2`; `finishedAt` is the existing camelCase nullable timestamp; proposal statuses use the existing persisted strings; the selector returns original proposal object references without mutation.
- Scope consistency: no new endpoint, table, configurable predicate, browser database call, or rule key is introduced.
