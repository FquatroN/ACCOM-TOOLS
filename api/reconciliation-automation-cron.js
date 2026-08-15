const { restQuery } = require("./_supabase");
const { isCronRequest, toAutomationPublicResult } = require("./_reconciliation-automation");

const SCHEDULE_ACTOR = "system:reconciliation";
const MAX_PROPOSALS_PER_HEARTBEAT = 25;
const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const SAFE_UNCLAIMED_REASONS = new Set([
  "schedule_disabled",
  "before_scheduled_time",
  "no_enabled_rules",
]);
const SAFE_RUN_STATUSES = new Set([
  "analyzing",
  "ready",
  "running",
  "completed",
  "partial",
  "failed",
]);
const COUNT_STATUSES = [
  "proposed",
  "ambiguous",
  "deselected",
  "executing",
  "completed",
  "stale",
  "failed",
];

function text(value) {
  return typeof value === "string" ? value : "";
}

function requireScheduledRun(value) {
  const run = toAutomationPublicResult(value);
  if (!run || typeof run !== "object" || Array.isArray(run)
    || !UUID_PATTERN.test(text(run.runId))
    || run.trigger !== "scheduled"
    || run.scope !== "batch"
    || !Array.isArray(run.definitions)
    || !Array.isArray(run.proposals)) {
    throw new Error("Scheduled reconciliation run response is invalid.");
  }
  return run;
}

function stablePendingProposals(run) {
  const priorities = new Map();
  for (const definition of run.definitions) {
    const priority = Number(definition?.priority);
    if (text(definition?.ruleKey) && Number.isSafeInteger(priority) && priority > 0) {
      priorities.set(definition.ruleKey, priority);
    }
  }

  return run.proposals
    .filter((proposal) => proposal?.status === "proposed" && UUID_PATTERN.test(text(proposal?.id)))
    .sort((left, right) => {
      const leftPriority = priorities.get(left.ruleKey) ?? Number.MAX_SAFE_INTEGER;
      const rightPriority = priorities.get(right.ruleKey) ?? Number.MAX_SAFE_INTEGER;
      return leftPriority - rightPriority
        || text(left.ruleKey).localeCompare(text(right.ruleKey))
        || text(left.baseSourceDate).localeCompare(text(right.baseSourceDate))
        || text(left.baseSourceId).localeCompare(text(right.baseSourceId))
        || text(left.id).localeCompare(text(right.id));
    });
}

function hasMoreWork(run) {
  return run.proposals.some((proposal) => proposal?.status === "proposed" || proposal?.status === "executing");
}

function publicCounts(run) {
  const counts = {
    bases: new Set(run.proposals.map((proposal) => text(proposal?.baseSourceId)).filter(Boolean)).size,
  };
  for (const status of COUNT_STATUSES) {
    counts[status] = run.proposals.filter((proposal) => proposal?.status === status).length;
  }
  return counts;
}

function publicRunResponse(claim, run, attemptedCount, hasMore) {
  return {
    ok: true,
    claimed: true,
    resumed: claim.resumed === true,
    runId: run.runId,
    status: SAFE_RUN_STATUSES.has(run.status) ? run.status : "running",
    counts: publicCounts(run),
    attemptedCount,
    hasMore,
  };
}

module.exports = async function handler(req, res) {
  if (req.method !== "GET" && req.method !== "POST") {
    res.setHeader("Allow", "GET, POST");
    return res.status(405).json({ error: "Method not allowed." });
  }
  if (!isCronRequest(req, process.env.CRON_SECRET)) {
    return res.status(401).json({ error: "Unauthorized." });
  }

  try {
    const claim = toAutomationPublicResult(await restQuery(
      "rpc/claim_financial_reconciliation_automatic_schedule",
      {
        method: "POST",
        body: { p_now: new Date().toISOString(), p_actor: SCHEDULE_ACTOR },
      },
    ));
    if (!claim || typeof claim !== "object" || Array.isArray(claim) || typeof claim.claimed !== "boolean") {
      throw new Error("Scheduled reconciliation claim response is invalid.");
    }
    if (!claim.claimed) {
      return res.status(200).json({
        ok: true,
        claimed: false,
        reason: SAFE_UNCLAIMED_REASONS.has(claim.reason) ? claim.reason : "not_claimed",
        hasMore: false,
      });
    }

    let run = requireScheduledRun(claim.run);
    if (run.finishedAt) {
      return res.status(200).json(publicRunResponse(claim, run, 0, false));
    }
    if (!run.analysisCompletedAt) {
      run = requireScheduledRun(await restQuery(
        "rpc/populate_financial_reconciliation_automatic_run",
        { method: "POST", body: { p_run_id: run.runId } },
      ));
    }

    const selected = stablePendingProposals(run).slice(0, MAX_PROPOSALS_PER_HEARTBEAT);
    for (const proposal of selected) {
      try {
        await restQuery("rpc/execute_financial_reconciliation_automatic_proposal", {
          method: "POST",
          body: { p_proposal_id: proposal.id, p_actor: SCHEDULE_ACTOR },
        });
      } catch {
        // A later heartbeat re-reads and safely resumes work that did not persist an outcome.
      }
    }

    if (selected.length > 0) {
      run = requireScheduledRun(await restQuery(
        "rpc/get_financial_reconciliation_automatic_run",
        { method: "POST", body: { p_run_id: run.runId } },
      ));
    }
    let hasMore = hasMoreWork(run);
    if (!hasMore && !run.finishedAt) {
      run = requireScheduledRun(await restQuery(
        "rpc/finish_financial_reconciliation_automatic_run",
        { method: "POST", body: { p_run_id: run.runId } },
      ));
      hasMore = hasMoreWork(run);
    }

    return res.status(200).json(publicRunResponse(claim, run, selected.length, hasMore));
  } catch {
    return res.status(500).json({ error: "Unexpected server error." });
  }
};
