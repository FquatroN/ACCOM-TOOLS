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
const PROPOSAL_STATUSES = new Set([
  "proposed",
  "ambiguous",
  "deselected",
  "executing",
  "completed",
  "stale",
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

function isPlainRecord(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function hasOwnFields(value, fields) {
  return fields.every((field) => Object.hasOwn(value, field));
}

function isTimestamp(value) {
  return typeof value === "string" && value !== "" && Number.isFinite(Date.parse(value));
}

function headerValue(headers, name) {
  if (!headers || typeof headers !== "object") return "";
  const matchedKey = Object.keys(headers).find((key) => key.toLowerCase() === name);
  return matchedKey === undefined ? "" : text(headers[matchedKey]);
}

function hasCronBearer(req, cronSecret) {
  return typeof cronSecret === "string"
    && cronSecret !== ""
    && headerValue(req?.headers, "authorization") === `Bearer ${cronSecret}`;
}

function compareText(left, right) {
  const leftText = text(left);
  const rightText = text(right);
  return leftText < rightText ? -1 : leftText > rightText ? 1 : 0;
}

function requireScheduledRun(value, expectedRunId = "") {
  const run = toAutomationPublicResult(value);
  if (!isPlainRecord(run)
    || !hasOwnFields(run, [
      "runId",
      "trigger",
      "scope",
      "status",
      "actor",
      "analysisCompletedAt",
      "finishedAt",
      "definitions",
      "proposals",
    ])
    || !UUID_PATTERN.test(text(run.runId))
    || (expectedRunId && run.runId !== expectedRunId)
    || run.trigger !== "scheduled"
    || run.scope !== "batch"
    || run.actor !== SCHEDULE_ACTOR
    || !SAFE_RUN_STATUSES.has(run.status)
    || (run.analysisCompletedAt !== null && !isTimestamp(run.analysisCompletedAt))
    || (run.finishedAt !== null && !isTimestamp(run.finishedAt))
    || !Array.isArray(run.definitions)
    || !Array.isArray(run.proposals)) {
    throw new Error("Scheduled reconciliation run response is invalid.");
  }

  const ruleKeys = new Set();
  const priorities = new Set();
  for (const definition of run.definitions) {
    const ruleKey = text(definition?.ruleKey);
    const priority = Number(definition?.priority);
    if (!isPlainRecord(definition)
      || !hasOwnFields(definition, ["ruleKey", "priority"])
      || !ruleKey || !Number.isSafeInteger(priority) || priority < 1
      || ruleKeys.has(ruleKey) || priorities.has(priority)) {
      throw new Error("Scheduled reconciliation rule snapshot is invalid.");
    }
    ruleKeys.add(ruleKey);
    priorities.add(priority);
  }
  if (ruleKeys.size === 0) throw new Error("Scheduled reconciliation rule snapshot is empty.");

  const proposalIds = new Set();
  for (const proposal of run.proposals) {
    if (!isPlainRecord(proposal)
      || !hasOwnFields(proposal, ["id", "runId", "ruleKey", "baseSourceDate", "baseSourceId", "status"])
      || !UUID_PATTERN.test(text(proposal.id))
      || proposal.runId !== run.runId
      || !ruleKeys.has(proposal.ruleKey)
      || !/^\d{4}-\d{2}-\d{2}$/.test(text(proposal.baseSourceDate))
      || !UUID_PATTERN.test(text(proposal.baseSourceId))
      || !PROPOSAL_STATUSES.has(proposal.status)
      || proposalIds.has(proposal.id)) {
      throw new Error("Scheduled reconciliation proposal response is invalid.");
    }
    proposalIds.add(proposal.id);
  }

  const analysisComplete = run.analysisCompletedAt !== null;
  const finished = run.finishedAt !== null;
  if ((run.status === "analyzing" && (analysisComplete || finished || run.proposals.length > 0))
    || ((run.status === "ready" || run.status === "running") && (!analysisComplete || finished))
    || ((run.status === "completed" || run.status === "partial" || run.status === "failed")
      && (!analysisComplete || !finished || hasMoreWork(run)))) {
    throw new Error("Scheduled reconciliation run lifecycle is invalid.");
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
        || compareText(left.baseSourceDate, right.baseSourceDate)
        || compareText(left.baseSourceId, right.baseSourceId)
        || compareText(left.id, right.id)
        || compareText(left.ruleKey, right.ruleKey);
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
  const cronSecret = process.env.CRON_SECRET;
  if (!isCronRequest(req, cronSecret) || !hasCronBearer(req, cronSecret)) {
    return res.status(401).json({ error: "Unauthorized." });
  }
  if (req.method !== "GET" && req.method !== "POST") {
    res.setHeader("Allow", "GET, POST");
    return res.status(405).json({ error: "Method not allowed." });
  }

  try {
    const claim = toAutomationPublicResult(await restQuery(
      "rpc/claim_financial_reconciliation_automatic_schedule",
      {
        method: "POST",
        body: { p_now: new Date().toISOString(), p_actor: SCHEDULE_ACTOR },
      },
    ));
    if (!isPlainRecord(claim)
      || !Object.hasOwn(claim, "claimed")
      || typeof claim.claimed !== "boolean") {
      throw new Error("Scheduled reconciliation claim response is invalid.");
    }
    if (!claim.claimed) {
      if (!Object.hasOwn(claim, "reason")) {
        throw new Error("Scheduled reconciliation claim reason is invalid.");
      }
      return res.status(200).json({
        ok: true,
        claimed: false,
        reason: SAFE_UNCLAIMED_REASONS.has(claim.reason) ? claim.reason : "not_claimed",
        hasMore: false,
      });
    }
    if (!hasOwnFields(claim, ["resumed", "run"]) || typeof claim.resumed !== "boolean") {
      throw new Error("Scheduled reconciliation claimed response is invalid.");
    }

    let run = requireScheduledRun(claim.run);
    const claimedRunId = run.runId;
    if (run.finishedAt) {
      return res.status(200).json(publicRunResponse(claim, run, 0, false));
    }
    if (!run.analysisCompletedAt) {
      run = requireScheduledRun(await restQuery(
        "rpc/populate_financial_reconciliation_automatic_run",
        { method: "POST", body: { p_run_id: run.runId } },
      ), claimedRunId);
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
      ), claimedRunId);
    }
    let hasMore = hasMoreWork(run);
    if (!hasMore && !run.finishedAt) {
      run = requireScheduledRun(await restQuery(
        "rpc/finish_financial_reconciliation_automatic_run",
        { method: "POST", body: { p_run_id: run.runId } },
      ), claimedRunId);
      hasMore = hasMoreWork(run);
    }

    return res.status(200).json(publicRunResponse(claim, run, selected.length, hasMore));
  } catch {
    return res.status(500).json({ error: "Unexpected server error." });
  }
};
