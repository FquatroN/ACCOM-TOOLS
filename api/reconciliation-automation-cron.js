const { restQuery } = require("./_supabase");
const {
  ADYEN_MONTHLY_RULE_KEY,
  AUTOMATIC_RULE_VERSIONS,
  BANK_RESERVATION_RULE_KEY,
  isCronRequest,
  MONTHLY_INCOME_RULE_KEY,
  toAutomationPublicResult,
} = require("./_reconciliation-automation");

const SCHEDULE_ACTOR = "system:reconciliation";
const MAX_PROPOSALS_PER_HEARTBEAT = 25;
const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const SAFE_UNCLAIMED_REASONS = new Set([
  "schedule_disabled",
  "before_scheduled_time",
  "no_enabled_rules",
  "unsupported_rule_set",
  "slot_failed",
  "batch_complete",
]);
const SAFE_RUN_STATUSES = new Set([
  "analyzing",
  "ready",
  "running",
  "completed",
  "partial",
  "failed",
]);
const TERMINAL_RUN_STATUSES = new Set(["completed", "partial", "failed"]);
const PROPOSAL_STATUSES = new Set([
  "proposed",
  "ambiguous",
  "skipped",
  "deselected",
  "executing",
  "completed",
  "stale",
  "failed",
]);
const COUNT_STATUSES = [
  "proposed",
  "ambiguous",
  "skipped",
  "deselected",
  "executing",
  "completed",
  "stale",
  "failed",
];

function analysisUnitForRule(ruleKey) {
  if (ruleKey === BANK_RESERVATION_RULE_KEY) return "bank_anchors";
  if (ruleKey === MONTHLY_INCOME_RULE_KEY || ruleKey === ADYEN_MONTHLY_RULE_KEY) {
    return "calendar_months";
  }
  return "records";
}

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
  return typeof value === "string"
    && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,6})?(?:Z|[+-]\d{2}:\d{2})$/.test(value)
    && Number.isFinite(Date.parse(value));
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

function requireScheduledRun(value, expectedRun = null) {
  const run = toAutomationPublicResult(value);
  if (!isPlainRecord(run)
    || !hasOwnFields(run, [
      "runId",
      "trigger",
      "scope",
      "status",
      "actor",
      "batchId",
      "batchRuleKey",
      "batchRulePosition",
      "batchRuleCount",
      "analysisCursorDate",
      "analysisCursorId",
      "analysisProcessed",
      "analysisTotal",
      "analysisErrorCode",
      "analysisErrorAt",
      "analysisUnit",
      "analysisComplete",
      "analysisCompletedAt",
      "finishedAt",
      "definitions",
      "proposals",
    ])
    || !UUID_PATTERN.test(text(run.runId))
    || run.trigger !== "scheduled"
    || run.scope !== "rule"
    || run.actor !== SCHEDULE_ACTOR
    || !UUID_PATTERN.test(text(run.batchId))
    || !Number.isSafeInteger(run.batchRulePosition) || run.batchRulePosition < 1
    || !Number.isSafeInteger(run.batchRuleCount) || run.batchRulePosition > run.batchRuleCount
    || !SAFE_RUN_STATUSES.has(run.status)
    || (run.analysisCursorDate !== null && !/^\d{4}-\d{2}-\d{2}$/.test(text(run.analysisCursorDate)))
    || (run.analysisCursorId !== null && !UUID_PATTERN.test(text(run.analysisCursorId)))
    || ((run.analysisCursorDate === null) !== (run.analysisCursorId === null))
    || !Number.isSafeInteger(run.analysisProcessed) || run.analysisProcessed < 0
    || !Number.isSafeInteger(run.analysisTotal) || run.analysisTotal < 0
    || run.analysisProcessed > run.analysisTotal
    || (run.analysisErrorCode !== null && typeof run.analysisErrorCode !== "string")
    || (run.analysisErrorAt !== null && !isTimestamp(run.analysisErrorAt))
    || !new Set(["records", "bank_anchors", "calendar_months"]).has(run.analysisUnit)
    || typeof run.analysisComplete !== "boolean"
    || (run.analysisCompletedAt !== null && !isTimestamp(run.analysisCompletedAt))
    || (run.finishedAt !== null && !isTimestamp(run.finishedAt))
    || !Array.isArray(run.definitions)
    || !Array.isArray(run.proposals)) {
    throw new Error("Scheduled reconciliation run response is invalid.");
  }

  if (expectedRun && (run.runId !== expectedRun.runId
    || run.batchId !== expectedRun.batchId
    || run.batchRuleKey !== expectedRun.batchRuleKey
    || run.batchRulePosition !== expectedRun.batchRulePosition
    || run.batchRuleCount !== expectedRun.batchRuleCount)) {
    throw new Error("Scheduled reconciliation run identity changed.");
  }

  if (run.definitions.length !== 1) {
    throw new Error("Scheduled reconciliation run must contain one rule snapshot.");
  }
  const definition = run.definitions[0];
  const ruleKey = text(definition?.ruleKey);
  const priority = definition?.priority;
  if (!isPlainRecord(definition)
    || !hasOwnFields(definition, ["ruleKey", "ruleVersion", "priority"])
    || !Object.hasOwn(AUTOMATIC_RULE_VERSIONS, ruleKey)
    || definition.ruleVersion !== AUTOMATIC_RULE_VERSIONS[ruleKey]
    || !Number.isSafeInteger(priority) || priority < 1
    || run.batchRuleKey !== ruleKey
    || run.analysisUnit !== analysisUnitForRule(ruleKey)) {
    throw new Error("Scheduled reconciliation rule snapshot is invalid.");
  }

  const strategyUsesProjectedCursor = ruleKey === BANK_RESERVATION_RULE_KEY
    || ruleKey === ADYEN_MONTHLY_RULE_KEY;
  if (strategyUsesProjectedCursor
    && ((run.analysisProcessed === 0 && run.analysisCursorDate !== null)
      || (run.analysisProcessed > 0 && run.analysisCursorDate === null)
      || (ruleKey === ADYEN_MONTHLY_RULE_KEY
        && run.analysisCursorDate !== null
        && !/^\d{4}-\d{2}-01$/.test(run.analysisCursorDate))
      || (run.analysisComplete
        && run.analysisProcessed !== run.analysisTotal))) {
    throw new Error("Scheduled reconciliation strategy progress is invalid.");
  }

  const proposalIds = new Set();
  for (const proposal of run.proposals) {
    if (!isPlainRecord(proposal)
      || !hasOwnFields(proposal, ["id", "runId", "ruleKey", "baseSourceDate", "baseSourceId", "status"])
      || !UUID_PATTERN.test(text(proposal.id))
      || proposal.runId !== run.runId
      || proposal.ruleKey !== ruleKey
      || !/^\d{4}-\d{2}-\d{2}$/.test(text(proposal.baseSourceDate))
      || !UUID_PATTERN.test(text(proposal.baseSourceId))
      || !PROPOSAL_STATUSES.has(proposal.status)
      || proposalIds.has(proposal.id)) {
      throw new Error("Scheduled reconciliation proposal response is invalid.");
    }
    proposalIds.add(proposal.id);
  }

  const analysisComplete = run.analysisCompletedAt !== null;
  if (run.analysisComplete !== analysisComplete) {
    throw new Error("Scheduled reconciliation analysis progress is inconsistent.");
  }
  const finished = run.finishedAt !== null;
  const analysisFailed = run.status === "failed"
    && !analysisComplete
    && finished
    && typeof run.analysisErrorCode === "string"
    && run.analysisErrorCode.length > 0
    && isTimestamp(run.analysisErrorAt);
  if ((run.status === "analyzing" && (analysisComplete || finished || run.proposals.length > 0))
    || ((run.status === "ready" || run.status === "running") && (!analysisComplete || finished))
    || ((run.status === "completed" || run.status === "partial")
      && (!analysisComplete || !finished || hasMoreWork(run)))
    || (run.status === "failed"
      && !analysisFailed
      && (!analysisComplete || !finished || hasMoreWork(run)))) {
    throw new Error("Scheduled reconciliation run lifecycle is invalid.");
  }
  return run;
}

function requireScheduleClaim(value) {
  const claim = toAutomationPublicResult(value);
  if (!isPlainRecord(claim)
    || !Object.hasOwn(claim, "claimed")
    || typeof claim.claimed !== "boolean") {
    throw new Error("Scheduled reconciliation claim response is invalid.");
  }
  if (!claim.claimed) {
    if (!Object.hasOwn(claim, "reason") || !SAFE_UNCLAIMED_REASONS.has(claim.reason)) {
      throw new Error("Scheduled reconciliation claim reason is invalid.");
    }
    if (claim.reason === "batch_complete"
      && (!Object.hasOwn(claim, "batchId") || !UUID_PATTERN.test(text(claim.batchId)))) {
      throw new Error("Scheduled reconciliation completed batch response is invalid.");
    }
    return claim;
  }

  if (!hasOwnFields(claim, [
    "resumed",
    "batchId",
    "batchRulePosition",
    "batchRuleCount",
    "run",
  ])
    || typeof claim.resumed !== "boolean"
    || !UUID_PATTERN.test(text(claim.batchId))
    || !Number.isSafeInteger(claim.batchRulePosition) || claim.batchRulePosition < 1
    || !Number.isSafeInteger(claim.batchRuleCount)
    || claim.batchRulePosition > claim.batchRuleCount) {
    throw new Error("Scheduled reconciliation claimed response is invalid.");
  }

  const run = requireScheduledRun(claim.run);
  if (claim.batchId !== run.batchId
    || claim.batchRulePosition !== run.batchRulePosition
    || claim.batchRuleCount !== run.batchRuleCount) {
    throw new Error("Scheduled reconciliation claim does not match its child run.");
  }
  return { ...claim, run };
}

function requireContinuedAnalysisRun(value) {
  const run = toAutomationPublicResult(value);
  if (!isPlainRecord(run)
    || !hasOwnFields(run, [
      "runId", "status", "analysisCursorDate", "analysisCursorId",
      "analysisProcessed", "analysisTotal", "analysisErrorCode", "analysisErrorAt",
      "analysisUnit", "analysisComplete", "analysisCompletedAt",
    ])
    || !UUID_PATTERN.test(text(run.runId))
    || !SAFE_RUN_STATUSES.has(run.status)
    || (run.analysisCursorDate !== null && !/^\d{4}-\d{2}-\d{2}$/.test(text(run.analysisCursorDate)))
    || (run.analysisCursorId !== null && !UUID_PATTERN.test(text(run.analysisCursorId)))
    || ((run.analysisCursorDate === null) !== (run.analysisCursorId === null))
    || !Number.isSafeInteger(run.analysisProcessed) || run.analysisProcessed < 0
    || !Number.isSafeInteger(run.analysisTotal) || run.analysisTotal < 0
    || run.analysisProcessed > run.analysisTotal
    || (run.analysisErrorCode !== null && typeof run.analysisErrorCode !== "string")
    || (run.analysisErrorAt !== null && !isTimestamp(run.analysisErrorAt))
    || !new Set(["records", "bank_anchors", "calendar_months"]).has(run.analysisUnit)
    || typeof run.analysisComplete !== "boolean"
    || (run.analysisCompletedAt !== null && !isTimestamp(run.analysisCompletedAt))
    || run.analysisComplete !== (run.analysisCompletedAt !== null)) {
    throw new Error("Continued reconciliation analysis response is invalid.");
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

function requireAnalyzedRun(value, expectedRun) {
  const run = requireScheduledRun(value, expectedRun);
  if (run.analysisCompletedAt === null) {
    throw new Error("Scheduled reconciliation analysis did not complete.");
  }
  return run;
}

function requireFinishedRun(value, expectedRun) {
  const run = requireAnalyzedRun(value, expectedRun);
  if (run.finishedAt === null || !TERMINAL_RUN_STATUSES.has(run.status)) {
    throw new Error("Scheduled reconciliation finalization did not complete.");
  }
  return run;
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
    batchId: run.batchId,
    ruleKey: run.definitions[0].ruleKey,
    rulePosition: run.batchRulePosition,
    ruleCount: run.batchRuleCount,
    runId: run.runId,
    status: SAFE_RUN_STATUSES.has(run.status) ? run.status : "running",
    analysisProcessed: run.analysisProcessed,
    analysisTotal: run.analysisTotal,
    analysisUnit: run.analysisUnit,
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
    const continued = toAutomationPublicResult(await restQuery(
      "rpc/continue_financial_reconciliation_automatic_oldest_analysis",
      { method: "POST", body: { p_worker: SCHEDULE_ACTOR } },
    ));
    if (!isPlainRecord(continued)
      || !Object.hasOwn(continued, "continued")
      || typeof continued.continued !== "boolean") {
      throw new Error("Scheduled reconciliation continuation response is invalid.");
    }
    if (continued.continued) {
      let run = requireContinuedAnalysisRun(continued.run);
      const scheduled = run.trigger === "scheduled";
      if (scheduled) run = requireScheduledRun(run);
      return res.status(200).json({
        ok: true,
        claimed: false,
        continuedAnalysis: true,
        ...(scheduled ? {
          batchId: run.batchId,
          ruleKey: run.definitions[0].ruleKey,
          rulePosition: run.batchRulePosition,
          ruleCount: run.batchRuleCount,
        } : {}),
        runId: run.runId,
        status: run.status,
        analysisProcessed: run.analysisProcessed,
        analysisTotal: run.analysisTotal,
        analysisUnit: run.analysisUnit,
        hasMore: run.status === "analyzing" && !run.analysisComplete,
      });
    }

    const claim = requireScheduleClaim(await restQuery(
      "rpc/claim_financial_reconciliation_automatic_schedule",
      {
        method: "POST",
        body: { p_now: new Date().toISOString(), p_actor: SCHEDULE_ACTOR },
      },
    ));
    if (!claim.claimed) {
      return res.status(200).json({
        ok: true,
        claimed: false,
        reason: claim.reason,
        ...(claim.reason === "batch_complete" ? { batchId: claim.batchId } : {}),
        hasMore: false,
      });
    }

    let run = claim.run;
    const claimedRun = {
      runId: run.runId,
      batchId: run.batchId,
      batchRuleKey: run.batchRuleKey,
      batchRulePosition: run.batchRulePosition,
      batchRuleCount: run.batchRuleCount,
    };
    if (run.finishedAt) {
      return res.status(200).json(publicRunResponse(claim, run, 0, false));
    }
    if (!run.analysisCompletedAt) {
      run = requireScheduledRun(await restQuery(
        "rpc/continue_financial_reconciliation_automatic_analysis",
        { method: "POST", body: { p_run_id: run.runId, p_actor: SCHEDULE_ACTOR } },
      ), claimedRun);
      if (!run.analysisCompletedAt) {
        return res.status(200).json(publicRunResponse(claim, run, 0, run.status === "analyzing"));
      }
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
      run = requireAnalyzedRun(await restQuery(
        "rpc/get_financial_reconciliation_automatic_run",
        { method: "POST", body: { p_run_id: run.runId } },
      ), claimedRun);
    }
    let hasMore = hasMoreWork(run);
    if (!hasMore && !run.finishedAt) {
      run = requireFinishedRun(await restQuery(
        "rpc/finish_financial_reconciliation_automatic_run",
        { method: "POST", body: { p_run_id: run.runId } },
      ), claimedRun);
      hasMore = hasMoreWork(run);
    }

    return res.status(200).json(publicRunResponse(claim, run, selected.length, hasMore));
  } catch {
    return res.status(500).json({ error: "Unexpected server error." });
  }
};
