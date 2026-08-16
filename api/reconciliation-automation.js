const { cleanText, parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const { mapRpcError } = require("./_reconciliation");
const {
  normalizeAnalyzePayload,
  normalizeAutomationAction,
  normalizeContinueAnalysisPayload,
  normalizeExecutePayload,
  toAutomationPublicResult,
} = require("./_reconciliation-automation");

const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function inputError(message) {
  const error = new Error(message);
  error.statusCode = 400;
  return error;
}

function statusError(message, statusCode) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

function safePublicError(error) {
  const mapped = mapRpcError(error);
  if (!error?.supabasePayload && mapped.statusCode < 500) return mapped;
  const safe = new Error(mapped.statusCode === 409
    ? "The reconciliation automation state changed. Refresh and try again."
    : mapped.statusCode >= 500
      ? "Unexpected server error."
      : "Reconciliation automation request could not be completed.");
  safe.statusCode = mapped.statusCode;
  return safe;
}

function actorFor(auth) {
  return cleanText(auth.user?.email) || cleanText(auth.user?.id);
}

async function requireManagedFeature(req, area) {
  const auth = await requireFeature(req, area, "financial-reconciliation");
  if (!cleanText(auth.access?.profile?.id)) {
    throw statusError("You do not have permission for this feature.", 403);
  }
  return auth;
}

function normalizeRunId(value) {
  if (typeof value !== "string" || !UUID_PATTERN.test(value)) {
    throw inputError("Run ID must be a valid UUID.");
  }
  return value.toLowerCase();
}

function requireClientRequestId(input) {
  if (!input.clientRequestId) throw inputError("Client request ID is required.");
  return input;
}

async function createAnalysis(input, actor, mode) {
  const result = await restQuery("rpc/create_financial_reconciliation_automatic_analysis", {
    method: "POST",
    body: {
      p_rule_keys: input.ruleKeys,
      p_mode: mode,
      p_actor: actor,
      p_client_request_id: input.clientRequestId,
    },
  });
  return toAutomationPublicResult(result);
}

async function analyzeRule(req, body) {
  const auth = await requireManagedFeature(req, "app");
  const input = requireClientRequestId(normalizeAnalyzePayload(body));
  if (input.action !== "analyze_rule" || input.ruleKeys.length !== 1) {
    throw inputError("Analyze rule requires exactly one manually enabled rule.");
  }
  return createAnalysis(input, actorFor(auth), "manual_rule");
}

async function continueAnalysis(req, body) {
  const auth = await requireManagedFeature(req, "app");
  const input = normalizeContinueAnalysisPayload(body);
  const result = await restQuery("rpc/continue_financial_reconciliation_automatic_analysis", {
    method: "POST",
    body: { p_run_id: input.runId, p_actor: actorFor(auth) },
  });
  return toAutomationPublicResult(result);
}

async function executeSelected(req, body) {
  const auth = await requireManagedFeature(req, "app");
  const normalizedInput = normalizeExecutePayload(body);
  const input = {
    ...normalizedInput,
    runId: normalizedInput.runId.toLowerCase(),
    proposalIds: normalizedInput.proposalIds.map((proposalId) => proposalId.toLowerCase()),
  };
  if (new Set(input.proposalIds).size !== input.proposalIds.length) {
    throw inputError("Proposal IDs must contain between 1 and 100 unique proposal IDs.");
  }
  const actor = actorFor(auth);
  const outcomes = [];
  const run = toAutomationPublicResult(await restQuery(
    "rpc/get_financial_reconciliation_automatic_run",
    { method: "POST", body: { p_run_id: input.runId } },
  ));
  if (run?.runId !== input.runId || run?.trigger !== "manual") {
    throw inputError("Selected proposals must belong to the requested manual run.");
  }
  if (cleanText(run.actor) !== actor) {
    throw statusError("You do not have permission for this automation run.", 403);
  }
  const proposals = new Map((Array.isArray(run.proposals) ? run.proposals : [])
    .map((proposal) => [cleanText(proposal?.id).toLowerCase(), proposal]));
  if (input.proposalIds.some((proposalId) => !proposals.has(proposalId))) {
    throw inputError("Selected proposals must belong to the requested manual run.");
  }
  if (run.finishedAt) {
    return {
      run,
      outcomes: input.proposalIds.map((proposalId) => {
        const proposal = proposals.get(proposalId);
        const outcome = {
          proposalId,
          runId: input.runId,
          status: proposal.status,
        };
        if (proposal.reason) outcome.reason = proposal.reason;
        if (proposal.reconciliationId) outcome.reconciliationId = proposal.reconciliationId;
        return outcome;
      }),
    };
  }

  for (const proposalId of input.proposalIds) {
    try {
      const result = await restQuery("rpc/execute_financial_reconciliation_automatic_proposal", {
        method: "POST",
        body: { p_proposal_id: proposalId, p_actor: actor },
      });
      outcomes.push(toAutomationPublicResult(result));
    } catch {
      outcomes.push({ proposalId, status: "failed", reason: "execution_failed" });
    }
  }

  const refreshedRun = toAutomationPublicResult(await restQuery(
    "rpc/get_financial_reconciliation_automatic_run",
    { method: "POST", body: { p_run_id: input.runId } },
  ));
  const refreshedProposals = new Map((Array.isArray(refreshedRun?.proposals) ? refreshedRun.proposals : [])
    .map((proposal) => [cleanText(proposal?.id).toLowerCase(), proposal]));
  const hasUnresolvedSelection = input.proposalIds.some((proposalId) => {
    const status = cleanText(refreshedProposals.get(proposalId)?.status).toLowerCase();
    return !status || status === "proposed" || status === "executing";
  });
  if (hasUnresolvedSelection) return { run: refreshedRun, outcomes };

  const finalizedRun = await restQuery("rpc/finish_financial_reconciliation_automatic_run", {
    method: "POST",
    body: { p_run_id: input.runId },
  });
  return { run: toAutomationPublicResult(finalizedRun), outcomes };
}

module.exports = async function handler(req, res) {
  try {
    if (req.method === "GET") {
      const auth = await requireManagedFeature(req, "app");
      const view = cleanText(req.query?.view);
      if (view) {
        if (!new Set(["rules", "active_run"]).has(view) || cleanText(req.query?.run_id)) {
          throw inputError("Automation view is invalid.");
        }
        const resource = view === "rules"
          ? "rpc/get_financial_reconciliation_automatic_manual_rules"
          : "rpc/get_financial_reconciliation_automatic_active_run";
        const body = view === "rules" ? {} : { p_actor: actorFor(auth) };
        const result = await restQuery(resource, {
          method: "POST",
          body,
        });
        return res.status(200).json(toAutomationPublicResult(result));
      }
      const runId = normalizeRunId(req.query?.run_id);
      const run = await restQuery("rpc/get_financial_reconciliation_automatic_run", {
        method: "POST",
        body: { p_run_id: runId },
      });
      return res.status(200).json(toAutomationPublicResult(run));
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const action = normalizeAutomationAction(body.action);
      const result = action === "analyze_rule"
        ? await analyzeRule(req, body)
        : action === "continue_analysis"
          ? await continueAnalysis(req, body)
          : await executeSelected(req, body);
      return res.status(200).json(result);
    }

    await requireManagedFeature(req, "app");
    res.setHeader("Allow", "GET, POST");
    return res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    return sendError(res, safePublicError(error));
  }
};
