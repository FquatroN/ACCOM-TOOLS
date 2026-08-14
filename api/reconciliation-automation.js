const { cleanText, parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const { mapRpcError } = require("./_reconciliation");
const {
  normalizeAnalyzePayload,
  normalizeAutomationAction,
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

function normalizeRunId(value) {
  if (typeof value !== "string" || !UUID_PATTERN.test(value)) {
    throw inputError("Run ID must be a valid UUID.");
  }
  return value;
}

function requireClientRequestId(input) {
  if (!input.clientRequestId) throw inputError("Client request ID is required.");
  return input;
}

function requireBatchFields(body) {
  if (!body || typeof body !== "object" || Array.isArray(body)) {
    throw inputError("Analyze payload must be an object.");
  }
  for (const key of Object.keys(body)) {
    if (key !== "action" && key !== "clientRequestId") {
      throw inputError("Analyze batch payload contains an unsupported field.");
    }
  }
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
  const auth = await requireFeature(req, "app", "financial-reconciliation");
  const input = requireClientRequestId(normalizeAnalyzePayload(body));
  if (input.action !== "analyze_rule" || input.ruleKeys.length !== 1) {
    throw inputError("Analyze rule requires exactly one manually enabled rule.");
  }
  return createAnalysis(input, actorFor(auth), "manual_rule");
}

async function analyzeBatch(req, body) {
  const auth = await requireFeature(req, "settings", "financial-reconciliation");
  if (!cleanText(auth.access?.profile?.id)) {
    throw statusError("You do not have permission for this feature.", 403);
  }
  requireBatchFields(body);
  const settings = toAutomationPublicResult(await restQuery(
    "rpc/get_financial_reconciliation_automation_settings",
    { method: "POST", body: {} },
  ));
  const ruleKeys = (Array.isArray(settings?.rules) ? settings.rules : [])
    .filter((rule) => rule?.enabled === true && rule?.includeInScheduledBatch === true)
    .map((rule) => rule.ruleKey);
  const input = requireClientRequestId(normalizeAnalyzePayload({
    action: "analyze_batch",
    ruleKeys,
    clientRequestId: body.clientRequestId,
  }));
  return createAnalysis(input, actorFor(auth), "manual_batch");
}

async function executeSelected(req, body) {
  const auth = await requireFeature(req, "app", "financial-reconciliation");
  const input = normalizeExecutePayload(body);
  const actor = actorFor(auth);
  const outcomes = [];
  const run = toAutomationPublicResult(await restQuery(
    "rpc/get_financial_reconciliation_automatic_run",
    { method: "POST", body: { p_run_id: input.runId } },
  ));
  if (run?.runId !== input.runId || run?.trigger !== "manual") {
    throw inputError("Selected proposals must belong to the requested manual run.");
  }
  if (run.finishedAt) {
    throw statusError("The requested automation run is already finished.", 409);
  }
  if (cleanText(run.actor) !== actor) {
    throw statusError("You do not have permission for this automation run.", 403);
  }
  const proposalIds = new Set((Array.isArray(run.proposals) ? run.proposals : []).map((proposal) => proposal?.id));
  if (input.proposalIds.some((proposalId) => !proposalIds.has(proposalId))) {
    throw inputError("Selected proposals must belong to the requested manual run.");
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

  const finalizedRun = await restQuery("rpc/finish_financial_reconciliation_automatic_run", {
    method: "POST",
    body: { p_run_id: input.runId },
  });
  return { run: toAutomationPublicResult(finalizedRun), outcomes };
}

module.exports = async function handler(req, res) {
  try {
    if (req.method === "GET") {
      await requireFeature(req, "app", "financial-reconciliation");
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
        : action === "analyze_batch"
          ? await analyzeBatch(req, body)
          : await executeSelected(req, body);
      return res.status(200).json(result);
    }

    await requireFeature(req, "app", "financial-reconciliation");
    res.setHeader("Allow", "GET, POST");
    return res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    return sendError(res, safePublicError(error));
  }
};
