const { cleanText, requireFeature, restQuery, sendError } = require("./_supabase");
const { mapRpcError } = require("./_reconciliation");
const { toAutomationPublicResult } = require("./_reconciliation-automation");

const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const MEMBER_ROLES = new Set(["source", "destination"]);

function statusError(message, statusCode) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

function inputError(message) {
  return statusError(message, 400);
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

async function requireManagedFeature(req) {
  const auth = await requireFeature(req, "app", "financial-reconciliation");
  if (!cleanText(auth.access?.profile?.id)) {
    throw statusError("You do not have permission for this feature.", 403);
  }
  return auth;
}

function actorFor(auth) {
  const actor = cleanText(auth.user?.email) || cleanText(auth.user?.id);
  if (!actor) throw statusError("Authenticated user identity is unavailable.", 403);
  return actor;
}

function normalizeUuid(value, label) {
  if (typeof value !== "string" || !UUID_PATTERN.test(value)) {
    throw inputError(`${label} must be a valid UUID.`);
  }
  return value.toLowerCase();
}

function normalizeInteger(value, label, minimum, maximum = Number.MAX_SAFE_INTEGER) {
  if (typeof value !== "string" || !/^(?:0|[1-9]\d*)$/.test(value)) {
    throw inputError(`${label} must be between ${minimum} and ${maximum}.`);
  }
  const normalized = Number(value);
  if (!Number.isSafeInteger(normalized) || normalized < minimum || normalized > maximum) {
    throw inputError(`${label} must be between ${minimum} and ${maximum}.`);
  }
  return normalized;
}

module.exports = async function handler(req, res) {
  try {
    const auth = await requireManagedFeature(req);
    if (req.method !== "GET") {
      res.setHeader("Allow", "GET");
      return res.status(405).json({ error: "Method not allowed." });
    }

    const runId = normalizeUuid(req.query?.run_id, "Run ID");
    const proposalId = normalizeUuid(req.query?.proposal_id, "Proposal ID");
    const role = req.query?.role;
    if (typeof role !== "string" || !MEMBER_ROLES.has(role)) {
      throw inputError("Member role must be source or destination.");
    }
    const offset = normalizeInteger(req.query?.offset, "Offset", 0);
    const limit = normalizeInteger(req.query?.limit, "Limit", 1, 50);
    const actor = actorFor(auth);
    const result = await restQuery(
      "rpc/get_financial_reconciliation_automatic_proposal_members",
      {
        method: "POST",
        body: {
          p_run_id: runId,
          p_proposal_id: proposalId,
          p_role: role,
          p_offset: offset,
          p_limit: limit,
          p_actor: actor,
        },
      },
    );
    return res.status(200).json(toAutomationPublicResult(result));
  } catch (error) {
    return sendError(res, safePublicError(error));
  }
};
