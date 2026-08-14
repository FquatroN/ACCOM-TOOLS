const { cleanText, parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const { mapRpcError } = require("./_reconciliation");
const {
  normalizeAutomationSettingsPayload,
  toAutomationPublicResult,
  toAutomationSettingsRpcPayload,
} = require("./_reconciliation-automation");

function actorFor(auth) {
  return cleanText(auth.user?.email) || cleanText(auth.user?.id);
}

function permissionError() {
  const error = new Error("You do not have permission for this feature.");
  error.statusCode = 403;
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

async function requireManagedSettingsFeature(req) {
  const auth = await requireFeature(req, "settings", "financial-reconciliation");
  if (!cleanText(auth.access?.profile?.id)) throw permissionError();
  return auth;
}

module.exports = async function handler(req, res) {
  try {
    const auth = await requireManagedSettingsFeature(req);

    if (req.method === "GET") {
      const result = await restQuery("rpc/get_financial_reconciliation_automation_settings", {
        method: "POST",
        body: {},
      });
      return res.status(200).json(toAutomationPublicResult(result));
    }

    if (req.method === "PUT") {
      const settings = normalizeAutomationSettingsPayload(await parseBody(req));
      const result = await restQuery("rpc/replace_financial_reconciliation_automation_settings", {
        method: "POST",
        body: toAutomationSettingsRpcPayload(settings, actorFor(auth)),
      });
      return res.status(200).json(toAutomationPublicResult(result));
    }

    res.setHeader("Allow", "GET, PUT");
    return res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    return sendError(res, safePublicError(error));
  }
};
