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

module.exports = async function handler(req, res) {
  try {
    const auth = await requireFeature(req, "settings", "financial-reconciliation");

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
    return sendError(res, mapRpcError(error));
  }
};
