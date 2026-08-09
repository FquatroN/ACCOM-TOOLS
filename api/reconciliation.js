const { cleanText, parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const { validateMutation, validateWorkspaceQuery, mapRpcError } = require("./_reconciliation");

module.exports = async function handler(req, res) {
  try {
    const auth = await requireFeature(req, "app", "financial-reconciliation");

    if (req.method === "GET") {
      const query = validateWorkspaceQuery(req.query || {});
      const workspace = await restQuery("rpc/get_financial_reconciliation_workspace", {
        method: "POST",
        body: {
          p_reconciliation_id: query.reconciliationId || null,
          p_source_type: query.sourceType,
          p_matching_source_types: query.matchingSourceTypes,
          p_filters: query.filters,
          p_page: query.page,
          p_page_size: query.pageSize,
        },
      });
      return res.status(200).json(workspace);
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const input = validateMutation(cleanText(body.action), body);
      const result = await restQuery("rpc/financial_reconciliation_action", {
        method: "POST",
        body: {
          p_action: input.action,
          p_actor: cleanText(auth.user?.email) || cleanText(auth.user?.id),
          p_reconciliation_id: input.reconciliationId || null,
          p_base_source_type: input.baseSourceType || null,
          p_matching_source_types: input.matchingSourceTypes || null,
          p_source_type: input.sourceType || null,
          p_source_id: input.sourceId || null,
          p_comment: input.comment || null,
        },
      });
      return res.status(200).json(result);
    }

    res.setHeader("Allow", "GET, POST");
    return res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    return sendError(res, mapRpcError(error));
  }
};
