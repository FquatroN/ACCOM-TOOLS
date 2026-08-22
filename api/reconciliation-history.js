const { requireFeature, restQuery, sendError } = require("./_supabase");
const { mapRpcError, validateHistoryQuery } = require("./_reconciliation");

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "financial-reconciliation");
    if (req.method !== "GET") {
      res.setHeader("Allow", "GET");
      return res.status(405).json({ error: "Method not allowed." });
    }

    const query = validateHistoryQuery(req.query || {});
    const result = await restQuery("rpc/get_financial_reconciliation_history", {
      method: "POST",
      body: {
        p_created_from: query.createdFrom || null,
        p_created_to: query.createdTo || null,
        p_origin: query.origin || null,
        p_status: query.status || null,
        p_difference_from: query.differenceFrom,
        p_difference_to: query.differenceTo,
        p_page: query.page,
        p_page_size: query.pageSize,
      },
    });
    return res.status(200).json(result);
  } catch (error) {
    return sendError(res, mapRpcError(error));
  }
};
