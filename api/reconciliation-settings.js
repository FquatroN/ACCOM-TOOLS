const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const { normalizeReconciliationRules } = require("./_reconciliation");

const toRule = (row) => ({
  baseSourceType: row.base_source_type,
  matchingSourceType: row.matching_source_type,
  operator: row.operator,
});

const toRow = (rule) => ({
  base_source_type: rule.baseSourceType,
  matching_source_type: rule.matchingSourceType,
  operator: rule.operator,
});

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "settings", "financial-reconciliation");

    if (req.method === "GET") {
      const rules = await restQuery("financial_reconciliation_source_rules?select=base_source_type,matching_source_type,operator&order=base_source_type.asc,matching_source_type.asc", { method: "GET" });
      res.status(200).json({ rules: rules.map(toRule) });
      return;
    }

    if (req.method === "PUT") {
      const input = normalizeReconciliationRules((await parseBody(req))?.rules);
      await restQuery("rpc/replace_financial_reconciliation_source_rules", {
        method: "POST",
        body: { p_rules: input.map(toRow) },
      });
      res.status(200).json({ rules: input });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
