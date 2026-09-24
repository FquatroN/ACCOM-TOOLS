const { requireFeature, sendError } = require("./_supabase");
const { processDueSupplierEmailSchedules } = require("./_financial-docs-supplier-email-automation");

module.exports = async function handler(req, res) {
  try {
    if (req.method !== "GET" && req.method !== "POST") {
      res.setHeader("Allow", "GET,POST"); res.status(405).json({ error: "Method not allowed." }); return;
    }
    const bearer = String(req.headers.authorization || "").replace(/^Bearer\s+/i, "");
    const secret = String(process.env.CRON_SECRET || "").trim();
    const isCron = !!req.headers["x-vercel-cron"] || (!!secret && bearer === secret) || String(req.headers["user-agent"] || "").toLowerCase().includes("vercel-cron");
    if (!isCron) await requireFeature(req, "settings", "financial-docs");
    const results = await processDueSupplierEmailSchedules();
    res.status(200).json({ ok: true, processed: results.length, results });
  } catch (error) { sendError(res, error); }
};
