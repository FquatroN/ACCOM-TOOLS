const { cleanText, parseBody, requireFeature, sendError } = require("./_supabase");
const {
  createSupplierEmailSchedule,
  deleteSupplierEmailSchedule,
  listSupplierEmailSchedules,
  updateSupplierEmailSchedule,
} = require("./_financial-docs-supplier-email-service");

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "settings", "financial-docs");
    if (req.method === "GET") {
      res.status(200).json({ rows: await listSupplierEmailSchedules() });
      return;
    }
    const body = await parseBody(req);
    const id = cleanText(req.query?.id || body?.id);
    if (req.method === "POST") {
      res.status(200).json({ row: await createSupplierEmailSchedule(body) });
      return;
    }
    if (req.method === "PUT") {
      res.status(200).json({ row: await updateSupplierEmailSchedule(id, body) });
      return;
    }
    if (req.method === "DELETE") {
      await deleteSupplierEmailSchedule(id);
      res.status(200).json({ ok: true });
      return;
    }
    res.setHeader("Allow", "GET,POST,PUT,DELETE");
    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
