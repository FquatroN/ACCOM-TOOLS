const { cleanText, parseBody, requireFeature, sendError } = require("./_supabase");
const {
  createSupplierEmailSchedule,
  deleteSupplierEmailSchedule,
  listSupplierEmailSchedules,
  updateSupplierEmailSchedule,
} = require("./_financial-docs-supplier-email-service");
const { sendSupplierInvoiceTestEmail } = require("./_financial-docs-supplier-email-automation");

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "settings", "financial-docs");
    if (req.method === "GET") {
      res.status(200).json({ rows: await listSupplierEmailSchedules() });
      return;
    }
    const body = await parseBody(req);
    const id = cleanText(req.query?.id || body?.id);
    if (req.method === "POST" && String(req.query?.action || "") === "test") {
      const schedule = (await listSupplierEmailSchedules()).find((row) => row.id === id);
      if (!schedule) { res.status(404).json({ error: "Supplier email schedule not found." }); return; }
      const result = await sendSupplierInvoiceTestEmail({ schedule, recipient: body?.testEmail });
      res.status(200).json({ ok: true, messageId: cleanText(result?.id) });
      return;
    }
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
