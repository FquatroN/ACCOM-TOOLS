const { parseBody, requireFeature, sendError } = require("./_supabase");
const {
  loadFinancialDocsSettings,
  saveFinancialDocsSettings,
  safeFinancialDocsSettings,
} = require("./_financial-docs-service");
const { sanitizeFinancialDocsSettings } = require("./_financial-docs");

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "settings", "financial-docs");
    if (req.method === "GET") {
      const settings = await loadFinancialDocsSettings();
      res.status(200).json({ settings: safeFinancialDocsSettings(settings) });
      return;
    }
    if (req.method === "PUT") {
      const body = await parseBody(req);
      const current = await loadFinancialDocsSettings();
      const next = sanitizeFinancialDocsSettings({
        ...current,
        attributes: body?.attributes || current.attributes,
        drive: {
          ...current.drive,
          folderPath: body?.drive?.folderPath ?? current.drive.folderPath,
          baseFolderId: body?.drive?.baseFolderId ?? current.drive.baseFolderId,
        },
      });
      const saved = await saveFinancialDocsSettings(next);
      res.status(200).json({ settings: safeFinancialDocsSettings(saved) });
      return;
    }
    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
