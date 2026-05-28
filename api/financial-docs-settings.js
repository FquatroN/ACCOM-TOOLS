const { parseBody, requireFeature, sendError } = require("./_supabase");
const {
  loadFinancialDocsSettings,
  redirectUri,
  saveFinancialDocsSettings,
  safeFinancialDocsSettings,
} = require("./_financial-docs-service");
const { extractGoogleDriveFolderId, sanitizeFinancialDocsSettings } = require("./_financial-docs");

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "settings", "financial-docs");
    const clientId = process.env.GOOGLE_CLIENT_ID || "";
    if (req.method === "GET") {
      const settings = await loadFinancialDocsSettings();
      const safeSettings = safeFinancialDocsSettings(settings);
      res.status(200).json({ settings: { ...safeSettings, drive: { ...safeSettings.drive, redirectUri: redirectUri(req), clientId } } });
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
          baseFolderId:
            body?.drive?.baseFolderId ??
            extractGoogleDriveFolderId(body?.drive?.folderPath) ??
            current.drive.baseFolderId,
        },
        rules: body?.rules ?? current.rules,
      });
      const saved = await saveFinancialDocsSettings(next);
      const safeSettings = safeFinancialDocsSettings(saved);
      res.status(200).json({ settings: { ...safeSettings, drive: { ...safeSettings.drive, redirectUri: redirectUri(req), clientId } } });
      return;
    }
    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
