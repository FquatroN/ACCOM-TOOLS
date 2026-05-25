const { cleanText, requireFeature, sendError } = require("./_supabase");
const {
  downloadDriveFile,
  loadFinancialDocumentRowById,
  loadFinancialDocsSettings,
  refreshDriveAccessToken,
} = require("./_financial-docs-service");

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "financial-docs");
    const id = cleanText(req.query?.id);
    if (!id) {
      res.status(400).json({ error: "Document id is required." });
      return;
    }
    const row = await loadFinancialDocumentRowById(id);
    if (!row) {
      res.status(404).json({ error: "Financial document not found." });
      return;
    }
    const driveFileId = cleanText(row.drive_file_id);
    if (!driveFileId) {
      res.status(404).json({ error: "This document has no attachment." });
      return;
    }
    const settings = await loadFinancialDocsSettings();
    const refreshed = await refreshDriveAccessToken(settings);
    const file = await downloadDriveFile(refreshed.accessToken, driveFileId);
    const download = cleanText(req.query?.download) === "1";
    const filename = cleanText(row.stored_filename || row.original_filename) || "document.pdf";
    res.setHeader("Content-Type", cleanText(row.mime_type) || file.mimeType || "application/pdf");
    res.setHeader("Content-Disposition", `${download ? "attachment" : "inline"}; filename="${filename.replace(/"/g, "")}"`);
    res.status(200).send(file.buffer);
  } catch (error) {
    sendError(res, error);
  }
};
