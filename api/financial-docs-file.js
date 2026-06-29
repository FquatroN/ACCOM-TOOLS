const { cleanText, requireFeature, sendError } = require("./_supabase");
const {
  deleteDriveFile,
  downloadDriveFile,
  insertFinancialDocumentHistory,
  loadFinancialDocumentRowById,
  loadFinancialDocumentWithHistory,
  loadFinancialDocsSettings,
  refreshDriveAccessToken,
  updateFinancialDocumentRow,
} = require("./_financial-docs-service");

module.exports = async function handler(req, res) {
  try {
    const auth = await requireFeature(req, "app", "financial-docs");
    const userEmail = cleanText(auth.user?.email) || cleanText(auth.user?.id);
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

    if (req.method === "DELETE") {
      await deleteDriveFile(refreshed.accessToken, driveFileId);
      await updateFinancialDocumentRow(id, {
        drive_file_id: null,
        drive_folder_id: null,
        drive_file_url: null,
        original_filename: null,
        stored_filename: null,
        mime_type: null,
        file_size: null,
        file_hash: null,
        uploaded_by: null,
        uploaded_at: null,
      });
      await insertFinancialDocumentHistory({
        document_id: id,
        action_type: "file_deleted",
        field_name: "",
        message: "Attachment deleted.",
        old_value: cleanText(row.stored_filename || row.original_filename || driveFileId),
        new_value: null,
        metadata: { driveFileId },
        created_by: userEmail,
      });
      const updated = await loadFinancialDocumentWithHistory(id);
      res.status(200).json({ row: updated });
      return;
    }

    if (req.method !== "GET") {
      res.status(405).json({ error: "Method not allowed." });
      return;
    }

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
