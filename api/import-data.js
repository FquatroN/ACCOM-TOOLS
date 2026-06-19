const crypto = require("crypto");

const { cleanText, parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const { normalizeImportDataType, sanitizeFdmAccountsImportRow } = require("./_import-data");

function importBatchId(type) {
  const stamp = new Date().toISOString().replace(/[-:.TZ]/g, "").slice(0, 14);
  return `${type}-${stamp}-${crypto.randomUUID().slice(0, 8)}`;
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "import-data");
    if (req.method === "GET") {
      const type = normalizeImportDataType(req.query?.type);
      if (type !== "fdm-accounts") {
        res.status(400).json({ error: "Unsupported import data type." });
        return;
      }
      const rows = await restQuery("import_fdm_accounts?select=*&order=created_at.desc&limit=120", { method: "GET" });
      res.status(200).json({ type, rows: Array.isArray(rows) ? rows : [] });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const type = normalizeImportDataType(body?.type);
      if (type !== "fdm-accounts") {
        res.status(400).json({ error: "Unsupported import data type." });
        return;
      }
      const sourceName = String(body?.sourceName || "").trim();
      const previewRows = Array.isArray(body?.rows) ? body.rows : [];
      if (!previewRows.length) {
        res.status(400).json({ error: "No rows to import." });
        return;
      }
      const batch = importBatchId(type);
      const rows = previewRows.map((row, index) =>
        sanitizeFdmAccountsImportRow(row, {
          importBatch: batch,
          sourceName,
          sourceRowNumber: row?.sourceRowNumber || index + 2,
        })
      );
      const created = await restQuery("import_fdm_accounts?select=id,import_batch,source_row_number", {
        method: "POST",
        body: rows,
        preferRepresentation: true,
      });
      res.status(200).json({
        ok: true,
        type,
        importBatch: batch,
        insertedCount: Array.isArray(created) ? created.length : rows.length,
      });
      return;
    }

    if (req.method === "PUT") {
      const body = await parseBody(req);
      const type = normalizeImportDataType(body?.type);
      const id = cleanText(body?.id);
      if (type !== "fdm-accounts") {
        res.status(400).json({ error: "Unsupported import data type." });
        return;
      }
      if (!id) {
        res.status(400).json({ error: "Record id is required." });
        return;
      }
      const existingRows = await restQuery(`import_fdm_accounts?select=*&id=eq.${encodeURIComponent(id)}&limit=1`, { method: "GET" });
      const existing = Array.isArray(existingRows) ? existingRows[0] : null;
      if (!existing) {
        res.status(404).json({ error: "Imported record not found." });
        return;
      }
      const updatedPayload = sanitizeFdmAccountsImportRow(body?.row || {}, {
        importBatch: body?.row?.import_batch || existing.import_batch,
        sourceName: body?.row?.source_name || existing.source_name,
        sourceRowNumber: body?.row?.source_row_number || existing.source_row_number,
      });
      const updatedRows = await restQuery(`import_fdm_accounts?id=eq.${encodeURIComponent(id)}&select=*`, {
        method: "PATCH",
        body: updatedPayload,
        preferRepresentation: true,
      });
      res.status(200).json({
        ok: true,
        row: Array.isArray(updatedRows) ? updatedRows[0] || null : null,
      });
      return;
    }

    if (req.method === "DELETE") {
      const type = normalizeImportDataType(req.query?.type);
      const id = cleanText(req.query?.id);
      if (type !== "fdm-accounts") {
        res.status(400).json({ error: "Unsupported import data type." });
        return;
      }
      if (!id) {
        res.status(400).json({ error: "Record id is required." });
        return;
      }
      await restQuery(`import_fdm_accounts?id=eq.${encodeURIComponent(id)}`, { method: "DELETE" });
      res.status(200).json({ ok: true });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
