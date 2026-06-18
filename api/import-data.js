const crypto = require("crypto");

const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
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

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
