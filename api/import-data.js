const crypto = require("crypto");

const { cleanText, parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  normalizeImportDataType,
  sanitizeFdmAccountsImportRow,
  sanitizeFdmBookingsImportRow,
} = require("./_import-data");

const IMPORT_DATA_CONFIG = {
  "fdm-accounts": {
    table: "import_fdm_accounts",
    sanitize: sanitizeFdmAccountsImportRow,
    listQuery: "select=*&order=created_at.desc&limit=120",
    insertSelect: "select=id,import_batch,source_row_number",
  },
  "fdm-bookings": {
    table: "import_fdm_bookings",
    sanitize: sanitizeFdmBookingsImportRow,
    listQuery: "select=*&order=created_at.desc&limit=120",
    insertSelect: "select=id,booking_number",
    onConflict: "booking_number",
  },
};

function importBatchId(type) {
  const stamp = new Date().toISOString().replace(/[-:.TZ]/g, "").slice(0, 14);
  return `${type}-${stamp}-${crypto.randomUUID().slice(0, 8)}`;
}

function importDataConfig(type) {
  return IMPORT_DATA_CONFIG[normalizeImportDataType(type)] || IMPORT_DATA_CONFIG["fdm-accounts"];
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "import-data");
    if (req.method === "GET") {
      const type = normalizeImportDataType(req.query?.type);
      const config = importDataConfig(type);
      const rows = await restQuery(`${config.table}?${config.listQuery}`, { method: "GET" });
      res.status(200).json({ type, rows: Array.isArray(rows) ? rows : [] });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const type = normalizeImportDataType(body?.type);
      const config = importDataConfig(type);
      const sourceName = String(body?.sourceName || "").trim();
      const previewRows = Array.isArray(body?.rows) ? body.rows : [];
      if (!previewRows.length) {
        res.status(400).json({ error: "No rows to import." });
        return;
      }
      const batch = importBatchId(type);
      const rows = previewRows.map((row, index) =>
        config.sanitize(row, {
          importBatch: batch,
          sourceName,
          sourceRowNumber: row?.sourceRowNumber || index + 2,
        })
      );
      const rowsToInsert = config.onConflict
        ? Array.from(rows.reduce((map, row) => {
          map.set(cleanText(row[config.onConflict]), row);
          return map;
        }, new Map()).values())
        : rows;
      const created = await restQuery(`${config.table}?${config.insertSelect}${config.onConflict ? `&on_conflict=${encodeURIComponent(config.onConflict)}` : ""}`, {
        method: "POST",
        body: rowsToInsert,
        preferRepresentation: true,
        prefer: config.onConflict ? "resolution=merge-duplicates" : "",
      });
      res.status(200).json({
        ok: true,
        type,
        importBatch: batch,
        insertedCount: Array.isArray(created) ? created.length : rowsToInsert.length,
      });
      return;
    }

    if (req.method === "PUT") {
      const body = await parseBody(req);
      const type = normalizeImportDataType(body?.type);
      const config = importDataConfig(type);
      const id = cleanText(body?.id);
      if (!id) {
        res.status(400).json({ error: "Record id is required." });
        return;
      }
      const existingRows = await restQuery(`${config.table}?select=*&id=eq.${encodeURIComponent(id)}&limit=1`, { method: "GET" });
      const existing = Array.isArray(existingRows) ? existingRows[0] : null;
      if (!existing) {
        res.status(404).json({ error: "Imported record not found." });
        return;
      }
      const updatedPayload = config.sanitize(body?.row || {}, {
        importBatch: body?.row?.import_batch || existing.import_batch,
        sourceName: body?.row?.source_name || existing.source_name,
        sourceRowNumber: body?.row?.source_row_number || existing.source_row_number,
      });
      const updatedRows = await restQuery(`${config.table}?id=eq.${encodeURIComponent(id)}&select=*`, {
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
      const config = importDataConfig(type);
      const id = cleanText(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Record id is required." });
        return;
      }
      await restQuery(`${config.table}?id=eq.${encodeURIComponent(id)}`, { method: "DELETE" });
      res.status(200).json({ ok: true });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
