const crypto = require("crypto");

const { cleanText, parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  normalizeImportDataType,
  sanitizeFdmAccountsImportRow,
  sanitizeFdmBookingsImportRow,
  sanitizeFdmSalesImportRow,
} = require("./_import-data");

const IMPORT_DATA_CONFIG = {
  "fdm-accounts": {
    table: "import_fdm_accounts",
    sanitize: sanitizeFdmAccountsImportRow,
    listQuery: "select=*&order=created_at.desc&limit=120",
    insertSelect: "select=id,import_batch,source_row_number",
    summaryQueries: {
      importDate: "select=created_at&order=created_at.desc&limit=1",
      specific: "select=date_time_raw,event_date,event_time,created_at&order=event_date.desc.nullslast,event_time.desc.nullslast,created_at.desc&limit=1",
    },
  },
  "fdm-bookings": {
    table: "import_fdm_bookings",
    sanitize: sanitizeFdmBookingsImportRow,
    listQuery: "select=*&order=created_at.desc&limit=120",
    insertSelect: "select=id,booking_number",
    onConflictColumns: ["booking_number"],
    summaryQueries: {
      importDate: "select=created_at&order=created_at.desc&limit=1",
      specific: "select=check_in_date&order=check_in_date.desc.nullslast,created_at.desc&limit=1",
    },
  },
  "fdm-sales": {
    table: "import_fdm_sales",
    sanitize: sanitizeFdmSalesImportRow,
    listQuery: "select=*&order=created_at.desc&limit=120",
    insertSelect: "select=id,reservation_id,sale_date,sale_time,sale_item,quantity,guest",
    summaryQueries: {
      importDate: "select=created_at&order=created_at.desc&limit=1",
      specific: "select=sale_date&order=sale_date.desc.nullslast,created_at.desc&limit=1",
    },
  },
};

function importBatchId(type) {
  const stamp = new Date().toISOString().replace(/[-:.TZ]/g, "").slice(0, 14);
  return `${type}-${stamp}-${crypto.randomUUID().slice(0, 8)}`;
}

function importDataConfig(type) {
  return IMPORT_DATA_CONFIG[normalizeImportDataType(type)] || IMPORT_DATA_CONFIG["fdm-accounts"];
}

function normalizeOnConflictColumns(config = {}) {
  if (Array.isArray(config.onConflictColumns) && config.onConflictColumns.length) {
    return config.onConflictColumns.map((column) => cleanText(column)).filter(Boolean);
  }
  const legacy = cleanText(config.onConflict);
  return legacy ? legacy.split(",").map((column) => cleanText(column)).filter(Boolean) : [];
}

function buildConflictKey(row, columns) {
  return columns.map((column) => cleanText(row?.[column])).join("\u001F");
}

function dedupeRowsByConflict(rows, columns) {
  if (!Array.isArray(rows) || !rows.length || !Array.isArray(columns) || !columns.length) return rows;
  return Array.from(rows.reduce((map, row) => {
    map.set(buildConflictKey(row, columns), row);
    return map;
  }, new Map()).values());
}

async function fetchImportDataMeta(type) {
  const normalizedType = normalizeImportDataType(type);
  const config = importDataConfig(normalizedType);
  const importRows = await restQuery(`${config.table}?${config.summaryQueries.importDate}`, { method: "GET" });
  const specificRows = await restQuery(`${config.table}?${config.summaryQueries.specific}`, { method: "GET" });
  const importRow = Array.isArray(importRows) ? importRows[0] || null : null;
  const specificRow = Array.isArray(specificRows) ? specificRows[0] || null : null;
  if (normalizedType === "fdm-bookings") {
    return {
      maxImportDate: cleanText(importRow?.created_at),
      maxCheckInDate: cleanText(specificRow?.check_in_date),
    };
  }
  if (normalizedType === "fdm-sales") {
    return {
      maxImportDate: cleanText(importRow?.created_at),
      maxDate: cleanText(specificRow?.sale_date),
    };
  }
  return {
    maxImportDate: cleanText(importRow?.created_at),
    maxDateTimeRaw: cleanText(specificRow?.date_time_raw),
  };
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "import-data");
    if (req.method === "GET") {
      const type = normalizeImportDataType(req.query?.type);
      if (cleanText(req.query?.summary) === "1") {
        const meta = await fetchImportDataMeta(type);
        res.status(200).json({ type, meta });
        return;
      }
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
      const onConflictColumns = normalizeOnConflictColumns(config);
      const rowsToInsert = dedupeRowsByConflict(rows, onConflictColumns);
      const created = await restQuery(`${config.table}?${config.insertSelect}${onConflictColumns.length ? `&on_conflict=${encodeURIComponent(onConflictColumns.join(","))}` : ""}`, {
        method: "POST",
        body: rowsToInsert,
        preferRepresentation: true,
        prefer: onConflictColumns.length ? "resolution=merge-duplicates" : "",
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
