const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  CASH_CONTROL_SETTING_KEY,
  sanitizeCashControlPayload,
  sanitizeCashControlRecord,
  sanitizeCashControlRecords,
  validateCashControlRecord,
} = require("./_cash-control");

function cleanId(value) {
  return String(value || "").trim();
}

async function loadCashPayloadRow() {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${encodeURIComponent(CASH_CONTROL_SETTING_KEY)}&limit=1`, {
    method: "GET",
  });
  const row = Array.isArray(rows) && rows[0] ? rows[0] : null;
  return {
    rowId: row?.id || "",
    payload: sanitizeCashControlPayload(row?.payload || {}),
  };
}

async function saveCashPayload(rowId, payload) {
  const safe = sanitizeCashControlPayload(payload);
  if (rowId) {
    await restQuery(`app_settings?id=eq.${encodeURIComponent(rowId)}`, {
      method: "PATCH",
      body: { payload: safe, updated_at: new Date().toISOString() },
    });
    return safe;
  }
  const created = await restQuery("app_settings", {
    method: "POST",
    body: [{ setting_key: CASH_CONTROL_SETTING_KEY, payload: safe }],
  });
  return sanitizeCashControlPayload(Array.isArray(created) && created[0]?.payload ? created[0].payload : safe);
}

function mergeCashRecord(records, input, settings, id = "") {
  const existing = records.find((row) => cleanId(row.id) === cleanId(id)) || {};
  const nextRecord = sanitizeCashControlRecord({ ...existing, ...input, id: id || input?.id || existing.id }, settings, existing);
  validateCashControlRecord(records, nextRecord, settings, { excludeId: id || nextRecord.id, isCreate: !id });
  const nextRows = [...records];
  const index = id ? nextRows.findIndex((row) => cleanId(row.id) === cleanId(id)) : -1;
  if (index >= 0) {
    nextRows[index] = {
      ...nextRecord,
      id: nextRows[index].id,
      createdAt: nextRows[index].createdAt || nextRecord.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
  } else {
    nextRows.push({
      ...nextRecord,
      createdAt: nextRecord.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });
  }
  return sanitizeCashControlRecords(nextRows, settings);
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "cash");

    if (req.method === "GET") {
      const { payload } = await loadCashPayloadRow();
      res.status(200).json({ rows: payload.records, settings: payload.settings });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const { rowId, payload } = await loadCashPayloadRow();
      const nextRows = mergeCashRecord(payload.records, body, payload.settings);
      const saved = await saveCashPayload(rowId, { settings: payload.settings, records: nextRows });
      res.status(200).json({ rows: saved.records, settings: saved.settings });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanId(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Record id is required." });
        return;
      }
      const body = await parseBody(req);
      const { rowId, payload } = await loadCashPayloadRow();
      const existing = payload.records.find((row) => cleanId(row.id) === id);
      if (!existing) {
        res.status(404).json({ error: "Record not found." });
        return;
      }
      const nextRows = mergeCashRecord(payload.records, body, payload.settings, id);
      const saved = await saveCashPayload(rowId, { settings: payload.settings, records: nextRows });
      res.status(200).json({ rows: saved.records, settings: saved.settings });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
