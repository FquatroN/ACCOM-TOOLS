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

function isMissingCashTableError(error) {
  const message = String(error?.message || "").toLowerCase();
  return message.includes("cash_control_records") && (
    message.includes("could not find") ||
    message.includes("schema cache") ||
    message.includes("relation") ||
    message.includes("does not exist")
  );
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

function mapCashTableRow(row, settings) {
  return sanitizeCashControlRecord({
    id: row?.id,
    day: row?.record_day ?? row?.day,
    shiftId: row?.shift_id,
    shiftName: row?.shift_name,
    name: row?.name,
    denominations: row?.denominations,
    cardPos: row?.card_pos,
    cashFdm: row?.cash_fdm,
    cardFdm: row?.card_fdm,
    justification: row?.justification,
    itemCounts: row?.item_counts,
    itemJustifications: row?.item_justifications,
    createdAt: row?.created_at,
    updatedAt: row?.updated_at,
  }, settings);
}

async function loadCashTableRows(settings) {
  const rows = await restQuery("cash_control_records?select=*", { method: "GET" });
  return sanitizeCashControlRecords((Array.isArray(rows) ? rows : []).map((row) => mapCashTableRow(row, settings)), settings);
}

function buildCashTableBody(record, existing = {}) {
  return {
    id: cleanId(record.id || existing.id) || undefined,
    record_day: record.day,
    shift_id: record.shiftId,
    shift_name: record.shiftName,
    name: record.name,
    denominations: record.denominations,
    card_pos: Number(record.cardPos || 0),
    cash_fdm: Number(record.cashFdm || 0),
    card_fdm: Number(record.cardFdm || 0),
    justification: record.justification || "",
    item_counts: record.itemCounts || {},
    item_justifications: record.itemJustifications || {},
    created_at: existing.createdAt || record.createdAt || new Date().toISOString(),
    updated_at: new Date().toISOString(),
  };
}

async function createCashTableRows(records) {
  if (!Array.isArray(records) || records.length === 0) return;
  await restQuery("cash_control_records", {
    method: "POST",
    body: records.map((record) => buildCashTableBody(record)),
  });
}

async function updateCashTableRow(id, record, existing = {}) {
  await restQuery(`cash_control_records?id=eq.${encodeURIComponent(id)}`, {
    method: "PATCH",
    body: buildCashTableBody(record, existing),
    preferRepresentation: true,
  });
}

function mergeCashRecord(records, input, settings, id = "") {
  const existing = records.find((row) => cleanId(row.id) === cleanId(id)) || {};
  const nextRecord = sanitizeCashControlRecord({ ...existing, ...input, id: id || input?.id || existing.id }, settings, existing);
  validateCashControlRecord(nextRecord, records, settings, { excludeId: id || nextRecord.id, isCreate: !id });
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

async function seedCashTableIfEmpty(settings, legacyRecords) {
  const safe = sanitizeCashControlRecords(legacyRecords, settings);
  if (!safe.length) return [];
  await createCashTableRows(safe.map((record) => ({
    ...record,
    createdAt: record.createdAt || new Date().toISOString(),
    updatedAt: record.updatedAt || new Date().toISOString(),
  })));
  return loadCashTableRows(settings);
}

async function loadRecordsAndSettings() {
  const { payload } = await loadCashPayloadRow();
  const settings = payload.settings;
  try {
    const rows = await loadCashTableRows(settings);
    if (rows.length > 0) return { mode: "table", settings, rows };
    const seededRows = await seedCashTableIfEmpty(settings, payload.records);
    return { mode: "table", settings, rows: seededRows };
  } catch (error) {
    if (!isMissingCashTableError(error)) throw error;
    return { mode: "legacy", settings, rows: payload.records };
  }
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "cash");

    if (req.method === "GET") {
      const current = await loadRecordsAndSettings();
      res.status(200).json({ rows: current.rows, settings: current.settings });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const current = await loadRecordsAndSettings();
      if (current.mode === "legacy") {
        const { rowId, payload } = await loadCashPayloadRow();
        const nextRows = mergeCashRecord(payload.records, body, payload.settings);
        const saved = await saveCashPayload(rowId, { settings: payload.settings, records: nextRows });
        res.status(200).json({ rows: saved.records, settings: saved.settings });
        return;
      }
      const nextRecord = sanitizeCashControlRecord(body, current.settings);
      validateCashControlRecord(nextRecord, current.rows, current.settings, { isCreate: true });
      await createCashTableRows([{
        ...nextRecord,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      }]);
      const rows = await loadCashTableRows(current.settings);
      res.status(200).json({ rows, settings: current.settings });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanId(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Record id is required." });
        return;
      }
      const body = await parseBody(req);
      const current = await loadRecordsAndSettings();
      if (current.mode === "legacy") {
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
      const existing = current.rows.find((row) => cleanId(row.id) === id);
      if (!existing) {
        res.status(404).json({ error: "Record not found." });
        return;
      }
      const updated = sanitizeCashControlRecord({ ...existing, ...body, id: existing.id }, current.settings, existing);
      validateCashControlRecord(updated, current.rows, current.settings, { excludeId: existing.id, isCreate: false });
      await updateCashTableRow(existing.id, updated, existing);
      const rows = await loadCashTableRows(current.settings);
      res.status(200).json({ rows, settings: current.settings });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
