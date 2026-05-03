const { randomUUID } = require("node:crypto");
const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  DEFAULT_LAUNDRY_SETTINGS,
  LAUNDRY_SETTING_KEY,
  sanitizeLaundryPayload,
  sanitizeLaundryRecord,
  sanitizeLaundryRecords,
  sanitizeLaundrySettings,
} = require("./_laundry");

function recordKey(record) {
  return `${record.property}::${record.date}`;
}

function shiftDate(value, days) {
  const date = new Date(`${String(value || "").trim()}T00:00:00`);
  if (Number.isNaN(date.getTime())) return "";
  date.setDate(date.getDate() + days);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function latestDateForProperty(rows, property, excludeId = "") {
  return rows
    .filter((row) => row.property === property && row.id !== excludeId)
    .map((row) => row.date)
    .filter(Boolean)
    .sort()
    .at(-1) || "";
}

function validateLaundrySave(existingRows, record, { isCreate = false, excludeId = "" } = {}) {
  const duplicate = existingRows.find((row) => row.property === record.property && row.date === record.date && row.id !== excludeId);
  if (duplicate) {
    const error = new Error(`A laundry record for ${record.property} on ${record.date} already exists.`);
    error.statusCode = 400;
    throw error;
  }
  if (!isCreate) return;
  const latest = latestDateForProperty(existingRows, record.property);
  if (!latest) return;
  const expected = shiftDate(latest, 1);
  if (record.date !== expected) {
    const error = new Error(`The next record for ${record.property} must be created for ${expected}.`);
    error.statusCode = 400;
    throw error;
  }
}

function isMissingLaundryTableError(error) {
  const message = String(error?.message || "").toLowerCase();
  return message.includes("laundry_records") && (
    message.includes("could not find") ||
    message.includes("schema cache") ||
    message.includes("relation") ||
    message.includes("does not exist")
  );
}

async function loadLaundryPayloadRow() {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${encodeURIComponent(LAUNDRY_SETTING_KEY)}&limit=1`, {
    method: "GET",
  });
  const row = Array.isArray(rows) && rows[0] ? rows[0] : null;
  return {
    rowId: row?.id || "",
    payload: sanitizeLaundryPayload(row?.payload || { settings: DEFAULT_LAUNDRY_SETTINGS, records: [] }),
  };
}

async function saveLaundryPayload(rowId, payload) {
  const safe = sanitizeLaundryPayload(payload);
  if (rowId) {
    await restQuery(`app_settings?id=eq.${encodeURIComponent(rowId)}`, {
      method: "PATCH",
      body: { payload: safe, updated_at: new Date().toISOString() },
    });
    return safe;
  }
  const created = await restQuery("app_settings", {
    method: "POST",
    body: [{ setting_key: LAUNDRY_SETTING_KEY, payload: safe }],
  });
  return sanitizeLaundryPayload(Array.isArray(created) && created[0]?.payload ? created[0].payload : safe);
}

function mapLaundryTableRow(row, settings) {
  return sanitizeLaundryRecord({
    id: row?.id,
    property: row?.property,
    date: row?.record_date,
    sentItems: row?.sent_items,
    receivedItems: row?.received_items,
    receivedWeightKg: row?.received_weight_kg,
    notes: row?.notes,
    createdAt: row?.created_at,
    updatedAt: row?.updated_at,
  }, settings);
}

async function loadLaundryTableRows(settings) {
  const rows = await restQuery(
    "laundry_records?select=id,property,record_date,sent_items,received_items,received_weight_kg,notes,created_at,updated_at&order=record_date.desc,property.asc",
    { method: "GET" }
  );
  return sanitizeLaundryRecords((Array.isArray(rows) ? rows : []).map((row) => mapLaundryTableRow(row, settings)), settings);
}

function buildLaundryTableBody(record, existing = {}) {
  return {
    id: record.id || existing.id || randomUUID(),
    property: record.property,
    record_date: record.date,
    sent_items: record.sentItems,
    received_items: record.receivedItems,
    received_weight_kg: Number(record.receivedWeightKg || 0),
    notes: record.notes || "",
    created_at: existing.createdAt || record.createdAt || new Date().toISOString(),
    updated_at: new Date().toISOString(),
  };
}

async function createLaundryTableRow(record, existing = {}) {
  const created = await restQuery("laundry_records", {
    method: "POST",
    body: [buildLaundryTableBody(record, existing)],
    preferRepresentation: true,
  });
  const row = Array.isArray(created) ? created[0] : created;
  return mapLaundryTableRow(row, sanitizeLaundrySettings(DEFAULT_LAUNDRY_SETTINGS));
}

async function updateLaundryTableRow(id, record, existing = {}) {
  const updated = await restQuery(`laundry_records?id=eq.${encodeURIComponent(id)}`, {
    method: "PATCH",
    body: buildLaundryTableBody(record, existing),
    preferRepresentation: true,
  });
  const row = Array.isArray(updated) ? updated[0] : updated;
  return mapLaundryTableRow(row, sanitizeLaundrySettings(DEFAULT_LAUNDRY_SETTINGS));
}

function mergeRecords(existingRecords, incomingRecords, settings, { allowMerge = false } = {}) {
  const safeSettings = sanitizeLaundrySettings(settings);
  const next = [...existingRecords];
  incomingRecords.forEach((item) => {
    const safe = sanitizeLaundryRecord(item, safeSettings);
    const existingIndex = next.findIndex((record) => record.id && record.id === safe.id);
    const duplicateIndex = next.findIndex((record) => recordKey(record) === recordKey(safe));
    if (!allowMerge && existingIndex === -1 && duplicateIndex !== -1) {
      const error = new Error(`A laundry record for ${safe.property} on ${safe.date} already exists.`);
      error.statusCode = 400;
      throw error;
    }
    const targetIndex = existingIndex !== -1 ? existingIndex : duplicateIndex;
    const existing = targetIndex !== -1 ? next[targetIndex] : {};
    const merged = sanitizeLaundryRecord(
      {
        ...existing,
        ...safe,
        id: safe.id || existing.id || randomUUID(),
        createdAt: existing.createdAt || new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      },
      safeSettings,
      existing
    );
    if (targetIndex !== -1) next[targetIndex] = merged;
    else next.push(merged);
  });
  return sanitizeLaundryRecords(next, safeSettings);
}

async function loadRecordsAndSettings() {
  const { payload } = await loadLaundryPayloadRow();
  const settings = payload.settings;
  try {
    const rows = await loadLaundryTableRows(settings);
    return { mode: "table", settings, rows };
  } catch (error) {
    if (!isMissingLaundryTableError(error)) throw error;
    return { mode: "legacy", settings, rows: payload.records };
  }
}

async function saveSingleRecordToTable(input, settings, { allowMerge = false } = {}) {
  const safeSettings = sanitizeLaundrySettings(settings);
  const safe = sanitizeLaundryRecord(input, safeSettings);
  const existingRows = await loadLaundryTableRows(safeSettings);
  const existing = existingRows.find((record) => record.id === safe.id || recordKey(record) === recordKey(safe));
  if (existing && !allowMerge && !safe.id && recordKey(existing) === recordKey(safe)) {
    const error = new Error(`A laundry record for ${safe.property} on ${safe.date} already exists.`);
    error.statusCode = 400;
    throw error;
  }
  validateLaundrySave(existingRows, safe, { isCreate: !existing && !allowMerge, excludeId: existing?.id || "" });
  if (existing) {
    const merged = sanitizeLaundryRecord({
      ...existing,
      ...safe,
      id: existing.id,
      createdAt: existing.createdAt,
      updatedAt: new Date().toISOString(),
    }, safeSettings, existing);
    await updateLaundryTableRow(existing.id, merged, existing);
  } else {
    const created = sanitizeLaundryRecord({
      ...safe,
      id: safe.id || randomUUID(),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    }, safeSettings);
    await createLaundryTableRow(created, created);
  }
  return loadLaundryTableRows(safeSettings);
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "laundry");

    if (req.method === "GET") {
      const current = await loadRecordsAndSettings();
      res.status(200).json({ rows: current.rows, settings: current.settings });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const items = Array.isArray(body?.rows) ? body.rows : Array.isArray(body) ? body : [body];
      if (!items.length) {
        res.status(400).json({ error: "Request body is empty." });
        return;
      }
      const current = await loadRecordsAndSettings();
      if (current.mode === "legacy") {
        const { rowId, payload } = await loadLaundryPayloadRow();
        if (!(items.length > 1 || !!body?.allowMerge)) {
          const safeSingle = sanitizeLaundryRecord(items[0], payload.settings);
          validateLaundrySave(payload.records, safeSingle, { isCreate: true });
        }
        const nextRecords = mergeRecords(payload.records, items, payload.settings, { allowMerge: items.length > 1 || !!body?.allowMerge });
        const saved = await saveLaundryPayload(rowId, { settings: payload.settings, records: nextRecords });
        res.status(200).json({ rows: saved.records, settings: saved.settings });
        return;
      }
      let rows = current.rows;
      for (const item of items) {
        rows = await saveSingleRecordToTable(item, current.settings, { allowMerge: items.length > 1 || !!body?.allowMerge });
      }
      res.status(200).json({ rows, settings: current.settings });
      return;
    }

    if (req.method === "PUT") {
      const id = String(req.query?.id || "").trim();
      if (!id) {
        res.status(400).json({ error: "Missing id query parameter." });
        return;
      }
      const body = await parseBody(req);
      const current = await loadRecordsAndSettings();
      if (current.mode === "legacy") {
        const { rowId, payload } = await loadLaundryPayloadRow();
        const existing = payload.records.find((record) => record.id === id);
        if (!existing) {
          res.status(404).json({ error: "Laundry record not found." });
          return;
        }
        const updated = sanitizeLaundryRecord(
          {
            ...existing,
            ...body,
            id,
            createdAt: existing.createdAt,
            updatedAt: new Date().toISOString(),
          },
          payload.settings,
          existing
        );
        validateLaundrySave(payload.records, updated, { isCreate: false, excludeId: id });
        const nextRecords = mergeRecords(payload.records.filter((record) => record.id !== id), [updated], payload.settings, { allowMerge: true });
        const saved = await saveLaundryPayload(rowId, { settings: payload.settings, records: nextRecords });
        res.status(200).json({ rows: saved.records, settings: saved.settings });
        return;
      }
      const existing = current.rows.find((record) => record.id === id);
      if (!existing) {
        res.status(404).json({ error: "Laundry record not found." });
        return;
      }
      const merged = sanitizeLaundryRecord(
        {
          ...existing,
          ...body,
          id,
          createdAt: existing.createdAt,
          updatedAt: new Date().toISOString(),
        },
        current.settings,
        existing
      );
      validateLaundrySave(current.rows, merged, { isCreate: false, excludeId: id });
      await updateLaundryTableRow(id, merged, existing);
      const rows = await loadLaundryTableRows(current.settings);
      res.status(200).json({ rows, settings: current.settings });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
