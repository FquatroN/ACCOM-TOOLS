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

function recordKey(record) {
  return `${record.property}::${record.date}`;
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

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "laundry");

    if (req.method === "GET") {
      const { payload } = await loadLaundryPayloadRow();
      res.status(200).json({ rows: payload.records, settings: payload.settings });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const items = Array.isArray(body?.rows) ? body.rows : Array.isArray(body) ? body : [body];
      if (!items.length) {
        res.status(400).json({ error: "Request body is empty." });
        return;
      }
      const { rowId, payload } = await loadLaundryPayloadRow();
      const nextRecords = mergeRecords(payload.records, items, payload.settings, { allowMerge: items.length > 1 || !!body?.allowMerge });
      const saved = await saveLaundryPayload(rowId, { settings: payload.settings, records: nextRecords });
      res.status(200).json({ rows: saved.records, settings: saved.settings });
      return;
    }

    if (req.method === "PUT") {
      const id = String(req.query?.id || "").trim();
      if (!id) {
        res.status(400).json({ error: "Missing id query parameter." });
        return;
      }
      const body = await parseBody(req);
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
      const nextRecords = mergeRecords(payload.records.filter((record) => record.id !== id), [updated], payload.settings, { allowMerge: true });
      const saved = await saveLaundryPayload(rowId, { settings: payload.settings, records: nextRecords });
      res.status(200).json({ rows: saved.records, settings: saved.settings });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
