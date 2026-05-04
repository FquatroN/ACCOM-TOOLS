const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  DEFAULT_HOURS_SETTINGS,
  DEFAULT_HOURS_RECORDS,
  HOURS_REGISTER_SETTING_KEY,
  sanitizeHoursPayload,
  sanitizeHoursRecord,
  sanitizeHoursRecords,
  sanitizeHoursSettings,
} = require("./_hours-register");

async function loadHoursPayloadRow() {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${encodeURIComponent(HOURS_REGISTER_SETTING_KEY)}&limit=1`, {
    method: "GET",
  });
  const row = Array.isArray(rows) && rows[0] ? rows[0] : null;
  return {
    rowId: row?.id || "",
    payload: sanitizeHoursPayload(row?.payload || { settings: DEFAULT_HOURS_SETTINGS, records: DEFAULT_HOURS_RECORDS }),
  };
}

async function saveHoursPayload(rowId, payload) {
  const safe = sanitizeHoursPayload(payload);
  if (rowId) {
    await restQuery(`app_settings?id=eq.${encodeURIComponent(rowId)}`, {
      method: "PATCH",
      body: { payload: safe, updated_at: new Date().toISOString() },
    });
    return safe;
  }
  const created = await restQuery("app_settings", {
    method: "POST",
    body: [{ setting_key: HOURS_REGISTER_SETTING_KEY, payload: safe }],
  });
  return sanitizeHoursPayload(Array.isArray(created) && created[0]?.payload ? created[0].payload : safe);
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "hours");

    if (req.method === "GET") {
      const { payload } = await loadHoursPayloadRow();
      res.status(200).json({ rows: payload.records, settings: payload.settings });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const { rowId, payload } = await loadHoursPayloadRow();
      const record = sanitizeHoursRecord(body, payload.settings);
      const duplicate = payload.records.find((item) => item.id === record.id);
      if (duplicate) {
        const error = new Error("A record with this id already exists.");
        error.statusCode = 400;
        throw error;
      }
      const next = {
        settings: payload.settings,
        records: sanitizeHoursRecords([
          ...payload.records,
          {
            ...record,
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString(),
          },
        ], payload.settings),
      };
      const saved = await saveHoursPayload(rowId, next);
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
      const { rowId, payload } = await loadHoursPayloadRow();
      const existing = payload.records.find((item) => cleanId(item.id) === id);
      if (!existing) {
        res.status(404).json({ error: "Record not found." });
        return;
      }
      const updated = sanitizeHoursRecord({ ...existing, ...body, id: existing.id }, payload.settings, existing);
      const next = {
        settings: payload.settings,
        records: sanitizeHoursRecords(payload.records.map((item) => (
          cleanId(item.id) === id
            ? { ...updated, createdAt: existing.createdAt || new Date().toISOString(), updatedAt: new Date().toISOString() }
            : item
        )), payload.settings),
      };
      const saved = await saveHoursPayload(rowId, next);
      res.status(200).json({ rows: saved.records, settings: saved.settings });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};

function cleanId(value) {
  return String(value || "").trim();
}
