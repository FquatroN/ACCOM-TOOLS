const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  DEFAULT_HOURS_SETTINGS,
  DEFAULT_HOURS_RECORDS,
  HOURS_REGISTER_SETTING_KEY,
  sanitizeHoursPayload,
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
    let canEdit = false;
    try {
      await requireFeature(req, "settings", "hours");
      canEdit = true;
    } catch {
      await requireFeature(req, "app", "hours");
    }

    if (req.method === "GET") {
      const { payload } = await loadHoursPayloadRow();
      res.status(200).json({ settings: payload.settings });
      return;
    }

    if (req.method === "PUT") {
      if (!canEdit) await requireFeature(req, "settings", "hours");
      const body = await parseBody(req);
      const { rowId, payload } = await loadHoursPayloadRow();
      const next = {
        settings: sanitizeHoursSettings(body?.settings),
        records: payload.records,
      };
      const saved = await saveHoursPayload(rowId, next);
      res.status(200).json({ settings: saved.settings });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
