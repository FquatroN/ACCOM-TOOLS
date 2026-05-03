const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  DEFAULT_LAUNDRY_SETTINGS,
  LAUNDRY_SETTING_KEY,
  sanitizeLaundryPayload,
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

module.exports = async function handler(req, res) {
  try {
    let canEdit = false;
    try {
      await requireFeature(req, "settings", "laundry");
      canEdit = true;
    } catch {
      await requireFeature(req, "app", "laundry");
    }

    if (req.method === "GET") {
      const { payload } = await loadLaundryPayloadRow();
      res.status(200).json({ settings: payload.settings });
      return;
    }

    if (req.method === "PUT") {
      if (!canEdit) await requireFeature(req, "settings", "laundry");
      const body = await parseBody(req);
      const { rowId, payload } = await loadLaundryPayloadRow();
      const next = {
        settings: sanitizeLaundrySettings(body?.settings),
        records: payload.records,
      };
      const saved = await saveLaundryPayload(rowId, next);
      res.status(200).json({ settings: saved.settings });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
