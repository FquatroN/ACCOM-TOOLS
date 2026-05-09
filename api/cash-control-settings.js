const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  CASH_CONTROL_SETTING_KEY,
  sanitizeCashControlPayload,
  sanitizeCashControlSettings,
} = require("./_cash-control");

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

module.exports = async function handler(req, res) {
  try {
    let canEdit = false;
    try {
      await requireFeature(req, "settings", "cash");
      canEdit = true;
    } catch {
      await requireFeature(req, "app", "cash");
    }

    if (req.method === "GET") {
      const { payload } = await loadCashPayloadRow();
      res.status(200).json({ settings: payload.settings });
      return;
    }

    if (req.method === "PUT") {
      if (!canEdit) await requireFeature(req, "settings", "cash");
      const body = await parseBody(req);
      const { rowId, payload } = await loadCashPayloadRow();
      const next = {
        settings: sanitizeCashControlSettings(body?.settings),
        records: payload.records,
      };
      const saved = await saveCashPayload(rowId, next);
      res.status(200).json({ settings: saved.settings });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
