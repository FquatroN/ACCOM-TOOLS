const { hasFeature, loadAccessForUser, parseBody, requireFeature, restQuery, sendError, verifyUser } = require("./_supabase");
const { COUNTRIES, DEFAULT_GUESTS_SETTINGS, GUESTS_SETTING_KEY, sanitizeGuestsPayload, normalizeGuestSettings } = require("./_guests");

async function loadGuestsPayloadRow() {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${encodeURIComponent(GUESTS_SETTING_KEY)}&limit=1`, {
    method: "GET",
  });
  const row = Array.isArray(rows) && rows[0] ? rows[0] : null;
  return {
    rowId: row?.id || "",
    payload: sanitizeGuestsPayload(row?.payload || { settings: DEFAULT_GUESTS_SETTINGS, rows: [], blacklist: [], lastFileNumber: 0 }),
  };
}

async function saveGuestsPayload(rowId, payload) {
  const safe = sanitizeGuestsPayload(payload);
  if (rowId) {
    await restQuery(`app_settings?id=eq.${encodeURIComponent(rowId)}`, {
      method: "PATCH",
      body: { payload: safe, updated_at: new Date().toISOString() },
    });
    return safe;
  }
  const created = await restQuery("app_settings", {
    method: "POST",
    body: [{ setting_key: GUESTS_SETTING_KEY, payload: safe }],
  });
  return sanitizeGuestsPayload(Array.isArray(created) && created[0]?.payload ? created[0].payload : safe);
}

module.exports = async function handler(req, res) {
  try {
    if (req.method === "GET") {
      const user = await verifyUser(req);
      const access = await loadAccessForUser(user.id);
      if (!hasFeature(access, "settings", "guests") && !hasFeature(access, "app", "guests")) {
        const err = new Error("You do not have permission for this feature.");
        err.statusCode = 403;
        throw err;
      }
      const { payload } = await loadGuestsPayloadRow();
      res.status(200).json({ settings: payload.settings, countries: COUNTRIES });
      return;
    }

    if (req.method === "PUT") {
      await requireFeature(req, "settings", "guests");
      const body = await parseBody(req);
      const { rowId, payload } = await loadGuestsPayloadRow();
      const next = {
        ...payload,
        settings: normalizeGuestSettings(body?.settings),
      };
      const saved = await saveGuestsPayload(rowId, next);
      res.status(200).json({ settings: saved.settings, countries: COUNTRIES });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
