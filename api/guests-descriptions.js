const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  DEFAULT_GUESTS_SETTINGS,
  GUESTS_SETTING_KEY,
  sanitizeGuestDescriptionRecord,
  sanitizeGuestsPayload,
} = require("./_guests");

function cleanId(value) {
  return String(value || "").trim();
}

async function loadGuestsPayloadRow() {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${encodeURIComponent(GUESTS_SETTING_KEY)}&limit=1`, {
    method: "GET",
  });
  const row = Array.isArray(rows) && rows[0] ? rows[0] : null;
  return {
    rowId: row?.id || "",
    payload: sanitizeGuestsPayload(row?.payload || { settings: DEFAULT_GUESTS_SETTINGS, rows: [], blacklist: [], descriptions: [], lastFileNumber: 0 }),
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
    await requireFeature(req, "app", "guests");

    if (req.method === "GET") {
      const { payload } = await loadGuestsPayloadRow();
      res.status(200).json({ rows: payload.descriptions || [] });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanId(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Description id is required." });
        return;
      }
      const body = await parseBody(req);
      const { rowId, payload } = await loadGuestsPayloadRow();
      const existing = (payload.descriptions || []).find((item) => cleanId(item.id) === id);
      if (!existing) {
        res.status(404).json({ error: "Guest description row not found." });
        return;
      }
      const updated = sanitizeGuestDescriptionRecord({ ...existing, guestDescription: body?.guestDescription }, existing);
      const nextDescriptions = (payload.descriptions || []).map((item) => (cleanId(item.id) === id ? updated : item));
      const saved = await saveGuestsPayload(rowId, { ...payload, descriptions: nextDescriptions });
      res.status(200).json({ rows: saved.descriptions || [] });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
