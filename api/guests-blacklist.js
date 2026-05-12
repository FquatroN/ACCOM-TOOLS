const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const { COUNTRIES, DEFAULT_GUESTS_SETTINGS, GUESTS_SETTING_KEY, sanitizeBlacklistRecord, sanitizeGuestsPayload } = require("./_guests");

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
    await requireFeature(req, "app", "guests");

    if (req.method === "GET") {
      const { payload } = await loadGuestsPayloadRow();
      res.status(200).json({ rows: payload.blacklist, countries: COUNTRIES });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const { rowId, payload } = await loadGuestsPayloadRow();
      const record = sanitizeBlacklistRecord(body);
      const nextRows = sanitizeGuestsPayload({
        ...payload,
        blacklist: [...payload.blacklist, { ...record, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() }],
      }).blacklist;
      const saved = await saveGuestsPayload(rowId, { ...payload, blacklist: nextRows });
      res.status(200).json({ rows: saved.blacklist, countries: COUNTRIES });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanId(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Blacklist id is required." });
        return;
      }
      const body = await parseBody(req);
      const { rowId, payload } = await loadGuestsPayloadRow();
      const existing = payload.blacklist.find((item) => cleanId(item.id) === id);
      if (!existing) {
        res.status(404).json({ error: "Blacklist record not found." });
        return;
      }
      const record = sanitizeBlacklistRecord({ ...existing, ...body, id }, existing);
      const nextRows = payload.blacklist.map((item) => (cleanId(item.id) === id ? { ...record, createdAt: existing.createdAt || item.createdAt || new Date().toISOString(), updatedAt: new Date().toISOString() } : item));
      const saved = await saveGuestsPayload(rowId, { ...payload, blacklist: nextRows });
      res.status(200).json({ rows: saved.blacklist, countries: COUNTRIES });
      return;
    }

    if (req.method === "DELETE") {
      const id = cleanId(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Blacklist id is required." });
        return;
      }
      const { rowId, payload } = await loadGuestsPayloadRow();
      const nextRows = payload.blacklist.filter((item) => cleanId(item.id) !== id);
      const saved = await saveGuestsPayload(rowId, { ...payload, blacklist: nextRows });
      res.status(200).json({ rows: saved.blacklist, countries: COUNTRIES });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
