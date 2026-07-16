const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  BI_SETTINGS_KEY,
  safeBiSettings,
  sanitizeBiSettings,
} = require("./_bi-settings");

async function loadBiSettings() {
  const rows = await restQuery(`app_settings?select=payload&setting_key=eq.${encodeURIComponent(BI_SETTINGS_KEY)}&limit=1`, {
    method: "GET",
  });
  return sanitizeBiSettings(Array.isArray(rows) && rows[0] ? rows[0].payload : {});
}

async function saveBiSettings(settings) {
  const safe = sanitizeBiSettings(settings);
  const existing = await restQuery(`app_settings?select=id&setting_key=eq.${encodeURIComponent(BI_SETTINGS_KEY)}&limit=1`, {
    method: "GET",
  });
  if (Array.isArray(existing) && existing[0]?.id) {
    await restQuery(`app_settings?setting_key=eq.${encodeURIComponent(BI_SETTINGS_KEY)}`, {
      method: "PATCH",
      body: { payload: safe },
    });
  } else {
    await restQuery("app_settings", {
      method: "POST",
      body: [{ setting_key: BI_SETTINGS_KEY, payload: safe }],
    });
  }
  return safe;
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "settings", "bi-settings");
    if (req.method === "GET") {
      const settings = await loadBiSettings();
      res.status(200).json({ settings: safeBiSettings(settings) });
      return;
    }
    if (req.method === "PUT") {
      const body = await parseBody(req);
      const settings = await saveBiSettings(body?.settings);
      res.status(200).json({ settings: safeBiSettings(settings) });
      return;
    }
    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
