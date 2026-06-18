const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  IMPORT_DATA_SETTINGS_KEY,
  safeImportDataSettings,
  sanitizeImportDataSettings,
} = require("./_import-data");

async function loadImportDataSettings() {
  const rows = await restQuery(`app_settings?select=payload&setting_key=eq.${encodeURIComponent(IMPORT_DATA_SETTINGS_KEY)}&limit=1`, {
    method: "GET",
  });
  return sanitizeImportDataSettings(Array.isArray(rows) && rows[0] ? rows[0].payload : {});
}

async function saveImportDataSettings(settings) {
  const safe = sanitizeImportDataSettings(settings);
  const existing = await restQuery(`app_settings?select=id&setting_key=eq.${encodeURIComponent(IMPORT_DATA_SETTINGS_KEY)}&limit=1`, {
    method: "GET",
  });
  if (Array.isArray(existing) && existing[0]?.id) {
    await restQuery(`app_settings?setting_key=eq.${encodeURIComponent(IMPORT_DATA_SETTINGS_KEY)}`, {
      method: "PATCH",
      body: { payload: safe },
    });
  } else {
    await restQuery("app_settings", {
      method: "POST",
      body: [{ setting_key: IMPORT_DATA_SETTINGS_KEY, payload: safe }],
    });
  }
  return safe;
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "settings", "import-data");
    if (req.method === "GET") {
      const settings = await loadImportDataSettings();
      res.status(200).json({ settings: safeImportDataSettings(settings) });
      return;
    }
    if (req.method === "PUT") {
      const body = await parseBody(req);
      const settings = await saveImportDataSettings(body?.settings);
      res.status(200).json({ settings: safeImportDataSettings(settings) });
      return;
    }
    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
