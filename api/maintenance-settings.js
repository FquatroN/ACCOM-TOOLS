const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const { DEFAULT_MAINTENANCE_SETTINGS, MAINTENANCE_SETTING_KEY, sanitizeMaintenanceSettings } = require("./_maintenance");

async function loadMaintenanceSettings() {
  const rows = await restQuery(`app_settings?select=payload&setting_key=eq.${encodeURIComponent(MAINTENANCE_SETTING_KEY)}&limit=1`, {
    method: "GET",
  });
  const payload = Array.isArray(rows) && rows[0]?.payload ? rows[0].payload : DEFAULT_MAINTENANCE_SETTINGS;
  return sanitizeMaintenanceSettings(payload);
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "settings", "maintenance");

    if (req.method === "GET") {
      const settings = await loadMaintenanceSettings();
      res.status(200).json({ settings });
      return;
    }

    if (req.method === "PUT") {
      const body = await parseBody(req);
      const settings = sanitizeMaintenanceSettings(body?.settings);
      const existing = await restQuery(`app_settings?select=id&setting_key=eq.${encodeURIComponent(MAINTENANCE_SETTING_KEY)}&limit=1`, {
        method: "GET",
      });
      if (Array.isArray(existing) && existing[0]) {
        await restQuery(`app_settings?setting_key=eq.${encodeURIComponent(MAINTENANCE_SETTING_KEY)}`, {
          method: "PATCH",
          body: { payload: settings, updated_at: new Date().toISOString() },
        });
      } else {
        await restQuery("app_settings", {
          method: "POST",
          body: [{ setting_key: MAINTENANCE_SETTING_KEY, payload: settings }],
        });
      }
      res.status(200).json({ settings });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
