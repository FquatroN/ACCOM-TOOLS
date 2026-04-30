const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const { DEFAULT_SHOPPING_SETTINGS, sanitizeShoppingSettings } = require("./_shopping");

async function loadShoppingSettings() {
  const rows = await restQuery("app_settings?select=payload&setting_key=eq.shopping&limit=1", {
    method: "GET",
  });
  const payload = Array.isArray(rows) && rows[0]?.payload ? rows[0].payload : DEFAULT_SHOPPING_SETTINGS;
  return sanitizeShoppingSettings(payload);
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "settings", "shopping");

    if (req.method === "GET") {
      const settings = await loadShoppingSettings();
      res.status(200).json({ settings });
      return;
    }

    if (req.method === "PUT") {
      const body = await parseBody(req);
      const safe = sanitizeShoppingSettings(body?.settings);
      const existing = await restQuery("app_settings?select=id&setting_key=eq.shopping&limit=1", {
        method: "GET",
      });
      if (Array.isArray(existing) && existing[0]) {
        await restQuery("app_settings?setting_key=eq.shopping", {
          method: "PATCH",
          body: { payload: safe, updated_at: new Date().toISOString() },
        });
      } else {
        await restQuery("app_settings", {
          method: "POST",
          body: [{ setting_key: "shopping", payload: safe }],
        });
      }
      res.status(200).json({ settings: safe });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
