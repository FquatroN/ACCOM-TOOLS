const {
  cleanText,
  parseBody,
  requireSettingsAdmin,
  restQuery,
  sanitizeProfileInput,
  sendError,
} = require("./_supabase");

module.exports = async function handler(req, res) {
  try {
    await requireSettingsAdmin(req);

    if (req.method === "GET") {
      const rows = await restQuery("app_profiles?select=id,name,app_features,settings_features&order=name.asc", {
        method: "GET",
      });
      res.status(200).json({ profiles: Array.isArray(rows) ? rows : [] });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const safe = sanitizeProfileInput(body);
      const created = await restQuery("app_profiles?select=id,name,app_features,settings_features", {
        method: "POST",
        body: [{ name: safe.name, app_features: safe.appFeatures, settings_features: safe.settingsFeatures }],
      });
      res.status(200).json({ profile: Array.isArray(created) && created[0] ? created[0] : null });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanText(req.query?.id);
      if (!id) return res.status(400).json({ error: "Missing id query parameter." });
      const body = await parseBody(req);
      const safe = sanitizeProfileInput(body);
      const updated = await restQuery(
        `app_profiles?id=eq.${encodeURIComponent(id)}&select=id,name,app_features,settings_features`,
        {
          method: "PATCH",
          body: { name: safe.name, app_features: safe.appFeatures, settings_features: safe.settingsFeatures },
        }
      );
      res.status(200).json({ profile: Array.isArray(updated) && updated[0] ? updated[0] : null });
      return;
    }

    if (req.method === "DELETE") {
      const id = cleanText(req.query?.id);
      if (!id) return res.status(400).json({ error: "Missing id query parameter." });

      const linked = await restQuery(
        `user_profile_assignments?select=user_id&profile_id=eq.${encodeURIComponent(id)}&limit=1`,
        { method: "GET" }
      );
      if (Array.isArray(linked) && linked[0]) {
        return res.status(400).json({ error: "Cannot delete a profile that is assigned to users." });
      }

      await restQuery(`app_profiles?id=eq.${encodeURIComponent(id)}`, { method: "DELETE" });
      res.status(200).json({ ok: true });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
