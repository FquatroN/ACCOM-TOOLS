const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "reviews");

    if (req.method === "GET") {
      const rows = await restQuery("properties?select=id,name,active,created_at,updated_at&order=name.asc", { method: "GET" });
      res.status(200).json({ rows: Array.isArray(rows) ? rows : [] });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const name = String(body?.name || "").trim();
      if (!name) {
        res.status(400).json({ error: "Property name is required." });
        return;
      }
      const created = await restQuery("properties?select=id,name,active,created_at,updated_at", {
        method: "POST",
        body: [{ name, active: body?.active !== false }],
      });
      res.status(200).json({ row: Array.isArray(created) && created[0] ? created[0] : null });
      return;
    }

    if (req.method === "PATCH") {
      const id = String(req.query?.id || "").trim();
      if (!id) {
        res.status(400).json({ error: "Missing id query parameter." });
        return;
      }
      const body = await parseBody(req);
      const name = String(body?.name || "").trim();
      if (!name) {
        res.status(400).json({ error: "Property name is required." });
        return;
      }
      const updated = await restQuery(`properties?id=eq.${encodeURIComponent(id)}&select=id,name,active,created_at,updated_at`, {
        method: "PATCH",
        body: { name, active: body?.active !== false },
      });
      res.status(200).json({ row: Array.isArray(updated) && updated[0] ? updated[0] : null });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
