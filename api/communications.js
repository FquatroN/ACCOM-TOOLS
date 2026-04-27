const {
  cleanText,
  parseBody,
  requireFeature,
  restQuery,
  sanitizeEntry,
  sendError,
} = require("./_supabase");

function disallowMethod(res) {
  res.status(405).json({ error: "Method not allowed." });
}

function withSelect(path) {
  return `${path}${path.includes("?") ? "&" : "?"}select=id,date,time,person,status,category,message,created_at,updated_at`;
}

async function getExistingById(id) {
  const rows = await restQuery(
    `communications?select=id,date,time,person,status,category,message,created_at,updated_at&id=eq.${encodeURIComponent(id)}&limit=1`,
    { method: "GET" }
  );
  return Array.isArray(rows) && rows[0] ? rows[0] : null;
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "communications");

    if (req.method === "GET") {
      const rows = await restQuery(
        "communications?select=id,date,time,person,status,category,message,created_at,updated_at&order=date.desc,time.desc",
        { method: "GET" }
      );
      res.status(200).json({ rows: Array.isArray(rows) ? rows : [] });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const items = Array.isArray(body) ? body : [body];
      if (items.length === 0) {
        res.status(400).json({ error: "Request body is empty." });
        return;
      }

      const nowIso = new Date().toISOString();
      const payload = items.map((item) => ({ ...sanitizeEntry(item), updated_at: nowIso }));
      const created = await restQuery(withSelect("communications"), {
        method: "POST",
        body: payload,
      });
      res.status(200).json({ rows: Array.isArray(created) ? created : [] });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanText(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Missing id query parameter." });
        return;
      }

      const body = await parseBody(req);
      const existing = await getExistingById(id);
      if (!existing) {
        res.status(404).json({ error: "Communication not found." });
        return;
      }

      // Keep original date/time by default; only change if explicitly provided.
      const merged = {
        date: body?.date ?? existing.date,
        time: body?.time ?? existing.time,
        person: body?.person ?? existing.person,
        status: body?.status ?? existing.status,
        category: body?.category ?? existing.category,
        message: body?.message ?? existing.message,
      };
      const payload = { ...sanitizeEntry(merged), updated_at: new Date().toISOString() };
      const updated = await restQuery(withSelect(`communications?id=eq.${encodeURIComponent(id)}`), {
        method: "PATCH",
        body: payload,
      });
      res.status(200).json({ rows: Array.isArray(updated) ? updated : [] });
      return;
    }

    if (req.method === "DELETE") {
      const id = cleanText(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Missing id query parameter." });
        return;
      }

      const deleted = await restQuery(withSelect(`communications?id=eq.${encodeURIComponent(id)}`), {
        method: "DELETE",
      });
      res.status(200).json({ rows: Array.isArray(deleted) ? deleted : [] });
      return;
    }

    disallowMethod(res);
  } catch (error) {
    sendError(res, error);
  }
};
