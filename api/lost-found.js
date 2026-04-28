const {
  cleanText,
  parseBody,
  requireFeature,
  restQuery,
  sendError,
} = require("./_supabase");

const STORED_OPTIONS = ["Receção", "Arrecadação 21"];

function normalizeStatus(value) {
  const raw = cleanText(value).toLowerCase();
  return raw === "closed" ? "Closed" : "Open";
}

function normalizeStored(value) {
  const raw = cleanText(value).toLowerCase();
  const match = STORED_OPTIONS.find((item) => item.toLowerCase() === raw);
  return match || STORED_OPTIONS[0];
}

function sanitizeLostFoundInput(input = {}, existing = null) {
  const objectDescription = cleanText(input.object_description ?? input.objectDescription ?? existing?.object_description);
  if (!objectDescription) {
    const error = new Error("Object description is required.");
    error.statusCode = 400;
    throw error;
  }

  return {
    who_found: cleanText(input.who_found ?? input.whoFound ?? existing?.who_found),
    who_recorded: cleanText(input.who_recorded ?? input.whoRecorded ?? existing?.who_recorded),
    location_found: cleanText(input.location_found ?? input.locationFound ?? existing?.location_found),
    object_description: objectDescription,
    notes: cleanText(input.notes ?? existing?.notes),
    stored_location: normalizeStored(input.stored_location ?? input.storedLocation ?? existing?.stored_location),
    status: normalizeStatus(input.status ?? existing?.status),
  };
}

function withSelect(path) {
  return `${path}${path.includes("?") ? "&" : "?"}select=id,item_number,created_at,updated_at,closed_at,who_found,who_recorded,location_found,object_description,notes,stored_location,status`;
}

async function getExistingById(id) {
  const rows = await restQuery(
    `lost_found?select=id,item_number,created_at,updated_at,closed_at,who_found,who_recorded,location_found,object_description,notes,stored_location,status&id=eq.${encodeURIComponent(id)}&limit=1`,
    { method: "GET" }
  );
  return Array.isArray(rows) && rows[0] ? rows[0] : null;
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "lost-found");

    if (req.method === "GET") {
      const rows = await restQuery(
        "lost_found?select=id,item_number,created_at,updated_at,closed_at,who_found,who_recorded,location_found,object_description,notes,stored_location,status&order=created_at.desc,item_number.desc",
        { method: "GET" }
      );
      res.status(200).json({ rows: Array.isArray(rows) ? rows : [] });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const items = Array.isArray(body) ? body : [body];
      if (!items.length) {
        res.status(400).json({ error: "Request body is empty." });
        return;
      }
      const payload = items.map((item) => sanitizeLostFoundInput(item));
      const created = await restQuery(withSelect("lost_found"), {
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
        res.status(404).json({ error: "Lost & Found record not found." });
        return;
      }

      const merged = sanitizeLostFoundInput(body, existing);
      let closedAt = existing.closed_at || null;
      if (existing.status !== "Closed" && merged.status === "Closed") closedAt = new Date().toISOString();
      if (existing.status === "Closed" && merged.status !== "Closed") closedAt = null;

      const updated = await restQuery(withSelect(`lost_found?id=eq.${encodeURIComponent(id)}`), {
        method: "PATCH",
        body: {
          ...merged,
          closed_at: closedAt,
        },
      });
      res.status(200).json({ rows: Array.isArray(updated) ? updated : [] });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
