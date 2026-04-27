const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");

const DEFAULT_SOURCES = [
  { key: "booking", label: "Booking.com", active: true },
  { key: "hostelworld", label: "Hostelworld", active: true },
  { key: "expedia", label: "Expedia", active: true },
  { key: "airbnb", label: "Airbnb", active: true },
  { key: "vrbo", label: "VRBO", active: true },
  { key: "tripadvisor", label: "Tripadvisor", active: true },
  { key: "google", label: "Google", active: true },
];

function normalizeSources(value) {
  const list = Array.isArray(value) ? value : [];
  return DEFAULT_SOURCES.map((fallback) => {
    const match = list.find((item) => String(item?.key || "").trim() === fallback.key);
    return {
      key: fallback.key,
      label: String(match?.label || fallback.label).trim() || fallback.label,
      active: typeof match?.active === "boolean" ? match.active : fallback.active,
    };
  });
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "settings", "reviews");

    if (req.method === "GET") {
      const rows = await restQuery("app_settings?select=payload&setting_key=eq.reviews&limit=1", { method: "GET" });
      const payload = Array.isArray(rows) && rows[0]?.payload ? rows[0].payload : {};
      res.status(200).json({ settings: { sources: normalizeSources(payload.sources) } });
      return;
    }

    if (req.method === "PUT") {
      const body = await parseBody(req);
      const existing = await restQuery("app_settings?select=id,payload&setting_key=eq.reviews&limit=1", { method: "GET" });
      const existingPayload = Array.isArray(existing) && existing[0]?.payload ? existing[0].payload : {};
      const safe = { ...existingPayload, sources: normalizeSources(body?.settings?.sources) };
      if (Array.isArray(existing) && existing[0]) {
        await restQuery("app_settings?setting_key=eq.reviews", {
          method: "PATCH",
          body: { payload: safe, updated_at: new Date().toISOString() },
        });
      } else {
        await restQuery("app_settings", {
          method: "POST",
          body: [{ setting_key: "reviews", payload: safe }],
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
