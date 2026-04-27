const {
  cleanText,
  parseBody,
  requireFeature,
  restQuery,
  sanitizeReviewInput,
  sendError,
} = require("./_supabase");

function withSelect(path) {
  return `${path}${path.includes("?") ? "&" : "?"}select=id,property_id,import_run_id,source,source_review_id,source_reservation_id,review_date,reviewer_name,reviewer_country,language,rating_raw,rating_scale,rating_normalized_100,title,positive_review_text,negative_review_text,body,subscores,host_reply_text,host_reply_date,raw_text,parse_confidence,dedupe_fingerprint,raw_payload,created_at,updated_at`;
}

function buildFilters(query) {
  const filters = [];
  const propertyId = cleanText(query?.propertyId);
  const source = cleanText(query?.source).toLowerCase();
  const search = cleanText(query?.search);
  const dateFrom = cleanText(query?.dateFrom);
  const dateTo = cleanText(query?.dateTo);
  if (propertyId) filters.push(`property_id=eq.${encodeURIComponent(propertyId)}`);
  if (source) filters.push(`source=eq.${encodeURIComponent(source)}`);
  if (dateFrom) filters.push(`review_date=gte.${encodeURIComponent(dateFrom)}`);
  if (dateTo) filters.push(`review_date=lte.${encodeURIComponent(dateTo)}`);
  if (search) filters.push(`or=${encodeURIComponent(`(title.ilike.*${search}*,body.ilike.*${search}*,reviewer_name.ilike.*${search}*)`)}`);
  return filters;
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "reviews");

    if (req.method === "GET") {
      const filters = buildFilters(req.query);
      const limit = Math.min(Math.max(Number.parseInt(cleanText(req.query?.limit), 10) || 1500, 1), 1500);
      const offset = Math.max(Number.parseInt(cleanText(req.query?.offset), 10) || 0, 0);
      const path = `reviews?select=*,properties(name)&order=review_date.desc,created_at.desc&limit=${limit}&offset=${offset}${filters.length ? `&${filters.join("&")}` : ""}`;
      const rows = await restQuery(path, { method: "GET" });
      res.status(200).json({ rows: Array.isArray(rows) ? rows : [], limit, offset });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const items = Array.isArray(body) ? body : [body];
      const payload = items.map((item) => sanitizeReviewInput(item));
      const created = await restQuery(withSelect("reviews"), {
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
      const payload = sanitizeReviewInput(body);
      const updated = await restQuery(withSelect(`reviews?id=eq.${encodeURIComponent(id)}`), {
        method: "PATCH",
        body: payload,
      });
      res.status(200).json({ rows: Array.isArray(updated) ? updated : [] });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
