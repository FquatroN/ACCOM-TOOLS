const { requireFeature, restQuery, sendError } = require("./_supabase");

const CACHE_TTL_MS = 60 * 1000;
const cache = new Map();
const pendingLoads = new Map();

function clean(value) {
  return String(value ?? "").trim();
}

function parseYear(value, fallback) {
  const year = Number.parseInt(value, 10);
  return Number.isFinite(year) && year >= 2000 && year <= 2100 ? year : fallback;
}

function filtersFor(req) {
  const currentYear = new Date().getFullYear();
  const query = req.query || {};
  const rawFrom = Array.isArray(query.yearFrom) ? query.yearFrom[0] : query.yearFrom;
  const rawTo = Array.isArray(query.yearTo) ? query.yearTo[0] : query.yearTo;
  const from = parseYear(rawFrom, currentYear - 5);
  const to = parseYear(rawTo, currentYear);
  return { yearFrom: Math.min(from, to), yearTo: Math.max(from, to) };
}

function normalizeRows(rows) {
  return (Array.isArray(rows) ? rows : []).map((row) => {
    const sales = Number(row?.sales || 0);
    const buy = Number(row?.buy || 0);
    return {
      analysis: clean(row?.analysis || "nespresso").toLowerCase(),
      year: Number.parseInt(row?.year, 10) || 0,
      month: Number.parseInt(row?.month, 10) || 0,
      yearMonth: clean(row?.year_month || row?.yearMonth),
      sales: Number.isFinite(sales) ? sales : 0,
      buy: Number.isFinite(buy) ? buy : 0,
      revenue: Number(row?.revenue ?? (sales - buy)) || 0,
    };
  }).filter((row) => row.year && row.month);
}

async function loadRows(filters) {
  const key = `${filters.yearFrom}:${filters.yearTo}`;
  const now = Date.now();
  for (const [cacheKey, entry] of cache.entries()) {
    if (!entry || entry.expiresAt <= now) cache.delete(cacheKey);
  }
  if (cache.has(key)) return cache.get(key).rows;
  if (pendingLoads.has(key)) return pendingLoads.get(key);
  const load = restQuery("rpc/get_bi_financial_items_revenue_payload", {
    method: "POST",
    body: { p_year_from: filters.yearFrom, p_year_to: filters.yearTo },
  }).then((payload) => {
    const first = Array.isArray(payload) ? payload[0] : payload;
    if (!first || !Array.isArray(first.rows)) {
      const error = new Error("Items Revenue returned an invalid aggregate payload. Run the database migration.");
      error.statusCode = 502;
      throw error;
    }
    const rows = normalizeRows(first.rows);
    cache.set(key, { rows, expiresAt: Date.now() + CACHE_TTL_MS });
    return rows;
  }).finally(() => pendingLoads.delete(key));
  pendingLoads.set(key, load);
  return load;
}

module.exports = async function handler(req, res) {
  try {
    if (req.method !== "GET") {
      res.setHeader("Allow", "GET");
      const error = new Error("Method not allowed.");
      error.statusCode = 405;
      throw error;
    }
    await requireFeature(req, "app", "sales-bi");
    const filters = filtersFor(req);
    const rows = await loadRows(filters);
    res.status(200).json({ filters, rows, rowCount: rows.length });
  } catch (error) {
    console.error("[financial-bi-items-revenue] load failed", { message: error.message, statusCode: error.statusCode });
    sendError(res, error);
  }
};
