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
  const from = parseYear(Array.isArray(query.yearFrom) ? query.yearFrom[0] : query.yearFrom, currentYear - 5);
  const to = parseYear(Array.isArray(query.yearTo) ? query.yearTo[0] : query.yearTo, currentYear);
  return { yearFrom: Math.min(from, to), yearTo: Math.max(from, to) };
}

function number(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeRows(rows) {
  return (Array.isArray(rows) ? rows : []).map((row) => ({
    year: Number.parseInt(row?.year, 10) || 0,
    month: Number.parseInt(row?.month, 10) || 0,
    yearMonth: clean(row?.year_month || row?.yearMonth),
    accommodation: number(row?.accommodation),
    drinks: number(row?.drinks),
    bookingSales: number(row?.booking_sales ?? row?.bookingSales),
    tmt: number(row?.tmt),
    tours: number(row?.tours),
    incomeBookings: number(row?.income_bookings ?? row?.incomeBookings),
    reservation: number(row?.reservation),
    accountSales: number(row?.account_sales ?? row?.accountSales),
    incomeAccounts: number(row?.income_accounts ?? row?.incomeAccounts),
    difference: number(row?.difference),
  })).filter((row) => row.year && row.month);
}

async function loadRows(filters) {
  const key = `${filters.yearFrom}:${filters.yearTo}`;
  const now = Date.now();
  for (const [cacheKey, entry] of cache.entries()) {
    if (!entry || entry.expiresAt <= now) cache.delete(cacheKey);
  }
  if (cache.has(key)) return cache.get(key).rows;
  if (pendingLoads.has(key)) return pendingLoads.get(key);
  const load = restQuery("rpc/get_bi_financial_income_book_vs_accounts_payload", {
    method: "POST",
    body: { p_year_from: filters.yearFrom, p_year_to: filters.yearTo },
  }).then((payload) => {
    const first = Array.isArray(payload) ? payload[0] : payload;
    if (!first || !Array.isArray(first.rows)) {
      const error = new Error("Income comparison returned an invalid aggregate payload. Run the database migration.");
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
    await requireFeature(req, "app", "financial-bi");
    const filters = filtersFor(req);
    const rows = await loadRows(filters);
    res.status(200).json({ filters, rows, rowCount: rows.length });
  } catch (error) {
    console.error("[financial-bi-income-book-vs-accounts] load failed", { message: error.message, statusCode: error.statusCode });
    sendError(res, error);
  }
};
