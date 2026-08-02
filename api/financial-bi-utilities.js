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
  const supplierNifs = [...new Set((Array.isArray(query.supplierNif) ? query.supplierNif : [query.supplierNif]).map(clean).filter(Boolean))].sort((left, right) => left.localeCompare(right));
  const from = parseYear(rawFrom, currentYear - 10);
  const to = parseYear(rawTo, currentYear);
  return { yearFrom: Math.min(from, to), yearTo: Math.max(from, to), supplierNifs };
}

function normalizeRows(rows) {
  return (Array.isArray(rows) ? rows : []).map((row) => ({
    year: Number.parseInt(row?.year, 10) || 0,
    month: Number.parseInt(row?.month, 10) || 0,
    yearMonth: clean(row?.year_month || row?.yearMonth),
    cc: clean(row?.cc).toUpperCase(),
    sumAmount: Number(row?.sum_amount || row?.sumAmount || 0),
    documentCount: Number(row?.document_count || row?.documentCount || 0),
  })).filter((row) => row.year && row.month && (row.cc === "H" || row.cc === "A"));
}

async function loadPayload(filters) {
  const key = `${filters.yearFrom}:${filters.yearTo}:${filters.supplierNifs.join("|")}`;
  const now = Date.now();
  for (const [cacheKey, entry] of cache.entries()) {
    if (!entry || entry.expiresAt <= now) cache.delete(cacheKey);
  }
  if (cache.has(key)) return cache.get(key).payload;
  if (pendingLoads.has(key)) return pendingLoads.get(key);
  const load = restQuery("rpc/get_bi_financial_utilities_payload", {
    method: "POST",
    body: {
      p_year_from: filters.yearFrom,
      p_year_to: filters.yearTo,
      p_supplier_nifs: filters.supplierNifs.length ? filters.supplierNifs : null,
    },
  }).then((payload) => {
    const first = Array.isArray(payload) ? payload[0] : payload;
    if (!first || !Array.isArray(first.rows)) {
      const error = new Error("Utilities returned an invalid aggregate payload. Run the database migration.");
      error.statusCode = 502;
      throw error;
    }
    const normalized = {
      rows: normalizeRows(first.rows),
      suppliers: Array.isArray(first.suppliers)
        ? first.suppliers.map((supplier) => ({
          nif: clean(supplier?.nif),
          name: clean(supplier?.name),
        })).filter((supplier) => supplier.nif)
        : [],
    };
    cache.set(key, { payload: normalized, expiresAt: Date.now() + CACHE_TTL_MS });
    return normalized;
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
    const payload = await loadPayload(filters);
    res.status(200).json({ filters, ...payload, rowCount: payload.rows.length });
  } catch (error) {
    console.error("[financial-bi-utilities] load failed", { message: error.message, statusCode: error.statusCode });
    sendError(res, error);
  }
};
