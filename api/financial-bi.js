const { requireFeature, restQuery, sendError } = require("./_supabase");

const FINANCIAL_BI_CACHE_TTL_MS = 60 * 1000;
const financialBiCache = new Map();
const financialBiLoads = new Map();

function clean(value) {
  return String(value ?? "").trim();
}

function parseNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const raw = clean(value);
  if (!raw) return 0;
  const normalized = raw
    .replace(/[^\d,.-]/g, "")
    .replace(/\s+/g, "")
    .replace(/\.(?=\d{3}(?:\D|$))/g, "")
    .replace(",", ".");
  const num = Number(normalized);
  return Number.isFinite(num) ? num : 0;
}

function firstValue(row, keys) {
  for (const key of keys) {
    if (row && Object.prototype.hasOwnProperty.call(row, key) && row[key] !== null && row[key] !== undefined) return row[key];
  }
  return "";
}

function normalizeRows(rows) {
  return (Array.isArray(rows) ? rows : []).map((row) => ({
    type: clean(firstValue(row, ["type", "TYPE", "Type"])).toUpperCase() === "EXPENSE" ? "EXPENSE" : "INCOME",
    year: Number.parseInt(firstValue(row, ["year", "YEAR", "Year"]), 10) || 0,
    month: Number.parseInt(firstValue(row, ["month", "MONTH", "Month"]), 10) || 0,
    yearMonth: clean(firstValue(row, ["year_month", "yearMonth", "YEAR_MONTH", "YearMonth"])),
    cc: clean(firstValue(row, ["cc", "CC"])).toUpperCase(),
    category: clean(firstValue(row, ["category", "CATEGORY", "Category", "sale_category", "saleCategory"])),
    totalAmount: parseNumber(firstValue(row, ["total_amount", "totalAmount", "TOTAL_AMOUNT", "total_charge", "total_sales", "total"])),
  })).filter((row) => row.year && row.month);
}

function categorySort(a, b, type) {
  const left = clean(a);
  const right = clean(b);
  if (type === "INCOME") {
    if (left === "Accomodation" && right !== "Accomodation") return -1;
    if (right === "Accomodation" && left !== "Accomodation") return 1;
  }
  if (!left && right) return 1;
  if (!right && left) return -1;
  return left.localeCompare(right);
}

function addAmount(bucket, category, amount) {
  const key = clean(category);
  bucket[key] = parseNumber(bucket[key]) + parseNumber(amount);
}

function buildEmptyTotals() {
  return {
    income: {},
    expense: {},
    incomeTotal: 0,
    expenseTotal: 0,
    grandTotal: 0,
  };
}

function buildPivot(rows) {
  const incomeCategories = [...new Set(rows.filter((row) => row.type === "INCOME").map((row) => row.category))]
    .sort((a, b) => categorySort(a, b, "INCOME"));
  const expenseCategories = [...new Set(rows.filter((row) => row.type === "EXPENSE").map((row) => row.category))]
    .sort((a, b) => categorySort(a, b, "EXPENSE"));
  const years = new Map();
  const totals = buildEmptyTotals();

  rows.forEach((row) => {
    const yearKey = String(row.year);
    const monthKey = String(row.month).padStart(2, "0");
    if (!years.has(yearKey)) {
      years.set(yearKey, {
        year: row.year,
        months: new Map(),
        ...buildEmptyTotals(),
      });
    }
    const yearBucket = years.get(yearKey);
    if (!yearBucket.months.has(monthKey)) {
      yearBucket.months.set(monthKey, {
        year: row.year,
        month: row.month,
        yearMonth: row.yearMonth || `${row.year}-${monthKey}`,
        ...buildEmptyTotals(),
      });
    }
    const monthBucket = yearBucket.months.get(monthKey);
    const targetName = row.type === "EXPENSE" ? "expense" : "income";
    const totalName = row.type === "EXPENSE" ? "expenseTotal" : "incomeTotal";

    addAmount(monthBucket[targetName], row.category, row.totalAmount);
    addAmount(yearBucket[targetName], row.category, row.totalAmount);
    addAmount(totals[targetName], row.category, row.totalAmount);
    monthBucket[totalName] += row.totalAmount;
    yearBucket[totalName] += row.totalAmount;
    totals[totalName] += row.totalAmount;
  });

  const yearRows = Array.from(years.values())
    .sort((a, b) => Number(a.year) - Number(b.year))
    .map((year) => {
      const months = Array.from(year.months.values())
        .sort((a, b) => Number(a.month) - Number(b.month))
        .map((month) => ({
          ...month,
          grandTotal: month.incomeTotal - month.expenseTotal,
        }));
      return {
        ...year,
        months,
        grandTotal: year.incomeTotal - year.expenseTotal,
      };
    });

  return {
    incomeCategories,
    expenseCategories,
    years: yearRows,
    totals: {
      ...totals,
      grandTotal: totals.incomeTotal - totals.expenseTotal,
    },
  };
}

function parseYear(value, fallback) {
  const year = Number.parseInt(value, 10);
  return Number.isFinite(year) && year >= 2000 && year <= 2100 ? year : fallback;
}

function financialBiFilters(req) {
  const currentYear = new Date().getFullYear();
  const query = req.query || {};
  const rawFrom = Array.isArray(query.yearFrom) ? query.yearFrom[0] : query.yearFrom;
  const rawTo = Array.isArray(query.yearTo) ? query.yearTo[0] : query.yearTo;
  const from = parseYear(rawFrom, currentYear - 4);
  const to = parseYear(rawTo, currentYear);
  const cc = clean(Array.isArray(query.cc) ? query.cc[0] : query.cc).toUpperCase();
  return {
    yearFrom: Math.min(from, to),
    yearTo: Math.max(from, to),
    cc: cc === "H" || cc === "A" ? cc : "",
  };
}

async function loadFinancialBiRows(filters) {
  try {
    const payload = await restQuery("rpc/get_bi_financial_analysis_sales_payload", {
      method: "POST",
      body: {
        p_year_from: filters.yearFrom,
        p_year_to: filters.yearTo,
        p_cc: filters.cc || null,
      },
    });
    const firstRow = Array.isArray(payload) ? payload[0] : payload;
    if (!firstRow || !Array.isArray(firstRow.rows)) {
      const invalidPayload = new Error("Financial BI returned an invalid aggregate payload.");
      invalidPayload.statusCode = 502;
      throw invalidPayload;
    }
    return firstRow.rows;
  } catch (error) {
    const message = clean(error.message).toLowerCase();
    const isMissingRpc = error.statusCode === 404 || message.includes("could not find the function") || message.includes("schema cache");
    if (!isMissingRpc) throw error;

    const legacyRows = await restQuery("rpc/get_bi_financial_analysis_sales", {
      method: "POST",
      body: {
        p_year_from: filters.yearFrom,
        p_year_to: filters.yearTo,
        p_cc: filters.cc || null,
      },
    });
    const rows = Array.isArray(legacyRows) ? legacyRows : [];
    if (rows.length >= 1000) {
      const truncated = new Error("Financial BI database migration is required to load the complete result set.");
      truncated.statusCode = 503;
      throw truncated;
    }
    return rows;
  }
}

function financialBiCacheKey(filters) {
  return `${filters.yearFrom}:${filters.yearTo}:${filters.cc || "ALL"}`;
}

async function loadCachedFinancialBiRows(filters) {
  const key = financialBiCacheKey(filters);
  const now = Date.now();
  for (const [cacheKey, cached] of financialBiCache.entries()) {
    if (!cached || cached.expiresAt <= now) financialBiCache.delete(cacheKey);
  }
  const cached = financialBiCache.get(key);
  if (cached) {
    console.log("[financial-bi] aggregate cache hit", { key, rowCount: cached.rows.length });
    return cached.rows;
  }
  if (financialBiLoads.has(key)) {
    console.log("[financial-bi] joining in-flight aggregate", { key });
    return financialBiLoads.get(key);
  }

  const load = loadFinancialBiRows(filters)
    .then((rows) => {
      financialBiCache.set(key, { rows, expiresAt: Date.now() + FINANCIAL_BI_CACHE_TTL_MS });
      return rows;
    })
    .finally(() => financialBiLoads.delete(key));
  financialBiLoads.set(key, load);
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

    const filters = financialBiFilters(req);
    console.log("[financial-bi] loading aggregates", filters);
    const currentYear = new Date().getFullYear();
    const comparisonFilters = {
      ...filters,
      yearFrom: Math.max(2000, Math.min(filters.yearFrom - 1, currentYear - 10)),
      yearTo: Math.max(filters.yearTo, currentYear),
    };
    const comparisonRows = normalizeRows(await loadCachedFinancialBiRows(comparisonFilters))
      .sort((a, b) => Number(a.year) - Number(b.year) || Number(a.month) - Number(b.month) || clean(a.type).localeCompare(clean(b.type)) || clean(a.category).localeCompare(clean(b.category)));
    const rows = comparisonRows.filter((row) => Number(row.year) >= filters.yearFrom && Number(row.year) <= filters.yearTo);
    const typeCounts = rows.reduce((counts, row) => {
      counts[row.type] = (counts[row.type] || 0) + 1;
      return counts;
    }, {});
    console.log("[financial-bi] aggregates loaded", { rowCount: rows.length, typeCounts });

    res.status(200).json({
      rowCount: rows.length,
      filters,
      rows,
      comparisonRows,
      pivot: buildPivot(rows),
    });
  } catch (error) {
    console.error("[financial-bi] load failed", { message: error.message, statusCode: error.statusCode });
    sendError(res, error);
  }
};
