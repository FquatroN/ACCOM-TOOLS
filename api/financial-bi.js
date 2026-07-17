const { requireFeature, restQuery, sendError } = require("./_supabase");

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

async function loadAllFinancialBiRows() {
  const pageSize = 1000;
  const allRows = [];
  for (let offset = 0; offset < 200000; offset += pageSize) {
    const params = new URLSearchParams();
    params.set("select", "*");
    params.set("limit", String(pageSize));
    params.set("offset", String(offset));
    const page = await restQuery(`bi_financial_analysis_sales?${params.toString()}`);
    const rows = Array.isArray(page) ? page : [];
    allRows.push(...rows);
    if (rows.length < pageSize) break;
  }
  return allRows;
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

    const rows = normalizeRows(await loadAllFinancialBiRows())
      .sort((a, b) => Number(a.year) - Number(b.year) || Number(a.month) - Number(b.month) || clean(a.type).localeCompare(clean(b.type)) || clean(a.category).localeCompare(clean(b.category)));

    res.status(200).json({
      rowCount: rows.length,
      rows,
      pivot: buildPivot(rows),
    });
  } catch (error) {
    sendError(res, error);
  }
};
