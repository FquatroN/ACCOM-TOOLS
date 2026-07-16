const { requireFeature, restQuery, sendError } = require("./_supabase");

function clean(value) {
  return String(value ?? "").trim();
}

function parseCcValue(value) {
  const raw = clean(value).toUpperCase();
  return raw === "H" || raw === "A" ? raw : "";
}

function parseNumber(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num : 0;
}

function normalizeRows(rows) {
  return (Array.isArray(rows) ? rows : []).map((row) => ({
    type: clean(row?.type || row?.TYPE).toUpperCase() === "EXPENSE" ? "EXPENSE" : "INCOME",
    year: Number.parseInt(row?.year, 10) || 0,
    month: Number.parseInt(row?.month, 10) || 0,
    yearMonth: clean(row?.year_month || row?.yearMonth),
    cc: clean(row?.cc).toUpperCase(),
    category: clean(row?.category || row?.sale_category),
    totalAmount: parseNumber(row?.total_amount || row?.totalAmount),
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

module.exports = async function handler(req, res) {
  try {
    if (req.method !== "GET") {
      res.setHeader("Allow", "GET");
      const error = new Error("Method not allowed.");
      error.statusCode = 405;
      throw error;
    }

    await requireFeature(req, "app", "financial-bi");

    const cc = parseCcValue(req.query?.cc);
    const params = new URLSearchParams();
    params.set("select", "type,year,month,year_month,cc,category,total_amount");
    if (cc) params.set("cc", `eq.${cc}`);
    params.append("order", "year.asc");
    params.append("order", "month.asc");
    params.append("order", "type.asc");
    params.append("order", "category.asc");

    const rows = normalizeRows(await restQuery(`bi_financial_analysis_sales?${params.toString()}`));

    res.status(200).json({
      cc,
      rows,
      pivot: buildPivot(rows),
    });
  } catch (error) {
    sendError(res, error);
  }
};
