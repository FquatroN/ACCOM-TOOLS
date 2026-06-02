const { requireFeature, restQuery, sendError } = require("./_supabase");

function clean(value) {
  return String(value ?? "").trim();
}

function parseYearValue(value) {
  const raw = clean(value);
  if (!/^\d{4}$/.test(raw)) return null;
  const year = Number.parseInt(raw, 10);
  return Number.isFinite(year) ? year : null;
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "guests-bi");
    if (req.method !== "GET") {
      res.setHeader("Allow", "GET");
      res.status(405).json({ error: "Method not allowed." });
      return;
    }

    const yearRows = await restQuery("rpc/guests_bi_years", {
      method: "POST",
      body: {},
    });
    const availableYears = [...new Set((Array.isArray(yearRows) ? yearRows : [])
      .map((row) => clean(row?.year))
      .filter((year) => /^\d{4}$/.test(year)))]
      .sort((a, b) => b.localeCompare(a));

    const requestedYear = parseYearValue(req.query?.year);
    const fallbackYear = parseYearValue(availableYears[0]) || new Date().getUTCFullYear();
    const selectedYear = requestedYear || fallbackYear;

    const rows = await restQuery("rpc/guests_bi_tmt", {
      method: "POST",
      body: { p_year: selectedYear },
    });

    const mappedRows = (Array.isArray(rows) ? rows : []).map((row) => ({
      yearMonth: clean(row?.year_month),
      totalNights: Number(row?.total_nights || 0),
      exempt7Days: Number(row?.exempt_7days || 0),
      exempt13Year: Number(row?.exempt_13_year || 0),
    }));

    const totals = mappedRows.reduce((acc, row) => ({
      totalNights: acc.totalNights + row.totalNights,
      exempt7Days: acc.exempt7Days + row.exempt7Days,
      exempt13Year: acc.exempt13Year + row.exempt13Year,
    }), { totalNights: 0, exempt7Days: 0, exempt13Year: 0 });

    res.status(200).json({
      year: String(selectedYear),
      years: availableYears.length ? availableYears : [String(selectedYear)],
      rows: mappedRows,
      totals,
    });
  } catch (error) {
    sendError(res, error);
  }
};
