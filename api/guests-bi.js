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
    const [pieRows, lineRows, monthLineRows] = await Promise.all([
      restQuery("rpc/guests_bi_nationality_pies", {
        method: "POST",
        body: {},
      }),
      restQuery("rpc/guests_bi_nationality_line", {
        method: "POST",
        body: {},
      }),
      restQuery("rpc/guests_bi_nationality_month_line", {
        method: "POST",
        body: {},
      }),
    ]);

    const mappedRows = (Array.isArray(rows) ? rows : []).map((row) => ({
      yearMonth: clean(row?.year_month),
      totalNights: Number(row?.total_nights || 0),
      exempt7Days: Number(row?.exempt_7days || 0),
      exempt13Year: Number(row?.exempt_13_year || 0),
    }));

    const mappedPieRows = (Array.isArray(pieRows) ? pieRows : []).map((row) => ({
      chartYear: clean(row?.chart_year),
      countryLabel: clean(row?.country_label) || "Unknown",
      guestCount: Number(row?.guest_count || 0),
      sortOrder: Number(row?.sort_order || 999),
    }));
    const mappedLineRows = (Array.isArray(lineRows) ? lineRows : []).map((row) => ({
      chartYear: clean(row?.chart_year),
      countryLabel: clean(row?.country_label) || "Unknown",
      guestCount: Number(row?.guest_count || 0),
      sortOrder: Number(row?.sort_order || 999),
    }));
    const mappedMonthLineRows = (Array.isArray(monthLineRows) ? monthLineRows : []).map((row) => ({
      chartMonth: clean(row?.chart_month),
      countryLabel: clean(row?.country_label) || "Unknown",
      guestCount: Number(row?.guest_count || 0),
      sortOrder: Number(row?.sort_order || 999),
    }));

    const currentCalendarYear = new Date().getFullYear();
    const pieYears = [0, 1, 2].map((offset) => String(currentCalendarYear - offset));
    const pieCharts = pieYears.map((year) => ({
      year,
      rows: mappedPieRows
        .filter((row) => row.chartYear === year)
        .sort((a, b) => (a.sortOrder - b.sortOrder) || b.guestCount - a.guestCount || a.countryLabel.localeCompare(b.countryLabel))
        .map((row) => ({
          countryLabel: row.countryLabel,
          guestCount: row.guestCount,
        })),
    }));

    const lineYears = [...new Set(mappedLineRows.map((row) => row.chartYear).filter((year) => /^\d{4}$/.test(year)))].sort((a, b) => a.localeCompare(b));
    const lineSeriesMap = new Map();
    mappedLineRows.forEach((row) => {
      const key = row.countryLabel;
      if (!lineSeriesMap.has(key)) {
        lineSeriesMap.set(key, {
          countryLabel: row.countryLabel,
          sortOrder: row.sortOrder,
          valueByYear: {},
        });
      }
      lineSeriesMap.get(key).valueByYear[row.chartYear] = row.guestCount;
    });
    const lineSeries = [...lineSeriesMap.values()]
      .sort((a, b) => {
        const totalA = Object.values(a.valueByYear).reduce((sum, value) => sum + Number(value || 0), 0);
        const totalB = Object.values(b.valueByYear).reduce((sum, value) => sum + Number(value || 0), 0);
        if (a.countryLabel === "Others" && b.countryLabel !== "Others") return 1;
        if (b.countryLabel === "Others" && a.countryLabel !== "Others") return -1;
        return totalB - totalA || a.countryLabel.localeCompare(b.countryLabel);
      })
      .map((item) => ({
        countryLabel: item.countryLabel,
        values: lineYears.map((year) => Number(item.valueByYear[year] || 0)),
      }));

    const monthLabels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const monthSeriesMap = new Map();
    mappedMonthLineRows.forEach((row) => {
      const key = row.countryLabel;
      if (!monthSeriesMap.has(key)) {
        monthSeriesMap.set(key, {
          countryLabel: row.countryLabel,
          valueByMonth: {},
        });
      }
      monthSeriesMap.get(key).valueByMonth[row.chartMonth] = row.guestCount;
    });
    const monthSeries = [...monthSeriesMap.values()]
      .sort((a, b) => {
        const totalA = Object.values(a.valueByMonth).reduce((sum, value) => sum + Number(value || 0), 0);
        const totalB = Object.values(b.valueByMonth).reduce((sum, value) => sum + Number(value || 0), 0);
        if (a.countryLabel === "Others" && b.countryLabel !== "Others") return 1;
        if (b.countryLabel === "Others" && a.countryLabel !== "Others") return -1;
        return totalB - totalA || a.countryLabel.localeCompare(b.countryLabel);
      })
      .map((item) => ({
        countryLabel: item.countryLabel,
        values: monthLabels.map((_, index) => Number(item.valueByMonth[String(index + 1)] || 0)),
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
      nationalities: {
        pieYears,
        pieCharts,
        lineYears,
        lineSeries,
        monthLabels,
        monthSeries,
      },
    });
  } catch (error) {
    sendError(res, error);
  }
};
