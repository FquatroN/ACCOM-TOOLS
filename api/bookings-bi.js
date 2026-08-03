const { requireFeature, restQuery, sendError } = require("./_supabase");

function clean(value) {
  return String(value ?? "").trim();
}

function parseYearValue(value) {
  const raw = clean(value);
  if (!/^\d{4}$/.test(raw)) return new Date().getUTCFullYear();
  const year = Number.parseInt(raw, 10);
  return Number.isFinite(year) ? year : new Date().getUTCFullYear();
}

function parseHaValue(value) {
  const raw = clean(value).toUpperCase();
  return raw === "H" || raw === "A" ? raw : "";
}

function defaultStatuses() {
  return ["Checked Out", "Confirmed", "Arriving", "Late", "Leaving", "Checked-in", "Checked In"];
}

function parseStatuses(value) {
  const source = Array.isArray(value) ? value : clean(value).split(",");
  const statuses = source.map((item) => clean(item)).filter(Boolean);
  return statuses.length ? [...new Set(statuses)] : defaultStatuses();
}

function buildPieYears(selectedYear) {
  return Array.from({ length: 4 }, (_, index) => selectedYear - index);
}

function normalizeRows(rows) {
  return (Array.isArray(rows) ? rows : []).map((row) => ({
    year: Number(row?.chart_year || row?.year || 0),
    ha: clean(row?.ha).toUpperCase(),
    status: clean(row?.status) || "Unknown",
    channel: clean(row?.channel_label || row?.channel) || "Unknown",
    bookingCount: Number(row?.booking_count || row?.bookingCount || 0),
  })).filter((row) => row.year && row.bookingCount > 0);
}

function normalizeMonthlyShareRows(rows) {
  return (Array.isArray(rows) ? rows : []).map((row) => ({
    yearMonth: clean(row?.year_month || row?.yearMonth),
    channel: clean(row?.channel_label || row?.channel) || "Unknown",
    bookingCount: Number(row?.booking_count || row?.bookingCount || 0),
  })).filter((row) => /^\d{4}-\d{2}$/.test(row.yearMonth) && row.bookingCount > 0);
}

function previousTwelveMonths() {
  const current = new Date();
  const result = [];
  for (let offset = 12; offset >= 1; offset -= 1) {
    const month = new Date(Date.UTC(current.getUTCFullYear(), current.getUTCMonth() - offset, 1));
    result.push(month.toISOString().slice(0, 7));
  }
  return result;
}

function buildPieCharts(rows, pieYears) {
  return pieYears.map((year) => {
    const byChannel = new Map();
    rows
      .filter((row) => Number(row.year) === Number(year))
      .forEach((row) => {
        const channel = clean(row.channel) || "Unknown";
        byChannel.set(channel, (byChannel.get(channel) || 0) + Number(row.bookingCount || 0));
      });
    const channelRows = Array.from(byChannel.entries())
      .map(([channel, bookingCount]) => ({ channel, bookingCount }))
      .sort((a, b) => Number(b.bookingCount || 0) - Number(a.bookingCount || 0) || clean(a.channel).localeCompare(clean(b.channel)));
    return {
      year: String(year),
      rows: channelRows,
    };
  });
}

function normalizeBookingWindowResult(result = {}) {
  const source = Array.isArray(result) ? result[0] || {} : result || {};
  return {
    years: Array.isArray(source.years) ? source.years.map((year) => String(year)).filter(Boolean) : [],
    summary: source.summary && typeof source.summary === "object" ? source.summary : {},
    distribution: Array.isArray(source.distribution) ? source.distribution : [],
    months: Array.isArray(source.months) ? source.months : [],
    channels: Array.isArray(source.channels) ? source.channels : [],
    yearTrend: Array.isArray(source.yearTrend) ? source.yearTrend : [],
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

    await requireFeature(req, "app", "bookings-bi");

    const selectedYear = parseYearValue(req.query?.year);
    const selectedHa = parseHaValue(req.query?.ha);
    const statuses = parseStatuses(req.query?.statuses || req.query?.booking_statuses);
    const tab = clean(req.query?.tab).toLowerCase();
    if (tab === "booking-window") {
      const bookingWindow = normalizeBookingWindowResult(await restQuery("rpc/bookings_bi_booking_window", {
        method: "POST",
        body: {
          p_year: selectedYear,
          p_ha: selectedHa || null,
          p_statuses: statuses,
        },
      }));
      res.status(200).json({
        year: String(selectedYear),
        ha: selectedHa,
        statuses,
        ...bookingWindow,
      });
      return;
    }
    const params = {
      method: "POST",
      body: {
        p_year: selectedYear,
        p_ha: selectedHa || null,
        p_statuses: statuses,
      },
    };
    const [rpcRows, monthlyShareRows] = await Promise.all([
      restQuery("rpc/bookings_bi_channels", params),
      restQuery("rpc/bookings_bi_channel_share_last_12_months", {
        method: "POST",
        body: {
          p_ha: selectedHa || null,
          p_statuses: statuses,
        },
      }),
    ]);
    const rows = normalizeRows(rpcRows);
    const pieYears = buildPieYears(selectedYear);
    const years = [...new Set([
      ...pieYears.map(String),
      ...rows.map((row) => String(row.year)),
    ])].sort((a, b) => Number(b) - Number(a));

    res.status(200).json({
      year: String(selectedYear),
      ha: selectedHa,
      statuses,
      years,
      channels: {
        pieYears: pieYears.map(String),
        pieCharts: buildPieCharts(rows, pieYears),
        monthlyShare: normalizeMonthlyShareRows(monthlyShareRows),
        monthlyShareMonths: previousTwelveMonths(),
      },
    });
  } catch (error) {
    sendError(res, error);
  }
};
