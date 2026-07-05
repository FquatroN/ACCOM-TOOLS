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
    const rpcRows = await restQuery("rpc/bookings_bi_channels", {
      method: "POST",
      body: {
        p_year: selectedYear,
        p_ha: selectedHa || null,
        p_statuses: statuses,
      },
    });
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
      },
    });
  } catch (error) {
    sendError(res, error);
  }
};
