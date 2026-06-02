const { requireFeature, restQuery, sendError } = require("./_supabase");

function clean(value) {
  return String(value ?? "").trim();
}

function toUtcDate(isoDate) {
  const raw = clean(isoDate);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return null;
  const date = new Date(`${raw}T00:00:00Z`);
  return Number.isNaN(date.getTime()) ? null : date;
}

function toYearMonth(isoDate) {
  return clean(isoDate).slice(0, 7);
}

function nightsBetween(checkIn, checkOut) {
  const start = toUtcDate(checkIn);
  const end = toUtcDate(checkOut);
  if (!start || !end) return 0;
  const diffDays = Math.round((end.getTime() - start.getTime()) / 86400000);
  return diffDays > 0 ? diffDays : 0;
}

function ageAtDate(birthDate, checkIn) {
  const birth = toUtcDate(birthDate);
  const start = toUtcDate(checkIn);
  if (!birth || !start) return null;
  let years = start.getUTCFullYear() - birth.getUTCFullYear();
  const birthMonth = birth.getUTCMonth();
  const startMonth = start.getUTCMonth();
  if (
    startMonth < birthMonth ||
    (startMonth === birthMonth && start.getUTCDate() < birth.getUTCDate())
  ) {
    years -= 1;
  }
  return years;
}

function buildMonthRows(year) {
  return Array.from({ length: 12 }, (_, index) => {
    const month = String(index + 1).padStart(2, "0");
    return {
      yearMonth: `${year}-${month}`,
      totalNights: 0,
      exempt7Days: 0,
      exempt13Year: 0,
    };
  });
}

async function loadAllGuestBiRows(basePath) {
  const pageSize = 5000;
  const rows = [];
  let offset = 0;
  while (true) {
    const page = await restQuery(`${basePath}&limit=${pageSize}&offset=${offset}`, { method: "GET" });
    const batch = Array.isArray(page) ? page : [];
    rows.push(...batch);
    if (batch.length < pageSize) break;
    offset += batch.length;
  }
  return rows;
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "guests-bi");
    if (req.method !== "GET") {
      res.setHeader("Allow", "GET");
      res.status(405).json({ error: "Method not allowed." });
      return;
    }

    const yearRows = await loadAllGuestBiRows(
      "guest_records?select=check_in&check_in=not.is.null&order=check_in.desc"
    );
    const availableYears = [...new Set((Array.isArray(yearRows) ? yearRows : [])
      .map((row) => clean(row?.check_in).slice(0, 4))
      .filter((year) => /^\d{4}$/.test(year)))]
      .sort((a, b) => b.localeCompare(a));

    const requestedYear = clean(req.query?.year);
    const selectedYear = /^\d{4}$/.test(requestedYear)
      ? requestedYear
      : (availableYears[0] || String(new Date().getUTCFullYear()));

    const nextYear = String(Number(selectedYear) + 1);
    const rows = await loadAllGuestBiRows(
      `guest_records?select=check_in,check_out,birth_date&check_in=gte.${selectedYear}-01-01&check_in=lt.${nextYear}-01-01&order=check_in.asc`
    );

    const monthRows = buildMonthRows(selectedYear);
    const monthMap = new Map(monthRows.map((row) => [row.yearMonth, row]));
    const totals = { totalNights: 0, exempt7Days: 0, exempt13Year: 0 };

    for (const row of Array.isArray(rows) ? rows : []) {
      const checkIn = clean(row?.check_in);
      const checkOut = clean(row?.check_out);
      const birthDate = clean(row?.birth_date);
      const yearMonth = toYearMonth(checkIn);
      const month = monthMap.get(yearMonth);
      if (!month) continue;
      const nights = nightsBetween(checkIn, checkOut);
      if (!nights) continue;
      const exempt7Days = Math.max(0, nights - 7);
      const age = ageAtDate(birthDate, checkIn);
      const exempt13Year = age !== null && age < 13 ? nights : 0;
      month.totalNights += nights;
      month.exempt7Days += exempt7Days;
      month.exempt13Year += exempt13Year;
      totals.totalNights += nights;
      totals.exempt7Days += exempt7Days;
      totals.exempt13Year += exempt13Year;
    }

    res.status(200).json({
      year: selectedYear,
      years: availableYears.length ? availableYears : [selectedYear],
      rows: monthRows,
      totals,
    });
  } catch (error) {
    sendError(res, error);
  }
};
