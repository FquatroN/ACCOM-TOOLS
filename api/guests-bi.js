const { requireFeature, restQuery, sendError } = require("./_supabase");
const GUESTS_BI_CACHE_TTL_MS = 120000;
const guestsBiSourceCache = new Map();

function clean(value) {
  return String(value ?? "").trim();
}

function parseYearValue(value) {
  const raw = clean(value);
  if (!/^\d{4}$/.test(raw)) return null;
  const year = Number.parseInt(raw, 10);
  return Number.isFinite(year) ? year : null;
}

function parseHaValue(value) {
  const raw = clean(value).toUpperCase();
  return raw === "H" || raw === "A" ? raw : "";
}

function parseDateOnly(value) {
  const raw = clean(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return null;
  const [year, month, day] = raw.split("-").map((part) => Number.parseInt(part, 10));
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return null;
  return { year, month, day, raw };
}

function diffDays(startRaw, endRaw) {
  const start = parseDateOnly(startRaw);
  const end = parseDateOnly(endRaw);
  if (!start || !end) return 0;
  const startUtc = Date.UTC(start.year, start.month - 1, start.day);
  const endUtc = Date.UTC(end.year, end.month - 1, end.day);
  return Math.max(Math.round((endUtc - startUtc) / 86400000), 0);
}

function ageAtDate(birthRaw, dateRaw) {
  const birth = parseDateOnly(birthRaw);
  const date = parseDateOnly(dateRaw);
  if (!birth || !date) return null;
  let age = date.year - birth.year;
  if (date.month < birth.month || (date.month === birth.month && date.day < birth.day)) age -= 1;
  return age;
}

function normalizeCountryLabel(value, fallbackCode = "") {
  const label = clean(value);
  if (label) return label;
  const code = clean(fallbackCode);
  return code || "Unknown";
}

async function fetchAllGuestsBiSourceRows(selectedHa = "") {
  const rows = [];
  const pageSize = 1000;
  let offset = 0;
  while (true) {
    const query = [
      "guest_records?select=check_in,check_out,birth_date,ha,nationality,nationality_code",
      "check_in=not.is.null",
      "order=check_in.desc",
      `limit=${pageSize}`,
      `offset=${offset}`,
      selectedHa ? `ha=eq.${encodeURIComponent(selectedHa)}` : "",
    ].filter(Boolean).join("&");
    const batch = await restQuery(query, { method: "GET" });
    const list = Array.isArray(batch) ? batch : [];
    rows.push(...list);
    if (list.length < pageSize) break;
    offset += pageSize;
  }
  return rows;
}

async function getCachedGuestsBiSourceRows(selectedHa = "") {
  const cacheKey = selectedHa || "ALL";
  const cached = guestsBiSourceCache.get(cacheKey);
  if (cached && (Date.now() - cached.loadedAt) < GUESTS_BI_CACHE_TTL_MS) {
    return cached.rows;
  }
  const rows = await fetchAllGuestsBiSourceRows(selectedHa);
  guestsBiSourceCache.set(cacheKey, {
    loadedAt: Date.now(),
    rows,
  });
  return rows;
}

function buildGuestsBiFallbackPayload(sourceRows, selectedYear) {
  const rows = Array.isArray(sourceRows) ? sourceRows : [];
  const availableYears = [...new Set(
    rows
      .map((row) => clean(row?.check_in).slice(0, 4))
      .filter((value) => /^\d{4}$/.test(value))
  )].sort((a, b) => b.localeCompare(a));
  const yearsForFilter = [...new Set([...availableYears, String(selectedYear)])].sort((a, b) => b.localeCompare(a));

  const monthMap = new Map();
  for (let month = 1; month <= 12; month += 1) {
    const yearMonth = `${selectedYear}-${String(month).padStart(2, "0")}`;
    monthMap.set(yearMonth, {
      yearMonth,
      totalNights: 0,
      exempt7Days: 0,
      exempt13Year: 0,
    });
  }

  const anchorYear = Number.parseInt(String(selectedYear), 10);
  const safeAnchorYear = Number.isFinite(anchorYear) ? anchorYear : new Date().getFullYear();
  const pieYears = [0, 1, 2, 3].map((offset) => String(safeAnchorYear - offset));
  const globalTotals = new Map();
  const pieYearCountry = new Map();
  const lineYearCountry = new Map();
  const monthCountry = new Map();

  rows.forEach((row) => {
    const checkIn = clean(row?.check_in);
    const checkOut = clean(row?.check_out);
    const birthDate = clean(row?.birth_date);
    const chartYear = checkIn.slice(0, 4);
    const chartMonth = checkIn.slice(5, 7);
    const label = normalizeCountryLabel(row?.nationality, row?.nationality_code);

    if (/^\d{4}$/.test(chartYear)) {
      globalTotals.set(label, Number(globalTotals.get(label) || 0) + 1);
      const pieKey = `${chartYear}||${label}`;
      pieYearCountry.set(pieKey, Number(pieYearCountry.get(pieKey) || 0) + 1);
      const lineKey = `${chartYear}||${label}`;
      lineYearCountry.set(lineKey, Number(lineYearCountry.get(lineKey) || 0) + 1);
    }
    if (/^\d{2}$/.test(chartMonth)) {
      const monthKey = `${chartMonth}||${label}`;
      monthCountry.set(monthKey, Number(monthCountry.get(monthKey) || 0) + 1);
    }

    if (chartYear !== String(selectedYear) || !checkOut) return;
    const bucket = monthMap.get(checkIn.slice(0, 7));
    if (!bucket) return;
    const nights = diffDays(checkIn, checkOut);
    bucket.totalNights += nights;
    bucket.exempt7Days += Math.max(nights - 7, 0);
    const age = ageAtDate(birthDate, checkIn);
    if (age !== null && age < 13) bucket.exempt13Year += nights;
  });

  const topCountries = [...globalTotals.entries()]
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .slice(0, 20)
    .map(([label]) => label);
  const topCountrySet = new Set(topCountries);

  const pieCharts = pieYears.map((year) => {
    const entries = topCountries
      .map((label) => ({
        countryLabel: label,
        guestCount: Number(pieYearCountry.get(`${year}||${label}`) || 0),
      }))
      .filter((item) => item.guestCount > 0);
    const others = [...pieYearCountry.entries()]
      .filter(([key]) => key.startsWith(`${year}||`))
      .reduce((sum, [key, value]) => {
        const label = key.split("||")[1];
        return topCountrySet.has(label) ? sum : sum + Number(value || 0);
      }, 0);
    if (others > 0) entries.push({ countryLabel: "Others", guestCount: others });
    entries.sort((a, b) => {
      if (a.countryLabel === "Others" && b.countryLabel !== "Others") return 1;
      if (b.countryLabel === "Others" && a.countryLabel !== "Others") return -1;
      return b.guestCount - a.guestCount || a.countryLabel.localeCompare(b.countryLabel);
    });
    return { year, rows: entries };
  });

  const lineYears = availableYears.slice().sort((a, b) => a.localeCompare(b));
  const lineSeries = topCountries.map((label) => ({
    countryLabel: label,
    values: lineYears.map((year) => Number(lineYearCountry.get(`${year}||${label}`) || 0)),
  }));
  const othersLine = {
    countryLabel: "Others",
    values: lineYears.map((year) => [...lineYearCountry.entries()]
      .filter(([key]) => key.startsWith(`${year}||`))
      .reduce((sum, [key, value]) => {
        const label = key.split("||")[1];
        return topCountrySet.has(label) ? sum : sum + Number(value || 0);
      }, 0)),
  };
  if (othersLine.values.some((value) => value > 0)) lineSeries.push(othersLine);

  const monthLabels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const monthSeries = topCountries.map((label) => ({
    countryLabel: label,
    values: monthLabels.map((_, index) => Number(monthCountry.get(`${String(index + 1).padStart(2, "0")}||${label}`) || 0)),
  }));
  const othersMonth = {
    countryLabel: "Others",
    values: monthLabels.map((_, index) => [...monthCountry.entries()]
      .filter(([key]) => key.startsWith(`${String(index + 1).padStart(2, "0")}||`))
      .reduce((sum, [key, value]) => {
        const label = key.split("||")[1];
        return topCountrySet.has(label) ? sum : sum + Number(value || 0);
      }, 0)),
  };
  if (othersMonth.values.some((value) => value > 0)) monthSeries.push(othersMonth);

  const mappedRows = [...monthMap.values()];
  const totals = mappedRows.reduce((acc, row) => ({
    totalNights: acc.totalNights + row.totalNights,
    exempt7Days: acc.exempt7Days + row.exempt7Days,
    exempt13Year: acc.exempt13Year + row.exempt13Year,
  }), { totalNights: 0, exempt7Days: 0, exempt13Year: 0 });

  return {
    year: String(selectedYear),
    years: yearsForFilter.length ? yearsForFilter : [String(selectedYear)],
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
  };
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "guests-bi");
    if (req.method !== "GET") {
      res.setHeader("Allow", "GET");
      res.status(405).json({ error: "Method not allowed." });
      return;
    }

    const selectedHa = parseHaValue(req.query?.ha);
    const fallbackSourceRows = await getCachedGuestsBiSourceRows(selectedHa);
    const availableYears = [...new Set(
      fallbackSourceRows
        .map((row) => clean(row?.check_in).slice(0, 4))
        .filter((year) => /^\d{4}$/.test(year))
    )].sort((a, b) => b.localeCompare(a));
    const requestedYear = parseYearValue(req.query?.year);
    const selectedYear = requestedYear || new Date().getUTCFullYear();
    const payload = buildGuestsBiFallbackPayload(fallbackSourceRows, selectedYear);
    res.status(200).json({
      ...payload,
      ha: selectedHa,
    });
  } catch (error) {
    sendError(res, error);
  }
};
