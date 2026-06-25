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

function parseYearMonthValue(value) {
  const raw = clean(value);
  if (!/^\d{4}-\d{2}$/.test(raw)) return null;
  const [year, month] = raw.split("-").map((part) => Number.parseInt(part, 10));
  if (!Number.isFinite(year) || !Number.isFinite(month) || month < 1 || month > 12) return null;
  return `${year}-${String(month).padStart(2, "0")}`;
}

function parseHaValue(value) {
  const raw = clean(value).toUpperCase();
  return raw === "H" || raw === "A" ? raw : "";
}

function defaultBookingStatuses() {
  return ["Checked Out", "Checked In", "Confirmed"];
}

function parseBookingStatuses(value) {
  const source = Array.isArray(value) ? value : clean(value).split(",");
  const statuses = source.map((item) => clean(item)).filter(Boolean);
  return statuses.length ? [...new Set(statuses)] : defaultBookingStatuses();
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

function defaultPreviousMonthKey() {
  const now = new Date();
  const year = now.getUTCFullYear();
  const month = now.getUTCMonth() + 1;
  const previousMonth = month === 1 ? 12 : month - 1;
  const previousYear = month === 1 ? year - 1 : year;
  return `${previousYear}-${String(previousMonth).padStart(2, "0")}`;
}

function defaultPreviousMonthValue() {
  return defaultPreviousMonthKey().slice(5, 7);
}

function parseMonthValue(value) {
  const raw = clean(value).padStart(2, "0");
  if (!/^\d{2}$/.test(raw)) return defaultPreviousMonthValue();
  const month = Number.parseInt(raw, 10);
  if (!Number.isFinite(month) || month < 1 || month > 12) return defaultPreviousMonthValue();
  return String(month).padStart(2, "0");
}

function monthStartUtc(yearMonth) {
  const parsed = parseYearMonthValue(yearMonth);
  if (!parsed) return null;
  const [year, month] = parsed.split("-").map((part) => Number.parseInt(part, 10));
  return Date.UTC(year, month - 1, 1);
}

function nextMonthKey(yearMonth) {
  const parsed = parseYearMonthValue(yearMonth);
  if (!parsed) return null;
  const [year, month] = parsed.split("-").map((part) => Number.parseInt(part, 10));
  const nextMonth = month === 12 ? 1 : month + 1;
  const nextYear = month === 12 ? year + 1 : year;
  return `${nextYear}-${String(nextMonth).padStart(2, "0")}`;
}

function overlapNightsInMonth(checkInRaw, checkOutRaw, yearMonth) {
  const start = parseDateOnly(checkInRaw);
  const end = parseDateOnly(checkOutRaw);
  const startMonthUtc = monthStartUtc(yearMonth);
  const endMonthUtc = monthStartUtc(nextMonthKey(yearMonth));
  if (!start || !end || startMonthUtc == null || endMonthUtc == null) return 0;
  const startUtc = Date.UTC(start.year, start.month - 1, start.day);
  const endUtc = Date.UTC(end.year, end.month - 1, end.day);
  const overlapStart = Math.max(startUtc, startMonthUtc);
  const overlapEnd = Math.min(endUtc, endMonthUtc);
  return Math.max(Math.round((overlapEnd - overlapStart) / 86400000), 0);
}

function isPortugalCountry(label, code) {
  const normalizedCode = clean(code).toUpperCase();
  if (normalizedCode === "PRT" || normalizedCode === "PTR") return true;
  return clean(label).toUpperCase() === "PORTUGAL";
}

function monthsTouchedByStay(checkInRaw, checkOutRaw) {
  const start = parseDateOnly(checkInRaw);
  if (!start) return [];
  const end = parseDateOnly(checkOutRaw);
  const startUtc = Date.UTC(start.year, start.month - 1, start.day);
  const endUtc = end ? Date.UTC(end.year, end.month - 1, end.day) : startUtc;
  const lastNightUtc = endUtc > startUtc ? endUtc - 86400000 : startUtc;
  const months = [];
  let cursorYear = start.year;
  let cursorMonth = start.month;
  const endDate = new Date(lastNightUtc);
  const endYear = endDate.getUTCFullYear();
  const endMonth = endDate.getUTCMonth() + 1;
  while (cursorYear < endYear || (cursorYear === endYear && cursorMonth <= endMonth)) {
    months.push(`${cursorYear}-${String(cursorMonth).padStart(2, "0")}`);
    cursorMonth += 1;
    if (cursorMonth > 12) {
      cursorMonth = 1;
      cursorYear += 1;
    }
  }
  return months;
}

async function fetchAllGuestsBiSourceRows(selectedHa = "") {
  const rows = [];
  const pageSize = 1000;
  let offset = 0;
  while (true) {
    const query = [
      "guest_records?select=check_in,check_out,birth_date,ha,nationality,nationality_code,issuer_country,issuer_country_code",
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

function nextMonthStartDate(yearMonth) {
  const next = nextMonthKey(yearMonth);
  return next ? `${next}-01` : "";
}

async function fetchFdmBookingsTmtRows(yearMonth, selectedHa = "", bookingStatuses = defaultBookingStatuses()) {
  const rows = [];
  const pageSize = 1000;
  let offset = 0;
  const startDate = `${yearMonth}-01`;
  const endDate = nextMonthStartDate(yearMonth);
  const statuses = parseBookingStatuses(bookingStatuses);
  if (!parseYearMonthValue(yearMonth) || !endDate) return rows;
  while (true) {
    const query = [
      "import_fdm_bookings?select=channel,nights,guests,room_type,check_in_date",
      `check_in_date=gte.${encodeURIComponent(startDate)}`,
      `check_in_date=lt.${encodeURIComponent(endDate)}`,
      statuses.length ? `status=in.(${statuses.map((status) => encodeURIComponent(status)).join(",")})` : "",
      "order=channel.asc",
      `limit=${pageSize}`,
      `offset=${offset}`,
    ].filter(Boolean).join("&");
    const batch = await restQuery(query, { method: "GET" });
    const list = Array.isArray(batch) ? batch : [];
    rows.push(...list);
    if (list.length < pageSize) break;
    offset += pageSize;
  }
  return selectedHa
    ? rows.filter((row) => {
      const rowHa = clean(row?.room_type).toLowerCase().includes("cruz") ? "A" : "H";
      return rowHa === selectedHa;
    })
    : rows;
}

async function buildFdmBookingsTmt(yearMonth, selectedHa = "", bookingStatuses = defaultBookingStatuses()) {
  const statuses = parseBookingStatuses(bookingStatuses);
  const rows = await fetchFdmBookingsTmtRows(yearMonth, selectedHa, statuses);
  const byChannel = new Map();
  rows.forEach((row) => {
    const channel = clean(row?.channel) || "Unknown";
    const nights = Number(row?.nights || 0);
    const guests = Number(row?.guests || 0);
    const totalNights = Math.max(nights, 0) * Math.max(guests, 0);
    byChannel.set(channel, Number(byChannel.get(channel) || 0) + totalNights);
  });
  const mappedRows = [...byChannel.entries()]
    .map(([channel, totalNights]) => ({ channel, totalNights }))
    .sort((a, b) => Number(b.totalNights || 0) - Number(a.totalNights || 0) || a.channel.localeCompare(b.channel));
  return {
    yearMonth,
    bookingStatuses: statuses,
    rows: mappedRows,
    totalNights: mappedRows.reduce((sum, row) => sum + Number(row.totalNights || 0), 0),
  };
}

function buildGuestsBiPivotRowsFromSource(sourceRows) {
  const rows = Array.isArray(sourceRows) ? sourceRows : [];
  const years = [...new Set(
    rows
      .map((row) => clean(row?.check_in).slice(0, 4))
      .filter((value) => /^\d{4}$/.test(value))
  )].sort((a, b) => b.localeCompare(a));
  const countryYearMap = new Map();
  const totalByCountry = new Map();
  rows.forEach((row) => {
    const year = clean(row?.check_in).slice(0, 4);
    if (!/^\d{4}$/.test(year)) return;
    const label = normalizeCountryLabel(row?.nationality, row?.nationality_code);
    const key = `${label}||${year}`;
    countryYearMap.set(key, Number(countryYearMap.get(key) || 0) + 1);
    totalByCountry.set(label, Number(totalByCountry.get(label) || 0) + 1);
  });
  const pivotRows = [...totalByCountry.entries()]
    .map(([countryLabel, total]) => ({
      countryLabel,
      total,
      values: Object.fromEntries(years.map((year) => [year, Number(countryYearMap.get(`${countryLabel}||${year}`) || 0)])),
    }))
    .sort((a, b) => Number(b.total || 0) - Number(a.total || 0) || a.countryLabel.localeCompare(b.countryLabel));
  return { years, rows: pivotRows };
}

function buildGuestsBiIneFromSource(sourceRows, selectedYearMonth) {
  const rows = Array.isArray(sourceRows) ? sourceRows : [];
  const monthSet = new Set([selectedYearMonth, defaultPreviousMonthKey()]);
  const summaryMap = new Map([
    ["pt-residents", { rowLabel: "Portugueses residentes em Portugal", guestsEntered: 0, guestsSlept: 0, nights: 0 }],
    ["foreign-residents", { rowLabel: "Estrangeiros residentes em Portugal", guestsEntered: 0, guestsSlept: 0, nights: 0 }],
  ]);
  const detailMap = new Map();

  rows.forEach((row) => {
    const checkIn = clean(row?.check_in);
    const checkOut = clean(row?.check_out);
    if (!checkIn) return;
    monthsTouchedByStay(checkIn, checkOut).forEach((month) => monthSet.add(month));
    const nationalityLabel = normalizeCountryLabel(row?.nationality, row?.nationality_code);
    const issuerLabel = normalizeCountryLabel(row?.issuer_country, row?.issuer_country_code);
    const nationalityPortugal = isPortugalCountry(nationalityLabel, row?.nationality_code);
    const issuerPortugal = isPortugalCountry(issuerLabel, row?.issuer_country_code);
    const guestsEntered = checkIn.startsWith(`${selectedYearMonth}-`) ? 1 : 0;
    const nights = overlapNightsInMonth(checkIn, checkOut, selectedYearMonth);
    const guestsSlept = nights > 0 ? 1 : 0;

    if (issuerPortugal) {
      const key = nationalityPortugal ? "pt-residents" : "foreign-residents";
      const bucket = summaryMap.get(key);
      bucket.guestsEntered += guestsEntered;
      bucket.guestsSlept += guestsSlept;
      bucket.nights += nights;
    }

    if (!nationalityPortugal && nights > 0) {
      const key = nationalityLabel || "Unknown";
      const bucket = detailMap.get(key) || { rowLabel: key, guestsEntered: 0, guestsSlept: 0, nights: 0 };
      bucket.guestsEntered += guestsEntered;
      bucket.guestsSlept += guestsSlept;
      bucket.nights += nights;
      detailMap.set(key, bucket);
    }
  });

  const months = [...monthSet].filter((value) => /^\d{4}-\d{2}$/.test(value)).sort((a, b) => b.localeCompare(a));
  const detailRows = [...detailMap.values()]
    .sort((a, b) => Number(b.nights || 0) - Number(a.nights || 0) || a.rowLabel.localeCompare(b.rowLabel));

  return {
    yearMonth: selectedYearMonth,
    months,
    summaryRows: [...summaryMap.values()],
    detailRows,
  };
}

function detailsAgeSegment(age) {
  if (!Number.isFinite(age)) return { label: "Unknown", sortOrder: 8 };
  if (age <= 12) return { label: "0-12", sortOrder: 1 };
  if (age <= 17) return { label: "13-17", sortOrder: 2 };
  if (age <= 25) return { label: "18-25", sortOrder: 3 };
  if (age <= 35) return { label: "26-35", sortOrder: 4 };
  if (age <= 45) return { label: "36-45", sortOrder: 5 };
  if (age <= 55) return { label: "46-55", sortOrder: 6 };
  if (age <= 65) return { label: "56-65", sortOrder: 7 };
  return { label: "66+", sortOrder: 8 };
}

function averageNumber(values) {
  const safeValues = (Array.isArray(values) ? values : []).map((value) => Number(value)).filter(Number.isFinite);
  if (!safeValues.length) return null;
  return Number((safeValues.reduce((sum, value) => sum + value, 0) / safeValues.length).toFixed(2));
}

function buildGuestsBiDetailsFromSource(sourceRows, selectedYear) {
  const rows = Array.isArray(sourceRows) ? sourceRows : [];
  const years = [...new Set(
    rows
      .map((row) => clean(row?.check_in).slice(0, 4))
      .filter((value) => /^\d{4}$/.test(value))
  )].sort((a, b) => a.localeCompare(b));
  const selectedRows = rows.filter((row) => clean(row?.check_in).slice(0, 4) === String(selectedYear));
  const selectedAges = selectedRows.map((row) => ageAtDate(row?.birth_date, row?.check_in)).filter((age) => Number.isFinite(age));
  const selectedStays = selectedRows.map((row) => diffDays(row?.check_in, row?.check_out)).filter((nights) => Number.isFinite(nights));
  const segmentMap = new Map([
    ["0-12", { ageSegment: "0-12", sortOrder: 1, guestCount: 0 }],
    ["13-17", { ageSegment: "13-17", sortOrder: 2, guestCount: 0 }],
    ["18-25", { ageSegment: "18-25", sortOrder: 3, guestCount: 0 }],
    ["26-35", { ageSegment: "26-35", sortOrder: 4, guestCount: 0 }],
    ["36-45", { ageSegment: "36-45", sortOrder: 5, guestCount: 0 }],
    ["46-55", { ageSegment: "46-55", sortOrder: 6, guestCount: 0 }],
    ["56-65", { ageSegment: "56-65", sortOrder: 7, guestCount: 0 }],
    ["66+", { ageSegment: "66+", sortOrder: 8, guestCount: 0 }],
    ["Unknown", { ageSegment: "Unknown", sortOrder: 9, guestCount: 0 }],
  ]);
  selectedRows.forEach((row) => {
    const segment = detailsAgeSegment(ageAtDate(row?.birth_date, row?.check_in));
    const bucket = segmentMap.get(segment.label) || segmentMap.get("Unknown");
    bucket.guestCount += 1;
  });
  const trendRows = years.map((year) => {
    const yearRows = rows.filter((row) => clean(row?.check_in).slice(0, 4) === year);
    return {
      year,
      guestCount: yearRows.length,
      averageAge: averageNumber(yearRows.map((row) => ageAtDate(row?.birth_date, row?.check_in))),
      averageStay: averageNumber(yearRows.map((row) => diffDays(row?.check_in, row?.check_out))),
    };
  });
  return {
    year: String(selectedYear),
    summary: {
      guestCount: selectedRows.length,
      averageAge: averageNumber(selectedAges),
      averageStay: averageNumber(selectedStays),
    },
    ageSegments: [...segmentMap.values()],
    trends: trendRows,
  };
}

function buildGuestsBiFallbackPayload(sourceRows, selectedYear, selectedMonthFilterYear = "") {
  const rows = Array.isArray(sourceRows) ? sourceRows : [];
  const availableYears = [...new Set(
    rows
      .map((row) => clean(row?.check_in).slice(0, 4))
      .filter((value) => /^\d{4}$/.test(value))
  )].sort((a, b) => b.localeCompare(a));
  const yearsForFilter = [...new Set([...availableYears, String(selectedYear)])].sort((a, b) => b.localeCompare(a));
  const resolvedMonthFilterYear = /^\d{4}$/.test(clean(selectedMonthFilterYear)) ? clean(selectedMonthFilterYear) : "";

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
    if (/^\d{2}$/.test(chartMonth) && (!resolvedMonthFilterYear || chartYear === resolvedMonthFilterYear)) {
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

  const pivotPayload = buildGuestsBiPivotRowsFromSource(rows);

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
      monthFilterYears: availableYears,
      monthFilterYear: resolvedMonthFilterYear,
      pivotYears: pivotPayload.years,
      pivotRows: pivotPayload.rows,
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
    const requestedYear = parseYearValue(req.query?.year);
    const requestedMonthFilterYear = parseYearValue(req.query?.month_year);
    const selectedYear = requestedYear || new Date().getUTCFullYear();
    const selectedMonth = parseMonthValue(req.query?.month);
    const selectedBookingStatuses = parseBookingStatuses(req.query?.booking_statuses);
    const selectedTmtYearMonth = `${selectedYear}-${selectedMonth}`;
    const selectedYearMonth = parseYearMonthValue(req.query?.year_month) || defaultPreviousMonthKey();
    let pivotPayload = { years: [], rows: [] };
    try {
      const pivotRows = await restQuery("rpc/guests_bi_nationality_pivot", {
        method: "POST",
        body: selectedHa ? { p_ha: selectedHa } : {},
      });
      const mappedPivotRows = Array.isArray(pivotRows) ? pivotRows : [];
      const pivotYears = [...new Set(mappedPivotRows.map((row) => clean(row?.chart_year)).filter((year) => /^\d{4}$/.test(year)))].sort((a, b) => b.localeCompare(a));
      const byCountry = new Map();
      mappedPivotRows.forEach((row) => {
        const label = normalizeCountryLabel(row?.country_label);
        if (!byCountry.has(label)) {
          byCountry.set(label, {
            countryLabel: label,
            total: Number(row?.row_total || 0),
            values: {},
          });
        }
        byCountry.get(label).values[clean(row?.chart_year)] = Number(row?.guest_count || 0);
        if (!byCountry.get(label).total) byCountry.get(label).total = Number(row?.row_total || 0);
      });
      pivotPayload = {
        years: pivotYears,
        rows: [...byCountry.values()].sort((a, b) => Number(b.total || 0) - Number(a.total || 0) || a.countryLabel.localeCompare(b.countryLabel)),
      };
    } catch {
      pivotPayload = buildGuestsBiPivotRowsFromSource(fallbackSourceRows);
    }
    const payload = buildGuestsBiFallbackPayload(
      fallbackSourceRows,
      selectedYear,
      requestedMonthFilterYear ? String(requestedMonthFilterYear) : ""
    );
    let bookingsTmtPayload;
    try {
      bookingsTmtPayload = await buildFdmBookingsTmt(selectedTmtYearMonth, selectedHa, selectedBookingStatuses);
    } catch {
      bookingsTmtPayload = { yearMonth: selectedTmtYearMonth, bookingStatuses: selectedBookingStatuses, rows: [], totalNights: 0 };
    }
    let inePayload;
    try {
      const ineRows = await restQuery("rpc/guests_bi_ine", {
        method: "POST",
        body: {
          p_year_month: selectedYearMonth,
          ...(selectedHa ? { p_ha: selectedHa } : {}),
        },
      });
      const mappedIneRows = Array.isArray(ineRows) ? ineRows : [];
      const summaryRows = mappedIneRows
        .filter((row) => clean(row?.section) === "summary")
        .sort((a, b) => Number(a?.sort_order || 0) - Number(b?.sort_order || 0))
        .map((row) => ({
          rowLabel: clean(row?.row_label),
          guestsEntered: Number(row?.guests_entered || 0),
          guestsSlept: Number(row?.guests_slept || 0),
          nights: Number(row?.nights || 0),
        }));
      const detailRows = mappedIneRows
        .filter((row) => clean(row?.section) === "detail")
        .sort((a, b) => Number(a?.sort_order || 0) - Number(b?.sort_order || 0) || clean(a?.row_label).localeCompare(clean(b?.row_label)))
        .map((row) => ({
          rowLabel: clean(row?.row_label),
          guestsEntered: Number(row?.guests_entered || 0),
          guestsSlept: Number(row?.guests_slept || 0),
          nights: Number(row?.nights || 0),
        }));
      inePayload = {
        ...buildGuestsBiIneFromSource(fallbackSourceRows, selectedYearMonth),
        summaryRows,
        detailRows,
      };
    } catch {
      inePayload = buildGuestsBiIneFromSource(fallbackSourceRows, selectedYearMonth);
    }
    let detailsPayload;
    try {
      const detailsRows = await restQuery("rpc/guests_bi_details", {
        method: "POST",
        body: {
          p_year: selectedYear,
          ...(selectedHa ? { p_ha: selectedHa } : {}),
        },
      });
      const mappedDetailsRows = Array.isArray(detailsRows) ? detailsRows : [];
      const summaryRow = mappedDetailsRows.find((row) => clean(row?.section) === "summary") || {};
      const ageSegments = mappedDetailsRows
        .filter((row) => clean(row?.section) === "segment")
        .sort((a, b) => Number(a?.sort_order || 0) - Number(b?.sort_order || 0))
        .map((row) => ({
          ageSegment: clean(row?.age_segment),
          sortOrder: Number(row?.sort_order || 0),
          guestCount: Number(row?.guest_count || 0),
        }));
      const trends = mappedDetailsRows
        .filter((row) => clean(row?.section) === "trend")
        .sort((a, b) => Number(a?.chart_year || 0) - Number(b?.chart_year || 0))
        .map((row) => ({
          year: clean(row?.chart_year),
          guestCount: Number(row?.guest_count || 0),
          averageAge: row?.average_age === null || row?.average_age === undefined ? null : Number(row.average_age),
          averageStay: row?.average_stay === null || row?.average_stay === undefined ? null : Number(row.average_stay),
        }));
      detailsPayload = {
        year: String(selectedYear),
        summary: {
          guestCount: Number(summaryRow?.guest_count || 0),
          averageAge: summaryRow?.average_age === null || summaryRow?.average_age === undefined ? null : Number(summaryRow.average_age),
          averageStay: summaryRow?.average_stay === null || summaryRow?.average_stay === undefined ? null : Number(summaryRow.average_stay),
        },
        ageSegments,
        trends,
      };
    } catch {
      detailsPayload = buildGuestsBiDetailsFromSource(fallbackSourceRows, selectedYear);
    }
    res.status(200).json({
      ...payload,
      nationalities: {
        ...payload.nationalities,
        pivotYears: pivotPayload.years,
        pivotRows: pivotPayload.rows,
      },
      bookingsTmt: bookingsTmtPayload,
      ine: inePayload,
      details: detailsPayload,
      ha: selectedHa,
    });
  } catch (error) {
    sendError(res, error);
  }
};
