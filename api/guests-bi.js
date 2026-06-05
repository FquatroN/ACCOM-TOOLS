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
    const availableYears = [...new Set(
      fallbackSourceRows
        .map((row) => clean(row?.check_in).slice(0, 4))
        .filter((year) => /^\d{4}$/.test(year))
    )].sort((a, b) => b.localeCompare(a));
    const requestedYear = parseYearValue(req.query?.year);
    const selectedYear = requestedYear || new Date().getUTCFullYear();
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
    const payload = buildGuestsBiFallbackPayload(fallbackSourceRows, selectedYear);
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
    res.status(200).json({
      ...payload,
      nationalities: {
        ...payload.nationalities,
        pivotYears: pivotPayload.years,
        pivotRows: pivotPayload.rows,
      },
      ine: inePayload,
      ha: selectedHa,
    });
  } catch (error) {
    sendError(res, error);
  }
};
