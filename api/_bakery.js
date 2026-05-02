const bakeryDefaults = require("./_bakery-defaults.json");

function cleanText(value) {
  return String(value ?? "").trim();
}

function parseBool(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  return ["true", "1", "yes", "sim", "on"].includes(cleanText(value).toLowerCase());
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanText(value));
}

function normalizeRecipients(value) {
  const source = Array.isArray(value) ? value.join(",") : String(value || "");
  const seen = new Set();
  return source
    .split(/[\n,;]/)
    .map((item) => cleanText(item).toLowerCase())
    .filter((item) => isValidEmail(item))
    .filter((item) => {
      if (seen.has(item)) return false;
      seen.add(item);
      return true;
    });
}

function normalizeBase(value) {
  const raw = cleanText(value).toLowerCase().replace(/\s+/g, "-");
  if (["base-baixa", "baixa", "low"].includes(raw)) return "base-baixa";
  if (["base-alta", "alta", "high"].includes(raw)) return "base-alta";
  return "base-media";
}

function sanitizeBreadTableRow(row = {}, index = 0) {
  return {
    guests: Math.max(1, Number.parseInt(row.guests ?? row.persons ?? row.people ?? index + 1, 10) || index + 1),
    baseBaixa: Math.max(0, Number.parseInt(row.baseBaixa ?? row.base_baixa ?? row.baseLow ?? 0, 10) || 0),
    baseMedia: Math.max(0, Number.parseInt(row.baseMedia ?? row.base_media ?? row.baseMedium ?? 0, 10) || 0),
    baseAlta: Math.max(0, Number.parseInt(row.baseAlta ?? row.base_alta ?? row.baseHigh ?? 0, 10) || 0),
  };
}

function sanitizeBreadType(row = {}, index = 0) {
  return {
    id: cleanText(row.id) || `bread-type-${index + 1}`,
    name: cleanText(row.name || row.type || row.breadType),
    percentage: Math.max(0, Number(row.percentage ?? row.percent ?? 0) || 0),
  };
}

const DEFAULT_BAKERY_SETTINGS = {
  selectedBase: normalizeBase(bakeryDefaults.selectedBase),
  hostelCapacity: Math.max(1, Number(bakeryDefaults.hostelCapacity) || 83),
  emailRecipients: [],
  emailConfig: {
    provider: "resend",
    smtpHost: "smtp.gmail.com",
    smtpPort: 465,
    smtpSecure: true,
    smtpUser: "",
    smtpPassword: "",
    fromEmail: "",
    fromName: "Lisboa Central Hostel",
  },
  breadTable: (Array.isArray(bakeryDefaults.breadTable) ? bakeryDefaults.breadTable : []).map(sanitizeBreadTableRow),
  breadTypes: (Array.isArray(bakeryDefaults.breadTypes) ? bakeryDefaults.breadTypes : []).map(sanitizeBreadType).filter((item) => item.name),
};

function sanitizeBakeryEmailConfig(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const provider = cleanText(source.provider).toLowerCase() === "smtp" ? "smtp" : "resend";
  const smtpHost = cleanText(source.smtpHost || source.smtp_host) || DEFAULT_BAKERY_SETTINGS.emailConfig.smtpHost;
  const smtpPort = Math.max(1, Number.parseInt(source.smtpPort ?? source.smtp_port ?? DEFAULT_BAKERY_SETTINGS.emailConfig.smtpPort, 10) || DEFAULT_BAKERY_SETTINGS.emailConfig.smtpPort);
  const smtpSecure = source.smtpSecure === undefined && source.smtp_secure === undefined
    ? !!DEFAULT_BAKERY_SETTINGS.emailConfig.smtpSecure
    : parseBool(source.smtpSecure ?? source.smtp_secure);
  const smtpUser = cleanText(source.smtpUser || source.smtp_user).toLowerCase();
  const smtpPassword = String(source.smtpPassword ?? source.smtp_password ?? "");
  const fromEmail = cleanText(source.fromEmail || source.from_email).toLowerCase();
  const fromName = cleanText(source.fromName || source.from_name) || DEFAULT_BAKERY_SETTINGS.emailConfig.fromName;
  if (provider === "smtp") {
    if (!smtpHost) {
      const error = new Error("SMTP host is required when Bakery email provider is SMTP.");
      error.statusCode = 400;
      throw error;
    }
    if (!smtpUser || !isValidEmail(smtpUser)) {
      const error = new Error("SMTP user must be a valid email when Bakery email provider is SMTP.");
      error.statusCode = 400;
      throw error;
    }
    if (!smtpPassword) {
      const error = new Error("SMTP password is required when Bakery email provider is SMTP.");
      error.statusCode = 400;
      throw error;
    }
    if (!fromEmail || !isValidEmail(fromEmail)) {
      const error = new Error("From email must be a valid email when Bakery email provider is SMTP.");
      error.statusCode = 400;
      throw error;
    }
  }
  return {
    provider,
    smtpHost,
    smtpPort,
    smtpSecure,
    smtpUser,
    smtpPassword,
    fromEmail,
    fromName,
  };
}

function sanitizeBakerySettings(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const breadTable = (Array.isArray(source.breadTable) ? source.breadTable : DEFAULT_BAKERY_SETTINGS.breadTable)
    .map(sanitizeBreadTableRow)
    .sort((a, b) => a.guests - b.guests);
  const breadTypes = (Array.isArray(source.breadTypes) ? source.breadTypes : DEFAULT_BAKERY_SETTINGS.breadTypes)
    .map(sanitizeBreadType)
    .filter((item) => item.name);
  const totalPercentage = breadTypes.reduce((sum, item) => sum + Number(item.percentage || 0), 0);
  if (breadTypes.length && Math.round(totalPercentage * 100) !== 10000) {
    const error = new Error("Bread type percentages must total 100%.");
    error.statusCode = 400;
    throw error;
  }
  return {
    selectedBase: normalizeBase(source.selectedBase || source.selected_base),
    hostelCapacity: Math.max(1, Number.parseInt(source.hostelCapacity ?? source.hostel_capacity ?? DEFAULT_BAKERY_SETTINGS.hostelCapacity, 10) || DEFAULT_BAKERY_SETTINGS.hostelCapacity),
    emailRecipients: normalizeRecipients(source.emailRecipients || source.email_recipients),
    emailConfig: sanitizeBakeryEmailConfig(source.emailConfig || source.email_config),
    breadTable,
    breadTypes,
  };
}

function formatLisbonIso(date = new Date()) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/Lisbon",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(date);
  const get = (type) => parts.find((item) => item.type === type)?.value || "";
  return `${get("year")}-${get("month")}-${get("day")}`;
}

function nextIsoDate(iso) {
  const date = new Date(`${cleanText(iso)}T00:00:00`);
  if (Number.isNaN(date.getTime())) return "";
  date.setDate(date.getDate() + 1);
  return date.toISOString().slice(0, 10);
}

function easterSunday(year) {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(Date.UTC(year, month - 1, day));
}

function shiftUtcDays(date, days) {
  const clone = new Date(date.getTime());
  clone.setUTCDate(clone.getUTCDate() + days);
  return clone;
}

function lisbonHolidaySet(year) {
  const easter = easterSunday(year);
  const goodFriday = shiftUtcDays(easter, -2);
  const corpusChristi = shiftUtcDays(easter, 60);
  const fixed = [
    `${year}-01-01`,
    `${year}-04-25`,
    `${year}-05-01`,
    `${year}-06-10`,
    `${year}-06-13`,
    `${year}-08-15`,
    `${year}-10-05`,
    `${year}-11-01`,
    `${year}-12-01`,
    `${year}-12-08`,
    `${year}-12-25`,
  ];
  return new Set([
    ...fixed,
    goodFriday.toISOString().slice(0, 10),
    corpusChristi.toISOString().slice(0, 10),
  ]);
}

function isWorkingDayLisbon(iso) {
  const raw = cleanText(iso);
  const date = new Date(`${raw}T00:00:00`);
  if (Number.isNaN(date.getTime())) return false;
  const day = date.getDay();
  if (day === 0 || day === 6) return false;
  const holidays = lisbonHolidaySet(date.getFullYear());
  return !holidays.has(raw);
}

function generateTargetDates(orderDate = formatLisbonIso(new Date())) {
  const dates = [];
  let cursor = nextIsoDate(orderDate);
  while (cursor) {
    dates.push(cursor);
    if (isWorkingDayLisbon(cursor)) break;
    cursor = nextIsoDate(cursor);
  }
  return dates;
}

function lookupBreadTotal(settings, guests) {
  const table = Array.isArray(settings?.breadTable) ? settings.breadTable : [];
  const normalizedGuests = Math.max(0, Number.parseInt(guests, 10) || 0);
  if (!table.length || normalizedGuests <= 0) return 0;
  const sorted = [...table].sort((a, b) => a.guests - b.guests);
  const found = sorted.find((item) => item.guests >= normalizedGuests) || sorted[sorted.length - 1];
  if (!found) return 0;
  if (normalizeBase(settings?.selectedBase) === "base-baixa") return Number(found.baseBaixa || 0);
  if (normalizeBase(settings?.selectedBase) === "base-alta") return Number(found.baseAlta || 0);
  return Number(found.baseMedia || 0);
}

function allocateBreadTypes(total, types = []) {
  const normalized = (Array.isArray(types) ? types : []).map((item) => ({
    ...item,
    percentage: Number(item.percentage || 0),
  })).filter((item) => item.name);
  if (!normalized.length) return [];
  const totalBreads = Math.max(0, Number(total || 0));
  const base = normalized.map((item) => {
    const exact = totalBreads * (item.percentage / 100);
    return { ...item, quantity: Math.floor(exact), remainder: exact - Math.floor(exact) };
  });
  let remaining = totalBreads - base.reduce((sum, item) => sum + item.quantity, 0);
  base.sort((a, b) => b.remainder - a.remainder || b.percentage - a.percentage);
  for (let i = 0; i < base.length && remaining > 0; i += 1, remaining -= 1) base[i].quantity += 1;
  return normalized.map((item) => {
    const found = base.find((entry) => entry.name === item.name);
    return { name: item.name, percentage: item.percentage, quantity: found ? found.quantity : 0 };
  });
}

function parseOptionalCount(value) {
  if (value === "" || value === null || value === undefined) return "";
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? Math.max(0, parsed) : "";
}

function normalizeBakeryDay(day = {}, settings = DEFAULT_BAKERY_SETTINGS, fallbackDate = "") {
  const date = cleanText(day.date || fallbackDate);
  const availableBeds = parseOptionalCount(day.availableBeds ?? day.available_beds);
  const cruzCheckins = parseOptionalCount(day.cruzCheckins ?? day.cruz_checkins);
  const hasAvailableBeds = availableBeds !== "";
  const hasCruzCheckins = cruzCheckins !== "";
  const hostelGuests = hasAvailableBeds ? Math.max(0, Number(settings.hostelCapacity || 0) - Number(availableBeds)) : "";
  const totalBreads = hostelGuests === "" ? "" : lookupBreadTotal(settings, hostelGuests);
  const breadBreakdown = totalBreads === "" ? allocateBreadTypes(0, settings.breadTypes).map((item) => ({ ...item, quantity: "" })) : allocateBreadTypes(totalBreads, settings.breadTypes);
  return {
    date,
    availableBeds,
    cruzCheckins,
    hostelGuests,
    totalBreads,
    breadBreakdown,
    pasteisDeNata: hasCruzCheckins ? cruzCheckins : "",
  };
}

function sanitizeBakeryDays(days = [], settings = DEFAULT_BAKERY_SETTINGS, targetDates = []) {
  const source = Array.isArray(days) ? days : [];
  const byDate = new Map(source.map((item) => [cleanText(item.date), item]));
  return (Array.isArray(targetDates) ? targetDates : []).map((date) => normalizeBakeryDay(byDate.get(date) || { date }, settings, date));
}

function formatDatePt(value) {
  const raw = cleanText(value);
  if (!raw) return "-";
  const date = new Date(`${raw}T00:00:00`);
  if (Number.isNaN(date.getTime())) return raw;
  return new Intl.DateTimeFormat("pt-PT", {
    timeZone: "Europe/Lisbon",
    weekday: "long",
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  }).format(date);
}

function orderDatesLabel(days = []) {
  return (Array.isArray(days) ? days : [])
    .map((item) => cleanText(item.date))
    .filter(Boolean)
    .join(", ");
}

function buildBakeryGeneratedText(order = {}, settings = DEFAULT_BAKERY_SETTINGS, submittedByName = "") {
  const days = Array.isArray(order.days) ? order.days : [];
  const lines = [
    `SUBJECT: Lisboa Central Hostel - Encomenda p\u00e3es e bolos para dias ${orderDatesLabel(days) || "-"}`,
    "",
    "Bom dia,",
    "",
    "Segue a encomenda de p\u00e3es e bolos:",
    "",
  ];
  days.forEach((day) => {
    lines.push(`${formatDatePt(day.date)}`);
    (Array.isArray(day.breadBreakdown) ? day.breadBreakdown : []).forEach((item) => {
      lines.push(`${item.name}: ${Number(item.quantity || 0)}`);
    });
    lines.push(`Past\u00e9is de nata: ${Number(day.pasteisDeNata || 0)}`);
    lines.push("");
  });
  lines.push("Cumprimentos,");
  lines.push(cleanText(submittedByName || order.submittedByName) || "[Name]");
  return lines.join("\n");
}

function sanitizeBakeryOrderRow(row = {}, settings = DEFAULT_BAKERY_SETTINGS) {
  const targetDates = Array.isArray(row.target_dates || row.targetDates)
    ? (row.target_dates || row.targetDates).map((item) => cleanText(item)).filter(Boolean)
    : ((Array.isArray(row.days) ? row.days : []).map((item) => cleanText(item?.date)).filter(Boolean));
  const effectiveTargetDates = targetDates.length ? targetDates : generateTargetDates(cleanText(row.order_date || row.orderDate) || formatLisbonIso(new Date()));
  const order = {
    id: cleanText(row.id),
    orderNumber: Number(row.order_number || row.orderNumber || 0) || 0,
    status: cleanText(row.status).toLowerCase() === "submitted" ? "submitted" : "open",
    orderDate: cleanText(row.order_date || row.orderDate) || formatLisbonIso(new Date()),
    createdAt: cleanText(row.created_at || row.createdAt),
    updatedAt: cleanText(row.updated_at || row.updatedAt),
    submittedAt: cleanText(row.submitted_at || row.submittedAt),
    submittedByName: cleanText(row.submitted_by_name || row.submittedByName),
    submittedByUserEmail: cleanText(row.submitted_by_user_email || row.submittedByUserEmail).toLowerCase(),
    targetDates: effectiveTargetDates,
    days: sanitizeBakeryDays(row.days, settings, effectiveTargetDates),
  };
  order.generatedText = cleanText(row.generated_text || row.generatedText) || buildBakeryGeneratedText(order, settings, order.submittedByName);
  return order;
}

module.exports = {
  DEFAULT_BAKERY_SETTINGS,
  buildBakeryGeneratedText,
  cleanText,
  formatDatePt,
  formatLisbonIso,
  generateTargetDates,
  isWorkingDayLisbon,
  normalizeBakeryDay,
  normalizeBase,
  normalizeRecipients,
  orderDatesLabel,
  sanitizeBakeryDays,
  sanitizeBakeryOrderRow,
  sanitizeBakerySettings,
};
