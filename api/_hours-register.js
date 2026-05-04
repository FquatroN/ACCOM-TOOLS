const { randomUUID } = require("node:crypto");
const { cleanText, normalizeDate, normalizeTime } = require("./_supabase");

const HOURS_REGISTER_SETTING_KEY = "hours_register";
const DEFAULT_HOURS_SETTINGS = {
  people: ["Fernanda Pereira"],
};

const INITIAL_HOURS_PERSON = "Fernanda Pereira";
const INITIAL_HOURS_LINES = `
2026-01-02|10:50|16:00
2026-01-05|11:00|17:40
2026-01-07|10:30|16:00
2026-01-10|10:55|16:10
2026-01-11|11:00|16:00
2026-01-12|11:20|16:20
2026-01-15|10:07|14:50
2026-01-16|09:00|13:30
2026-01-18|09:30|14:15
2026-01-19|10:00|14:00
2026-01-20|10:05|15:30
2026-01-21|10:40|14:30
2026-01-26|10:41|14:45
2026-01-30|11:00|15:30
2026-01-31|10:40|16:00
2026-02-01|11:00|17:40
2026-02-02|09:45|14:35
2026-02-05|10:46|14:45
2026-02-09|11:25|12:45
2026-02-10|11:00|14:20
2026-02-14|11:00|13:18
2026-02-15|11:00|14:10
2026-02-16|08:10|12:05
2026-02-17|09:45|15:04
2026-02-18|10:20|13:20
2026-02-20|11:00|16:25
2026-02-22|10:45|15:05
2026-02-24|11:00|16:10
2026-02-27|10:50|15:00
2026-02-28|10:50|15:25
2026-03-02|10:00|14:50
2026-03-05|09:45|14:15
2026-03-06|09:45|14:40
2026-03-07|09:45|13:05
2026-03-08|10:30|16:20
2026-03-11|10:00|15:10
2026-03-12|09:40|15:15
2026-03-14|11:00|14:00
2026-03-15|11:00|15:25
2026-03-18|10:25|15:57
2026-03-19|10:58|15:50
2026-03-21|11:00|15:32
2026-03-22|10:30|15:23
2026-03-24|10:30|14:40
2026-03-25|11:00|14:00
2026-03-31|11:00|15:00
2026-04-01|10:50|14:50
2026-04-02|11:00|16:10
2026-04-04|10:40|14:35
2026-04-05|09:08|12:43
2026-04-06|10:20|15:40
2026-04-09|10:40|14:59
2026-04-11|11:00|15:50
2026-04-12|11:00|15:00
2026-04-13|11:40|14:50
2026-04-15|11:00|13:40
2026-04-16|11:00|16:05
2026-04-21|10:58|15:51
2026-04-23|09:10|14:00
2026-04-24|10:00|15:00
2026-04-26|11:00|15:41
2026-04-27|11:10|13:31
2026-04-28|10:00|14:40
2026-04-30|09:30|13:37
2026-05-01|11:00|15:15
2026-05-03|10:45|15:36
2026-05-04|10:20|14:30
`.trim();

function normalizePersonList(value) {
  const source = Array.isArray(value) ? value : String(value || "").split(/[\n,;]/);
  const seen = new Set();
  return source
    .map((item) => cleanText(item))
    .filter(Boolean)
    .filter((item) => {
      const key = item.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function parseMinutes(value) {
  const safe = normalizeTime(value);
  if (!/^\d{2}:\d{2}$/.test(safe)) return null;
  const [hours, minutes] = safe.split(":").map(Number);
  if (!Number.isFinite(hours) || !Number.isFinite(minutes)) return null;
  return hours * 60 + minutes;
}

function calculateDurationMinutes(start, finish) {
  const startMinutes = parseMinutes(start);
  const finishMinutes = parseMinutes(finish);
  if (startMinutes == null || finishMinutes == null) return null;
  return finishMinutes - startMinutes;
}

function calculateDurationHours(start, finish) {
  const minutes = calculateDurationMinutes(start, finish);
  if (minutes == null || minutes <= 0) return 0;
  return Number((minutes / 60).toFixed(2));
}

function sanitizeHoursSettings(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const people = normalizePersonList(source.people || source.persons);
  return {
    people: people.length ? people : [...DEFAULT_HOURS_SETTINGS.people],
  };
}

function sanitizeHoursRecord(input = {}, settings = DEFAULT_HOURS_SETTINGS, existing = {}) {
  const safeSettings = sanitizeHoursSettings(settings);
  const person = cleanText(input.person ?? existing.person);
  const date = normalizeDate(input.date ?? existing.date);
  const start = normalizeTime(input.start ?? existing.start);
  const finish = normalizeTime(input.finish ?? existing.finish);
  if (!person) {
    const error = new Error("Person is required.");
    error.statusCode = 400;
    throw error;
  }
  if (!safeSettings.people.some((item) => cleanText(item).toLowerCase() === person.toLowerCase())) {
    const error = new Error("Person must exist in the configured list.");
    error.statusCode = 400;
    throw error;
  }
  if (!date) {
    const error = new Error("Date is required.");
    error.statusCode = 400;
    throw error;
  }
  if (!start || !finish) {
    const error = new Error("Start and finish time are required.");
    error.statusCode = 400;
    throw error;
  }
  const minutes = calculateDurationMinutes(start, finish);
  if (minutes == null || minutes <= 0) {
    const error = new Error("Finish time must be after start time.");
    error.statusCode = 400;
    throw error;
  }
  return {
    id: cleanText(input.id || existing.id) || randomUUID(),
    person,
    date,
    start,
    finish,
    createdAt: cleanText(input.createdAt || input.created_at || existing.createdAt || existing.created_at),
    updatedAt: cleanText(input.updatedAt || input.updated_at || existing.updatedAt || existing.updated_at),
  };
}

function sortHoursRecords(rows) {
  return [...rows].sort((a, b) => {
    const dateCompare = cleanText(b.date).localeCompare(cleanText(a.date));
    if (dateCompare !== 0) return dateCompare;
    const personCompare = cleanText(a.person).localeCompare(cleanText(b.person));
    if (personCompare !== 0) return personCompare;
    return cleanText(b.start).localeCompare(cleanText(a.start));
  });
}

function sanitizeHoursRecords(value, settings = DEFAULT_HOURS_SETTINGS) {
  const source = Array.isArray(value) ? value : [];
  const safeSettings = sanitizeHoursSettings(settings);
  const seen = new Set();
  const rows = source
    .map((item) => sanitizeHoursRecord(item, safeSettings))
    .filter((item) => {
      const key = cleanText(item.id);
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  return sortHoursRecords(rows);
}

function buildInitialHoursRecords() {
  return INITIAL_HOURS_LINES
    .split(/\r?\n/)
    .map((line, index) => line.trim())
    .filter(Boolean)
    .map((line, index) => {
      const [date, start, finish] = line.split("|");
      return {
        id: `hours-${date}-${String(start || "").replace(":", "")}-${String(finish || "").replace(":", "")}-${index + 1}`,
        person: INITIAL_HOURS_PERSON,
        date,
        start,
        finish,
      };
    });
}

const DEFAULT_HOURS_RECORDS = buildInitialHoursRecords();

function sanitizeHoursPayload(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const settings = sanitizeHoursSettings(source.settings || source.hoursSettings || source);
  const rawRecords = Array.isArray(source.records || source.hoursRecords)
    ? source.records || source.hoursRecords
    : DEFAULT_HOURS_RECORDS;
  const records = sanitizeHoursRecords(rawRecords, settings);
  return { settings, records };
}

module.exports = {
  HOURS_REGISTER_SETTING_KEY,
  DEFAULT_HOURS_SETTINGS,
  DEFAULT_HOURS_RECORDS,
  calculateDurationHours,
  calculateDurationMinutes,
  normalizePersonList,
  parseMinutes,
  sanitizeHoursPayload,
  sanitizeHoursRecord,
  sanitizeHoursRecords,
  sanitizeHoursSettings,
  sortHoursRecords,
};
