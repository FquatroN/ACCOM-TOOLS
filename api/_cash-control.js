const { randomUUID } = require("node:crypto");
const cashDefaults = require("./_cash-control-defaults.json");
const { cleanText, normalizeDate, normalizeTime } = require("./_supabase");

const CASH_CONTROL_SETTING_KEY = "cash_control";
const CASH_DENOMINATIONS = Object.freeze([
  { key: "500", value: 500 },
  { key: "200", value: 200 },
  { key: "100", value: 100 },
  { key: "50", value: 50 },
  { key: "20", value: 20 },
  { key: "10", value: 10 },
  { key: "5", value: 5 },
  { key: "2", value: 2 },
  { key: "1", value: 1 },
  { key: "0.5", value: 0.5 },
  { key: "0.2", value: 0.2 },
  { key: "0.1", value: 0.1 },
  { key: "0.05", value: 0.05 },
  { key: "0.02", value: 0.02 },
  { key: "0.01", value: 0.01 },
]);
const CASH_MIN_ALERT_DENOMINATIONS = Object.freeze(["20", "10", "5", "2", "1", "0.5", "0.2", "0.1"]);

function normalizeCashStatus(value, fallback = "C") {
  const raw = cleanText(value).toUpperCase();
  if (raw === "O") return "O";
  if (raw === "C") return "C";
  return fallback === "O" ? "O" : "C";
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanText(value));
}

function isOpenCashStatus(value) {
  return normalizeCashStatus(value) === "O";
}

function slugify(value) {
  return cleanText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
}

function normalizeCashNumber(value) {
  const raw = cleanText(value).replace(/\s/g, "").replace(/€/g, "").replace(",", ".");
  if (!raw) return null;
  const numeric = Number(raw);
  return Number.isFinite(numeric) ? numeric : null;
}

function normalizeCashInteger(value, defaultValue = null) {
  const numeric = normalizeCashNumber(value);
  if (numeric == null) return defaultValue;
  return Math.max(0, Math.round(numeric));
}

function normalizeMoney(value) {
  const numeric = normalizeCashNumber(value);
  if (numeric == null) return 0;
  return Number(numeric.toFixed(2));
}

function normalizeShiftId(value) {
  return slugify(value || "shift");
}

function normalizeShift(input = {}, fallbackIndex = 0) {
  const name = cleanText(input.name) || `Shift ${fallbackIndex + 1}`;
  return {
    id: cleanText(input.id) || normalizeShiftId(name),
    name,
    startTime: normalizeTime(input.startTime || input.start_time) || "00:00",
  };
}

function normalizeItem(input = {}, fallbackIndex = 0) {
  const name = cleanText(input.name) || `Item ${fallbackIndex + 1}`;
  const defaultQuantity = normalizeCashInteger(input.defaultQuantity ?? input.default_quantity, 0);
  return {
    id: cleanText(input.id) || slugify(name || `item-${fallbackIndex + 1}`),
    name,
    defaultQuantity: defaultQuantity == null ? 0 : defaultQuantity,
  };
}

function sanitizeCashMinSettings(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  return CASH_MIN_ALERT_DENOMINATIONS.reduce((acc, key) => {
    acc[key] = normalizeCashInteger(source[key], 0) || 0;
    return acc;
  }, {});
}

function buildItemMaps(settings) {
  const byId = new Map();
  const byName = new Map();
  const bySlug = new Map();
  (settings?.items || []).forEach((item) => {
    byId.set(cleanText(item.id), item);
    byName.set(cleanText(item.name).toLowerCase(), item);
    bySlug.set(slugify(item.name), item);
  });
  return { byId, byName, bySlug };
}

function sanitizeCashControlSettings(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const shiftsSource = Array.isArray(source.shifts) ? source.shifts : cashDefaults.shifts;
  const itemsSource = Array.isArray(source.items) ? source.items : cashDefaults.items;
  const seenShiftIds = new Set();
  const shifts = shiftsSource
    .map((shift, index) => normalizeShift(shift, index))
    .filter((shift) => {
      const key = cleanText(shift.id).toLowerCase();
      if (!key || seenShiftIds.has(key)) return false;
      seenShiftIds.add(key);
      return true;
    });
  const seenItemIds = new Set();
  const items = itemsSource
    .map((item, index) => normalizeItem(item, index))
    .filter((item) => {
      const key = cleanText(item.id).toLowerCase();
      if (!key || seenItemIds.has(key)) return false;
      seenItemIds.add(key);
      return true;
    });
  return {
    shifts: shifts.length ? shifts : cashDefaults.shifts.map(normalizeShift),
    items: items.length ? items : cashDefaults.items.map(normalizeItem),
    minCash: sanitizeCashMinSettings(source.minCash || source.min_cash),
    managerAlertEmail: isValidEmail(source.managerAlertEmail || source.manager_alert_email)
      ? cleanText(source.managerAlertEmail || source.manager_alert_email).toLowerCase()
      : "",
  };
}

function sanitizeDenominations(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  return CASH_DENOMINATIONS.reduce((acc, denom) => {
    acc[denom.key] = normalizeCashInteger(source[denom.key], 0) || 0;
    return acc;
  }, {});
}

function sanitizeItemCounts(input = {}, settings) {
  const source = input && typeof input === "object" ? input : {};
  const maps = buildItemMaps(settings);
  const result = {};
  (settings?.items || []).forEach((item) => {
    const sourceKey = Object.keys(source).find((key) => slugify(key) === slugify(item.name));
    const raw = source[item.id] ?? source[item.name] ?? (sourceKey ? source[sourceKey] : undefined);
    result[item.id] = raw === "" || raw === null || raw === undefined ? null : normalizeCashInteger(raw, null);
  });
  return result;
}

function sanitizeItemJustifications(input = {}, settings) {
  const source = input && typeof input === "object" ? input : {};
  const result = {};
  (settings?.items || []).forEach((item) => {
    const sourceKey = Object.keys(source).find((key) => slugify(key) === slugify(item.name));
    result[item.id] = cleanText(source[item.id] ?? source[item.name] ?? (sourceKey ? source[sourceKey] : ""));
  });
  return result;
}

function sanitizeCashControlRecord(input = {}, settings, existing = {}) {
  const safeSettings = sanitizeCashControlSettings(settings);
  const day = normalizeDate(input.day ?? input.date ?? existing.day ?? existing.date);
  const shiftId = cleanText(input.shiftId ?? input.shift_id ?? existing.shiftId ?? existing.shift_id)
    || normalizeShiftId(input.shift ?? input.shiftName ?? existing.shift ?? existing.shiftName);
  const shift = safeSettings.shifts.find((item) => cleanText(item.id) === cleanText(shiftId))
    || safeSettings.shifts.find((item) => cleanText(item.name).toLowerCase() === cleanText(input.shift ?? input.shiftName ?? existing.shift ?? existing.shiftName).toLowerCase());
  const id = cleanText(input.id || existing.id) || randomUUID();
  return {
    id,
    day,
    shiftId: cleanText(shift?.id),
    shiftName: cleanText(shift?.name),
    status: normalizeCashStatus(input.status ?? existing.status, "C"),
    name: cleanText(input.name ?? existing.name),
    denominations: sanitizeDenominations(input.denominations ?? existing.denominations),
    cardPos: normalizeMoney(input.cardPos ?? input.card_pos ?? existing.cardPos ?? existing.card_pos),
    cashFdm: normalizeMoney(input.cashFdm ?? input.cash_fdm ?? existing.cashFdm ?? existing.cash_fdm),
    cardFdm: normalizeMoney(input.cardFdm ?? input.card_fdm ?? existing.cardFdm ?? existing.card_fdm),
    justification: cleanText(input.justification ?? existing.justification),
    itemCounts: sanitizeItemCounts(input.itemCounts ?? input.item_counts ?? existing.itemCounts ?? existing.item_counts, safeSettings),
    itemJustifications: sanitizeItemJustifications(input.itemJustifications ?? input.item_justifications ?? existing.itemJustifications ?? existing.item_justifications, safeSettings),
    createdAt: cleanText(input.createdAt || input.created_at || existing.createdAt || existing.created_at),
    updatedAt: cleanText(input.updatedAt || input.updated_at || existing.updatedAt || existing.updated_at),
  };
}

function shiftSequence(settings) {
  return sanitizeCashControlSettings(settings).shifts;
}

function shiftOrderMap(settings) {
  const map = new Map();
  shiftSequence(settings).forEach((shift, index) => map.set(cleanText(shift.id), index));
  return map;
}

function sortCashControlRecords(records, settings) {
  const order = shiftOrderMap(settings);
  return [...records].sort((a, b) => {
    const dayCompare = cleanText(a.day).localeCompare(cleanText(b.day));
    if (dayCompare !== 0) return dayCompare;
    return (order.get(cleanText(a.shiftId)) ?? 999) - (order.get(cleanText(b.shiftId)) ?? 999);
  });
}

function recordKey(record) {
  return `${cleanText(record.day)}::${cleanText(record.shiftId)}`;
}

function calculateCashTotal(denominations = {}) {
  const total = CASH_DENOMINATIONS.reduce((sum, denom) => {
    return sum + Number(denominations?.[denom.key] || 0) * denom.value;
  }, 0);
  return Number(total.toFixed(2));
}

function getPreviousShiftDescriptor(day, shiftId, settings) {
  const shifts = shiftSequence(settings);
  const index = shifts.findIndex((item) => cleanText(item.id) === cleanText(shiftId));
  if (index === -1) return null;
  if (index > 0) {
    return { day, shiftId: shifts[index - 1].id };
  }
  if (!day) return null;
  const date = new Date(`${day}T00:00:00`);
  if (Number.isNaN(date.getTime())) return null;
  date.setDate(date.getDate() - 1);
  return { day: date.toISOString().slice(0, 10), shiftId: shifts[shifts.length - 1].id };
}

function getNextShiftDescriptor(day, shiftId, settings) {
  const shifts = shiftSequence(settings);
  const index = shifts.findIndex((item) => cleanText(item.id) === cleanText(shiftId));
  if (index === -1) return null;
  if (index < shifts.length - 1) {
    return { day, shiftId: shifts[index + 1].id };
  }
  if (!day) return null;
  const date = new Date(`${day}T00:00:00`);
  if (Number.isNaN(date.getTime())) return null;
  date.setDate(date.getDate() + 1);
  return { day: date.toISOString().slice(0, 10), shiftId: shifts[0].id };
}

function buildComputedCashRows(records, settings) {
  const sorted = sortCashControlRecords(records, settings);
  const byKey = new Map(sorted.map((row) => [recordKey(row), row]));
  return sorted.map((row) => {
    const previousDescriptor = getPreviousShiftDescriptor(row.day, row.shiftId, settings);
    const previous = previousDescriptor ? byKey.get(`${previousDescriptor.day}::${previousDescriptor.shiftId}`) : null;
    const cashTotal = calculateCashTotal(row.denominations);
    const calculatedCash = previous ? Number((calculateCashTotal(previous.denominations) + Number(row.cashFdm || 0)).toFixed(2)) : null;
    const diffCash = calculatedCash == null ? null : Number((cashTotal - calculatedCash).toFixed(2));
    const diffCard = Number((Number(row.cardPos || 0) - Number(row.cardFdm || 0)).toFixed(2));
    const itemDiffs = (settings?.items || []).map((item) => {
      const counted = row.itemCounts?.[item.id];
      const diff = counted == null ? null : counted - Number(item.defaultQuantity || 0);
      return { itemId: item.id, diff, counted, defaultQuantity: Number(item.defaultQuantity || 0) };
    });
    return {
      ...row,
      cashTotal,
      calculatedCash,
      diffCash,
      diffCard,
      itemDiffs,
      hasItemDiffs: itemDiffs.some((item) => item.diff != null && item.diff !== 0),
    };
  });
}

function validateCashControlRecord(record, records, settings, { excludeId = "", isCreate = false } = {}) {
  if (!cleanText(record.day)) {
    const error = new Error("Day is required.");
    error.statusCode = 400;
    throw error;
  }
  if (!cleanText(record.shiftId)) {
    const error = new Error("Shift is required.");
    error.statusCode = 400;
    throw error;
  }
  if (!cleanText(record.name)) {
    const error = new Error("Name is required.");
    error.statusCode = 400;
    throw error;
  }
  const duplicate = records.find((row) => recordKey(row) === recordKey(record) && cleanText(row.id) !== cleanText(excludeId));
  if (duplicate) {
    const error = new Error(`A cash control record for ${record.day} ${record.shiftName || record.shiftId} already exists.`);
    error.statusCode = 400;
    throw error;
  }
  const existing = records.find((row) => cleanText(row.id) === cleanText(excludeId)) || null;
  const otherOpen = records.find((row) => isOpenCashStatus(row.status) && cleanText(row.id) !== cleanText(excludeId));
  if (isCreate && otherOpen) {
    const error = new Error("Close the current open shift before adding a new record.");
    error.statusCode = 400;
    throw error;
  }
  if (isOpenCashStatus(record.status) && otherOpen) {
    const error = new Error("Only one cash control shift can stay open at a time.");
    error.statusCode = 400;
    throw error;
  }
  if (existing && normalizeCashStatus(existing.status) === "C" && normalizeCashStatus(record.status) !== "C") {
    const error = new Error("A closed cash control shift cannot be reopened.");
    error.statusCode = 400;
    throw error;
  }
  if (isCreate) {
    const next = getNextExpectedCashRecord(records, settings);
    if (next.day && next.shiftId) {
      if (cleanText(record.day) !== cleanText(next.day) || cleanText(record.shiftId) !== cleanText(next.shiftId)) {
        const error = new Error(`The next record must be ${next.day} ${next.shiftName}.`);
        error.statusCode = 400;
        throw error;
      }
    }
  }
  if (isOpenCashStatus(record.status)) return;
  const computedRows = buildComputedCashRows(
    sortCashControlRecords(
      [...records.filter((row) => cleanText(row.id) !== cleanText(excludeId)), record],
      settings
    ),
    settings
  );
  const current = computedRows.find((row) => cleanText(row.id) === cleanText(record.id));
  if (!current) return;
  if (current.diffCash != null && current.diffCash !== 0 && !cleanText(record.justification)) {
    const error = new Error("Justification is required when Dif. Cash is not zero.");
    error.statusCode = 400;
    throw error;
  }
  if (current.diffCard !== 0 && !cleanText(record.justification)) {
    const error = new Error("Justification is required when Dif. Card is not zero.");
    error.statusCode = 400;
    throw error;
  }
  (settings?.items || []).forEach((item) => {
    const counted = record.itemCounts?.[item.id];
    if (counted == null) {
      const error = new Error(`Count is required for item ${item.name}.`);
      error.statusCode = 400;
      throw error;
    }
    if (counted !== Number(item.defaultQuantity || 0) && !cleanText(record.itemJustifications?.[item.id])) {
      const error = new Error(`Justification is required for item ${item.name} when the quantity is different.`);
      error.statusCode = 400;
      throw error;
    }
  });
}

function sanitizeCashControlRecords(input = [], settings = cashDefaults) {
  const safeSettings = sanitizeCashControlSettings(settings);
  const source = Array.isArray(input) ? input : [];
  const seen = new Set();
  const rows = source
    .map((row) => sanitizeCashControlRecord(row, safeSettings))
    .filter((row) => {
      const key = cleanText(row.id).toLowerCase();
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  return sortCashControlRecords(rows, safeSettings);
}

function getNextExpectedCashRecord(records = [], settings = cashDefaults) {
  const safeSettings = sanitizeCashControlSettings(settings);
  const sorted = sortCashControlRecords(records, safeSettings);
  const last = sorted.at(-1);
  if (!last) {
    const today = new Date().toISOString().slice(0, 10);
    return {
      day: today,
      shiftId: safeSettings.shifts[0].id,
      shiftName: safeSettings.shifts[0].name,
    };
  }
  const next = getNextShiftDescriptor(last.day, last.shiftId, safeSettings);
  const shift = safeSettings.shifts.find((item) => cleanText(item.id) === cleanText(next?.shiftId)) || safeSettings.shifts[0];
  return {
    day: next?.day || last.day,
    shiftId: cleanText(shift.id),
    shiftName: cleanText(shift.name),
  };
}

function sanitizeCashControlPayload(input = {}) {
  const settings = sanitizeCashControlSettings(input.settings || input.cashSettings || input);
  const rawRecords = Array.isArray(input.records || input.cashRecords)
    ? input.records || input.cashRecords
    : [];
  const records = sanitizeCashControlRecords(rawRecords, settings);
  return { settings, records };
}

module.exports = {
  CASH_CONTROL_SETTING_KEY,
  CASH_DENOMINATIONS,
  CASH_MIN_ALERT_DENOMINATIONS,
  calculateCashTotal,
  buildComputedCashRows,
  getNextExpectedCashRecord,
  isOpenCashStatus,
  normalizeCashStatus,
  normalizeCashInteger,
  normalizeCashNumber,
  normalizeMoney,
  sanitizeCashControlPayload,
  sanitizeCashControlRecord,
  sanitizeCashControlRecords,
  sanitizeCashControlSettings,
  sortCashControlRecords,
  validateCashControlRecord,
};
