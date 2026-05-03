const { cleanText, normalizeBool, normalizeDate, normalizeNumeric } = require("./_supabase");

const LAUNDRY_SETTING_KEY = "laundry_control";
const LAUNDRY_PROPERTY_OPTIONS = ["Hostel", "Cruz"];

const DEFAULT_LAUNDRY_SETTINGS = {
  pricePerKg: 0,
  emailRecipients: [],
  itemTypes: [
    { id: "single-baixo", name: "single baixo", weightKg: 0.48 },
    { id: "single-cima", name: "single cima", weightKg: 0.5 },
    { id: "casal-baixo", name: "casal baixo", weightKg: 0.72 },
    { id: "casal-cima", name: "casal cima", weightKg: 0.75 },
  ],
};

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

function slugify(value) {
  return cleanText(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function normalizeLaundryProperty(value) {
  const raw = cleanText(value).toLowerCase();
  if (!raw) return "Hostel";
  if (raw.includes("hostel")) return "Hostel";
  if (raw.includes("apart")) return "Cruz";
  if (raw.includes("cruz")) return "Cruz";
  const exact = LAUNDRY_PROPERTY_OPTIONS.find((item) => item.toLowerCase() === raw);
  return exact || "Hostel";
}

function sanitizeLaundryItemType(item = {}, fallbackIndex = 0) {
  const name = cleanText(item.name || item.label || item.item);
  return {
    id: cleanText(item.id) || `laundry-item-${fallbackIndex + 1}-${slugify(name || `item-${fallbackIndex + 1}`)}`,
    name: name || `item ${fallbackIndex + 1}`,
    weightKg: Math.max(0, Number(normalizeNumeric(item.weightKg ?? item.weight_kg) || 0)),
  };
}

function sanitizeLaundrySettings(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const itemTypes = (Array.isArray(source.itemTypes || source.item_types) ? source.itemTypes || source.item_types : [])
    .map(sanitizeLaundryItemType)
    .filter((item, index, items) => item.name && items.findIndex((candidate) => cleanText(candidate.id) === cleanText(item.id)) === index);
  return {
    pricePerKg: Math.max(0, Number(normalizeNumeric(source.pricePerKg ?? source.price_per_kg) || 0)),
    emailRecipients: normalizeRecipients(source.emailRecipients || source.email_recipients),
    itemTypes: itemTypes.length ? itemTypes : DEFAULT_LAUNDRY_SETTINGS.itemTypes.map((item, index) => sanitizeLaundryItemType(item, index)),
  };
}

function sanitizeLaundryCountMap(value, itemTypes = DEFAULT_LAUNDRY_SETTINGS.itemTypes) {
  const source = value && typeof value === "object" && !Array.isArray(value) ? value : {};
  return itemTypes.reduce((acc, item) => {
    const raw = source[item.id];
    const numeric = Number.parseInt(String(raw ?? "").trim(), 10);
    acc[item.id] = Number.isFinite(numeric) && numeric > 0 ? numeric : 0;
    return acc;
  }, {});
}

function sanitizeLaundryRecord(input = {}, settings = DEFAULT_LAUNDRY_SETTINGS, existing = {}) {
  const safeSettings = sanitizeLaundrySettings(settings);
  const property = normalizeLaundryProperty(input.property ?? existing.property);
  const date = normalizeDate(input.date ?? existing.date);
  if (!date) {
    const error = new Error("Date is required.");
    error.statusCode = 400;
    throw error;
  }
  return {
    id: cleanText(input.id || existing.id),
    property,
    date,
    sentItems: sanitizeLaundryCountMap(input.sentItems ?? input.sent_items ?? existing.sentItems ?? existing.sent_items, safeSettings.itemTypes),
    receivedItems: sanitizeLaundryCountMap(input.receivedItems ?? input.received_items ?? existing.receivedItems ?? existing.received_items, safeSettings.itemTypes),
    receivedWeightKg: Math.max(0, Number(normalizeNumeric(input.receivedWeightKg ?? input.received_weight_kg ?? existing.receivedWeightKg ?? existing.received_weight_kg) || 0)),
    notes: cleanText(input.notes ?? existing.notes),
    createdAt: cleanText(input.createdAt || input.created_at || existing.createdAt || existing.created_at),
    updatedAt: cleanText(input.updatedAt || input.updated_at || existing.updatedAt || existing.updated_at),
  };
}

function sanitizeLaundryRecords(value, settings = DEFAULT_LAUNDRY_SETTINGS) {
  const source = Array.isArray(value) ? value : [];
  const safeSettings = sanitizeLaundrySettings(settings);
  const seen = new Set();
  return source
    .map((item) => sanitizeLaundryRecord(item, safeSettings))
    .filter((item) => {
      const key = `${item.property}::${item.date}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .sort((a, b) => {
      const dateCompare = cleanText(b.date).localeCompare(cleanText(a.date));
      if (dateCompare !== 0) return dateCompare;
      if (a.property === b.property) return 0;
      if (a.property === "Hostel") return -1;
      if (b.property === "Hostel") return 1;
      return cleanText(a.property).localeCompare(cleanText(b.property));
    });
}

function sanitizeLaundryPayload(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const settings = sanitizeLaundrySettings(source.settings || source.laundrySettings || source);
  const records = sanitizeLaundryRecords(source.records || source.laundryRecords, settings);
  return { settings, records };
}

function countMapWeightKg(counts, itemTypes = DEFAULT_LAUNDRY_SETTINGS.itemTypes) {
  const safeCounts = sanitizeLaundryCountMap(counts, itemTypes);
  return Number(
    itemTypes.reduce((sum, item) => sum + (Number(safeCounts[item.id] || 0) * Number(item.weightKg || 0)), 0).toFixed(2)
  );
}

module.exports = {
  DEFAULT_LAUNDRY_SETTINGS,
  LAUNDRY_PROPERTY_OPTIONS,
  LAUNDRY_SETTING_KEY,
  cleanText,
  countMapWeightKg,
  normalizeLaundryProperty,
  normalizeRecipients,
  sanitizeLaundryCountMap,
  sanitizeLaundryItemType,
  sanitizeLaundryPayload,
  sanitizeLaundryRecord,
  sanitizeLaundryRecords,
  sanitizeLaundrySettings,
};
