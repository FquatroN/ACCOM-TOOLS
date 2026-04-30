const shoppingDefaults = require("./_shopping-defaults.json");

const SHOPPING_CATEGORY_OPTIONS = ["Breakfast", "Cleaning", "Sales", "Activities", "Other", "Tapas", "Utensils"];
const SHOPPING_WEEKDAY_OPTIONS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"];
const SHOPPING_STORED_OPTIONS = Array.isArray(shoppingDefaults?.storedOptions)
  ? shoppingDefaults.storedOptions.map((value) => String(value || "").trim()).filter(Boolean)
  : [
      "20 (10) -Frigorificos",
      "11-Armario",
      "11-Escritorio",
      "20-Lavandaria",
      "20-Limpeza",
      "21-Comidas",
      "146-Arrecadacao",
    ];

const DEFAULT_SHOPPING_CATEGORY_COLORS = Object.freeze(
  SHOPPING_CATEGORY_OPTIONS.reduce((acc, category) => {
    const provided = String(shoppingDefaults?.categoryColors?.[category] || "").trim();
    acc[category] = /^#[0-9a-f]{6}$/i.test(provided) ? provided.toUpperCase() : "#F3E7DB";
    return acc;
  }, {})
);

function cleanText(value) {
  return String(value ?? "").trim();
}

function normalizeShoppingCategory(value) {
  const raw = cleanText(value).toLowerCase().replace(/\s+/g, " ");
  if (raw === "breakfast" || raw === "pequeno almoço" || raw === "pequeno almoco") return "Breakfast";
  if (raw === "cleaning" || raw === "limpeza") return "Cleaning";
  if (raw === "sales" || raw === "vendas") return "Sales";
  if (raw === "activities" || raw === "atividades" || raw === "actividades") return "Activities";
  if (raw === "other" || raw === "outros") return "Other";
  if (raw === "tapas" || raw.includes("ver só à terça-feira") || raw.includes("ver so a terca-feira")) return "Tapas";
  if (raw === "utensils" || raw === "utensilios" || raw === "utensílios") return "Utensils";
  return SHOPPING_CATEGORY_OPTIONS.includes(cleanText(value)) ? cleanText(value) : "Other";
}

function normalizeShoppingStored(value) {
  const raw = cleanText(value);
  if (!raw) return "";
  const exact = SHOPPING_STORED_OPTIONS.find((option) => option === raw);
  if (exact) return exact;
  const normalized = raw.toLowerCase();
  const match = SHOPPING_STORED_OPTIONS.find((option) => option.toLowerCase() === normalized);
  return match || raw;
}

function normalizeColor(value, fallback = "#F3E7DB") {
  const raw = cleanText(value).toUpperCase();
  return /^#[0-9A-F]{6}$/.test(raw) ? raw : fallback;
}

function parseShoppingBool(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  const raw = cleanText(value).toLowerCase();
  if (!raw) return false;
  return ["true", "1", "yes", "y", "sim", "on"].includes(raw);
}

function slugify(value) {
  return cleanText(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function shoppingItemKey(category, item) {
  return `${normalizeShoppingCategory(category)}::${cleanText(item).toLowerCase()}`;
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

function normalizeWeekdays(value) {
  const source = Array.isArray(value) ? value : [];
  const seen = new Set();
  return source
    .map((item) => cleanText(item).toLowerCase())
    .filter((item) => SHOPPING_WEEKDAY_OPTIONS.includes(item))
    .filter((item) => {
      if (seen.has(item)) return false;
      seen.add(item);
      return true;
    });
}

function sanitizeShoppingCategoryColors(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  return SHOPPING_CATEGORY_OPTIONS.reduce((acc, category) => {
    acc[category] = normalizeColor(source[category], DEFAULT_SHOPPING_CATEGORY_COLORS[category]);
    return acc;
  }, {});
}

function sanitizeShoppingItem(item = {}, fallbackIndex = 0) {
  const category = normalizeShoppingCategory(item.category);
  const label = cleanText(item.item || item.name);
  return {
    id: cleanText(item.id) || `shopping-item-${fallbackIndex + 1}-${slugify(label || `item-${fallbackIndex + 1}`)}`,
    category,
    item: label,
    supplier: cleanText(item.supplier || item.suppliers),
    stored: normalizeShoppingStored(item.stored),
    quantityRequired: parseShoppingBool(item.quantityRequired ?? item.quantity_required ?? item.mandatoryExistingQuantity ?? item.mandatory_existing_quantity),
  };
}

const DEFAULT_SHOPPING_ITEMS = (Array.isArray(shoppingDefaults?.items) ? shoppingDefaults.items : [])
  .map((item, index) => sanitizeShoppingItem(item, index))
  .filter((item) => item.item);

const DEFAULT_SHOPPING_SETTINGS = {
  mandatoryWeekdays: [],
  emailRecipients: [],
  categoryColors: { ...DEFAULT_SHOPPING_CATEGORY_COLORS },
  items: DEFAULT_SHOPPING_ITEMS,
};

function sanitizeShoppingSettings(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const sourceItems = Array.isArray(source.items) ? source.items : [];
  const sourceHasStored = sourceItems.some((item) => cleanText(item?.stored));
  const sourceCategoryColors = source.categoryColors || source.category_colors;
  const sourceHasCategoryColors =
    sourceCategoryColors && typeof sourceCategoryColors === "object" && Object.keys(sourceCategoryColors).length > 0;
  const defaultsByKey = new Map(DEFAULT_SHOPPING_ITEMS.map((item) => [shoppingItemKey(item.category, item.item), item]));
  const seen = new Set();

  let items = [];
  if (sourceItems.length && sourceHasStored) {
    items = sourceItems
      .map((item, index) => {
        const base = defaultsByKey.get(shoppingItemKey(item?.category, item?.item)) || {};
        return sanitizeShoppingItem({ ...base, ...item, id: cleanText(item?.id) || cleanText(base.id) }, index);
      })
      .filter((item) => item.item)
      .filter((item) => {
        const key = cleanText(item.id) || shoppingItemKey(item.category, item.item);
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    DEFAULT_SHOPPING_ITEMS.forEach((defaultItem, index) => {
      const key = shoppingItemKey(defaultItem.category, defaultItem.item);
      if (items.some((item) => shoppingItemKey(item.category, item.item) === key)) return;
      items.push(sanitizeShoppingItem(defaultItem, sourceItems.length + index));
    });
  } else {
    items = DEFAULT_SHOPPING_ITEMS.map((item, index) => sanitizeShoppingItem(item, index));
  }

  return {
    mandatoryWeekdays: normalizeWeekdays(source.mandatoryWeekdays || source.mandatory_weekdays),
    emailRecipients: normalizeRecipients(source.emailRecipients || source.email_recipients),
    categoryColors: sourceHasCategoryColors
      ? sanitizeShoppingCategoryColors(sourceCategoryColors)
      : { ...DEFAULT_SHOPPING_CATEGORY_COLORS },
    items,
  };
}

function buildOrderItemsFromSettings(settings) {
  const normalized = sanitizeShoppingSettings(settings);
  return normalized.items.map((item) => ({
    id: item.id,
    category: item.category,
    item: item.item,
    supplier: item.supplier,
    stored: item.stored,
    quantityRequired: !!item.quantityRequired,
    existingQuantity: "",
    order: false,
  }));
}

function sanitizeOrderItems(value, settingsItems = []) {
  const source = Array.isArray(value) ? value : [];
  const configMap = new Map((Array.isArray(settingsItems) ? settingsItems : []).map((item) => [cleanText(item.id), item]));
  return source
    .map((item, index) => {
      const id = cleanText(item?.id);
      const config = configMap.get(id) || {};
      const hasConfigQuantityRequired = Object.prototype.hasOwnProperty.call(config, "quantityRequired");
      return {
        id: id || cleanText(config.id) || `shopping-order-item-${index + 1}`,
        category: normalizeShoppingCategory(item?.category || config.category),
        item: cleanText(item?.item || config.item),
        supplier: cleanText(item?.supplier || config.supplier),
        stored: normalizeShoppingStored(item?.stored || config.stored),
        quantityRequired: hasConfigQuantityRequired
          ? parseShoppingBool(config.quantityRequired)
          : parseShoppingBool(item?.quantityRequired ?? item?.quantity_required),
        existingQuantity: cleanText(item?.existingQuantity ?? item?.existing_quantity),
        order: parseShoppingBool(item?.order),
      };
    })
    .filter((item) => item.item);
}

function countOrderedItems(items) {
  return (Array.isArray(items) ? items : []).filter((item) => !!item?.order).length;
}

module.exports = {
  DEFAULT_SHOPPING_CATEGORY_COLORS,
  DEFAULT_SHOPPING_ITEMS,
  DEFAULT_SHOPPING_SETTINGS,
  SHOPPING_CATEGORY_OPTIONS,
  SHOPPING_STORED_OPTIONS,
  SHOPPING_WEEKDAY_OPTIONS,
  buildOrderItemsFromSettings,
  cleanText,
  countOrderedItems,
  normalizeRecipients,
  normalizeShoppingCategory,
  normalizeShoppingStored,
  normalizeWeekdays,
  sanitizeOrderItems,
  sanitizeShoppingCategoryColors,
  sanitizeShoppingItem,
  sanitizeShoppingSettings,
};
