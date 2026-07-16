const { cleanText } = require("./_supabase");

const BI_SETTINGS_KEY = "bi_settings";

const DEFAULT_BI_SALE_ITEM_CATEGORIES = [
  ["Beer", "Drinks"],
  ["Beer Mini", "Drinks"],
  ["Carristur - YellowBus", "Tours"],
  ["Coca-Cola", "Drinks"],
  ["EarPlug", "Sales"],
  ["Extra Accomodation", "Accomodation"],
  ["Extra Breakfast", "Sales"],
  ["Fanta", "Drinks"],
  ["Gota - Surf", "Tours"],
  ["Guarana", "Drinks"],
  ["Indemnização - Danos", "Sales"],
  ["Indemnização - Perda Chave", "Sales"],
  ["Laundry", "Sales"],
  ["Limpeza Extra", "Sales"],
  ["Luggage", "Sales"],
  ["Nespresso", "Sales"],
  ["Oceanário Ticket", "Tours"],
  ["Parking", "Sales"],
  ["Printings", "Sales"],
  ["PubCrawl - Discovery Lisbon", "Tours"],
  ["Red Bull", "Drinks"],
  ["Sangria Happy Hour", "Sales"],
  ["Shampoo/Shower Gel", "Sales"],
  ["Soap", "Sales"],
  ["Souvenir - Caixa 5 Portos", "Sales"],
  ["Souvenir - Copo Porto / Sagres", "Sales"],
  ["Souvenir - Copo Shot Lisboa", "Sales"],
  ["Souvenir - Iman", "Sales"],
  ["Souvenir - Pin", "Sales"],
  ["Souvenir - Postal", "Sales"],
  ["Souvenir -Other", "Sales"],
  ["Sprite/7UP", "Drinks"],
  ["Sumol", "Drinks"],
  ["TMT", "TMT"],
  ["Tooth Brush Set", "Sales"],
  ["Tour - Discovery - Boat Party", "Tours"],
  ["Tour - Sintra - Keep It Local", "Tours"],
  ["Towel", "Sales"],
  ["Transfer - SmartBus", "Tours"],
  ["Transfer - Uber", "Tours"],
  ["Water", "Drinks"],
  ["Wine - Large (75cl)", "Drinks"],
  ["Wine - Small (37cl)", "Drinks"],
].map(([saleItem, saleCategory]) => ({ saleItem, saleCategory }));

const DEFAULT_BI_SETTINGS = {
  saleItemCategories: DEFAULT_BI_SALE_ITEM_CATEGORIES,
};

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function sanitizeBiSettings(source = {}) {
  const settings = source && typeof source === "object" ? source : {};
  const rawRows = Array.isArray(settings.saleItemCategories)
    ? settings.saleItemCategories
    : Array.isArray(settings.sale_item_categories)
      ? settings.sale_item_categories
      : DEFAULT_BI_SALE_ITEM_CATEGORIES;
  const seen = new Set();
  const saleItemCategories = rawRows
    .map((row) => {
      const saleItem = cleanText(row?.saleItem || row?.sale_item || row?.SALE_ITEM);
      const saleCategory = cleanText(row?.saleCategory || row?.sale_category || row?.SALE_CATEGORY);
      return { saleItem, saleCategory };
    })
    .filter((row) => row.saleItem || row.saleCategory)
    .filter((row) => {
      const key = row.saleItem.toLowerCase();
      if (!key) return true;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  return {
    saleItemCategories: saleItemCategories.length ? saleItemCategories : clone(DEFAULT_BI_SALE_ITEM_CATEGORIES),
  };
}

function safeBiSettings(settings = DEFAULT_BI_SETTINGS) {
  return sanitizeBiSettings(settings);
}

module.exports = {
  BI_SETTINGS_KEY,
  DEFAULT_BI_SETTINGS,
  DEFAULT_BI_SALE_ITEM_CATEGORIES,
  sanitizeBiSettings,
  safeBiSettings,
};
