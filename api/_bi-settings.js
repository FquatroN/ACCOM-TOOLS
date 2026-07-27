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

const DEFAULT_BI_FDM_ACCOUNT_CATEGORIES = [
  ["POSSales", "Sales", true],
  ["ReservationSales", "Sales", true],
  ["ReservationPayments", "Reservation", true],
  ["DepositReturn", "Deposit", false],
  ["RefundsAndReturns", "Reservation", true],
  ["DepositTaken", "Deposit", false],
  ["TransferInFromAccount", "Transfer", false],
  ["TransferOutToAccount", "Transfer", false],
  ["StaffWithdrawals", "Deposit", false],
  ["Compras", "Transfer", false],
  ["StaffDeposit", "Deposit", false],
  ["CancellationCharges", "Reservation", true],
  ["Tip", "Sales", true],
  ["Coffee Machine", "Sales", true],
  ["NoShowCharges", "Reservation", true],
].map(([category, macroCategory, isResult]) => ({ category, macroCategory, isResult }));

const DEFAULT_BI_SETTINGS = {
  saleItemCategories: DEFAULT_BI_SALE_ITEM_CATEGORIES,
  fdmAccountCategories: DEFAULT_BI_FDM_ACCOUNT_CATEGORIES,
};

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function sanitizeBiSettings(source = {}) {
  const settings = source && typeof source === "object" ? source : {};
  const rawSaleItemRows = Array.isArray(settings.saleItemCategories)
    ? settings.saleItemCategories
    : Array.isArray(settings.sale_item_categories)
      ? settings.sale_item_categories
      : DEFAULT_BI_SALE_ITEM_CATEGORIES;
  const seenSaleItems = new Set();
  const saleItemCategories = rawSaleItemRows
    .map((row) => {
      const saleItem = cleanText(row?.saleItem || row?.sale_item || row?.SALE_ITEM);
      const saleCategory = cleanText(row?.saleCategory || row?.sale_category || row?.SALE_CATEGORY);
      return { saleItem, saleCategory };
    })
    .filter((row) => row.saleItem || row.saleCategory)
    .filter((row) => {
      const key = row.saleItem.toLowerCase();
      if (!key) return true;
      if (seenSaleItems.has(key)) return false;
      seenSaleItems.add(key);
      return true;
    });
  const rawFdmAccountRows = Array.isArray(settings.fdmAccountCategories)
    ? settings.fdmAccountCategories
    : Array.isArray(settings.fdm_account_categories)
      ? settings.fdm_account_categories
      : DEFAULT_BI_FDM_ACCOUNT_CATEGORIES;
  const seenAccountCategories = new Set();
  const fdmAccountCategories = rawFdmAccountRows
    .map((row) => {
      const category = cleanText(row?.category || row?.Category || row?.CATEGORY);
      const macroCategory = cleanText(row?.macroCategory || row?.macro_category || row?.["Macro Category"] || row?.MACRO_CATEGORY);
      const rawIsResult = row?.isResult ?? row?.is_result ?? row?.IsResult ?? row?.IS_RESULT;
      const isResult = rawIsResult === true || ["yes", "true", "1", "y"].includes(cleanText(rawIsResult).toLowerCase());
      return { category, macroCategory, isResult };
    })
    .filter((row) => row.category || row.macroCategory)
    .filter((row) => {
      const key = row.category.toLowerCase();
      if (!key) return true;
      if (seenAccountCategories.has(key)) return false;
      seenAccountCategories.add(key);
      return true;
    });
  return {
    saleItemCategories: saleItemCategories.length ? saleItemCategories : clone(DEFAULT_BI_SALE_ITEM_CATEGORIES),
    fdmAccountCategories: fdmAccountCategories.length ? fdmAccountCategories : clone(DEFAULT_BI_FDM_ACCOUNT_CATEGORIES),
  };
}

function safeBiSettings(settings = DEFAULT_BI_SETTINGS) {
  return sanitizeBiSettings(settings);
}

module.exports = {
  BI_SETTINGS_KEY,
  DEFAULT_BI_SETTINGS,
  DEFAULT_BI_SALE_ITEM_CATEGORIES,
  DEFAULT_BI_FDM_ACCOUNT_CATEGORIES,
  sanitizeBiSettings,
  safeBiSettings,
};
