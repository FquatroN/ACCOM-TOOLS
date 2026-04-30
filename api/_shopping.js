const SHOPPING_CATEGORY_OPTIONS = ["Breakfast", "Cleaning", "Sales", "Activities", "Other", "Tapas", "Utensils"];
const SHOPPING_WEEKDAY_OPTIONS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"];

const DEFAULT_SHOPPING_ITEM_ROWS = [
  ["Pequeno Almoço", "Cereais Chocopic", "Recheio"],
  ["Pequeno Almoço", "Cereais Muesli", "Recheio"],
  ["Pequeno Almoço", "Cereais CornFlakes", "Recheio"],
  ["Pequeno Almoço", "Chá Preto", "Recheio"],
  ["Pequeno Almoço", "Chá Tília", "Recheio"],
  ["Pequeno Almoço", "Chá Verde", "Recheio"],
  ["Pequeno Almoço", "Chá Descafeinado", "Recheio"],
  ["Pequeno Almoço", "Sumo Pacote individual (early breakfast)", "Recheio"],
  ["Pequeno Almoço", "Mel (Caixas)", "Recheio"],
  ["Pequeno Almoço", "Manteigas (caixas)", "Recheio"],
  ["Pequeno Almoço", "Café (frascos)", "Recheio"],
  ["Pequeno Almoço", "Madalenas", "Lidl/ Recheio"],
  ["Pequeno Almoço", "Pão de Forma", "Lidl / Recheio"],
  ["Pequeno Almoço", "Garrafas de Sumo Sunquick", "Recheio"],
  ["Pequeno Almoço", "Chocolate em Pó", "Recheio"],
  ["Pequeno Almoço", "Açúcar (saquetas)", "Recheio"],
  ["Pequeno Almoço", "Açúcar (kg)", "Recheio"],
  ["Pequeno Almoço", "Adoçante", "Recheio"],
  ["Pequeno Almoço", "Farinha (kg)", "Recheio"],
  ["Pequeno Almoço", "Bebida Vegetal (Soja)", "Recheio"],
  ["Pequeno Almoço", "Leite", "Recheio"],
  ["Pequeno Almoço", "Ovos", "Recheio"],
  ["Pequeno Almoço", "Doces morango (Caixas)", "Recheio"],
  ["Pequeno Almoço", "Doces Pessego (caixas)", "Recheio"],
  ["Pequeno Almoço", "Canela", "Recheio"],
  ["Pequeno Almoço", "Croissants (pacotes)", "Lidl/ Recheio"],
  ["Pequeno Almoço", "Creme barrar choco duo", "LIdl"],
  ["Pequeno Almoço", "Chocolate para barrar", "Lidl"],
  ["Pequeno Almoço", "Manteiga Amendoim", "Recheio"],
  ["Pequeno Almoço", "Fruta - Limões", "Recheio"],
  ["Pequeno Almoço", "Fruta - Maças ou Peras", "Recheio"],
  ["Pequeno Almoço", "Fruta - Bananas", "Recheio"],
  ["Pequeno Almoço", "Fruta - Laranjas", "Recheio"],
  ["Pequeno Almoço", "Queijo", "Recheio"],
  ["Pequeno Almoço", "Fiambre", "Recheio"],
  ["Pequeno Almoço", "Rolos de papel de Cozinha", "Renova"],
  ["Pequeno Almoço", "Papel Zig Zag", "Renova"],
  ["Pequeno Almoço", "Guardanapos (Pacotes)", "Renova"],
  ["Cleaning", "Papel Higienico (12x)", "Renova"],
  ["Cleaning", "Toalhetes apartamento (wipes)", "Recheio"],
  ["Cleaning", "Água Destilada", "Recheio"],
  ["Cleaning", "Lixívia", "Recheio"],
  ["Cleaning", "Panos cozinha lava-loiça (coloridos)", "Recheio"],
  ["Cleaning", "Esfregão inox", "Recheio"],
  ["Cleaning", "Esfregão (salva unhas)", "Recheio"],
  ["Cleaning", "Spray Lixivia", "Recheio"],
  ["Cleaning", "Spray Anti fungos", "Recheio / Leroy"],
  ["Cleaning", "Spray tira nodoas", "Recheio"],
  ["Cleaning", "Tira gorduras", "Recheio"],
  ["Cleaning", "Limpa Vidros", "Recheio"],
  ["Cleaning", "Gel Sanitario (wc pato)", "Recheio"],
  ["Cleaning", "Luvas (S/M)", "MateriaAtiva"],
  ["Cleaning", "Sabonete liquido (wcs)", "Renova"],
  ["Cleaning", "Det. Loiça Cozinha", "Recheio"],
  ["Cleaning", "Det. Máquina loiça em capsula (apt)", "Recheio"],
  ["Cleaning", "Det. Máquina Roupa Liquido (Hospedes e Apart)", "Recheio"],
  ["Cleaning", "Det. Máquina Roupa  Liquido", "Siali"],
  ["Cleaning", "Branqueador Roupa", "Siali"],
  ["Cleaning", "Det. Máquina Roupa capsulas  individuais (apt)", "Recheio"],
  ["Cleaning", "Amaciador Roupa", "Recheio"],
  ["Cleaning", "Saco Lixo Grande Cozinha e Apartamento \"Preto AD 80x90 (40rx10s)\"", "MateriaAtiva"],
  ["Cleaning", "Saco Lixo Brancos Caixotes Quartos e Lavandaria\"15 L 45x50 (35rx50s)\"", "MateriaAtiva"],
  ["Cleaning", "Sacos Lixo 10L WCS - Brancos", "Recheio"],
  ["Cleaning", "H40 - Lemon - Det. Chão", "MateriaAtiva"],
  ["Cleaning", "C90 - Cozinha - Det. Cozinha", "MateriaAtiva"],
  ["Cleaning", "H30 - Multiusos - Det. Multiusos", "MateriaAtiva"],
  ["Cleaning", "H150 - Det. Casas de Banho", "MateriaAtiva"],
  ["Cleaning", "AIRNOR 13 ECO - Ambientador", "MateriaAtiva"],
  ["Cleaning", "Higisol - Alcool Gel", "MateriaAtiva"],
  ["Cleaning", "Pau de madeira para esfregonas/vassouras", "Recheio"],
  ["Cleaning", "Cabeças de pá", "Chinês"],
  ["Cleaning", "Cabeças de vassoura", "Recheio"],
  ["Cleaning", "Cabeças de Esfregona", "Recheio"],
  ["Sales", "Água", "Recheio"],
  ["Sales", "Coca-Cola", "Recheio"],
  ["Sales", "Coca-cola zero", "Recheio"],
  ["Sales", "Guaraná", "Recheio"],
  ["Sales", "Sumol", "Recheio"],
  ["Sales", "Vinho Branco grande", "Recheio"],
  ["Sales", "Vinho Branco pequeno", "Recheio"],
  ["Sales", "Vinho Tinto grande", "Recheio"],
  ["Sales", "Vinho Tinto pequeno", "Recheio"],
  ["Sales", "7UP", "Recheio"],
  ["Sales", "Cerveja", "Recheio"],
  ["Sales", "Fanta", "Recheio"],
  ["Sales", "Café nespresso", "Nespresso"],
  ["Sales", "Chocolate  em pó maquina", "Nestle"],
  ["Sales", "Leite em pó maquina", "Nestle"],
  ["Sales", "Café em pó maquina", "Nestle"],
  ["Activities", "Gelados baun/choc/mor (caixa)", "Recheio"],
  ["Activities", "Ginjinha", "Recheio"],
  ["Activities", "Pipocas doces e salgadas", "Recheio"],
  ["Activities", "Chantilly", "Recheio"],
  ["Activities", "Copos de Papel grandes (Sangria)", "Alpha"],
  ["Activities", "Copos de papel pequenos (gelados)", "Alpha"],
  ["Activities", "Groselha", "Recheio"],
  ["Activities", "Vinho tinto para sangria (pacotes)", "Recheio"],
  ["Activities", "Sangria Refill", "Recheio"],
  ["Activities", "Gasosa para sangria", "Recheio"],
  ["Activities", "Sumo de laranja para sangria", "Recheio"],
  ["Other", "Azeite - garrafão", "Recheio"],
  ["Other", "Sal Grosso", "Recheio"],
  ["Other", "Feijão Branco", "Recheio"],
  ["Other", "Grão de Bico", "Recheio"],
  ["Other", "Abobora Congelada", "Recheio"],
  ["Other", "Cenoura Congelada", "Recheio"],
  ["Other", "Cogumelos", "Recheio"],
  ["Other", "Cebola", "Recheio"],
  ["Other", "Gengibre", "Recheio"],
  ["Other", "Fita-Cola", "Staples"],
  ["Other", "Rolos POS", "Recheio"],
  ["Other", "Ear Plugs", "Miguel"],
  ["Other", "escovas de dentes", "Lousani"],
  ["Other", "Shampoo/ sabonete pequeno venda hóspedes", "Lousani"],
  ["Other", "Agrafos", "Miguel"],
  ["Other", "Pilhas Médias AA", "Recheio/ Leroy"],
  ["Other", "Pilhas pequenas AAA", "Recheio/ Leroy"],
  ["TAPAS\n (ver só à terça-feira)", "Baguetes Normais (x4)", "LIDL\\Continente"],
  ["TAPAS\n (ver só à terça-feira)", "Baguetes Escuras (x4)", "LIDL\\Continente"],
  ["TAPAS\n (ver só à terça-feira)", "Batata Frita Pacote", "Recheio"],
  ["TAPAS\n (ver só à terça-feira)", "Paio", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Presunto", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Sardinha em Lata", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Tortilha de Batata", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Queijo Brie", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Queijo Fresco", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Tomate Cherry", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Manjericão", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Kiwi", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Melão", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Uvas", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Abacaxi", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Pão de Alho", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Queijo Barrar", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Geleia Morango", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Geleia Abobora", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Azeitonas", "Recheio/LIDL"],
  ["TAPAS\n (ver só à terça-feira)", "Abacate", "Recheio/LIDL"],
  ["Utensils", "Copos", "IKEA"],
  ["Utensils", "Canecas", "IKEA"],
  ["Utensils", "Facas", "tramontina"],
  ["Utensils", "Garfos", "tramontina"],
  ["Utensils", "Colheres de Sopa", "tramontina"],
  ["Utensils", "Colheres de café", "tramontina"],
  ["Utensils", "Pratos Grandes", "IKEA"],
  ["Utensils", "Taparueres ikea", "IKEA"],
  ["Utensils", "Pano de cozinha Hostel", "IKEA"],
  ["Utensils", "Pano de cozinha ikea APT", "IKEA"],
  ["Utensils", "Pratos Pequenos", "IKEA"],
  ["Utensils", "Palitos Tapas", "Recheio"],
  ["Utensils", "Saco papel Breakfast Box", "Recheio"],
];

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

function slugify(value) {
  return cleanText(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
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

function sanitizeShoppingItem(item = {}, fallbackIndex = 0) {
  const category = normalizeShoppingCategory(item.category);
  const label = cleanText(item.item || item.name);
  return {
    id: cleanText(item.id) || `shopping-item-${fallbackIndex + 1}-${slugify(label || `item-${fallbackIndex + 1}`)}`,
    category,
    item: label,
    supplier: cleanText(item.supplier || item.suppliers),
    quantityRequired: !!(item.quantityRequired ?? item.quantity_required ?? item.mandatoryExistingQuantity ?? item.mandatory_existing_quantity),
  };
}

const DEFAULT_SHOPPING_ITEMS = DEFAULT_SHOPPING_ITEM_ROWS
  .map(([category, item, supplier], index) =>
    sanitizeShoppingItem({ category, item, supplier, quantityRequired: true }, index)
  );

const DEFAULT_SHOPPING_SETTINGS = {
  mandatoryWeekdays: [],
  emailRecipients: [],
  items: DEFAULT_SHOPPING_ITEMS,
};

function sanitizeShoppingSettings(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const sourceItems = Array.isArray(source.items) ? source.items : [];
  const seen = new Set();
  const items = sourceItems
    .map((item, index) => sanitizeShoppingItem(item, index))
    .filter((item) => item.item)
    .filter((item) => {
      const key = cleanText(item.id) || `${item.category}::${item.item}`.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  return {
    mandatoryWeekdays: normalizeWeekdays(source.mandatoryWeekdays || source.mandatory_weekdays),
    emailRecipients: normalizeRecipients(source.emailRecipients || source.email_recipients),
    items: items.length ? items : DEFAULT_SHOPPING_ITEMS,
  };
}

function buildOrderItemsFromSettings(settings) {
  const normalized = sanitizeShoppingSettings(settings);
  return normalized.items.map((item) => ({
    id: item.id,
    category: item.category,
    item: item.item,
    supplier: item.supplier,
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
      return {
        id: id || cleanText(config.id) || `shopping-order-item-${index + 1}`,
        category: normalizeShoppingCategory(item?.category || config.category),
        item: cleanText(item?.item || config.item),
        supplier: cleanText(item?.supplier || config.supplier),
        quantityRequired: !!(item?.quantityRequired ?? item?.quantity_required ?? config.quantityRequired),
        existingQuantity: cleanText(item?.existingQuantity ?? item?.existing_quantity),
        order: !!item?.order,
      };
    })
    .filter((item) => item.item);
}

function countOrderedItems(items) {
  return (Array.isArray(items) ? items : []).filter((item) => !!item?.order).length;
}

module.exports = {
  DEFAULT_SHOPPING_ITEMS,
  DEFAULT_SHOPPING_SETTINGS,
  SHOPPING_CATEGORY_OPTIONS,
  SHOPPING_WEEKDAY_OPTIONS,
  buildOrderItemsFromSettings,
  cleanText,
  countOrderedItems,
  normalizeRecipients,
  normalizeShoppingCategory,
  normalizeWeekdays,
  sanitizeOrderItems,
  sanitizeShoppingItem,
  sanitizeShoppingSettings,
};
