const { cleanText, requireFeature, restQuery, sendError } = require("./_supabase");

const MAX_ROWS = 500;

const SOURCES = Object.freeze({
  "cgd-extrato": {
    table: "import_cgd_extrato_ordem",
    dateColumn: "data",
    descriptionColumn: "descritivo",
    select: "id,data,data_valor,descritivo,montante,saldo",
  },
  "cartao-credito": {
    table: "import_cgd_cartao_credito",
    dateColumn: "data",
    descriptionColumn: "descricao",
    select: "id,data,data_valor,descricao,debito,credito,valor",
  },
});

function normalizeSource(value) {
  const source = cleanText(value).toLowerCase();
  return SOURCES[source] ? source : "cgd-extrato";
}

function normalizeDate(value) {
  const date = cleanText(value);
  return /^\d{4}-\d{2}-\d{2}$/.test(date) ? date : "";
}

function normalizeSort(value) {
  return cleanText(value).toLowerCase() === "asc" ? "asc" : "desc";
}

function buildQuery(source, query = {}) {
  const config = SOURCES[source];
  const params = [
    `select=${config.select}`,
    `order=${config.dateColumn}.${normalizeSort(query.sort)},id.${normalizeSort(query.sort)}`,
    `limit=${MAX_ROWS}`,
  ];
  const dateFrom = normalizeDate(query.date_from);
  const dateTo = normalizeDate(query.date_to);
  if (dateFrom) params.push(`${config.dateColumn}=gte.${encodeURIComponent(dateFrom)}`);
  if (dateTo) params.push(`${config.dateColumn}=lte.${encodeURIComponent(dateTo)}`);

  const description = cleanText(query.description).slice(0, 160).replace(/[,%]/g, " ");
  if (description) {
    params.push(`${config.descriptionColumn}=ilike.${encodeURIComponent(`%${description}%`)}`);
  }
  return `${config.table}?${params.join("&")}`;
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "bank-accounts");
    if (req.method !== "GET") {
      res.setHeader("Allow", "GET");
      res.status(405).json({ error: "Method not allowed." });
      return;
    }
    const source = normalizeSource(req.query?.source);
    const rows = await restQuery(buildQuery(source, req.query), { method: "GET" });
    const result = Array.isArray(rows) ? rows : [];
    res.status(200).json({ source, rows: result, truncated: result.length >= MAX_ROWS });
  } catch (error) {
    sendError(res, error);
  }
};
