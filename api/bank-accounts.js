const { cleanText, requireFeature, restQuery, sendError } = require("./_supabase");

const MAX_ROWS = 500;

const SOURCES = Object.freeze({
  "cgd-extrato": {
    table: "import_cgd_extrato_ordem",
    reconciliationSourceType: "import_cgd_extrato_ordem",
    dateColumn: "data",
    descriptionColumn: "descritivo",
    select: "id,data,data_valor,descritivo,montante,saldo",
  },
  "cartao-credito": {
    table: "import_cgd_cartao_credito",
    reconciliationSourceType: "import_cgd_cartao_credito",
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

function chunks(values, size = 100) {
  const result = [];
  for (let index = 0; index < values.length; index += size) result.push(values.slice(index, index + size));
  return result;
}

async function enrichWithReconciliation(rows, source) {
  const config = SOURCES[source];
  const ids = [...new Set((Array.isArray(rows) ? rows : []).map((row) => cleanText(row?.id)).filter(Boolean))];
  if (!ids.length) return Array.isArray(rows) ? rows : [];

  const itemRows = [];
  for (const batch of chunks(ids)) {
    const encodedIds = batch.map((id) => encodeURIComponent(id)).join(",");
    const result = await restQuery(
      `financial_reconciliation_items?select=source_id,reconciliation_id&source_type=eq.${config.reconciliationSourceType}&source_id=in.(${encodedIds})`,
      { method: "GET" },
    );
    if (Array.isArray(result)) itemRows.push(...result);
  }

  const reconciliationIds = [...new Set(itemRows
    .map((item) => cleanText(item?.reconciliation_id || item?.reconciliationId))
    .filter(Boolean))];
  const reconciliationRows = [];
  for (const batch of chunks(reconciliationIds)) {
    const encodedIds = batch.map((id) => encodeURIComponent(id)).join(",");
    const result = await restQuery(
      `financial_reconciliations?select=id,status,difference_amount&deleted_at=is.null&id=in.(${encodedIds})`,
      { method: "GET" },
    );
    if (Array.isArray(result)) reconciliationRows.push(...result);
  }

  const reconciliationById = new Map(reconciliationRows.map((row) => [cleanText(row?.id), row]));
  const reconciliationBySourceId = new Map();
  itemRows.forEach((item) => {
    const reconciliation = reconciliationById.get(cleanText(item?.reconciliation_id || item?.reconciliationId));
    const sourceId = cleanText(item?.source_id || item?.sourceId);
    if (sourceId && reconciliation) reconciliationBySourceId.set(sourceId, reconciliation);
  });

  return rows.map((row) => {
    const reconciliation = reconciliationBySourceId.get(cleanText(row?.id));
    if (!reconciliation) return row;
    return {
      ...row,
      reconciliationId: cleanText(reconciliation.id),
      reconciliationStatus: cleanText(reconciliation.status),
      reconciliationDifferenceAmount: reconciliation.difference_amount ?? reconciliation.differenceAmount,
    };
  });
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
    const result = await enrichWithReconciliation(Array.isArray(rows) ? rows : [], source);
    res.status(200).json({ source, rows: result, truncated: result.length >= MAX_ROWS });
  } catch (error) {
    sendError(res, error);
  }
};
