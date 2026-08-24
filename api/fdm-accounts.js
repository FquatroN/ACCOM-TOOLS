const { cleanText, requireFeature, restQuery, sendError } = require("./_supabase");

const MAX_ROWS = 500;
const OPTIONS_LIMIT = 1000;
const TABLE = "import_fdm_accounts";
const SELECT = "id,event_date,event_time,account,category,amount,reservation_id,guest,description,reporting_date,user_name,bill_number,item,invoice_number,currency,invoice_amount,designation,invoice";

function normalizeDate(value) {
  const date = cleanText(value);
  return /^\d{4}-\d{2}-\d{2}$/.test(date) ? date : "";
}

function normalizeSort(value) {
  return cleanText(value).toLowerCase() === "asc" ? "asc" : "desc";
}

function normalizeAmount(value) {
  const text = cleanText(value).replace(",", ".");
  if (!/^[+-]?(?:\d+|\d*\.\d+)$/.test(text)) return "";
  const amount = Number(text);
  return Number.isFinite(amount) ? amount : "";
}

function normalizeTextFilter(value) {
  return cleanText(value).slice(0, 160).replace(/[,%()]/g, " ").trim();
}

function buildQuery(query = {}) {
  const sort = normalizeSort(query.sort);
  const params = [
    `select=${SELECT}`,
    `order=event_date.${sort},event_time.${sort},id.${sort}`,
    `limit=${MAX_ROWS}`,
  ];
  const dateFrom = normalizeDate(query.date_from);
  const dateTo = normalizeDate(query.date_to);
  if (dateFrom) params.push(`event_date=gte.${encodeURIComponent(dateFrom)}`);
  if (dateTo) params.push(`event_date=lte.${encodeURIComponent(dateTo)}`);

  const description = normalizeTextFilter(query.description);
  if (description) {
    params.push(`or=${encodeURIComponent(`(description.ilike.*${description}*,guest.ilike.*${description}*)`)}`);
  }
  const reservationId = normalizeTextFilter(query.reservation_id);
  if (reservationId) params.push(`reservation_id=ilike.${encodeURIComponent(`*${reservationId}*`)}`);
  const account = normalizeTextFilter(query.account);
  if (account) params.push(`account=eq.${encodeURIComponent(account)}`);
  const category = normalizeTextFilter(query.category);
  if (category) params.push(`category=eq.${encodeURIComponent(category)}`);
  const amountFrom = normalizeAmount(query.amount_from);
  const amountTo = normalizeAmount(query.amount_to);
  if (amountFrom !== "") params.push(`amount=gte.${amountFrom}`);
  if (amountTo !== "") params.push(`amount=lte.${amountTo}`);
  return `${TABLE}?${params.join("&")}`;
}

function chunks(values, size = 100) {
  const result = [];
  for (let index = 0; index < values.length; index += size) result.push(values.slice(index, index + size));
  return result;
}

async function enrichWithReconciliation(rows) {
  const ids = [...new Set((Array.isArray(rows) ? rows : []).map((row) => cleanText(row?.id)).filter(Boolean))];
  if (!ids.length) return Array.isArray(rows) ? rows : [];

  const itemRows = [];
  for (const batch of chunks(ids)) {
    const encodedIds = batch.map((id) => encodeURIComponent(id)).join(",");
    const result = await restQuery(
      `financial_reconciliation_items?select=source_id,reconciliation_id&source_type=eq.import_fdm_accounts&source_id=in.(${encodedIds})`,
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

function uniqueValues(rows, field) {
  return [...new Set((Array.isArray(rows) ? rows : [])
    .map((row) => cleanText(row?.[field]))
    .filter(Boolean))]
    .sort((left, right) => left.localeCompare(right, undefined, { sensitivity: "base" }));
}

async function loadOptions() {
  const [accountRows, categoryRows] = await Promise.all([
    restQuery(`${TABLE}?select=account&account=not.is.null&order=account.asc&limit=${OPTIONS_LIMIT}`, { method: "GET" }),
    restQuery(`${TABLE}?select=category&category=not.is.null&order=category.asc&limit=${OPTIONS_LIMIT}`, { method: "GET" }),
  ]);
  return {
    accounts: uniqueValues(accountRows, "account"),
    categories: uniqueValues(categoryRows, "category"),
  };
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "fdm-accounts");
    if (req.method !== "GET") {
      res.setHeader("Allow", "GET");
      res.status(405).json({ error: "Method not allowed." });
      return;
    }
    const includeOptions = cleanText(req.query?.include_options) === "1";
    const [rows, options] = await Promise.all([
      restQuery(buildQuery(req.query), { method: "GET" }),
      includeOptions ? loadOptions() : Promise.resolve(null),
    ]);
    const result = await enrichWithReconciliation(Array.isArray(rows) ? rows : []);
    res.status(200).json({
      rows: result,
      ...(options || {}),
      truncated: result.length >= MAX_ROWS,
    });
  } catch (error) {
    sendError(res, error);
  }
};
