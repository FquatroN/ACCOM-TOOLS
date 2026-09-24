const { cleanText, restQuery } = require("./_supabase");
const { normalizeSupplierEmailSchedule } = require("./_financial-docs-supplier-emails");
const { listFinancialDocumentEntities } = require("./_financial-docs-service");

const SCHEDULE_SELECT = "id,supplier_nif,supplier_name,recipients,day_of_month,subject,enabled,created_at,updated_at";
const RUN_SELECT = "id,schedule_id,period_start,period_end,attempt_number,trigger_type,status,attachment_count,attachment_raw_bytes,resend_message_id,error_code,error_message,retryable,started_at,completed_at,created_at";

function toClientSchedule(row = {}) {
  return {
    id: cleanText(row.id),
    supplierNif: cleanText(row.supplier_nif),
    supplierName: cleanText(row.supplier_name),
    recipients: Array.isArray(row.recipients) ? row.recipients.map(cleanText).filter(Boolean) : [],
    dayOfMonth: Number(row.day_of_month),
    subject: cleanText(row.subject),
    enabled: row.enabled !== false,
    createdAt: cleanText(row.created_at),
    updatedAt: cleanText(row.updated_at),
  };
}

function toClientRun(row = {}) {
  return {
    id: cleanText(row.id),
    scheduleId: cleanText(row.schedule_id),
    periodStart: cleanText(row.period_start),
    periodEnd: cleanText(row.period_end),
    attemptNumber: Number(row.attempt_number || 0),
    triggerType: cleanText(row.trigger_type),
    status: cleanText(row.status),
    attachmentCount: Number(row.attachment_count || 0),
    attachmentRawBytes: Number(row.attachment_raw_bytes || 0),
    resendMessageId: cleanText(row.resend_message_id),
    errorCode: cleanText(row.error_code),
    errorMessage: cleanText(row.error_message),
    retryable: row.retryable === true,
    startedAt: cleanText(row.started_at),
    completedAt: cleanText(row.completed_at),
    createdAt: cleanText(row.created_at),
  };
}

function buildSupplierPeriodDocumentsPath({ supplierNif, periodStart, periodEnd }) {
  const nif = cleanText(supplierNif).replace(/\D+/g, "");
  const start = cleanText(periodStart);
  const end = cleanText(periodEnd);
  if (!nif || !/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) {
    throw new Error("Supplier and invoice period are required.");
  }
  return `financial_documents?select=id,document_date,supplier_nif,supplier_name,description,amount,drive_file_id,stored_filename,mime_type,file_size&supplier_nif=eq.${encodeURIComponent(nif)}&document_date=gte.${encodeURIComponent(start)}&document_date=lte.${encodeURIComponent(end)}&order=document_date.asc,id.asc`;
}

async function listSupplierPeriodDocuments({ supplierNif, periodStart, periodEnd }) {
  const rows = await restQuery(buildSupplierPeriodDocumentsPath({ supplierNif, periodStart, periodEnd }), { method: "GET" });
  return Array.isArray(rows) ? rows : [];
}

async function listSupplierEmailRuns(scheduleId) {
  const id = cleanText(scheduleId);
  if (!id) return [];
  const rows = await restQuery(`financial_document_supplier_email_runs?select=${RUN_SELECT}&schedule_id=eq.${encodeURIComponent(id)}&order=created_at.desc&limit=50`, { method: "GET" });
  return (Array.isArray(rows) ? rows : []).map(toClientRun);
}

async function listSupplierEmailSchedules() {
  const rows = await restQuery(`financial_document_supplier_email_schedules?select=${SCHEDULE_SELECT}&order=supplier_name.asc`, { method: "GET" });
  const schedules = (Array.isArray(rows) ? rows : []).map(toClientSchedule);
  const withRuns = await Promise.all(schedules.map(async (schedule) => ({ ...schedule, runs: await listSupplierEmailRuns(schedule.id) })));
  return withRuns.map(({ runs, ...schedule }) => ({ ...schedule, lastRun: runs[0] || null }));
}

async function validateScheduleInput(input) {
  const entities = await listFinancialDocumentEntities();
  return normalizeSupplierEmailSchedule(input, entities);
}

function toStorageSchedule(schedule) {
  return {
    supplier_nif: schedule.supplierNif,
    supplier_name: schedule.supplierName,
    recipients: schedule.recipients,
    day_of_month: schedule.dayOfMonth,
    subject: schedule.subject,
    enabled: schedule.enabled,
  };
}

async function createSupplierEmailSchedule(input) {
  const schedule = await validateScheduleInput(input);
  const rows = await restQuery("financial_document_supplier_email_schedules", {
    method: "POST",
    body: [toStorageSchedule(schedule)],
    preferRepresentation: true,
  });
  return toClientSchedule(Array.isArray(rows) ? rows[0] : rows);
}

async function updateSupplierEmailSchedule(id, input) {
  const safeId = cleanText(id);
  if (!safeId) throw new Error("Schedule id is required.");
  const schedule = await validateScheduleInput(input);
  const rows = await restQuery(`financial_document_supplier_email_schedules?id=eq.${encodeURIComponent(safeId)}`, {
    method: "PATCH",
    body: toStorageSchedule(schedule),
    preferRepresentation: true,
  });
  if (!Array.isArray(rows) || !rows[0]) {
    const error = new Error("Supplier email schedule not found.");
    error.statusCode = 404;
    throw error;
  }
  return toClientSchedule(rows[0]);
}

async function deleteSupplierEmailSchedule(id) {
  const safeId = cleanText(id);
  if (!safeId) throw new Error("Schedule id is required.");
  await restQuery(`financial_document_supplier_email_schedules?id=eq.${encodeURIComponent(safeId)}`, { method: "DELETE" });
}

module.exports = {
  buildSupplierPeriodDocumentsPath,
  createSupplierEmailSchedule,
  deleteSupplierEmailSchedule,
  listSupplierEmailRuns,
  listSupplierEmailSchedules,
  listSupplierPeriodDocuments,
  toClientRun,
  toClientSchedule,
  updateSupplierEmailSchedule,
};
