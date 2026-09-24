const { cleanText } = require("./_supabase");
const { buildSupplierInvoiceEmail, isSupplierEmailScheduleDue, previousCalendarMonth, validateAttachmentBundle } = require("./_financial-docs-supplier-emails");
const { claimSupplierEmailRun, completeSupplierEmailRun, listSupplierEmailSchedules, listSupplierPeriodDocuments } = require("./_financial-docs-supplier-email-service");
const { downloadDriveFile, loadFinancialDocsSettings, refreshDriveAccessToken } = require("./_financial-docs-service");

const SAFE_RAW_ATTACHMENT_BYTES = 28 * 1024 * 1024;
const ALLOWED_ATTACHMENT_TYPES = new Set(["application/pdf", "image/jpeg", "image/png", "text/csv", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]);

function safeError(error) {
  const message = cleanText(error?.message) || "Email delivery failed.";
  return message.slice(0, 500);
}

async function sendSupplierInvoiceEmail({ schedule, period, html, text, attachments, fetchImpl = fetch, env = process.env }) {
  const apiKey = cleanText(env.RESEND_API_KEY);
  const rawFrom = cleanText(env.EMAIL_FROM);
  if (!apiKey || !rawFrom) throw new Error("Resend email configuration is incomplete.");
  const from = /<[^>]+>/.test(rawFrom) ? rawFrom : `ACCOM Tools - LCH <${rawFrom}>`;
  let response;
  try {
    response = await fetchImpl("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
        "Idempotency-Key": `financial-document-supplier-email/${schedule.id}/${period.key}`,
      },
      body: JSON.stringify({ from, to: schedule.recipients, subject: schedule.subject, html, text, attachments }),
    });
  } catch (error) {
    const uncertain = new Error(safeError(error));
    uncertain.code = "TRANSPORT_UNCERTAIN";
    throw uncertain;
  }
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(cleanText(payload?.message || payload?.error) || `Resend rejected the email (${response.status}).`);
    error.code = "RESEND_REJECTED";
    throw error;
  }
  return payload;
}

function toSnapshot(document) {
  return {
    id: cleanText(document.id), documentDate: cleanText(document.document_date), supplierName: cleanText(document.supplier_name),
    description: cleanText(document.description), amount: Number(document.amount || 0), storedFilename: cleanText(document.stored_filename),
  };
}

async function processSupplierEmailSchedule({ schedule, now = new Date(), triggerType = "cron", dependencies = {} }) {
  const period = previousCalendarMonth(now, "Europe/Lisbon");
  const claim = dependencies.claimSupplierEmailRun || claimSupplierEmailRun;
  const complete = dependencies.completeSupplierEmailRun || completeSupplierEmailRun;
  const listDocuments = dependencies.listSupplierPeriodDocuments || listSupplierPeriodDocuments;
  const claimed = await claim({ scheduleId: schedule.id, periodStart: period.start, periodEnd: period.end, triggerType });
  if (!claimed?.claimed) return { status: "already_processed", run: claimed?.run || null };
  const runId = cleanText(claimed?.run?.id);
  const documents = await listDocuments({ supplierNif: schedule.supplierNif, periodStart: period.start, periodEnd: period.end });
  const snapshots = documents.map(toSnapshot);
  if (!documents.length) return complete({ runId, status: "skipped", documents: snapshots });
  const missing = documents.filter((document) => !cleanText(document.drive_file_id));
  if (missing.length) return complete({ runId, status: "failed", documents: snapshots, errorCode: "missing_attachment", errorMessage: `Missing attachment for ${missing.length} invoice document(s).`, retryable: true });
  try {
    const settings = await (dependencies.loadFinancialDocsSettings || loadFinancialDocsSettings)();
    const access = await (dependencies.refreshDriveAccessToken || refreshDriveAccessToken)(settings);
    const download = dependencies.downloadDriveFile || downloadDriveFile;
    const files = await Promise.all(documents.map(async (document) => {
      const downloaded = await download(access.accessToken, document.drive_file_id);
      const mimeType = cleanText(document.mime_type || downloaded.mimeType).split(";")[0].toLowerCase();
      if (!ALLOWED_ATTACHMENT_TYPES.has(mimeType)) throw new Error(`Unsupported attachment type: ${mimeType || "unknown"}.`);
      return { buffer: downloaded.buffer, filename: cleanText(document.stored_filename) || "invoice", contentType: mimeType };
    }));
    const bundle = validateAttachmentBundle(files, SAFE_RAW_ATTACHMENT_BYTES);
    const content = buildSupplierInvoiceEmail({ schedule, period, documents: snapshots });
    const sent = await sendSupplierInvoiceEmail({ schedule, period, html: content.html, text: content.text, attachments: files.map((file) => ({ filename: file.filename, content: file.buffer.toString("base64") })), fetchImpl: dependencies.fetchImpl || fetch, env: dependencies.env || process.env });
    return complete({ runId, status: "sent", documents: snapshots, attachmentCount: bundle.count, attachmentRawBytes: bundle.rawBytes, resendMessageId: cleanText(sent?.id) });
  } catch (error) {
    const uncertain = error?.code === "TRANSPORT_UNCERTAIN";
    return complete({ runId, status: uncertain ? "uncertain" : "failed", documents: snapshots, errorCode: uncertain ? "transport_uncertain" : cleanText(error?.code) || "delivery_failed", errorMessage: safeError(error), retryable: !uncertain });
  }
}

async function processDueSupplierEmailSchedules({ now = new Date(), dependencies = {} } = {}) {
  const schedules = await (dependencies.listSupplierEmailSchedules || listSupplierEmailSchedules)();
  const due = schedules.filter((schedule) => isSupplierEmailScheduleDue(schedule, now, "Europe/Lisbon").due);
  const process = dependencies.processSupplierEmailSchedule || processSupplierEmailSchedule;
  return Promise.all(due.map((schedule) => process({ schedule, now, dependencies })));
}

module.exports = { ALLOWED_ATTACHMENT_TYPES, SAFE_RAW_ATTACHMENT_BYTES, processDueSupplierEmailSchedules, processSupplierEmailSchedule, sendSupplierInvoiceEmail };
