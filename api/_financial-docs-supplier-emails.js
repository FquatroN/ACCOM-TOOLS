function cleanText(value) {
  return String(value ?? "").trim();
}

function normalizeNif(value) {
  return cleanText(value).replace(/\D+/g, "");
}

function normalizeName(value) {
  return cleanText(value).replace(/\s+/g, " ").toLowerCase();
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanText(value));
}

function normalizeRecipients(value) {
  const raw = Array.isArray(value) ? value.join(",") : String(value ?? "");
  const seen = new Set();
  return raw
    .split(/[\n,;]/)
    .map((item) => cleanText(item).toLowerCase())
    .filter((item) => isValidEmail(item))
    .filter((item) => {
      if (seen.has(item)) return false;
      seen.add(item);
      return true;
    });
}

function normalizeSupplierEmailSchedule(input = {}, entities = []) {
  const supplierNif = normalizeNif(input.supplierNif ?? input.supplier_nif);
  const supplierName = cleanText(input.supplierName ?? input.supplier_name).replace(/\s+/g, " ");
  const entity = (Array.isArray(entities) ? entities : []).find(
    (item) => normalizeNif(item?.nif) === supplierNif && normalizeName(item?.name) === normalizeName(supplierName)
  );
  if (!supplierNif || !supplierName || !entity) throw new Error("Select an existing supplier.");
  const recipients = normalizeRecipients(input.recipients);
  if (!recipients.length) throw new Error("At least one valid email is required.");
  const dayOfMonth = Number(input.dayOfMonth ?? input.day_of_month);
  if (!Number.isInteger(dayOfMonth) || dayOfMonth < 1 || dayOfMonth > 31) {
    throw new Error("Day of the month must be between 1 and 31.");
  }
  const subject = cleanText(input.subject);
  if (!subject) throw new Error("Subject is required.");
  return {
    supplierNif: normalizeNif(entity.nif),
    supplierName: cleanText(entity.name).replace(/\s+/g, " "),
    recipients,
    dayOfMonth,
    subject,
    enabled: input.enabled !== false && input.enabled !== "false",
  };
}

function getClockParts(date, timeZone) {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hourCycle: "h23",
  }).formatToParts(date);
  const values = Object.fromEntries(parts.filter((part) => part.type !== "literal").map((part) => [part.type, part.value]));
  return {
    year: Number(values.year),
    month: Number(values.month),
    day: Number(values.day),
    hour: Number(values.hour),
    minute: Number(values.minute),
  };
}

function isoDate(date) {
  return date.toISOString().slice(0, 10);
}

function previousCalendarMonth(now = new Date(), timeZone = "Europe/Lisbon") {
  const { year, month } = getClockParts(now, timeZone);
  const end = new Date(Date.UTC(year, month - 1, 0));
  const start = new Date(Date.UTC(end.getUTCFullYear(), end.getUTCMonth(), 1));
  return { start: isoDate(start), end: isoDate(end), key: isoDate(start).slice(0, 7) };
}

function isSupplierEmailScheduleDue(schedule = {}, now = new Date(), timeZone = "Europe/Lisbon") {
  const clock = getClockParts(now, timeZone);
  const recipients = Array.isArray(schedule.recipients) ? schedule.recipients : [];
  const dayOfMonth = Number(schedule.dayOfMonth ?? schedule.day_of_month);
  return {
    due: schedule.enabled !== false && recipients.length > 0 && clock.hour === 9 && clock.minute === 0 && clock.day === dayOfMonth,
    slotKey: `${clock.year}-${String(clock.month).padStart(2, "0")}-${String(clock.day).padStart(2, "0")}:09:00`,
  };
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function normalizedDocument(document = {}) {
  const amount = Number(document.amount);
  return {
    documentDate: cleanText(document.documentDate ?? document.document_date),
    supplierName: cleanText(document.supplierName ?? document.supplier_name),
    description: cleanText(document.description),
    amount: Number.isFinite(amount) ? amount : 0,
  };
}

function buildSupplierInvoiceEmail({ schedule = {}, period = {}, documents = [] } = {}) {
  const rows = (Array.isArray(documents) ? documents : []).map(normalizedDocument);
  const supplierName = cleanText(schedule.supplierName ?? schedule.supplier_name);
  const periodLabel = cleanText(period.key);
  const htmlRows = rows
    .map(
      (row) => `<tr><td>${escapeHtml(row.documentDate)}</td><td>${escapeHtml(row.supplierName)}</td><td>${escapeHtml(row.description)}</td><td>${escapeHtml(row.amount.toFixed(2))}</td></tr>`
    )
    .join("");
  const textRows = rows.map((row) => `${row.documentDate} | ${row.supplierName} | ${row.description} | ${row.amount.toFixed(2)}`);
  return {
    subject: cleanText(schedule.subject),
    html: `<!doctype html><html><body><p>Invoices for ${escapeHtml(supplierName)} — ${escapeHtml(periodLabel)}.</p><table><thead><tr><th>Invoice date</th><th>Supplier</th><th>Description</th><th>Amount</th></tr></thead><tbody>${htmlRows}</tbody></table></body></html>`,
    text: [`Invoices for ${supplierName} — ${periodLabel}.`, "", "Invoice date | Supplier | Description | Amount", ...textRows].join("\n"),
  };
}

function validateAttachmentBundle(files, rawByteLimit) {
  const attachments = Array.isArray(files) ? files : [];
  const limit = Number(rawByteLimit);
  if (!Number.isFinite(limit) || limit <= 0) throw new Error("Attachment size limit is invalid.");
  let rawBytes = 0;
  attachments.forEach((file) => {
    if (!Buffer.isBuffer(file?.buffer)) throw new Error("Attachment content is missing.");
    rawBytes += file.buffer.length;
  });
  if (rawBytes > limit) throw new Error("Attachment bundle exceeds the safe email size limit.");
  return { count: attachments.length, rawBytes };
}

module.exports = {
  buildSupplierInvoiceEmail,
  escapeHtml,
  getClockParts,
  isSupplierEmailScheduleDue,
  normalizeRecipients,
  normalizeSupplierEmailSchedule,
  previousCalendarMonth,
  validateAttachmentBundle,
};
