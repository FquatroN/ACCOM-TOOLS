const assert = require("node:assert/strict");
const { readFileSync } = require("node:fs");
const test = require("node:test");

const {
  normalizeSupplierEmailSchedule,
  previousCalendarMonth,
  isSupplierEmailScheduleDue,
  buildSupplierInvoiceEmail,
  validateAttachmentBundle,
} = require("../api/_financial-docs-supplier-emails");
const { buildSupplierPeriodDocumentsPath } = require("../api/_financial-docs-supplier-email-service");
const { sendSupplierInvoiceEmail, sendSupplierInvoiceTestEmail } = require("../api/_financial-docs-supplier-email-automation");

test("normalizes one supplier schedule and rejects invalid values", () => {
  assert.deepEqual(
    normalizeSupplierEmailSchedule(
      {
        supplierNif: "PT 123",
        supplierName: "A & B",
        recipients: "A@EXAMPLE.COM; a@example.com\nb@example.com",
        dayOfMonth: "5",
        subject: "Invoices",
      },
      [{ nif: "123", name: "A & B" }]
    ),
    {
      supplierNif: "123",
      supplierName: "A & B",
      recipients: ["a@example.com", "b@example.com"],
      dayOfMonth: 5,
      subject: "Invoices",
      enabled: true,
    }
  );
  assert.throws(
    () => normalizeSupplierEmailSchedule({ supplierNif: "123", supplierName: "A & B", recipients: "a@example.com", dayOfMonth: 32, subject: "Invoices" }, [{ nif: "123", name: "A & B" }]),
    /day/i
  );
});

test("uses the complete prior calendar month and exact existing day in Lisbon", () => {
  assert.deepEqual(previousCalendarMonth(new Date("2026-03-05T09:00:00Z"), "Europe/Lisbon"), {
    start: "2026-02-01",
    end: "2026-02-28",
    key: "2026-02",
  });
  assert.equal(
    isSupplierEmailScheduleDue(
      { enabled: true, dayOfMonth: 31, recipients: ["a@example.com"] },
      new Date("2026-04-30T08:00:00Z"),
      "Europe/Lisbon"
    ).due,
    false
  );
});

test("escapes supplier invoice data in the email body and preserves the configured subject", () => {
  const email = buildSupplierInvoiceEmail({
    schedule: { supplierName: "A & B", subject: "September invoices" },
    period: { key: "2026-09" },
    documents: [{ documentDate: "2026-09-03", supplierName: "A & B", description: "<Invoice>", amount: 12.5 }],
  });
  assert.equal(email.subject, "September invoices");
  assert.match(email.html, /A &amp; B/);
  assert.match(email.html, /&lt;Invoice&gt;/);
  assert.match(email.text, /2026-09-03 \| A & B \| <Invoice> \| 12.50/);
});

test("rejects an attachment bundle before Base64 expansion exceeds its safe raw limit", () => {
  assert.deepEqual(validateAttachmentBundle([{ buffer: Buffer.from("abc"), filename: "one.pdf" }], 3), {
    count: 1,
    rawBytes: 3,
  });
  assert.throws(
    () => validateAttachmentBundle([{ buffer: Buffer.alloc(4), filename: "too-large.pdf" }], 3),
    /size/i
  );
});

test("migration provides atomic schedule-period claiming and no duplicate sent run", () => {
  const sql = readFileSync("supabase-migrations/2026-09-24-financial-document-supplier-email-schedules.sql", "utf8");
  assert.match(sql, /financial_document_supplier_email_schedules/);
  assert.match(sql, /claim_financial_document_supplier_email_run/);
  assert.match(sql, /unique/i);
});

test("supplier-period lookup uses supplier identity and document date only, not status", () => {
  const path = buildSupplierPeriodDocumentsPath({ supplierNif: "123", periodStart: "2026-09-01", periodEnd: "2026-09-30" });
  assert.match(path, /supplier_nif=eq\.123/);
  assert.match(path, /document_date=gte\.2026-09-01/);
  assert.match(path, /document_date=lte\.2026-09-30/);
  assert.doesNotMatch(path, /status=/);
});

test("sends the configured invoice email with stable Resend idempotency", async () => {
  let request;
  const result = await sendSupplierInvoiceEmail({
    schedule: { id: "schedule-1", recipients: ["a@example.com"], subject: "Invoices" },
    period: { key: "2026-09" },
    html: "<p>Invoices</p>",
    text: "Invoices",
    attachments: [{ filename: "invoice.pdf", content: Buffer.from("pdf").toString("base64") }],
    fetchImpl: async (url, options) => {
      request = { url, options };
      return { ok: true, json: async () => ({ id: "resend-1" }) };
    },
    env: { RESEND_API_KEY: "key", EMAIL_FROM: "from@example.com" },
  });
  assert.equal(result.id, "resend-1");
  assert.equal(request.options.headers["Idempotency-Key"], "financial-document-supplier-email/schedule-1/2026-09");
  assert.equal(JSON.parse(request.options.body).subject, "Invoices");
});

test("test invoice email sends the prior-month attachment pack only to its test recipient", async () => {
  let request;
  await sendSupplierInvoiceTestEmail({
    schedule: { id: "schedule-1", supplierNif: "123", supplierName: "Supplier", recipients: ["real@example.com"], subject: "Invoices" },
    recipient: "test@example.com",
    now: new Date("2026-10-05T08:00:00Z"),
    dependencies: {
      listSupplierPeriodDocuments: async () => [{ id: "doc-1", document_date: "2026-09-03", supplier_name: "Supplier", description: "Invoice", amount: 10, drive_file_id: "drive-1", stored_filename: "invoice.pdf", mime_type: "application/pdf" }],
      loadFinancialDocsSettings: async () => ({}),
      refreshDriveAccessToken: async () => ({ accessToken: "drive-token" }),
      downloadDriveFile: async () => ({ buffer: Buffer.from("pdf"), mimeType: "application/pdf" }),
      fetchImpl: async (_url, options) => { request = options; return { ok: true, json: async () => ({ id: "test-message" }) }; },
      env: { RESEND_API_KEY: "key", EMAIL_FROM: "from@example.com" },
    },
  });
  const payload = JSON.parse(request.body);
  assert.deepEqual(payload.to, ["test@example.com"]);
  assert.equal(payload.attachments.length, 1);
  assert.equal(request.headers["Idempotency-Key"], "financial-document-supplier-email/test/schedule-1/2026-09/test@example.com");
});

test("Financial Documents settings exposes the Supplier Emails configuration tab", () => {
  const html = readFileSync("index.html", "utf8");
  assert.match(html, /id="financial-docs-settings-supplier-emails-tab"/);
  assert.match(html, /id="financial-docs-settings-supplier-email-subject"/);
  assert.match(html, /id="financial-docs-settings-supplier-emails-body"/);
  assert.match(readFileSync("app-main.js", "utf8"), /data-supplier-email-test/);
});
