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
