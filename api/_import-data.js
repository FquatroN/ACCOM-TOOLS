const { cleanText, normalizeDate, normalizeNumeric } = require("./_supabase");

const IMPORT_DATA_SETTINGS_KEY = "import_data";
const IMPORT_DATA_TYPES = [
  { key: "fdm-accounts", label: "FDM Accounts" },
  { key: "fdm-bookings", label: "FDM Bookings" },
];

const DEFAULT_IMPORT_DATA_SETTINGS = {
  types: [
    { type: "fdm-accounts", description: "Import FDM account movements from pasted text or uploaded tabular files." },
    { type: "fdm-bookings", description: "Import FDM reservation bookings from uploaded or pasted reservation exports." },
  ],
};

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function normalizeImportDataType(value) {
  const raw = cleanText(value).toLowerCase();
  return IMPORT_DATA_TYPES.find((item) => item.key === raw)?.key || IMPORT_DATA_TYPES[0].key;
}

function sanitizeImportDataSettings(source = {}) {
  const settings = source && typeof source === "object" ? source : {};
  const defaults = clone(DEFAULT_IMPORT_DATA_SETTINGS);
  const rawTypes = Array.isArray(settings.types) ? settings.types : defaults.types;
  const map = new Map(
    rawTypes.map((item) => [
      normalizeImportDataType(item?.type),
      {
        type: normalizeImportDataType(item?.type),
        description: cleanText(item?.description),
      },
    ])
  );
  return {
    types: IMPORT_DATA_TYPES.map((type) => ({
      type: type.key,
      description: cleanText(map.get(type.key)?.description) || cleanText(defaults.types.find((item) => item.type === type.key)?.description),
    })),
  };
}

function safeImportDataSettings(settings = DEFAULT_IMPORT_DATA_SETTINGS) {
  return sanitizeImportDataSettings(settings);
}

function normalizeImportMoney(value) {
  const raw = cleanText(value).replace(/\s+/g, "").replace(/[€$£]/g, "").replace(/\.(?=\d{3}(?:\D|$))/g, "").replace(",", ".");
  const numeric = normalizeNumeric(raw);
  return numeric === null || numeric === undefined || Number.isNaN(numeric) ? null : Number(Number(numeric).toFixed(2));
}

function normalizeImportTime(value) {
  const raw = cleanText(value);
  const match = raw.match(/(\d{1,2}):(\d{2})/);
  if (!match) return "";
  return `${String(match[1]).padStart(2, "0")}:${match[2]}`;
}

function normalizeImportLooseDate(value) {
  const normalized = normalizeDate(value);
  if (normalized) return normalized;
  const dottedMatch = cleanText(value).match(/^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})$/);
  if (dottedMatch) {
    const yearNum = Number.parseInt(dottedMatch[3], 10);
    if (Number.isFinite(yearNum)) {
      const year = dottedMatch[3].length === 2 ? 2000 + yearNum : yearNum;
      return `${year}-${String(Number.parseInt(dottedMatch[2], 10)).padStart(2, "0")}-${String(Number.parseInt(dottedMatch[1], 10)).padStart(2, "0")}`;
    }
  }
  const raw = cleanText(value).toLowerCase().replace(/\./g, "");
  const match = raw.match(/^(\d{1,2})\/([a-zç]+)\/(\d{2,4})$/i);
  if (!match) return "";
  const months = {
    jan: 1, january: 1, janeiro: 1,
    feb: 2, fev: 2, february: 2, fevereiro: 2,
    mar: 3, march: 3, marco: 3, março: 3,
    apr: 4, abr: 4, april: 4, abril: 4,
    may: 5, mai: 5, maio: 5,
    jun: 6, june: 6, junho: 6,
    jul: 7, july: 7, julho: 7,
    aug: 8, ago: 8, august: 8, agosto: 8,
    sep: 9, set: 9, september: 9, setembro: 9,
    oct: 10, out: 10, october: 10, outubro: 10,
    nov: 11, november: 11, novembro: 11,
    dec: 12, dez: 12, december: 12, dezembro: 12,
  };
  const month = months[match[2]];
  if (!month) return "";
  const yearNum = Number.parseInt(match[3], 10);
  if (!Number.isFinite(yearNum)) return "";
  const year = match[3].length === 2 ? 2000 + yearNum : yearNum;
  return `${year}-${String(month).padStart(2, "0")}-${String(Number.parseInt(match[1], 10)).padStart(2, "0")}`;
}

function normalizeImportDateTime(value) {
  const raw = cleanText(value);
  if (!raw) return { raw: "", eventDate: "", eventTime: "" };
  const match = raw.match(/^(.+?)\s+(\d{1,2}:\d{2})(?::\d{2})?$/);
  if (match) {
    return {
      raw,
      eventDate: normalizeImportLooseDate(match[1]),
      eventTime: normalizeImportTime(match[2]),
    };
  }
  return {
    raw,
    eventDate: normalizeImportLooseDate(raw),
    eventTime: "",
  };
}

function normalizeInvoiceFlag(value) {
  const raw = cleanText(value).toUpperCase();
  if (!raw) return null;
  if (["S", "Y", "YES", "TRUE", "1"].includes(raw)) return true;
  if (["N", "NO", "FALSE", "0"].includes(raw)) return false;
  return null;
}

function badRequest(message) {
  const error = new Error(message);
  error.statusCode = 400;
  return error;
}

function sanitizeFdmAccountsImportRow(input = {}, meta = {}) {
  const dateTime = normalizeImportDateTime(input.dateTimeRaw || input.date_time_raw || input.dateTime);
  const account = cleanText(input.account);
  const category = cleanText(input.category);
  const amount = normalizeImportMoney(input.amount);
  if (!account) throw badRequest("Account is required.");
  if (!dateTime.raw) throw badRequest("Date Time is required.");
  if (!category) throw badRequest("Category is required.");
  if (!Number.isFinite(amount)) throw badRequest("Amount is required.");
  return {
    import_batch: cleanText(meta.importBatch),
    source_type: "fdm-accounts",
    source_name: cleanText(meta.sourceName),
    source_row_number: Number(meta.sourceRowNumber) || 0,
    account,
    date_time_raw: dateTime.raw,
    event_date: dateTime.eventDate || null,
    event_time: dateTime.eventTime || null,
    category,
    amount,
    reservation_id: cleanText(input.reservationId || input.reservation_id),
    guest: cleanText(input.guest),
    reporting_date_raw: cleanText(input.reportingDateRaw || input.reporting_date_raw || input.reportingDate),
    reporting_date: normalizeImportLooseDate(input.reportingDateRaw || input.reporting_date_raw || input.reportingDate) || null,
    user_name: cleanText(input.userName || input.user_name || input.user),
    description: cleanText(input.description),
    bill_number: cleanText(input.billNumber || input.bill_number),
    item: cleanText(input.item),
    invoice_number: cleanText(input.invoiceNumber || input.invoice_number),
    currency: cleanText(input.currency) || "EUR",
    invoice_amount: cleanText(input.invoiceAmount || input.invoice_amount) === "" ? null : normalizeImportMoney(input.invoiceAmount || input.invoice_amount),
    designation: cleanText(input.designation),
    invoice: cleanText(input.invoice),
    invoice_flag: normalizeInvoiceFlag(input.invoice),
    raw_payload: input && typeof input === "object" ? input : {},
  };
}

function sanitizeFdmBookingsImportRow(input = {}, meta = {}) {
  const bookingNumber = cleanText(input.bookingNumber || input.booking_number || input.reservationId || input.reservation_id);
  const bookingDate = normalizeImportDateTime(input.bookingDate || input.booking_date || input.booking_date_raw);
  const arrivalRaw = cleanText(input.arrival || input.arrival_raw);
  const checkInRaw = cleanText(input.checkIn || input.check_in || input.check_in_raw);
  const checkOutRaw = cleanText(input.checkOut || input.check_out || input.check_out_raw);
  const balanceDueRaw = input.balanceDue ?? input.balance_due;
  const invoiceTotalRaw = input.invoiceTotal ?? input.invoice_total;

  if (!bookingNumber) throw badRequest("Booking Number is required.");

  return {
    import_batch: cleanText(meta.importBatch),
    source_type: "fdm-bookings",
    source_name: cleanText(meta.sourceName),
    source_row_number: Number(meta.sourceRowNumber) || 0,
    booking_number: bookingNumber,
    room_type: cleanText(input.roomType || input.room_type),
    room: cleanText(input.room),
    rate: cleanText(input.rate),
    guest_name: cleanText(input.name || input.guest_name || input.guest),
    arrival_raw: arrivalRaw,
    arrival_time: normalizeImportTime(arrivalRaw) || null,
    check_in_raw: checkInRaw,
    check_in_date: normalizeImportLooseDate(checkInRaw) || null,
    check_out_raw: checkOutRaw,
    check_out_date: normalizeImportLooseDate(checkOutRaw) || null,
    nights: cleanText(input.nights) === "" ? null : Number.parseInt(input.nights, 10),
    guests: cleanText(input.guests) === "" ? null : Number.parseInt(input.guests, 10),
    room_assigned: cleanText(input.isRoomAssigned || input.roomAssigned || input.room_assigned),
    status: cleanText(input.status),
    payment_status: cleanText(input.paymentStatus || input.payment_status),
    balance_due: cleanText(balanceDueRaw) === "" ? null : normalizeImportMoney(balanceDueRaw),
    channel: cleanText(input.channel),
    booking_date_raw: bookingDate.raw,
    booking_date: bookingDate.eventDate || null,
    booking_time: bookingDate.eventTime || null,
    country: cleanText(input.country),
    city: cleanText(input.city),
    invoice_total: cleanText(invoiceTotalRaw) === "" ? null : normalizeImportMoney(invoiceTotalRaw),
    currency: cleanText(input.currency) || "EUR",
    raw_payload: input && typeof input === "object" ? input : {},
  };
}

module.exports = {
  DEFAULT_IMPORT_DATA_SETTINGS,
  IMPORT_DATA_SETTINGS_KEY,
  IMPORT_DATA_TYPES,
  normalizeImportDataType,
  safeImportDataSettings,
  sanitizeFdmAccountsImportRow,
  sanitizeFdmBookingsImportRow,
  sanitizeImportDataSettings,
};
