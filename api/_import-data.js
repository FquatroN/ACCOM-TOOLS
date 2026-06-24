const crypto = require("crypto");

const { cleanText, normalizeDate, normalizeNumeric } = require("./_supabase");

const IMPORT_DATA_SETTINGS_KEY = "import_data";
const IMPORT_DATA_TYPES = [
  { key: "fdm-accounts", label: "FDM Accounts" },
  { key: "fdm-bookings", label: "FDM Bookings" },
  { key: "fdm-sales", label: "FDM Sales" },
  { key: "cgd-extrato-ordem", label: "CGD Extrato Ordem" },
  { key: "cgd-cartao-credito", label: "CGD Cartao Credito" },
];

const DEFAULT_IMPORT_DATA_SETTINGS = {
  types: [
    { type: "fdm-accounts", description: "Import FDM account movements from pasted text or uploaded tabular files." },
    { type: "fdm-bookings", description: "Import FDM reservation bookings from uploaded or pasted reservation exports." },
    { type: "fdm-sales", description: "Import FDM sales lines from uploaded or pasted sales report exports." },
    { type: "cgd-extrato-ordem", description: "Import CGD Conta Ordem bank statement movements from uploaded or pasted extracts." },
    { type: "cgd-cartao-credito", description: "Import CGD Cartao Credito statement movements from PDF, pasted text, or tabular files." },
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

function compactImportNumericText(value) {
  return cleanText(value)
    .replace(/[\s\u00a0\u1680\u180e\u2000-\u200d\u2028\u2029\u202f\u205f\u2060\u3000\ufeff]+/g, "")
    .replace(/[€$£]/g, "");
}

function normalizeImportNumericText(value) {
  const raw = compactImportNumericText(value).replace(/[^0-9,.\-+]/g, "");
  if (!raw) return "";
  const lastComma = raw.lastIndexOf(",");
  const lastDot = raw.lastIndexOf(".");
  const decimalIndex = Math.max(lastComma, lastDot);
  if (decimalIndex < 0) return raw.replace(/[,.]/g, "");
  const decimalSeparator = raw[decimalIndex];
  const decimalPart = raw.slice(decimalIndex + 1);
  const integerPart = raw.slice(0, decimalIndex).replace(/[,.]/g, "");
  if (decimalPart.length === 3 && !raw.slice(decimalIndex + 1).includes(decimalSeparator) && !/[,.]/.test(raw.slice(decimalIndex + 1))) {
    const singleSeparator = (lastComma >= 0 ? 1 : 0) + (lastDot >= 0 ? 1 : 0) === 1;
    if (singleSeparator) return raw.replace(/[,.]/g, "");
  }
  return `${integerPart}.${decimalPart}`;
}

function normalizeImportMoney(value) {
  const raw = normalizeImportNumericText(value);
  const numeric = normalizeNumeric(raw);
  return numeric === null || numeric === undefined || Number.isNaN(numeric) ? null : Number(Number(numeric).toFixed(2));
}

function normalizeImportDecimal(value) {
  const raw = normalizeImportNumericText(value);
  const numeric = normalizeNumeric(raw);
  return numeric === null || numeric === undefined || Number.isNaN(numeric) ? null : Number(numeric);
}

function normalizeImportTime(value) {
  const raw = cleanText(value);
  const match = raw.match(/(\d{1,2}):(\d{2})/);
  if (!match) return "";
  return `${String(match[1]).padStart(2, "0")}:${match[2]}`;
}

function normalizeImportLooseDate(value) {
  const compact = cleanText(value);
  const numericMatch = compact.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4})$/);
  if (numericMatch) {
    const yearNum = Number.parseInt(numericMatch[3], 10);
    if (Number.isFinite(yearNum)) {
      const year = numericMatch[3].length === 2 ? 2000 + yearNum : yearNum;
      return `${year}-${String(Number.parseInt(numericMatch[2], 10)).padStart(2, "0")}-${String(Number.parseInt(numericMatch[1], 10)).padStart(2, "0")}`;
    }
  }
  const normalized = normalizeDate(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(normalized)) return normalized;
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

function normalizeImportCgdDate(value) {
  const raw = cleanText(value);
  if (/^\d+(?:\.\d+)?$/.test(raw)) {
    const serial = Number(raw);
    if (Number.isFinite(serial) && serial > 20000 && serial < 80000) {
      const excelEpoch = Date.UTC(1899, 11, 30);
      const date = new Date(excelEpoch + Math.round(serial) * 86400000);
      return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
    }
  }
  return normalizeImportLooseDate(value);
}

function normalizeImportDateTime(value) {
  const raw = cleanText(value);
  if (!raw) return { raw: "", eventDate: "", eventTime: "" };
  const meridiemMatch = raw.match(/^(.+?)\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)$/i);
  if (meridiemMatch) {
    let hours = Number.parseInt(meridiemMatch[2], 10);
    const minutes = meridiemMatch[3];
    const suffix = meridiemMatch[5].toUpperCase();
    if (Number.isFinite(hours)) {
      if (suffix === "AM") hours = hours === 12 ? 0 : hours;
      if (suffix === "PM") hours = hours === 12 ? 12 : hours + 12;
      return {
        raw,
        eventDate: normalizeImportLooseDate(meridiemMatch[1]),
        eventTime: `${String(hours).padStart(2, "0")}:${minutes}`,
      };
    }
  }
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

function importDataRowKey(parts) {
  return crypto.createHash("sha256")
    .update((Array.isArray(parts) ? parts : []).map((part) => cleanText(part)).join("\u001F"))
    .digest("hex");
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
    check_in_raw: cleanText(checkInRaw),
    check_in_date: normalizeImportLooseDate(checkInRaw) || null,
    check_out_raw: cleanText(checkOutRaw),
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

function sanitizeFdmSalesImportRow(input = {}, meta = {}) {
  const reservationId = cleanText(input.reservationId || input.reservation_id || input.bookingNumber || input.booking_number);
  const saleDate = normalizeImportDateTime(input.dateRaw || input.date || input.saleDate || input.sale_date_raw);
  const saleItem = cleanText(input.saleItem || input.sale_item);
  const quantityValue = cleanText(input.quantity) === "" ? null : normalizeImportDecimal(input.quantity);
  if (!saleDate.raw) throw badRequest("Date is required.");
  if (!saleDate.eventDate) throw badRequest("Date is invalid.");
  if (!saleItem) throw badRequest("Sale Item is required.");
  if (!Number.isFinite(quantityValue)) throw badRequest("Quantity is required.");
  return {
    import_batch: cleanText(meta.importBatch),
    source_type: "fdm-sales",
    source_name: cleanText(meta.sourceName),
    source_row_number: Number(meta.sourceRowNumber) || 0,
    reservation_id: reservationId,
    sale_date_raw: saleDate.raw,
    sale_date: saleDate.eventDate,
    sale_time: saleDate.eventTime || "",
    sale_item: saleItem,
    quantity: quantityValue,
    price: cleanText(input.price) === "" ? null : normalizeImportMoney(input.price),
    net_price: cleanText(input.netPrice || input.net_price) === "" ? null : normalizeImportMoney(input.netPrice || input.net_price),
    tax: cleanText(input.tax) === "" ? null : normalizeImportMoney(input.tax),
    total: cleanText(input.total) === "" ? null : normalizeImportMoney(input.total),
    total_net: cleanText(input.totalNet || input.total_net) === "" ? null : normalizeImportMoney(input.totalNet || input.total_net),
    total_tax: cleanText(input.totalTax || input.total_tax) === "" ? null : normalizeImportMoney(input.totalTax || input.total_tax),
    user_name: cleanText(input.user || input.user_name),
    guest: cleanText(input.guest),
    financial_account: cleanText(input.financialAccount || input.financial_account),
    note: cleanText(input.note),
    raw_payload: input && typeof input === "object" ? input : {},
  };
}

function sanitizeCgdExtratoOrdemImportRow(input = {}, meta = {}) {
  const dataRaw = cleanText(input.dataRaw || input.data_raw || input.data || input.DATA);
  const dataValorRaw = cleanText(input.dataValorRaw || input.data_valor_raw || input.dataValor || input.datavalor || input.DATAVALOR);
  const descritivo = cleanText(input.descritivo || input.DESCRITIVO);
  const montanteInput = input.montante ?? input.MONTANTE;
  const saldoInput = input.saldo ?? input.SALDO;
  const montante = cleanText(montanteInput) === "" ? null : normalizeImportMoney(montanteInput);
  const saldo = cleanText(saldoInput) === "" ? null : normalizeImportMoney(saldoInput);
  if (!dataRaw) throw badRequest("Data is required.");
  if (!dataValorRaw) throw badRequest("Data Valor is required.");
  if (!descritivo) throw badRequest("Descritivo is required.");
  if (!Number.isFinite(montante)) throw badRequest("Montante is required.");
  if (!Number.isFinite(saldo)) throw badRequest("Saldo is required.");
  const data = normalizeImportCgdDate(dataRaw);
  const dataValor = normalizeImportCgdDate(dataValorRaw);
  if (!data) throw badRequest("Data is invalid.");
  if (!dataValor) throw badRequest("Data Valor is invalid.");
  const normalizedMontante = Number(Number(montante).toFixed(2));
  const normalizedSaldo = Number(Number(saldo).toFixed(2));
  return {
    import_batch: cleanText(meta.importBatch),
    source_type: "cgd-extrato-ordem",
    source_name: cleanText(meta.sourceName),
    source_row_number: Number(meta.sourceRowNumber) || 0,
    row_key: importDataRowKey([data, dataValor, descritivo, normalizedMontante, normalizedSaldo]),
    data_raw: dataRaw,
    data,
    data_valor_raw: dataValorRaw,
    data_valor: dataValor,
    descritivo,
    montante: normalizedMontante,
    saldo: normalizedSaldo,
    raw_payload: input && typeof input === "object" ? input : {},
  };
}

function sanitizeCgdCartaoCreditoImportRow(input = {}, meta = {}) {
  const dataRaw = cleanText(input.dataRaw || input.data_raw || input.data || input.DATA);
  const dataValorRaw = cleanText(input.dataValorRaw || input.data_valor_raw || input.dataValor || input.data_valor || input["Data Valor"]);
  const descricao = cleanText(input.descricao || input.descrição || input.descritivo || input.description || input.DESCRICAO || input["Descrição"]);
  const debitoInput = input.debito ?? input.débito ?? input.debit ?? input.DEBITO ?? input["Débito"];
  const creditoInput = input.credito ?? input.crédito ?? input.credit ?? input.CREDITO ?? input["Crédito"];
  const hasDebito = cleanText(debitoInput) !== "";
  const hasCredito = cleanText(creditoInput) !== "";
  const debito = hasDebito ? normalizeImportMoney(debitoInput) : null;
  const credito = hasCredito ? normalizeImportMoney(creditoInput) : null;
  if (!dataRaw) throw badRequest("Data is required.");
  if (!dataValorRaw) throw badRequest("Data Valor is required.");
  if (!descricao) throw badRequest("Descrição is required.");
  if (!hasDebito && !hasCredito) throw badRequest("Débito or Crédito is required.");
  if (hasDebito && !Number.isFinite(debito)) throw badRequest("Débito is invalid.");
  if (hasCredito && !Number.isFinite(credito)) throw badRequest("Crédito is invalid.");
  const data = normalizeImportCgdDate(dataRaw);
  const dataValor = normalizeImportCgdDate(dataValorRaw);
  if (!data) throw badRequest("Data is invalid.");
  if (!dataValor) throw badRequest("Data Valor is invalid.");
  const normalizedDebito = Number.isFinite(debito) ? Number(Number(debito).toFixed(2)) : null;
  const normalizedCredito = Number.isFinite(credito) ? Number(Number(credito).toFixed(2)) : null;
  return {
    import_batch: cleanText(meta.importBatch),
    source_type: "cgd-cartao-credito",
    source_name: cleanText(meta.sourceName),
    source_row_number: Number(meta.sourceRowNumber) || 0,
    row_key: importDataRowKey([data, descricao, normalizedDebito ?? "", normalizedCredito ?? ""]),
    data_raw: dataRaw,
    data,
    data_valor_raw: dataValorRaw,
    data_valor: dataValor,
    descricao,
    debito: normalizedDebito,
    credito: normalizedCredito,
    raw_payload: input && typeof input === "object" ? input : {},
  };
}

module.exports = {
  DEFAULT_IMPORT_DATA_SETTINGS,
  IMPORT_DATA_SETTINGS_KEY,
  IMPORT_DATA_TYPES,
  normalizeImportDataType,
  safeImportDataSettings,
  sanitizeCgdCartaoCreditoImportRow,
  sanitizeCgdExtratoOrdemImportRow,
  sanitizeFdmAccountsImportRow,
  sanitizeFdmBookingsImportRow,
  sanitizeFdmSalesImportRow,
  sanitizeImportDataSettings,
};
