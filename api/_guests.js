const { randomUUID } = require("node:crypto");
const countries = require("./_guests-countries.json");
const { cleanText, normalizeDate } = require("./_supabase");

const GUESTS_SETTING_KEY = "guests";
const DEFAULT_GUESTS_INTEGRATION_MAPPING = {
  name: "name",
  nationality: "nationality",
  birthDate: "birthDate",
  docNumber: "docNumber",
  docType: "docType",
  issuerCountry: "issuerCountry",
  residenceCountry: "issuerCountry",
  residenceCity: "issuerCountry",
  checkIn: "checkIn",
  checkOut: "checkOut",
};
const DEFAULT_GUESTS_SETTINGS = {
  sendTime: "18:00",
  integrationMapping: { ...DEFAULT_GUESTS_INTEGRATION_MAPPING },
};

const SEF_ENDPOINT = "https://siba.sef.pt/baws/boletinsalojamento.asmx";
const SEF_CONFIG = {
  unitCode: "508459893",
  establishment: 0,
  establishmentLabel: "00",
  accessKey: "102907025181",
  hotel: {
    code: "508459893",
    establishment: "00",
    name: "Lisboa Central Hostel",
    abbreviation: "LCH",
    address: "Rua Rodrigues Sampaio 160",
    city: "Lisboa",
    postalCode: "1150",
    postalZone: "282",
    phone: "309881038",
    fax: "309881038",
    contactName: "Lisboa Central Hostel",
    contactEmail: "global@lisboacentralhostel.com",
  },
};

function normalizeTime(value, fallback = "18:00") {
  const raw = cleanText(value);
  return /^\d{2}:\d{2}$/.test(raw) ? raw : fallback;
}

function normalizeGuestText(value, maxLength = 200) {
  return cleanText(value).slice(0, maxLength);
}

function normalizeGuestName(value) {
  return normalizeGuestText(value, 120);
}

function normalizeDocType(value) {
  const raw = cleanText(value).toUpperCase();
  return raw === "B" || raw === "O" ? raw : "P";
}

function normalizeHA(value) {
  return cleanText(value).toUpperCase() === "A" ? "A" : "H";
}

function normalizeDocNumber(value) {
  return cleanText(value).toUpperCase().replace(/\s+/g, "");
}

function normalizeCountryLookupKey(value) {
  return cleanText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .trim();
}

function buildCountryLookups() {
  const exact = new Map();
  countries.forEach((item) => {
    const code = cleanText(item.code).toUpperCase();
    const name = cleanText(item.name);
    const abbr = cleanText(item.abbr);
    const entry = { code, name, abbr };
    [code, name, abbr].forEach((value) => {
      const key = normalizeCountryLookupKey(value);
      if (key && !exact.has(key)) exact.set(key, entry);
    });
  });
  return { exact };
}

const COUNTRY_LOOKUPS = buildCountryLookups();

function resolveCountry(value) {
  const raw = normalizeGuestText(value, 140);
  const key = normalizeCountryLookupKey(raw);
  const entry = COUNTRY_LOOKUPS.exact.get(key) || null;
  if (entry) {
    return {
      input: raw,
      code: entry.code,
      name: entry.name,
      abbr: entry.abbr,
      isValid: true,
    };
  }
  if (/^[A-Z]{3}$/.test(key)) {
    return {
      input: raw || key,
      code: key,
      name: raw,
      abbr: raw,
      isValid: true,
    };
  }
  return {
    input: raw,
    code: "",
    name: raw,
    abbr: raw,
    isValid: false,
  };
}

function normalizeGuestMappingValue(value, fallback) {
  const raw = cleanText(value);
  return raw || fallback;
}

function normalizeGuestSettings(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const mappingSource = source.integrationMapping && typeof source.integrationMapping === "object" ? source.integrationMapping : {};
  return {
    sendTime: normalizeTime(source.sendTime ?? source.send_time, DEFAULT_GUESTS_SETTINGS.sendTime),
    integrationMapping: {
      name: normalizeGuestMappingValue(mappingSource.name, DEFAULT_GUESTS_INTEGRATION_MAPPING.name),
      nationality: normalizeGuestMappingValue(mappingSource.nationality, DEFAULT_GUESTS_INTEGRATION_MAPPING.nationality),
      birthDate: normalizeGuestMappingValue(mappingSource.birthDate ?? mappingSource.birth_date, DEFAULT_GUESTS_INTEGRATION_MAPPING.birthDate),
      docNumber: normalizeGuestMappingValue(mappingSource.docNumber ?? mappingSource.doc_number, DEFAULT_GUESTS_INTEGRATION_MAPPING.docNumber),
      docType: normalizeGuestMappingValue(mappingSource.docType ?? mappingSource.doc_type, DEFAULT_GUESTS_INTEGRATION_MAPPING.docType),
      issuerCountry: normalizeGuestMappingValue(mappingSource.issuerCountry ?? mappingSource.issuer_country, DEFAULT_GUESTS_INTEGRATION_MAPPING.issuerCountry),
      residenceCountry: normalizeGuestMappingValue(mappingSource.residenceCountry ?? mappingSource.residence_country, DEFAULT_GUESTS_INTEGRATION_MAPPING.residenceCountry),
      residenceCity: normalizeGuestMappingValue(mappingSource.residenceCity ?? mappingSource.residence_city, DEFAULT_GUESTS_INTEGRATION_MAPPING.residenceCity),
      checkIn: normalizeGuestMappingValue(mappingSource.checkIn ?? mappingSource.check_in, DEFAULT_GUESTS_INTEGRATION_MAPPING.checkIn),
      checkOut: normalizeGuestMappingValue(mappingSource.checkOut ?? mappingSource.check_out, DEFAULT_GUESTS_INTEGRATION_MAPPING.checkOut),
    },
  };
}

function normalizeSentStatus(value) {
  const raw = cleanText(value).toLowerCase();
  if (raw === "sent") return "sent";
  if (raw === "error") return "error";
  return "pending";
}

function sanitizeGuestRecord(input = {}, existing = {}) {
  const name = normalizeGuestName(input.name ?? existing.name);
  const birthDate = normalizeDate(input.birthDate ?? input.birth_date ?? existing.birthDate ?? existing.birth_date);
  const checkIn = normalizeDate(input.checkIn ?? input.check_in ?? existing.checkIn ?? existing.check_in);
  const checkOut = normalizeDate(input.checkOut ?? input.check_out ?? existing.checkOut ?? existing.check_out);
  const nationality = resolveCountry(input.nationality ?? existing.nationality);
  const issuerCountry = resolveCountry(input.issuerCountry ?? input.issuer_country ?? existing.issuerCountry ?? existing.issuer_country);
  const residenceCountry = resolveCountry(input.residenceCountry ?? input.residence_country ?? existing.residenceCountry ?? existing.residence_country);
  const birthPlace = normalizeGuestText(input.birthPlace ?? input.birth_place ?? existing.birthPlace ?? existing.birth_place, 60);
  const residenceCity = normalizeGuestText(input.residenceCity ?? input.residence_city ?? existing.residenceCity ?? existing.residence_city, 60);
  const record = {
    id: cleanText(input.id || existing.id) || randomUUID(),
    ha: normalizeHA(input.ha ?? existing.ha),
    name,
    nationality: nationality.input,
    nationalityCode: nationality.code,
    birthDate,
    birthPlace,
    docNumber: normalizeDocNumber(input.docNumber ?? input.doc_number ?? existing.docNumber ?? existing.doc_number),
    docType: normalizeDocType(input.docType ?? input.doc_type ?? existing.docType ?? existing.doc_type),
    issuerCountry: issuerCountry.input,
    issuerCountryCode: issuerCountry.code,
    residenceCountry: residenceCountry.input,
    residenceCountryCode: residenceCountry.code,
    residenceCity,
    checkIn,
    checkOut,
    sentStatus: normalizeSentStatus(input.sentStatus ?? input.sent_status ?? existing.sentStatus ?? existing.sent_status),
    sentAt: cleanText(input.sentAt || input.sent_at || existing.sentAt || existing.sent_at),
    sendError: normalizeGuestText(input.sendError ?? input.send_error ?? existing.sendError ?? existing.send_error, 400),
    sendBatchNumber: Number.parseInt(input.sendBatchNumber ?? input.send_batch_number ?? existing.sendBatchNumber ?? existing.send_batch_number, 10) || 0,
    createdAt: cleanText(input.createdAt || input.created_at || existing.createdAt || existing.created_at),
    updatedAt: cleanText(input.updatedAt || input.updated_at || existing.updatedAt || existing.updated_at),
  };
  validateGuestRecord(record);
  return record;
}

function sanitizeBlacklistRecord(input = {}, existing = {}) {
  const nationality = resolveCountry(input.nationality ?? existing.nationality);
  const record = {
    id: cleanText(input.id || existing.id) || randomUUID(),
    name: normalizeGuestName(input.name ?? existing.name),
    nationality: nationality.input,
    nationalityCode: nationality.code,
    birthDate: normalizeDate(input.birthDate ?? input.birth_date ?? existing.birthDate ?? existing.birth_date),
    docNumber: normalizeDocNumber(input.docNumber ?? input.doc_number ?? existing.docNumber ?? existing.doc_number),
    whatHappened: normalizeGuestText(input.whatHappened ?? input.what_happened ?? existing.whatHappened ?? existing.what_happened, 1000),
    occurrenceDate: normalizeDate(input.occurrenceDate ?? input.occurrence_date ?? existing.occurrenceDate ?? existing.occurrence_date),
    whoReported: normalizeGuestText(input.whoReported ?? input.who_reported ?? existing.whoReported ?? existing.who_reported, 140),
    createdAt: cleanText(input.createdAt || input.created_at || existing.createdAt || existing.created_at),
    updatedAt: cleanText(input.updatedAt || input.updated_at || existing.updatedAt || existing.updated_at),
  };
  validateBlacklistRecord(record);
  return record;
}

function validateGuestRecord(record) {
  if (!record.name) {
    const error = new Error("Guest name is required.");
    error.statusCode = 400;
    throw error;
  }
  if (!record.birthDate) {
    const error = new Error("Birth date is required.");
    error.statusCode = 400;
    throw error;
  }
  if (!record.docNumber) {
    const error = new Error("Document number is required.");
    error.statusCode = 400;
    throw error;
  }
  if (!record.checkIn) {
    const error = new Error("Check-in is required.");
    error.statusCode = 400;
    throw error;
  }
  if (!record.checkOut) {
    const error = new Error("Check-out is required.");
    error.statusCode = 400;
    throw error;
  }
  if (record.checkOut < record.checkIn) {
    const error = new Error("Check-out must be after or equal to check-in.");
    error.statusCode = 400;
    throw error;
  }
}

function validateBlacklistRecord(record) {
  if (!record.name && !record.docNumber) {
    const error = new Error("Blacklist record requires a name or document number.");
    error.statusCode = 400;
    throw error;
  }
  if (!record.occurrenceDate) {
    const error = new Error("Occurrence date is required.");
    error.statusCode = 400;
    throw error;
  }
}

function sanitizeGuestsPayload(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const settings = normalizeGuestSettings(source.settings || source.guestsSettings || {});
  const rows = (Array.isArray(source.rows || source.records) ? source.rows || source.records : [])
    .map((item) => sanitizeGuestRecord(item))
    .sort((a, b) => cleanText(b.checkIn).localeCompare(cleanText(a.checkIn)) || cleanText(a.name).localeCompare(cleanText(b.name)));
  const blacklist = (Array.isArray(source.blacklist) ? source.blacklist : [])
    .map((item) => sanitizeBlacklistRecord(item))
    .sort((a, b) => cleanText(b.occurrenceDate).localeCompare(cleanText(a.occurrenceDate)) || cleanText(a.name).localeCompare(cleanText(b.name)));
  const lastFileNumber = Math.max(0, Number.parseInt(source.lastFileNumber ?? source.last_file_number, 10) || 0);
  return { settings, rows, blacklist, lastFileNumber };
}

function calculateAge(birthDate, todayIso) {
  const birth = cleanText(birthDate);
  const today = cleanText(todayIso);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(birth) || !/^\d{4}-\d{2}-\d{2}$/.test(today)) return null;
  let age = Number(today.slice(0, 4)) - Number(birth.slice(0, 4));
  const monthDayToday = today.slice(5);
  const monthDayBirth = birth.slice(5);
  if (monthDayToday < monthDayBirth) age -= 1;
  return age >= 0 ? age : null;
}

function normalizeMatchName(value) {
  return cleanText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9 ]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function tokenizeName(value) {
  return normalizeMatchName(value)
    .split(" ")
    .map((item) => item.trim())
    .filter((item) => item.length >= 3);
}

function guestBlacklistMatches(record, blacklistRecord) {
  const guestDoc = normalizeDocNumber(record?.docNumber);
  const blackDoc = normalizeDocNumber(blacklistRecord?.docNumber);
  if (guestDoc && blackDoc && guestDoc === blackDoc) return "document";
  const guestName = normalizeMatchName(record?.name);
  const blackName = normalizeMatchName(blacklistRecord?.name);
  if (guestName && blackName && guestName === blackName) return "exact-name";
  if (cleanText(record?.birthDate) && cleanText(record?.birthDate) === cleanText(blacklistRecord?.birthDate)) {
    const guestTokens = tokenizeName(record?.name);
    const blacklistTokens = tokenizeName(blacklistRecord?.name);
    if (guestTokens.some((token) => blacklistTokens.includes(token))) return "name-birthdate";
  }
  return "";
}

function lisbonTodayIso() {
  const formatter = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/Lisbon",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });
  return formatter.format(new Date());
}

function shouldShowGuestsAlert(rows, settings, todayIso = lisbonTodayIso()) {
  const sendTime = normalizeTime(settings?.sendTime, DEFAULT_GUESTS_SETTINGS.sendTime);
  const timeFormatter = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/Lisbon",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });
  const currentTime = timeFormatter.format(new Date());
  if (currentTime < sendTime) return false;
  return (Array.isArray(rows) ? rows : []).some((row) => cleanText(row.sentStatus) !== "sent" && cleanText(row.checkIn) && cleanText(row.checkIn) < todayIso);
}

function splitGuestNameForApi(name) {
  const cleaned = normalizeMatchName(name).replace(/[^A-ZÇÃÁÀÉÊÍÕÔÓÚ' -]+/g, " ").replace(/\s+/g, " ").trim();
  if (!cleaned) return { surname: "", givenNames: "" };
  const parts = cleaned.split(" ").filter(Boolean);
  if (parts.length === 1) return { surname: parts[0], givenNames: "" };
  return {
    surname: parts[parts.length - 1].slice(0, 40),
    givenNames: parts.slice(0, -1).join(" ").slice(0, 40),
  };
}

function escapeXml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function xmlDateTime(date) {
  const safe = normalizeDate(date);
  return safe ? `${safe}T00:00:00` : "";
}

function buildBalXml(records, fileNumber) {
  const hotel = SEF_CONFIG.hotel;
  const boletins = records.map((record) => {
    const names = splitGuestNameForApi(record.name);
    return `  <Boletim_Alojamento>
    <Apelido>${escapeXml(names.surname || record.name)}</Apelido>
    ${names.givenNames ? `<Nome>${escapeXml(names.givenNames)}</Nome>` : ""}
    <Nacionalidade>${escapeXml(record.nationalityCode)}</Nacionalidade>
    <Data_Nascimento>${escapeXml(xmlDateTime(record.birthDate))}</Data_Nascimento>
    ${record.birthPlace ? `<Local_Nascimento>${escapeXml(record.birthPlace.slice(0, 30))}</Local_Nascimento>` : ""}
    <Documento_Identificacao>${escapeXml(record.docNumber.slice(0, 16))}</Documento_Identificacao>
    <Pais_Emissor_Documento>${escapeXml(record.issuerCountryCode)}</Pais_Emissor_Documento>
    <Tipo_Documento>${escapeXml(record.docType)}</Tipo_Documento>
    <Data_Entrada>${escapeXml(xmlDateTime(record.checkIn))}</Data_Entrada>
    ${record.checkOut ? `<Data_Saida>${escapeXml(xmlDateTime(record.checkOut))}</Data_Saida>` : ""}
    <Pais_Residencia_Origem>${escapeXml(record.residenceCountryCode)}</Pais_Residencia_Origem>
    <Local_Residencia_Origem>${escapeXml(record.residenceCity.slice(0, 30))}</Local_Residencia_Origem>
  </Boletim_Alojamento>`;
  }).join("\n");
  return `<?xml version="1.0" encoding="utf-8"?>
<MovimentoBAL xmlns="http://sef.pt/BAws">
  <Unidade_Hoteleira>
    <Codigo_Unidade_Hoteleira>${escapeXml(hotel.code)}</Codigo_Unidade_Hoteleira>
    <Estabelecimento>${escapeXml(hotel.establishment)}</Estabelecimento>
    <Nome>${escapeXml(hotel.name.slice(0, 40))}</Nome>
    <Abreviatura>${escapeXml(hotel.abbreviation.slice(0, 15))}</Abreviatura>
    <Morada>${escapeXml(hotel.address.slice(0, 40))}</Morada>
    <Localidade>${escapeXml(hotel.city.slice(0, 30))}</Localidade>
    <Codigo_Postal>${escapeXml(hotel.postalCode)}</Codigo_Postal>
    <Zona_Postal>${escapeXml(hotel.postalZone)}</Zona_Postal>
    <Telefone>${escapeXml(hotel.phone)}</Telefone>
    <Fax>${escapeXml(hotel.fax)}</Fax>
    <Nome_Contacto>${escapeXml(hotel.contactName.slice(0, 40))}</Nome_Contacto>
    <Email_Contacto>${escapeXml(hotel.contactEmail.slice(0, 140))}</Email_Contacto>
  </Unidade_Hoteleira>
${boletins}
  <Envio>
    <Numero_Ficheiro>${fileNumber}</Numero_Ficheiro>
    <Data_Movimento>${escapeXml(new Date().toISOString().slice(0, 19))}</Data_Movimento>
  </Envio>
</MovimentoBAL>`;
}

function buildSoapEnvelope(base64Payload) {
  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <EntregaBoletinsAlojamento xmlns="http://sef.pt/">
      <UnidadeHoteleira>${SEF_CONFIG.unitCode}</UnidadeHoteleira>
      <Estabelecimento>${SEF_CONFIG.establishment}</Estabelecimento>
      <ChaveAcesso>${SEF_CONFIG.accessKey}</ChaveAcesso>
      <Boletins>${escapeXml(base64Payload)}</Boletins>
    </EntregaBoletinsAlojamento>
  </soap:Body>
</soap:Envelope>`;
}

module.exports = {
  COUNTRIES: countries,
  DEFAULT_GUESTS_INTEGRATION_MAPPING,
  DEFAULT_GUESTS_SETTINGS,
  GUESTS_SETTING_KEY,
  SEF_CONFIG,
  SEF_ENDPOINT,
  buildBalXml,
  buildSoapEnvelope,
  calculateAge,
  guestBlacklistMatches,
  lisbonTodayIso,
  normalizeCountryLookupKey,
  normalizeGuestSettings,
  resolveCountry,
  sanitizeBlacklistRecord,
  sanitizeGuestRecord,
  sanitizeGuestsPayload,
  shouldShowGuestsAlert,
};
