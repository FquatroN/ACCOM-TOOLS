const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  COUNTRIES,
  DEFAULT_GUESTS_SETTINGS,
  GUESTS_SETTING_KEY,
  lisbonTodayIso,
  sefDocType,
  sanitizeGuestRecord,
  sanitizeGuestsPayload,
} = require("./_guests");

function cleanId(value) {
  return String(value || "").trim();
}

function cleanQueryValue(value) {
  return Array.isArray(value) ? cleanId(value[0]) : cleanId(value);
}

function shiftIsoDate(isoDate, days) {
  const raw = cleanId(isoDate);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return "";
  const date = new Date(`${raw}T00:00:00Z`);
  if (Number.isNaN(date.getTime())) return "";
  date.setUTCDate(date.getUTCDate() + Number(days || 0));
  return date.toISOString().slice(0, 10);
}

function normalizeGuestListFilters(query = {}) {
  return {
    ha: cleanQueryValue(query.ha).toUpperCase(),
    search: cleanQueryValue(query.search),
    nationality: cleanQueryValue(query.nationality),
    checkInFrom: cleanQueryValue(query.checkInFrom || query.check_in_from),
    checkInTo: cleanQueryValue(query.checkInTo || query.check_in_to),
    checkOutFrom: cleanQueryValue(query.checkOutFrom || query.check_out_from),
    checkOutTo: cleanQueryValue(query.checkOutTo || query.check_out_to),
  };
}

function hasExplicitGuestListFilters(filters = {}) {
  return !!(
    cleanId(filters.ha) ||
    cleanId(filters.search) ||
    cleanId(filters.nationality) ||
    cleanId(filters.checkInFrom) ||
    cleanId(filters.checkInTo) ||
    cleanId(filters.checkOutFrom) ||
    cleanId(filters.checkOutTo)
  );
}

function normalizeGuestCountryKey(value) {
  return cleanId(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .trim();
}

function guestNationalityMatchesFilter(record, filter) {
  const normalizedFilter = normalizeGuestCountryKey(filter);
  if (!normalizedFilter) return true;
  return [record?.nationality, record?.nationalityCode]
    .map((value) => normalizeGuestCountryKey(value))
    .some((value) => value === normalizedFilter || value.includes(normalizedFilter));
}

function sanitizeGuestFilterTerm(value) {
  return cleanId(value)
    .replace(/[(),]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function sortGuestListRows(rows) {
  return [...(Array.isArray(rows) ? rows : [])].sort((a, b) => {
    const checkInCompare = cleanId(b?.checkIn).localeCompare(cleanId(a?.checkIn));
    if (checkInCompare) return checkInCompare;
    const at = new Date(cleanId(a?.createdAt)).getTime() || 0;
    const bt = new Date(cleanId(b?.createdAt)).getTime() || 0;
    return bt - at || cleanId(a?.name).localeCompare(cleanId(b?.name));
  });
}

function applyGuestListFilters(rows, filters = {}) {
  const safe = normalizeGuestListFilters(filters);
  const ha = cleanId(safe.ha).toUpperCase();
  const search = cleanId(safe.search).toLowerCase();
  const checkInFrom = cleanId(safe.checkInFrom);
  const checkInTo = cleanId(safe.checkInTo);
  const checkOutFrom = cleanId(safe.checkOutFrom);
  const checkOutTo = cleanId(safe.checkOutTo);
  return sortGuestListRows(rows)
    .filter((row) => !ha || cleanId(row?.ha).toUpperCase() === ha)
    .filter((row) => !search || cleanId(row?.name).toLowerCase().includes(search) || cleanId(row?.docNumber).toLowerCase().includes(search))
    .filter((row) => guestNationalityMatchesFilter(row, safe.nationality))
    .filter((row) => !checkInFrom || cleanId(row?.checkIn) >= checkInFrom)
    .filter((row) => !checkInTo || cleanId(row?.checkIn) <= checkInTo)
    .filter((row) => !checkOutFrom || cleanId(row?.checkOut) >= checkOutFrom)
    .filter((row) => !checkOutTo || cleanId(row?.checkOut) <= checkOutTo);
}

function scopeGuestListRows(rows, filters = {}, { includeDefaultRecent = true } = {}) {
  const safe = normalizeGuestListFilters(filters);
  if (hasExplicitGuestListFilters(safe)) return applyGuestListFilters(rows, safe);
  const sorted = sortGuestListRows(rows);
  if (!includeDefaultRecent) return sorted;
  const cutoff = shiftIsoDate(lisbonTodayIso(), -60);
  if (!cutoff) return sorted;
  return sorted.filter((row) => !cleanId(row?.checkOut) || cleanId(row?.checkOut) >= cutoff);
}

async function loadGuestsPayloadRow() {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${encodeURIComponent(GUESTS_SETTING_KEY)}&limit=1`, {
    method: "GET",
  });
  const row = Array.isArray(rows) && rows[0] ? rows[0] : null;
  return {
    rowId: row?.id || "",
    payload: sanitizeGuestsPayload(row?.payload || { settings: DEFAULT_GUESTS_SETTINGS, rows: [], blacklist: [], lastFileNumber: 0 }),
  };
}

async function saveGuestsPayload(rowId, payload) {
  const safe = sanitizeGuestsPayload(payload);
  if (rowId) {
    await restQuery(`app_settings?id=eq.${encodeURIComponent(rowId)}`, {
      method: "PATCH",
      body: { payload: safe, updated_at: new Date().toISOString() },
    });
    return safe;
  }
  const created = await restQuery("app_settings", {
    method: "POST",
    body: [{ setting_key: GUESTS_SETTING_KEY, payload: safe }],
  });
  return sanitizeGuestsPayload(Array.isArray(created) && created[0]?.payload ? created[0].payload : safe);
}

function isMissingGuestsTableError(error) {
  const message = String(error?.message || "").toLowerCase();
  return message.includes("guest_records") && (
    message.includes("could not find") ||
    message.includes("schema cache") ||
    message.includes("relation") ||
    message.includes("does not exist")
  );
}

function mapGuestsSchemaError(error) {
  const message = String(error?.message || "");
  if (
    message.includes('null value in column "check_out" of relation "guest_records" violates not-null constraint') ||
    message.includes('null value in column "check_in" of relation "guest_records" violates not-null constraint')
  ) {
    const next = new Error("The Guests table still requires check-in/check-out. Please run the migration file 2026-05-12-guests-open-dates.sql in Supabase.");
    next.statusCode = 400;
    return next;
  }
  return error;
}

function mapGuestTableRow(row) {
  return sanitizeGuestRecord({
    id: row?.id,
    ha: row?.ha,
    name: row?.name,
    nationality: row?.nationality,
    nationalityCode: row?.nationality_code,
    birthDate: row?.birth_date,
    birthPlace: row?.birth_place,
    docNumber: row?.doc_number,
    docType: row?.doc_type,
    issuerCountry: row?.issuer_country,
    issuerCountryCode: row?.issuer_country_code,
    residenceCountry: row?.residence_country,
    residenceCountryCode: row?.residence_country_code,
    residenceCity: row?.residence_city,
    checkIn: row?.check_in,
    checkOut: row?.check_out,
    sentStatus: row?.sent_status,
    sentAt: row?.sent_at,
    sendError: row?.send_error,
    sendBatchNumber: row?.send_batch_number,
    createdAt: row?.created_at,
    updatedAt: row?.updated_at,
  });
}

function buildGuestTableQuery(filters = {}, { recentOnly = false } = {}) {
  const safe = normalizeGuestListFilters(filters);
  const params = new URLSearchParams();
  params.set("select", "*");
  params.set("order", "check_in.desc,created_at.desc,name.asc");
  params.set("limit", "10000");

  const expressions = [];
  if (cleanId(safe.ha)) expressions.push(`ha.eq.${sanitizeGuestFilterTerm(safe.ha).toUpperCase()}`);
  if (cleanId(safe.checkInFrom)) expressions.push(`check_in.gte.${sanitizeGuestFilterTerm(safe.checkInFrom)}`);
  if (cleanId(safe.checkInTo)) expressions.push(`check_in.lte.${sanitizeGuestFilterTerm(safe.checkInTo)}`);
  if (cleanId(safe.checkOutFrom)) expressions.push(`check_out.gte.${sanitizeGuestFilterTerm(safe.checkOutFrom)}`);
  if (cleanId(safe.checkOutTo)) expressions.push(`check_out.lte.${sanitizeGuestFilterTerm(safe.checkOutTo)}`);
  if (cleanId(safe.search)) {
    const search = sanitizeGuestFilterTerm(safe.search);
    if (search) expressions.push(`or(name.ilike.*${search}*,doc_number.ilike.*${search}*)`);
  }
  if (cleanId(safe.nationality)) {
    const nationality = sanitizeGuestFilterTerm(safe.nationality);
    if (nationality) expressions.push(`or(nationality.ilike.*${nationality}*,nationality_code.ilike.*${nationality}*)`);
  }
  if (recentOnly) {
    const cutoff = shiftIsoDate(lisbonTodayIso(), -60);
    if (cutoff) expressions.push(`or(check_out.gte.${cutoff},check_out.is.null)`);
  }
  if (expressions.length) params.set("and", `(${expressions.join(",")})`);
  return `guest_records?${params.toString()}`;
}

async function loadGuestTableRows(filters = {}, options = {}) {
  const safe = normalizeGuestListFilters(filters);
  const mode = cleanId(options?.mode).toLowerCase();
  const recentOnly = mode !== "all" && !hasExplicitGuestListFilters(safe);
  const rows = await restQuery(buildGuestTableQuery(safe, { recentOnly }), {
    method: "GET",
  });
  const mapped = (Array.isArray(rows) ? rows : []).map(mapGuestTableRow);
  if (mode === "all") return sortGuestListRows(mapped);
  return sortGuestListRows(mapped);
}

function buildGuestTableBody(record, existing = {}) {
  return {
    id: cleanId(record.id || existing.id) || undefined,
    ha: record.ha,
    name: record.name,
    nationality: record.nationality || "",
    nationality_code: record.nationalityCode || "",
    birth_date: record.birthDate,
    birth_place: record.birthPlace || "",
    doc_number: record.docNumber,
    doc_type: sefDocType(record.docType),
    issuer_country: record.issuerCountry || "",
    issuer_country_code: record.issuerCountryCode || "",
    residence_country: record.residenceCountry || "",
    residence_country_code: record.residenceCountryCode || "",
    residence_city: record.residenceCity || "",
    check_in: record.checkIn || null,
    check_out: record.checkOut || null,
    sent_status: record.sentStatus || "pending",
    sent_at: record.sentAt || null,
    send_error: record.sendError || "",
    send_batch_number: Math.max(0, Number.parseInt(record.sendBatchNumber, 10) || 0),
    created_at: existing.createdAt || record.createdAt || new Date().toISOString(),
    updated_at: new Date().toISOString(),
  };
}

async function createGuestTableRow(record, existing = {}) {
  await restQuery("guest_records", {
    method: "POST",
    body: [buildGuestTableBody(record, existing)],
    preferRepresentation: true,
  });
}

async function updateGuestTableRow(id, record, existing = {}) {
  await restQuery(`guest_records?id=eq.${encodeURIComponent(id)}`, {
    method: "PATCH",
    body: buildGuestTableBody(record, existing),
    preferRepresentation: true,
  });
}

async function deleteGuestTableRow(id) {
  await restQuery(`guest_records?id=eq.${encodeURIComponent(id)}`, {
    method: "DELETE",
  });
}

function validateGuestSave(existingRows, record, { excludeId = "" } = {}) {
  const duplicate = existingRows.find((item) =>
    cleanId(item.id) !== cleanId(excludeId) &&
    cleanId(item.docNumber) &&
    cleanId(record.checkIn) &&
    cleanId(item.docNumber) === cleanId(record.docNumber) &&
    cleanId(item.checkIn) === cleanId(record.checkIn)
  );
  if (duplicate) {
    const error = new Error(`A guest with the same document number and check-in already exists: ${duplicate.name || "Unknown guest"} (${duplicate.docNumber}, ${duplicate.checkIn}).`);
    error.statusCode = 400;
    throw error;
  }
}

function mergeLegacyRows(existingRows, input, id = "") {
  const existing = existingRows.find((item) => cleanId(item.id) === cleanId(id)) || {};
  const nextRecord = sanitizeGuestRecord({ ...existing, ...input, id: id || input?.id || existing.id }, existing);
  validateGuestSave(existingRows, nextRecord, { excludeId: id || nextRecord.id });
  const nextRows = [...existingRows];
  const index = id ? nextRows.findIndex((item) => cleanId(item.id) === cleanId(id)) : -1;
  if (index >= 0) {
    nextRows[index] = {
      ...nextRecord,
      id: nextRows[index].id,
      createdAt: nextRows[index].createdAt || nextRecord.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
  } else {
    nextRows.push({
      ...nextRecord,
      createdAt: nextRecord.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });
  }
  return sanitizeGuestsPayload({ rows: nextRows }).rows;
}

async function loadRowsAndSettings(filters = {}, options = {}) {
  const { payload } = await loadGuestsPayloadRow();
  try {
    const rows = await loadGuestTableRows(filters, options);
    return { mode: "table", settings: payload.settings, rows };
  } catch (error) {
    if (!isMissingGuestsTableError(error)) throw error;
    return { mode: "legacy", settings: payload.settings, rows: scopeGuestListRows(payload.rows, filters, { includeDefaultRecent: cleanId(options?.mode).toLowerCase() !== "all" }) };
  }
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "guests");
    const responseFilters = normalizeGuestListFilters(req.query);

    if (req.method === "GET") {
      const current = await loadRowsAndSettings(responseFilters);
      res.status(200).json({ rows: current.rows, settings: current.settings, countries: COUNTRIES });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const current = await loadRowsAndSettings({}, { mode: "all" });
      if (current.mode === "legacy") {
        const { rowId, payload } = await loadGuestsPayloadRow();
        const nextRows = mergeLegacyRows(payload.rows, body);
        const saved = await saveGuestsPayload(rowId, { ...payload, rows: nextRows });
        res.status(200).json({ rows: scopeGuestListRows(saved.rows, responseFilters), settings: saved.settings, countries: COUNTRIES });
        return;
      }
      const nextRecord = sanitizeGuestRecord(body);
      validateGuestSave(current.rows, nextRecord);
      await createGuestTableRow({
        ...nextRecord,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      });
      const rows = await loadGuestTableRows(responseFilters);
      res.status(200).json({ rows, settings: current.settings, countries: COUNTRIES });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanId(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Guest id is required." });
        return;
      }
      const body = await parseBody(req);
      const current = await loadRowsAndSettings({}, { mode: "all" });
      if (current.mode === "legacy") {
        const { rowId, payload } = await loadGuestsPayloadRow();
        const existing = payload.rows.find((item) => cleanId(item.id) === id);
        if (!existing) {
          res.status(404).json({ error: "Guest record not found." });
          return;
        }
        const nextRows = mergeLegacyRows(payload.rows, { ...existing, ...body, id: existing.id }, id);
        const saved = await saveGuestsPayload(rowId, { ...payload, rows: nextRows });
        res.status(200).json({ rows: scopeGuestListRows(saved.rows, responseFilters), settings: saved.settings, countries: COUNTRIES });
        return;
      }
      const existing = current.rows.find((item) => cleanId(item.id) === id);
      if (!existing) {
        res.status(404).json({ error: "Guest record not found." });
        return;
      }
      const nextRecord = sanitizeGuestRecord({ ...existing, ...body, id }, existing);
      validateGuestSave(current.rows, nextRecord, { excludeId: id });
      await updateGuestTableRow(id, nextRecord, existing);
      const rows = await loadGuestTableRows(responseFilters);
      res.status(200).json({ rows, settings: current.settings, countries: COUNTRIES });
      return;
    }

    if (req.method === "DELETE") {
      const id = cleanId(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Guest id is required." });
        return;
      }
      const current = await loadRowsAndSettings({}, { mode: "all" });
      if (current.mode === "legacy") {
        const { rowId, payload } = await loadGuestsPayloadRow();
        const nextRows = payload.rows.filter((item) => cleanId(item.id) !== id);
        const saved = await saveGuestsPayload(rowId, { ...payload, rows: nextRows });
        res.status(200).json({ rows: scopeGuestListRows(saved.rows, responseFilters), settings: saved.settings, countries: COUNTRIES });
        return;
      }
      await deleteGuestTableRow(id);
      const rows = await loadGuestTableRows(responseFilters);
      res.status(200).json({ rows, settings: current.settings, countries: COUNTRIES });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, mapGuestsSchemaError(error));
  }
};
