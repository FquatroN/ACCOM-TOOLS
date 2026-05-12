const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  COUNTRIES,
  DEFAULT_GUESTS_SETTINGS,
  GUESTS_SETTING_KEY,
  sanitizeGuestRecord,
  sanitizeGuestsPayload,
} = require("./_guests");

function cleanId(value) {
  return String(value || "").trim();
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

async function loadGuestTableRows() {
  const rows = await restQuery("guest_records?select=*&order=check_in.desc,check_out.desc,name.asc", {
    method: "GET",
  });
  return (Array.isArray(rows) ? rows : []).map(mapGuestTableRow);
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
    doc_type: record.docType,
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
    const error = new Error("A guest with the same document number and check-in already exists.");
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

async function loadRowsAndSettings() {
  const { payload } = await loadGuestsPayloadRow();
  try {
    const rows = await loadGuestTableRows();
    return { mode: "table", settings: payload.settings, rows };
  } catch (error) {
    if (!isMissingGuestsTableError(error)) throw error;
    return { mode: "legacy", settings: payload.settings, rows: payload.rows };
  }
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "guests");

    if (req.method === "GET") {
      const current = await loadRowsAndSettings();
      res.status(200).json({ rows: current.rows, settings: current.settings, countries: COUNTRIES });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const current = await loadRowsAndSettings();
      if (current.mode === "legacy") {
        const { rowId, payload } = await loadGuestsPayloadRow();
        const nextRows = mergeLegacyRows(payload.rows, body);
        const saved = await saveGuestsPayload(rowId, { ...payload, rows: nextRows });
        res.status(200).json({ rows: saved.rows, settings: saved.settings, countries: COUNTRIES });
        return;
      }
      const nextRecord = sanitizeGuestRecord(body);
      validateGuestSave(current.rows, nextRecord);
      await createGuestTableRow({
        ...nextRecord,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      });
      const rows = await loadGuestTableRows();
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
      const current = await loadRowsAndSettings();
      if (current.mode === "legacy") {
        const { rowId, payload } = await loadGuestsPayloadRow();
        const existing = payload.rows.find((item) => cleanId(item.id) === id);
        if (!existing) {
          res.status(404).json({ error: "Guest record not found." });
          return;
        }
        if (cleanId(existing.sentStatus).toLowerCase() === "sent") {
          res.status(400).json({ error: "Sent guest records cannot be modified." });
          return;
        }
        const nextRows = mergeLegacyRows(payload.rows, { ...existing, ...body, id: existing.id }, id);
        const saved = await saveGuestsPayload(rowId, { ...payload, rows: nextRows });
        res.status(200).json({ rows: saved.rows, settings: saved.settings, countries: COUNTRIES });
        return;
      }
      const existing = current.rows.find((item) => cleanId(item.id) === id);
      if (!existing) {
        res.status(404).json({ error: "Guest record not found." });
        return;
      }
      if (cleanId(existing.sentStatus).toLowerCase() === "sent") {
        res.status(400).json({ error: "Sent guest records cannot be modified." });
        return;
      }
      const nextRecord = sanitizeGuestRecord({ ...existing, ...body, id }, existing);
      validateGuestSave(current.rows, nextRecord, { excludeId: id });
      await updateGuestTableRow(id, nextRecord, existing);
      const rows = await loadGuestTableRows();
      res.status(200).json({ rows, settings: current.settings, countries: COUNTRIES });
      return;
    }

    if (req.method === "DELETE") {
      const id = cleanId(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Guest id is required." });
        return;
      }
      const current = await loadRowsAndSettings();
      if (current.mode === "legacy") {
        const { rowId, payload } = await loadGuestsPayloadRow();
        const existing = payload.rows.find((item) => cleanId(item.id) === id);
        if (cleanId(existing?.sentStatus).toLowerCase() === "sent") {
          res.status(400).json({ error: "Sent guest records cannot be modified." });
          return;
        }
        const nextRows = payload.rows.filter((item) => cleanId(item.id) !== id);
        const saved = await saveGuestsPayload(rowId, { ...payload, rows: nextRows });
        res.status(200).json({ rows: saved.rows, settings: saved.settings, countries: COUNTRIES });
        return;
      }
      const existing = current.rows.find((item) => cleanId(item.id) === id);
      if (cleanId(existing?.sentStatus).toLowerCase() === "sent") {
        res.status(400).json({ error: "Sent guest records cannot be modified." });
        return;
      }
      await deleteGuestTableRow(id);
      const rows = await loadGuestTableRows();
      res.status(200).json({ rows, settings: current.settings, countries: COUNTRIES });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
