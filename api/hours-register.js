const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  DEFAULT_HOURS_SETTINGS,
  DEFAULT_HOURS_RECORDS,
  HOURS_REGISTER_SETTING_KEY,
  hasPendingFinish,
  sanitizeHoursPayload,
  sanitizeHoursRecord,
  sanitizeHoursRecords,
  sanitizeHoursSettings,
} = require("./_hours-register");

function cleanId(value) {
  return String(value || "").trim();
}

function isMissingHoursTableError(error) {
  const message = String(error?.message || "").toLowerCase();
  return message.includes("hours_register_records") && (
    message.includes("could not find") ||
    message.includes("schema cache") ||
    message.includes("relation") ||
    message.includes("does not exist")
  );
}

function recordKey(record) {
  return [
    cleanId(record?.person).toLowerCase(),
    cleanId(record?.date),
    cleanId(record?.start),
    cleanId(record?.finish),
  ].join("::");
}

async function loadHoursPayloadRow() {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${encodeURIComponent(HOURS_REGISTER_SETTING_KEY)}&limit=1`, {
    method: "GET",
  });
  const row = Array.isArray(rows) && rows[0] ? rows[0] : null;
  return {
    rowId: row?.id || "",
    payload: sanitizeHoursPayload(row?.payload || { settings: DEFAULT_HOURS_SETTINGS, records: DEFAULT_HOURS_RECORDS }),
  };
}

async function saveHoursPayload(rowId, payload) {
  const safe = sanitizeHoursPayload(payload);
  if (rowId) {
    await restQuery(`app_settings?id=eq.${encodeURIComponent(rowId)}`, {
      method: "PATCH",
      body: { payload: safe, updated_at: new Date().toISOString() },
    });
    return safe;
  }
  const created = await restQuery("app_settings", {
    method: "POST",
    body: [{ setting_key: HOURS_REGISTER_SETTING_KEY, payload: safe }],
  });
  return sanitizeHoursPayload(Array.isArray(created) && created[0]?.payload ? created[0].payload : safe);
}

function mapHoursTableRow(row, settings) {
  return sanitizeHoursRecord({
    id: row?.id,
    person: row?.person,
    date: row?.record_date ?? row?.date,
    start: row?.start_time ?? row?.start,
    finish: row?.finish_time ?? row?.finish,
    createdAt: row?.created_at,
    updatedAt: row?.updated_at,
  }, settings);
}

async function loadHoursTableRows(settings) {
  const rows = await restQuery("hours_register_records?select=*", { method: "GET" });
  return sanitizeHoursRecords((Array.isArray(rows) ? rows : []).map((row) => mapHoursTableRow(row, settings)), settings);
}

function buildHoursTableBody(record, existing = {}) {
  return {
    id: record.id || existing.id,
    person: record.person,
    record_date: record.date,
    start_time: record.start,
    finish_time: record.finish || null,
    created_at: existing.createdAt || record.createdAt || new Date().toISOString(),
    updated_at: new Date().toISOString(),
  };
}

async function createHoursTableRow(record, existing = {}) {
  await restQuery("hours_register_records", {
    method: "POST",
    body: [buildHoursTableBody(record, existing)],
    preferRepresentation: true,
  });
}

async function updateHoursTableRow(id, record, existing = {}) {
  await restQuery(`hours_register_records?id=eq.${encodeURIComponent(id)}`, {
    method: "PATCH",
    body: buildHoursTableBody(record, existing),
    preferRepresentation: true,
  });
}

function validateHoursSave(existingRows, record, { excludeId = "", isCreate = false } = {}) {
  const duplicate = existingRows.find((row) => recordKey(row) === recordKey(record) && cleanId(row.id) !== cleanId(excludeId));
  if (duplicate) {
    const rangeText = record.finish ? ` from ${record.start} to ${record.finish}` : ` starting at ${record.start}`;
    const error = new Error(`A hours record for ${record.person} on ${record.date}${rangeText} already exists.`);
    error.statusCode = 400;
    throw error;
  }
  const otherPending = existingRows.find((row) => hasPendingFinish(row) && cleanId(row.id) !== cleanId(excludeId));
  if (isCreate && otherPending) {
    const error = new Error(`Finish time is still missing for ${otherPending.person} on ${otherPending.date}. Please fill it before adding a new record.`);
    error.statusCode = 400;
    throw error;
  }
  if (!isCreate && hasPendingFinish(record) && otherPending) {
    const error = new Error(`There is already another hours record waiting for finish time on ${otherPending.date}.`);
    error.statusCode = 400;
    throw error;
  }
}

function mergeLegacyRecords(existingRecords, incomingRecord, settings, id = "") {
  const safe = sanitizeHoursRecord(incomingRecord, settings, existingRecords.find((item) => cleanId(item.id) === cleanId(id)) || {});
  const next = [...existingRecords];
  validateHoursSave(next, safe, { excludeId: id || safe.id, isCreate: !id });
  const index = id ? next.findIndex((item) => cleanId(item.id) === cleanId(id)) : -1;
  if (index >= 0) {
    next[index] = {
      ...safe,
      id: next[index].id,
      createdAt: next[index].createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
  } else {
    next.push({
      ...safe,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });
  }
  return sanitizeHoursRecords(next, settings);
}

async function loadRecordsAndSettings() {
  const { payload } = await loadHoursPayloadRow();
  const settings = payload.settings;
  try {
    const rows = await loadHoursTableRows(settings);
    return { mode: "table", settings, rows };
  } catch (error) {
    if (!isMissingHoursTableError(error)) throw error;
    return { mode: "legacy", settings, rows: payload.records };
  }
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "hours");

    if (req.method === "GET") {
      const current = await loadRecordsAndSettings();
      res.status(200).json({ rows: current.rows, settings: current.settings });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const current = await loadRecordsAndSettings();
      if (current.mode === "legacy") {
        const { rowId, payload } = await loadHoursPayloadRow();
        const nextRecords = mergeLegacyRecords(payload.records, body, payload.settings);
        const saved = await saveHoursPayload(rowId, { settings: payload.settings, records: nextRecords });
        res.status(200).json({ rows: saved.records, settings: saved.settings });
        return;
      }
      const nextRecord = sanitizeHoursRecord(body, current.settings);
      validateHoursSave(current.rows, nextRecord, { isCreate: true });
      await createHoursTableRow({
        ...nextRecord,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      });
      const rows = await loadHoursTableRows(current.settings);
      res.status(200).json({ rows, settings: current.settings });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanId(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Record id is required." });
        return;
      }
      const body = await parseBody(req);
      const current = await loadRecordsAndSettings();
      if (current.mode === "legacy") {
        const { rowId, payload } = await loadHoursPayloadRow();
        const existing = payload.records.find((item) => cleanId(item.id) === id);
        if (!existing) {
          res.status(404).json({ error: "Record not found." });
          return;
        }
        const nextRecords = mergeLegacyRecords(payload.records, { ...existing, ...body, id: existing.id }, payload.settings, id);
        const saved = await saveHoursPayload(rowId, { settings: payload.settings, records: nextRecords });
        res.status(200).json({ rows: saved.records, settings: saved.settings });
        return;
      }
      const existing = current.rows.find((item) => cleanId(item.id) === id);
      if (!existing) {
        res.status(404).json({ error: "Record not found." });
        return;
      }
      const updated = sanitizeHoursRecord({ ...existing, ...body, id: existing.id }, current.settings, existing);
      validateHoursSave(current.rows, updated, { excludeId: existing.id, isCreate: false });
      await updateHoursTableRow(existing.id, updated, existing);
      const rows = await loadHoursTableRows(current.settings);
      res.status(200).json({ rows, settings: current.settings });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
