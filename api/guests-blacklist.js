const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const { COUNTRIES, DEFAULT_GUESTS_SETTINGS, GUESTS_SETTING_KEY, sanitizeBlacklistRecord, sanitizeGuestsPayload } = require("./_guests");

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

function isMissingGuestsBlacklistTableError(error) {
  const message = String(error?.message || "").toLowerCase();
  return message.includes("guests_blacklist") && (
    message.includes("could not find") ||
    message.includes("schema cache") ||
    message.includes("relation") ||
    message.includes("does not exist")
  );
}

function mapBlacklistTableRow(row) {
  return sanitizeBlacklistRecord({
    id: row?.id,
    name: row?.name,
    nationality: row?.nationality,
    nationalityCode: row?.nationality_code,
    birthDate: row?.birth_date,
    docNumber: row?.doc_number,
    whatHappened: row?.what_happened,
    occurrenceDate: row?.occurrence_date,
    whoReported: row?.who_reported,
    createdAt: row?.created_at,
    updatedAt: row?.updated_at,
  });
}

async function loadBlacklistTableRows() {
  const rows = await restQuery("guests_blacklist?select=*&order=occurrence_date.desc,name.asc", {
    method: "GET",
  });
  return (Array.isArray(rows) ? rows : []).map(mapBlacklistTableRow);
}

function buildBlacklistTableBody(record, existing = {}) {
  return {
    id: cleanId(record.id || existing.id) || undefined,
    name: record.name,
    nationality: record.nationality || "",
    nationality_code: record.nationalityCode || "",
    birth_date: record.birthDate || null,
    doc_number: record.docNumber || "",
    what_happened: record.whatHappened || "",
    occurrence_date: record.occurrenceDate,
    who_reported: record.whoReported || "",
    created_at: existing.createdAt || record.createdAt || new Date().toISOString(),
    updated_at: new Date().toISOString(),
  };
}

async function createBlacklistTableRow(record, existing = {}) {
  await restQuery("guests_blacklist", {
    method: "POST",
    body: [buildBlacklistTableBody(record, existing)],
    preferRepresentation: true,
  });
}

async function updateBlacklistTableRow(id, record, existing = {}) {
  await restQuery(`guests_blacklist?id=eq.${encodeURIComponent(id)}`, {
    method: "PATCH",
    body: buildBlacklistTableBody(record, existing),
    preferRepresentation: true,
  });
}

async function deleteBlacklistTableRow(id) {
  await restQuery(`guests_blacklist?id=eq.${encodeURIComponent(id)}`, {
    method: "DELETE",
  });
}

async function loadBlacklistRows() {
  const { payload } = await loadGuestsPayloadRow();
  try {
    const rows = await loadBlacklistTableRows();
    return { mode: "table", rows };
  } catch (error) {
    if (!isMissingGuestsBlacklistTableError(error)) throw error;
    return { mode: "legacy", rows: payload.blacklist };
  }
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "guests");

    if (req.method === "GET") {
      const current = await loadBlacklistRows();
      res.status(200).json({ rows: current.rows, countries: COUNTRIES });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const current = await loadBlacklistRows();
      if (current.mode === "legacy") {
        const { rowId, payload } = await loadGuestsPayloadRow();
        const record = sanitizeBlacklistRecord(body);
        const nextRows = sanitizeGuestsPayload({
          ...payload,
          blacklist: [...payload.blacklist, { ...record, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() }],
        }).blacklist;
        const saved = await saveGuestsPayload(rowId, { ...payload, blacklist: nextRows });
        res.status(200).json({ rows: saved.blacklist, countries: COUNTRIES });
        return;
      }
      const record = sanitizeBlacklistRecord(body);
      await createBlacklistTableRow({
        ...record,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      });
      const rows = await loadBlacklistTableRows();
      res.status(200).json({ rows, countries: COUNTRIES });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanId(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Blacklist id is required." });
        return;
      }
      const body = await parseBody(req);
      const current = await loadBlacklistRows();
      if (current.mode === "legacy") {
        const { rowId, payload } = await loadGuestsPayloadRow();
        const existing = payload.blacklist.find((item) => cleanId(item.id) === id);
        if (!existing) {
          res.status(404).json({ error: "Blacklist record not found." });
          return;
        }
        const record = sanitizeBlacklistRecord({ ...existing, ...body, id }, existing);
        const nextRows = payload.blacklist.map((item) => (cleanId(item.id) === id ? { ...record, createdAt: existing.createdAt || item.createdAt || new Date().toISOString(), updatedAt: new Date().toISOString() } : item));
        const saved = await saveGuestsPayload(rowId, { ...payload, blacklist: nextRows });
        res.status(200).json({ rows: saved.blacklist, countries: COUNTRIES });
        return;
      }
      const existing = current.rows.find((item) => cleanId(item.id) === id);
      if (!existing) {
        res.status(404).json({ error: "Blacklist record not found." });
        return;
      }
      const record = sanitizeBlacklistRecord({ ...existing, ...body, id }, existing);
      await updateBlacklistTableRow(id, record, existing);
      const rows = await loadBlacklistTableRows();
      res.status(200).json({ rows, countries: COUNTRIES });
      return;
    }

    if (req.method === "DELETE") {
      const id = cleanId(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Blacklist id is required." });
        return;
      }
      const current = await loadBlacklistRows();
      if (current.mode === "legacy") {
        const { rowId, payload } = await loadGuestsPayloadRow();
        const nextRows = payload.blacklist.filter((item) => cleanId(item.id) !== id);
        const saved = await saveGuestsPayload(rowId, { ...payload, blacklist: nextRows });
        res.status(200).json({ rows: saved.blacklist, countries: COUNTRIES });
        return;
      }
      await deleteBlacklistTableRow(id);
      const rows = await loadBlacklistTableRows();
      res.status(200).json({ rows, countries: COUNTRIES });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
