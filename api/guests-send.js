const { requireFeature, restQuery, sendError } = require("./_supabase");
const {
  DEFAULT_GUESTS_SETTINGS,
  GUESTS_SETTING_KEY,
  SEF_ENDPOINT,
  buildBalXml,
  buildSoapEnvelope,
  lisbonTodayIso,
  sanitizeGuestRecord,
  sanitizeGuestsPayload,
} = require("./_guests");

function cleanText(value) {
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

function cleanId(value) {
  return String(value || "").trim();
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

async function updateGuestSendStatuses(rows, { ok, nextFileNumber, message, timestamp }) {
  const updates = rows.map((row) => restQuery(`guest_records?id=eq.${encodeURIComponent(cleanId(row.id))}`, {
    method: "PATCH",
    body: ok ? {
      sent_status: "sent",
      sent_at: timestamp,
      send_error: "",
      send_batch_number: nextFileNumber,
      updated_at: timestamp,
    } : {
      sent_status: "error",
      send_error: message,
      updated_at: timestamp,
    },
  }));
  await Promise.all(updates);
}

async function loadRowsAndSettings() {
  const { rowId, payload } = await loadGuestsPayloadRow();
  try {
    const rows = await loadGuestTableRows();
    return { mode: "table", rowId, payload, settings: payload.settings, rows };
  } catch (error) {
    if (!isMissingGuestsTableError(error)) throw error;
    return { mode: "legacy", rowId, payload, settings: payload.settings, rows: payload.rows };
  }
}

function validateSendableGuest(record) {
  if (!record.nationalityCode) return "Nationality must match a valid ICAO country.";
  if (!record.issuerCountryCode) return "Issuer Country must match a valid ICAO country.";
  if (!record.residenceCountryCode) return "Residence Country must match a valid ICAO country.";
  if (!cleanText(record.birthPlace)) return "Birth Place is required for SEF send.";
  if (!cleanText(record.residenceCity)) return "Residence City is required for SEF send.";
  return "";
}

function parseSefResult(resultText) {
  const safe = cleanText(resultText);
  if (safe === "0") return { ok: true, message: "0" };
  const codeMatch = safe.match(/<Codigo_Retorno>(.*?)<\/Codigo_Retorno>/i);
  const descMatch = safe.match(/<Descricao>(.*?)<\/Descricao>/i);
  return {
    ok: false,
    message: [codeMatch?.[1], descMatch?.[1]].filter(Boolean).join(" - ") || safe || "Unknown SEF response",
  };
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "guests");
    if (req.method !== "POST") {
      res.status(405).json({ error: "Method not allowed." });
      return;
    }

    const current = await loadRowsAndSettings();
    const today = lisbonTodayIso();
    const sendable = current.rows.filter((row) => cleanText(row.sentStatus) !== "sent" && cleanText(row.checkIn) && cleanText(row.checkIn) <= today);
    if (!sendable.length) {
      res.status(200).json({ rows: current.rows, settings: current.settings, sent: 0, message: "No pending guest records ready to send." });
      return;
    }

    const validationErrors = sendable
      .map((row) => ({ row, error: validateSendableGuest(row) }))
      .filter((item) => item.error);
    if (validationErrors.length) {
      const first = validationErrors[0];
      res.status(400).json({ error: `${first.row.name}: ${first.error}` });
      return;
    }

    const nextFileNumber = Math.max(1, Number(current.payload.lastFileNumber || 0) + 1);
    const xml = buildBalXml(sendable, nextFileNumber);
    const base64Payload = Buffer.from(xml, "utf-8").toString("base64");
    const soap = buildSoapEnvelope(base64Payload);
    const response = await fetch(SEF_ENDPOINT, {
      method: "POST",
      headers: {
        "Content-Type": "text/xml; charset=utf-8",
        SOAPAction: "http://sef.pt/EntregaBoletinsAlojamento",
      },
      body: soap,
    });
    const body = await response.text();
    const resultMatch = body.match(/<EntregaBoletinsAlojamentoResult>([\s\S]*?)<\/EntregaBoletinsAlojamentoResult>/i);
    const resultText = resultMatch ? resultMatch[1].replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&") : body;
    const result = parseSefResult(resultText);
    const timestamp = new Date().toISOString();
    let savedRows = current.rows;
    let savedSettings = current.settings;
    if (current.mode === "table") {
      await updateGuestSendStatuses(sendable, {
        ok: result.ok,
        nextFileNumber,
        message: result.message,
        timestamp,
      });
      savedRows = await loadGuestTableRows();
      const savedPayload = await saveGuestsPayload(current.rowId, {
        ...current.payload,
        lastFileNumber: result.ok ? nextFileNumber : current.payload.lastFileNumber,
      });
      savedSettings = savedPayload.settings;
    } else {
      const attemptedIds = new Set(sendable.map((row) => cleanId(row.id)));
      const nextRows = current.payload.rows.map((row) => {
        if (!attemptedIds.has(cleanId(row.id))) return row;
        if (result.ok) {
          return {
            ...row,
            sentStatus: "sent",
            sentAt: timestamp,
            sendError: "",
            sendBatchNumber: nextFileNumber,
            updatedAt: timestamp,
          };
        }
        return {
          ...row,
          sentStatus: "error",
          sendError: result.message,
          updatedAt: timestamp,
        };
      });
      const saved = await saveGuestsPayload(current.rowId, {
        ...current.payload,
        rows: nextRows,
        lastFileNumber: result.ok ? nextFileNumber : current.payload.lastFileNumber,
      });
      savedRows = saved.rows;
      savedSettings = saved.settings;
    }

    if (!result.ok) {
      res.status(502).json({ error: result.message, rows: savedRows, settings: savedSettings, sent: 0 });
      return;
    }
    res.status(200).json({ rows: savedRows, settings: savedSettings, sent: sendable.length, message: "Guests sent successfully." });
  } catch (error) {
    sendError(res, error);
  }
};
