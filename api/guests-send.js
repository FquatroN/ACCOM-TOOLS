const { requireFeature, restQuery, sendError } = require("./_supabase");
const {
  DEFAULT_GUESTS_SETTINGS,
  GUESTS_SETTING_KEY,
  SEF_ENDPOINT,
  buildBalXml,
  buildSoapEnvelope,
  lisbonTodayIso,
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

    const { rowId, payload } = await loadGuestsPayloadRow();
    const today = lisbonTodayIso();
    const sendable = payload.rows.filter((row) => cleanText(row.sentStatus) !== "sent" && cleanText(row.checkIn) && cleanText(row.checkIn) <= today);
    if (!sendable.length) {
      res.status(200).json({ rows: payload.rows, settings: payload.settings, sent: 0, message: "No pending guest records ready to send." });
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

    const nextFileNumber = Math.max(1, Number(payload.lastFileNumber || 0) + 1);
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
    const attemptedIds = new Set(sendable.map((row) => cleanText(row.id)));
    const nextRows = payload.rows.map((row) => {
      if (!attemptedIds.has(cleanText(row.id))) return row;
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
    const saved = await saveGuestsPayload(rowId, {
      ...payload,
      rows: nextRows,
      lastFileNumber: result.ok ? nextFileNumber : payload.lastFileNumber,
    });

    if (!result.ok) {
      res.status(502).json({ error: result.message, rows: saved.rows, settings: saved.settings, sent: 0 });
      return;
    }
    res.status(200).json({ rows: saved.rows, settings: saved.settings, sent: sendable.length, message: "Guests sent successfully." });
  } catch (error) {
    sendError(res, error);
  }
};
