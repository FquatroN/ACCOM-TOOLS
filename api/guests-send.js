const { requireFeature, restQuery, sendError } = require("./_supabase");
const http = require("node:http");
const https = require("node:https");
const {
  DEFAULT_GUESTS_INTEGRATION_MAPPING,
  DEFAULT_GUESTS_SETTINGS,
  GUESTS_SETTING_KEY,
  SEF_ENDPOINT,
  SEF_ENDPOINTS,
  buildBalXml,
  buildSoapEnvelope,
  lisbonTodayIso,
  resolveCountry,
  resolveSefConfig,
  sanitizeGuestRecord,
  sanitizeGuestsPayload,
} = require("./_guests");

function cleanText(value) {
  return String(value || "").trim();
}

function decodeXmlEntities(value) {
  return String(value || "")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"')
    .replace(/&apos;/gi, "'");
}

function isMissingGuestApiCallsTableError(error) {
  const message = String(error?.message || "").toLowerCase();
  return message.includes("guest_api_calls") && (
    message.includes("could not find") ||
    message.includes("schema cache") ||
    message.includes("relation") ||
    message.includes("does not exist")
  );
}

function isMissingGuestApiCallsColumnError(error, columnName) {
  const message = String(error?.message || "").toLowerCase();
  return message.includes(columnName.toLowerCase()) && message.includes("guest_api_calls") && message.includes("schema cache");
}

function maskSefSoap(soap) {
  return String(soap || "").replace(/(<ChaveAcesso>)([\s\S]*?)(<\/ChaveAcesso>)/i, "$1***$3");
}

function normalizeResponseHeaders(headers) {
  const source = headers && typeof headers === "object" ? headers : {};
  return Object.fromEntries(
    Object.entries(source)
      .map(([key, value]) => [cleanText(key).toLowerCase(), Array.isArray(value) ? value.join(", ") : cleanText(value)])
      .filter(([key, value]) => key && value)
  );
}

async function saveGuestApiCallLog(entry) {
  const payload = {
    endpoint: cleanText(entry?.endpoint),
    request_method: cleanText(entry?.requestMethod) || "POST",
    soap_action: cleanText(entry?.soapAction) || "http://sef.pt/EntregaBoletinsAlojamento",
    http_status: Math.max(0, Number.parseInt(entry?.httpStatus, 10) || 0),
    file_number: Math.max(0, Number.parseInt(entry?.fileNumber, 10) || 0),
    guest_count: Math.max(0, Number.parseInt(entry?.guestCount, 10) || 0),
    success: !!entry?.success,
    response_message: cleanText(entry?.responseMessage),
    error_message: cleanText(entry?.errorMessage),
    request_details: entry?.requestDetails && typeof entry.requestDetails === "object" ? entry.requestDetails : {},
    request_body: String(entry?.requestBody || "").slice(0, 200000),
    response_body: String(entry?.responseBody || "").slice(0, 200000),
    response_headers: normalizeResponseHeaders(entry?.responseHeaders),
  };
  try {
    await restQuery("guest_api_calls", {
      method: "POST",
      body: [payload],
    });
  } catch (error) {
    if (isMissingGuestApiCallsColumnError(error, "response_headers")) {
      const fallbackPayload = { ...payload };
      delete fallbackPayload.response_headers;
      await restQuery("guest_api_calls", {
        method: "POST",
        body: [fallbackPayload],
      });
      return;
    }
    if (!isMissingGuestApiCallsTableError(error)) throw error;
  }
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
  const rows = await restQuery("guest_records?select=*&order=created_at.desc,name.asc", {
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
  if (!cleanText(record.residenceCity)) return "Residence City is required for SEF send.";
  return "";
}

function getMappedGuestValue(record, settings, key) {
  const sourceKey = cleanText(settings?.integrationMapping?.[key]) || DEFAULT_GUESTS_INTEGRATION_MAPPING[key] || "";
  return sourceKey ? record?.[sourceKey] : "";
}

function buildGuestForSend(record, settings) {
  const nationality = resolveCountry(getMappedGuestValue(record, settings, "nationality"));
  const issuerCountry = resolveCountry(getMappedGuestValue(record, settings, "issuerCountry"));
  const residenceCountrySource = getMappedGuestValue(record, settings, "residenceCountry");
  const residenceCountry = resolveCountry(residenceCountrySource);
  const residenceCitySource = getMappedGuestValue(record, settings, "residenceCity");
  const residenceCityCountry = resolveCountry(residenceCitySource);
  return {
    ...record,
    name: cleanText(getMappedGuestValue(record, settings, "name")) || record.name,
    nationality: nationality.input,
    nationalityCode: nationality.code,
    birthDate: cleanText(getMappedGuestValue(record, settings, "birthDate")) || record.birthDate,
    birthPlace: "",
    docNumber: cleanText(getMappedGuestValue(record, settings, "docNumber")) || record.docNumber,
    docType: cleanText(getMappedGuestValue(record, settings, "docType")) || record.docType,
    issuerCountry: issuerCountry.input,
    issuerCountryCode: issuerCountry.code,
    residenceCountry: residenceCountry.input,
    residenceCountryCode: residenceCountry.code,
    residenceCity: residenceCityCountry.name || cleanText(residenceCitySource),
    checkIn: cleanText(getMappedGuestValue(record, settings, "checkIn")) || record.checkIn,
    checkOut: cleanText(getMappedGuestValue(record, settings, "checkOut")) || record.checkOut,
  };
}

function parseSefResult(resultText) {
  const safe = cleanText(decodeXmlEntities(resultText));
  if (safe === "0") return { ok: true, message: "0" };
  if (!safe) {
    return { ok: false, message: "SEF returned an empty response." };
  }
  const faultMatch = safe.match(/<(?:[\w-]+:)?faultstring(?:\s[^>]*)?>([\s\S]*?)<\/(?:[\w-]+:)?faultstring>/i);
  if (faultMatch?.[1]) {
    return { ok: false, message: cleanText(decodeXmlEntities(faultMatch[1])) || "SEF SOAP fault" };
  }
  const codeMatch = safe.match(/<Codigo_Retorno>(.*?)<\/Codigo_Retorno>/i);
  const descMatch = safe.match(/<Descricao>(.*?)<\/Descricao>/i);
  const statusTextMatch = safe.match(/<(?:[\w-]+:)?(?:EntregaBoletinsAlojamentoResult|Resultado)(?:\s[^>]*)?>([\s\S]*?)<\/(?:[\w-]+:)?(?:EntregaBoletinsAlojamentoResult|Resultado)>/i);
  const statusText = cleanText(decodeXmlEntities(statusTextMatch?.[1] || ""));
  if (statusText === "0") return { ok: true, message: "0" };
  return {
    ok: false,
    message: [codeMatch?.[1], descMatch?.[1]].filter(Boolean).join(" - ")
      || statusText
      || `Unknown SEF response: ${safe.slice(0, 240)}`,
  };
}

async function postToSef(soap, settings) {
  const endpoints = Array.isArray(SEF_ENDPOINTS) && SEF_ENDPOINTS.length ? SEF_ENDPOINTS : [SEF_ENDPOINT];
  const sefConfig = resolveSefConfig(settings);
  const failures = [];
  for (const endpoint of endpoints) {
    try {
      const result = await new Promise((resolve, reject) => {
        const url = new URL(endpoint);
        const client = url.protocol === "https:" ? https : http;
        const request = client.request(url, {
          method: "POST",
          headers: {
            "Content-Type": "text/xml; charset=utf-8",
            SOAPAction: "http://sef.pt/EntregaBoletinsAlojamento",
            "Content-Length": Buffer.byteLength(soap, "utf8"),
          },
          ...(url.protocol === "https:" && sefConfig.caCertificate ? { ca: sefConfig.caCertificate } : {}),
        }, (response) => {
          let body = "";
          response.setEncoding("utf8");
          response.on("data", (chunk) => {
            body += chunk;
          });
          response.on("end", () => resolve({
            statusCode: response.statusCode || 0,
            body,
            headers: normalizeResponseHeaders(response.headers),
          }));
        });
        request.on("error", reject);
        request.write(soap);
        request.end();
      });
      if ([301, 302, 307, 308].includes(result.statusCode)) {
        const location = cleanText(result.headers?.location);
        failures.push(`${endpoint}: HTTP ${result.statusCode}${location ? ` redirect to ${location}` : " redirect with no location header"}`);
        continue;
      }
      return { ...result, endpoint };
    } catch (error) {
      failures.push(`${endpoint}: ${cleanText(error?.cause?.message || error?.message || "fetch failed")}`);
    }
  }
  const joined = failures.join(" | ") || "fetch failed";
  const hint = joined.toLowerCase().includes("certificate") && !sefConfig.caCertificate
    ? " Configure the SEF Root CA certificate in Guests Settings -> SEF Credentials."
    : "";
  const err = new Error(`Could not reach SEF endpoint. ${joined}.${hint}`.replace(/\.\s*\./g, "."));
  err.statusCode = 502;
  throw err;
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

    const sendableMapped = sendable.map((row) => buildGuestForSend(row, current.settings));

    const validationErrors = sendableMapped
      .map((row) => ({ row, error: validateSendableGuest(row) }))
      .filter((item) => item.error);
    if (validationErrors.length) {
      const first = validationErrors[0];
      res.status(400).json({ error: `${first.row.name}: ${first.error}` });
      return;
    }

    const sefConfig = resolveSefConfig(current.settings);
    if (!cleanText(sefConfig.unitCode) || !cleanText(sefConfig.establishment) || !cleanText(sefConfig.accessKey)) {
      res.status(400).json({ error: "SEF credentials are required in Guests Settings." });
      return;
    }

    const nextFileNumber = Math.max(1, Number(current.payload.lastFileNumber || 0) + 1);
    const xml = buildBalXml(sendableMapped, nextFileNumber, current.settings);
    const base64Payload = Buffer.from(xml, "utf-8").toString("base64");
    const soap = buildSoapEnvelope(base64Payload, current.settings);
    const maskedSoap = maskSefSoap(soap);
    const requestDetails = {
      endpoints: Array.isArray(SEF_ENDPOINTS) ? SEF_ENDPOINTS : [SEF_ENDPOINT],
      requestMethod: "POST",
      soapAction: "http://sef.pt/EntregaBoletinsAlojamento",
      fileNumber: nextFileNumber,
      guestCount: sendableMapped.length,
      guestNames: sendableMapped.map((row) => cleanText(row.name)).filter(Boolean),
      unitCode: cleanText(sefConfig.unitCode),
      establishment: cleanText(sefConfig.establishment),
    };
    let body = "";
    let statusCode = 0;
    let endpoint = "";
    let responseHeaders = {};
    let result;
    try {
      const postResult = await postToSef(soap, current.settings);
      body = String(postResult?.body || "");
      statusCode = Number.parseInt(postResult?.statusCode, 10) || 0;
      endpoint = cleanText(postResult?.endpoint);
      responseHeaders = normalizeResponseHeaders(postResult?.headers);
      const resultMatch = body.match(/<(?:[\w-]+:)?EntregaBoletinsAlojamentoResult(?:\s[^>]*)?>([\s\S]*?)<\/(?:[\w-]+:)?EntregaBoletinsAlojamentoResult>/i);
      const resultText = resultMatch ? decodeXmlEntities(resultMatch[1]) : body;
      result = parseSefResult(resultText);
      await saveGuestApiCallLog({
        endpoint,
        requestMethod: "POST",
        soapAction: "http://sef.pt/EntregaBoletinsAlojamento",
        httpStatus: statusCode,
        fileNumber: nextFileNumber,
        guestCount: sendableMapped.length,
        success: result.ok,
        responseMessage: result.message,
        errorMessage: result.ok ? "" : result.message,
        requestDetails,
        requestBody: maskedSoap,
        responseBody: body,
        responseHeaders,
      });
    } catch (error) {
      await saveGuestApiCallLog({
        endpoint,
        requestMethod: "POST",
        soapAction: "http://sef.pt/EntregaBoletinsAlojamento",
        httpStatus: statusCode || Number.parseInt(error?.statusCode, 10) || 0,
        fileNumber: nextFileNumber,
        guestCount: sendableMapped.length,
        success: false,
        responseMessage: "",
        errorMessage: cleanText(error?.message || "Unknown SEF transport error."),
        requestDetails,
        requestBody: maskedSoap,
        responseBody: body,
        responseHeaders,
      });
      throw error;
    }
    if (!result.ok && result.message.startsWith("Unknown SEF response:")) {
      result.message = `${result.message} [HTTP ${statusCode || 0} via ${endpoint}]`;
    }
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
    res.status(200).json({ rows: savedRows, settings: savedSettings, sent: sendableMapped.length, message: "Guests sent successfully." });
  } catch (error) {
    sendError(res, error);
  }
};
