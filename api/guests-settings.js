const { hasFeature, loadAccessForUser, parseBody, requireFeature, restQuery, sendError, verifyUser } = require("./_supabase");
const { COUNTRIES, DEFAULT_GUESTS_SETTINGS, GUESTS_SETTING_KEY, sanitizeGuestsPayload, normalizeGuestSettings } = require("./_guests");

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

function isMissingGuestApiCallsTableError(error) {
  const message = String(error?.message || "").toLowerCase();
  return message.includes("guest_api_calls") && (
    message.includes("could not find") ||
    message.includes("schema cache") ||
    message.includes("relation") ||
    message.includes("does not exist")
  );
}

function mapGuestApiCallRow(row) {
  return {
    id: String(row?.id || "").trim(),
    createdAt: String(row?.created_at || "").trim(),
    endpoint: String(row?.endpoint || "").trim(),
    requestMethod: String(row?.request_method || "").trim(),
    soapAction: String(row?.soap_action || "").trim(),
    httpStatus: Number.parseInt(row?.http_status, 10) || 0,
    fileNumber: Number.parseInt(row?.file_number, 10) || 0,
    guestCount: Number.parseInt(row?.guest_count, 10) || 0,
    success: !!row?.success,
    responseMessage: String(row?.response_message || "").trim(),
    errorMessage: String(row?.error_message || "").trim(),
    requestDetails: row?.request_details && typeof row.request_details === "object" ? row.request_details : {},
    requestBody: String(row?.request_body || ""),
    responseBody: String(row?.response_body || ""),
  };
}

async function loadGuestApiCalls() {
  const rows = await restQuery("guest_api_calls?select=*&order=created_at.desc&limit=100", {
    method: "GET",
  });
  return (Array.isArray(rows) ? rows : []).map(mapGuestApiCallRow);
}

module.exports = async function handler(req, res) {
  try {
    if (req.method === "GET") {
      const user = await verifyUser(req);
      const access = await loadAccessForUser(user.id);
      const canSettings = hasFeature(access, "settings", "guests");
      const canApp = hasFeature(access, "app", "guests");
      if (!canSettings && !canApp) {
        const err = new Error("You do not have permission for this feature.");
        err.statusCode = 403;
        throw err;
      }
      const { payload } = await loadGuestsPayloadRow();
      let apiCalls = [];
      let apiCallsEnabled = false;
      if (canSettings) {
        try {
          apiCalls = await loadGuestApiCalls();
          apiCallsEnabled = true;
        } catch (error) {
          if (!isMissingGuestApiCallsTableError(error)) throw error;
        }
      }
      res.status(200).json({ settings: payload.settings, countries: COUNTRIES, apiCalls, apiCallsEnabled });
      return;
    }

    if (req.method === "PUT") {
      await requireFeature(req, "settings", "guests");
      const body = await parseBody(req);
      const { rowId, payload } = await loadGuestsPayloadRow();
      const next = {
        ...payload,
        settings: normalizeGuestSettings(body?.settings),
      };
      const saved = await saveGuestsPayload(rowId, next);
      let apiCalls = [];
      let apiCallsEnabled = false;
      try {
        apiCalls = await loadGuestApiCalls();
        apiCallsEnabled = true;
      } catch (error) {
        if (!isMissingGuestApiCallsTableError(error)) throw error;
      }
      res.status(200).json({ settings: saved.settings, countries: COUNTRIES, apiCalls, apiCallsEnabled });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
