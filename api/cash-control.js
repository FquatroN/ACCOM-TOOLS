const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  CASH_CONTROL_SETTING_KEY,
  CASH_DENOMINATIONS,
  CASH_MIN_ALERT_DENOMINATIONS,
  calculateCashTotal,
  sanitizeCashControlPayload,
  sanitizeCashControlRecord,
  sanitizeCashControlRecords,
  validateCashControlRecord,
  normalizeCashStatus,
} = require("./_cash-control");
const { sanitizeBakerySettings } = require("./_bakery");

function cleanId(value) {
  return String(value || "").trim();
}

function cleanText(value) {
  return String(value || "").trim();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function formatCashDenominationLabel(value) {
  const amount = Number(value);
  if (!Number.isFinite(amount)) return `${value}€`;
  if (Math.abs(amount - Math.round(amount)) < 0.000001) return `${Math.round(amount)}€`;
  return `${String(amount).replace(".", ",")}€`;
}

function formatCashAlertDay(value) {
  const raw = cleanText(value);
  const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return raw;
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const monthLabel = monthNames[Math.max(0, Number(match[2]) - 1)] || "";
  return `${match[3]}${monthLabel}`;
}

function formatCashAlertShift(record) {
  const raw = cleanText(record?.shiftName || record?.shiftId).toLowerCase();
  if (raw === "night" || raw.startsWith("night")) return "N";
  if (raw === "morning" || raw.startsWith("morning")) return "M";
  if (raw === "afternoon" || raw.startsWith("afternoon")) return "T";
  return cleanText(record?.shiftName || record?.shiftId);
}

async function loadGeneralEmailConfig() {
  const rows = await restQuery("app_settings?select=payload&setting_key=eq.communications&limit=1", { method: "GET" });
  const payload = Array.isArray(rows) && rows[0]?.payload ? rows[0].payload : {};
  const generalEmailConfig = payload?.general?.emailConfig || payload?.general?.email_config || payload?.general?.bakeryEmailConfig || payload?.general?.bakery_email_config;
  return sanitizeBakerySettings({ emailConfig: generalEmailConfig || {} }).emailConfig;
}

async function sendWithResend({ to, subject, html, text }, emailConfig = {}) {
  const apiKey = process.env.RESEND_API_KEY;
  const rawFrom = process.env.EMAIL_FROM;
  const configuredFromEmail = cleanText(emailConfig.fromEmail || "").toLowerCase();
  const configuredFromName = cleanText(emailConfig.fromName || "");
  const replyTo = configuredFromEmail || "global@lisboacentralhostel.com";
  if (!apiKey) {
    const error = new Error("Missing server environment variable: RESEND_API_KEY");
    error.statusCode = 500;
    throw error;
  }
  if (!rawFrom) {
    const error = new Error("Missing server environment variable: EMAIL_FROM");
    error.statusCode = 500;
    throw error;
  }
  const envFromEmail = cleanText(rawFrom.replace(/^.*<([^>]+)>.*$/, "$1") || rawFrom).toLowerCase();
  const effectiveFromEmail = configuredFromEmail || envFromEmail;
  const effectiveFromName = configuredFromName || "ACCOM Tools - LCH";
  const from = `${effectiveFromName} <${effectiveFromEmail}>`;
  const response = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ from, reply_to: replyTo, to, subject, html, text }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload?.message || payload?.error || `Email provider failed (${response.status})`);
    error.statusCode = 502;
    throw error;
  }
  return payload;
}

function findLowCashDenominations(record, settings) {
  const minCash = settings?.minCash || {};
  return CASH_MIN_ALERT_DENOMINATIONS
    .map((key) => {
      const minimum = Math.max(0, Number(minCash?.[key] || 0));
      const quantity = Math.max(0, Number(record?.denominations?.[key] || 0));
      return { key, label: formatCashDenominationLabel(key), minimum, quantity };
    })
    .filter((item) => item.quantity < item.minimum);
}

function findHighCashDenominations(record, settings) {
  const maxCashByDenomination = settings?.maxCashByDenomination || {};
  return CASH_MIN_ALERT_DENOMINATIONS
    .map((key) => {
      const maximum = Math.max(0, Number(maxCashByDenomination?.[key] || 0));
      const quantity = Math.max(0, Number(record?.denominations?.[key] || 0));
      return { key, label: formatCashDenominationLabel(key), maximum, quantity };
    })
    .filter((item) => item.maximum > 0 && item.quantity >= item.maximum);
}

function buildLowCashAlertContent(items = []) {
  const subject = items.length === 1
    ? `Notas/Moedas de ${items[0].label} abaixo de ${items[0].minimum}`
    : "Notas/Moedas abaixo do minimo configurado";
  const bodyRows = items.map((item) => `<tr>
      <td style="border:1px solid #d8d0c7;padding:6px;">${escapeHtml(item.label)}</td>
      <td style="border:1px solid #d8d0c7;padding:6px;text-align:center;">${escapeHtml(String(item.quantity))}</td>
    </tr>`).join("");
  const html = `<!doctype html><html><body style="font-family:Arial,sans-serif;color:#1f2937;">
    <p>A quantidade das seguintes notas moedas precisa de ser reforcadas:</p>
    <table style="border-collapse:collapse;min-width:280px;">
      <thead>
        <tr>
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:left;">Nota/Moeda</th>
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:center;">Quantidade existente</th>
        </tr>
      </thead>
      <tbody>${bodyRows}</tbody>
    </table>
  </body></html>`;
  const text = [
    "A quantidade das seguintes notas moedas precisa de ser reforcadas:",
    "",
    "Nota/Moeda | Quantidade existente",
    ...items.map((item) => `${item.label} | ${item.quantity}`),
  ].join("\n");
  return { subject, html, text };
}

async function maybeSendLowCashAlert(record, settings) {
  const recipients = Array.isArray(settings?.managerAlertEmails) ? settings.managerAlertEmails.filter(Boolean) : [];
  if (!recipients.length || !settings?.minimumCashEmailEnabled) return null;
  const lowItems = findLowCashDenominations(record, settings);
  if (!lowItems.length) return null;
  const emailConfig = await loadGeneralEmailConfig();
  const mail = buildLowCashAlertContent(lowItems);
  return sendWithResend({ to: recipients, ...mail }, emailConfig);
}

function buildHighCashAlertContent(record, items = []) {
  const cashTotalValue = calculateCashTotal(record?.denominations || {});
  const cashTotal = Number(cashTotalValue || 0).toFixed(2).replace(".", ",");
  const noteList = CASH_DENOMINATIONS
    .map((item) => ({ label: formatCashDenominationLabel(item.key), quantity: Math.max(0, Number(record?.denominations?.[item.key] || 0)) }))
    .filter((item) => item.quantity > 0)
    .map((item) => `${item.label}: ${item.quantity}`)
    .join(", ");
  const dayLabel = formatCashAlertDay(record?.day);
  const shiftLabel = formatCashAlertShift(record);
  const subject = `Deposito Necessario - ${dayLabel}${shiftLabel ? ` ${shiftLabel}` : ""}`;
  const html = `<!doctype html><html><body style="font-family:Arial,sans-serif;color:#1f2937;">
    <p>O valor em caixa e ${escapeHtml(cashTotal)}€, existindo as seguintes notas: ${escapeHtml(noteList || "-")}.</p>
  </body></html>`;
  const text = `O valor em caixa e ${cashTotal}€, existindo as seguintes notas: ${noteList || "-"}.`;
  return { subject, html, text };
}

async function maybeSendHighCashAlert(record, settings) {
  const recipients = Array.isArray(settings?.managerAlertEmails) ? settings.managerAlertEmails.filter(Boolean) : [];
  if (!recipients.length || !settings?.maximumCashEmailEnabled) return null;
  const highItems = findHighCashDenominations(record, settings);
  const maximumCash = Math.max(0, Number(settings?.maximumCash || 0));
  const cashTotal = calculateCashTotal(record?.denominations || {});
  const exceedsCashLimit = maximumCash > 0 && cashTotal > maximumCash;
  if (!highItems.length && !exceedsCashLimit) return null;
  const emailConfig = await loadGeneralEmailConfig();
  const mail = buildHighCashAlertContent(record, highItems);
  return sendWithResend({ to: recipients, ...mail }, emailConfig);
}

function isMissingCashTableError(error) {
  const message = String(error?.message || "").toLowerCase();
  return message.includes("cash_control_records") && (
    message.includes("could not find") ||
    message.includes("schema cache") ||
    message.includes("relation") ||
    message.includes("does not exist")
  );
}

async function loadCashPayloadRow() {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${encodeURIComponent(CASH_CONTROL_SETTING_KEY)}&limit=1`, {
    method: "GET",
  });
  const row = Array.isArray(rows) && rows[0] ? rows[0] : null;
  return {
    rowId: row?.id || "",
    payload: sanitizeCashControlPayload(row?.payload || {}),
  };
}

async function saveCashPayload(rowId, payload) {
  const safe = sanitizeCashControlPayload(payload);
  if (rowId) {
    await restQuery(`app_settings?id=eq.${encodeURIComponent(rowId)}`, {
      method: "PATCH",
      body: { payload: safe, updated_at: new Date().toISOString() },
    });
    return safe;
  }
  const created = await restQuery("app_settings", {
    method: "POST",
    body: [{ setting_key: CASH_CONTROL_SETTING_KEY, payload: safe }],
  });
  return sanitizeCashControlPayload(Array.isArray(created) && created[0]?.payload ? created[0].payload : safe);
}

function mapCashTableRow(row, settings) {
  return sanitizeCashControlRecord({
    id: row?.id,
    day: row?.record_day ?? row?.day,
    shiftId: row?.shift_id,
    shiftName: row?.shift_name,
    status: row?.status,
    name: row?.name,
    denominations: row?.denominations,
    cardPos: row?.card_pos,
    cashFdm: row?.cash_fdm,
    cardFdm: row?.card_fdm,
    justification: row?.justification,
    itemCounts: row?.item_counts,
    itemJustifications: row?.item_justifications,
    createdAt: row?.created_at,
    updatedAt: row?.updated_at,
  }, settings);
}

async function loadCashTableRows(settings) {
  const rows = await restQuery("cash_control_records?select=*", { method: "GET" });
  return sanitizeCashControlRecords((Array.isArray(rows) ? rows : []).map((row) => mapCashTableRow(row, settings)), settings);
}

function buildCashTableBody(record, existing = {}) {
  return {
    id: cleanId(record.id || existing.id) || undefined,
    record_day: record.day,
    shift_id: record.shiftId,
    shift_name: record.shiftName,
    status: record.status,
    name: record.name,
    denominations: record.denominations,
    card_pos: Number(record.cardPos || 0),
    cash_fdm: Number(record.cashFdm || 0),
    card_fdm: Number(record.cardFdm || 0),
    justification: record.justification || "",
    item_counts: record.itemCounts || {},
    item_justifications: record.itemJustifications || {},
    created_at: existing.createdAt || record.createdAt || new Date().toISOString(),
    updated_at: new Date().toISOString(),
  };
}

async function createCashTableRows(records) {
  if (!Array.isArray(records) || records.length === 0) return;
  await restQuery("cash_control_records", {
    method: "POST",
    body: records.map((record) => buildCashTableBody(record)),
  });
}

async function updateCashTableRow(id, record, existing = {}) {
  await restQuery(`cash_control_records?id=eq.${encodeURIComponent(id)}`, {
    method: "PATCH",
    body: buildCashTableBody(record, existing),
    preferRepresentation: true,
  });
}

function mergeCashRecord(records, input, settings, id = "") {
  const existing = records.find((row) => cleanId(row.id) === cleanId(id)) || {};
  const nextRecord = sanitizeCashControlRecord({ ...existing, ...input, id: id || input?.id || existing.id }, settings, existing);
  validateCashControlRecord(nextRecord, records, settings, { excludeId: id || nextRecord.id, isCreate: !id });
  const nextRows = [...records];
  const index = id ? nextRows.findIndex((row) => cleanId(row.id) === cleanId(id)) : -1;
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
  return sanitizeCashControlRecords(nextRows, settings);
}

async function maybeSendCashAlertEmails(record, settings) {
  const errors = [];
  try {
    await maybeSendLowCashAlert(record, settings);
  } catch (emailError) {
    errors.push(emailError.message || "Could not send minimum cash alert email.");
  }
  try {
    await maybeSendHighCashAlert(record, settings);
  } catch (emailError) {
    errors.push(emailError.message || "Could not send maximum cash alert email.");
  }
  if (!errors.length) return null;
  return { error: errors.join(" | ") };
}

async function loadRecordsAndSettings() {
  const { payload } = await loadCashPayloadRow();
  const settings = payload.settings;
  try {
    const rows = await loadCashTableRows(settings);
    return { mode: "table", settings, rows };
  } catch (error) {
    if (!isMissingCashTableError(error)) throw error;
    return { mode: "legacy", settings, rows: payload.records };
  }
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "cash");

    if (req.method === "GET") {
      const current = await loadRecordsAndSettings();
      res.status(200).json({ rows: current.rows, settings: current.settings });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const current = await loadRecordsAndSettings();
      if (current.mode === "legacy") {
        const { rowId, payload } = await loadCashPayloadRow();
        const nextRows = mergeCashRecord(payload.records, body, payload.settings);
        const saved = await saveCashPayload(rowId, { settings: payload.settings, records: nextRows });
        res.status(200).json({ rows: saved.records, settings: saved.settings });
        return;
      }
      const nextRecord = sanitizeCashControlRecord({ ...body, status: "O" }, current.settings);
      validateCashControlRecord(nextRecord, current.rows, current.settings, { isCreate: true });
      await createCashTableRows([{
        ...nextRecord,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      }]);
      const rows = await loadCashTableRows(current.settings);
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
        const { rowId, payload } = await loadCashPayloadRow();
        const existing = payload.records.find((row) => cleanId(row.id) === id);
        if (!existing) {
          res.status(404).json({ error: "Record not found." });
          return;
        }
        const nextRows = mergeCashRecord(payload.records, body, payload.settings, id);
        const saved = await saveCashPayload(rowId, { settings: payload.settings, records: nextRows });
        const updatedRecord = saved.records.find((row) => cleanId(row.id) === id) || null;
        let alertEmailResult = null;
        if (updatedRecord && normalizeCashStatus(existing.status) !== "C" && normalizeCashStatus(updatedRecord.status) === "C") {
          alertEmailResult = await maybeSendCashAlertEmails(updatedRecord, saved.settings);
        }
        res.status(200).json({ rows: saved.records, settings: saved.settings, alertEmailResult });
        return;
      }
      const existing = current.rows.find((row) => cleanId(row.id) === id);
      if (!existing) {
        res.status(404).json({ error: "Record not found." });
        return;
      }
      const updated = sanitizeCashControlRecord({ ...existing, ...body, id: existing.id, status: body?.status ?? existing.status }, current.settings, existing);
      validateCashControlRecord(updated, current.rows, current.settings, { excludeId: existing.id, isCreate: false });
      await updateCashTableRow(existing.id, updated, existing);
      const rows = await loadCashTableRows(current.settings);
      const persisted = rows.find((row) => cleanId(row.id) === id) || updated;
      let alertEmailResult = null;
      if (normalizeCashStatus(existing.status) !== "C" && normalizeCashStatus(persisted.status) === "C") {
        alertEmailResult = await maybeSendCashAlertEmails(persisted, current.settings);
      }
      res.status(200).json({ rows, settings: current.settings, alertEmailResult });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
