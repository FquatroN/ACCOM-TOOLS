const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  DEFAULT_LAUNDRY_SETTINGS,
  LAUNDRY_SETTING_KEY,
  sanitizeLaundryPayload,
  sanitizeLaundryRecord,
  sanitizeLaundrySettings,
} = require("./_laundry");
const { sanitizeBakerySettings, cleanText } = require("./_bakery");
const { sendWithSmtp } = require("./_smtp");

const LAST_SENT_KEY = "laundry_email_automation_last_sent";
const MANAGEMENT_LAST_SENT_KEY = "laundry_management_email_automation_last_sent";
const DEFAULT_TZ = process.env.AUTOMATION_TIMEZONE || "Europe/Lisbon";
const PROPERTIES = ["Cruz", "Hostel"];

function normalizeTime(value) {
  const raw = String(value || "").trim();
  return /^\d{2}:\d{2}$/.test(raw) ? raw : "00:00";
}

function getClockParts(date, timeZone) {
  const fmt = new Intl.DateTimeFormat("en-GB", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hourCycle: "h23",
  });
  const parts = fmt.formatToParts(date);
  const map = {};
  parts.forEach((part) => {
    if (part.type !== "literal") map[part.type] = part.value;
  });
  return {
    year: Number(map.year),
    month: Number(map.month),
    day: Number(map.day),
    hour: Number(map.hour),
    minute: Number(map.minute),
  };
}

function shouldSendNow(timeOfDay, now, timeZone) {
  const [hourRaw, minuteRaw] = normalizeTime(timeOfDay).split(":");
  const clock = getClockParts(now, timeZone);
  const hour = Number(hourRaw);
  const minute = Number(minuteRaw);
  if (clock.hour !== hour || clock.minute !== minute) return { due: false };
  return { due: true, slotKey: `${clock.year}-${clock.month}-${clock.day}:${clock.hour}:${clock.minute}` };
}

function lisbonTodayIso(now = new Date(), timeZone = DEFAULT_TZ) {
  const parts = getClockParts(now, timeZone);
  return `${parts.year}-${String(parts.month).padStart(2, "0")}-${String(parts.day).padStart(2, "0")}`;
}

function shiftIsoDate(value, days) {
  const date = new Date(`${String(value || "").trim()}T00:00:00`);
  if (Number.isNaN(date.getTime())) return "";
  date.setDate(date.getDate() + days);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatDatePt(value) {
  const date = new Date(`${String(value || "").trim()}T00:00:00`);
  if (Number.isNaN(date.getTime())) return String(value || "").trim();
  return new Intl.DateTimeFormat("pt-PT", {
    timeZone: DEFAULT_TZ,
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  }).format(date);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function itemCountsFilled(counts, itemTypes) {
  return itemTypes.every((item) => counts?.[item.id] != null);
}

function loadCounts(counts, itemId) {
  const value = counts?.[itemId];
  return value == null ? "" : Number(value || 0);
}

function formatLaundryItemsSummary(counts, itemTypes) {
  return itemTypes
    .map((item) => {
      const value = counts?.[item.id];
      const normalized = value == null || String(value).trim() === "" ? "" : String(Number(value || 0));
      return `${item.name}: ${normalized}`;
    })
    .join("\n");
}

function buildLaundryDifference(record, settings) {
  const itemTypes = Array.isArray(settings?.itemTypes) ? settings.itemTypes : [];
  let totalDiff = 0;
  const lines = [];
  itemTypes.forEach((item) => {
    const sent = Number(record?.sentItems?.[item.id] || 0);
    const received = Number(record?.receivedItems?.[item.id] || 0);
    const diff = received - sent;
    totalDiff += diff;
    if (diff !== 0) {
      lines.push(`${item.name}: ${sent} -> ${received} (${diff > 0 ? "+" : ""}${diff})`);
    }
  });
  lines.push(`Total counts difference: ${totalDiff > 0 ? "+" : ""}${totalDiff}`);
  return { totalDiff, lines };
}

function laundryRowColor(totalDiff) {
  if (totalDiff > 0) return "rgba(55, 140, 92, 0.15)";
  if (totalDiff < 0) return "rgba(177, 32, 48, 0.15)";
  return "rgba(67, 127, 211, 0.15)";
}

async function loadGeneralEmailConfig() {
  const rows = await restQuery("app_settings?select=payload&setting_key=eq.communications&limit=1", { method: "GET" });
  const payload = Array.isArray(rows) && rows[0]?.payload ? rows[0].payload : {};
  const generalEmailConfig = payload?.general?.emailConfig || payload?.general?.email_config || payload?.general?.bakeryEmailConfig || payload?.general?.bakery_email_config;
  return sanitizeBakerySettings({ emailConfig: generalEmailConfig || {} }).emailConfig;
}

async function loadLaundryPayloadRow() {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${encodeURIComponent(LAUNDRY_SETTING_KEY)}&limit=1`, {
    method: "GET",
  });
  const row = Array.isArray(rows) && rows[0] ? rows[0] : null;
  return {
    rowId: row?.id || "",
    payload: sanitizeLaundryPayload(row?.payload || { settings: DEFAULT_LAUNDRY_SETTINGS, records: [] }),
  };
}

function isMissingLaundryTableError(error) {
  const message = String(error?.message || "").toLowerCase();
  return message.includes("laundry_records") && (
    message.includes("could not find") ||
    message.includes("schema cache") ||
    message.includes("relation") ||
    message.includes("does not exist")
  );
}

async function loadLaundryRecords(settings) {
  try {
    const rows = await restQuery(
      "laundry_records?select=id,property,record_date,received_date,sent_items,received_items,received_weight_kg,notes,created_at,updated_at&order=record_date.desc,property.asc",
      { method: "GET" }
    );
    return (Array.isArray(rows) ? rows : []).map((row) => sanitizeLaundryRecord({
      id: row?.id,
      property: row?.property,
      date: row?.record_date,
      receivedDate: row?.received_date,
      sentItems: row?.sent_items,
      receivedItems: row?.received_items,
      receivedWeightKg: row?.received_weight_kg,
      notes: row?.notes,
      createdAt: row?.created_at,
      updatedAt: row?.updated_at,
    }, settings));
  } catch (error) {
    if (!isMissingLaundryTableError(error)) throw error;
    const legacy = await loadLaundryPayloadRow();
    return legacy.payload.records;
  }
}

async function loadLastSentSlot(settingKey = LAST_SENT_KEY) {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${encodeURIComponent(settingKey)}&limit=1`, { method: "GET" });
  if (!Array.isArray(rows) || !rows[0]) return { id: null, slotKey: "" };
  return {
    id: rows[0].id,
    slotKey: String(rows[0]?.payload?.slotKey || ""),
  };
}

async function saveLastSentSlot(id, slotKey, sentAt, providerMessageId, settingKey = LAST_SENT_KEY) {
  const payload = {
    slotKey: String(slotKey || ""),
    sentAt,
    providerMessageId: String(providerMessageId || ""),
  };
  if (id) {
    await restQuery(`app_settings?setting_key=eq.${encodeURIComponent(settingKey)}`, {
      method: "PATCH",
      body: { payload, updated_at: new Date().toISOString() },
    });
    return;
  }
  await restQuery("app_settings", {
    method: "POST",
    body: [{ setting_key: settingKey, payload }],
  });
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

async function sendConfiguredEmail(emailConfig, mail) {
  if (cleanText(emailConfig?.provider).toLowerCase() === "smtp") {
    return sendWithSmtp(emailConfig, mail);
  }
  return sendWithResend(mail, emailConfig);
}

function buildLaundrySupplierEmailContent(records, settings, sentDate) {
  const itemTypes = Array.isArray(settings?.itemTypes) ? settings.itemTypes : [];
  const byProperty = new Map(records.map((record) => [record.property, record]));
  const subject = `ROUPA ENVIADA DIA ${formatDatePt(sentDate)}`;
  const headerCells = itemTypes
    .map((item) => `<th style="border:1px solid #d8d0c7;padding:6px;text-align:center;">${escapeHtml(item.name)}</th>`)
    .join("");
  const bodyRows = PROPERTIES.map((property) => {
    const record = byProperty.get(property);
    const cells = itemTypes
      .map((item) => `<td style="border:1px solid #d8d0c7;padding:6px;text-align:center;">${escapeHtml(String(loadCounts(record?.sentItems, item.id)))}</td>`)
      .join("");
    return `<tr><td style="border:1px solid #d8d0c7;padding:6px;font-weight:600;">${escapeHtml(property)}</td>${cells}</tr>`;
  }).join("");
  const html = `<!doctype html><html><body style="font-family:Arial,sans-serif;color:#1f2937;">
    <p>Bom dia,</p>
    <p>Segue a roupa enviada dia ${escapeHtml(formatDatePt(sentDate))}</p>
    <table style="border-collapse:collapse;width:100%;max-width:760px;">
      <thead>
        <tr>
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:left;"></th>
          ${headerCells}
        </tr>
      </thead>
      <tbody>${bodyRows}</tbody>
    </table>
    <p>Cumprimentos,<br />
    Lisboa Central Hostel<br /><br />
    +351 309 881 038<br />
    +351 925 222 809<br />
    global@lisboacentralhostel.com<br /><br />
    Rua Rodrigues Sampaio 160, 1150-282 Lisboa</p>
  </body></html>`;
  const lines = [
    `ROUPA ENVIADA DIA ${formatDatePt(sentDate)}`,
    "",
    "Bom dia,",
    `Segue a roupa enviada dia ${formatDatePt(sentDate)}`,
    "",
    ["", ...itemTypes.map((item) => item.name)].join(" | "),
    ...PROPERTIES.map((property) => {
      const record = byProperty.get(property);
      return [property, ...itemTypes.map((item) => String(loadCounts(record?.sentItems, item.id)))].join(" | ");
    }),
    "",
    "Cumprimentos,",
    "Lisboa Central Hostel",
    "",
    "+351 309 881 038",
    "+351 925 222 809",
    "global@lisboacentralhostel.com",
    "",
    "Rua Rodrigues Sampaio 160, 1150-282 Lisboa",
  ];
  return { subject, html, text: lines.join("\n") };
}

function buildLaundryManagementEmailContent(records, settings, currentDate) {
  const subject = `LCH - Diferenças de Lavandaria ${currentDate} (últimos 3 dias)`;
  const rowsHtml = records.length
    ? records.map((record) => {
      const diff = buildLaundryDifference(record, settings);
      const rowColor = laundryRowColor(diff.totalDiff);
      const tdStyle = `border:1px solid #d8d0c7;padding:6px;vertical-align:top;background:${rowColor};background-color:${rowColor};`;
      return `<tr>
        <td style="${tdStyle}">${escapeHtml(record.date)}</td>
        <td style="${tdStyle}">${escapeHtml(record.property)}</td>
        <td style="${tdStyle}">${escapeHtml(formatLaundryItemsSummary(record.sentItems, settings.itemTypes)).replace(/\n/g, "<br />")}</td>
        <td style="${tdStyle}">${escapeHtml(record.receivedDate || "")}</td>
        <td style="${tdStyle}">${escapeHtml(formatLaundryItemsSummary(record.receivedItems, settings.itemTypes)).replace(/\n/g, "<br />")}</td>
        <td style="${tdStyle}">${escapeHtml(diff.lines.join("\n")).replace(/\n/g, "<br />")}</td>
        <td style="${tdStyle}">${escapeHtml(record.notes || "-")}</td>
      </tr>`;
    }).join("")
    : '<tr><td colspan="7" style="border:1px solid #d8d0c7;padding:6px;">No laundry records received in the last 3 days.</td></tr>';
  const html = `<!doctype html><html><body style="font-family:Arial,sans-serif;color:#1f2937;">
    <p>Bom dia,</p>
    <p>Segue o resumo das diferenças de lavandaria dos últimos 3 dias.</p>
    <table style="border-collapse:collapse;width:100%;max-width:980px;">
      <thead>
        <tr>
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:left;">Sent Date</th>
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:left;">Property</th>
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:left;">Sent</th>
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:left;">Received Date</th>
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:left;">Received</th>
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:left;">Difference</th>
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:left;">Notes</th>
        </tr>
      </thead>
      <tbody>${rowsHtml}</tbody>
    </table>
  </body></html>`;
  const textLines = [
    subject,
    "",
    "Bom dia,",
    "Segue o resumo das diferenças de lavandaria dos últimos 3 dias.",
    "",
    ...(
      records.length
        ? records.flatMap((record) => {
          const diff = buildLaundryDifference(record, settings);
          return [
            `${record.date} | ${record.property} | ${record.receivedDate || "-"}`,
            `Sent: ${formatLaundryItemsSummary(record.sentItems, settings.itemTypes).replace(/\n/g, " ; ")}`,
            `Received: ${formatLaundryItemsSummary(record.receivedItems, settings.itemTypes).replace(/\n/g, " ; ")}`,
            `Difference: ${diff.lines.join(" ; ")}`,
            `Notes: ${record.notes || "-"}`,
            "",
          ];
        })
        : ["No laundry records received in the last 3 days."]
    ),
  ];
  return { subject, html, text: textLines.join("\n") };
}

async function processSupplierEmail({ settings, records, now, force, testRecipient = "" }) {
  if (!settings.emailEnabled && !force) {
    return { ok: true, status: "skipped", reason: "disabled", mode: "supplier" };
  }
  const due = shouldSendNow(settings.emailTime, now, DEFAULT_TZ);
  if (!force && !due.due) {
    return { ok: true, status: "skipped", reason: "time_mismatch", mode: "supplier" };
  }
  const lastSent = await loadLastSentSlot(LAST_SENT_KEY);
  if (!force && lastSent.slotKey === due.slotKey) {
    return { ok: true, status: "skipped", reason: "already_sent_for_slot", mode: "supplier" };
  }
  const recipients = testRecipient ? [testRecipient] : (Array.isArray(settings.emailRecipients) ? settings.emailRecipients : []);
  if (!recipients.length) {
    return { ok: true, status: "skipped", reason: "no_recipients", mode: "supplier" };
  }
  const sentDate = lisbonTodayIso(now, DEFAULT_TZ);
  const todayRecords = records.filter((record) => record.date === sentDate);
  const missingProperties = PROPERTIES.filter((property) => !todayRecords.some((record) => record.property === property));
  if (missingProperties.length) {
    return { ok: true, status: "skipped", reason: "missing_records", missingProperties, sentDate, mode: "supplier" };
  }
  const incompleteProperties = todayRecords
    .filter((record) => !itemCountsFilled(record.sentItems, settings.itemTypes))
    .map((record) => record.property);
  if (incompleteProperties.length) {
    return { ok: true, status: "skipped", reason: "incomplete_counts", incompleteProperties, sentDate, mode: "supplier" };
  }
  const generalEmailConfig = await loadGeneralEmailConfig();
  const content = buildLaundrySupplierEmailContent(todayRecords, settings, sentDate);
  const sent = await sendConfiguredEmail(generalEmailConfig, {
    to: recipients,
    subject: content.subject,
    html: content.html,
    text: content.text,
  });
  if (!testRecipient) {
    await saveLastSentSlot(lastSent.id, due.slotKey || sentDate, now.toISOString(), sent?.id, LAST_SENT_KEY);
  }
  return {
    ok: true,
    status: "sent",
    mode: "supplier",
    sentDate,
    recipients,
    providerMessageId: sent?.id || "",
    count: todayRecords.length,
  };
}

async function processManagementEmail({ settings, records, now, force, testRecipient = "" }) {
  if (!settings.managementEmailEnabled && !force) {
    return { ok: true, status: "skipped", reason: "disabled", mode: "management" };
  }
  const due = shouldSendNow(settings.managementEmailTime, now, DEFAULT_TZ);
  if (!force && !due.due) {
    return { ok: true, status: "skipped", reason: "time_mismatch", mode: "management" };
  }
  const lastSent = await loadLastSentSlot(MANAGEMENT_LAST_SENT_KEY);
  if (!force && lastSent.slotKey === due.slotKey) {
    return { ok: true, status: "skipped", reason: "already_sent_for_slot", mode: "management" };
  }
  const recipients = testRecipient ? [testRecipient] : (Array.isArray(settings.managementEmailRecipients) ? settings.managementEmailRecipients : []);
  if (!recipients.length) {
    return { ok: true, status: "skipped", reason: "no_recipients", mode: "management" };
  }
  const currentDate = lisbonTodayIso(now, DEFAULT_TZ);
  const fromDate = shiftIsoDate(currentDate, -2);
  const windowRecords = records
    .filter((record) => {
      const receivedDate = cleanText(record.receivedDate || "");
      if (!receivedDate || receivedDate < fromDate || receivedDate > currentDate) return false;
      return itemCountsFilled(record.receivedItems, settings.itemTypes);
    })
    .sort((a, b) => {
      const receivedCompare = cleanText(b.receivedDate).localeCompare(cleanText(a.receivedDate));
      if (receivedCompare !== 0) return receivedCompare;
      const sentCompare = cleanText(b.date).localeCompare(cleanText(a.date));
      if (sentCompare !== 0) return sentCompare;
      return cleanText(a.property).localeCompare(cleanText(b.property));
    });
  const content = buildLaundryManagementEmailContent(windowRecords, settings, currentDate);
  const sent = await sendWithResend({
    to: recipients,
    subject: content.subject,
    html: content.html,
    text: content.text,
  });
  if (!testRecipient) {
    await saveLastSentSlot(lastSent.id, due.slotKey || currentDate, now.toISOString(), sent?.id, MANAGEMENT_LAST_SENT_KEY);
  }
  return {
    ok: true,
    status: "sent",
    mode: "management",
    currentDate,
    recipients,
    providerMessageId: sent?.id || "",
    count: windowRecords.length,
  };
}

module.exports = async function handler(req, res) {
  try {
    if (req.method !== "GET" && req.method !== "POST") {
      res.status(405).json({ error: "Method not allowed." });
      return;
    }

    const authHeader = String(req.headers.authorization || "");
    const bearerToken = authHeader.startsWith("Bearer ") ? authHeader.slice(7).trim() : "";
    const cronSecret = String(process.env.CRON_SECRET || "").trim();
    const userAgent = String(req.headers["user-agent"] || "").toLowerCase();
    const isCronRequest = !!req.headers["x-vercel-cron"] || (!!cronSecret && bearerToken === cronSecret) || userAgent.includes("vercel-cron");
    if (!isCronRequest) await requireFeature(req, "settings", "laundry");

    const force = String(req.query?.force || "") === "1";
    const body = req.method === "POST" ? await parseBody(req) : {};
    const testRecipient = cleanText(body?.testRecipient || "").toLowerCase();
    const requestedMode = cleanText(body?.mode || req.query?.mode).toLowerCase();
    const { payload } = await loadLaundryPayloadRow();
    const settings = sanitizeLaundrySettings(payload.settings);
    const now = new Date();
    const records = await loadLaundryRecords(settings);

    if (requestedMode === "management") {
      const result = await processManagementEmail({ settings, records, now, force, testRecipient });
      res.status(200).json(result);
      return;
    }

    if (requestedMode === "supplier" || requestedMode === "laundry") {
      const result = await processSupplierEmail({ settings, records, now, force, testRecipient });
      res.status(200).json(result);
      return;
    }

    if (isCronRequest) {
      const [supplier, management] = await Promise.all([
        processSupplierEmail({ settings, records, now, force: false }),
        processManagementEmail({ settings, records, now, force: false }),
      ]);
      res.status(200).json({ ok: true, supplier, management });
      return;
    }

    const result = await processSupplierEmail({ settings, records, now, force, testRecipient });
    res.status(200).json(result);
  } catch (error) {
    sendError(res, error);
  }
};
