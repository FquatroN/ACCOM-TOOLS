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

async function loadLastSentSlot() {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${LAST_SENT_KEY}&limit=1`, { method: "GET" });
  if (!Array.isArray(rows) || !rows[0]) return { id: null, slotKey: "" };
  return {
    id: rows[0].id,
    slotKey: String(rows[0]?.payload?.slotKey || ""),
  };
}

async function saveLastSentSlot(id, slotKey, sentAt, providerMessageId) {
  const payload = {
    slotKey: String(slotKey || ""),
    sentAt,
    providerMessageId: String(providerMessageId || ""),
  };
  if (id) {
    await restQuery(`app_settings?setting_key=eq.${LAST_SENT_KEY}`, {
      method: "PATCH",
      body: { payload, updated_at: new Date().toISOString() },
    });
    return;
  }
  await restQuery("app_settings", {
    method: "POST",
    body: [{ setting_key: LAST_SENT_KEY, payload }],
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

function buildLaundryEmailContent(records, settings, sentDate) {
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
    const { payload } = await loadLaundryPayloadRow();
    const settings = sanitizeLaundrySettings(payload.settings);
    if (!settings.emailEnabled && !force) {
      res.status(200).json({ ok: true, status: "skipped", reason: "disabled" });
      return;
    }

    const now = new Date();
    const due = shouldSendNow(settings.emailTime, now, DEFAULT_TZ);
    if (!force && !due.due) {
      res.status(200).json({ ok: true, status: "skipped", reason: "time_mismatch" });
      return;
    }

    const lastSent = await loadLastSentSlot();
    if (!force && lastSent.slotKey === due.slotKey) {
      res.status(200).json({ ok: true, status: "skipped", reason: "already_sent_for_slot" });
      return;
    }

    const recipients = testRecipient ? [testRecipient] : (Array.isArray(settings.emailRecipients) ? settings.emailRecipients : []);
    if (!recipients.length) {
      res.status(200).json({ ok: true, status: "skipped", reason: "no_recipients" });
      return;
    }

    const sentDate = lisbonTodayIso(now, DEFAULT_TZ);
    const allRecords = await loadLaundryRecords(settings);
    const records = allRecords.filter((record) => record.date === sentDate);
    const missingProperties = PROPERTIES.filter((property) => !records.some((record) => record.property === property));
    if (missingProperties.length) {
      res.status(200).json({ ok: true, status: "skipped", reason: "missing_records", missingProperties, sentDate });
      return;
    }

    const incompleteProperties = records
      .filter((record) => !itemCountsFilled(record.sentItems, settings.itemTypes))
      .map((record) => record.property);
    if (incompleteProperties.length) {
      res.status(200).json({ ok: true, status: "skipped", reason: "incomplete_counts", incompleteProperties, sentDate });
      return;
    }

    const generalEmailConfig = await loadGeneralEmailConfig();
    const content = buildLaundryEmailContent(records, settings, sentDate);
    const sent = await sendConfiguredEmail(generalEmailConfig, {
      to: recipients,
      subject: content.subject,
      html: content.html,
      text: content.text,
    });

    if (!testRecipient) {
      await saveLastSentSlot(lastSent.id, due.slotKey || sentDate, now.toISOString(), sent?.id);
    }
    res.status(200).json({
      ok: true,
      status: "sent",
      sentDate,
      recipients,
      providerMessageId: sent?.id || "",
      count: records.length,
    });
  } catch (error) {
    sendError(res, error);
  }
};
