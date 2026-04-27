const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");

const SETTINGS_KEY = "communications";
const LAST_SENT_KEY = "communications_email_automation_last_sent";
const DEFAULT_TZ = process.env.AUTOMATION_TIMEZONE || "Europe/Lisbon";
const DEFAULT_CATEGORIES = [
  { name: "Warning", color: "#ffd89b" },
  { name: "Maintenance", color: "#a9f0df" },
  { name: "Information", color: "#add4ff" },
  { name: "very important", color: "#ffb3c2" },
];

function normalizeFrequency(value) {
  return value === "every1h" || value === "every4h" || value === "every8h" ? value : "everyday";
}

function emailFrequencyStep(value) {
  if (value === "every1h") return 1;
  if (value === "every4h") return 4;
  return 8;
}

function normalizeTime(value) {
  const raw = String(value || "").trim();
  return /^\d{2}:\d{2}$/.test(raw) ? raw : "00:00";
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || "").trim());
}

function normalizeRecipients(value) {
  const source = Array.isArray(value) ? value.join(",") : String(value || "");
  const seen = new Set();
  return source
    .split(/[\n,;]/)
    .map((item) => item.trim().toLowerCase())
    .filter((item) => isValidEmail(item))
    .filter((item) => {
      if (seen.has(item)) return false;
      seen.add(item);
      return true;
    });
}

function normalizeSchedule(raw, suffix = "") {
  return {
    key: suffix ? `schedule${suffix}` : "schedule1",
    frequency: normalizeFrequency(raw?.[`frequency${suffix}`] ?? raw?.frequency),
    timeOfDay: normalizeTime(raw?.[`timeOfDay${suffix}`] ?? raw?.timeOfDay),
    recipients: normalizeRecipients(raw?.[`recipients${suffix}`] ?? raw?.recipients),
  };
}

function normalizeHex(value) {
  const raw = String(value || "").trim();
  if (/^#[0-9a-fA-F]{6}$/.test(raw)) return raw.toLowerCase();
  if (/^#[0-9a-fA-F]{3}$/.test(raw)) return `#${raw[1]}${raw[1]}${raw[2]}${raw[2]}${raw[3]}${raw[3]}`.toLowerCase();
  return "#d8d8d8";
}

function normalizeAutoCloseDays(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  const num = Number(raw);
  if (!Number.isFinite(num) || num <= 0) return null;
  return Math.floor(num);
}

function normalizeCategories(value) {
  const list = Array.isArray(value) ? value : [];
  const seen = new Set();
  const cleaned = list
    .map((item) => ({
      name: String(item?.name || "").trim(),
      color: normalizeHex(item?.color),
      autoCloseDays: normalizeAutoCloseDays(item?.autoCloseDays),
    }))
    .filter((item) => item.name)
    .filter((item) => {
      const key = item.name.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  return cleaned.length > 0 ? cleaned : DEFAULT_CATEGORIES;
}

function isClosedStatus(value) {
  const raw = String(value || "").trim().toLowerCase();
  return raw === "closed" || raw === "resolved" || raw === "archived";
}

function entryCreatedTime(entry) {
  const fromCreatedAt = new Date(String(entry?.created_at || ""));
  if (!Number.isNaN(fromCreatedAt.getTime())) return fromCreatedAt;
  const date = String(entry?.date || "").trim();
  const time = String(entry?.time || "").trim().slice(0, 5) || "00:00";
  if (!date) return null;
  const fallback = new Date(`${date}T${time}:00`);
  return Number.isNaN(fallback.getTime()) ? null : fallback;
}

function entryUpdatedTime(entry) {
  const fromUpdatedAt = new Date(String(entry?.updated_at || ""));
  if (!Number.isNaN(fromUpdatedAt.getTime())) return fromUpdatedAt;
  return entryCreatedTime(entry);
}

function wasClosedInLast24h(entry) {
  if (!isClosedStatus(entry?.status)) return false;
  const updated = entryUpdatedTime(entry);
  if (!updated) return false;
  return Date.now() - updated.getTime() <= 24 * 60 * 60 * 1000;
}

function splitActiveRows(rows) {
  const openRows = [];
  const recentlyClosedRows = [];
  rows.forEach((row) => {
    if (!isClosedStatus(row?.status)) {
      openRows.push(row);
      return;
    }
    if (wasClosedInLast24h(row)) recentlyClosedRows.push(row);
  });
  return { openRows, recentlyClosedRows };
}

function autoCloseDaysForCategory(categories, categoryName) {
  const key = String(categoryName || "").trim().toLowerCase();
  const match = categories.find((category) => category.name.toLowerCase() === key);
  return normalizeAutoCloseDays(match?.autoCloseDays);
}

function shouldAutoClose(row, categories, now) {
  if (isClosedStatus(row?.status)) return false;
  const days = autoCloseDaysForCategory(categories, row?.category);
  if (!days) return false;
  const created = entryCreatedTime(row);
  if (!created) return false;
  return now.getTime() - created.getTime() > days * 24 * 60 * 60 * 1000;
}

async function autoCloseExpiredRows(rows, categories, now) {
  const candidates = rows.filter((row) => shouldAutoClose(row, categories, now));
  for (const row of candidates) {
    await restQuery(`communications?id=eq.${encodeURIComponent(row.id)}`, {
      method: "PATCH",
      body: { status: "Closed", updated_at: now.toISOString() },
    });
    row.status = "Closed";
    row.updated_at = now.toISOString();
  }
  return candidates.length;
}

function hexToRgba(hex, alpha) {
  const raw = normalizeHex(hex).slice(1);
  const r = parseInt(raw.slice(0, 2), 16);
  const g = parseInt(raw.slice(2, 4), 16);
  const b = parseInt(raw.slice(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function tintHexOverWhite(hex, alpha) {
  const raw = normalizeHex(hex).slice(1);
  const r = parseInt(raw.slice(0, 2), 16);
  const g = parseInt(raw.slice(2, 4), 16);
  const b = parseInt(raw.slice(4, 6), 16);
  const a = Math.max(0, Math.min(1, Number(alpha)));
  const mix = (channel) => Math.round((1 - a) * 255 + a * channel);
  const rr = mix(r).toString(16).padStart(2, "0");
  const gg = mix(g).toString(16).padStart(2, "0");
  const bb = mix(b).toString(16).padStart(2, "0");
  return `#${rr}${gg}${bb}`;
}

function contrastText(hex) {
  const x = normalizeHex(hex).slice(1);
  const r = parseInt(x.slice(0, 2), 16);
  const g = parseInt(x.slice(2, 4), 16);
  const b = parseInt(x.slice(4, 6), 16);
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255 > 0.6 ? "#1d1714" : "#ffffff";
}

function categoryColor(categories, categoryName) {
  const key = String(categoryName || "").trim().toLowerCase();
  return (categories.find((c) => c.name.toLowerCase() === key) || categories[0] || DEFAULT_CATEGORIES[0]).color;
}

function rowBackground(status, categoryHex) {
  if (isClosedStatus(status)) return tintHexOverWhite("#2e9f42", 0.25);
  return tintHexOverWhite(categoryHex, 0.25);
}

function formatDateTimeShort(value, timeZone = DEFAULT_TZ) {
  const date = value instanceof Date ? value : new Date(String(value || ""));
  if (Number.isNaN(date.getTime())) return "";
  return new Intl.DateTimeFormat("en-GB", {
    timeZone,
    dateStyle: "short",
    timeStyle: "short",
  }).format(date);
}

function closedStatusStamp(row, timeZone) {
  if (!isClosedStatus(row?.status)) return "";
  return formatDateTimeShort(entryUpdatedTime(row), timeZone);
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
  parts.forEach((p) => {
    if (p.type !== "literal") map[p.type] = p.value;
  });
  return {
    year: Number(map.year),
    month: Number(map.month),
    day: Number(map.day),
    hour: Number(map.hour),
    minute: Number(map.minute),
  };
}

function shouldScheduleSendNow(schedule, now, timeZone) {
  if (!Array.isArray(schedule.recipients) || schedule.recipients.length === 0) {
    return { due: false, reason: "no_recipients" };
  }

  const frequency = normalizeFrequency(schedule.frequency);
  const timeOfDay = normalizeTime(schedule.timeOfDay);
  const [startHourRaw, startMinuteRaw] = timeOfDay.split(":");
  const startHour = Number(startHourRaw);
  const startMinute = Number(startMinuteRaw);
  const clock = getClockParts(now, timeZone);

  if (clock.minute !== startMinute) return { due: false, reason: "minute_mismatch" };

  if (frequency === "everyday") {
    if (clock.hour !== startHour) return { due: false, reason: "hour_mismatch" };
    return { due: true, slotKey: `${clock.year}-${clock.month}-${clock.day}:${clock.hour}:${clock.minute}` };
  }

  const step = emailFrequencyStep(frequency);
  const hourDelta = (clock.hour - startHour + 24) % 24;
  if (hourDelta % step !== 0) return { due: false, reason: "interval_mismatch" };
  return { due: true, slotKey: `${clock.year}-${clock.month}-${clock.day}:${clock.hour}:${clock.minute}` };
}

function dueSchedules(emailAutomation, now, timeZone) {
  if (!emailAutomation?.enabled) return { schedules: [], reason: "disabled" };
  const schedules = [emailAutomation.schedule1, emailAutomation.schedule2].filter(Boolean);
  const due = [];
  let lastReason = "no_recipients";
  for (const schedule of schedules) {
    const result = shouldScheduleSendNow(schedule, now, timeZone);
    if (result.due) due.push({ ...schedule, slotKey: result.slotKey });
    else lastReason = result.reason;
  }
  return { schedules: due, reason: due.length ? "" : lastReason };
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

async function loadSettings() {
  const rows = await restQuery(`app_settings?select=payload&setting_key=eq.${SETTINGS_KEY}&limit=1`, {
    method: "GET",
  });
  const payload = Array.isArray(rows) && rows[0] ? rows[0].payload : {};
  const comm = payload?.communications || {};
  const raw = payload?.communications?.emailAutomation || {};
  return {
    categories: normalizeCategories(comm.categories),
    enabled: !!raw.enabled,
    schedule1: normalizeSchedule(raw, ""),
    schedule2: normalizeSchedule(raw, "2"),
  };
}

async function loadLastSentSlot() {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${LAST_SENT_KEY}&limit=1`, {
    method: "GET",
  });
  if (!Array.isArray(rows) || !rows[0]) return { id: null, schedule1: "", schedule2: "" };
  const payload = rows[0].payload || {};
  return {
    id: rows[0].id,
    schedule1: String(payload?.schedule1 || payload?.slotKey || ""),
    schedule2: String(payload?.schedule2 || ""),
  };
}

async function saveLastSentSlot(id, slotKeys, sentAt, providerMessageIds) {
  const ids = Array.isArray(providerMessageIds) ? providerMessageIds.filter(Boolean) : [providerMessageIds].filter(Boolean);
  const payload = {
    slotKey: String(slotKeys?.schedule1 || ""),
    schedule1: String(slotKeys?.schedule1 || ""),
    schedule2: String(slotKeys?.schedule2 || ""),
    sentAt,
    providerMessageId: ids[0] || "",
    providerMessageIds: ids,
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

async function fetchRecentCommunications() {
  const rows = await restQuery(
    "communications?select=id,date,time,person,status,category,message,created_at,updated_at&order=date.desc,time.desc&limit=1000",
    { method: "GET" }
  );
  return Array.isArray(rows) ? rows : [];
}

function tableRowsHtml(rows, categories, timeZone) {
  const itemsHtml = rows
    .map(
      (row) => {
        const catColor = categoryColor(categories, row.category);
        const rowBg = rowBackground(row.status, catColor);
        const chipText = contrastText(catColor);
        const tdStyle = `border:1px solid #d1d5db;padding:6px;background:${rowBg};background-color:${rowBg};`;
        const closedStamp = closedStatusStamp(row, timeZone);
        const statusHtml = `${escapeHtml(row.status)}${closedStamp ? `<div style="margin-top:3px;color:#4b5563;font-size:11px;line-height:1.25;">${escapeHtml(closedStamp)}</div>` : ""}`;
        return `<tr>
          <td bgcolor="${rowBg}" style="${tdStyle}">${escapeHtml(row.date)}</td>
          <td bgcolor="${rowBg}" style="${tdStyle}">${escapeHtml(String(row.time || "").slice(0, 5))}</td>
          <td bgcolor="${rowBg}" style="${tdStyle}">${escapeHtml(row.person)}</td>
          <td bgcolor="${rowBg}" style="${tdStyle}"><span style="display:inline-block;padding:3px 10px;border-radius:999px;border:1px solid ${catColor};background:${catColor};color:${chipText};font-weight:600;">${escapeHtml(row.category)}</span></td>
          <td bgcolor="${rowBg}" style="${tdStyle}">${escapeHtml(row.message)}</td>
          <td bgcolor="${rowBg}" style="${tdStyle}">${statusHtml}</td>
        </tr>`;
      }
    )
    .join("");
  return itemsHtml || '<tr><td colspan="6" style="border:1px solid #d1d5db;padding:6px;">No records found.</td></tr>';
}

function buildEmailContent(openRows, recentlyClosedRows, timeZone, categories) {
  const stamp = new Intl.DateTimeFormat("en-GB", {
    timeZone,
    dateStyle: "medium",
    timeStyle: "short",
  }).format(new Date());

  const html = `<!doctype html>
<html>
  <body style="font-family:Arial,sans-serif;color:#1f2937;">
    <h2 style="margin-bottom:8px;">Comunicações ${escapeHtml(stamp)}</h2>
    <h3 style="margin:12px 0 6px;">Open records</h3>
    <table style="border-collapse:collapse;width:100%;">
      <thead>
        <tr>
          <th style="border:1px solid #d1d5db;padding:6px;text-align:left;">Date</th>
          <th style="border:1px solid #d1d5db;padding:6px;text-align:left;">Time</th>
          <th style="border:1px solid #d1d5db;padding:6px;text-align:left;">Person</th>
          <th style="border:1px solid #d1d5db;padding:6px;text-align:left;">Category</th>
          <th style="border:1px solid #d1d5db;padding:6px;text-align:left;">What happened?</th>
          <th style="border:1px solid #d1d5db;padding:6px;text-align:left;">Status</th>
        </tr>
      </thead>
      <tbody>${tableRowsHtml(openRows, categories, timeZone)}</tbody>
    </table>
    <h3 style="margin:16px 0 6px;">Closed in last 24h</h3>
    <table style="border-collapse:collapse;width:100%;">
      <thead>
        <tr>
          <th style="border:1px solid #d1d5db;padding:6px;text-align:left;">Date</th>
          <th style="border:1px solid #d1d5db;padding:6px;text-align:left;">Time</th>
          <th style="border:1px solid #d1d5db;padding:6px;text-align:left;">Person</th>
          <th style="border:1px solid #d1d5db;padding:6px;text-align:left;">Category</th>
          <th style="border:1px solid #d1d5db;padding:6px;text-align:left;">What happened?</th>
          <th style="border:1px solid #d1d5db;padding:6px;text-align:left;">Status</th>
        </tr>
      </thead>
      <tbody>${tableRowsHtml(recentlyClosedRows, categories, timeZone)}</tbody>
    </table>
    <p style="margin-top:14px;color:#6b7280;">Open: ${openRows.length} | Closed in last 24h: ${recentlyClosedRows.length}.</p>
  </body>
</html>`;

  const textLines = [
    `Comunicações ${stamp} (${timeZone})`,
    "",
    "Open records:",
    ...openRows.map(
      (row) =>
        `${row.date} ${String(row.time || "").slice(0, 5)} | ${row.person} | ${row.status} | ${row.category} | ${row.message}`
    ),
    "",
    "Closed in last 24h:",
    ...recentlyClosedRows.map(
      (row) => {
        const closedStamp = closedStatusStamp(row, timeZone);
        const status = closedStamp ? `${row.status} (${closedStamp})` : row.status;
        return `${row.date} ${String(row.time || "").slice(0, 5)} | ${row.person} | ${status} | ${row.category} | ${row.message}`;
      }
    ),
  ];

  return { html, text: textLines.join("\n") };
}

async function sendWithResend({ to, subject, html, text }) {
  const apiKey = process.env.RESEND_API_KEY;
  const from = process.env.EMAIL_FROM;
  if (!apiKey) {
    const error = new Error("Missing server environment variable: RESEND_API_KEY");
    error.statusCode = 500;
    throw error;
  }
  if (!from) {
    const error = new Error("Missing server environment variable: EMAIL_FROM");
    error.statusCode = 500;
    throw error;
  }

  const response = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      from,
      to,
      subject,
      html,
      text,
    }),
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload?.message || payload?.error || `Email provider failed (${response.status})`);
    error.statusCode = response.status;
    throw error;
  }
  return payload;
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
    if (!isCronRequest) await requireFeature(req, "settings", "communications");

    const force = String(req.query?.force || "") === "1";
    const body = req.method === "POST" ? await parseBody(req) : {};
    const testRecipient = String(body?.testRecipient || "").trim().toLowerCase();
    if (testRecipient && !isValidEmail(testRecipient)) {
      const err = new Error("Invalid test recipient email.");
      err.statusCode = 400;
      throw err;
    }
    const emailAutomation = await loadSettings();
    const now = new Date();
    const allRows = await fetchRecentCommunications();
    const autoClosedCount = await autoCloseExpiredRows(allRows, emailAutomation.categories, now);
    const dueCheck = dueSchedules(emailAutomation, now, DEFAULT_TZ);

    if (!force && !dueCheck.schedules.length) {
      res.status(200).json({ ok: true, status: "skipped", reason: dueCheck.reason, autoClosedCount });
      return;
    }

    const slotKeys = Object.fromEntries(dueCheck.schedules.map((schedule) => [schedule.key, schedule.slotKey]));
    let lastSent = { id: null, schedule1: "", schedule2: "" };
    if (!testRecipient) {
      lastSent = await loadLastSentSlot();
      const unsentSchedules = dueCheck.schedules.filter((schedule) => lastSent[schedule.key] !== schedule.slotKey);
      if (!force && !unsentSchedules.length) {
        res.status(200).json({ ok: true, status: "skipped", reason: "already_sent_for_slot", autoClosedCount });
        return;
      }
      dueCheck.schedules = unsentSchedules;
    }

    const recipients = testRecipient
      ? [testRecipient]
      : Array.from(new Set(dueCheck.schedules.flatMap((schedule) => schedule.recipients)));
    if (!Array.isArray(recipients) || recipients.length === 0) {
      res.status(200).json({ ok: true, status: "skipped", reason: "no_recipients", autoClosedCount });
      return;
    }

    const { openRows, recentlyClosedRows } = splitActiveRows(allRows);
    const content = buildEmailContent(openRows, recentlyClosedRows, DEFAULT_TZ, emailAutomation.categories);
    const day = new Intl.DateTimeFormat("pt-PT", {
      timeZone: DEFAULT_TZ,
      dateStyle: "short",
    }).format(now);
    const hour = new Intl.DateTimeFormat("pt-PT", {
      timeZone: DEFAULT_TZ,
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
    }).format(now);
    const subject = `Comunicações ${day} - ${hour}`;
    const sent = await sendWithResend({
      to: recipients,
      subject,
      html: content.html,
      text: content.text,
    });

    if (!testRecipient) {
      await saveLastSentSlot(lastSent.id, {
        schedule1: dueCheck.schedules.some((schedule) => schedule.key === "schedule1") ? slotKeys.schedule1 : lastSent.schedule1,
        schedule2: dueCheck.schedules.some((schedule) => schedule.key === "schedule2") ? slotKeys.schedule2 : lastSent.schedule2,
      }, now.toISOString(), sent?.id);
    }
    res.status(200).json({
      ok: true,
      status: "sent",
      recipients,
      providerMessageId: sent?.id || "",
      count: openRows.length + recentlyClosedRows.length,
      openCount: openRows.length,
      closedRecentCount: recentlyClosedRows.length,
      autoClosedCount,
    });
  } catch (error) {
    sendError(res, error);
  }
};
