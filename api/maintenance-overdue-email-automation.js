const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  DEFAULT_MAINTENANCE_SETTINGS,
  MAINTENANCE_OVERDUE_EMAIL_LAST_SENT_KEY,
  MAINTENANCE_SETTING_KEY,
  buildMaintenanceOverdueRows,
  isoTodayInTimeZone,
  sanitizeMaintenanceLog,
  sanitizeMaintenanceSettings,
} = require("./_maintenance");

const DEFAULT_TZ = process.env.AUTOMATION_TIMEZONE || "Europe/Lisbon";
const WEEKLY_SEND_DAY = 1; // Monday
const WEEKLY_SEND_HOUR = 9;

function clean(value) {
  return String(value || "").trim();
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

function isoWeekSlotKey(dateIso) {
  const date = new Date(`${dateIso}T00:00:00Z`);
  const day = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  return `${date.getUTCFullYear()}-W${String(weekNo).padStart(2, "0")}`;
}

function isDueThisWeek(now = new Date(), timeZone = DEFAULT_TZ) {
  const todayIso = isoTodayInTimeZone(now, timeZone);
  const day = new Date(`${todayIso}T00:00:00Z`).getUTCDay() || 7;
  const clock = getClockParts(now, timeZone);
  if (day !== WEEKLY_SEND_DAY || clock.hour < WEEKLY_SEND_HOUR) {
    return { due: false, reason: "time_mismatch", slotKey: isoWeekSlotKey(todayIso), todayIso };
  }
  return { due: true, slotKey: isoWeekSlotKey(todayIso), todayIso };
}

function formatDatePt(value, timeZone = DEFAULT_TZ) {
  const date = new Date(`${clean(value)}T00:00:00`);
  if (Number.isNaN(date.getTime())) return clean(value);
  return new Intl.DateTimeFormat("pt-PT", {
    timeZone,
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

async function loadMaintenanceSettings() {
  const rows = await restQuery(`app_settings?select=payload&setting_key=eq.${encodeURIComponent(MAINTENANCE_SETTING_KEY)}&limit=1`, {
    method: "GET",
  });
  const payload = Array.isArray(rows) && rows[0]?.payload ? rows[0].payload : DEFAULT_MAINTENANCE_SETTINGS;
  return sanitizeMaintenanceSettings(payload);
}

function mapMaintenanceLogRow(row) {
  return sanitizeMaintenanceLog({
    id: row?.id,
    taskId: row?.task_id,
    taskName: row?.task_name,
    whereValue: row?.where_value,
    doneDate: row?.done_date,
    type: row?.type,
    who: row?.who,
    note: row?.note,
    createdAt: row?.created_at,
    updatedAt: row?.updated_at,
  });
}

async function loadMaintenanceLogs() {
  const rows = await restQuery("maintenance_logs?select=*&order=done_date.desc,created_at.desc", {
    method: "GET",
  });
  return (Array.isArray(rows) ? rows : []).map(mapMaintenanceLogRow);
}

async function loadLastSentSlot() {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${encodeURIComponent(MAINTENANCE_OVERDUE_EMAIL_LAST_SENT_KEY)}&limit=1`, {
    method: "GET",
  });
  if (!Array.isArray(rows) || !rows[0]) return { id: null, slotKey: "" };
  return {
    id: rows[0].id,
    slotKey: clean(rows[0]?.payload?.slotKey),
  };
}

async function saveLastSentSlot(id, slotKey, sentAt, providerMessageId) {
  const payload = { slotKey: clean(slotKey), sentAt, providerMessageId: clean(providerMessageId) };
  if (id) {
    await restQuery(`app_settings?setting_key=eq.${encodeURIComponent(MAINTENANCE_OVERDUE_EMAIL_LAST_SENT_KEY)}`, {
      method: "PATCH",
      body: { payload, updated_at: new Date().toISOString() },
    });
    return;
  }
  await restQuery("app_settings", {
    method: "POST",
    body: [{ setting_key: MAINTENANCE_OVERDUE_EMAIL_LAST_SENT_KEY, payload }],
  });
}

async function sendWithResend({ to, subject, html, text }) {
  const apiKey = process.env.RESEND_API_KEY;
  if (!apiKey) {
    const error = new Error("Missing server environment variable: RESEND_API_KEY");
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
      from: "ACCOM Tools - LCH <info@accomtools.com>",
      reply_to: "info@accomtools.com",
      to,
      subject,
      html,
      text,
    }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload?.message || payload?.error || `Email provider failed (${response.status})`);
    error.statusCode = 502;
    throw error;
  }
  return payload;
}

function buildEmailContent(rows, todayIso) {
  const subject = `Tarefas em atraso a ${formatDatePt(todayIso)}`;
  if (!rows.length) {
    return {
      subject,
      html: `<p>As seguintes tarefas estão em atraso:</p><p>Não existem tarefas em atraso.</p>`,
      text: `As seguintes tarefas estão em atraso:\n\nNão existem tarefas em atraso.`,
    };
  }
  const htmlRows = rows.map((row) => `<tr>
      <td style="border:1px solid #d8dee4;padding:8px;text-align:left;">${escapeHtml(row.task || "-")}</td>
      <td style="border:1px solid #d8dee4;padding:8px;text-align:left;">${escapeHtml(row.whereValue || "-")}</td>
      <td style="border:1px solid #d8dee4;padding:8px;text-align:left;">${escapeHtml(row.lastTask ? formatDatePt(row.lastTask) : "Never")}</td>
      <td style="border:1px solid #d8dee4;padding:8px;text-align:left;">${escapeHtml(row.overdueLabel || "-")}</td>
    </tr>`).join("");
  const html = `<p>As seguintes tarefas estão em atraso:</p>
    <table style="border-collapse:collapse;width:100%;font-family:Arial,sans-serif;font-size:13px;">
      <thead>
        <tr>
          <th style="border:1px solid #d8dee4;padding:8px;background:#f3f4f6;text-align:left;">Task</th>
          <th style="border:1px solid #d8dee4;padding:8px;background:#f3f4f6;text-align:left;">Where</th>
          <th style="border:1px solid #d8dee4;padding:8px;background:#f3f4f6;text-align:left;">Last Task</th>
          <th style="border:1px solid #d8dee4;padding:8px;background:#f3f4f6;text-align:left;">Overdue By</th>
        </tr>
      </thead>
      <tbody>${htmlRows}</tbody>
    </table>`;
  const text = [
    "As seguintes tarefas estão em atraso:",
    "",
    "Task | Where | Last Task | Overdue By",
    ...rows.map((row) => `${row.task || "-"} | ${row.whereValue || "-"} | ${row.lastTask ? formatDatePt(row.lastTask) : "Never"} | ${row.overdueLabel || "-"}`),
  ].join("\n");
  return { subject, html, text };
}

module.exports = async function handler(req, res) {
  try {
    const authHeader = String(req.headers.authorization || "");
    const bearerToken = authHeader.startsWith("Bearer ") ? authHeader.slice(7).trim() : "";
    const cronSecret = String(process.env.CRON_SECRET || "").trim();
    const userAgent = String(req.headers["user-agent"] || "").toLowerCase();
    const isCronRequest = !!req.headers["x-vercel-cron"] || (!!cronSecret && bearerToken === cronSecret) || userAgent.includes("vercel-cron");
    if (!isCronRequest) await requireFeature(req, "settings", "maintenance");

    const force = String(req.query?.force || "") === "1";
    if (req.method !== "POST") {
      res.status(405).json({ error: "Method not allowed." });
      return;
    }
    await parseBody(req);
    const due = isDueThisWeek(new Date(), DEFAULT_TZ);
    if (!force && !due.due) {
      res.status(200).json({ ok: true, status: "skipped", reason: due.reason });
      return;
    }
    const [settings, logs] = await Promise.all([loadMaintenanceSettings(), loadMaintenanceLogs()]);
    const recipients = Array.isArray(settings.overdueEmailRecipients) ? settings.overdueEmailRecipients : [];
    if (!recipients.length) {
      res.status(200).json({ ok: true, status: "skipped", reason: "no_recipients" });
      return;
    }
    const overdueRows = buildMaintenanceOverdueRows(settings, logs, due.todayIso);
    if (!force && !overdueRows.length) {
      res.status(200).json({ ok: true, status: "skipped", reason: "no_overdue_tasks" });
      return;
    }
    const lastSent = await loadLastSentSlot();
    if (!force && lastSent.slotKey === due.slotKey) {
      res.status(200).json({ ok: true, status: "skipped", reason: "already_sent_for_slot" });
      return;
    }
    const content = buildEmailContent(overdueRows, due.todayIso);
    const provider = await sendWithResend({
      to: recipients,
      subject: content.subject,
      html: content.html,
      text: content.text,
    });
    if (!force) {
      await saveLastSentSlot(lastSent.id, due.slotKey, new Date().toISOString(), clean(provider?.id));
    }
    res.status(200).json({
      ok: true,
      status: "sent",
      recipients,
      count: overdueRows.length,
      providerMessageId: clean(provider?.id),
    });
  } catch (error) {
    sendError(res, error);
  }
};
