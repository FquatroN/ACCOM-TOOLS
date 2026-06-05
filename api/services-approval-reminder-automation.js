const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");

const SETTINGS_KEY = "services";
const LAST_SENT_KEY = "services_approval_reminder_last_sent";
const DEFAULT_TZ = process.env.AUTOMATION_TIMEZONE || "Europe/Lisbon";

function clean(value) {
  return String(value || "").trim();
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean(value).toLowerCase());
}

function normalizeBool(value, fallback = false) {
  if (typeof value === "boolean") return value;
  const raw = clean(value).toLowerCase();
  if (!raw) return fallback;
  if (["true", "1", "yes", "on"].includes(raw)) return true;
  if (["false", "0", "no", "off"].includes(raw)) return false;
  return fallback;
}

function normalizeTime(value, fallback = "09:00") {
  const raw = clean(value);
  return /^\d{2}:\d{2}$/.test(raw) ? raw : fallback;
}

function normalizeServiceSettings(payload) {
  const source = payload && typeof payload === "object" ? payload : {};
  const rawConfigs = Array.isArray(source.serviceConfigs || source.service_configs)
    ? source.serviceConfigs || source.service_configs
    : [];
  return {
    approvalReminderEnabled: normalizeBool(source.approvalReminderEnabled ?? source.approval_reminder_enabled, false),
    approvalReminderTime: normalizeTime(source.approvalReminderTime ?? source.approval_reminder_time, "09:00"),
    approvalReminderTestEmail: clean(source.approvalReminderTestEmail ?? source.approval_reminder_test_email).toLowerCase(),
    serviceConfigs: rawConfigs.map((item) => ({
      serviceType: clean(item?.serviceType || item?.service_type),
      providerEmail: clean(item?.providerEmail || item?.provider_email).toLowerCase(),
    })).filter((item) => item.serviceType),
  };
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
  const currentMinutes = (clock.hour * 60) + clock.minute;
  const targetMinutes = (hour * 60) + minute;
  if (currentMinutes < targetMinutes) return { due: false, reason: "before_time" };
  return {
    due: true,
    slotKey: `${clock.year}-${String(clock.month).padStart(2, "0")}-${String(clock.day).padStart(2, "0")}:${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`,
  };
}

function formatDateDisplay(value) {
  const raw = clean(value);
  if (!raw) return "-";
  const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return raw;
  return `${match[3]}/${match[2]}/${match[1]}`;
}

function formatTimeDisplay(value) {
  const raw = clean(value);
  return raw ? raw.slice(0, 5) : "-";
}

function formatMoney(value) {
  const amount = Number(value || 0);
  if (!Number.isFinite(amount)) return "0,00 €";
  return `${new Intl.NumberFormat("pt-PT", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(amount)} €`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function appBaseUrl() {
  return (
    clean(process.env.APP_BASE_URL) ||
    clean(process.env.PUBLIC_APP_URL) ||
    clean(process.env.NEXT_PUBLIC_APP_URL) ||
    "https://accomtools.com"
  ).replace(/\/+$/, "");
}

function serviceDeepLink(row) {
  const serviceKey = clean(row?.request_number || row?.requestNumber || row?.id);
  if (!serviceKey) return `${appBaseUrl()}/index.html?view=services`;
  return `${appBaseUrl()}/index.html?view=services&service=${encodeURIComponent(serviceKey)}`;
}

async function loadServiceSettings() {
  const rows = await restQuery(`app_settings?select=payload&setting_key=eq.${encodeURIComponent(SETTINGS_KEY)}&limit=1`, {
    method: "GET",
  });
  const payload = Array.isArray(rows) && rows[0]?.payload ? rows[0].payload : {};
  return normalizeServiceSettings(payload);
}

async function loadLastSentSlot() {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${encodeURIComponent(LAST_SENT_KEY)}&limit=1`, {
    method: "GET",
  });
  if (!Array.isArray(rows) || !rows[0]) return { id: "", slotKey: "" };
  return {
    id: clean(rows[0]?.id),
    slotKey: clean(rows[0]?.payload?.slotKey),
  };
}

async function saveLastSentSlot(id, slotKey, sentAt, sentCount, providerCount) {
  const payload = {
    slotKey: clean(slotKey),
    sentAt,
    sentCount: Number(sentCount || 0),
    providerCount: Number(providerCount || 0),
  };
  if (id) {
    await restQuery(`app_settings?setting_key=eq.${encodeURIComponent(LAST_SENT_KEY)}`, {
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

async function loadSubmittedServices() {
  const rows = await restQuery(
    "services?select=id,request_number,service_type,customer_name,service_date,service_time,pickup_location,dropoff_location,price,status,provider_email&or=(status.eq.Submitted,status.eq.submitted)&order=provider_email.asc,service_date.asc,service_time.asc,request_number.asc",
    { method: "GET" }
  );
  return Array.isArray(rows) ? rows : [];
}

function rowProviderEmail(row, settings) {
  const direct = clean(row?.provider_email).toLowerCase();
  if (isValidEmail(direct)) return direct;
  const fallback = settings.serviceConfigs.find((item) => item.serviceType === clean(row?.service_type))?.providerEmail || "";
  return isValidEmail(fallback) ? fallback : "";
}

function groupRowsByProvider(rows, settings) {
  const groups = new Map();
  let missingProviderCount = 0;
  rows.forEach((row) => {
    const email = rowProviderEmail(row, settings);
    if (!email) {
      missingProviderCount += 1;
      return;
    }
    if (!groups.has(email)) groups.set(email, []);
    groups.get(email).push(row);
  });
  return { groups, missingProviderCount };
}

function buildEmailContent(rows) {
  const subject = "Pedidos de serviço por Aprovar";
  if (!rows.length) {
    return {
      subject,
      html: "<p>Não existem pedidos de serviço em estado Submitted por aprovar.</p>",
      text: "Não existem pedidos de serviço em estado Submitted por aprovar.",
    };
  }
  const htmlRows = rows.map((row) => `<tr>
      <td style="border:1px solid #d8dee4;padding:8px;text-align:left;">${escapeHtml(clean(row?.service_type) || "-")}</td>
      <td style="border:1px solid #d8dee4;padding:8px;text-align:left;">${escapeHtml(clean(row?.customer_name) || "-")}</td>
      <td style="border:1px solid #d8dee4;padding:8px;text-align:left;">${escapeHtml(formatDateDisplay(row?.service_date))}</td>
      <td style="border:1px solid #d8dee4;padding:8px;text-align:left;">${escapeHtml(formatTimeDisplay(row?.service_time))}</td>
      <td style="border:1px solid #d8dee4;padding:8px;text-align:left;">${escapeHtml(clean(row?.pickup_location) || "-")}</td>
      <td style="border:1px solid #d8dee4;padding:8px;text-align:left;">${escapeHtml(clean(row?.dropoff_location) || "-")}</td>
      <td style="border:1px solid #d8dee4;padding:8px;text-align:left;">${escapeHtml(formatMoney(row?.price))}</td>
      <td style="border:1px solid #d8dee4;padding:8px;text-align:left;">${escapeHtml(clean(row?.status) || "-")}</td>
      <td style="border:1px solid #d8dee4;padding:8px;text-align:left;"><a href="${escapeHtml(serviceDeepLink(row))}" target="_blank" rel="noopener noreferrer">Open</a></td>
    </tr>`).join("");
  const html = `<p>Os seguintes pedidos de serviço aguardam aprovação:</p>
    <table style="border-collapse:collapse;width:100%;font-family:Arial,sans-serif;font-size:13px;">
      <thead>
        <tr>
          <th style="border:1px solid #d8dee4;padding:8px;background:#f3f4f6;text-align:left;">Type</th>
          <th style="border:1px solid #d8dee4;padding:8px;background:#f3f4f6;text-align:left;">Customer Name</th>
          <th style="border:1px solid #d8dee4;padding:8px;background:#f3f4f6;text-align:left;">Service Date</th>
          <th style="border:1px solid #d8dee4;padding:8px;background:#f3f4f6;text-align:left;">Pickup Time</th>
          <th style="border:1px solid #d8dee4;padding:8px;background:#f3f4f6;text-align:left;">Pickup Place</th>
          <th style="border:1px solid #d8dee4;padding:8px;background:#f3f4f6;text-align:left;">Drop Off</th>
          <th style="border:1px solid #d8dee4;padding:8px;background:#f3f4f6;text-align:left;">Price</th>
          <th style="border:1px solid #d8dee4;padding:8px;background:#f3f4f6;text-align:left;">Current Status</th>
          <th style="border:1px solid #d8dee4;padding:8px;background:#f3f4f6;text-align:left;">Link</th>
        </tr>
      </thead>
      <tbody>${htmlRows}</tbody>
    </table>`;
  const text = [
    "Os seguintes pedidos de serviço aguardam aprovação:",
    "",
    "Type | Customer Name | Service Date | Pickup Time | Pickup Place | Drop Off | Price | Current Status | Link",
    ...rows.map((row) => [
      clean(row?.service_type) || "-",
      clean(row?.customer_name) || "-",
      formatDateDisplay(row?.service_date),
      formatTimeDisplay(row?.service_time),
      clean(row?.pickup_location) || "-",
      clean(row?.dropoff_location) || "-",
      formatMoney(row?.price),
      clean(row?.status) || "-",
      serviceDeepLink(row),
    ].join(" | ")),
  ].join("\n");
  return { subject, html, text };
}

async function sendWithResend({ to, subject, html, text }) {
  const apiKey = process.env.RESEND_API_KEY;
  const rawFrom = process.env.EMAIL_FROM;
  const replyTo = "global@lisboacentralhostel.com";
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
  const from = /<[^>]+>/.test(rawFrom) ? rawFrom : `ACCOM Tools - LCH <${rawFrom}>`;
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
    error.statusCode = response.status;
    throw error;
  }
  return payload;
}

module.exports = async function handler(req, res) {
  try {
    const authHeader = clean(req.headers.authorization || "");
    const bearerToken = authHeader.startsWith("Bearer ") ? authHeader.slice(7).trim() : "";
    const cronSecret = clean(process.env.CRON_SECRET);
    const userAgent = clean(req.headers["user-agent"]).toLowerCase();
    const isCronRequest = !!req.headers["x-vercel-cron"] || (!!cronSecret && bearerToken === cronSecret) || userAgent.includes("vercel-cron");
    if (!isCronRequest) await requireFeature(req, "settings", "services");

    if (req.method !== "POST") {
      res.status(405).json({ error: "Method not allowed." });
      return;
    }

    const body = await parseBody(req);
    const force = clean(req.query?.force) === "1";
    const isTest = clean(req.query?.test) === "1";
    const now = new Date();
    const settings = await loadServiceSettings();

    if (isTest) {
      const recipient = clean(body?.email || settings.approvalReminderTestEmail).toLowerCase();
      if (!isValidEmail(recipient)) {
        res.status(400).json({ error: "Approval reminder test email is required." });
        return;
      }
      const rows = await loadSubmittedServices();
      const content = buildEmailContent(rows);
      const provider = await sendWithResend({
        to: [recipient],
        subject: content.subject,
        html: content.html,
        text: content.text,
      });
      res.status(200).json({
        ok: true,
        status: "sent",
        test: true,
        recipients: [recipient],
        serviceCount: rows.length,
        providerMessageId: clean(provider?.id),
      });
      return;
    }

    if (!settings.approvalReminderEnabled) {
      res.status(200).json({ ok: true, status: "skipped", reason: "disabled" });
      return;
    }

    const due = shouldSendNow(settings.approvalReminderTime, now, DEFAULT_TZ);
    if (!force && !due.due) {
      res.status(200).json({ ok: true, status: "skipped", reason: due.reason || "time_mismatch" });
      return;
    }

    const lastSent = await loadLastSentSlot();
    if (!force && lastSent.slotKey === due.slotKey) {
      res.status(200).json({ ok: true, status: "skipped", reason: "already_sent_for_slot" });
      return;
    }

    const rows = await loadSubmittedServices();
    if (!rows.length) {
      res.status(200).json({ ok: true, status: "skipped", reason: "no_submitted_services" });
      return;
    }

    const { groups, missingProviderCount } = groupRowsByProvider(rows, settings);
    if (!groups.size) {
      res.status(200).json({ ok: true, status: "skipped", reason: "no_provider_emails", missingProviderCount });
      return;
    }

    let sentServices = 0;
    let providerCount = 0;
    let firstProviderMessageId = "";
    for (const [recipient, providerRows] of groups.entries()) {
      const content = buildEmailContent(providerRows);
      const provider = await sendWithResend({
        to: [recipient],
        subject: content.subject,
        html: content.html,
        text: content.text,
      });
      if (!firstProviderMessageId) firstProviderMessageId = clean(provider?.id);
      sentServices += providerRows.length;
      providerCount += 1;
    }

    if (!force) {
      await saveLastSentSlot(lastSent.id, due.slotKey, now.toISOString(), sentServices, providerCount);
    }

    res.status(200).json({
      ok: true,
      status: "sent",
      providerCount,
      serviceCount: sentServices,
      missingProviderCount,
      providerMessageId: firstProviderMessageId,
    });
  } catch (error) {
    sendError(res, error);
  }
};
