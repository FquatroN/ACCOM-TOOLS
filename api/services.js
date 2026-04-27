const { cleanText, parseBody, requireFeature, restQuery, sendError } = require("./_supabase");

function normalizeDate(value) {
  const raw = cleanText(value);
  if (!raw) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return null;
  return date.toISOString().slice(0, 10);
}

function normalizeTime(value) {
  const raw = cleanText(value);
  if (!raw) return null;
  if (/^\d{2}:\d{2}/.test(raw)) return raw.slice(0, 5);
  const date = new Date(`1970-01-01T${raw}`);
  if (Number.isNaN(date.getTime())) return raw.slice(0, 5) || null;
  return `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
}

function normalizeNumber(value, fallback = 0) {
  const num = Number(String(value ?? "").replace(",", "."));
  return Number.isFinite(num) ? num : fallback;
}

function normalizeBool(value, fallback = false) {
  if (typeof value === "boolean") return value;
  const raw = cleanText(value).toLowerCase();
  if (["true", "1", "yes"].includes(raw)) return true;
  if (["false", "0", "no"].includes(raw)) return false;
  return fallback;
}

function normalizeStatus(value) {
  const raw = cleanText(value).toLowerCase();
  if (raw === "approved") return "Approved";
  if (raw === "cancelled" || raw === "canceled") return "Cancelled";
  if (raw === "completed") return "Completed";
  return "Submitted";
}

function normalizeServiceType(value) {
  const raw = cleanText(value);
  if (!raw) return "";
  return raw;
}

function normalizeAuditLog(value) {
  return (Array.isArray(value) ? value : [])
    .map((item) => ({
      at: cleanText(item?.at),
      action: cleanText(item?.action),
      user: cleanText(item?.user),
      summary: cleanText(item?.summary),
    }))
    .filter((item) => item.at && item.action)
    .slice(-50);
}

function normalizeServiceConfigs(payload) {
  const source = payload && typeof payload === "object" ? payload : {};
  const configs = Array.isArray(source.serviceConfigs || source.service_configs)
    ? source.serviceConfigs || source.service_configs
    : [];
  return configs
    .map((item) => ({
      serviceType: cleanText(item?.serviceType || item?.service_type),
      providerUserId: cleanText(item?.providerUserId || item?.provider_user_id),
      providerEmail: cleanText(item?.providerEmail || item?.provider_email).toLowerCase(),
    }))
    .filter((item) => item.serviceType);
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanText(value));
}

function normalizeRecipients(value) {
  const source = Array.isArray(value) ? value.join(",") : String(value || "");
  const seen = new Set();
  return source
    .split(/[\n,;]/)
    .map((item) => cleanText(item).toLowerCase())
    .filter((item) => isValidEmail(item))
    .filter((item) => {
      if (seen.has(item)) return false;
      seen.add(item);
      return true;
    });
}

function normalizeServiceSettingsPayload(payload) {
  const source = payload && typeof payload === "object" ? payload : {};
  return {
    automaticEmailRecipients: normalizeRecipients(source.automaticEmailRecipients || source.automatic_email_recipients),
    serviceConfigs: normalizeServiceConfigs(source),
  };
}

function isValidInternationalPhone(value) {
  const raw = cleanText(value);
  if (!raw) return true;
  if (!/^\+[0-9][0-9\s().-]{5,20}$/.test(raw)) return false;
  const digits = raw.replace(/\D/g, "");
  return digits.length >= 6 && digits.length <= 15;
}

async function loadServiceSettingsPayload() {
  const rows = await restQuery("app_settings?select=payload&setting_key=eq.services&limit=1", { method: "GET" });
  const payload = Array.isArray(rows) && rows[0]?.payload ? rows[0].payload : {};
  return normalizeServiceSettingsPayload(payload);
}

async function loadServiceScope(user, access) {
  const profileName = cleanText(access?.profile?.name).toLowerCase();
  const isServiceProvider = profileName === "service provider";
  const canManageSettings = (access?.settingsFeatures || []).includes("services");
  if (!isServiceProvider || canManageSettings) return { restricted: false, allowedTypes: new Set() };
  const settings = await loadServiceSettingsPayload();
  const configs = settings.serviceConfigs;
  const allowedTypes = new Set(
    configs
      .filter((item) => item.providerUserId === cleanText(user?.id) || item.providerEmail === cleanText(user?.email).toLowerCase())
      .map((item) => item.serviceType)
      .filter(Boolean)
  );
  return {
    restricted: true,
    allowedTypes,
    userId: cleanText(user?.id),
    userEmail: cleanText(user?.email).toLowerCase(),
  };
}

function canAccessServiceRow(row, scope) {
  if (!scope?.restricted) return true;
  return (
    scope.allowedTypes.has(cleanText(row?.service_type || row?.serviceType)) ||
    cleanText(row?.provider_user_id || row?.providerUserId) === cleanText(scope.userId) ||
    cleanText(row?.provider_email || row?.providerEmail).toLowerCase() === cleanText(scope.userEmail).toLowerCase()
  );
}

function withSelect(path) {
  return `${path}${path.includes("?") ? "&" : "?"}select=id,request_number,service_type,customer_name,customer_email,customer_phone,pax,notes,service_date,service_time,pickup_location,dropoff_location,flight_number,has_return,return_pickup_location,return_dropoff_location,return_date,return_time,return_flight_number,price,status,provider_user_id,provider_email,audit_log,created_at,updated_at&order=service_date.asc,service_time.asc,request_number.asc`;
}

function sanitizeService(input = {}, existing = null) {
  const serviceType = normalizeServiceType(input.service_type ?? input.serviceType ?? existing?.service_type);
  const customerName = cleanText(input.customer_name ?? input.customerName ?? existing?.customer_name);
  const customerEmail = cleanText(input.customer_email ?? input.customerEmail ?? existing?.customer_email).toLowerCase();
  const customerPhone = cleanText(input.customer_phone ?? input.customerPhone ?? existing?.customer_phone);
  const pax = Math.max(1, Math.min(60, Math.round(normalizeNumber(input.pax ?? input.nr_persons ?? input.nrPersons, existing?.pax ?? 1))));
  const serviceDate = normalizeDate(input.service_date ?? input.date ?? input.serviceDate ?? existing?.service_date);
  const serviceTime = normalizeTime(input.service_time ?? input.time ?? input.serviceTime ?? existing?.service_time);
  const price = Math.max(0, normalizeNumber(input.price, existing?.price ?? 0));
  const status = normalizeStatus(input.status ?? existing?.status);
  const hasReturn = normalizeBool(input.has_return ?? input.hasReturn ?? existing?.has_return);
  const auditLog = normalizeAuditLog(input.audit_log ?? input.auditLog ?? existing?.audit_log);

  if (!serviceType) throw Object.assign(new Error("Service Type is required."), { statusCode: 400 });
  if (!customerName) throw Object.assign(new Error("Customer Name is required."), { statusCode: 400 });
  if (!serviceDate) throw Object.assign(new Error("Date is required."), { statusCode: 400 });
  if (!serviceTime) throw Object.assign(new Error("Time is required."), { statusCode: 400 });
  if (!pax) throw Object.assign(new Error("Nr persons is required."), { statusCode: 400 });
  if (customerEmail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(customerEmail)) {
    throw Object.assign(new Error("Customer Email is invalid."), { statusCode: 400 });
  }
  if (customerPhone && !isValidInternationalPhone(customerPhone)) {
    throw Object.assign(new Error("Customer Phone must include country code, for example +351 912 345 678."), { statusCode: 400 });
  }

  return {
    service_type: serviceType,
    customer_name: customerName,
    customer_email: customerEmail,
    customer_phone: customerPhone,
    pax,
    notes: cleanText(input.notes ?? existing?.notes),
    service_date: serviceDate,
    service_time: serviceTime,
    pickup_location: cleanText(input.pickup_location ?? input.pickupLocation ?? existing?.pickup_location),
    dropoff_location: cleanText(input.dropoff_location ?? input.dropoffLocation ?? existing?.dropoff_location),
    flight_number: cleanText(input.flight_number ?? input.flightNumber ?? existing?.flight_number),
    has_return: hasReturn,
    return_pickup_location: hasReturn ? cleanText(input.return_pickup_location ?? input.returnPickupLocation ?? existing?.return_pickup_location) : "",
    return_dropoff_location: hasReturn ? cleanText(input.return_dropoff_location ?? input.returnDropoffLocation ?? existing?.return_dropoff_location) : "",
    return_date: hasReturn ? normalizeDate(input.return_date ?? input.returnDate ?? existing?.return_date) : null,
    return_time: hasReturn ? normalizeTime(input.return_time ?? input.returnTime ?? existing?.return_time) : null,
    return_flight_number: hasReturn ? cleanText(input.return_flight_number ?? input.returnFlightNumber ?? existing?.return_flight_number) : "",
    price,
    status,
    provider_user_id: cleanText(input.provider_user_id ?? input.providerUserId ?? existing?.provider_user_id) || null,
    provider_email: cleanText(input.provider_email ?? input.providerEmail ?? existing?.provider_email).toLowerCase(),
    audit_log: auditLog,
  };
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function formatDateDisplay(value) {
  const raw = cleanText(value);
  if (!raw) return "-";
  const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return raw;
  return `${match[3]}/${match[2]}/${match[1]}`;
}

function formatTimeDisplay(value) {
  const raw = cleanText(value);
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

function formatBool(value) {
  return value ? "Sim" : "Nao";
}

function serviceDetailRows(row) {
  const items = [
    ["Request #", cleanText(row?.request_number)],
    ["Service Type", cleanText(row?.service_type)],
    ["Customer", cleanText(row?.customer_name)],
    ["Customer Email", cleanText(row?.customer_email)],
    ["Customer Phone", cleanText(row?.customer_phone)],
    ["Date", formatDateDisplay(row?.service_date)],
    ["Time", formatTimeDisplay(row?.service_time)],
    ["Pax", String(Math.max(1, Math.round(Number(row?.pax || 1))))],
    ["Pickup", cleanText(row?.pickup_location)],
    ["Drop Off", cleanText(row?.dropoff_location)],
    ["Flight Number", cleanText(row?.flight_number)],
    ["Return?", formatBool(!!row?.has_return)],
    ["Return Pickup", cleanText(row?.return_pickup_location)],
    ["Return Drop Off", cleanText(row?.return_dropoff_location)],
    ["Return Date", formatDateDisplay(row?.return_date)],
    ["Return Time", formatTimeDisplay(row?.return_time)],
    ["Return Flight", cleanText(row?.return_flight_number)],
    ["Price", formatMoney(row?.price)],
    ["Status", cleanText(row?.status)],
    ["Service Provider", cleanText(row?.provider_email)],
    ["Notes", cleanText(row?.notes)],
  ];
  return items.filter(([, value]) => cleanText(value) && value !== "-");
}

function serviceChangedFields(before, after) {
  const changed = [];
  const push = (label, oldValue, newValue) => {
    if (oldValue !== newValue) changed.push({ label, oldValue, newValue });
  };
  push("Service Type", cleanText(before?.service_type), cleanText(after?.service_type));
  push("Customer", cleanText(before?.customer_name), cleanText(after?.customer_name));
  push("Customer Email", cleanText(before?.customer_email), cleanText(after?.customer_email));
  push("Customer Phone", cleanText(before?.customer_phone), cleanText(after?.customer_phone));
  push("Date", formatDateDisplay(before?.service_date), formatDateDisplay(after?.service_date));
  push("Time", formatTimeDisplay(before?.service_time), formatTimeDisplay(after?.service_time));
  push("Pax", String(Math.max(1, Math.round(Number(before?.pax || 1)))), String(Math.max(1, Math.round(Number(after?.pax || 1)))));
  push("Pickup", cleanText(before?.pickup_location) || "-", cleanText(after?.pickup_location) || "-");
  push("Drop Off", cleanText(before?.dropoff_location) || "-", cleanText(after?.dropoff_location) || "-");
  push("Flight Number", cleanText(before?.flight_number) || "-", cleanText(after?.flight_number) || "-");
  push("Return?", formatBool(!!before?.has_return), formatBool(!!after?.has_return));
  push("Return Pickup", cleanText(before?.return_pickup_location) || "-", cleanText(after?.return_pickup_location) || "-");
  push("Return Drop Off", cleanText(before?.return_dropoff_location) || "-", cleanText(after?.return_dropoff_location) || "-");
  push("Return Date", formatDateDisplay(before?.return_date), formatDateDisplay(after?.return_date));
  push("Return Time", formatTimeDisplay(before?.return_time), formatTimeDisplay(after?.return_time));
  push("Return Flight", cleanText(before?.return_flight_number) || "-", cleanText(after?.return_flight_number) || "-");
  push("Price", formatMoney(before?.price), formatMoney(after?.price));
  push("Status", cleanText(before?.status) || "-", cleanText(after?.status) || "-");
  push("Service Provider", cleanText(before?.provider_email) || "-", cleanText(after?.provider_email) || "-");
  push("Notes", cleanText(before?.notes) || "-", cleanText(after?.notes) || "-");
  return changed;
}

function serviceEmailSubject(row, isUpdate) {
  const prefix = isUpdate ? "Alteracao" : "Novo Pedido";
  return `${prefix}: ${cleanText(row?.service_type)}, ${cleanText(row?.request_number)}, - ${cleanText(row?.customer_name)} - ${formatDateDisplay(row?.service_date)}, ${formatTimeDisplay(row?.service_time)}, Pax: ${Math.max(1, Math.round(Number(row?.pax || 1)))}`;
}

function serviceCreatedEmailContent(row) {
  const rows = serviceDetailRows(row)
    .map(
      ([label, value]) =>
        `<tr><td style="border:1px solid #d1d5db;padding:6px;font-weight:600;background:#f8fafc;">${escapeHtml(label)}</td><td style="border:1px solid #d1d5db;padding:6px;">${escapeHtml(value)}</td></tr>`
    )
    .join("");
  const html = `<!doctype html>
<html>
  <body style="font-family:Arial,sans-serif;color:#1f2937;">
    <p>Foi criado um novo pedido de servico com os seguintes dados:</p>
    <table style="border-collapse:collapse;width:100%;max-width:760px;">${rows}</table>
  </body>
</html>`;
  const text = [
    "Foi criado um novo pedido de servico com os seguintes dados:",
    ...serviceDetailRows(row).map(([label, value]) => `${label}: ${value}`),
  ].join("\n");
  return { html, text };
}

function serviceUpdatedEmailContent(before, after) {
  const changes = serviceChangedFields(before, after);
  const rows = (changes.length ? changes : [{ label: "Alteracoes", oldValue: "-", newValue: "Sem alteracoes visiveis" }])
    .map(
      (item) =>
        `<tr><td style="border:1px solid #d1d5db;padding:6px;font-weight:600;background:#f8fafc;">${escapeHtml(item.label)}</td><td style="border:1px solid #d1d5db;padding:6px;">${escapeHtml(item.oldValue)}</td><td style="border:1px solid #d1d5db;padding:6px;">${escapeHtml(item.newValue)}</td></tr>`
    )
    .join("");
  const html = `<!doctype html>
<html>
  <body style="font-family:Arial,sans-serif;color:#1f2937;">
    <p>Foram feitas as seguintes alteracoes neste servico:</p>
    <table style="border-collapse:collapse;width:100%;max-width:860px;">
      <thead>
        <tr>
          <th style="border:1px solid #d1d5db;padding:6px;text-align:left;">Campo</th>
          <th style="border:1px solid #d1d5db;padding:6px;text-align:left;">Valor anterior</th>
          <th style="border:1px solid #d1d5db;padding:6px;text-align:left;">Novo valor</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  </body>
</html>`;
  const text = [
    "Foram feitas as seguintes alteracoes neste servico:",
    ...(changes.length
      ? changes.map((item) => `${item.label}: ${item.oldValue} -> ${item.newValue}`)
      : ["Alteracoes: - -> Sem alteracoes visiveis"]),
  ].join("\n");
  return { html, text };
}

async function sendWithResend({ to, subject, html, text }) {
  const apiKey = process.env.RESEND_API_KEY;
  const rawFrom = process.env.EMAIL_FROM;
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
  const from = /<[^>]+>/.test(rawFrom) ? rawFrom : `ACOOM Tools <${rawFrom}>`;
  const response = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ from, to, subject, html, text }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload?.message || payload?.error || `Email provider failed (${response.status})`);
    error.statusCode = response.status;
    throw error;
  }
  return payload;
}

async function sendServiceNotification({ previousRow, row, settings }) {
  const providerEmail =
    cleanText(row?.provider_email).toLowerCase() ||
    cleanText((settings?.serviceConfigs || []).find((item) => cleanText(item.serviceType) === cleanText(row?.service_type))?.providerEmail).toLowerCase();
  const recipients = Array.from(new Set([
    ...normalizeRecipients(settings?.automaticEmailRecipients),
    providerEmail,
  ].filter(Boolean)));
  if (!recipients.length) return { skipped: true };
  const content = previousRow ? serviceUpdatedEmailContent(previousRow, row) : serviceCreatedEmailContent(row);
  return sendWithResend({
    to: recipients,
    subject: serviceEmailSubject(row, !!previousRow),
    html: content.html,
    text: content.text,
  });
}

async function loadExisting(id) {
  const rows = await restQuery(
    `services?select=id,request_number,service_type,customer_name,customer_email,customer_phone,pax,notes,service_date,service_time,pickup_location,dropoff_location,flight_number,has_return,return_pickup_location,return_dropoff_location,return_date,return_time,return_flight_number,price,status,provider_user_id,provider_email,audit_log&id=eq.${encodeURIComponent(id)}&limit=1`,
    { method: "GET" }
  );
  return Array.isArray(rows) && rows[0] ? rows[0] : null;
}

module.exports = async function handler(req, res) {
  try {
    const { user, access } = await requireFeature(req, "app", "services");
    const scope = await loadServiceScope(user, access);

    if (req.method === "GET") {
      const rows = await restQuery(withSelect("services"), { method: "GET" });
      const visibleRows = (Array.isArray(rows) ? rows : []).filter((row) => canAccessServiceRow(row, scope));
      res.status(200).json({ rows: visibleRows });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      if (scope.restricted && !scope.allowedTypes.has(cleanText(body?.serviceType || body?.service_type))) {
        res.status(403).json({ error: "You do not have permission for this service type." });
        return;
      }
      const serviceSettings = await loadServiceSettingsPayload();
      const created = await restQuery(withSelect("services"), {
        method: "POST",
        body: [sanitizeService(body)],
        preferRepresentation: true,
      });
      const row = Array.isArray(created) && created[0] ? created[0] : null;
      let emailWarning = "";
      if (row) {
        try {
          await sendServiceNotification({ previousRow: null, row, settings: serviceSettings });
        } catch (error) {
          emailWarning = error.message || "Could not send notification email.";
        }
      }
      res.status(200).json({ row, emailWarning });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanText(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Missing id query parameter." });
        return;
      }
      const existing = await loadExisting(id);
      if (!existing) {
        res.status(404).json({ error: "Service not found." });
        return;
      }
      if (!canAccessServiceRow(existing, scope)) {
        res.status(403).json({ error: "You do not have permission for this service." });
        return;
      }
      const body = await parseBody(req);
      if (scope.restricted && !scope.allowedTypes.has(cleanText(body?.serviceType || body?.service_type || existing?.service_type))) {
        res.status(403).json({ error: "You do not have permission for this service type." });
        return;
      }
      const serviceSettings = await loadServiceSettingsPayload();
      const updated = await restQuery(withSelect(`services?id=eq.${encodeURIComponent(id)}`), {
        method: "PATCH",
        body: sanitizeService(body, existing),
        preferRepresentation: true,
      });
      const row = Array.isArray(updated) && updated[0] ? updated[0] : null;
      let emailWarning = "";
      if (row) {
        try {
          await sendServiceNotification({ previousRow: existing, row, settings: serviceSettings });
        } catch (error) {
          emailWarning = error.message || "Could not send notification email.";
        }
      }
      res.status(200).json({ row, emailWarning });
      return;
    }

    if (req.method === "DELETE") {
      const id = cleanText(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Missing id query parameter." });
        return;
      }
      const existing = await loadExisting(id);
      if (!existing) {
        res.status(404).json({ error: "Service not found." });
        return;
      }
      if (!canAccessServiceRow(existing, scope)) {
        res.status(403).json({ error: "You do not have permission for this service." });
        return;
      }
      await restQuery(`services?id=eq.${encodeURIComponent(id)}`, { method: "DELETE" });
      res.status(200).json({ ok: true });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
