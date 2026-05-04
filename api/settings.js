const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");

const DEFAULT_SETTINGS = {
  general: {
    bakeryEmailConfig: {
      provider: "resend",
      smtpHost: "smtp.gmail.com",
      smtpPort: 465,
      smtpSecure: true,
      smtpUser: "",
      smtpPassword: "",
      fromEmail: "",
      fromName: "Lisboa Central Hostel",
    },
  },
  communications: {
    categories: [
      { name: "Warning", color: "#ffd89b" },
      { name: "Maintenance", color: "#a9f0df" },
      { name: "Information", color: "#add4ff" },
      { name: "very important", color: "#ffb3c2" },
    ],
    emailAutomation: {
      enabled: false,
      frequency: "everyday",
      timeOfDay: "00:00",
      recipients: [],
      frequency2: "everyday",
      timeOfDay2: "00:00",
      recipients2: [],
    },
  },
};

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

function normalizeFrequency(value) {
  return value === "every1h" || value === "every4h" || value === "every8h" ? value : "everyday";
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

function normalizeBakeryEmailProvider(value) {
  return String(value || "").trim().toLowerCase() === "smtp" ? "smtp" : "resend";
}

function normalizeBakerySecure(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  return ["true", "1", "yes", "sim", "on"].includes(String(value || "").trim().toLowerCase());
}

function sanitizeSettings(input) {
  const output = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
  const source = input && typeof input === "object" ? input : {};
  const general = source.general || {};
  const generalEmail = general.bakeryEmailConfig || general.bakery_email_config || {};
  output.general.bakeryEmailConfig.provider = normalizeBakeryEmailProvider(generalEmail.provider);
  output.general.bakeryEmailConfig.smtpHost = String(generalEmail.smtpHost || generalEmail.smtp_host || DEFAULT_SETTINGS.general.bakeryEmailConfig.smtpHost).trim();
  output.general.bakeryEmailConfig.smtpPort = Math.max(1, Number.parseInt(generalEmail.smtpPort ?? generalEmail.smtp_port ?? DEFAULT_SETTINGS.general.bakeryEmailConfig.smtpPort, 10) || DEFAULT_SETTINGS.general.bakeryEmailConfig.smtpPort);
  output.general.bakeryEmailConfig.smtpSecure = generalEmail.smtpSecure === undefined && generalEmail.smtp_secure === undefined
    ? !!DEFAULT_SETTINGS.general.bakeryEmailConfig.smtpSecure
    : normalizeBakerySecure(generalEmail.smtpSecure ?? generalEmail.smtp_secure);
  output.general.bakeryEmailConfig.smtpUser = String(generalEmail.smtpUser || generalEmail.smtp_user || "").trim().toLowerCase();
  output.general.bakeryEmailConfig.smtpPassword = String(generalEmail.smtpPassword ?? generalEmail.smtp_password ?? "");
  output.general.bakeryEmailConfig.fromEmail = String(generalEmail.fromEmail || generalEmail.from_email || "").trim().toLowerCase();
  output.general.bakeryEmailConfig.fromName = String(generalEmail.fromName || generalEmail.from_name || DEFAULT_SETTINGS.general.bakeryEmailConfig.fromName).trim() || DEFAULT_SETTINGS.general.bakeryEmailConfig.fromName;
  const comm = source.communications || {};
  const categories = Array.isArray(comm.categories) ? comm.categories : [];
  const seen = new Set();

  const cleanCategories = categories
    .map((cat) => ({
      name: String(cat?.name || "").trim(),
      color: normalizeHex(cat?.color),
      autoCloseDays: normalizeAutoCloseDays(cat?.autoCloseDays),
    }))
    .filter((cat) => cat.name)
    .filter((cat) => {
      const key = cat.name.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });

  if (cleanCategories.length > 0) {
    output.communications.categories = cleanCategories;
  }

  const email = comm.emailAutomation || {};
  output.communications.emailAutomation.enabled = !!email.enabled;
  output.communications.emailAutomation.frequency = normalizeFrequency(email.frequency);
  output.communications.emailAutomation.timeOfDay = normalizeTime(email.timeOfDay);
  output.communications.emailAutomation.recipients = normalizeRecipients(email.recipients);
  output.communications.emailAutomation.frequency2 = normalizeFrequency(email.frequency2);
  output.communications.emailAutomation.timeOfDay2 = normalizeTime(email.timeOfDay2);
  output.communications.emailAutomation.recipients2 = normalizeRecipients(email.recipients2);
  return output;
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "settings", "communications");

    if (req.method === "GET") {
      const rows = await restQuery("app_settings?select=payload&setting_key=eq.communications&limit=1", {
        method: "GET",
      });
      const payload = Array.isArray(rows) && rows[0] && rows[0].payload ? rows[0].payload : DEFAULT_SETTINGS;
      res.status(200).json({ settings: sanitizeSettings(payload) });
      return;
    }

    if (req.method === "PUT") {
      const body = await parseBody(req);
      const safe = sanitizeSettings(body?.settings);
      const existing = await restQuery(
        "app_settings?select=id&setting_key=eq.communications&limit=1",
        { method: "GET" }
      );

      if (Array.isArray(existing) && existing[0]) {
        await restQuery("app_settings?setting_key=eq.communications", {
          method: "PATCH",
          body: { payload: safe, updated_at: new Date().toISOString() },
        });
      } else {
        await restQuery("app_settings", {
          method: "POST",
          body: [{ setting_key: "communications", payload: safe }],
        });
      }
      res.status(200).json({ settings: safe });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
