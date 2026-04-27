const {
  authAdminQuery,
  cleanText,
  loadProfiles,
  parseBody,
  requireFeature,
  restQuery,
  sendError,
} = require("./_supabase");

const DEFAULT_PRICE_MATRIX = {
  oneWay: {
    "1-3": 35,
    "4-7": 55,
    "8-11": 90,
    "12-16": 110,
  },
  returnTrip: {
    "1-3": 63,
    "4-7": 99,
    "8-11": 162,
    "12-16": 198,
  },
};

const DEFAULT_SERVICE_TYPES = [
  {
    id: "airport-transfer",
    serviceType: "Airport Transfer",
    providerUserId: "",
    providerEmail: "odete@netcabo.pt",
    airportTransfer: true,
    hasReturn: true,
    approvedByDefault: false,
    priceMode: "airport_matrix",
    priceMatrix: DEFAULT_PRICE_MATRIX,
    confirmationTemplate: `Dear {{customer_name}},

Your transfer is confirmed with the following details:

{{service_table}}

For pick-up at the airport, the transfer company will be waiting for you at arrivals with your name on a board. The pickup time is based on the flight arrival time and the transfer company will track your flight.

For pick-up in other locations, please be ready 5 minutes before the scheduled pickup time.

If you have any trouble, please use the shuttle service number: +351 917921578. It is also available on WhatsApp. Please contact the company if you have any problem finding them, otherwise we will have to charge the service amount.

Payment should be made at the check-in desk, not to the driver.

Cancellation Policy

Any cancellations must be informed 48h before service, otherwise the full amount of the service will be charged.

Best regards,
Lisboa Central Hostel`,
  },
  {
    id: "other-transfer",
    serviceType: "Other Transfer",
    providerUserId: "",
    providerEmail: "odete@netcabo.pt",
    airportTransfer: true,
    hasReturn: true,
    approvedByDefault: false,
    priceMode: "open",
    priceMatrix: {},
    confirmationTemplate: `Dear {{customer_name}},

Your transfer is confirmed with the following details:

{{service_table}}

For pick-up at the airport, the transfer company will be waiting for you at arrivals with your name on a board. The pickup time is based on the flight arrival time and the transfer company will track your flight.

For pick-up in other locations, please be ready 5 minutes before the scheduled pickup time.

If you have any trouble, please use the shuttle service number: +351 917921578. It is also available on WhatsApp. Please contact the company if you have any problem finding them, otherwise we will have to charge the service amount.

Payment should be made at the check-in desk, not to the driver.

Cancellation Policy

Any cancellations must be informed 48h before service, otherwise the full amount of the service will be charged.

Best regards,
Lisboa Central Hostel`,
  },
  {
    id: "tour",
    serviceType: "Tour",
    providerUserId: "",
    providerEmail: "",
    airportTransfer: false,
    hasReturn: false,
    approvedByDefault: false,
    priceMode: "open",
    priceMatrix: {},
    confirmationTemplate: `Dear {{customer_name}},

Your service is confirmed with the following details:

{{service_table}}

Please be ready 5 minutes before the scheduled pickup time.

If you have any trouble, please use the shuttle service number: +351 917921578. It is also available on WhatsApp. Please contact the company if you have any problem finding them, otherwise we will have to charge the service amount.

Cancellation Policy

Any cancellations must be informed 48h before service, otherwise the full amount of the service will be charged.

Best regards,
Lisboa Central Hostel`,
  },
  {
    id: "boat-tour",
    serviceType: "Boat Tour",
    providerUserId: "",
    providerEmail: "",
    airportTransfer: false,
    hasReturn: false,
    approvedByDefault: false,
    priceMode: "open",
    priceMatrix: {},
    confirmationTemplate: `Dear {{customer_name}},

Your service is confirmed with the following details:

{{service_table}}

Please be ready 5 minutes before the scheduled pickup time.

If you have any trouble, please use the shuttle service number: +351 917921578. It is also available on WhatsApp. Please contact the company if you have any problem finding them, otherwise we will have to charge the service amount.

Cancellation Policy

Any cancellations must be informed 48h before service, otherwise the full amount of the service will be charged.

Best regards,
Lisboa Central Hostel`,
  },
];

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

function normalizeBool(value, fallback = false) {
  if (typeof value === "boolean") return value;
  const raw = cleanText(value).toLowerCase();
  if (["true", "1", "yes"].includes(raw)) return true;
  if (["false", "0", "no"].includes(raw)) return false;
  return fallback;
}

function normalizeNumber(value, fallback = 0) {
  const num = Number(String(value ?? "").replace(",", "."));
  return Number.isFinite(num) ? num : fallback;
}

function normalizePriceMode(value) {
  return cleanText(value) === "airport_matrix" ? "airport_matrix" : "open";
}

function slugify(value) {
  return cleanText(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function normalizePriceMatrix(value) {
  const source = value && typeof value === "object" && !Array.isArray(value) ? value : {};
  const normalizeBandGroup = (item) => {
    const group = item && typeof item === "object" && !Array.isArray(item) ? item : {};
    return {
      "1-3": Math.max(0, normalizeNumber(group["1-3"], 0)),
      "4-7": Math.max(0, normalizeNumber(group["4-7"], 0)),
      "8-11": Math.max(0, normalizeNumber(group["8-11"], 0)),
      "12-16": Math.max(0, normalizeNumber(group["12-16"], 0)),
    };
  };
  return {
    oneWay: normalizeBandGroup(source.oneWay),
    returnTrip: normalizeBandGroup(source.returnTrip),
  };
}

function normalizeServiceConfig(item = {}) {
  const serviceType = cleanText(item.serviceType || item.service_type);
  const airportTransfer = normalizeBool(item.airportTransfer ?? item.airport_transfer);
  return {
    id: cleanText(item.id) || slugify(serviceType),
    serviceType,
    providerUserId: cleanText(item.providerUserId || item.provider_user_id),
    providerEmail: cleanText(item.providerEmail || item.provider_email).toLowerCase(),
    airportTransfer,
    hasReturn: normalizeBool(item.hasReturn ?? item.has_return),
    approvedByDefault: normalizeBool(item.approvedByDefault ?? item.approved_by_default),
    priceMode: normalizePriceMode(item.priceMode || item.price_mode),
    priceMatrix: normalizePriceMatrix(item.priceMatrix || item.price_matrix),
    confirmationTemplate: cleanText(item.confirmationTemplate || item.confirmation_template) || (DEFAULT_SERVICE_TYPES.find((entry) => cleanText(entry.id) === cleanText(item.id) || cleanText(entry.serviceType) === serviceType)?.confirmationTemplate || (airportTransfer ? DEFAULT_SERVICE_TYPES[0].confirmationTemplate : DEFAULT_SERVICE_TYPES[2].confirmationTemplate)),
  };
}

function sanitizeSettings(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const configs = Array.isArray(source.serviceConfigs || source.service_configs)
    ? source.serviceConfigs || source.service_configs
    : [];
  const seen = new Set();
  const normalized = configs
    .map(normalizeServiceConfig)
    .filter((item) => item.serviceType)
    .filter((item) => {
      const key = item.id || slugify(item.serviceType);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  return {
    automaticEmailRecipients: normalizeRecipients(source.automaticEmailRecipients || source.automatic_email_recipients),
    serviceConfigs: normalized.length ? normalized : DEFAULT_SERVICE_TYPES,
  };
}

function filterSettingsForProvider(settings, user) {
  const userId = cleanText(user?.id);
  const userEmail = cleanText(user?.email).toLowerCase();
  return {
    automaticEmailRecipients: normalizeRecipients(settings?.automaticEmailRecipients),
    serviceConfigs: (Array.isArray(settings?.serviceConfigs) ? settings.serviceConfigs : []).filter(
      (item) => cleanText(item.providerUserId) === userId || cleanText(item.providerEmail).toLowerCase() === userEmail
    ),
  };
}

async function loadServiceProviderUsers() {
  const profiles = await loadProfiles();
  const providerProfile = profiles.find((profile) => cleanText(profile.name).toLowerCase() === "service provider");
  if (!providerProfile?.id) return [];
  const assignments = await restQuery(
    `user_profile_assignments?select=user_id,profile_id&profile_id=eq.${encodeURIComponent(providerProfile.id)}`,
    { method: "GET" }
  );
  const providerUserIds = new Set((Array.isArray(assignments) ? assignments : []).map((row) => cleanText(row.user_id)).filter(Boolean));
  if (!providerUserIds.size) return [];
  const payload = await authAdminQuery("users?page=1&per_page=500", { method: "GET" });
  const users = Array.isArray(payload?.users) ? payload.users : [];
  return users
    .filter((user) => providerUserIds.has(cleanText(user?.id)))
    .map((user) => ({
      id: cleanText(user?.id),
      email: cleanText(user?.email).toLowerCase(),
    }))
    .filter((user) => user.id && user.email)
    .sort((a, b) => a.email.localeCompare(b.email));
}

module.exports = async function handler(req, res) {
  try {
    let canEdit = false;
    let auth = null;
    try {
      auth = await requireFeature(req, "settings", "services");
      canEdit = true;
    } catch {
      auth = await requireFeature(req, "app", "services");
    }

    if (req.method === "GET") {
      const rows = await restQuery("app_settings?select=payload&setting_key=eq.services&limit=1", { method: "GET" });
      const payload = Array.isArray(rows) && rows[0]?.payload ? rows[0].payload : {};
      const settings = sanitizeSettings(payload);
      const profileName = cleanText(auth?.access?.profile?.name).toLowerCase();
      const visibleSettings = !canEdit && profileName === "service provider" ? filterSettingsForProvider(settings, auth?.user) : settings;
      const providers = canEdit ? await loadServiceProviderUsers() : [];
      res.status(200).json({ settings: visibleSettings, providers });
      return;
    }

    if (req.method === "PUT") {
      if (!canEdit) await requireFeature(req, "settings", "services");
      const body = await parseBody(req);
      const safe = sanitizeSettings(body?.settings);
      const existing = await restQuery("app_settings?select=id&setting_key=eq.services&limit=1", { method: "GET" });
      if (Array.isArray(existing) && existing[0]) {
        await restQuery("app_settings?setting_key=eq.services", {
          method: "PATCH",
          body: { payload: safe, updated_at: new Date().toISOString() },
        });
      } else {
        await restQuery("app_settings", {
          method: "POST",
          body: [{ setting_key: "services", payload: safe }],
        });
      }
      const providers = await loadServiceProviderUsers();
      res.status(200).json({ settings: safe, providers });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
