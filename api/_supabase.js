const REQUIRED_ENV_VARS = ["SUPABASE_URL", "SUPABASE_ANON_KEY", "SUPABASE_SERVICE_ROLE_KEY"];
const APP_FEATURES = ["communications", "lost-found", "reviews", "groups", "services"];
const SETTINGS_FEATURES = ["communications", "reviews", "groups", "services", "admin-users"];
const FALLBACK_PROFILE = {
  id: "",
  name: "Full access (fallback)",
  appFeatures: [...APP_FEATURES],
  settingsFeatures: [...SETTINGS_FEATURES],
};

function readEnv() {
  const env = {
    url: process.env.SUPABASE_URL,
    anonKey: process.env.SUPABASE_ANON_KEY,
    serviceRoleKey: process.env.SUPABASE_SERVICE_ROLE_KEY,
  };

  const missing = REQUIRED_ENV_VARS.filter((key) => !process.env[key]);
  if (missing.length > 0) {
    const error = new Error(`Missing server environment variables: ${missing.join(", ")}`);
    error.statusCode = 500;
    throw error;
  }

  return env;
}

function cleanText(value) {
  return String(value ?? "").trim();
}

function normalizeCategory(value) {
  const category = cleanText(value);
  return category || "Information";
}

function normalizeStatus(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (raw === "closed" || raw === "resolved" || raw === "archived") return "Closed";
  return "Open";
}

function toIsoDate(date) {
  const d = new Date(date);
  if (Number.isNaN(d.getTime())) return "";
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function toIsoTime(date) {
  const d = new Date(date);
  if (Number.isNaN(d.getTime())) return "";
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}

function normalizeDate(value) {
  const raw = cleanText(value);
  if (!raw) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const dmyMatch = raw.match(/^(\d{2})[\/.-](\d{2})[\/.-](\d{4})$/);
  if (dmyMatch) return `${dmyMatch[3]}-${dmyMatch[2]}-${dmyMatch[1]}`;
  const portugueseDate = normalizePortugueseDate(raw);
  if (portugueseDate) return portugueseDate;
  return toIsoDate(raw) || raw;
}

function normalizePortugueseDate(value) {
  const raw = cleanText(value).toLowerCase().replace(/[\u2009\u202f]/g, " ").replace(/\s+/g, " ");
  const dateSeparator = "[\\s\\u2009\\u202f]*[^\\da-z.]+[\\s\\u2009\\u202f]*";
  const crossYearMatch = raw.match(new RegExp(`(\\d{1,2})\\s+de\\s+([a-zç.]+)\\s+de\\s+\\d{4}${dateSeparator}(\\d{1,2})\\s+de\\s+([a-zç.]+)\\s+de\\s+(\\d{4})`, "i"));
  if (crossYearMatch) {
    return formatPortugueseDateParts(crossYearMatch[3], crossYearMatch[4], crossYearMatch[5]);
  }
  const crossMonthMatch = raw.match(new RegExp(`(\\d{1,2})\\s+de\\s+([a-zç.]+)${dateSeparator}(\\d{1,2})\\s+de\\s+([a-zç.]+)\\s+de\\s+(\\d{4})`, "i"));
  if (crossMonthMatch) {
    return formatPortugueseDateParts(crossMonthMatch[3], crossMonthMatch[4], crossMonthMatch[5]);
  }
  const match = raw.match(new RegExp(`(\\d{1,2})(?:${dateSeparator}(\\d{1,2}))?\\s+(?:de\\s+)?([a-zç.]+)\\s+de\\s+(\\d{4})`, "i"));
  if (!match) return "";
  const day = match[2] || match[1];
  return formatPortugueseDateParts(day, match[3], match[4]);
}

function formatPortugueseDateParts(day, monthName, year) {
  const monthKey = cleanText(monthName).toLowerCase().replace(".", "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const months = {
    jan: 1, janeiro: 1,
    fev: 2, fevereiro: 2,
    mar: 3, marco: 3,
    abr: 4, abril: 4,
    mai: 5, maio: 5,
    jun: 6, junho: 6,
    jul: 7, julho: 7,
    ago: 8, agosto: 8,
    set: 9, setembro: 9,
    out: 10, outubro: 10,
    nov: 11, novembro: 11,
    dez: 12, dezembro: 12,
  };
  const month = months[monthKey];
  if (!month) return "";
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function normalizeTime(value) {
  const raw = cleanText(value);
  if (!raw) return "";
  if (/^\d{2}:\d{2}/.test(raw)) return raw.slice(0, 5);
  return toIsoTime(`1970-01-01T${raw}`) || raw;
}

function normalizeFeatureList(value, allowed) {
  const source = Array.isArray(value) ? value : [];
  const seen = new Set();
  return source
    .map((item) => cleanText(item).toLowerCase())
    .filter((item) => allowed.includes(item))
    .filter((item) => {
      if (seen.has(item)) return false;
      seen.add(item);
      return true;
    });
}

function sanitizeProfileInput(input) {
  const name = cleanText(input?.name);
  if (!name) {
    const err = new Error("Profile name is required.");
    err.statusCode = 400;
    throw err;
  }
  const appFeatures = normalizeFeatureList(input?.appFeatures, APP_FEATURES);
  const settingsFeatures = normalizeFeatureList(input?.settingsFeatures, SETTINGS_FEATURES);
  return { name, appFeatures, settingsFeatures };
}

function sanitizeEntry(input) {
  const now = new Date();
  const person = cleanText(input.person);
  const message = cleanText(input.message);
  if (!person || !message) {
    const error = new Error("Both person and message are required.");
    error.statusCode = 400;
    throw error;
  }

  return {
    date: normalizeDate(input.date) || toIsoDate(now),
    time: normalizeTime(input.time) || toIsoTime(now),
    person,
    status: normalizeStatus(input.status),
    category: normalizeCategory(input.category),
    message,
  };
}

function normalizeSource(value) {
  const raw = cleanText(value).toLowerCase();
  if (raw.includes("booking")) return "booking";
  if (raw.includes("hostelworld")) return "hostelworld";
  if (raw.includes("expedia") || raw.includes("hotel")) return "expedia";
  if (raw.includes("airbnb")) return "airbnb";
  if (raw.includes("vrbo") || raw.includes("homeaway")) return "vrbo";
  if (raw.includes("trip")) return "tripadvisor";
  if (raw.includes("google")) return "google";
  return raw || "unknown";
}

function normalizeNumeric(value) {
  const raw = cleanText(value).replace(",", ".");
  if (!raw) return null;
  const num = Number(raw);
  return Number.isFinite(num) ? num : null;
}

function normalizeBool(value, fallback = false) {
  if (typeof value === "boolean") return value;
  const raw = cleanText(value).toLowerCase();
  if (["true", "1", "yes"].includes(raw)) return true;
  if (["false", "0", "no"].includes(raw)) return false;
  return fallback;
}

function normalizeJsonArray(value) {
  if (Array.isArray(value)) return value;
  if (typeof value === "string" && value.trim()) {
    try {
      const parsed = JSON.parse(value);
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return value.split(/[|,;\n]/).map((item) => cleanText(item)).filter(Boolean);
    }
  }
  return [];
}

function normalizeJsonObject(value) {
  if (value && typeof value === "object" && !Array.isArray(value)) return value;
  if (typeof value === "string" && value.trim()) {
    try {
      const parsed = JSON.parse(value);
      return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
    } catch {
      return { raw: value };
    }
  }
  return {};
}

function compactObject(value) {
  const source = normalizeJsonObject(value);
  const entries = Object.entries(source).filter(([, item]) => {
    if (item === null || item === undefined) return false;
    if (typeof item === "string") return cleanText(item) !== "";
    return true;
  });
  return Object.fromEntries(entries);
}

function normalizeRatingNormalized(raw, scale) {
  if (raw === null || raw === undefined || scale === null || scale === undefined) return null;
  const rawNum = Number(raw);
  const scaleNum = Number(scale);
  if (!Number.isFinite(rawNum) || !Number.isFinite(scaleNum) || scaleNum <= 0) return null;
  return Math.max(0, Math.min(100, Number(((rawNum / scaleNum) * 100).toFixed(2))));
}

function fingerprintReview(input) {
  const reference = cleanText(input.source_review_id || input.sourceReviewId || input.source_reservation_id || input.sourceReservationId);
  const stableParts = [
    normalizeSource(input.source),
    cleanText(input.property_id || input.propertyId),
    cleanText(input.review_date || input.reviewDate),
    cleanText(input.reviewer_name || input.reviewerName).toLowerCase(),
    cleanText(input.rating_raw || input.ratingRaw),
  ];
  if (stableParts.every(Boolean)) return stableParts.join("::");
  return [...stableParts, reference || cleanText(input.body).toLowerCase().slice(0, 400)].join("::");
}

function sanitizeReviewInput(input, defaults = {}) {
  const source = normalizeSource(input.source ?? defaults.source);
  const propertyId = cleanText(input.property_id ?? input.propertyId ?? defaults.propertyId);
  if (!propertyId) {
    const err = new Error("property_id is required.");
    err.statusCode = 400;
    throw err;
  }

  const ratingRaw = normalizeNumeric(input.rating_raw ?? input.ratingRaw);
  const ratingScale = normalizeNumeric(input.rating_scale ?? input.ratingScale);
  const parsed = {
    property_id: propertyId,
    import_run_id: cleanText(input.import_run_id ?? input.importRunId ?? defaults.importRunId) || null,
    source,
    source_review_id: cleanText(input.source_review_id ?? input.sourceReviewId),
    source_reservation_id: cleanText(input.source_reservation_id ?? input.sourceReservationId ?? input.reservationNumber),
    review_date: normalizeDate(input.review_date ?? input.reviewDate) || null,
    reviewer_name: cleanText(input.reviewer_name ?? input.reviewerName),
    reviewer_country: cleanText(input.reviewer_country ?? input.reviewerCountry),
    language: cleanText(input.language),
    rating_raw: ratingRaw,
    rating_scale: ratingScale,
    rating_normalized_100:
      normalizeNumeric(input.rating_normalized_100 ?? input.ratingNormalized100) ??
      normalizeRatingNormalized(ratingRaw, ratingScale),
    title: cleanText(input.title),
    positive_review_text: cleanText(input.positive_review_text ?? input.positiveReviewText),
    negative_review_text: cleanText(input.negative_review_text ?? input.negativeReviewText),
    body: cleanText(input.body),
    subscores: compactObject(input.subscores),
    host_reply_text: cleanText(input.host_reply_text ?? input.hostReplyText),
    host_reply_date: normalizeDate(input.host_reply_date ?? input.hostReplyDate) || null,
    raw_text: cleanText(input.raw_text ?? input.rawText),
    parse_confidence: normalizeNumeric(input.parse_confidence ?? input.parseConfidence),
    dedupe_fingerprint: cleanText(input.dedupe_fingerprint ?? input.dedupeFingerprint),
    warning_flags: normalizeJsonArray(input.warning_flags ?? input.warningFlags),
    raw_payload: normalizeJsonObject(input.raw_payload ?? input.rawPayload),
    selected_for_import: normalizeBool(input.selected_for_import ?? input.selectedForImport, true),
    is_valid: normalizeBool(input.is_valid ?? input.isValid, true),
  };

  if (!parsed.body && !parsed.title) {
    parsed.body = [parsed.positive_review_text, parsed.negative_review_text]
      .filter(Boolean)
      .map((text, index) => `${index === 0 ? "Positive" : "Negative"}: ${text}`)
      .join("\n");
  }

  if (!parsed.body && !parsed.title && parsed.rating_raw !== null) {
    parsed.title = "Rating only review";
    parsed.warning_flags = [...new Set([...parsed.warning_flags, "rating_only"])];
  }

  if (!parsed.body && !parsed.title) {
    const err = new Error("Each review needs at least a title or body.");
    err.statusCode = 400;
    throw err;
  }

  parsed.dedupe_fingerprint = parsed.dedupe_fingerprint || fingerprintReview(parsed);
  return parsed;
}

async function parseBody(req) {
  if (req.body && typeof req.body === "object") return req.body;
  if (typeof req.body === "string" && req.body.length > 0) return JSON.parse(req.body);
  return {};
}

async function verifyUser(req) {
  const { url, anonKey } = readEnv();
  const authHeader = req.headers.authorization || "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : "";

  if (!token) {
    const error = new Error("Missing bearer token.");
    error.statusCode = 401;
    throw error;
  }

  const response = await fetch(`${url}/auth/v1/user`, {
    method: "GET",
    headers: {
      apikey: anonKey,
      Authorization: `Bearer ${token}`,
    },
  });

  if (!response.ok) {
    const error = new Error("Authentication required.");
    error.statusCode = 401;
    throw error;
  }

  return response.json();
}

async function restQuery(path, { method = "GET", body, preferRepresentation = false } = {}) {
  const { url, serviceRoleKey } = readEnv();
  const headers = {
    apikey: serviceRoleKey,
    Authorization: `Bearer ${serviceRoleKey}`,
    "Content-Type": "application/json",
  };
  if (preferRepresentation) headers.Prefer = "return=representation";

  const response = await fetch(`${url}/rest/v1/${path}`, {
    method,
    headers,
    body: body === undefined ? undefined : JSON.stringify(body),
  });

  const text = await response.text();
  let payload;
  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = text || null;
  }

  if (!response.ok) {
    const message =
      (payload && (payload.message || payload.error || payload.hint || payload.details)) ||
      `Supabase REST failed (${response.status})`;
    const error = new Error(message);
    error.statusCode = response.status;
    throw error;
  }

  return payload;
}

async function authAdminQuery(path, { method = "GET", body } = {}) {
  const { url, serviceRoleKey } = readEnv();
  const headers = {
    apikey: serviceRoleKey,
    Authorization: `Bearer ${serviceRoleKey}`,
    "Content-Type": "application/json",
  };

  const response = await fetch(`${url}/auth/v1/admin/${path}`, {
    method,
    headers,
    body: body === undefined ? undefined : JSON.stringify(body),
  });

  const text = await response.text();
  let payload;
  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = text || null;
  }

  if (!response.ok) {
    const message =
      (payload && (payload.message || payload.error || payload.msg)) ||
      `Supabase Auth admin failed (${response.status})`;
    const error = new Error(message);
    error.statusCode = response.status;
    throw error;
  }

  return payload;
}

function mapProfileRow(row) {
  return {
    id: cleanText(row?.id),
    name: cleanText(row?.name),
    appFeatures: normalizeFeatureList(row?.app_features, APP_FEATURES),
    settingsFeatures: normalizeFeatureList(row?.settings_features, SETTINGS_FEATURES),
  };
}

async function loadProfiles() {
  const rows = await restQuery("app_profiles?select=id,name,app_features,settings_features&order=name.asc", {
    method: "GET",
  });
  return Array.isArray(rows) ? rows.map(mapProfileRow) : [];
}

async function loadAccessForUser(userId) {
  try {
    const profiles = await loadProfiles();
    const assignments = await restQuery(
      `user_profile_assignments?select=user_id,profile_id&user_id=eq.${encodeURIComponent(userId)}&limit=1`,
      { method: "GET" }
    );
    const assigned = Array.isArray(assignments) && assignments[0] ? cleanText(assignments[0].profile_id) : "";
    const profile = profiles.find((item) => item.id === assigned) || FALLBACK_PROFILE;
    return {
      profile,
      appFeatures: profile.appFeatures,
      settingsFeatures: profile.settingsFeatures,
    };
  } catch {
    return {
      profile: FALLBACK_PROFILE,
      appFeatures: FALLBACK_PROFILE.appFeatures,
      settingsFeatures: FALLBACK_PROFILE.settingsFeatures,
    };
  }
}

function hasFeature(access, area, feature) {
  const key = cleanText(feature).toLowerCase();
  if (!key) return false;
  if (area === "app") return (access?.appFeatures || []).includes(key);
  if (area === "settings") return (access?.settingsFeatures || []).includes(key);
  return false;
}

async function requireFeature(req, area, feature) {
  const user = await verifyUser(req);
  const access = await loadAccessForUser(user.id);
  if (!hasFeature(access, area, feature)) {
    const err = new Error("You do not have permission for this feature.");
    err.statusCode = 403;
    throw err;
  }
  return { user, access };
}

async function requireSettingsAdmin(req) {
  return requireFeature(req, "settings", "admin-users");
}

function sendError(res, error) {
  const status = Number(error.statusCode || 500);
  res.status(status).json({ error: error.message || "Unexpected server error." });
}

module.exports = {
  APP_FEATURES,
  SETTINGS_FEATURES,
  authAdminQuery,
  cleanText,
  fingerprintReview,
  hasFeature,
  loadAccessForUser,
  loadProfiles,
  mapProfileRow,
  normalizeBool,
  normalizeDate,
  normalizeFeatureList,
  normalizeNumeric,
  normalizeSource,
  parseBody,
  requireFeature,
  requireSettingsAdmin,
  restQuery,
  sanitizeReviewInput,
  sanitizeEntry,
  sanitizeProfileInput,
  sendError,
  verifyUser,
};
