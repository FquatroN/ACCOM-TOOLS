const { randomUUID } = require("crypto");

const {
  cleanText,
  parseBody,
  requireFeature,
  restQuery,
  sanitizeReviewInput,
  sendError,
} = require("./_supabase");

const SETTINGS_KEY = "reviews";
const GOOGLE_SCOPE = "https://www.googleapis.com/auth/business.manage";
const GOOGLE_AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth";
const GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token";
const GOOGLE_ACCOUNTS_URL = "https://mybusinessaccountmanagement.googleapis.com/v1/accounts";
const GOOGLE_BUSINESS_INFO_URL = "https://mybusinessbusinessinformation.googleapis.com/v1";
const GOOGLE_REVIEWS_URL = "https://mybusiness.googleapis.com/v4";
const LOCATION_LOAD_COOLDOWN_MS = 90 * 1000;

function requireGoogleEnv() {
  const clientId = cleanText(process.env.GOOGLE_CLIENT_ID);
  const clientSecret = cleanText(process.env.GOOGLE_CLIENT_SECRET);
  if (!clientId || !clientSecret) {
    const error = new Error("Missing server environment variables: GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET");
    error.statusCode = 500;
    throw error;
  }
  return { clientId, clientSecret };
}

function appBaseUrl(req) {
  const protocol = cleanText(req.headers["x-forwarded-proto"]) || (req.headers.host?.includes("localhost") ? "http" : "https");
  return `${protocol}://${req.headers.host}`;
}

function redirectUri(req) {
  return cleanText(process.env.GOOGLE_REDIRECT_URI) || `${appBaseUrl(req)}/api/google-business`;
}

async function loadSettingsPayload() {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${SETTINGS_KEY}&limit=1`, { method: "GET" });
  const row = Array.isArray(rows) && rows[0] ? rows[0] : null;
  return { id: cleanText(row?.id), payload: row?.payload && typeof row.payload === "object" ? row.payload : {} };
}

async function saveSettingsPayload(payload) {
  const existing = await loadSettingsPayload();
  if (existing.id) {
    await restQuery(`app_settings?setting_key=eq.${SETTINGS_KEY}`, {
      method: "PATCH",
      body: { payload, updated_at: new Date().toISOString() },
    });
    return;
  }
  await restQuery("app_settings", {
    method: "POST",
    body: [{ setting_key: SETTINGS_KEY, payload }],
  });
}

function safeGoogleSettings(payload) {
  const google = payload?.google && typeof payload.google === "object" ? payload.google : {};
  return {
    connected: !!cleanText(google.refreshToken),
    connectedAt: cleanText(google.connectedAt),
    locations: Array.isArray(google.locations) ? google.locations.map(safeLocation) : [],
    propertyLocations: google.propertyLocations && typeof google.propertyLocations === "object" ? google.propertyLocations : {},
    locationsLoadedAt: cleanText(google.locationsLoadedAt),
    locationsLastAttemptAt: cleanText(google.locationsLastAttemptAt),
    locationsLastError: cleanText(google.locationsLastError),
  };
}

function safeLocation(location) {
  return {
    accountName: cleanText(location?.accountName),
    locationName: cleanText(location?.locationName),
    reviewParent: cleanText(location?.reviewParent),
    title: cleanText(location?.title),
    address: cleanText(location?.address),
  };
}

async function googleJson(url, accessToken, options = {}) {
  const response = await fetch(url, {
    method: options.method || "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: options.body === undefined ? undefined : JSON.stringify(options.body),
  });
  const text = await response.text();
  let payload = {};
  try {
    payload = text ? JSON.parse(text) : {};
  } catch {
    payload = { raw: text };
  }
  if (!response.ok) {
    const error = new Error(payload?.error?.message || payload?.error || `Google API request failed (${response.status})`);
    error.statusCode = response.status;
    throw error;
  }
  return payload;
}

function isGoogleQuotaError(error) {
  const message = cleanText(error?.message).toLowerCase();
  return Number(error?.statusCode) === 429 || message.includes("quota exceeded") || message.includes("rate limit");
}

function cachedLocationsResponse(payload, message, cooldownSeconds = 0) {
  const google = safeGoogleSettings(payload);
  return {
    locations: google.locations,
    google,
    fromCache: true,
    cooldownSeconds,
    message,
  };
}

async function exchangeCodeForTokens(req, code) {
  const { clientId, clientSecret } = requireGoogleEnv();
  const response = await fetch(GOOGLE_TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      code,
      client_id: clientId,
      client_secret: clientSecret,
      redirect_uri: redirectUri(req),
      grant_type: "authorization_code",
    }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload?.error_description || payload?.error || "Google OAuth token exchange failed.");
    error.statusCode = response.status;
    throw error;
  }
  return payload;
}

async function refreshAccessToken(req, payload) {
  const google = payload?.google && typeof payload.google === "object" ? payload.google : {};
  const refreshToken = cleanText(google.refreshToken);
  if (!refreshToken) {
    const error = new Error("Google Business Profile is not connected yet.");
    error.statusCode = 400;
    throw error;
  }

  const existingToken = cleanText(google.accessToken);
  const expiresAt = Date.parse(cleanText(google.tokenExpiresAt));
  if (existingToken && Number.isFinite(expiresAt) && expiresAt > Date.now() + 60_000) {
    return existingToken;
  }

  const { clientId, clientSecret } = requireGoogleEnv();
  const response = await fetch(GOOGLE_TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      refresh_token: refreshToken,
      grant_type: "refresh_token",
    }),
  });
  const tokenPayload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(tokenPayload?.error_description || tokenPayload?.error || "Could not refresh Google access token.");
    error.statusCode = response.status;
    throw error;
  }

  const nextPayload = {
    ...payload,
    google: {
      ...google,
      accessToken: cleanText(tokenPayload.access_token),
      tokenExpiresAt: new Date(Date.now() + (Number(tokenPayload.expires_in) || 3600) * 1000).toISOString(),
    },
  };
  await saveSettingsPayload(nextPayload);
  return nextPayload.google.accessToken;
}

async function listGoogleLocations(accessToken) {
  const accountsPayload = await googleJson(`${GOOGLE_ACCOUNTS_URL}?pageSize=100`, accessToken);
  const accounts = Array.isArray(accountsPayload.accounts) ? accountsPayload.accounts : [];
  const locations = [];

  for (const account of accounts) {
    const accountName = cleanText(account.name);
    if (!accountName) continue;
    let pageToken = "";
    do {
      const params = new URLSearchParams({
        readMask: "name,title,storefrontAddress",
        pageSize: "100",
      });
      if (pageToken) params.set("pageToken", pageToken);
      const payload = await googleJson(`${GOOGLE_BUSINESS_INFO_URL}/${accountName}/locations?${params.toString()}`, accessToken);
      (Array.isArray(payload.locations) ? payload.locations : []).forEach((location) => {
        const locationName = cleanText(location.name);
        if (!locationName) return;
        locations.push(safeLocation({
          accountName,
          locationName,
          reviewParent: locationName.startsWith("accounts/") ? locationName : `${accountName}/${locationName}`,
          title: cleanText(location.title) || locationName,
          address: formatAddress(location.storefrontAddress),
        }));
      });
      pageToken = cleanText(payload.nextPageToken);
    } while (pageToken);
  }

  return locations.sort((a, b) => a.title.localeCompare(b.title));
}

function formatAddress(address) {
  if (!address || typeof address !== "object") return "";
  return [
    ...(Array.isArray(address.addressLines) ? address.addressLines : []),
    address.locality,
    address.postalCode,
    address.regionCode,
  ].map(cleanText).filter(Boolean).join(", ");
}

function googleStarRatingToNumber(value) {
  if (typeof value === "number") return value;
  const raw = cleanText(value).toUpperCase();
  const map = {
    ONE: 1,
    TWO: 2,
    THREE: 3,
    FOUR: 4,
    FIVE: 5,
  };
  return map[raw] || null;
}

function isoDate(value) {
  const date = new Date(cleanText(value));
  if (Number.isNaN(date.getTime())) return "";
  return date.toISOString().slice(0, 10);
}

function googleReviewToPayload(review, propertyId, importRunId) {
  const rating = googleStarRatingToNumber(review.starRating);
  const comment = cleanText(review.comment);
  return sanitizeReviewInput({
    propertyId,
    importRunId,
    source: "google",
    sourceReviewId: cleanText(review.name),
    sourceReservationId: "",
    reviewDate: isoDate(review.createTime || review.updateTime),
    reviewerName: cleanText(review.reviewer?.displayName) || "Google user",
    ratingRaw: rating,
    ratingScale: 5,
    title: comment ? "" : "Rating only review",
    body: comment,
    hostReplyText: cleanText(review.reviewReply?.comment),
    hostReplyDate: isoDate(review.reviewReply?.updateTime),
    rawPayload: review,
  });
}

async function findExistingReviewId(payload) {
  const sourceReviewId = cleanText(payload.source_review_id);
  if (sourceReviewId) {
    const rows = await restQuery(
      `reviews?select=id&source=eq.google&source_review_id=eq.${encodeURIComponent(sourceReviewId)}&limit=1`,
      { method: "GET" }
    );
    if (Array.isArray(rows) && rows[0]?.id) return rows[0].id;
  }

  if (payload.review_date && payload.reviewer_name && payload.rating_raw !== null && payload.rating_raw !== undefined) {
    const rows = await restQuery(
      `reviews?select=id&property_id=eq.${encodeURIComponent(payload.property_id)}&source=eq.google&review_date=eq.${encodeURIComponent(payload.review_date)}&reviewer_name=eq.${encodeURIComponent(payload.reviewer_name)}&rating_raw=eq.${encodeURIComponent(payload.rating_raw)}&limit=1`,
      { method: "GET" }
    );
    if (Array.isArray(rows) && rows[0]?.id) return rows[0].id;
  }

  const fingerprint = cleanText(payload.dedupe_fingerprint);
  if (fingerprint) {
    const rows = await restQuery(`reviews?select=id&dedupe_fingerprint=eq.${encodeURIComponent(fingerprint)}&limit=1`, { method: "GET" });
    if (Array.isArray(rows) && rows[0]?.id) return rows[0].id;
  }
  return "";
}

function isDuplicateReviewError(error) {
  const message = cleanText(error?.message).toLowerCase();
  return message.includes("duplicate key") || message.includes("unique constraint");
}

async function upsertReview(payload) {
  const existingId = await findExistingReviewId(payload);
  if (existingId) {
    await restQuery(`reviews?id=eq.${encodeURIComponent(existingId)}`, { method: "PATCH", body: payload });
    return "replaced";
  }
  try {
    await restQuery("reviews", { method: "POST", body: payload });
    return "inserted";
  } catch (error) {
    if (!isDuplicateReviewError(error)) throw error;
    const duplicateId = await findExistingReviewId(payload);
    if (!duplicateId) throw error;
    await restQuery(`reviews?id=eq.${encodeURIComponent(duplicateId)}`, { method: "PATCH", body: payload });
    return "replaced";
  }
}

async function createApiImportRun(propertyId, reviewCount, userId) {
  const id = randomUUID();
  const created = await restQuery("review_import_runs?select=*", {
    method: "POST",
    body: [{
      id,
      property_id: propertyId,
      source: "google",
      file_name: "Google Business Profile API",
      file_type: "api",
      upload_kind: "google_api",
      status: "parsed",
      row_count_detected: reviewCount,
      row_count_imported: 0,
      created_by: userId || null,
    }],
  });
  return Array.isArray(created) && created[0] ? created[0] : { id };
}

async function fetchReviewsForLocation(accessToken, reviewParent) {
  const reviews = [];
  let pageToken = "";
  do {
    const params = new URLSearchParams({ pageSize: "50", orderBy: "updateTime desc" });
    if (pageToken) params.set("pageToken", pageToken);
    const payload = await googleJson(`${GOOGLE_REVIEWS_URL}/${reviewParent}/reviews?${params.toString()}`, accessToken);
    reviews.push(...(Array.isArray(payload.reviews) ? payload.reviews : []));
    pageToken = cleanText(payload.nextPageToken);
  } while (pageToken);
  return reviews;
}

async function syncGoogleReviews(req, userId) {
  const body = await parseBody(req);
  const propertyIdFilter = cleanText(body?.propertyId);
  const { payload } = await loadSettingsPayload();
  const google = payload.google && typeof payload.google === "object" ? payload.google : {};
  const mappings = google.propertyLocations && typeof google.propertyLocations === "object" ? google.propertyLocations : {};
  const accessToken = await refreshAccessToken(req, payload);
  const properties = Object.entries(mappings)
    .map(([propertyId, reviewParent]) => ({ propertyId: cleanText(propertyId), reviewParent: cleanText(reviewParent) }))
    .filter((item) => item.propertyId && item.reviewParent && (!propertyIdFilter || item.propertyId === propertyIdFilter));

  if (properties.length === 0) {
    const error = new Error("No Google location mapping is configured for sync.");
    error.statusCode = 400;
    throw error;
  }

  let insertedCount = 0;
  let replacedCount = 0;
  let importedCount = 0;
  const syncedProperties = [];

  for (const item of properties) {
    const googleReviews = await fetchReviewsForLocation(accessToken, item.reviewParent);
    const run = await createApiImportRun(item.propertyId, googleReviews.length, userId);
    let propertyImported = 0;
    let propertyInserted = 0;
    let propertyReplaced = 0;

    for (const review of googleReviews) {
      const payloadReview = googleReviewToPayload(review, item.propertyId, run.id);
      const action = await upsertReview(payloadReview);
      propertyImported += 1;
      if (action === "inserted") propertyInserted += 1;
      if (action === "replaced") propertyReplaced += 1;
    }

    await restQuery(`review_import_runs?id=eq.${encodeURIComponent(run.id)}`, {
      method: "PATCH",
      body: {
        status: "imported",
        row_count_imported: propertyImported,
        error_summary: propertyReplaced ? `${propertyReplaced} duplicate review${propertyReplaced === 1 ? " was" : "s were"} replaced.` : "",
      },
    });

    insertedCount += propertyInserted;
    replacedCount += propertyReplaced;
    importedCount += propertyImported;
    syncedProperties.push({
      propertyId: item.propertyId,
      reviewParent: item.reviewParent,
      importedCount: propertyImported,
      insertedCount: propertyInserted,
      replacedCount: propertyReplaced,
    });
  }

  return { importedCount, insertedCount, replacedCount, syncedProperties };
}

module.exports = async function handler(req, res) {
  try {
    if (req.method === "GET" && req.query?.code && req.query?.state) {
      const code = cleanText(req.query.code);
      const state = cleanText(req.query.state);
      const { payload } = await loadSettingsPayload();
      const google = payload.google && typeof payload.google === "object" ? payload.google : {};
      if (!state || state !== cleanText(google.oauthState)) {
        res.writeHead(302, { Location: "/index.html?google=failed" });
        res.end();
        return;
      }
      const tokenPayload = await exchangeCodeForTokens(req, code);
      const refreshToken = cleanText(tokenPayload.refresh_token || google.refreshToken);
      if (!refreshToken) {
        const error = new Error("Google did not return a refresh token. Please reconnect and approve offline access.");
        error.statusCode = 400;
        throw error;
      }
      await saveSettingsPayload({
        ...payload,
        google: {
          ...google,
          oauthState: "",
          refreshToken,
          accessToken: cleanText(tokenPayload.access_token),
          tokenExpiresAt: new Date(Date.now() + (Number(tokenPayload.expires_in) || 3600) * 1000).toISOString(),
          connectedAt: new Date().toISOString(),
        },
      });
      res.writeHead(302, { Location: "/index.html?google=connected" });
      res.end();
      return;
    }

    const auth = await requireFeature(req, "settings", "reviews");
    const action = cleanText(req.query?.action || "status").toLowerCase();

    if (req.method === "GET" && action === "status") {
      const { payload } = await loadSettingsPayload();
      res.status(200).json({ google: safeGoogleSettings(payload) });
      return;
    }

    if (req.method === "POST" && action === "auth-url") {
      const { clientId } = requireGoogleEnv();
      const { payload } = await loadSettingsPayload();
      const oauthState = randomUUID();
      await saveSettingsPayload({
        ...payload,
        google: {
          ...(payload.google && typeof payload.google === "object" ? payload.google : {}),
          oauthState,
        },
      });
      const params = new URLSearchParams({
        client_id: clientId,
        redirect_uri: redirectUri(req),
        response_type: "code",
        scope: GOOGLE_SCOPE,
        access_type: "offline",
        prompt: "consent",
        state: oauthState,
      });
      res.status(200).json({ authUrl: `${GOOGLE_AUTH_URL}?${params.toString()}` });
      return;
    }

    if (req.method === "POST" && action === "locations") {
      const { payload } = await loadSettingsPayload();
      const google = payload.google && typeof payload.google === "object" ? payload.google : {};
      const lastAttemptAt = Date.parse(cleanText(google.locationsLastAttemptAt));
      const elapsed = Number.isFinite(lastAttemptAt) ? Date.now() - lastAttemptAt : LOCATION_LOAD_COOLDOWN_MS;
      if (elapsed < LOCATION_LOAD_COOLDOWN_MS) {
        const cooldownSeconds = Math.ceil((LOCATION_LOAD_COOLDOWN_MS - elapsed) / 1000);
        res.status(200).json(cachedLocationsResponse(
          payload,
          `Google location loading is cooling down to avoid quota limits. Try again in ${cooldownSeconds} seconds.`,
          cooldownSeconds
        ));
        return;
      }

      await saveSettingsPayload({
        ...payload,
        google: {
          ...google,
          locationsLastAttemptAt: new Date().toISOString(),
        },
      });

      const accessToken = await refreshAccessToken(req, payload);
      let locations = [];
      try {
        locations = await listGoogleLocations(accessToken);
      } catch (error) {
        const latestAfterError = (await loadSettingsPayload()).payload;
        const cachedLocations = Array.isArray(latestAfterError.google?.locations) ? latestAfterError.google.locations : [];
        await saveSettingsPayload({
          ...latestAfterError,
          google: {
            ...(latestAfterError.google && typeof latestAfterError.google === "object" ? latestAfterError.google : {}),
            locationsLastError: cleanText(error.message),
          },
        });
        if (isGoogleQuotaError(error) && cachedLocations.length > 0) {
          res.status(200).json(cachedLocationsResponse(
            { ...latestAfterError, google: { ...latestAfterError.google, locationsLastError: cleanText(error.message) } },
            "Google quota is currently exceeded, so cached locations are being used. Wait a few minutes before refreshing locations."
          ));
          return;
        }
        if (isGoogleQuotaError(error)) {
          const friendly = new Error("Google quota is currently exceeded while loading locations. Wait a few minutes, then press Load Locations once. If this continues, request a quota increase in Google Cloud for the My Business Account Management API.");
          friendly.statusCode = 429;
          throw friendly;
        }
        throw error;
      }
      const latest = (await loadSettingsPayload()).payload;
      await saveSettingsPayload({
        ...latest,
        google: {
          ...(latest.google && typeof latest.google === "object" ? latest.google : {}),
          locations,
          locationsLoadedAt: new Date().toISOString(),
          locationsLastError: "",
        },
      });
      res.status(200).json({ locations, google: safeGoogleSettings({ ...latest, google: { ...latest.google, locations } }) });
      return;
    }

    if (req.method === "POST" && action === "mapping") {
      const body = await parseBody(req);
      const propertyLocations = body?.propertyLocations && typeof body.propertyLocations === "object" ? body.propertyLocations : {};
      const cleanMapping = Object.fromEntries(
        Object.entries(propertyLocations)
          .map(([propertyId, reviewParent]) => [cleanText(propertyId), cleanText(reviewParent)])
          .filter(([propertyId, reviewParent]) => propertyId && reviewParent)
      );
      const { payload } = await loadSettingsPayload();
      await saveSettingsPayload({
        ...payload,
        google: {
          ...(payload.google && typeof payload.google === "object" ? payload.google : {}),
          propertyLocations: cleanMapping,
        },
      });
      res.status(200).json({ propertyLocations: cleanMapping });
      return;
    }

    if (req.method === "POST" && action === "sync") {
      const result = await syncGoogleReviews(req, auth.user?.id);
      res.status(200).json(result);
      return;
    }

    res.status(405).json({ error: "Method/action not allowed." });
  } catch (error) {
    if (req.query?.code && req.query?.state && !res.headersSent) {
      res.writeHead(302, { Location: `/index.html?google=failed&message=${encodeURIComponent(error.message || "Google connection failed")}` });
      res.end();
      return;
    }
    sendError(res, error);
  }
};
