const { randomUUID } = require("crypto");

const { cleanText, requireFeature, sendError } = require("./_supabase");
const {
  GOOGLE_AUTH_URL,
  GOOGLE_DRIVE_SCOPE,
  exchangeCodeForDriveTokens,
  loadDriveAccountEmail,
  loadFinancialDocsSettings,
  redirectUri,
  refreshDriveAccessToken,
  saveFinancialDocsSettings,
} = require("./_financial-docs-service");

module.exports = async function handler(req, res) {
  try {
    if (req.method === "GET" && req.query?.code && req.query?.state) {
      const code = cleanText(req.query.code);
      const state = cleanText(req.query.state);
      const settings = await loadFinancialDocsSettings();
      if (!state || state !== cleanText(settings.drive.oauthState)) {
        res.writeHead(302, { Location: "/index.html?fd_drive=failed" });
        res.end();
        return;
      }
      const tokenPayload = await exchangeCodeForDriveTokens(req, code);
      const refreshToken = cleanText(tokenPayload.refresh_token || settings.drive.refreshToken);
      if (!refreshToken) {
        const error = new Error("Google did not return a refresh token. Please reconnect and approve offline access.");
        error.statusCode = 400;
        throw error;
      }
      const tempSettings = await saveFinancialDocsSettings({
        ...settings,
        drive: {
          ...settings.drive,
          oauthState: "",
          refreshToken,
          accessToken: cleanText(tokenPayload.access_token),
          tokenExpiresAt: new Date(Date.now() + (Number(tokenPayload.expires_in) || 3600) * 1000).toISOString(),
          connectedAt: new Date().toISOString(),
          connected: true,
        },
      });
      const email = await loadDriveAccountEmail(cleanText(tempSettings.drive.accessToken)).catch(() => "");
      await saveFinancialDocsSettings({
        ...tempSettings,
        drive: {
          ...tempSettings.drive,
          accountEmail: email,
          connected: true,
        },
      });
      res.writeHead(302, { Location: "/index.html?fd_drive=connected" });
      res.end();
      return;
    }

    await requireFeature(req, "settings", "financial-docs");
    const action = cleanText(req.query?.action || "status").toLowerCase();

    if (req.method === "GET" && action === "status") {
      const settings = await loadFinancialDocsSettings();
      res.status(200).json({ drive: { ...settings.drive, refreshToken: "", accessToken: "", tokenExpiresAt: "", oauthState: "", redirectUri: redirectUri(req) } });
      return;
    }

    if (req.method === "POST" && action === "auth-url") {
      const settings = await loadFinancialDocsSettings();
      const oauthState = randomUUID();
      const saved = await saveFinancialDocsSettings({
        ...settings,
        drive: {
          ...settings.drive,
          oauthState,
        },
      });
      const clientId = cleanText(process.env.GOOGLE_CLIENT_ID);
      if (!clientId) {
        const error = new Error("Missing server environment variable: GOOGLE_CLIENT_ID");
        error.statusCode = 500;
        throw error;
      }
      const params = new URLSearchParams({
        client_id: clientId,
        redirect_uri: redirectUri(req),
        response_type: "code",
        scope: GOOGLE_DRIVE_SCOPE,
        access_type: "offline",
        prompt: "consent",
        state: oauthState,
      });
      res.status(200).json({
        authUrl: `${GOOGLE_AUTH_URL}?${params.toString()}`,
        redirectUri: redirectUri(req),
        drive: saved.drive,
      });
      return;
    }

    if (req.method === "POST" && action === "disconnect") {
      const settings = await loadFinancialDocsSettings();
      const saved = await saveFinancialDocsSettings({
        ...settings,
        drive: {
          ...settings.drive,
          connected: false,
          connectedAt: "",
          accountEmail: "",
          accessToken: "",
          refreshToken: "",
          tokenExpiresAt: "",
          oauthState: "",
          baseFolderId: "",
        },
      });
      res.status(200).json({ drive: { ...saved.drive, redirectUri: redirectUri(req) } });
      return;
    }

    if (req.method === "POST" && action === "refresh") {
      const settings = await loadFinancialDocsSettings();
      const refreshed = await refreshDriveAccessToken(settings);
      const accountEmail = settings.drive.accountEmail || await loadDriveAccountEmail(refreshed.accessToken).catch(() => "");
      const saved = await saveFinancialDocsSettings({
        ...refreshed.settings,
        drive: {
          ...refreshed.settings.drive,
          accountEmail,
          connected: true,
        },
      });
      res.status(200).json({ drive: { ...saved.drive, refreshToken: "", accessToken: "", tokenExpiresAt: "", oauthState: "", redirectUri: redirectUri(req) } });
      return;
    }

    res.status(405).json({ error: "Method/action not allowed." });
  } catch (error) {
    if (req.query?.code && req.query?.state && !res.headersSent) {
      res.writeHead(302, { Location: `/index.html?fd_drive=failed&message=${encodeURIComponent(error.message || "Google Drive connection failed")}` });
      res.end();
      return;
    }
    sendError(res, error);
  }
};
