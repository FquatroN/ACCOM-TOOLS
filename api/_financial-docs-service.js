const { randomUUID } = require("crypto");

const { cleanText, restQuery } = require("./_supabase");
const {
  FINANCIAL_DOCS_SETTINGS_KEY,
  sanitizeFinancialDocsSettings,
  safeFinancialDocsSettings,
  toClientDocument,
  toClientHistory,
  monthlyFolderKey,
  buildStoredFileName,
} = require("./_financial-docs");

const GOOGLE_AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth";
const GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token";
const GOOGLE_DRIVE_API = "https://www.googleapis.com/drive/v3";
const GOOGLE_DRIVE_UPLOAD_API = "https://www.googleapis.com/upload/drive/v3/files";
const GOOGLE_USERINFO_URL = "https://www.googleapis.com/oauth2/v2/userinfo";
const GOOGLE_DRIVE_SCOPE = "https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/userinfo.email";
const FOLDER_MIME = "application/vnd.google-apps.folder";

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
  return cleanText(process.env.GOOGLE_DRIVE_REDIRECT_URI) || `${appBaseUrl(req)}/api/financial-docs-drive`;
}

async function loadFinancialDocsSettingsRecord() {
  const rows = await restQuery(`app_settings?select=id,payload&setting_key=eq.${FINANCIAL_DOCS_SETTINGS_KEY}&limit=1`, { method: "GET" });
  const row = Array.isArray(rows) && rows[0] ? rows[0] : null;
  return { id: cleanText(row?.id), payload: row?.payload && typeof row.payload === "object" ? row.payload : {} };
}

async function loadFinancialDocsSettings() {
  const record = await loadFinancialDocsSettingsRecord();
  return sanitizeFinancialDocsSettings(record.payload);
}

async function saveFinancialDocsSettings(nextSettings) {
  const existing = await loadFinancialDocsSettingsRecord();
  const payload = sanitizeFinancialDocsSettings(nextSettings);
  if (existing.id) {
    await restQuery(`app_settings?setting_key=eq.${FINANCIAL_DOCS_SETTINGS_KEY}`, {
      method: "PATCH",
      body: { payload, updated_at: new Date().toISOString() },
    });
    return payload;
  }
  await restQuery("app_settings", {
    method: "POST",
    body: [{ setting_key: FINANCIAL_DOCS_SETTINGS_KEY, payload }],
  });
  return payload;
}

async function updateFinancialDocsSettings(mutator) {
  const current = await loadFinancialDocsSettings();
  const next = await mutator(current);
  return saveFinancialDocsSettings(next);
}

async function googleTokenRequest(params) {
  const response = await fetch(GOOGLE_TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(params),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload?.error_description || payload?.error || "Google token request failed.");
    error.statusCode = response.status;
    throw error;
  }
  return payload;
}

async function exchangeCodeForDriveTokens(req, code) {
  const { clientId, clientSecret } = requireGoogleEnv();
  return googleTokenRequest({
    code,
    client_id: clientId,
    client_secret: clientSecret,
    redirect_uri: redirectUri(req),
    grant_type: "authorization_code",
  });
}

async function refreshDriveAccessToken(settings) {
  const refreshToken = cleanText(settings?.drive?.refreshToken);
  if (!refreshToken) {
    const error = new Error("Google Drive is not connected yet.");
    error.statusCode = 400;
    throw error;
  }
  const tokenExpiresAt = Date.parse(cleanText(settings?.drive?.tokenExpiresAt));
  const currentAccessToken = cleanText(settings?.drive?.accessToken);
  if (currentAccessToken && Number.isFinite(tokenExpiresAt) && tokenExpiresAt > Date.now() + 60_000) {
    return { accessToken: currentAccessToken, settings };
  }
  const { clientId, clientSecret } = requireGoogleEnv();
  const tokenPayload = await googleTokenRequest({
    client_id: clientId,
    client_secret: clientSecret,
    refresh_token: refreshToken,
    grant_type: "refresh_token",
  });
  const nextSettings = await updateFinancialDocsSettings((current) => ({
    ...current,
    drive: {
      ...current.drive,
      accessToken: cleanText(tokenPayload.access_token),
      tokenExpiresAt: new Date(Date.now() + (Number(tokenPayload.expires_in) || 3600) * 1000).toISOString(),
    },
  }));
  return { accessToken: cleanText(nextSettings.drive.accessToken), settings: nextSettings };
}

async function googleJson(url, accessToken, options = {}) {
  const response = await fetch(url, {
    method: options.method || "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options.body === undefined ? {} : { "Content-Type": "application/json" }),
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

async function loadDriveAccountEmail(accessToken) {
  const payload = await googleJson(GOOGLE_USERINFO_URL, accessToken);
  return cleanText(payload.email);
}

function driveFolderQuery(name, parentId) {
  const escapedName = cleanText(name).replace(/'/g, "\\'");
  return `mimeType='${FOLDER_MIME}' and trashed=false and name='${escapedName}' and '${cleanText(parentId || "root")}' in parents`;
}

async function findDriveFolder(accessToken, name, parentId) {
  const params = new URLSearchParams({
    q: driveFolderQuery(name, parentId),
    fields: "files(id,name)",
    pageSize: "10",
    supportsAllDrives: "true",
    includeItemsFromAllDrives: "true",
  });
  const payload = await googleJson(`${GOOGLE_DRIVE_API}/files?${params.toString()}`, accessToken);
  const match = Array.isArray(payload.files) && payload.files[0] ? payload.files[0] : null;
  return cleanText(match?.id);
}

async function createDriveFolder(accessToken, name, parentId) {
  const payload = await googleJson(`${GOOGLE_DRIVE_API}/files?supportsAllDrives=true`, accessToken, {
    method: "POST",
    body: {
      name,
      mimeType: FOLDER_MIME,
      parents: parentId ? [parentId] : undefined,
    },
  });
  return cleanText(payload.id);
}

async function ensureDriveFolder(accessToken, name, parentId) {
  const existing = await findDriveFolder(accessToken, name, parentId);
  if (existing) return existing;
  return createDriveFolder(accessToken, name, parentId);
}

async function ensureDriveFolderPath(accessToken, folderPath) {
  const segments = cleanText(folderPath)
    .split(/[\\/]/)
    .map((segment) => cleanText(segment))
    .filter(Boolean);
  let parentId = "root";
  for (const segment of segments) {
    parentId = await ensureDriveFolder(accessToken, segment, parentId);
  }
  return parentId;
}

function joinMultipartRelated(boundary, metadata, buffer, mimeType) {
  const prefix = Buffer.from(
    `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`,
    "utf8"
  );
  const suffix = Buffer.from(`\r\n--${boundary}--`, "utf8");
  return Buffer.concat([prefix, buffer, suffix]);
}

async function uploadDriveFile(accessToken, { folderId, fileName, mimeType, base64Content }) {
  const boundary = `accomtools-${randomUUID()}`;
  const buffer = Buffer.from(String(base64Content || ""), "base64");
  const body = joinMultipartRelated(boundary, { name: fileName, parents: folderId ? [folderId] : undefined }, buffer, mimeType || "application/pdf");
  const response = await fetch(`${GOOGLE_DRIVE_UPLOAD_API}?uploadType=multipart&supportsAllDrives=true`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": `multipart/related; boundary=${boundary}`,
    },
    body,
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload?.error?.message || payload?.error || `Google Drive upload failed (${response.status})`);
    error.statusCode = response.status;
    throw error;
  }
  return payload;
}

async function renameDriveFile(accessToken, fileId, fileName) {
  const payload = await googleJson(`${GOOGLE_DRIVE_API}/files/${encodeURIComponent(fileId)}?supportsAllDrives=true`, accessToken, {
    method: "PATCH",
    body: { name: fileName },
  });
  return payload;
}

async function deleteDriveFile(accessToken, fileId) {
  if (!cleanText(fileId)) return;
  const response = await fetch(`${GOOGLE_DRIVE_API}/files/${encodeURIComponent(fileId)}?supportsAllDrives=true`, {
    method: "DELETE",
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!response.ok && response.status !== 404) {
    const text = await response.text().catch(() => "");
    const error = new Error(text || `Google Drive delete failed (${response.status})`);
    error.statusCode = response.status;
    throw error;
  }
}

async function downloadDriveFile(accessToken, fileId) {
  const response = await fetch(`${GOOGLE_DRIVE_API}/files/${encodeURIComponent(fileId)}?alt=media&supportsAllDrives=true`, {
    method: "GET",
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!response.ok) {
    const text = await response.text().catch(() => "");
    const error = new Error(text || `Google Drive download failed (${response.status})`);
    error.statusCode = response.status;
    throw error;
  }
  const arrayBuffer = await response.arrayBuffer();
  return {
    buffer: Buffer.from(arrayBuffer),
    mimeType: cleanText(response.headers.get("content-type")),
  };
}

async function prepareDriveDestination(req, createdAt, settings) {
  const { accessToken } = await refreshDriveAccessToken(settings);
  const baseFolderPath = cleanText(settings?.drive?.folderPath) || "Financial Documents";
  const baseFolderId = await ensureDriveFolderPath(accessToken, baseFolderPath);
  const monthlyKey = monthlyFolderKey(createdAt);
  const monthlyFolderId = monthlyKey ? await ensureDriveFolder(accessToken, monthlyKey, baseFolderId) : baseFolderId;
  return { accessToken, baseFolderId, folderId: monthlyFolderId };
}

async function attachDocumentFile(req, documentRecord, upload, settings) {
  if (!upload?.base64Content) return null;
  const destination = await prepareDriveDestination(req, documentRecord.created_at || documentRecord.createdAt, settings);
  const fileName = buildStoredFileName(documentRecord, upload.originalFilename || upload.fileName || "document.pdf");
  const uploaded = await uploadDriveFile(destination.accessToken, {
    folderId: destination.folderId,
    fileName,
    mimeType: cleanText(upload.mimeType) || "application/pdf",
    base64Content: upload.base64Content,
  });
  const fileId = cleanText(uploaded.id);
  return {
    drive_file_id: fileId,
    drive_folder_id: destination.folderId,
    drive_file_url: fileId ? `https://drive.google.com/file/d/${fileId}/view` : "",
    original_filename: cleanText(upload.originalFilename || upload.fileName),
    stored_filename: fileName,
    mime_type: cleanText(upload.mimeType) || "application/pdf",
    file_size: Number(upload.fileSize || 0),
    file_hash: cleanText(upload.fileHash),
    uploaded_by: cleanText(upload.uploadedBy),
    uploaded_at: new Date().toISOString(),
  };
}

async function renameExistingDriveFileIfNeeded(req, documentRecord, settings) {
  const driveFileId = cleanText(documentRecord.drive_file_id || documentRecord.driveFileId);
  if (!driveFileId) return null;
  const { accessToken } = await refreshDriveAccessToken(settings);
  const nextName = buildStoredFileName(documentRecord, documentRecord.original_filename || documentRecord.originalFilename || documentRecord.stored_filename || "document.pdf");
  if (nextName === cleanText(documentRecord.stored_filename || documentRecord.storedFilename)) return null;
  await renameDriveFile(accessToken, driveFileId, nextName);
  return nextName;
}

async function listFinancialDocuments() {
  const rows = await restQuery("financial_documents?select=*&order=created_at.desc", { method: "GET" });
  return Array.isArray(rows) ? rows.map(toClientDocument) : [];
}

async function loadFinancialDocumentRowById(id) {
  const rows = await restQuery(`financial_documents?select=*&id=eq.${encodeURIComponent(id)}&limit=1`, { method: "GET" });
  return Array.isArray(rows) && rows[0] ? rows[0] : null;
}

async function loadFinancialDocumentWithHistory(id) {
  const row = await loadFinancialDocumentRowById(id);
  if (!row) return null;
  const historyRows = await restQuery(`financial_document_history?select=*&document_id=eq.${encodeURIComponent(id)}&order=created_at.desc`, { method: "GET" });
  const mapped = toClientDocument(row);
  mapped.history = Array.isArray(historyRows) ? historyRows.map(toClientHistory) : [];
  return mapped;
}

async function insertFinancialDocument(body) {
  const payload = Array.isArray(body) ? body : [body];
  const rows = await restQuery("financial_documents", { method: "POST", body: payload, preferRepresentation: true });
  return Array.isArray(rows) && rows[0] ? rows[0] : null;
}

async function updateFinancialDocumentRow(id, body) {
  const rows = await restQuery(`financial_documents?id=eq.${encodeURIComponent(id)}`, {
    method: "PATCH",
    body,
    preferRepresentation: true,
  });
  return Array.isArray(rows) && rows[0] ? rows[0] : null;
}

async function insertFinancialDocumentHistory(entries) {
  const payload = (Array.isArray(entries) ? entries : [entries]).filter(Boolean);
  if (!payload.length) return;
  await restQuery("financial_document_history", { method: "POST", body: payload });
}

module.exports = {
  GOOGLE_AUTH_URL,
  GOOGLE_DRIVE_SCOPE,
  appBaseUrl,
  attachDocumentFile,
  deleteDriveFile,
  downloadDriveFile,
  exchangeCodeForDriveTokens,
  loadDriveAccountEmail,
  loadFinancialDocumentRowById,
  loadFinancialDocumentWithHistory,
  loadFinancialDocsSettings,
  loadFinancialDocsSettingsRecord,
  listFinancialDocuments,
  redirectUri,
  refreshDriveAccessToken,
  renameExistingDriveFileIfNeeded,
  saveFinancialDocsSettings,
  safeFinancialDocsSettings,
  insertFinancialDocument,
  insertFinancialDocumentHistory,
  updateFinancialDocumentRow,
  updateFinancialDocsSettings,
};
