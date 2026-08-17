const { randomUUID } = require("crypto");

const { cleanText, restQuery } = require("./_supabase");
const {
  FINANCIAL_DOCS_SETTINGS_KEY,
  sanitizeFinancialDocsSettings,
  safeFinancialDocsSettings,
  toClientDocument,
  toClientEntity,
  latestDuplicateWarningMessage,
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
const FINANCIAL_DOCUMENT_DUPLICATE_SELECT = [
  "id",
  "document_date",
  "doc_number",
  "supplier_nif",
  "supplier_name",
  "amount",
  "document_type",
  "status",
  "file_hash",
].join(",");

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
  const sharedGoogleRedirect = cleanText(process.env.GOOGLE_REDIRECT_URI);
  if (sharedGoogleRedirect) return sharedGoogleRedirect;
  const driveRedirect = cleanText(process.env.GOOGLE_DRIVE_REDIRECT_URI);
  if (driveRedirect) return driveRedirect;
  return `${appBaseUrl(req)}/api/google-business`;
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
    const detailParts = [];
    if (payload?.error && typeof payload.error === "object") {
      [payload.error.message, payload.error.status, payload.error.code]
        .map(cleanText)
        .filter(Boolean)
        .forEach((item) => {
          if (!detailParts.includes(item)) detailParts.push(item);
        });
      if (Array.isArray(payload.error.errors)) {
        payload.error.errors.forEach((item) => {
          [item.message, item.reason, item.location]
            .map(cleanText)
            .filter(Boolean)
            .forEach((part) => {
              if (!detailParts.includes(part)) detailParts.push(part);
            });
        });
      }
    } else if (cleanText(payload?.error)) {
      detailParts.push(cleanText(payload.error));
    }
    const error = new Error(detailParts.join(" ") || `Google API request failed (${response.status})`);
    error.statusCode = response.status;
    error.googlePayload = payload;
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
  const configuredFolderId = cleanText(settings?.drive?.baseFolderId);
  const baseFolderId = configuredFolderId || await ensureDriveFolderPath(accessToken, baseFolderPath);
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

function nextIsoDate(value) {
  const raw = cleanText(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return "";
  const date = new Date(`${raw}T00:00:00Z`);
  if (Number.isNaN(date.getTime())) return "";
  date.setUTCDate(date.getUTCDate() + 1);
  return date.toISOString().slice(0, 10);
}

function buildFinancialDocumentsListPath(filters = {}) {
  const listSelect = [
    "id",
    "created_at",
    "updated_at",
    "cc",
    "document_date",
    "doc_number",
    "description",
    "supplier_nif",
    "supplier_name",
    "amount",
    "vat_amount",
    "payment",
    "document_type",
    "fat",
    "category",
    "status",
    "drive_file_id",
    "drive_folder_id",
    "drive_file_url",
    "original_filename",
    "stored_filename",
    "mime_type",
    "file_size",
    "file_hash",
    "uploaded_by",
    "uploaded_at",
  ].join(",");
  const allowedSortColumns = new Set(["created_at", "document_date", "supplier_name", "supplier_nif", "amount"]);
  const requestedSortColumn = cleanText(filters.sortBy).toLowerCase();
  const sortColumn = allowedSortColumns.has(requestedSortColumn) ? requestedSortColumn : "created_at";
  const sortDirection = cleanText(filters.sortDirection).toLowerCase() === "asc" ? "asc" : "desc";
  const order = sortColumn === "created_at"
    ? `created_at.${sortDirection},id.desc`
    : `${sortColumn}.${sortDirection}.nullslast,created_at.desc,id.desc`;
  const params = [`select=${listSelect}`, `order=${order}`];
  const createdFrom = cleanText(filters.createdFrom);
  const createdTo = cleanText(filters.createdTo);
  const dateFrom = cleanText(filters.dateFrom);
  const dateTo = cleanText(filters.dateTo);
  const supplierSearch = cleanText(filters.supplierSearch);
  const descriptionSearch = cleanText(filters.descriptionSearch);
  const payment = cleanText(filters.payment);
  const docType = cleanText(filters.docType);
  const fat = cleanText(filters.fat);
  const category = cleanText(filters.category);
  const status = cleanText(filters.status);
  if (createdFrom) params.push(`created_at=gte.${encodeURIComponent(`${createdFrom}T00:00:00.000Z`)}`);
  if (createdTo) {
    const nextDate = nextIsoDate(createdTo);
    if (nextDate) params.push(`created_at=lt.${encodeURIComponent(`${nextDate}T00:00:00.000Z`)}`);
  }
  if (dateFrom) params.push(`document_date=gte.${encodeURIComponent(dateFrom)}`);
  if (dateTo) params.push(`document_date=lte.${encodeURIComponent(dateTo)}`);
  if (supplierSearch) params.push(`or=${encodeURIComponent(`(supplier_nif.ilike.*${supplierSearch}*,supplier_name.ilike.*${supplierSearch}*)`)}`);
  if (descriptionSearch) params.push(`description=ilike.${encodeURIComponent(`*${descriptionSearch}*`)}`);
  if (payment) params.push(`payment=eq.${encodeURIComponent(payment)}`);
  if (docType) params.push(`document_type=eq.${encodeURIComponent(docType)}`);
  if (fat) params.push(`fat=eq.${encodeURIComponent(fat)}`);
  if (category) params.push(`category=eq.${encodeURIComponent(category)}`);
  if (status) params.push(`status=eq.${encodeURIComponent(status)}`);
  return `financial_documents?${params.join("&")}`;
}

async function listFinancialDocuments(filters = {}) {
  const rows = await restQuery(buildFinancialDocumentsListPath(filters), { method: "GET" });
  if (!Array.isArray(rows) || !rows.length) return [];
  const ids = rows.map((row) => cleanText(row.id)).filter(Boolean);
  const sourceIds = ids.map((id) => encodeURIComponent(id)).join(",");
  const [warningRows, reconciliationItems] = ids.length
    ? await Promise.all([
      restQuery(`financial_document_history?select=document_id,action_type,message,created_at&action_type=in.(duplicate_warning,duplicate_warning_resolved)&document_id=in.(${sourceIds})&order=created_at.desc`, { method: "GET" }),
      restQuery(`financial_reconciliation_items?select=source_id,reconciliation_id&source_type=eq.financial_documents&source_id=in.(${sourceIds})`, { method: "GET" }),
    ])
    : [[], []];
  const reconciliationIds = [...new Set((Array.isArray(reconciliationItems) ? reconciliationItems : [])
    .map((item) => cleanText(item.reconciliation_id || item.reconciliationId))
    .filter(Boolean))];
  const reconciliationRows = reconciliationIds.length
    ? await restQuery(`financial_reconciliations?select=id,status&deleted_at=is.null&id=in.(${reconciliationIds.map((id) => encodeURIComponent(id)).join(",")})`, { method: "GET" })
    : [];
  const historyByDocumentId = new Map();
  (Array.isArray(warningRows) ? warningRows : []).forEach((item) => {
    const documentId = cleanText(item.document_id || item.documentId);
    if (!documentId) return;
    if (!historyByDocumentId.has(documentId)) historyByDocumentId.set(documentId, []);
    historyByDocumentId.get(documentId).push(item);
  });
  const reconciliationById = new Map((Array.isArray(reconciliationRows) ? reconciliationRows : []).map((item) => [cleanText(item.id), item]));
  const reconciliationByDocumentId = new Map();
  (Array.isArray(reconciliationItems) ? reconciliationItems : []).forEach((item) => {
    const reconciliation = reconciliationById.get(cleanText(item.reconciliation_id || item.reconciliationId));
    const documentId = cleanText(item.source_id || item.sourceId);
    if (documentId && reconciliation) reconciliationByDocumentId.set(documentId, reconciliation);
  });
  return rows.map((row) => toClientDocument({
    ...row,
    reconciliation_id: reconciliationByDocumentId.get(cleanText(row.id))?.id || "",
    reconciliation_status: reconciliationByDocumentId.get(cleanText(row.id))?.status || "",
    duplicate_warning_message: latestDuplicateWarningMessage(historyByDocumentId.get(cleanText(row.id)) || []),
  }));
}

async function listPotentialDuplicateFinancialDocuments(candidate = {}) {
  const supplierNif = cleanText(candidate.supplierNif || candidate.supplier_nif);
  const docNumber = cleanText(candidate.docNumber || candidate.doc_number);
  const documentType = cleanText(candidate.docType || candidate.document_type);
  const documentDate = cleanText(candidate.documentDate || candidate.document_date);
  const amount = Number(candidate.amount);
  const fileHash = cleanText(candidate.fileHash || candidate.file_hash);
  const paths = [];
  const addPath = (params) => {
    const safe = params.filter(Boolean);
    if (!safe.length) return;
    paths.push(`financial_documents?select=${FINANCIAL_DOCUMENT_DUPLICATE_SELECT}&${safe.join("&")}&limit=200`);
  };
  const amountFilter = Number.isFinite(amount) ? `amount=eq.${encodeURIComponent(amount.toFixed(2))}` : "";

  if (supplierNif && docNumber && documentType) {
    addPath([
      `supplier_nif=eq.${encodeURIComponent(supplierNif)}`,
      `doc_number=eq.${encodeURIComponent(docNumber)}`,
      `document_type=eq.${encodeURIComponent(documentType)}`,
    ]);
  }
  if (supplierNif && documentDate && amountFilter) {
    addPath([
      `supplier_nif=eq.${encodeURIComponent(supplierNif)}`,
      `document_date=eq.${encodeURIComponent(documentDate)}`,
      amountFilter,
    ]);
  }
  if (fileHash) {
    addPath([`file_hash=eq.${encodeURIComponent(fileHash)}`]);
  }
  if (docNumber && documentDate && amountFilter) {
    addPath([
      `doc_number=eq.${encodeURIComponent(docNumber)}`,
      `document_date=eq.${encodeURIComponent(documentDate)}`,
      amountFilter,
    ]);
  }
  if (documentDate && amountFilter) {
    addPath([
      `document_date=eq.${encodeURIComponent(documentDate)}`,
      amountFilter,
    ]);
  }

  if (!paths.length) return [];
  const seen = new Map();
  const batches = await Promise.all(paths.map((path) => restQuery(path, { method: "GET" })));
  batches.flat().forEach((row) => {
    const id = cleanText(row?.id);
    if (!id || seen.has(id)) return;
    seen.set(id, row);
  });
  return [...seen.values()];
}

async function loadFinancialDocumentRowById(id) {
  const rows = await restQuery(`financial_documents?select=*&id=eq.${encodeURIComponent(id)}&limit=1`, { method: "GET" });
  return Array.isArray(rows) && rows[0] ? rows[0] : null;
}

async function loadFinancialDocumentWithHistory(id) {
  const row = await loadFinancialDocumentRowById(id);
  if (!row) return null;
  const historyRows = await restQuery(`financial_document_history?select=*&document_id=eq.${encodeURIComponent(id)}&order=created_at.desc`, { method: "GET" });
  return toClientDocument({
    ...row,
    history: Array.isArray(historyRows) ? historyRows : [],
  });
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

async function deleteFinancialDocumentRow(id) {
  await restQuery(`financial_documents?id=eq.${encodeURIComponent(id)}`, {
    method: "DELETE",
  });
}

async function insertFinancialDocumentHistory(entries) {
  const payload = (Array.isArray(entries) ? entries : [entries]).filter(Boolean);
  if (!payload.length) return;
  await restQuery("financial_document_history", { method: "POST", body: payload });
}

async function listFinancialDocumentEntities() {
  const rows = await restQuery("financial_document_entities?select=*&order=name.asc", { method: "GET" });
  return Array.isArray(rows) ? rows.map(toClientEntity) : [];
}

async function loadFinancialDocumentEntityRowById(id) {
  const rows = await restQuery(`financial_document_entities?select=*&id=eq.${encodeURIComponent(id)}&limit=1`, { method: "GET" });
  return Array.isArray(rows) && rows[0] ? rows[0] : null;
}

async function insertFinancialDocumentEntity(body) {
  const payload = Array.isArray(body) ? body : [body];
  const rows = await restQuery("financial_document_entities", { method: "POST", body: payload, preferRepresentation: true });
  return Array.isArray(rows) && rows[0] ? rows[0] : null;
}

async function updateFinancialDocumentEntityRow(id, body) {
  const rows = await restQuery(`financial_document_entities?id=eq.${encodeURIComponent(id)}`, {
    method: "PATCH",
    body,
    preferRepresentation: true,
  });
  return Array.isArray(rows) && rows[0] ? rows[0] : null;
}

async function deleteFinancialDocumentEntityRow(id) {
  await restQuery(`financial_document_entities?id=eq.${encodeURIComponent(id)}`, {
    method: "DELETE",
  });
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
  loadFinancialDocumentEntityRowById,
  loadFinancialDocumentRowById,
  loadFinancialDocumentWithHistory,
  loadFinancialDocsSettings,
  loadFinancialDocsSettingsRecord,
  listFinancialDocumentEntities,
  listFinancialDocuments,
  listPotentialDuplicateFinancialDocuments,
  redirectUri,
  refreshDriveAccessToken,
  renameExistingDriveFileIfNeeded,
  saveFinancialDocsSettings,
  safeFinancialDocsSettings,
  insertFinancialDocument,
  insertFinancialDocumentEntity,
  insertFinancialDocumentHistory,
  updateFinancialDocumentRow,
  updateFinancialDocumentEntityRow,
  deleteFinancialDocumentRow,
  deleteFinancialDocumentEntityRow,
  updateFinancialDocsSettings,
};
