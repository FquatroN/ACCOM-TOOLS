const crypto = require("crypto");

const { cleanText, normalizeDate, normalizeNumeric } = require("./_supabase");

const FINANCIAL_DOCS_SETTINGS_KEY = "financial_documents";

const DEFAULT_FINANCIAL_DOCS_SETTINGS = {
  attributes: {
    cc: ["H", "A"],
    payment: ["Banco", "Visa", "Cash", "Caixa", "Miguel", "Carlos", "Odete"],
    docType: ["R", "F"],
    fat: ["S", "N"],
    category: ["Renda", "Alimentacao", "Utility", "Setup", "Imposto", "Financ", "Ordenados", "Outros", "Servicos"],
    status: ["Draft", "Confirmed"],
  },
  drive: {
    connected: false,
    connectedAt: "",
    accountEmail: "",
    folderPath: "Financial Documents",
  },
  rules: [],
};

function extractGoogleDriveFolderId(value) {
  const raw = cleanText(value);
  if (!raw) return "";
  try {
    const url = new URL(raw);
    if (!/drive\.google\.com$/i.test(url.hostname)) return "";
    const folderMatch = url.pathname.match(/\/folders\/([a-zA-Z0-9_-]+)/);
    if (folderMatch?.[1]) return cleanText(folderMatch[1]);
    const queryId = cleanText(url.searchParams.get("id"));
    if (queryId) return queryId;
  } catch {}
  return "";
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function uniqueList(values, fallback = []) {
  const source = Array.isArray(values) ? values : fallback;
  const seen = new Set();
  const output = [];
  source.forEach((item) => {
    const next = cleanText(item);
    if (!next) return;
    const key = next.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    output.push(next);
  });
  return output;
}

function sanitizeFinancialDocsSettings(source = {}) {
  const settings = source && typeof source === "object" ? source : {};
  const defaults = clone(DEFAULT_FINANCIAL_DOCS_SETTINGS);
  const rawAttributes = settings.attributes && typeof settings.attributes === "object" ? settings.attributes : {};
  const rawDrive = settings.drive && typeof settings.drive === "object" ? settings.drive : {};
  const rawGoogle = settings.google && typeof settings.google === "object" ? settings.google : {};
  const rawRules = Array.isArray(settings.rules) ? settings.rules : defaults.rules;
  return {
    attributes: {
      cc: uniqueList(rawAttributes.cc, defaults.attributes.cc),
      payment: uniqueList(rawAttributes.payment, defaults.attributes.payment),
      docType: uniqueList(rawAttributes.docType || rawAttributes.type, defaults.attributes.docType),
      fat: uniqueList(rawAttributes.fat, defaults.attributes.fat),
      category: uniqueList(rawAttributes.category, defaults.attributes.category),
      status: uniqueList(rawAttributes.status, defaults.attributes.status),
    },
    drive: {
      connected: !!cleanText(rawGoogle.refreshToken || rawDrive.refreshToken),
      connectedAt: cleanText(rawGoogle.connectedAt || rawDrive.connectedAt),
      accountEmail: cleanText(rawGoogle.accountEmail || rawDrive.accountEmail),
      folderPath: cleanText(rawDrive.folderPath || rawDrive.path) || defaults.drive.folderPath,
      baseFolderId: cleanText(rawDrive.baseFolderId) || extractGoogleDriveFolderId(rawDrive.folderPath || rawDrive.path),
      accessToken: cleanText(rawDrive.accessToken || rawGoogle.accessToken),
      refreshToken: cleanText(rawDrive.refreshToken || rawGoogle.refreshToken),
      tokenExpiresAt: cleanText(rawDrive.tokenExpiresAt || rawGoogle.tokenExpiresAt),
      oauthState: cleanText(rawDrive.oauthState || rawGoogle.oauthState),
    },
    rules: rawRules
      .map((rule, index) => ({
        id: cleanText(rule?.id) || `rule-${index + 1}`,
        nif: cleanText(rule?.nif),
        name: cleanText(rule?.name).replace(/\s+/g, " ").trim(),
        cc: cleanText(rule?.cc),
        payment: cleanText(rule?.payment),
        docType: cleanText(rule?.docType || rule?.type),
        fat: cleanText(rule?.fat),
        category: cleanText(rule?.category),
      }))
      .filter((rule) => normalizeEntityNif(rule.nif) && normalizeEntityName(rule.name)),
  };
}

function safeFinancialDocsSettings(settings = DEFAULT_FINANCIAL_DOCS_SETTINGS) {
  const safe = sanitizeFinancialDocsSettings(settings);
  return {
    attributes: safe.attributes,
    drive: {
      connected: !!safe.drive.connected,
      connectedAt: safe.drive.connectedAt,
      accountEmail: safe.drive.accountEmail,
      folderPath: safe.drive.folderPath,
      baseFolderId: safe.drive.baseFolderId,
    },
    rules: safe.rules,
  };
}

function normalizeEntityNif(value) {
  return cleanText(value).replace(/\D+/g, "");
}

function normalizeEntityName(value) {
  return cleanText(value).replace(/\s+/g, " ").trim().toLowerCase();
}

function sanitizeFinancialDocumentEntityInput(input = {}) {
  const entity = {
    nif: cleanText(input.nif),
    name: cleanText(input.name).replace(/\s+/g, " ").trim(),
    address: cleanText(input.address),
  };
  if (!normalizeEntityNif(entity.nif)) throw badRequest("NIF is required.");
  if (!normalizeEntityName(entity.name)) throw badRequest("Name is required.");
  return entity;
}

function normalizeStatus(value, settings = DEFAULT_FINANCIAL_DOCS_SETTINGS) {
  const raw = cleanText(value);
  if (raw) {
    const match = sanitizeFinancialDocsSettings(settings).attributes.status.find((item) => item.toLowerCase() === raw.toLowerCase());
    if (match) return match;
  }
  return sanitizeFinancialDocsSettings(settings).attributes.status[0] || "Draft";
}

function normalizeMoney(value, fallback = 0) {
  const normalized = normalizeNumeric(value);
  if (normalized === null || normalized === undefined || Number.isNaN(normalized)) return fallback;
  return Number(Number(normalized).toFixed(2));
}

function sanitizeFinancialDocumentInput(input = {}, settings = DEFAULT_FINANCIAL_DOCS_SETTINGS) {
  const safeSettings = sanitizeFinancialDocsSettings(settings);
  const doc = {
    cc: cleanText(input.cc),
    documentDate: normalizeDate(input.documentDate || input.date),
    docNumber: cleanText(input.docNumber || input.documentNumber),
    description: cleanText(input.description),
    supplierNif: cleanText(input.supplierNif || input.nif),
    supplierName: cleanText(input.supplierName || input.name),
    amount: normalizeMoney(input.amount, Number.NaN),
    vatAmount: cleanText(input.vatAmount) === "" ? null : normalizeMoney(input.vatAmount, Number.NaN),
    payment: cleanText(input.payment),
    docType: cleanText(input.docType || input.type),
    fat: cleanText(input.fat),
    category: cleanText(input.category),
    status: normalizeStatus(input.status, safeSettings),
  };
  const isDraft = cleanText(doc.status).toLowerCase() === "draft";

  if (!doc.documentDate) throw badRequest("Date is required.");
  if (!doc.description) throw badRequest("Description is required.");
  if (!doc.supplierName) throw badRequest("Name is required.");
  if (!Number.isFinite(doc.amount)) throw badRequest("Amount is required.");
  if (!isDraft && !doc.cc) throw badRequest("CC is required.");
  if (!isDraft && !doc.payment) throw badRequest("Payment is required.");
  if (!isDraft && !doc.docType) throw badRequest("Type is required.");
  if (!isDraft && !doc.fat) throw badRequest("Fat is required.");
  if (!isDraft && !doc.category) throw badRequest("Category is required.");
  if (doc.vatAmount !== null && !Number.isFinite(doc.vatAmount)) throw badRequest("VAT Amount is invalid.");
  return doc;
}

function badRequest(message) {
  const error = new Error(message);
  error.statusCode = 400;
  return error;
}

function isoDateOnly(value) {
  const normalized = normalizeDate(value);
  return /^\d{4}-\d{2}-\d{2}$/.test(normalized) ? normalized : "";
}

function isoTimestamp(value) {
  const raw = cleanText(value);
  const ms = Date.parse(raw);
  return Number.isFinite(ms) ? new Date(ms).toISOString() : "";
}

function monthlyFolderKey(createdAt) {
  const iso = isoTimestamp(createdAt);
  if (!iso) return "";
  const dt = new Date(iso);
  return `${dt.getUTCFullYear()}${String(dt.getUTCMonth() + 1).padStart(2, "0")}`;
}

function sanitizeFileStemPart(value, fallback = "Document") {
  const normalized = cleanText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const clipped = normalized.slice(0, 72).trim();
  return clipped || fallback;
}

function buildStoredFileName(recordLike = {}, originalFilename = "document.pdf") {
  const extMatch = cleanText(originalFilename).match(/(\.[a-z0-9]{1,8})$/i);
  const ext = extMatch ? extMatch[1].toLowerCase() : ".pdf";
  const docDate = isoDateOnly(recordLike.documentDate || recordLike.document_date).replace(/-/g, "") || "00000000";
  const supplierName = sanitizeFileStemPart(recordLike.supplierName || recordLike.supplier_name, "Supplier");
  const description = sanitizeFileStemPart(recordLike.description, "Document");
  return `${docDate} - ${supplierName} - ${description}${ext}`;
}

function sha256Base64Content(base64) {
  const buffer = Buffer.from(String(base64 || ""), "base64");
  return crypto.createHash("sha256").update(buffer).digest("hex");
}

function normalizeNameForDuplicate(value) {
  return cleanText(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function diceCoefficient(a, b) {
  const left = normalizeNameForDuplicate(a);
  const right = normalizeNameForDuplicate(b);
  if (!left || !right) return 0;
  if (left === right) return 1;
  if (left.includes(right) || right.includes(left)) return 0.92;
  const pairs = (value) => {
    const output = [];
    for (let i = 0; i < value.length - 1; i += 1) output.push(value.slice(i, i + 2));
    return output;
  };
  const leftPairs = pairs(left);
  const rightPairs = pairs(right);
  if (!leftPairs.length || !rightPairs.length) return 0;
  const counts = new Map();
  leftPairs.forEach((pair) => counts.set(pair, (counts.get(pair) || 0) + 1));
  let intersection = 0;
  rightPairs.forEach((pair) => {
    const current = counts.get(pair) || 0;
    if (current > 0) {
      intersection += 1;
      counts.set(pair, current - 1);
    }
  });
  return (2 * intersection) / (leftPairs.length + rightPairs.length);
}

function findPossibleDuplicates(candidate, existingRows, { currentId = "", checksum = "" } = {}) {
  const source = Array.isArray(existingRows) ? existingRows : [];
  const safeCurrentId = cleanText(currentId);
  const candidateNif = cleanText(candidate.supplierNif || candidate.supplier_nif);
  const candidateDocNumber = cleanText(candidate.docNumber || candidate.doc_number);
  const candidateType = cleanText(candidate.docType || candidate.document_type);
  const candidateDate = isoDateOnly(candidate.documentDate || candidate.document_date);
  const candidateAmount = normalizeMoney(candidate.amount, Number.NaN);
  const candidateName = cleanText(candidate.supplierName || candidate.supplier_name);
  const candidateChecksum = cleanText(checksum || candidate.fileHash || candidate.file_hash);
  const matches = [];

  source.forEach((row) => {
    const rowId = cleanText(row.id);
    if (!rowId || rowId === safeCurrentId) return;
    const rowNif = cleanText(row.supplier_nif || row.supplierNif);
    const rowDocNumber = cleanText(row.doc_number || row.docNumber);
    const rowType = cleanText(row.document_type || row.docType);
    const rowDate = isoDateOnly(row.document_date || row.documentDate);
    const rowAmount = normalizeMoney(row.amount, Number.NaN);
    const rowName = cleanText(row.supplier_name || row.supplierName);
    const rowChecksum = cleanText(row.file_hash || row.fileHash);
    let reason = "";

    if (candidateNif && candidateDocNumber && candidateType && rowNif === candidateNif && rowDocNumber === candidateDocNumber && rowType === candidateType) {
      reason = "same supplier NIF + document number + document type";
    } else if (candidateNif && candidateDate && Number.isFinite(candidateAmount) && rowNif === candidateNif && rowDate === candidateDate && rowAmount === candidateAmount) {
      reason = "same supplier NIF + document date + amount";
    } else if (candidateChecksum && rowChecksum && rowChecksum === candidateChecksum) {
      reason = "same file hash/checksum";
    } else if (candidateDocNumber && candidateDate && Number.isFinite(candidateAmount) && rowDocNumber === candidateDocNumber && rowDate === candidateDate && rowAmount === candidateAmount) {
      reason = "same document number + amount + date";
    } else if (candidateName && candidateDate && Number.isFinite(candidateAmount) && rowDate === candidateDate && rowAmount === candidateAmount && diceCoefficient(candidateName, rowName) >= 0.78) {
      reason = "similar supplier name + same date + same amount";
    }

    if (reason) {
      matches.push({
        id: rowId,
        reason,
        supplierName: rowName,
        supplierNif: rowNif,
        docNumber: rowDocNumber,
        documentDate: rowDate,
        amount: rowAmount,
        status: cleanText(row.status),
      });
    }
  });
  return matches;
}

function latestDuplicateWarningMessage(history = [], fallbackMessage = "") {
  const entries = (Array.isArray(history) ? history : [])
    .filter((item) => {
      const actionType = cleanText(item.actionType || item.action_type);
      return actionType === "duplicate_warning" || actionType === "duplicate_warning_resolved";
    })
    .sort((a, b) => {
      const left = Date.parse(cleanText(a.createdAt || a.created_at));
      const right = Date.parse(cleanText(b.createdAt || b.created_at));
      return (Number.isFinite(right) ? right : 0) - (Number.isFinite(left) ? left : 0);
    });
  const latest = entries[0];
  const latestAction = cleanText(latest?.actionType || latest?.action_type);
  if (latestAction === "duplicate_warning_resolved") return "";
  if (latestAction === "duplicate_warning") return cleanText(latest.message) || cleanText(fallbackMessage);
  return cleanText(fallbackMessage);
}

function toClientDocument(row = {}) {
  const history = Array.isArray(row.history) ? row.history.map(toClientHistory) : [];
  const hasDuplicateWarningOverride =
    Object.prototype.hasOwnProperty.call(row, "duplicate_warning_message") ||
    Object.prototype.hasOwnProperty.call(row, "duplicateWarningMessage");
  const duplicateWarningMessage = hasDuplicateWarningOverride
    ? cleanText(row.duplicate_warning_message || row.duplicateWarningMessage)
    : latestDuplicateWarningMessage(history);
  return {
    id: cleanText(row.id),
    createdAt: cleanText(row.created_at || row.createdAt),
    updatedAt: cleanText(row.updated_at || row.updatedAt),
    cc: cleanText(row.cc),
    documentDate: cleanText(row.document_date || row.documentDate),
    docNumber: cleanText(row.doc_number || row.docNumber),
    description: cleanText(row.description),
    supplierNif: cleanText(row.supplier_nif || row.supplierNif),
    supplierName: cleanText(row.supplier_name || row.supplierName),
    amount: normalizeMoney(row.amount, 0),
    vatAmount: row.vat_amount === null || row.vatAmount === null || cleanText(row.vat_amount ?? row.vatAmount) === "" ? null : normalizeMoney(row.vat_amount ?? row.vatAmount, 0),
    payment: cleanText(row.payment),
    docType: cleanText(row.document_type || row.docType),
    fat: cleanText(row.fat),
    category: cleanText(row.category),
    status: cleanText(row.status),
    driveFileId: cleanText(row.drive_file_id || row.driveFileId),
    driveFolderId: cleanText(row.drive_folder_id || row.driveFolderId),
    driveFileUrl: cleanText(row.drive_file_url || row.driveFileUrl),
    originalFilename: cleanText(row.original_filename || row.originalFilename),
    storedFilename: cleanText(row.stored_filename || row.storedFilename),
    mimeType: cleanText(row.mime_type || row.mimeType),
    fileSize: Number(row.file_size || row.fileSize || 0),
    fileHash: cleanText(row.file_hash || row.fileHash),
    uploadedBy: cleanText(row.uploaded_by || row.uploadedBy),
    uploadedAt: cleanText(row.uploaded_at || row.uploadedAt),
    ocrFields: row.ocr_fields && typeof row.ocr_fields === "object" ? row.ocr_fields : (row.ocrFields && typeof row.ocrFields === "object" ? row.ocrFields : {}),
    ocrRawText: cleanText(row.ocr_raw_text || row.ocrRawText),
    duplicateWarningMessage,
    history,
  };
}

function toClientHistory(row = {}) {
  return {
    id: cleanText(row.id),
    documentId: cleanText(row.document_id || row.documentId),
    actionType: cleanText(row.action_type || row.actionType),
    fieldName: cleanText(row.field_name || row.fieldName),
    message: cleanText(row.message),
    oldValue: row.old_value ?? row.oldValue ?? null,
    newValue: row.new_value ?? row.newValue ?? null,
    metadata: row.metadata && typeof row.metadata === "object" ? row.metadata : {},
    createdBy: cleanText(row.created_by || row.createdBy),
    createdAt: cleanText(row.created_at || row.createdAt),
  };
}

function toClientEntity(row = {}) {
  return {
    id: cleanText(row.id),
    nif: cleanText(row.nif),
    name: cleanText(row.name).replace(/\s+/g, " ").trim(),
    address: cleanText(row.address),
    createdAt: cleanText(row.created_at || row.createdAt),
    updatedAt: cleanText(row.updated_at || row.updatedAt),
  };
}

module.exports = {
  DEFAULT_FINANCIAL_DOCS_SETTINGS,
  FINANCIAL_DOCS_SETTINGS_KEY,
  buildStoredFileName,
  clone,
  diceCoefficient,
  normalizeEntityName,
  normalizeEntityNif,
  findPossibleDuplicates,
  latestDuplicateWarningMessage,
  monthlyFolderKey,
  safeFinancialDocsSettings,
  sanitizeFinancialDocumentEntityInput,
  sanitizeFinancialDocumentInput,
  sanitizeFinancialDocsSettings,
  extractGoogleDriveFolderId,
  sha256Base64Content,
  toClientDocument,
  toClientEntity,
  toClientHistory,
};
