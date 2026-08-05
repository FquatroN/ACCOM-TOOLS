const { cleanText, parseBody, requireFeature, sendError } = require("./_supabase");
const financialDocsParseHandler = require("./financial-docs-parse");
const {
  buildStoredFileName,
  findPossibleDuplicates,
  normalizeEntityName,
  normalizeEntityNif,
  safeFinancialDocsSettings,
  sanitizeFinancialDocumentInput,
  sha256Base64Content,
} = require("./_financial-docs");
const {
  attachDocumentFile,
  deleteDriveFile,
  deleteFinancialDocumentRow,
  insertFinancialDocument,
  insertFinancialDocumentEntity,
  insertFinancialDocumentHistory,
  listFinancialDocumentEntities,
  listFinancialDocuments,
  listPotentialDuplicateFinancialDocuments,
  loadFinancialDocumentRowById,
  loadFinancialDocumentWithHistory,
  loadFinancialDocsSettings,
  refreshDriveAccessToken,
  renameExistingDriveFileIfNeeded,
  updateFinancialDocumentRow,
} = require("./_financial-docs-service");

const TRACKED_FIELDS = [
  ["cc", "CC"],
  ["document_date", "Date"],
  ["doc_number", "Doc Number"],
  ["description", "Description"],
  ["supplier_nif", "Supplier NIF"],
  ["supplier_name", "Name"],
  ["amount", "Amount"],
  ["vat_amount", "VAT Amount"],
  ["payment", "Payment"],
  ["document_type", "Type"],
  ["fat", "Fat"],
  ["category", "Category"],
  ["status", "Status"],
];

function normalizeUpload(upload, userEmail) {
  if (!upload || typeof upload !== "object") return null;
  const base64Content = cleanText(upload.base64Content);
  if (!base64Content) return null;
  return {
    base64Content,
    originalFilename: cleanText(upload.originalFilename || upload.fileName) || "document.pdf",
    mimeType: cleanText(upload.mimeType) || "application/pdf",
    fileSize: Number(upload.fileSize || 0),
    fileHash: cleanText(upload.fileHash) || sha256Base64Content(base64Content),
    uploadedBy: cleanText(upload.uploadedBy) || cleanText(userEmail),
  };
}

function buildDbPayload(doc, userEmail) {
  return {
    cc: doc.cc,
    document_date: doc.documentDate,
    doc_number: doc.docNumber,
    description: doc.description,
    supplier_nif: doc.supplierNif,
    supplier_name: doc.supplierName,
    amount: doc.amount,
    vat_amount: doc.vatAmount,
    payment: doc.payment,
    document_type: doc.docType,
    fat: doc.fat,
    category: doc.category,
    status: doc.status,
    created_by: cleanText(userEmail),
  };
}

function duplicateConflict(duplicates) {
  const error = new Error("Possible duplicate found.");
  error.statusCode = 409;
  error.duplicates = duplicates;
  return error;
}

function trackFieldChanges(before, after) {
  return TRACKED_FIELDS
    .map(([key, label]) => {
      const previous = before?.[key] ?? null;
      const next = after?.[key] ?? null;
      return JSON.stringify(previous) === JSON.stringify(next)
        ? null
        : {
            action_type: "field_update",
            field_name: key,
            message: `${label} updated.`,
            old_value: previous,
            new_value: next,
          };
    })
    .filter(Boolean);
}

function findMatchingEntity(entities, doc) {
  const targetNif = normalizeEntityNif(doc?.supplierNif || doc?.supplier_nif);
  const targetName = normalizeEntityName(doc?.supplierName || doc?.supplier_name);
  return (Array.isArray(entities) ? entities : []).find((row) => {
    const rowNif = normalizeEntityNif(row?.nif);
    const rowName = normalizeEntityName(row?.name);
    return (
      (targetNif && rowNif === targetNif) ||
      (targetName && rowName === targetName)
    );
  }) || null;
}

async function ensureFinancialDocumentEntityForDoc(doc) {
  const supplierName = cleanText(doc?.supplierName || doc?.supplier_name);
  const supplierNif = cleanText(doc?.supplierNif || doc?.supplier_nif);
  if (!supplierName || !supplierNif || !normalizeEntityNif(supplierNif)) {
    return { entity: null, created: false };
  }

  const entities = await listFinancialDocumentEntities();
  const existing = findMatchingEntity(entities, { supplierName, supplierNif });
  if (existing) {
    return { entity: existing, created: false };
  }

  try {
    const created = await insertFinancialDocumentEntity({
      nif: supplierNif,
      name: supplierName,
      address: "",
    });
    return { entity: created, created: true };
  } catch (error) {
    if (error?.statusCode === 409) {
      const retryEntities = await listFinancialDocumentEntities();
      const matched = findMatchingEntity(retryEntities, { supplierName, supplierNif });
      if (matched) return { entity: matched, created: false };
    }
    return { entity: null, created: false, error };
  }
}

module.exports = async function handler(req, res) {
  let operationStep = "";
  try {
    operationStep = "checking financial documents access";
    const auth = await requireFeature(req, "app", "financial-docs");
    const userEmail = cleanText(auth.user?.email) || cleanText(auth.user?.id);
    const action = cleanText(req.query?.action).toLowerCase();

    if (req.method === "GET") {
      const id = cleanText(req.query?.id);
      const settings = await loadFinancialDocsSettings();
      if (id) {
        const document = await loadFinancialDocumentWithHistory(id);
        if (!document) {
          res.status(404).json({ error: "Financial document not found." });
          return;
        }
        res.status(200).json({ row: document, settings: safeFinancialDocsSettings(settings) });
        return;
      }
      const rows = await listFinancialDocuments({
        createdFrom: cleanText(req.query?.created_from),
        createdTo: cleanText(req.query?.created_to),
        dateFrom: cleanText(req.query?.date_from),
        dateTo: cleanText(req.query?.date_to),
        supplierSearch: cleanText(req.query?.supplier),
        descriptionSearch: cleanText(req.query?.description),
        payment: cleanText(req.query?.payment),
        docType: cleanText(req.query?.doc_type),
        fat: cleanText(req.query?.fat),
        category: cleanText(req.query?.category),
        status: cleanText(req.query?.status),
        sortBy: cleanText(req.query?.sort_by),
        sortDirection: cleanText(req.query?.sort_direction),
      });
      res.status(200).json({ rows, settings: safeFinancialDocsSettings(settings) });
      return;
    }

    if (req.method === "POST") {
      operationStep = "reading financial document request";
      const body = await parseBody(req);
      if (action === "parse") {
        operationStep = "parsing financial document attachment";
        const result = await financialDocsParseHandler.parseFinancialDocumentRequest(body);
        res.status(200).json(result);
        return;
      }
      operationStep = "loading financial document settings";
      const settings = await loadFinancialDocsSettings();
      operationStep = "validating financial document fields";
      const sanitized = sanitizeFinancialDocumentInput(body, settings);
      const upload = normalizeUpload(body?.attachmentUpload, userEmail);
      operationStep = "checking financial document duplicates";
      const duplicateCandidates = await listPotentialDuplicateFinancialDocuments({
        ...sanitized,
        fileHash: upload?.fileHash,
      });
      const duplicates = findPossibleDuplicates(sanitized, duplicateCandidates, {
        checksum: upload?.fileHash,
      });
      if (duplicates.length && !body?.confirmDuplicate) throw duplicateConflict(duplicates);

      operationStep = "creating financial document";
      let created = await insertFinancialDocument({
        ...buildDbPayload(sanitized, userEmail),
        ocr_fields: body?.ocrFields && typeof body.ocrFields === "object" ? body.ocrFields : {},
        ocr_raw_text: cleanText(body?.ocrRawText),
      });
      operationStep = "syncing financial document entity";
      const entitySync = await ensureFinancialDocumentEntityForDoc(created);

      const historyEntries = [{
        document_id: cleanText(created?.id),
        action_type: "created",
        field_name: "",
        message: "Document created.",
        old_value: null,
        new_value: null,
        metadata: { source: cleanText(body?.source || "manual") },
        created_by: userEmail,
      }];

      if (duplicates.length) {
        historyEntries.push({
          document_id: cleanText(created?.id),
          action_type: "duplicate_warning",
          field_name: "",
          message: "Possible duplicate found. Save was confirmed by the user.",
          old_value: null,
          new_value: null,
          metadata: { duplicates },
          created_by: userEmail,
        });
      }

      if (entitySync.created) {
        historyEntries.push({
          document_id: cleanText(created?.id),
          action_type: "entity_created",
          field_name: "",
          message: "Entity created automatically from document data.",
          old_value: null,
          new_value: entitySync.entity?.name || sanitized.supplierName,
          metadata: {
            nif: entitySync.entity?.nif || sanitized.supplierNif,
            name: entitySync.entity?.name || sanitized.supplierName,
          },
          created_by: userEmail,
        });
      }

      if (upload) {
        operationStep = "uploading financial document attachment";
        const attachmentFields = await attachDocumentFile(req, created, upload, settings);
        operationStep = "saving financial document attachment fields";
        created = await updateFinancialDocumentRow(cleanText(created.id), attachmentFields);
        historyEntries.push({
          document_id: cleanText(created?.id),
          action_type: "file_uploaded",
          field_name: "",
          message: "Attachment uploaded.",
          old_value: null,
          new_value: attachmentFields.stored_filename,
          metadata: { originalFilename: attachmentFields.original_filename, fileHash: attachmentFields.file_hash },
          created_by: userEmail,
        });
      }

      if (body?.ocrFields || body?.ocrRawText) {
        historyEntries.push({
          document_id: cleanText(created?.id),
          action_type: "ocr_parse",
          field_name: "",
          message: "OCR/AI extracted draft values.",
          old_value: null,
          new_value: null,
          metadata: { fields: body?.ocrFields || {} },
          created_by: userEmail,
        });
      }

      operationStep = "saving financial document history";
      await insertFinancialDocumentHistory(historyEntries);
      operationStep = "reloading financial document after save";
      const row = await loadFinancialDocumentWithHistory(cleanText(created.id));
      res.status(200).json({ row, duplicates });
      return;
    }

    if (req.method === "PUT") {
      operationStep = "reading financial document id";
      const id = cleanText(req.query?.id || req.body?.id);
      if (!id) {
        res.status(400).json({ error: "Document id is required." });
        return;
      }
      operationStep = "reading financial document request";
      const body = await parseBody(req);
      operationStep = "loading existing financial document";
      const existing = await loadFinancialDocumentRowById(id);
      if (!existing) {
        res.status(404).json({ error: "Financial document not found." });
        return;
      }
      operationStep = "loading financial document settings";
      const settings = await loadFinancialDocsSettings();
      operationStep = "validating financial document fields";
      const sanitized = sanitizeFinancialDocumentInput(body, settings);
      const upload = normalizeUpload(body?.attachmentUpload, userEmail);
      operationStep = "checking financial document duplicates";
      const duplicateCandidates = await listPotentialDuplicateFinancialDocuments({
        ...sanitized,
        fileHash: upload?.fileHash,
      });
      const duplicates = findPossibleDuplicates(sanitized, duplicateCandidates, {
        currentId: id,
        checksum: upload?.fileHash,
      });
      if (duplicates.length && !body?.confirmDuplicate) throw duplicateConflict(duplicates);
      operationStep = "loading financial document history";
      const existingWithHistory = await loadFinancialDocumentWithHistory(id);

      operationStep = "updating financial document";
      let updated = await updateFinancialDocumentRow(id, {
        ...buildDbPayload(sanitized, existing.created_by || userEmail),
        ocr_fields: body?.ocrFields && typeof body.ocrFields === "object" ? body.ocrFields : (existing.ocr_fields || {}),
        ocr_raw_text: cleanText(body?.ocrRawText) || cleanText(existing.ocr_raw_text),
      });
      operationStep = "syncing financial document entity";
      const entitySync = await ensureFinancialDocumentEntityForDoc(updated);

      const historyEntries = trackFieldChanges(existing, updated).map((item) => ({
        document_id: id,
        ...item,
        metadata: {},
        created_by: userEmail,
      }));

      if (duplicates.length) {
        historyEntries.push({
          document_id: id,
          action_type: "duplicate_warning",
          field_name: "",
          message: "Possible duplicate found. Save was confirmed by the user.",
          old_value: null,
          new_value: null,
          metadata: { duplicates },
          created_by: userEmail,
        });
      } else if (cleanText(existingWithHistory?.duplicateWarningMessage)) {
        historyEntries.push({
          document_id: id,
          action_type: "duplicate_warning_resolved",
          field_name: "",
          message: "Duplicate warning resolved.",
          old_value: cleanText(existingWithHistory.duplicateWarningMessage),
          new_value: null,
          metadata: {},
          created_by: userEmail,
        });
      }

      if (entitySync.created) {
        historyEntries.push({
          document_id: id,
          action_type: "entity_created",
          field_name: "",
          message: "Entity created automatically from document data.",
          old_value: null,
          new_value: entitySync.entity?.name || sanitized.supplierName,
          metadata: {
            nif: entitySync.entity?.nif || sanitized.supplierNif,
            name: entitySync.entity?.name || sanitized.supplierName,
          },
          created_by: userEmail,
        });
      }

      if (upload) {
        const oldDriveFileId = cleanText(existing.drive_file_id);
        operationStep = "uploading financial document attachment";
        const attachmentFields = await attachDocumentFile(req, updated, upload, settings);
        operationStep = "saving financial document attachment fields";
        updated = await updateFinancialDocumentRow(id, attachmentFields);
        if (oldDriveFileId && oldDriveFileId !== cleanText(attachmentFields.drive_file_id)) {
          try {
            operationStep = "deleting replaced financial document attachment";
            const refreshed = await refreshDriveAccessToken(settings);
            await deleteDriveFile(refreshed.accessToken, oldDriveFileId);
          } catch {}
        }
        historyEntries.push({
          document_id: id,
          action_type: oldDriveFileId ? "file_replaced" : "file_uploaded",
          field_name: "",
          message: oldDriveFileId ? "Attachment replaced." : "Attachment uploaded.",
          old_value: oldDriveFileId || null,
          new_value: attachmentFields.stored_filename,
          metadata: { originalFilename: attachmentFields.original_filename, fileHash: attachmentFields.file_hash },
          created_by: userEmail,
        });
      } else if (cleanText(existing.drive_file_id)) {
        let nextName = "";
        let renameError = null;
        try {
          operationStep = "renaming financial document attachment";
          nextName = await renameExistingDriveFileIfNeeded(req, updated, settings);
        } catch (error) {
          renameError = error;
        }
        if (nextName) {
          updated = await updateFinancialDocumentRow(id, { stored_filename: nextName });
          historyEntries.push({
            document_id: id,
            action_type: "file_renamed",
            field_name: "",
            message: "Attachment renamed to match document data.",
            old_value: cleanText(existing.stored_filename),
            new_value: nextName,
            metadata: {},
            created_by: userEmail,
          });
        } else if (renameError) {
          historyEntries.push({
            document_id: id,
            action_type: "file_rename_failed",
            field_name: "",
            message: `Attachment rename failed: ${cleanText(renameError.message) || "Unknown error"}`,
            old_value: cleanText(existing.stored_filename),
            new_value: null,
            metadata: { statusCode: renameError.statusCode || null },
            created_by: userEmail,
          });
        }
      }

      if (body?.ocrFields || body?.ocrRawText) {
        historyEntries.push({
          document_id: id,
          action_type: "ocr_parse",
          field_name: "",
          message: "OCR/AI extracted draft values.",
          old_value: null,
          new_value: null,
          metadata: { fields: body?.ocrFields || {} },
          created_by: userEmail,
        });
      }

      operationStep = "saving financial document history";
      if (historyEntries.length) await insertFinancialDocumentHistory(historyEntries);
      operationStep = "reloading financial document after save";
      const row = await loadFinancialDocumentWithHistory(id);
      res.status(200).json({ row, duplicates });
      return;
    }

    if (req.method === "DELETE") {
      const id = cleanText(req.query?.id || req.body?.id);
      if (!id) {
        res.status(400).json({ error: "Document id is required." });
        return;
      }
      const existing = await loadFinancialDocumentRowById(id);
      if (!existing) {
        res.status(404).json({ error: "Financial document not found." });
        return;
      }
      const driveFileId = cleanText(existing.drive_file_id);
      if (driveFileId) {
        const settings = await loadFinancialDocsSettings();
        const refreshed = await refreshDriveAccessToken(settings);
        await deleteDriveFile(refreshed.accessToken, driveFileId);
      }
      await deleteFinancialDocumentRow(id);
      res.status(200).json({ ok: true });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    if (error.statusCode === 409) {
      res.status(409).json({ error: error.message, duplicates: error.duplicates || [] });
      return;
    }
    if (operationStep && error && !String(error.message || "").toLowerCase().includes(operationStep.toLowerCase())) {
      error.message = `Failed while ${operationStep}: ${error.message || "Unexpected error"}`;
    }
    sendError(res, error);
  }
};
