const { cleanText, parseBody, requireFeature, sendError } = require("./_supabase");
const {
  normalizeEntityName,
  normalizeEntityNif,
  sanitizeFinancialDocumentEntityInput,
  toClientEntity,
} = require("./_financial-docs");
const {
  deleteFinancialDocumentEntityRow,
  insertFinancialDocumentEntity,
  listFinancialDocumentEntities,
  listFinancialDocuments,
  loadFinancialDocumentEntityRowById,
  updateFinancialDocumentEntityRow,
} = require("./_financial-docs-service");

function conflict(message) {
  const error = new Error(message);
  error.statusCode = 409;
  return error;
}

function validateEntityUniqueness(candidate, existingRows, currentId = "") {
  const safeId = cleanText(currentId);
  const candidateNif = normalizeEntityNif(candidate.nif);
  const candidateName = normalizeEntityName(candidate.name);
  for (const row of Array.isArray(existingRows) ? existingRows : []) {
    if (cleanText(row.id) === safeId) continue;
    if (normalizeEntityNif(row.nif) === candidateNif) throw conflict("An entity with this NIF already exists.");
    if (normalizeEntityName(row.name) === candidateName) throw conflict("An entity with this Name already exists.");
  }
}

function isDocumentUsingEntity(documentRow, entityRow) {
  const docNif = normalizeEntityNif(documentRow.supplierNif || documentRow.supplier_nif);
  const docName = normalizeEntityName(documentRow.supplierName || documentRow.supplier_name);
  return (
    (docNif && docNif === normalizeEntityNif(entityRow.nif)) ||
    (docName && docName === normalizeEntityName(entityRow.name))
  );
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "financial-docs");

    if (req.method === "GET") {
      const rows = await listFinancialDocumentEntities();
      res.status(200).json({ rows });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const entity = sanitizeFinancialDocumentEntityInput(body);
      const existing = await listFinancialDocumentEntities();
      validateEntityUniqueness(entity, existing);
      const created = await insertFinancialDocumentEntity(entity);
      res.status(200).json({ row: toClientEntity(created) });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanText(req.query?.id || req.body?.id);
      if (!id) {
        res.status(400).json({ error: "Entity id is required." });
        return;
      }
      const existingRow = await loadFinancialDocumentEntityRowById(id);
      if (!existingRow) {
        res.status(404).json({ error: "Entity not found." });
        return;
      }
      const body = await parseBody(req);
      const entity = sanitizeFinancialDocumentEntityInput(body);
      const existing = await listFinancialDocumentEntities();
      validateEntityUniqueness(entity, existing, id);
      const updated = await updateFinancialDocumentEntityRow(id, entity);
      res.status(200).json({ row: toClientEntity(updated) });
      return;
    }

    if (req.method === "DELETE") {
      const id = cleanText(req.query?.id || req.body?.id);
      if (!id) {
        res.status(400).json({ error: "Entity id is required." });
        return;
      }
      const entity = await loadFinancialDocumentEntityRowById(id);
      if (!entity) {
        res.status(404).json({ error: "Entity not found." });
        return;
      }
      const docs = await listFinancialDocuments();
      if (docs.some((row) => isDocumentUsingEntity(row, entity))) {
        const error = new Error("This entity is used by one or more financial documents and cannot be deleted.");
        error.statusCode = 409;
        throw error;
      }
      await deleteFinancialDocumentEntityRow(id);
      res.status(200).json({ ok: true });
      return;
    }

    res.setHeader("Allow", "GET,POST,PUT,DELETE");
    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
