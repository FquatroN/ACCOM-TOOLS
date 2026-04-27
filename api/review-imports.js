const { randomUUID } = require("crypto");

const {
  cleanText,
  parseBody,
  requireFeature,
  restQuery,
  sanitizeReviewInput,
  sendError,
  verifyUser,
} = require("./_supabase");

async function loadRun(id) {
  const rows = await restQuery(
    `review_import_runs?select=*,properties(name)&id=eq.${encodeURIComponent(id)}&limit=1`,
    { method: "GET" }
  );
  return Array.isArray(rows) && rows[0] ? rows[0] : null;
}

async function loadStagingRows(importRunId) {
  const rows = await restQuery(
    `review_import_staging?select=*&import_run_id=eq.${encodeURIComponent(importRunId)}&order=created_at.asc`,
    { method: "GET" }
  );
  return Array.isArray(rows) ? rows : [];
}

async function findReviewIdByPayload(payload) {
  const source = cleanText(payload.source);
  const sourceReviewId = cleanText(payload.source_review_id);
  if (source && sourceReviewId) {
    const rows = await restQuery(
      `reviews?select=id&source=eq.${encodeURIComponent(source)}&source_review_id=eq.${encodeURIComponent(sourceReviewId)}&limit=1`,
      { method: "GET" }
    );
    if (Array.isArray(rows) && rows[0]?.id) return rows[0].id;
  }
  const stableKey = stableReviewDuplicateKey(payload);
  if (stableKey) {
    const rows = await restQuery(
      `reviews?select=id,property_id,source,review_date,reviewer_name,rating_raw&property_id=eq.${encodeURIComponent(payload.property_id)}&source=eq.${encodeURIComponent(source)}&review_date=eq.${encodeURIComponent(payload.review_date)}&reviewer_name=eq.${encodeURIComponent(payload.reviewer_name)}&rating_raw=eq.${encodeURIComponent(payload.rating_raw)}&limit=1`,
      { method: "GET" }
    );
    if (Array.isArray(rows) && rows[0]?.id) return rows[0].id;
  }
  const fingerprint = cleanText(payload.dedupe_fingerprint);
  if (fingerprint) {
    const rows = await restQuery(
      `reviews?select=id&dedupe_fingerprint=eq.${encodeURIComponent(fingerprint)}&limit=1`,
      { method: "GET" }
    );
    if (Array.isArray(rows) && rows[0]?.id) return rows[0].id;
  }
  return "";
}

function isDuplicateReviewError(error) {
  const message = cleanText(error?.message).toLowerCase();
  return message.includes("duplicate key") || message.includes("unique constraint");
}

function stableReviewDuplicateKey(row) {
  const source = cleanText(row.source).toLowerCase();
  const parts = [
    cleanText(row.property_id || row.propertyId),
    source,
    cleanText(row.review_date || row.reviewDate),
    cleanText(row.reviewer_name || row.reviewerName).toLowerCase(),
    cleanText(row.rating_raw || row.ratingRaw),
  ];
  return parts.every(Boolean) ? parts.join("::") : "";
}

function mapStagedReviewPayload(row, importRunId) {
  return {
    property_id: row.property_id,
    import_run_id: importRunId,
    source: row.source,
    source_review_id: row.source_review_id,
    source_reservation_id: row.source_reservation_id,
    review_date: row.review_date,
    reviewer_name: row.reviewer_name,
    reviewer_country: row.reviewer_country,
    language: row.language,
    rating_raw: row.rating_raw,
    rating_scale: row.rating_scale,
    rating_normalized_100: row.rating_normalized_100,
    title: row.title,
    positive_review_text: row.positive_review_text,
    negative_review_text: row.negative_review_text,
    body: row.body,
    subscores: row.subscores,
    host_reply_text: row.host_reply_text,
    host_reply_date: row.host_reply_date,
    raw_text: row.raw_text,
    parse_confidence: row.parse_confidence,
    dedupe_fingerprint: row.dedupe_fingerprint,
    raw_payload: row.raw_payload,
  };
}

async function createImportRun({ propertyId, source, fileName, fileType, uploadKind, createdBy, rowCountDetected, errorSummary }) {
  const runId = randomUUID();
  const created = await restQuery("review_import_runs?select=*", {
    method: "POST",
    body: [{
      id: runId,
      property_id: propertyId,
      source,
      file_name: fileName,
      file_type: fileType,
      upload_kind: uploadKind,
      status: "parsed",
      row_count_detected: rowCountDetected,
      row_count_imported: 0,
      error_summary: errorSummary || "",
      created_by: createdBy || null,
    }],
  });
  if (Array.isArray(created) && created[0]) return created[0];
  return loadRun(runId);
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "reviews");

    if (req.method === "GET") {
      const id = cleanText(req.query?.id);
      if (id) {
        const run = await loadRun(id);
        if (!run) {
          res.status(404).json({ error: "Import run not found." });
          return;
        }
        const rows = await loadStagingRows(id);
        res.status(200).json({ run, rows });
        return;
      }
      const runs = await restQuery("review_import_runs?select=*,properties(name)&order=created_at.desc&limit=20", { method: "GET" });
      res.status(200).json({ rows: Array.isArray(runs) ? runs : [] });
      return;
    }

    if (req.method === "POST") {
      const action = cleanText(req.query?.action || "stage").toLowerCase();
      const body = await parseBody(req);

      if (action === "stage") {
        const user = await verifyUser(req);
        const propertyId = cleanText(body?.propertyId);
        const source = cleanText(body?.source);
        const rowsInput = Array.isArray(body?.rows) ? body.rows : [];
        if (!propertyId || !source) {
          res.status(400).json({ error: "propertyId and source are required." });
          return;
        }
        if (rowsInput.length === 0) {
          res.status(400).json({ error: "No rows were provided for staging." });
          return;
        }

        const run = await createImportRun({
          propertyId,
          source,
          fileName: cleanText(body?.fileName),
          fileType: cleanText(body?.fileType),
          uploadKind: cleanText(body?.uploadKind) || "csv",
          createdBy: user.id,
          rowCountDetected: rowsInput.length,
          errorSummary: cleanText(body?.errorSummary),
        });
        if (!run?.id) {
          const err = new Error("Could not create the import run.");
          err.statusCode = 500;
          throw err;
        }

        const payload = rowsInput.map((row) => {
          const safe = sanitizeReviewInput(row, { propertyId, source, importRunId: run.id });
          return {
            import_run_id: run.id,
            property_id: safe.property_id,
            source: safe.source,
            source_review_id: safe.source_review_id,
            source_reservation_id: safe.source_reservation_id,
            review_date: safe.review_date,
            reviewer_name: safe.reviewer_name,
            reviewer_country: safe.reviewer_country,
            language: safe.language,
            rating_raw: safe.rating_raw,
            rating_scale: safe.rating_scale,
            rating_normalized_100: safe.rating_normalized_100,
            title: safe.title,
            positive_review_text: safe.positive_review_text,
            negative_review_text: safe.negative_review_text,
            body: safe.body,
            subscores: safe.subscores,
            host_reply_text: safe.host_reply_text,
            host_reply_date: safe.host_reply_date,
            raw_text: safe.raw_text,
            parse_confidence: safe.parse_confidence,
            dedupe_fingerprint: safe.dedupe_fingerprint,
            warning_flags: safe.warning_flags,
            raw_payload: safe.raw_payload,
            selected_for_import: safe.selected_for_import,
            is_valid: safe.is_valid,
          };
        });

        const staged = await restQuery("review_import_staging?select=*", {
          method: "POST",
          body: payload,
        });
        res.status(200).json({ run, rows: Array.isArray(staged) ? staged : [] });
        return;
      }

      if (action === "confirm") {
        const importRunId = cleanText(body?.importRunId);
        if (!importRunId) {
          res.status(400).json({ error: "importRunId is required." });
          return;
        }
        const rowIds = Array.isArray(body?.rowIds) ? body.rowIds.map((id) => cleanText(id)).filter(Boolean) : [];
        const stagedRows = await loadStagingRows(importRunId);
        const selected = stagedRows.filter((row) => row.is_valid && row.selected_for_import && (!rowIds.length || rowIds.includes(row.id)));
        if (selected.length === 0) {
          res.status(400).json({ error: "No valid staged rows selected for import." });
          return;
        }

        const existing = await restQuery("reviews?select=id,source,source_review_id,dedupe_fingerprint", { method: "GET" });
        const existingBySourceId = new Map();
        const existingByFingerprint = new Map();
        const existingByStableKey = new Map();
        (Array.isArray(existing) ? existing : []).forEach((row) => {
          const sourceReviewId = cleanText(row.source_review_id);
          const fingerprint = cleanText(row.dedupe_fingerprint);
          if (sourceReviewId) existingBySourceId.set(`${cleanText(row.source).toLowerCase()}::${sourceReviewId}`, row.id);
          if (fingerprint) existingByFingerprint.set(fingerprint, row.id);
        });

        let insertedCount = 0;
        let replacedCount = 0;
        for (const row of selected) {
          const sourceReviewKey = row.source_review_id ? `${cleanText(row.source).toLowerCase()}::${cleanText(row.source_review_id)}` : "";
          const fingerprint = cleanText(row.dedupe_fingerprint);
          const payload = mapStagedReviewPayload(row, importRunId);
          const stableKey = stableReviewDuplicateKey(payload);
          const existingId =
            (sourceReviewKey && existingBySourceId.get(sourceReviewKey)) ||
            (sourceReviewKey && (await findReviewIdByPayload(payload))) ||
            (stableKey && existingByStableKey.get(stableKey)) ||
            (stableKey && (await findReviewIdByPayload(payload))) ||
            (fingerprint && existingByFingerprint.get(fingerprint)) ||
            "";
          if (existingId) {
            await restQuery(`reviews?id=eq.${encodeURIComponent(existingId)}`, {
              method: "PATCH",
              body: payload,
            });
            replacedCount += 1;
          } else {
            try {
              await restQuery("reviews", {
                method: "POST",
                body: payload,
              });
              insertedCount += 1;
              const insertedId = await findReviewIdByPayload(payload);
              if (insertedId && sourceReviewKey) existingBySourceId.set(sourceReviewKey, insertedId);
              if (insertedId && stableKey) existingByStableKey.set(stableKey, insertedId);
              if (insertedId && fingerprint) existingByFingerprint.set(fingerprint, insertedId);
            } catch (error) {
              if (!isDuplicateReviewError(error)) throw error;
              const duplicateId = await findReviewIdByPayload(payload);
              if (!duplicateId) throw error;
              await restQuery(`reviews?id=eq.${encodeURIComponent(duplicateId)}`, {
                method: "PATCH",
                body: payload,
              });
              replacedCount += 1;
              if (sourceReviewKey) existingBySourceId.set(sourceReviewKey, duplicateId);
              if (stableKey) existingByStableKey.set(stableKey, duplicateId);
              if (fingerprint) existingByFingerprint.set(fingerprint, duplicateId);
            }
          }
        }

        const importedCount = selected.length;
        const skippedCount = 0;

        await restQuery(`review_import_runs?id=eq.${encodeURIComponent(importRunId)}`, {
          method: "PATCH",
          body: {
            status: "imported",
            row_count_imported: importedCount,
            error_summary: replacedCount ? `${replacedCount} duplicate review${replacedCount === 1 ? " was" : "s were"} replaced.` : "",
          },
        });

        const run = await loadRun(importRunId);
        res.status(200).json({
          run,
          importedCount,
          insertedCount,
          replacedCount,
          skippedCount,
        });
        return;
      }

      res.status(400).json({ error: "Unsupported action." });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanText(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Missing id query parameter." });
        return;
      }
      const body = await parseBody(req);
      const payload = sanitizeReviewInput(body);
      const updated = await restQuery(`review_import_staging?id=eq.${encodeURIComponent(id)}&select=*`, {
        method: "PATCH",
        body: {
          property_id: payload.property_id,
          source: payload.source,
          source_review_id: payload.source_review_id,
          source_reservation_id: payload.source_reservation_id,
          review_date: payload.review_date,
          reviewer_name: payload.reviewer_name,
          reviewer_country: payload.reviewer_country,
          language: payload.language,
          rating_raw: payload.rating_raw,
          rating_scale: payload.rating_scale,
          rating_normalized_100: payload.rating_normalized_100,
          title: payload.title,
          positive_review_text: payload.positive_review_text,
          negative_review_text: payload.negative_review_text,
          body: payload.body,
          subscores: payload.subscores,
          host_reply_text: payload.host_reply_text,
          host_reply_date: payload.host_reply_date,
          raw_text: payload.raw_text,
          parse_confidence: payload.parse_confidence,
          dedupe_fingerprint: payload.dedupe_fingerprint,
          warning_flags: payload.warning_flags,
          raw_payload: payload.raw_payload,
          selected_for_import: payload.selected_for_import,
          is_valid: payload.is_valid,
        },
      });
      res.status(200).json({ row: Array.isArray(updated) && updated[0] ? updated[0] : null });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
