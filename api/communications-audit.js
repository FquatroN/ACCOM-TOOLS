const {
  cleanText,
  requireFeature,
  restQuery,
  sendError,
} = require("./_supabase");

function disallowMethod(res) {
  res.status(405).json({ error: "Method not allowed." });
}

function sanitizeAuditEvent(row) {
  return {
    id: cleanText(row?.id),
    action: cleanText(row?.action),
    communicationId: cleanText(row?.communication_id),
    actorUserId: cleanText(row?.actor_user_id),
    actorEmail: cleanText(row?.actor_email),
    occurredAt: cleanText(row?.occurred_at),
    recordSnapshot: row?.record_snapshot && typeof row.record_snapshot === "object" && !Array.isArray(row.record_snapshot)
      ? row.record_snapshot
      : {},
  };
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "settings", "communications");

    if (req.method !== "GET") {
      disallowMethod(res);
      return;
    }

    const rows = await restQuery(
      "communications_audit_events?select=id,action,communication_id,actor_user_id,actor_email,occurred_at,record_snapshot&order=occurred_at.desc&limit=200",
      { method: "GET" }
    );
    res.status(200).json({ rows: Array.isArray(rows) ? rows.map(sanitizeAuditEvent) : [] });
  } catch (error) {
    sendError(res, error);
  }
};
