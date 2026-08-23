const { cleanText, requireFeature, restQuery, sendError } = require("./_supabase");
const { mapRpcError } = require("./_reconciliation");
const {
  ADYEN_MONTHLY_RULE_KEY,
  ADYEN_MONTHLY_RULE_VERSION,
  BANK_RESERVATION_RULE_KEY,
  BANK_RESERVATION_RULE_VERSION,
  MONTHLY_INCOME_RULE_KEY,
  toAutomationPublicResult,
} = require("./_reconciliation-automation");

const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const MEMBER_ROLES = new Set(["source", "destination"]);

function statusError(message, statusCode) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

function inputError(message) {
  return statusError(message, 400);
}

function safePublicError(error) {
  const mapped = mapRpcError(error);
  if (!error?.supabasePayload && mapped.statusCode < 500) return mapped;
  const safe = new Error(mapped.statusCode === 409
    ? "The reconciliation automation state changed. Refresh and try again."
    : mapped.statusCode >= 500
      ? "Unexpected server error."
      : "Reconciliation automation request could not be completed.");
  safe.statusCode = mapped.statusCode;
  return safe;
}

async function requireManagedFeature(req) {
  const auth = await requireFeature(req, "app", "financial-reconciliation");
  if (!cleanText(auth.access?.profile?.id)) {
    throw statusError("You do not have permission for this feature.", 403);
  }
  return auth;
}

function actorFor(auth) {
  const actor = cleanText(auth.user?.email) || cleanText(auth.user?.id);
  if (!actor) throw statusError("Authenticated user identity is unavailable.", 403);
  return actor;
}

function normalizeUuid(value, label) {
  if (typeof value !== "string" || !UUID_PATTERN.test(value)) {
    throw inputError(`${label} must be a valid UUID.`);
  }
  return value.toLowerCase();
}

function normalizeInteger(value, label, minimum, maximum = Number.MAX_SAFE_INTEGER) {
  if (typeof value !== "string" || !/^(?:0|[1-9]\d*)$/.test(value)) {
    throw inputError(`${label} must be between ${minimum} and ${maximum}.`);
  }
  const normalized = Number(value);
  if (!Number.isSafeInteger(normalized) || normalized < minimum || normalized > maximum) {
    throw inputError(`${label} must be between ${minimum} and ${maximum}.`);
  }
  return normalized;
}

function isPlainRecord(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function isPagedRulePair(ruleKey, ruleVersion) {
  return (ruleKey === MONTHLY_INCOME_RULE_KEY && (ruleVersion === 1 || ruleVersion === 2))
    || (ruleKey === BANK_RESERVATION_RULE_KEY && ruleVersion === BANK_RESERVATION_RULE_VERSION)
    || (ruleKey === ADYEN_MONTHLY_RULE_KEY && ruleVersion === ADYEN_MONTHLY_RULE_VERSION);
}

function isDecimal(value) {
  return (typeof value === "number" && Number.isFinite(value))
    || (typeof value === "string" && /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(value));
}

function pickOwn(record, key, fallbackKey) {
  if (Object.hasOwn(record, key)) return record[key];
  return fallbackKey && Object.hasOwn(record, fallbackKey) ? record[fallbackKey] : undefined;
}

function safeOptionalString(value) {
  if (value === undefined || value === null || typeof value === "string") return value;
  throw new Error("Unexpected reconciliation member page response.");
}

function safeOptionalDecimal(value) {
  if (value === undefined || value === null || isDecimal(value)) return value;
  throw new Error("Unexpected reconciliation member page response.");
}

function safeOptionalCount(value) {
  if (value === undefined || value === null
    || (Number.isSafeInteger(value) && value >= 0)) return value;
  throw new Error("Unexpected reconciliation member page response.");
}

function compactPublicFields(fields) {
  return Object.fromEntries(Object.entries(fields).filter(([, nested]) => nested !== undefined));
}

function shapeSummary(ruleKey, summary) {
  if (ruleKey === BANK_RESERVATION_RULE_KEY) {
    return compactPublicFields({
      classification: safeOptionalString(pickOwn(summary, "classification")),
      reason: safeOptionalString(pickOwn(summary, "reason")),
      candidateCount: safeOptionalCount(pickOwn(summary, "candidateCount", "candidate_count")),
      bankAnchorDate: safeOptionalString(pickOwn(summary, "bankAnchorDate", "bank_anchor_date")),
      sourceCount: safeOptionalCount(pickOwn(summary, "sourceCount", "source_count")),
      sourceTotal: safeOptionalDecimal(pickOwn(summary, "sourceTotal", "source_total")),
      destinationCount: safeOptionalCount(pickOwn(summary, "destinationCount", "destination_count")),
      destinationTotal: safeOptionalDecimal(pickOwn(summary, "destinationTotal", "destination_total")),
    });
  }
  return compactPublicFields({
    calendarMonth: safeOptionalString(pickOwn(summary, "calendarMonth", "calendar_month")),
    sourceCount: safeOptionalCount(pickOwn(summary, "sourceCount", "source_count")),
    sourceTotal: safeOptionalDecimal(pickOwn(summary, "sourceTotal", "source_total")),
    destinationCount: safeOptionalCount(pickOwn(summary, "destinationCount", "destination_count")),
    destinationTotal: safeOptionalDecimal(pickOwn(summary, "destinationTotal", "destination_total")),
  });
}

function shapeRowSnapshot(snapshot) {
  return compactPublicFields({
    id: safeOptionalString(pickOwn(snapshot, "id")),
    importBatch: safeOptionalString(pickOwn(snapshot, "importBatch", "import_batch")),
    rowKey: safeOptionalString(pickOwn(snapshot, "rowKey", "row_key")),
    date: safeOptionalString(pickOwn(snapshot, "date", "data")),
    description: safeOptionalString(pickOwn(snapshot, "description", "descritivo")),
    amount: safeOptionalDecimal(pickOwn(snapshot, "amount", "montante")),
    account: safeOptionalString(pickOwn(snapshot, "account", "conta")),
    sourceType: safeOptionalString(pickOwn(snapshot, "sourceType", "source_type")),
  });
}

function requireMemberPage(value, expected) {
  const page = toAutomationPublicResult(value);
  if (!isPlainRecord(page)
    || page.runId !== expected.runId
    || page.proposalId !== expected.proposalId
    || page.role !== expected.role
    || page.offset !== expected.offset
    || page.limit !== expected.limit
    || !isPagedRulePair(page.ruleKey, page.ruleVersion)
    || typeof page.groupingKey !== "string" || !page.groupingKey
    || !isPlainRecord(page.summarySnapshot)
    || !Number.isSafeInteger(page.sourceCount) || page.sourceCount < 0
    || !isDecimal(page.sourceTotal)
    || !Number.isSafeInteger(page.destinationCount) || page.destinationCount < 0
    || !isDecimal(page.destinationTotal)
    || !Number.isSafeInteger(page.totalCount) || page.totalCount < 0
    || page.totalCount !== (expected.role === "source" ? page.sourceCount : page.destinationCount)
    || !Array.isArray(page.members)
    || page.members.length > expected.limit
    || (page.members.length === 0
      ? expected.offset < page.totalCount
      : expected.offset + page.members.length > page.totalCount)) {
    throw new Error("Unexpected reconciliation member page response.");
  }

  const sourceIds = new Set();
  const members = page.members.map((member, index) => {
    if (!isPlainRecord(member)
      || member.role !== expected.role
      || typeof member.sourceType !== "string" || !member.sourceType
      || typeof member.sourceId !== "string" || !UUID_PATTERN.test(member.sourceId)
      || member.ordinal !== expected.offset + index + 1
      || typeof member.sourceDate !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(member.sourceDate)
      || !isDecimal(member.amount)
      || (member.description !== null && typeof member.description !== "string")
      || (member.account !== null && typeof member.account !== "string")
      || !isPlainRecord(member.rowSnapshot)
      || sourceIds.has(member.sourceId)) {
      throw new Error("Unexpected reconciliation member page response.");
    }
    sourceIds.add(member.sourceId);
    return {
      role: member.role,
      sourceType: member.sourceType,
      sourceId: member.sourceId,
      ordinal: member.ordinal,
      sourceDate: member.sourceDate,
      amount: member.amount,
      description: member.description,
      account: member.account,
      rowSnapshot: shapeRowSnapshot(member.rowSnapshot),
    };
  });

  return {
    runId: page.runId,
    proposalId: page.proposalId,
    ruleKey: page.ruleKey,
    ruleVersion: page.ruleVersion,
    groupingKey: page.groupingKey,
    summarySnapshot: shapeSummary(page.ruleKey, page.summarySnapshot),
    sourceCount: page.sourceCount,
    sourceTotal: page.sourceTotal,
    destinationCount: page.destinationCount,
    destinationTotal: page.destinationTotal,
    role: page.role,
    offset: page.offset,
    limit: page.limit,
    totalCount: page.totalCount,
    members,
  };
}

module.exports = async function handler(req, res) {
  try {
    const auth = await requireManagedFeature(req);
    if (req.method !== "GET") {
      res.setHeader("Allow", "GET");
      return res.status(405).json({ error: "Method not allowed." });
    }

    const runId = normalizeUuid(req.query?.run_id, "Run ID");
    const proposalId = normalizeUuid(req.query?.proposal_id, "Proposal ID");
    const role = req.query?.role;
    if (typeof role !== "string" || !MEMBER_ROLES.has(role)) {
      throw inputError("Member role must be source or destination.");
    }
    const offset = normalizeInteger(req.query?.offset, "Offset", 0);
    const limit = normalizeInteger(req.query?.limit, "Limit", 1, 50);
    const actor = actorFor(auth);
    const result = await restQuery(
      "rpc/get_financial_reconciliation_automatic_proposal_members",
      {
        method: "POST",
        body: {
          p_run_id: runId,
          p_proposal_id: proposalId,
          p_role: role,
          p_offset: offset,
          p_limit: limit,
          p_actor: actor,
        },
      },
    );
    return res.status(200).json(requireMemberPage(result, {
      runId,
      proposalId,
      role,
      offset,
      limit,
    }));
  } catch (error) {
    return sendError(res, safePublicError(error));
  }
};
