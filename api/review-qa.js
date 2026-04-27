const {
  cleanText,
  parseBody,
  requireFeature,
  sendError,
} = require("./_supabase");

const DEFAULT_MODEL = process.env.OPENAI_MODEL || "gpt-5";
const MAX_ROWS = 250;

function defaultReasoningEffort(model) {
  const raw = cleanText(model).toLowerCase();
  if (raw.startsWith("gpt-5.4")) return "none";
  return "minimal";
}

function reasoningEffortForModel(model) {
  const requested = cleanText(process.env.OPENAI_REASONING_EFFORT);
  return requested || defaultReasoningEffort(model);
}

function summarizeFilters(filters) {
  const parts = [];
  const propertyId = cleanText(filters?.propertyId);
  const source = cleanText(filters?.source);
  const dateFrom = cleanText(filters?.dateFrom);
  const dateTo = cleanText(filters?.dateTo);
  const scoreFrom = cleanText(filters?.scoreFrom);
  const scoreTo = cleanText(filters?.scoreTo);
  const search = cleanText(filters?.search);
  if (propertyId) parts.push(`property filter active (${propertyId})`);
  if (source) parts.push(`source filter active (${source})`);
  if (dateFrom || dateTo) parts.push(`date range ${dateFrom || "start"} to ${dateTo || "end"}`);
  if (scoreFrom || scoreTo) parts.push(`score range ${scoreFrom || "0"} to ${scoreTo || "100"}`);
  if (search) parts.push(`search text "${search}"`);
  return parts.length ? parts.join("; ") : "no extra filters besides the current dashboard scope";
}

function normalizeRows(value) {
  const rows = Array.isArray(value) ? value : [];
  return rows.slice(0, MAX_ROWS).map((row) => ({
    id: cleanText(row?.id),
    reviewDate: cleanText(row?.reviewDate),
    property: cleanText(row?.property),
    source: cleanText(row?.source),
    reviewerName: cleanText(row?.reviewerName),
    ratingNormalized100: row?.ratingNormalized100 ?? null,
    ratingRaw: row?.ratingRaw ?? null,
    ratingScale: row?.ratingScale ?? null,
    title: cleanText(row?.title),
    positiveReviewText: cleanText(row?.positiveReviewText),
    negativeReviewText: cleanText(row?.negativeReviewText),
    body: cleanText(row?.body),
    hostReplyText: cleanText(row?.hostReplyText),
    subscores: row?.subscores && typeof row.subscores === "object" ? row.subscores : {},
  }));
}

function extractResponseText(payload) {
  const direct = cleanText(payload?.output_text);
  if (direct) return direct;

  const parts = [];
  const visit = (value) => {
    if (!value) return;
    if (typeof value === "string") return;
    if (Array.isArray(value)) {
      value.forEach(visit);
      return;
    }
    if (typeof value !== "object") return;

    const type = cleanText(value.type);
    if ((type === "output_text" || type === "text") && cleanText(value.text)) {
      parts.push(cleanText(value.text));
    }
    if (cleanText(value.output_text)) parts.push(cleanText(value.output_text));
    if (value.content) visit(value.content);
    if (value.output) visit(value.output);
  };

  visit(payload?.output);
  return parts.join("\n\n").trim();
}

async function askOpenAI({ question, filters, rows, totalCount }) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    const err = new Error("Missing server environment variable: OPENAI_API_KEY");
    err.statusCode = 500;
    throw err;
  }

  const instructions = [
    "You are analyzing accommodation reviews for internal property management.",
    "Answer only from the provided reviews.",
    "If the evidence is limited, say so clearly.",
    "Be concise but useful.",
    "Cite supporting review snippets in paraphrased form with dates/source/property when helpful.",
    "End with 2-4 practical takeaways when appropriate.",
  ].join(" ");

  const dataset = {
    question,
    scope_summary: summarizeFilters(filters),
    total_filtered_reviews: Number(totalCount || rows.length),
    analyzed_reviews: rows.length,
    reviews: rows,
  };

  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: DEFAULT_MODEL,
      reasoning: { effort: reasoningEffortForModel(DEFAULT_MODEL) },
      instructions,
      input: [
        {
          role: "user",
          content: [
            {
              type: "input_text",
              text: `Analyze this review dataset and answer the user's question.\n\n${JSON.stringify(dataset)}`,
            },
          ],
        },
      ],
    }),
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const message = payload?.error?.message || payload?.message || `OpenAI request failed (${response.status})`;
    const err = new Error(message);
    err.statusCode = response.status;
    throw err;
  }

  const answer = extractResponseText(payload);
  if (!answer) {
    const err = new Error("OpenAI returned no answer text. Please try again with a shorter or more specific question.");
    err.statusCode = 502;
    throw err;
  }
  return answer;
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "reviews");
    if (req.method !== "POST") {
      res.status(405).json({ error: "Method not allowed." });
      return;
    }
    const body = await parseBody(req);
    const question = cleanText(body?.question);
    const rows = normalizeRows(body?.rows);
    const totalCount = Number(body?.totalCount || rows.length);
    if (!question) {
      res.status(400).json({ error: "Question is required." });
      return;
    }
    if (rows.length === 0) {
      res.status(400).json({ error: "At least one review is required for analysis." });
      return;
    }

    const answer = await askOpenAI({
      question,
      filters: body?.filters || {},
      rows,
      totalCount,
    });

    res.status(200).json({
      answer,
      analyzedCount: rows.length,
      totalCount,
      note: rows.length < totalCount ? `Analysis based on the most recent ${rows.length} filtered reviews.` : "Analysis based on the full filtered review set.",
    });
  } catch (error) {
    sendError(res, error);
  }
};
