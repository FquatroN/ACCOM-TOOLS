const { cleanText, parseBody, requireFeature, sendError } = require("./_supabase");

const OPENAI_MODEL = cleanText(process.env.OPENAI_MODEL) || "gpt-5";

function extractResponseText(payload) {
  const direct = cleanText(payload?.output_text);
  if (direct) return direct;
  const parts = [];
  const visit = (value) => {
    if (!value) return;
    if (Array.isArray(value)) {
      value.forEach(visit);
      return;
    }
    if (typeof value !== "object") return;
    if (cleanText(value.text)) parts.push(cleanText(value.text));
    if (cleanText(value.output_text)) parts.push(cleanText(value.output_text));
    if (value.content) visit(value.content);
    if (value.output) visit(value.output);
  };
  visit(payload?.output);
  return parts.join("\n\n").trim();
}

function parseJsonText(text) {
  const raw = cleanText(text);
  if (!raw) return {};
  const fenced = raw.replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/```$/i, "").trim();
  try {
    return JSON.parse(fenced);
  } catch {
    const firstBrace = fenced.indexOf("{");
    const lastBrace = fenced.lastIndexOf("}");
    if (firstBrace >= 0 && lastBrace > firstBrace) {
      try {
        return JSON.parse(fenced.slice(firstBrace, lastBrace + 1));
      } catch {}
    }
  }
  return {};
}

function normalizeUploadFilename(value, fallback = "document.pdf") {
  const raw = cleanText(value) || fallback;
  return raw.replace(/(\.[a-z0-9]{1,8})$/i, (match) => match.toLowerCase());
}

async function uploadOpenAiFile(file) {
  const apiKey = cleanText(process.env.OPENAI_API_KEY);
  if (!apiKey) {
    const error = new Error("Missing server environment variable: OPENAI_API_KEY");
    error.statusCode = 500;
    throw error;
  }
  const form = new FormData();
  form.append("purpose", "user_data");
  form.append(
    "file",
    new Blob([Buffer.from(String(file.base64Content || ""), "base64")], { type: file.mimeType || "application/pdf" }),
    normalizeUploadFilename(file.originalFilename, "document.pdf")
  );
  const response = await fetch("https://api.openai.com/v1/files", {
    method: "POST",
    headers: { Authorization: `Bearer ${apiKey}` },
    body: form,
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload?.error?.message || payload?.message || `OpenAI file upload failed (${response.status})`);
    error.statusCode = response.status;
    throw error;
  }
  return cleanText(payload.id);
}

async function parseCgdCartaoCredito(file) {
  const apiKey = cleanText(process.env.OPENAI_API_KEY);
  if (!apiKey) {
    const error = new Error("Missing server environment variable: OPENAI_API_KEY");
    error.statusCode = 500;
    throw error;
  }
  const prompt = [
    "Extract transaction rows from this CGD Cartao Credito credit-card statement.",
    "Return only JSON with key rows, where rows is an array.",
    "Each row must have: data, dataValor, descricao, debito, credito.",
    "Only include rows where data, dataValor, descricao, and at least one of debito or credito are visible.",
    "Do not include opening balance, closing balance, totals, headers, page footers, or summary lines.",
    "Use the exact transaction description text as descricao.",
    "Use ISO date YYYY-MM-DD when possible; otherwise keep the visible date text.",
    "Use numeric values for debito and credito when possible; use null for blank values.",
    "Preserve debit values as positive numbers in debito, and credit values as positive numbers in credito.",
  ].join(" ");
  const content = [{ type: "input_text", text: prompt }];
  if ((file.mimeType || "").startsWith("image/")) {
    content.push({ type: "input_image", image_url: `data:${file.mimeType};base64,${file.base64Content}` });
  } else {
    const fileId = await uploadOpenAiFile(file);
    content.push({ type: "input_file", file_id: fileId });
  }
  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: OPENAI_MODEL,
      reasoning: { effort: "low" },
      input: [{ role: "user", content }],
    }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload?.error?.message || payload?.message || `OpenAI parse failed (${response.status})`);
    error.statusCode = response.status;
    throw error;
  }
  const rawText = extractResponseText(payload);
  const parsed = parseJsonText(rawText);
  const rows = Array.isArray(parsed?.rows) ? parsed.rows : [];
  return { rows, rawText };
}

async function handler(req, res) {
  try {
    await requireFeature(req, "app", "import-data");
    if (req.method !== "POST") {
      res.status(405).json({ error: "Method not allowed." });
      return;
    }
    const body = await parseBody(req);
    const type = cleanText(body?.type).toLowerCase();
    if (type !== "cgd-cartao-credito") {
      res.status(400).json({ error: "Unsupported import parse type." });
      return;
    }
    const file = body?.file && typeof body.file === "object" ? body.file : {};
    if (!cleanText(file.base64Content)) {
      res.status(400).json({ error: "A document file is required." });
      return;
    }
    const result = await parseCgdCartaoCredito(file);
    res.status(200).json({ type, ...result });
  } catch (error) {
    sendError(res, error);
  }
}

module.exports = handler;
