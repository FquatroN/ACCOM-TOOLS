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

async function uploadOpenAiFile(file) {
  const apiKey = cleanText(process.env.OPENAI_API_KEY);
  if (!apiKey) {
    const error = new Error("Missing server environment variable: OPENAI_API_KEY");
    error.statusCode = 500;
    throw error;
  }
  const form = new FormData();
  form.append("purpose", "user_data");
  form.append("file", new Blob([Buffer.from(String(file.base64Content || ""), "base64")], { type: file.mimeType || "application/pdf" }), file.originalFilename || "document.pdf");
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

async function parseFinancialDocument(file) {
  const apiKey = cleanText(process.env.OPENAI_API_KEY);
  if (!apiKey) {
    const error = new Error("Missing server environment variable: OPENAI_API_KEY");
    error.statusCode = 500;
    throw error;
  }
  const prompt = [
    "Analyze this financial document and extract likely fields.",
    "Return only JSON with these keys:",
    "documentDate, docNumber, description, supplierNif, supplierName, amount, vatAmount, notes.",
    "Use ISO date format YYYY-MM-DD when possible.",
    "Use numbers for amount and vatAmount when possible, otherwise null.",
    "If a field is unknown, return an empty string or null.",
    "Description should be a short practical description of the expense/income document.",
  ].join(" ");
  const content = [{ type: "input_text", text: prompt }];
  if ((file.mimeType || "").startsWith("image/")) {
    const dataUrl = `data:${file.mimeType};base64,${file.base64Content}`;
    content.push({ type: "input_image", image_url: dataUrl });
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
      reasoning: { effort: "minimal" },
      input: [{ role: "user", content }],
    }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload?.error?.message || payload?.message || `OpenAI parse failed (${response.status})`);
    error.statusCode = response.status;
    throw error;
  }
  const text = extractResponseText(payload);
  const parsed = parseJsonText(text);
  return { parsed, rawText: text };
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "financial-docs");
    if (req.method !== "POST") {
      res.status(405).json({ error: "Method not allowed." });
      return;
    }
    const body = await parseBody(req);
    const file = body?.file && typeof body.file === "object" ? body.file : {};
    if (!cleanText(file.base64Content)) {
      res.status(400).json({ error: "A document file is required." });
      return;
    }
    const { parsed, rawText } = await parseFinancialDocument(file);
    const row = {
      cc: "",
      documentDate: cleanText(parsed.documentDate),
      docNumber: cleanText(parsed.docNumber),
      description: cleanText(parsed.description),
      supplierNif: cleanText(parsed.supplierNif).slice(0, 15),
      supplierName: cleanText(parsed.supplierName),
      amount: parsed.amount === null || parsed.amount === undefined || cleanText(parsed.amount) === "" ? "" : Number(parsed.amount),
      vatAmount: parsed.vatAmount === null || parsed.vatAmount === undefined || cleanText(parsed.vatAmount) === "" ? "" : Number(parsed.vatAmount),
      payment: "",
      docType: "",
      fat: "",
      category: "",
      status: "Draft",
    };
    res.status(200).json({
      row,
      ocrFields: parsed,
      ocrRawText: rawText,
      notes: cleanText(parsed.notes),
    });
  } catch (error) {
    sendError(res, error);
  }
};
