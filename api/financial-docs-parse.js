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

function normalizeWhitespace(value) {
  return cleanText(String(value || "").replace(/\s+/g, " "));
}

function isUtilitySupplier(name) {
  const value = normalizeWhitespace(name).toUpperCase();
  return value.includes("EDP") || value.includes("EPAL");
}

function cleanUtilityAddressPart(value) {
  return normalizeWhitespace(value)
    .replace(/\b(andar|piso)\b/gi, "")
    .replace(/\b(frac(?:c|ç)?(?:ao|ão)?|porta)\b/gi, "")
    .replace(/\s*[,;]\s*/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function extractUtilityFloorDoor(text) {
  const raw = normalizeWhitespace(text);
  if (!raw) return "";
  const floorDoorPatterns = [
    /\b(\d+\s*(?:esq|dto|dt|frt|frente|tras|traseiras?|a|b|c))\b/i,
    /\b(rc\s*(?:esq|dto|dt|frt|frente|tras|traseiras?)?)\b/i,
    /\b(c\/?v\s*(?:esq|dto|dt|frt|frente|tras|traseiras?)?)\b/i,
    /\b(loja\s*[a-z0-9-]*)\b/i,
  ];
  for (const pattern of floorDoorPatterns) {
    const match = raw.match(pattern);
    if (match?.[1]) return cleanUtilityAddressPart(match[1]);
  }
  return "";
}

function normalizeUtilityStreetAddress(value) {
  const raw = normalizeWhitespace(value);
  if (!raw) return "";
  const primary = cleanUtilityAddressPart(raw.split(/\s*-\s*/)[0] || raw);
  const streetMatch = primary.match(/^(.*?\b\d+[A-Z0-9/-]*)\b/i);
  const street = cleanUtilityAddressPart(streetMatch?.[1] || primary);
  const floorDoor = extractUtilityFloorDoor(raw);
  return normalizeWhitespace([street, floorDoor].filter(Boolean).join(" "));
}

function buildPortugueseDescription(parsed) {
  const baseDescription = normalizeWhitespace(parsed?.description);
  const serviceAddressShort = normalizeUtilityStreetAddress(parsed?.serviceAddressShort || parsed?.service_address_short);
  if (!isUtilitySupplier(parsed?.supplierName)) return baseDescription;
  if (!serviceAddressShort) return baseDescription;
  const normalizedBase = baseDescription.toLowerCase();
  const normalizedAddress = serviceAddressShort.toLowerCase();
  if (!baseDescription) return serviceAddressShort;
  if (normalizedBase.startsWith(normalizedAddress)) return baseDescription;
  return `${serviceAddressShort} - ${baseDescription}`;
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
    "documentDate, docNumber, description, supplierNif, supplierName, amount, vatAmount, serviceAddressShort, notes.",
    "Use ISO date format YYYY-MM-DD when possible.",
    "Use numbers for amount and vatAmount when possible, otherwise null.",
    "If a field is unknown, return an empty string or null.",
    "Description must always be written in Portuguese (Portugal).",
    "Description should be a short practical description of the expense/income document.",
    "If the supplier is EDP or EPAL, identify the supply address and return serviceAddressShort using only street name with number(s), plus floor and door when available.",
    "For EDP or EPAL, the description should begin with that short supply address when available.",
    "Do not include city, postcode, country, or extra address lines in serviceAddressShort.",
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
      description: buildPortugueseDescription(parsed),
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
