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

function normalizeUploadFilename(value, fallback = "document.pdf") {
  const raw = cleanText(value) || fallback;
  return raw.replace(/(\.[a-z0-9]{1,8})$/i, (match) => match.toLowerCase());
}

function isUtilitySupplier(name) {
  const value = normalizeWhitespace(name).toUpperCase();
  return value.includes("EDP") || value.includes("EPAL");
}

function isEpalSupplier(name) {
  return normalizeWhitespace(name).toUpperCase().includes("EPAL");
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
  const numberIndex = primary.search(/\b\d+[A-Z0-9/-]*\b/i);
  const tail = numberIndex >= 0 ? cleanUtilityAddressPart(primary.slice(numberIndex)) : "";
  const floorDoor = extractUtilityFloorDoor(raw);
  const normalizedTail = normalizeWhitespace(tail);
  if (!normalizedTail) return floorDoor;
  if (floorDoor && !normalizedTail.toLowerCase().includes(floorDoor.toLowerCase())) {
    return normalizeWhitespace(`${normalizedTail} ${floorDoor}`);
  }
  return normalizedTail;
}

function buildPortugueseDescription(parsed) {
  const baseDescription = normalizeWhitespace(parsed?.description);
  const serviceAddressShort = normalizeUtilityStreetAddress(parsed?.serviceAddressShort || parsed?.service_address_short);
  const postalAddressPrincipalShort = normalizeUtilityStreetAddress(parsed?.postalAddressPrincipalShort || parsed?.postal_address_principal_short);
  const supplyAddressShort = normalizeUtilityStreetAddress(parsed?.supplyAddressShort || parsed?.supply_address_short);
  const preferredAddress = isEpalSupplier(parsed?.supplierName)
    ? postalAddressPrincipalShort
    : (serviceAddressShort || supplyAddressShort || postalAddressPrincipalShort);
  if (!isUtilitySupplier(parsed?.supplierName)) return baseDescription;
  if (!preferredAddress) return baseDescription;
  const normalizedBase = baseDescription.toLowerCase();
  const normalizedAddress = preferredAddress.toLowerCase();
  if (!baseDescription) return preferredAddress;
  if (normalizedBase.startsWith(normalizedAddress)) return baseDescription;
  return `${preferredAddress} - ${baseDescription}`;
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
    "documentDate, docNumber, description, supplierNif, supplierName, amount, vatAmount, serviceAddressShort, postalAddressPrincipalShort, supplyAddressShort, notes.",
    "Use ISO date format YYYY-MM-DD when possible.",
    "Use numbers for amount and vatAmount when possible, otherwise null.",
    "If a field is unknown, return an empty string or null.",
    "Description must always be written in Portuguese (Portugal).",
    "Description should be a short practical description of the expense/income document.",
    "If the supplier is EDP or EPAL, extract utility addresses separately.",
    "Use postalAddressPrincipalShort for 'Morada Postal (Principal)'.",
    "Use supplyAddressShort for 'Morada Abastecimento'.",
    "Use serviceAddressShort only as a fallback short address if the wording is different.",
    "For utility addresses, return only what comes after the street name: building number(s), plus floor and door when available.",
    "For EPAL, use ONLY the address shown under 'Morada Postal (Principal)' for the description prefix. Do NOT use 'Morada Abastecimento' for the description prefix.",
    "Example: if the address is 'RUA RODRIGUES SAMPAIO 146 5 ESQ', return '146 5 ESQ'.",
    "EPAL example with both fields present: if 'Morada Postal (Principal)' is 'RUA RODRIGUES SAMPAIO 146 5 ESQ' and 'Morada Abastecimento' is 'RUA CAMILO CASTELO BRANCO 9-A 5 ESQ', then postalAddressPrincipalShort must be '146 5 ESQ' and the description prefix must also be '146 5 ESQ'.",
    "For EDP or EPAL, the description should begin with that short supply address when available.",
    "Do not include city, postcode, country, or extra address lines in any returned short address.",
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
