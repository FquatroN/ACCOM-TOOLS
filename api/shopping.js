const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  DEFAULT_SHOPPING_SETTINGS,
  buildOrderItemsFromSettings,
  cleanText,
  countOrderedItems,
  sanitizeOrderItems,
  sanitizeShoppingSettings,
} = require("./_shopping");

function todayInLisbon() {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/Lisbon",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date());
  const get = (type) => parts.find((item) => item.type === type)?.value || "";
  return `${get("year")}-${get("month")}-${get("day")}`;
}

function formatDateDisplay(value) {
  const raw = cleanText(value);
  if (!raw) return "-";
  const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) return `${match[3]}/${match[2]}/${match[1]}`;
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return raw;
  return new Intl.DateTimeFormat("pt-PT", { dateStyle: "short", timeZone: "Europe/Lisbon" }).format(date);
}

function formatDateTimeShort(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return cleanText(value);
  return new Intl.DateTimeFormat("pt-PT", {
    timeZone: "Europe/Lisbon",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

async function loadShoppingSettings() {
  const rows = await restQuery("app_settings?select=payload&setting_key=eq.shopping&limit=1", {
    method: "GET",
  });
  const payload = Array.isArray(rows) && rows[0]?.payload ? rows[0].payload : DEFAULT_SHOPPING_SETTINGS;
  return sanitizeShoppingSettings(payload);
}

function sanitizeShoppingOrderRow(row = {}) {
  const items = sanitizeOrderItems(row.items, []);
  return {
    id: cleanText(row.id),
    orderNumber: Number(row.order_number || row.orderNumber || 0) || 0,
    status: cleanText(row.status).toLowerCase() === "submitted" ? "submitted" : "open",
    createdAt: cleanText(row.created_at || row.createdAt),
    updatedAt: cleanText(row.updated_at || row.updatedAt),
    submittedAt: cleanText(row.submitted_at || row.submittedAt),
    submittedByName: cleanText(row.submitted_by_name || row.submittedByName),
    submittedByUserEmail: cleanText(row.submitted_by_user_email || row.submittedByUserEmail).toLowerCase(),
    notes: cleanText(row.notes),
    reopenedFromId: cleanText(row.reopened_from_id || row.reopenedFromId),
    items,
    orderedCount: countOrderedItems(items),
  };
}

function buildShoppingOrderFallback(existing = {}, patch = {}) {
  return sanitizeShoppingOrderRow({
    id: existing.id,
    order_number: existing.orderNumber,
    status: patch.status || existing.status,
    created_at: existing.createdAt,
    updated_at: patch.updated_at || existing.updatedAt,
    submitted_at: patch.submitted_at || existing.submittedAt,
    submitted_by_name: patch.submitted_by_name || existing.submittedByName,
    submitted_by_user_email: patch.submitted_by_user_email || existing.submittedByUserEmail,
    notes: patch.notes ?? existing.notes,
    reopened_from_id: existing.reopenedFromId,
    items: patch.items || existing.items,
  });
}

async function loadOpenOrder() {
  const rows = await restQuery(
    "shopping_orders?select=id,order_number,status,submitted_by_name,submitted_by_user_email,submitted_at,notes,reopened_from_id,items,created_at,updated_at&status=eq.open&limit=1",
    { method: "GET" }
  );
  return Array.isArray(rows) && rows[0] ? sanitizeShoppingOrderRow(rows[0]) : null;
}

async function loadSubmittedHistory() {
  const rows = await restQuery(
    "shopping_orders?select=id,order_number,status,submitted_by_name,submitted_by_user_email,submitted_at,notes,reopened_from_id,items,created_at,updated_at&status=eq.submitted&order=submitted_at.desc,updated_at.desc,order_number.desc",
    { method: "GET" }
  );
  return Array.isArray(rows) ? rows.map(sanitizeShoppingOrderRow) : [];
}

async function loadOrderById(id) {
  const rows = await restQuery(
    `shopping_orders?select=id,order_number,status,submitted_by_name,submitted_by_user_email,submitted_at,notes,reopened_from_id,items,created_at,updated_at&id=eq.${encodeURIComponent(id)}&limit=1`,
    { method: "GET" }
  );
  return Array.isArray(rows) && rows[0] ? sanitizeShoppingOrderRow(rows[0]) : null;
}

function validateSubmittableOrder(order) {
  const items = Array.isArray(order?.items) ? order.items : [];
  const orderedItems = items.filter((item) => !!item.order);
  if (!orderedItems.length) {
    const error = new Error("Select at least one shopping item before submitting.");
    error.statusCode = 400;
    throw error;
  }
  const missingQuantity = orderedItems.find((item) => item.quantityRequired && !cleanText(item.existingQuantity));
  if (missingQuantity) {
    const error = new Error(`Existing Quantity is required for "${missingQuantity.item}".`);
    error.statusCode = 400;
    throw error;
  }
}

function buildSelectedItemsTable(items, settings) {
  const rows = sortShoppingItemsServer(items.filter((item) => !!item.order))
    .map(
      (item) => `<tr style="background:${escapeHtml(blendShoppingColorOnWhiteServer(getShoppingCategoryColorServer(settings, item.category), 0.10))}">
          <td>${escapeHtml(item.category)}</td>
          <td>${escapeHtml(item.item)}</td>
          <td>${escapeHtml(item.supplier || "-")}</td>
          <td>${escapeHtml(item.existingQuantity || "-")}</td>
          <td>Yes</td>
      </tr>`
    )
    .join("");
    return `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%">
      <thead>
        <tr>
          <th align="left" style="background:#f1ece6">Category</th>
          <th align="left" style="background:#f1ece6">Item</th>
          <th align="left" style="background:#f1ece6">Supplier</th>
          <th align="left" style="background:#f1ece6">Existing quantity</th>
          <th align="left" style="background:#f1ece6">Order</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>`;
}

function shoppingOrderFileStem(order) {
  return `shopping_order_${String(order?.orderNumber || "order").padStart(4, "0")}`;
}

function shoppingOrderMetaRows(order) {
  return [
    ["Order #", String(order?.orderNumber || "-")],
    ["Order Date", formatDateTimeShort(order?.submittedAt || order?.updatedAt || order?.createdAt)],
    ["Name", order?.submittedByName || "-"],
    ["Number Items", String(order?.orderedCount || 0)],
    ["Notes", order?.notes || "-"],
  ];
}

function shoppingOrderSelectedItems(order) {
  return sortShoppingItemsServer(Array.isArray(order?.items) ? order.items.filter((item) => !!item.order) : []);
}

function compareShoppingItemsServer(a, b) {
  const categoryCompare = cleanText(a?.category).localeCompare(cleanText(b?.category), undefined, { sensitivity: "base" });
  if (categoryCompare !== 0) return categoryCompare;
  const itemCompare = cleanText(a?.item).localeCompare(cleanText(b?.item), undefined, { sensitivity: "base" });
  if (itemCompare !== 0) return itemCompare;
  return cleanText(a?.supplier).localeCompare(cleanText(b?.supplier), undefined, { sensitivity: "base" });
}

function sortShoppingItemsServer(items) {
  return [...(Array.isArray(items) ? items : [])].sort(compareShoppingItemsServer);
}

function normalizeHexColorServer(hex, fallback = "#F3E7DB") {
  const raw = cleanText(hex).toUpperCase();
  return /^#[0-9A-F]{6}$/.test(raw) ? raw : fallback;
}

function getShoppingCategoryColorServer(settings, category) {
  const configured = settings?.categoryColors?.[category] || settings?.category_colors?.[category];
  return normalizeHexColorServer(configured, "#F3E7DB");
}

function blendShoppingColorOnWhiteServer(hex, alpha = 0.15) {
  const normalized = normalizeHexColorServer(hex, "#F3E7DB").replace("#", "");
  const r = Number.parseInt(normalized.slice(0, 2), 16);
  const g = Number.parseInt(normalized.slice(2, 4), 16);
  const b = Number.parseInt(normalized.slice(4, 6), 16);
  const blend = (channel) => Math.round((255 * (1 - alpha)) + (channel * alpha));
  return `#${[blend(r), blend(g), blend(b)].map((item) => item.toString(16).padStart(2, "0")).join("")}`;
}

function buildShoppingOrderExcelHtml(order, settings) {
  const rows = shoppingOrderSelectedItems(order);
  const metaRows = shoppingOrderMetaRows(order);
  return `<!doctype html><html><head><meta charset="utf-8"><style>
    @page { size: A4 portrait; margin: 12mm; }
    body { font-family: Calibri, Arial, sans-serif; color: #222; font-size: 11px; }
    h1 { font-size: 18px; margin: 0 0 8px; color: #3b2f24; }
    table { width: 100%; border-collapse: collapse; margin-top: 10px; mso-page-orientation: portrait; }
    th, td { border: 1px solid #cfc7bf; padding: 6px 7px; vertical-align: top; text-align: left; }
    th { background: #e8ded4; font-weight: 700; }
    .meta td:first-child { width: 150px; font-weight: 700; background: #faf7f2; }
    .items th:nth-child(1) { width: 16%; }
    .items th:nth-child(2) { width: 28%; }
    .items th:nth-child(3) { width: 20%; }
    .items th:nth-child(4) { width: 16%; }
    .items th:nth-child(5) { width: 12%; }
    .items th:nth-child(6) { width: 8%; }
  </style></head><body>
    <h1>Shopping Order Detail</h1>
    <table class="meta"><tbody>${metaRows.map(([label, value]) => `<tr><td>${escapeHtml(label)}</td><td>${escapeHtml(value)}</td></tr>`).join("")}</tbody></table>
    <table class="items"><thead><tr><th>Category</th><th>Item</th><th>Supplier</th><th>Stored</th><th>Existing Quantity</th><th>Order</th></tr></thead>
    <tbody>${rows.map((item) => `<tr style="background:${escapeHtml(blendShoppingColorOnWhiteServer(getShoppingCategoryColorServer(settings, item.category), 0.15))}"><td>${escapeHtml(item.category || "-")}</td><td>${escapeHtml(item.item || "-")}</td><td>${escapeHtml(item.supplier || "-")}</td><td>${escapeHtml(item.stored || "-")}</td><td>${escapeHtml(item.existingQuantity || "-")}</td><td>Yes</td></tr>`).join("")}</tbody></table>
  </body></html>`;
}

function wrapShoppingPdfText(text, width = 88) {
  const words = String(text ?? "").split(/\s+/).filter(Boolean);
  if (!words.length) return [""];
  const lines = [];
  let current = "";
  words.forEach((word) => {
    const candidate = current ? `${current} ${word}` : word;
    if (candidate.length > width && current) {
      lines.push(current);
      current = word;
    } else {
      current = candidate;
    }
  });
  if (current) lines.push(current);
  return lines;
}

function pdfEscape(text) {
  return String(text ?? "").replaceAll("\\", "\\\\").replaceAll("(", "\\(").replaceAll(")", "\\)");
}

function normalizePdfTextServer(value) {
  return String(value ?? "")
    .replaceAll("\u00a0", " ")
    .replaceAll("\u2018", "'")
    .replaceAll("\u2019", "'")
    .replaceAll("\u201c", '"')
    .replaceAll("\u201d", '"')
    .replaceAll("\u2013", "-")
    .replaceAll("\u2014", "-")
    .replaceAll("\u2026", "...")
    .replaceAll("\u2022", "-")
    .replaceAll("\u200b", "")
    .split("")
    .map((char) => (char.charCodeAt(0) <= 255 ? char : "?"))
    .join("");
}

function pdfRgbServer(hex) {
  const normalized = normalizeHexColorServer(hex, "#F3E7DB").replace("#", "");
  return [
    Number.parseInt(normalized.slice(0, 2), 16) / 255,
    Number.parseInt(normalized.slice(2, 4), 16) / 255,
    Number.parseInt(normalized.slice(4, 6), 16) / 255,
  ];
}

function drawPdfRectServer(commands, x, y, width, height, fillHex, strokeHex = "#cfc7bf") {
  const [fr, fg, fb] = pdfRgbServer(fillHex);
  const [sr, sg, sb] = pdfRgbServer(strokeHex);
  commands.push(`${fr.toFixed(3)} ${fg.toFixed(3)} ${fb.toFixed(3)} rg`);
  commands.push(`${x.toFixed(2)} ${y.toFixed(2)} ${width.toFixed(2)} ${height.toFixed(2)} re f`);
  commands.push(`${sr.toFixed(3)} ${sg.toFixed(3)} ${sb.toFixed(3)} RG`);
  commands.push(`0.6 w ${x.toFixed(2)} ${y.toFixed(2)} ${width.toFixed(2)} ${height.toFixed(2)} re S`);
}

function buildShoppingOrderPdfBuffer(order, settings) {
  const rows = shoppingOrderSelectedItems(order);
  const pageWidth = 595;
  const pageHeight = 842;
  const margin = 32;
  const colWidths = [78, 160, 105, 78, 78, 52];
  const headers = ["Category", "Item", "Supplier", "Stored", "Existing Qty", "Order"];
  const metaRows = shoppingOrderMetaRows(order);
  const pages = [];
  let commands = [];
  let y = pageHeight - margin;
  const startPage = () => {
    commands = [];
    y = pageHeight - margin;
    commands.push("0 0 0 rg");
    commands.push("BT");
    commands.push(`/F2 14 Tf 1 0 0 1 ${margin} ${y} Tm (${pdfEscape(normalizePdfTextServer("Shopping Order Detail"))}) Tj`);
    commands.push("ET");
    y -= 20;
    metaRows.forEach(([label, value]) => {
      commands.push("0 0 0 rg");
      commands.push("BT");
      commands.push(`/F2 8 Tf 1 0 0 1 ${margin} ${y} Tm (${pdfEscape(normalizePdfTextServer(`${label}:`))}) Tj`);
      commands.push(`/F1 8 Tf 1 0 0 1 ${margin + 90} ${y} Tm (${pdfEscape(normalizePdfTextServer(value))}) Tj`);
      commands.push("ET");
      y -= 11;
    });
    y -= 4;
  };
  const flushPage = () => {
    pages.push(commands.join("\n"));
  };
  const drawHeader = () => {
    const rowHeight = 15;
    let x = margin;
    headers.forEach((header, index) => {
      drawPdfRectServer(commands, x, y - rowHeight, colWidths[index], rowHeight, "#f1ece6");
      commands.push("0 0 0 rg");
      commands.push("BT");
      commands.push(`/F2 7.5 Tf 1 0 0 1 ${x + 3} ${y - 10.5} Tm (${pdfEscape(normalizePdfTextServer(header))}) Tj`);
      commands.push("ET");
      x += colWidths[index];
    });
    y -= rowHeight;
  };
  const wrapCell = (text, width) => wrapShoppingPdfText(text, Math.max(8, Math.floor((width - 6) / 5.0)));
  startPage();
  drawHeader();
  rows.forEach((item) => {
    const cells = [
      item.category || "-",
      item.item || "-",
      item.supplier || "-",
      item.stored || "-",
      item.existingQuantity || "-",
      "Yes",
    ];
    const wrapped = cells.map((cell, index) => wrapCell(cell, colWidths[index]));
    const rowLines = Math.max(...wrapped.map((cell) => cell.length), 1);
    const rowHeight = Math.max(14, rowLines * 8.5 + 4);
    if (y - rowHeight < margin) {
      flushPage();
      startPage();
      drawHeader();
    }
    let x = margin;
    const fill = blendShoppingColorOnWhiteServer(getShoppingCategoryColorServer(settings, item.category), 0.10);
    wrapped.forEach((cellLines, index) => {
      drawPdfRectServer(commands, x, y - rowHeight, colWidths[index], rowHeight, fill);
      cellLines.forEach((line, lineIndex) => {
        commands.push("0 0 0 rg");
        commands.push("BT");
        commands.push(`/F1 7.5 Tf 1 0 0 1 ${x + 3} ${y - 10.5 - lineIndex * 8.5} Tm (${pdfEscape(normalizePdfTextServer(line))}) Tj`);
        commands.push("ET");
      });
      x += colWidths[index];
    });
    y -= rowHeight;
  });
  if (!rows.length) {
    const rowHeight = 18;
    drawPdfRectServer(commands, margin, y - rowHeight, colWidths.reduce((sum, value) => sum + value, 0), rowHeight, "#faf7f2");
    commands.push("0 0 0 rg");
    commands.push("BT");
    commands.push(`/F1 8 Tf 1 0 0 1 ${margin + 3} ${y - 11.5} Tm (${pdfEscape(normalizePdfTextServer("No selected items in this order."))}) Tj`);
    commands.push("ET");
    y -= rowHeight;
  }
  flushPage();
  const objects = [];
  objects.push("<< /Type /Catalog /Pages 2 0 R >>");
  objects.push(`<< /Type /Pages /Count ${pages.length} /Kids [${pages.map((_, i) => `${3 + i * 2} 0 R`).join(" ")}] >>`);
  pages.forEach((stream, index) => {
    const pageObj = 3 + index * 2;
    const contentObj = pageObj + 1;
    objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /Font << /F1 ${3 + pages.length * 2} 0 R /F2 ${4 + pages.length * 2} 0 R >> >> /Contents ${contentObj} 0 R >>`);
    objects.push(`<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`);
  });
  objects.push("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
  objects.push("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>");
  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  objects.forEach((obj, index) => {
    offsets.push(pdf.length);
    pdf += `${index + 1} 0 obj\n${obj}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  offsets.slice(1).forEach((offset) => {
    pdf += `${String(offset).padStart(10, "0")} 00000 n \n`;
  });
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`;
  return Buffer.from(pdf, "latin1");
}

function buildShoppingOrderAttachments(order, settings) {
  const stem = shoppingOrderFileStem(order);
  const excelHtml = buildShoppingOrderExcelHtml(order, settings);
  const pdfBuffer = buildShoppingOrderPdfBuffer(order, settings);
  return [
    {
      filename: `${stem}.xls`,
      content: Buffer.from(excelHtml, "utf8").toString("base64"),
    },
    {
      filename: `${stem}.pdf`,
      content: pdfBuffer.toString("base64"),
    },
  ];
}

async function sendWithResend({ to, subject, html, text, attachments }) {
  const apiKey = process.env.RESEND_API_KEY;
  const rawFrom = process.env.EMAIL_FROM;
  if (!apiKey) {
    const error = new Error("Missing server environment variable: RESEND_API_KEY");
    error.statusCode = 500;
    throw error;
  }
  if (!rawFrom) {
    const error = new Error("Missing server environment variable: EMAIL_FROM");
    error.statusCode = 500;
    throw error;
  }
  const from = /<[^>]+>/.test(rawFrom) ? rawFrom : `ACOOM Tools <${rawFrom}>`;
  const response = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ from, to, subject, html, text, attachments }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload?.message || payload?.error || `Email provider failed (${response.status})`);
    error.statusCode = 502;
    throw error;
  }
  return payload;
}

async function sendShoppingEmail(order, settings, notes = "") {
  const recipients = Array.isArray(settings?.emailRecipients) ? settings.emailRecipients : [];
  if (!recipients.length) return { skipped: true };
  const shoppingDate = todayInLisbon();
  const selectedItems = shoppingOrderSelectedItems(order);
  const tableHtml = buildSelectedItemsTable(selectedItems, settings);
  const subject = `Lista de Compras - data ${shoppingDate}`;
  const cleanNotes = cleanText(notes);
  const html = `<p>Foi submetida uma nova lista de compras.</p>
    <p><strong>Data:</strong> ${escapeHtml(shoppingDate)}<br /><strong>Nome:</strong> ${escapeHtml(order.submittedByName || "-")}${cleanNotes ? `<br /><strong>Notes:</strong> ${escapeHtml(cleanNotes)}` : ""}</p>
    ${tableHtml}`;
  const text = [
    "Foi submetida uma nova lista de compras.",
    `Data: ${shoppingDate}`,
    `Nome: ${cleanText(order.submittedByName) || "-"}`,
    ...(cleanNotes ? [`Notes: ${cleanNotes}`] : []),
    "",
    ...selectedItems.map((item) => `${item.category} | ${item.item} | ${item.supplier || "-"} | Qt existente: ${item.existingQuantity || "-"} | Order: Yes`),
  ].join("\n");
  return sendWithResend({ to: recipients, subject, html, text, attachments: buildShoppingOrderAttachments(order, settings) });
}

module.exports = async function handler(req, res) {
  try {
    const { user } = await requireFeature(req, "app", "shopping");

    if (req.method === "GET") {
      const [openOrder, history, settings] = await Promise.all([
        loadOpenOrder(),
        loadSubmittedHistory(),
        loadShoppingSettings(),
      ]);
      res.status(200).json({ openOrder, history, settings });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const action = cleanText(body?.action).toLowerCase();
      const openOrder = await loadOpenOrder();
      if (openOrder) {
        const error = new Error("There is already one open shopping order.");
        error.statusCode = 400;
        throw error;
      }

      if (action === "reopen") {
        const history = await loadSubmittedHistory();
        const latest = history[0];
        const sourceId = cleanText(body?.sourceOrderId);
        if (!latest || !sourceId || sourceId !== latest.id) {
          const error = new Error("Only the latest submitted shopping order can be reopened.");
          error.statusCode = 400;
          throw error;
        }
        const payload = {
          status: "open",
          reopened_from_id: latest.id,
          items: latest.items,
          submitted_by_name: "",
          submitted_by_user_email: "",
          submitted_at: null,
          notes: "",
        };
        const created = await restQuery("shopping_orders?select=id,order_number,status,submitted_by_name,submitted_by_user_email,submitted_at,notes,reopened_from_id,items,created_at,updated_at", {
          method: "POST",
          body: [payload],
          preferRepresentation: true,
        });
        res.status(201).json({ order: sanitizeShoppingOrderRow(Array.isArray(created) ? created[0] : {}) });
        return;
      }

      const settings = await loadShoppingSettings();
      const payload = {
        status: "open",
        notes: "",
        items: buildOrderItemsFromSettings(settings),
      };
      const created = await restQuery("shopping_orders?select=id,order_number,status,submitted_by_name,submitted_by_user_email,submitted_at,notes,reopened_from_id,items,created_at,updated_at", {
        method: "POST",
        body: [payload],
        preferRepresentation: true,
      });
      res.status(201).json({ order: sanitizeShoppingOrderRow(Array.isArray(created) ? created[0] : {}) });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanText(req.query.id);
      if (!id) {
        const error = new Error("Shopping order id is required.");
        error.statusCode = 400;
        throw error;
      }
      const existing = await loadOrderById(id);
      if (!existing) {
        const error = new Error("Shopping order not found.");
        error.statusCode = 404;
        throw error;
      }
      if (existing.status !== "open") {
        const error = new Error("Only open shopping orders can be updated.");
        error.statusCode = 400;
        throw error;
      }

      const body = await parseBody(req);
      const action = cleanText(body?.action).toLowerCase();
      const settings = await loadShoppingSettings();
      const items = sanitizeOrderItems(body?.items, settings.items);
      if (!items.length) {
        const error = new Error("Shopping order must include items.");
        error.statusCode = 400;
        throw error;
      }

      const patch = {
        items,
        updated_at: new Date().toISOString(),
      };

      let emailResult = null;
      if (action === "submit") {
        const submittedByName = cleanText(body?.submittedByName || body?.name);
        if (!submittedByName) {
          const error = new Error("Name is required to submit the shopping order.");
          error.statusCode = 400;
          throw error;
        }
        const orderForValidation = { ...existing, items, submittedByName };
        validateSubmittableOrder(orderForValidation);
        patch.status = "submitted";
        patch.submitted_by_name = submittedByName;
        patch.submitted_by_user_email = cleanText(user?.email).toLowerCase();
        patch.submitted_at = new Date().toISOString();
        patch.notes = cleanText(body?.notes);
      }

      const updatedRows = await restQuery(
        `shopping_orders?id=eq.${encodeURIComponent(id)}&select=id,order_number,status,submitted_by_name,submitted_by_user_email,submitted_at,notes,reopened_from_id,items,created_at,updated_at`,
        {
          method: "PATCH",
          body: patch,
          preferRepresentation: true,
        }
      );
      const order = sanitizeShoppingOrderRow(Array.isArray(updatedRows) ? updatedRows[0] : {});
      const effectiveOrder = order.id ? order : buildShoppingOrderFallback(existing, patch);
      if (action === "submit") {
        try {
          emailResult = await sendShoppingEmail(effectiveOrder, settings, cleanText(body?.notes));
        } catch (error) {
          emailResult = { error: error.message || "Could not send shopping email." };
        }
      }
      res.status(200).json({ order: effectiveOrder, emailResult });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
