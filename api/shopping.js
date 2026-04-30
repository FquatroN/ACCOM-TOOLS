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

function buildSelectedItemsTable(items) {
  const rows = items
    .filter((item) => !!item.order)
    .map(
      (item) => `<tr>
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
        <th align="left">Category</th>
        <th align="left">Item</th>
        <th align="left">Supplier</th>
        <th align="left">Existing quantity</th>
        <th align="left">Order</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
  </table>`;
}

async function sendWithResend({ to, subject, html, text }) {
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
    body: JSON.stringify({ from, to, subject, html, text }),
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
  const selectedItems = Array.isArray(order?.items) ? order.items.filter((item) => !!item.order) : [];
  const tableHtml = buildSelectedItemsTable(selectedItems);
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
  return sendWithResend({ to: recipients, subject, html, text });
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
        }
      );
      const order = sanitizeShoppingOrderRow(Array.isArray(updatedRows) ? updatedRows[0] : {});
      if (action === "submit") {
        try {
          emailResult = await sendShoppingEmail(order, settings, cleanText(body?.notes));
        } catch (error) {
          emailResult = { error: error.message || "Could not send shopping email." };
        }
      }
      res.status(200).json({ order, emailResult });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
