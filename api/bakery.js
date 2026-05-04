const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  DEFAULT_BAKERY_SETTINGS,
  buildBakeryGeneratedText,
  cleanText,
  formatLisbonIso,
  generateTargetDates,
  orderDatesLabel,
  sanitizeBakeryDays,
  sanitizeBakeryOrderRow,
  sanitizeBakerySettings,
} = require("./_bakery");
const { sendWithSmtp } = require("./_smtp");

async function loadBakerySettings() {
  const rows = await restQuery("app_settings?select=payload&setting_key=eq.bakery&limit=1", { method: "GET" });
  const payload = Array.isArray(rows) && rows[0]?.payload ? rows[0].payload : DEFAULT_BAKERY_SETTINGS;
  const settings = sanitizeBakerySettings(payload);
  const generalRows = await restQuery("app_settings?select=payload&setting_key=eq.communications&limit=1", { method: "GET" });
  const generalPayload = Array.isArray(generalRows) && generalRows[0]?.payload ? generalRows[0].payload : {};
  const generalEmailConfig = generalPayload?.general?.emailConfig || generalPayload?.general?.email_config || generalPayload?.general?.bakeryEmailConfig || generalPayload?.general?.bakery_email_config;
  if (generalEmailConfig && typeof generalEmailConfig === "object") {
    settings.emailConfig = sanitizeBakerySettings({ emailConfig: generalEmailConfig }).emailConfig;
  }
  return settings;
}

async function loadOpenOrder(settings) {
  const rows = await restQuery(
    "bakery_orders?select=id,order_number,status,order_date,target_dates,days,generated_text,submitted_by_name,submitted_by_user_email,submitted_at,created_at,updated_at&status=eq.open&limit=1",
    { method: "GET" }
  );
  return Array.isArray(rows) && rows[0] ? sanitizeBakeryOrderRow(rows[0], settings) : null;
}

async function loadHistory(settings) {
  const rows = await restQuery(
    "bakery_orders?select=id,order_number,status,order_date,target_dates,days,generated_text,submitted_by_name,submitted_by_user_email,submitted_at,created_at,updated_at&status=eq.submitted&order=submitted_at.desc,updated_at.desc,order_number.desc",
    { method: "GET" }
  );
  return Array.isArray(rows) ? rows.map((row) => sanitizeBakeryOrderRow(row, settings)) : [];
}

async function loadOrderById(id, settings) {
  const rows = await restQuery(
    `bakery_orders?select=id,order_number,status,order_date,target_dates,days,generated_text,submitted_by_name,submitted_by_user_email,submitted_at,created_at,updated_at&id=eq.${encodeURIComponent(id)}&limit=1`,
    { method: "GET" }
  );
  return Array.isArray(rows) && rows[0] ? sanitizeBakeryOrderRow(rows[0], settings) : null;
}

async function sendWithResend({ to, subject, html, text }, emailConfig = {}) {
  const apiKey = process.env.RESEND_API_KEY;
  const rawFrom = process.env.EMAIL_FROM;
  const configuredFromEmail = cleanText(emailConfig.fromEmail || "").toLowerCase();
  const configuredFromName = cleanText(emailConfig.fromName || "");
  const replyTo = configuredFromEmail || "global@lisboacentralhostel.com";
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
  const envFromEmail = cleanText(rawFrom.replace(/^.*<([^>]+)>.*$/, "$1") || rawFrom).toLowerCase();
  const effectiveFromEmail = configuredFromEmail || envFromEmail;
  const effectiveFromName = configuredFromName || "ACCOM Tools - LCH";
  const from = `${effectiveFromName} <${effectiveFromEmail}>`;
  const response = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ from, reply_to: replyTo, to, subject, html, text }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload?.message || payload?.error || `Email provider failed (${response.status})`);
    error.statusCode = 502;
    throw error;
  }
  return payload;
}

async function sendBakeryMessage(settings, mail) {
  const emailConfig = settings?.emailConfig || {};
  if (cleanText(emailConfig.provider).toLowerCase() === "smtp") {
    return sendWithSmtp(emailConfig, mail);
  }
  return sendWithResend(mail, emailConfig);
}

async function sendBakeryEmail(order, settings) {
  const recipients = Array.isArray(settings?.emailRecipients) ? settings.emailRecipients : [];
  if (!recipients.length) return { skipped: true };
  const subject = `Encomenda pães e bolos para dias ${orderDatesLabel(order.days)}`;
  const htmlRows = (Array.isArray(order.days) ? order.days : []).map((day) => `
    <tr>
      <td style="border:1px solid #d8d0c7;padding:6px;">${day.date}</td>
      <td style="border:1px solid #d8d0c7;padding:6px;">${day.hostelGuests}</td>
      <td style="border:1px solid #d8d0c7;padding:6px;">${(day.breadBreakdown || []).map((item) => `${item.name}: ${item.quantity}`).join("<br />")}</td>
      <td style="border:1px solid #d8d0c7;padding:6px;">${day.pasteisDeNata}</td>
    </tr>`).join("");
  const html = `<!doctype html><html><body style="font-family:Arial,sans-serif;color:#1f2937;">
    <p>Bom dia,</p>
    <p>Segue a encomenda de pães e bolos:</p>
    <table style="border-collapse:collapse;width:100%;max-width:760px;">
      <thead>
        <tr>
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:left;">Data</th>
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:left;">Hóspedes Hostel</th>
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:left;">Pães</th>
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:left;">Past\u00e9is de nata</th>
        </tr>
      </thead>
      <tbody>${htmlRows}</tbody>
    </table>
    <p>Cumprimentos,<br />${order.submittedByName || "-"}<br /><br />
    Lisboa Central Hostel<br /><br />
    +351 309 881 038<br />
    +351 925 222 809<br />
    global@lisboacentralhostel.com<br /><br />
    Rua Rodrigues Sampaio 160, 1150-282 Lisboa</p>
  </body></html>`;
  const text = buildBakeryGeneratedText(order, settings, order.submittedByName);
  return sendWithResend({ to: recipients, subject, html, text });
}

async function sendBakeryEmail(order, settings) {
  const recipients = Array.isArray(settings?.emailRecipients) ? settings.emailRecipients : [];
  if (!recipients.length) return { skipped: true };
  const subject = `Lisboa Central Hostel - Encomenda p\u00e3es e bolos para dias ${orderDatesLabel(order.days)}`;
  const breadTypes = (Array.isArray(settings?.breadTypes) ? settings.breadTypes : [])
    .map((item) => cleanText(item?.name))
    .filter(Boolean);
  const htmlRows = (Array.isArray(order.days) ? order.days : []).map((day) => `
    <tr>
      <td style="border:1px solid #d8d0c7;padding:6px;">${day.date}</td>
      ${breadTypes.map((breadType) => {
        const found = (Array.isArray(day.breadBreakdown) ? day.breadBreakdown : []).find((item) => cleanText(item?.name).toLowerCase() === breadType.toLowerCase());
        return `<td style="border:1px solid #d8d0c7;padding:6px;text-align:center;">${found && found.quantity !== "" ? Number(found.quantity || 0) : "-"}</td>`;
      }).join("")}
      <td style="border:1px solid #d8d0c7;padding:6px;text-align:center;">${day.pasteisDeNata === "" ? "-" : Number(day.pasteisDeNata || 0)}</td>
    </tr>`).join("");
  const html = `<!doctype html><html><body style="font-family:Arial,sans-serif;color:#1f2937;">
    <p>Bom dia,</p>
    <p>Segue a encomenda de p\u00e3es e bolos:</p>
    <table style="border-collapse:collapse;width:100%;max-width:760px;">
      <thead>
        <tr>
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:left;">Data</th>
          ${breadTypes.map((breadType) => `<th style="border:1px solid #d8d0c7;padding:6px;text-align:center;">${breadType}</th>`).join("")}
          <th style="border:1px solid #d8d0c7;padding:6px;text-align:left;">Past\u00e9is de nata</th>
        </tr>
      </thead>
      <tbody>${htmlRows}</tbody>
    </table>
    <p>Cumprimentos,<br />${order.submittedByName || "-"}<br /><br />
    Lisboa Central Hostel<br /><br />
    +351 309 881 038<br />
    +351 925 222 809<br />
    global@lisboacentralhostel.com<br /><br />
    Rua Rodrigues Sampaio 160, 1150-282 Lisboa</p>
  </body></html>`;
  const text = buildBakeryGeneratedText(order, settings, order.submittedByName);
  return sendBakeryMessage(settings, { to: recipients, subject, html, text });
}

module.exports = async function handler(req, res) {
  try {
    const { user } = await requireFeature(req, "app", "bakery");
    const settings = await loadBakerySettings();

    if (req.method === "GET") {
      const [openOrder, history] = await Promise.all([loadOpenOrder(settings), loadHistory(settings)]);
      res.status(200).json({ openOrder, history, settings });
      return;
    }

    if (req.method === "POST") {
      const openOrder = await loadOpenOrder(settings);
      if (openOrder) {
        const error = new Error("There is already one open bakery order.");
        error.statusCode = 400;
        throw error;
      }
      const orderDate = formatLisbonIso(new Date());
      const targetDates = generateTargetDates(orderDate);
      const days = sanitizeBakeryDays([], settings, targetDates);
      const generatedText = buildBakeryGeneratedText({ days }, settings, "");
      const created = await restQuery(
        "bakery_orders?select=id,order_number,status,order_date,target_dates,days,generated_text,submitted_by_name,submitted_by_user_email,submitted_at,created_at,updated_at",
        {
          method: "POST",
          body: [{
            status: "open",
            order_date: orderDate,
            target_dates: targetDates,
            days,
            generated_text: generatedText,
          }],
          preferRepresentation: true,
        }
      );
      res.status(201).json({ order: sanitizeBakeryOrderRow(Array.isArray(created) ? created[0] : {}, settings) });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanText(req.query?.id);
      if (!id) {
        const error = new Error("Bakery order id is required.");
        error.statusCode = 400;
        throw error;
      }
      const existing = await loadOrderById(id, settings);
      if (!existing) {
        const error = new Error("Bakery order not found.");
        error.statusCode = 404;
        throw error;
      }
      const body = await parseBody(req);
      const action = cleanText(body?.action).toLowerCase();
      if (action === "resend") {
        if (existing.status !== "submitted") {
          const error = new Error("Only submitted bakery orders can resend email.");
          error.statusCode = 400;
          throw error;
        }
        let emailResult = null;
        try {
          emailResult = await sendBakeryEmail(existing, settings);
        } catch (error) {
          emailResult = { error: error.message || "Could not send bakery email." };
        }
        res.status(200).json({ order: existing, emailResult });
        return;
      }
      if (existing.status !== "open") {
        const error = new Error("Only open bakery orders can be updated.");
        error.statusCode = 400;
        throw error;
      }
      const days = sanitizeBakeryDays(body?.days, settings, existing.targetDates);
      const submittedByName = cleanText(body?.submittedByName || body?.name);
      if (action === "submit" && !submittedByName) {
        const error = new Error("Name is required to submit the bakery order.");
        error.statusCode = 400;
        throw error;
      }
      const orderForText = { ...existing, days, submittedByName: submittedByName || existing.submittedByName };
      const patch = {
        days,
        generated_text: buildBakeryGeneratedText(orderForText, settings, orderForText.submittedByName),
        updated_at: new Date().toISOString(),
      };
      if (action === "submit") {
        patch.status = "submitted";
        patch.submitted_by_name = submittedByName;
        patch.submitted_by_user_email = cleanText(user?.email).toLowerCase();
        patch.submitted_at = new Date().toISOString();
      }
      const updated = await restQuery(
        `bakery_orders?id=eq.${encodeURIComponent(id)}&select=id,order_number,status,order_date,target_dates,days,generated_text,submitted_by_name,submitted_by_user_email,submitted_at,created_at,updated_at`,
        {
          method: "PATCH",
          body: patch,
          preferRepresentation: true,
        }
      );
      const order = sanitizeBakeryOrderRow(Array.isArray(updated) ? updated[0] : { ...existing, ...patch }, settings);
      let emailResult = null;
      if (action === "submit") {
        try {
          emailResult = await sendBakeryEmail(order, settings);
        } catch (error) {
          emailResult = { error: error.message || "Could not send bakery email." };
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

