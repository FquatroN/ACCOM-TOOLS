const { cleanText, parseBody, requireFeature, restQuery, sendError } = require("./_supabase");

function normalizeDate(value) {
  const raw = cleanText(value);
  if (!raw) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return null;
  return date.toISOString().slice(0, 10);
}

function normalizeNumber(value, fallback = 0) {
  const num = Number(String(value ?? "").replace(",", "."));
  return Number.isFinite(num) ? num : fallback;
}

function normalizeJsonArray(value) {
  return Array.isArray(value) ? value : [];
}

function normalizeStatus(value) {
  const raw = cleanText(value).toLowerCase();
  if (raw === "accepted") return "Accepted";
  if (raw === "refused" || raw === "rejected") return "Refused";
  return "Proposal";
}

function sanitizeGroup(input = {}) {
  const name = cleanText(input.name);
  const email = cleanText(input.email).toLowerCase();
  const checkIn = normalizeDate(input.check_in ?? input.checkIn);
  const checkOut = normalizeDate(input.check_out ?? input.checkOut);
  const guests = Math.max(1, Math.min(60, Math.round(normalizeNumber(input.guests, 0))));
  if (!name) throw Object.assign(new Error("Name is required."), { statusCode: 400 });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw Object.assign(new Error("A valid email is required."), { statusCode: 400 });
  if (!checkIn || !checkOut) throw Object.assign(new Error("Check-in and check-out are required."), { statusCode: 400 });
  if (checkOut <= checkIn) throw Object.assign(new Error("Check-out must be after check-in."), { statusCode: 400 });
  if (!guests || guests > 60) throw Object.assign(new Error("Guests must be between 1 and 60."), { statusCode: 400 });

  return {
    reservation_number: cleanText(input.reservation_number ?? input.reservationNumber),
    name,
    email,
    check_in: checkIn,
    check_out: checkOut,
    guests,
    guest_groups: normalizeJsonArray(input.guest_groups ?? input.guestGroups),
    room_items: normalizeJsonArray(input.room_items ?? input.roomItems),
    total_value: normalizeNumber(input.total_value ?? input.totalValue, 0),
    option_date: normalizeDate(input.option_date ?? input.optionDate),
    status: normalizeStatus(input.status),
  };
}

function withSelect(path) {
  return `${path}${path.includes("?") ? "&" : "?"}select=id,reservation_number,creation_date,name,email,check_in,check_out,guests,guest_groups,room_items,total_value,option_date,status,created_at,updated_at`;
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "groups");

    if (req.method === "GET") {
      const rows = await restQuery(withSelect("group_proposals?order=check_in.asc,created_at.desc"), { method: "GET" });
      res.status(200).json({ rows: Array.isArray(rows) ? rows : [] });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const created = await restQuery(withSelect("group_proposals"), {
        method: "POST",
        body: [sanitizeGroup(body)],
      });
      res.status(200).json({ row: Array.isArray(created) && created[0] ? created[0] : null });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanText(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Missing id query parameter." });
        return;
      }
      const body = await parseBody(req);
      const updated = await restQuery(withSelect(`group_proposals?id=eq.${encodeURIComponent(id)}`), {
        method: "PATCH",
        body: sanitizeGroup(body),
      });
      res.status(200).json({ row: Array.isArray(updated) && updated[0] ? updated[0] : null });
      return;
    }

    if (req.method === "DELETE") {
      const id = cleanText(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Missing id query parameter." });
        return;
      }
      await restQuery(`group_proposals?id=eq.${encodeURIComponent(id)}`, { method: "DELETE" });
      res.status(200).json({ ok: true });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
