const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");

const DEFAULT_ROOM_TYPES = [
  ["11 Bed Dorm Shared Bathroom", "105"],
  ["10 Bed Dorm Shared Bathroom", "102, 105"],
  ["9 Bed Dorm Shared Bathroom", "102, 105, 206"],
  ["8 Bed Dorm Shared Bathroom", "206, 203, 113, 217, 102, 105"],
  ["7 Bed Dorm Shared Bathroom", "206, 203, 113, 217, 213"],
  ["6 Bed Dorm Shared Bathroom", "206, 203, 113, 217, 213"],
  ["5 Bed Dorm Shared Bathroom", "206, 203, 113, 217, 201, 211, 213, 111"],
  ["4 Bed Dorm Shared Bathroom", "201, 202, 211, 213, 111"],
  ["4 Bed Dorm Private Bathroom", "213, 111"],
  ["3 Bed Dorm Shared Bathroom", "201, 202, 211, 213, 111, 205, 216, 212, 204, 214"],
  ["3 Bed Dorm Private Bathroom", "213, 111, 205, 204, 214"],
  ["2 Bed Dorm Shared Bathroom", "204, 205, 218, 212, 214, 215, 216, 218"],
  ["2 Bed Dorm Private Bathroom", "204, 205, 214, 215"],
  ["Twin Private with Private Bathroom", "204, 205, 214, 215"],
  ["Twin Private with Shared Bathroom", "204, 205, 218, 212, 214, 215, 216, 218"],
  ["Single Private with Private Bathroom", "204, 205, 214, 215, 112"],
  ["Single Private with Shared Bathroom", "204, 205, 218, 212, 214, 215, 216, 218, 114, 112"],
].map(([name, rooms]) => ({ name, guestsPerRoom: inferGuestsPerRoom(name), rooms: rooms.split(",").map((room) => room.trim()).filter(Boolean) }));

const DEFAULT_EMAIL_TEMPLATE = `Dear {{name}},

Thank you for contacting us.
Please find below our proposal based on your request:

Arrival: {{arrival}}
Departure: {{departure}} ({{nights}} nights)

{{room_table}}

Accommodation Total = {{accommodation_total}}

City Tax {{guests}} guests x {{city_tax_nights}} nights x 4€ = {{city_tax_total}}

Total = {{total}}

The price includes: bed sheets, fully equipped kitchen, 24h reception, free internet (computers in the lobby and Wi-Fi throughout the entire hostel) plus lots of information about Lisbon.
We also offer a free breakfast served daily from 8:00 AM to 11:00 AM, which includes a generous variety of options: three types of cereals, three types of bread, muffins, mini croissants, jam, honey, butter, peanut butter, chocolate cream, fruit, coffee, tea, cocoa, milk, juice and our homemade pancakes!

A {{deposit_percentage}}% non-refundable deposit ({{deposit_value}}) is required to confirm the reservation. The remaining balance must be paid up to {{last_payment_days}} days before arrival.

Payment can be made by:
Bank transfer
Credit card (we can send a secure payment link)

Bank details:
IBAN: PT50 0035 0137 00004852230 14
BIC/SWIFT: CGDIPTPL

We can also provide some tours and activities like Lisbon Walking Tours, Surf Lessons, PubCrawls. Please let us know if you need further information.

Cancelation Policy: {{deposit_percentage}}% after booking confirmation, 100% if canceled less than {{last_payment_days}} days before check-in.

Please note that there is a city tax of EUR 4 per person, per night that applies to all guests aged 13 and older. The amount of this TAX is already in the price total above. It is subject to a maximum amount of EUR 28 per guest.

Please let us know if you need any additional information.

Hope to hear from you soon,`;

const DEFAULT_CONFIRMATION_TEMPLATE = `Dear {{name}},

Thank you for your contact.
Your reservation has been confirmed as follows:

{{confirmation_table}}

To make any changes to an existing reservation, please contact us.
Please also let us know your expected arrival time.
Please note that any cancellations must be notified at least {{last_payment_days}} days in advance (only full rooms are accepted), otherwise the total of the reservation will be charged.

Our bank details are:
IBAN: PT50 0035 0137 00004852230 14
BIC SWIFT: CGDIPTPL

Best regards,`;

const DEFAULT_FINAL_CONFIRMATION_TEMPLATE = `Dear {{name}},

Thank you for your payment.
Your reservation is now fully paid and confirmed as follows:

{{confirmation_table}}

To make any changes to an existing reservation, please contact us.
Please also let us know your expected arrival time.

Best regards,`;

function clean(value) {
  return String(value ?? "").trim();
}

function normalizeNumber(value, fallback = 0) {
  const num = Number(String(value ?? "").replace(",", "."));
  return Number.isFinite(num) ? num : fallback;
}

function inferGuestsPerRoom(name) {
  const raw = clean(name);
  const leadingNumber = Number(raw.match(/^\d+/)?.[0]);
  if (Number.isFinite(leadingNumber) && leadingNumber > 0) return Math.min(20, leadingNumber);
  if (/twin/i.test(raw)) return 2;
  if (/single/i.test(raw)) return 1;
  return 1;
}

function normalizeGuestsPerRoom(value, fallbackName) {
  const fallback = inferGuestsPerRoom(fallbackName);
  const num = Math.round(normalizeNumber(value, fallback));
  return Math.max(1, Math.min(20, num));
}

function normalizeLastPaymentDays(value) {
  const num = Math.round(normalizeNumber(value, 14));
  return Math.max(0, Math.min(365, num));
}

function normalizeSettings(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const roomTypes = Array.isArray(source.roomTypes) ? source.roomTypes : [];
  const seen = new Set();
  const cleanRoomTypes = roomTypes
    .map((item) => ({
      name: clean(item?.name),
      guestsPerRoom: normalizeGuestsPerRoom(item?.guestsPerRoom ?? item?.guests_per_room, item?.name),
      rooms: Array.isArray(item?.rooms)
        ? item.rooms.map(clean).filter(Boolean)
        : clean(item?.rooms).split(",").map(clean).filter(Boolean),
    }))
    .filter((item) => item.name)
    .filter((item) => {
      const key = item.name.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  return {
    depositPercentage: Math.max(0, Math.min(100, normalizeNumber(source.depositPercentage, 30))),
    lastPaymentDaysBeforeArrival: normalizeLastPaymentDays(source.lastPaymentDaysBeforeArrival ?? source.last_payment_days_before_arrival),
    emailTemplate: clean(source.emailTemplate) || DEFAULT_EMAIL_TEMPLATE,
    confirmationTemplate: clean(source.confirmationTemplate) || DEFAULT_CONFIRMATION_TEMPLATE,
    finalConfirmationTemplate: clean(source.finalConfirmationTemplate) || DEFAULT_FINAL_CONFIRMATION_TEMPLATE,
    roomTypes: cleanRoomTypes.length ? cleanRoomTypes : DEFAULT_ROOM_TYPES,
  };
}

module.exports = async function handler(req, res) {
  try {
    if (req.method === "GET") {
      try {
        await requireFeature(req, "app", "groups");
      } catch {
        await requireFeature(req, "settings", "groups");
      }
      const rows = await restQuery("app_settings?select=payload&setting_key=eq.groups&limit=1", { method: "GET" });
      const payload = Array.isArray(rows) && rows[0]?.payload ? rows[0].payload : {};
      res.status(200).json({ settings: normalizeSettings(payload) });
      return;
    }

    if (req.method === "PUT") {
      await requireFeature(req, "settings", "groups");
      const body = await parseBody(req);
      const safe = normalizeSettings(body?.settings);
      const existing = await restQuery("app_settings?select=id&setting_key=eq.groups&limit=1", { method: "GET" });
      if (Array.isArray(existing) && existing[0]) {
        await restQuery("app_settings?setting_key=eq.groups", {
          method: "PATCH",
          body: { payload: safe, updated_at: new Date().toISOString() },
        });
      } else {
        await restQuery("app_settings", {
          method: "POST",
          body: [{ setting_key: "groups", payload: safe }],
        });
      }
      res.status(200).json({ settings: safe });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
