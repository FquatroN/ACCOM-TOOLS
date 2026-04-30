const { cleanText, requireFeature, restQuery, sendError } = require("./_supabase");

function lisbonMonthParts(date = new Date()) {
  const formatter = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/Lisbon",
    year: "numeric",
    month: "2-digit",
  });
  const parts = Object.fromEntries(formatter.formatToParts(date).map((part) => [part.type, part.value]));
  return {
    year: Number(parts.year),
    month: Number(parts.month),
  };
}

function shiftMonth(year, month, delta) {
  const value = new Date(Date.UTC(year, month - 1 + delta, 1));
  return {
    year: value.getUTCFullYear(),
    month: value.getUTCMonth() + 1,
  };
}

function isoMonthStart(year, month) {
  return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-01`;
}

function monthLabel(year, month) {
  return new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/Lisbon",
    month: "short",
    year: "numeric",
  }).format(new Date(Date.UTC(year, month - 1, 1)));
}

function propertyBucket(name) {
  const raw = cleanText(name).toLowerCase();
  if (raw.includes("cruz apartments") || raw === "cruz") return "cruz";
  if (raw.includes("lisboa central hostel") || raw === "hostel") return "hostel";
  return "";
}

function average(values) {
  const nums = values.map((value) => Number(value)).filter((value) => Number.isFinite(value));
  if (!nums.length) return null;
  return Number((nums.reduce((sum, value) => sum + value, 0) / nums.length).toFixed(1));
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "communications");
    if (req.method !== "GET") {
      res.status(405).json({ error: "Method not allowed." });
      return;
    }

    const current = lisbonMonthParts();
    const previous = shiftMonth(current.year, current.month, -1);
    const next = shiftMonth(current.year, current.month, 1);
    const previousStart = isoMonthStart(previous.year, previous.month);
    const currentStart = isoMonthStart(current.year, current.month);
    const nextStart = isoMonthStart(next.year, next.month);

    const path = `reviews?select=review_date,rating_normalized_100,properties(name)&review_date=gte.${encodeURIComponent(previousStart)}&review_date=lt.${encodeURIComponent(nextStart)}&rating_normalized_100=not.is.null&order=review_date.desc`;
    const rows = await restQuery(path, { method: "GET" });

    const buckets = {
      hostel: { current: [], previous: [] },
      cruz: { current: [], previous: [] },
    };

    (Array.isArray(rows) ? rows : []).forEach((row) => {
      const key = propertyBucket(row?.properties?.name);
      const reviewDate = cleanText(row?.review_date);
      const rating = Number(row?.rating_normalized_100);
      if (!key || !reviewDate || !Number.isFinite(rating)) return;
      if (reviewDate >= currentStart && reviewDate < nextStart) buckets[key].current.push(rating);
      else if (reviewDate >= previousStart && reviewDate < currentStart) buckets[key].previous.push(rating);
    });

    res.status(200).json({
      summary: {
        months: {
          currentLabel: monthLabel(current.year, current.month),
          previousLabel: monthLabel(previous.year, previous.month),
          currentShortLabel: "Current",
          previousShortLabel: "Past",
        },
        properties: {
          hostel: {
            currentAverage: average(buckets.hostel.current),
            previousAverage: average(buckets.hostel.previous),
          },
          cruz: {
            currentAverage: average(buckets.cruz.current),
            previousAverage: average(buckets.cruz.previous),
          },
        },
      },
    });
  } catch (error) {
    sendError(res, error);
  }
};
