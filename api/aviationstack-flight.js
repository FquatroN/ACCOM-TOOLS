const { cleanText, requireFeature, sendError } = require("./_supabase");

function normalizeFlightCode(value) {
  return cleanText(value).toUpperCase().replace(/\s+/g, "");
}

function normalizeDate(value) {
  const raw = cleanText(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  return "";
}

function parseFlightParts(flightCode) {
  const normalized = normalizeFlightCode(flightCode);
  const match = normalized.match(/^([A-Z0-9]{2,3})(\d{1,4}[A-Z]?)$/);
  return {
    normalized,
    airlineIata: match ? match[1] : "",
    flightNumber: match ? match[2] : normalized,
  };
}

function rowAirlineIata(row) {
  return cleanText(row?.airline?.iata || row?.airline?.iataCode).toUpperCase();
}

function rowFlightNumber(row) {
  return cleanText(row?.flight?.number);
}

function rowFlightIata(row) {
  return normalizeFlightCode(
    row?.flight?.iata ||
      row?.flight?.iataNumber ||
      row?.codeshared?.flight?.iataNumber ||
      row?.flight?.codeshared?.flight?.iata
  );
}

function rowCodeshareAirlineIata(row) {
  return cleanText(row?.codeshared?.airline?.iataCode || row?.flight?.codeshared?.airline?.iata).toUpperCase();
}

function rowCodeshareFlightNumber(row) {
  return cleanText(row?.codeshared?.flight?.number || row?.flight?.codeshared?.flight?.number);
}

function rowCodeshareFlightIata(row) {
  return normalizeFlightCode(row?.codeshared?.flight?.iataNumber || row?.flight?.codeshared?.flight?.iata);
}

function scoreFlightCandidate(row, flightCode, airlineIata, flightNumber) {
  const directIata = rowFlightIata(row);
  const directNumber = rowFlightNumber(row);
  const directAirline = rowAirlineIata(row);
  const codeshareIata = rowCodeshareFlightIata(row);
  const codeshareNumber = rowCodeshareFlightNumber(row);
  const codeshareAirline = rowCodeshareAirlineIata(row);
  if (directIata && directIata === flightCode) return 100;
  if (codeshareIata && codeshareIata === flightCode) return 95;
  if (directAirline === airlineIata && directNumber === flightNumber) return 85;
  if (codeshareAirline === airlineIata && codeshareNumber === flightNumber) return 80;
  if (directNumber === flightNumber) return 60;
  return 0;
}

function legTimeByKind(legBlock, kind = "best") {
  const normalizedKind = cleanText(kind).toLowerCase();
  if (normalizedKind === "scheduled") {
    return cleanText(legBlock?.scheduled) || cleanText(legBlock?.scheduledTime) || cleanText(legBlock?.estimated) || cleanText(legBlock?.estimatedTime) || cleanText(legBlock?.actual) || cleanText(legBlock?.actualTime);
  }
  return cleanText(legBlock?.estimated) || cleanText(legBlock?.estimatedTime) || cleanText(legBlock?.actual) || cleanText(legBlock?.actualTime) || cleanText(legBlock?.scheduled) || cleanText(legBlock?.scheduledTime);
}

module.exports = async function handler(req, res) {
  try {
    if (req.method !== "GET") {
      res.status(405).json({ error: "Method not allowed." });
      return;
    }

    await requireFeature(req, "app", "services");

    const apiKey = process.env.AVIATIONSTACK_API_KEY;
    if (!apiKey) {
      const error = new Error("Missing server environment variable: AVIATIONSTACK_API_KEY");
      error.statusCode = 500;
      throw error;
    }

    const flightCode = normalizeFlightCode(req.query?.flight);
    const flightDate = normalizeDate(req.query?.date);
    if (!flightCode || !flightDate) {
      res.status(400).json({ error: "Missing flight or date query parameter." });
      return;
    }

    const parts = parseFlightParts(flightCode);
    const params = new URLSearchParams({ access_key: apiKey });
    const endpointPath = "flights";
    const leg = cleanText(req.query?.leg).toLowerCase() === "departure" ? "departure" : "arrival";
    const timeKind = cleanText(req.query?.time_kind).toLowerCase() === "scheduled" ? "scheduled" : "best";
    params.set("flight_iata", parts.normalized);
    params.set("limit", "100");

    const response = await fetch(`https://api.aviationstack.com/v1/${endpointPath}?${params.toString()}`);
    const payload = await response.json().catch(() => ({}));
    if (!response.ok || payload?.error) {
      const apiMessage = payload?.error?.message || payload?.error?.type || `Aviationstack failed (${response.status})`;
      const error = new Error(apiMessage);
      error.statusCode = response.status || 502;
      throw error;
    }

    const rows = Array.isArray(payload?.data) ? payload.data : [];
    const dateMatchedRows = rows.filter((row) => cleanText(row?.flight_date) === flightDate);
    const candidateRows = dateMatchedRows.length ? dateMatchedRows : rows;
    const best = candidateRows
      .map((row) => ({ row, score: scoreFlightCandidate(row, parts.normalized, parts.airlineIata, parts.flightNumber) }))
      .sort((a, b) => b.score - a.score)[0];

    if (!best?.row || best.score <= 0) {
      res.status(404).json({ error: dateMatchedRows.length ? "Flight not found for that date." : "Flight not found." });
      return;
    }

    const legBlock = best.row?.[leg] || {};
    const predictedTime = legTimeByKind(legBlock, timeKind);

    res.status(200).json({
      predictedTime: predictedTime || "",
      timeZone: cleanText(legBlock?.timezone),
      airport: cleanText(legBlock?.airport),
      status: cleanText(best.row?.flight_status || best.row?.status),
      leg,
      flightIata: rowFlightIata(best.row) || rowCodeshareFlightIata(best.row),
      message: "",
    });
  } catch (error) {
    sendError(res, error);
  }
};
