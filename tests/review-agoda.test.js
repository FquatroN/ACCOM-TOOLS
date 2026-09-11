const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const { normalizeSource } = require("../api/_supabase");
const root = path.join(__dirname, "..");
const appMain = fs.readFileSync(path.join(root, "app-main.js"), "utf8");
const reviewSettings = fs.readFileSync(path.join(root, "api", "review-settings.js"), "utf8");

const agodaSample = `**8.4Excellent**

**Farman** from Spain

Solo traveler

1 Person in 10-Bed Dormitory - Mixed

Stayed 1 night in September 2026

#### “Satisfactory services and nice place to stay”

During my stay the overall services was satisfactory and the staff are very cooperative.

Reviewed September 06, 2026

Did you find this review helpful?

**YesNo**

1. **10.0Exceptional**

**DJ** from Portugal

Solo traveler

1 Person in 10-Bed Dormitory - Mixed

Stayed 1 night in August 2026

#### “Excellent”

Sweet home away home place. Highly recommend!

Reviewed September 01, 2026

Did you find this review helpful?

**YesNo**

2. **10.0Exceptional**

**Ingrid** from Brazil

Solo traveler

1 Person in 6-Bed Dormitory - Mixed

Stayed 2 nights in August 2026

#### “Melhor hostel de Lisboa!”

Staff sensacional; cafe da manhã maravilhoso.

Reviewed September 01, 2026

Did you find this review helpful?

**YesNo**

3. **10.0Exceptional**

**Yamila** from Argentina

Solo traveler

1 Person in 10-Bed Dormitory - Mixed

Stayed 1 night in July 2026

#### “Hermosa experiencia”

Una experiencia verdaderamente maravillosa.

Reviewed July 27, 2026

Did you find this review helpful?

**YesNo**

4. **10.0Exceptional**

**Hei** from Hong Kong SAR, China

Solo traveler

1 Person in 10-Bed Dormitory - Mixed

Stayed 1 night in May 2026

#### “Very good”

Very good hostel. Clean and tidy.

Reviewed May 28, 2026

Did you find this review helpful?

**YesNo**`;

function appFunctionSource(name) {
  const functionStart = appMain.indexOf(`function ${name}(`);
  assert.notEqual(functionStart, -1, `missing function ${name}`);
  const bodyStart = appMain.indexOf("{", appMain.indexOf(")", functionStart));
  let depth = 0;
  for (let index = bodyStart; index < appMain.length; index += 1) {
    if (appMain[index] === "{") depth += 1;
    if (appMain[index] === "}") depth -= 1;
    if (depth === 0) return appMain.slice(functionStart, index + 1);
  }
  throw new Error(`could not extract function ${name}`);
}

function agodaParser() {
  return new Function(`
    ${appFunctionSource("clean")}
    ${appFunctionSource("normalizeNumber")}
    ${appFunctionSource("normalizePortugueseDate")}
    ${appFunctionSource("formatPortugueseDateParts")}
    ${appFunctionSource("formatDate")}
    ${appFunctionSource("normalizeDate")}
    ${appFunctionSource("agodaPlainLine")}
    ${appFunctionSource("isAgodaRatingLine")}
    ${appFunctionSource("splitAgodaReviewBlocks")}
    ${appFunctionSource("isAgodaStayDetailLine")}
    ${appFunctionSource("agodaSourceReviewId")}
    ${appFunctionSource("parseAgodaReviewBlock")}
    ${appFunctionSource("extractAgodaReviewCandidatesFromText")}
    return extractAgodaReviewCandidatesFromText;
  `)();
}

test("Agoda is registered as an active review source on the client and server", () => {
  assert.match(appMain, /\{ key: "agoda", label: "Agoda", active: true \}/);
  assert.match(reviewSettings, /\{ key: "agoda", label: "Agoda", active: true \}/);
  assert.equal(normalizeSource("Agoda"), "agoda");
});

test("Agoda reviews use the Agoda logo domain", () => {
  const reviewSourceIconDomain = new Function(`${appFunctionSource("clean")}\n${appFunctionSource("reviewSourceIconDomain")}\nreturn reviewSourceIconDomain;`)();
  assert.equal(reviewSourceIconDomain("agoda"), "agoda.com");
});

test("Agoda pasted review text parses the supplied five-review format", () => {
  const rows = agodaParser()(agodaSample, "pasted Agoda reviews");
  assert.equal(rows.length, 5);
  assert.deepEqual(rows[0], {
    source: "agoda",
    sourceReviewId: "farman:2026-09-06:satisfactory-services-and-nice-place-to-stay",
    sourceReservationId: "",
    reviewDate: "2026-09-06",
    reviewerName: "Farman",
    reviewerCountry: "Spain",
    language: "",
    ratingRaw: 8.4,
    ratingScale: 10,
    title: "Satisfactory services and nice place to stay",
    positiveReviewText: "",
    negativeReviewText: "",
    body: "During my stay the overall services was satisfactory and the staff are very cooperative.",
    subscores: {},
    hostReplyText: "",
    hostReplyDate: "",
    rawText: rows[0].rawText,
    parseConfidence: 0.9,
    warningFlags: ["agoda_text"],
    isValid: true,
    selectedForImport: true,
    rawPayload: rows[0].rawPayload,
  });
  assert.equal(rows[2].reviewerName, "Ingrid");
  assert.equal(rows[2].reviewerCountry, "Brazil");
  assert.equal(rows[3].reviewDate, "2026-07-27");
  assert.equal(rows[4].title, "Very good");
});
