const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");

const appMain = fs.readFileSync(path.join(__dirname, "..", "app-main.js"), "utf8");

function profileMatrixRows() {
  const declaration = appMain.match(/const PROFILE_MATRIX_ROWS = \[[\s\S]*?\n\];/);
  assert.ok(declaration, "profile permission matrix must be defined");
  return new Function(`${declaration[0]}\nreturn PROFILE_MATRIX_ROWS;`)();
}

function collectProfilePayload(enabledFeatures) {
  const declaration = appMain.match(/function collectProfilePayload\(id\) \{[\s\S]*?\r?\n\}\r?\n\r?\nasync function saveUserProfile/);
  assert.ok(declaration, "profile payload collector must be defined");
  const collect = new Function("els", "clean", `${declaration[0].replace(/\r?\n\r?\nasync function saveUserProfile$/, "")}\nreturn collectProfilePayload;`)(
    {
      profilesBody: {
        querySelector(selector) {
          if (selector.includes("data-profile-name")) return { value: "Operations" };
          const match = selector.match(/data-profile-(app|settings)-([^=]+)=/);
          return { checked: Boolean(match && enabledFeatures.has(`${match[1]}:${match[2]}`)) };
        },
      },
    },
    (value) => String(value || "").trim(),
  );
  return collect("profile-1");
}

test("Admin Users exposes and saves Reconciliation as an app permission", () => {
  assert.deepEqual(
    profileMatrixRows().find((row) => row.kind === "app" && row.key === "financial-reconciliation"),
    { label: "App: Reconciliation", kind: "app", key: "financial-reconciliation" },
  );

  assert.deepEqual(
    collectProfilePayload(new Set(["app:financial-reconciliation"])).appFeatures,
    ["financial-reconciliation"],
  );
});
