const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const css = fs.readFileSync(path.join(root, "styles.css"), "utf8");

test("reconciliation density rules are scoped to workbench and eligible records", () => {
  assert.match(html, /class="card financial-reconciliation-workbench-card"/);
  assert.match(html, /class="card financial-reconciliation-eligible-card"/);
  assert.match(css, /\.financial-reconciliation-workbench-card h2,\s*\.financial-reconciliation-eligible-card h2\s*\{\s*font-size:\s*\.94rem;/);
  assert.match(css, /\.financial-reconciliation-filters label\s*\{\s*font-size:\s*\.72rem;/);
  assert.match(css, /\.financial-reconciliation-filters input,\s*\.financial-reconciliation-filters select\s*\{\s*font-size:\s*\.70rem;/);
  assert.match(css, /\.financial-reconciliation-table th\s*\{\s*font-size:\s*\.70rem;\s*padding:\s*\.54rem;/);
  assert.match(css, /\.financial-reconciliation-table td\s*\{\s*font-size:\s*\.74rem;\s*padding:\s*\.54rem;/);
  assert.match(css, /\.financial-reconciliation-table button\s*\{\s*font-size:\s*\.84rem;/);
  assert.match(css, /@media \(max-width:\s*768px\)\s*\{\s*\.financial-reconciliation-filters input,\s*\.financial-reconciliation-filters select\s*\{\s*font-size:\s*16px;/);
  assert.match(css, /\.financial-reconciliation-table th\s*\{\s*font-size:\s*\.70rem;\s*padding:\s*\.46rem;/);
  assert.match(css, /\.financial-reconciliation-table td\s*\{\s*font-size:\s*\.70rem;\s*padding:\s*\.46rem;/);
  assert.match(css, /\.financial-reconciliation-table button\s*\{\s*font-size:\s*\.8rem;/);
});
