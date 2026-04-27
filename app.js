const STORAGE_KEY = "communications_log_v1";
const SEED_FLAG_KEY = "communications_recent20_seeded_v1";
const SHEET_NAME = "Comunica\u00e7\u00f5es";

const CATEGORY_COLORS = {
  Warning: "warning",
  Maintenance: "maintenance",
  Information: "information",
  "very important": "very-important",
};

const DEFAULT_STATUS = "Open";
const DEFAULT_CATEGORY = "Information";

const state = {
  entries: [],
  editingId: null,
  newDraft: {
    person: "",
    status: DEFAULT_STATUS,
    category: DEFAULT_CATEGORY,
    message: "",
  },
  editDraft: null,
  supabase: null,
  dbConfigured: false,
  user: null,
  authLoading: false,
};

const els = {
  authScreen: document.getElementById("auth-screen"),
  appShell: document.getElementById("app-shell"),
  rows: document.getElementById("rows"),
  count: document.getElementById("count"),
  search: document.getElementById("search"),
  statusFilter: document.getElementById("status-filter"),
  categoryFilter: document.getElementById("category-filter"),
  fromDate: document.getElementById("from-date"),
  toDate: document.getElementById("to-date"),
  excelInput: document.getElementById("excel-input"),
  exportCsv: document.getElementById("export-csv"),
  dbStatus: document.getElementById("db-status"),
  authEmail: document.getElementById("auth-email"),
  authPassword: document.getElementById("auth-password"),
  authSignin: document.getElementById("auth-signin"),
  authSignup: document.getElementById("auth-signup"),
  authLogout: document.getElementById("auth-logout"),
  authUser: document.getElementById("auth-user"),
  authStatus: document.getElementById("auth-status"),
};

init().catch((err) => {
  console.error(err);
  setDbStatus("Failed to initialize app.");
});

async function init() {
  bindEvents();
  await initializeDataMode();
  render();
}

function bindEvents() {
  els.rows.addEventListener("click", onRowAction);
  els.rows.addEventListener("input", onDraftInput);

  [els.search, els.statusFilter, els.categoryFilter, els.fromDate, els.toDate].forEach((el) =>
    el.addEventListener("input", render)
  );

  els.excelInput.addEventListener("change", importFromExcel);
  els.exportCsv.addEventListener("click", exportToCsv);

  if (els.authSignin) els.authSignin.addEventListener("click", onSignIn);
  if (els.authSignup) els.authSignup.addEventListener("click", onSignUp);
  if (els.authLogout) els.authLogout.addEventListener("click", onSignOut);
}

async function initializeDataMode() {
  const config = window.APP_CONFIG || {};
  const url = cleanText(config.SUPABASE_URL);
  const key = cleanText(config.SUPABASE_ANON_KEY);

  if (window.supabase && url && key) {
    state.supabase = window.supabase.createClient(url, key);
    state.dbConfigured = true;
    await initializeAuth();
    return;
  }

  state.dbConfigured = false;
  setDbStatus("Configuration missing: set SUPABASE_URL and SUPABASE_ANON_KEY in config.js.");
  window.location.replace("/gate.html");
  updateAuthUi();
}

async function initializeAuth() {
  setDbStatus("Connected to database.");
  const { data, error } = await state.supabase.auth.getSession();
  if (error) {
    setAuthStatus(error.message);
    updateAuthUi();
    return;
  }

  state.user = data.session?.user || null;

  state.supabase.auth.onAuthStateChange(async (_event, session) => {
    state.user = session?.user || null;
    updateAuthUi();

    if (state.user) {
      setAppVisibility(true);
      await loadEntriesFromDb();
    } else {
      window.location.replace("/gate.html");
    }
  });

  updateAuthUi();

  if (state.user) {
    setAuthStatus(`Signed in as ${state.user.email || "user"}`);
    setAppVisibility(true);
    await loadEntriesFromDb();
  } else {
    window.location.replace("/gate.html");
  }
}

function setAppVisibility(showApp) {
  if (els.appShell) els.appShell.hidden = !showApp;
  if (!showApp) {
    window.location.replace("/gate.html");
  }
}

function updateAuthUi() {
  const configured = state.dbConfigured;
  const authed = !!state.user;
  const busy = state.authLoading;

  if (els.authLogout) els.authLogout.disabled = !configured || !authed || busy;
  if (els.authUser) {
    els.authUser.textContent = authed
      ? `Signed in: ${state.user.email || "user"}`
      : "Not signed in";
  }

  els.excelInput.disabled = configured && !authed;
}

function setAuthBusy(busy) {
  state.authLoading = busy;
  updateAuthUi();
}

function setAuthStatus(text) {
  if (els.authStatus) els.authStatus.textContent = text;
}

async function onSignIn() {
  if (!state.dbConfigured) return;
  const email = cleanText(els.authEmail.value);
  const password = cleanText(els.authPassword.value);
  if (!email || !password) {
    setAuthStatus("Enter email and password.");
    return;
  }

  setAuthBusy(true);
  const { error } = await state.supabase.auth.signInWithPassword({ email, password });
  setAuthBusy(false);

  if (error) {
    setAuthStatus(error.message);
    return;
  }

  setAuthStatus("Signed in.");
}

async function onSignUp() {
  if (!state.dbConfigured) return;
  const email = cleanText(els.authEmail.value);
  const password = cleanText(els.authPassword.value);
  if (!email || !password) {
    setAuthStatus("Enter email and password.");
    return;
  }

  setAuthBusy(true);
  const { error } = await state.supabase.auth.signUp({ email, password });
  setAuthBusy(false);

  if (error) {
    setAuthStatus(error.message);
    return;
  }

  setAuthStatus("Account created. If email confirmation is enabled, confirm first.");
}

async function onSignOut() {
  if (!state.dbConfigured) return;
  setAuthBusy(true);
  const { error } = await state.supabase.auth.signOut();
  setAuthBusy(false);

  if (error) {
    setAuthStatus(error.message);
    return;
  }

  setAuthStatus("Signed out.");
}

function setDbStatus(text) {
  if (els.dbStatus) els.dbStatus.textContent = text;
}

function formatDbError(prefix, error) {
  if (!error) return prefix;
  const message = cleanText(error.message || error.error || String(error));
  return message ? `${prefix}: ${message}` : prefix;
}

async function getAccessToken() {
  if (!state.supabase) return null;
  const { data, error } = await state.supabase.auth.getSession();
  if (error) throw new Error(error.message || "Failed to read user session.");
  return data?.session?.access_token || null;
}

async function callApi(path, options = {}) {
  const token = await getAccessToken();
  if (!token) throw new Error("Session expired. Please sign in again.");

  const method = options.method || "GET";
  const headers = {
    Authorization: `Bearer ${token}`,
  };

  if (options.body !== undefined) {
    headers["Content-Type"] = "application/json";
  }

  const response = await fetch(path, {
    method,
    headers,
    body: options.body === undefined ? undefined : JSON.stringify(options.body),
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error || `Request failed (${response.status})`);
  }

  return payload;
}

async function loadEntriesFromDb() {
  if (!state.dbConfigured || !state.user) return;

  let rows = [];
  try {
    const result = await callApi("/api/communications");
    rows = Array.isArray(result.rows) ? result.rows : [];
  } catch (error) {
    setDbStatus(`DB error: ${error.message || "Failed to fetch records."}`);
    return;
  }

  state.entries = rows.map((row) => ({
    id: row.id,
    date: normalizeDate(cleanText(row.date)),
    time: normalizeTime(cleanText(row.time)),
    person: cleanText(row.person),
    status: cleanText(row.status) || DEFAULT_STATUS,
    category: normalizeCategory(row.category),
    message: cleanText(row.message),
    createdAt: row.created_at || new Date().toISOString(),
  }));

  setDbStatus(`Loaded ${state.entries.length} records.`);
  render();
}

function onDraftInput(event) {
  const target = event.target;
  if (!target.dataset.field) return;
  const scope = target.dataset.scope;
  if (scope === "new") {
    state.newDraft[target.dataset.field] = cleanText(target.value);
    return;
  }
  if (scope === "edit" && state.editingId && target.dataset.id === state.editingId) {
    if (!state.editDraft) state.editDraft = {};
    state.editDraft[target.dataset.field] = cleanText(target.value);
  }
}

async function onRowAction(event) {
  const button = event.target.closest("button[data-action]");
  if (!button) return;

  const action = button.dataset.action;

  if (action === "save-inline") {
    await saveNewEntry();
    return;
  }

  if (action === "save-edit") {
    await saveRowEdit(button.dataset.id);
    return;
  }

  if (action === "cancel-edit") {
    cancelRowEdit();
    render();
    return;
  }

  const id = button.dataset.id;
  if (!id) return;
  const entry = state.entries.find((item) => item.id === id);
  if (!entry) return;

  if (action === "edit") {
    state.editingId = id;
    state.editDraft = {
      person: entry.person || "",
      status: entry.status || DEFAULT_STATUS,
      category: normalizeCategory(entry.category),
      message: entry.message || "",
    };
    render();
    return;
  }

  if (action === "delete") {
    const confirmed = window.confirm("Delete this communication?");
    if (!confirmed) return;

    if (state.dbConfigured) {
      if (!state.user) {
        alert("Sign in first.");
        return;
      }
      try {
        await callApi(`/api/communications?id=${encodeURIComponent(id)}`, { method: "DELETE" });
      } catch (error) {
        alert(formatDbError("Delete failed", error));
        return;
      }
      await loadEntriesFromDb();
      return;
    }

    state.entries = state.entries.filter((item) => item.id !== id);
    if (state.editingId === id) {
      cancelRowEdit();
    }
    persistStateLocal();
    render();
  }
}

async function saveNewEntry() {
  const person = cleanText(state.newDraft.person);
  const message = cleanText(state.newDraft.message);

  if (!person || !message) {
    alert("Please fill Person and What happened.");
    return;
  }

  const now = new Date();
  const payload = {
    date: formatDate(now),
    time: formatTime(now),
    person,
    status: cleanText(state.newDraft.status) || DEFAULT_STATUS,
    category: normalizeCategory(state.newDraft.category),
    message,
  };

  if (state.dbConfigured) {
    if (!state.user) {
      alert("Sign in first.");
      return;
    }
    try {
      await callApi("/api/communications", { method: "POST", body: payload });
    } catch (error) {
      alert(formatDbError("Save failed", error));
      return;
    }
    resetNewDraft();
    await loadEntriesFromDb();
    return;
  }

  state.entries.unshift({
    id: crypto.randomUUID(),
    ...payload,
    createdAt: now.toISOString(),
  });
  persistStateLocal();
  resetNewDraft();
  render();
}

async function saveRowEdit(id) {
  if (!id || state.editingId !== id || !state.editDraft) return;
  const person = cleanText(state.editDraft.person);
  const message = cleanText(state.editDraft.message);
  if (!person || !message) {
    alert("Please fill Person and What happened.");
    return;
  }

  const payload = {
    person,
    status: cleanText(state.editDraft.status) || DEFAULT_STATUS,
    category: normalizeCategory(state.editDraft.category),
    message,
  };

  if (state.dbConfigured) {
    if (!state.user) {
      alert("Sign in first.");
      return;
    }
    try {
      await callApi(`/api/communications?id=${encodeURIComponent(id)}`, {
        method: "PUT",
        body: payload,
      });
    } catch (error) {
      alert(formatDbError("Update failed", error));
      return;
    }
    cancelRowEdit();
    await loadEntriesFromDb();
    return;
  }

  const idx = state.entries.findIndex((item) => item.id === id);
  if (idx === -1) return;
  state.entries[idx] = { ...state.entries[idx], ...payload };
  persistStateLocal();
  cancelRowEdit();
  render();
}

function cancelRowEdit() {
  state.editingId = null;
  state.editDraft = null;
}

function resetNewDraft() {
  state.newDraft = {
    person: "",
    status: DEFAULT_STATUS,
    category: DEFAULT_CATEGORY,
    message: "",
  };
}

function getFilteredEntries() {
  const search = cleanText(els.search.value).toLowerCase();
  const statusFilter = els.statusFilter.value;
  const categoryFilter = els.categoryFilter.value;
  const fromDate = els.fromDate.value;
  const toDate = els.toDate.value;

  return state.entries.filter((entry) => {
    const text = `${entry.person} ${entry.message}`.toLowerCase();
    const inSearch = !search || text.includes(search);
    const inStatus = !statusFilter || entry.status === statusFilter;
    const inCategory = !categoryFilter || normalizeCategory(entry.category) === categoryFilter;
    const inFrom = !fromDate || entry.date >= fromDate;
    const inTo = !toDate || entry.date <= toDate;

    return inSearch && inStatus && inCategory && inFrom && inTo;
  });
}

function render() {
  const filtered = getFilteredEntries();
  els.count.textContent = `${filtered.length} record${filtered.length === 1 ? "" : "s"}`;
  els.rows.innerHTML = "";

  els.rows.appendChild(buildInlineRow());

  if (filtered.length === 0) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="7" class="empty">No communications found.</td>`;
    els.rows.appendChild(tr);
    return;
  }

  filtered.forEach((entry) => {
    if (state.editingId === entry.id) {
      els.rows.appendChild(buildEditableRow(entry));
      return;
    }
    els.rows.appendChild(buildReadOnlyRow(entry));
  });
}

function buildInlineRow() {
  const tr = document.createElement("tr");
  tr.className = "inline-editor";
  const now = new Date();

  tr.innerHTML = `
    <td><span class="auto-stamp">New row</span><div>${formatDate(now)}</div></td>
    <td>${formatTime(now)}</td>
    <td>
      <input data-field="person" data-scope="new" type="text" placeholder="Person" value="${escapeHtml(
        state.newDraft.person
      )}" />
    </td>
    <td>
      <select data-field="status" data-scope="new">
        ${statusOption("Open", state.newDraft.status)}
        ${statusOption("In progress", state.newDraft.status)}
        ${statusOption("Resolved", state.newDraft.status)}
        ${statusOption("Archived", state.newDraft.status)}
      </select>
    </td>
    <td>
      <select data-field="category" data-scope="new">
        ${categoryOption("Information", state.newDraft.category)}
        ${categoryOption("Maintenance", state.newDraft.category)}
        ${categoryOption("Warning", state.newDraft.category)}
        ${categoryOption("very important", state.newDraft.category)}
      </select>
    </td>
    <td>
      <textarea data-field="message" data-scope="new" rows="2" placeholder="What happened?">${escapeHtml(
        state.newDraft.message
      )}</textarea>
    </td>
    <td class="row-actions">
      <button type="button" data-action="save-inline">Add</button>
    </td>
  `;

  return tr;
}

function buildReadOnlyRow(entry) {
  const tr = document.createElement("tr");
  const category = normalizeCategory(entry.category);
  const colorClass = CATEGORY_COLORS[category] || CATEGORY_COLORS[DEFAULT_CATEGORY];
  tr.innerHTML = `
    <td>${escapeHtml(entry.date || "")}</td>
    <td>${escapeHtml(entry.time || "")}</td>
    <td>${escapeHtml(entry.person || "")}</td>
    <td><span class="chip status">${escapeHtml(entry.status || DEFAULT_STATUS)}</span></td>
    <td><span class="chip ${colorClass}">${escapeHtml(category)}</span></td>
    <td class="message">${escapeHtml(entry.message || "")}</td>
    <td class="row-actions">
      <button type="button" data-action="edit" data-id="${entry.id}">Edit</button>
      <button type="button" data-action="delete" data-id="${entry.id}" class="danger">Delete</button>
    </td>
  `;
  return tr;
}

function buildEditableRow(entry) {
  const tr = document.createElement("tr");
  tr.className = "inline-editor";
  const draft = state.editDraft || {
    person: entry.person || "",
    status: entry.status || DEFAULT_STATUS,
    category: normalizeCategory(entry.category),
    message: entry.message || "",
  };
  tr.innerHTML = `
    <td>${escapeHtml(entry.date || "")}</td>
    <td>${escapeHtml(entry.time || "")}</td>
    <td>
      <input data-field="person" data-scope="edit" data-id="${entry.id}" type="text" value="${escapeHtml(
        draft.person
      )}" />
    </td>
    <td>
      <select data-field="status" data-scope="edit" data-id="${entry.id}">
        ${statusOption("Open", draft.status)}
        ${statusOption("In progress", draft.status)}
        ${statusOption("Resolved", draft.status)}
        ${statusOption("Archived", draft.status)}
      </select>
    </td>
    <td>
      <select data-field="category" data-scope="edit" data-id="${entry.id}">
        ${categoryOption("Information", draft.category)}
        ${categoryOption("Maintenance", draft.category)}
        ${categoryOption("Warning", draft.category)}
        ${categoryOption("very important", draft.category)}
      </select>
    </td>
    <td>
      <textarea data-field="message" data-scope="edit" data-id="${entry.id}" rows="2">${escapeHtml(
        draft.message
      )}</textarea>
    </td>
    <td class="row-actions">
      <button type="button" data-action="save-edit" data-id="${entry.id}">Save</button>
      <button type="button" data-action="cancel-edit" data-id="${entry.id}" class="ghost">Cancel</button>
    </td>
  `;
  return tr;
}

function statusOption(label, selectedValue) {
  const selected = cleanText(selectedValue) === label ? "selected" : "";
  return `<option value="${escapeHtml(label)}" ${selected}>${escapeHtml(label)}</option>`;
}

function categoryOption(label, selectedValue) {
  const normalized = normalizeCategory(label);
  const selected = normalizeCategory(selectedValue) === normalized ? "selected" : "";
  return `<option value="${escapeHtml(label)}" ${selected}>${escapeHtml(label)}</option>`;
}

function persistStateLocal() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.entries));
}

function loadStateLocal() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        state.entries = parsed
          .map((entry) => ({
            id: entry.id || crypto.randomUUID(),
            date: entry.date || "",
            time: entry.time || "",
            person: cleanText(entry.person),
            status: cleanText(entry.status) || DEFAULT_STATUS,
            category: normalizeCategory(entry.category),
            message: cleanText(entry.message),
            createdAt: entry.createdAt || new Date().toISOString(),
          }))
          .filter((entry) => entry.person || entry.message);
      }
    } catch {
      state.entries = [];
    }
  }
  applyRecent20SeedOnce();
}

function applyRecent20SeedOnce() {
  if (localStorage.getItem(SEED_FLAG_KEY) === "1") return;
  const seeded = Array.isArray(window.__RECENT20__) ? window.__RECENT20__ : [];
  if (seeded.length === 0) {
    localStorage.setItem(SEED_FLAG_KEY, "1");
    return;
  }

  const signatures = new Set(state.entries.map(entrySignature));
  const additions = seeded
    .map((entry) => ({
      id: crypto.randomUUID(),
      date: normalizeDate(cleanText(entry.date)),
      time: normalizeTime(cleanText(entry.time)),
      person: cleanText(entry.person),
      status: cleanText(entry.status) || DEFAULT_STATUS,
      category: normalizeCategory(entry.category),
      message: cleanText(entry.message),
      createdAt: new Date().toISOString(),
    }))
    .filter((entry) => entry.person || entry.message)
    .filter((entry) => {
      const sig = entrySignature(entry);
      if (signatures.has(sig)) return false;
      signatures.add(sig);
      return true;
    });

  if (additions.length > 0) {
    state.entries = [...additions, ...state.entries].sort((a, b) =>
      `${b.date} ${b.time}`.localeCompare(`${a.date} ${a.time}`)
    );
    persistStateLocal();
  }
  localStorage.setItem(SEED_FLAG_KEY, "1");
}

function entrySignature(entry) {
  return `${entry.date}|${entry.time}|${entry.person}|${entry.message}`;
}

async function importFromExcel(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  if (!window.XLSX) {
    alert("Excel parser not loaded. Check internet connection and refresh.");
    return;
  }

  if (state.dbConfigured && !state.user) {
    alert("Sign in first.");
    els.excelInput.value = "";
    return;
  }

  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      if (!workbook.Sheets[SHEET_NAME]) {
        alert(`Sheet "${SHEET_NAME}" not found.`);
        return;
      }

      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[SHEET_NAME], {
        header: 1,
        raw: false,
        defval: "",
      });

      const imported = parseComunicacoesRows(rows);
      if (imported.length === 0) {
        alert(`No valid rows found in ${SHEET_NAME} sheet.`);
        return;
      }

      if (state.dbConfigured) {
        const payload = imported.map(({ id, createdAt, ...rest }) => rest);
        try {
          await callApi("/api/communications", { method: "POST", body: payload });
        } catch (error) {
          alert(formatDbError("Import failed", error));
          return;
        }
        await loadEntriesFromDb();
      } else {
        state.entries = [...imported, ...state.entries];
        persistStateLocal();
        render();
      }

      alert(`Imported ${imported.length} communications from "${SHEET_NAME}".`);
    } catch {
      alert("Failed to import file. Please try again.");
    } finally {
      els.excelInput.value = "";
    }
  };

  reader.readAsArrayBuffer(file);
}

function parseComunicacoesRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return [];

  const headerIndex = rows.findIndex((row) => {
    const normalized = row.map((cell) => cleanText(cell).toLowerCase());
    return normalized.includes("data") && normalized.includes("hora") && normalized.includes("pessoa");
  });

  if (headerIndex === -1) return [];

  const headerRow = rows[headerIndex].map((c) => cleanText(c).toLowerCase());
  const col = {
    date: headerRow.indexOf("data"),
    time: headerRow.indexOf("hora"),
    person: headerRow.indexOf("pessoa"),
    message: headerRow.findIndex((h) => h.includes("o que aconteceu")),
    status: headerRow.indexOf("status"),
    category: Math.max(headerRow.indexOf("category"), headerRow.indexOf("categoria")),
  };

  return rows
    .slice(headerIndex + 1)
    .map((row) => {
      const person = cleanText(row[col.person]);
      const message = cleanText(row[col.message]);
      if (!person && !message) return null;

      return {
        id: crypto.randomUUID(),
        date: normalizeDate(cleanText(row[col.date])),
        time: normalizeTime(cleanText(row[col.time])),
        person,
        status: cleanText(row[col.status]) || DEFAULT_STATUS,
        category: normalizeCategory(cleanText(row[col.category])),
        message,
        createdAt: new Date().toISOString(),
      };
    })
    .filter(Boolean);
}

function exportToCsv() {
  const header = ["Data", "Hora", "Pessoa", "Status", "Category", "O que aconteceu?"];
  const lines = state.entries.map((e) => [
    e.date,
    e.time,
    e.person,
    e.status,
    normalizeCategory(e.category),
    e.message,
  ]);

  const csv = [header, ...lines]
    .map((row) => row.map((val) => `"${String(val ?? "").replaceAll('"', '""')}"`).join(","))
    .join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "communications_log.csv";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function normalizeCategory(value) {
  const raw = cleanText(value).toLowerCase();
  if (raw === "warning") return "Warning";
  if (raw === "maintenance") return "Maintenance";
  if (raw === "information" || raw === "info") return "Information";
  if (raw === "very important" || raw === "veryimportant") return "very important";
  return DEFAULT_CATEGORY;
}

function normalizeDate(value) {
  if (!value) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
  const dt = new Date(value);
  if (Number.isNaN(dt.getTime())) return value;
  return formatDate(dt);
}

function normalizeTime(value) {
  if (!value) return "";
  if (/^\d{2}:\d{2}/.test(value)) return value.slice(0, 5);
  const dt = new Date(`1970-01-01T${value}`);
  if (Number.isNaN(dt.getTime())) return value;
  return formatTime(dt);
}

function formatDate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(
    date.getDate()
  ).padStart(2, "0")}`;
}

function formatTime(date) {
  return `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
}

function cleanText(value) {
  return String(value ?? "").trim();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
