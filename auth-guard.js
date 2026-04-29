(async function authGuard() {
  function showPage() {
    if (document.body) document.body.style.visibility = "visible";
  }

  function safeNextPath(raw) {
    const value = String(raw || "").trim();
    if (!value.startsWith("/") || value.startsWith("//")) return "/index.html";
    return value;
  }

  function currentPathWithQuery() {
    return `${window.location.pathname || "/index.html"}${window.location.search || ""}${window.location.hash || ""}`;
  }

  function gateUrlFor(nextPath) {
    const target = safeNextPath(nextPath || currentPathWithQuery());
    return `/gate.html?next=${encodeURIComponent(target)}`;
  }

  const path = (window.location.pathname || "").toLowerCase();
  const isGate = path.endsWith("/gate.html");
  const search = new URLSearchParams(window.location.search || "");
  const hash = new URLSearchParams(String(window.location.hash || "").replace(/^#/, ""));
  const recoveryIntent = search.get("mode") === "recovery" || search.get("type") === "recovery" || hash.get("type") === "recovery";

  const config = window.APP_CONFIG || {};
  const url = String(config.SUPABASE_URL || "").trim();
  const key = String(config.SUPABASE_ANON_KEY || "").trim();

  if (!window.supabase || !url || !key) {
    if (isGate) {
      showPage();
      return;
    }
    window.location.replace(gateUrlFor());
    return;
  }

  try {
    const client = window.supabase.createClient(url, key);
    const { data, error } = await client.auth.getSession();
    const authed = !error && !!data?.session?.user;

    if (isGate) {
      if (authed && !recoveryIntent) {
        window.location.replace(safeNextPath(search.get("next")));
        return;
      }
      showPage();
      return;
    }

    if (!authed) {
      window.location.replace(gateUrlFor());
      return;
    }

    showPage();
  } catch {
    if (isGate) {
      showPage();
      return;
    }
    window.location.replace(gateUrlFor());
  }
})();
