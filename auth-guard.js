(async function authGuard() {
  function showPage() {
    if (document.body) document.body.style.visibility = "visible";
  }

  const path = (window.location.pathname || "").toLowerCase();
  const isGate = path === "/" || path.endsWith("/gate.html");
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
    window.location.replace("/gate.html");
    return;
  }

  try {
    const client = window.supabase.createClient(url, key);
    const { data, error } = await client.auth.getSession();
    const authed = !error && !!data?.session?.user;

    if (isGate) {
      if (authed && !recoveryIntent) {
        window.location.replace("/index.html");
        return;
      }
      showPage();
      return;
    }

    if (!authed) {
      window.location.replace("/gate.html");
      return;
    }

    showPage();
  } catch {
    if (isGate) {
      showPage();
      return;
    }
    window.location.replace("/gate.html");
  }
})();
