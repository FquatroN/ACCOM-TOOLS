const {
  authAdminQuery,
  cleanText,
  parseBody,
  requireSettingsAdmin,
  restQuery,
  sendError,
} = require("./_supabase");

function maskUser(user) {
  return {
    id: user?.id || "",
    email: user?.email || "",
    createdAt: user?.created_at || "",
    lastSignInAt: user?.last_sign_in_at || "",
    emailConfirmedAt: user?.email_confirmed_at || "",
    profileId: "",
  };
}

async function loadAssignmentsMap() {
  const map = new Map();
  try {
    const rows = await restQuery("user_profile_assignments?select=user_id,profile_id", { method: "GET" });
    if (Array.isArray(rows)) {
      rows.forEach((row) => {
        const userId = cleanText(row?.user_id);
        if (userId) map.set(userId, cleanText(row?.profile_id));
      });
    }
  } catch {
    return map;
  }
  return map;
}

async function assignProfile(userId, profileId, options = {}) {
  const clearFirst = options.clearFirst !== false;
  if (clearFirst) {
    await restQuery(`user_profile_assignments?user_id=eq.${encodeURIComponent(userId)}`, { method: "DELETE" });
  }
  if (!profileId) return;
  await restQuery("user_profile_assignments", {
    method: "POST",
    body: [{ user_id: userId, profile_id: profileId }],
  });
}

module.exports = async function handler(req, res) {
  try {
    await requireSettingsAdmin(req);

    if (req.method === "GET") {
      const payload = await authAdminQuery("users?page=1&per_page=200", { method: "GET" });
      const assignments = await loadAssignmentsMap();
      const users = Array.isArray(payload?.users) ? payload.users.map((user) => {
        const item = maskUser(user);
        item.profileId = assignments.get(item.id) || "";
        return item;
      }) : [];
      res.status(200).json({ users });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const email = cleanText(body?.email).toLowerCase();
      const password = String(body?.password || "");
      const profileId = cleanText(body?.profileId);

      if (!email || !email.includes("@")) {
        res.status(400).json({ error: "A valid email is required." });
        return;
      }

      if (password.length < 8) {
        res.status(400).json({ error: "Password must have at least 8 characters." });
        return;
      }

      const created = await authAdminQuery("users", {
        method: "POST",
        body: { email, password, email_confirm: true },
      });
      const user = maskUser(created?.user || created);
      if (profileId) await assignProfile(user.id, profileId);
      user.profileId = profileId;
      res.status(200).json({ user });
      return;
    }

    if (req.method === "PATCH") {
      const body = await parseBody(req);
      const userId = cleanText(body?.userId);
      const profileId = cleanText(body?.profileId);
      const password = String(body?.password || "");
      if (!userId) {
        res.status(400).json({ error: "userId is required." });
        return;
      }
      if (password) {
        if (password.length < 8) {
          res.status(400).json({ error: "Password must have at least 8 characters." });
          return;
        }
        await authAdminQuery(`users/${encodeURIComponent(userId)}`, {
          method: "PUT",
          body: { password },
        });
        res.status(200).json({ ok: true });
        return;
      }
      await assignProfile(userId, profileId);
      res.status(200).json({ ok: true });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, error);
  }
};
