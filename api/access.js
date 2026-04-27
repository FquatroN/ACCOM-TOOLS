const { loadAccessForUser, sendError, verifyUser } = require("./_supabase");

module.exports = async function handler(req, res) {
  try {
    if (req.method !== "GET") {
      res.status(405).json({ error: "Method not allowed." });
      return;
    }
    const user = await verifyUser(req);
    const access = await loadAccessForUser(user.id);
    res.status(200).json({
      user: { id: user.id, email: user.email || "" },
      profile: access.profile,
      appFeatures: access.appFeatures,
      settingsFeatures: access.settingsFeatures,
    });
  } catch (error) {
    sendError(res, error);
  }
};
