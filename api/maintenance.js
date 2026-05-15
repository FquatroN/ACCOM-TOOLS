const { parseBody, requireFeature, restQuery, sendError } = require("./_supabase");
const {
  DEFAULT_MAINTENANCE_SETTINGS,
  MAINTENANCE_SETTING_KEY,
  findMaintenanceTask,
  sanitizeMaintenanceLog,
  sanitizeMaintenanceSettings,
} = require("./_maintenance");

function cleanId(value) {
  return String(value || "").trim();
}

async function loadMaintenanceSettings() {
  const rows = await restQuery(`app_settings?select=payload&setting_key=eq.${encodeURIComponent(MAINTENANCE_SETTING_KEY)}&limit=1`, {
    method: "GET",
  });
  const payload = Array.isArray(rows) && rows[0]?.payload ? rows[0].payload : DEFAULT_MAINTENANCE_SETTINGS;
  return sanitizeMaintenanceSettings(payload);
}

function isMissingMaintenanceTableError(error) {
  const message = String(error?.message || "").toLowerCase();
  return message.includes("maintenance_logs") && (
    message.includes("could not find")
    || message.includes("schema cache")
    || message.includes("relation")
    || message.includes("does not exist")
  );
}

function mapMaintenanceSchemaError(error) {
  if (isMissingMaintenanceTableError(error)) {
    const next = new Error("The Maintenance table is not available yet. Please run the migration file 2026-05-15-maintenance-logs.sql in Supabase.");
    next.statusCode = 400;
    return next;
  }
  return error;
}

function mapMaintenanceLogRow(row) {
  return sanitizeMaintenanceLog({
    id: row?.id,
    taskId: row?.task_id,
    taskName: row?.task_name,
    whereValue: row?.where_value,
    doneDate: row?.done_date,
    type: row?.type,
    who: row?.who,
    note: row?.note,
    createdAt: row?.created_at,
    updatedAt: row?.updated_at,
  });
}

async function loadMaintenanceLogs() {
  const rows = await restQuery("maintenance_logs?select=*&order=done_date.desc,created_at.desc", {
    method: "GET",
  });
  return (Array.isArray(rows) ? rows : []).map(mapMaintenanceLogRow);
}

function buildMaintenanceLogBody(record, existing = {}) {
  return {
    id: cleanId(record.id || existing.id) || undefined,
    task_id: record.taskId,
    task_name: record.taskName || "",
    where_value: record.whereValue,
    done_date: record.doneDate,
    type: record.type,
    who: record.who || "",
    note: record.note || "",
    created_at: existing.createdAt || record.createdAt || new Date().toISOString(),
    updated_at: new Date().toISOString(),
  };
}

async function createMaintenanceLog(record) {
  await restQuery("maintenance_logs", {
    method: "POST",
    body: [buildMaintenanceLogBody(record)],
    preferRepresentation: true,
  });
}

async function updateMaintenanceLog(id, record, existing = {}) {
  await restQuery(`maintenance_logs?id=eq.${encodeURIComponent(id)}`, {
    method: "PATCH",
    body: buildMaintenanceLogBody(record, existing),
    preferRepresentation: true,
  });
}

async function deleteMaintenanceLog(id) {
  await restQuery(`maintenance_logs?id=eq.${encodeURIComponent(id)}`, {
    method: "DELETE",
  });
}

module.exports = async function handler(req, res) {
  try {
    await requireFeature(req, "app", "maintenance");

    if (req.method === "GET") {
      const [settings, rows] = await Promise.all([loadMaintenanceSettings(), loadMaintenanceLogs()]);
      res.status(200).json({ settings, rows });
      return;
    }

    if (req.method === "POST") {
      const body = await parseBody(req);
      const settings = await loadMaintenanceSettings();
      const taskConfig = findMaintenanceTask(settings, body?.taskId || body?.task_id);
      if (!taskConfig) {
        res.status(400).json({ error: "Select a configured maintenance task before adding a log row." });
        return;
      }
      const record = sanitizeMaintenanceLog(body, {}, taskConfig);
      await createMaintenanceLog({
        ...record,
        taskName: taskConfig.task,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      });
      const rows = await loadMaintenanceLogs();
      res.status(200).json({ rows, settings });
      return;
    }

    if (req.method === "PUT") {
      const id = cleanId(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Maintenance log id is required." });
        return;
      }
      const body = await parseBody(req);
      const settings = await loadMaintenanceSettings();
      const rows = await loadMaintenanceLogs();
      const existing = rows.find((row) => cleanId(row.id) === id);
      if (!existing) {
        res.status(404).json({ error: "Maintenance log record not found." });
        return;
      }
      const taskConfig = findMaintenanceTask(settings, body?.taskId || body?.task_id || existing.taskId);
      if (!taskConfig) {
        res.status(400).json({ error: "The selected maintenance task is no longer configured." });
        return;
      }
      const record = sanitizeMaintenanceLog({ ...existing, ...body, id }, existing, taskConfig);
      await updateMaintenanceLog(id, { ...record, taskName: taskConfig.task }, existing);
      const nextRows = await loadMaintenanceLogs();
      res.status(200).json({ rows: nextRows, settings });
      return;
    }

    if (req.method === "DELETE") {
      const id = cleanId(req.query?.id);
      if (!id) {
        res.status(400).json({ error: "Maintenance log id is required." });
        return;
      }
      const settings = await loadMaintenanceSettings();
      await deleteMaintenanceLog(id);
      const rows = await loadMaintenanceLogs();
      res.status(200).json({ rows, settings });
      return;
    }

    res.status(405).json({ error: "Method not allowed." });
  } catch (error) {
    sendError(res, mapMaintenanceSchemaError(error));
  }
};
