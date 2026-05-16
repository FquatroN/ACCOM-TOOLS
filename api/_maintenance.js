const { randomUUID } = require("node:crypto");
const { cleanText, normalizeDate } = require("./_supabase");

const MAINTENANCE_SETTING_KEY = "maintenance";

const DEFAULT_MAINTENANCE_TASKS = [
  {
    task: "Socket Camas",
    taskDescription: "Manutenção das tomadas das camas, verificação, reparação ou Nova",
    where: "102.1, 102.2, 102.3, 102.4, 102.5, 102.6, 102.7, 102.8, 102.9, 102.10, 105.1, 105.2, 105.3, 105.4, 105.5, 105.6, 105.7, 105.8, 105.9, 105.10, 201.1, 201.2, 201.3, 201.4, 202.1, 202.2, 202.3, 202.4, 203.1, 203.2, 203.3, 203.4, 203.5, 203.6, 206.1, 206.2, 206.3, 206.4, 206.5, 206.6, 211.1, 211.2, 211.3, 211.4, 213.1, 213.2, 213.3, 213.4, 217.1, 217.2, 217.3, 217.4, 217.5, 217.6, 217.7, 111.1, 111.2, 111.3, 111.4, 113.1, 113.2, 113.3, 113.4, 113.5, 113.6",
    type: "Verificado, Reparado, Nova",
  },
  {
    task: "Filtros Ar Condicionados",
    taskDescription: "Limpeza dos filtros das maquinas interiores de AC",
    where: "102, 10Almofadas, 105, 201, 202, 203, 204, 205, 206, 207, 211, 212, 213, 214, 215, 216, 217, 218, 111, 112, 113, 114, 2D, 2E, 3E, 4E, 4D, 5D, 5E",
    type: "Verificado, Lavado",
  },
  {
    task: "Lavgem Colchas Casal Apartamentos",
    taskDescription: "Lavagem das colchas das camas de Casal dos Apartamentos",
    where: "2D, 2E, 3E, 4D, 4E, 5D, 5E",
    type: "Lavado",
  },
  {
    task: "Ventoinhas Dorms",
    taskDescription: "Limpeza de Ventoinhas dos Dorms – Fazer com compressor ou soprador",
    where: "102, 105, 201, 202, 203, 206, 211, 213, 217, 111, 113",
    type: "Limpo",
  },
  {
    task: "Cortinas de Janelas",
    taskDescription: "Lavagem das Cortinas das Janelas",
    where: "102, 10Almofadas, 10Cozinha, 105, 201, 202, 203, 204, 205, 206, 207, 211, 212, 213, 214, 215, 216, 217, 218, 111, 112, 113, 114, 2D, 2E, 3E, 4E, 4D, 5D, 5E",
    type: "Lavado",
  },
  {
    task: "Cortinas Camas",
    taskDescription: "Lavagem das Cortinas de Camas",
    where: "102, 105, 201, 202, 203, 206, 211, 213, 217, 111, 113, 2E",
    type: "Lavado",
  },
  {
    task: "Esfregonas Pas e Vassouras Apartamentos",
    taskDescription: "Verificação do estado das Pás, Vassouras e Esfregonas nos Apartamentos",
    where: "2D, 2E, 3E, 4D, 4E, 5D, 5E",
    type: "Verificado, Nova",
  },
  {
    task: "Resguardos Camas",
    taskDescription: "Lavagem de Resguardos Camas",
    where: "102, 105, 201, 202, 203, 204, 205, 206, 207, 211, 212, 213, 214, 215, 216, 217, 218, 111, 112, 113, 114, 2D, 2E, 3E, 4E, 4D, 5D, 5E",
    type: "Lavado",
  },
  {
    task: "Kits de Emergencia",
    taskDescription: "Verificação e revisão dos items nos Kits de Emergência (Compressas - ligaduras - Luvas descartaveis - Pensos - Adesivos - Tesoura - Pinça - Beatdine - Soro )",
    where: "2D, 2E, 3E, 4D, 4E, 5D, 5E, Receção",
    type: "Verificado",
  },
  {
    task: "Bedbugs",
    taskDescription: "Registar deteções, inspeções, ou intervenções (nossas ou de empresa)",
    where: "102, 105, 201, 202, 203, 204, 205, 206, 207, 211, 212, 213, 214, 215, 216, 217, 218, 111, 112, 113, 114, 2D, 2E, 3E, 4E, 4D, 5D, 5E",
    type: "Verificado, Detetado, Desinfestação, Intervenção",
  },
  {
    task: "AQS",
    taskDescription: "Registar problemas, reparações e manutenções nos equipamentos de aquecimento de Águas",
    where: "10 Termoacumulador, 10 CaldeiraDeposito, 11 StaffTermoacumulador, 112 Termoacumulador, 20 CaldeiraDeposito, 21 CaldeiraDeposito, 2D CaldeiraDeposito, 2E CaldeiraDeposito, 3E CaldeiraDeposito, 4D CaldeiraDeposito, 4E CaldeiraDeposito, 5D Esquentador, 5E Esquentador",
    type: "Avaria, Reparação, Manutenção",
  },
  {
    task: "Colchoes",
    taskDescription: "Registar substituição de colchões por novos",
    where: "102, 105, 201, 202, 203, 204, 205, 206, 207, 211, 212, 213, 214, 215, 216, 217, 218, 111, 112, 113, 114, 2D, 2E, 3E, 4E, 4D, 5D, 5E",
    type: "Novo",
  },
];

const DEFAULT_MAINTENANCE_SETTINGS = {
  tasks: DEFAULT_MAINTENANCE_TASKS.map((task, index) => sanitizeMaintenanceTask(task, index)),
};

function slugify(value) {
  return cleanText(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function normalizeCsvList(value) {
  const source = Array.isArray(value) ? value.join(",") : String(value || "");
  const seen = new Set();
  return source
    .split(/[\n,;]/)
    .map((item) => cleanText(item))
    .filter(Boolean)
    .filter((item) => {
      const key = item.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function sanitizeMaintenanceTask(input = {}, index = 0) {
  const name = cleanText(input.task || input.name);
  const maxDaysRaw = Number.parseInt(cleanText(input.maxDays || input.max_days), 10);
  const task = {
    id: cleanText(input.id) || `maintenance-task-${index + 1}-${slugify(name || `task-${index + 1}`)}`,
    task: name,
    taskDescription: cleanText(input.taskDescription || input.task_description || input.description).slice(0, 1000),
    whereOptions: normalizeCsvList(input.whereOptions || input.where_options || input.where),
    typeOptions: normalizeCsvList(input.typeOptions || input.type_options || input.type),
    maxDays: Number.isFinite(maxDaysRaw) && maxDaysRaw > 0 ? maxDaysRaw : "",
  };
  return task;
}

function sanitizeMaintenanceSettings(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const seen = new Set();
  const tasks = (Array.isArray(source.tasks) ? source.tasks : [])
    .map((task, index) => sanitizeMaintenanceTask(task, index))
    .filter((task) => task.task)
    .filter((task) => {
      const key = cleanText(task.id).toLowerCase();
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  return {
    tasks: tasks.length ? tasks : cloneDefaultMaintenanceTasks(),
  };
}

function cloneDefaultMaintenanceTasks() {
  return DEFAULT_MAINTENANCE_SETTINGS.tasks.map((task) => ({
    id: task.id,
    task: task.task,
    taskDescription: task.taskDescription,
    whereOptions: [...task.whereOptions],
    typeOptions: [...task.typeOptions],
    maxDays: task.maxDays,
  }));
}

function findMaintenanceTask(settings, taskId) {
  const tasks = Array.isArray(settings?.tasks) ? settings.tasks : [];
  return tasks.find((task) => cleanText(task.id) === cleanText(taskId)) || null;
}

function sanitizeMaintenanceLog(input = {}, existing = {}, taskConfig = null) {
  const taskId = cleanText(input.taskId || input.task_id || existing.taskId || existing.task_id || taskConfig?.id);
  const whereValue = cleanText(input.whereValue || input.where_value || input.where || existing.whereValue || existing.where_value || existing.where);
  const type = cleanText(input.type || existing.type);
  const record = {
    id: cleanText(input.id || existing.id) || randomUUID(),
    taskId,
    taskName: cleanText(input.taskName || input.task_name || existing.taskName || existing.task_name || taskConfig?.task).slice(0, 200),
    whereValue,
    doneDate: normalizeDate(input.doneDate || input.done_date || input.date || existing.doneDate || existing.done_date || existing.date),
    type,
    who: cleanText(input.who || existing.who).slice(0, 120),
    note: String(input.note ?? existing.note ?? "").trim().slice(0, 4000),
    createdAt: cleanText(input.createdAt || input.created_at || existing.createdAt || existing.created_at),
    updatedAt: cleanText(input.updatedAt || input.updated_at || existing.updatedAt || existing.updated_at),
  };
  validateMaintenanceLog(record, taskConfig);
  return record;
}

function validateMaintenanceLog(record, taskConfig = null) {
  if (!cleanText(record.taskId)) {
    const error = new Error("Task is required.");
    error.statusCode = 400;
    throw error;
  }
  if (!cleanText(record.doneDate)) {
    const error = new Error("Date is required.");
    error.statusCode = 400;
    throw error;
  }
  if (!cleanText(record.whereValue)) {
    const error = new Error("Where is required.");
    error.statusCode = 400;
    throw error;
  }
  if (!cleanText(record.type)) {
    const error = new Error("Type is required.");
    error.statusCode = 400;
    throw error;
  }
  if (taskConfig) {
    if (Array.isArray(taskConfig.whereOptions) && taskConfig.whereOptions.length && !taskConfig.whereOptions.includes(record.whereValue)) {
      const error = new Error("Where must be one of the configured values for the selected task.");
      error.statusCode = 400;
      throw error;
    }
    if (Array.isArray(taskConfig.typeOptions) && taskConfig.typeOptions.length && !taskConfig.typeOptions.includes(record.type)) {
      const error = new Error("Type must be one of the configured values for the selected task.");
      error.statusCode = 400;
      throw error;
    }
  }
}

module.exports = {
  DEFAULT_MAINTENANCE_SETTINGS,
  MAINTENANCE_SETTING_KEY,
  cloneDefaultMaintenanceTasks,
  findMaintenanceTask,
  normalizeCsvList,
  sanitizeMaintenanceLog,
  sanitizeMaintenanceSettings,
  sanitizeMaintenanceTask,
  slugify,
};
