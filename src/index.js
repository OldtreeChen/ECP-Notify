import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

const REQUIRED_ENV = [
  "ECP_BASE_URL",
  "ECP_LOGIN_NAME",
  "ECP_PASSWORD"
];

const CLOSED_STATUSES = new Set(["Finished", "Cancel", "Discarded"]);
const DATA_DIR = path.resolve(".data");
const STATE_FILE = path.join(DATA_DIR, "notifications.json");
const DEFAULT_LIST_ID = "296aa935-f6c0-4a8e-9ab9-32254ea39861";
const DGPA_NON_WORKDAYS = {
  2026: new Set([
    "2026-01-01",
    "2026-02-16",
    "2026-02-17",
    "2026-02-18",
    "2026-02-19",
    "2026-02-20",
    "2026-02-27",
    "2026-04-03",
    "2026-04-06",
    "2026-05-01",
    "2026-06-19",
    "2026-09-25",
    "2026-09-28",
    "2026-10-09",
    "2026-10-26",
    "2026-12-25"
  ])
};

export const DEPARTMENT_SCHEMAS = {
  project1: {
    name: "\u5c08\u6848\u4e00\u90e8",
    title: "ECP \u5c08\u6848\u4e00\u90e8\u4efb\u52d9",
    schemaId: "ffffff19-cc0e-0b66-5801-6006b23cf204",
    identityId:
      process.env.ECP_PROJECT1_IDENTITY_ID || "d8bf955b-0ac0-11ea-9186-0a79a042dc0a"
  },
  project2: {
    name: "\u5c08\u6848\u4e8c\u90e8",
    title: "ECP \u5c08\u6848\u4e8c\u90e8\u4efb\u52d9",
    schemaId: "ffffff19-cc0e-1122-5001-6006b23cf204",
    identityId:
      process.env.ECP_PROJECT2_IDENTITY_ID || "d8bf955b-0ac0-11ea-9186-0a79a042dc0a"
  },
  cloud: {
    name: "\u96f2\u7aef\u670d\u52d9\u90e8",
    title: "ECP \u96f2\u7aef\u670d\u52d9\u90e8\u4efb\u52d9",
    schemaId: "ffffff19-d0f3-baca-7006-9bd1843cf204",
    dueSoonSchemaId: "ffffff19-d0f5-28b8-2806-9bd1843cf204",
    identityId:
      process.env.ECP_CLOUD_IDENTITY_ID || "9a4ea951-eac3-11f0-92f3-0607bbc2ee97"
  }
};

export async function runReminderJob(options = {}) {
  loadEnvFile();
  validateEnv();

  const config = getConfig(options);
  if (!isDgpaWorkday(new Date(), config.scheduleTimezone)) {
    console.log(`Skip reminders: ${formatDateInTimezone(new Date(), config.scheduleTimezone)} is not a DGPA workday.`);
    return;
  }

  const ecp = new EcpClient(config);
  await ecp.login();

  const state = loadState();
  const nextState = structuredClone(state);
  let sentRuleCount = 0;
  let overdueCount = 0;
  let dueSoonCount = 0;

  for (const rule of config.rules) {
    if (!shouldRunRuleToday(rule, config.scheduleTimezone)) {
      continue;
    }

    if (rule.identityId) {
      await ecp.switchIdentity(rule.identityId);
    }

    const result = await queryRule(ecp, rule);
    const notification = buildNotification(result, nextState, rule);
    if (!notification) {
      continue;
    }

    if (config.dryRun) {
      console.log(`[DRY RUN] ${rule.name}`);
      console.log(notification.previewText);
    } else {
      const deliveryResult = await deliverNotification(config, rule, notification);
      if (!deliveryResult.delivered) {
        throw new Error(
          `No notification destinations succeeded for ${rule.name}: ${deliveryResult.errors.join(" | ")}`
        );
      }
    }

    if (!config.dryRun) {
      Object.assign(nextState, notification.nextState);
    }
    sentRuleCount += 1;
    overdueCount += notification.counts.overdue;
    dueSoonCount += notification.counts.dueSoon;
  }

  if (sentRuleCount === 0) {
    console.log("No new reminders to send.");
    return;
  }

  if (!config.dryRun) {
    saveState(nextState);
  }
  console.log(
    `Sent reminders. rules=${sentRuleCount}, overdue=${overdueCount}, dueSoon=${dueSoonCount}`
  );
}

async function deliverNotification(config, rule, notification) {
  const errors = [];
  let delivered = false;

  if (rule.to && config.lineChannelAccessToken) {
    try {
      await sendLinePush(config.lineChannelAccessToken, rule.to, notification.messages);
      delivered = true;
    } catch (error) {
      errors.push(formatDeliveryError("LINE", error));
    }
  }

  const teamsWebhookUrl = rule.teamsWebhookUrl;
  if (teamsWebhookUrl) {
    try {
      await sendTeamsCards(teamsWebhookUrl, notification.teamsCards);
      delivered = true;
    } catch (error) {
      errors.push(formatDeliveryError("Teams", error));
    }
  }

  return { delivered, errors };
}

function formatDeliveryError(target, error) {
  if (error instanceof Error) {
    return `${target}: ${error.message}`;
  }
  return `${target}: ${String(error)}`;
}

export async function queryDepartmentTasks({ department, mode = "dueSoon", hours = 168 }) {
  loadEnvFile();
  validateEnv();

  const target = resolveDepartment(department);
  if (!target) {
    throw new Error(`Unknown department: ${department}`);
  }

  const config = getBaseConfig();
  const ecp = new EcpClient(config);
  await ecp.login();
  if (target.identityId) {
    await ecp.switchIdentity(target.identityId);
  }

  const rule = {
    id: `${target.key}-${mode}`,
    name: `${target.name}-${mode}`,
    title: target.title,
    enabled: true,
    to: "reply",
    schemaId: getSchemaIdForMode(target, mode),
    identityId: target.identityId,
    listId: process.env.ECP_TASK_LIST_ID || DEFAULT_LIST_ID,
    pageSize: Number.parseInt(process.env.ECP_PAGE_SIZE || "200", 10),
    dueSoonHours: Number.parseInt(String(hours), 10),
    executorAllowlist: null
  };

  const tasks = await ecp.getTasks({
    listId: rule.listId,
    schemaId: rule.schemaId,
    pageSize: rule.pageSize
  });

  const result =
    mode === "open"
      ? classifyOpenTasks(filterTasks(tasks, rule.executorAllowlist, { requireDueDate: false }))
      : classifyTasks(filterTasks(tasks, rule.executorAllowlist), rule.dueSoonHours);

  return buildInteractivePayload(result, rule, mode);
}

function loadEnvFile() {
  const envPath = path.resolve(".env");
  if (!fs.existsSync(envPath)) {
    return;
  }

  const content = fs.readFileSync(envPath, "utf8");
  for (const rawLine of content.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) {
      continue;
    }
    const equalsIndex = line.indexOf("=");
    if (equalsIndex <= 0) {
      continue;
    }
    const key = line.slice(0, equalsIndex).trim();
    let value = line.slice(equalsIndex + 1).trim();
    if (
      (value.startsWith("\"") && value.endsWith("\"")) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    if (!(key in process.env)) {
      process.env[key] = value;
    }
  }
}

function validateEnv() {
  const missing = REQUIRED_ENV.filter((key) => !process.env[key]);
  if (missing.length > 0) {
    throw new Error(`Missing required env: ${missing.join(", ")}`);
  }
}

function getBaseConfig() {
  return {
    baseUrl: process.env.ECP_BASE_URL.replace(/\/+$/, ""),
    loginName: process.env.ECP_LOGIN_NAME,
    password: process.env.ECP_PASSWORD
  };
}

function getConfig(options = {}) {
  const defaultRule = {
    name: "project-dept-2",
    title: DEPARTMENT_SCHEMAS.project2.title,
    enabled: true,
    to: process.env.LINE_TO || "",
    schemaId: process.env.ECP_TASK_SCHEMA_ID || DEPARTMENT_SCHEMAS.project2.schemaId,
    identityId: process.env.ECP_TASK_IDENTITY_ID || DEPARTMENT_SCHEMAS.project2.identityId,
    listId: process.env.ECP_TASK_LIST_ID || DEFAULT_LIST_ID,
    pageSize: Number.parseInt(process.env.ECP_PAGE_SIZE || "200", 10),
    dueSoonHours: Number.parseInt(process.env.DUE_SOON_HOURS || "72", 10),
    executorAllowlist: parseList(process.env.EXECUTOR_ALLOWLIST)
  };

  const runSource = getRunSource(options);
  const rules = applyNonSchedulerSafety(loadRules(defaultRule), runSource);
  if (rules.length === 0) {
    throw new Error("No enabled reminder rules configured.");
  }

  return {
    ...getBaseConfig(),
    lineChannelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
    scheduleTimezone: process.env.SCHEDULE_TIMEZONE || "Asia/Taipei",
    dryRun: getDryRunValue(options),
    runSource,
    rules
  };
}

function getRunSource(options = {}) {
  return String(options.runSource || process.env.REMINDER_RUN_SOURCE || "manual");
}

function getDryRunValue(options = {}) {
  if (typeof options.dryRun === "boolean") {
    return options.dryRun;
  }

  if (String(process.env.DRY_RUN || "").toLowerCase() === "true") {
    return true;
  }

  const runSource = getRunSource(options);
  const allowManualLiveSend =
    String(process.env.ALLOW_MANUAL_LIVE_SEND || "false").toLowerCase() === "true";

  if (runSource !== "scheduler" && !allowManualLiveSend) {
    return true;
  }

  return false;
}

function applyNonSchedulerSafety(rules, runSource) {
  if (runSource === "scheduler") {
    return rules;
  }

  const personalLineId = String(process.env.LINE_ALERT_TO || "").trim();
  return rules.map((rule) => ({
    ...rule,
    to: personalLineId,
    teamsWebhookUrl: ""
  }));
}

function loadRules(defaultRule) {
  const fromFile = process.env.REMINDER_RULES_FILE;
  const fromJson = process.env.REMINDER_RULES_JSON;
  let rawRules = null;

  if (fromFile) {
    rawRules = readJsonFile(path.resolve(fromFile));
  } else if (fromJson) {
    rawRules = JSON.parse(fromJson);
  }

  if (!Array.isArray(rawRules)) {
    return defaultRule.to || defaultRule.teamsWebhookUrl ? [normalizeRule(defaultRule, 0)] : [];
  }

  return rawRules
    .map((rule, index) => normalizeRule({ ...defaultRule, ...rule }, index))
    .filter((rule) => rule.enabled);
}

function normalizeRule(rule, index) {
  const enabled = Boolean(rule.enabled ?? true);
  const teamsWebhookEnv = String(rule.teamsWebhookEnv || "").trim();
  const normalized = {
    id: String(rule.id || rule.name || `rule-${index + 1}`),
    name: String(rule.name || `rule-${index + 1}`),
    title: String(rule.title || rule.name || "ECP 任務提醒"),
    enabled,
    to: String(rule.to || "").trim(),
    teamsWebhookUrl: String(rule.teamsWebhookUrl || process.env[teamsWebhookEnv] || "").trim(),
    schemaId: String(rule.schemaId || "").trim(),
    identityId: String(rule.identityId || "").trim(),
    listId: String(rule.listId || "").trim(),
    pageSize: Number.parseInt(String(rule.pageSize || "200"), 10),
    dueSoonHours: Number.parseInt(String(rule.dueSoonHours || "72"), 10),
    executorAllowlist: parseRuleAllowlist(rule.executorAllowlist),
    scheduleWeekdays: parseRuleWeekdays(rule.scheduleWeekdays)
  };

  if (!normalized.enabled) {
    return normalized;
  }
  if (!normalized.to && !normalized.teamsWebhookUrl) {
    throw new Error(`Reminder rule ${normalized.name} must define "to" or "teamsWebhookUrl".`);
  }
  if (!normalized.schemaId) {
    throw new Error(`Reminder rule ${normalized.name} is missing "schemaId".`);
  }
  if (!normalized.listId) {
    throw new Error(`Reminder rule ${normalized.name} is missing "listId".`);
  }

  return normalized;
}

function parseRuleAllowlist(value) {
  if (!value) {
    return null;
  }
  if (Array.isArray(value)) {
    const items = value.map((item) => String(item).trim()).filter(Boolean);
    return items.length > 0 ? new Set(items) : null;
  }
  return parseList(String(value));
}

function parseRuleWeekdays(value) {
  if (!value) {
    return null;
  }
  if (Array.isArray(value)) {
    const items = value
      .map((item) => Number.parseInt(String(item), 10))
      .filter((item) => item >= 1 && item <= 7);
    return items.length > 0 ? new Set(items) : null;
  }

  const items = String(value)
    .split(",")
    .map((item) => Number.parseInt(item.trim(), 10))
    .filter((item) => item >= 1 && item <= 7);
  return items.length > 0 ? new Set(items) : null;
}

function shouldRunRuleToday(rule, timeZone) {
  if (!isDgpaWorkday(new Date(), timeZone)) {
    return false;
  }
  if (!rule.scheduleWeekdays || rule.scheduleWeekdays.size === 0) {
    return true;
  }
  return rule.scheduleWeekdays.has(getWeekdayNumber(new Date(), timeZone));
}

function isDgpaWorkday(date, timeZone) {
  const weekday = getWeekdayNumber(date, timeZone);
  if (weekday === 6 || weekday === 7) {
    return false;
  }

  const dateKey = formatDateInTimezone(date, timeZone);
  const year = Number.parseInt(dateKey.slice(0, 4), 10);
  const nonWorkdays = DGPA_NON_WORKDAYS[year];
  if (!nonWorkdays) {
    return true;
  }

  return !nonWorkdays.has(dateKey);
}

function formatDateInTimezone(date, timeZone) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(date);

  const map = {};
  for (const part of parts) {
    if (part.type !== "literal") {
      map[part.type] = part.value;
    }
  }

  return `${map.year}-${map.month}-${map.day}`;
}

function readJsonFile(filePath) {
  const text = fs.readFileSync(filePath, "utf8").replace(/^\uFEFF/, "");
  return JSON.parse(text);
}

function parseList(value) {
  if (!value) {
    return null;
  }
  const items = value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
  return items.length > 0 ? new Set(items) : null;
}

class EcpClient {
  constructor(config) {
    this.config = config;
    this.cookie = "";
  }

  async login() {
    await this.fetchText("/Qs.OnlineUser.Login.page");
    const publicKeyResponse = await this.postJson("/Qs.Misc.getLoginPublicKey.data", {});
    const encryptedPassword = encryptPassword(publicKeyResponse.publicKey, this.config.password);

    await this.postJson("/Qs.OnlineUser.login.data", {
      loginName: this.config.loginName,
      password: encryptedPassword,
      language: "zh-tw",
      extraArgs: null,
      checkRelogin: true
    });
  }

  async switchIdentity(identityId) {
    await this.postJson("/Qs.OnlineUser.switchIdentity.data", { identityId });
  }

  async getTasks({ listId, schemaId, pageSize }) {
    const result = await this.postJson("/qsvd-list/Ecp.Task.getListData.data", {
      listId,
      schemaId,
      pageNo: 1,
      pageSize
    });
    return result?.data?.records || [];
  }

  async fetchText(pathname, options = {}) {
    const response = await fetch(this.config.baseUrl + pathname, {
      ...options,
      headers: this.buildHeaders(options.headers)
    });
    this.captureCookies(response);
    if (!response.ok) {
      throw new Error(`ECP request failed: ${response.status} ${pathname}`);
    }
    return response.text();
  }

  async postJson(pathname, body) {
    const text = await this.fetchText(pathname, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(body)
    });
    const json = JSON.parse(text || "{}");
    if (json && json._failed) {
      throw new Error(`ECP API failed for ${pathname}: ${json.message || json.code}`);
    }
    return json;
  }

  buildHeaders(extraHeaders) {
    const headers = {
      Accept: "application/json, text/plain, */*",
      ...extraHeaders
    };
    if (this.cookie) {
      headers.Cookie = this.cookie;
    }
    return headers;
  }

  captureCookies(response) {
    const setCookies =
      typeof response.headers.getSetCookie === "function"
        ? response.headers.getSetCookie()
        : response.headers.get("set-cookie")
          ? [response.headers.get("set-cookie")]
          : [];
    if (setCookies.length === 0) {
      return;
    }

    const jar = new Map();
    if (this.cookie) {
      for (const item of this.cookie.split(/;\s*/)) {
        const [name, ...rest] = item.split("=");
        jar.set(name, rest.join("="));
      }
    }

    for (const cookieText of setCookies) {
      const first = cookieText.split(";")[0];
      const [name, ...rest] = first.split("=");
      jar.set(name.trim(), rest.join("=").trim());
    }

    this.cookie = [...jar.entries()].map(([k, v]) => `${k}=${v}`).join("; ");
  }
}

function encryptPassword(publicKeyBase64, password) {
  const pem = [
    "-----BEGIN PUBLIC KEY-----",
    publicKeyBase64,
    "-----END PUBLIC KEY-----"
  ].join("\n");

  return crypto.publicEncrypt(
    {
      key: pem,
      padding: crypto.constants.RSA_PKCS1_PADDING
    },
    Buffer.from(password, "utf8")
  ).toString("base64");
}

async function queryRule(ecp, rule) {
  const tasks = await ecp.getTasks({
    listId: rule.listId,
    schemaId: rule.schemaId,
    pageSize: rule.pageSize
  });

  const filteredTasks = filterTasks(tasks, rule.executorAllowlist);
  return classifyTasks(filteredTasks, rule.dueSoonHours);
}

function filterTasks(tasks, executorAllowlist, options = {}) {
  const requireDueDate = options.requireDueDate !== false;
  return tasks.filter((task) => {
    if (!task.FId) {
      return false;
    }
    if (requireDueDate && !task.FPredictEndDate) {
      return false;
    }
    if (CLOSED_STATUSES.has(task.FStatus)) {
      return false;
    }
    if (executorAllowlist && !executorAllowlist.has(task["FUserId$"])) {
      return false;
    }
    return true;
  });
}

function classifyTasks(tasks, dueSoonHours) {
  const now = new Date();
  const dueSoonMs = dueSoonHours * 60 * 60 * 1000;
  const overdue = [];
  const dueSoon = [];
  const open = [];

  for (const task of tasks) {
    const dueDate = parseEcpDate(task.FPredictEndDate);
    if (!dueDate) {
      continue;
    }
    const taskWithDue = { ...task, __dueDate: dueDate };
    open.push(taskWithDue);
    const diffMs = dueDate.getTime() - now.getTime();
    if (diffMs < 0) {
      overdue.push(taskWithDue);
    }
    if (diffMs <= dueSoonMs) {
      dueSoon.push(taskWithDue);
    }
  }

  overdue.sort(sortByDueDate);
  dueSoon.sort(sortByDueDate);
  open.sort(sortByDueDate);
  return { overdue, dueSoon, open, now };
}

function classifyOpenTasks(tasks) {
  const now = new Date();
  const open = tasks
    .map((task) => ({
      ...task,
      __dueDate: parseEcpDate(task.FPredictEndDate)
    }))
    .sort((a, b) => {
      const aDue = a.__dueDate ? a.__dueDate.getTime() : Number.MAX_SAFE_INTEGER;
      const bDue = b.__dueDate ? b.__dueDate.getTime() : Number.MAX_SAFE_INTEGER;
      return aDue - bDue;
    });

  const overdue = open.filter((task) => task.__dueDate && task.__dueDate.getTime() < now.getTime());
  return { overdue, dueSoon: [], open, now };
}

function parseEcpDate(value) {
  if (!value) {
    return null;
  }
  const date = new Date(value.replace(" ", "T"));
  return Number.isNaN(date.getTime()) ? null : date;
}

function sortByDueDate(a, b) {
  const aDue = a.__dueDate ? a.__dueDate.getTime() : Number.MAX_SAFE_INTEGER;
  const bDue = b.__dueDate ? b.__dueDate.getTime() : Number.MAX_SAFE_INTEGER;
  return aDue - bDue;
}

function buildNotification(summary, state, rule) {
  const todayKey = formatDateKey(summary.now);
  const newOverdue = [];
  const newDueSoon = [];
  const nextState = {};

  for (const task of summary.overdue) {
    const stateKey = buildStateKey(rule, task, "overdue");
    const record = state[stateKey];
    if (record?.lastSentDate === todayKey) {
      continue;
    }
    newOverdue.push(task);
    nextState[stateKey] = {
      ruleId: rule.id,
      type: "overdue",
      lastSentDate: todayKey,
      dueAt: task.FPredictEndDate
    };
  }

  for (const task of summary.dueSoon) {
    const stateKey = buildStateKey(rule, task, "soon");
    if (state[stateKey]) {
      continue;
    }
    newDueSoon.push(task);
    nextState[stateKey] = {
      ruleId: rule.id,
      type: "soon",
      lastSentDate: todayKey,
      dueAt: task.FPredictEndDate
    };
  }

  if (newOverdue.length === 0 && newDueSoon.length === 0) {
    return null;
  }

  const dueSoonOnly = newDueSoon.filter(
    (task) => !newOverdue.some((overdueTask) => overdueTask.FId === task.FId)
  );

  const payload = buildLinePayload({
    now: summary.now,
    ruleTitle: rule.title,
    subtitle: "\u8fd17\u5929\u5230\u671f\u4efb\u52d9",
    overdue: newOverdue,
    dueSoon: dueSoonOnly
  });

  return {
    previewText: payload.previewText,
    messages: payload.messages,
    teamsCards: payload.teamsCards,
    nextState,
    counts: {
      overdue: newOverdue.length,
      dueSoon: dueSoonOnly.length
    }
  };
}

function buildInteractivePayload(summary, rule, mode) {
  const selection = selectTasksByMode(summary, mode);
  const subtitle = getSubtitleByMode(mode);
  const payload = buildLinePayload({
    now: summary.now,
    ruleTitle: rule.title,
    subtitle,
    overdue: selection.overdue,
    dueSoon: selection.dueSoon
  });

  return {
    mode,
    counts: {
      overdue: selection.overdue.length,
      dueSoon: selection.dueSoon.length,
      open: selection.open.length
    },
    messages: payload.messages
  };
}

function selectTasksByMode(summary, mode) {
  if (mode === "overdue") {
    return {
      overdue: summary.overdue,
      dueSoon: [],
      open: summary.overdue
    };
  }

  if (mode === "open") {
    return {
      overdue: summary.overdue,
      dueSoon: summary.open.filter((task) => !summary.overdue.some((item) => item.FId === task.FId)),
      open: summary.open
    };
  }

  return {
    overdue: summary.overdue,
    dueSoon: summary.dueSoon.filter((task) => !summary.overdue.some((item) => item.FId === task.FId)),
    open: summary.dueSoon
  };
}

function getSubtitleByMode(mode) {
  if (mode === "overdue") {
    return "\u903e\u671f\u4efb\u52d9";
  }
  if (mode === "open") {
    return "\u672a\u5b8c\u6210\u4efb\u52d9";
  }
  return "\u8fd17\u5929\u5230\u671f\u4efb\u52d9";
}

function buildStateKey(rule, task, type) {
  return `${rule.id}:${rule.to}:${type}:${task.FId}:${task.FPredictEndDate || "none"}`;
}

function loadState() {
  if (!fs.existsSync(STATE_FILE)) {
    return {};
  }
  return JSON.parse(fs.readFileSync(STATE_FILE, "utf8") || "{}");
}

function saveState(state) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(STATE_FILE, JSON.stringify(state, null, 2), "utf8");
}

function formatDateKey(date) {
  return date.toISOString().slice(0, 10);
}

function buildLinePayload({ now, ruleTitle, subtitle, overdue, dueSoon }) {
  const previewText = formatPlainTextMessage({ now, ruleTitle, subtitle, overdue, dueSoon });
  const taskItems = [
    ...overdue.map((task) => ({ task, overdue: true })),
    ...dueSoon.map((task) => ({ task, overdue: false }))
  ];

  if (taskItems.length === 0) {
    return {
      previewText,
      teamsCards: [buildTeamsEmptyCard({ now, ruleTitle, subtitle })],
      messages: [
        {
          type: "text",
          text: previewText
        }
      ]
    };
  }

  const bubbles = chunkArray(taskItems, 4).map((items, index, all) =>
    buildFlexBubble({
      now,
      items,
      page: index + 1,
      totalPages: all.length,
      overdueCount: overdue.length,
      dueSoonCount: dueSoon.length,
      ruleTitle,
      subtitle
    })
  );

  const messages = bubbles.map((bubble) => ({
    type: "flex",
    altText: previewText.slice(0, 400),
    contents: bubble
  }));

  const teamsCards = chunkArray(taskItems, 4).map((items, index, all) =>
    buildTeamsCard({
      now,
      items,
      page: index + 1,
      totalPages: all.length,
      overdueCount: overdue.length,
      dueSoonCount: dueSoon.length,
      ruleTitle,
      subtitle
    })
  );

  return { previewText, messages, teamsCards };
}

function formatPlainTextMessage({ now, ruleTitle, subtitle, overdue, dueSoon }) {
  const lines = [];
  lines.push(ruleTitle);
  lines.push(subtitle);
  lines.push(`\u6642\u9593: ${formatTimestamp(now)}`);
  lines.push(`\u903e\u671f: ${overdue.length}  \u5176\u4ed6: ${dueSoon.length}`);

  for (const task of [...overdue, ...dueSoon].slice(0, 12)) {
    lines.push(formatTaskLine(task));
  }

  if (overdue.length + dueSoon.length === 0) {
    lines.push("\u6c92\u6709\u7b26\u5408\u689d\u4ef6\u7684\u4efb\u52d9\u3002");
  }

  return lines.join("\n");
}

function formatTaskLine(task) {
  const serial = task.FSerialNumber || "(\u7121\u7de8\u865f)";
  const name = task.FName || "(\u7121\u540d\u7a31)";
  const executor = task["FUserId$"] || "(\u7121\u57f7\u884c\u4eba)";
  const dueAt = task.FPredictEndDate || "(\u7121\u65e5\u671f)";
  const status = task["FStatus$"] || task.FStatus || "(\u7121\u72c0\u614b)";
  return `- ${serial} | ${name} | ${executor} | ${dueAt} | ${status}`;
}

function buildFlexBubble({
  now,
  items,
  page,
  totalPages,
  overdueCount,
  dueSoonCount,
  ruleTitle,
  subtitle
}) {
  return {
    type: "bubble",
    size: "giga",
    body: {
      type: "box",
      layout: "vertical",
      spacing: "md",
      contents: [
        {
          type: "text",
          text: ruleTitle,
          weight: "bold",
          size: "lg",
          wrap: true
        },
        {
          type: "text",
          text: `${subtitle}  ${page}/${totalPages}`,
          size: "sm",
          color: "#ef6c00",
          weight: "bold",
          wrap: true
        },
        {
          type: "box",
          layout: "baseline",
          spacing: "md",
          contents: [
            {
              type: "text",
              text: `\u6642\u9593 ${formatTimestamp(now)}`,
              size: "xs",
              color: "#666666",
              flex: 3
            },
            {
              type: "text",
              text: `\u903e\u671f ${overdueCount}  \u5176\u4ed6 ${dueSoonCount}`,
              size: "xs",
              color: "#666666",
              align: "end",
              flex: 2
            }
          ]
        },
        {
          type: "separator",
          margin: "sm"
        },
        ...items.flatMap((item, index) => {
          const row = buildTaskBox(item);
          if (index === items.length - 1) {
            return [row];
          }
          return [row, { type: "separator", margin: "md" }];
        })
      ]
    }
  };
}

function buildTaskBox({ task, overdue }) {
  const serial = task.FSerialNumber || "-";
  const name = task.FName || "-";
  const executor = task["FUserId$"] || "-";
  const dueAt = task.FPredictEndDate || "-";
  const status = task["FStatus$"] || task.FStatus || "-";

  return {
    type: "box",
    layout: "vertical",
    spacing: "sm",
    margin: "md",
    contents: [
      {
        type: "box",
        layout: "baseline",
        contents: [
          {
            type: "text",
            text: status,
            size: "xs",
            color: overdue ? "#d32f2f" : "#666666",
            weight: overdue ? "bold" : "regular",
            align: "end",
            flex: 1
          }
        ]
      },
      {
        type: "text",
        text: name,
        size: "sm",
        weight: "bold",
        wrap: true
      },
      {
        type: "text",
        text: serial,
        size: "xs",
        color: "#666666",
        wrap: true
      },
      {
        type: "box",
        layout: "baseline",
        spacing: "md",
        contents: [
          {
            type: "text",
            text: executor,
            size: "sm",
            color: "#1565c0",
            weight: "bold",
            flex: 2,
            wrap: true
          },
          {
            type: "text",
            text: dueAt,
            size: "sm",
            color: "#c62828",
            align: "end",
            flex: 3,
            wrap: true
          }
        ]
      }
    ]
  };
}

function buildTeamsEmptyCard({ now, ruleTitle, subtitle }) {
  return {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: ruleTitle,
        weight: "Bolder",
        size: "Large",
        wrap: true
      },
      {
        type: "TextBlock",
        text: subtitle,
        color: "Warning",
        weight: "Bolder",
        spacing: "Small",
        wrap: true
      },
      {
        type: "TextBlock",
        text: `時間 ${formatTimestamp(now)}`,
        isSubtle: true,
        spacing: "Small",
        wrap: true
      },
      {
        type: "TextBlock",
        text: "目前沒有符合條件的任務。",
        spacing: "Medium",
        wrap: true
      }
    ]
  };
}

function buildTeamsCard({
  now,
  items,
  page,
  totalPages,
  overdueCount,
  dueSoonCount,
  ruleTitle,
  subtitle
}) {
  return {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: ruleTitle,
        weight: "Bolder",
        size: "Large",
        wrap: true
      },
      {
        type: "TextBlock",
        text: `${subtitle}  ${page}/${totalPages}`,
        color: "Warning",
        weight: "Bolder",
        spacing: "Small",
        wrap: true
      },
      {
        type: "ColumnSet",
        spacing: "Small",
        columns: [
          {
            type: "Column",
            width: "stretch",
            items: [
              {
                type: "TextBlock",
                text: `時間 ${formatTimestamp(now)}`,
                isSubtle: true,
                wrap: true
              }
            ]
          },
          {
            type: "Column",
            width: "auto",
            items: [
              {
                type: "TextBlock",
                text: `逾期 ${overdueCount}  其他 ${dueSoonCount}`,
                isSubtle: true,
                horizontalAlignment: "Right",
                wrap: true
              }
            ]
          }
        ]
      },
      ...items.flatMap((item, index) => {
        const cardItem = buildTeamsTaskBlock(item);
        if (index === items.length - 1) {
          return [cardItem];
        }
        return [
          cardItem,
          {
            type: "TextBlock",
            text: " ",
            separator: true,
            spacing: "Medium"
          }
        ];
      })
    ]
  };
}

function buildTeamsTaskBlock({ task, overdue }) {
  const serial = task.FSerialNumber || "-";
  const name = task.FName || "-";
  const executor = task["FUserId$"] || "-";
  const dueAt = task.FPredictEndDate || "-";
  const status = task["FStatus$"] || task.FStatus || "-";

  return {
    type: "Container",
    spacing: "Medium",
    items: [
      {
        type: "ColumnSet",
        columns: [
          {
            type: "Column",
            width: "stretch",
            items: [
              {
                type: "TextBlock",
                text: name,
                weight: "Bolder",
                wrap: true
              }
            ]
          },
          {
            type: "Column",
            width: "auto",
            items: [
              {
                type: "TextBlock",
                text: status,
                color: overdue ? "Attention" : "Default",
                weight: overdue ? "Bolder" : "Default",
                horizontalAlignment: "Right",
                wrap: true
              }
            ]
          }
        ]
      },
      {
        type: "TextBlock",
        text: serial,
        isSubtle: true,
        spacing: "Small",
        wrap: true
      },
      {
        type: "ColumnSet",
        spacing: "Small",
        columns: [
          {
            type: "Column",
            width: "stretch",
            items: [
              {
                type: "TextBlock",
                text: executor,
                color: "Accent",
                weight: "Bolder",
                wrap: true
              }
            ]
          },
          {
            type: "Column",
            width: "auto",
            items: [
              {
                type: "TextBlock",
                text: dueAt,
                color: "Attention",
                horizontalAlignment: "Right",
                wrap: true
              }
            ]
          }
        ]
      }
    ]
  };
}

function chunkArray(items, size) {
  const result = [];
  for (let index = 0; index < items.length; index += size) {
    result.push(items.slice(index, index + size));
  }
  return result;
}

function formatTimestamp(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mi = String(date.getMinutes()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd} ${hh}:${mi}`;
}

function getWeekdayNumber(date, timeZone) {
  const weekday = new Intl.DateTimeFormat("en-US", {
    timeZone,
    weekday: "short"
  }).format(date);

  const map = {
    Mon: 1,
    Tue: 2,
    Wed: 3,
    Thu: 4,
    Fri: 5,
    Sat: 6,
    Sun: 7
  };

  return map[weekday];
}

function resolveDepartment(value) {
  const text = String(value || "").replace(/\s+/g, "").toLowerCase();
  if (!text) {
    return null;
  }

  const aliases = [
    {
      key: "project1",
      names: ["\u5c08\u6848\u4e00\u90e8", "\u4e00\u90e8", "project1", "p1"]
    },
    {
      key: "project2",
      names: ["\u5c08\u6848\u4e8c\u90e8", "\u4e8c\u90e8", "project2", "p2"]
    },
    {
      key: "cloud",
      names: ["\u96f2\u7aef\u670d\u52d9\u90e8", "\u96f2\u7aef", "cloud", "cloudservice"]
    }
  ];

  for (const item of aliases) {
    if (item.names.some((name) => text.includes(name.toLowerCase()))) {
      return { key: item.key, ...DEPARTMENT_SCHEMAS[item.key] };
    }
  }
  return null;
}

function getSchemaIdForMode(target, mode) {
  if ((mode === "dueSoon" || mode === "overdue") && target.dueSoonSchemaId) {
    return target.dueSoonSchemaId;
  }
  if (mode === "open" && target.openSchemaId) {
    return target.openSchemaId;
  }
  return target.schemaId;
}

async function sendLinePush(channelAccessToken, to, messages) {
  const validMessages = Array.isArray(messages) ? messages.filter(Boolean) : [];
  if (validMessages.length === 0) {
    throw new Error("LINE push failed: no messages to send");
  }

  for (const messageBatch of chunkArray(validMessages, 5)) {
    const response = await fetch("https://api.line.me/v2/bot/message/push", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${channelAccessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        to,
        messages: messageBatch
      })
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`LINE push failed: ${response.status} ${body}`);
    }
  }
}

async function sendTeamsCards(webhookUrl, cards) {
  const validCards = Array.isArray(cards) ? cards.filter(Boolean) : [];
  if (validCards.length === 0) {
    throw new Error("Teams webhook failed: no cards to send");
  }

  for (const card of validCards) {
    const response = await fetch(webhookUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(card)
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`Teams webhook failed: ${response.status} ${body}`);
    }
  }
}

export async function sendLineText(channelAccessToken, to, text) {
  await sendLinePush(channelAccessToken, to, [
    {
      type: "text",
      text: String(text || "").slice(0, 5000)
    }
  ]);
}

const isDirectRun =
  process.argv[1] &&
  import.meta.url === pathToFileURL(process.argv[1]).href;

if (isDirectRun) {
  runReminderJob().catch((error) => {
    console.error(error instanceof Error ? error.stack : error);
    process.exitCode = 1;
  });
}
