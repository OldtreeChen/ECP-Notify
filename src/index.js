import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

const REQUIRED_ENV = [
  "ECP_BASE_URL",
  "ECP_LOGIN_NAME",
  "ECP_PASSWORD",
  "LINE_CHANNEL_ACCESS_TOKEN"
];

const CLOSED_STATUSES = new Set(["Finished", "Cancel", "Discarded"]);
const DATA_DIR = path.resolve(".data");
const STATE_FILE = path.join(DATA_DIR, "notifications.json");

export async function runReminderJob(options = {}) {
  loadEnvFile();
  validateEnv();

  const config = getConfig(options);
  const ecp = new EcpClient(config);
  await ecp.login();

  const state = loadState();
  const nextState = structuredClone(state);
  let sentRuleCount = 0;
  let overdueCount = 0;
  let dueSoonCount = 0;

  for (const rule of config.rules) {
    const tasks = await ecp.getTasks({
      listId: rule.listId,
      schemaId: rule.schemaId,
      pageSize: rule.pageSize
    });

    const filteredTasks = filterTasks(tasks, rule.executorAllowlist);
    const summary = classifyTasks(filteredTasks, rule.dueSoonHours);
    const notification = buildNotification(summary, nextState, rule);

    if (!notification) {
      continue;
    }

    if (config.dryRun) {
      console.log(`[DRY RUN] ${rule.name}`);
      console.log(notification.previewText);
    } else {
      await sendLinePush(config.lineChannelAccessToken, rule.to, notification.messages);
    }

    Object.assign(nextState, notification.nextState);
    sentRuleCount += 1;
    overdueCount += notification.counts.overdue;
    dueSoonCount += notification.counts.dueSoon;
  }

  if (sentRuleCount === 0) {
    console.log("No new reminders to send.");
    return;
  }

  saveState(nextState);
  console.log(
    `Sent LINE reminder. rules=${sentRuleCount}, overdue=${overdueCount}, dueSoon=${dueSoonCount}`
  );
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

function getConfig(options = {}) {
  const defaultRule = {
    name: "project-dept-2",
    title: "ECP 專案二部任務",
    to: process.env.LINE_TO || "",
    schemaId: process.env.ECP_TASK_SCHEMA_ID || "ffffff19-cc0e-1122-5001-6006b23cf204",
    listId: process.env.ECP_TASK_LIST_ID || "296aa935-f6c0-4a8e-9ab9-32254ea39861",
    pageSize: Number.parseInt(process.env.ECP_PAGE_SIZE || "200", 10),
    dueSoonHours: Number.parseInt(process.env.DUE_SOON_HOURS || "72", 10),
    executorAllowlist: parseList(process.env.EXECUTOR_ALLOWLIST)
  };

  const rules = loadRules(defaultRule);
  if (rules.length === 0) {
    throw new Error("No reminder rules configured. Set LINE_TO or provide REMINDER_RULES_FILE.");
  }

  return {
    baseUrl: process.env.ECP_BASE_URL.replace(/\/+$/, ""),
    loginName: process.env.ECP_LOGIN_NAME,
    password: process.env.ECP_PASSWORD,
    lineChannelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
    lineAlertTo: process.env.LINE_ALERT_TO || process.env.LINE_TO || "",
    runSource: String(options.runSource || process.env.REMINDER_RUN_SOURCE || "manual"),
    dryRun: getDryRunValue(options),
    rules
  };
}

function getDryRunValue(options = {}) {
  if (typeof options.dryRun === "boolean") {
    return options.dryRun;
  }

  if (String(process.env.DRY_RUN || "").toLowerCase() === "true") {
    return true;
  }

  const runSource = String(options.runSource || process.env.REMINDER_RUN_SOURCE || "manual");
  const allowManualLiveSend =
    String(process.env.ALLOW_MANUAL_LIVE_SEND || "false").toLowerCase() === "true";

  if (runSource !== "scheduler" && !allowManualLiveSend) {
    return true;
  }

  return false;
}

function loadRules(defaultRule) {
  const fromFile = process.env.REMINDER_RULES_FILE;
  const fromJson = process.env.REMINDER_RULES_JSON;
  let rawRules = null;

  if (fromFile) {
    const filePath = path.resolve(fromFile);
    rawRules = JSON.parse(fs.readFileSync(filePath, "utf8"));
  } else if (fromJson) {
    rawRules = JSON.parse(fromJson);
  }

  if (!Array.isArray(rawRules)) {
    return defaultRule.to ? [normalizeRule(defaultRule, 0)] : [];
  }

  return rawRules.map((rule, index) => normalizeRule({ ...defaultRule, ...rule }, index));
}

function normalizeRule(rule, index) {
  const normalized = {
    id: String(rule.id || rule.name || `rule-${index + 1}`),
    name: String(rule.name || `rule-${index + 1}`),
    title: String(rule.title || rule.name || "ECP 任務提醒"),
    to: String(rule.to || "").trim(),
    schemaId: String(rule.schemaId || "").trim(),
    listId: String(rule.listId || "").trim(),
    pageSize: Number.parseInt(String(rule.pageSize || "200"), 10),
    dueSoonHours: Number.parseInt(String(rule.dueSoonHours || "72"), 10),
    executorAllowlist: parseRuleAllowlist(rule.executorAllowlist)
  };

  if (!normalized.to) {
    throw new Error(`Reminder rule ${normalized.name} is missing "to".`);
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

function filterTasks(tasks, executorAllowlist) {
  return tasks.filter((task) => {
    if (!task.FId || !task.FPredictEndDate) {
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

  for (const task of tasks) {
    const dueDate = parseEcpDate(task.FPredictEndDate);
    if (!dueDate) {
      continue;
    }
    const diffMs = dueDate.getTime() - now.getTime();
    if (diffMs < 0) {
      overdue.push(task);
    } else if (diffMs <= dueSoonMs) {
      dueSoon.push(task);
    }
  }

  overdue.sort(sortByDueDate);
  dueSoon.sort(sortByDueDate);
  return { overdue, dueSoon, now };
}

function parseEcpDate(value) {
  if (!value) {
    return null;
  }
  const date = new Date(value.replace(" ", "T"));
  return Number.isNaN(date.getTime()) ? null : date;
}

function sortByDueDate(a, b) {
  return String(a.FPredictEndDate).localeCompare(String(b.FPredictEndDate));
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

  const payload = buildLinePayload({
    now: summary.now,
    overdue: newOverdue,
    dueSoon: newDueSoon,
    rule
  });

  return {
    previewText: payload.previewText,
    messages: payload.messages,
    nextState,
    counts: {
      overdue: newOverdue.length,
      dueSoon: newDueSoon.length
    }
  };
}

function buildStateKey(rule, task, type) {
  return `${rule.id}:${rule.to}:${type}:${task.FId}:${task.FPredictEndDate}`;
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

function buildLinePayload({ now, overdue, dueSoon, rule }) {
  const previewText = formatPlainTextMessage({ now, overdue, dueSoon, rule });
  const taskItems = [
    ...overdue.map((task) => ({ task, overdue: true })),
    ...dueSoon.map((task) => ({ task, overdue: false }))
  ];

  const bubbles = chunkArray(taskItems, 4).map((items, index, all) =>
    buildFlexBubble({
      now,
      items,
      page: index + 1,
      totalPages: all.length,
      overdueCount: overdue.length,
      dueSoonCount: dueSoon.length,
      rule
    })
  );

  const messages = bubbles.map((bubble) => ({
    type: "flex",
    altText: previewText.slice(0, 400),
    contents: bubble
  }));

  return { previewText, messages };
}

function formatPlainTextMessage({ now, overdue, dueSoon, rule }) {
  const lines = [];
  lines.push(rule.title);
  lines.push(`時間: ${formatTimestamp(now)}`);
  lines.push(`總計: ${overdue.length + dueSoon.length}`);

  if (overdue.length > 0) {
    lines.push("");
    lines.push(`已逾期: ${overdue.length}`);
    for (const task of overdue.slice(0, 10)) {
      lines.push(formatTaskLine(task));
    }
  }

  if (dueSoon.length > 0) {
    lines.push("");
    lines.push(`近7天到期: ${dueSoon.length}`);
    for (const task of dueSoon.slice(0, 10)) {
      lines.push(formatTaskLine(task));
    }
  }

  return lines.join("\n");
}

function formatTaskLine(task) {
  const serial = task.FSerialNumber || "(無編號)";
  const name = task.FName || "(無名稱)";
  const executor = task["FUserId$"] || "(無執行人)";
  const dueAt = task.FPredictEndDate || "(無日期)";
  const status = task["FStatus$"] || task.FStatus || "(無狀態)";
  return `- ${serial} | ${name} | ${executor} | ${dueAt} | ${status}`;
}

function buildFlexBubble({ now, items, page, totalPages, overdueCount, dueSoonCount, rule }) {
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
          text: rule.title,
          weight: "bold",
          size: "lg",
          wrap: true
        },
        {
          type: "text",
          text: `近7天到期任務  ${page}/${totalPages}`,
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
              text: `時間 ${formatTimestamp(now)}`,
              size: "xs",
              color: "#666666",
              flex: 3
            },
            {
              type: "text",
              text: `逾期 ${overdueCount}  到期 ${dueSoonCount}`,
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

async function sendLinePush(channelAccessToken, to, messages) {
  const response = await fetch("https://api.line.me/v2/bot/message/push", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${channelAccessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      to,
      messages
    })
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`LINE push failed: ${response.status} ${body}`);
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
