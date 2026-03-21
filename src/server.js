import "dotenv/config";
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import express from "express";
import { runReminderJob, sendLineText } from "./index.js";

const DATA_DIR = path.resolve(".data");
const SOURCE_FILE = path.join(DATA_DIR, "line-sources.json");

const app = express();
const port = Number(process.env.PORT || 3000);
const host = process.env.HOST || "0.0.0.0";
const scheduleTimezone = process.env.SCHEDULE_TIMEZONE || "Asia/Taipei";
const scheduleHour = Number(process.env.SCHEDULE_HOUR || "10");
const scheduleMinute = Number(process.env.SCHEDULE_MINUTE || "0");
const scheduleWeekdays = new Set(["1", "2", "3", "4", "5"]);

let lastScheduleKey = null;

app.post(
  ["/line/webhook", "/webhook/line"],
  express.json({
    verify: (req, _res, buffer) => {
      req.rawBody = buffer;
    }
  }),
  async (req, res) => {
    try {
      const channelSecret = process.env.LINE_CHANNEL_SECRET || "";
      const channelAccessToken = process.env.LINE_CHANNEL_ACCESS_TOKEN || "";
      const signature = req.get("x-line-signature");

      if (!verifyLineSignature(req.rawBody, signature, channelSecret)) {
        res.status(401).json({ ok: false, error: "Invalid LINE signature" });
        return;
      }

      const events = Array.isArray(req.body?.events) ? req.body.events : [];
      res.json({ ok: true });

      for (const event of events) {
        try {
          const sourceInfo = normalizeSource(event.source);
          if (sourceInfo) {
            upsertSource(sourceInfo);
          }

          if (
            sourceInfo?.type === "user" &&
            event.type === "message" &&
            event.message?.type === "text" &&
            event.replyToken
          ) {
            const messageText = buildReplyMessage(event, sourceInfo);
            if (channelAccessToken) {
              await replyMessage(channelAccessToken, event.replyToken, messageText);
            }
          }
        } catch (error) {
          console.error("LINE webhook event failed:", error);
        }
      }
    } catch (error) {
      console.error(error instanceof Error ? error.stack : error);
      if (!res.headersSent) {
        res.status(500).json({ ok: false, error: "Internal server error" });
      }
    }
  }
);

app.get("/health", (_req, res) => {
  res.json({
    ok: true,
    service: "ecp-line-webhook",
    timestamp: new Date().toISOString()
  });
});

app.get("/sources", (_req, res) => {
  res.json(loadSources());
});

app.listen(port, host, () => {
  console.log(`ECP LINE server listening on http://${host}:${port}`);
});

startReminderScheduler();

function verifyLineSignature(rawBody, signature, secret) {
  if (!rawBody || !signature || !secret) {
    return false;
  }

  const expected = crypto.createHmac("sha256", secret).update(rawBody).digest("base64");
  const actualBuffer = Buffer.from(signature);
  const expectedBuffer = Buffer.from(expected);

  if (actualBuffer.length !== expectedBuffer.length) {
    return false;
  }

  return crypto.timingSafeEqual(actualBuffer, expectedBuffer);
}

function normalizeSource(source) {
  if (!source?.type) {
    return null;
  }
  const id = source.userId || source.groupId || source.roomId || null;
  if (!id) {
    return null;
  }
  return {
    id,
    type: source.type,
    userId: source.userId || null,
    groupId: source.groupId || null,
    roomId: source.roomId || null,
    timestamp: new Date().toISOString()
  };
}

function buildReplyMessage(event, sourceInfo) {
  const text = event.message?.text?.trim() || "";
  if (/^id$/i.test(text) || /^whoami$/i.test(text) || /source/i.test(text)) {
    return [
      "LINE source captured.",
      `type: ${sourceInfo?.type || "unknown"}`,
      `id: ${sourceInfo?.id || "unknown"}`
    ].join("\n");
  }

  return [
    "Webhook is working.",
    `type: ${sourceInfo?.type || "unknown"}`,
    `id: ${sourceInfo?.id || "unknown"}`,
    "Reply with `id` to show the target ID again."
  ].join("\n");
}

function loadSources() {
  if (!fs.existsSync(SOURCE_FILE)) {
    return [];
  }
  return JSON.parse(fs.readFileSync(SOURCE_FILE, "utf8") || "[]");
}

function upsertSource(sourceInfo) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  const sources = loadSources();
  const index = sources.findIndex((item) => item.id === sourceInfo.id);
  if (index >= 0) {
    sources[index] = { ...sources[index], ...sourceInfo };
  } else {
    sources.push(sourceInfo);
  }
  fs.writeFileSync(SOURCE_FILE, JSON.stringify(sources, null, 2), "utf8");
}

async function replyMessage(accessToken, replyToken, text) {
  const response = await fetch("https://api.line.me/v2/bot/message/reply", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      replyToken,
      messages: [
        {
          type: "text",
          text: String(text || "").slice(0, 5000)
        }
      ]
    })
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`LINE reply failed: ${response.status} ${body}`);
  }
}

function startReminderScheduler() {
  const tick = async () => {
    const now = new Date();
    const parts = getTimezoneParts(now, scheduleTimezone);
    const weekday = getWeekdayNumber(now, scheduleTimezone);
    const minute = Number(parts.minute);
    const hour = Number(parts.hour);
    const key = `${parts.year}-${parts.month}-${parts.day}-${hour}-${minute}`;

    if (
      scheduleWeekdays.has(String(weekday)) &&
      hour === scheduleHour &&
      minute === scheduleMinute &&
      lastScheduleKey !== key
    ) {
      lastScheduleKey = key;
      try {
        console.log(`Running scheduled reminder for ${scheduleTimezone} ${key}`);
        await runReminderJob({ runSource: "scheduler", dryRun: false });
      } catch (error) {
        console.error("Scheduled reminder failed:", error);
        await notifyReminderFailure(error, key);
      }
    }
  };

  tick();
  setInterval(tick, 30 * 1000);
}

async function notifyReminderFailure(error, scheduleKey) {
  const channelAccessToken = process.env.LINE_CHANNEL_ACCESS_TOKEN || "";
  const alertTo = process.env.LINE_ALERT_TO || "";
  if (!channelAccessToken || !alertTo) {
    return;
  }

  const message = [
    "ECP 專案二部任務提醒失敗",
    `時間: ${formatAlertTimestamp(new Date(), scheduleTimezone)}`,
    `排程: 上班日 ${String(scheduleHour).padStart(2, "0")}:${String(scheduleMinute).padStart(2, "0")}`,
    `批次: ${scheduleKey}`,
    `錯誤: ${formatErrorMessage(error)}`
  ].join("\n");

  try {
    await sendLineText(channelAccessToken, alertTo, message);
  } catch (notifyError) {
    console.error("Reminder failure alert failed:", notifyError);
  }
}

function formatErrorMessage(error) {
  if (error instanceof Error) {
    return error.message.slice(0, 1000);
  }
  return String(error || "Unknown error").slice(0, 1000);
}

function formatAlertTimestamp(date, timeZone) {
  const parts = getTimezoneParts(date, timeZone);
  return `${parts.year}-${parts.month}-${parts.day} ${parts.hour}:${parts.minute}`;
}

function getTimezoneParts(date, timeZone) {
  const formatter = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false
  });

  const map = {};
  for (const part of formatter.formatToParts(date)) {
    if (part.type !== "literal") {
      map[part.type] = part.value;
    }
  }
  return map;
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
