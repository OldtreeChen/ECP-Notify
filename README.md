# ECP LINE Webhook + Reminder

This project contains two modes and follows the same basic style as your `C:\\Codex\\OpenClaw` project:

- `npm run web`: LINE webhook service for Railway
- `npm run reminder`: ECP overdue / due-soon reminder sender

## Files

- [src/server.js](/C:/Codex/ECP-Notify/src/server.js): LINE webhook server
- [src/index.js](/C:/Codex/ECP-Notify/src/index.js): ECP reminder job
- [.env.example](/C:/Codex/ECP-Notify/.env.example): environment template
- [config/reminder-rules.example.json](/C:/Codex/ECP-Notify/config/reminder-rules.example.json): multi-group rule example

## Required env

```env
ECP_BASE_URL=https://econtact.ai3.cloud/ecp
ECP_LOGIN_NAME=oldtree.chen
ECP_PASSWORD=...
ECP_TASK_SCHEMA_ID=ffffff19-cc0e-1122-5001-6006b23cf204
ECP_TASK_LIST_ID=296aa935-f6c0-4a8e-9ab9-32254ea39861
ECP_PAGE_SIZE=200
DUE_SOON_HOURS=72
REMINDER_RULES_FILE=./config/reminder-rules.json

LINE_CHANNEL_SECRET=...
LINE_CHANNEL_ACCESS_TOKEN=...
LINE_CHANNEL_ID=2009553650
LINE_TO=U... / C... / R...
LINE_ALERT_TO=U...
DRY_RUN=false
ALLOW_MANUAL_LIVE_SEND=false
```

## Multi-group rules

You can define multiple reminder targets in a JSON file. When `REMINDER_RULES_FILE` is set, the job will evaluate every rule independently.

Rule fields:

- `id`: stable identifier for de-duplication
- `name`: readable rule name
- `title`: LINE card title
- `enabled`: whether the rule is active
- `to`: LINE user/group/room id
- `schemaId`: ECP query schema
- `listId`: ECP list id
- `pageSize`: max rows to fetch
- `dueSoonHours`: due-soon window in hours
- `executorAllowlist`: optional executor name array

Each rule keeps its own notification state, so one group sending does not suppress another group.

## Webhook endpoints

- `GET /health`
- `GET /sources`
- `POST /line/webhook`
- `POST /webhook/line` (compatibility alias)

When a user sends a message to the LINE bot, the webhook will:

- verify the LINE signature
- store the source id in `.data/line-sources.json`
- only reply in 1:1 chat

## Local run

Webhook:

```bash
npm run web
```

Reminder:

```bash
npm run reminder
```

Dry run reminder:

```bash
DRY_RUN=true npm run reminder
```

Manual runs are safe by default. Unless `ALLOW_MANUAL_LIVE_SEND=true`, `npm run reminder` will fall back to dry-run and will not push to LINE groups. Railway scheduled runs still send normally.

## Railway

Railway should run:

```bash
npm run web
```

After deployment, set the LINE webhook URL to:

```text
https://<your-railway-domain>/line/webhook
```

Health check:

```text
https://<your-railway-domain>/health
```
