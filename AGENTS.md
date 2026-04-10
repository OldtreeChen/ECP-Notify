# Agent Rules

## Notification Safety

- During development, debugging, and manual testing, do not send any LINE or Teams messages.
- Only scheduled production runs are allowed to send external notifications.
- Any non-scheduler execution must remain `dry-run` and must not override targets to personal or group accounts.
- If a change needs notification verification, inspect payloads and logs locally instead of sending live messages.

## Current Production Targets

- LINE production target: Cloud Service Department group only.
- Teams production target: Project Department 2 only.
- Project Department 2 LINE target has been removed and must not be restored without explicit user approval.
