---
name: outlook-calendar
description: Retrieves and summarizes Outlook calendar information and availability (read-only). Use for "what's on my calendar", "show my schedule", date ranges, and free/busy checks.
allowed-tools: Read, mcp__outlook__list_events, mcp__outlook__get_free_busy, mcp__outlook__find_free_slots, mcp__outlook__get_calendars
argument-hint: "[date|date-range] (optional)"
---

# Outlook Calendar (Read-only)

## Preconditions
- If Outlook MCP tools are missing/unavailable, have the user run:
  - `/mcp` to view server status
  - `/mcp enable <server-name>` to start it (commonly `outlook`)
- Read `~/.claude/config.md` to confirm the user's email/timezone.

## Tool naming note
MCP tool names are namespaced as `mcp__<server>__<tool>`.
This skill assumes server name `outlook` (example: `mcp__outlook__list_events`).
If your server name differs, replace accordingly.

## Core tools (typical)
- `mcp__outlook__list_events` (by date range; prefer smallest possible ranges)
- `mcp__outlook__get_free_busy` (availability blocks)
- `mcp__outlook__find_free_slots` (candidate slots)
- `mcp__outlook__get_calendars` (if multiple calendars exist)

## Output rules
- Always show events with:
  - start-end time
  - subject
  - busy status (Busy/Tentative/Free/OOF)
  - location (if present)
- Group by date, then "Morning / Afternoon / Evening".
- Use absolute dates when the user uses relative terms ("next Tuesday").

## Workflows

### 1) Show calendar for a specific day
1. Call `mcp__outlook__list_events` for that single day (startDate=endDate in MM/DD/YYYY if required by the server).
2. Format as:

**Morning**
- 9:00-9:30: Standup [Busy] @ Teams

**Afternoon**
- 1:00-2:00: Project review [Tentative] @ Conf Room

### 2) Show calendar for a date range
- Prefer splitting into smaller chunks (day-by-day) if responses get large.

### 3) Check availability (user + others)
- Use `mcp__outlook__get_free_busy` across all attendees.
- Present:
  - Candidate slots
  - Who is Free/Tentative/Busy per slot
- If the user's calendar detail is needed for a candidate slot, also call `list_events` for that day so you can show meeting titles.
