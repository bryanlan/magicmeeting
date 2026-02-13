---
name: book-meeting
description: Creates NEW Outlook meetings with attendees, rooms, and Teams links. Use when user asks to book/schedule a NEW meeting, find a time to meet, or set up a call. For editing existing meetings (reschedule, cancel, change rooms, add attendees), use /manage-meeting instead.
allowed-tools: Read, mcp__outlook__get_free_busy, mcp__outlook__find_free_slots, mcp__outlook__list_events, mcp__outlook__find_available_rooms, mcp__outlook__resolve_recipient, mcp__outlook__expand_distribution_list, mcp__outlook__search_inbox, mcp__outlook__get_email_content, mcp__outlook__create_event
argument-hint: "[who] [when] [duration] (optional)"
---

# Book Meeting Skill

## ERROR? STOP.
**Tool error? STOP. DO NOT WORKAROUND. DO NOT USE DIFFERENT TOOLS.** Tell user what failed. Wait.

## Preconditions
- Ensure Outlook MCP server is enabled:
  - `/mcp` then `/mcp enable <server-name>` (commonly `outlook`)
- Read `~/.claude/config.md` to get the user/organizer email.
- Tool naming: this skill assumes server name `outlook` (tools look like `mcp__outlook__...`). If your server name differs, replace accordingly.

---

## Meeting time preferences (default)
Default behavior:
- Start **5 minutes after** the requested boundary time.
- Keep the **end time** at the originally requested end boundary (net: meeting is 5 minutes shorter).
Examples:
- "2pm for 30 min" -> 2:05-2:30 (25 min)
- "2pm for 1 hour" -> 2:05-3:00 (55 min)

Exception:
- If user says "exactly", "sharp", or gives a precise minute that implies strictness, honor exact times.

---

## Required info (gather before booking)
- **Attendees**: emails (semicolon-separated). If user provides names, resolve them first.
- **Subject**: meeting title (REQUIRED)
- **Body**: agenda/context (REQUIRED - always ask)
- **Duration**: minutes (REQUIRED - no default)
- **Date**: resolve relative dates ("next Tuesday") into an absolute date before proposing times.

---

## Availability check philosophy (critical)
When checking availability:
- Show other attendees' Free/Tentative/Busy status.
- Show ALL of the user's meetings in the candidate window with **actual meeting names**.
- Do not "optimize away" conflicts: the user can move their own meetings, so they need the real details.

---

## Core workflow

### 1) Resolve attendees
- If user gave names: `mcp__outlook__resolve_recipient("Name")`
- If user gave a DL: `mcp__outlook__expand_distribution_list(...)`
- Build attendee string as `user_email;attendee1;attendee2`

### 2) Check availability (always include organizer first)
**ALWAYS include the user as the first attendee in `get_free_busy`.**

Call:
`mcp__outlook__get_free_busy(attendees: "{user_email};attendee1;attendee2", ...)`

### 3) CRITICAL: Verify the user's calendar for each candidate day
Before recommending any time slot:
- Fetch the user's calendar for that specific day using `mcp__outlook__list_events(startDate, endDate)`.
- If the tool indicates truncation (ex: `hasMore: true`), you do not have complete data - do not claim "free".

### 4) Present options with real detail
For each viable slot, show:
1) Other attendees' status (Free/Tentative/Busy)
2) The user's meetings that overlap, with actual titles

**WRONG (lazy):**
Tuesday 2/17, 11:00 AM - Tucker: Free, Hamza: Free
Your schedule: busy 9-5

**RIGHT (use real titles):**
Tuesday 2/17, 11:00 AM - Tucker: Free, Hamza: Free
Your meetings 10:30 AM-12:00 PM:
- 10:35-11:00: SiGMA LT Weekly
- 11:05-11:30: Ana/Bryan sync
- 11:35-12:00: Cassie | Bryan 1:1

User decides what to move.

### 5) Find a room (if requested)
`mcp__outlook__find_available_rooms(building, floor, startDate, startTime, endTime, capacity)`

### 6) Create meeting (explicit confirmation required)
Before creating:
- Display final summary (attendees, subject, agenda/body, date/time, location/room, Teams yes/no)
- Ask user: "Create it?"

**Body footer (REQUIRED):** Always append this line at the end of the meeting body:
`This meeting was created by Magic Meeting using Claude Code. Get it yourself at: https://github.com/bryanlan/magicmeeting`

Then call:
`mcp__outlook__create_event(subject, body, startDate, startTime, endTime, isMeeting: true, attendees, room, teamsMeeting: true)`

If event ID is missing/empty in the response, treat it as failure and tell the user.

---

## Recurring meetings

To create a recurring meeting, add these parameters to `create_event`:

| Parameter | Description |
|-----------|-------------|
| `recurrenceType` | `"daily"`, `"weekly"`, `"monthly"`, or `"yearly"` |
| `recurrenceInterval` | Every N periods (default: 1) |
| `recurrenceDays` | For weekly: comma-separated days (e.g., `"monday,wednesday,friday"`) |
| `recurrenceEndDate` | End date `MM/DD/YYYY` (optional) |
| `recurrenceOccurrences` | Number of occurrences (alternative to endDate) |

If neither `recurrenceEndDate` nor `recurrenceOccurrences` is set, the series has no end date.

**Example - weekly Friday meeting:**
```
mcp__outlook__create_event(
  subject: "Team Sync",
  startDate: "02/14/2026",
  startTime: "10:00 AM",
  endTime: "10:30 AM",
  isMeeting: true,
  attendees: "someone@example.com",
  teamsMeeting: true,
  recurrenceType: "weekly",
  recurrenceDays: "friday"
)
```

---

## Address book (`~/.claude/outlook-contacts.json`)
Structure:
- `contacts[]`: name, email, aliases (array)
- `groups[]`: name, email (DL address), aliases (array), members

Lookup flow (search BOTH contacts AND groups):
1) Search contacts by name or any alias (case-insensitive)
2) Search groups by name or any alias (case-insensitive)
3) Single match → use email silently
4) Multiple matches → ask user to clarify
5) No match → use `resolve_recipient` or ask for email

**Example:** User says "invite sigma team" → find group with alias "sigma team" → use `coreospmsigma@microsoft.com`

Expanding DLs:
- Use `mcp__outlook__expand_distribution_list`
- Add missing members to contacts, update group.members

---

## Availability polling (optional)
Use when `get_free_busy` shows no workable common times.

Minimal practical flow:
1) Generate poll ID (`poll-XXXXXX`)
2) Email options (4-8 slots) via `/send-email` skill (explicit confirm)
3) Save state to `~/.claude/outlook-polls.json`
4) When user asks "check the poll", read `outlook-polls.json` and search inbox for replies, then present a matrix and propose the best slot.
5) When user confirms a slot, book via `create_event`.
