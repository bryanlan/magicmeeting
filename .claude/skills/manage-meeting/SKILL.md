---
name: manage-meeting
description: Modify existing meetings - cancel, reschedule, add/remove attendees, change rooms, add Teams links, update details. Use when user wants to edit, cancel, or change an existing meeting.
allowed-tools: Read, mcp__outlook__list_events, mcp__outlook__update_event, mcp__outlook__delete_event, mcp__outlook__cancel_event, mcp__outlook__add_attendee, mcp__outlook__remove_attendee, mcp__outlook__get_attendee_status, mcp__outlook__find_available_rooms, mcp__outlook__resolve_recipient, mcp__outlook__expand_distribution_list, mcp__outlook__get_free_busy
argument-hint: "[action] [meeting name/details]"
---

## ERROR? STOP.
**Tool error? STOP. DO NOT WORKAROUND. DO NOT USE DIFFERENT TOOLS.** Tell user what failed. Wait.

## Preconditions
- Outlook MCP server enabled
- Read `.claude/config.md` for user/organizer email
- Read `.claude/outlook-contacts.json` for address book

## Attendee resolution (MANDATORY)
**ALWAYS check address book BEFORE adding any attendee.** NEVER call `add_attendee` without first:
1. Read `.claude/outlook-contacts.json`
2. Search for the name in contacts and groups
3. If single match → use that email
4. If multiple matches → ask user to clarify
5. If no match → use `resolve_recipient` or ask user for email

**DO NOT guess emails. DO NOT assume email formats.**

## Meeting time preferences
When rescheduling, start **5 minutes after** the requested boundary time (e.g., "1:30" → 1:35).
Exception: honor exact times if user says "exactly", "sharp", or gives a precise non-boundary minute.

## Find meeting first
Use `list_events` with filters: `subjectContains`, `attendeeEmail`, `locationContains`. Confirm with user before modifying.

## Recurring meetings (critical)
Always ask: **"Just this instance or the entire series?"**

### How recurring item IDs work
`list_events` returns these fields for recurring items:
- `id` = **Master's EntryID** (always the series master, even for exceptions)
- `originalStart` = **Original scheduled time** (the time to pass to GetOccurrence)
- `start` = Current displayed time (may differ from originalStart for exceptions)
- `recurrenceState` = `occurrence` | `exception` | `master` | `notRecurring`

### Updating a single occurrence or exception
**ALWAYS** use `id` + `originalStart` from list_events:
```
mcp__outlook__update_event(eventId: <id>, originalStart: <originalStart>, ...)
```
The `originalStart` field is pre-computed to work with Outlook's GetOccurrence API.

### Updating the entire series
```
mcp__outlook__update_event(eventId: <id>, updateSeries: true, ...)
```

**NEVER** try to use the occurrence's EntryID directly - Outlook doesn't work that way.

## Cancel meeting
Use `cancel_event` to cancel with custom message (notifies attendees); use `delete_event` for silent deletion.

**Single occurrence of recurring meeting:**
```
mcp__outlook__cancel_event(eventId, occurrenceStart: "2/13/2026 10:35 AM", comment: "Your message")
```
Use the `start` value from `list_events` for that occurrence.

**Entire recurring series (with notifications):**
```
mcp__outlook__cancel_event(eventId, cancelSeries: true, comment: "Your message")
```

**Non-recurring meeting:**
```
mcp__outlook__cancel_event(eventId, comment: "Your message")
```

**Silent deletion (no notification to attendees):**
```
mcp__outlook__delete_event(eventId)
```

Only organizer can cancel. Confirmation required.

## Reschedule / Update / Add-Remove Attendees

For these operations on recurring meetings:

| Scope | Parameter |
|-------|-----------|
| Single instance | `originalStart: "MM/DD/YYYY HH:MM AM/PM"` |
| Entire series | `updateSeries: true` |

### Reschedule
**AUTOMATICALLY check all attendee availability - NEVER ask user if they want this.**

1. Find meeting
2. **IMMEDIATELY** call `get_free_busy` for ALL attendees (extract emails from meeting)
3. Call `list_events` for user's calendar during times other attendees are free
4. Present options showing:
   - Times when ALL attendees are free
   - User's conflicting meetings by **ACTUAL NAME** (never "busy")
5. Confirm new time
6. **PRE-FLIGHT CHECKLIST** before calling update_event:
   - [ ] Start time ends in :05 or :35? (e.g., 2pm → 2:05 PM)
   - [ ] Duration preserved from original meeting?
   - [ ] Using correct originalStart or updateSeries for recurring?
   - [ ] sendUpdate: true?
7. `mcp__outlook__update_event(eventId, startDate, startTime, endDate, endTime, [originalStart|updateSeries], sendUpdate: true)`

### Add/remove attendee (forwarding a meeting)
**ALWAYS ask user which method before adding:**

| Method | Pros | Cons |
|--------|------|------|
| **Add as attendee** | Proper invite (Accept/Decline), gets updates, tracked | May notify ALL existing attendees |
| **Forward only** | Won't spam other attendees | No Accept/Decline, no tracking, no updates |

Present this choice. Then:
- **Add as attendee:** `mcp__outlook__add_attendee(eventId, attendee, sendUpdate: true)`
- **Forward only:** `mcp__outlook__add_attendee(eventId, attendee, sendUpdate: true, forwardAsVcal: true)`

```
mcp__outlook__remove_attendee(eventId, attendee, sendUpdate: true)
```

Add `updateSeries: true` or `originalStart` for recurring.

### Change room
1. `mcp__outlook__find_available_rooms(building, startDate, startTime, endTime, capacity)`
2. `mcp__outlook__remove_attendee(eventId, oldRoomEmail, updateSeries: true)`
3. `mcp__outlook__add_attendee(eventId, newRoomEmail, type: "resource")`
4. `mcp__outlook__update_event(eventId, location: "New Room Name", updateSeries: true)`

**Recurring series:** Check 8-12 future occurrences. Exceptions (modified instances) have different IDs - changes to master don't propagate. Fix each unique ID separately.

### Add Teams link
```
mcp__outlook__update_event(eventId, teamsMeeting: true, sendUpdate: true)
```

### Update subject/body
```
mcp__outlook__update_event(eventId, subject: "New", body: "New", sendUpdate: true)
```

## Confirmation required
Before any modify/delete: show subject, date/time, affected attendees, what changes. Ask "Proceed?"
