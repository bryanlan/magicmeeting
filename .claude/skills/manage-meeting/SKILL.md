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
- Read `~/.claude/config.md` for user/organizer email

## Find meeting first
Use `list_events` with filters: `subjectContains`, `attendeeEmail`, `locationContains`. Confirm with user before modifying.

## Recurring meetings (critical)
Always ask: **"Just this instance or the entire series?"**

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
1. Find meeting
2. Check availability with `get_free_busy` (all attendees)
3. Present options with user's real meeting titles
4. Confirm new time
5. `mcp__outlook__update_event(eventId, startDate, startTime, endDate, endTime, [originalStart|updateSeries], sendUpdate: true)`

### Add/remove attendee
```
mcp__outlook__add_attendee(eventId, attendee, type: "required"|"optional"|"resource", sendUpdate: true)
mcp__outlook__remove_attendee(eventId, attendee, sendUpdate: true)
```
Add `updateSeries: true` or `originalStart` for recurring. Resolve names via `mcp__outlook__resolve_recipient` or `~/.claude/outlook-contacts.json`.

Never forward .vcs files - use `add_attendee` instead.

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
