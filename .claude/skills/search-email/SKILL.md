---
name: search-email
description: Searches Outlook email by sender, subject, body keywords, folder, and date range; can fetch full message content. Use when the user asks to find an email, thread, attachment reference, or what someone said.
allowed-tools: mcp__outlook__search_inbox, mcp__outlook__get_email_content, mcp__outlook__resolve_recipient
argument-hint: "[keywords] (optional)"
---

# Search Email

## ERROR? STOP.
**Tool error? STOP. DO NOT WORKAROUND. DO NOT USE DIFFERENT TOOLS.** Tell user what failed. Wait.

## Preconditions
- If Outlook MCP tools are missing/unavailable: use `/mcp` and `/mcp enable <server-name>`.
- Use `mcp__outlook__resolve_recipient` when the user provides a person name but no email.

## Core tools
- `mcp__outlook__search_inbox(...)`
- `mcp__outlook__get_email_content(emailId)`
- `mcp__outlook__resolve_recipient(query)`

## Search strategy (fast -> slow)
1) Subject filter first (highest signal, fastest)
2) Sender filter next
3) Date window next
4) Body keyword search last (noisy)

## Typical parameters (adapt to your MCP schema)
- `subjectContains`
- `fromAddresses` (semicolon-separated for OR logic)
- `bodyContains` (space-separated keywords; confirm whether it's AND/OR behavior)
- `folder` ("inbox", "sent", or subfolder name)
- `receivedAfter` / `receivedBefore` (MM/DD/YYYY)
- `limit`
- `includeBody` (preview)

## Output format
Return a short ranked list:
- #, date, from, subject, and a 1-2 line snippet.
If the user asks for details, call `get_email_content(emailId)` and show the relevant excerpt.

## Guardrails
- This skill is read-only. Do not send replies or new emails here-use `/send-email`.
