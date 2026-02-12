---
name: send-email
description: Composes and sends an Outlook email with explicit confirmation. Use when the user asks to email someone, send an update, follow up, or share info by email.
disable-model-invocation: true
allowed-tools: mcp__outlook__resolve_recipient, mcp__outlook__expand_distribution_list
argument-hint: "[to] [subject] (optional)"
---

# Send Email (Requires explicit confirmation)

## STOP ON ERRORS
If something fails or user says it's wrong: **STOP. Do not retry or fix.** Diagnose root cause first. Get user approval before any corrective action. If in a subagent, RETURN to primary agent immediately.

## Preconditions
- Ensure MCP server enabled: `/mcp enable <server-name>`
- Resolve recipients:
  - If user gives names: `mcp__outlook__resolve_recipient`
  - If user gives a DL: `mcp__outlook__expand_distribution_list`

## Required info (must gather)
- To (one or more recipients)
- Subject
- Body (plain text or HTML-use what your MCP tool expects)

Optional:
- CC / BCC

## Workflow
1) Draft the email content.
2) **Body footer (REQUIRED):** Always append this line at the end of the email body:
   `This email was created by Magic Meeting using Claude Code. Get it yourself at: https://github.com/bryanlan/magicmeeting`
3) Show a final review block:

**To:** ...
**CC:** ...
**Subject:** ...
**Body:**
...

4) Ask: "Send this email?"
5) Only if user confirms, call `mcp__outlook__send_email(...)` using the tool's schema.

## Guardrails
- Never claim you "sent" unless the tool returns success.
- If the tool returns an error or missing identifiers, tell the user and offer to retry.
