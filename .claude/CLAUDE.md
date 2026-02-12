# Outlook Assistant (Claude Code)

This workspace configures Claude Code as an Outlook-centric assistant (calendar + email) via an MCP server.

## Non-negotiables

1) **ALL Outlook actions MUST follow skill playbooks.** Every tool call and workflow step should be grounded in a skill file in `.claude/skills/`. No freestyling.
2) **If no skill covers the action, STOP.** Do not improvise. Instead:
   - Pause and explain what action is needed but not covered
   - Ask the user how to proceed
   - Update or create the relevant skill file BEFORE executing the action
3) Never create/update/delete calendar events or send emails without an explicit user confirmation step.
4) Always read `~/.claude/config.md` early to confirm the organizer identity (name + email).
5) When presenting calendar availability, show the user's real meeting titles (not "busy blocks") so the user can decide what to move.

## MCP server lifecycle (START / STOP / RESTART)

Claude Code can toggle MCP servers from inside the session:

- **Status / manage:** `/mcp`
- **Start (enable):** `/mcp enable <server-name>`
- **Stop (disable):** `/mcp disable <server-name>`
- **Restart:** `/mcp disable <server-name>` then `/mcp enable <server-name>`

Use the exact `<server-name>` shown in `/mcp` (example: `outlook`).

If OAuth/auth is required for the server, `/mcp` is also where authentication is initiated.

### What NOT to do
- Don't use OS-level "kill node.exe" as a normal restart mechanism.
- Don't manually run the server process in a separate terminal unless the MCP configuration explicitly requires it.

## Where important state lives

- `~/.claude/config.md`
  - Source of truth for user identity and organizer email.

- `~/.claude/outlook-contacts.json`
  - Address book (contacts + groups).

- `~/.claude/outlook-polls.json`
  - Availability poll tracking (if you use polls).

## Editing skills and config

**CRITICAL: Always edit skills in THIS PROJECT's `.claude/skills/` folder, never `~/.claude/skills/`.** The project folder is the source of truth for version-controlled skill definitions.

Keep edits to SKILL.md and CLAUDE.md **minimalist and crisp**. No verbose explanations or lengthy examples. One-liners preferred.

## Mistake handling (systemic fix first)

**DO NOT JUMP TO FIXING THE IMMEDIATE PROBLEM.** When a mistake happens:

1) **STOP** - Do not attempt to fix the specific instance
2) **Identify root cause** - What assumption, missing info, or incorrect tool usage caused this?
   - **Do NOT guess.** If unsure, provide a research prompt for the user to investigate.
3) **Update skill instructions** - Edit the relevant SKILL.md to prevent recurrence
4) **WAIT** - Ask user for confirmation before proceeding to fix the specific instance

## Available skills (slash commands)

- `/outlook-calendar` - view calendar, free/busy checks, date-range summaries
- `/book-meeting` - scheduling workflow (availability, room, Teams link, recurring rules, poll option)
- `/search-email` - search mailbox and retrieve message content
- `/send-email` - compose and send an email (explicit confirm)
- `/lookup-contact` - resolve an email address + optionally add to contacts file
