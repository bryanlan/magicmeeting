---
name: lookup-contact
description: Resolves a person's email address via GAL or email history, and can add them to the local contacts file. Use for "look up", "find email for", or "add [name] to contacts".
allowed-tools: Read, mcp__outlook__resolve_recipient, mcp__outlook__search_inbox
argument-hint: "[person-name]"
---

# Lookup Contact

## ERROR? STOP.
**Tool error? STOP. DO NOT WORKAROUND. DO NOT USE DIFFERENT TOOLS.** Tell user what failed. Wait.

Find someone's email address by name, then optionally save it into `~/.claude/outlook-contacts.json`.

## Workflow

### 1) Try GAL first
Call:
- `mcp__outlook__resolve_recipient(query: "firstname lastname")`

If it returns a single strong match, present it and confirm with the user.

### 2) Search email history if GAL fails (or user requests)
**Emails FROM this person:**
- `mcp__outlook__search_inbox(fromAddresses: "firstname", limit: 20)`

**Emails TO this person (sent folder):**
- `mcp__outlook__search_inbox(folder: "sent", toAddresses: "firstname", limit: 20)`
  - If your MCP tool doesn't support `toAddresses`, fall back to searching the sent folder via subject/body keywords.

### 3) Present options (dedupe by email)
Show unique contacts found:

| # | Name | Email | Source |
|---|------|-------|--------|
| 1 | John Smith | jsmith@example.com | GAL |
| 2 | John Doe | johnd@other.com | Inbox |

### 4) Add to address book (only if user confirms)
Update `~/.claude/outlook-contacts.json`:

- Normalize email to lowercase.
- Add reasonable aliases (first name, last name, common shorthand).
- Preserve existing JSON structure and formatting.

Example contact entry:
```json
{
  "name": "John Smith",
  "email": "jsmith@example.com",
  "aliases": ["john", "jsmith"]
}
```

## Tips

* Start broad (first name), then narrow (add last name) if too many matches.
* Search both inbox and sent for coverage.
* Never invent an email; if unresolved, ask the user for the address.
