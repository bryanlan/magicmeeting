# Outlook Assistant for Claude Code

Claude Code skills and MCP server for Outlook calendar and email management on Windows.

## Requirements

- Windows with Outlook desktop app (uses COM automation)
- Node.js 18+
- Python 3.10+ (for calendar parsing scripts)
- [Claude Code CLI](https://claude.ai/claude-code)

## Setup

### 1. Install MCP Server

```bash
cd mcp-server
npm install
```

### 2. Configure Claude Code

Add to your Claude Code MCP settings (`~/.claude/settings.json` or via `/mcp`):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["C:/path/to/scheduler/mcp-server/src/index.js"]
    }
  }
}
```

### 3. Copy Skills

Copy `.claude/skills/` contents to your `~/.claude/skills/` directory.

### 4. Create Config

Create `~/.claude/config.md` with your identity:

```markdown
# User Config
- Name: Your Name
- Email: you@example.com
```

### 5. Create Contacts File

Create `~/.claude/outlook-contacts.json`:

```json
{
  "version": "1.2",
  "contacts": [],
  "groups": []
}
```

## Skills

| Skill | Description |
|-------|-------------|
| `/outlook-calendar` | View calendar, check availability |
| `/book-meeting` | Schedule meetings with attendees, rooms, Teams |
| `/search-email` | Search mailbox |
| `/send-email` | Compose and send email (with confirmation) |
| `/lookup-contact` | Resolve names to email addresses |

## Usage

```
> /outlook-calendar today
> /book-meeting with Alice and Bob next week
> /search-email from:manager subject:review
> /lookup-contact John Smith
```

## Key Behaviors

- **Tentative = Free**: When checking availability, tentative slots are treated as available
- **Confirmation required**: Calendar changes and emails require explicit user approval
- **Real meeting names**: Shows actual meeting titles, not just "busy" blocks

## License

MIT
