# Outlook Assistant for Claude Code

Claude Code skills and MCP server for Outlook calendar and email management on Windows.

## What It Can Do

**Calendar**
- View your calendar for any date range
- Check free/busy status for yourself and others
- Find common available times across multiple attendees
- Book meetings with attendees, conference rooms, and Teams links
- Reschedule single instances or entire recurring series
- Change rooms for recurring meetings (handles exceptions)
- Add/remove attendees from existing meetings

**Email**
- Search emails by sender, subject, body, date range, or folder
- Read full email content
- Send emails with HTML formatting (requires confirmation)
- Resolve names to email addresses via Global Address List

**Smart Behaviors**
- Treats "Tentative" as available when checking free/busy
- Shows actual meeting titles (not just "busy blocks") so you can decide what to move
- Books meetings at :05 or :35 by default (respects meeting buffer time)
- Maintains a personal contacts/aliases file for quick name resolution
- Requires explicit confirmation before creating/modifying calendar events or sending email

## Issues & Feedback

Something not working? [Open an issue](https://github.com/bryanlan/magicmeeting/issues)

## Requirements

- Windows with Outlook desktop app installed and signed in
- [Claude Code CLI](https://claude.ai/claude-code)

## Setup

### 1. Open Claude Code

### 2. Copy and paste the text below into Claude Code

```
Set up the Outlook Assistant from https://github.com/bryanlan/magicmeeting

Do the following steps in order:

1. CHECK PREREQUISITES:
   - Verify Node.js is installed (run: node --version). If not installed, tell user to download from https://nodejs.org/ and restart Claude Code after installing.
   - Verify Git is installed (run: git --version). If not installed, tell user to download from https://git-scm.com/ and restart Claude Code after installing.
   - If either is missing, STOP and tell the user what to install.

2. CLONE THE REPO:
   - Clone https://github.com/bryanlan/magicmeeting.git to a local folder (suggest ~/projects/magicmeeting or similar)
   - cd into the cloned folder

3. INSTALL MCP SERVER DEPENDENCIES:
   - cd mcp-server
   - npm install
   - cd ..

4. COPY SKILLS TO USER CONFIG:
   - Copy the contents of .claude/skills/ to ~/.claude/skills/ (create if needed)
   - Do NOT overwrite existing skills the user may have

5. CREATE CONFIG FILE:
   - Ask the user for their full name and email address
   - Create ~/.claude/config.md with this content:
     ```
     # User Configuration

     ## Identity
     - **User Email**: {their_email}
     - **User Name**: {their_name}

     ## Notes
     Edit this file to configure skills for your own use.
     ```

6. CREATE CONTACTS FILE (if it doesn't exist):
   - Create ~/.claude/outlook-contacts.json with:
     ```json
     {
       "version": "1.2",
       "contacts": [],
       "groups": []
     }
     ```

7. CONFIGURE MCP SERVER:
   - Get the absolute path to mcp-server/src/index.js
   - Tell the user to add this to their Claude Code MCP settings (via /mcp or ~/.claude/settings.json):
     ```json
     {
       "mcpServers": {
         "outlook": {
           "command": "node",
           "args": ["ABSOLUTE_PATH_TO/mcp-server/src/index.js"]
         }
       }
     }
     ```
   - Replace ABSOLUTE_PATH_TO with the actual path, using forward slashes

8. FINAL INSTRUCTIONS:
   - Tell the user: "Setup complete! Please restart Claude Code for changes to take effect."
   - After restart, tell them to run /mcp to verify the outlook server is connected
   - Suggest they try: /outlook-calendar today
```

## Skills

| Skill | Description |
|-------|-------------|
| `/outlook-calendar` | View calendar, check availability |
| `/book-meeting` | Schedule meetings with attendees, rooms, Teams |
| `/search-email` | Search mailbox |
| `/send-email` | Compose and send email (with confirmation) |
| `/lookup-contact` | Resolve names to email addresses |

## License

MIT
