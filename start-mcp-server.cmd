@echo off
echo Starting Outlook Calendar MCP Server...
cd /d "%~dp0mcp-server"
node src\index.js
