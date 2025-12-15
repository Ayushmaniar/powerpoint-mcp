# PowerPoint MCP Server

Model Context Protocol server for PowerPoint automation on Windows.

## Launch Video

[![Watch the PowerPoint MCP Demo](video_thumbnail.jpg)](https://www.youtube.com/watch?v=5p24Vr36py8)

---

## Windows Only

**This MCP server works exclusively on Windows** because it uses `pywin32` to automate Microsoft PowerPoint through COM automation.

---

# Installation

## Prerequisites
- Windows 10/11
- Microsoft PowerPoint installed
- Python 3.10+

## Install with uvx (One step/One line installation across most MCP Clients)

### Claude Code
```bash
claude mcp add powerpoint -- uvx powerpoint-mcp

## Making it avaiable for across all project (Install once, use everywhere)

claude mcp add powerpoint --scope user -- uvx powerpoint-mcp
```

### Cursor
1) Click on Settings -> Tools & MCP -> New MCP Server
2) `~/.cursor/mcp.json` would open up
3) Copy-paste this json
```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uvx",
      "args": ["powerpoint-mcp"]
    }
  }
}
```
4) *Restart your IDE after configuration.*


### VS Code (GitHub Copilot)
1) Open `C:\Users\Your_User_Name\AppData\Roaming\Code\User\mcp.json`
2) Copy-paste this json
```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uvx",
      "args": ["powerpoint-mcp"]
    }
  }
}
```
3) *Restart your IDE after configuration.*
