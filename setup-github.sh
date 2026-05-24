#!/usr/bin/env bash
# Run this from inside the outlook-mcp-server folder to create the GitHub repo and push.
# Requires: git, gh (GitHub CLI) — install gh from https://cli.github.com

set -e

REPO_NAME="outlook-mcp-server"
DESCRIPTION="MCP server for Microsoft Outlook — intelligent email triage (Action Required / Waiting On / Reference/Read / Archive-Delete) via Microsoft Graph API"

echo "→ Authenticating GitHub CLI (if not already)"
gh auth status 2>/dev/null || gh auth login

echo "→ Initialising git repo"
git init
git add .
git commit -m "feat: initial Outlook MCP server with four-category triage

- outlook_triage_unread_emails: sorts unread mail into Action Required,
  Waiting On, Reference/Read, Archive/Delete with per-email rationale
- 11 tools total: triage, list, get, reply, forward, mark-read,
  flag, move, delete, search, list-folders
- MSAL device-code auth with file-backed token cache
- Microsoft Graph API client with typed responses
- Heuristic triage engine (spam keywords, expired notifications,
  automated senders, action keywords, CC vs To detection)"

echo "→ Creating GitHub repo: $REPO_NAME"
gh repo create "$REPO_NAME" \
  --description "$DESCRIPTION" \
  --public \
  --source=. \
  --remote=origin \
  --push

echo ""
echo "✅ Done! Your repo is live at:"
gh repo view --json url -q .url
