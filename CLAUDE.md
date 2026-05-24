# Outlook MCP Server — Claude Code Context

## What this project is
A TypeScript MCP (Model Context Protocol) server that connects to Microsoft Outlook via the Microsoft Graph API. It acts as an intelligent email assistant, triaging unread emails into four categories and providing full email management tools.

## Architecture

```
src/
├── index.ts          # Entry point — validates env, starts stdio MCP server
├── auth.ts           # MSAL device-code OAuth flow + file-backed token cache
├── graph-client.ts   # Typed Microsoft Graph API client (axios-based)
├── triage.ts         # Heuristic email categorisation engine
├── types.ts          # Shared TypeScript interfaces
├── constants.ts      # API URLs, category labels, heuristic regex patterns
└── tools/
    └── email-tools.ts  # All 11 MCP tool registrations
```

## Four triage categories (priority order)
1. **Archive/Delete** — spam, expired notifications, CI alerts, digests (evaluated FIRST)
2. **Action Required** — needs human response/decision
3. **Waiting On** — replies expected from others
4. **Reference/Read** — newsletters, receipts, FYI

## Key commands
```bash
npm install       # install dependencies (requires Node 18+)
npm run build     # compile TypeScript → dist/
npm start         # run the MCP server (stdio transport)
npm run dev       # run with tsx watch (auto-reload)
```

## Required environment variables
| Variable | Required | Notes |
|---|---|---|
| `OUTLOOK_CLIENT_ID` | ✅ | Azure app registration client ID |
| `OUTLOOK_TENANT_ID` | No | Defaults to "common" |
| `OUTLOOK_USER_EMAIL` | No | Improves triage accuracy (direct vs CC detection) |

## Adding a new tool
1. Add the registration function to `src/tools/email-tools.ts`
2. Call it inside `registerAllEmailTools()` at the bottom of that file
3. If it needs a new Graph API endpoint, add the call to `src/graph-client.ts`
4. If it needs new heuristic patterns, add them to `src/constants.ts`

## Adding a new triage category
1. Add the label to `CATEGORY` in `src/constants.ts`
2. Add detection patterns/signals in `src/constants.ts`
3. Import the patterns in `src/triage.ts` and add signal detection
4. Add the category to the decision tree in `triageMessage()` — respect priority order
5. Add the bucket to `TriageReport.categories` in `src/types.ts`
6. Add the filter and markdown section in `registerTriageUnreadEmails()` in email-tools.ts
7. Update README.md

## Graph API scopes used
`Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `User.Read`

## Token cache
Stored at `~/.outlook-mcp-token-cache.json` with `0600` permissions.
Delete it to force re-authentication.
