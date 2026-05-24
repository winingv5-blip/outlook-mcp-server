# Outlook MCP Server

A Model Context Protocol server for Microsoft Outlook that acts as a highly efficient digital assistant. It fetches your unread emails and sorts them into three actionable categories:

| Category | Description |
|---|---|
| ⚡ **Action Required** | Needs your decision, response, or action |
| ⏳ **Waiting On** | Replies you are expecting from others |
| 📚 **Reference/Read** | Newsletters, receipts, confirmations, FYI mail |
| 🗑️ **Archive/Delete** | Obvious spam, expired notifications, CI alerts, daily digests |

---

## Prerequisites

- Node.js 18+
- A Microsoft account (personal Outlook.com or work/school account)
- An Azure App Registration (free, takes ~5 minutes)

---

## Azure App Registration (one-time setup)

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Name it anything (e.g. "Outlook MCP")
3. Supported account types: **"Accounts in any organizational directory and personal Microsoft accounts"** (or restrict to your tenant)
4. Redirect URI: **Mobile and desktop applications** → `http://localhost`
5. Click **Register**
6. Copy the **Application (client) ID** and **Directory (tenant) ID**
7. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**, add:
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `User.Read`
8. Click **Grant admin consent** (if you have admin rights), or just proceed — you'll be prompted during sign-in

---

## Installation

```bash
git clone <your-repo>   # or unzip the downloaded folder
cd outlook-mcp-server
npm install
npm run build
```

---

## Configuration

Set these environment variables before running:

| Variable | Required | Description |
|---|---|---|
| `OUTLOOK_CLIENT_ID` | ✅ Yes | Azure app client ID from step 6 above |
| `OUTLOOK_TENANT_ID` | No | Tenant ID, or `common` (default) for personal accounts |
| `OUTLOOK_USER_EMAIL` | No | Your email address — improves triage accuracy |

```bash
export OUTLOOK_CLIENT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
export OUTLOOK_TENANT_ID="common"          # or your tenant ID
export OUTLOOK_USER_EMAIL="you@outlook.com"
```

---

## First Run (Device Code Auth)

On first run the server prints a short URL and code to your terminal. Open the URL in any browser, type the code, and sign in to your Microsoft account. The token is cached in `~/.outlook-mcp-token-cache.json` and silently refreshed on subsequent runs — you won't be prompted again unless the token expires.

---

## Claude Desktop / Claude Code Setup

Add to your MCP config (`claude_desktop_config.json` or `.mcp.json`):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/absolute/path/to/outlook-mcp-server/dist/index.js"],
      "env": {
        "OUTLOOK_CLIENT_ID": "your-client-id-here",
        "OUTLOOK_TENANT_ID": "common",
        "OUTLOOK_USER_EMAIL": "you@outlook.com"
      }
    }
  }
}
```

Restart Claude Desktop. The Outlook tools will appear in the tool list.

---

## Available Tools

### Email Triage

| Tool | Description |
|---|---|
| `outlook_triage_unread_emails` | **The main tool.** Fetches all unread emails and sorts them into Action Required / Waiting On / Reference/Read with a rationale for each |
| `outlook_list_unread_emails` | Paginated list of unread emails (raw, no categorisation) |

### Email Reading

| Tool | Description |
|---|---|
| `outlook_get_email` | Full email body + metadata by ID |
| `outlook_search_emails` | Full-text search across your mailbox |
| `outlook_list_folders` | List all folders with unread counts |

### Email Actions

| Tool | Description |
|---|---|
| `outlook_reply_email` | Reply (or reply-all) to an email |
| `outlook_forward_email` | Forward to one or more recipients |
| `outlook_mark_as_read` | Mark one or more emails read/unread |
| `outlook_flag_email` | Flag / unflag for follow-up |
| `outlook_move_email` | Move to a folder (by name or ID) |
| `outlook_delete_email` | Delete (moves to Deleted Items) |

---

## Example Prompts

```
Triage my unread emails and tell me what needs action today.

Search my email for anything from alice@company.com this month.

Mark all the newsletters in my triage report as read and move them to the "Newsletters" folder.

Reply to the email from Bob about the Q3 report — tell him I'll have it by Friday.
```

---

## Triage Heuristics

The server uses these signals to categorise emails (the LLM can override). **Archive/Delete is evaluated first** and wins over all other categories.

**🗑️ Archive/Delete** — spam keywords ("you have won", "claim your prize", "verify your account now"), mailer-daemon / postmaster / bounce senders, expired notifications (password reset links, CI build results, daily/weekly digests), system report senders

**⚡ Action Required** — direct-to-To recipient, action keywords ("please review", "can you", "deadline", "ASAP"), high importance, flagged

**⏳ Waiting On** — "Re:" subject where you are not the sender (suggests a reply you were waiting for arrived)

**📚 Reference/Read** — automated senders (noreply@, newsletter@), receipt/confirmation keywords, CC-only addressing

---

## Security Notes

- The token cache is stored at `~/.outlook-mcp-token-cache.json` with `0600` permissions (readable only by you)
- No credentials are stored in the server code
- To sign out: delete `~/.outlook-mcp-token-cache.json`
- The app uses **delegated permissions** only — it acts as you, not as an admin

---

## Troubleshooting

**"OUTLOOK_CLIENT_ID environment variable is required"** → Set the env var and restart.

**403 Permission denied** → Go back to your Azure app registration and verify the Mail permissions are added and (if required by your tenant) admin-consented.

**Token expired** → Delete `~/.outlook-mcp-token-cache.json` and restart — you'll be prompted to sign in again.

**Rate limit errors** → Microsoft Graph throttles at ~10,000 requests/10 minutes for personal accounts. Wait a moment and retry.
