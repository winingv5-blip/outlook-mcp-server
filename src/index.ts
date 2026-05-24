#!/usr/bin/env node
/**
 * Outlook MCP Server
 *
 * A highly efficient digital assistant for Microsoft Outlook that analyses
 * unread emails and sorts them into:
 *   • Action Required  — needs your decision or response
 *   • Waiting On       — replies you're expecting from others
 *   • Reference/Read   — newsletters, receipts, FYI mail
 *
 * Transport: stdio (for use with Claude Desktop / Claude Code)
 *
 * Required environment variables:
 *   OUTLOOK_CLIENT_ID   – Azure app registration client ID
 *   OUTLOOK_TENANT_ID   – Tenant ID (or "common" / "consumers")
 *   OUTLOOK_USER_EMAIL  – (optional) your email address, improves triage accuracy
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerAllEmailTools } from "./tools/email-tools.js";

// ─── Validate required environment ───────────────────────────────────────────

const REQUIRED_ENV = ["OUTLOOK_CLIENT_ID"] as const;

function validateEnv(): void {
  const missing = REQUIRED_ENV.filter((key) => !process.env[key]);
  if (missing.length > 0) {
    console.error("════════════════════════════════════════════════════════");
    console.error("  Outlook MCP Server — Configuration Required");
    console.error("════════════════════════════════════════════════════════");
    console.error("");
    console.error("  Missing required environment variables:");
    missing.forEach((key) => console.error(`    • ${key}`));
    console.error("");
    console.error("  Setup steps:");
    console.error("  1. Go to https://portal.azure.com → App registrations");
    console.error("  2. Create a new registration (public client / native)");
    console.error('  3. Add redirect URI: http://localhost (type: "Mobile and desktop")');
    console.error("  4. Under API Permissions, add:");
    console.error("       Microsoft Graph → Delegated:");
    console.error("         Mail.Read, Mail.ReadWrite, Mail.Send, User.Read");
    console.error("  5. Set environment variables and restart.");
    console.error("");
    console.error("  Example (bash):");
    console.error("    export OUTLOOK_CLIENT_ID=<your-app-client-id>");
    console.error("    export OUTLOOK_TENANT_ID=common");
    console.error("    export OUTLOOK_USER_EMAIL=you@example.com");
    console.error("════════════════════════════════════════════════════════");
    process.exit(1);
  }
}

// ─── Server setup ─────────────────────────────────────────────────────────────

const server = new McpServer({
  name: "outlook-mcp-server",
  version: "1.0.0",
});

registerAllEmailTools(server);

// ─── Start ────────────────────────────────────────────────────────────────────

async function main(): Promise<void> {
  validateEnv();

  const transport = new StdioServerTransport();
  await server.connect(transport);

  console.error("[outlook-mcp] Server running via stdio. Awaiting tool calls.");
}

main().catch((err) => {
  console.error("[outlook-mcp] Fatal error:", err);
  process.exit(1);
});
