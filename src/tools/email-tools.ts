/**
 * MCP tool registrations for Outlook email management.
 */

import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import * as graph from "../graph-client.js";
import { handleGraphError } from "../graph-client.js";
import { toEmailSummary, toTriagedEmail } from "../triage.js";
import {
  CATEGORY,
  CHARACTER_LIMIT,
  DEFAULT_LIMIT,
  MAX_LIMIT,
} from "../constants.js";
import type {
  EmailSummary,
  TriagedEmail,
  TriageReport,
} from "../types.js";

// ─── Shared schema pieces ─────────────────────────────────────────────────────

enum ResponseFormat {
  MARKDOWN = "markdown",
  JSON = "json",
}

const paginationSchema = {
  limit: z
    .number()
    .int()
    .min(1)
    .max(MAX_LIMIT)
    .default(DEFAULT_LIMIT)
    .describe("Maximum number of emails to return (1-100, default 20)"),
  offset: z
    .number()
    .int()
    .min(0)
    .default(0)
    .describe("Number of emails to skip for pagination (default 0)"),
};

const responseFormatSchema = z
  .nativeEnum(ResponseFormat)
  .default(ResponseFormat.MARKDOWN)
  .describe("Output format: 'markdown' for human-readable, 'json' for programmatic use");

// ─── Markdown formatters ──────────────────────────────────────────────────────

function formatEmailSummaryMd(e: EmailSummary): string {
  const lines = [
    `### ${e.subject}`,
    `- **From**: ${e.from}`,
    `- **To**: ${e.to.join(", ") || "—"}`,
    `- **Received**: ${new Date(e.receivedAt).toLocaleString()}`,
    `- **Importance**: ${e.importance}${e.flagged ? " 🚩" : ""}${e.hasAttachments ? " 📎" : ""}`,
    `- **Preview**: ${e.preview}`,
    `- **ID**: \`${e.id}\``,
    `- **Open**: ${e.webLink}`,
  ];
  return lines.join("\n");
}

function formatTriagedEmailMd(e: TriagedEmail): string {
  return [
    formatEmailSummaryMd(e),
    `- **Category**: ${e.suggestedCategory}`,
    `- **Rationale**: ${e.rationale}`,
  ].join("\n");
}

function truncateIfNeeded(text: string): string {
  if (text.length <= CHARACTER_LIMIT) return text;
  return (
    text.slice(0, CHARACTER_LIMIT) +
    "\n\n… [truncated — use pagination or filters to see more]"
  );
}

// ─── Tool: outlook_triage_unread_emails ───────────────────────────────────────

export function registerTriageUnreadEmails(server: McpServer): void {
  server.registerTool(
    "outlook_triage_unread_emails",
    {
      title: "Triage Unread Emails",
      description: `Fetch all unread emails and sort them into four categories with a one-sentence rationale for each:

1. **Action Required** — emails that require a human decision, response, or action (e.g. direct requests, approvals, urgent matters)
2. **Waiting On** — replies you were expecting from others (detected by "Re:" threading heuristics)
3. **Reference/Read** — newsletters, receipts, order confirmations, FYI-only mail, CC-only
4. **Archive/Delete** — obvious spam, expired notifications (password resets, CI build alerts, daily digests), mailer-daemon bounces

Archive/Delete is evaluated first and wins over other categories — if something is clearly disposable, no action keywords can override it.

The tool uses heuristic signals (sender patterns, keywords, importance, CC vs To, reply detection) for initial categorisation. The LLM should review and may reclassify based on full context.

Args:
  - limit: Max emails to analyse (default 50, max 100)
  - folder_id: Analyse a specific folder by ID instead of all mail (optional)
  - response_format: 'markdown' (default) or 'json'

Returns:
  A triage report with emails grouped by category, each including:
  - id, subject, from, to, cc, preview, receivedAt, importance, hasAttachments
  - suggestedCategory, rationale, signals[]`,
      inputSchema: z
        .object({
          limit: z
            .number()
            .int()
            .min(1)
            .max(MAX_LIMIT)
            .default(50)
            .describe("Max unread emails to analyse (default 50)"),
          folder_id: z
            .string()
            .optional()
            .describe("Optional: analyse a specific folder ID instead of all mail"),
          response_format: responseFormatSchema,
        })
        .strict(),
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async ({ limit, folder_id, response_format }) => {
      try {
        const data = await graph.listUnreadMessages({
          limit,
          skip: 0,
          folderId: folder_id,
        });

        const messages = data.value ?? [];
        if (messages.length === 0) {
          return {
            content: [{ type: "text", text: "✅ No unread emails found. Your inbox is clear!" }],
          };
        }

        const triaged = messages.map(toTriagedEmail);

        const report: TriageReport = {
          total: triaged.length,
          analyzedAt: new Date().toISOString(),
          categories: {
            actionRequired: triaged.filter(
              (e) => e.suggestedCategory === CATEGORY.ACTION_REQUIRED
            ),
            waitingOn: triaged.filter(
              (e) => e.suggestedCategory === CATEGORY.WAITING_ON
            ),
            referenceRead: triaged.filter(
              (e) => e.suggestedCategory === CATEGORY.REFERENCE_READ
            ),
            archiveDelete: triaged.filter(
              (e) => e.suggestedCategory === CATEGORY.ARCHIVE_DELETE
            ),
          },
        };

        let text: string;
        if (response_format === ResponseFormat.MARKDOWN) {
          const lines = [
            `# 📬 Email Triage Report`,
            `**Analysed**: ${report.total} unread emails — ${new Date(report.analyzedAt).toLocaleString()}`,
            "",
            `## ⚡ Action Required (${report.categories.actionRequired.length})`,
            ...report.categories.actionRequired.map((e) => formatTriagedEmailMd(e) + "\n---"),
            "",
            `## ⏳ Waiting On (${report.categories.waitingOn.length})`,
            ...report.categories.waitingOn.map((e) => formatTriagedEmailMd(e) + "\n---"),
            "",
            `## 📚 Reference/Read (${report.categories.referenceRead.length})`,
            ...report.categories.referenceRead.map((e) => formatTriagedEmailMd(e) + "\n---"),
            "",
            `## 🗑️ Archive/Delete (${report.categories.archiveDelete.length})`,
            ...report.categories.archiveDelete.map((e) => formatTriagedEmailMd(e) + "\n---"),
          ];
          text = lines.join("\n");
        } else {
          text = JSON.stringify(report, null, 2);
        }

        return {
          content: [{ type: "text", text: truncateIfNeeded(text) }],
          structuredContent: report,
        };
      } catch (err) {
        return { content: [{ type: "text", text: handleGraphError(err) }] };
      }
    }
  );
}

// ─── Tool: outlook_list_unread_emails ─────────────────────────────────────────

export function registerListUnreadEmails(server: McpServer): void {
  server.registerTool(
    "outlook_list_unread_emails",
    {
      title: "List Unread Emails",
      description: `Fetch a paginated list of unread emails with metadata (no categorisation).
Use when you want raw unread email data, or when fetching a specific page.

Args:
  - limit: Max emails per page (default 20, max 100)
  - offset: Skip N emails for pagination
  - folder_id: Restrict to a specific mail folder (optional)
  - response_format: 'markdown' or 'json'

Returns:
  Paginated email list with: id, subject, from, to, cc, preview, receivedAt,
  importance, hasAttachments, flagged, webLink, plus pagination metadata.`,
      inputSchema: z
        .object({
          ...paginationSchema,
          folder_id: z
            .string()
            .optional()
            .describe("Optional folder ID to restrict results"),
          response_format: responseFormatSchema,
        })
        .strict(),
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async ({ limit, offset, folder_id, response_format }) => {
      try {
        const data = await graph.listUnreadMessages({
          limit,
          skip: offset,
          folderId: folder_id,
        });

        const messages = data.value ?? [];
        const total = data["@odata.count"] ?? messages.length;

        if (messages.length === 0) {
          return {
            content: [{ type: "text", text: "No unread emails found." }],
          };
        }

        const summaries: EmailSummary[] = messages.map(toEmailSummary);
        const hasMore = total > offset + summaries.length;

        const output = {
          total,
          count: summaries.length,
          offset,
          items: summaries,
          has_more: hasMore,
          ...(hasMore ? { next_offset: offset + summaries.length } : {}),
        };

        let text: string;
        if (response_format === ResponseFormat.MARKDOWN) {
          const lines = [
            `# 📬 Unread Emails (${summaries.length} of ${total})`,
            "",
            ...summaries.map((e) => formatEmailSummaryMd(e) + "\n---"),
            hasMore
              ? `\n*Use offset=${offset + summaries.length} to load more.*`
              : "",
          ];
          text = lines.join("\n");
        } else {
          text = JSON.stringify(output, null, 2);
        }

        return {
          content: [{ type: "text", text: truncateIfNeeded(text) }],
          structuredContent: output,
        };
      } catch (err) {
        return { content: [{ type: "text", text: handleGraphError(err) }] };
      }
    }
  );
}

// ─── Tool: outlook_get_email ──────────────────────────────────────────────────

export function registerGetEmail(server: McpServer): void {
  server.registerTool(
    "outlook_get_email",
    {
      title: "Get Email",
      description: `Fetch the full content of a single email by ID, including body (HTML or plain text), all recipients, and attachments indicator.

Args:
  - message_id: The email ID (from list/triage tools)
  - response_format: 'markdown' or 'json'

Returns:
  Full email details including body content, all recipients, and metadata.`,
      inputSchema: z
        .object({
          message_id: z
            .string()
            .min(1)
            .describe("The email message ID"),
          response_format: responseFormatSchema,
        })
        .strict(),
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async ({ message_id, response_format }) => {
      try {
        const msg = await graph.getMessage(message_id);
        const summary = toEmailSummary(msg);

        const bodyText =
          msg.body?.content
            ?.replace(/<[^>]+>/g, " ")
            .replace(/\s+/g, " ")
            .trim() ?? msg.bodyPreview ?? "";

        const output = {
          ...summary,
          body: bodyText,
          bodyContentType: msg.body?.contentType ?? "text",
          isDraft: msg.isDraft,
          sentAt: msg.sentDateTime,
          conversationId: msg.conversationId,
          internetMessageId: msg.internetMessageId,
        };

        let text: string;
        if (response_format === ResponseFormat.MARKDOWN) {
          text = [
            `# ${summary.subject}`,
            `**From**: ${summary.from}`,
            `**To**: ${summary.to.join(", ")}`,
            summary.cc.length ? `**CC**: ${summary.cc.join(", ")}` : "",
            `**Received**: ${new Date(summary.receivedAt).toLocaleString()}`,
            `**Importance**: ${summary.importance}`,
            summary.flagged ? "**Flagged**: Yes" : "",
            summary.hasAttachments ? "**Has Attachments**: Yes" : "",
            `**Web Link**: ${summary.webLink}`,
            "",
            "## Body",
            bodyText,
          ]
            .filter(Boolean)
            .join("\n");
        } else {
          text = JSON.stringify(output, null, 2);
        }

        return {
          content: [{ type: "text", text: truncateIfNeeded(text) }],
          structuredContent: output,
        };
      } catch (err) {
        return { content: [{ type: "text", text: handleGraphError(err) }] };
      }
    }
  );
}

// ─── Tool: outlook_reply_email ────────────────────────────────────────────────

export function registerReplyEmail(server: McpServer): void {
  server.registerTool(
    "outlook_reply_email",
    {
      title: "Reply to Email",
      description: `Send a reply (or reply-all) to an email.

Args:
  - message_id: The email to reply to
  - comment: The reply body text (plain text)
  - reply_all: If true, replies to all recipients; if false (default), replies only to sender

Returns:
  Confirmation message on success.`,
      inputSchema: z
        .object({
          message_id: z.string().min(1).describe("Email message ID to reply to"),
          comment: z
            .string()
            .min(1)
            .max(50000)
            .describe("Reply body text (plain text)"),
          reply_all: z
            .boolean()
            .default(false)
            .describe("If true, reply to all recipients; default false"),
        })
        .strict(),
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    async ({ message_id, comment, reply_all }) => {
      try {
        await graph.replyToMessage(message_id, comment, reply_all);
        return {
          content: [
            {
              type: "text",
              text: `✅ Reply sent successfully${reply_all ? " (reply-all)" : ""}.`,
            },
          ],
        };
      } catch (err) {
        return { content: [{ type: "text", text: handleGraphError(err) }] };
      }
    }
  );
}

// ─── Tool: outlook_forward_email ──────────────────────────────────────────────

export function registerForwardEmail(server: McpServer): void {
  server.registerTool(
    "outlook_forward_email",
    {
      title: "Forward Email",
      description: `Forward an email to one or more recipients with an optional comment.

Args:
  - message_id: The email to forward
  - to_addresses: Array of recipient email addresses
  - comment: Optional introductory message to prepend

Returns:
  Confirmation message on success.`,
      inputSchema: z
        .object({
          message_id: z.string().min(1).describe("Email message ID to forward"),
          to_addresses: z
            .array(z.string().email())
            .min(1)
            .describe("Recipient email addresses"),
          comment: z
            .string()
            .max(10000)
            .optional()
            .describe("Optional message to prepend to the forwarded email"),
        })
        .strict(),
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    async ({ message_id, to_addresses, comment }) => {
      try {
        await graph.forwardMessage(message_id, to_addresses, comment);
        return {
          content: [
            {
              type: "text",
              text: `✅ Email forwarded to: ${to_addresses.join(", ")}`,
            },
          ],
        };
      } catch (err) {
        return { content: [{ type: "text", text: handleGraphError(err) }] };
      }
    }
  );
}

// ─── Tool: outlook_mark_as_read ───────────────────────────────────────────────

export function registerMarkAsRead(server: McpServer): void {
  server.registerTool(
    "outlook_mark_as_read",
    {
      title: "Mark Email as Read/Unread",
      description: `Mark one or more emails as read or unread.

Args:
  - message_ids: Array of email IDs to update
  - is_read: true to mark as read, false to mark as unread

Returns:
  Summary of how many emails were updated.`,
      inputSchema: z
        .object({
          message_ids: z
            .array(z.string().min(1))
            .min(1)
            .max(50)
            .describe("Array of email message IDs (up to 50)"),
          is_read: z
            .boolean()
            .default(true)
            .describe("true = mark as read, false = mark as unread"),
        })
        .strict(),
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async ({ message_ids, is_read }) => {
      const results: string[] = [];
      const failed: string[] = [];

      for (const id of message_ids) {
        try {
          await graph.markMessageRead(id, is_read);
          results.push(id);
        } catch (err) {
          failed.push(`${id}: ${handleGraphError(err)}`);
        }
      }

      const status = is_read ? "read" : "unread";
      const lines = [`✅ Marked ${results.length} email(s) as ${status}.`];
      if (failed.length > 0) {
        lines.push(`\n⚠️ Failed (${failed.length}):`);
        lines.push(...failed.map((f) => `  - ${f}`));
      }

      return { content: [{ type: "text", text: lines.join("\n") }] };
    }
  );
}

// ─── Tool: outlook_flag_email ─────────────────────────────────────────────────

export function registerFlagEmail(server: McpServer): void {
  server.registerTool(
    "outlook_flag_email",
    {
      title: "Flag / Unflag Email",
      description: `Set or clear a follow-up flag on an email.

Args:
  - message_id: Email to flag/unflag
  - flag_status: 'flagged' (follow up), 'complete' (done), or 'notFlagged' (clear)

Returns:
  Confirmation message.`,
      inputSchema: z
        .object({
          message_id: z.string().min(1).describe("Email message ID"),
          flag_status: z
            .enum(["flagged", "complete", "notFlagged"])
            .describe("Flag status to set"),
        })
        .strict(),
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async ({ message_id, flag_status }) => {
      try {
        await graph.flagMessage(message_id, flag_status);
        return {
          content: [
            {
              type: "text",
              text: `✅ Email flag updated to "${flag_status}".`,
            },
          ],
        };
      } catch (err) {
        return { content: [{ type: "text", text: handleGraphError(err) }] };
      }
    }
  );
}

// ─── Tool: outlook_move_email ─────────────────────────────────────────────────

export function registerMoveEmail(server: McpServer): void {
  server.registerTool(
    "outlook_move_email",
    {
      title: "Move Email to Folder",
      description: `Move one or more emails to a different mail folder.
Use outlook_list_folders first to get valid folder IDs or names.

Args:
  - message_ids: Array of email message IDs to move (up to 20)
  - destination_folder: Destination folder ID or well-known name (e.g. "Inbox", "Archive", "DeletedItems", "Junk")

Returns:
  Count of emails moved and any failures.`,
      inputSchema: z
        .object({
          message_ids: z
            .array(z.string().min(1))
            .min(1)
            .max(20)
            .describe("Email message IDs to move (up to 20)"),
          destination_folder: z
            .string()
            .min(1)
            .describe(
              "Destination folder ID or well-known name (Inbox, Archive, DeletedItems, Junk, Drafts, SentItems)"
            ),
        })
        .strict(),
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    async ({ message_ids, destination_folder }) => {
      // Resolve name → ID if needed
      let folderId = destination_folder;
      const wellKnown: Record<string, string> = {
        inbox: "inbox",
        archive: "archive",
        deleteditems: "deleteditems",
        junk: "junkemail",
        drafts: "drafts",
        sentitems: "sentitems",
        clutter: "clutter",
      };
      const normalized = destination_folder.toLowerCase().replace(/\s/g, "");
      if (wellKnown[normalized]) {
        folderId = wellKnown[normalized];
      } else {
        // Try to resolve by display name
        try {
          const folder = await graph.resolveFolder(destination_folder);
          if (folder) folderId = folder.id;
        } catch {
          // Use as-is (may be a raw ID)
        }
      }

      const moved: string[] = [];
      const failed: string[] = [];

      for (const id of message_ids) {
        try {
          await graph.moveMessage(id, folderId);
          moved.push(id);
        } catch (err) {
          failed.push(`${id}: ${handleGraphError(err)}`);
        }
      }

      const lines = [`✅ Moved ${moved.length} email(s) to "${destination_folder}".`];
      if (failed.length > 0) {
        lines.push(`\n⚠️ Failed (${failed.length}):`);
        lines.push(...failed.map((f) => `  - ${f}`));
      }

      return { content: [{ type: "text", text: lines.join("\n") }] };
    }
  );
}

// ─── Tool: outlook_delete_email ───────────────────────────────────────────────

export function registerDeleteEmail(server: McpServer): void {
  server.registerTool(
    "outlook_delete_email",
    {
      title: "Delete Email",
      description: `Permanently delete an email. This moves it to Deleted Items; it does NOT bypass the Deleted Items folder.
Prefer outlook_move_email to "DeletedItems" for recoverable deletion.

Args:
  - message_id: Email to delete

Returns:
  Confirmation message.`,
      inputSchema: z
        .object({
          message_id: z.string().min(1).describe("Email message ID to delete"),
        })
        .strict(),
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    async ({ message_id }) => {
      try {
        await graph.deleteMessage(message_id);
        return {
          content: [{ type: "text", text: "✅ Email deleted (moved to Deleted Items)." }],
        };
      } catch (err) {
        return { content: [{ type: "text", text: handleGraphError(err) }] };
      }
    }
  );
}

// ─── Tool: outlook_search_emails ──────────────────────────────────────────────

export function registerSearchEmails(server: McpServer): void {
  server.registerTool(
    "outlook_search_emails",
    {
      title: "Search Emails",
      description: `Full-text search across your mailbox using Microsoft Search syntax.
Searches subject, body, and sender fields.

Supported query examples:
  - "meeting notes"          → phrase search
  - "from:alice@example.com" → from a sender
  - "subject:invoice"        → subject contains word
  - "hasAttachment:true"     → has attachments

Args:
  - query: Search query string
  - limit: Max results (default 20)
  - offset: Skip N results for pagination
  - folder_id: Restrict to a folder (optional)
  - response_format: 'markdown' or 'json'

Returns:
  Matching emails with metadata.`,
      inputSchema: z
        .object({
          query: z
            .string()
            .min(1)
            .max(500)
            .describe("Search query string"),
          ...paginationSchema,
          folder_id: z
            .string()
            .optional()
            .describe("Optional folder ID to restrict search"),
          response_format: responseFormatSchema,
        })
        .strict(),
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async ({ query, limit, offset, folder_id, response_format }) => {
      try {
        const data = await graph.searchMessages({
          query,
          limit,
          skip: offset,
          folderId: folder_id,
        });

        const messages = data.value ?? [];
        if (messages.length === 0) {
          return {
            content: [{ type: "text", text: `No emails found matching: "${query}"` }],
          };
        }

        const summaries = messages.map(toEmailSummary);
        const hasMore = Boolean(data["@odata.nextLink"]);
        const output = {
          query,
          count: summaries.length,
          offset,
          items: summaries,
          has_more: hasMore,
          ...(hasMore ? { next_offset: offset + summaries.length } : {}),
        };

        let text: string;
        if (response_format === ResponseFormat.MARKDOWN) {
          text = [
            `# 🔍 Search Results for "${query}" (${summaries.length} results)`,
            "",
            ...summaries.map((e) => formatEmailSummaryMd(e) + "\n---"),
          ].join("\n");
        } else {
          text = JSON.stringify(output, null, 2);
        }

        return {
          content: [{ type: "text", text: truncateIfNeeded(text) }],
          structuredContent: output,
        };
      } catch (err) {
        return { content: [{ type: "text", text: handleGraphError(err) }] };
      }
    }
  );
}

// ─── Tool: outlook_list_folders ───────────────────────────────────────────────

export function registerListFolders(server: McpServer): void {
  server.registerTool(
    "outlook_list_folders",
    {
      title: "List Mail Folders",
      description: `List all mail folders in your Outlook mailbox with unread counts.
Use this to get folder IDs for use in other tools.

Args:
  - include_hidden: Include hidden system folders (default false)
  - response_format: 'markdown' or 'json'

Returns:
  List of folders with: id, displayName, totalItemCount, unreadItemCount`,
      inputSchema: z
        .object({
          include_hidden: z
            .boolean()
            .default(false)
            .describe("Include hidden system folders (default false)"),
          response_format: responseFormatSchema,
        })
        .strict(),
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async ({ include_hidden, response_format }) => {
      try {
        const folders = await graph.listFolders(include_hidden);

        if (response_format === ResponseFormat.MARKDOWN) {
          const lines = [
            "# 📁 Mail Folders",
            "",
            "| Folder | Unread | Total | ID |",
            "|--------|--------|-------|----|",
            ...folders.map(
              (f) =>
                `| ${f.displayName} | ${f.unreadItemCount} | ${f.totalItemCount} | \`${f.id}\` |`
            ),
          ];
          return {
            content: [{ type: "text", text: lines.join("\n") }],
            structuredContent: folders,
          };
        } else {
          return {
            content: [{ type: "text", text: JSON.stringify(folders, null, 2) }],
            structuredContent: folders,
          };
        }
      } catch (err) {
        return { content: [{ type: "text", text: handleGraphError(err) }] };
      }
    }
  );
}

// ─── R