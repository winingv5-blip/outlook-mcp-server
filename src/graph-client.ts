/**
 * Microsoft Graph API client.
 *
 * Wraps axios with automatic token injection, structured error handling,
 * and typed responses.
 */

import axios, { type AxiosError, type AxiosRequestConfig } from "axios";
import { getAccessToken } from "./auth.js";
import { GRAPH_BASE_URL } from "./constants.js";
import type {
  GraphMessage,
  GraphMessageListResponse,
  GraphFolder,
  GraphFolderListResponse,
} from "./types.js";

// ─── Error handling ───────────────────────────────────────────────────────────

export function handleGraphError(error: unknown): string {
  if (axios.isAxiosError(error)) {
    const axiosErr = error as AxiosError<{ error?: { message?: string } }>;
    if (axiosErr.response) {
      const status = axiosErr.response.status;
      const msg = axiosErr.response.data?.error?.message ?? "";
      switch (status) {
        case 400:
          return `Error: Bad request — ${msg}. Check your query parameters.`;
        case 401:
          return "Error: Authentication failed. Run with OUTLOOK_CLIENT_ID set and re-authenticate.";
        case 403:
          return `Error: Permission denied — ${msg}. Ensure the app has Mail.Read and Mail.ReadWrite scopes.`;
        case 404:
          return "Error: Resource not found. The email or folder ID may be invalid or already deleted.";
        case 429:
          return "Error: Rate limit exceeded. Please wait a moment before retrying.";
        case 503:
          return "Error: Microsoft Graph is temporarily unavailable. Please retry.";
        default:
          return `Error: Graph API returned HTTP ${status} — ${msg}`;
      }
    } else if (axiosErr.code === "ECONNABORTED") {
      return "Error: Request timed out. Check your network connection.";
    } else if (axiosErr.code === "ENOTFOUND") {
      return "Error: Cannot reach graph.microsoft.com. Check your internet connection.";
    }
  }
  return `Error: ${error instanceof Error ? error.message : String(error)}`;
}

// ─── Core request helper ──────────────────────────────────────────────────────

async function graphRequest<T>(
  method: "GET" | "POST" | "PATCH" | "DELETE",
  endpoint: string,
  data?: unknown,
  params?: Record<string, string | number | boolean>
): Promise<T> {
  const token = await getAccessToken();

  const config: AxiosRequestConfig = {
    method,
    url: endpoint.startsWith("http") ? endpoint : `${GRAPH_BASE_URL}${endpoint}`,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    data,
    params,
    timeout: 30_000,
  };

  const response = await axios(config);
  return response.data as T;
}

// ─── Select fields for list queries (perf optimisation) ──────────────────────

const MESSAGE_SELECT_FIELDS = [
  "id",
  "subject",
  "bodyPreview",
  "from",
  "toRecipients",
  "ccRecipients",
  "receivedDateTime",
  "sentDateTime",
  "isRead",
  "importance",
  "hasAttachments",
  "flag",
  "categories",
  "conversationId",
  "webLink",
  "parentFolderId",
].join(",");

const MESSAGE_FULL_SELECT_FIELDS = [
  MESSAGE_SELECT_FIELDS,
  "body",
  "replyTo",
  "bccRecipients",
  "internetMessageId",
  "isDraft",
].join(",");

// ─── Message APIs ─────────────────────────────────────────────────────────────

export async function listUnreadMessages(opts: {
  limit: number;
  skip: number;
  folderId?: string;
  orderBy?: string;
}): Promise<GraphMessageListResponse> {
  const folder = opts.folderId ? `/mailFolders/${opts.folderId}` : "";
  return graphRequest<GraphMessageListResponse>(
    "GET",
    `/me${folder}/messages`,
    undefined,
    {
      $filter: "isRead eq false",
      $select: MESSAGE_SELECT_FIELDS,
      $top: opts.limit,
      $skip: opts.skip,
      $orderby: opts.orderBy ?? "receivedDateTime desc",
      $count: "true",
    }
  );
}

export async function searchMessages(opts: {
  query: string;
  limit: number;
  skip: number;
  folderId?: string;
}): Promise<GraphMessageListResponse> {
  const folder = opts.folderId ? `/mailFolders/${opts.folderId}` : "";
  return graphRequest<GraphMessageListResponse>(
    "GET",
    `/me${folder}/messages`,
    undefined,
    {
      $search: `"${opts.query}"`,
      $select: MESSAGE_SELECT_FIELDS,
      $top: opts.limit,
      $skip: opts.skip,
    }
  );
}

export async function getMessage(messageId: string): Promise<GraphMessage> {
  return graphRequest<GraphMessage>(
    "GET",
    `/me/messages/${messageId}`,
    undefined,
    { $select: MESSAGE_FULL_SELECT_FIELDS }
  );
}

export async function markMessageRead(
  messageId: string,
  isRead: boolean
): Promise<void> {
  await graphRequest<unknown>("PATCH", `/me/messages/${messageId}`, { isRead });
}

export async function moveMessage(
  messageId: string,
  destinationFolderId: string
): Promise<GraphMessage> {
  return graphRequest<GraphMessage>(
    "POST",
    `/me/messages/${messageId}/move`,
    { destinationId: destinationFolderId }
  );
}

export async function replyToMessage(
  messageId: string,
  comment: string,
  replyAll: boolean
): Promise<void> {
  const endpoint = replyAll
    ? `/me/messages/${messageId}/replyAll`
    : `/me/messages/${messageId}/reply`;
  await graphRequest<unknown>("POST", endpoint, { comment });
}

export async function forwardMessage(
  messageId: string,
  toAddresses: string[],
  comment?: string
): Promise<void> {
  await graphRequest<unknown>("POST", `/me/messages/${messageId}/forward`, {
    comment: comment ?? "",
    toRecipients: toAddresses.map((addr) => ({
      emailAddress: { address: addr },
    })),
  });
}

export async function deleteMessage(messageId: string): Promise<void> {
  await graphRequest<unknown>("DELETE", `/me/messages/${messageId}`);
}

export async function flagMessage(
  messageId: string,
  flagStatus: "flagged" | "notFlagged" | "complete"
): Promise<void> {
  await graphRequest<unknown>("PATCH", `/me/messages/${messageId}`, {
    flag: { flagStatus },
  });
}

// ─── Folder APIs ──────────────────────────────────────────────────────────────

export async function listFolders(includeHidden = false): Promise<GraphFolder[]> {
  const data = await graphRequest<GraphFolderListResponse>(
    "GET",
    "/me/mailFolders",
    undefined,
    {
      $select: "id,displayName,totalItemCount,unreadItemCount,childFolderCount",
      includeHiddenFolders: includeHidden ? "true" : "false",
      $top: 100,
    }
  );
  return data.value;
}

export async function resolveFolder(nameOrId: string): Promise<GraphFolder | null> {
  const folders = await listFolders();
  const found = folders.find(
    (f) =>
      f.id === nameOrId ||
      f.displayName.toLowerCase() === nameOrId.toLowerCase()
  );
  return found ?? null;
}
