import type { TriageCategory } from "./constants.js";

// ─── Microsoft Graph types ────────────────────────────────────────────────────

export interface GraphEmailAddress {
  name: string;
  address: string;
}

export interface GraphRecipient {
  emailAddress: GraphEmailAddress;
}

export interface GraphBodyContent {
  contentType: "text" | "html";
  content: string;
}

export interface GraphMessage {
  id: string;
  subject: string;
  bodyPreview: string;
  body?: GraphBodyContent;
  from: GraphRecipient;
  toRecipients: GraphRecipient[];
  ccRecipients: GraphRecipient[];
  bccRecipients: GraphRecipient[];
  replyTo: GraphRecipient[];
  receivedDateTime: string;
  sentDateTime: string;
  isRead: boolean;
  isDraft: boolean;
  importance: "low" | "normal" | "high";
  hasAttachments: boolean;
  conversationId: string;
  internetMessageId: string;
  categories: string[];
  flag: {
    flagStatus: "notFlagged" | "flagged" | "complete";
  };
  parentFolderId: string;
  webLink: string;
}

export interface GraphMessageListResponse {
  value: GraphMessage[];
  "@odata.nextLink"?: string;
  "@odata.count"?: number;
}

export interface GraphFolder {
  id: string;
  displayName: string;
  totalItemCount: number;
  unreadItemCount: number;
  childFolderCount: number;
}

export interface GraphFolderListResponse {
  value: GraphFolder[];
  "@odata.nextLink"?: string;
}

// ─── MCP-layer types ─────────────────────────────────────────────────────────

export interface EmailSummary {
  id: string;
  subject: string;
  from: string;
  fromAddress: string;
  to: string[];
  cc: string[];
  preview: string;
  receivedAt: string;
  isRead: boolean;
  importance: string;
  hasAttachments: boolean;
  flagged: boolean;
  webLink: string;
}

export interface TriagedEmail extends EmailSummary {
  suggestedCategory: TriageCategory;
  rationale: string;
  signals: string[];
}

export interface TriageReport {
  total: number;
  analyzedAt: string;
  categories: {
    actionRequired: TriagedEmail[];
    waitingOn: TriagedEmail[];
    referenceRead: TriagedEmail[];
    archiveDelete: TriagedEmail[];
  };
}

export interface PaginatedResponse<T> {
  total: number;
  count: number;
  offset: number;
  items: T[];
  has_more: boolean;
  next_offset