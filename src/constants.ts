// Microsoft Graph API base URL
export const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

// Default scopes for mail access
export const MAIL_SCOPES = [
  "Mail.Read",
  "Mail.ReadWrite",
  "Mail.Send",
  "User.Read",
];

// Response size cap (characters)
export const CHARACTER_LIMIT = 25000;

// Default pagination
export const DEFAULT_LIMIT = 20;
export const MAX_LIMIT = 100;

// Token cache file (relative to CWD / home dir)
export const TOKEN_CACHE_FILE = ".outlook-mcp-token-cache.json";

// Triage category labels
export const CATEGORY = {
  ACTION_REQUIRED: "Action Required",
  WAITING_ON: "Waiting On",
  REFERENCE_READ: "Reference/Read",
  ARCHIVE_DELETE: "Archive/Delete",
} as const;

export type TriageCategory = (typeof CATEGORY)[keyof typeof CATEGORY];

// Heuristic sender patterns that indicate newsletters / automated mail
export const NEWSLETTER_SENDER_PATTERNS = [
  /noreply/i,
  /no-reply/i,
  /donotreply/i,
  /notifications?@/i,
  /newsletter/i,
  /mailer@/i,
  /bounce@/i,
  /updates?@/i,
];

// Keywords that strongly suggest the recipient must act
export const ACTION_KEYWORDS = [
  /please\s+(review|respond|confirm|approve|complete|sign|submit|reply|let\s+me\s+know)/i,
  /action\s+required/i,
  /your\s+(approval|response|input|feedback|decision)\s+is\s+(needed|required|requested)/i,
  /can\s+you\s+/i,
  /could\s+you\s+/i,
  /would\s+you\s+/i,
  /follow\s+up/i,
  /deadline/i,
  /due\s+(date|by)/i,
  /asap/i,
  /urgent/i,
  /by\s+(end\s+of\s+(day|week|month)|(monday|tuesday|wednesday|thursday|friday|saturday|sunday))/i,
];

// Receipt / reference keywords
export const RECEIPT_KEYWORDS = [
  /your\s+order/i,
  /order\s+confirmation/i,
  /receipt/i,
  /invoice/i,
  /statement/i,
  /payment\s+(received|confirmed)/i,
  /subscription\s+(renewal|confirmed)/i,
  /shipping\s+confirm