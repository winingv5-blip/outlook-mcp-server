/**
 * Heuristic email triage engine.
 *
 * Analyses an email's metadata and body preview and assigns one of three
 * categories:
 *   - Action Required   → you need to do something
 *   - Waiting On        → you sent something and are awaiting a reply
 *   - Reference/Read    → newsletter, receipt, FYI, or informational
 *
 * The LLM that calls this MCP will have rich context and can override the
 * suggestion — the heuristics are a first-pass signal, not a final verdict.
 */

import type { GraphMessage } from "./types.js";
import {
  CATEGORY,
  type TriageCategory,
  NEWSLETTER_SENDER_PATTERNS,
  ACTION_KEYWORDS,
  RECEIPT_KEYWORDS,
  SPAM_KEYWORDS,
  ARCHIVE_SENDER_PATTERNS,
  EXPIRED_NOTIFICATION_PATTERNS,
} from "./constants.js";
import type { EmailSummary, TriagedEmail } from "./types.js";

// ─── Helpers ──────────────────────────────────────────────────────────────────

function formatAddress(recipient: { emailAddress: { name: string; address: string } }): string {
  const { name, address } = recipient.emailAddress;
  return name ? `${name} <${address}>` : address;
}

function recipientAddresses(recipients: Array<{ emailAddress: { address: string } }>): string[] {
  return recipients.map((r) => r.emailAddress.address.toLowerCase());
}

function userEmail(): string {
  return (process.env.OUTLOOK_USER_EMAIL ?? "").toLowerCase();
}

/** Extract signals and pick a category + one-sentence rationale */
export function triageMessage(
  msg: GraphMessage
): { category: TriageCategory; rationale: string; signals: string[] } {
  const signals: string[] = [];
  const text = `${msg.subject ?? ""} ${msg.bodyPreview ?? ""}`;
  const senderAddress = msg.from?.emailAddress?.address?.toLowerCase() ?? "";

  // ── Waiting On signals ──────────────────────────────────────────────────────
  //  We can't inspect "Sent" items here, but we can flag emails where the
  //  user is NOT listed as a recipient (meaning this was auto-generated or
  //  the user is checking a shared folder). More reliably: if the email
  //  subject starts with "Re:" and the user's own address is the original
  //  sender, that's a reply they were waiting for.
  const myEmail = userEmail();
  const toAddresses = recipientAddresses(msg.toRecipients ?? []);
  const ccAddresses = recipientAddresses(msg.ccRecipients ?? []);
  const isDirectlyAddressed = myEmail
    ? toAddresses.includes(myEmail)
    : toAddresses.length > 0;
  const isOnlyCC = myEmail
    ? !toAddresses.includes(myEmail) && ccAddresses.includes(myEmail)
    : false;
  const isReply = /^\s*Re:/i.test(msg.subject ?? "");

  if (isReply && !isDirectlyAddressed) {
    signals.push("reply-to-your-message");
  }

  // ── Newsletter / automated sender ──────────────────────────────────────────
  const isAutomatedSender = NEWSLETTER_SENDER_PATTERNS.some((re) =>
    re.test(senderAddress)
  );
  if (isAutomatedSender) signals.push("automated-sender");

  // ── Archive/Delete signals ─────────────────────────────────────────────────
  const isArchiveSender = ARCHIVE_SENDER_PATTERNS.some((re) =>
    re.test(senderAddress)
  );
  if (isArchiveSender) signals.push("archive-sender");

  const isSpam = SPAM_KEYWORDS.some((re) => re.test(text));
  if (isSpam) signals.push("spam-keyword");

  const isExpiredNotification = EXPIRED_NOTIFICATION_PATTERNS.some((re) =>
    re.test(text)
  );
  if (isExpiredNotification) signals.push("expired-notification");

  // ── Receipt / reference content ────────────────────────────────────────────
  const isReceipt = RECEIPT_KEYWORDS.some((re) => re.test(text));
  if (isReceipt) signals.push("receipt-or-confirmation");

  // ── Action keywords ────────────────────────────────────────────────────────
  const hasActionKeywords = ACTION_KEYWORDS.some((re) => re.test(text));
  if (hasActionKeywords) signals.push("action-keyword");

  // ── Importance / flag ──────────────────────────────────────────────────────
  if (msg.importance === "high") signals.push("marked-high-importance");
  if (msg.flag?.flagStatus === "flagged") signals.push("flagged");
  if (isOnlyCC) signals.push("cc-only");
  if (isDirectlyAddressed && !isOnlyCC) signals.push("direct-recipient");

  // ── Decision ───────────────────────────────────────────────────────────────
  // Archive/Delete wins first — if something is clearly spam or an expired
  // notification, no amount of action keywords should override that.
  let category: TriageCategory;
  let rationale: string;

  if (signals.includes("spam-keyword")) {
    category = CATEGORY.ARCHIVE_DELETE;
    rationale =
      "Contains language patterns consistent with spam or unsolicited promotional mail — safe to delete.";
  } else if (
    signals.includes("archive-sender") ||
    signals.includes("expired-notification")
  ) {
    const reason = signals.includes("expired-notification")
      ? "expired or automated CI/digest notification"
      : "system/mailer-daemon sender";
    category = CATEGORY.ARCHIVE_DELETE;
    rationale = `Appears to be a ${reason} with no ongoing value — safe to archive.`;
  } else if (
    signals.includes("reply-to-your-message") &&
    !signals.includes("automated-sender")
  ) {
    category = CATEGORY.WAITING_ON;
    rationale =
      "This appears to be a reply to a message you sent, suggesting you were awaiting a response.";
  } else if (
    signals.includes("automated-sender") ||
    signals.includes("receipt-or-confirmation") ||
    signals.includes("cc-only")
  ) {
    const reason = signals.includes("receipt-or-confirmation")
      ? "receipt or order confirmation"
      : signals.includes("automated-sender")
        ? "automated sender"
        : "you are CC'd only";
    category = CATEGORY.REFERENCE_READ;
    rationale = `Classified as reference material because it appears to be a ${reason} that doesn't require action.`;
  } else if (
    signals.includes("action-keyword") ||
    signals.includes("marked-high-importance") ||
    signals.includes("flagged") ||
    signals.includes("direct-recipient")
  ) {
    category = CATEGORY.ACTION_REQUIRED;
    const reasons: string[] = [];
    if (signals.includes("action-keyword")) reasons.push("action-oriented language");
    if (signals.includes("marked-high-importance")) reasons.push("high importance");
    if (signals.includes("flagged")) reasons.push("fla