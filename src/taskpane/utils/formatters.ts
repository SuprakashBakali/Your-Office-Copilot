// ============================================================
// Formatters — Data display, export, and conversion helpers
// ============================================================

import { CellData, ChatConversation, ChatMessage, ExportFormat } from "../types";

/** Format a number for display (with locale-aware separators) */
export function formatNumber(n: number, decimals = 2): string {
  if (Number.isInteger(n) && decimals === 2) return n.toLocaleString();
  return n.toLocaleString(undefined, {
    minimumFractionDigits: 0,
    maximumFractionDigits: decimals,
  });
}

/** Format bytes to human readable */
export function formatBytes(bytes: number): string {
  if (bytes === 0) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${(bytes / Math.pow(k, i)).toFixed(1)} ${sizes[i]}`;
}

/** Format timestamp to relative time */
export function timeAgo(timestamp: number): string {
  const seconds = Math.floor((Date.now() - timestamp) / 1000);
  if (seconds < 60) return "just now";
  if (seconds < 3600) return `${Math.floor(seconds / 60)}m ago`;
  if (seconds < 86400) return `${Math.floor(seconds / 3600)}h ago`;
  if (seconds < 604800) return `${Math.floor(seconds / 86400)}d ago`;
  return new Date(timestamp).toLocaleDateString();
}

/** Format timestamp to readable date/time */
export function formatDateTime(timestamp: number): string {
  return new Date(timestamp).toLocaleString(undefined, {
    month: "short",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

/** Convert cell data to a readable text table */
export function cellDataToText(data: CellData): string {
  if (!data.values || data.values.length === 0) return "(empty selection)";

  const rows = data.values;
  // Calculate column widths
  const colWidths: number[] = [];
  for (const row of rows) {
    for (let c = 0; c < row.length; c++) {
      const len = String(row[c] ?? "").length;
      colWidths[c] = Math.max(colWidths[c] || 0, Math.min(len, 30));
    }
  }

  const lines = rows.map((row) =>
    row.map((cell, i) => String(cell ?? "").padEnd(colWidths[i] || 0)).join(" | ")
  );

  // Add header separator
  if (lines.length > 1) {
    const sep = colWidths.map((w) => "-".repeat(w)).join("-+-");
    lines.splice(1, 0, sep);
  }

  return `Range: ${data.address} (${data.sheetName})\n${data.rowCount}×${data.columnCount}\n\n${lines.join("\n")}`;
}

/** Convert cell data to CSV string */
export function cellDataToCSV(data: CellData): string {
  return data.values
    .map((row) =>
      row.map((cell) => {
        const s = String(cell ?? "");
        return s.includes(",") || s.includes('"') || s.includes("\n")
          ? `"${s.replace(/"/g, '""')}"`
          : s;
      }).join(",")
    )
    .join("\n");
}

/** Truncate text with ellipsis */
export function truncate(text: string, maxLength: number): string {
  if (text.length <= maxLength) return text;
  return text.substring(0, maxLength - 3) + "...";
}

/** Create a title for a conversation from the first message */
export function generateConversationTitle(firstMessage: string): string {
  const cleaned = firstMessage.replace(/\n/g, " ").trim();
  return truncate(cleaned, 50);
}

/** Export a conversation to the given format */
export function exportConversation(
  conversation: ChatConversation,
  format: ExportFormat
): string {
  switch (format) {
    case "markdown":
      return conversationToMarkdown(conversation);
    case "json":
      return JSON.stringify(conversation, null, 2);
    case "text":
      return conversationToText(conversation);
    case "html":
      return conversationToHTML(conversation);
    default:
      return conversationToText(conversation);
  }
}

function conversationToMarkdown(convo: ChatConversation): string {
  const lines = [
    `# ${convo.title}`,
    `*${new Date(convo.createdAt).toLocaleString()}* | ${convo.provider} / ${convo.model}`,
    "",
  ];
  for (const msg of convo.messages) {
    if (msg.role === "system") continue;
    const label = msg.role === "user" ? "**You**" : "**AI**";
    lines.push(`### ${label}`, "", msg.content, "");
  }
  return lines.join("\n");
}

function conversationToText(convo: ChatConversation): string {
  const lines = [`${convo.title}`, `${new Date(convo.createdAt).toLocaleString()}`, ""];
  for (const msg of convo.messages) {
    if (msg.role === "system") continue;
    const label = msg.role === "user" ? "You" : "AI";
    lines.push(`[${label}]`, msg.content, "");
  }
  return lines.join("\n");
}

function conversationToHTML(convo: ChatConversation): string {
  const msgs = convo.messages
    .filter((m) => m.role !== "system")
    .map(
      (m) =>
        `<div class="message ${m.role}"><strong>${m.role === "user" ? "You" : "AI"}</strong><p>${escapeHtml(m.content)}</p></div>`
    )
    .join("\n");

  return `<!DOCTYPE html>
<html><head><title>${escapeHtml(convo.title)}</title>
<style>body{font-family:system-ui;max-width:800px;margin:0 auto;padding:20px}.user{background:#e8f0fe;padding:12px;border-radius:8px;margin:8px 0}.assistant{background:#f0f0f0;padding:12px;border-radius:8px;margin:8px 0}</style>
</head><body><h1>${escapeHtml(convo.title)}</h1>${msgs}</body></html>`;
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Download a string as a file */
export function downloadAsFile(content: string, filename: string, mimeType: string): void {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
