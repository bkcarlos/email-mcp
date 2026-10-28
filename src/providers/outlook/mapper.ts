import type { Folder, Email, Contact, AttachmentMeta, FolderTypeValue, SearchQuery } from '../../models/types.js';
import { FolderType } from '../../models/types.js';

const WELL_KNOWN_FOLDER_MAP: Record<string, FolderTypeValue> = {
  inbox: FolderType.Inbox,
  sentitems: FolderType.Sent,
  drafts: FolderType.Drafts,
  deleteditems: FolderType.Trash,
  junkemail: FolderType.Spam,
  archive: FolderType.Archive,
};

// Maps localized display names to Graph API well-known folder names
const DISPLAY_NAME_TO_WELL_KNOWN: Record<string, string> = {
  // English
  'inbox': 'inbox',
  'sent items': 'sentitems',
  'sent': 'sentitems',
  'drafts': 'drafts',
  'deleted items': 'deleteditems',
  'trash': 'deleteditems',
  'junk email': 'junkemail',
  'junk': 'junkemail',
  'spam': 'junkemail',
  'archive': 'archive',
  // German
  'posteingang': 'inbox',
  'gesendete elemente': 'sentitems',
  'gesendet': 'sentitems',
  'entwürfe': 'drafts',
  'gelöschte elemente': 'deleteditems',
  'papierkorb': 'deleteditems',
  'junk-e-mail': 'junkemail',
  'archiv': 'archive',
  // Spanish
  'bandeja de entrada': 'inbox',
  'elementos enviados': 'sentitems',
  'enviados': 'sentitems',
  'borradores': 'drafts',
  'elementos eliminados': 'deleteditems',
  'correo no deseado': 'junkemail',
  'archivo': 'archive',
  // French
  'boîte de réception': 'inbox',
  'éléments envoyés': 'sentitems',
  'brouillons': 'drafts',
  'éléments supprimés': 'deleteditems',
  'courrier indésirable': 'junkemail',
  'archives': 'archive',
};

/**
 * Resolves a display name or well-known alias to a Graph API well-known folder name.
 * Returns undefined if not a recognized name.
 */
export function resolveWellKnownFolder(nameOrId: string): string | undefined {
  const lower = nameOrId.toLowerCase().trim();
  // Check if it's already a well-known name
  if (WELL_KNOWN_FOLDER_MAP[lower]) return lower;
  // Check display name mappings
  return DISPLAY_NAME_TO_WELL_KNOWN[lower];
}

export function mapGraphFolder(graphFolder: any): Folder {
  const wellKnown = graphFolder.wellKnownName?.toLowerCase();
  return {
    id: graphFolder.id,
    name: graphFolder.displayName,
    path: graphFolder.displayName,
    type: (wellKnown && WELL_KNOWN_FOLDER_MAP[wellKnown]) || FolderType.Other,
    totalCount: graphFolder.totalItemCount,
    unreadCount: graphFolder.unreadItemCount,
  };
}

function mapGraphContact(recipient: any): Contact {
  return {
    name: recipient?.emailAddress?.name || undefined,
    email: recipient?.emailAddress?.address || '',
  };
}

function mapGraphContacts(recipients: any[]): Contact[] {
  if (!recipients) return [];
  return recipients.map(mapGraphContact);
}

export function mapGraphMessage(message: any, accountId: string): Email {
  const bodyType = message.body?.contentType?.toLowerCase();

  return {
    id: message.id,
    accountId,
    threadId: message.conversationId,
    folder: message.parentFolderId || '',
    from: mapGraphContact(message.from),
    to: mapGraphContacts(message.toRecipients),
    cc: message.ccRecipients?.length ? mapGraphContacts(message.ccRecipients) : undefined,
    bcc: message.bccRecipients?.length ? mapGraphContacts(message.bccRecipients) : undefined,
    subject: message.subject || '(no subject)',
    date: message.receivedDateTime || new Date().toISOString(),
    body: {
      text: bodyType === 'text' ? message.body?.content : undefined,
      html: bodyType === 'html' ? message.body?.content : undefined,
    },
    snippet: message.bodyPreview,
    attachments: [],
    categories: message.categories || [],
    flags: {
      read: message.isRead ?? false,
      starred: message.importance === 'high',
      flagged: message.flag?.flagStatus === 'flagged',
      draft: message.isDraft ?? false,
    },
  };
}

export function mapGraphAttachment(attachment: any): AttachmentMeta {
  return {
    id: attachment.id,
    filename: attachment.name || 'unknown',
    contentType: attachment.contentType || 'application/octet-stream',
    size: attachment.size || 0,
  };
}

/** Date-only inputs mean midnight UTC; timestamps must include a timezone. */
export function normalizeGraphSearchQuery(query: SearchQuery): SearchQuery {
  const normalized = { ...query };
  for (const field of ['since', 'before'] as const) {
    const value = query[field];
    if (value === undefined) continue;
    const dateOnly = /^\d{4}-\d{2}-\d{2}$/.test(value);
    const timestamp = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})$/i.test(value);
    const date = new Date(value);
    if ((!dateOnly && !timestamp) || !Number.isFinite(date.getTime()) ||
        (dateOnly && date.toISOString().slice(0, 10) !== value)) {
      throw new Error(`Invalid Outlook search ${field}: use YYYY-MM-DD or an ISO 8601 timestamp with a timezone.`);
    }
    normalized[field] = date.toISOString();
  }
  return normalized;
}

export function buildGraphFilter(query: SearchQuery): { filter?: string; search?: string } {
  const filters: string[] = [];
  const searchParts: string[] = [];

  // Graph requires the orderby property first in the filter, even when the
  // caller only supplied a sender or a flag. Do not impose a recent-date cutoff.
  if (query.since) filters.push(`receivedDateTime ge ${query.since}`);
  if (query.before) filters.push(`receivedDateTime lt ${query.before}`);

  if (query.from) {
    filters.push(`from/emailAddress/address eq '${query.from.replace(/'/g, "''")}'`);
  }
  if (query.to) {
    filters.push(`toRecipients/any(r:r/emailAddress/address eq '${query.to.replace(/'/g, "''")}')`);
  }
  if (query.unreadOnly) {
    filters.push('isRead eq false');
  }
  if (query.starredOnly) {
    filters.push("importance eq 'high'");
  }
  if (query.hasAttachment) {
    filters.push('hasAttachments eq true');
  }
  if (filters.length && !query.since && !query.before) {
    filters.unshift('receivedDateTime ge 0001-01-01T00:00:00Z');
  }

  if (query.subject) {
    searchParts.push(`subject:${query.subject}`);
  }
  if (query.body) {
    searchParts.push(`body:${query.body}`);
  }

  return {
    filter: filters.length > 0 ? filters.join(' and ') : undefined,
    search: searchParts.length > 0 ? JSON.stringify(searchParts.join(' AND ')) : undefined,
  };
}

/** $search cannot be combined with $filter; check structured criteria locally. */
export function matchesGraphSearchFilters(message: any, query: SearchQuery): boolean {
  if (query.from && message.from?.emailAddress?.address?.toLowerCase() !== query.from.toLowerCase()) return false;
  if (query.to && !message.toRecipients?.some((recipient: any) =>
    recipient.emailAddress?.address?.toLowerCase() === query.to!.toLowerCase())) return false;
  if (query.unreadOnly && message.isRead !== false) return false;
  if (query.starredOnly && message.importance !== 'high') return false;
  if (query.hasAttachment && message.hasAttachments !== true) return false;
  const received = Date.parse(message.receivedDateTime);
  if (query.since && !(received >= Date.parse(query.since))) return false;
  if (query.before && !(received < Date.parse(query.before))) return false;
  return true;
}
