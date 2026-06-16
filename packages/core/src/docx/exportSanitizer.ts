import type {
  BlockContent,
  Document,
  HeaderFooter,
  Paragraph,
  ParagraphContent,
} from '../types/document';
import { generateHexId, MAX_HEX_ID_EXCLUSIVE } from '../utils/hexId';

const LONG_HEX_ID_RE = /^[0-9A-Fa-f]{8}$/;

function isValidLongHexId(id: string | undefined): boolean {
  if (!id || !LONG_HEX_ID_RE.test(id)) return false;
  return parseInt(id, 16) < MAX_HEX_ID_EXCLUSIVE;
}

function newLongHexId(): string {
  let id = generateHexId();
  while (!isValidLongHexId(id)) {
    id = generateHexId();
  }
  return id;
}

function sanitizeLongHexId(id: string | undefined): string | undefined {
  if (id == null) return undefined;
  return isValidLongHexId(id) ? id : newLongHexId();
}

function hasKnownCommentId(item: ParagraphContent, commentIds: Set<number>): boolean {
  if (item.type !== 'commentRangeStart' && item.type !== 'commentRangeEnd') {
    return true;
  }

  return commentIds.has(item.id);
}

function sanitizeParagraph(paragraph: Paragraph, commentIds: Set<number>): void {
  paragraph.paraId = sanitizeLongHexId(paragraph.paraId);
  paragraph.textId = sanitizeLongHexId(paragraph.textId);
  paragraph.content = paragraph.content.filter((item) => hasKnownCommentId(item, commentIds));
}

function sanitizeBlocks(blocks: BlockContent[] | undefined, commentIds: Set<number>): void {
  if (!blocks) return;

  for (const block of blocks) {
    if (block.type === 'paragraph') {
      sanitizeParagraph(block, commentIds);
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          sanitizeBlocks(cell.content, commentIds);
        }
      }
    } else if (block.type === 'blockSdt') {
      sanitizeBlocks(block.content, commentIds);
    }
  }
}

function sanitizeHeaderFooterParts(
  parts: Map<string, HeaderFooter> | undefined,
  commentIds: Set<number>
): void {
  if (!parts) return;

  for (const part of parts.values()) {
    sanitizeBlocks(part.content, commentIds);
  }
}

/**
 * Normalize model fields that Word validates globally, even though they are
 * not user-visible document content.
 */
export function sanitizeDocumentForExport(doc: Document): void {
  const commentIds = new Set((doc.package.document.comments ?? []).map((comment) => comment.id));

  sanitizeBlocks(doc.package.document.content, commentIds);
  sanitizeHeaderFooterParts(doc.package.headers, commentIds);
  sanitizeHeaderFooterParts(doc.package.footers, commentIds);

  for (const footnote of doc.package.footnotes ?? []) {
    sanitizeBlocks(footnote.content, commentIds);
  }
  for (const footnote of doc.package.footnoteSeparators ?? []) {
    sanitizeBlocks(footnote.content, commentIds);
  }
  for (const endnote of doc.package.endnotes ?? []) {
    sanitizeBlocks(endnote.content, commentIds);
  }
  for (const endnote of doc.package.endnoteSeparators ?? []) {
    sanitizeBlocks(endnote.content, commentIds);
  }
}
