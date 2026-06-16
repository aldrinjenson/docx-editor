/**
 * New Hyperlink Registration
 *
 * On save, scan all parts for external hyperlinks and ensure each emitted rId
 * resolves in the owning part's rels file.
 */

import type JSZip from 'jszip';
import type { BlockContent, Hyperlink } from '../../types/content';
import { RELATIONSHIP_TYPES } from '../relsParser';
import { escapeXml } from '../serializer/xmlUtils';
import { findMaxRId, readRelsOrStub, type Part } from './parts';

/**
 * Collect all external hyperlinks from block content.
 */
function collectExternalHyperlinks(blocks: BlockContent[]): Hyperlink[] {
  const hyperlinks: Hyperlink[] = [];

  for (const block of blocks) {
    if (block.type === 'paragraph') {
      for (const item of block.content) {
        if (item.type === 'hyperlink' && (item.href || item.rId) && !item.anchor) {
          hyperlinks.push(item);
        }
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          hyperlinks.push(...collectExternalHyperlinks(cell.content));
        }
      }
    } else if (block.type === 'blockSdt') {
      hyperlinks.push(...collectExternalHyperlinks(block.content));
    }
  }

  return hyperlinks;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function relationshipElementForId(relsXml: string, rId: string): string | undefined {
  const match = relsXml.match(
    new RegExp(`<Relationship\\b(?=[^>]*\\bId="${escapeRegExp(rId)}")[^>]*/?>`)
  );
  return match?.[0];
}

function relationshipTargetsHref(relationshipXml: string | undefined, href: string): boolean {
  if (!relationshipXml) return false;

  return (
    relationshipXml.includes(`Type="${RELATIONSHIP_TYPES.hyperlink}"`) &&
    relationshipXml.includes(`Target="${escapeXml(href)}"`) &&
    relationshipXml.includes('TargetMode="External"')
  );
}

function findExistingHyperlinkRelationshipId(relsXml: string, href: string): string | undefined {
  const escapedHref = escapeXml(href);

  for (const match of relsXml.matchAll(/<Relationship\b[^>]*\/?>/g)) {
    const relationshipXml = match[0];
    if (
      !relationshipXml.includes(`Type="${RELATIONSHIP_TYPES.hyperlink}"`) ||
      !relationshipXml.includes(`Target="${escapedHref}"`) ||
      !relationshipXml.includes('TargetMode="External"')
    ) {
      continue;
    }

    return relationshipXml.match(/\bId="([^"]+)"/)?.[1];
  }

  return undefined;
}

/**
 * Process newly created hyperlinks across all parts (body, headers, footers):
 * assign rIds and add relationship entries to the owning part's rels file.
 *
 * Mutates each hyperlink's rId in-place.
 */
export async function processNewHyperlinks(
  parts: Part[],
  zip: JSZip,
  compressionLevel: number
): Promise<void> {
  for (const { relsPath, blocks } of parts) {
    const hyperlinks = collectExternalHyperlinks(blocks);
    if (hyperlinks.length === 0) continue;

    const relsXml = await readRelsOrStub(zip, relsPath);
    let maxId = findMaxRId(relsXml);
    const relEntries: string[] = [];

    for (const hyperlink of hyperlinks) {
      const href = hyperlink.href;
      const currentRelationship = hyperlink.rId
        ? relationshipElementForId(relsXml, hyperlink.rId)
        : undefined;
      if (!href) {
        if (!currentRelationship) {
          hyperlink.rId = undefined;
        }
        continue;
      }

      if (relationshipTargetsHref(currentRelationship, href)) {
        continue;
      }

      const existingRId = findExistingHyperlinkRelationshipId(relsXml, href);
      if (existingRId) {
        hyperlink.rId = existingRId;
        continue;
      }

      maxId++;
      const newRId = `rId${maxId}`;

      relEntries.push(
        `<Relationship Id="${newRId}" Type="${RELATIONSHIP_TYPES.hyperlink}" Target="${escapeXml(href)}" TargetMode="External"/>`
      );

      hyperlink.rId = newRId;
    }

    const updatedRelsXml = relsXml.replace(
      '</Relationships>',
      relEntries.join('') + '</Relationships>'
    );
    zip.file(relsPath, updatedRelsXml, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });
  }
}
