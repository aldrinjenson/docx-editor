import { describe, expect, test } from 'bun:test';
import type { Document, Paragraph } from '../../../types/document';
import { serializeParagraph } from '../../../docx/serializer/paragraphSerializer';
import { fromProseDoc } from '../fromProseDoc';
import { toProseDoc } from '../toProseDoc';

function docOf(...content: Paragraph[]): Document {
  return { package: { document: { content } } };
}

function textParagraph(text: string): Paragraph {
  return {
    type: 'paragraph',
    content: [{ type: 'run', content: [{ type: 'text', text }] }],
  };
}

function childTypes(doc: ReturnType<typeof toProseDoc>): string[] {
  const types: string[] = [];
  doc.forEach((node) => {
    types.push(node.type.name);
  });
  return types;
}

describe('issue #830 leading hard page break round-trip', () => {
  test('does not export a leading hard page break as an extra empty paragraph', () => {
    const leadingBreakParagraph: Paragraph = {
      type: 'paragraph',
      renderedPageBreakBefore: true,
      formatting: { styleId: 'CenterSingle' },
      content: [
        { type: 'run', content: [{ type: 'break', breakType: 'page' }] },
        { type: 'run', content: [{ type: 'text', text: 'After hard break' }] },
      ],
    };
    const input = docOf(textParagraph('Before'), leadingBreakParagraph);

    const pmDoc = toProseDoc(input);
    expect(childTypes(pmDoc)).toEqual(['paragraph', 'paragraph']);
    expect(pmDoc.child(1).attrs.pageBreakBefore).toBe(true);
    expect(pmDoc.child(1).attrs.renderedPageBreakBefore).toBe(true);

    const roundTripped = fromProseDoc(pmDoc, input);
    expect(roundTripped.package.document.content).toHaveLength(2);

    const outputParagraph = roundTripped.package.document.content[1];
    expect(outputParagraph?.type).toBe('paragraph');
    if (outputParagraph?.type !== 'paragraph') {
      throw new Error('Expected second body block to remain a paragraph');
    }

    const xml = serializeParagraph(outputParagraph);
    expect(xml).toContain('<w:pageBreakBefore/>');
    expect(xml).toMatch(/<w:r[^>]*><w:lastRenderedPageBreak\/><w:t>After hard break/);
  });

  test('keeps non-leading hard page breaks as explicit PM page break blocks', () => {
    const midParagraph: Paragraph = {
      type: 'paragraph',
      content: [
        { type: 'run', content: [{ type: 'text', text: 'Before break' }] },
        { type: 'run', content: [{ type: 'break', breakType: 'page' }] },
        { type: 'run', content: [{ type: 'text', text: 'After break' }] },
      ],
    };

    const pmDoc = toProseDoc(docOf(midParagraph));
    expect(childTypes(pmDoc)).toEqual(['paragraph', 'pageBreak']);
  });
});
