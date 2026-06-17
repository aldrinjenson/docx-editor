import { describe, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { DocumentAgent } from '../agent/DocumentAgent';
import { createDocx } from './rezip';
import { parseDocx } from './parser';
import { serializeStyles } from './serializer/styleSerializer';
import { parseStyleDefinitions } from './styleParser';
import { updateDocDefaults, updateStyleDefinition, addStyleDefinition } from './styleDefinitions';
import { createEmptyDocument } from '../utils/createDocument';
import { toProseDoc } from '../prosemirror/conversion';
import { createStyleResolver } from '../prosemirror/styles';
import type { Document, Paragraph, StyleDefinitions } from '../types/document';

function firstTextMarks(doc: Document) {
  const pmDoc = toProseDoc(doc, { styles: doc.package.styles });
  return pmDoc.firstChild?.firstChild?.marks ?? [];
}

describe('style definition API', () => {
  test('headless updateStyle resolves newly-created heading content', () => {
    const doc = createEmptyDocument({ initialText: 'Styled heading' });
    const paragraph = doc.package.document.content[0] as Paragraph;
    paragraph.formatting = { styleId: 'Heading1' };
    const run = paragraph.content[0];
    if (run.type === 'run') run.formatting = undefined;

    const agent = DocumentAgent.fromDocument(doc).updateStyle('Heading1', {
      rPr: { color: { rgb: '123456' } },
    });
    const updatedDoc = agent.getDocument();

    const textColor = firstTextMarks(updatedDoc).find((mark) => mark.type.name === 'textColor');
    expect(textColor?.attrs.rgb).toBe('123456');
  });

  test('basedOn inheritance is resolved dynamically after parent updates', () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: 'BaseHeading',
          type: 'paragraph',
          rPr: { color: { rgb: 'AA0000' }, fontSize: 30 },
        },
        {
          styleId: 'ChildHeading',
          type: 'paragraph',
          basedOn: 'BaseHeading',
          rPr: { bold: true },
        },
      ],
    };

    const updated = updateStyleDefinition(styles, 'BaseHeading', {
      rPr: { color: { rgb: '0055AA' } },
    });
    const resolved = createStyleResolver(updated).resolveParagraphStyle('ChildHeading');

    expect(resolved.runFormatting?.color?.rgb).toBe('0055AA');
    expect(resolved.runFormatting?.fontSize).toBe(30);
    expect(resolved.runFormatting?.bold).toBe(true);
  });

  test('table style overrides and basedOn table inheritance are merged', () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: 'BaseTable',
          type: 'table',
          tblPr: {
            cellMargins: {
              left: { value: 120, type: 'dxa' },
              right: { value: 120, type: 'dxa' },
            },
          },
          tblStylePr: [{ type: 'firstRow', rPr: { bold: true } }],
        },
        {
          styleId: 'ChildTable',
          type: 'table',
          basedOn: 'BaseTable',
          tblPr: {
            borders: { top: { style: 'single', size: 8, color: { rgb: '000000' } } },
          },
        },
      ],
    };

    const updated = updateStyleDefinition(styles, 'ChildTable', {
      tblPr: {
        cellMargins: { left: { value: 240, type: 'dxa' } },
      },
      tblStylePr: [{ type: 'firstRow', rPr: { color: { rgb: 'FFFFFF' } } }],
    });
    const child = createStyleResolver(updated).getStyle('ChildTable');

    expect(child?.tblPr?.cellMargins?.left?.value).toBe(240);
    expect(child?.tblPr?.cellMargins?.right?.value).toBe(120);
    expect(child?.tblPr?.borders?.top?.size).toBe(8);
    const firstRow = child?.tblStylePr?.find((part) => part.type === 'firstRow');
    expect(firstRow?.rPr?.bold).toBe(true);
    expect(firstRow?.rPr?.color?.rgb).toBe('FFFFFF');
  });

  test('docDefaults and theme color references serialize to styles.xml', () => {
    const originalStylesXml =
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/></w:style></w:styles>';
    const parsed = parseStyleDefinitions(originalStylesXml, null);
    const withDefaults = updateDocDefaults(parsed, {
      rPr: { color: { themeColor: 'accent1', themeTint: '80' } },
    });
    const updated = updateStyleDefinition(withDefaults, 'Heading1', {
      rPr: { color: { themeColor: 'accent2', themeShade: '40' } },
    });
    const serialized = serializeStyles(originalStylesXml, updated);

    expect(serialized).toContain('w:themeColor="accent1"');
    expect(serialized).toContain('w:themeTint="80"');
    expect(serialized).toContain('w:themeColor="accent2"');
    expect(serialized).toContain('w:themeShade="40"');
  });

  test('adds a new style definition', () => {
    const updated = addStyleDefinition(undefined, {
      styleId: 'BrandHeading',
      type: 'paragraph',
      name: 'Brand Heading',
      basedOn: 'Heading1',
      rPr: { color: { rgb: '7722AA' } },
    });

    expect(updated.styles.some((style) => style.styleId === 'BrandHeading')).toBe(true);
    expect(updated.modifiedStyleIds).toContain('BrandHeading');
  });

  test('exported docx round-trips style definition changes through styles.xml', async () => {
    const doc = createEmptyDocument({ initialText: 'Round trip' });
    doc.package.styles = updateStyleDefinition(doc.package.styles, 'Heading1', {
      rPr: { color: { rgb: 'BADA55' } },
    });

    const buffer = await createDocx(doc);
    const zip = await JSZip.loadAsync(buffer);
    const stylesXml = await zip.file('word/styles.xml')?.async('text');

    expect(stylesXml).toContain('w:styleId="Heading1"');
    expect(stylesXml).toContain('w:val="BADA55"');

    const reparsed = await parseDocx(buffer);
    const heading1 = reparsed.package.styles?.styles.find((style) => style.styleId === 'Heading1');
    expect(heading1?.rPr?.color?.rgb).toBe('BADA55');
  });
});
