import { describe, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { executeInsertTable } from '../agent/executor/structureCommands';
import type { Table } from '../types/document';
import { MAX_HEX_ID_EXCLUSIVE } from '../utils/hexId';
import { createEmptyDocument } from '../utils/createDocument';
import { createDocx } from './rezip';

function textRun(text: string) {
  return { type: 'run' as const, content: [{ type: 'text' as const, text }] };
}

function bareTable(): Table {
  return {
    type: 'table',
    rows: [
      {
        type: 'tableRow',
        cells: [
          { type: 'tableCell', content: [{ type: 'paragraph', content: [textRun('A')] }] },
          { type: 'tableCell', content: [{ type: 'paragraph', content: [textRun('B')] }] },
        ],
      },
    ],
  };
}

function longHexIds(xml: string): string[] {
  return [...xml.matchAll(/w14:(?:paraId|textId)="([0-9A-Fa-f]{8})"/g)].map((match) => match[1]);
}

describe('DOCX export sanitization', () => {
  test('emits Word-valid relationships, table grids, comment anchors, and long-hex ids', async () => {
    const doc = createEmptyDocument();
    doc.package.document.content = [
      {
        type: 'paragraph',
        paraId: 'F2345678',
        textId: '9ABCDEF0',
        content: [
          { type: 'commentRangeStart', id: 100 },
          textRun('Visit '),
          {
            type: 'hyperlink',
            rId: 'rIdMissing',
            href: 'https://example.com?a=1&b=2',
            children: [textRun('example')],
          },
          textRun(' and '),
          {
            type: 'hyperlink',
            rId: 'rIdMissingNoHref',
            children: [textRun('stale')],
          },
          { type: 'commentRangeEnd', id: 100 },
        ],
      },
      bareTable(),
    ];

    const buffer = await createDocx(doc);
    const zip = await JSZip.loadAsync(buffer);
    const documentXml = await zip.file('word/document.xml')!.async('text');
    const documentRelsXml = await zip.file('word/_rels/document.xml.rels')!.async('text');

    expect(documentXml).not.toContain('F2345678');
    expect(documentXml).not.toContain('9ABCDEF0');
    for (const id of longHexIds(documentXml)) {
      expect(parseInt(id, 16)).toBeLessThan(MAX_HEX_ID_EXCLUSIVE);
    }

    expect(documentXml).not.toContain('w:commentRangeStart');
    expect(documentXml).not.toContain('w:commentRangeEnd');
    expect(documentXml).not.toContain('w:commentReference');

    expect(documentXml).toContain(
      '<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>'
    );

    const hyperlinkRId = documentXml.match(/<w:hyperlink [^>]*r:id="([^"]+)"/)?.[1];
    expect(hyperlinkRId).toBeDefined();
    expect(hyperlinkRId).not.toBe('rIdMissing');
    expect(documentRelsXml).toContain(`Id="${hyperlinkRId}"`);
    expect(documentRelsXml).toContain('Target="https://example.com?a=1&amp;b=2"');
    expect(documentRelsXml).toContain('TargetMode="External"');
    expect(documentXml).not.toContain('rIdMissingNoHref');
  });

  test('agent-inserted tables carry explicit export widths', () => {
    const doc = createEmptyDocument();
    const edited = executeInsertTable(doc, {
      type: 'insertTable',
      position: { paragraphIndex: 0, offset: 0 },
      rows: 2,
      columns: 3,
      data: [
        ['A', 'B', 'C'],
        ['D', 'E', 'F'],
      ],
    });

    const table = edited.package.document.content.find((block) => block.type === 'table') as Table;
    expect(table.formatting?.width).toEqual({ value: 9360, type: 'dxa' });
    expect(table.formatting?.layout).toBe('fixed');
    expect(table.columnWidths).toEqual([3120, 3120, 3120]);
    expect(table.rows[0].cells.map((cell) => cell.formatting?.width?.value)).toEqual([
      3120, 3120, 3120,
    ]);
  });
});
