import { describe, test, expect } from 'bun:test';

import { layoutDocument } from '../index';
import type { FlowBlock, Measure, TableBlock, TableFragment, TableMeasure } from '../types';

import { makeParagraphBlock, makeLine, makeParagraphMeasure, makeLayoutOptions } from './helpers';

describe('Section Breaks', () => {
  test('nextPage section break forces new page', () => {
    const blocks: FlowBlock[] = [
      makeParagraphBlock(0, 'Before section', 1),
      { kind: 'sectionBreak', id: 1, type: 'nextPage' },
      makeParagraphBlock(2, 'After section', 18),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 14, 120, 24)]),
      { kind: 'sectionBreak' },
      makeParagraphMeasure([makeLine(0, 0, 0, 13, 110, 24)]),
    ];

    const layout = layoutDocument(blocks, measures, makeLayoutOptions());

    expect(layout.pages.length).toBe(2);
    expect(layout.pages[0].fragments.some((f) => f.blockId === 0)).toBe(true);
    expect(layout.pages[1].fragments.some((f) => f.blockId === 2)).toBe(true);
  });

  test('continuous section break does not force new page', () => {
    const blocks: FlowBlock[] = [
      makeParagraphBlock(0, 'Before section', 1),
      { kind: 'sectionBreak', id: 1, type: 'continuous' },
      makeParagraphBlock(2, 'After section', 18),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 14, 120, 24)]),
      { kind: 'sectionBreak' },
      makeParagraphMeasure([makeLine(0, 0, 0, 13, 110, 24)]),
    ];

    const layout = layoutDocument(blocks, measures, makeLayoutOptions());

    expect(layout.pages.length).toBe(1);
    expect(layout.pages[0].fragments.length).toBe(2);
  });

  test('paragraph fragments in multi-column sections use column width', () => {
    const blocks: FlowBlock[] = [makeParagraphBlock(0, 'Column body text', 1)];
    const measures: Measure[] = [makeParagraphMeasure([makeLine(0, 0, 0, 16, 120, 24)])];

    const layout = layoutDocument(
      blocks,
      measures,
      makeLayoutOptions({ columns: { count: 2, gap: 48 } })
    );

    const fragment = layout.pages[0].fragments[0];

    expect(fragment.kind).toBe('paragraph');
    expect(fragment.x).toBe(96);
    expect(fragment.width).toBe(288);
  });

  test('table fragments keep the advanced column x position', () => {
    const fillFirstColumn = makeParagraphBlock(0, 'First column filler', 1);
    const table = {
      kind: 'table',
      id: 1,
      columnWidths: [200],
      rows: [{ id: 2, cells: [{ id: 3, blocks: [] }] }],
    } as unknown as TableBlock;
    const tableMeasure: TableMeasure = {
      kind: 'table',
      columnWidths: [200],
      totalWidth: 200,
      totalHeight: 36,
      rows: [{ height: 36, cells: [{ blocks: [], width: 200, height: 36 }] }],
    };
    const blocks: FlowBlock[] = [fillFirstColumn, table];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 19, 10, 850)]),
      tableMeasure,
    ];

    const layout = layoutDocument(
      blocks,
      measures,
      makeLayoutOptions({ columns: { count: 2, gap: 48 } })
    );

    const tableFragment = layout.pages[0].fragments.find(
      (fragment): fragment is TableFragment => fragment.kind === 'table'
    );

    expect(tableFragment).toBeDefined();
    expect(tableFragment!.x).toBe(432);
    expect(tableFragment!.width).toBe(200);
  });
});
