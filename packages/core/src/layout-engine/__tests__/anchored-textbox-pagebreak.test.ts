import { describe, expect, test } from 'bun:test';
import { layoutDocument } from '../index';
import type {
  FlowBlock,
  Measure,
  ParagraphBlock,
  ParagraphMeasure,
  SectionBreakBlock,
  TextBoxBlock,
} from '../types';

const pageSize = { w: 800, h: 600 };
const margins = { top: 50, right: 50, bottom: 50, left: 50 };

describe('layoutDocument anchored text boxes', () => {
  test('moves following-block anchored text box with a pageBreakBefore paragraph', () => {
    const intro = paragraph(1, 'intro');
    const anchoredTitle: TextBoxBlock = {
      kind: 'textBox',
      id: 2,
      width: 300,
      height: 80,
      content: [],
      displayMode: 'float',
      wrapType: 'inFront',
      anchorTarget: 'followingBlock',
      zIndex: 42,
    };
    const heading = paragraph(3, 'heading', { pageBreakBefore: true });

    const blocks: FlowBlock[] = [intro, anchoredTitle, heading];
    const measures: Measure[] = [
      paragraphMeasure(24),
      { kind: 'textBox', width: 300, height: 80, innerMeasures: [] },
      paragraphMeasure(30),
    ];

    const layout = layoutDocument(blocks, measures, { pageSize, margins });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages[0].fragments.map((fragment) => fragment.blockId)).toEqual([1]);
    expect(layout.pages[1].fragments.map((fragment) => fragment.blockId)).toEqual([2, 3]);

    const titleFragment = layout.pages[1].fragments[0];
    expect(titleFragment.kind).toBe('textBox');
    if (titleFragment.kind === 'textBox') {
      expect(titleFragment.zIndex).toBe(42);
    }
  });

  test('promotes a continuous section break before a page-centered anchored title', () => {
    const note = paragraph(1, 'note');
    const sectionBreak: SectionBreakBlock = {
      kind: 'sectionBreak',
      id: 2,
      type: 'continuous',
      pageSize,
      margins,
    };
    const spacer = paragraph(3, '');
    const decorativeBox: TextBoxBlock = {
      kind: 'textBox',
      id: 4,
      width: 700,
      height: 40,
      content: [],
      displayMode: 'float',
      wrapType: 'inFront',
      anchorTarget: 'followingBlock',
      position: {
        horizontal: { relativeTo: 'margin', align: 'center' },
        vertical: { relativeTo: 'paragraph', posOffset: 120 },
      },
    };
    const anchoredTitle: TextBoxBlock = {
      kind: 'textBox',
      id: 5,
      width: 700,
      height: 120,
      content: [paragraph(50, 'Section title')],
      displayMode: 'float',
      wrapType: 'inFront',
      anchorTarget: 'followingBlock',
      position: {
        horizontal: { relativeTo: 'page', align: 'left' },
        vertical: { relativeTo: 'margin', align: 'center' },
      },
    };
    const anchorParagraph = paragraph(6, '');

    const blocks: FlowBlock[] = [
      note,
      sectionBreak,
      spacer,
      decorativeBox,
      anchoredTitle,
      anchorParagraph,
    ];
    const measures: Measure[] = [
      paragraphMeasure(24),
      { kind: 'sectionBreak' },
      emptyParagraphMeasure(),
      { kind: 'textBox', width: 700, height: 40, innerMeasures: [] },
      { kind: 'textBox', width: 700, height: 120, innerMeasures: [] },
      emptyParagraphMeasure(),
    ];

    const layout = layoutDocument(blocks, measures, { pageSize, margins });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages[0].fragments.map((fragment) => fragment.blockId)).toEqual([1]);
    expect(layout.pages[1].fragments.map((fragment) => fragment.blockId)).toEqual([3, 4, 5, 6]);
  });
});

function paragraph(id: number, text: string, attrs?: ParagraphBlock['attrs']): ParagraphBlock {
  return {
    kind: 'paragraph',
    id,
    runs: text ? [{ kind: 'text', text }] : [],
    attrs,
  };
}

function paragraphMeasure(height: number): ParagraphMeasure {
  return {
    kind: 'paragraph',
    totalHeight: height,
    lines: [
      {
        fromRun: 0,
        fromChar: 0,
        toRun: 0,
        toChar: 1,
        width: 100,
        ascent: Math.round(height * 0.75),
        descent: Math.round(height * 0.25),
        lineHeight: height,
      },
    ],
  };
}

function emptyParagraphMeasure(): ParagraphMeasure {
  return {
    kind: 'paragraph',
    totalHeight: 16,
    lines: [
      {
        fromRun: 0,
        fromChar: 0,
        toRun: 0,
        toChar: 0,
        width: 0,
        ascent: 12,
        descent: 4,
        lineHeight: 16,
      },
    ],
  };
}
