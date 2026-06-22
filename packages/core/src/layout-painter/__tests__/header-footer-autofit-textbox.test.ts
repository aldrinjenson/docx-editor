import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { measureTextBoxBlock, type MeasureParagraphFn } from '../../layout-bridge/measuring';
import type { ParagraphBlock, ParagraphMeasure, TextBoxBlock } from '../../layout-engine/types';
import type { RenderContext } from '../renderPage';
import {
  renderHeaderFooterContent,
  type HeaderFooterContent,
  type HeaderFooterLayoutInfo,
} from '../renderPage/headerFooter';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const context: RenderContext = {
  pageNumber: 1,
  totalPages: 1,
  section: 'header',
  contentWidth: 600,
};
const layout: HeaderFooterLayoutInfo = {
  flowTop: 48,
  flowLeft: 72,
  contentWidth: 600,
  pageWidth: 744,
  pageHeight: 1056,
  margins: { top: 96, right: 72, bottom: 96, left: 72 },
};

function paragraph(id: string, text: string): ParagraphBlock {
  return { kind: 'paragraph', id, runs: [{ kind: 'text', text }] };
}

const measureParagraph: MeasureParagraphFn = (block): ParagraphMeasure => ({
  kind: 'paragraph',
  lines: [
    {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: block.runs[0]?.kind === 'text' ? block.runs[0].text.length : 0,
      width: block.id === 'active-p' ? 70 : 78,
      ascent: 8,
      descent: 3,
      lineHeight: 12,
    },
  ],
  totalHeight: 12,
});

function labelBlock(
  id: string,
  text: string,
  width: number,
  left: number,
  autoFit?: TextBoxBlock['autoFit']
): TextBoxBlock {
  return {
    kind: 'textBox',
    id,
    width,
    height: 28,
    autoFit,
    displayMode: 'float',
    wrapType: 'inFront',
    position: { horizontal: { posOffset: left * 9525 }, vertical: { posOffset: 0 } },
    margins: { top: 4, right: 10, bottom: 4, left: 10 },
    content: [paragraph(`${id}-p`, text)],
  };
}

describe('header/footer shape-autofit text boxes', () => {
  test('autofit label hit boxes do not extend into neighboring header tabs', () => {
    const active = labelBlock('active', 'INTRODUCTION', 192, -60, 'shape');
    const next = labelBlock('next', 'CORPORATE', 100, 32);
    const measures = [
      measureTextBoxBlock(active, measureParagraph),
      measureTextBoxBlock(next, measureParagraph),
    ];
    const content: HeaderFooterContent = {
      blocks: [active, next],
      measures,
      height: 28,
      flowHeight: 0,
      visualTop: 0,
      visualBottom: 28,
    };

    const el = renderHeaderFooterContent(content, context, { document }, layout);
    const boxes = Array.from(el.querySelectorAll<HTMLElement>('.layout-textbox'));
    const activeEl = boxes.find((box) => box.dataset.blockId === 'active')!;
    const nextEl = boxes.find((box) => box.dataset.blockId === 'next')!;

    const activeRight = parseFloat(activeEl.style.left) + parseFloat(activeEl.style.width);
    const nextLeft = parseFloat(nextEl.style.left);
    expect(activeEl.style.width).toBe('90px');
    expect(activeRight).toBeLessThan(nextLeft);
  });
});
