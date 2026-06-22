import { describe, expect, test } from 'bun:test';
import { measureTextBoxBlock, type MeasureParagraphFn } from '../measureTextBox';
import type { ParagraphBlock, ParagraphMeasure, TextBoxBlock } from '../../../layout-engine/types';

const paragraph: ParagraphBlock = {
  kind: 'paragraph',
  id: 'p1',
  runs: [{ kind: 'text', text: 'INTRODUCTION' }],
};

const measureParagraph: MeasureParagraphFn = (): ParagraphMeasure => ({
  kind: 'paragraph',
  lines: [
    {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 12,
      width: 70,
      ascent: 8,
      descent: 3,
      lineHeight: 12,
    },
  ],
  totalHeight: 12,
});

function textBox(autoFit?: TextBoxBlock['autoFit']): TextBoxBlock {
  return {
    kind: 'textBox',
    id: 'tb',
    width: 192,
    height: 32,
    autoFit,
    margins: { top: 5, right: 10, bottom: 5, left: 10 },
    content: [paragraph],
  };
}

describe('measureTextBoxBlock', () => {
  test('preserves authored width when shape auto-fit is absent', () => {
    const measure = measureTextBoxBlock(textBox(), measureParagraph);
    expect(measure.width).toBe(192);
  });

  test('fits DrawingML shape-autofit text boxes to measured content width', () => {
    const measure = measureTextBoxBlock(textBox('shape'), measureParagraph);
    expect(measure.width).toBe(90);
    expect(measure.height).toBe(32);
  });
});
