import {
  DEFAULT_TEXTBOX_MARGINS,
  DEFAULT_TEXTBOX_WIDTH,
  type ParagraphBlock,
  type ParagraphMeasure,
  type TextBoxBlock,
  type TextBoxMeasure,
} from '../../layout-engine/types';
import { measureParagraph } from './measureParagraph';

export type MeasureParagraphFn = (block: ParagraphBlock, maxWidth: number) => ParagraphMeasure;

function maxMeasuredLineWidth(measures: ParagraphMeasure[]): number {
  return measures.reduce((max, measure) => {
    const lineMax = measure.lines.reduce((lineAcc, line) => Math.max(lineAcc, line.width), 0);
    return Math.max(max, lineMax);
  }, 0);
}

function measureInnerParagraphs(
  textBox: TextBoxBlock,
  width: number,
  measureParagraphFn: MeasureParagraphFn
): ParagraphMeasure[] {
  const margins = textBox.margins ?? DEFAULT_TEXTBOX_MARGINS;
  const innerWidth = Math.max(1, width - margins.left - margins.right);
  return textBox.content.map((paragraph) => measureParagraphFn(paragraph, innerWidth));
}

/**
 * Measure a DOCX text box. DrawingML `spAutoFit` means the shape itself fits
 * its text; for transparent positioned labels this prevents the authored hit
 * box from centering text across neighboring tab artwork.
 */
export function measureTextBoxBlock(
  textBox: TextBoxBlock,
  measureParagraphFn: MeasureParagraphFn = measureParagraph
): TextBoxMeasure {
  const margins = textBox.margins ?? DEFAULT_TEXTBOX_MARGINS;
  let width = textBox.width ?? DEFAULT_TEXTBOX_WIDTH;
  let innerMeasures = measureInnerParagraphs(textBox, width, measureParagraphFn);

  if (textBox.autoFit === 'shape') {
    const preferredInnerWidth = Math.ceil(maxMeasuredLineWidth(innerMeasures));
    if (preferredInnerWidth > 0) {
      width = Math.max(1, preferredInnerWidth + margins.left + margins.right);
      innerMeasures = measureInnerParagraphs(textBox, width, measureParagraphFn);
    }
  }

  const contentHeight = innerMeasures.reduce((sum, measure) => sum + measure.totalHeight, 0);
  return {
    kind: 'textBox',
    width,
    height: textBox.height ?? contentHeight + margins.top + margins.bottom,
    innerMeasures,
  };
}
