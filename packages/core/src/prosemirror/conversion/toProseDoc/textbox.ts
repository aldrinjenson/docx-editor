/**
 * Document TextBox/anchored shape → PM textBox node + sibling-paragraph
 * extraction (Document → ProseMirror direction).
 *
 * Word stores text boxes as `<w:r><mc:AlternateContent><...><w:txbxContent>`
 * embedded in a host paragraph. This module pulls them out into sibling PM
 * `textBox` nodes (anchored before the host paragraph, in-flow ones after)
 * and converts the shape body into a paragraph-bearing PM node. The host
 * paragraph is dropped if extracting the text boxes leaves it empty.
 */

import type { Node as PMNode } from 'prosemirror-model';
import { schema } from '../../schema';
import type { Paragraph, Run, ShapeContent, TextBox, Shape, Theme } from '../../../types/document';
import { emuToPixels } from '../../../docx/imageParser';
import { resolveColorToHex } from '../../../utils/colorResolver';
import type { StyleResolver } from '../../styles';
import { isAnchoredDocxTextBox, textBoxAnchorAttrsFromDocx } from '../textBoxAnchors';
import { convertParagraph } from './paragraph';

/**
 * Convert a paragraph block to PM nodes, extracting text boxes as sibling nodes.
 * Skips ghost empty paragraphs that only contained text box drawings.
 */
export function convertParagraphWithTextBoxes(
  block: Paragraph,
  styleResolver: StyleResolver | null,
  theme: Theme | null = null
): PMNode[] {
  const { textBoxes, extractedShapes } = extractTextBoxesFromParagraph(block);
  const paragraphWithoutExtractedShapes =
    extractedShapes.size > 0 ? omitExtractedShapes(block, extractedShapes) : block;
  const pmParagraph = convertParagraph(paragraphWithoutExtractedShapes, styleResolver);
  const nodes: PMNode[] = [];
  const isEmptyAfterExtraction = textBoxes.length > 0 && pmParagraph.content.size === 0;
  const { anchored, inFlow } = partitionTextBoxesByAnchor(textBoxes);

  for (const tb of anchored) {
    nodes.push(convertTextBox(tb, styleResolver, theme));
  }

  if (!isEmptyAfterExtraction) {
    nodes.push(pmParagraph);
  }

  for (const tb of inFlow) {
    nodes.push(convertTextBox(tb, styleResolver, theme));
  }
  return nodes;
}

function partitionTextBoxesByAnchor(textBoxes: TextBox[]): {
  anchored: TextBox[];
  inFlow: TextBox[];
} {
  const anchored: TextBox[] = [];
  const inFlow: TextBox[] = [];

  for (const textBox of textBoxes) {
    if (isAnchoredDocxTextBox(textBox)) {
      anchored.push(textBox);
    } else {
      inFlow.push(textBox);
    }
  }

  return { anchored, inFlow };
}

/**
 * Extract text boxes from paragraph runs.
 * Text boxes appear as ShapeContent where the shape has textBody,
 * or as anchored visual-only rectangles that can use the textBox block painter
 * for fill/outline/position fidelity.
 */
function extractTextBoxesFromParagraph(paragraph: Paragraph): {
  textBoxes: TextBox[];
  extractedShapes: Set<ShapeContent>;
} {
  const textBoxes: TextBox[] = [];
  const extractedShapes = new Set<ShapeContent>();
  for (const content of paragraph.content) {
    if (content.type === 'run') {
      for (const rc of content.content) {
        if (rc.type === 'shape' && 'shape' in rc) {
          const shape = rc.shape as Shape;
          if (shape.textBody && shape.textBody.content.length > 0) {
            // Convert shape with text body to TextBox
            textBoxes.push({
              type: 'textBox',
              id: shape.id,
              size: shape.size,
              position: shape.position,
              wrap: shape.wrap,
              relativeHeight: shape.relativeHeight,
              fill: shape.fill,
              outline: shape.outline,
              content: shape.textBody.content,
              margins: shape.textBody.margins,
            });
            extractedShapes.add(rc);
          } else if (isAnchoredVisualRectangle(shape)) {
            textBoxes.push(shapeToDecorativeTextBox(shape));
            extractedShapes.add(rc);
          }
        }
      }
    }
  }
  return { textBoxes, extractedShapes };
}

function omitExtractedShapes(paragraph: Paragraph, extractedShapes: Set<ShapeContent>): Paragraph {
  return {
    ...paragraph,
    content: paragraph.content.map((content) => {
      if (content.type !== 'run') return content;
      return {
        ...(content as Run),
        content: content.content.filter((rc) => rc.type !== 'shape' || !extractedShapes.has(rc)),
      };
    }),
  };
}

function isAnchoredVisualRectangle(shape: Shape): boolean {
  const hasVisualSurface =
    (shape.fill !== undefined && shape.fill.type !== 'none') || shape.outline !== undefined;
  const isRectangle = shape.shapeType === 'rect' || shape.shapeType === 'roundRect';
  return isRectangle && hasVisualSurface && (!!shape.position || isAnchoredDocxTextBox(shape));
}

function shapeToDecorativeTextBox(shape: Shape): TextBox {
  const textBox: TextBox = {
    type: 'textBox',
    id: shape.id,
    size: shape.size,
    position: shape.position,
    wrap: shape.wrap,
    relativeHeight: shape.relativeHeight,
    fill: shape.fill,
    outline: shape.outline,
    content: [],
    margins: { top: 0, right: 0, bottom: 0, left: 0 },
  };

  return textBox;
}

/**
 * Convert a TextBox to a ProseMirror textBox node
 */
function convertTextBox(
  textBox: TextBox,
  styleResolver: StyleResolver | null,
  theme: Theme | null
): PMNode {
  const widthPx = textBox.size?.width ? emuToPixels(textBox.size.width) : 200;
  const heightPx = textBox.size?.height ? emuToPixels(textBox.size.height) : undefined;

  // Convert fill color
  let fillColor: string | undefined;
  const fillHex =
    textBox.fill?.type === 'solid' ? resolveColorToHex(textBox.fill.color, theme) : null;
  if (fillHex) {
    fillColor = `#${fillHex}`;
  }

  // Convert outline
  let outlineWidth: number | undefined;
  let outlineColor: string | undefined;
  let outlineStyle: string | undefined;
  if (textBox.outline && textBox.outline.width) {
    outlineWidth = Math.round((textBox.outline.width / 914400) * 96 * 100) / 100;
    const outlineHex = resolveColorToHex(textBox.outline.color, theme);
    if (outlineHex) {
      outlineColor = `#${outlineHex}`;
    }
    outlineStyle = textBox.outline.style || 'solid';
  }

  // Convert margins from EMU to pixels
  const marginTop = textBox.margins?.top != null ? emuToPixels(textBox.margins.top) : 4;
  const marginBottom = textBox.margins?.bottom != null ? emuToPixels(textBox.margins.bottom) : 4;
  const marginLeft = textBox.margins?.left != null ? emuToPixels(textBox.margins.left) : 7;
  const marginRight = textBox.margins?.right != null ? emuToPixels(textBox.margins.right) : 7;

  // Convert text box content (paragraphs) to PM nodes
  const contentNodes: PMNode[] = [];
  for (const para of textBox.content) {
    contentNodes.push(convertParagraph(para, styleResolver));
  }

  // Ensure at least one paragraph
  if (contentNodes.length === 0) {
    contentNodes.push(schema.node('paragraph', {}, []));
  }

  return schema.node(
    'textBox',
    {
      width: widthPx,
      height: heightPx,
      textBoxId: textBox.id,
      fillColor,
      outlineWidth,
      outlineColor,
      outlineStyle,
      marginTop,
      marginBottom,
      marginLeft,
      marginRight,
      relativeHeight: textBox.relativeHeight,
      ...textBoxAnchorAttrsFromDocx(textBox),
    },
    contentNodes
  );
}
