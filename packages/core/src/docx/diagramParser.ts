/**
 * Conservative DrawingML diagram fallback.
 *
 * Office SmartArt/diagram drawings are stored as a `<dgm:relIds>` shell in the
 * document story plus cached `word/diagrams/drawing*.xml` parts. We do not
 * implement a SmartArt layout engine here; instead we read the cached drawing
 * shapes and project their positioned text boxes into the existing shape/text
 * box pipeline. This preserves navigational header strips and similar diagram
 * labels instead of silently dropping the whole drawing.
 */

import type {
  ColorValue,
  ImagePosition,
  ImageSize,
  ImageWrap,
  MediaFile,
  Paragraph,
  RelationshipMap,
  Shape,
  ShapeContent,
  TextFormatting,
} from '../types/document';
import { parseAnchorPosition, parseAnchorWrap } from './drawingUtils';
import {
  findByFullName,
  getAttribute,
  getChildElements,
  getTextContent,
  parseNumericAttribute,
  parseXmlDocument,
  type XmlElement,
} from './xmlParser';

const DIAGRAM_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const DIAGRAM_DRAWING_REL_SUFFIX = '/diagramDrawing';

export function parseDiagramDrawingContent(
  drawingEl: XmlElement,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null
): ShapeContent[] {
  if (!rels || !media || !isDiagramDrawing(drawingEl)) return [];

  const drawingXml = resolveDiagramDrawingXml(rels, media);
  if (!drawingXml) return [];

  const root = parseXmlDocument(drawingXml);
  if (!root) return [];

  const container = getChildElements(drawingEl).find(
    (el) => el.name === 'wp:inline' || el.name === 'wp:anchor'
  );
  if (!container) return [];

  const baseExtent = findByFullName(container, 'wp:extent');
  const baseSize: ImageSize = {
    width: parseNumericAttribute(baseExtent, null, 'cx') ?? 0,
    height: parseNumericAttribute(baseExtent, null, 'cy') ?? 0,
  };
  const basePosition = container.name === 'wp:anchor' ? parseAnchorPosition(container) : undefined;
  const baseWrap: ImageWrap =
    container.name === 'wp:anchor'
      ? (parseAnchorWrap(container) ?? { type: 'inFront' })
      : { type: 'inline' };
  const baseRelativeHeight =
    container.name === 'wp:anchor'
      ? parseNumericAttribute(container, null, 'relativeHeight')
      : undefined;

  return findDescendantsByLocalName(root, 'sp')
    .map((shapeEl, index) =>
      parseDiagramShape(shapeEl, index, basePosition, baseWrap, baseSize, baseRelativeHeight)
    )
    .filter((shape): shape is ShapeContent => shape !== null);
}

function isDiagramDrawing(drawingEl: XmlElement): boolean {
  const graphicData = findDescendantsByLocalName(drawingEl, 'graphicData')[0] ?? null;
  return getAttribute(graphicData, null, 'uri') === DIAGRAM_URI;
}

function resolveDiagramDrawingXml(
  rels: RelationshipMap,
  media: Map<string, MediaFile>
): string | null {
  const drawingRel = Array.from(rels.values()).find((rel) =>
    rel.type.endsWith(DIAGRAM_DRAWING_REL_SUFFIX)
  );
  if (!drawingRel) return null;

  const candidates = [
    drawingRel.target,
    `word/${drawingRel.target}`,
    drawingRel.target.replace(/^\/+/, ''),
  ];
  for (const candidate of candidates) {
    const file = media.get(candidate);
    if (file?.text) return file.text;
  }
  return null;
}

function parseDiagramShape(
  shapeEl: XmlElement,
  index: number,
  basePosition: ImagePosition | undefined,
  baseWrap: ImageWrap,
  baseSize: ImageSize,
  baseRelativeHeight: number | undefined
): ShapeContent | null {
  const text = findDescendantsByLocalName(shapeEl, 't')
    .map((el) => getTextContent(el))
    .join('');
  if (!text.trim()) return null;

  // Cached SmartArt shapes often carry two transforms: `spPr/a:xfrm` for the
  // geometry surface (e.g. a chevron) and `dsp:txXfrm` for the text rectangle
  // inside that shape. Use the text transform when present so projected labels
  // line up with Word instead of using the wider decorative shape bounds.
  const textXfrm = findDirectChildByLocalName(shapeEl, 'txXfrm');
  const xfrm = textXfrm ?? findDescendantsByLocalName(shapeEl, 'xfrm')[0] ?? null;
  const off = xfrm ? findDescendantsByLocalName(xfrm, 'off')[0] : null;
  const ext = xfrm ? findDescendantsByLocalName(xfrm, 'ext')[0] : null;

  const offset = {
    x: parseNumericAttribute(off, null, 'x') ?? 0,
    y: parseNumericAttribute(off, null, 'y') ?? 0,
  };
  const size: ImageSize = {
    width: parseNumericAttribute(ext, null, 'cx') ?? Math.max(baseSize.width, 1),
    height: parseNumericAttribute(ext, null, 'cy') ?? Math.max(baseSize.height, 1),
  };

  const prstGeom = findDescendantsByLocalName(shapeEl, 'prstGeom')[0] ?? null;
  const shapeType = getAttribute(prstGeom, null, 'prst') ?? 'rect';
  const formatting = parseDiagramTextFormatting(shapeEl);

  const bodyPr = findDescendantsByLocalName(shapeEl, 'bodyPr')[0] ?? null;
  const paragraph: Paragraph = {
    type: 'paragraph',
    formatting: { alignment: 'center' },
    content: [
      {
        type: 'run',
        formatting,
        content: [{ type: 'text', text }],
      },
    ],
  };

  const shape: Shape = {
    type: 'shape',
    shapeType: shapeType as Shape['shapeType'],
    id: `diagram-${index}`,
    size,
    wrap: baseWrap,
    textBody: {
      margins: parseDiagramTextMargins(bodyPr),
      content: [paragraph],
    },
  };

  const position = offsetPosition(basePosition, offset);
  if (position) shape.position = position;
  if (baseRelativeHeight !== undefined) shape.relativeHeight = baseRelativeHeight;

  return { type: 'shape', shape };
}

function parseDiagramTextFormatting(shapeEl: XmlElement): TextFormatting {
  const rPr = findDescendantsByLocalName(shapeEl, 'rPr')[0] ?? null;
  const formatting: TextFormatting = {};

  const size = parseNumericAttribute(rPr, null, 'sz');
  if (size !== undefined) {
    // DrawingML text size is hundredths of a point; TextFormatting uses
    // half-points.
    formatting.fontSize = Math.max(1, Math.round(size / 50));
  }

  if (getAttribute(rPr, null, 'b') === '1') formatting.bold = true;
  if (getAttribute(rPr, null, 'i') === '1') formatting.italic = true;

  const latin = rPr ? findDescendantsByLocalName(rPr, 'latin')[0] : null;
  const typeface = getAttribute(latin, null, 'typeface');
  if (typeface) {
    formatting.fontFamily = { ascii: typeface, hAnsi: typeface };
  }

  const color = parseFirstSolidColor(rPr);
  if (color) formatting.color = color;

  return formatting;
}

function parseDiagramTextMargins(
  bodyPr: XmlElement | null
): NonNullable<Shape['textBody']>['margins'] {
  return {
    top: parseNumericAttribute(bodyPr, null, 'tIns') ?? 0,
    right: parseNumericAttribute(bodyPr, null, 'rIns') ?? 0,
    bottom: parseNumericAttribute(bodyPr, null, 'bIns') ?? 0,
    left: parseNumericAttribute(bodyPr, null, 'lIns') ?? 0,
  };
}

function parseFirstSolidColor(root: XmlElement | null): ColorValue | undefined {
  if (!root) return undefined;
  const solidFill = findDescendantsByLocalName(root, 'solidFill')[0] ?? null;
  if (!solidFill) return undefined;

  const srgb = findDescendantsByLocalName(solidFill, 'srgbClr')[0] ?? null;
  const rgb = getAttribute(srgb, null, 'val');
  if (rgb) return { rgb };

  const scheme = findDescendantsByLocalName(solidFill, 'schemeClr')[0] ?? null;
  const themeColor = mapThemeColor(getAttribute(scheme, null, 'val'));
  return themeColor ? { themeColor } : undefined;
}

function mapThemeColor(value: string | null): ColorValue['themeColor'] | undefined {
  switch (value) {
    case 'tx1':
      return 'text1';
    case 'tx2':
      return 'text2';
    case 'bg1':
      return 'background1';
    case 'bg2':
      return 'background2';
    case 'dk1':
    case 'lt1':
    case 'dk2':
    case 'lt2':
    case 'accent1':
    case 'accent2':
    case 'accent3':
    case 'accent4':
    case 'accent5':
    case 'accent6':
    case 'hlink':
    case 'folHlink':
    case 'background1':
    case 'background2':
    case 'text1':
    case 'text2':
      return value;
    default:
      return undefined;
  }
}

function offsetPosition(
  base: ImagePosition | undefined,
  offset: { x: number; y: number }
): ImagePosition | undefined {
  if (!base) return undefined;

  return {
    horizontal: {
      relativeTo: base.horizontal.relativeTo,
      posOffset: (base.horizontal.posOffset ?? 0) + offset.x,
    },
    vertical: {
      relativeTo: base.vertical.relativeTo,
      posOffset: (base.vertical.posOffset ?? 0) + offset.y,
    },
  };
}

function findDirectChildByLocalName(root: XmlElement, localName: string): XmlElement | null {
  return getChildElements(root).find((child) => getLocalName(child.name) === localName) ?? null;
}

function findDescendantsByLocalName(root: XmlElement, localName: string): XmlElement[] {
  const result: XmlElement[] = [];
  const visit = (node: XmlElement): void => {
    if (getLocalName(node.name) === localName) result.push(node);
    for (const child of getChildElements(node)) visit(child);
  };
  visit(root);
  return result;
}

function getLocalName(name: string | undefined): string {
  if (!name) return '';
  const colonIndex = name.indexOf(':');
  return colonIndex >= 0 ? name.substring(colonIndex + 1) : name;
}
