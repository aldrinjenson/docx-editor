/**
 * Style serializer for `word/styles.xml`.
 *
 * The serializer patches only style definitions changed through the supported
 * style-definition API. Untouched `<w:style>` elements are preserved byte-for-
 * byte at the archive level because callers skip this serializer unless the
 * package carries modification markers.
 */

import type { Element as XmlElement } from 'xml-js';
import type { DocDefaults, Style, StyleDefinitions } from '../../types/document';
import {
  elementToXml,
  findChild,
  findChildren,
  getAttribute,
  getLocalName,
  parseXmlDocument,
} from '../xmlParser';
import { serializeParagraphFormatting } from './paragraphSerializer';
import { serializeTextFormatting } from './runSerializer';
import {
  serializeTableCellFormatting,
  serializeTableFormatting,
  serializeTableRowFormatting,
} from './tableSerializer';

const STYLES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';

const STYLE_CHILD_ORDER = [
  'name',
  'aliases',
  'basedOn',
  'next',
  'link',
  'autoRedefine',
  'hidden',
  'uiPriority',
  'semiHidden',
  'unhideWhenUsed',
  'qFormat',
  'locked',
  'personal',
  'personalCompose',
  'personalReply',
  'rsid',
  'pPr',
  'rPr',
  'tblPr',
  'trPr',
  'tcPr',
  'tblStylePr',
];

const MODELED_STYLE_CHILDREN = new Set([
  'name',
  'basedOn',
  'next',
  'link',
  'uiPriority',
  'hidden',
  'semiHidden',
  'unhideWhenUsed',
  'qFormat',
  'personal',
  'pPr',
  'rPr',
  'tblPr',
  'trPr',
  'tcPr',
  'tblStylePr',
]);

function createElement(
  name: string,
  attributes?: Record<string, string | number | boolean | undefined>,
  elements?: XmlElement[]
): XmlElement {
  const cleanedAttributes: Record<string, string> = {};
  for (const [key, value] of Object.entries(attributes ?? {})) {
    if (value !== undefined) cleanedAttributes[key] = String(value);
  }

  return {
    type: 'element',
    name,
    ...(Object.keys(cleanedAttributes).length > 0 ? { attributes: cleanedAttributes } : {}),
    ...(elements && elements.length > 0 ? { elements } : {}),
  };
}

function parseElement(xml: string): XmlElement | null {
  if (!xml) return null;
  return parseXmlDocument(xml);
}

function createValElement(
  localName: string,
  value: string | number | undefined
): XmlElement | null {
  if (value === undefined) return null;
  return createElement(`w:${localName}`, { 'w:val': String(value) });
}

function createBoolElement(localName: string, value: boolean | undefined): XmlElement | null {
  if (value === undefined) return null;
  return value
    ? createElement(`w:${localName}`)
    : createElement(`w:${localName}`, { 'w:val': '0' });
}

function groupExistingChildren(styleEl: XmlElement): {
  byLocalName: Map<string, XmlElement[]>;
  unmodeledRemainder: XmlElement[];
} {
  const byLocalName = new Map<string, XmlElement[]>();
  const unmodeledRemainder: XmlElement[] = [];

  for (const child of styleEl.elements ?? []) {
    if (child.type !== 'element') continue;
    const localName = getLocalName(child.name ?? '');
    if (STYLE_CHILD_ORDER.includes(localName)) {
      const existing = byLocalName.get(localName) ?? [];
      existing.push(child);
      byLocalName.set(localName, existing);
    } else if (!MODELED_STYLE_CHILDREN.has(localName)) {
      unmodeledRemainder.push(child);
    }
  }

  return { byLocalName, unmodeledRemainder };
}

function styleMetadataElements(style: Style): Map<string, XmlElement[]> {
  const entries = new Map<string, XmlElement[]>();
  const add = (localName: string, element: XmlElement | null) => {
    if (element) entries.set(localName, [element]);
  };

  add('name', createValElement('name', style.name));
  add('basedOn', createValElement('basedOn', style.basedOn));
  add('next', createValElement('next', style.next));
  add('link', createValElement('link', style.link));
  add('uiPriority', createValElement('uiPriority', style.uiPriority));
  add('hidden', createBoolElement('hidden', style.hidden));
  add('semiHidden', createBoolElement('semiHidden', style.semiHidden));
  add('unhideWhenUsed', createBoolElement('unhideWhenUsed', style.unhideWhenUsed));
  add('qFormat', createBoolElement('qFormat', style.qFormat));
  add('personal', createBoolElement('personal', style.personal));

  const pPr = parseElement(serializeParagraphFormatting(style.pPr));
  add('pPr', pPr);

  const rPr = parseElement(serializeTextFormatting(style.rPr));
  add('rPr', rPr);

  const tblPr = parseElement(serializeTableFormatting(style.tblPr));
  add('tblPr', tblPr);

  const trPr = parseElement(serializeTableRowFormatting(style.trPr));
  add('trPr', trPr);

  const tcPr = parseElement(serializeTableCellFormatting(style.tcPr));
  add('tcPr', tcPr);

  if (style.tblStylePr?.length) {
    entries.set(
      'tblStylePr',
      style.tblStylePr.map((part) => {
        const children = [
          parseElement(serializeParagraphFormatting(part.pPr)),
          parseElement(serializeTextFormatting(part.rPr)),
          parseElement(serializeTableFormatting(part.tblPr)),
          parseElement(serializeTableRowFormatting(part.trPr)),
          parseElement(serializeTableCellFormatting(part.tcPr)),
        ].filter(Boolean) as XmlElement[];
        return createElement('w:tblStylePr', { 'w:type': part.type }, children);
      })
    );
  }

  return entries;
}

function styleElementFromDefinition(style: Style, existing?: XmlElement): XmlElement {
  const styleEl: XmlElement = existing
    ? {
        ...existing,
        attributes: { ...(existing.attributes ?? {}) },
        elements: [...(existing.elements ?? [])],
      }
    : createElement('w:style');

  styleEl.attributes = {
    ...(styleEl.attributes ?? {}),
    'w:type': style.type,
    'w:styleId': style.styleId,
  };

  if (style.default !== undefined) {
    styleEl.attributes['w:default'] = style.default ? '1' : '0';
  } else {
    delete styleEl.attributes['w:default'];
  }

  const existingGroups = groupExistingChildren(styleEl);
  const modeled = styleMetadataElements(style);
  const nextChildren: XmlElement[] = [];
  const consumed = new Set<string>();

  for (const localName of STYLE_CHILD_ORDER) {
    const replacement = modeled.get(localName);
    if (replacement) {
      nextChildren.push(...replacement);
      consumed.add(localName);
      continue;
    }

    if (!MODELED_STYLE_CHILDREN.has(localName)) {
      const existingChildren = existingGroups.byLocalName.get(localName);
      if (existingChildren) {
        nextChildren.push(...existingChildren);
        consumed.add(localName);
      }
    }
  }

  for (const [localName, children] of existingGroups.byLocalName) {
    if (!consumed.has(localName) && !MODELED_STYLE_CHILDREN.has(localName)) {
      nextChildren.push(...children);
    }
  }

  nextChildren.push(...existingGroups.unmodeledRemainder);
  styleEl.elements = nextChildren;
  return styleEl;
}

function createEmptyStylesRoot(): XmlElement {
  return createElement('w:styles', {
    'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'mc:Ignorable': 'w14 w15',
  });
}

function getStyleId(styleEl: XmlElement): string | undefined {
  return getAttribute(styleEl, 'w', 'styleId') ?? undefined;
}

function getExistingStyleIds(root: XmlElement): Set<string> {
  return new Set(findChildren(root, 'w', 'style').map(getStyleId).filter(Boolean) as string[]);
}

function updateDocDefaultsElement(root: XmlElement, docDefaults: DocDefaults | undefined): void {
  if (!docDefaults) return;

  let docDefaultsEl = findChild(root, 'w', 'docDefaults');
  if (!docDefaultsEl) {
    docDefaultsEl = createElement('w:docDefaults');
    root.elements = [docDefaultsEl, ...(root.elements ?? [])];
  }

  const nextChildren: XmlElement[] = [];
  for (const child of docDefaultsEl.elements ?? []) {
    if (child.type !== 'element') continue;
    const localName = getLocalName(child.name ?? '');
    if (localName !== 'rPrDefault' && localName !== 'pPrDefault') {
      nextChildren.push(child);
    }
  }

  const rPr = parseElement(serializeTextFormatting(docDefaults.rPr));
  if (rPr) {
    nextChildren.push(createElement('w:rPrDefault', undefined, [rPr]));
  }

  const pPr = parseElement(serializeParagraphFormatting(docDefaults.pPr));
  if (pPr) {
    nextChildren.push(createElement('w:pPrDefault', undefined, [pPr]));
  }

  docDefaultsEl.elements = nextChildren;
}

function shouldSerializeStyles(root: XmlElement, styleDefinitions: StyleDefinitions): boolean {
  if (styleDefinitions.docDefaultsModified) return true;
  if (styleDefinitions.modifiedStyleIds && styleDefinitions.modifiedStyleIds.length > 0) {
    return true;
  }

  const existingIds = getExistingStyleIds(root);
  return styleDefinitions.styles.some((style) => !existingIds.has(style.styleId));
}

export function hasStyleDefinitionUpdates(
  styleDefinitions: StyleDefinitions | null | undefined,
  originalStylesXml?: string | null
): boolean {
  if (!styleDefinitions) return false;
  if (styleDefinitions.docDefaultsModified) return true;
  if (styleDefinitions.modifiedStyleIds && styleDefinitions.modifiedStyleIds.length > 0) {
    return true;
  }
  if (!originalStylesXml) return styleDefinitions.styles.length > 0;

  const root = parseXmlDocument(originalStylesXml);
  if (!root) return false;
  const existingIds = getExistingStyleIds(root);
  return styleDefinitions.styles.some((style) => !existingIds.has(style.styleId));
}

export function serializeStyles(
  originalStylesXml: string | undefined,
  styleDefinitions: StyleDefinitions
): string {
  const root = originalStylesXml ? parseXmlDocument(originalStylesXml) : null;
  const stylesRoot = root ?? createEmptyStylesRoot();

  if (root && !shouldSerializeStyles(root, styleDefinitions)) {
    return originalStylesXml ?? elementToXml(stylesRoot);
  }

  if (styleDefinitions.docDefaultsModified || !root) {
    updateDocDefaultsElement(stylesRoot, styleDefinitions.docDefaults);
  }

  const existingStyleElements = findChildren(stylesRoot, 'w', 'style');
  const styleElementsById = new Map<string, XmlElement>();
  for (const styleEl of existingStyleElements) {
    const styleId = getStyleId(styleEl);
    if (styleId) styleElementsById.set(styleId, styleEl);
  }

  const modifiedIds = new Set(styleDefinitions.modifiedStyleIds ?? []);
  for (const style of styleDefinitions.styles) {
    if (!style.styleId) continue;
    const existing = styleElementsById.get(style.styleId);
    if (!existing && !modifiedIds.has(style.styleId)) {
      modifiedIds.add(style.styleId);
    }
  }

  for (const style of styleDefinitions.styles) {
    if (!modifiedIds.has(style.styleId)) continue;
    const existing = styleElementsById.get(style.styleId);
    const nextStyleEl = styleElementFromDefinition(style, existing);

    if (existing) {
      Object.assign(existing, nextStyleEl);
    } else {
      stylesRoot.elements = [...(stylesRoot.elements ?? []), nextStyleEl];
    }
  }

  return elementToXml(stylesRoot);
}

export function ensureStylesContentType(contentTypesXml: string): string {
  if (contentTypesXml.includes('/word/styles.xml')) {
    return contentTypesXml;
  }
  return contentTypesXml.replace(
    '</Types>',
    `<Override PartName="/word/styles.xml" ContentType="${STYLES_CONTENT_TYPE}"/></Types>`
  );
}
