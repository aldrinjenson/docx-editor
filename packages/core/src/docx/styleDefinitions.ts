/**
 * Public helpers for editing a document's style package.
 *
 * These functions operate on `document.package.styles` without requiring an
 * EditorView. They are intentionally immutable so adapter state, collaboration
 * providers, and headless callers can all share the same update path.
 */

import type {
  ColorValue,
  DocDefaults,
  ParagraphFormatting,
  Style,
  StyleDefinitions,
  TableCellFormatting,
  TableFormatting,
  TableRowFormatting,
  TextFormatting,
} from '../types/document';
import { mergeTextFormatting } from '../utils/textFormattingMerge';

export type TableStyleCondition = NonNullable<Style['tblStylePr']>[number];

export interface StyleDefinitionPatch extends Partial<Omit<Style, 'styleId'>> {
  /** Optional replacement style ID when adding/upserting a style. */
  styleId?: string;
  /** Convenience shorthand for `rPr.color`. Strings are RGB hex without `#`. */
  color?: string | ColorValue;
  /** Convenience shorthand for `rPr.bold`. */
  bold?: boolean;
  /** Convenience shorthand for `rPr.italic`. */
  italic?: boolean;
  /** Convenience shorthand for `rPr.fontSize` (half-points). */
  fontSize?: number;
  /** Convenience shorthand for `pPr.alignment`. */
  alignment?: ParagraphFormatting['alignment'];
}

export type StyleOverrides = Record<string, StyleDefinitionPatch | null | undefined>;

function cloneValue<T>(value: T): T {
  if (Array.isArray(value)) {
    return value.map((item) => cloneValue(item)) as T;
  }

  if (value && typeof value === 'object') {
    const cloned: Record<string, unknown> = {};
    for (const [key, nestedValue] of Object.entries(value)) {
      cloned[key] = cloneValue(nestedValue);
    }
    return cloned as T;
  }

  return value;
}

function ensureStyleDefinitions(styleDefinitions: StyleDefinitions | undefined): StyleDefinitions {
  return styleDefinitions ? cloneStyleDefinitions(styleDefinitions) : { styles: [] };
}

export function cloneStyleDefinitions(styleDefinitions: StyleDefinitions): StyleDefinitions {
  return cloneValue(styleDefinitions);
}

function markStyleModified(styleDefinitions: StyleDefinitions, styleId: string): void {
  const existing = styleDefinitions.modifiedStyleIds ?? [];
  styleDefinitions.modifiedStyleIds = existing.includes(styleId)
    ? existing
    : [...existing, styleId];
}

function mergePlainObject<T extends Record<string, unknown>>(
  base: T | undefined,
  patch: T | undefined
): T | undefined {
  if (!patch) return base ? cloneValue(base) : undefined;
  if (!base) return cloneValue(patch);

  const merged: Record<string, unknown> = cloneValue(base);
  for (const [key, patchValue] of Object.entries(patch)) {
    if (patchValue === undefined) continue;
    const baseValue = merged[key];
    if (
      patchValue &&
      typeof patchValue === 'object' &&
      !Array.isArray(patchValue) &&
      baseValue &&
      typeof baseValue === 'object' &&
      !Array.isArray(baseValue)
    ) {
      merged[key] = mergePlainObject(
        baseValue as Record<string, unknown>,
        patchValue as Record<string, unknown>
      );
    } else {
      merged[key] = cloneValue(patchValue);
    }
  }
  return merged as T;
}

function mergeParagraphFormatting(
  base: ParagraphFormatting | undefined,
  patch: ParagraphFormatting | undefined
): ParagraphFormatting | undefined {
  if (!patch) return base ? cloneValue(base) : undefined;
  if (!base) return cloneValue(patch);

  const merged: ParagraphFormatting = cloneValue(base);
  for (const key of Object.keys(patch) as (keyof ParagraphFormatting)[]) {
    const value = patch[key];
    if (value === undefined) continue;

    if (key === 'runProperties') {
      merged.runProperties = mergeTextFormatting(merged.runProperties, patch.runProperties);
    } else if (Array.isArray(value)) {
      (merged as Record<string, unknown>)[key] = cloneValue(value);
    } else if (value && typeof value === 'object') {
      (merged as Record<string, unknown>)[key] = mergePlainObject(
        (merged as Record<string, unknown>)[key] as Record<string, unknown> | undefined,
        value as Record<string, unknown>
      );
    } else {
      (merged as Record<string, unknown>)[key] = value;
    }
  }
  return merged;
}

function mergeTableFormatting(
  base: TableFormatting | undefined,
  patch: TableFormatting | undefined
): TableFormatting | undefined {
  return mergePlainObject(
    base as Record<string, unknown> | undefined,
    patch as Record<string, unknown> | undefined
  ) as TableFormatting | undefined;
}

function mergeTableRowFormatting(
  base: TableRowFormatting | undefined,
  patch: TableRowFormatting | undefined
): TableRowFormatting | undefined {
  return mergePlainObject(
    base as Record<string, unknown> | undefined,
    patch as Record<string, unknown> | undefined
  ) as TableRowFormatting | undefined;
}

function mergeTableCellFormatting(
  base: TableCellFormatting | undefined,
  patch: TableCellFormatting | undefined
): TableCellFormatting | undefined {
  return mergePlainObject(
    base as Record<string, unknown> | undefined,
    patch as Record<string, unknown> | undefined
  ) as TableCellFormatting | undefined;
}

function mergeTableStyleCondition(
  base: TableStyleCondition | undefined,
  patch: TableStyleCondition
): TableStyleCondition {
  return {
    type: patch.type,
    pPr: mergeParagraphFormatting(base?.pPr, patch.pPr),
    rPr: mergeTextFormatting(base?.rPr, patch.rPr),
    tblPr: mergeTableFormatting(base?.tblPr, patch.tblPr),
    trPr: mergeTableRowFormatting(base?.trPr, patch.trPr),
    tcPr: mergeTableCellFormatting(base?.tcPr, patch.tcPr),
  };
}

function mergeTableStyleConditions(
  base: Style['tblStylePr'],
  patch: Style['tblStylePr']
): Style['tblStylePr'] {
  if (!patch) return base ? cloneValue(base) : undefined;

  const byType = new Map<TableStyleCondition['type'], TableStyleCondition>();
  const order: TableStyleCondition['type'][] = [];

  for (const condition of base ?? []) {
    byType.set(condition.type, cloneValue(condition));
    order.push(condition.type);
  }

  for (const condition of patch) {
    if (!byType.has(condition.type)) {
      order.push(condition.type);
    }
    byType.set(condition.type, mergeTableStyleCondition(byType.get(condition.type), condition));
  }

  return order.map((type) => byType.get(type)).filter(Boolean) as TableStyleCondition[];
}

function normalizeStylePatch(styleId: string, patch: StyleDefinitionPatch): StyleDefinitionPatch {
  const normalized: StyleDefinitionPatch = {
    ...patch,
    styleId: patch.styleId ?? styleId,
  };

  const rPr: TextFormatting = { ...(patch.rPr ?? {}) };
  if (patch.color !== undefined) {
    rPr.color =
      typeof patch.color === 'string' ? { rgb: patch.color.replace(/^#/, '') } : patch.color;
  }
  if (patch.bold !== undefined) rPr.bold = patch.bold;
  if (patch.italic !== undefined) rPr.italic = patch.italic;
  if (patch.fontSize !== undefined) rPr.fontSize = patch.fontSize;
  if (Object.keys(rPr).length > 0) normalized.rPr = rPr;

  const pPr: ParagraphFormatting = { ...(patch.pPr ?? {}) };
  if (patch.alignment !== undefined) pPr.alignment = patch.alignment;
  if (Object.keys(pPr).length > 0) normalized.pPr = pPr;

  delete normalized.color;
  delete normalized.bold;
  delete normalized.italic;
  delete normalized.fontSize;
  delete normalized.alignment;

  return normalized;
}

function mergeStyle(base: Style | undefined, styleId: string, patch: StyleDefinitionPatch): Style {
  const normalized = normalizeStylePatch(styleId, patch);
  const next: Style = base
    ? cloneValue(base)
    : {
        styleId: normalized.styleId ?? styleId,
        type: normalized.type ?? 'paragraph',
      };

  if (normalized.styleId) next.styleId = normalized.styleId;
  if (normalized.type) next.type = normalized.type;

  const scalarKeys: Array<keyof Style> = [
    'name',
    'basedOn',
    'next',
    'link',
    'uiPriority',
    'hidden',
    'semiHidden',
    'unhideWhenUsed',
    'qFormat',
    'default',
    'personal',
  ];

  for (const key of scalarKeys) {
    const value = normalized[key];
    if (value !== undefined) {
      (next as unknown as Record<string, unknown>)[key] = value;
    }
  }

  next.pPr = mergeParagraphFormatting(next.pPr, normalized.pPr);
  next.rPr = mergeTextFormatting(next.rPr, normalized.rPr);
  next.tblPr = mergeTableFormatting(next.tblPr, normalized.tblPr);
  next.trPr = mergeTableRowFormatting(next.trPr, normalized.trPr);
  next.tcPr = mergeTableCellFormatting(next.tcPr, normalized.tcPr);
  next.tblStylePr = mergeTableStyleConditions(next.tblStylePr, normalized.tblStylePr);

  return next;
}

export function updateStyleDefinition(
  styleDefinitions: StyleDefinitions | undefined,
  styleId: string,
  patch: StyleDefinitionPatch
): StyleDefinitions {
  const next = ensureStyleDefinitions(styleDefinitions);
  const index = next.styles.findIndex((style) => style.styleId === styleId);
  const updated = mergeStyle(index >= 0 ? next.styles[index] : undefined, styleId, patch);

  if (index >= 0) {
    next.styles[index] = updated;
  } else {
    next.styles = [...next.styles, updated];
  }

  markStyleModified(next, updated.styleId);
  return next;
}

export function addStyleDefinition(
  styleDefinitions: StyleDefinitions | undefined,
  style: Style
): StyleDefinitions {
  return updateStyleDefinition(styleDefinitions, style.styleId, style);
}

export function updateDocDefaults(
  styleDefinitions: StyleDefinitions | undefined,
  patch: Partial<DocDefaults>
): StyleDefinitions {
  const next = ensureStyleDefinitions(styleDefinitions);
  next.docDefaults = {
    rPr: mergeTextFormatting(next.docDefaults?.rPr, patch.rPr),
    pPr: mergeParagraphFormatting(next.docDefaults?.pPr, patch.pPr),
  };
  next.docDefaultsModified = true;
  return next;
}

export function applyStyleOverrides(
  styleDefinitions: StyleDefinitions | undefined,
  overrides: StyleOverrides | null | undefined
): StyleDefinitions | undefined {
  if (!overrides || Object.keys(overrides).length === 0) {
    return styleDefinitions;
  }

  let next = ensureStyleDefinitions(styleDefinitions);
  for (const [styleId, patch] of Object.entries(overrides)) {
    if (!patch) continue;
    next = updateStyleDefinition(next, styleId, patch);
  }
  return next;
}
