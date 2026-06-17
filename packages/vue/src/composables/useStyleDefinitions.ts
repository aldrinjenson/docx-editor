import { shallowRef, unref, watch, type MaybeRef, type ShallowRef } from 'vue';
import type { Selection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import {
  applyStyleOverrides,
  cloneStyleDefinitions,
  updateStyleDefinition,
  type StyleDefinitionPatch,
  type StyleOverrides,
} from '@eigenpal/docx-editor-core/docx';
import type { Document, StyleDefinitions } from '@eigenpal/docx-editor-core/types/document';

interface UseStyleDefinitionsOptions {
  document: ShallowRef<Document | null>;
  editorView: ShallowRef<EditorView | null>;
  styleOverrides?: MaybeRef<StyleOverrides | null | undefined>;
  styleDefinitions?: MaybeRef<StyleDefinitions | null | undefined>;
  onChange?: (doc: Document) => void;
  onStyleDefinitionsChange?: (styles: StyleDefinitions) => void;
  refreshEditorFromDocument: (previousSelection?: Selection) => void;
  onSelectionUpdate?: () => void;
}

function withStyleDefinitions(doc: Document, styles: StyleDefinitions): Document {
  return {
    ...doc,
    package: {
      ...doc.package,
      styles,
    },
  };
}

export function useStyleDefinitions({
  document,
  editorView,
  styleDefinitions,
  styleOverrides,
  onChange,
  onStyleDefinitionsChange,
  refreshEditorFromDocument,
  onSelectionUpdate,
}: UseStyleDefinitionsOptions) {
  const baseStyleDefinitions = shallowRef<StyleDefinitions | undefined>(undefined);

  function prepareLoadedDocument(doc: Document): Document {
    const loadedStyles = doc.package.styles;
    baseStyleDefinitions.value = loadedStyles ? cloneStyleDefinitions(loadedStyles) : undefined;

    if (!unref(styleDefinitions) && !unref(styleOverrides)) {
      return doc;
    }

    const effectiveStyles =
      unref(styleDefinitions) ??
      applyStyleOverrides(baseStyleDefinitions.value, unref(styleOverrides));
    return effectiveStyles ? withStyleDefinitions(doc, effectiveStyles) : doc;
  }

  watch([() => unref(styleDefinitions), () => unref(styleOverrides)], () => {
    const current = document.value;
    if (!current || (!unref(styleDefinitions) && !unref(styleOverrides))) return;

    const baseStyles = baseStyleDefinitions.value ?? current.package.styles;
    const effectiveStyles =
      unref(styleDefinitions) ?? applyStyleOverrides(baseStyles, unref(styleOverrides));
    if (!effectiveStyles) return;

    if (JSON.stringify(effectiveStyles) === JSON.stringify(current.package.styles)) {
      return;
    }

    const previousSelection = editorView.value?.state.selection;
    document.value = withStyleDefinitions(current, effectiveStyles);
    refreshEditorFromDocument(previousSelection);
    onSelectionUpdate?.();
  });

  function updateStyle(styleId: string, patch: StyleDefinitionPatch): StyleDefinitions | null {
    const current = document.value;
    if (!current) return null;

    const nextStyles = updateStyleDefinition(current.package.styles, styleId, patch);
    if (!unref(styleDefinitions)) {
      baseStyleDefinitions.value = cloneStyleDefinitions(nextStyles);
    }

    const previousSelection = editorView.value?.state.selection;
    const nextDocument = withStyleDefinitions(current, nextStyles);
    document.value = nextDocument;
    onStyleDefinitionsChange?.(nextStyles);
    refreshEditorFromDocument(previousSelection);
    onChange?.(nextDocument);
    onSelectionUpdate?.();
    return nextStyles;
  }

  return { prepareLoadedDocument, updateStyle };
}
