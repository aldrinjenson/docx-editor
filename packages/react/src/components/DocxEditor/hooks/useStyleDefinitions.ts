import { useCallback, useEffect, useRef, type MutableRefObject, type RefObject } from 'react';
import { DocumentAgent } from '@eigenpal/docx-editor-core/agent';
import {
  applyStyleOverrides,
  cloneStyleDefinitions,
  updateStyleDefinition,
  type StyleDefinitionPatch,
  type StyleOverrides,
} from '@eigenpal/docx-editor-core/docx';
import type { Document, StyleDefinitions } from '@eigenpal/docx-editor-core/types/document';
import type { PagedEditorRef } from '../PagedEditor';
import type { UseHistoryReturn } from '../../../hooks/useHistory';

interface UseStyleDefinitionsOptions {
  initialStyles?: StyleDefinitions;
  styleOverrides?: StyleOverrides;
  styleDefinitions?: StyleDefinitions | null;
  onStyleDefinitionsChange?: (styles: StyleDefinitions) => void;
  history: Pick<UseHistoryReturn<Document | null>, 'transformAll'>;
  historyStateRef: MutableRefObject<Document | null>;
  agentRef: MutableRefObject<DocumentAgent | null>;
  pagedEditorRef: RefObject<PagedEditorRef | null>;
  styleResolverCacheRef: MutableRefObject<unknown | null>;
  handleDocumentChangeRef: MutableRefObject<(doc: Document) => void>;
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
  initialStyles,
  styleDefinitions,
  styleOverrides,
  onStyleDefinitionsChange,
  history,
  historyStateRef,
  agentRef,
  pagedEditorRef,
  styleResolverCacheRef,
  handleDocumentChangeRef,
}: UseStyleDefinitionsOptions) {
  const { transformAll } = history;
  const baseStyleDefinitionsRef = useRef<StyleDefinitions | undefined>(
    initialStyles ? cloneStyleDefinitions(initialStyles) : undefined
  );

  const prepareLoadedDocument = useCallback(
    (doc: Document): Document => {
      const loadedStyles = doc.package.styles;
      baseStyleDefinitionsRef.current = loadedStyles
        ? cloneStyleDefinitions(loadedStyles)
        : undefined;

      if (!styleDefinitions && !styleOverrides) {
        return doc;
      }

      const effectiveStyles =
        styleDefinitions ?? applyStyleOverrides(baseStyleDefinitionsRef.current, styleOverrides);
      return effectiveStyles ? withStyleDefinitions(doc, effectiveStyles) : doc;
    },
    [styleDefinitions, styleOverrides]
  );

  useEffect(() => {
    if (!styleDefinitions && !styleOverrides) return;
    const current = historyStateRef.current;
    if (!current) return;

    const baseStyles = baseStyleDefinitionsRef.current ?? current.package.styles;
    const effectiveStyles = styleDefinitions ?? applyStyleOverrides(baseStyles, styleOverrides);
    if (!effectiveStyles) return;

    if (JSON.stringify(effectiveStyles) === JSON.stringify(current.package.styles)) {
      return;
    }

    const applyStyles = (doc: Document | null): Document | null =>
      doc ? withStyleDefinitions(doc, effectiveStyles) : doc;
    const nextDocument = applyStyles(current);
    if (!nextDocument) return;

    styleResolverCacheRef.current = null;
    transformAll(applyStyles);
    historyStateRef.current = nextDocument;
    agentRef.current = new DocumentAgent(nextDocument);
    requestAnimationFrame(() => pagedEditorRef.current?.relayout());
  }, [
    agentRef,
    historyStateRef,
    pagedEditorRef,
    styleDefinitions,
    styleOverrides,
    styleResolverCacheRef,
    transformAll,
  ]);

  const updateStyle = useCallback(
    (styleId: string, patch: StyleDefinitionPatch): StyleDefinitions | null => {
      const current = historyStateRef.current;
      if (!current) return null;

      const nextStyles = updateStyleDefinition(current.package.styles, styleId, patch);
      if (!styleDefinitions) {
        baseStyleDefinitionsRef.current = cloneStyleDefinitions(nextStyles);
      }

      const nextDocument = withStyleDefinitions(current, nextStyles);
      styleResolverCacheRef.current = null;
      historyStateRef.current = nextDocument;
      agentRef.current = new DocumentAgent(nextDocument);
      onStyleDefinitionsChange?.(nextStyles);
      handleDocumentChangeRef.current(nextDocument);
      requestAnimationFrame(() => pagedEditorRef.current?.relayout());
      return nextStyles;
    },
    [
      agentRef,
      handleDocumentChangeRef,
      historyStateRef,
      onStyleDefinitionsChange,
      pagedEditorRef,
      styleDefinitions,
      styleResolverCacheRef,
    ]
  );

  return { prepareLoadedDocument, updateStyle };
}
