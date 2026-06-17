import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createEmptyDocument } from '../../../packages/core/src/utils/createDocument';
import { updateStyleDefinition } from '../../../packages/core/src/docx/styleDefinitions';
import type { StyleDefinitions } from '../../../packages/core/src/types/document';

const STYLE_DEFINITIONS_KEY = 'definitions';

function getStyleMap(doc: Y.Doc) {
  return doc.getMap<StyleDefinitions>('styles');
}

describe('collaboration style package sync', () => {
  test('syncs style definition updates between two Y.Doc clients', () => {
    const clientA = new Y.Doc();
    const clientB = new Y.Doc();
    const seedStyles = createEmptyDocument().package.styles!;

    getStyleMap(clientA).set(STYLE_DEFINITIONS_KEY, seedStyles);
    Y.applyUpdate(clientB, Y.encodeStateAsUpdate(clientA));

    const updated = updateStyleDefinition(
      getStyleMap(clientA).get(STYLE_DEFINITIONS_KEY),
      'Heading1',
      {
        rPr: { color: { rgb: '336699' } },
      }
    );
    getStyleMap(clientA).set(STYLE_DEFINITIONS_KEY, updated);
    Y.applyUpdate(clientB, Y.encodeStateAsUpdate(clientA));

    const syncedHeading = getStyleMap(clientB)
      .get(STYLE_DEFINITIONS_KEY)
      ?.styles.find((style) => style.styleId === 'Heading1');
    expect(syncedHeading?.rPr?.color?.rgb).toBe('336699');
  });

  test('reload persistence keeps style definitions in encoded Y.Doc state', () => {
    const beforeReload = new Y.Doc();
    const seedStyles = createEmptyDocument().package.styles!;
    const updated = updateStyleDefinition(seedStyles, 'Heading2', {
      rPr: { color: { rgb: 'AA5500' } },
    });
    getStyleMap(beforeReload).set(STYLE_DEFINITIONS_KEY, updated);

    const persistedState = Y.encodeStateAsUpdate(beforeReload);
    const afterReload = new Y.Doc();
    Y.applyUpdate(afterReload, persistedState);

    const reloadedHeading = getStyleMap(afterReload)
      .get(STYLE_DEFINITIONS_KEY)
      ?.styles.find((style) => style.styleId === 'Heading2');
    expect(reloadedHeading?.rPr?.color?.rgb).toBe('AA5500');
  });
});
