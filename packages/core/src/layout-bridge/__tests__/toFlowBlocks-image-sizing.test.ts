import { describe, expect, test } from 'bun:test';
import { Schema } from 'prosemirror-model';
import { toFlowBlocks } from '../toFlowBlocks';
import type { ImageRun, ParagraphBlock } from '../../layout-engine/types';

const schema = new Schema({
  nodes: {
    doc: { content: 'paragraph+' },
    paragraph: {
      content: 'inline*',
      group: 'block',
    },
    text: { group: 'inline' },
    image: {
      inline: true,
      group: 'inline',
      atom: true,
      attrs: {
        src: { default: '' },
        width: { default: null },
        height: { default: null },
        displayMode: { default: null },
        wrapType: { default: null },
        cssFloat: { default: null },
      },
    },
  },
});

function firstImageRun(displayMode: 'inline' | 'float'): ImageRun {
  const doc = schema.node('doc', null, [
    schema.node('paragraph', null, [
      schema.node('image', {
        src: 'data:image/png;base64,synthetic',
        width: 1123,
        height: 792,
        displayMode,
        wrapType: displayMode === 'float' ? 'inFront' : 'inline',
        cssFloat: 'none',
      }),
    ]),
  ]);
  const blocks = toFlowBlocks(doc, { pageContentHeight: 640 });
  const paragraph = blocks.find((block) => block.kind === 'paragraph') as ParagraphBlock;
  return paragraph.runs.find((run) => run.kind === 'image') as ImageRun;
}

describe('toFlowBlocks image sizing', () => {
  test('preserves floating image extents even when taller than the body area', () => {
    const image = firstImageRun('float');

    expect(image.width).toBe(1123);
    expect(image.height).toBe(792);
  });

  test('continues constraining in-flow images to the body area', () => {
    const image = firstImageRun('inline');

    expect(image.height).toBe(640);
    expect(image.width).toBe(Math.round((1123 * 640) / 792));
  });
});
