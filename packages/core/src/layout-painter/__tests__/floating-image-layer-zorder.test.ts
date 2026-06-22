import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { renderFloatingImagesLayer } from '../floatingImageLayer';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

describe('renderFloatingImagesLayer', () => {
  test('keeps front floating image z-order on each item instead of the whole layer', () => {
    const layer = renderFloatingImagesLayer(
      [
        {
          src: 'data:image/png;base64,iVBORw0KGgo=',
          width: 20,
          height: 10,
          x: 4,
          y: 6,
          zIndex: 17,
        },
      ],
      document,
      {
        layerClass: 'floating-layer',
        itemClass: 'floating-item',
        sizing: 'inset0',
        layerMode: 'front',
      }
    );

    const item = layer.firstElementChild as HTMLElement;
    expect(layer.style.zIndex).toBe('');
    expect(item.style.zIndex).toBe('17');
  });
});
