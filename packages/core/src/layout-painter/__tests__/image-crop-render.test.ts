/**
 * Issue #811 — an OOXML `srcRect` crop must render by cropping the source
 * (object-fit: cover + object-position) so it keeps its aspect ratio, NOT by
 * stretching the whole source into the display box (which squashes it). The img
 * element keeps its display-box size so selection/resize overlays still measure
 * the correct rect.
 */
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import {
  applyImageCrop,
  applyImageSourceForBrowser,
  hasImageCrop,
  applyImageVisualAttrs,
  isBrowserRenderableImageSrc,
} from '../renderImage';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

describe('applyImageCrop (#811)', () => {
  test('a vertically-cropped image crops via object-fit, keeping its box size', () => {
    // Format_test.docx: srcRect t=25.793% b=30.317%, displayed 624×183.
    const img = document.createElement('img');
    img.style.width = '624px';
    img.style.height = '183px';
    applyImageCrop(img, { cropTop: 0.25793, cropBottom: 0.30317 });

    expect(img.style.objectFit).toBe('cover');
    // No horizontal crop → 50%; vertical = top / (top + bottom) ≈ 45.97%.
    const [posX, posY] = img.style.objectPosition.split(' ');
    expect(posX).toBe('50%');
    expect(parseFloat(posY)).toBeCloseTo(45.97, 1);

    // Element box unchanged — overlays measure the visible 624×183, not a
    // scaled/offset image.
    expect(img.style.width).toBe('624px');
    expect(img.style.height).toBe('183px');
  });

  test('horizontal crop maps to object-position X', () => {
    const img = document.createElement('img');
    applyImageCrop(img, { cropLeft: 0.1, cropRight: 0.3 });
    const [posX] = img.style.objectPosition.split(' ');
    expect(parseFloat(posX)).toBeCloseTo(25, 1); // 0.1 / (0.1 + 0.3) = 25%
  });

  test('an uncropped image is left untouched', () => {
    const img = document.createElement('img');
    expect(hasImageCrop({})).toBe(false);
    applyImageCrop(img, {});
    expect(img.style.objectFit).toBe('');
    expect(img.style.objectPosition).toBe('');
  });

  test('applyImageVisualAttrs applies crop AND opacity (single entry point)', () => {
    // Guards the footgun: a caller using the one visual-attrs entry point must
    // get the crop too, not silently drop it (regression for floating images).
    const img = document.createElement('img');
    applyImageVisualAttrs(img, { cropTop: 0.2, cropBottom: 0.1, opacity: 0.5 });
    expect(img.style.objectFit).toBe('cover');
    expect(img.style.opacity).toBe('0.5');
  });

  test('unsupported browser image formats are hidden instead of painted broken', () => {
    const img = document.createElement('img');
    const tiff = 'data:image/tiff;base64,AAAA';

    expect(isBrowserRenderableImageSrc(tiff)).toBe(false);
    applyImageSourceForBrowser(img, tiff);

    expect(img.getAttribute('src')).toBeNull();
    expect(img.dataset.unsupportedImage).toBe('true');
    expect(img.dataset.unsupportedImageMime).toBe('image/tiff');
    expect(img.style.visibility).toBe('hidden');
  });

  test('missing image sources are hidden without unsupported-image diagnostics', () => {
    const img = document.createElement('img');

    expect(isBrowserRenderableImageSrc(undefined)).toBe(false);
    applyImageSourceForBrowser(img, undefined);

    expect(img.getAttribute('src')).toBeNull();
    expect(img.dataset.unsupportedImage).toBeUndefined();
    expect(img.dataset.unsupportedImageMime).toBeUndefined();
    expect(img.style.visibility).toBe('hidden');
  });

  test('browser-renderable image sources are kept', () => {
    const img = document.createElement('img');
    const png = 'data:image/png;base64,AAAA';

    expect(isBrowserRenderableImageSrc(png)).toBe(true);
    applyImageSourceForBrowser(img, png);

    expect(img.getAttribute('src')).toBe(png);
    expect(img.dataset.unsupportedImage).toBeUndefined();
  });
});
