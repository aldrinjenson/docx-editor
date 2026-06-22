import { describe, expect, test } from 'bun:test';
import { createTiffPreviewDataUrl } from '../tiffPreview';

describe('createTiffPreviewDataUrl', () => {
  test('converts uncompressed CMYK TIFF data to a browser-renderable BMP data URL', () => {
    const tiff = makeCmykTiff(1, 1, [0, 0, 0, 255]);
    const dataUrl = createTiffPreviewDataUrl(toArrayBuffer(tiff));

    expect(dataUrl?.startsWith('data:image/bmp;base64,')).toBe(true);

    const bmp = Buffer.from(dataUrl!.split(',')[1], 'base64');
    expect(bmp[0]).toBe(0x42);
    expect(bmp[1]).toBe(0x4d);
    expect(bmp.readInt32LE(18)).toBe(1);
    expect(bmp.readInt32LE(22)).toBe(-1);
  });
});

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

function makeCmykTiff(width: number, height: number, pixelBytes: number[]): Uint8Array {
  const entryCount = 9;
  const ifdOffset = 8;
  const ifdByteLength = 2 + entryCount * 12 + 4;
  const bitsOffset = ifdOffset + ifdByteLength;
  const pixelOffset = bitsOffset + 8;
  const out = new Uint8Array(pixelOffset + pixelBytes.length);
  const view = new DataView(out.buffer);

  out[0] = 0x49;
  out[1] = 0x49;
  view.setUint16(2, 42, true);
  view.setUint32(4, ifdOffset, true);
  view.setUint16(ifdOffset, entryCount, true);

  let entryOffset = ifdOffset + 2;
  const writeEntry = (tag: number, type: number, count: number, value: number): void => {
    view.setUint16(entryOffset, tag, true);
    view.setUint16(entryOffset + 2, type, true);
    view.setUint32(entryOffset + 4, count, true);
    if (type === 3 && count === 1) view.setUint16(entryOffset + 8, value, true);
    else view.setUint32(entryOffset + 8, value, true);
    entryOffset += 12;
  };

  writeEntry(256, 4, 1, width);
  writeEntry(257, 4, 1, height);
  writeEntry(258, 3, 4, bitsOffset);
  writeEntry(259, 3, 1, 1);
  writeEntry(262, 3, 1, 5);
  writeEntry(273, 4, 1, pixelOffset);
  writeEntry(277, 3, 1, 4);
  writeEntry(279, 4, 1, pixelBytes.length);
  writeEntry(284, 3, 1, 1);

  view.setUint32(entryOffset, 0, true);
  for (let i = 0; i < 4; i++) {
    view.setUint16(bitsOffset + i * 2, 8, true);
  }
  out.set(pixelBytes, pixelOffset);

  return out;
}
