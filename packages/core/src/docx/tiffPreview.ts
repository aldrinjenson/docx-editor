/**
 * Minimal TIFF preview decoder for browser rendering.
 *
 * DOCX files can embed TIFF images that Word renders but browsers usually do
 * not. This decoder intentionally supports the conservative subset we can
 * render safely without a dependency: uncompressed, chunky 8-bit grayscale,
 * RGB, RGBA, or CMYK TIFFs. The output is a BMP data URL because BMP encoding
 * is tiny and browser-renderable; the original TIFF bytes remain untouched for
 * round-trip/export.
 */

const TAG_IMAGE_WIDTH = 256;
const TAG_IMAGE_LENGTH = 257;
const TAG_BITS_PER_SAMPLE = 258;
const TAG_COMPRESSION = 259;
const TAG_PHOTOMETRIC = 262;
const TAG_STRIP_OFFSETS = 273;
const TAG_SAMPLES_PER_PIXEL = 277;
const TAG_STRIP_BYTE_COUNTS = 279;
const TAG_PLANAR_CONFIGURATION = 284;

const TYPE_SIZES: Record<number, number> = {
  1: 1, // BYTE
  2: 1, // ASCII
  3: 2, // SHORT
  4: 4, // LONG
};

type TiffEndian = 'le' | 'be';

type TiffEntry = {
  type: number;
  count: number;
  valueFieldOffset: number;
  valueOffset: number;
};

export function createTiffPreviewDataUrl(data: ArrayBuffer): string | undefined {
  const bytes = new Uint8Array(data);
  if (bytes.length < 8) return undefined;

  const endian: TiffEndian =
    bytes[0] === 0x49 && bytes[1] === 0x49
      ? 'le'
      : bytes[0] === 0x4d && bytes[1] === 0x4d
        ? 'be'
        : 'le';
  if (readU16(bytes, 2, endian) !== 42) return undefined;

  const ifdOffset = readU32(bytes, 4, endian);
  if (ifdOffset <= 0 || ifdOffset + 2 > bytes.length) return undefined;

  const entries = readIfd(bytes, ifdOffset, endian);
  const width = firstValue(entries.get(TAG_IMAGE_WIDTH), bytes, endian);
  const height = firstValue(entries.get(TAG_IMAGE_LENGTH), bytes, endian);
  if (!width || !height) return undefined;

  const compression = firstValue(entries.get(TAG_COMPRESSION), bytes, endian) ?? 1;
  const photometric = firstValue(entries.get(TAG_PHOTOMETRIC), bytes, endian);
  const samplesPerPixel = firstValue(entries.get(TAG_SAMPLES_PER_PIXEL), bytes, endian) ?? 1;
  const planarConfiguration = firstValue(entries.get(TAG_PLANAR_CONFIGURATION), bytes, endian) ?? 1;
  if (compression !== 1 || planarConfiguration !== 1 || photometric === undefined) {
    return undefined;
  }

  const bits = valuesFor(entries.get(TAG_BITS_PER_SAMPLE), bytes, endian);
  const effectiveBits = bits.length > 0 ? bits : Array.from({ length: samplesPerPixel }, () => 8);
  if (effectiveBits.some((bit) => bit !== 8)) return undefined;

  const stripOffsets = valuesFor(entries.get(TAG_STRIP_OFFSETS), bytes, endian);
  const stripByteCounts = valuesFor(entries.get(TAG_STRIP_BYTE_COUNTS), bytes, endian);
  if (stripOffsets.length === 0 || stripByteCounts.length === 0) return undefined;

  const pixelBytes = concatenateStrips(bytes, stripOffsets, stripByteCounts);
  const requiredBytes = width * height * samplesPerPixel;
  if (pixelBytes.length < requiredBytes) return undefined;

  const bmp = tiffPixelsToBmp({
    pixelBytes,
    width,
    height,
    samplesPerPixel,
    photometric,
  });
  return bmp ? `data:image/bmp;base64,${bytesToBase64(bmp)}` : undefined;
}

function readIfd(bytes: Uint8Array, offset: number, endian: TiffEndian): Map<number, TiffEntry> {
  const count = readU16(bytes, offset, endian);
  const entries = new Map<number, TiffEntry>();
  for (let i = 0; i < count; i++) {
    const entryOffset = offset + 2 + i * 12;
    if (entryOffset + 12 > bytes.length) break;
    entries.set(readU16(bytes, entryOffset, endian), {
      type: readU16(bytes, entryOffset + 2, endian),
      count: readU32(bytes, entryOffset + 4, endian),
      valueFieldOffset: entryOffset + 8,
      valueOffset: readU32(bytes, entryOffset + 8, endian),
    });
  }
  return entries;
}

function firstValue(
  entry: TiffEntry | undefined,
  bytes: Uint8Array,
  endian: TiffEndian
): number | undefined {
  return valuesFor(entry, bytes, endian)[0];
}

function valuesFor(entry: TiffEntry | undefined, bytes: Uint8Array, endian: TiffEndian): number[] {
  if (!entry) return [];
  const typeSize = TYPE_SIZES[entry.type];
  if (!typeSize) return [];

  const byteLength = entry.count * typeSize;
  const valueOffset = byteLength <= 4 ? entry.valueFieldOffset : entry.valueOffset;
  if (valueOffset < 0 || valueOffset + byteLength > bytes.length) return [];

  const values: number[] = [];
  for (let i = 0; i < entry.count; i++) {
    const offset = valueOffset + i * typeSize;
    if (entry.type === 3) values.push(readU16(bytes, offset, endian));
    else if (entry.type === 4) values.push(readU32(bytes, offset, endian));
    else values.push(bytes[offset]);
  }
  return values;
}

function concatenateStrips(
  bytes: Uint8Array,
  stripOffsets: number[],
  stripByteCounts: number[]
): Uint8Array {
  const total = stripByteCounts.reduce((sum, count) => sum + Math.max(0, count), 0);
  const out = new Uint8Array(total);
  let cursor = 0;
  for (let i = 0; i < stripOffsets.length; i++) {
    const offset = stripOffsets[i];
    const count = stripByteCounts[i] ?? 0;
    if (offset < 0 || count <= 0 || offset + count > bytes.length) continue;
    out.set(bytes.subarray(offset, offset + count), cursor);
    cursor += count;
  }
  return cursor === out.length ? out : out.subarray(0, cursor);
}

function tiffPixelsToBmp({
  pixelBytes,
  width,
  height,
  samplesPerPixel,
  photometric,
}: {
  pixelBytes: Uint8Array;
  width: number;
  height: number;
  samplesPerPixel: number;
  photometric: number;
}): Uint8Array | undefined {
  const rowStride = width * 4;
  const pixelDataOffset = 54;
  const bmp = new Uint8Array(pixelDataOffset + rowStride * height);
  const view = new DataView(bmp.buffer);

  bmp[0] = 0x42;
  bmp[1] = 0x4d;
  view.setUint32(2, bmp.length, true);
  view.setUint32(10, pixelDataOffset, true);
  view.setUint32(14, 40, true);
  view.setInt32(18, width, true);
  view.setInt32(22, -height, true); // top-down rows
  view.setUint16(26, 1, true);
  view.setUint16(28, 32, true);
  view.setUint32(34, rowStride * height, true);

  let sourceOffset = 0;
  let targetOffset = pixelDataOffset;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const color = readTiffPixel(pixelBytes, sourceOffset, samplesPerPixel, photometric);
      if (!color) return undefined;
      bmp[targetOffset++] = color.b;
      bmp[targetOffset++] = color.g;
      bmp[targetOffset++] = color.r;
      bmp[targetOffset++] = color.a;
      sourceOffset += samplesPerPixel;
    }
  }

  return bmp;
}

function readTiffPixel(
  bytes: Uint8Array,
  offset: number,
  samplesPerPixel: number,
  photometric: number
): { r: number; g: number; b: number; a: number } | undefined {
  if (offset + samplesPerPixel > bytes.length) return undefined;

  if (photometric === 0 || photometric === 1) {
    const raw = bytes[offset];
    const gray = photometric === 0 ? 255 - raw : raw;
    return { r: gray, g: gray, b: gray, a: 255 };
  }

  if (photometric === 2 && samplesPerPixel >= 3) {
    return {
      r: bytes[offset],
      g: bytes[offset + 1],
      b: bytes[offset + 2],
      a: samplesPerPixel >= 4 ? bytes[offset + 3] : 255,
    };
  }

  if (photometric === 5 && samplesPerPixel >= 4) {
    const c = bytes[offset];
    const m = bytes[offset + 1];
    const y = bytes[offset + 2];
    const k = bytes[offset + 3];
    return {
      r: 255 - Math.min(255, c + k),
      g: 255 - Math.min(255, m + k),
      b: 255 - Math.min(255, y + k),
      a: 255,
    };
  }

  return undefined;
}

function readU16(bytes: Uint8Array, offset: number, endian: TiffEndian): number {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint16(
    offset,
    endian === 'le'
  );
}

function readU32(bytes: Uint8Array, offset: number, endian: TiffEndian): number {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(
    offset,
    endian === 'le'
  );
}

function bytesToBase64(bytes: Uint8Array): string {
  const maybeBuffer = (
    globalThis as {
      Buffer?: { from(input: Uint8Array): { toString(encoding: 'base64'): string } };
    }
  ).Buffer;
  if (maybeBuffer) return maybeBuffer.from(bytes).toString('base64');

  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}
