import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  batchSlideMutations,
  capturePresentationChanges,
  createPresentation,
  emu,
  getPackagePartInfo,
  getShapeBounds,
  getShapeText,
  getSlideShapes,
  getSlides,
  loadPresentation,
  readPackagePart,
  savePresentation,
  setShapeBounds,
  setShapeText,
} from '../src/api/index.ts';
import { _internalPackageOf } from '../src/api/fn/package-introspection.ts';
import * as presentationml from '../src/internal/presentationml/index.ts';

const NAME = '/ppt/slides/slide1.xml';
const bounds = (x: number) => ({ x: emu(x), y: emu(2), w: emu(30), h: emu(40) });
const fixture = () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideTextBox(slide, { ...bounds(1), text: 'unchanged rich content' });
  return { pres, slide, shape };
};
afterEach(() => vi.restoreAllMocks());

describe('bounds edits materialize slide bytes on demand', () => {
  it('avoids serializing 60 live edits while keeping geometry and mutation revisions current', () => {
    const { pres, slide, shape } = fixture();
    const initial = readPackagePart(pres, NAME);
    const encode = vi.spyOn(TextEncoder.prototype, 'encode');
    let cursor = capturePresentationChanges(pres).token;
    for (let x = 2; x <= 61; x++) {
      setShapeBounds(shape, bounds(x));
      expect(getShapeBounds(shape)).toEqual(bounds(x));
      expect(getShapeText(shape)).toBe('unchanged rich content');
      expect(getSlideShapes(slide)[0]).toBe(shape);
      const changes = capturePresentationChanges(pres, cursor);
      expect(changes.changedPartNames).toEqual([NAME]);
      expect(changes.token).not.toBe(cursor);
      expect(
        getPackagePartInfo(pres, NAME, { includeByteLength: false })?.revision,
      ).toBeGreaterThan(0);
      cursor = changes.token;
    }
    expect(encode).toHaveBeenCalledTimes(0);
    const current = readPackagePart(pres, NAME);
    expect(encode).toHaveBeenCalledTimes(1);
    expect(getPackagePartInfo(pres, NAME)?.byteLength).toBe(current?.byteLength);
    expect(new TextDecoder().decode(current ?? undefined)).toContain('x="61"');
    expect(current).not.toBe(initial);
    expect(readPackagePart(pres, NAME)).toBe(current);
    expect(capturePresentationChanges(pres, cursor).token).toBe(cursor);
    expect(encode).toHaveBeenCalledTimes(1);
  });

  it('materializes immediate reads and save/reload without observing a journal first', async () => {
    const { pres, shape } = fixture();
    setShapeBounds(shape, bounds(7));
    expect(new TextDecoder().decode(readPackagePart(pres, NAME) ?? undefined)).toContain('x="7"');
    setShapeBounds(shape, bounds(9));
    const loaded = await loadPresentation(await savePresentation(pres));
    const slide = getSlides(loaded)[0];
    if (!slide) throw new Error('missing reloaded slide');
    const restored = getSlideShapes(slide)[0];
    if (!restored) throw new Error('missing reloaded shape');
    expect(getShapeBounds(restored)).toEqual(bounds(9));
    expect(getShapeText(restored)).toBe('unchanged rich content');
  });

  it('preserves eager batch boundaries and subsequent non-bounds mutations', () => {
    const { pres, slide, shape } = fixture();
    const before = readPackagePart(pres, NAME);
    batchSlideMutations(slide, () => {
      setShapeBounds(shape, bounds(8));
      expect(readPackagePart(pres, NAME)).toBe(before);
    });
    const encode = vi.spyOn(TextEncoder.prototype, 'encode');
    expect(new TextDecoder().decode(readPackagePart(pres, NAME) ?? undefined)).toContain('x="8"');
    expect(encode).toHaveBeenCalledTimes(0);
    setShapeBounds(shape, bounds(10));
    setShapeText(shape, 'updated');
    const xml = new TextDecoder().decode(readPackagePart(pres, NAME) ?? undefined);
    expect(xml).toContain('x="10"');
    expect(xml).toContain('updated');
  });

  it('preserves the preceding committed bytes inside a batch following an unread bounds edit', () => {
    const { pres, slide, shape } = fixture();
    setShapeBounds(shape, bounds(12));
    batchSlideMutations(slide, () => {
      setShapeBounds(shape, bounds(13));
      const committed = new TextDecoder().decode(readPackagePart(pres, NAME) ?? undefined);
      expect(committed).toContain('x="12"');
      expect(committed).not.toContain('x="13"');
    });
    expect(new TextDecoder().decode(readPackagePart(pres, NAME) ?? undefined)).toContain('x="13"');
  });

  it('direct package replacement cancels pending bytes and remains visible to rollback detection', () => {
    const { pres, shape } = fixture();
    const part = _internalPackageOf(pres).getPart(NAME);
    if (!part) throw new Error('missing slide part');
    const original = part.data;
    const before = capturePresentationChanges(pres);
    setShapeBounds(shape, bounds(11));
    part.data = original;
    expect(part.data).toBe(original);
    expect(capturePresentationChanges(pres, before.token).changedPartNames).toEqual([NAME]);
  });

  it('keeps a failed refresh visible to rollback without returning an old byte snapshot', () => {
    const { pres, shape } = fixture();
    const part = _internalPackageOf(pres).getPart(NAME);
    if (!part) throw new Error('missing slide part');
    const original = part.data;
    const before = capturePresentationChanges(pres);
    const failure = new Error('refresh failed');
    vi.spyOn(presentationml, 'readSlidePart').mockImplementationOnce(() => {
      throw failure;
    });
    expect(() => setShapeBounds(shape, bounds(14))).toThrow(failure);
    expect(capturePresentationChanges(pres, before.token).changedPartNames).toEqual([NAME]);
    expect(new TextDecoder().decode(part.data)).toContain('x="14"');
    expect(new TextDecoder().decode(original)).toContain('x="1"');
    part.data = original;
    expect(readPackagePart(pres, NAME)).toBe(original);
  });

  it('keeps default metadata a byte-length snapshot and metadata-only reads lazy', () => {
    const { pres, shape } = fixture();
    const before = getPackagePartInfo(pres, NAME);
    const length = before?.byteLength;
    setShapeBounds(shape, bounds(1234567));
    const encode = vi.spyOn(TextEncoder.prototype, 'encode');
    expect(
      getPackagePartInfo(pres, NAME.toUpperCase(), { includeByteLength: false }),
    ).toMatchObject({ byteLength: null });
    expect(getPackagePartInfo(pres, '/missing.xml', { includeByteLength: false })).toBeNull();
    expect(
      getPackagePartInfo(pres, '/[Content_Types].xml', { includeByteLength: false }),
    ).toMatchObject({ byteLength: null });
    expect(encode).toHaveBeenCalledTimes(0);
    const after = getPackagePartInfo(pres, NAME);
    expect(encode).toHaveBeenCalledTimes(1);
    expect(after?.byteLength).toBe((length ?? 0) + 6);
    expect(before?.byteLength).toBe(length);
    expect(after?.revision).not.toBe(before?.revision);
  });
});
