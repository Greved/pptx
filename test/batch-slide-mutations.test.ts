import { describe, expect, it, vi } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  batchSlideMutations,
  capturePresentationChanges,
  createPresentation,
  emu,
  getShapeText,
  getSlideShapes,
  getSlides,
  loadPresentation,
  readPackagePart,
  savePresentation,
  setShapeText,
} from '../src/api/index.ts';
import { _internalPackageOf } from '../src/api/fn/package-introspection.ts';

describe('batchSlideMutations', () => {
  it('keeps shape getters current and writes the final slide once, including nested batches', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const part = _internalPackageOf(pres).getPart('/ppt/slides/slide1.xml');
    if (!part) throw new Error('missing slide');
    const baseline = capturePresentationChanges(pres);
    const writes = vi.spyOn(part, 'data', 'set');
    const before = readPackagePart(pres, part.name);
    const result = batchSlideMutations(slide, () => {
      const shape = addSlideTextBox(slide, {
        x: emu(1),
        y: emu(1),
        w: emu(10),
        h: emu(10),
        text: 'first',
      });
      expect(getSlideShapes(slide)).toHaveLength(1);
      setShapeText(shape, 'second');
      expect(getShapeText(shape)).toBe('second');
      batchSlideMutations(slide, () => {
        setShapeText(shape, 'final');
        expect(getShapeText(shape)).toBe('final');
      });
      expect(writes).toHaveBeenCalledTimes(0);
      expect(readPackagePart(pres, part.name)).toBe(before);
      return 'batch result';
    });
    expect(result).toBe('batch result');
    expect(writes).toHaveBeenCalledTimes(1);
    expect(capturePresentationChanges(pres, baseline.token).changedPartNames).toEqual([part.name]);
    const loaded = await loadPresentation(await savePresentation(pres));
    const loadedSlide = getSlides(loaded)[0];
    if (!loadedSlide) throw new Error('missing reloaded slide');
    expect(getSlideShapes(loadedSlide).map(getShapeText)).toEqual(['final']);
  });

  it('flushes partial changes when a callback throws so rollback can observe the dirty journal', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shape = addSlideTextBox(slide, {
      x: emu(1),
      y: emu(1),
      w: emu(10),
      h: emu(10),
      text: 'before',
    });
    const baseline = capturePresentationChanges(pres);
    const failure = new Error('callback failure');
    expect(() =>
      batchSlideMutations(slide, () => {
        setShapeText(shape, 'partial');
        throw failure;
      }),
    ).toThrow(failure);
    expect(
      new TextDecoder().decode(readPackagePart(pres, '/ppt/slides/slide1.xml') ?? undefined),
    ).toContain('partial');
    expect(capturePresentationChanges(pres, baseline.token).changedPartNames).toEqual([
      '/ppt/slides/slide1.xml',
    ]);
    setShapeText(shape, 'after');
    expect(
      new TextDecoder().decode(readPackagePart(pres, '/ppt/slides/slide1.xml') ?? undefined),
    ).toContain('after');
  });

  it('does not write for no-op batches and rejects asynchronous callbacks', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const baseline = capturePresentationChanges(pres);
    expect(batchSlideMutations(slide, () => batchSlideMutations(slide, () => 42))).toBe(42);
    expect(capturePresentationChanges(pres, baseline.token).token).toBe(baseline.token);
    expect(() => batchSlideMutations(slide, async () => 42)).toThrow(TypeError);
    expect(() => batchSlideMutations(slide, () => Promise.resolve(42))).toThrow(TypeError);
    expect(capturePresentationChanges(pres, baseline.token).token).toBe(baseline.token);
  });
});
