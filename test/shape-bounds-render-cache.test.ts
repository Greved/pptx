import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  emu,
  readPackagePart,
  setShapeBounds,
  setShapeRunFormat,
} from '../src/api/index.ts';
import { renderSlidePaintCached } from '../packages/editor/src/core/slide-render-cache.ts';
import { renderSlideToSvgResult } from '../packages/preview/src/index.ts';

afterEach(() => vi.restoreAllMocks());

describe('live bounds through editor and preview revision caches', () => {
  it('renders 60 current frames without materializing slide bytes and matches uncached output', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shapes = Array.from({ length: 40 }, (_, index) => {
      const shape = addSlideTextBox(slide, {
        x: emu((index % 8) * 1_400_000),
        y: emu(Math.floor(index / 8) * 1_200_000),
        w: emu(1_300_000),
        h: emu(1_000_000),
        text: `Object ${index}: rich text with wrapping\nSecond paragraph`,
      });
      setShapeRunFormat(shape, 0, 0, { bold: true, size: 14 });
      return shape;
    });
    const shape = shapes[0];
    if (!shape) throw new Error('missing target');
    const encode = vi.spyOn(TextEncoder.prototype, 'encode');
    let previous = '';
    for (let index = 1; index <= 60; index++) {
      setShapeBounds(shape, {
        x: emu(index * 1000),
        y: emu(0),
        w: emu(1_300_000),
        h: emu(1_000_000),
      });
      const cached = renderSlidePaintCached(pres, slide);
      const preview = renderSlideToSvgResult(pres, slide);
      // An explicit background option bypasses root caching with the same paint.
      const reference = renderSlideToSvgResult(pres, slide, { background: true });
      expect(cached.svg).toBe(preview.svg);
      expect(cached.svg).toBe(reference.svg);
      expect(cached.svg).not.toBe(previous);
      previous = cached.svg;
    }
    expect(encode).toHaveBeenCalledTimes(0);
    expect(
      new TextDecoder().decode(readPackagePart(pres, '/ppt/slides/slide1.xml') ?? undefined),
    ).toContain('x="60000"');
    expect(encode).toHaveBeenCalledTimes(1);
  });
});
