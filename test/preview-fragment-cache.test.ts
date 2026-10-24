import { afterEach, describe, expect, it, vi } from 'vitest';
import type * as PptxModule from '../src/api/index.ts';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  emu,
  getSlideShapes,
  getGroupChildren,
  groupShapes,
  setShapeBounds,
  setShapeRunFormat,
  setShapeText,
  setParagraphTabs,
  setShapeRotation,
  setShapeSlideBackgroundFill,
  setPresentationTheme,
  setSlideBackground,
  setSlideSize,
} from '../src/api/index.ts';
import { getShapeBoundsResolved as measuredBounds } from '@office-kit/pptx';
import { renderSlideToSvg, renderSlideToSvgResult } from '../packages/preview/src/index.ts';
import { renderRootFragment } from '../packages/preview/src/root-fragment-cache.ts';

vi.mock('@office-kit/pptx', async (importOriginal) => {
  const original = await importOriginal<typeof PptxModule>();
  return { ...original, getShapeBoundsResolved: vi.fn(original.getShapeBoundsResolved) };
});

afterEach(() => {
  vi.mocked(measuredBounds).mockClear();
  vi.unstubAllGlobals();
});

function richSlide(count = 40) {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  for (let index = 0; index < count; index++) {
    const shape = addSlideTextBox(slide, {
      x: emu((index % 8) * 1_400_000),
      y: emu(Math.floor(index / 8) * 1_200_000),
      w: emu(1_300_000),
      h: emu(1_000_000),
      text: `Object ${index}: rich text with wrapping\nSecond paragraph`,
    });
    setShapeRunFormat(shape, 0, 0, { bold: true, size: 14 });
  }
  return { pres, slide };
}

describe('default preview root fragments', () => {
  it('remeasures browser-font-dependent roots after the available font metrics change', () => {
    let factor = 5;
    let measurements = 0;
    vi.stubGlobal('document', {
      createElement: () => ({
        getContext: () => ({
          measureText: (text: string) => {
            measurements++;
            return {
              width: text.length * factor,
              fontBoundingBoxAscent: 10,
              fontBoundingBoxDescent: 2,
              actualBoundingBoxAscent: 10,
              actualBoundingBoxDescent: 2,
            };
          },
        }),
      }),
    });
    const { pres, slide } = richSlide(1);
    const shape = getSlideShapes(slide)[0];
    if (!shape) throw new Error('Missing text shape');
    setShapeText(shape, 'Start\tRightAligned');
    setParagraphTabs(shape, 0, { tabStops: [{ positionEmu: emu(1_000_000), alignment: 'right' }] });
    const first = renderSlideToSvgResult(pres, slide);
    const before = measurements;
    factor = 15;
    const second = renderSlideToSvgResult(pres, slide);
    expect(first.dependsOnBrowserFontMetrics).toBe(true);
    expect(second.dependsOnBrowserFontMetrics).toBe(true);
    expect(measurements).toBeGreaterThan(before);
    expect(second.svg).not.toBe(first.svg);
  });

  it('exposes the exact canonical SVG with renderer-owned paint fragments', () => {
    const { pres, slide } = richSlide(2);
    for (const options of [{}, { background: false }, { textLayout: 'svg' as const }]) {
      const result = renderSlideToSvgResult(pres, slide, options);
      expect(result.svg).toBe(renderSlideToSvg(pres, slide, options));
      expect(result.svg).toBe(result.opening + result.fragments.join('') + '</svg>');
      expect(result.fragments).toHaveLength(3);
    }
    setSlideBackground(slide, '#123456');
    const result = renderSlideToSvgResult(pres, slide);
    expect(result.svg).toBe(renderSlideToSvg(pres, slide));
    expect(result.fragments[0]).toContain('#123456');
  });

  it('bounds retained fragments and keeps recently used roots', () => {
    const pres = createPresentation();
    let renders = 0;
    const render = () => String(++renders);
    for (let index = 0; index < 512; index++)
      renderRootFragment(pres, String(index), 'revision', 'xml', render);
    expect(renderRootFragment(pres, '0', 'revision', 'xml', render)).toBe('1');
    renderRootFragment(pres, '512', 'revision', 'xml', render);
    renderRootFragment(pres, '1', 'revision', 'xml', render);
    expect(renders).toBe(514);
    expect(renderRootFragment(pres, '0', 'revision', 'xml', render)).toBe('1');
  });

  it('reuses all 40 roots and rerenders only the single changed root', () => {
    const { pres, slide } = richSlide();
    renderSlideToSvg(pres, slide);
    const firstCalls = vi.mocked(measuredBounds).mock.calls.length;
    expect(firstCalls).toBeGreaterThanOrEqual(40);
    vi.mocked(measuredBounds).mockClear();
    renderSlideToSvg(pres, slide);
    expect(measuredBounds).not.toHaveBeenCalled();
    const shape = getSlideShapes(slide)[0];
    expect(shape).toBeDefined();
    if (!shape) throw new Error('Missing root');
    setShapeBounds(shape, { x: emu(50_000), y: emu(50_000), w: emu(1_300_000), h: emu(1_000_000) });
    renderSlideToSvg(pres, slide);
    expect(vi.mocked(measuredBounds).mock.calls.length).toBe(firstCalls / 40);
    vi.mocked(measuredBounds).mockClear();
    setShapeText(shape, 'Replacement rich content');
    expect(renderSlideToSvg(pres, slide)).toContain('Replacement rich content');
    expect(vi.mocked(measuredBounds).mock.calls.length).toBe(firstCalls / 40);
  });

  it('invalidates theme and background context and bypasses custom render options', () => {
    const { pres, slide } = richSlide(2);
    renderSlideToSvg(pres, slide);
    vi.mocked(measuredBounds).mockClear();
    setPresentationTheme(pres, { dark1: '#1267AB' });
    const themed = renderSlideToSvg(pres, slide);
    expect(measuredBounds).toHaveBeenCalled();
    expect(themed).toContain('#1267AB');
    vi.mocked(measuredBounds).mockClear();
    setSlideBackground(slide, '#AE1234');
    expect(renderSlideToSvg(pres, slide)).toContain('#AE1234');
    expect(measuredBounds).toHaveBeenCalled();
    vi.mocked(measuredBounds).mockClear();
    const custom = renderSlideToSvg(pres, slide, { background: false });
    expect(custom).not.toContain('fill="#AE1234"');
    expect(measuredBounds).toHaveBeenCalled();
    vi.mocked(measuredBounds).mockClear();
    setSlideSize(pres, { width: emu(10_000_000), height: emu(6_000_000) });
    expect(renderSlideToSvg(pres, slide)).toContain('viewBox="0 0 1049.87 629.92"');
    expect(measuredBounds).toHaveBeenCalled();
  });

  it('invalidates whole group roots for descendant edits and ancestor transforms', () => {
    const { pres, slide } = richSlide(2);
    const group = groupShapes(getSlideShapes(slide));
    renderSlideToSvg(pres, slide);
    vi.mocked(measuredBounds).mockClear();
    renderSlideToSvg(pres, slide);
    expect(measuredBounds).not.toHaveBeenCalled();
    const child = getGroupChildren(group)[0];
    if (!child) throw new Error('Missing group child');
    setShapeText(child, 'Changed nested child');
    expect(renderSlideToSvg(pres, slide)).toContain('Changed nested child');
    expect(measuredBounds).toHaveBeenCalled();
    vi.mocked(measuredBounds).mockClear();
    setShapeRotation(group, 35);
    expect(renderSlideToSvg(pres, slide)).toContain('rotate(35');
    expect(measuredBounds).toHaveBeenCalled();
  });

  it('rebuilds background-fill roots with valid per-render definition references', () => {
    const { pres, slide } = richSlide(2);
    const shape = getSlideShapes(slide)[0];
    if (!shape) throw new Error('Missing root');
    setShapeSlideBackgroundFill(shape);
    for (let iteration = 0; iteration < 3; iteration++) {
      setSlideBackground(slide, iteration % 2 === 0 ? '#AE1234' : '#1267AB');
      const svg = renderSlideToSvg(pres, slide);
      const ids = new Set([...svg.matchAll(/\bid="([^"]+)"/g)].map((match) => match[1]));
      for (const reference of svg.matchAll(/(?:href="#|url\(#)([^"\s)]+)/g)) {
        expect(ids.has(reference[1])).toBe(true);
      }
    }
  });
});
