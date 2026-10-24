import { describe, expect, it, vi } from 'vitest';
import {
  addBlankSlide,
  addSlideMaster,
  addSlideShape,
  capturePresentationChanges,
  createPresentation,
  getEffectiveColorMap,
  getSlideLayout,
  getSlideLayoutBackground,
  getSlideMasterBackground,
  getSlideMasterBackgroundStyles,
  getPresentationTheme,
  getShapeFillEffective,
  getSlideMasterLayouts,
  emu,
  setPresentationTheme,
  setSlideLayoutBackground,
  setSlideMasterBackgroundStyle,
  setSlideLayout,
  setShapeStyle,
} from '../src/api/index.ts';
import * as xml from '../src/internal/xml/index.ts';

describe('read-only inheritance metadata', () => {
  it('reuses unchanged master/layout XML across repeated color reads', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const parse = vi.spyOn(xml, 'parseXml');
    try {
      const expected = getEffectiveColorMap(slide);
      const initialParses = parse.mock.calls.length;
      expect(initialParses).toBeGreaterThan(0);
      for (let index = 0; index < 40; index++) {
        expect(getEffectiveColorMap(slide)).toEqual(expected);
      }
      expect(parse).toHaveBeenCalledTimes(initialParses);
    } finally {
      parse.mockRestore();
    }
  });

  it('observes public master/layout/theme edits and keeps reads independent of returned objects', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const layout = getSlideLayout(slide);
    if (!layout) throw new Error('missing layout');
    const original = getEffectiveColorMap(slide);
    original.bg1 = 'accent6';
    expect(getEffectiveColorMap(slide).bg1).toBe('lt1');
    getSlideMasterBackground(pres, layout);
    setSlideMasterBackgroundStyle(slide, 4);
    expect(getEffectiveColorMap(slide).bg1).toBe('dk1');
    setSlideLayoutBackground(layout, '#123456');
    setPresentationTheme(pres, { accent1: '#123456' });
    expect(getSlideLayoutBackground(layout)).toMatchObject({ kind: 'solid', color: '#123456' });
    expect(getPresentationTheme(pres)?.accent1).toBe('#123456');
    setSlideLayoutBackground(layout, '#ABCDEF');
    setPresentationTheme(pres, { accent1: '#ABCDEF' });
    expect(getSlideLayoutBackground(layout)).toMatchObject({ kind: 'solid', color: '#ABCDEF' });
    expect(getPresentationTheme(pres)?.accent1).toBe('#ABCDEF');
    const baseline = capturePresentationChanges(pres);
    const master = getSlideMasterBackground(pres, layout);
    const mapping = getEffectiveColorMap(slide);
    getSlideMasterBackgroundStyles(slide);
    expect(getSlideMasterBackground(pres, layout)).toEqual(master);
    expect(getEffectiveColorMap(slide)).toEqual(mapping);
    expect(capturePresentationChanges(pres, baseline.token).token).toBe(baseline.token);
  });

  it('does not parse mutable layout documents while repeatedly resolving a shape style', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: emu(1),
      y: emu(1),
      w: emu(10),
      h: emu(10),
    });
    setShapeStyle(shape, {
      line: { idx: 1, color: 'accent1' },
      fill: { idx: 1, color: 'accent1' },
      effect: { idx: 0, color: 'accent1' },
      font: { idx: 'minor', color: 'accent1' },
    });
    const expected = getShapeFillEffective(pres, shape);
    const parse = vi.spyOn(xml, 'parseXml');
    try {
      for (let index = 0; index < 40; index++)
        expect(getShapeFillEffective(pres, shape)).toEqual(expected);
      expect(parse).not.toHaveBeenCalled();
    } finally {
      parse.mockRestore();
    }
  });

  it('follows public layout rebindings and theme/master changes when resolving a shape style', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const firstLayout = getSlideLayout(slide);
    if (!firstLayout) throw new Error('missing layout');
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: emu(1),
      y: emu(1),
      w: emu(10),
      h: emu(10),
    });
    setShapeStyle(shape, {
      line: { idx: 1, color: 'bg1' },
      fill: { idx: 1, color: 'bg1' },
      effect: { idx: 0, color: 'bg1' },
      font: { idx: 'minor', color: 'bg1' },
    });
    setPresentationTheme(pres, { dark1: '#123456', light1: '#ABCDEF' });
    expect(getShapeFillEffective(pres, shape)).toEqual({ kind: 'solid', color: '#ABCDEF' });
    setSlideMasterBackgroundStyle(slide, 4);
    expect(getShapeFillEffective(pres, shape)).toEqual({ kind: 'solid', color: '#123456' });
    const secondMaster = addSlideMaster(pres);
    const secondLayout = getSlideMasterLayouts(pres, secondMaster)[0];
    if (!secondLayout) throw new Error('missing second layout');
    setSlideLayout(slide, secondLayout);
    const secondFill = getShapeFillEffective(pres, shape);
    expect(secondFill).not.toEqual({ kind: 'solid', color: '#123456' });
    setPresentationTheme(pres, { dark1: '#FEDCBA' });
    expect(getShapeFillEffective(pres, shape)).toEqual(secondFill);
    setSlideLayout(slide, firstLayout);
    expect(getShapeFillEffective(pres, shape)).toEqual({ kind: 'solid', color: '#FEDCBA' });
  });
});
