import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideLine,
  findSlideLayout,
  getShapeId,
  getShapeStrokeColorResolved,
  getShapeStrokeEffective,
  getShapeXmlString,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  pt,
  savePresentation,
  setPresentationTheme,
  setShapeStyle,
} from '../src/api/index.ts';
import { expectSchemaValid } from './lib/expect-schema-valid.ts';

const newDeck = async () => {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
  );
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('expected fixture layout');
  const slide = addSlide(pres, { layout });
  return { pres, slide };
};

const endpoints = {
  from: { x: inches(1), y: inches(1) },
  to: { x: inches(3), y: inches(2) },
};

describe('authored connector style defaults', () => {
  it('retains theme paint when only width is explicit, including theme changes and reload', async () => {
    const { pres, slide } = await newDeck();
    const reference = addSlideLine(slide, endpoints);
    const line = addSlideLine(slide, { ...endpoints, widthEmu: pt(2) });

    expect(getShapeStrokeEffective(pres, line)).toMatchObject({
      kind: 'solid',
      color: 'scheme:accent1',
      widthEmu: pt(2),
    });
    const before = getShapeStrokeColorResolved(pres, line);
    expect(before).not.toBeNull();
    expect(before).toBe(getShapeStrokeColorResolved(pres, reference));

    setPresentationTheme(pres, { accent1: '#000000' });
    expect(getShapeStrokeColorResolved(pres, line)).toBe('#000000');
    expect(getShapeStrokeColorResolved(pres, line)).not.toBe(before);
    expect(getShapeStrokeColorResolved(pres, line)).toBe(
      getShapeStrokeColorResolved(pres, reference),
    );
    expectSchemaValid(getSlideXmlString(slide), 'pml');

    const saved = await savePresentation(pres);
    expect(await savePresentation(pres)).toEqual(saved);
    const restored = await loadPresentation(saved);
    const restoredSlide = getSlides(restored)[0];
    if (!restoredSlide) throw new Error('expected restored slide');
    const restoredLine = getSlideShapes(restoredSlide).find(
      (shape) => getShapeId(shape) === getShapeId(line),
    );
    if (!restoredLine) throw new Error('expected restored connector');
    expect(getShapeXmlString(restoredLine)).toBe(getShapeXmlString(line));
    expect(getShapeStrokeEffective(restored, restoredLine)).toEqual(
      getShapeStrokeEffective(pres, line),
    );
    expect(getShapeStrokeColorResolved(restored, restoredLine)).toBe('#000000');
  });

  it('inherits theme width for explicit paint and preserves direct colors across theme changes', async () => {
    const { pres, slide } = await newDeck();
    const reference = addSlideLine(slide, endpoints);
    const line = addSlideLine(slide, { ...endpoints, color: '#123456' });
    const referenceStroke = getShapeStrokeEffective(pres, reference);
    expect(referenceStroke.kind).toBe('solid');
    if (referenceStroke.kind !== 'solid') throw new Error('expected theme stroke');

    expect(getShapeStrokeEffective(pres, line)).toMatchObject({
      kind: 'solid',
      color: '#123456',
      widthEmu: referenceStroke.widthEmu,
    });
    setPresentationTheme(pres, { accent1: '#000000' });
    expect(getShapeStrokeColorResolved(pres, line)).toBe('#123456');
    expectSchemaValid(getSlideXmlString(slide), 'pml');
  });

  it('preserves an imported custom no-line reference instead of reapplying authoring defaults', async () => {
    const { pres, slide } = await newDeck();
    const line = addSlideLine(slide, { ...endpoints, widthEmu: pt(2) });
    setShapeStyle(line, {
      line: { idx: 0, color: 'accent2' },
      fill: { idx: 0, color: 'accent2' },
      effect: { idx: 0, color: 'accent2' },
      font: { idx: 'major', color: 'accent2' },
    });
    const source = getShapeXmlString(line);
    const restored = await loadPresentation(await savePresentation(pres));
    const restoredSlide = getSlides(restored)[0];
    if (!restoredSlide) throw new Error('expected restored slide');
    const restoredLine = getSlideShapes(restoredSlide).find(
      (shape) => getShapeId(shape) === getShapeId(line),
    );
    if (!restoredLine) throw new Error('expected restored connector');

    expect(getShapeXmlString(restoredLine)).toBe(source);
    expect(getShapeStrokeEffective(restored, restoredLine)).toEqual({ kind: 'none' });
    expectSchemaValid(getSlideXmlString(restoredSlide), 'pml');
  });

  it('keeps fully specified lines independent from the theme style', async () => {
    const { pres, slide } = await newDeck();
    const line = addSlideLine(slide, {
      ...endpoints,
      color: '#123456',
      widthEmu: pt(2),
    });
    expect(getShapeXmlString(line)).not.toContain('<p:style>');
    expect(getShapeStrokeEffective(pres, line)).toEqual({
      kind: 'solid',
      color: '#123456',
      widthEmu: pt(2),
    });
    setPresentationTheme(pres, { accent1: '#000000' });
    expect(getShapeStrokeColorResolved(pres, line)).toBe('#123456');
    expectSchemaValid(getSlideXmlString(slide), 'pml');
  });
});
