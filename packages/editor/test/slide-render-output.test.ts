import assert from 'node:assert/strict';
import test from 'node:test';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  emu,
  setShapeText,
  savePresentation,
  loadPresentation,
  getSlides,
  getSlideShapes,
  setParagraphTabs,
} from '@office-kit/pptx';
import { renderSlideCached } from '../src/core/slide-render-cache.ts';

test('browser font changes invalidate measured text at both preview cache layers', () => {
  let factor = 5;
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'document');
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => ({
        getContext: () => ({
          measureText: (text: string) => ({
            width: text.length * factor,
            fontBoundingBoxAscent: 10,
            fontBoundingBoxDescent: 2,
            actualBoundingBoxAscent: 10,
            actualBoundingBoxDescent: 2,
          }),
        }),
      }),
    },
  });
  try {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shape = addSlideTextBox(slide, {
      x: emu(0),
      y: emu(0),
      w: emu(4_000_000),
      h: emu(1_000_000),
      text: 'Start\tRightAligned',
    });
    setParagraphTabs(shape, 0, { tabStops: [{ positionEmu: emu(2_000_000), alignment: 'right' }] });
    const first = renderSlideCached(pres, slide);
    factor = 15;
    assert.notEqual(renderSlideCached(pres, slide), first);
  } finally {
    if (descriptor) Object.defineProperty(globalThis, 'document', descriptor);
    else Reflect.deleteProperty(globalThis, 'document');
  }
});

test('default cached rendering updates text through the built core and preview packages', () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideTextBox(slide, {
    x: emu(0),
    y: emu(0),
    w: emu(2_000_000),
    h: emu(1_000_000),
    text: 'Original',
  });
  assert.match(renderSlideCached(pres, slide), /Original/);
  setShapeText(shape, 'Changed');
  assert.match(renderSlideCached(pres, slide), /Changed/);
});

test('cached rendering updates loaded presentation text', async () => {
  const source = createPresentation();
  const initial = addBlankSlide(source);
  addSlideTextBox(initial, {
    x: emu(0),
    y: emu(0),
    w: emu(2_000_000),
    h: emu(1_000_000),
    text: 'Original',
  });
  const pres = await loadPresentation(await savePresentation(source));
  const slide = getSlides(pres)[0];
  if (!slide) throw new Error('Missing slide');
  const shape = getSlideShapes(slide)[0];
  if (!shape) throw new Error('Missing shape');
  assert.match(renderSlideCached(pres, slide), /Original/);
  setShapeText(shape, 'Changed');
  assert.match(renderSlideCached(pres, slide), /Changed/);
});
