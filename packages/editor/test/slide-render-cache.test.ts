import assert from 'node:assert/strict';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
import type * as PptxModule from '@office-kit/pptx';
import type * as CacheModule from '../src/core/slide-render-cache.ts';

const bundle = await build({
  stdin: {
    contents: `export { renderSlideCached } from './src/core/slide-render-cache.ts'; export * from '@office-kit/pptx'; export { renders, setFontDependent } from '@office-kit/pptx-preview';`,
    resolveDir: fileURLToPath(new URL('..', import.meta.url)),
  },
  bundle: true,
  write: false,
  platform: 'node',
  format: 'esm',
  plugins: [
    {
      name: 'count-slide-renders',
      setup(builder) {
        builder.onResolve({ filter: /^@office-kit\/pptx-preview$/ }, () => ({
          path: 'preview-probe',
          namespace: 'probe',
        }));
        builder.onLoad({ filter: /.*/, namespace: 'probe' }, () => ({
          contents: `export let renders = 0; let fontDependent = false; export function setFontDependent(value) { fontDependent = value; } export function renderSlideToSvgResult() { renders++; return { svg: 'svg-' + renders, opening: '', fragments: [], dependsOnBrowserFontMetrics: fontDependent }; }`,
        }));
      },
    },
  ],
});
const output = bundle.outputFiles[0];
assert.ok(output);
const runtime: typeof CacheModule &
  typeof PptxModule & { renders: number; setFontDependent: (value: boolean) => void } =
  await import(`data:text/javascript;base64,${Buffer.from(output.text).toString('base64')}`);

test('whole-slide previews do not retain browser-font-dependent rendering', () => {
  const pres = runtime.createPresentation();
  const slide = runtime.addTitleSlide(pres, 'Font-dependent');
  runtime.setFontDependent(true);
  try {
    assert.notEqual(runtime.renderSlideCached(pres, slide), runtime.renderSlideCached(pres, slide));
  } finally {
    runtime.setFontDependent(false);
  }
});

test('canvas and thumbnail consumers share one render and invalidate all resource revisions', () => {
  const pres = runtime.createPresentation();
  runtime.addTitleSlide(pres, 'First');
  runtime.addTitleSlide(pres, 'Second');
  const first = runtime.getSlides(pres)[0];
  const second = runtime.getSlides(pres)[1];
  assert.ok(first);
  assert.ok(second);
  const start = runtime.renders;
  const canvas = runtime.renderSlideCached(pres, first);
  assert.equal(runtime.renderSlideCached(pres, first), canvas);
  assert.equal(runtime.renders - start, 1);
  runtime.renderSlideCached(pres, second);
  const shape = runtime.getSlideShapes(first)[0];
  assert.ok(shape);
  runtime.setShapeText(shape, 'Changed');
  runtime.renderSlideCached(pres, first);
  runtime.renderSlideCached(pres, second);
  assert.equal(runtime.renders - start, 3);
  runtime.setPresentationTheme(pres, { dark1: '#1267AB' });
  runtime.renderSlideCached(pres, first);
  runtime.renderSlideCached(pres, second);
  assert.equal(runtime.renders - start, 5);
  const replacement = runtime.createPresentation();
  runtime.addTitleSlide(replacement, 'Replacement');
  const replaced = runtime.getSlides(replacement)[0];
  assert.ok(replaced);
  runtime.renderSlideCached(replacement, replaced);
  assert.equal(runtime.renders - start, 6);
});
