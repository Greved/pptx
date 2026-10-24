import assert from 'node:assert/strict';
import test from 'node:test';
import { SlidePreviewCache } from '../src/core/slide-render-cache.ts';

test('reuses unchanged dependencies across visible thumbnails', () => {
  const cache = new SlidePreviewCache<number, string>(8);
  let renders = 0;
  const render = () => String(++renders);
  for (let index = 0; index < 5; index++) cache.preview(index, 'initial', render);
  assert.equal(renders, 5);
  for (let index = 0; index < 5; index++) cache.preview(index, 'initial', render);
  assert.equal(renders, 5);
  assert.equal(cache.preview(0, 'theme-changed', render), '6');
  assert.equal(cache.preview(59, 'initial', render), '7');
});

test('evicts old thumbnails and clears replaced presentations', () => {
  const cache = new SlidePreviewCache<number, string>(2);
  let renders = 0;
  const render = () => String(++renders);
  assert.equal(cache.preview(0, 'a', render), '1');
  assert.equal(cache.preview(1, 'a', render), '2');
  assert.equal(cache.preview(2, 'a', render), '3');
  assert.equal(cache.preview(0, 'b', render), '4');
  cache.clear();
  assert.equal(cache.preview(0, 'b', render), '5');
});

test('bounds retained preview bytes and does not retain oversized renders', () => {
  const cache = new SlidePreviewCache<number, string>(32, {
    maxBytes: 8,
    size: (value) => value.length,
  });
  let renders = 0;
  const render = () => {
    renders++;
    return 'four';
  };
  cache.preview(0, 'a', render);
  cache.preview(1, 'a', render);
  cache.preview(2, 'a', render);
  cache.preview(0, 'a', render);
  assert.equal(renders, 4);
  cache.preview(3, 'a', () => 'oversized');
  cache.preview(3, 'a', () => {
    renders++;
    return 'oversized';
  });
  assert.equal(renders, 5);
});
