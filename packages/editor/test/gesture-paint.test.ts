import assert from 'node:assert/strict';
import test from 'node:test';
import { gesturePaintTransform, canTransformGesturePaint } from '../src/canvas/gesture-paint.ts';

test('moving rendered rich text translates its existing SVG without scaling glyphs', () => {
  assert.equal(gesturePaintTransform('move', { x: 9525, y: 19050 }), 'translate(1 2)');
});

test('rotation uses each original shape centre in SVG pixels', () => {
  assert.equal(gesturePaintTransform('rotate', { x: 19050, y: 28575 }, 45), 'rotate(45 2 3)');
});

test('resize and nested group editing retain full render for layout fidelity', () => {
  assert.equal(canTransformGesturePaint('move', false), true);
  assert.equal(canTransformGesturePaint('rotate', false), true);
  assert.equal(canTransformGesturePaint('resize', false), false);
  assert.equal(canTransformGesturePaint('move', true), false);
});
