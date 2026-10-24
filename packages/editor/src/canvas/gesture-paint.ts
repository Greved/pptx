type TransformMode = 'move' | 'resize' | 'rotate';
const EMU_PER_SVG_PIXEL = 9525;

/** Nested groups and resize need the renderer's geometry and text layout. */
export function canTransformGesturePaint(mode: TransformMode, nested: boolean): boolean {
  return mode !== 'resize' && !nested;
}

export function gesturePaintTransform(
  mode: 'move' | 'rotate',
  point: { x: number; y: number },
  degrees = 0,
): string {
  const x = point.x / EMU_PER_SVG_PIXEL;
  const y = point.y / EMU_PER_SVG_PIXEL;
  return mode === 'move' ? `translate(${x} ${y})` : `rotate(${degrees} ${x} ${y})`;
}
