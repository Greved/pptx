import { expect, it, vi } from 'vitest';

it('shares the package journal across independently loaded public API copies', async () => {
  const first = await import('../src/api/index.ts');
  const pres = first.createPresentation();
  const slide = first.addBlankSlide(pres);
  const shape = first.addSlideTextBox(slide, {
    x: first.emu(1),
    y: first.emu(1),
    w: first.emu(10),
    h: first.emu(10),
    text: 'before',
  });
  const baseline = first.capturePresentationChanges(pres);
  vi.resetModules();
  const second = await import('../src/api/index.ts');
  expect(second.capturePresentationChanges).not.toBe(first.capturePresentationChanges);
  const fromSecond = second.capturePresentationChanges(pres, baseline.token);
  expect(fromSecond.resetRequired).toBe(false);
  expect(fromSecond.token).toBe(baseline.token);
  first.setShapeText(shape, 'first copy writes');
  const firstRevision = first.getPackagePartInfo(pres, '/ppt/slides/slide1.xml')?.revision;
  expect(firstRevision).toBeGreaterThan(0);
  expect(second.getPackagePartInfo(pres, '/ppt/slides/slide1.xml')?.revision).toBe(firstRevision);
  const afterFirst = second.capturePresentationChanges(pres, baseline.token);
  expect(afterFirst.changedPartNames).toEqual(['/ppt/slides/slide1.xml']);
  second.setShapeText(shape, 'second copy writes');
  const afterSecond = first.capturePresentationChanges(pres, afterFirst.token);
  expect(afterSecond.resetRequired).toBe(false);
  expect(afterSecond.changedPartNames).toEqual(['/ppt/slides/slide1.xml']);
  expect(first.getPackagePartInfo(pres, '/ppt/slides/slide1.xml')?.revision).toBeGreaterThan(
    firstRevision ?? 0,
  );
  expect(first.getPackagePartInfo(pres, '/ppt/slides/slide1.xml')?.revision).toBe(
    second.getPackagePartInfo(pres, '/ppt/slides/slide1.xml')?.revision,
  );
});
