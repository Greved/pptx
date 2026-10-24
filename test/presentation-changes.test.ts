import { Buffer } from 'node:buffer';
import { describe, expect, it, vi } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  addSlideImage,
  capturePresentationChanges,
  createPresentation,
  emu,
  getPackagePartInfo,
  getSlideRenderDependencies,
  getSlideShapes,
  loadPresentation,
  savePresentation,
  readPackagePart,
  setCoreProperties,
  setMediaPartBytes,
  setShapeText,
  setThumbnail,
} from '../src/api/index.ts';
import { _internalPackageOf } from '../src/api/fn/package-introspection.ts';
import { partName, resolveTarget } from '../src/internal/opc/index.ts';

describe('public presentation change tracking', () => {
  it.each(['metadata', 'dependencies'])('activates revisions from the first %s query', (kind) => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shape = addSlideTextBox(slide, {
      x: emu(1),
      y: emu(1),
      w: emu(10),
      h: emu(10),
      text: 'before',
    });
    if (kind === 'dependencies') getSlideRenderDependencies(slide);
    const before = getPackagePartInfo(pres, '/ppt/slides/slide1.xml')?.revision;
    expect(before).toBe(0);
    setShapeText(shape, 'after');
    expect(getPackagePartInfo(pres, '/ppt/slides/slide1.xml')?.revision).toBeGreaterThan(0);
    expect(capturePresentationChanges(pres).resetRequired).toBe(true);
  });

  it('tracks active/offslide edits and metadata without reading or serializing unrelated parts', () => {
    const pres = createPresentation();
    const first = addBlankSlide(pres);
    const second = addBlankSlide(pres);
    const firstShape = addSlideTextBox(first, {
      x: emu(1),
      y: emu(1),
      w: emu(10),
      h: emu(10),
      text: 'A',
    });
    const secondShape = addSlideTextBox(second, {
      x: emu(1),
      y: emu(1),
      w: emu(10),
      h: emu(10),
      text: 'B',
    });
    const baseline = capturePresentationChanges(pres);
    setShapeText(firstShape, 'edited');
    setShapeText(secondShape, 'offslide');
    setCoreProperties(pres, { title: 'metadata' });
    const pkg = _internalPackageOf(pres);
    const partsRead = vi.spyOn(pkg, 'parts', 'get').mockImplementation(() => {
      throw new Error('all-part scan');
    });
    const saved = vi.spyOn(pkg, 'save').mockImplementation(() => {
      throw new Error('ZIP serialization');
    });
    const changes = capturePresentationChanges(pres, baseline.token);
    expect(changes.resetRequired).toBe(false);
    expect(changes.changedPartNames).toEqual(
      expect.arrayContaining([
        '/ppt/slides/slide1.xml',
        '/ppt/slides/slide2.xml',
        '/docProps/core.xml',
      ]),
    );
    expect(getPackagePartInfo(pres, '/PPT/SLIDES/SLIDE1.XML')?.revision).toBeGreaterThan(0);
    partsRead.mockRestore();
    saved.mockRestore();
    const after = capturePresentationChanges(pres);
    getSlideShapes(first);
    expect(capturePresentationChanges(pres, after.token).changedPartNames).toEqual([]);
  });

  it('reports shared resources, global content types and transitive render dependencies', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    addBlankSlide(pres);
    const pkg = _internalPackageOf(pres);
    const imageName = partName('/ppt/media/shared.png');
    pkg.addPart(imageName, 'image/png', new Uint8Array([1]));
    const slideName = partName('/ppt/slides/slide1.xml');
    const rels = pkg.getRels(slideName);
    if (!rels) throw new Error('missing slide relationships');
    rels.items.push({
      id: 'rId99',
      type: 'image',
      target: '../media/shared.png',
      targetMode: 'Internal',
    });
    pkg.setRels(slideName, rels);
    const dependencies = getSlideRenderDependencies(slide);
    expect(dependencies).toEqual(
      expect.arrayContaining([
        slideName,
        '/ppt/slides/_rels/slide1.xml.rels',
        imageName,
        '/ppt/slideMasters/slideMaster1.xml',
        '/ppt/theme/theme1.xml',
        '/ppt/presentation.xml',
        '/[Content_Types].xml',
      ]),
    );
    expect(dependencies).not.toContain('/ppt/slides/slide2.xml');
    const baseline = capturePresentationChanges(pres);
    const revision = getPackagePartInfo(pres, imageName)?.revision;
    expect(setMediaPartBytes(pres, imageName, new Uint8Array([2]))).toBe(true);
    expect(capturePresentationChanges(pres, baseline.token).changedPartNames).toEqual([imageName]);
    expect(getPackagePartInfo(pres, imageName)?.revision).not.toBe(revision);
    expect(getPackagePartInfo(pres, '/[Content_Types].xml')?.revision).toBe(0);
  });

  it('invalidates a cursor when a presentation is replaced by a loaded copy', async () => {
    const pres = createPresentation();
    addBlankSlide(pres);
    const baseline = capturePresentationChanges(pres);
    const replacement = await loadPresentation(await savePresentation(pres));
    expect(capturePresentationChanges(replacement, baseline.token).resetRequired).toBe(true);
  });

  it('finds encoded media targets and excludes sibling layouts even with custom master part names', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const pkg = _internalPackageOf(pres);
    const slideName = partName('/ppt/slides/slide1.xml');
    const slideRels = pkg.getRels(slideName);
    const layoutRel = slideRels?.items.find((rel) => rel.type.endsWith('/slideLayout'));
    if (!slideRels || !layoutRel) throw new Error('missing slide layout');
    const layoutName = resolveTarget(slideName, layoutRel.target);
    const master = pkg.getPart('/ppt/slideMasters/slideMaster1.xml');
    if (!master) throw new Error('missing master');
    const customMaster = partName('/custom/master.xml');
    pkg.addPart(customMaster, master.contentType, master.data);
    const sibling = partName('/custom/unrelated-layout.xml');
    pkg.addPart(sibling, 'application/xml', new Uint8Array());
    pkg.setRels(layoutName, {
      items: [
        { id: 'rId1', type: 'test/slideMaster', target: customMaster, targetMode: 'Internal' },
      ],
    });
    pkg.setRels(customMaster, {
      items: [
        { id: 'rId1', type: 'test/slideLayout', target: sibling, targetMode: 'Internal' },
        { id: 'rId2', type: 'test/theme', target: '/ppt/theme/theme1.xml', targetMode: 'Internal' },
      ],
    });
    const imageName = partName('/ppt/media/shared logo.png');
    pkg.addPart(imageName, 'image/png', new Uint8Array([1]));
    slideRels.items.push({
      id: 'rId99',
      type: 'test/image',
      target: '../media/shared%20logo.png',
      targetMode: 'Internal',
    });
    pkg.setRels(slideName, slideRels);
    const dependencies = getSlideRenderDependencies(slide);
    expect(dependencies).toContain(imageName);
    expect(dependencies).toContain(customMaster);
    expect(dependencies).not.toContain(sibling);
  });

  it.each([
    { name: 'Uint8Array', from: (values: number[]): Uint8Array => new Uint8Array(values) },
    { name: 'Buffer', from: (values: number[]): Uint8Array => Buffer.from(values) },
  ])(
    'owns public $name inputs so retained buffers cannot silently change a package',
    ({ from }) => {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const image = from([1, 2, 3]);
      addSlideImage(slide, image, { x: emu(1), y: emu(1), w: emu(10), h: emu(10), format: 'png' });
      const thumbnail = from([4, 5, 6]);
      setThumbnail(pres, thumbnail, { format: 'png' });
      const baseline = capturePresentationChanges(pres);
      image.fill(0);
      thumbnail.fill(0);
      expect(readPackagePart(pres, '/ppt/media/image1.png')).toEqual(new Uint8Array([1, 2, 3]));
      expect(readPackagePart(pres, '/docProps/thumbnail.png')).toEqual(new Uint8Array([4, 5, 6]));
      expect(capturePresentationChanges(pres, baseline.token).changedPartNames).toEqual([]);
      const replacement = from([7, 8, 9]);
      setMediaPartBytes(pres, '/ppt/media/image1.png', replacement);
      const after = capturePresentationChanges(pres);
      replacement.fill(0);
      expect(readPackagePart(pres, '/ppt/media/image1.png')).toEqual(new Uint8Array([7, 8, 9]));
      expect(capturePresentationChanges(pres, after.token).changedPartNames).toEqual([]);
    },
  );

  it('owns loaded Buffer archives, including stored binary parts', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    addSlideImage(slide, new Uint8Array([1, 2, 3]), {
      x: emu(1),
      y: emu(1),
      w: emu(10),
      h: emu(10),
      format: 'png',
    });
    const archive = Buffer.from(await savePresentation(pres));
    const loaded = await loadPresentation(archive);
    const baseline = capturePresentationChanges(loaded);
    archive.fill(0);
    expect(readPackagePart(loaded, '/ppt/media/image1.png')).toEqual(new Uint8Array([1, 2, 3]));
    expect(capturePresentationChanges(loaded, baseline.token).changedPartNames).toEqual([]);
  });
});
