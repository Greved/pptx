import { describe, expect, it } from 'vitest';
import { partName } from '../opc/index.ts';
import { OpcPackage } from './package.ts';
import {
  ensurePackageJournal,
  type PackageChangeToken,
  type PackageChanges,
} from './package-changes.ts';

const NAME = partName('/ppt/slides/slide1.xml');
const bytes = (): Uint8Array => new TextEncoder().encode('<slide/>');
const capture = (pkg: OpcPackage, previous?: PackageChangeToken): PackageChanges =>
  ensurePackageJournal(pkg).capture(previous);

describe('OPC mutation journal', () => {
  it('keeps reads and missing removals unchanged, but records direct part writes', () => {
    const pkg = OpcPackage.empty();
    const part = pkg.addPart(NAME, 'application/xml', bytes());
    const baseline = capture(pkg);
    pkg.getPart(NAME);
    pkg.getRels(NAME);
    pkg.removePart(partName('/missing.xml'));
    expect(capture(pkg, baseline.token)).toEqual({
      token: baseline.token,
      changedPartNames: [],
      resetRequired: false,
    });
    part.data = bytes();
    part.contentType = 'application/x-slide';
    expect(capture(pkg, baseline.token).changedPartNames).toEqual([NAME]);
  });

  it('records relationships, content type array replacement, insertion and entry edits', () => {
    const pkg = OpcPackage.empty();
    const baseline = capture(pkg);
    pkg.setRels(NAME, { items: [] });
    pkg.setRootRels({ items: [] });
    pkg.contentTypes.defaults.push({ extension: 'png', contentType: 'image/png' });
    pkg.contentTypes.overrides = [{ partName: NAME, contentType: 'application/xml' }];
    const entry = pkg.contentTypes.defaults.find((item) => item.extension === 'png');
    if (!entry) throw new Error('missing content type');
    entry.contentType = 'image/x-test';
    expect(new Set(capture(pkg, baseline.token).changedPartNames)).toEqual(
      new Set(['/ppt/slides/_rels/slide1.xml.rels', '/_rels/.rels', '/[Content_Types].xml']),
    );
    const after = capture(pkg);
    entry.contentType = 'image/x-other';
    expect(capture(pkg, after.token).changedPartNames).toEqual(['/[Content_Types].xml']);
  });

  it('records removals and retainParts, including content type overrides', () => {
    const pkg = OpcPackage.empty();
    pkg.addPart(NAME, 'custom/type', bytes());
    const extra = partName('/ppt/media/logo.png');
    pkg.addPart(extra, 'image/png', bytes());
    const baseline = capture(pkg);
    expect(pkg.removePart(NAME)).toBe(true);
    pkg.retainParts(() => false);
    expect(ensurePackageJournal(pkg).revisionOf(NAME)).toBe(0);
    expect(ensurePackageJournal(pkg).revisionOf(extra)).toBe(0);
    expect(new Set(capture(pkg, baseline.token).changedPartNames)).toEqual(
      new Set([NAME, extra, '/[Content_Types].xml']),
    );
  });

  it('fails closed for foreign and expired cursors without scanning package bytes', () => {
    const pkg = OpcPackage.empty();
    const part = pkg.addPart(NAME, 'application/xml', bytes());
    const baseline = capture(pkg);
    expect(capture(OpcPackage.empty(), baseline.token).resetRequired).toBe(true);
    for (let i = 0; i < 4097; i++) part.data = bytes();
    expect(capture(pkg, baseline.token)).toEqual({
      token: capture(pkg).token,
      changedPartNames: [],
      resetRequired: true,
    });
    const recent = capture(pkg);
    part.data = bytes();
    expect(capture(pkg, recent.token).changedPartNames).toEqual([NAME]);
  });

  it('keeps rollback writes visible and loading the same bytes creates a new identity', () => {
    const pkg = OpcPackage.empty();
    const part = pkg.addPart(NAME, 'application/xml', bytes());
    const original = part.data;
    const baseline = capture(pkg);
    part.data = new TextEncoder().encode('<edited/>');
    part.data = original;
    expect(capture(pkg, baseline.token).changedPartNames).toEqual([NAME]);
    expect(capture(OpcPackage.load(pkg.save()), baseline.token).resetRequired).toBe(true);
  });
});
