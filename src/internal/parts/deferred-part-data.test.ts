import { describe, expect, it } from 'vitest';
import { partName } from '../opc/index.ts';
import { OpcPackage } from './package.ts';
import { ensurePackageJournal } from './package-changes.ts';

const NAME = partName('/ppt/slides/slide1.xml');
const bytes = (value: string): Uint8Array => new TextEncoder().encode(value);
const fixture = () => {
  const pkg = OpcPackage.empty();
  const part = pkg.addPart(NAME, 'application/xml', bytes('<original/>'));
  return { pkg, part };
};

describe('owned deferred part bytes', () => {
  it('records every pending change without materializing during observation, then caches the current bytes', () => {
    const { pkg, part } = fixture();
    const original = part.data;
    let produced = 0;
    ensurePackageJournal(pkg).deferPartData(NAME, () => {
      produced++;
      return bytes('<first/>');
    });
    const journal = ensurePackageJournal(pkg);
    const before = journal.capture();
    const oldRevision = journal.revisionOf(NAME);
    ensurePackageJournal(pkg).deferPartData(NAME, () => {
      produced++;
      return bytes('<final/>');
    });
    expect(journal.revisionOf(NAME)).toBeGreaterThan(oldRevision);
    expect(journal.capture(before.token).changedPartNames).toEqual([NAME]);
    const current = journal.capture();
    expect(produced).toBe(0);
    expect(new TextDecoder().decode(original)).toBe('<original/>');
    expect(new TextDecoder().decode(part.data)).toBe('<final/>');
    expect(pkg.getPart(NAME)?.data).toBe(part.data);
    expect(produced).toBe(1);
    expect(journal.capture(current.token).token).toBe(current.token);
  });

  it('does not report stale bytes after serializer failure and accepts a direct rollback write', () => {
    const { pkg, part } = fixture();
    const original = part.data;
    const journal = ensurePackageJournal(pkg);
    const failure = new Error('serializer failed');
    let attempts = 0;
    ensurePackageJournal(pkg).deferPartData(NAME, () => {
      attempts++;
      throw failure;
    });
    expect(() => part.data).toThrow(failure);
    expect(() => pkg.save()).toThrow(failure);
    expect(attempts).toBe(2);
    const pending = journal.capture();
    part.data = original;
    expect(journal.capture(pending.token).changedPartNames).toEqual([NAME]);
    expect(part.data).toBe(original);
    expect(OpcPackage.load(pkg.save()).getPart(NAME)?.data).toEqual(original);
    expect(attempts).toBe(2);
  });

  it('isolates packages and removed/recreated parts, and rejects absent targets', () => {
    const { pkg, part } = fixture();
    const other = fixture();
    ensurePackageJournal(pkg).deferPartData(NAME, () => bytes('<pending/>'));
    expect(new TextDecoder().decode(other.part.data)).toBe('<original/>');
    pkg.removePart(NAME);
    expect(() => ensurePackageJournal(pkg).deferPartData(NAME, () => bytes('<missing/>'))).toThrow(
      'part missing',
    );
    const replacement = pkg.addPart(NAME, 'application/xml', bytes('<replacement/>'));
    expect(new TextDecoder().decode(part.data)).toBe('<pending/>');
    expect(new TextDecoder().decode(replacement.data)).toBe('<replacement/>');
    expect(OpcPackage.load(pkg.save()).getPart(NAME)?.data).toEqual(replacement.data);
  });

  it('keeps direct writes enumerable and idempotent, including parts loaded from ZIP', () => {
    const { pkg } = fixture();
    const loaded = OpcPackage.load(pkg.save());
    const part = loaded.getPart(NAME);
    if (!part) throw new Error('missing loaded part');
    const journal = ensurePackageJournal(loaded);
    const baseline = journal.capture();
    const original = part.data;
    part.data = original;
    expect(journal.capture(baseline.token).token).toBe(baseline.token);
    ensurePackageJournal(loaded).deferPartData(NAME, () => bytes('<changed/>'));
    expect({ ...part }.data).toEqual(bytes('<changed/>'));
    const before = journal.capture();
    part.data = bytes('<direct/>');
    expect(journal.capture(before.token).changedPartNames).toEqual([NAME]);
    expect(part.data).toEqual(bytes('<direct/>'));
  });
});
