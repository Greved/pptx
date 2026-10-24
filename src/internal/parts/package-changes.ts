import type { ContentTypes } from '../opc/index.ts';
import type { OpcPackage, Part } from './package.ts';
import type { PackageMutationObserver } from './package-mutation-observer.ts';

const TOKEN = Symbol('package-change-token');
const JOURNAL_LIMIT = 2048;

/** Opaque, in-memory cursor. It cannot be persisted or reused for another package. */
export interface PackageChangeToken {
  readonly [TOKEN]: true;
}

export interface PackageChanges {
  readonly token: PackageChangeToken;
  readonly changedPartNames: readonly string[];
  readonly resetRequired: boolean;
}

interface CursorState {
  readonly owner: PackageMutationJournal;
  readonly revision: number;
}

const cursors = new WeakMap<PackageChangeToken, CursorState>();
export const ensurePackageJournal = (pkg: OpcPackage): PackageMutationObserver => {
  // A presentation can cross independently bundled API copies. Reuse the
  // observer held by its package, including that observer's opaque-token registry.
  const existing = pkg.getMutationObserver();
  if (existing !== undefined) return existing;
  const journal = new PackageMutationJournal();
  const deferred = new WeakMap<Part, (produce: () => Uint8Array) => void>();
  const observer: PackageMutationObserver = {
    capture: (previous) => journal.capture(previous),
    revisionOf: (name) => journal.revisionOf(name),
    changed: (name) => journal.record(name),
    removed: (name) => journal.recordRemoval(name),
    observePart: (part) => {
      deferred.set(
        part,
        observePart(part, () => journal.record(part.name)),
      );
    },
    deferPartData: (name, produce) => {
      const part = pkg.getPart(name);
      const replace = part === null ? undefined : deferred.get(part);
      if (replace === undefined) throw new Error(`part missing: ${name}`);
      replace(produce);
    },
  };
  pkg.observeMutations(observer);
  const contentTypesChanged = (): void => journal.record('/[Content_Types].xml');
  let contentTypes = observeContentTypes(pkg.contentTypes, contentTypesChanged);
  Object.defineProperty(pkg, 'contentTypes', {
    get: (): ContentTypes => contentTypes,
    set(value: ContentTypes) {
      contentTypes = observeContentTypes(value, contentTypesChanged);
      contentTypesChanged();
    },
  });
  return observer;
};

const observePart = (part: Part, changed: () => void): ((produce: () => Uint8Array) => void) => {
  let data: Uint8Array | (() => Uint8Array) = part.data;
  let contentType = part.contentType;
  Object.defineProperties(part, {
    data: {
      get: (): Uint8Array => {
        // A failure retains the producer, never reports stale bytes as current.
        if (typeof data === 'function') data = data();
        return data;
      },
      set(value: Uint8Array) {
        if (data === value) return;
        data = value;
        changed();
      },
    },
    contentType: {
      get: (): string => contentType,
      set(value: string) {
        if (contentType === value) return;
        contentType = value;
        changed();
      },
    },
  });
  // One producer per owned part; replace it without evaluating earlier work.
  return (produce) => {
    data = produce;
    changed();
  };
};

/** Tracks writes rather than comparing serialized content, including writes later rolled back. */
export class PackageMutationJournal {
  private revision = 0;
  private token: PackageChangeToken = this.makeToken();
  private readonly entries = new Map<number, string>();
  private readonly partRevisions = new Map<string, number>();

  private makeToken(): PackageChangeToken {
    const token: PackageChangeToken = { [TOKEN]: true };
    Object.freeze(token);
    cursors.set(token, { owner: this, revision: this.revision });
    return token;
  }

  record(name: string): void {
    this.revision++;
    this.entries.set(this.revision, name);
    this.entries.delete(this.revision - JOURNAL_LIMIT);
    this.partRevisions.set(name.toLowerCase(), this.revision);
    this.token = this.makeToken();
  }

  recordRemoval(name: string): void {
    this.record(name);
    // Revisions are needed only for current parts. Keep removed names in the
    // bounded journal, without accumulating one tombstone per historical asset.
    this.partRevisions.delete(name.toLowerCase());
  }

  revisionOf(name: string): number {
    return this.partRevisions.get(name.toLowerCase()) ?? 0;
  }

  capture(previous?: PackageChangeToken): PackageChanges {
    const state = previous === undefined ? undefined : cursors.get(previous);
    if (
      state === undefined ||
      state.owner !== this ||
      state.revision < this.revision - JOURNAL_LIMIT
    ) {
      return { token: this.token, changedPartNames: [], resetRequired: true };
    }
    const names = new Set<string>();
    for (let revision = state.revision + 1; revision <= this.revision; revision++) {
      const name = this.entries.get(revision);
      if (name !== undefined) names.add(name);
    }
    return { token: this.token, changedPartNames: [...names], resetRequired: false };
  }
}

/** Copy caller-owned entries so retained input objects cannot bypass the observer. */
const observeEntry = <T extends object>(entry: T, changed: () => void): T =>
  new Proxy(
    { ...entry },
    {
      set(target, key, value: unknown): boolean {
        if (Object.is(Reflect.get(target, key), value)) return true;
        const result = Reflect.set(target, key, value);
        if (result) changed();
        return result;
      },
      deleteProperty(target, key): boolean {
        const existed = Reflect.has(target, key);
        const result = Reflect.deleteProperty(target, key);
        if (result && existed) changed();
        return result;
      },
      defineProperty(target, key, descriptor): boolean {
        const result = Reflect.defineProperty(target, key, descriptor);
        if (result) changed();
        return result;
      },
    },
  );

const observeEntries = <T extends object>(entries: readonly T[], changed: () => void): T[] =>
  new Proxy(
    entries.map((entry) => observeEntry(entry, changed)),
    {
      set(target, key, value: unknown): boolean {
        if (Object.is(Reflect.get(target, key), value)) return true;
        const observed =
          typeof value === 'object' && value !== null ? observeEntry(value, changed) : value;
        const result = Reflect.set(target, key, observed);
        if (result) changed();
        return result;
      },
      deleteProperty(target, key): boolean {
        const existed = Reflect.has(target, key);
        const result = Reflect.deleteProperty(target, key);
        if (result && existed) changed();
        return result;
      },
      defineProperty(target, key, descriptor): boolean {
        const result = Reflect.defineProperty(target, key, descriptor);
        if (result) changed();
        return result;
      },
    },
  );
const observeContentTypes = (value: ContentTypes, changed: () => void): ContentTypes => {
  let defaults = observeEntries(value.defaults, changed);
  let overrides = observeEntries(value.overrides, changed);
  return {
    get defaults() {
      return defaults;
    },
    set defaults(entries) {
      defaults = observeEntries(entries, changed);
      changed();
    },
    get overrides() {
      return overrides;
    },
    set overrides(entries) {
      overrides = observeEntries(entries, changed);
      changed();
    },
  };
};
