import type { Part } from './package.ts';
import type { PackageChangeToken, PackageChanges } from './package-changes.ts';

/** Optional instrumentation installed only when a caller observes changes. */
export interface PackageMutationObserver {
  capture(previous?: PackageChangeToken): PackageChanges;
  revisionOf(name: string): number;
  changed(name: string): void;
  removed(name: string): void;
  observePart(part: Part): void;
}
