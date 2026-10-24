import type { Part } from './package.ts';
import type { PackageChangeToken, PackageChanges } from './package-changes.ts';

/** Installed when changes are observed or a mutation defers its part bytes. */
export interface PackageMutationObserver {
  capture(previous?: PackageChangeToken): PackageChanges;
  revisionOf(name: string): number;
  changed(name: string): void;
  removed(name: string): void;
  observePart(part: Part): void;
  deferPartData(name: string, produce: () => Uint8Array): void;
}
