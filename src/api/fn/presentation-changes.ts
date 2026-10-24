import { partName, relsPartNameFor, resolveTarget } from '../../internal/opc/index.ts';
import {
  ensurePackageJournal,
  type PackageChangeToken,
  type PackageChanges,
} from '../../internal/parts/package-changes.ts';
import {
  INTERNAL_PACKAGE,
  SLIDE_PART_NAME,
  type PresentationData,
  type SlideData,
} from '../_internal-symbols.ts';

export type PresentationChangeToken = PackageChangeToken;
export type PresentationChanges = PackageChanges;

const CONTENT_TYPES = '/[Content_Types].xml';
const PRESENTATION = partName('/ppt/presentation.xml');
const PRESENTATION_RELS = '/ppt/_rels/presentation.xml.rels';
const SLIDE_MASTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml';
const GLOBAL_DEPENDENCIES = [CONTENT_TYPES, PRESENTATION, PRESENTATION_RELS, '/_rels/.rels'];

/**
 * After initial setup, captures writes made through the public mutation API
 * since an earlier cursor without serializing XML, reading part bodies, or
 * enumerating the package.
 * Writes later rolled back remain changes. Identical-byte writes may also appear.
 *
 * The first call, a cursor from another presentation, or a cursor older than the
 * bounded journal returns `resetRequired: true` and no incremental names. Use
 * that result's token as a new baseline after refreshing the caller's state.
 * Tokens are opaque in-memory values; do not serialize or construct them.
 * Tracking activates lazily on the first change capture, part metadata query or
 * render-dependency query. This performs one metadata pass to instrument existing
 * parts, which start at revision zero. Subsequent captures read only the journal.
 *
 * `[Content_Types].xml` changes use the virtual name `/[Content_Types].xml`.
 * Binary mutation inputs are copied; read-only byte views returned by inspection
 * APIs must never be mutated.
 * Internal escape hatches are outside the public mutation contract.
 */
export const capturePresentationChanges = (
  pres: PresentationData,
  previous?: PresentationChangeToken,
): PresentationChanges => ensurePackageJournal(pres[INTERNAL_PACKAGE]).capture(previous);

export interface PackagePartRevisionInfo {
  readonly name: string;
  readonly contentType: string;
  /** `null` for the virtual content-types part, which is not serialized by this query. */
  readonly byteLength: number | null;
  /** Monotonic within this presentation instance; compare only for the same part. */
  readonly revision: number;
}

/** After one-time tracking setup, O(1) metadata lookup; does not read or serialize XML. */
export const getPackagePartInfo = (
  pres: PresentationData,
  name: string,
): PackagePartRevisionInfo | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const journal = ensurePackageJournal(pkg);
  if (name.toLowerCase() === CONTENT_TYPES.toLowerCase()) {
    return {
      name: CONTENT_TYPES,
      contentType: 'application/xml',
      byteLength: null,
      revision: journal.revisionOf(CONTENT_TYPES),
    };
  }
  const part = pkg.getPart(name);
  return part === null
    ? null
    : {
        name: part.name,
        contentType: part.contentType,
        byteLength: part.data.byteLength,
        revision: journal.revisionOf(part.name),
      };
};

const relationshipKind = (type: string): string => type.slice(type.lastIndexOf('/') + 1);

/**
 * Names whose revisions determine a slide render: the slide, its relationships,
 * transitive layout/master/theme/chart/media resources and global presentation
 * metadata. Missing relationship parts are included so their later creation
 * invalidates a cached render. This reads relationship metadata only; it never
 * enumerates slides/shapes or serializes XML/ZIP. Master-to-layout backreferences
 * and slide navigation links are excluded so unrelated slides are not traversed.
 */
export const getSlideRenderDependencies = (slide: SlideData): readonly string[] => {
  const pkg = slide[INTERNAL_PACKAGE];
  ensurePackageJournal(pkg);
  const names = new Set<string>(GLOBAL_DEPENDENCIES);
  const visited = new Set<string>();
  const queue = [slide[SLIDE_PART_NAME], PRESENTATION];
  while (queue.length > 0) {
    const source = queue.pop();
    if (source === undefined) break;
    const lower = source.toLowerCase();
    if (visited.has(lower)) continue;
    visited.add(lower);
    const sourcePart = pkg.getPart(source);
    names.add(sourcePart?.name ?? source);
    names.add(relsPartNameFor(source));
    const rels = pkg.getRels(source);
    if (rels === null) continue;
    for (const rel of rels.items) {
      if (rel.targetMode === 'External' || rel.target === '') continue;
      const kind = relationshipKind(rel.type);
      if (kind === 'slide') continue;
      if (sourcePart?.contentType === SLIDE_MASTER_CONTENT_TYPE && kind === 'slideLayout') continue;
      if (
        lower === PRESENTATION &&
        (kind === 'slideMaster' ||
          kind === 'slideLayout' ||
          kind === 'notesMaster' ||
          kind === 'handoutMaster')
      )
        continue;
      const target = resolveTarget(source, rel.target);
      queue.push(target);
      // Some producers encode URI characters while storing decoded ZIP names.
      // Keep both names so either an existing decoded target or a later exact
      // target invalidates the render without searching the package.
      if (target.includes('%')) {
        let decoded: string;
        try {
          decoded = decodeURIComponent(target);
        } catch (error) {
          if (error instanceof URIError) continue;
          throw error;
        }
        const decodedPart = pkg.getPart(decoded);
        if (decodedPart !== null) queue.push(decodedPart.name);
      }
    }
  }
  return [...names];
};
