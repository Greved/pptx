import type { PresentationData } from '@office-kit/pptx';

interface Fragment {
  readonly revision: string;
  readonly xml: string;
  readonly svg: string;
  readonly bytes: number;
}

const MAX_FRAGMENTS = 512;
const MAX_CACHE_BYTES = 16 * 1024 * 1024;
const caches = new WeakMap<PresentationData, { entries: Map<string, Fragment>; bytes: number }>();

/** Only complete root subtrees are cached: descendants depend on their ancestors. */
export function renderRootFragment(
  pres: PresentationData,
  key: string,
  revision: string,
  xml: string,
  render: () => string,
  cacheable: () => boolean = () => true,
): string {
  let cache = caches.get(pres);
  if (!cache) {
    cache = { entries: new Map(), bytes: 0 };
    caches.set(pres, cache);
  }
  const previous = cache.entries.get(key);
  if (previous) {
    cache.entries.delete(key);
    if (previous.revision === revision && previous.xml === xml) {
      cache.entries.set(key, previous);
      return previous.svg;
    }
    cache.bytes -= previous.bytes;
  }
  const svg = render();
  const bytes = (xml.length + svg.length + revision.length) * 2;
  if (bytes <= MAX_CACHE_BYTES && cacheable()) {
    cache.entries.set(key, { revision, xml, svg, bytes });
    cache.bytes += bytes;
  }
  while (cache.entries.size > MAX_FRAGMENTS || cache.bytes > MAX_CACHE_BYTES) {
    const oldest = cache.entries.entries().next();
    if (oldest.done) break;
    cache.entries.delete(oldest.value[0]);
    cache.bytes -= oldest.value[1].bytes;
  }
  return svg;
}
