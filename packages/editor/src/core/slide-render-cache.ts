import {
  getPackagePartInfo,
  getSlidePartName,
  getSlideRenderDependencies,
  type PresentationData,
  type SlideData,
} from '@office-kit/pptx';
import { renderSlideToSvgResult, type SlideSvgResult } from '@office-kit/pptx-preview';

/** A bounded LRU shared by the canvas and thumbnail consumers. */
export class SlidePreviewCache<Key, Value> {
  readonly #limit: number;
  readonly #budget:
    | { maxBytes: number; size: (value: Value) => number; cacheable?: (value: Value) => boolean }
    | undefined;
  #bytes = 0;
  readonly #entries = new Map<Key, { revision: string; value: Value; bytes: number }>();

  constructor(
    limit: number,
    budget?: {
      maxBytes: number;
      size: (value: Value) => number;
      cacheable?: (value: Value) => boolean;
    },
  ) {
    this.#limit = limit;
    this.#budget = budget;
  }

  preview(key: Key, revision: string, render: () => Value): Value {
    const cached = this.#entries.get(key);
    if (cached?.revision === revision) {
      this.#entries.delete(key);
      this.#entries.set(key, cached);
      return cached.value;
    }
    const value = render();
    if (cached) this.#bytes -= cached.bytes;
    this.#entries.delete(key);
    if (this.#budget?.cacheable?.(value) === false) return value;
    const bytes = this.#budget?.size(value) ?? 0;
    if (bytes > (this.#budget?.maxBytes ?? Infinity)) return value;
    this.#entries.set(key, { revision, value, bytes });
    this.#bytes += bytes;
    while (this.#entries.size > this.#limit || this.#bytes > (this.#budget?.maxBytes ?? Infinity)) {
      const oldest = this.#entries.keys().next();
      if (oldest.done) break;
      this.#bytes -= this.#entries.get(oldest.value)?.bytes ?? 0;
      this.#entries.delete(oldest.value);
    }
    return value;
  }

  clear(): void {
    this.#entries.clear();
    this.#bytes = 0;
  }
}

const MAX_CACHED_SLIDES = 32;
const presentationCaches = new WeakMap<
  PresentationData,
  SlidePreviewCache<string, SlideSvgResult>
>();

/** Weak presentation keys release replaced decks; linked-part revisions prevent stale visuals. */
export function renderSlideCached(pres: PresentationData, slide: SlideData): string {
  return renderSlidePaintCached(pres, slide).svg;
}

export function renderSlidePaintCached(pres: PresentationData, slide: SlideData): SlideSvgResult {
  let cache = presentationCaches.get(pres);
  if (!cache) {
    cache = new SlidePreviewCache(MAX_CACHED_SLIDES, {
      maxBytes: 32 * 1024 * 1024,
      cacheable: (value) => !value.dependsOnBrowserFontMetrics,
      // Count all strings conservatively even when fragments share renderer cache references.
      size: (value) =>
        2 *
        (value.svg.length +
          value.opening.length +
          value.fragments.reduce((size, fragment) => size + fragment.length, 0)),
    });
    presentationCaches.set(pres, cache);
  }
  const revision = JSON.stringify(
    getSlideRenderDependencies(slide).map((name) => [
      name,
      getPackagePartInfo(pres, name)?.revision,
    ]),
  );
  return cache.preview(getSlidePartName(slide), revision, () =>
    renderSlideToSvgResult(pres, slide),
  );
}
