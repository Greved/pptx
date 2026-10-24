import type { SlideData } from '../_internal-symbols.ts';
import { runSlideMutationBatch } from './_helpers.ts';

/**
 * Batches synchronous shape edits on one slide into a single slide-XML write.
 * Shape getters and handles update immediately; nested batches flush only when
 * the outermost callback exits. The callback's return value is preserved.
 *
 * This is batching, not a transaction: a thrown callback still flushes completed
 * edits. The caller owns rollback. Raw package bytes, mutation cursors and saves
 * inside the callback may contain the previous persisted slide XML; inspect or
 * save the package after this function returns. Other parts, including media
 * and relationships, may be written immediately by their respective APIs.
 *
 * Async functions and Promise-like return values are rejected. The callback
 * must complete all its edits synchronously and must not schedule later edits.
 */
export const batchSlideMutations = <T>(slide: SlideData, fn: () => T): T => {
  if (Object.prototype.toString.call(fn) === '[object AsyncFunction]') {
    throw new TypeError('batchSlideMutations requires a synchronous callback');
  }
  return runSlideMutationBatch(slide, () => {
    const result = fn();
    if (
      result !== null &&
      (typeof result === 'object' || typeof result === 'function') &&
      'then' in result &&
      typeof result.then === 'function'
    ) {
      throw new TypeError('batchSlideMutations does not accept Promise-like results');
    }
    return result;
  });
};
