---
'@office-kit/pptx': minor
---

Add `batchSlideMutations(slide, callback)` to coalesce synchronous shape edits into one slide-XML write. Shape getters remain current inside nested batches. Completed edits are flushed even if the callback throws, allowing callers to detect partial changes and perform their own rollback.
