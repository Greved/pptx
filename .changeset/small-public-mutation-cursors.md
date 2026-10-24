---
'@office-kit/pptx': minor
---

Track presentation mutations with bounded, in-memory change cursors. `capturePresentationChanges` reports changed OPC parts without scanning slide XML or saving the archive, and requires a reset for replaced presentations or expired cursors. `getPackagePartInfo` and `getSlideRenderDependencies` support render caches that invalidate for layout, master, theme, chart and shared media changes.
