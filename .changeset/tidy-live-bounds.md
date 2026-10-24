---
'@office-kit/pptx': patch
'@office-kit/pptx-preview': patch
'@office-kit/pptx-editor': patch
---

Avoid repeatedly serializing slide XML during live bounds edits. Package reads and saves still return current bytes, and render caches can request revision metadata without materializing byte lengths.
