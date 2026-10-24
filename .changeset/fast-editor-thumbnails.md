---
'@office-kit/pptx-editor': patch
'@office-kit/pptx-preview': patch
---

Keep large decks responsive by rendering only nearby slide thumbnails, caching previews with their linked resources, and moving or rotating existing canvas objects without rebuilding rich text on every pointer frame.

Preserve unchanged canvas and thumbnail objects and their text DOM when another object changes.

Expose renderer-owned SVG fragments alongside the complete SVG and parse only changed fragments when repainting the canvas. Bound shared previews by both count and retained string size.

Remeasure text that depends on browser font metrics instead of retaining it across font-loading changes.

Reuse unchanged root objects in default slide previews so editing one object does not repeat text layout for every object on the slide.
