---
'@office-kit/pptx-editor': minor
---

Allow embedding applications to supply synchronous canonical document history through `mountEditor({ externalHistory })`. The adapter handles committed edits, document settings, gestures, undo/redo, cancellation and rollback without creating internal PPTX history snapshots. Hosts can refresh a directly accepted canonical change with `refreshFromExternalHistory()` while retaining the same undo stack. The existing snapshot history remains the default.
