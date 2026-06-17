---
'@eigenpal/docx-editor-core': minor
'@eigenpal/docx-editor-react': minor
'@eigenpal/docx-editor-vue': minor
---

Add a supported style-definition API for reading, updating, adding, syncing, and exporting document styles.

The core API now supports headless style updates through `DocumentAgent.updateStyle`, style package helpers, basedOn-aware resolution, docDefaults/table style merging, and `styles.xml` serialization. React and Vue expose declarative `styleOverrides`/`styleDefinitions`, `onStyleDefinitionsChange`, and an imperative `updateStyle` ref method. The collaboration example mirrors style definitions through the same Y.Doc as body/comments so style edits sync and survive local reloads.
