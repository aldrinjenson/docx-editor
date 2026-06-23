---
'@eigenpal/docx-editor-react': minor
---

Add a `toolbarAlignment` prop to `<DocxEditor>` (`'start' | 'center' | 'end'`, default `'start'`) that sets the horizontal alignment of the formatting toolbar's contents. The underlying `Toolbar` / `EditorToolbar` gains the same `toolbarAlignment` prop. `'center'` and `'end'` use safe alignment, so the leading controls stay scroll-reachable when the toolbar overflows a narrow viewport. The default preserves the existing left-aligned layout, so current callers are unaffected.
