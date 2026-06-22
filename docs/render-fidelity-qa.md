# Render Fidelity QA

Use this checklist when a DOCX renders differently from Word or another known-good renderer. Keep committed regressions synthetic and sanitized; local reference documents can guide the investigation, but must not be committed, uploaded, or named in PRs unless they are explicitly approved fixtures.

## Regression Loop

1. Reproduce the issue in a local editor build with `?e2e=1` enabled.
2. Compare the local render against a trusted local baseline, such as Word, LibreOffice, or an approved screenshot.
3. Identify the smallest OOXML construct behind the mismatch: section geometry, columns, anchored objects, text boxes, unsupported media formats, table width, header/footer placement, or field output.
4. Add or update synthetic unit/e2e fixtures that preserve the problematic structure without private content.
5. Run focused tests first, then typecheck and the closest Playwright regression when the issue involves painted DOM.
6. Repeat the local smoke pass against the original document only on the developer machine.

## Useful Signals

- Page count and scroll-to-page behavior should match the expected pagination window. Virtualized pages can be empty until scrolled into view, so sample pages through the e2e hook instead of reading only the initial DOM.
- Broken image icons usually mean the parser created an image with no resolved relationship, or the painter assigned a browser-unsupported source. Preserve the document model for round-trip, but avoid painting undecodable browser formats as visible broken icons.
- Text boxes should pass relationships and media into their nested paragraph parser. Missing media in nested content often shows up as empty picture boxes.
- Header and footer navigation bands may be composed from `mc:AlternateContent`, WPS shapes, theme-colored rectangles, and cached diagram drawing parts. Verify repeated labels, visual strip backgrounds, anchored vertical offsets, footer text, and page fields together.
- Multi-column continuous sections need fragment-local x/y coordinates. Paragraphs, tables, and floating content should be placed relative to the fragment column rather than the physical page origin.

## Validation Commands

```bash
bun test packages/core/src/docx packages/core/src/layout-painter
bun run --filter '@eigenpal/docx-editor-core' typecheck
bun run test:e2e -- e2e/tests/generic-rendering-regression.spec.ts
```

Narrow the test paths while iterating, then expand validation before opening a PR.
