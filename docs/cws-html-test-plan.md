# CWS HTML Test Plan

Status: Draft 0.1

Date: 2026-06-12

## Goals

The CWS HTML test plan demonstrates that the format is recognizable,
AI-editable, and round-trip safe for workbook data.

## Test Layers

1. Static conformance tests.
2. JSON model validation.
3. Reader/writer round-trip tests.
4. AI edit tests.
5. Security and no-script inspection tests.

## Static Conformance Tests

Command:

```bash
npm run spec:validate
```

Current coverage:

- valid minimal CWS HTML;
- valid AI edit before/after pair;
- valid compressed image asset example;
- invalid top-level compressed workbook model;
- invalid missing CWS markers;
- invalid duplicate model scripts;
- invalid JSON;
- invalid cell keys;
- invalid range ordering;
- invalid or missing image assets.

## Required Example Classes

- [ ] Minimal workbook.
- [ ] Multi-sheet workbook.
- [ ] Formula workbook.
- [ ] Merged range workbook.
- [ ] Styled cells workbook.
- [ ] Image asset workbook.
- [ ] Compressed image asset workbook.
- [ ] Filter/table workbook.
- [ ] Comment/note workbook.
- [ ] Unknown-field preservation workbook.

## AI Edit Test Matrix

| Test | Edit | Expected Result |
| --- | --- | --- |
| Cell text | Change `sheets[0].cells["3:2"].raw` | CWS reloads and displays new text. |
| Formula | Change formula and cached value | Formula remains syntactically visible and reloadable. |
| Style | Add fill and bold style | Existing cell data remains unchanged. |
| Merge | Add merge range | Top-left value remains visible. |
| Sheet | Rename sheet | Formula references are handled by the editor or flagged for review. |
| Image metadata | Move image object | Image payload is preserved. |
| Image payload | Preserve `assets.images` | Image displays after reload. |
| Unknown field | Add and preserve extension field | Writer does not delete unknown data. |

## Round-Trip Procedure

1. Open source workbook in Cht WebSheet.
2. Save as CWS HTML.
3. Run `node scripts/validate-cws-html.mjs <file>`.
4. Modify only `script#websheet-model`.
5. Reopen in Cht WebSheet.
6. Save again.
7. Validate again.
8. Compare workbook model before and after for expected changes only.

Useful commands:

```bash
node scripts/validate-cws-html.mjs path/to/original.html
node scripts/validate-cws-html.mjs path/to/edited.html
node scripts/validate-cws-html.mjs path/to/resaved.html
node scripts/diff-cws-model.mjs path/to/original.html path/to/edited.html
node scripts/diff-cws-model.mjs path/to/edited.html path/to/resaved.html
node scripts/report-cws-roundtrip.mjs --out docs/evidence/report.md path/to/original.html path/to/edited.html path/to/resaved.html
```

Automated real export smoke:

```bash
npm run dev
npm run spec:evidence
```

The automated smoke imports an xlsx workbook through the browser UI, saves CWS
HTML, edits only `script#websheet-model`, reopens/resaves the edited file, and
checks validation plus preservation of an unknown extension field.

## Validator Requirements

The reference validator should reject:

- missing root CWS markers;
- missing or duplicate `script#websheet-model`;
- invalid JSON;
- JSON Schema violations;
- top-level encoded workbook wrappers;
- missing required workbook fields;
- invalid sheet names;
- invalid cell keys;
- invalid range ordering;
- invalid image asset references;
- unsupported image asset encodings.

The reference validator should warn, not fail, for:

- cells outside declared `rowCount` or `colCount`;
- missing recommended AI instruction script;
- optional fields not recognized by the draft.

## Future Automation

- Add a test fixture generator for larger workbooks.
- Add browser-based reload tests for image assets with `dataEncoding`.
- Add a test that verifies no scripts are executed during validation.
