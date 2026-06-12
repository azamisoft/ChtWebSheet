# CWS HTML Implementation Report

Status: Draft 0.1

Date: 2026-06-12

## Summary

This report records the implementation status for CWS HTML 1.0 draft
conformance in the Cht WebSheet project.

## Implementations

| Implementation | Type | Status | Notes |
| --- | --- | --- | --- |
| Cht WebSheet writer | Writer | Implemented draft support | Saves CWS HTML markers, AI instructions, schema URL, and plain `script#websheet-model` workbook JSON. |
| Cht WebSheet reader | Reader | Implemented draft support | Reads `script#websheet-model`, supports plain workbook models, and can decode image assets with `dataEncoding: "gzip-base64"`. |
| CWS standalone runtime | Reader/runtime | Implemented draft support | Opens saved HTML workbooks and materializes `assets.images` into renderable image sources. |
| Reference validator | Validator | Implemented draft support | Checks required markers, JSON validity, JSON Schema conformance, workbook fields, sheet/cell/range invariants, image assets, and selected invalid cases. |
| Model diff helper | Test utility | Implemented draft support | Extracts `script#websheet-model` from two CWS HTML files and reports structural differences with large payload summaries. |
| Round-trip report helper | Test utility | Implemented draft support | Runs validation and model diffs for original, edited, and resaved CWS HTML files and writes a Markdown report. |
| Evidence generator | Test utility | Implemented draft support | Opens a workbook in the browser UI, saves CWS HTML, applies an AI-style JSON edit, reopens/resaves, validates, and writes report artifacts. |
| JSON Schema | Schema | Draft and executable | Describes the workbook model shape for `script#websheet-model` and is executed by the reference validator through Ajv. |
| AI guide | Agent guidance | Draft | Describes recognition and editing rules for AI agents. |

## Feature Coverage

| Feature | Writer | Reader | Validator | Evidence |
| --- | --- | --- | --- | --- |
| CWS HTML root markers | Yes | Yes | Yes | Valid examples pass `npm run spec:validate`. |
| `meta name="cws:*"` markers | Yes | Yes | Yes | Valid examples pass `npm run spec:validate`. |
| Plain workbook JSON in `script#websheet-model` | Yes | Yes | Yes | Top-level compressed model invalid example is rejected. |
| AI instructions script | Yes | Informational | Warning | Validator warns when missing. |
| Schema URL | Yes | Informational | Yes | Required by validator. |
| Workbook required fields | Yes | Yes | Yes | Validator rejects missing model essentials through schema and CWS-specific checks. |
| Sheet and cell coordinates | Yes | Yes | Yes | Invalid cell key example is rejected. |
| Merge/range ordering | Yes | Yes | Yes | Invalid merge example is rejected. |
| Image asset references | Yes | Yes | Yes | Valid compressed image asset example passes; missing asset example is rejected. |
| Per-image `gzip-base64` payloads | Conditional | Yes | Yes | Encoded image asset example passes. |
| Top-level workbook compression | No for draft writer | Legacy read only | Rejected | `compressed-model.cws.html` is intentionally invalid. |
| Unknown field preservation | Intended | Intended | Scripted smoke | `npm run spec:evidence` adds and checks a top-level extension field across standalone reopen/save. |

## Current Conformance Command

```bash
npm run spec:validate
```

## Current Example Coverage

Valid examples:

- `examples/cws-html/valid/minimal.cws.html`
- `examples/cws-html/valid/ai-edit-before.cws.html`
- `examples/cws-html/valid/ai-edit-after.cws.html`
- `examples/cws-html/valid/image-asset-gzip.cws.html`

Invalid examples:

- `examples/cws-html/invalid/compressed-model.cws.html`
- `examples/cws-html/invalid/missing-cws-marker.html`
- `examples/cws-html/invalid/duplicate-model-script.cws.html`
- `examples/cws-html/invalid/invalid-json.cws.html`
- `examples/cws-html/invalid/invalid-cell-key.cws.html`
- `examples/cws-html/invalid/invalid-merge-range.cws.html`
- `examples/cws-html/invalid/missing-image-asset.cws.html`
- `examples/cws-html/invalid/invalid-image-asset-encoding.cws.html`

## Round-Trip Support Tooling

Model diff command:

```bash
node scripts/diff-cws-model.mjs before.html after.html
```

The diff helper is intended for round-trip evidence. It compares workbook model
structure without executing the HTML file and summarizes long image/base64
payloads by length and hash.

Round-trip report command:

```bash
node scripts/report-cws-roundtrip.mjs --out docs/evidence/report.md original.html edited.html resaved.html
```

Real export smoke evidence command:

```bash
npm run dev
npm run spec:evidence
```

The evidence generator writes `exported.cws.html`, `edited.cws.html`,
`resaved.cws.html`, `roundtrip-report.md`, and `summary.json` under
`docs/evidence/generated/real-export-smoke/` by default.

Current generated smoke evidence:

- `docs/evidence/generated/real-export-smoke/summary.json`
- `docs/evidence/generated/real-export-smoke/roundtrip-report.md`
- `docs/evidence/generated/real-export-smoke/exported.cws.html`
- `docs/evidence/generated/real-export-smoke/edited.cws.html`
- `docs/evidence/generated/real-export-smoke/resaved.cws.html`

## Known Gaps

- Independent third-party reader/writer implementation is not yet available.
- Round-trip evidence from real business workbooks still needs to be collected.
- Browser reload tests for compressed image assets should be added to e2e.
- Public repository exists; exact public review URLs for this draft package are
  pending commit/push and placeholder replacement.
- License and contribution rights need final confirmation before public
  submission.
- Accessibility, internationalization, security, and privacy review feedback has
  not yet been collected.

## Evidence Needed Before External Submission

- Real exported CWS HTML files that pass validation.
- Round-trip reports using `docs/cws-html-round-trip-report-template.md`.
- Exact public URLs for the specification, schema, guide, examples, and
  validator after the draft package is published.
- A final license decision for specification text and conformance assets.
- At least one external reviewer or independent tool author willing to evaluate
  the profile.
