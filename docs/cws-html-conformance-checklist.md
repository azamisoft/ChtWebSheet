# CWS HTML Conformance Checklist

Status: Draft 0.1

Date: 2026-06-12

Use this checklist when reviewing a CWS HTML writer, reader, validator, or test
file.

## Document Identity

- [ ] The file is a valid HTML document.
- [ ] The root `html` element has `data-cws-format="CWS_HTML"`.
- [ ] The root `html` element has `data-cws-schema-version="1"`.
- [ ] The root `html` element has `data-cws-app="CWS"`.
- [ ] The document contains `meta[name="cws:format"][content="CWS_HTML"]`.
- [ ] The document contains `meta[name="cws:schema-version"][content="1"]`.
- [ ] The document contains `meta[name="cws:guide"]`.
- [ ] The document contains `meta[name="cws:schema"]`.

## Workbook Model Script

- [ ] Exactly one `script#websheet-model` exists.
- [ ] `script#websheet-model` has `type="application/json"`.
- [ ] `script#websheet-model` has `data-cws-model="workbook"`.
- [ ] `script#websheet-model` has `data-cws-format="CWS_HTML"`.
- [ ] `script#websheet-model` has `data-cws-schema-version="1"`.
- [ ] `script#websheet-model` has `data-cws-editable="true"`.
- [ ] The script text parses as JSON.
- [ ] The top-level JSON value is the workbook object, not a top-level encoded
  gzip/base64 wrapper.

## Workbook Model

- [ ] `version` exists.
- [ ] `sourceName` is a non-empty string.
- [ ] `sheets` is a non-empty array.
- [ ] `activeSheetIndex`, when present, points to an existing sheet.
- [ ] Unknown workbook-level fields are preserved during read/write.

## Sheets

- [ ] Each sheet has a non-empty `name`.
- [ ] Sheet names are unique.
- [ ] Sheet names do not contain `: \ / ? * [ ]`.
- [ ] Each sheet has positive integer `rowCount` and `colCount`.
- [ ] Each sheet has a `cells` object.
- [ ] Unknown sheet-level fields are preserved during read/write.

## Cells

- [ ] Cell keys use 1-based `row:col` syntax.
- [ ] `cell.row` and `cell.col`, when present, match the cell key.
- [ ] Formula strings begin with `=`.
- [ ] Literal values use `raw`.
- [ ] Display strings use `display`.
- [ ] Styles use `css`, rich text uses `richText`, and borders use `borders`.
- [ ] Unknown cell-level fields are preserved during read/write.

## Ranges

- [ ] Ranges use `top`, `left`, `bottom`, and `right`.
- [ ] Range coordinates are positive integers.
- [ ] `top <= bottom` and `left <= right`.
- [ ] Merge ranges keep the visible value in the top-left cell unless the user
  explicitly requests another layout.

## Images and Assets

- [ ] Sheet image metadata remains plain JSON.
- [ ] A sheet image may use either `src` or `assetId`.
- [ ] `assetId`, when present, resolves to `workbook.assets.images`.
- [ ] Image asset `data`, when `dataEncoding` is absent, is base64 image bytes.
- [ ] Image asset `data`, when `dataEncoding` is `gzip-base64`, decompresses to
  image bytes.
- [ ] AI or validation tools preserve image payload data unless the user asks to
  replace or remove the image.

## AI Instructions

- [ ] The document contains one `script#cws-ai-instructions`, or documents why it
  is omitted.
- [ ] The AI instruction JSON declares `format`, `schemaVersion`,
  `editableRegion`, `guideUrl`, and `schemaUrl`.
- [ ] Editing rules instruct agents to edit only `script#websheet-model` for
  workbook-content changes.

## Security

- [ ] Validators inspect the file without executing scripts.
- [ ] Agents do not fetch external links during structural edits.
- [ ] Runtime JavaScript, CSS, loader code, script attributes, and meta tags are
  not changed for workbook-content edits.
