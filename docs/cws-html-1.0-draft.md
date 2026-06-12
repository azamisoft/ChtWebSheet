# CWS HTML 1.0 Draft Specification

Status: Draft 0.1

Last updated: 2026-06-12

## Abstract

CWS HTML is an HTML-based workbook format for Cht WebSheet. A conforming CWS
HTML file is both a user-openable HTML document and an AI-editable workbook
package. The actual workbook model is stored as plain JSON in
`script#websheet-model`, while the HTML shell identifies the document as CWS and
points automated agents to editing guidance.

This specification defines the CWS HTML 1.0 file profile, its identification
markers, the editable JSON region, the AI instruction region, and the minimum
workbook model requirements needed for interoperable readers, writers, and
validators.

## Goals

- Make CWS HTML files recognizable without relying on product text.
- Make the editable workbook data discoverable at a stable HTML location.
- Keep workbook content AI-readable and AI-editable by default.
- Preserve compatibility with ordinary HTML tools and browsers.
- Allow CWS runtime code to evolve independently from saved workbook data.
- Preserve unknown fields so future CWS versions and imported workbook metadata
  are not lost.

## Non-Goals

- This specification does not define a complete spreadsheet calculation engine.
- This specification does not require non-CWS spreadsheet applications to render
  every CWS feature.
- This specification does not standardize Excel, OOXML, BIFF, or Google Sheets
  export formats.
- This specification does not permit automated agents to rewrite runtime code
  when a user asks only to edit workbook content.

## Conformance Language

The words MUST, MUST NOT, SHOULD, SHOULD NOT, and MAY are to be interpreted as
normative requirements for this draft.

## File Type

A CWS HTML file MUST be a valid HTML document. The default media type SHOULD
remain `text/html`.

CWS HTML 1.0 is an HTML profile. A separate registered media type is not required
for normal web delivery because browsers already process the file as HTML. If a
future registration is needed for interchange systems that cannot use `text/html`
with profile metadata, the registration request SHOULD reference this
specification and the CWS HTML identification markers.

## Versioning

CWS HTML has two separate version concepts:

- `data-cws-schema-version` and `meta[name="cws:schema-version"]` identify the
  CWS HTML shell/profile version. This draft defines version `1`.
- `websheet-model.version` identifies the internal CWS workbook model version.
  Writers MUST preserve it unless they intentionally migrate the workbook model.

Readers MUST NOT assume that these two versions are the same.

## Required HTML Markers

A conforming CWS HTML 1.0 document MUST contain all of the following markers:

```html
<html data-cws-format="CWS_HTML" data-cws-schema-version="1" data-cws-app="CWS">
```

```html
<meta name="cws:format" content="CWS_HTML" />
<meta name="cws:schema-version" content="1" />
<meta name="cws:guide" content="https://chtec.co.jp/cws/ai/cws-html-guide-v1.json" />
<meta name="cws:schema" content="https://chtec.co.jp/cws/schema/cws-html-workbook-model-v1.schema.json" />
```

The document MUST contain exactly one editable workbook model script:

```html
<script
  id="websheet-model"
  type="application/json"
  data-cws-model="workbook"
  data-cws-format="CWS_HTML"
  data-cws-schema-version="1"
  data-cws-editable="true"
>{...}</script>
```

The document SHOULD contain an AI instruction script:

```html
<script id="cws-ai-instructions" type="application/json">{...}</script>
```

## Recognition Rules

A reader or automated agent SHOULD treat an HTML file as high-confidence CWS HTML
when:

- `script#websheet-model[type="application/json"]` exists, and
- either `html[data-cws-format="CWS_HTML"]` or
  `meta[name="cws:format"][content="CWS_HTML"]` exists.

A reader MAY treat a document as medium-confidence CWS HTML when
`script#websheet-model` and `script#cws-ai-instructions` exist but one of the
top-level markers is missing.

A reader MUST NOT edit a document as CWS HTML based only on product text such as
`Cht WebSheet` or `WebSheet`.

## Editable Region

For workbook-content changes, automated agents MUST edit only the JSON text
inside `script#websheet-model`.

Automated agents MUST NOT modify runtime JavaScript, CSS, loader code, external
links, script attributes, or meta tags unless the user explicitly asks to change
those parts of the document.

Writers MUST escape JSON for safe embedding in HTML script text. At minimum, the
character `<` MUST be escaped as `\u003c`, and U+2028/U+2029 SHOULD be escaped.

## AI-Editable Profile Requirement

CWS HTML 1.0 requires `script#websheet-model` to contain the plain workbook
object as JSON. Encoded wrappers for the whole workbook, such as a top-level
gzip/base64 payload, are not conforming to the AI-editable profile because the
workbook structure is not directly readable or editable by generic agents.

Readers MAY support legacy top-level encoded payloads for backward
compatibility, but writers SHOULD save CWS HTML 1.0 files with a plain model
object.

Image binary payloads MAY be encoded independently when workbook structure and
image metadata remain plain JSON. For example, `assets.images[id].data` MAY use
`dataEncoding: "gzip-base64"` to indicate that the image bytes are gzip
compressed and then base64 encoded. This does not break the AI-editable profile
because sheets, cells, formulas, styles, object positions, image dimensions, and
asset references remain directly readable and editable.

## Workbook Model

The JSON inside `script#websheet-model` MUST be an object.

The workbook model MUST contain:

- `version`: internal CWS workbook model version.
- `sourceName`: display or source file name.
- `sheets`: non-empty ordered array of worksheet objects.

The workbook model SHOULD contain:

- `saveType`: normally `html` for CWS HTML files.
- `createdAt`: ISO 8601 creation timestamp.
- `activeSheetIndex`: zero-based active sheet index.
- `definedNames`: workbook-level defined names.
- `workbookView` and window-view state when present.
- `assets.images` when large image payloads are shared or encoded separately
  from sheet object placement metadata.

Unknown workbook-level fields MUST be preserved.

## Image Assets

Images SHOULD keep their workbook-facing metadata in plain JSON. A sheet image
object MAY use `assetId` to refer to a shared entry in `workbook.assets.images`.

An image asset SHOULD contain:

- `mimeType`: image media type, such as `image/png` or `image/webp`.
- `extension`: preferred file extension.
- `width` and `height`: intrinsic or optimized pixel dimensions.
- `data`: base64 image bytes, or an encoded payload when `dataEncoding` is set.

When `dataEncoding` is absent, `data` MUST be base64 image bytes. When
`dataEncoding` is `gzip-base64`, `data` MUST be base64-encoded gzip bytes whose
decompressed result is the original image byte sequence.

Automated agents that edit cells, formulas, styles, rows, columns, sheet names,
or layout SHOULD preserve image asset payloads exactly unless the user asks to
replace or compress images.

## Worksheet Model

Each sheet object MUST contain:

- `name`: sheet tab name.
- `rowCount`: positive integer row count.
- `colCount`: positive integer column count.
- `cells`: object map keyed by `row:col`.

Each sheet SHOULD contain:

- `columns`: array of column records.
- `rows`: array of row records.
- `merges`: array of merged ranges.
- `images`, `shapes`, `drawingObjects`, or `inkStrokes` when visual objects are
  present.
- `autoFilter`, `tables`, `validations`, `pageLayout`, and view fields when
  those workbook features are present.

Unknown sheet-level fields MUST be preserved.

## Cells

Cell map keys MUST use 1-based `row:col` coordinates. For example, A1 is `1:1`
and B3 is `3:2`.

A cell object SHOULD contain `row` and `col` fields matching its map key.

Cell values SHOULD use:

- `raw` for literal user-entered values.
- `formula` for formulas beginning with `=`.
- `cached` for cached formula results.
- `display` for rendered display text.
- `css` for visual style.
- `richText` for partial text styling.
- `borders` for borders.
- `hyperlink`, `note`, `comment`, and `validation` for attached metadata.

Unknown cell-level fields MUST be preserved.

## Ranges and Merges

Ranges MUST use 1-based inclusive coordinates:

```json
{ "top": 1, "left": 1, "bottom": 3, "right": 4 }
```

For merged ranges, the visible value and primary style SHOULD remain in the
top-left cell unless a user explicitly requests another structure.

## AI Instruction Script

`script#cws-ai-instructions` SHOULD contain a JSON object with:

- `notice`: short human-readable editing notice.
- `format`: `CWS_HTML`.
- `schemaVersion`: `1`.
- `modelScriptId`: `websheet-model`.
- `editableRegion`: `script#websheet-model`.
- `guideUrl`: URL of the public AI guide.
- `schemaUrl`: URL of the public workbook model JSON Schema.
- `rules`: array of short editing rules.

Agents SHOULD prefer the latest guide URL declared in the file over hard-coded
knowledge.

## Security and Privacy

CWS HTML files may contain formulas, external links, comments, image data URLs,
and imported workbook metadata. Tools MUST NOT fetch external links or execute
runtime code merely to inspect or edit `script#websheet-model`.

Validators and AI agents SHOULD parse the HTML as data and extract the model
script without executing scripts.

Writers SHOULD avoid embedding secrets in workbook metadata and SHOULD preserve
existing privacy-sensitive fields unless the user requests their removal.

## Validation

A CWS HTML 1.0 validator SHOULD check:

- required HTML markers;
- exactly one `script#websheet-model`;
- valid JSON in the model script;
- plain object model, not an encoded wrapper;
- required workbook and sheet fields;
- valid `row:col` cell keys;
- valid range ordering;
- duplicate sheet names;
- conformance to the JSON Schema published at
  `public/schema/cws-html-workbook-model-v1.schema.json`.

## Reference Artifacts

The initial reference package contains:

- Specification draft: `docs/cws-html-1.0-draft.md`
- Standardization plan: `docs/standardization-application-plan.md`
- W3C proposal starter: `docs/w3c-community-group-proposal.md`
- JSON Schema: `public/schema/cws-html-workbook-model-v1.schema.json`
- Valid examples: `examples/cws-html/valid/`
- Invalid examples: `examples/cws-html/invalid/`
- Validator: `scripts/validate-cws-html.mjs`

## Change Control

Until a standards body adopts the work, changes to this draft are controlled by
the Cht WebSheet project maintainers. Breaking changes SHOULD require a new CWS
HTML schema version.
