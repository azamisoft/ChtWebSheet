# Cht WebSheet

Cht WebSheet, also called CWS, is a native HTML worksheet application.
Its primary workbook format is an HTML file that contains workbook data,
styles, formulas, drawing data, and document metadata. CWS can read and
write Excel-family file formats, but it is not an Excel wrapper: the
saved CWS workbook is HTML, and the shared application runtime is loaded
separately.

The project is developed by CHT Co., Ltd. (株式会社CHT).

## What It Does

CWS lets users open a workbook in a browser, edit it with a spreadsheet-like
interface, and save the result as a compact HTML workbook. The HTML file
keeps workbook-specific data, while common UI, rendering, menu, and editing
logic live in the shared CWS runtime.

This keeps saved documents small and makes it possible to update the runtime
without rewriting every saved workbook.

## Supported Features

- Native CWS HTML workbook files.
- Import from `.xlsx` and `.xlsm` workbooks.
- Save as CWS HTML; Save As also exposes Excel-family export choices where
  supported, including `.xlsx`, `.xlsm`, and `.xls`.
- Multi-sheet workbooks with sheet tabs, sheet insertion, sheet context menus,
  tab colors, hidden sheets, and sheet selection.
- Excel-like ribbon UI, formula bar, name box, status bar, zoom controls,
  horizontal and vertical scrolling, and workbook view controls.
- Canvas-based cell rendering for backgrounds, text, and borders, with DOM
  overlays for selection, editing, images, shapes, dialogs, and menus.
- Cell selection, multi-selection, range selection, row/column selection,
  keyboard navigation, fill handle, copy/cut/paste, undo/redo history, and
  basic clipboard integration.
- Cell editing, formula-bar editing, in-cell line breaks, rich text editing
  behavior, and common keyboard shortcuts.
- Common cell formatting: font family, font size, bold, italic, underline,
  double underline, text color, fill color, borders, alignment, wrapping,
  merge/unmerge, number formats, row height, and column width.
- Formula recalculation through HyperFormula, including dependency tracking
  and many common spreadsheet functions.
- Hyperlinks, defined names, comments, notes, dropdown lists, data validation,
  phonetic display when phonetic data exists, and basic translation menu hooks.
- Images, pasted images, imported drawing shapes, inserted shapes, shape
  selection, multi-selection, resizing, rotation, ordering, grouping, and
  shape-format tools.
- Ink drawing tools, including pen/highlighter strokes and eraser behavior.
- Data-oriented tools such as sort, filter, text/CSV import flows, remove
  duplicates, data validation dialogs, data queries/connections UI,
  relationships UI, what-if analysis flows, forecast/data-analysis dialogs,
  outline/grouping, and subtotal/consolidation flows.
- View tools such as normal/page-layout/page-break-preview modes, gridline and
  heading toggles, formula bar visibility, zoom, freeze/split panes, workbook
  windows, custom views, sheet views, full-screen mode, and status-bar
  calculation summaries.
- Local persistence while using the web app, so a refresh restores the current
  workbook instead of returning to the upload screen.

Some advanced workbook features are still evolving. VBA/macro execution is not
currently a supported runtime feature. Complex Excel object-model behavior,
full binary `.xls` compatibility, every chart/pivot-table edge case, and every
Excel UI detail should be verified before relying on them in production.

## Use Online

Open the hosted CWS page:

```text
https://chtec.co.jp/cws
```

You can use CWS online without installing the common runtime.

Typical online workflow:

1. Open `https://chtec.co.jp/cws`.
2. Drag a workbook onto the page, or choose `File > Open`.
3. Edit the workbook in the browser.
4. Use `File > Save` to save the current workbook as a CWS HTML workbook.
5. Use `File > Save As` when you need to choose a target file type.

When the app is opened from the web service, CWS saves to the HTML workbook
format by default.

## Use Locally

CWS can also run from saved local HTML workbooks.

Installing the common runtime is optional:

- Without the common runtime, you can still use CWS through the online service.
- With the common runtime installed, saved CWS HTML workbooks can be opened and
  edited offline because the browser loads the shared minified JS/CSS runtime
  from the local system.

Download the common runtime package from:

```text
https://chtec.co.jp/cws/manual/#download
```

The current runtime packages are:

- Windows: `.zip` package containing `install-runtime.ps1`.
- macOS: `.zip` package containing `install.command`.

Install on Windows:

1. Download the Windows runtime zip package.
2. Extract the zip file.
3. Right-click `install-runtime.ps1` and run it with PowerShell.
4. If Windows asks for permission, allow the script to install the runtime.

Install on macOS:

1. Download the macOS runtime zip package.
2. Extract the zip file.
3. Double-click `install.command`.
4. If macOS asks for permission or an administrator password, follow the prompt.

The installed runtime stores shared CWS files such as `websheet-app.js`,
`websheet-standalone.css`, `websheet-standalone.js`, and version metadata.
Saved HTML workbooks contain the workbook data and know how to locate the
runtime.

After the runtime is installed, users can open a saved CWS HTML workbook
directly in the browser. A local launcher HTML can also be used as a blank
starting point for opening supported workbook files from local storage.

Default local runtime locations used by the runtime packages are:

- macOS: `/Library/Application Support/WebSheet/runtime/current` and
  `/Library/Application Support/WebSheet/runtime/<version>`
- Windows: `C:\ProgramData\WebSheet\runtime\current` and
  `C:\ProgramData\WebSheet\runtime\<version>`
- Linux: `$XDG_DATA_HOME/WebSheet/runtime/<version>` or
  `~/.local/share/WebSheet/runtime/<version>`

## Local Development Setup

Requirements:

- Node.js.
- npm.

Install dependencies:

```bash
npm install
```

Start the local development server:

```bash
npm run dev
```

The dev server uses:

```text
http://127.0.0.1:5173/
```

Build the app and install the local standalone runtime used by development:

```bash
npm run build
```

Preview the production build:

```bash
npm run preview
```

Run the end-to-end test script:

```bash
npm run test:e2e
```

## Repository Layout

- `index.html`: application shell and ribbon markup.
- `src/main.js`: workbook model, import/export, rendering, editing, dialogs,
  menus, data tools, drawing tools, and runtime behavior.
- `src/styles.css`: application UI and worksheet styling.
- `docs/cws-html-1.0-draft.md`: draft CWS HTML profile specification for
  AI-readable and AI-editable workbook files.
- `docs/cws-html-application-index.md`: index of standardization application
  materials and recommended review order.
- `docs/cws-html-prior-art-and-positioning.md`: prior-art review and
  positioning against ODF, OOXML, CSVW, and adjacent W3C Community Groups.
- `docs/standardization-application-plan.md`: standardization application
  roadmap and material checklist.
- `docs/cws-html-w3c-cg-form-draft.md`: form-ready W3C Community Group
  proposal draft.
- `docs/cws-html-w3c-cg-submission-copy.md`: final copy-paste text for a W3C
  Community Group proposal.
- `docs/cws-html-w3c-submission-day-runbook.md`: submission-day checklist and
  post-submission follow-up sequence.
- `docs/cws-html-cg-operational-agreement-draft.md`: initial Community Group
  operating rules, scope, status language, and decision policy draft.
- `docs/cws-html-outreach-kit.md`: supporter recruitment, prior-art feedback,
  and launch announcement text.
- `docs/cws-html-publication-checklist.md`: public URL, status wording,
  license, privacy, and launch readiness checklist.
- `docs/cws-html-readiness-audit.md`: current readiness gaps, source-of-truth
  map, and repetition cleanup notes.
- `docs/cws-html-supporter-tracker.md`: private/public supporter tracking for
  W3C Community Group creation.
- `docs/cws-html-initial-issues-and-labels.md`: recommended labels and initial
  public issue texts.
- `.github/ISSUE_TEMPLATE/`: issue forms for CWS HTML prior-art, conformance,
  and AI-editing feedback.
- `public/schema/cws-html-workbook-model-v1.schema.json`: JSON Schema for the
  plain workbook model stored in `script#websheet-model`.
- `public/ai/cws-html-guide-v1.json`: public AI editing guide for CWS HTML.
- `examples/cws-html/`: valid and invalid CWS HTML conformance examples.
- `scripts/validate-cws-html.mjs`: reference validator for CWS HTML profile
  checks.
- `scripts/diff-cws-model.mjs`: helper for comparing workbook models extracted
  from two CWS HTML files.
- `scripts/report-cws-roundtrip.mjs`: helper for generating Markdown
  validation/diff reports for round-trip evidence.
- `scripts/generate-cws-evidence.mjs`: browser automation for real export,
  AI-style JSON edit, standalone reopen/save, validation, and report evidence.
- `scripts/install-standalone-runtime.mjs`: builds and installs the shared
  standalone runtime for local use.
- `scripts/e2e-real-upload.mjs`: browser-based validation script.
- `public/download/`: development download page, runtime version metadata, and
  runtime assets.
- `THIRD_PARTY_NOTICES.md`: third-party dependency license summary.

Validate the CWS HTML draft examples:

```bash
npm run spec:validate
```

Compare workbook models extracted from two CWS HTML files:

```bash
npm run spec:diff -- before.html after.html
```

Generate a round-trip evidence report:

```bash
npm run spec:roundtrip-report -- --out docs/evidence/report.md original.html edited.html resaved.html
```

Generate real export smoke evidence from a running dev server:

```bash
npm run dev
npm run spec:evidence
```

## Main Open-Source Components

The following direct dependencies are used by CWS. Their own licenses continue
to apply when they are used, bundled, or redistributed.

| Component | Version | License | Use |
| --- | ---: | --- | --- |
| `exceljs` | 4.4.0 | MIT | Reading and writing Excel-family workbook structures. |
| `hyperformula` | 3.2.0 | GPL-3.0-only | Formula engine, recalculation, and dependency tracking. |
| `jquery` | 3.7.1 | MIT | DOM events, UI wiring, and legacy-compatible interaction code. |
| `jszip` | 3.10.1 | MIT OR GPL-3.0-or-later | Reading and writing zipped Office Open XML packages. |
| `lucide` | 0.468.0 | ISC | Ribbon, menu, and dialog icons. |
| `@playwright/test` | 1.60.0 | Apache-2.0 | Browser-based development and E2E validation. |
| `ajv` | 8.20.0 | MIT | JSON Schema validation for CWS HTML conformance tooling. |
| `vite` | 8.0.13 | MIT | Local development server and production bundling. |

See `THIRD_PARTY_NOTICES.md` and `package-lock.json` for more detail,
including transitive dependencies.

## License

Cht WebSheet is released under `GPL-3.0-only`. See `LICENSE` for the full
license text.

This license mode was chosen because the bundled formula engine,
HyperFormula, is used under `GPL-3.0-only`. Keeping CWS under the same
GPL-compatible copyleft model makes redistribution obligations clear: users who
receive the software also receive the freedom to study, modify, and share the
corresponding source code under the same license terms.

If a proprietary or closed-source distribution is required, the HyperFormula
license situation must be handled separately, for example by obtaining an
appropriate commercial HyperFormula license or replacing the calculation engine.

## Trademark Note

Excel is a trademark of Microsoft Corporation. Cht WebSheet is an independent
HTML worksheet application that supports Excel-family file formats.
