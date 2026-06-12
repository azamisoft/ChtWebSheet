# CWS HTML Explainer

Status: Draft 0.1

Date: 2026-06-12

## What Is CWS HTML?

CWS HTML is an HTML document profile for workbook files. It lets the same file
act as:

- a browser-openable workbook document;
- a structured workbook data package;
- an AI-editable artifact with clear editing boundaries.

The core mechanism is simple: workbook data is stored as JSON inside
`script#websheet-model`.

## Why Not Ordinary Spreadsheet HTML Export?

Many spreadsheet tools can export HTML, but that HTML is usually optimized for
viewing. Workbook concepts are flattened into tables and styles. Editing the
workbook model from that output is difficult because the file does not reliably
expose formulas, sheet metadata, merged ranges, filters, comments, image
placement, or workbook-level settings in a stable structure.

CWS HTML keeps the workbook model in a predictable JSON location and treats the
HTML runtime as separate from user data.

## Key Design Choices

1. Use HTML as the container.

   This keeps files easy to open, host, attach, and inspect with common tools.

2. Use explicit markers.

   CWS HTML files are recognized through `data-cws-*`, `meta name="cws:*"`, and
   `script#websheet-model`, not by product text or file name.

3. Keep structure plain.

   Sheet names, cells, formulas, styles, ranges, image placement metadata, and
   references remain directly visible in JSON.

4. Allow binary payload optimization.

   Image bytes may live in `assets.images` and may use per-asset encoding such
   as `dataEncoding: "gzip-base64"`. This keeps large binary payloads away from
   sheet structure without hiding workbook semantics.

5. Preserve unknown fields.

   Readers and AI agents must preserve fields they do not understand so imported
   workbook metadata and future CWS extensions are not lost.

## Example

```html
<html data-cws-format="CWS_HTML" data-cws-schema-version="1" data-cws-app="CWS">
<head>
  <meta name="cws:format" content="CWS_HTML" />
  <meta name="cws:schema-version" content="1" />
  <meta name="cws:guide" content="https://chtec.co.jp/cws/ai/cws-html-guide-v1.json" />
</head>
<body>
  <script id="websheet-model" type="application/json" data-cws-model="workbook">
  {
    "version": 8,
    "sourceName": "book.html",
    "sheets": [
      {
        "name": "Sheet1",
        "rowCount": 1,
        "colCount": 1,
        "cells": {
          "1:1": { "row": 1, "col": 1, "raw": "Hello", "display": "Hello" }
        }
      }
    ]
  }
  </script>
</body>
</html>
```

## Primary Use Cases

- AI edits a design workbook by changing cells and styles.
- A browser app saves an offline editable workbook as a single HTML file.
- A document workflow system validates workbook structure without executing
  scripts.
- A reviewer inspects a structured diff of workbook changes.
- A tool extracts sheet data, formulas, and metadata for conversion or audit.

## Non-Goals

- Replace Excel, OOXML, BIFF, or Google Sheets formats.
- Standardize Cht WebSheet UI behavior.
- Require browser vendors to implement spreadsheet editing.
- Define a complete formula engine.

## Privacy and Security Model

Tools can inspect and edit the workbook model without running page scripts.
Automated agents should not fetch external links or rewrite runtime code unless
the user specifically asks them to.

## Success Criteria

- A saved CWS HTML file is recognizable by deterministic markers.
- Workbook structure is readable from `script#websheet-model`.
- AI agents can make targeted edits to cells, formulas, styles, and ranges.
- CWS can reopen AI-edited files.
- Unknown fields and image payloads are preserved.
- A validator can distinguish conforming and non-conforming files.
