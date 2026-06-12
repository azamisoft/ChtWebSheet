# CWS HTML Security and Privacy Considerations

Status: Draft 0.1

Date: 2026-06-12

## Summary

CWS HTML is an HTML document. It may contain executable runtime references, user
workbook data, formulas, comments, links, images, and imported metadata. Tools
that only inspect or edit workbook structure should operate on the HTML as data
and should not execute page scripts.

## Security Principles

- Validators should parse static HTML and JSON only.
- AI agents should edit only `script#websheet-model` for workbook-content
  changes.
- Runtime JavaScript, CSS, loader code, script attributes, meta tags, and
  external links should be left unchanged unless explicitly requested.
- External resources should not be fetched during structural validation.
- Unknown fields should be preserved but not executed.

## Potential Risks

| Risk | Mitigation |
| --- | --- |
| Script execution during inspection | Use static parsing; do not load the file in an executing browser context for validation. |
| External link leakage | Do not fetch links, image URLs, or runtime URLs during validation. |
| Formula injection | Treat formulas as workbook data; consumers that export to other formats should apply their own formula safety policy. |
| Hidden metadata | Preserve metadata by default, but provide user-controlled removal tools for privacy workflows. |
| Oversized image payloads | Allow asset deduplication and per-asset encoding; validators may enforce local size limits. |
| AI over-editing | Provide explicit AI instructions and validate that only intended JSON regions changed. |

## Privacy Considerations

CWS HTML may contain:

- personal data in cells;
- comments and notes;
- file names and timestamps;
- imported workbook properties;
- image data;
- hyperlinks and external references;
- hidden sheets, rows, columns, or formulas.

Tools should not assume that hidden workbook data is safe to disclose. Sharing a
CWS HTML file shares the workbook model and embedded assets in that file.

## Validation Guidance

A validator may reject or warn on:

- unexpectedly large `script#websheet-model` content;
- image assets above a configured size limit;
- external image or runtime URLs in environments that require self-contained
  documents;
- top-level encoded workbook wrappers;
- invalid JSON or malformed ranges.

## AI Editing Guidance

AI agents should:

- preserve unknown fields;
- preserve image payload data unless replacing images;
- avoid rewriting generated runtime code;
- keep JSON valid;
- report when an edit requires formula-reference updates that cannot be safely
  inferred.
