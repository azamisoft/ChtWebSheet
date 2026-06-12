# CWS HTML Outreach Kit

Status: Draft 0.1

Date: 2026-06-12

## Purpose

This kit contains short public wording for prior-art review, supporter
recruitment, and a possible W3C Community Group proposal. It is intentionally
conservative: the work should be presented as incubation, not as an approved
standard.

## One-Line Positioning

CWS HTML is an HTML workbook profile that keeps the editable workbook model as
plain JSON inside the HTML file so AI tools can make targeted spreadsheet edits
without rewriting runtime code or visual markup.

## Short Public Description

The proposed AI-Editable HTML Workbook Community Group would incubate an HTML
profile for workbook documents that are browser-openable and automation-editable.
The initial contribution, CWS HTML 1.0 Draft, stores workbook structure as plain
JSON in `script#websheet-model`, identifies the file with explicit HTML markers,
links to a JSON Schema and AI guide, and includes validators and round-trip
evidence.

## Longer Public Description

Spreadsheet applications can export HTML, but those exports are usually designed
for display rather than reliable workbook editing. AI agents and automation
tools need a predictable place to read and modify workbook structure while
preserving formulas, styles, merged cells, filters, images, metadata, and unknown
extension fields.

CWS HTML proposes a narrow HTML profile: the file remains ordinary HTML, but the
editable workbook model is stored as deterministic JSON in
`script#websheet-model`. Runtime code, styles, and presentation markup remain
separate from the model. Image payloads may be encoded per asset, while workbook
structure remains readable JSON. The draft includes schema markers, AI editing
instructions, validation tooling, valid and invalid examples, and round-trip
evidence.

The proposal does not replace ODF, OOXML, CSVW, Google Sheets, or browser-native
HTML tables. It targets a smaller interoperability gap: browser-openable
workbook files that AI tools can inspect, edit, validate, and round-trip.

## W3C Form Text

Use `docs/cws-html-w3c-cg-submission-copy.md` as the canonical W3C form text.
This outreach kit keeps only public messages and review requests so the same
proposal language is not maintained in two places.

## Prior-Art Feedback Request

Subject:

Prior-art feedback request: AI-editable HTML workbook profile

Body:

Hello,

I am preparing an incubation proposal for an AI-editable HTML workbook profile.
The draft, currently called CWS HTML 1.0, keeps a browser-openable workbook and
an editable JSON workbook model in the same HTML file. The model is stored in
`script#websheet-model`, with explicit profile markers, JSON Schema, AI editing
guidance, validation tooling, examples, and round-trip evidence.

Before proposing a new W3C Community Group, I would like feedback on whether
this overlaps with existing work or should be discussed in an existing venue.
The proposal is not intended to replace ODF, OOXML, CSVW, Google Sheets, or
HTML tables. The narrower goal is a saved HTML workbook profile that AI tools
can inspect, edit, validate, and round-trip.

I would especially appreciate comments on:

- whether this belongs in a new Community Group or an existing forum;
- related prior art we should acknowledge;
- whether the scope is narrow enough for incubation;
- conformance or security concerns that should be addressed first.

Draft materials:

- Explainer:
  https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-explainer.md
- Draft specification:
  https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-1.0-draft.md
- Prior-art note:
  https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-prior-art-and-positioning.md
- Schema:
  https://github.com/azamisoft/ChtWebSheet/blob/main/public/schema/cws-html-workbook-model-v1.schema.json
- Examples:
  https://github.com/azamisoft/ChtWebSheet/tree/main/examples/cws-html
- Validator:
  https://github.com/azamisoft/ChtWebSheet/blob/main/scripts/validate-cws-html.mjs
- Round-trip evidence:
  https://github.com/azamisoft/ChtWebSheet/tree/main/docs/evidence

Thank you.

## Supporter Recruitment Request

Subject:

Support request: proposed AI-Editable HTML Workbook Community Group

Body:

Hello,

I am seeking initial supporters for a proposed W3C Community Group on
AI-editable HTML workbook documents. W3C Community Group creation requires five
individual supporters. Supporting the proposal helps create the discussion venue;
it does not mean you endorse every technical detail of the current draft.

The initial contribution is CWS HTML 1.0 Draft, an HTML profile where workbook
data is stored as plain JSON in `script#websheet-model` with schema markers,
AI editing guidance, examples, validator tooling, and round-trip evidence.

The proposed scope is intentionally narrow:

- browser-openable HTML workbook files;
- deterministic AI-editable workbook JSON;
- validation and conformance tests;
- security/privacy guidance for inspection and automation;
- interoperability between readers, writers, validators, and AI editing tools.

Out of scope are replacing ODF or OOXML, standardizing spreadsheet application
internals, requiring browser-native spreadsheet editing, or claiming that the
draft is already a W3C standard.

Proposal draft:

https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-w3c-cg-submission-copy.md

If you are willing to support the creation of the discussion group, please let
me know what name and affiliation, if any, you would like listed.

## Launch Announcement Draft

Subject:

Proposed Community Group: AI-editable HTML workbook profile

Body:

I have published a draft incubation package for an AI-editable HTML workbook
profile. The initial draft, CWS HTML 1.0, defines a browser-openable HTML file
where the editable workbook model is plain JSON in `script#websheet-model`.

The package includes:

- draft specification;
- explainer and use cases;
- prior-art and positioning note;
- JSON Schema and AI guide;
- valid and invalid example files;
- reference validator;
- AI edit round-trip evidence.

The proposal is deliberately scoped as an incubation effort. It does not replace
ODF, OOXML, CSVW, Google Sheets, or browser-native HTML tables. It focuses on a
specific interoperability gap: saved HTML workbook files that AI tools can read,
edit, validate, and round-trip.

Feedback is welcome on scope, prior art, conformance, security, privacy,
accessibility, internationalization, and whether this should become a new W3C
Community Group or be discussed first in an existing venue.

Links:

- Proposal:
  https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-w3c-cg-submission-copy.md
- Explainer:
  https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-explainer.md
- Draft spec:
  https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-1.0-draft.md
- Prior-art note:
  https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-prior-art-and-positioning.md
- Examples:
  https://github.com/azamisoft/ChtWebSheet/tree/main/examples/cws-html
- Validator:
  https://github.com/azamisoft/ChtWebSheet/blob/main/scripts/validate-cws-html.mjs

## Words To Use

- Community Group incubation
- draft profile
- HTML workbook profile
- AI-editable workbook JSON
- browser-openable file
- validation and round-trip evidence
- complements existing office document formats

## Words To Avoid

- W3C standard
- W3C-approved
- official HTML spreadsheet format
- Excel replacement
- OOXML replacement
- browser-native spreadsheet
- universal spreadsheet engine

## Related Preparation Docs

- `docs/cws-html-w3c-cg-submission-copy.md`
- `docs/cws-html-w3c-submission-day-runbook.md`
- `docs/cws-html-publication-checklist.md`
- `docs/cws-html-supporter-tracker.md`
- `docs/cws-html-initial-issues-and-labels.md`

## References

- W3C Community and Business Groups: https://www.w3.org/community/
- W3C Community and Business Group Process:
  https://www.w3.org/community/about/process/
- W3C Community Groups FAQ: https://www.w3.org/community/about/faq/
