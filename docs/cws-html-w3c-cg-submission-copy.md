# CWS HTML W3C Community Group Submission Copy

Status: Draft 0.1

Date: 2026-06-12

## Purpose

This document contains concise copy for a W3C Community Group proposal form.
Replace bracketed placeholders before submission. Keep the wording conservative:
this is an incubation proposal, not a claim that CWS HTML is already a W3C
standard.

## Pre-Submission Status Statement

Use this near the proposal repository or announcement:

> CWS HTML 1.0 Draft is a project-level incubation draft for an AI-editable HTML
> workbook profile. It is not a W3C Recommendation or W3C Standard.

## Proposed Group Name

AI-Editable HTML Workbook Community Group

## Short Name

html-workbook

## One-Sentence Summary

The group incubates an HTML workbook profile that keeps a browser-openable
workbook and an AI-editable JSON workbook model in the same HTML file.

## Short Description

The AI-Editable HTML Workbook Community Group incubates an HTML profile for
workbook documents whose structure can be read and edited by automation tools.
The initial contribution is CWS HTML 1.0 Draft, where workbook data is plain JSON
inside `script#websheet-model`, with explicit HTML markers, an AI editing guide,
JSON Schema, validation tests, and optional per-image asset encoding. This group
may publish Specifications.

## Scope

This group will discuss and refine an HTML-based workbook profile that separates
document identity, workbook data, AI editing guidance, and runtime behavior. The
group will focus on deterministic recognition of workbook documents, a stable
editable JSON region, conformance requirements, test suites, security and
privacy guidance, and interoperability between tools that read, write, validate,
or AI-edit workbook-like HTML files.

The group may use CWS HTML 1.0 Draft as its initial technical contribution, but
the group may rename concepts, generalize vendor-specific fields, or split the
work into profiles if that improves interoperability.

## Out of Scope

The group will not standardize Microsoft Excel, OOXML, BIFF, ODF, or Google
Sheets internals; require browser vendors to implement spreadsheet editing;
define a complete spreadsheet formula calculation engine; standardize Cht
WebSheet user interface behavior; or claim W3C endorsement before any formal
W3C process has occurred.

## Initial Deliverables

- AI-editable HTML workbook profile specification.
- Use cases and requirements.
- Prior-art and positioning note.
- Workbook model JSON Schema.
- AI editing guide.
- Conformance checklist.
- Test plan and example corpus.
- Reference validator.
- Implementation and round-trip reports.
- Community Group Report, if the group reaches sufficient agreement.

## Initial Contribution

The initial contribution is the CWS HTML 1.0 Draft package:

- draft specification:
  https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-1.0-draft.md
- explainer:
  https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-explainer.md
- use cases and requirements:
  https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-use-cases-and-requirements.md
- prior-art and positioning note:
  https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-prior-art-and-positioning.md
- JSON Schema:
  https://github.com/azamisoft/ChtWebSheet/blob/main/public/schema/cws-html-workbook-model-v1.schema.json
- AI editing guide:
  https://github.com/azamisoft/ChtWebSheet/blob/main/public/ai/cws-html-guide-v1.json
- examples:
  https://github.com/azamisoft/ChtWebSheet/tree/main/examples/cws-html
- validator:
  https://github.com/azamisoft/ChtWebSheet/blob/main/scripts/validate-cws-html.mjs
- evidence reports:
  https://github.com/azamisoft/ChtWebSheet/tree/main/docs/evidence

## Expected Participants

- Web document format implementers.
- Spreadsheet and workbook application developers.
- AI coding and document editing tool builders.
- Enterprise document workflow teams.
- Standards, security, privacy, accessibility, and internationalization
  reviewers.
- Developers of import/export, conversion, validation, and diff tools.

## Success Criteria

During incubation, success means:

- the problem statement is clear and validated by real workflows;
- at least one implementation writes conforming files;
- at least one reader or validator independent of the writer can inspect files;
- test examples cover common workbook features and failure cases;
- AI edit round-trip reports demonstrate practical editability;
- open issues are tracked publicly;
- naming, licensing, and IPR questions are resolved enough for broader review.

Success does not require immediate transition to the W3C Recommendation Track.

## Proposed Communication Channels

- Public Git repository: https://github.com/azamisoft/ChtWebSheet/
- Public issue tracker: https://github.com/azamisoft/ChtWebSheet/issues
- W3C Community Group discussion channel, if the group is created.
- Periodic public progress reports.

## Initial 90-Day Plan

1. Publish the draft package and invite review.
2. Collect real CWS HTML export examples.
3. Produce at least three round-trip reports.
4. Expand JSON Schema validation and invalid examples.
5. Identify vendor-specific terms that should be generalized.
6. Recruit participants from workbook, Web document, and AI tooling
   communities.
7. Decide whether the group should keep CWS HTML as the profile name or adopt a
   vendor-neutral name.
8. Publish or confirm the group's initial operational agreement.
9. Ask adjacent communities for prior-art and venue feedback if this has not
   already happened.

## Post-Submission Supporter Message

Use this after the W3C proposed group page exists:

Hello,

The proposed AI-Editable HTML Workbook Community Group page is now available:

`[W3C_PROPOSED_GROUP_URL]`

If you are still willing to support creating the discussion venue, please visit
the proposal page and support the group there. Supporting group creation means
supporting a public incubation venue; it does not mean endorsing every technical
detail of CWS HTML 1.0 Draft.

The current draft remains an incubation draft and is not a W3C Recommendation or
W3C Standard.

Thank you.

## Remaining Placeholder

`[W3C_PROPOSED_GROUP_URL]` is intentionally left in the supporter follow-up
message until W3C creates the proposed group page.
