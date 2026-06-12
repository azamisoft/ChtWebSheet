# W3C Community Group Proposal Starter

Status: Draft 0.2

Date: 2026-06-12

## Purpose

This is the narrative proposal starter for explaining why a Community Group is a
reasonable incubation venue. It is not the canonical form text. Use
`docs/cws-html-w3c-cg-submission-copy.md` for the copy-paste W3C proposal.

## Proposed Group

Name:

AI-Editable HTML Workbook Community Group

Short name:

`html-workbook`

## Process Fit

W3C Community Groups are a reasonable first venue because CWS HTML is
Web-adjacent, needs public discussion, and needs conformance tests before any
formal standards-track proposal. The initial goal is incubation,
interoperability feedback, and possibly a Community Group Report.

The proposal must not claim that CWS HTML is already a W3C Standard or W3C
Recommendation.

## Problem Summary

Spreadsheet HTML exports are usually render targets, not stable workbook
packages. AI agents and automation tools need a predictable place to read and
edit workbook structure without rewriting runtime JavaScript, CSS, or
presentation-only markup.

Existing standards such as ODF and OOXML cover rich office workbook packages,
and W3C CSV on the Web covers tabular data metadata. CWS HTML targets a narrower
gap: browser-openable HTML files with a deterministic AI-editable workbook JSON
model, schema markers, AI guidance, validation, and round-trip evidence.

## Technical Shape

CWS HTML separates:

- workbook data: plain JSON in `script#websheet-model`;
- image payloads: optional per-asset encoding under `assets.images`;
- AI editing instructions: `script#cws-ai-instructions` and a public guide;
- runtime behavior: separately loaded application code;
- HTML identity: explicit `data-cws-*` and `meta name="cws:*"` markers.

## Evidence To Lead With

- Cht WebSheet writes and reopens CWS HTML.
- A reference validator accepts valid examples and rejects invalid examples.
- AI-style edits can be made by changing only `script#websheet-model`.
- Edited files can be reopened and resaved while preserving unknown fields.

## Where Details Live

- Final W3C form copy: `docs/cws-html-w3c-cg-submission-copy.md`
- Prior art: `docs/cws-html-prior-art-and-positioning.md`
- Operational agreement draft: `docs/cws-html-cg-operational-agreement-draft.md`
- Submission day runbook: `docs/cws-html-w3c-submission-day-runbook.md`
- Publication checklist: `docs/cws-html-publication-checklist.md`

## Process References

- W3C Community and Business Groups: https://www.w3.org/community/
- W3C Community and Business Group Process:
  https://www.w3.org/community/about/process/
- W3C Recommendation Track readiness guidance:
  https://www.w3.org/guide/standards-track/
