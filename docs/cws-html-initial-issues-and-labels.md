# CWS HTML Initial Issues and Labels

Status: Draft 0.1

Date: 2026-06-12

## Purpose

This document prepares the first public issue tracker setup for CWS HTML
incubation. It is useful whether the work is hosted in this repository, a new
neutral repository, or a future W3C Community Group repository.

## Recommended Labels

| Label | Purpose |
| --- | --- |
| `cws-html` | General label for the profile work. |
| `requirements` | Use cases, user needs, and acceptance criteria. |
| `prior-art` | Related standards, formats, venues, and overlap checks. |
| `conformance` | Normative requirements, validator behavior, and examples. |
| `schema` | JSON Schema and workbook model shape. |
| `ai-editing` | AI guide, edit boundaries, model diff, and round-trip behavior. |
| `assets` | Images, binary payloads, asset references, and compression. |
| `security` | Script execution, unsafe inspection, links, metadata, and trust. |
| `privacy` | Hidden data, workbook metadata, personal data, and clean export. |
| `i18n` | Internationalization, locale, text direction, formula names, dates. |
| `a11y` | Accessibility implications and assistive technology needs. |
| `interop` | Independent readers, writers, validators, and converters. |
| `evidence` | Real-file exports, round-trip reports, and implementation reports. |
| `editorial` | Wording, examples, broken links, and non-normative cleanup. |
| `process` | W3C CG setup, decision policy, licensing, and governance. |

## Initial Issue 1: Scope and Venue

Title:

Confirm scope and venue for AI-editable HTML workbook incubation

Labels:

`process`, `prior-art`, `requirements`

Body:

The draft proposes an HTML workbook profile where the editable workbook model is
stored as plain JSON in `script#websheet-model`.

Questions:

- Should this be proposed as a new W3C Community Group?
- Should it first be discussed in an existing Web, AI, document, or tabular-data
  venue?
- Is the scope narrow enough?
- Which related work should be acknowledged before public submission?

Relevant drafts:

- `docs/cws-html-prior-art-and-positioning.md`
- `docs/cws-html-w3c-cg-form-draft.md`
- `docs/cws-html-cg-operational-agreement-draft.md`

## Initial Issue 2: Vendor-Neutral Naming

Title:

Decide whether the public profile name should remain CWS HTML

Labels:

`process`, `requirements`, `editorial`

Body:

The initial implementation uses CWS HTML markers and comes from Cht WebSheet.
For broader incubation, the group may want a neutral name such as
AI-Editable HTML Workbook or HTML Workbook Profile.

Questions:

- Should `CWS HTML` remain the profile name?
- Should CWS remain only the initial implementation name?
- Should future markers use a neutral profile value?
- What migration path would avoid breaking existing CWS files?

## Initial Issue 3: Core Model Boundary

Title:

Define which workbook model fields belong in the core profile

Labels:

`schema`, `conformance`, `interop`

Body:

The current model contains both general workbook concepts and implementation
fields. Incubation needs a clear boundary between required core fields,
recommended fields, and extension fields.

Questions:

- Which top-level fields are required for CWS HTML 1.0?
- Which worksheet fields are required?
- Which Cht WebSheet-specific fields should become extensions?
- How should unknown fields be preserved?

Relevant drafts:

- `docs/cws-html-1.0-draft.md`
- `public/schema/cws-html-workbook-model-v1.schema.json`
- `docs/cws-html-conformance-checklist.md`

## Initial Issue 4: AI Editing Requirements

Title:

Define AI editing boundaries and round-trip expectations

Labels:

`ai-editing`, `conformance`, `evidence`

Body:

The central claim is that AI tools can edit workbook content by modifying only
`script#websheet-model`.

Questions:

- Which edits must a conforming AI editor be able to make safely?
- Should AI instructions be normative or informative?
- Should edits include an edit manifest or change log?
- What round-trip evidence is enough for the first draft?

Relevant drafts:

- `public/ai/cws-html-guide-v1.json`
- `docs/cws-html-round-trip-report-template.md`
- `docs/evidence/`

## Initial Issue 5: Image Asset Encoding

Title:

Decide the core profile rules for image asset encoding

Labels:

`assets`, `schema`, `conformance`

Body:

CWS HTML keeps workbook structure as plain JSON while allowing binary-heavy
image payloads to be encoded per asset.

Questions:

- Should `gzip-base64` be the only compressed image encoding in 1.0?
- Should validators decompress assets or only check declared structure?
- Should image assets include hashes?
- Should remote image URLs be disallowed in the core profile?

## Initial Issue 6: Security and Privacy Review

Title:

Review security and privacy implications of CWS HTML inspection and editing

Labels:

`security`, `privacy`, `conformance`

Body:

CWS HTML is ordinary HTML and may include runtime code, links, metadata, hidden
sheets, comments, or embedded assets. Tools need safe inspection guidance.

Questions:

- Should validators inspect without executing scripts?
- Should the profile define a privacy-clean export option?
- Which hidden workbook data should trigger warnings?
- Should external runtime URLs be allowed in regulated environments?

Relevant draft:

- `docs/cws-html-security-privacy.md`

## Initial Issue 7: Independent Implementation Evidence

Title:

Collect independent reader, validator, or converter implementation evidence

Labels:

`interop`, `evidence`, `conformance`

Body:

The proposal is stronger if at least one tool outside the initial writer can
read, validate, convert, or AI-edit CWS HTML.

Questions:

- What is the smallest useful independent implementation?
- Should it be a validator, model extractor, converter, or AI edit tool?
- Which examples should it pass?
- How should implementation evidence be reported?

Relevant drafts:

- `docs/cws-html-implementation-report.md`
- `docs/cws-html-test-plan.md`

## Initial Issue 8: Public Examples and Test Corpus

Title:

Expand public valid and invalid examples for first incubation review

Labels:

`conformance`, `evidence`, `schema`

Body:

The current example corpus should grow as requirements become clearer.

Suggested examples:

- workbook with formulas and cached values;
- workbook with merged ranges;
- workbook with hidden sheets or hidden rows;
- workbook with image assets;
- workbook with unknown extension fields;
- invalid duplicate model script;
- invalid compressed top-level model;
- invalid asset reference;
- invalid cell coordinates.

Relevant directories:

- `examples/cws-html/valid/`
- `examples/cws-html/invalid/`

## Optional GitHub CLI Label Commands

If the repository uses GitHub Issues, labels can be created with:

```bash
gh label create cws-html --description "CWS HTML profile work"
gh label create requirements --description "Use cases, user needs, and acceptance criteria"
gh label create prior-art --description "Related standards, formats, venues, and overlap checks"
gh label create conformance --description "Normative requirements, validator behavior, and examples"
gh label create schema --description "JSON Schema and workbook model shape"
gh label create ai-editing --description "AI guide, edit boundaries, model diff, and round-trip behavior"
gh label create assets --description "Images, binary payloads, asset references, and compression"
gh label create security --description "Script execution, unsafe inspection, links, metadata, and trust"
gh label create privacy --description "Hidden data, workbook metadata, personal data, and clean export"
gh label create i18n --description "Internationalization, locale, text direction, formula names, dates"
gh label create a11y --description "Accessibility implications and assistive technology needs"
gh label create interop --description "Independent readers, writers, validators, and converters"
gh label create evidence --description "Real-file exports, round-trip reports, and implementation reports"
gh label create editorial --description "Wording, examples, broken links, and non-normative cleanup"
gh label create process --description "W3C CG setup, decision policy, licensing, and governance"
```
