# CWS HTML Standardization Submission Package

Status: Draft 0.1

Date: 2026-06-12

## Executive Summary

CWS HTML is an HTML profile for AI-readable and AI-editable workbook documents.
A CWS HTML file can be opened by a browser as ordinary HTML, while the workbook
structure is stored as plain JSON in `script#websheet-model`. Automated agents
can edit workbook content by modifying that JSON region without rewriting the
runtime, loader, styles, or presentation-only HTML.

The first standardization target should be incubation through a W3C Community
Group. W3C Community Groups are intended for open discussion, specification
development, and test-suite work, and their reports may later become input to a
formal standards-track process.

## Proposed Name

CWS HTML Workbook Format 1.0

## Short Description

An HTML workbook profile that uses explicit CWS markers, a stable JSON workbook
model script, optional separately encoded image assets, and AI editing guidance
to make spreadsheet-like HTML workbooks directly inspectable and editable by
automation tools.

## Problem

Existing spreadsheet HTML exports are usually render targets. They are useful
for viewing a workbook-like table, but they do not expose a stable workbook
model that AI agents can safely modify. Agents must infer structure from visual
markup, risk rewriting unrelated runtime code, and cannot reliably preserve
features such as formulas, styles, merged ranges, filters, images, and workbook
metadata.

This proposal does not claim to replace existing office document standards.
ODF and OOXML already cover rich office workbook packages, and W3C CSV on the
Web covers tabular data metadata. The proposed contribution is an HTML workbook
profile with a deterministic AI-editable JSON region and conformance tooling.

## Proposed Solution

CWS HTML separates the document into four layers:

- Identification: `data-cws-*` and `meta name="cws:*"` markers.
- Workbook model: plain JSON in `script#websheet-model`.
- AI guidance: `script#cws-ai-instructions` and a public guide URL.
- Runtime: shared application code loaded separately from the saved workbook.

Image binary payloads may be stored in `assets.images` and optionally encoded
per asset, but workbook structure, sheet metadata, cell data, formulas, styles,
object positions, and image references remain plain JSON.

## Target Standardization Path

1. Publish the project-level draft and examples.
2. Use the reference validator to build conformance evidence.
3. Open discussion through a W3C Community Group or an existing related Web
   document incubation forum.
4. Collect implementation and interoperability feedback.
5. Consider Recommendation Track or other formal standardization only after
   broader implementation evidence exists.

## Initial Artifacts

- Specification: `docs/cws-html-1.0-draft.md`
- Explainer: `docs/cws-html-explainer.md`
- Use cases and requirements: `docs/cws-html-use-cases-and-requirements.md`
- Prior art and positioning: `docs/cws-html-prior-art-and-positioning.md`
- Conformance checklist: `docs/cws-html-conformance-checklist.md`
- Test plan: `docs/cws-html-test-plan.md`
- Implementation report: `docs/cws-html-implementation-report.md`
- Round-trip report template: `docs/cws-html-round-trip-report-template.md`
- Governance and IPR note: `docs/cws-html-governance-and-ipr.md`
- W3C Community Group proposal: `docs/w3c-community-group-proposal.md`
- W3C Community Group form draft: `docs/cws-html-w3c-cg-form-draft.md`
- W3C Community Group submission copy:
  `docs/cws-html-w3c-cg-submission-copy.md`
- W3C submission day runbook:
  `docs/cws-html-w3c-submission-day-runbook.md`
- Community Group operational agreement draft:
  `docs/cws-html-cg-operational-agreement-draft.md`
- Outreach kit: `docs/cws-html-outreach-kit.md`
- Publication checklist: `docs/cws-html-publication-checklist.md`
- Readiness audit: `docs/cws-html-readiness-audit.md`
- Supporter tracker: `docs/cws-html-supporter-tracker.md`
- Initial issues and labels: `docs/cws-html-initial-issues-and-labels.md`
- Incubation roadmap: `docs/cws-html-incubation-roadmap.md`
- Open issues: `docs/cws-html-open-issues.md`
- JSON Schema: `public/schema/cws-html-workbook-model-v1.schema.json`
- AI guide: `public/ai/cws-html-guide-v1.json`
- Reference validator: `scripts/validate-cws-html.mjs`
- Model diff helper: `scripts/diff-cws-model.mjs`
- Examples: `examples/cws-html/`

## Readiness Evidence

Current:

- Cht WebSheet writes CWS HTML markers and `script#websheet-model`.
- The saved workbook model remains plain JSON.
- Image payloads can be separated into `assets.images`.
- A reference validator checks required markers, JSON Schema conformance, and
  CWS-specific model invariants.
- Valid and invalid conformance examples are present.
- Model diff and round-trip report helpers are present for evidence collection.

Needed before external submission:

- Public URLs for the specification, schema, AI guide, and examples.
- At least one real exported workbook from CWS that validates.
- A round-trip test showing AI edit, reload, and preservation of unknown fields.
- A short issue list for open questions and known gaps.
- A prior-art review covering W3C CGs, ODF, OOXML, CSVW, and related AI/Web
  groups.
- A clear license statement for specification text and test assets.
- Five individual supporters for W3C Community Group creation.
- Public wording that clearly says the work is incubation, not an approved W3C
  standard.
- Early accessibility, internationalization, security, and privacy review
  requests.

## Open Questions

- Should the format name remain `CWS HTML`, or should a vendor-neutral profile
  name be used for wider standardization?
- Should `script#cws-ai-instructions` be mandatory or recommended?
- Should image asset encoding be limited to `gzip-base64`, or should the
  extension point allow additional encodings?
- Should a future media type registration be pursued, or should `text/html`
  with profile markers remain the only delivery mechanism?
- Which workbook features belong in CWS HTML 1.0 conformance versus optional
  implementation extensions?

## Official Process References

- W3C Community and Business Groups: https://www.w3.org/community/
- W3C Community and Business Group Process:
  https://www.w3.org/community/about/process/
- W3C Recommendation Track readiness guidance:
  https://www.w3.org/guide/standards-track/
- W3C Process Document: https://www.w3.org/policies/process/
