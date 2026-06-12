# CWS HTML Application Material Index

Status: Draft 0.1

Date: 2026-06-12

This index lists the files that make up the current CWS HTML standardization
application package.

## Core Proposal

- `docs/standardization-submission-package.md`
- `docs/cws-html-explainer.md`
- `docs/cws-html-use-cases-and-requirements.md`
- `docs/cws-html-prior-art-and-positioning.md`
- `docs/cws-html-1.0-draft.md`
- `docs/w3c-community-group-proposal.md`
- `docs/cws-html-w3c-cg-form-draft.md`
- `docs/cws-html-w3c-cg-submission-copy.md`
- `docs/cws-html-w3c-submission-day-runbook.md`
- `docs/cws-html-cg-operational-agreement-draft.md`
- `docs/cws-html-outreach-kit.md`
- `docs/cws-html-publication-checklist.md`
- `docs/cws-html-readiness-audit.md`
- `docs/cws-html-supporter-tracker.md`
- `docs/cws-html-initial-issues-and-labels.md`
- `docs/cws-html-incubation-roadmap.md`
- `docs/cws-html-open-issues.md`

## Conformance and Testing

- `docs/cws-html-conformance-checklist.md`
- `docs/cws-html-test-plan.md`
- `scripts/validate-cws-html.mjs`
- `scripts/diff-cws-model.mjs`
- `scripts/report-cws-roundtrip.mjs`
- `scripts/generate-cws-evidence.mjs`
- `examples/cws-html/valid/`
- `examples/cws-html/invalid/`
- `docs/evidence/`

## Governance and Risk

- `docs/cws-html-governance-and-ipr.md`
- `docs/cws-html-security-privacy.md`
- `docs/cws-html-readiness-audit.md`
- `docs/standardization-application-plan.md`

## Machine-Readable Materials

- `public/schema/cws-html-workbook-model-v1.schema.json`
- `public/ai/cws-html-guide-v1.json`

## Recommended Review Order

1. Read `docs/cws-html-explainer.md`.
2. Read `docs/cws-html-use-cases-and-requirements.md`.
3. Read `docs/cws-html-prior-art-and-positioning.md`.
4. Read `docs/cws-html-1.0-draft.md`.
5. Run `npm run spec:validate`.
6. Run `npm run spec:evidence` with a dev server when browser export evidence
   is needed.
7. Review `docs/cws-html-conformance-checklist.md`.
8. Review `docs/cws-html-test-plan.md`.
9. Review `docs/cws-html-security-privacy.md`.
10. Use `docs/cws-html-implementation-report.md` for current implementation
   evidence.
11. Review `docs/cws-html-open-issues.md`.
12. Use `docs/cws-html-cg-operational-agreement-draft.md` to check process,
   decision, deliverable, and status wording.
13. Use `docs/cws-html-w3c-cg-submission-copy.md` as the source of truth for
   final copy-paste W3C proposal text.
14. Use `docs/cws-html-w3c-submission-day-runbook.md` for the submission-day
   sequence.
15. Use `docs/cws-html-publication-checklist.md` to prepare public URLs,
   status wording, license checks, and launch order.
16. Use `docs/cws-html-supporter-tracker.md` to recruit five initial
   individual supporters.
17. Use `docs/cws-html-initial-issues-and-labels.md` to set up the public issue
   tracker.
18. Review `docs/cws-html-readiness-audit.md` before broad public sharing.
19. Use `docs/cws-html-outreach-kit.md`,
   `docs/standardization-submission-package.md`, and
   `docs/w3c-community-group-proposal.md` for external submission preparation.
