# CWS HTML Publication Checklist

Status: Draft 0.1

Date: 2026-06-12

## Purpose

This checklist prepares the CWS HTML incubation package for public review. It
keeps public URLs, status wording, licenses, and launch order in one place so
the proposal can be shared without accidentally overstating its standards
status.

## Publication Status Wording

Use this wording before a W3C Community Group is created:

> CWS HTML 1.0 Draft is a project-level incubation draft for an AI-editable HTML
> workbook profile. It is not a W3C Recommendation or W3C Standard.

Use this wording if a W3C Community Group is created:

> CWS HTML 1.0 Draft is an initial contribution to a W3C Community Group
> incubation effort. Community Group work is not automatically a W3C
> Recommendation or W3C Standard.

## Recommended Public URL Map

The public repository already exists:

https://github.com/azamisoft/ChtWebSheet/

The table below assumes the draft package has been pushed to `main`.

| Artifact | Recommended public URL |
| --- | --- |
| Application index | `https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-application-index.md` |
| Submission package | `https://github.com/azamisoft/ChtWebSheet/blob/main/docs/standardization-submission-package.md` |
| Explainer | `https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-explainer.md` |
| Use cases and requirements | `https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-use-cases-and-requirements.md` |
| Draft specification | `https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-1.0-draft.md` |
| Prior-art note | `https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-prior-art-and-positioning.md` |
| CG form draft | `https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-w3c-cg-form-draft.md` |
| CG submission copy | `https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-w3c-cg-submission-copy.md` |
| Submission day runbook | `https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-w3c-submission-day-runbook.md` |
| Operational agreement draft | `https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-cg-operational-agreement-draft.md` |
| Outreach kit | `https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-outreach-kit.md` |
| Open issues | `https://github.com/azamisoft/ChtWebSheet/blob/main/docs/cws-html-open-issues.md` |
| JSON Schema | `https://github.com/azamisoft/ChtWebSheet/blob/main/public/schema/cws-html-workbook-model-v1.schema.json` |
| AI guide | `https://github.com/azamisoft/ChtWebSheet/blob/main/public/ai/cws-html-guide-v1.json` |
| Valid examples | `https://github.com/azamisoft/ChtWebSheet/tree/main/examples/cws-html/valid` |
| Invalid examples | `https://github.com/azamisoft/ChtWebSheet/tree/main/examples/cws-html/invalid` |
| Validator | `https://github.com/azamisoft/ChtWebSheet/blob/main/scripts/validate-cws-html.mjs` |
| Evidence reports | `https://github.com/azamisoft/ChtWebSheet/tree/main/docs/evidence` |

## Link Replacement Checklist

Before public outreach, confirm these URLs resolve after the draft package is
pushed to `main`:

- proposal;
- explainer;
- draft specification;
- prior-art note;
- schema and examples;
- validator and round-trip evidence.

## Minimum Public Package

Do not request broad review until these are public:

- draft specification;
- explainer;
- use cases and requirements;
- prior-art and positioning note;
- JSON Schema;
- AI guide;
- valid and invalid examples;
- validator;
- at least one implementation or evidence report;
- open issues list;
- license or contribution note.

## Recommended Launch Order

1. Publish the draft package to the repository branch or documentation site.
2. Run the validator on examples and evidence.
3. Replace outreach links.
4. Ask a small prior-art review circle before broad posting.
5. Collect at least five supporter permissions.
6. Submit or update the W3C Community Group proposal.
7. Post the launch announcement only after the proposal link is stable.

## Pre-Submission Verification

Run:

```bash
npm run spec:validate
git diff --check
```

For implementation evidence, also run:

```bash
npm run spec:evidence -- --app-url <local-dev-server-url>
```

## License and Rights Checklist

Before public submission:

- confirm who owns the initial specification text;
- confirm that CHT Co., Ltd. permits contribution of the initial draft package;
- decide license for specification text;
- decide license for schema, examples, and validator scripts;
- remove any private or customer data from evidence files;
- confirm whether `CWS`, `Cht WebSheet`, or a neutral profile name should be
  used in public-facing group names.

## Privacy Checklist

Before publishing real workbook evidence:

- remove personal names, customer names, email addresses, and private URLs;
- remove hidden sheets if they contain sensitive data;
- remove comments, document properties, custom metadata, and external links
  unless they are needed for the test;
- prefer synthetic workbooks for public examples;
- keep raw customer files out of the public repository.

## Readiness Gate

The package is ready for public prior-art feedback when:

- all required draft materials have exact public file URLs;
- status wording says "draft" and "not a W3C Recommendation";
- the validator passes examples;
- known open issues are visible;
- evidence is synthetic or cleared for publication;
- supporter outreach text is ready.
