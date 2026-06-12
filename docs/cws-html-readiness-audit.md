# CWS HTML Readiness Audit

Status: Draft 0.1

Date: 2026-06-12

## Purpose

This audit records the current readiness of the CWS HTML incubation package and
the document cleanup decisions made before public submission.

## Current Assessment

The package is strong enough for private prior-art review and supporter
outreach. The public repository already exists:

https://github.com/azamisoft/ChtWebSheet/

It is not yet ready for broad public submission until this draft package is
committed/pushed to a stable public branch or documentation site, and the exact
public file URLs are substituted into the submission templates.

## Source of Truth

| Topic | Source of truth |
| --- | --- |
| Normative profile requirements | `docs/cws-html-1.0-draft.md` |
| Final W3C copy-paste text | `docs/cws-html-w3c-cg-submission-copy.md` |
| Submission-day sequence | `docs/cws-html-w3c-submission-day-runbook.md` |
| Public URL and privacy checks | `docs/cws-html-publication-checklist.md` |
| Supporter recruitment | `docs/cws-html-supporter-tracker.md` |
| Prior-art positioning | `docs/cws-html-prior-art-and-positioning.md` |
| First public issues and labels | `docs/cws-html-initial-issues-and-labels.md` |

Other documents should point to these files instead of duplicating their full
text.

## Repetition Review

Repeated cautionary language about W3C status, ODF/OOXML/CSVW positioning, and
supporter meaning is intentional only in:

- public announcement or email templates that may be copied alone;
- final W3C submission copy;
- submission-day checklist.

The W3C form draft and proposal starter have been shortened to reference the
canonical submission copy instead of repeating the same proposal sections.

## Remaining Gaps Before Broad Public Submission

- Exact public file URLs must be assigned after the draft package is published
  to a stable branch or documentation site.
- License and contribution rights for specification text, examples, schema, and
  validator code must be confirmed.
- At least five credible individual supporters must be ready to support the
  proposed group through the W3C page.
- Evidence files must be confirmed synthetic or cleared for publication.
- An independent reader, validator, converter, or AI-editing proof would
  materially strengthen the proposal.
- Accessibility, internationalization, security, and privacy reviews should be
  requested early, even if they are not complete before Community Group
  creation.
- Vendor-neutral naming remains unresolved.
- The boundary between core profile fields and Cht WebSheet implementation
  fields remains unresolved.

## Current Verification Commands

Run before sharing the package:

```bash
npm run spec:validate
git diff --check
```

Run when issue templates change:

```bash
ruby -e 'require "yaml"; ARGV.each { |path| YAML.load_file(path); puts "OK #{path}" }' .github/ISSUE_TEMPLATE/*.yml
```

## Next Editorial Rule

When adding new material, avoid copying the W3C form text. Link to
`docs/cws-html-w3c-cg-submission-copy.md` unless the new document is itself a
standalone public email, announcement, or form response.
