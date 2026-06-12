# CWS HTML W3C Submission Day Runbook

Status: Draft 0.1

Date: 2026-06-12

## Purpose

This runbook is the step-by-step checklist for the day a W3C Community Group
proposal is submitted. It assumes the work is still an incubation draft and
must not be described as a W3C Recommendation or W3C Standard.

## Stop Conditions

Do not submit yet if any of these are true:

- exact public file URLs for the draft package are not ready;
- license and rights approval is unresolved;
- evidence files contain private, customer, or personal data;
- the proposal still contains public-link placeholders instead of `main`
  branch URLs;
- the status wording suggests W3C endorsement;
- no likely initial supporters have been contacted.

## Morning Preflight

1. Confirm the public repository is reachable:
   `https://github.com/azamisoft/ChtWebSheet/`.
2. Confirm the draft package has been pushed to the branch or documentation site
   whose URLs will be submitted.
3. Confirm issue tracker access is public.
4. Confirm the draft package has a clear status statement.
5. Confirm private evidence or customer files are not published.
6. Confirm the proposed group name is still acceptable.
7. Confirm the short name `html-workbook` is still acceptable.
8. Confirm at least five people have been privately asked whether they are
   willing to support group creation once the proposed group page exists.

## Verification Commands

Run from the project root:

```bash
npm run spec:validate
git diff --check
```

If publishing fresh implementation evidence, also run:

```bash
npm run spec:evidence -- --app-url <local-dev-server-url>
```

Record the command results in the submission notes.

## URL Confirmation

Confirm the public links in these files point to the `main` branch URLs that
will be submitted:

- `docs/cws-html-w3c-cg-submission-copy.md`
- `docs/cws-html-outreach-kit.md`
- `docs/cws-html-publication-checklist.md`

Check for remaining public-link placeholders:

```bash
rg -n "\\[link\\]|\\[PUBLIC_|<PUBLIC_BLOB_BASE_URL>" docs README.md
```

`[W3C_PROPOSED_GROUP_URL]` may remain in the supporter follow-up message until
W3C creates the proposed group page.

## Submit Proposed Community Group

1. Sign in with the W3C account that will submit the proposal.
2. Open the W3C Community Group proposal flow.
3. Paste the group name from `docs/cws-html-w3c-cg-submission-copy.md`.
4. Paste the short name.
5. Paste the short description.
6. Paste the scope.
7. Include the phrase: "This group may publish Specifications."
8. Include links to the public draft package.
9. Review the final preview for overclaiming.
10. Submit the proposed group.

## Immediately After Submission

1. Save the proposed group URL.
2. Add the proposed group URL to local notes.
3. Replace `[W3C_PROPOSED_GROUP_URL]` in the supporter follow-up message.
4. Ask the five prepared supporters to support the proposed group through the
   W3C page.
5. Update `docs/cws-html-supporter-tracker.md` privately with who has been
   asked and who has responded.
6. Do not announce that the group has launched until W3C shows the group as
   created.

## Public Announcement After Proposal

Use a conservative announcement:

> A proposed W3C Community Group is open for support and feedback: AI-Editable
> HTML Workbook Community Group. The initial contribution is CWS HTML 1.0 Draft,
> an incubation draft for browser-openable, AI-editable workbook HTML. This is
> not a W3C Recommendation or W3C Standard.

Include links to:

- proposed group page;
- explainer;
- draft specification;
- prior-art note;
- issue tracker;
- examples and validator.

## If the Proposal Gets Questions

Respond with the narrow positioning:

- The work complements ODF, OOXML, and CSVW.
- It does not replace office document package standards.
- It does not require browser-native spreadsheet editing.
- It focuses on saved HTML workbook files with an explicit editable JSON model.
- It seeks incubation feedback before any formal standards-track discussion.

## If the Proposal Is Rejected or Redirected

Do not treat rejection as failure of the technical idea. Record the reason and
choose one of these next actions:

- narrow the scope;
- rename the proposal to avoid confusion;
- request feedback in an existing W3C group or Community Group;
- continue as a project/community specification;
- gather more implementation evidence before retrying.

## Submission Record Template

Date:

Submitted by:

W3C account:

Proposed group URL:

Public repository URL:

Spec URL:

Validation command result:

Evidence command result:

Supporters asked:

Supporters confirmed on W3C page:

Follow-up issues:

## Post-Creation Checklist

After the group is created:

1. Update status wording from "proposed group" to "Community Group incubation".
2. Add the W3C group URL to the public repository.
3. Publish or confirm the operational agreement.
4. Open the initial issues from `docs/cws-html-initial-issues-and-labels.md`.
5. Invite prior-art feedback publicly.
6. Schedule the first asynchronous call for scope and naming.
7. Keep all technical decisions archived in public.
