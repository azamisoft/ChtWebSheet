# CWS HTML Incubation Roadmap

Status: Draft 0.1

Date: 2026-06-12

## Goal

Prepare CWS HTML 1.0 Draft for W3C Community Group incubation as an
AI-editable HTML workbook profile.

## Phase 0: Project Draft

Status: In progress.

Exit criteria:

- Specification draft exists.
- Explainer exists.
- Use cases and requirements exist.
- JSON Schema exists.
- AI guide exists.
- Validator exists.
- Valid and invalid examples exist.
- CWS writer emits draft-conforming files.

## Phase 1: Evidence Collection

Target duration: 2 to 4 weeks.

Tasks:

- Collect 3 to 5 real exported CWS HTML workbooks.
- Run validation on every exported file.
- Produce at least 3 round-trip reports.
- Confirm image asset preservation across save, AI edit, reload, and resave.
- Add a fixture for unknown-field preservation.
- Expand JSON Schema coverage as new workbook features become part of the draft.

Exit criteria:

- Implementation report includes real-file evidence.
- Round-trip reports have pass/fail outcomes.
- Validator covers known invalid cases.
- Open issues are documented.

## Phase 2: Public Incubation Proposal

Target duration: 2 weeks.

Tasks:

- Publish the draft package in a public repository or documentation site.
- Confirm specification and test asset licenses.
- Prepare W3C Community Group form text.
- Prepare initial Community Group operational agreement text.
- Prepare outreach text for supporters and prior-art feedback.
- Prepare public URL and publication checklist.
- Prepare final W3C submission copy and submission-day runbook.
- Recruit initial supporters.
- Decide whether to propose a CWS-specific or vendor-neutral group name.
- Create public issue labels for requirements, conformance, security, privacy,
  interoperability, and editorial work.

Exit criteria:

- Public URL exists for the draft package.
- Supporters are identified.
- Proposal text is ready to submit.
- Initial operational agreement and outreach kit are ready.
- Issue labels and initial issue texts are ready.
- Final W3C copy-paste text and submission-day runbook are ready.
- Known risks are documented honestly.

## Phase 3: Community Group Incubation

Target duration: 3 to 6 months.

Tasks:

- Hold issue-based technical discussion.
- Refine terminology and conformance language.
- Split CWS-specific behavior from general HTML workbook profile behavior.
- Encourage independent reader, validator, or converter implementations.
- Expand examples and tests.
- Review security, privacy, accessibility, and internationalization concerns.

Exit criteria:

- Community consensus on problem statement and scope.
- At least one independent tool can consume or validate the format.
- Test suite is useful outside Cht WebSheet.
- A Community Group Report draft is possible.

## Phase 4: Recommendation Track Readiness Review

This phase is optional and should not be assumed.

Readiness indicators:

- Clear real-world problem statement.
- Explicit success criteria.
- Evidence of usage or implementer interest.
- Reasonable technical maturity.
- Open issues and points of contention are documented.
- Security, privacy, accessibility, and internationalization reviews are
  requested or planned.

If these indicators are not met, the work should remain a project/community
specification rather than pursuing a formal standards track.
