# CWS HTML Community Group Operational Agreement Draft

Status: Draft 0.1

Date: 2026-06-12

## Purpose

This document is a draft initial operational agreement for a possible W3C
Community Group. It is not a W3C charter and does not claim W3C endorsement.
Its purpose is to make the proposed group's scope, process, deliverables, and
status clear before asking people to support or join the group.

W3C Community Group process allows a group to adopt public operational
agreements such as a charter. These agreements must be fair, public, and
consistent with W3C Community Group process and contribution agreements.

## Proposed Group Name

AI-Editable HTML Workbook Community Group

Alternative names under consideration:

- HTML Workbook Interoperability Community Group
- CWS HTML Workbook Community Group
- AI-Editable Workbook Documents Community Group

## Mission

The mission of the group is to incubate an HTML profile for workbook documents
that can be opened in a browser, inspected as ordinary HTML, and edited by
automation tools through a stable JSON workbook model embedded in the file.

The initial contribution is CWS HTML 1.0 Draft. The group may rename concepts,
generalize vendor-specific fields, split profiles, or reject parts of the
initial contribution if doing so improves interoperability.

## Scope

The group is in scope to discuss, specify, test, and document:

- HTML markers that identify a file as an AI-editable workbook profile.
- A deterministic JSON location for the editable workbook model.
- JSON Schema and validation requirements for the workbook model.
- AI editing guidance embedded in or referenced by the HTML file.
- Conformance requirements for writers, readers, validators, and AI editors.
- Valid and invalid example files.
- Round-trip evidence for AI edits.
- Security and privacy guidance for inspecting workbook HTML without executing
  untrusted runtime code.
- Interoperability with existing workbook import/export systems.

## Out of Scope

The group will not:

- standardize Microsoft Excel, OOXML, BIFF, ODF, or Google Sheets internals;
- require browser engines to implement native spreadsheet editing;
- define a complete formula calculation engine;
- standardize the Cht WebSheet product user interface;
- claim that any draft is a W3C Standard or W3C Recommendation.

## Deliverables

This group may publish Specifications.

Initial deliverables may include:

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

All deliverables should be publicly available and archived.

## Status Language

The group and its participants should describe the work as:

- "a Community Group incubation proposal";
- "a draft HTML workbook profile";
- "a possible Community Group Report";
- "not a W3C Recommendation".

The group and its participants should not describe the work as:

- "a W3C standard";
- "W3C-approved";
- "browser-native spreadsheet support";
- "a replacement for ODF, OOXML, CSVW, or spreadsheet applications".

## Participation

The group should invite participation from:

- workbook application developers;
- Web document format implementers;
- AI coding and document editing tool builders;
- validator, converter, and import/export tool authors;
- enterprise document workflow teams;
- accessibility, internationalization, privacy, and security reviewers;
- people familiar with ODF, OOXML, CSVW, and Web document architecture.

The initial creation proposal needs at least five individual supporters under
the W3C Community Group creation process.

## Roles

Suggested initial roles:

- Chair: coordinates meetings, consensus calls, and process hygiene.
- Specification editor: maintains the draft profile text.
- Schema editor: maintains JSON Schema and schema change notes.
- Test lead: maintains examples, validation tests, and evidence reports.
- Security/privacy reviewer: tracks inspection, execution, and metadata risks.

One person may hold multiple roles during early incubation, but the group should
work toward shared ownership before publishing a stable report.

## Decision Policy

The group should prefer issue-based discussion and consensus.

Suggested default policy:

- Editorial fixes may be accepted by editor discretion.
- Normative changes should have an issue, rationale, and example or validator
  impact note.
- Substantive resolutions should use an asynchronous call for consensus of at
  least seven calendar days.
- Objections should be recorded with the reason and proposed alternative.
- If consensus is not clear, the chair may defer the change, narrow the change,
  or call a recorded vote according to the group's agreed process.

## Communications

Technical work should happen in public:

- public repository;
- public issue tracker;
- public mailing list or W3C-provided discussion channel;
- public meeting agenda and minutes if meetings are held.

Administrative communications may use private channels when needed for personal
information or logistics.

## Contribution and IPR Notes

Participants should expect W3C Community Group contribution rules to apply when
the group is created. Before contributing employer-owned text, code, examples,
or patented technology, participants should confirm that they have permission to
contribute under the applicable W3C Community Group agreements.

The project should keep a clear contribution history for specification text,
schema, examples, and validator code.

## Licensing

Initial recommendation:

- specification and explanatory documents: W3C-compatible document license or a
  standards-compatible open document license;
- schema, examples, validators, and helper scripts: permissive open-source
  license suitable for reuse in conformance tooling.

The final license choice should be confirmed before external submission.

## Initial 90-Day Work Plan

1. Publish the draft package and proposal text.
2. Collect at least five initial supporters.
3. Request prior-art feedback from adjacent Web, AI, workbook, and tabular-data
   communities.
4. Run validation against real exported CWS HTML files.
5. Produce at least three AI edit round-trip reports.
6. Decide whether the profile name should remain CWS HTML or use a neutral
   group-owned name.
7. Separate core profile requirements from implementation-specific Cht WebSheet
   fields.
8. Prepare a first Community Group Report outline if there is sustained
   interest.

## Success Criteria

Incubation is successful if:

- the problem statement is accepted as real by multiple participants;
- at least one implementation can write conforming files;
- at least one independent reader, validator, converter, or AI-editing tool can
  consume the profile;
- the example corpus includes real files and invalid cases;
- the security, privacy, accessibility, and internationalization implications
  are explicitly reviewed;
- open issues and disagreements are documented.

Transition to the W3C Recommendation Track is out of scope for the initial
proposal and should be considered only after broader implementer interest and
review evidence exist.

## References

- W3C Community and Business Groups: https://www.w3.org/community/
- W3C Community and Business Group Process:
  https://www.w3.org/community/about/process/
- W3C Community Groups FAQ: https://www.w3.org/community/about/faq/
- W3C Charter template reference:
  https://raw.githubusercontent.com/w3c/charter-drafts/gh-pages/charter-template.html
