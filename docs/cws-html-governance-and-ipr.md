# CWS HTML Governance and IPR Note

Status: Draft 0.1

Date: 2026-06-12

## Purpose

This note describes how the CWS HTML draft should be maintained before and
during standards incubation.

## Current Steward

The current steward is the Cht WebSheet project, developed by CHT Co., Ltd.

## Draft Change Control

Before a standards group adopts the work:

- changes are reviewed in the Cht WebSheet repository;
- normative changes should update the specification, schema, examples, AI guide,
  and validator together;
- breaking changes should require a new CWS HTML schema version;
- workbook internal model migrations should not be confused with CWS HTML schema
  version changes.

## Version Policy

Patch-level clarifications may keep the same CWS HTML schema version when they
do not change conformance requirements.

A new schema version is needed when:

- required markers change;
- required model fields change;
- an existing field changes meaning incompatibly;
- AI editing boundaries change;
- a validator would reject files that were previously conforming.

## Contribution Policy

The project should accept issues and pull requests for:

- specification clarity;
- validator behavior;
- example coverage;
- security and privacy concerns;
- internationalization and accessibility concerns;
- implementation feedback from independent tools.

## Licensing Recommendation

For external standardization, use separate licenses for specification text and
code:

- Specification text and diagrams: CC BY 4.0 or a standards-body-compatible
  document license.
- JSON Schema, validator, and reference code: a permissive license such as MIT
  or Apache-2.0 if they are split into a standards repository.
- The Cht WebSheet application itself may continue to use its project license.

Before external submission, confirm the final license choice with legal review.

## Patent and IPR Considerations

W3C Community Groups have their own contributor and final specification
agreement mechanisms. If the work moves toward a formal standards track, the
project should review patent commitments, contributor agreements, and any
company approval requirements before contributing text or technology.

## Trademark and Naming

`CWS`, `Cht WebSheet`, and related marks should be reviewed before external
standardization. If broader neutral adoption is desired, the group may choose a
vendor-neutral profile name while keeping CWS as the initial implementation.

## Decision Policy

During project-level drafting:

- editorial fixes may be merged after normal review;
- normative changes should include a validator or example update;
- incompatible changes require explicit versioning discussion;
- unresolved issues should be listed in the submission package.

During community incubation, decisions should follow the host group's process.
