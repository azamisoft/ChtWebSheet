# CWS HTML Open Issues

Status: Draft 0.1

Date: 2026-06-12

This document tracks issues that should be discussed before or during Community
Group incubation.

## Naming

- ISSUE-001: Should the incubation group use a vendor-neutral name such as
  `AI-Editable HTML Workbook`, while CWS HTML remains the initial
  implementation?
- ISSUE-002: Should `CWS_HTML` remain the format marker, or should a future
  profile marker use a neutral value?

## Conformance

- ISSUE-010: Should `script#cws-ai-instructions` be required or recommended?
- ISSUE-011: Should `meta[name="cws:schema"]` be required for every conforming
  file?
- ISSUE-012: Should remote image URLs be allowed in the core profile or only in
  an extension profile?
- ISSUE-013: Should validators fail or warn when cells exist outside
  `rowCount`/`colCount`?

## Workbook Model

- ISSUE-020: Which fields are required for CWS HTML 1.0 versus Cht WebSheet
  implementation-specific?
- ISSUE-021: Should charts and shapes be standardized in the first draft or
  treated as extension fields?
- ISSUE-022: How much formula semantics should be specified?
- ISSUE-023: Should sheet IDs become mandatory, or are sheet names/order enough
  for 1.0?

## Assets

- ISSUE-030: Should image asset encoding be limited to `gzip-base64`?
- ISSUE-031: Should validators verify gzip payloads by decompressing them, or
  only check declared structure?
- ISSUE-032: Should per-asset hashes be recommended for payload integrity?

## AI Editing

- ISSUE-040: Should AI guide content be normative or informative?
- ISSUE-041: Should an AI edit manifest be added to describe intended changes?
- ISSUE-042: Should the format define a canonical model diff representation?

## Security and Privacy

- ISSUE-050: Should the specification define a privacy-clean export profile?
- ISSUE-051: Should external runtime URLs be allowed in regulated environments?
- ISSUE-052: Should validators warn about hidden sheets, comments, links, and
  metadata?

## Standardization Path

- ISSUE-060: Is a W3C Community Group the right venue, or should the work first
  be discussed in an existing incubator?
- ISSUE-061: What evidence is enough before proposing any Recommendation Track
  transition?
- ISSUE-062: What license should be used for specification text, examples,
  schema, and validator code?
- ISSUE-063: Should the group be proposed as a new Community Group, or should
  prior-art feedback first be requested from WICG, Web & AI groups, OASIS ODF,
  OOXML/ECMA, or CSVW-adjacent reviewers?
- ISSUE-064: Who are the first five individual supporters for W3C Community
  Group creation?
- ISSUE-065: Should the initial operational agreement require seven-day
  asynchronous calls for consensus on normative changes?
- ISSUE-066: Which stable branch or documentation page should provide the exact
  public file URLs for the initial submission package?
- ISSUE-067: Who can approve license, contribution rights, and trademark/naming
  use for public submission?
- ISSUE-068: What is the smallest independent implementation evidence needed
  before broad outreach?
- ISSUE-069: When should accessibility, internationalization, security, and
  privacy reviewers be asked for horizontal feedback?
