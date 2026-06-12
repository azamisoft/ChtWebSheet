# CWS HTML Prior Art and Positioning

Status: Draft 0.1

Date: 2026-06-12

## Purpose

This document records the initial prior-art review for the CWS HTML 1.0 Draft
incubation proposal. It is intended to keep the proposal honest: CWS HTML should
not claim to invent spreadsheet standards or office document formats. Its
specific contribution is narrower:

> an HTML workbook profile that keeps a browser-openable workbook and a stable,
> AI-editable workbook JSON model in the same file.

## Preliminary Finding

As of 2026-06-12, a preliminary review of W3C current, proposed, and past
Community Groups did not find an existing group with the same scope as CWS HTML:
an AI-editable HTML workbook file profile with a deterministic workbook model in
`script#websheet-model`.

This is not evidence that no related work exists. Several adjacent standards and
Community Groups are directly relevant and should be acknowledged in any public
proposal.

## W3C Community Group Landscape

W3C Community Group pages explicitly advise proposers to check whether a related
group already exists before proposing a new one. The proposed groups page also
states that a proposed group needs sufficient support before W3C announces its
creation and lists it with current groups.

Relevant W3C CG observations:

| Group or area | Relevance | Positioning |
| --- | --- | --- |
| Web Content Browser for AI Agents Community Group | AI agents reading Web pages; JSON representation optimized for AI consumption. | Adjacent AI/Web representation work. CWS HTML is a workbook file profile with round-trip editing, validation, and saved-file conformance. |
| AI Agent Protocol Community Group | AI agent discovery, identity, metadata, communication, and collaboration protocols. | Adjacent AI agent infrastructure. CWS HTML is document-format work, not an inter-agent protocol. |
| AI-Driven Web Standards Specification Community Group | AI assistance for standards authoring and evaluation. | Process/tooling adjacency only. CWS HTML is a technical document profile, not an AI standards-authoring methodology. |
| Web Platform Incubator Community Group | General venue for Web platform feature incubation. | A possible discussion venue if the proposal becomes a broader Web platform feature. CWS HTML can start as a narrower document-profile incubation effort. |

## Related Standards and Formats

| Work | What it covers | Relationship to CWS HTML |
| --- | --- | --- |
| ODF / OpenDocument | Open office document format for text, spreadsheet, presentation, chart, and graphical documents; standardized through OASIS and ISO/IEC. | Strong prior art for open office document packages. CWS HTML should not compete as a full office suite format. It targets browser-openable HTML plus AI-editable JSON. |
| OOXML / ECMA-376 / ISO/IEC 29500 | Office Open XML vocabularies, document representation, packaging, and producer/consumer requirements. | Strong prior art for rich workbook package representation. CWS HTML should not replace OOXML; it can import/export or map to workbook semantics while exposing a simpler AI-editable HTML profile. |
| W3C CSV on the Web / Tabular Data Model | A model and metadata vocabulary for tabular data on the Web, including CSV-like and related tabular sources. | Strong W3C prior art for tabular data metadata. CWS HTML focuses on rich workbook documents: sheets, cells, formulas, styles, merges, filters, images, and round-trip editing. |
| HTML with JSON script data islands | Common Web pattern for embedding machine-readable JSON in an HTML document. | CWS HTML uses this pattern, but adds explicit CWS markers, schema URL, AI guide, workbook model constraints, validation, and round-trip requirements. |
| Google Sheets HTML export / published sheets | Browser-viewable sheet or table renderings. | Useful comparison point, but generally optimized for display/publishing, not direct workbook model editing by an AI agent. |

## What CWS HTML Should Not Claim

CWS HTML should not claim:

- to be the first spreadsheet or workbook standard;
- to replace ODF, OOXML, CSVW, Google Sheets, or browser-native HTML tables;
- to standardize all Excel or Google Sheets internals;
- to define a universal formula engine;
- to make arbitrary HTML safely AI-editable.

## What CWS HTML Can Claim

CWS HTML can more defensibly claim:

- It defines a recognizable HTML workbook profile.
- It puts the editable workbook model in one deterministic JSON location.
- It separates editable workbook data from runtime JavaScript, CSS, and loader
  code.
- It allows binary-heavy image payloads to be encoded per asset while preserving
  workbook structure as plain JSON.
- It provides AI-facing instructions and a schema URL inside the file.
- It includes validation, invalid examples, model diff tooling, and real export
  round-trip evidence.
- It can complement existing workbook formats by providing an AI-editable HTML
  interchange/profile layer.

## Proposed Public Positioning

Recommended wording for a W3C Community Group proposal:

> Existing standards such as ODF and OOXML standardize office workbook packages,
> and W3C CSVW standardizes tabular data metadata. CWS HTML targets a narrower
> gap: an HTML document profile where the rendered workbook and the editable
> workbook model coexist in one browser-openable file, with deterministic
> AI-editable JSON, schema markers, AI guidance, validation, and round-trip
> conformance evidence.

## Review Checklist Before Submission

- Re-check W3C current, proposed, and past groups for `workbook`, `spreadsheet`,
  `tabular`, `AI agent`, `AI content`, and `HTML document`.
- Re-check non-W3C standardization venues: OASIS, Ecma, ISO/IEC JTC 1/SC 34,
  IETF, WHATWG, and relevant open-source document-format communities.
- Decide whether to propose a new CG or first ask for feedback in an existing
  W3C forum such as WICG, Web & AI-related groups, or a document/data group.
- Prepare a public comparison table and avoid vendor-specific naming where it
  weakens adoption.
- Invite feedback from at least one implementer or reviewer familiar with ODF,
  OOXML, CSVW, or Web document architecture.

## References

- W3C Current Community and Business Groups:
  https://www.w3.org/community/groups/
- W3C Proposed Community and Business Groups:
  https://www.w3.org/community/groups/proposed/
- W3C Past Community and Business Groups:
  https://www.w3.org/community/groups/past/
- W3C Model for Tabular Data and Metadata on the Web:
  https://www.w3.org/TR/tabular-data-model/
- OASIS Open Document Format for Office Applications TC:
  https://www.oasis-open.org/committees/tc_home.php?wg_abbrev=office
- ECMA-376 Office Open XML File Formats:
  https://ecma-international.org/publications-and-standards/standards/ecma-376/
