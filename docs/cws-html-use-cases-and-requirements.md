# CWS HTML Use Cases and Requirements

Status: Draft 0.1

Date: 2026-06-12

## Purpose

This document captures the real-world needs that CWS HTML 1.0 Draft is intended
to address during Community Group incubation.

## Actors

- Workbook author: creates and edits workbook-like documents.
- AI editing agent: modifies workbook content based on user instructions.
- Validator: checks conformance without executing scripts.
- Reader: opens and renders a CWS HTML workbook.
- Writer: saves a workbook as CWS HTML.
- Reviewer: inspects structured changes.
- Enterprise workflow system: stores, audits, or transforms documents.

## Use Case 1: AI Edits a Design Workbook

A user sends a CWS HTML design workbook to an AI editing tool and asks it to
update several cells, add formulas, and apply formatting.

Requirements:

- The AI can identify the file as CWS HTML.
- The editable workbook model is in a stable location.
- The AI can edit cells and styles without rewriting runtime code.
- Unknown fields and image payloads are preserved.
- The file can be reopened by Cht WebSheet after the edit.

## Use Case 2: Offline HTML Workbook Sharing

A user saves a workbook as a single HTML file and sends it to another computer.

Requirements:

- Workbook structure and embedded assets are stored in the HTML file.
- Image references such as `assetId` resolve within the same workbook model.
- The runtime can materialize image assets for rendering.
- The file remains valid even when opened from local storage.

## Use Case 3: Static Validation in a Document Pipeline

An enterprise workflow system receives CWS HTML files and wants to reject broken
or unsafe files before storing them.

Requirements:

- Validation does not execute scripts.
- The validator can detect required markers, invalid JSON, invalid ranges, and
  missing image assets.
- The validator can reject top-level encoded workbook wrappers for the
  AI-editable profile.
- Validation results are deterministic and explainable.

## Use Case 4: Structured Review of AI Edits

A reviewer needs to confirm that an AI changed only the intended workbook data.

Requirements:

- Workbook model changes can be compared structurally.
- Large image payloads are summarized rather than printed in full.
- Runtime and loader code should remain unchanged for workbook-content edits.
- The review can distinguish expected edits from accidental data loss.

## Use Case 5: Import/Export Tool Interoperability

A conversion tool extracts sheet data, formulas, styles, and image metadata from
CWS HTML and writes another format.

Requirements:

- Sheet, cell, range, formula, and style structures are readable as JSON.
- Coordinates are stable and 1-based.
- Images can be found either as direct data URLs or as reusable assets.
- Unknown fields can be copied through without interpretation.

## Functional Requirements

- FR1: A conforming file must be recognizable by explicit HTML markers.
- FR2: A conforming file must contain exactly one editable workbook model script.
- FR3: The workbook model must be plain JSON at the top level.
- FR4: Sheets must expose names, sizes, cells, ranges, and visual object metadata
  as JSON.
- FR5: Cell keys must use 1-based `row:col` coordinates.
- FR6: Image assets may be separately encoded, but image placement metadata must
  remain plain JSON.
- FR7: AI instruction metadata must identify the editable region and guide URL.
- FR8: Unknown fields must be preserved by tools that do not understand them.

## Non-Functional Requirements

- NFR1: Validation should not execute page scripts.
- NFR2: Common workbook edits should produce reviewable diffs.
- NFR3: Files should remain ordinary HTML for browser delivery and local opening.
- NFR4: Large binary payloads should be optimizable without hiding workbook
  structure.
- NFR5: The format should allow versioned evolution.
- NFR6: The specification, examples, and tests should be publicly reviewable.

## Open Requirement Questions

- Should `script#cws-ai-instructions` become required for conformance?
- Should the initial profile require all assets to be embedded, or allow remote
  image/runtime references with warnings?
- Should a future vendor-neutral profile define a smaller common workbook model?
- Which formula semantics are required for interoperability versus left to
  implementations?
