# CWS Agent Progress Report

Date: 2026-11-17

Status: Available for initial use.

## Current Progress

CWS Agent can now be used for AI-assisted design and diagram creation in
Cht WebSheet. Users can describe what they want to create, or provide a
screenshot as a design reference, and work with the resulting editable
content in a CWS workbook.

This makes CWS an environment for creating and refining design documents,
diagrams, and page layouts with AI assistance, alongside its worksheet
editing capabilities.

## What You Can Do Now

| Workflow | Current capability | Example request |
| --- | --- | --- |
| Design documents and layouts | Create and refine structured designs using worksheet content, formatting, and drawing objects. | "Create an order-entry screen design with customer details, an item table, and a totals section." |
| Flowcharts | Create process diagrams with steps, decision points, and connections. | "Draw an approval flowchart with submission, review, approval, and a return-for-revision branch." |
| Sequence diagrams | Create diagrams showing participants and the order of their interactions. | "Create a login sequence diagram with a user, browser, application server, and database." |
| Page design from screenshots | Use a screenshot as a visual reference to create an editable page or UI layout inside CWS. | "Use this screenshot as a reference for a page design, including its header, navigation, form fields, and buttons." |

These are example requests for the available workflows. The resulting
designs remain workbook content that users can inspect and continue editing
in CWS.

## Typical Workflow

1. Open [CWS online](https://chtec.co.jp/cws/) and use CWS Agent with the
   workbook you want to work on.
2. Describe the design or diagram to create. For a screenshot-based design,
   provide the reference image and describe the desired result.
3. Review the generated content in CWS and refine it with follow-up
   instructions or direct editing.
4. Save the result as a CWS HTML workbook so it can be reopened and edited.

For example, a user can start with a screenshot of a business application,
create a page design in CWS, then request changes to labels, layout, or
formatting as the design develops.

## Relationship to CWS HTML

CWS Agent provides implementation experience for the AI-editable HTML workbook
approach: AI assistance can operate on editable workbook content, including
layout and drawing objects, while CWS provides the editing and saving
environment.

This progress report describes application capabilities. The
[implementation report](cws-html-implementation-report.md) and
[evidence reports](evidence/README.md) separately record CWS HTML conformance
work and validation evidence.

CWS HTML 1.0 Draft is being incubated through the
[AI-Editable HTML Workbook Community Group](https://www.w3.org/community/html-workbook/).
This is W3C Community Group incubation work, not a W3C Recommendation or W3C
Standard.

## Feedback

Feedback from initial use is welcome through the
[ChtWebSheet issue tracker](https://github.com/azamisoft/ChtWebSheet/issues).
Describe the requested design, the result, and the changes needed. Examples
of flowcharts, sequence diagrams, and screenshot-based page designs can help
guide further improvements.
