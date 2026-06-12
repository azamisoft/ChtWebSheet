# CWS HTML Evidence Reports

This directory stores validation, model-diff, and round-trip reports for the CWS
HTML 1.0 Draft incubation package.

Report types:

- Sample reports: generated from small repository fixtures to demonstrate the
  reporting workflow.
- Real export reports: generated from Cht WebSheet exports of real or realistic
  workbooks.
- AI edit reports: generated after an AI agent edits only `script#websheet-model`
  and the file is reopened/resaved.

Generate a report:

```bash
npm run spec:roundtrip-report -- --out docs/evidence/report.md original.html edited.html resaved.html
```

Generate real export smoke evidence from a running Cht WebSheet dev server:

```bash
npm run dev
npm run spec:evidence
```

The evidence generator creates a small xlsx fixture when `--workbook` is not
provided, saves it as CWS HTML through the browser UI, edits only
`script#websheet-model`, reopens/resaves the edited HTML, validates all three
HTML files, and writes a round-trip report plus `summary.json`.

For formal submission evidence, prefer real export reports over sample fixture
reports.

Current sample reports:

- `sample-ai-edit-roundtrip.md`: generated from repository fixture files. This
  demonstrates the reporting workflow only; it is not real-world export
  evidence.

Current real export smoke evidence:

- `generated/real-export-smoke/summary.json`: machine-readable summary of the
  automated browser export, AI-style edit, standalone reopen/save, validation,
  and assertions.
- `generated/real-export-smoke/roundtrip-report.md`: generated validation and
  model-diff report for the same run.
- `generated/real-export-smoke/exported.cws.html`: CWS HTML saved through the
  browser UI from `source.xlsx`.
- `generated/real-export-smoke/edited.cws.html`: same file after editing only
  `script#websheet-model`.
- `generated/real-export-smoke/resaved.cws.html`: edited file reopened through
  standalone runtime and saved again.
