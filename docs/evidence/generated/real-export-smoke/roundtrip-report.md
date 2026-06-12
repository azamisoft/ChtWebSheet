# CWS HTML Real Export Smoke Round-Trip Report

Test ID: real-export-smoke

Generated: 2026-06-12T09:00:18.207Z

## Files

- Original: `docs/evidence/generated/real-export-smoke/exported.cws.html`
- Edited: `docs/evidence/generated/real-export-smoke/edited.cws.html`
- Resaved: `docs/evidence/generated/real-export-smoke/resaved.cws.html`

## Exit Code Notes

- Validation exit code `0` means the file passed validation.
- Diff exit code `0` means no model differences were found.
- Diff exit code `1` means model differences were found and listed below.

## Validation

### Original

Exit code: 0

```text
OK docs/evidence/generated/real-export-smoke/exported.cws.html
```

### Edited

Exit code: 0

```text
OK docs/evidence/generated/real-export-smoke/edited.cws.html
```

### Resaved

Exit code: 0

```text
OK docs/evidence/generated/real-export-smoke/resaved.cws.html
```

## Model Diffs

### Original to Edited

Exit code: 1

```text
Model differences: 4
added $.cwsEvidence
  + <object length=195 sha256=011e96f3a91b540f>
added $.sheets[0].cells["3:3"].display
  + "Done"
changed $.sheets[0].cells["3:3"].raw
  - "Todo"
  + "Done"
changed $.sourceName
  - "source.html"
  + "edited.cws.html"
```

### Edited to Resaved

Exit code: 1

```text
Model differences: 38
added $.fullScreenView
  + false
added $.macroRecording
  + null
added $.macroRelativeReferences
  + false
added $.macroSecurity
  + {"level":"disableWithNotification","trustVbaObjectModel":false,"updatedAt":""}
added $.macros
  + []
changed $.mainWindowView.name
  - "source.xlsx"
  + "edited.cws.html"
added $.sheets[0].activeSheetViewId
  + null
removed $.sheets[0].cells["3:3"].display
  - "Done"
added $.sheets[0].chartSheet
  + false
added $.sheets[0].codeModule
  + null
added $.sheets[0].dialogSheet
  + false
added $.sheets[0].formulaAudits
  + {"precedents":[],"dependents":[],"arrows":[],"ignoredErrors":[]}
added $.sheets[0].inkStrokes
  + []
added $.sheets[0].internationalMacroSheet
  + false
added $.sheets[0].macroSheet
  + false
changed $.sheets[0].rowCount
  - 33
  + 35
added $.sheets[0].rows[33]
  + {"height":20,"hidden":false}
added $.sheets[0].rows[34]
  + {"height":20,"hidden":false}
added $.sheets[0].sheetKind
  + "worksheet"
added $.sheets[0].sheetViews
  + []
added $.sheets[0].specialSheet
  + null
added $.sheets[0].xl4MacroSheet
  + false
added $.showFormulaBar
  + true
added $.synchronousScrolling
  + false
changed $.themeColors[0]
  - "FFFFFF"
  + "#FFFFFF"
changed $.themeColors[1]
  - "000000"
  + "#000000"
changed $.themeColors[2]
  - "EEECE1"
  + "#EEECE1"
changed $.themeColors[3]
  - "1F497D"
  + "#1F497D"
changed $.themeColors[4]
  - "4F81BD"
  + "#4F81BD"
changed $.themeColors[5]
  - "C0504D"
  + "#C0504D"
changed $.themeColors[6]
  - "9BBB59"
  + "#9BBB59"
changed $.themeColors[7]
  - "8064A2"
  + "#8064A2"
changed $.themeColors[8]
  - "4BACC6"
  + "#4BACC6"
changed $.themeColors[9]
  - "F79646"
  + "#F79646"
changed $.themeColors[10]
  - "0000FF"
  + "#0000FF"
changed $.themeColors[11]
  - "800080"
  + "#800080"
added $.viewSideBySide
  + null
added $.windowArrangement
  + null
```

## Result

- [ ] Pass
- [ ] Pass with notes
- [ ] Fail

## Notes

Add human review notes here.
