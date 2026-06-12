# CWS HTML Round-Trip Report Template

Status: Template

Date:

## Test Summary

- Test ID:
- Tester:
- CWS version:
- Runtime version:
- Source workbook:
- Source workbook type:
- Operating system:
- Browser:

## Purpose

Describe what this round-trip proves. Examples:

- AI can edit cells without changing runtime code.
- Image assets survive save, AI edit, reload, and save.
- Unknown fields are preserved.
- Formulas and styles remain reloadable.

## Source Workbook Features

Check all that apply:

- [ ] Multiple sheets
- [ ] Formulas
- [ ] Cell styles
- [ ] Merged ranges
- [ ] Images
- [ ] Compressed image assets
- [ ] Shapes
- [ ] Filters
- [ ] Tables
- [ ] Comments or notes
- [ ] Hidden rows, columns, or sheets
- [ ] Unknown extension fields

## Procedure

1. Open the source workbook in Cht WebSheet.
2. Save as CWS HTML.
3. Validate the saved file.
4. Edit only `script#websheet-model`.
5. Reopen the edited CWS HTML file.
6. Save again as CWS HTML.
7. Validate the second saved file.
8. Compare expected and actual model changes.

## Commands

```bash
node scripts/validate-cws-html.mjs path/to/original.html
node scripts/validate-cws-html.mjs path/to/edited.html
node scripts/validate-cws-html.mjs path/to/resaved.html
```

## AI Edit Description

Describe the exact requested edit:

```text
Example: Change Plan!C3 from Todo to Done, add green fill, and preserve all image assets.
```

## Expected Model Changes

List paths expected to change:

- `sheets[0].cells["3:3"].raw`
- `sheets[0].cells["3:3"].display`
- `sheets[0].cells["3:3"].css`

## Preservation Checks

- [ ] Runtime scripts and styles were not edited.
- [ ] HTML markers were preserved.
- [ ] `script#websheet-model` remained valid JSON.
- [ ] Sheet order was preserved.
- [ ] Formulas were preserved unless intentionally changed.
- [ ] Merged ranges were preserved unless intentionally changed.
- [ ] Image payloads were preserved unless intentionally changed.
- [ ] Unknown fields were preserved.

## Validation Result

Original saved file:

```text
Paste validator output here.
```

Edited file:

```text
Paste validator output here.
```

Resaved file:

```text
Paste validator output here.
```

## Observed Differences

Summarize the model diff:

```text
Paste or summarize diff output here.
```

## Result

- [ ] Pass
- [ ] Pass with notes
- [ ] Fail

## Notes

Record any warnings, unsupported features, or follow-up issues.
