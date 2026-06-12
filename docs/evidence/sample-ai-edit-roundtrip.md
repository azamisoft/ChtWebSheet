# Sample AI Edit Round-Trip Report

Test ID: sample-ai-edit

Generated: 2026-06-12T08:33:38.325Z

## Files

- Original: `examples/cws-html/valid/ai-edit-before.cws.html`
- Edited: `examples/cws-html/valid/ai-edit-after.cws.html`
- Resaved: `examples/cws-html/valid/ai-edit-after.cws.html`

## Exit Code Notes

- Validation exit code `0` means the file passed validation.
- Diff exit code `0` means no model differences were found.
- Diff exit code `1` means model differences were found and listed below.

## Validation

### Original

Exit code: 0

```text
OK examples/cws-html/valid/ai-edit-before.cws.html
```

### Edited

Exit code: 0

```text
OK examples/cws-html/valid/ai-edit-after.cws.html
```

### Resaved

Exit code: 0

```text
OK examples/cws-html/valid/ai-edit-after.cws.html
```

## Model Diffs

### Original to Edited

Exit code: 1

```text
Model differences: 9
added $.sheets[0].cells["2:4"]
  + {"row":2,"col":4,"raw":"Due","display":"Due","css":{"fontWeight":"700"}}
added $.sheets[0].cells["3:3"].css
  + {"backgroundColor":"#E2F0D9","color":"#375623"}
changed $.sheets[0].cells["3:3"].display
  - "Todo"
  + "Done"
changed $.sheets[0].cells["3:3"].raw
  - "Todo"
  + "Done"
added $.sheets[0].cells["3:4"]
  + {"row":3,"col":4,"raw":"2026-06-12","display":"2026-06-12"}
added $.sheets[0].cells["4:1"]
  + {"row":4,"col":1,"raw":"Publish schema","display":"Publish schema"}
added $.sheets[0].cells["4:2"]
  + {"row":4,"col":2,"raw":"CHT","display":"CHT"}
added $.sheets[0].cells["4:3"]
  + {"row":4,"col":3,"raw":"Todo","display":"Todo"}
changed $.sourceName
  - "ai-edit-before.cws.html"
  + "ai-edit-after.cws.html"
```

### Edited to Resaved

Exit code: 0

```text
No model differences.
```

## Result

- [ ] Pass
- [ ] Pass with notes
- [ ] Fail

## Notes

Add human review notes here.
