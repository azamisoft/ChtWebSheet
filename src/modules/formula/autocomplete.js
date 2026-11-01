// Feature module; shared bindings are imported explicitly.
import {
  $,
  $formulaInput,
  FORMULA_ARGUMENT_DEFAULT_OVERRIDES,
  FORMULA_COMPATIBILITY_FUNCTIONS,
  FORMULA_ENGINE_ARGUMENT_OVERRIDES,
  FORMULA_ENGINE_DESCRIPTION_OVERRIDES,
  FORMULA_ENGINE_SYNTAX_OVERRIDES,
  FORMULA_REFERENCE_COLORS,
  INSERT_FUNCTION_CATEGORY_ORDER,
  INSERT_FUNCTION_RECENT_NAMES,
  formulaEngineAutocompleteFunctionDefinitions,
  state,
} from "../../app-core.js";
import {
  isAltModifierEvent,
  isEnterKeyEvent,
  positiveModulo,
} from "../editing/autofill.js";
import {
  isImeCompositionEvent,
} from "../editing/keyboard-and-paste.js";
import {
  cellPointFromElement,
  rangeToLabel,
  syncFormulaEditPreview,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  syncEditableCellAlignment,
} from "../editing/selection-and-cell-text.js";
import {
  closeFormulaAutocomplete,
  formulaArgumentDescription,
  formulaEditorCaretOffset,
  formulaEditorSelectionOffsets,
  formulaEditorSetTextAndCaret,
  formulaEditorText,
  formulaFunctionArgumentSpecs,
  formulaTemplateReference,
  functionArgumentDisplaySpecs,
  openFunctionArgumentsDialog,
  recordFormulaFunctionUse,
  updateFormulaReferenceUi,
} from "./dialogs-and-names.js";
import {
  escapeAttr,
  escapeHtml,
} from "../ui/help-and-utilities.js";
import {
  firstCellReferenceFromRangeLabel,
  formulaAutocompleteExtraFunctionDefinitions,
  formulaFunctionByName,
  formulaFunctionDefinitions,
  quoteSheetNameForFormula,
} from "../ui/ribbon-and-links.js";
import {
  cleanStructuredReferencePart,
  structuredReferencePartHasEscapedSpecial,
  structuredReferenceTokenAt,
  unescapeStructuredReferencePart,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  escapeStructuredReferenceColumnName,
} from "../workbook/tables-and-validation.js";


function formulaAutocompleteShouldExposeEngineFunction(name) {
  const text = String(name || "").trim().toUpperCase();
  if (!text || text.startsWith("HF.")) return false;
  return !new Set(["ARRAY_CONSTRAIN"]).has(text);
}



function formulaEngineFunctionMetadata(hf, name) {
  const registry = hf?._functionRegistry;
  if (!registry) return null;
  const metadata = registry.functionsMetadata;
  return (
    metadata?.get?.(name) ||
    metadata?.[name] ||
    registry.getFunctionPlugin?.(name)?.implementedFunctions?.[name] ||
    null
  );
}



function formulaEngineFunctionIsCompatibility(name) {
  return FORMULA_COMPATIBILITY_FUNCTIONS.has(String(name || "").toUpperCase());
}



function formulaEngineFunctionCategory(name) {
  const text = String(name || "").toUpperCase();
  const inSet = (items) => items.includes(text);
  if (
    inSet(["DSUM", "DCOUNT", "DCOUNTA", "DAVERAGE", "DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP", "DVAR", "DVARP"])
  ) return "データベース";
  if (
    inSet(["FORMULATEXT", "ISFORMULA", "ISBLANK", "ISERROR", "ISERR", "ISNA", "ISNUMBER", "ISTEXT", "ISNONTEXT", "ISLOGICAL", "ISREF", "ISEVEN", "ISODD", "ERROR.TYPE", "N", "NA", "TYPE", "INFO", "CELL"])
  ) return "情報";
  if (
    inSet(["IF", "IFS", "LET", "AND", "OR", "XOR", "NOT", "TRUE", "FALSE", "SWITCH", "IFERROR", "IFNA"]) ||
    /^IS/.test(text)
  ) return "論理/情報";
  if (
    inSet(["DATE", "TIME", "TODAY", "NOW", "DAY", "MONTH", "YEAR", "HOUR", "MINUTE", "SECOND", "DATEVALUE", "TIMEVALUE", "EOMONTH", "EDATE", "DAYS", "DAYS360", "WEEKDAY", "WEEKNUM", "ISOWEEKNUM", "NETWORKDAYS", "NETWORKDAYS.INTL", "WORKDAY", "WORKDAY.INTL", "YEARFRAC"]) ||
    /(DAY|DAYS|DATE|TIME|WEEK|MONTH|YEAR|HOUR|MINUTE|SECOND)$/.test(text)
  ) return "日付/時刻";
  if (
    inSet(["LEFT", "RIGHT", "MID", "LEN", "FIND", "SEARCH", "SUBSTITUTE", "REPLACE", "TEXT", "TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER", "VALUE", "TRIM", "UPPER", "LOWER", "PROPER", "CONCATENATE", "CHAR", "UNICHAR", "CODE", "UNICODE", "EXACT", "REPT", "CLEAN", "SPLIT", "T"]) ||
    /TEXT|CHAR|CODE|UNICODE|STRING|CONCAT|SPLIT/.test(text)
  ) return "文字列";
  if (
    inSet(["XLOOKUP", "VLOOKUP", "HLOOKUP", "LOOKUP", "INDEX", "MATCH", "XMATCH", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "TAKE", "DROP", "VSTACK", "HSTACK", "MAKEARRAY", "BYROW", "BYCOL", "MAP", "REDUCE", "SCAN", "CHOOSEROWS", "CHOOSECOLS", "TOCOL", "TOROW", "WRAPROWS", "WRAPCOLS", "EXPAND", "OFFSET", "INDIRECT", "ADDRESS", "ROW", "ROWS", "COLUMN", "COLUMNS", "CHOOSE", "TRANSPOSE", "SHEET", "SHEETS", "HYPERLINK"]) ||
    /LOOKUP|MATCH|INDEX|ADDRESS|ROW|COLUMN|SHEET|FILTER/.test(text)
  ) return "検索/行列";
  if (
    inSet(["PMT", "PV", "FV", "FVSCHEDULE", "NPV", "XNPV", "XIRR", "IRR", "MIRR", "RATE", "NPER", "IPMT", "PPMT", "CUMIPMT", "CUMPRINC", "DB", "DDB", "VDB", "AMORLINC", "AMORDEGRC", "SLN", "SYD", "DOLLARDE", "DOLLARFR", "EFFECT", "NOMINAL", "RRI", "PDURATION", "DISC", "PRICEDISC", "YIELDDISC", "INTRATE", "RECEIVED", "ACCRINT", "ACCRINTM", "PRICEMAT", "YIELDMAT", "COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPPCD", "COUPNUM", "PRICE", "YIELD", "DURATION", "MDURATION", "TBILLEQ", "TBILLPRICE", "TBILLYIELD", "ISPMT"]) ||
    /COUP|PRICE|YIELD|RATE|DURATION|DOLLAR|TBILL/.test(text)
  ) return "財務";
  if (
    /^IM/.test(text) ||
    /^(BIN|DEC|HEX|OCT)2/.test(text) ||
    /^(BIT|BESSEL|COMPLEX|CONVERT|DELTA|GESTEP|ERF|ERFC)/.test(text)
  ) return "エンジニアリング";
  if (formulaEngineFunctionIsCompatibility(text)) return "互換性";
  if (
    inSet(["AVERAGE", "AVERAGEA", "AVERAGEIF", "AVERAGEIFS", "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS", "COUNTUNIQUE", "MAX", "MAXA", "MAXIFS", "MIN", "MINA", "MINIFS", "MEDIAN", "MODE.SNGL", "MODE.MULT", "FREQUENCY", "TRIMMEAN", "PROB", "LARGE", "SMALL", "RANK.EQ", "RANK.AVG", "PERCENTILE.INC", "PERCENTRANK.INC", "PERCENTRANK.EXC", "CORREL", "COVARIANCE.S", "COVARIANCE.P", "PEARSON", "RSQ", "STEYX", "SKEW", "SKEW.P", "STDEV", "STDEVA", "STDEVP", "STDEVPA", "STDEV.S", "STDEV.P", "VAR", "VARA", "VARP", "VARPA", "VAR.S", "VAR.P"]) ||
    /(DIST|INV|TEST|NORM|BINOM|CHISQ|FISHER|GAMMA|BETA|WEIBULL|POISSON|HYPGEOM|CONFIDENCE|SKEW|KURT|PEARSON|PROB|RSQ|SLOPE|STEYX|DEVSQ|MEAN|VAR|STDEV|COVAR|RANK|PERCENTILE|QUARTILE|FREQUENCY)/.test(text)
  ) return "統計";
  return "数学/三角";
}



function formulaEngineFunctionDescription(name) {
  const text = String(name || "").toUpperCase();
  return FORMULA_ENGINE_DESCRIPTION_OVERRIDES[text] || `${text} 関数を挿入します。`;
}



function formulaEngineFunctionSyntax(name, metadata) {
  const text = String(name || "").toUpperCase();
  const override = FORMULA_ENGINE_SYNTAX_OVERRIDES[text];
  if (override) return override;
  const parameters = Array.isArray(metadata?.parameters) ? metadata.parameters : [];
  if (!parameters.length) return `${text}()`;
  const repeat = Number(metadata?.repeatLastArgs) || 0;
  const args = parameters.map((param, index) => formulaEngineArgumentName(text, param, index, 1));
  if (repeat > 0 && repeat <= args.length) {
    const repeated = parameters
      .slice(parameters.length - repeat)
      .map((param, index) => formulaEngineArgumentName(text, param, parameters.length - repeat + index, 2));
    return `${text}(${args.join(", ")}, [${repeated.join(", ")}], ...)`;
  }
  return `${text}(${parameters.map((param, index) => {
    const arg = formulaEngineArgumentName(text, param, index, 1);
    return formulaEngineArgumentIsOptional(param) ? `[${arg}]` : arg;
  }).join(", ")})`;
}



function formulaEngineArgumentIsOptional(param) {
  return Boolean(param?.optionalArg || Object.prototype.hasOwnProperty.call(param || {}, "defaultValue"));
}



function formulaEngineArgumentName(functionName, param, index, sequence) {
  const overrides = FORMULA_ENGINE_ARGUMENT_OVERRIDES[functionName];
  if (overrides?.[index]) {
    return overrides[index].replace(/1$/u, String(sequence || 1));
  }
  const type = String(param?.argumentType || "ANY").toUpperCase();
  const base = {
    ANY: "value",
    BOOLEAN: "logical",
    COMPLEX: "inumber",
    INTEGER: "number",
    NOERROR: "value",
    NUMBER: "number",
    RANGE: formulaEngineFunctionCategory(functionName) === "検索/行列" ? "array" : "number",
    SCALAR: "value",
    STRING: "text",
  }[type] || "value";
  return `${base}${sequence || index + 1}`;
}



function formulaInsertFunctionDefinitions() {
  const byName = new Map();
  formulaAutocompleteDefinitions()
    .filter((definition) => definition.kind === "function")
    .forEach((definition) => {
      const name = String(definition.name || "").toUpperCase();
      if (!name) return;
      byName.set(name, {
        ...definition,
        name,
        label: definition.label || name,
        description: definition.description || formulaEngineFunctionDescription(name),
        syntax: definition.syntax || formulaDefaultSyntax(name),
        template: definition.template || ((ref) => formulaDefaultTemplate(name, ref)),
      });
    });
  return [...byName.values()].sort((left, right) => {
    const leftRecent = INSERT_FUNCTION_RECENT_NAMES.indexOf(left.name);
    const rightRecent = INSERT_FUNCTION_RECENT_NAMES.indexOf(right.name);
    const leftRank = leftRecent >= 0 ? leftRecent : 999;
    const rightRank = rightRecent >= 0 ? rightRecent : 999;
    return (
      leftRank - rightRank ||
      (left.priority || 999) - (right.priority || 999) ||
      formulaAutocompleteDisplayName(left).localeCompare(formulaAutocompleteDisplayName(right), "ja")
    );
  });
}



function formulaInsertFunctionDefinitionByName(functionName) {
  const name = String(functionName || "").trim().toUpperCase();
  return formulaInsertFunctionDefinitions().find((definition) => definition.name === name) || null;
}



function insertFunctionCategoryOptions(definitions) {
  const categories = new Set(definitions.map((item) => item.category).filter(Boolean));
  const ordered = INSERT_FUNCTION_CATEGORY_ORDER.filter((category) => category === "最近使用した関数" || category === "すべて表示" || categories.has(category));
  const extras = [...categories].filter((category) => !ordered.includes(category)).sort((left, right) => left.localeCompare(right, "ja"));
  return [...ordered, ...extras];
}



function formulaAutocompleteDefinitions() {
  const byName = new Map();
  formulaFunctionDefinitions().forEach((definition, index) => {
    const name = String(definition.name || "").toUpperCase();
    byName.set(name, {
      ...definition,
      kind: "function",
      name,
      label: definition.label || name,
      priority: index + 1,
      syntax: definition.syntax || formulaDefaultSyntax(name),
    });
  });
  formulaAutocompleteExtraFunctionDefinitions().forEach((definition) => {
    const name = String(definition.name || "").toUpperCase();
    byName.set(name, {
      ...(byName.get(name) || {}),
      ...definition,
      kind: "function",
      name,
      label: definition.label || name,
      template: byName.get(name)?.template || ((ref) => formulaDefaultTemplate(name, ref)),
      syntax: definition.syntax || formulaDefaultSyntax(name),
    });
  });
  formulaEngineAutocompleteFunctionDefinitions().forEach((definition) => {
    const name = String(definition.name || "").toUpperCase();
    if (!name) return;
    if (byName.has(name)) {
      byName.set(name, {
        ...definition,
        ...byName.get(name),
        engineSupported: Boolean(definition.engineSupported),
      });
      return;
    }
    byName.set(name, definition);
  });
  return [...byName.values(), ...formulaWorkbookAutocompleteDefinitions()].sort(
    (left, right) => (left.priority || 999) - (right.priority || 999) || formulaAutocompleteDisplayName(left).localeCompare(formulaAutocompleteDisplayName(right), "ja"),
  );
}



function formulaWorkbookAutocompleteDefinitions() {
  const entries = formulaWorkbookSheetReferenceDefinitions();
  const activeSheetIndex = state.activeSheetIndex;
  (state.model?.definedNames || []).forEach((definition, index) => {
    if (!definition?.name || definition.hidden) return;
    if (definition.scope === "sheet" && Number(definition.sheetIndex) !== activeSheetIndex) return;
    const scopeLabel = definition.scope === "sheet" && state.model?.sheets?.[definition.sheetIndex]
      ? state.model.sheets[definition.sheetIndex].name
      : "ブック";
    entries.push({
      kind: "name",
      name: String(definition.name),
      label: String(definition.name),
      category: definition.scope === "sheet" ? "名前 (シート)" : "名前",
      description: `${scopeLabel}: ${definition.refersTo || ""}`,
      priority: 200 + index,
      matchText: String(definition.name),
      insertText: String(definition.name),
    });
  });
  workbookFormulaTableRefs().forEach((entry, index) => {
    entries.push({
      kind: "table",
      name: entry.name,
      label: entry.name,
      category: "テーブル",
      description: `${entry.sheet.name}!${rangeToLabel(entry.table.range)}`,
      priority: 300 + index,
      matchText: entry.name,
      insertText: entry.name,
      tableRef: entry,
    });
  });
  return entries;
}



function formulaWorkbookSheetReferenceDefinitions() {
  return (state.model?.sheets || []).map((sheet, index) => {
    const name = String(sheet?.name || `Sheet${index + 1}`);
    const referencePrefix = `${quoteSheetNameForFormula(name)}!`;
    return {
      kind: "sheet",
      name,
      label: name,
      category: "シート",
      syntax: referencePrefix,
      description: `${name} シートへの参照を開始します。`,
      priority: 180 + index,
      matchText: name,
      insertText: referencePrefix,
      sheetIndex: index,
    };
  });
}



function workbookFormulaTableRefs() {
  const refs = [];
  (state.model?.sheets || []).forEach((sheet, sheetIndex) => {
    (sheet?.tables || []).forEach((table, tableIndex) => {
      if (!table?.range) return;
      refs.push({
        sheet,
        sheetIndex,
        table,
        tableIndex,
        name: String(table.displayName || table.name || `Table${tableIndex + 1}`),
        columns: formulaTableColumnNames(table),
      });
    });
  });
  return refs;
}



function formulaTableColumnNames(table) {
  if (Array.isArray(table?.columns) && table.columns.length) {
    return table.columns.map((column, index) => String(column?.name || `Column${index + 1}`));
  }
  const width = table?.range ? Math.max(1, table.range.right - table.range.left + 1) : 1;
  return Array.from({ length: width }, (_item, index) => `Column${index + 1}`);
}



function formulaAutocompleteDisplayName(candidate) {
  return String(candidate?.label || candidate?.name || "");
}



function formulaAutocompleteMatchText(candidate) {
  return String(candidate?.matchText || candidate?.label || candidate?.name || "").toLocaleUpperCase("ja-JP");
}



function formulaAutocompleteDefinitionByName(functionName) {
  const name = String(functionName || "").trim().toUpperCase();
  return formulaAutocompleteDefinitions().find((definition) => definition.kind === "function" && definition.name === name) || null;
}



function formulaDefaultSyntax(functionName) {
  return `${String(functionName || "").toUpperCase()}(argument1, [argument2], ...)`;
}



function formulaReferenceColorClass(colorIndex) {
  const index = positiveModulo(Number(colorIndex) || 0, FORMULA_REFERENCE_COLORS.length);
  return `formula-reference-color-${index}`;
}



function formulaDefaultTemplate(functionName, ref) {
  const name = String(functionName || "SUM").toUpperCase();
  const definition = formulaAutocompleteDefinitionByName(name) || { name, syntax: FORMULA_ENGINE_SYNTAX_OVERRIDES[name] || formulaDefaultSyntax(name) };
  const syntax = definition.syntax || formulaDefaultSyntax(name);
  if (/\(\s*\)$/.test(syntax)) return `=${name}()`;
  const specs = formulaFunctionArgumentSpecs(definition);
  if (!specs.length) return `=${name}(${ref || ""})`;
  const values = specs.map((spec, index) => formulaDefaultTemplateArgument(name, spec, index, ref || formulaTemplateReference()));
  while (values.length && !String(values[values.length - 1] || "").trim()) values.pop();
  return `=${name}(${values.join(",")})`;
}



function formulaDefaultTemplateArgument(functionName, spec, index, ref) {
  const name = String(functionName || "").toUpperCase();
  const override = FORMULA_ARGUMENT_DEFAULT_OVERRIDES[name]?.[index];
  if (typeof override === "function") return override();
  if (override != null) return String(override);
  if (spec?.optional) return "";
  const label = String(spec?.label || "").toLowerCase();
  const firstCell = firstCellReferenceFromRangeLabel(ref);
  if (/logical|cumulative|no_commas|ignore_empty|range_lookup|a1/.test(label)) return "TRUE";
  if (/criteria/.test(label)) return '"<>"';
  if (/delimiter|separator/.test(label)) return '","';
  if (/format/.test(label)) return '"0"';
  if (/date_text/.test(label)) return '"2000/1/1"';
  if (/time_text/.test(label)) return '"12:00"';
  if (/old_text|new_text|find_text|within_text|text1|text2|text|friendly|link_location|suffix|unit/.test(label)) return '""';
  if (/lookup_value/.test(label)) return firstCell;
  if (/table_array|lookup_array|return_array|known_y|known_x|actual_range|expected_range|array|range|database|criteria|ref|value|number|sum|min|max/.test(label)) return ref || "A1";
  if (/field/.test(label)) return "1";
  if (/rate/.test(label)) return "0.05";
  if (/nper|life|period|per|month|day|year|hour|minute|second|index|row|column|digits|type|mode|order|basis|guess|bottom|top|places|quart|tails/.test(label)) return "1";
  if (/pv|fv|pmt|cost|salvage|finance|reinvest|mean|standard|probability|x|k|angle|multiple|significance|denominator|numerator|chosen|decimals/.test(label)) return "1";
  return "1";
}



function updateFormulaAutocompleteForEditor(editor, options = {}) {
  if (state.suppressFormulaAutocompleteInput) return;
  if (!editor || isImeComposingForEditor(editor)) {
    if (!editor) closeFormulaAutocomplete({ restoreFocus: false });
    return;
  }
  const text = formulaEditorText(editor);
  const caret = formulaEditorCaretOffset(editor);
  const accepted = state.formulaAutocompleteAccepted;
  if (accepted?.editor === editor) {
    if (accepted.text === text && accepted.caret === caret) {
      closeFormulaAutocomplete({ restoreFocus: false });
      return;
    }
    state.formulaAutocompleteAccepted = null;
  }
  if (!formulaTextStartsFormula(text)) {
    closeFormulaAutocomplete({ restoreFocus: false });
    return;
  }
  const queryContext = formulaAutocompleteQueryContext(text, caret, formulaAutocompleteEditorContext(editor));
  const hintContext = formulaParameterHintContext(text, caret);
  const candidates = formulaAutocompleteShouldShowCandidates(queryContext) ? formulaAutocompleteCandidates(queryContext) : [];
  const hint = hintContext ? formulaParameterHintHtml(hintContext) : "";
  if (!candidates.length && !hint) {
    closeFormulaAutocomplete({ restoreFocus: false });
    return;
  }

  const previous = state.formulaAutocomplete;
  const previousName =
    previous?.candidates?.[previous.selectedIndex]?.name ||
    previous?.candidates?.[0]?.name ||
    "";
  let selectedIndex = 0;
  if (options.keepSelection && previousName) {
    const match = candidates.findIndex((candidate) => candidate.name === previousName);
    if (match >= 0) selectedIndex = match;
  }
  state.formulaAutocomplete = {
    editor,
    queryContext,
    hintContext,
    candidates,
    selectedIndex,
    hint,
  };
  renderFormulaAutocomplete();
}



function formulaAutocompleteShouldShowCandidates(context) {
  if (!context) return false;
  if (context.type === "identifier" && !String(context.query || "").trim()) return false;
  return true;
}



function scheduleFormulaAutocompleteRefreshForEditor(editor, options = {}) {
  if (!editor || state.suppressFormulaAutocompleteInput) return;
  const key = String(options.event?.key || "");
  if (["Shift", "Control", "Alt", "Meta", "CapsLock"].includes(key)) return;
  requestAnimationFrame(() => {
    if (!editor.isConnected || isImeComposingForEditor(editor)) return;
    updateFormulaAutocompleteForEditor(editor, { keepSelection: options.keepSelection !== false });
  });
}



function formulaAutocompleteEditorContext(editor) {
  if (!editor) return {};
  if (editor === $formulaInput[0]) {
    const point = state.formulaEdit || state.selected || {};
    return {
      sheetIndex: point.sheetIndex ?? state.activeSheetIndex,
      row: point.row,
      col: point.col,
    };
  }
  const point = cellPointFromElement(editor);
  return point || { sheetIndex: state.activeSheetIndex };
}



function isImeComposingForEditor(editor) {
  if (!editor) return false;
  return Boolean($(editor).data("imeComposing") || (editor === $formulaInput[0] ? state.formulaBarComposing : state.cellEditComposing));
}



function formulaTextStartsFormula(text) {
  return String(text ?? "").startsWith("=");
}



function formulaTextWithExcelAutoCompletion(text) {
  const value = String(text ?? "");
  if (!formulaTextStartsFormula(value)) return value;
  const missingCloseCount = formulaMissingClosingParenthesisCount(value);
  return missingCloseCount > 0 ? `${value}${")".repeat(missingCloseCount)}` : value;
}



function formulaMissingClosingParenthesisCount(text) {
  const value = String(text ?? "");
  let openParenCount = 0;
  let inString = false;
  let inQuotedSheetName = false;
  let arrayDepth = 0;
  let structuredDepth = 0;

  for (let index = 0; index < value.length; index += 1) {
    const char = value[index];
    if (inString) {
      if (char === '"' && value[index + 1] === '"') {
        index += 1;
      } else if (char === '"') {
        inString = false;
      }
      continue;
    }
    if (inQuotedSheetName) {
      if (char === "'" && value[index + 1] === "'") {
        index += 1;
      } else if (char === "'") {
        inQuotedSheetName = false;
      }
      continue;
    }
    if (char === '"') {
      inString = true;
      continue;
    }
    if (char === "'") {
      inQuotedSheetName = true;
      continue;
    }
    if (char === "{") {
      arrayDepth += 1;
      continue;
    }
    if (char === "}") {
      arrayDepth = Math.max(0, arrayDepth - 1);
      continue;
    }
    if (!arrayDepth && char === "[") {
      structuredDepth += 1;
      continue;
    }
    if (!arrayDepth && char === "]") {
      structuredDepth = Math.max(0, structuredDepth - 1);
      continue;
    }
    if (arrayDepth || structuredDepth) continue;
    if (char === "(") {
      openParenCount += 1;
    } else if (char === ")") {
      openParenCount = Math.max(0, openParenCount - 1);
    }
  }

  if (inString || inQuotedSheetName || arrayDepth || structuredDepth) return 0;
  return openParenCount;
}



function formulaAutocompleteCandidates(context) {
  if (context?.type === "structured") return formulaStructuredReferenceCandidates(context);
  if (context?.type === "sheet") return formulaSheetReferenceCandidates(context);
  const normalized = String(context?.query || "").toLocaleUpperCase("ja-JP");
  const definitions = formulaAutocompleteDefinitions();
  const matched = normalized
    ? definitions.filter((definition) => {
        const text = formulaAutocompleteMatchText(definition);
        return text.startsWith(normalized) || text.includes(normalized);
      })
    : definitions.slice(0, 14);
  return matched
    .sort((left, right) => {
      const leftStarts = formulaAutocompleteMatchText(left).startsWith(normalized) ? 0 : 1;
      const rightStarts = formulaAutocompleteMatchText(right).startsWith(normalized) ? 0 : 1;
      return leftStarts - rightStarts || (left.priority || 999) - (right.priority || 999) || formulaAutocompleteDisplayName(left).localeCompare(formulaAutocompleteDisplayName(right), "ja");
    })
    .slice(0, 12);
}



function formulaSheetReferenceCandidates(context) {
  const normalized = String(context?.query || "").toLocaleUpperCase("ja-JP");
  const rawNormalized = String(context?.rawQuery || context?.query || "").toLocaleUpperCase("ja-JP");
  return formulaWorkbookSheetReferenceDefinitions()
    .filter((definition) => {
      const text = formulaAutocompleteMatchText(definition);
      const syntax = String(definition?.syntax || "").toLocaleUpperCase("ja-JP");
      const insertText = String(definition?.insertText || "").toLocaleUpperCase("ja-JP");
      return !normalized ||
        text.startsWith(normalized) ||
        text.includes(normalized) ||
        (rawNormalized && (syntax.includes(rawNormalized) || insertText.includes(rawNormalized)));
    })
    .sort((left, right) => {
      const leftText = formulaAutocompleteMatchText(left);
      const rightText = formulaAutocompleteMatchText(right);
      const leftStarts = leftText.startsWith(normalized) ? 0 : 1;
      const rightStarts = rightText.startsWith(normalized) ? 0 : 1;
      return leftStarts - rightStarts || (left.priority || 999) - (right.priority || 999) || formulaAutocompleteDisplayName(left).localeCompare(formulaAutocompleteDisplayName(right), "ja");
    })
    .slice(0, 12);
}



function formulaAutocompleteQueryContext(text, caretOffset, editorContext = {}) {
  const value = String(text ?? "");
  const caret = Math.max(0, Math.min(value.length, Number(caretOffset) || 0));
  if (!formulaTextStartsFormula(value) || formulaCaretIsInsideString(value, caret)) return null;
  const structured = formulaStructuredReferenceContext(value, caret, editorContext);
  if (structured) return structured;
  const quotedSheet = formulaQuotedSheetNameContext(value, caret);
  if (quotedSheet) return quotedSheet;
  const before = value.slice(0, caret);
  const token = formulaIdentifierBeforeCaret(before);
  const query = token.text;
  const start = caret - query.length;
  const prefix = before.slice(0, start);
  const nextChar = value.charAt(caret);
  if (nextChar && formulaIdentifierChar(nextChar)) return null;
  if (query && /^\$?[A-Za-z]{1,3}\$?\d+$/.test(query)) return null;
  if (query && !/[=+\-*/^&,(;]\s*$/.test(prefix)) return null;
  if (!query && prefix.replace(/\s/g, "") !== "=") return null;
  return { type: "identifier", query, start, end: caret };
}



function formulaQuotedSheetNameContext(text, caret) {
  const before = String(text || "").slice(0, caret);
  for (let index = before.length - 1; index >= 0; index -= 1) {
    if (before[index] !== "'") continue;
    const prefix = before.slice(0, index);
    if (!/[=+\-*/^&,(;]\s*$/.test(prefix)) continue;
    const fragment = before.slice(index + 1);
    const parsed = formulaUnclosedQuotedSheetFragment(fragment);
    if (!parsed) return null;
    return {
      type: "sheet",
      query: parsed.query,
      rawQuery: fragment,
      start: index,
      end: caret,
      quoted: true,
    };
  }
  return null;
}



function formulaUnclosedQuotedSheetFragment(fragment) {
  const value = String(fragment || "");
  let query = "";
  for (let index = 0; index < value.length; index += 1) {
    const char = value[index];
    if (char === "!") return null;
    if (char === "'") {
      if (value[index + 1] === "'") {
        query += "'";
        index += 1;
        continue;
      }
      return null;
    }
    query += char;
  }
  return { query };
}



function formulaIdentifierBeforeCaret(before) {
  let index = String(before || "").length - 1;
  while (index >= 0 && formulaIdentifierChar(before[index])) index -= 1;
  const text = before.slice(index + 1);
  if (text && !formulaIdentifierStartChar(text[0])) return { text: "" };
  return { text };
}



function formulaIdentifierChar(char) {
  return Boolean(char && /[\p{L}\p{N}_.]/u.test(char));
}



function formulaIdentifierStartChar(char) {
  return Boolean(char && /[\p{L}_]/u.test(char));
}



function formulaStructuredReferenceContext(text, caret, editorContext = {}) {
  const before = String(text || "").slice(0, caret);
  const lastOpen = formulaStructuredReferenceLastOpenIndex(before);
  if (lastOpen < 0) return null;
  const namedContext = formulaNamedTableStructuredReferenceContext(before, caret, lastOpen);
  if (namedContext) return namedContext;
  return formulaCurrentTableStructuredReferenceContext(before, caret, lastOpen, editorContext);
}



function formulaStructuredReferenceLastOpenIndex(text) {
  const value = String(text || "");
  let depth = 0;
  let lastOpen = -1;
  for (let index = 0; index < value.length; index += 1) {
    const char = value[index];
    if (char === "'" && index + 1 < value.length) {
      index += 1;
      continue;
    }
    if (char === "[") {
      depth += 1;
      lastOpen = index;
      continue;
    }
    if (char === "]") {
      if (depth > 0) depth -= 1;
      if (depth === 0) lastOpen = -1;
    }
  }
  return depth > 0 ? lastOpen : -1;
}



function formulaNamedTableStructuredReferenceContext(before, caret, lastOpen) {
  const refs = workbookFormulaTableRefs()
    .map((entry) => ({ entry, index: formulaLastIndexOfCaseInsensitive(before, entry.name) }))
    .filter((item) => item.index >= 0 && item.index + item.entry.name.length <= lastOpen)
    .sort((left, right) => right.index - left.index || right.entry.name.length - left.entry.name.length);
  for (const item of refs) {
    const tableRef = item.entry;
    const index = item.index;
    const boundary = before[index - 1] || "=";
    if (!/[=+\-*/^&,(;]/.test(boundary)) continue;
    const afterName = before.slice(index + tableRef.name.length);
    const firstOpenOffset = afterName.indexOf("[");
    if (firstOpenOffset < 0) continue;
    const firstOpen = index + tableRef.name.length + firstOpenOffset;
    if (firstOpen > lastOpen) continue;
    const fragment = before.slice(firstOpen, caret);
    if (!formulaNamedStructuredReferenceFragmentIsOpen(fragment)) continue;
    const currentRow = /^\[@/.test(fragment);
    let start = lastOpen + 1;
    if (currentRow && firstOpen === lastOpen) start = firstOpen + 2;
    const nestedContext = formulaStructuredReferenceNestedContext(fragment, lastOpen - firstOpen);
    return {
      type: "structured",
      query: before.slice(start),
      start,
      end: caret,
      tableRef,
      tableOpenIndex: index + tableRef.name.length,
      currentRow,
      columnWrapped: /\[@\[[^\]]*$/u.test(fragment),
      ...nestedContext,
      suppressSpecialItems: formulaStructuredReferenceColumnPartIsOpen(fragment, lastOpen - firstOpen),
    };
  }
  return null;
}



function formulaNamedStructuredReferenceFragmentIsOpen(fragment) {
  const text = String(fragment || "");
  if (!text.startsWith("[")) return false;
  let depth = 0;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (char === "'" && index + 1 < text.length) {
      index += 1;
      continue;
    }
    if (char === "[") {
      depth += 1;
      continue;
    }
    if (char === "]") {
      depth -= 1;
      if (depth <= 0 && index < text.length - 1) return false;
    }
  }
  return depth > 0;
}



function formulaStructuredReferenceColumnPartIsOpen(fragment, lastOpenOffset) {
  const text = String(fragment || "").slice(0, Math.max(0, Number(lastOpenOffset) || 0));
  return /\],\s*$/.test(text) || /:\s*$/.test(text);
}



function formulaStructuredReferenceRangeStartColumn(prefix) {
  const text = String(prefix || "").trim();
  if (!text.startsWith("[")) return "";
  const inner = text.slice(1).trimStart();
  const first = structuredReferenceTokenAt(inner, 0);
  if (!first) return "";
  const afterToken = inner.slice(first.end).trim();
  if (afterToken !== ":") return "";
  return cleanStructuredReferencePart(first.text.slice(1, -1));
}



function formulaStructuredReferencePrefixContext(prefix) {
  const text = String(prefix || "").trim();
  if (!text.startsWith("[")) return {};
  const inner = text.slice(1).trimStart();
  if (inner.startsWith("@")) {
    const afterAt = inner.slice(1).trimStart();
    const first = structuredReferenceTokenAt(afterAt, 0);
    if (!first) return {};
    const afterFirst = afterAt.slice(first.end).trim();
    if (afterFirst !== ":") return {};
    return { rangeStartColumn: cleanStructuredReferencePart(first.text.slice(1, -1)) };
  }
  const rowQualifiers = [];
  let cursor = 0;
  let firstColumn = "";
  while (cursor < inner.length) {
    const token = structuredReferenceTokenAt(inner, cursor);
    if (!token) break;
    const qualifier = formulaStructuredReferenceRowQualifier(token.text);
    const tokenValue = cleanStructuredReferencePart(token.text.slice(1, -1));
    const afterToken = inner.slice(token.end).trimStart();
    if (qualifier && afterToken.startsWith(",")) {
      rowQualifiers.push(qualifier);
      cursor = token.end + inner.slice(token.end).length - afterToken.length + 1 + afterToken.slice(1).match(/^\s*/)?.[0].length;
      continue;
    }
    firstColumn = tokenValue;
    if (afterToken === ":") {
      return {
        rowQualifier: rowQualifiers[0] || "",
        rowQualifiers,
        rangeStartColumn: firstColumn,
      };
    }
    if (afterToken.startsWith(",")) {
      return {
        rowQualifier: rowQualifiers[0] || "",
        rowQualifiers,
      };
    }
    break;
  }
  if (rowQualifiers.length) return { rowQualifier: rowQualifiers[0], rowQualifiers };
  return {};
}



function formulaStructuredReferenceNestedContext(fragment, lastOpenOffset) {
  const prefix = String(fragment || "").slice(0, Math.max(0, Number(lastOpenOffset) || 0));
  if (/^\[\s*$/u.test(prefix)) {
    return { nestedSpecialItem: true };
  }
  const context = formulaStructuredReferencePrefixContext(prefix);
  if (context.rowQualifier || context.rangeStartColumn) {
    return context;
  }
  return {};
}



function formulaStructuredReferenceRowQualifier(part) {
  if (structuredReferencePartHasEscapedSpecial(part)) return "";
  const normalized = formulaStructuredReferenceNormalizeSpecialItem(cleanStructuredReferencePart(part));
  return ["#All", "#Data", "#Headers", "#Totals", "#This Row"].includes(normalized) ? normalized : "";
}



function formulaStructuredReferenceNormalizeSpecialItem(item) {
  const normalized = String(item || "").trim().replace(/\s+/g, " ").toLocaleLowerCase("ja-JP");
  if (normalized === "#all") return "#All";
  if (normalized === "#data") return "#Data";
  if (normalized === "#headers") return "#Headers";
  if (normalized === "#totals") return "#Totals";
  if (normalized === "#this row") return "#This Row";
  return String(item || "").trim();
}



function formulaCurrentTableStructuredReferenceContext(before, caret, lastOpen, editorContext = {}) {
  const tableRef = formulaCurrentTableRefForEditorContext(editorContext);
  if (!tableRef) return null;
  const firstOpen = formulaCurrentStructuredReferenceFirstOpen(before, lastOpen);
  if (firstOpen < 0) return null;
  const boundary = before[firstOpen - 1] || "=";
  if (!/[=+\-*/^&,(;]/.test(boundary)) return null;
  const fragment = before.slice(firstOpen, caret);
  if (!formulaCurrentStructuredReferenceFragmentIsOpen(fragment)) return null;
  let start = lastOpen + 1;
  let currentRow = false;
  if (fragment.startsWith("[@[")) {
    currentRow = true;
  } else if (fragment.startsWith("[@")) {
    currentRow = true;
    if (firstOpen === lastOpen) {
      start = firstOpen + 2;
    }
  }
  const nestedContext = formulaStructuredReferenceNestedContext(fragment, lastOpen - firstOpen);
  return {
    type: "structured",
    query: before.slice(start),
    start,
    end: caret,
    tableRef,
    tableOpenIndex: firstOpen,
    currentRow,
    columnWrapped: /\[@\[[^\]]*$/u.test(fragment),
    currentTable: true,
    ...nestedContext,
    suppressSpecialItems: formulaStructuredReferenceColumnPartIsOpen(fragment, lastOpen - firstOpen),
  };
}



function formulaCurrentTableRefForEditorContext(editorContext = {}) {
  const sheetIndex = Number(editorContext.sheetIndex ?? state.activeSheetIndex);
  const row = Number(editorContext.row);
  const col = Number(editorContext.col);
  if (!Number.isFinite(row) || !Number.isFinite(col)) return null;
  return workbookFormulaTableRefs().find((entry) => {
    if (entry.sheetIndex !== sheetIndex) return false;
    const range = entry.table?.range;
    return range && row >= range.top && row <= range.bottom && col >= range.left && col <= range.right;
  }) || null;
}



function formulaCurrentStructuredReferenceFirstOpen(before, lastOpen) {
  const text = String(before || "");
  for (let index = lastOpen; index >= 0; index -= 1) {
    if (text[index] !== "[") continue;
    const fragment = text.slice(index, lastOpen + 1);
    if (/^\[@/.test(fragment) || /^\[\[/.test(fragment)) {
      return index;
    }
  }
  return lastOpen;
}



function formulaCurrentStructuredReferenceFragmentIsOpen(fragment) {
  const text = String(fragment || "");
  if (!text.startsWith("[")) return false;
  return formulaNamedStructuredReferenceFragmentIsOpen(text);
}



function formulaLastIndexOfCaseInsensitive(text, search) {
  return String(text || "").toLocaleUpperCase("ja-JP").lastIndexOf(String(search || "").toLocaleUpperCase("ja-JP"));
}



function formulaStructuredReferenceCandidates(context) {
  const rawNormalized = String(context?.query || "").toLocaleUpperCase("ja-JP");
  const normalized = unescapeStructuredReferencePart(context?.query || "").toLocaleUpperCase("ja-JP");
  const specialItems = (context?.currentRow || context?.suppressSpecialItems ? [] : [
    { name: "#All", description: "テーブル全体" },
    { name: "#Data", description: "データ行" },
    { name: "#Headers", description: "見出し行" },
    { name: "#Totals", description: "集計行" },
    { name: "#This Row", description: "現在の行" },
  ]).map((item, index) => ({
    kind: "structured",
    name: item.name,
    label: item.name,
    category: "構造化参照",
    syntax: formulaStructuredSpecialItemSyntax(item.name, context),
    description: item.description,
    priority: index + 1,
    insertText: item.name,
  }));
  const columns = (context?.tableRef?.columns || []).map((name, index) => ({
    kind: "structured",
    name,
    label: name,
    category: "テーブル列",
    syntax: formulaStructuredReferenceCandidateSyntax(name, context),
    description: formulaStructuredReferenceCandidateDescription(name, context),
    priority: 20 + index,
    insertText: formulaStructuredReferenceCandidateInsertText(name, context),
  }));
  const visibleLimit = Math.max(24, Math.min(80, specialItems.length + columns.length));
  const candidateMatches = (candidate) => {
    if (!normalized && !rawNormalized) return true;
    const text = formulaAutocompleteMatchText(candidate);
    const syntax = String(candidate?.syntax || "").toLocaleUpperCase("ja-JP");
    const insertText = String(candidate?.insertText || "").toLocaleUpperCase("ja-JP");
    return Boolean(
      (normalized && (text.startsWith(normalized) || text.includes(normalized))) ||
      (rawNormalized && (syntax.includes(rawNormalized) || insertText.includes(rawNormalized))),
    );
  };
  const candidateStarts = (candidate) => {
    if (!normalized && !rawNormalized) return true;
    const text = formulaAutocompleteMatchText(candidate);
    const insertText = String(candidate?.insertText || "").toLocaleUpperCase("ja-JP");
    return Boolean(
      (normalized && text.startsWith(normalized)) ||
      (rawNormalized && insertText.startsWith(rawNormalized)),
    );
  };
  return [...specialItems, ...columns]
    .filter(candidateMatches)
    .sort((left, right) => {
      const leftStarts = candidateStarts(left) ? 0 : 1;
      const rightStarts = candidateStarts(right) ? 0 : 1;
      return leftStarts - rightStarts || (left.priority || 999) - (right.priority || 999) || formulaAutocompleteDisplayName(left).localeCompare(formulaAutocompleteDisplayName(right), "ja");
    })
    .slice(0, visibleLimit);
}



function formulaStructuredReferenceCandidateSyntax(name, context) {
  const columnText = escapeStructuredReferenceColumnName(name);
  const tablePrefix = context?.currentTable || !context?.tableRef?.name ? "" : context.tableRef.name;
  if (context?.currentRow && context?.rangeStartColumn) {
    return `${tablePrefix}[@[${escapeStructuredReferenceColumnName(context.rangeStartColumn)}]:[${columnText}]]`;
  }
  if (context?.rowQualifier && context?.rangeStartColumn) {
    const rowItems = formulaStructuredReferenceRowQualifierSyntaxItems(context);
    return `${tablePrefix}[${rowItems}[${escapeStructuredReferenceColumnName(context.rangeStartColumn)}]:[${columnText}]]`;
  }
  if (context?.rowQualifier) {
    const rowItems = formulaStructuredReferenceRowQualifierSyntaxItems(context);
    return `${tablePrefix}[${rowItems}[${columnText}]]`;
  }
  if (context?.rangeStartColumn) {
    return `${tablePrefix}[[${escapeStructuredReferenceColumnName(context.rangeStartColumn)}]:[${columnText}]]`;
  }
  if (context?.currentRow) {
    const body = structuredReferenceColumnNeedsInnerBrackets(name) ? `[@[${columnText}]]` : `[@${columnText}]`;
    return `${tablePrefix}${body}`;
  }
  const body = structuredReferenceColumnNeedsOuterBrackets(name) ? `[[${columnText}]]` : `[${columnText}]`;
  return `${tablePrefix}${body}`;
}



function formulaStructuredReferenceRowQualifierSyntaxItems(context) {
  const qualifiers = Array.isArray(context?.rowQualifiers) && context.rowQualifiers.length
    ? context.rowQualifiers
    : [context?.rowQualifier].filter(Boolean);
  return qualifiers.map((item) => `[${item}],`).join("");
}



function formulaStructuredSpecialItemSyntax(name, context) {
  const item = String(name || "");
  const body = context?.nestedSpecialItem ? `[[${item}]]` : `[${item}]`;
  return context?.currentTable || !context?.tableRef?.name ? body : `${context.tableRef.name}${body}`;
}



function formulaStructuredReferenceCandidateDescription(name, context) {
  return formulaStructuredReferenceCandidateSyntax(name, context);
}



function formulaStructuredReferenceCandidateInsertText(name, context) {
  const columnText = escapeStructuredReferenceColumnName(name);
  if (!context?.currentRow) {
    if (!context?.rowQualifier && !context?.rangeStartColumn && structuredReferenceColumnNeedsOuterBrackets(name)) {
      return `[${columnText}]`;
    }
    return columnText;
  }
  if (context?.rangeStartColumn) return columnText;
  if (context.columnWrapped) return columnText;
  return structuredReferenceColumnNeedsInnerBrackets(name) ? `[${columnText}]` : columnText;
}



function structuredReferenceColumnNeedsInnerBrackets(name) {
  const text = String(name || "");
  return /[\s\[\]#'"@]/u.test(text) || structuredReferenceColumnNeedsOuterBrackets(text);
}



function structuredReferenceColumnNeedsOuterBrackets(name) {
  const text = String(name || "");
  return /[\t\n\r,.:"{}$^&*+=\-<>\/\\!()%?`;~_]/u.test(text);
}



function formulaCaretIsInsideString(text, caretOffset) {
  const value = String(text || "");
  let inString = false;
  for (let index = 0; index < Math.min(value.length, caretOffset); index += 1) {
    if (value[index] !== '"') continue;
    if (inString && value[index + 1] === '"') {
      index += 1;
      continue;
    }
    inString = !inString;
  }
  return inString;
}



function formulaParameterHintContext(text, caretOffset) {
  const value = String(text ?? "");
  const caret = Math.max(0, Math.min(value.length, Number(caretOffset) || 0));
  if (!formulaTextStartsFormula(value)) return null;
  const stack = [];
  let inString = false;
  let arrayDepth = 0;
  let structuredDepth = 0;
  for (let index = 0; index < caret; index += 1) {
    const char = value[index];
    if (char === '"') {
      if (inString && value[index + 1] === '"') {
        index += 1;
      } else {
        inString = !inString;
      }
      continue;
    }
    if (inString) continue;
    if (char === "{") {
      arrayDepth += 1;
      continue;
    }
    if (char === "}") {
      arrayDepth = Math.max(0, arrayDepth - 1);
      continue;
    }
    if (!arrayDepth && char === "[") {
      structuredDepth += 1;
      continue;
    }
    if (!arrayDepth && char === "]") {
      structuredDepth = Math.max(0, structuredDepth - 1);
      continue;
    }
    if (arrayDepth || structuredDepth) continue;
    if (char === "(") {
      const functionName = formulaNameBeforeOpeningParen(value, index);
      stack.push(functionName ? { name: functionName, argIndex: 0 } : null);
      continue;
    }
    if (char === ")") {
      stack.pop();
      continue;
    }
    if ((char === "," || char === ";") && stack.length) {
      const active = stack[stack.length - 1];
      if (active) active.argIndex += 1;
    }
  }
  for (let index = stack.length - 1; index >= 0; index -= 1) {
    const active = stack[index];
    const definition = active ? formulaAutocompleteDefinitionByName(active.name) : null;
    if (definition) return { definition, argIndex: active.argIndex };
  }
  return null;
}



function formulaNameBeforeOpeningParen(text, parenIndex) {
  let index = parenIndex - 1;
  while (index >= 0 && /\s/.test(text[index])) index -= 1;
  const end = index + 1;
  while (index >= 0 && /[A-Za-z0-9_.]/.test(text[index])) index -= 1;
  const name = text.slice(index + 1, end).toUpperCase();
  if (!name || !/^[A-Z_][A-Z0-9_.]*$/.test(name)) return "";
  if (index >= 0 && !/[=+\-*/^&,(;]/.test(text[index])) return "";
  return name;
}



function formulaParameterHintHtml(context) {
  const definition = context?.definition;
  if (!definition) return "";
  const specs = formulaParameterHintSpecs(definition, context.argIndex || 0);
  if (!specs.length) {
    return `<div class="formula-autocomplete-hint"><strong>${escapeHtml(definition.name)}</strong></div>`;
  }
  const activeIndex = Math.min(Math.max(0, context.argIndex || 0), Math.max(0, specs.length - 1));
  const activeDescription = formulaArgumentDescription(definition, specs[activeIndex], activeIndex);
  return `
    <div class="formula-autocomplete-hint">
      <span class="formula-autocomplete-hint-name">${escapeHtml(definition.name)}</span><span>(</span>${specs
        .map((spec, index) => `<span class="${index === activeIndex ? "is-active" : ""}">${escapeHtml(formulaArgumentHintLabel(spec))}</span>`)
        .join('<span class="formula-autocomplete-comma">, </span>')}<span>)</span>
      ${activeDescription ? `<div class="formula-autocomplete-hint-detail"><span class="formula-autocomplete-hint-argument">引数 ${activeIndex + 1}: </span>${escapeHtml(activeDescription)}</div>` : ""}
    </div>
  `;
}



function formulaParameterHintSpecs(definition, argIndex) {
  const minRows = Math.max(0, Number(argIndex) || 0) + 1;
  return functionArgumentDisplaySpecs(definition, [], { minRows });
}



function formulaArgumentHintLabel(spec) {
  const label = String(spec?.label || "");
  return spec?.optional ? `[${label}]` : label;
}



function isFormulaArgumentExpansionShortcut(event) {
  const key = String(event?.key || "").toLowerCase();
  return key === "a" && event.shiftKey && event.ctrlKey && !event.metaKey && !event.altKey;
}



function isFormulaFunctionArgumentsShortcut(event) {
  const key = String(event?.key || "").toLowerCase();
  return key === "a" && event.ctrlKey && !event.metaKey && !event.shiftKey && !event.altKey;
}



function expandFormulaFunctionArgumentsAtEditor(editor) {
  if (!editor || isImeComposingForEditor(editor)) return false;
  const text = formulaEditorText(editor);
  const selection = formulaEditorSelectionOffsets(editor);
  const context = formulaArgumentExpansionContext(text, selection);
  if (!context) return false;
  const definition = formulaAutocompleteDefinitionByName(context.name);
  if (!definition) return false;
  const body = formulaFunctionArgumentPlaceholderText(definition);
  const replacement = context.replaceFunctionName
    ? `${definition.name}(${body})`
    : `${body}${context.appendClosingParen ? ")" : ""}`;
  const nextText = `${text.slice(0, context.start)}${replacement}${text.slice(context.end)}`;
  const nextCaret = context.replaceFunctionName
    ? context.start + definition.name.length + 1
    : context.start;
  state.formulaAutocompleteAccepted = null;
  state.suppressFormulaAutocompleteInput = true;
  try {
    formulaEditorSetTextAndCaret(editor, nextText, nextCaret);
  } finally {
    state.suppressFormulaAutocompleteInput = false;
  }
  closeFormulaAutocomplete({ restoreFocus: false });
  if (editor === $formulaInput[0]) {
    syncFormulaEditPreview();
  } else {
    syncEditableCellAlignment(editor);
  }
  updateFormulaReferenceUi();
  requestAnimationFrame(() => updateFormulaAutocompleteForEditor(editor));
  return true;
}



function formulaArgumentExpansionContext(text, selection) {
  const value = String(text || "");
  const start = Math.max(0, Math.min(value.length, Number(selection?.start) || 0));
  const end = Math.max(0, Math.min(value.length, Number(selection?.end) || start));
  const caret = end;
  if (!formulaTextStartsFormula(value) || formulaCaretIsInsideString(value, caret)) return null;
  const before = value.slice(0, caret);
  const openParenMatch = before.match(/([A-Za-z_][A-Za-z0-9_.]*)\(\s*$/u);
  if (openParenMatch) {
    const name = openParenMatch[1].toUpperCase();
    if (!formulaAutocompleteDefinitionByName(name)) return null;
    return { name, start: caret, end: caret, replaceFunctionName: false, appendClosingParen: value.charAt(caret) !== ")" };
  }
  const identifier = formulaIdentifierBeforeCaret(before);
  const query = identifier.text;
  if (!query) return null;
  const tokenStart = caret - query.length;
  if (start !== end && start !== tokenStart) return null;
  const prefix = value.slice(0, tokenStart);
  const nextChar = value.charAt(caret);
  if (nextChar && formulaIdentifierChar(nextChar)) return null;
  if (!/[=+\-*/^&,(;]\s*$/u.test(prefix)) return null;
  const name = query.toUpperCase();
  if (!formulaAutocompleteDefinitionByName(name)) return null;
  return { name, start: tokenStart, end: caret, replaceFunctionName: true };
}



function formulaFunctionArgumentPlaceholderText(definition) {
  const specs = formulaFunctionArgumentSpecs(definition);
  if (!specs.length) return "";
  const labels = specs.map((spec) => String(spec.label || "").trim()).filter(Boolean);
  if (specs.some((spec) => spec.repeat)) labels.push("...");
  return labels.join(", ");
}



function openFormulaFunctionArgumentsDialogFromEditor(editor) {
  if (!editor || isImeComposingForEditor(editor)) return false;
  const context = formulaFunctionArgumentsEditorContext(editor);
  if (!context) return false;
  if (context.editor !== $formulaInput[0]) {
    $(context.editor).data("skipBlurCommit", true);
  }
  openFunctionArgumentsDialog(context.functionName, {
    argumentValues: context.argumentValues,
    editorSource: context,
  });
  return true;
}



function formulaFunctionArgumentsEditorContext(editor) {
  const text = formulaEditorText(editor);
  const selection = formulaEditorSelectionOffsets(editor);
  const caret = Math.max(0, Math.min(text.length, Number(selection?.end) || 0));
  if (!formulaTextStartsFormula(text) || formulaCaretIsInsideString(text, caret)) return null;
  const call = formulaFunctionCallAtCaret(text, caret) || formulaFunctionCallBeforeCaret(text, caret) || formulaFunctionIdentifierAtCaret(text, selection);
  if (!call || !formulaAutocompleteDefinitionByName(call.functionName)) return null;
  const point = editor === $formulaInput[0]
    ? (state.formulaEdit || state.selected || {})
    : cellPointFromElement(editor);
  return {
    editor,
    isFormulaBar: editor === $formulaInput[0],
    text,
    start: call.start,
    end: call.end,
    functionName: call.functionName,
    argumentValues: formulaArgumentValuesForDialog(call.functionName, call.argumentValues || []),
    sheetIndex: point?.sheetIndex ?? state.activeSheetIndex,
    row: point?.row,
    col: point?.col,
  };
}



function formulaFunctionIdentifierAtCaret(text, selection) {
  const value = String(text || "");
  const start = Math.max(0, Math.min(value.length, Number(selection?.start) || 0));
  const end = Math.max(0, Math.min(value.length, Number(selection?.end) || start));
  const before = value.slice(0, end);
  const identifier = formulaIdentifierBeforeCaret(before);
  const query = identifier.text;
  if (!query) return null;
  const tokenStart = end - query.length;
  if (start !== end && start !== tokenStart) return null;
  const prefix = value.slice(0, tokenStart);
  const nextChar = value.charAt(end);
  if (nextChar && formulaIdentifierChar(nextChar)) return null;
  if (!/[=+\-*/^&,(;]\s*$/u.test(prefix)) return null;
  const functionName = query.toUpperCase();
  if (!formulaAutocompleteDefinitionByName(functionName)) return null;
  return { functionName, start: tokenStart, end, argumentValues: [] };
}



function formulaFunctionCallAtCaret(text, caretOffset) {
  const value = String(text || "");
  const caret = Math.max(0, Math.min(value.length, Number(caretOffset) || 0));
  const stack = [];
  let inString = false;
  for (let index = 0; index < caret; index += 1) {
    const char = value[index];
    if (char === '"') {
      if (inString && value[index + 1] === '"') index += 1;
      else inString = !inString;
      continue;
    }
    if (inString) continue;
    if (char === "(") {
      const nameRange = formulaNameRangeBeforeOpeningParen(value, index);
      stack.push(nameRange ? { ...nameRange, openIndex: index } : null);
      continue;
    }
    if (char === ")") {
      stack.pop();
    }
  }
  for (let index = stack.length - 1; index >= 0; index -= 1) {
    const active = stack[index];
    if (!active?.functionName || !formulaAutocompleteDefinitionByName(active.functionName)) continue;
    const closeIndex = formulaMatchingCloseParenIndex(value, active.openIndex);
    const end = closeIndex >= 0 ? closeIndex + 1 : value.length;
    if (caret > end) continue;
    const bodyEnd = closeIndex >= 0 ? closeIndex : value.length;
    return {
      functionName: active.functionName,
      start: active.start,
      end,
      openIndex: active.openIndex,
      closeIndex,
      argumentValues: splitFormulaArgumentValues(value.slice(active.openIndex + 1, bodyEnd)),
    };
  }
  return null;
}



function formulaFunctionCallBeforeCaret(text, caretOffset) {
  const value = String(text || "");
  let index = Math.max(0, Math.min(value.length, Number(caretOffset) || 0)) - 1;
  while (index >= 0 && /\s/.test(value[index])) index -= 1;
  if (index < 0 || value[index] !== ")") return null;
  const openIndex = formulaMatchingOpenParenIndex(value, index);
  if (openIndex < 0) return null;
  const nameRange = formulaNameRangeBeforeOpeningParen(value, openIndex);
  if (!nameRange?.functionName || !formulaAutocompleteDefinitionByName(nameRange.functionName)) return null;
  return {
    functionName: nameRange.functionName,
    start: nameRange.start,
    end: index + 1,
    openIndex,
    closeIndex: index,
    argumentValues: splitFormulaArgumentValues(value.slice(openIndex + 1, index)),
  };
}



function formulaMatchingOpenParenIndex(text, closeIndex) {
  const value = String(text || "");
  const target = Math.max(0, Math.min(value.length - 1, Number(closeIndex) || 0));
  const stack = [];
  let inString = false;
  for (let index = 0; index <= target; index += 1) {
    const char = value[index];
    if (char === '"') {
      if (inString && value[index + 1] === '"') index += 1;
      else inString = !inString;
      continue;
    }
    if (inString) continue;
    if (char === "(") {
      stack.push(index);
      continue;
    }
    if (char === ")") {
      const openIndex = stack.pop();
      if (index === target) return Number.isFinite(openIndex) ? openIndex : -1;
    }
  }
  return -1;
}



function formulaNameRangeBeforeOpeningParen(text, parenIndex) {
  let index = parenIndex - 1;
  while (index >= 0 && /\s/.test(text[index])) index -= 1;
  const end = index + 1;
  while (index >= 0 && /[A-Za-z0-9_.]/.test(text[index])) index -= 1;
  const start = index + 1;
  const functionName = text.slice(start, end).toUpperCase();
  if (!functionName || !/^[A-Z_][A-Z0-9_.]*$/.test(functionName)) return null;
  if (index >= 0 && !/[=+\-*/^&,(;]/.test(text[index])) return null;
  return { functionName, start, nameEnd: end };
}



function formulaMatchingCloseParenIndex(text, openIndex) {
  const value = String(text || "");
  let depth = 0;
  let inString = false;
  for (let index = Math.max(0, Number(openIndex) || 0); index < value.length; index += 1) {
    const char = value[index];
    if (char === '"') {
      if (inString && value[index + 1] === '"') index += 1;
      else inString = !inString;
      continue;
    }
    if (inString) continue;
    if (char === "(") depth += 1;
    if (char === ")") {
      depth -= 1;
      if (depth === 0) return index;
    }
  }
  return -1;
}



function splitFormulaArgumentValues(body) {
  const text = String(body || "");
  if (!text.trim()) return [];
  const parts = [];
  let current = "";
  let depth = 0;
  let arrayDepth = 0;
  let inString = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (char === '"') {
      current += char;
      if (inString && text[index + 1] === '"') {
        current += text[index + 1];
        index += 1;
      } else {
        inString = !inString;
      }
      continue;
    }
    if (!inString) {
      if (char === "{") arrayDepth += 1;
      if (char === "}") arrayDepth = Math.max(0, arrayDepth - 1);
      if (char === "(" || char === "[") depth += 1;
      if (char === ")" || char === "]") depth = Math.max(0, depth - 1);
      if ((char === "," || char === ";") && depth === 0 && arrayDepth === 0) {
        parts.push(current.trim());
        current = "";
        continue;
      }
    }
    current += char;
  }
  parts.push(current.trim());
  return parts;
}



function formulaArgumentValuesForDialog(functionName, values) {
  const definition = formulaFunctionByName(functionName);
  const specs = formulaFunctionArgumentSpecs(definition);
  const minimum = specs.length || values.length;
  const output = Array.from({ length: Math.max(minimum, values.length) }, (_item, index) => values[index] || "");
  return output;
}



function renderFormulaAutocomplete() {
  const autocomplete = state.formulaAutocomplete;
  if (!autocomplete?.editor?.isConnected) {
    closeFormulaAutocomplete({ restoreFocus: false });
    return;
  }
  let popup = document.getElementById("formulaAutocomplete");
  if (!popup) {
    popup = document.createElement("div");
    popup.id = "formulaAutocomplete";
    popup.className = "formula-autocomplete-popover";
    popup.setAttribute("role", "listbox");
    popup.setAttribute("aria-label", "数式オートコンプリート");
    document.body.appendChild(popup);
  }
  if (autocomplete.candidates.length) {
    popup.setAttribute("aria-activedescendant", `formulaAutocompleteItem${autocomplete.selectedIndex}`);
  } else {
    popup.removeAttribute("aria-activedescendant");
  }
  popup.classList.toggle("is-hint-only", autocomplete.candidates.length === 0);
  const candidatesHtml = autocomplete.candidates.length
    ? `<div class="formula-autocomplete-list">${autocomplete.candidates.map((candidate, index) => formulaAutocompleteItemHtml(candidate, index === autocomplete.selectedIndex, index)).join("")}</div>`
    : "";
  const detailHtml = autocomplete.candidates.length ? formulaAutocompleteActiveDetailHtml(autocomplete) : "";
  popup.innerHTML = `${candidatesHtml}${detailHtml}${autocomplete.hint || ""}`;
  positionFormulaAutocompletePopup(popup, autocomplete.editor);
  popup.querySelector(".formula-autocomplete-item.is-active")?.scrollIntoView({ block: "nearest" });
}



function formulaAutocompleteItemHtml(candidate, active, index) {
  const itemId = `formulaAutocompleteItem${index}`;
  const syntax = String(candidate?.syntax || "");
  const description = String(candidate?.description || "");
  return `
    <button id="${escapeAttr(itemId)}" type="button" class="formula-autocomplete-item${active ? " is-active" : ""}" role="option" aria-selected="${active ? "true" : "false"}" data-formula-autocomplete-item="${index}">
      <span class="formula-autocomplete-name">${escapeHtml(formulaAutocompleteDisplayName(candidate))}</span>
      <span class="formula-autocomplete-category">${escapeHtml(candidate.category || "")}</span>
      <span class="formula-autocomplete-detail">
        <span class="formula-autocomplete-syntax">${escapeHtml(syntax)}</span>
        ${description ? `<span class="formula-autocomplete-description">${escapeHtml(description)}</span>` : ""}
      </span>
    </button>
  `;
}



function formulaAutocompleteActiveDetailHtml(autocomplete) {
  const candidate = autocomplete?.candidates?.[autocomplete.selectedIndex];
  if (!candidate) return "";
  const name = formulaAutocompleteDisplayName(candidate);
  const syntax = String(candidate.syntax || "");
  const description = String(candidate.description || "");
  const category = String(candidate.category || "");
  return `
    <div class="formula-autocomplete-active-detail" data-formula-autocomplete-active-detail>
      <div class="formula-autocomplete-active-heading">
        <span class="formula-autocomplete-active-name">${escapeHtml(name)}</span>
        ${category ? `<span class="formula-autocomplete-active-category">${escapeHtml(category)}</span>` : ""}
      </div>
      ${syntax ? `<div class="formula-autocomplete-active-syntax">${escapeHtml(syntax)}</div>` : ""}
      ${description ? `<div class="formula-autocomplete-active-description">${escapeHtml(description)}</div>` : ""}
    </div>
  `;
}



function positionFormulaAutocompletePopup(popup, editor) {
  const editorRect = editor.getBoundingClientRect();
  const anchorRect = formulaAutocompleteAnchorRect(editor) || editorRect;
  const width = Math.min(460, Math.max(280, editorRect.width || 320));
  const height = Math.min(320, popup.scrollHeight || 220);
  let left = Math.max(8, Math.min(window.innerWidth - width - 8, anchorRect.left));
  let top = anchorRect.bottom + 4;
  if (top + height > window.innerHeight - 8) {
    top = Math.max(8, anchorRect.top - height - 4);
  }
  popup.style.left = `${Math.round(left)}px`;
  popup.style.top = `${Math.round(top)}px`;
  popup.style.width = `${Math.round(width)}px`;
}



function formulaAutocompleteAnchorRect(editor) {
  if (!editor?.isConnected) return null;
  if (editor === $formulaInput[0] || editor.tagName === "TEXTAREA" || editor.tagName === "INPUT") {
    return editor.getBoundingClientRect();
  }
  if (editor.getAttribute("contenteditable") === "true") {
    return formulaContentEditableCaretRect(editor) || editor.getBoundingClientRect();
  }
  return editor.getBoundingClientRect();
}



function formulaContentEditableCaretRect(editor) {
  const selection = window.getSelection?.();
  if (!selection?.rangeCount || !selection.anchorNode || !editor?.contains(selection.anchorNode)) return null;
  const range = selection.getRangeAt(0);
  if (!editor.contains(range.startContainer) || !editor.contains(range.endContainer)) return null;
  const rect = usableFormulaAutocompleteRect(range.getClientRects?.()[0]) || usableFormulaAutocompleteRect(range.getBoundingClientRect?.());
  if (rect) return rect;

  const markerRange = range.cloneRange();
  markerRange.collapse(false);
  const marker = document.createElement("span");
  marker.textContent = "\u200b";
  marker.style.display = "inline-block";
  marker.style.width = "0";
  marker.style.height = "1em";
  marker.style.overflow = "hidden";
  try {
    markerRange.insertNode(marker);
    const markerRect = usableFormulaAutocompleteRect(marker.getBoundingClientRect());
    marker.remove();
    selection.removeAllRanges();
    selection.addRange(range);
    return markerRect;
  } catch {
    marker.remove();
    return null;
  }
}



function usableFormulaAutocompleteRect(rect) {
  if (!rect) return null;
  const left = Number(rect.left);
  const top = Number(rect.top);
  if (!Number.isFinite(left) || !Number.isFinite(top)) return null;
  const height = Number(rect.height);
  return {
    left,
    right: Number.isFinite(Number(rect.right)) ? Number(rect.right) : left,
    top,
    bottom: Number.isFinite(Number(rect.bottom)) && Number(rect.bottom) > top
      ? Number(rect.bottom)
      : top + Math.max(14, Number.isFinite(height) ? height : 14),
  };
}



function handleFormulaAutocompleteKeydown(event, editor) {
  if (isEnterKeyEvent(event) && isAltModifierEvent(event)) return false;
  const scope = editor === $formulaInput[0] ? "formula" : "cell";
  if (isImeCompositionEvent(event, editor, scope)) return false;
  if (isFormulaFunctionArgumentsShortcut(event) && openFormulaFunctionArgumentsDialogFromEditor(editor)) {
    event.preventDefault();
    event.stopPropagation();
    return true;
  }
  if (isFormulaArgumentExpansionShortcut(event) && expandFormulaFunctionArgumentsAtEditor(editor)) {
    event.preventDefault();
    event.stopPropagation();
    return true;
  }
  let autocomplete = state.formulaAutocomplete;
  if (autocomplete && !document.getElementById("formulaAutocomplete")) {
    state.formulaAutocomplete = null;
    autocomplete = null;
  }
  if (!autocomplete || autocomplete.editor !== editor) {
    if (event.key?.length === 1 || event.key === "Backspace" || event.key === "Delete") {
      scheduleFormulaAutocompleteRefreshForEditor(editor, { keepSelection: true, event });
    }
    return false;
  }
  if (
    event.key === "(" &&
    autocomplete.queryContext?.type === "identifier" &&
    String(autocomplete.queryContext.query || "").trim() &&
    formulaAutocompleteSelectedCandidate(autocomplete)?.kind === "function"
  ) {
    event.preventDefault();
    event.stopPropagation();
    acceptFormulaAutocompleteSelection();
    return true;
  }
  if (event.key === "ArrowDown" && autocomplete.candidates.length) {
    event.preventDefault();
    event.stopPropagation();
    moveFormulaAutocompleteSelection(1);
    return true;
  }
  if (event.key === "ArrowUp" && autocomplete.candidates.length) {
    event.preventDefault();
    event.stopPropagation();
    moveFormulaAutocompleteSelection(-1);
    return true;
  }
  if (event.key === "PageDown" && autocomplete.candidates.length) {
    event.preventDefault();
    event.stopPropagation();
    moveFormulaAutocompleteSelection(6);
    return true;
  }
  if (event.key === "PageUp" && autocomplete.candidates.length) {
    event.preventDefault();
    event.stopPropagation();
    moveFormulaAutocompleteSelection(-6);
    return true;
  }
  if (event.key === "Home" && autocomplete.candidates.length) {
    event.preventDefault();
    event.stopPropagation();
    setFormulaAutocompleteSelection(0);
    return true;
  }
  if (event.key === "End" && autocomplete.candidates.length) {
    event.preventDefault();
    event.stopPropagation();
    setFormulaAutocompleteSelection(autocomplete.candidates.length - 1);
    return true;
  }
  if ((event.key === "Enter" || event.key === "Tab") && autocomplete.candidates.length) {
    event.preventDefault();
    event.stopPropagation();
    acceptFormulaAutocompleteSelection();
    return true;
  }
  if (event.key === "Escape") {
    event.preventDefault();
    event.stopPropagation();
    closeFormulaAutocomplete({ rememberClosedState: true });
    return true;
  }
  scheduleFormulaAutocompleteRefreshForEditor(editor, { keepSelection: true, event });
  return false;
}



function formulaAutocompleteSelectedCandidate(autocomplete = state.formulaAutocomplete) {
  if (!autocomplete?.candidates?.length) return null;
  const index = Math.max(0, Math.min(autocomplete.candidates.length - 1, Number(autocomplete.selectedIndex) || 0));
  return autocomplete.candidates[index] || null;
}



function moveFormulaAutocompleteSelection(delta) {
  const autocomplete = state.formulaAutocomplete;
  if (!autocomplete?.candidates?.length) return;
  const count = autocomplete.candidates.length;
  setFormulaAutocompleteSelection((autocomplete.selectedIndex + delta + count) % count);
}



function setFormulaAutocompleteSelection(index) {
  const autocomplete = state.formulaAutocomplete;
  if (!autocomplete?.candidates?.length) return;
  autocomplete.selectedIndex = Math.max(0, Math.min(autocomplete.candidates.length - 1, Number(index) || 0));
  renderFormulaAutocomplete();
}



function acceptFormulaAutocompleteSelection(index = null) {
  const autocomplete = state.formulaAutocomplete;
  if (!autocomplete?.candidates?.length || !autocomplete.queryContext) return false;
  const selectedIndex = Number.isFinite(Number(index)) ? Number(index) : autocomplete.selectedIndex;
  const candidate = autocomplete.candidates[Math.max(0, Math.min(autocomplete.candidates.length - 1, selectedIndex))];
  if (!candidate) return false;
  const editor = autocomplete.editor;
  const text = formulaEditorText(editor);
  const { start, end } = autocomplete.queryContext;
  const { insertText, suffix, reopen } = formulaAutocompleteReplacement(candidate, autocomplete.queryContext, text);
  const nextText = `${text.slice(0, start)}${insertText}${suffix}${text.slice(end)}`;
  const nextCaret = start + insertText.length;
  state.formulaAutocompleteAccepted = reopen ? null : { editor, text: nextText, caret: nextCaret };
  state.suppressFormulaAutocompleteInput = true;
  try {
    formulaEditorSetTextAndCaret(editor, nextText, nextCaret);
  } finally {
    state.suppressFormulaAutocompleteInput = false;
  }
  closeFormulaAutocomplete({ restoreFocus: false });
  if (editor === $formulaInput[0]) {
    syncFormulaEditPreview();
  } else {
    syncEditableCellAlignment(editor);
  }
  if (candidate.kind === "function") recordFormulaFunctionUse(candidate.name);
  if (reopen) requestAnimationFrame(() => updateFormulaAutocompleteForEditor(editor));
  return true;
}



function formulaAutocompleteReplacement(candidate, context, text) {
  if (context?.type === "structured") {
    const insertText = String(candidate.insertText || candidate.name || "");
    return {
      insertText,
      suffix: formulaStructuredClosingSuffix(text, context),
      reopen: false,
    };
  }
  if (candidate?.kind === "function") {
    const nextChar = String(text || "").charAt(context.end);
    return {
      insertText: `${candidate.name}${nextChar === "(" ? "" : "("}`,
      suffix: "",
      reopen: true,
    };
  }
  return {
    insertText: String(candidate?.insertText || candidate?.name || ""),
    suffix: "",
    reopen: false,
  };
}



function formulaStructuredClosingSuffix(text, context) {
  const value = String(text || "");
  if (value.charAt(context.end) === "]") return "";
  const fragment = value.slice(context.tableOpenIndex, context.end);
  let depth = 0;
  for (let index = 0; index < fragment.length; index += 1) {
    const char = fragment[index];
    if (char === "'" && index + 1 < fragment.length) {
      index += 1;
      continue;
    }
    if (char === "[") depth += 1;
    if (char === "]") depth = Math.max(0, depth - 1);
  }
  return "]".repeat(Math.max(1, depth));
}



function handleFormulaAutocompleteItemClick(event) {
  const index = Number(event.currentTarget.dataset.formulaAutocompleteItem);
  acceptFormulaAutocompleteSelection(index);
}

export {
  acceptFormulaAutocompleteSelection,
  expandFormulaFunctionArgumentsAtEditor,
  formulaArgumentExpansionContext,
  formulaArgumentHintLabel,
  formulaArgumentValuesForDialog,
  formulaAutocompleteActiveDetailHtml,
  formulaAutocompleteAnchorRect,
  formulaAutocompleteCandidates,
  formulaAutocompleteDefinitionByName,
  formulaAutocompleteDefinitions,
  formulaAutocompleteDisplayName,
  formulaAutocompleteEditorContext,
  formulaAutocompleteItemHtml,
  formulaAutocompleteMatchText,
  formulaAutocompleteQueryContext,
  formulaAutocompleteReplacement,
  formulaAutocompleteSelectedCandidate,
  formulaAutocompleteShouldExposeEngineFunction,
  formulaAutocompleteShouldShowCandidates,
  formulaCaretIsInsideString,
  formulaContentEditableCaretRect,
  formulaCurrentStructuredReferenceFirstOpen,
  formulaCurrentStructuredReferenceFragmentIsOpen,
  formulaCurrentTableRefForEditorContext,
  formulaCurrentTableStructuredReferenceContext,
  formulaDefaultSyntax,
  formulaDefaultTemplate,
  formulaDefaultTemplateArgument,
  formulaEngineArgumentIsOptional,
  formulaEngineArgumentName,
  formulaEngineFunctionCategory,
  formulaEngineFunctionDescription,
  formulaEngineFunctionIsCompatibility,
  formulaEngineFunctionMetadata,
  formulaEngineFunctionSyntax,
  formulaFunctionArgumentPlaceholderText,
  formulaFunctionArgumentsEditorContext,
  formulaFunctionCallAtCaret,
  formulaFunctionCallBeforeCaret,
  formulaFunctionIdentifierAtCaret,
  formulaIdentifierBeforeCaret,
  formulaIdentifierChar,
  formulaIdentifierStartChar,
  formulaInsertFunctionDefinitionByName,
  formulaInsertFunctionDefinitions,
  formulaLastIndexOfCaseInsensitive,
  formulaMatchingCloseParenIndex,
  formulaMatchingOpenParenIndex,
  formulaMissingClosingParenthesisCount,
  formulaNameBeforeOpeningParen,
  formulaNameRangeBeforeOpeningParen,
  formulaNamedStructuredReferenceFragmentIsOpen,
  formulaNamedTableStructuredReferenceContext,
  formulaParameterHintContext,
  formulaParameterHintHtml,
  formulaParameterHintSpecs,
  formulaQuotedSheetNameContext,
  formulaReferenceColorClass,
  formulaSheetReferenceCandidates,
  formulaStructuredClosingSuffix,
  formulaStructuredReferenceCandidateDescription,
  formulaStructuredReferenceCandidateInsertText,
  formulaStructuredReferenceCandidateSyntax,
  formulaStructuredReferenceCandidates,
  formulaStructuredReferenceColumnPartIsOpen,
  formulaStructuredReferenceContext,
  formulaStructuredReferenceLastOpenIndex,
  formulaStructuredReferenceNestedContext,
  formulaStructuredReferenceNormalizeSpecialItem,
  formulaStructuredReferencePrefixContext,
  formulaStructuredReferenceRangeStartColumn,
  formulaStructuredReferenceRowQualifier,
  formulaStructuredReferenceRowQualifierSyntaxItems,
  formulaStructuredSpecialItemSyntax,
  formulaTableColumnNames,
  formulaTextStartsFormula,
  formulaTextWithExcelAutoCompletion,
  formulaUnclosedQuotedSheetFragment,
  formulaWorkbookAutocompleteDefinitions,
  formulaWorkbookSheetReferenceDefinitions,
  handleFormulaAutocompleteItemClick,
  handleFormulaAutocompleteKeydown,
  insertFunctionCategoryOptions,
  isFormulaArgumentExpansionShortcut,
  isFormulaFunctionArgumentsShortcut,
  isImeComposingForEditor,
  moveFormulaAutocompleteSelection,
  openFormulaFunctionArgumentsDialogFromEditor,
  positionFormulaAutocompletePopup,
  renderFormulaAutocomplete,
  scheduleFormulaAutocompleteRefreshForEditor,
  setFormulaAutocompleteSelection,
  splitFormulaArgumentValues,
  structuredReferenceColumnNeedsInnerBrackets,
  structuredReferenceColumnNeedsOuterBrackets,
  updateFormulaAutocompleteForEditor,
  usableFormulaAutocompleteRect,
  workbookFormulaTableRefs,
};
