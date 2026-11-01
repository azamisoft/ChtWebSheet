// Feature module; shared bindings are imported explicitly.
import {
  $horizontalScrollbar,
  $sheetHost,
  createIcons,
  icons,
  state,
} from "../../app-core.js";
import {
  coerceDelimitedImportValue,
  queryTransformStepLabel,
  renderQueryTransformEditorDialog,
} from "./query-editor.js";
import {
  closeDataCommandMenu,
  odataImportRowsFromSource,
  openWebNavigatorTransformEditorFromSource,
  padQueryTransformRow,
  queryTransformNormalizeDataType,
  rememberRecentDataSourceFromQuery,
  rowsFromObjects,
  sheetDisplayRowsForRange,
  uniqueQueryTransformColumns,
  webImportRowsFromSource,
  workbookQueryRowsWithTransforms,
} from "./query-transforms.js";
import {
  connectionStringUrlForSourceType,
  detectDelimitedTextDelimiter,
  parseDelimitedRows,
  parsePrefixedConnectionString,
  parseQueryConnectionString,
  queryRefreshDurationLabel,
  rangeHasUserData,
  readableRefreshError,
  reapplyAutoFilter,
  refreshWorkbookData,
  refreshWorkbookQueryWithRemoteSource,
  scheduleQueryRefreshTimers,
  setConnectionStringValue,
  showDataOverwriteConfirmation,
} from "./sort-and-autofilter.js";
import {
  canCreateWorkbookRelationship,
  cloneRelationships,
  decodeWorkbookRange,
  relationshipSummaryLabel,
  workbookCellReference,
  workbookRelationships,
  workbookTableRefs,
} from "./validation-and-consolidation.js";
import {
  chartDataFromRange,
  chartDescriptor,
  closeDataDialog,
  pivotSummaryFromRange,
  uniqueSheetNameWithBase,
  writePivotSummaryAt,
  writePivotSummarySheet,
} from "../drawing/charts-and-insertion.js";
import {
  endingCellForPixelSize,
  fitObjectSize,
} from "../drawing/shape-rendering.js";
import {
  rangesEqualLoose,
} from "../editing/autofill.js";
import {
  uniqueWorksheetObjectId,
} from "../editing/ink-and-clipboard.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  coerceUserInput,
  getDisplayForCell,
  normalizeRange,
  rangeToLabel,
  selectionRangeFromPoints,
  updateCellModel,
  writeCellContent,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  rangesEqual,
  scrollRangeIntoView,
} from "../editing/selection-and-cell-text.js";
import {
  captureSheetForHistory,
  captureSheetsForHistory,
  cloneCellModel,
  cloneJsonLikeValue,
  internCellsInRange,
  pushHistory,
  pushSheetInsertHistory,
  writeClonedCellContent,
} from "../editing/structure-and-history.js";
import {
  rangesIntersect,
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  closeCellContextMenu,
  positionPopupMenu,
} from "../ui/annotations-and-cell-format.js";
import {
  accessKeyLabelHtml,
  cellKey,
  columnName,
  escapeAttr,
  escapeHtml,
  setStatus,
  showMessageBox,
} from "../ui/help-and-utilities.js";
import {
  consumeCommandMenuKey,
  enhanceCommandMenuAccessKeys,
} from "../ui/printing-and-home-tools.js";
import {
  applyExcelTableStyle,
  commitDataOperation,
  ensureSheetSize,
  nextTableName,
} from "../ui/ribbon-and-links.js";
import {
  closeSheetTabContextMenu,
  createBlankSheet,
  uniqueSheetName,
} from "../ui/sheet-protection-and-translation.js";
import {
  buildFormulaEngine,
  renderWorkbook,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  scheduleWorkbookPersistence,
} from "../workbook/import-model.js";
import {
  activeSheet,
  cellDisplayText,
  getSheetId,
} from "../workbook/saving-and-export.js";


function xmlRowsFromElements(elements) {
  return rowsFromObjects([...elements].map((element) => xmlElementToRowObject(element)));
}



function xmlElementToRowObject(element) {
  const row = {};
  [...(element?.attributes || [])].forEach((attr) => {
    row[`@${attr.name}`] = attr.value;
  });
  const childGroups = new Map();
  [...(element?.children || [])].forEach((child) => {
    const name = child.localName || child.tagName;
    if (!childGroups.has(name)) childGroups.set(name, []);
    childGroups.get(name).push(child);
  });
  childGroups.forEach((children, name) => {
    if (children.length === 1) {
      const [child] = children;
      row[name] = child.children.length ? JSON.stringify(xmlElementToRowObject(child)) : xmlElementText(child);
    } else {
      row[name] = children.map((child) => child.children.length ? JSON.stringify(xmlElementToRowObject(child)) : xmlElementText(child)).join(", ");
    }
  });
  if (!childGroups.size) {
    const text = xmlElementText(element);
    if (text) row.Value = text;
  }
  return row;
}



function xmlElementText(element) {
  return String(element?.textContent ?? "").replace(/\s+/g, " ").trim();
}



function htmlTableSources(sourceText) {
  try {
    const doc = new DOMParser().parseFromString(sourceText, "text/html");
    return [...doc.querySelectorAll("table")]
      .map((table, index) => ({
        kind: "table",
        itemIndex: index,
        name: htmlTableName(table, index),
        rows: htmlTableRows(table),
      }))
      .filter((source) => source.rows.some((row) => row.some((value) => value !== "")));
  } catch (error) {
    console.warn("Failed to parse web table", error);
    return [];
  }
}



function htmlTableName(table, index) {
  const caption = table.querySelector("caption")?.textContent.replace(/\s+/g, " ").trim();
  if (caption) return caption;
  const ariaLabel = table.getAttribute("aria-label")?.replace(/\s+/g, " ").trim();
  if (ariaLabel) return ariaLabel;
  const heading = previousHtmlHeadingText(table);
  if (heading) return heading;
  const id = table.getAttribute("id")?.trim();
  return id || `Table ${index + 1}`;
}



function previousHtmlHeadingText(node) {
  let current = node?.previousElementSibling || null;
  while (current) {
    if (/^H[1-6]$/i.test(current.tagName || "")) {
      const text = current.textContent.replace(/\s+/g, " ").trim();
      if (text) return text;
    }
    current = current.previousElementSibling;
  }
  return "";
}



function htmlTableRows(table) {
  const carry = [];
  return [...table.querySelectorAll("tr")]
    .map((row) => {
      const output = [];
      let colIndex = 0;
      const placeCarriedCells = () => {
        while (carry[colIndex]?.rowsLeft > 0) {
          output[colIndex] = carry[colIndex].value;
          carry[colIndex].rowsLeft -= 1;
          if (carry[colIndex].rowsLeft <= 0) delete carry[colIndex];
          colIndex += 1;
        }
      };
      [...row.querySelectorAll("th,td")].forEach((cell) => {
        placeCarriedCells();
        const value = cell.textContent.replace(/\s+/g, " ").trim();
        const colSpan = Math.max(1, Math.min(256, Math.round(Number(cell.getAttribute("colspan")) || 1)));
        const rowSpan = Math.max(1, Math.min(4096, Math.round(Number(cell.getAttribute("rowspan")) || 1)));
        for (let offset = 0; offset < colSpan; offset += 1) {
          const targetCol = colIndex + offset;
          output[targetCol] = offset === 0 ? value : "";
          if (rowSpan > 1) {
            carry[targetCol] = { value: offset === 0 ? value : "", rowsLeft: rowSpan - 1 };
          }
        }
        colIndex += colSpan;
      });
      placeCarriedCells();
      while (output.length && String(output[output.length - 1] ?? "") === "") output.pop();
      return output;
    })
    .filter((row) => row.some((value) => value !== ""));
}



function openWebNavigatorDialog(source) {
  const sources = source.sources || [];
  state.dataDialog = {
    type: "web-navigator",
    sourceType: source.sourceType || "web",
    sourceName: source.sourceName || source.sourceUrl || "Web",
    sourceUrl: source.sourceUrl,
    sourceText: source.sourceText,
    contentType: source.contentType,
    sourceAuthMode: source.sourceAuthMode || "",
    sourceAuthUsername: source.sourceAuthUsername || "",
    sourceCredentialKey: source.sourceCredentialKey || "",
    sources,
    selectedIndex: 0,
    targetPoint: { ...source.targetPoint },
  };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = webNavigatorDialogHtml(state.dataDialog);
  document.body.appendChild(overlay);
  if (window.lucide) window.lucide.createIcons({ icons });
  overlay.querySelector("[data-data-dialog-action='web-navigator-select']")?.focus({ preventScroll: true });
}



function webNavigatorDialogHtml(descriptor) {
  const selectedSource = webNavigatorSelectedSource(descriptor);
  return `
    <div class="excel-mini-dialog data-dialog web-navigator-dialog" role="dialog" aria-modal="true" aria-labelledby="webNavigatorTitle" data-data-dialog="web-navigator">
      <div class="excel-mini-titlebar" id="webNavigatorTitle">ナビゲーター</div>
      <div class="excel-mini-body data-dialog-body web-navigator-body">
        <aside class="web-navigator-list" aria-label="テーブル">
          ${webNavigatorSourceButtonsHtml(descriptor)}
        </aside>
        <section class="web-navigator-preview-pane">
          <div class="web-navigator-summary" data-web-navigator-summary>${webNavigatorSummaryHtml(selectedSource)}</div>
          <div class="query-preview web-navigator-preview" data-web-navigator-preview>${webNavigatorPreviewHtml(selectedSource?.rows || [])}</div>
        </section>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="transform">データの変換</button>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="load">読み込み</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
}



function webNavigatorSourceButtonsHtml(descriptor) {
  const selectedIndex = Math.max(0, Math.round(Number(descriptor?.selectedIndex) || 0));
  return (descriptor?.sources || []).map((source, index) => `
    <button class="${index === selectedIndex ? "is-active" : ""}" type="button" data-data-dialog-action="web-navigator-select" data-web-source-index="${index}">
      <i data-lucide="${webNavigatorSourceIconName(source)}"></i>
      <span>${escapeHtml(source.name || `Table ${index + 1}`)}</span>
    </button>`).join("");
}



function webNavigatorSourceIconName(source) {
  if (source?.kind === "table") return "table-2";
  if (source?.kind === "workbook") return "file-spreadsheet";
  if (source?.kind === "folder" || source?.kind === "folderTable" || source?.kind === "folderCombined") return "folder-open";
  if (source?.kind === "odata") return "database";
  if (source?.kind === "json") return "braces";
  if (source?.kind === "xml") return "file-code-2";
  return "file-text";
}



function webNavigatorSummaryHtml(source) {
  if (!source) return "選択できるデータがありません。";
  const width = source.rows.length ? Math.max(...source.rows.map((row) => row.length), 0) : 0;
  return `<strong>${escapeHtml(source.name || "Table")}</strong><span>${source.rows.length} 行 x ${width} 列</span>`;
}



function webNavigatorPreviewHtml(rows) {
  if (!rows?.length) return "<div class=\"query-preview-empty\">プレビューできるデータがありません。</div>";
  const width = Math.min(Math.max(...rows.map((row) => row.length), 1), 16);
  const body = rows
    .slice(0, 100)
    .map((row, rowIndex) => {
      const cellTag = rowIndex === 0 ? "th" : "td";
      return `<tr>${Array.from({ length: width }, (_item, index) => `<${cellTag}>${escapeHtml(row[index] ?? "")}</${cellTag}>`).join("")}</tr>`;
    })
    .join("");
  return `<table><tbody>${body}</tbody></table>`;
}



function webNavigatorSelectedSource(descriptor = state.dataDialog) {
  const sources = descriptor?.sources || [];
  const selectedIndex = Math.max(0, Math.min(sources.length - 1, Math.round(Number(descriptor?.selectedIndex) || 0)));
  return sources[selectedIndex] || sources[0] || null;
}



function selectWebNavigatorSource(dialog, index) {
  if (state.dataDialog?.type !== "web-navigator") return;
  state.dataDialog.selectedIndex = Math.max(0, Math.min((state.dataDialog.sources || []).length - 1, Math.round(Number(index) || 0)));
  dialog.querySelectorAll("[data-data-dialog-action='web-navigator-select']").forEach((button) => {
    button.classList.toggle("is-active", Number(button.dataset.webSourceIndex) === state.dataDialog.selectedIndex);
  });
  const selectedSource = webNavigatorSelectedSource();
  const summary = dialog.querySelector("[data-web-navigator-summary]");
  const preview = dialog.querySelector("[data-web-navigator-preview]");
  if (summary) summary.innerHTML = webNavigatorSummaryHtml(selectedSource);
  if (preview) preview.innerHTML = webNavigatorPreviewHtml(selectedSource?.rows || []);
}



function loadWebNavigatorSelection(dialog) {
  const descriptor = state.dataDialog;
  const selectedSource = webNavigatorSelectedSource(descriptor);
  if (!descriptor || !selectedSource) return;
  loadWebRowsIntoSelection(selectedSource.rows, {
    sourceType: descriptor.sourceType || "web",
    sourceName: descriptor.sourceName || descriptor.sourceUrl || "Web",
    sourceUrl: descriptor.sourceUrl,
    sourceText: descriptor.sourceText,
    contentType: descriptor.contentType,
    sourceAuthMode: descriptor.sourceAuthMode || "",
    sourceAuthUsername: descriptor.sourceAuthUsername || "",
    sourceCredentialKey: descriptor.sourceCredentialKey || "",
    sourceRows: selectedSource.rows,
    sourceItemIndex: selectedSource.itemIndex,
    sourceItemName: selectedSource.name,
    sourceItemKind: selectedSource.kind,
    targetPoint: descriptor.targetPoint,
  });
}



function openWebNavigatorTransformEditor(dialog) {
  const descriptor = state.dataDialog;
  const selectedSource = webNavigatorSelectedSource(descriptor);
  if (!descriptor || !selectedSource) return;
  openWebNavigatorTransformEditorFromSource(descriptor, selectedSource);
}



function loadWebRowsIntoSelection(rows, options = {}) {
  const sheet = activeSheet();
  if (!sheet || !rows?.length) return false;
  const sourceType = options.sourceType || "web";
  const point = options.targetPoint || state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  const width = Math.max(...rows.map((row) => row.length), 1);
  const targetRange = normalizeRange(point, { sheetIndex: state.activeSheetIndex, row: point.row + rows.length - 1, col: point.col + width - 1 });
  if (!options.skipOverwriteCheck && rangeHasUserData(sheet, targetRange)) {
    showDataOverwriteConfirmation(() => {
      loadWebRowsIntoSelection(rows, { ...options, skipOverwriteCheck: true });
    });
    return false;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  writeRowsToSheet(sheet, rows, targetRange.top, targetRange.left, { typeDetection: "first200" });
  applyExcelTableStyle(sheet, targetRange);
  sheet.tables = Array.isArray(sheet.tables) ? sheet.tables : [];
  sheet.tables.push({
    id: uniqueWorksheetObjectId("table"),
    name: nextTableName(sheet),
    range: { ...targetRange },
    showHeaderRow: true,
    showFilterButton: true,
    style: "TableStyleMedium2",
  });
  sheet.queries = Array.isArray(sheet.queries) ? sheet.queries : [];
  const query = {
    id: uniqueWorksheetObjectId("query"),
    name: sourceType === "web" ? nextWebQueryName(sheet) : nextWorkbookQueryName(options.sourceItemName || options.sourceName || "Query"),
    sourceType,
    sourceName: options.sourceName || options.sourceUrl || "",
    sourceUrl: options.sourceUrl || "",
    sourceText: options.sourceText || "",
    contentType: options.contentType || "",
    sourceAuthMode: options.sourceAuthMode || "",
    sourceAuthUsername: options.sourceAuthUsername || "",
    sourceCredentialKey: options.sourceCredentialKey || "",
    sourceRows: querySourceTypeUsesStoredRows(sourceType) ? cloneJsonLikeValue(options.sourceRows || rows) : null,
    sourceItemIndex: options.sourceItemIndex ?? 0,
    sourceItemName: options.sourceItemName || "",
    sourceItemKind: options.sourceItemKind || "document",
    range: { ...targetRange },
    hasHeader: true,
    loadedAt: new Date().toISOString(),
    refreshWithAll: true,
  };
  sheet.queries.push(query);
  rememberRecentDataSourceFromQuery(query);
  closeDataDialog();
  const loadLabel = sourceType === "workbook"
    ? "ブック データ"
    : sourceType === "folder"
      ? "フォルダー データ"
      : sourceType === "odata"
        ? "OData データ"
        : sourceType === "json"
          ? "JSON データ"
          : sourceType === "xml"
            ? "XML データ"
            : "Web データ";
  commitDataOperation(sheet, sheetBefore, targetRange, `${rangeToLabel(targetRange)} ${loadLabel}`, `${rangeToLabel(targetRange)} に ${loadLabel}を読み込みました。`, targetRange);
  return true;
}



function querySourceTypeUsesStoredRows(sourceType) {
  return sourceType === "workbook" || sourceType === "folder";
}



function querySourceIconName(query) {
  if (query?.sourceType === "blankQuery") return "file-plus";
  if (query?.sourceType === "queryReference") return "git-branch";
  if (query?.sourceType === "queryAppend") return "git-branch";
  if (query?.sourceType === "queryMerge") return "git-branch";
  if (query?.sourceType === "web") return "globe";
  if (query?.sourceType === "odata") return "database";
  if (query?.sourceType === "sqlServer") return "plug";
  if (query?.sourceType === "access") return "table-2";
  if (query?.sourceType === "json") return "braces";
  if (query?.sourceType === "xml") return "file-code-2";
  if (query?.sourceType === "textCsv") return "file-text";
  if (query?.sourceType === "workbook") return "file-spreadsheet";
  if (query?.sourceType === "folder") return "folder-open";
  return "table-2";
}



function writeRowsToSheet(sheet, rows, startRow, startCol, options = {}) {
  const width = Math.max(...rows.map((row) => row.length), 1);
  const targetRange = normalizeRange(
    { sheetIndex: state.activeSheetIndex, row: startRow, col: startCol },
    { sheetIndex: state.activeSheetIndex, row: startRow + rows.length - 1, col: startCol + width - 1 },
  );
  ensureSheetSize(sheet, targetRange.bottom, targetRange.right);
  const sheetId = getSheetId(sheet.name);
  rows.forEach((rowValues, rowIndex) => {
    for (let colIndex = 0; colIndex < width; colIndex += 1) {
      writeCellContent(sheet, sheetId, startRow + rowIndex, startCol + colIndex, coerceDelimitedImportValue(rowValues[colIndex] ?? "", options.typeDetection || "first200"));
    }
  });
  return targetRange;
}



function nextWebQueryName(sheet) {
  const existing = new Set((sheet.queries || []).map((query) => String(query.name || "").toLowerCase()));
  let index = 1;
  while (existing.has(`webquery${index}`)) index += 1;
  return `WebQuery${index}`;
}



function nextQueryName(sheet) {
  const existing = new Set((sheet.queries || []).map((query) => String(query.name || "").toLowerCase()));
  let index = (sheet.queries || []).length + 1;
  while (existing.has(`query${index}`)) index += 1;
  return `Query${index}`;
}



function workbookQueries() {
  const rows = [];
  (state.model?.sheets || []).forEach((sheet, sheetIndex) => {
    (sheet.queries || []).forEach((query) => {
      rows.push({ sheet, sheetIndex, query });
    });
  });
  return rows;
}



function querySortTime(query) {
  const date = new Date(query?.lastRefreshed || query?.loadedAt || 0);
  return Number.isNaN(date.getTime()) ? 0 : date.getTime();
}



function workbookQueryKey(sheetIndex, queryId) {
  return `${sheetIndex}::${queryId || ""}`;
}



function workbookQueryRefFromKey(key) {
  const [sheetIndexText, ...idParts] = String(key || "").split("::");
  return workbookQueryRef(Number(sheetIndexText), idParts.join("::"));
}



function workbookQueryRef(sheetIndex, queryId) {
  const sheet = state.model?.sheets?.[sheetIndex];
  const query = sheet?.queries?.find((item) => item.id === queryId);
  return sheet && query ? { sheet, sheetIndex, query } : null;
}



function queryReferenceSourceRef(query) {
  if (!query) return null;
  const sourceSheetIndex = Number(query.sourceQuerySheetIndex ?? query.sourceSheetIndex);
  if (Number.isFinite(sourceSheetIndex) && query.sourceQueryId) {
    const direct = workbookQueryRef(sourceSheetIndex, query.sourceQueryId);
    if (direct) return direct;
  }
  if (query.sourceQueryId) {
    const byId = workbookQueries().find(({ query: candidate }) => candidate.id === query.sourceQueryId);
    if (byId) return byId;
  }
  const sourceName = String(query.sourceQueryName || query.sourceName || "").trim().toLocaleLowerCase("ja-JP");
  if (!sourceName) return null;
  return workbookQueries().find(({ query: candidate }) =>
    candidate !== query && String(candidate.name || "").trim().toLocaleLowerCase("ja-JP") === sourceName,
  ) || null;
}



function queryReferenceSourceRows(query) {
  const sourceRef = queryReferenceSourceRef(query);
  if (!sourceRef || sourceRef.query === query) return [];
  const rows = workbookQueryRowsForImport(sourceRef, new Set([query.id].filter(Boolean)));
  return workbookQueryRowsToSourceValues(rows);
}



function queryAppendSourceRefs(query) {
  return (query?.appendQueryRefs || [])
    .map((item) => workbookQueryRef(Number(item.sheetIndex), item.queryId))
    .filter(Boolean);
}



function queryAppendSourceRows(query, visited = new Set()) {
  const refs = queryAppendSourceRefs(query).filter((ref) => ref.query !== query);
  let columns = [];
  const outputRows = [];
  refs.forEach((ref, refIndex) => {
    const rows = workbookQueryRowsToSourceValues(workbookQueryRowsForImport(ref, visited));
    if (!rows.length) return;
    const header = rows[0].map((value, index) => String(value || `Column${index + 1}`).trim() || `Column${index + 1}`);
    if (!columns.length) {
      columns = uniqueQueryTransformColumns(header);
      outputRows.push([...columns]);
    } else {
      header.forEach((name) => {
        if (columns.some((column) => column.toLocaleLowerCase("ja-JP") === name.toLocaleLowerCase("ja-JP"))) return;
        columns.push(uniqueQueryTransformColumns([...columns, name]).at(-1));
        outputRows.forEach((row) => row.push(""));
      });
    }
    const normalizedHeader = header.map((name) => name.toLocaleLowerCase("ja-JP"));
    rows.slice(1).forEach((row) => {
      const next = Array.from({ length: columns.length }, () => "");
      columns.forEach((column, columnIndex) => {
        const sourceIndex = normalizedHeader.indexOf(column.toLocaleLowerCase("ja-JP"));
        if (sourceIndex >= 0) next[columnIndex] = String(row[sourceIndex] ?? "");
      });
      outputRows.push(next);
    });
  });
  return outputRows;
}



function queryStoredRef(item) {
  if (!item) return null;
  return workbookQueryRef(Number(item.sheetIndex), item.queryId);
}



function queryMergeSourceRefs(query) {
  const left = queryStoredRef(query?.mergeLeftRef);
  const right = queryStoredRef(query?.mergeRightRef);
  return { left, right };
}



function queryHeaderIndex(headers, columnName) {
  const expected = String(columnName || "").trim().toLocaleLowerCase("ja-JP");
  if (!expected) return -1;
  return (headers || []).findIndex((header) => String(header || "").trim().toLocaleLowerCase("ja-JP") === expected);
}



function queryMergeRightColumns(leftHeaders, rightHeaders, rightKeyIndex, prefix = "") {
  const leftSet = new Set((leftHeaders || []).map((header) => String(header || "").toLocaleLowerCase("ja-JP")));
  const used = new Set(leftSet);
  return (rightHeaders || []).map((header, index) => {
    if (index === rightKeyIndex) return null;
    const base = String(prefix || "") + String(header || `Column${index + 1}`);
    let name = base || `Column${index + 1}`;
    let suffix = 1;
    while (used.has(name.toLocaleLowerCase("ja-JP"))) {
      suffix += 1;
      name = `${base}.${suffix}`;
    }
    used.add(name.toLocaleLowerCase("ja-JP"));
    return { sourceIndex: index, name };
  }).filter(Boolean);
}



function queryMergeSourceRows(query, visited = new Set()) {
  const { left, right } = queryMergeSourceRefs(query);
  if (!left || !right || left.query === query || right.query === query) return [];
  const leftRows = workbookQueryRowsToSourceValues(workbookQueryRowsForImport(left, visited));
  const rightRows = workbookQueryRowsToSourceValues(workbookQueryRowsForImport(right, visited));
  if (!leftRows.length || !rightRows.length) return [];
  const leftHeaders = uniqueQueryTransformColumns(leftRows[0].map((value, index) => String(value || `Column${index + 1}`).trim() || `Column${index + 1}`));
  const rightHeaders = uniqueQueryTransformColumns(rightRows[0].map((value, index) => String(value || `Column${index + 1}`).trim() || `Column${index + 1}`));
  const leftKeyIndex = queryHeaderIndex(leftHeaders, query.mergeLeftColumn) >= 0 ? queryHeaderIndex(leftHeaders, query.mergeLeftColumn) : 0;
  const rightKeyIndex = queryHeaderIndex(rightHeaders, query.mergeRightColumn) >= 0 ? queryHeaderIndex(rightHeaders, query.mergeRightColumn) : 0;
  const joinKind = ["inner", "leftOuter", "rightOuter", "fullOuter", "leftAnti", "rightAnti"].includes(query.mergeJoinKind)
    ? query.mergeJoinKind
    : "leftOuter";
  if (joinKind === "leftAnti") return queryMergeAntiRows(leftHeaders, leftRows.slice(1), rightRows.slice(1), leftKeyIndex, rightKeyIndex, true);
  if (joinKind === "rightAnti") return queryMergeAntiRows(rightHeaders, rightRows.slice(1), leftRows.slice(1), rightKeyIndex, leftKeyIndex, true);
  const rightColumns = queryMergeRightColumns(leftHeaders, rightHeaders, rightKeyIndex, query.mergeRightPrefix || `${right.query.name || "Right"}.`);
  const output = [[...leftHeaders, ...rightColumns.map((column) => column.name)]];
  const rightByKey = new Map();
  rightRows.slice(1).forEach((row, index) => {
    const key = queryMergeKey(row[rightKeyIndex]);
    const bucket = rightByKey.get(key) || [];
    bucket.push({ row, index });
    rightByKey.set(key, bucket);
  });
  const matchedRight = new Set();
  leftRows.slice(1).forEach((leftRow) => {
    const matches = rightByKey.get(queryMergeKey(leftRow[leftKeyIndex])) || [];
    if (matches.length) {
      matches.forEach((match) => {
        matchedRight.add(match.index);
        output.push([...padQueryTransformRow(leftRow, leftHeaders.length), ...rightColumns.map((column) => String(match.row[column.sourceIndex] ?? ""))]);
      });
    } else if (joinKind === "leftOuter" || joinKind === "fullOuter") {
      output.push([...padQueryTransformRow(leftRow, leftHeaders.length), ...rightColumns.map(() => "")]);
    }
  });
  if (joinKind === "rightOuter" || joinKind === "fullOuter") {
    rightRows.slice(1).forEach((rightRow, index) => {
      if (matchedRight.has(index)) return;
      output.push([...leftHeaders.map(() => ""), ...rightColumns.map((column) => String(rightRow[column.sourceIndex] ?? ""))]);
    });
  }
  return output;
}



function queryMergeAntiRows(headers, primaryRows, secondaryRows, primaryKeyIndex, secondaryKeyIndex) {
  const secondaryKeys = new Set(secondaryRows.map((row) => queryMergeKey(row[secondaryKeyIndex])));
  return [
    [...headers],
    ...primaryRows
      .filter((row) => !secondaryKeys.has(queryMergeKey(row[primaryKeyIndex])))
      .map((row) => padQueryTransformRow(row, headers.length)),
  ];
}



function queryMergeKey(value) {
  return String(value ?? "").trim().toLocaleLowerCase("ja-JP");
}



function queryMergeCommandJoinKind(value) {
  return {
    inner: "Inner",
    leftOuter: "LeftOuter",
    rightOuter: "RightOuter",
    fullOuter: "FullOuter",
    leftAnti: "LeftAnti",
    rightAnti: "RightAnti",
  }[value] || "LeftOuter";
}



function workbookQueryDependsOnQuery(query, sourceSheetIndex, sourceQueryId) {
  if (!query || !sourceQueryId) return false;
  if (query.sourceType === "queryReference" && query.sourceQueryId === sourceQueryId) {
    const refSheetIndex = Number(query.sourceQuerySheetIndex ?? query.sourceSheetIndex);
    return !Number.isFinite(refSheetIndex) || refSheetIndex === sourceSheetIndex;
  }
  if (query.sourceType === "queryAppend" && Array.isArray(query.appendQueryRefs)) {
    return query.appendQueryRefs.some((item) =>
      item.queryId === sourceQueryId &&
      (!Number.isFinite(Number(item.sheetIndex)) || Number(item.sheetIndex) === sourceSheetIndex),
    );
  }
  if (query.sourceType === "queryMerge") {
    return [query.mergeLeftRef, query.mergeRightRef].some((item) =>
      item?.queryId === sourceQueryId &&
      (!Number.isFinite(Number(item.sheetIndex)) || Number(item.sheetIndex) === sourceSheetIndex),
    );
  }
  return false;
}



function dependentWorkbookQueriesForQuery(sheetIndex, queryId) {
  return workbookQueries().filter(({ sheetIndex: itemSheetIndex, query }) =>
    !(itemSheetIndex === sheetIndex && query.id === queryId) &&
    workbookQueryDependsOnQuery(query, sheetIndex, queryId),
  );
}



function workbookQueryDeleteMessage(ref) {
  if (!ref) return "クエリを削除しますか?";
  const base = `クエリ「${ref.query.name || "Query"}」を削除しますか?`;
  const dependents = dependentWorkbookQueriesForQuery(ref.sheetIndex, ref.query.id);
  if (!dependents.length) return base;
  const list = dependents
    .slice(0, 8)
    .map(({ sheet, query }) => `・${query.name || "Query"} (${sheet.name || "Sheet"})`)
    .join("\n");
  const overflow = dependents.length > 8 ? `\nほか ${dependents.length - 8} 件` : "";
  return `${base}\n\nこのクエリは次のクエリから参照されています。\n${list}${overflow}\n削除すると、参照先の更新に失敗する可能性があります。`;
}



function activeWorkbookQueryRef() {
  const queries = workbookQueries();
  if (!queries.length) return null;
  const range = activeSelectionRange();
  if (range) {
    const matching = queries.find(({ sheetIndex, query }) =>
      sheetIndex === state.activeSheetIndex &&
      query.range &&
      rangesIntersect(query.range, range),
    );
    if (matching) return matching;
  }
  return queries.find(({ sheetIndex }) => sheetIndex === state.activeSheetIndex) || queries[0];
}



async function refreshActiveWorkbookConnection() {
  const ref = activeWorkbookQueryRef();
  if (!ref) {
    setStatus("更新する接続はありません。");
    return;
  }
  await refreshSingleQuery(ref.sheetIndex, ref.query.id);
  refreshQueriesConnectionsPane();
}



function openActiveQueryPropertiesDialog() {
  const ref = activeWorkbookQueryRef();
  if (!ref) {
    showMessageBox("接続がありません。", { title: "Cht WebSheet" });
    return;
  }
  openQueryPropertiesDialog(ref.sheetIndex, ref.query.id);
}



function openQueryTransformEditorForQuery(sheetIndex, queryId) {
  const ref = workbookQueryRef(sheetIndex, queryId);
  if (!ref) {
    showMessageBox("編集するクエリが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  const { sheet, query } = ref;
  const targetRange = query.range || null;
  const targetPoint = targetRange
    ? { sheetIndex: ref.sheetIndex, row: targetRange.top, col: targetRange.left }
    : state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  state.dataDialog = {
    type: "query-transform-editor",
    editingQuerySheetIndex: ref.sheetIndex,
    editingQueryId: query.id,
    sourceType: query.sourceType || "textCsv",
    sourceName: query.sourceName || query.name || "Query",
    sourceUrl: query.sourceUrl || "",
    sourceText: query.sourceText || "",
    contentType: query.contentType || "",
    sourceAuthMode: query.sourceAuthMode || "",
    sourceAuthUsername: query.sourceAuthUsername || "",
    sourceCredentialKey: query.sourceCredentialKey || "",
    sourceRows: Array.isArray(query.sourceRows) ? cloneJsonLikeValue(query.sourceRows) : null,
    sourceItemIndex: query.sourceItemIndex ?? 0,
    sourceItemName: query.sourceItemName || "",
    sourceItemKind: query.sourceItemKind || "document",
    sourceQueryId: query.sourceQueryId || "",
    sourceQueryName: query.sourceQueryName || "",
    sourceQuerySheetIndex: Number.isFinite(Number(query.sourceQuerySheetIndex)) ? Number(query.sourceQuerySheetIndex) : null,
    appendQueryRefs: Array.isArray(query.appendQueryRefs) ? cloneJsonLikeValue(query.appendQueryRefs) : [],
    mergeLeftRef: query.mergeLeftRef ? cloneJsonLikeValue(query.mergeLeftRef) : null,
    mergeRightRef: query.mergeRightRef ? cloneJsonLikeValue(query.mergeRightRef) : null,
    mergeLeftColumn: query.mergeLeftColumn || "",
    mergeRightColumn: query.mergeRightColumn || "",
    mergeJoinKind: query.mergeJoinKind || "leftOuter",
    mergeRightPrefix: query.mergeRightPrefix || "",
    sourceSheetIndex: Number.isFinite(Number(query.sourceSheetIndex)) ? Number(query.sourceSheetIndex) : ref.sheetIndex,
    sourceSheetName: query.sourceSheetName || sheet.name,
    sourceRange: query.sourceRange ? { ...query.sourceRange } : null,
    origin: query.origin || "65001",
    delimiter: query.delimiter || ",",
    targetPoint: { ...targetPoint },
    transformSteps: cloneJsonLikeValue(query.transformSteps || []),
    typeDetection: query.typeDetection || "first200",
    queryName: query.name || "Query",
  };
  renderQueryTransformEditorDialog();
}



function openExistingConnectionsDialog(selectedKey = "") {
  if (!state.model) {
    setStatus("先に Excel ファイルを開いてください。");
    return;
  }
  const queries = workbookQueries();
  const activeRef = selectedKey ? workbookQueryRefFromKey(selectedKey) : activeWorkbookQueryRef();
  const selected = activeRef ? workbookQueryKey(activeRef.sheetIndex, activeRef.query.id) : workbookQueryKey(queries[0]?.sheetIndex ?? 0, queries[0]?.query?.id || "");
  state.dataDialog = { type: "existing-connections", selectedKey: selected };
  closeDataDialog({ keepState: true });
  const options = queries.length
    ? queries.map(({ sheet, sheetIndex, query }) => {
      const key = workbookQueryKey(sheetIndex, query.id);
      const range = query.range ? ` - ${rangeToLabel(query.range)}` : "";
      return `<option value="${escapeAttr(key)}"${key === selected ? " selected" : ""}>${escapeHtml(query.name || "Query")} (${escapeHtml(sheet.name)}${escapeHtml(range)})</option>`;
    }).join("")
    : '<option value="">このブックには接続がありません</option>';
  const ref = workbookQueryRefFromKey(selected) || queries[0] || null;
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog existing-connections-dialog" role="dialog" aria-modal="true" aria-labelledby="existingConnectionsTitle" data-data-dialog="existing-connections">
      <div class="excel-mini-titlebar" id="existingConnectionsTitle">既存の接続</div>
      <div class="excel-mini-body data-dialog-body existing-connections-body">
        <label class="existing-connections-list-wrap">
          <span class="excel-dialog-caption">接続</span>
          <select class="existing-connections-list" size="9" data-data-field="existingConnectionList" data-existing-connection-list>${options}</select>
        </label>
	        <div class="existing-connections-details">
	          <span class="excel-dialog-caption">説明</span>
	          <div class="existing-connections-details-box" data-existing-connection-details>${existingConnectionDetailsHtml(ref)}</div>
	          <div class="existing-connections-actions">
	            <button class="excel-mini-button" type="button" data-data-dialog-action="existing-open" ${ref ? "" : "disabled"}>開く(O)</button>
	            <button class="excel-mini-button" type="button" data-data-dialog-action="existing-select" ${ref?.query?.range ? "" : "disabled"}>場所へ移動(G)</button>
	            <button class="excel-mini-button" type="button" data-data-dialog-action="existing-properties" ${ref ? "" : "disabled"}>プロパティ(P)...</button>
	            <button class="excel-mini-button" type="button" data-data-dialog-action="existing-change-source" ${ref ? "" : "disabled"}>接続先の変更(C)...</button>
	            <button class="excel-mini-button" type="button" data-data-dialog-action="existing-refresh" ${ref ? "" : "disabled"}>更新(R)</button>
	            <button class="excel-mini-button" type="button" data-data-dialog-action="existing-delete" ${ref ? "" : "disabled"}>削除(D)</button>
	          </div>
	        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">閉じる</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector("[data-existing-connection-list]")?.focus({ preventScroll: true });
}



function existingConnectionDetailsHtml(ref) {
  if (!ref) return "このブックには接続がありません。";
  const { sheet, query } = ref;
  const range = query.range ? rangeToLabel(query.range) : "";
  const dimensions = queryLoadedDimensions(query);
  const lines = [
    `名前: ${query.name || "Query"}`,
    ...(query.description ? [`説明: ${query.description}`] : []),
    `種類: ${queryConnectionTypeLabel(query)}`,
    `読み込み先: ${queryLoadDestinationLabel(query, sheet)}`,
    `場所: ${sheet.name}${range ? `!${range}` : ""}`,
    `行数: ${dimensions.rows.toLocaleString("ja-JP")}`,
    `列数: ${dimensions.columns.toLocaleString("ja-JP")}`,
    `接続文字列: ${queryConnectionString(query, sheet)}`,
    `コマンドの種類: ${queryCommandTypeLabel(query)}`,
    `コマンド: ${queryCommandText(query, sheet)}`,
    `更新: ${queryConnectionRefreshLabel(query)}`,
    `認証: ${query.sourceAuthMode === "basic" ? "基本認証" : "匿名"}`,
    ...(query.connectionFile ? [`接続ファイル: ${query.connectionFile}`] : []),
    ...(query.sourceItemName ? [`アイテム: ${query.sourceItemName}`] : []),
    `データ モデル: ${query.addToDataModel ? "はい" : "いいえ"}`,
  ];
  if (query.transformSteps?.length) lines.push(`適用したステップ: ${query.transformSteps.map(queryTransformStepLabel).join(" / ")}`);
  if (query.lastRefreshed || query.loadedAt) lines.push(`最終更新: ${formatQueryTimestamp(query.lastRefreshed || query.loadedAt)}`);
  if (query.lastRefreshDurationMs) lines.push(`前回の更新時間: ${queryRefreshDurationLabel(query)}`);
  return escapeHtml(lines.join("\n"));
}



function selectedExistingConnectionRef(dialog) {
  const key = dialog?.querySelector("[data-existing-connection-list]")?.value || state.dataDialog?.selectedKey || "";
  state.dataDialog = state.dataDialog && typeof state.dataDialog === "object" ? state.dataDialog : {};
  state.dataDialog.selectedKey = key;
  return workbookQueryRefFromKey(key);
}



function handleExistingConnectionDialogSync(event) {
  const dialog = event.target.closest("[data-data-dialog='existing-connections']");
  if (!dialog) return;
  const ref = selectedExistingConnectionRef(dialog);
  const details = dialog.querySelector("[data-existing-connection-details]");
  if (details) details.innerHTML = existingConnectionDetailsHtml(ref);
  dialog.querySelectorAll("[data-data-dialog-action^='existing-']").forEach((button) => {
    if (button.dataset.dataDialogAction === "existing-select") {
      button.disabled = !ref?.query?.range;
      return;
    }
    button.disabled = !ref;
  });
}



async function handleExistingConnectionsDialogAction(dialog, action) {
  const ref = selectedExistingConnectionRef(dialog);
  if (!ref) return;
  if (action === "existing-open") {
    openImportDataDialog(ref.sheetIndex, ref.query.id, { returnToExistingConnections: true });
    return;
  }
  if (action === "existing-select") {
    closeDataDialog();
    selectQueryRange(ref.sheetIndex, ref.query.id);
    return;
  }
  if (action === "existing-properties") {
    openQueryPropertiesDialog(ref.sheetIndex, ref.query.id, { returnToExistingConnections: true });
    return;
  }
  if (action === "existing-change-source") {
    openQuerySourceChangeDialog(ref.sheetIndex, ref.query.id, { returnToExistingConnections: true });
    return;
  }
  if (action === "existing-refresh") {
    await refreshSingleQuery(ref.sheetIndex, ref.query.id);
    openExistingConnectionsDialog(workbookQueryKey(ref.sheetIndex, ref.query.id));
    return;
  }
  if (action === "existing-delete") {
    const selectedKey = workbookQueryKey(ref.sheetIndex, ref.query.id);
    showMessageBox(workbookQueryDeleteMessage(ref), {
      title: "Cht WebSheet",
      tone: "warning",
      buttons: [
        {
          label: "削除",
          result: "delete",
          callback: () => {
            deleteWorkbookQuery(ref.sheetIndex, ref.query.id, { skipConfirm: true });
            const next = workbookQueries().find((item) => workbookQueryKey(item.sheetIndex, item.query.id) !== selectedKey);
            openExistingConnectionsDialog(next ? workbookQueryKey(next.sheetIndex, next.query.id) : "");
          },
        },
        { label: "キャンセル", result: "cancel" },
      ],
      cancelResult: "cancel",
    });
  }
}



function openQuerySourceChangeDialog(sheetIndex, queryId, options = {}) {
  const sheet = state.model?.sheets?.[sheetIndex];
  const query = sheet?.queries?.find((item) => item.id === queryId);
  if (!sheet || !query) return;
  state.dataDialog = {
    type: "query-source-change",
    sheetIndex,
    queryId,
    returnToExistingConnectionsKey: options.returnToExistingConnections ? workbookQueryKey(sheetIndex, queryId) : "",
  };
  closeDataDialog({ keepState: true });
  const currentSource = queryChangeSourceCurrentValue(query, sheet);
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog query-source-change-dialog" role="dialog" aria-modal="true" aria-labelledby="querySourceChangeTitle" data-data-dialog="query-source-change">
      <div class="excel-mini-titlebar" id="querySourceChangeTitle">接続先の変更</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="data-dialog-note">接続: ${escapeHtml(query.name || "Query")}</div>
        <div class="data-dialog-note">種類: ${escapeHtml(queryConnectionTypeLabel(query))}</div>
        <label class="excel-dialog-field">
          <span class="excel-dialog-caption">${escapeHtml(queryChangeSourceCaption(query))}</span>
          <input type="text" data-data-field="querySourceValue" value="${escapeAttr(currentSource)}" />
        </label>
        <label class="data-dialog-check"><input type="checkbox" data-data-field="queryClearCachedSource" ${querySourceChangeShouldClearCache(query) ? "checked" : ""} />保存済みのソース テキストをクリアし、次回更新時に再取得する</label>
        <div class="data-dialog-note">Web / OData / JSON / XML 接続では URL、SQL / Access 接続では接続文字列を変更します。</div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  const input = overlay.querySelector("[data-data-field='querySourceValue']");
  input?.focus({ preventScroll: true });
  input?.select?.();
}



function queryChangeSourceCaption(query) {
  if (["web", "odata", "json", "xml"].includes(query?.sourceType)) return "URL";
  if (query?.connectionString) return "接続文字列";
  if (query?.sourceType === "sqlServer" || query?.sourceType === "access") return "接続先";
  return "ソース";
}



function queryChangeSourceCurrentValue(query, sheet = null) {
  if (["web", "odata", "json", "xml"].includes(query?.sourceType)) {
    return query?.sourceUrl || connectionStringUrlForSourceType(query.sourceType, parsePrefixedConnectionString(query.connectionString || ""), parsePrefixedConnectionString(query.connectionString || "").payload) || "";
  }
  if (query?.connectionString) return query.connectionString;
  if (query?.sourceUrl) return query.sourceUrl;
  if (query?.sourceName) return query.sourceName;
  return queryCommandText(query, sheet);
}



function querySourceChangeShouldClearCache(query) {
  return ["web", "odata", "json", "xml"].includes(query?.sourceType);
}



function applyQuerySourceChangeDialog(dialog) {
  const { sheetIndex, queryId } = state.dataDialog || {};
  const returnToExistingConnectionsKey = state.dataDialog?.returnToExistingConnectionsKey || "";
  const sheet = state.model?.sheets?.[sheetIndex];
  const query = sheet?.queries?.find((item) => item.id === queryId);
  if (!sheet || !query) return;
  const nextSource = String(dialog.querySelector("[data-data-field='querySourceValue']")?.value || "").trim();
  if (!nextSource) {
    showMessageBox("接続先を入力してください。", { title: "Cht WebSheet" });
    return;
  }
  const previousSource = queryChangeSourceCurrentValue(query, sheet);
  if (previousSource === nextSource) {
    closeDataDialog();
    if (returnToExistingConnectionsKey) openExistingConnectionsDialog(returnToExistingConnectionsKey);
    return;
  }
  const clearCache = Boolean(dialog.querySelector("[data-data-field='queryClearCachedSource']")?.checked);
  const sheetBefore = captureSheetForHistory(sheet);
  applyQuerySourceChange(query, nextSource, { clearCache });
  const nextExistingConnectionKey = returnToExistingConnectionsKey ? workbookQueryKey(sheetIndex, query.id) : "";
  closeDataDialog();
  pushHistory({
    sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: query.range || null,
    selectionAfter: query.range || null,
    label: `${query.name || "Query"} 接続先変更`,
  });
  refreshQueriesConnectionsPane();
  setStatus(`${query.name || "Query"} の接続先を変更しました。`);
  scheduleQueryRefreshTimers();
  scheduleWorkbookPersistence();
  if (nextExistingConnectionKey) openExistingConnectionsDialog(nextExistingConnectionKey);
}



function applyQuerySourceChange(query, nextSource, options = {}) {
  if (!query) return;
  const normalizedSource = normalizeQuerySourceChangeValue(query, nextSource);
  if (["web", "odata", "json", "xml"].includes(query.sourceType)) {
    query.sourceUrl = normalizedSource;
    query.sourceName = normalizedSource;
    if (query.connectionString) query.connectionString = queryConnectionStringAfterSourceChange(query, normalizedSource);
  } else if (query.connectionString) {
    query.connectionString = normalizedSource;
    const parsed = parseQueryConnectionString(query, normalizedSource, query.commandText || "");
    if (parsed.sourceUrl !== undefined) query.sourceUrl = parsed.sourceUrl;
    if (parsed.sourceName !== undefined) query.sourceName = parsed.sourceName;
  } else if (query.sourceType === "sqlServer" || query.sourceType === "access") {
    query.sourceUrl = normalizedSource;
    query.sourceName = normalizedSource;
  } else {
    query.sourceName = normalizedSource;
    if (!query.sourceUrl || /^(?:https?:|data:)/i.test(normalizedSource)) query.sourceUrl = normalizedSource;
  }
  if (options.clearCache) {
    query.sourceText = "";
    if (!querySourceTypeUsesStoredRows(query.sourceType)) query.sourceRows = null;
  }
  query.loadedAt = new Date().toISOString();
  query.lastRefreshStatus = "";
  query.refreshError = "";
}



function normalizeQuerySourceChangeValue(query, nextSource) {
  const source = String(nextSource || "").trim();
  if (!["web", "odata", "json", "xml"].includes(query?.sourceType)) return source;
  const prefixed = parsePrefixedConnectionString(source);
  return connectionStringUrlForSourceType(query.sourceType, prefixed, prefixed.payload) || source;
}



function queryConnectionStringAfterSourceChange(query, nextSource) {
  const current = String(query?.connectionString || "").trim();
  const source = String(nextSource || "").trim();
  if (!current) return current;
  const prefixed = parsePrefixedConnectionString(current);
  if (["web", "odata", "json", "xml"].includes(query?.sourceType)) {
    const prefixMap = { web: "WEB", odata: "ODATA", json: "JSON", xml: "XML" };
    return `${prefixed.prefix || prefixMap[query.sourceType] || "WEB"};${source}`;
  }
  if (query?.sourceType === "textCsv") return `TEXT;${source}`;
  if (query?.sourceType === "workbook") return `WORKBOOK;${source}`;
  if (query?.sourceType === "folder") return `FOLDER;${source}`;
  if (query?.sourceType === "sqlServer" || query?.sourceType === "access") {
    return setConnectionStringValue(current, "Data Source", source);
  }
  return source;
}



function openQueriesConnectionsPane() {
  if (!state.model) {
    setStatus("先に Excel ファイルを開いてください。");
    return;
  }
  closeQueriesConnectionsPane();
  const pane = document.createElement("aside");
  pane.id = "queryPane";
  pane.className = "query-pane";
  pane.setAttribute("aria-label", "クエリと接続");
  state.queryPane = { activeTab: "queries" };
  pane.innerHTML = queriesConnectionsPaneHtml();
  document.body.appendChild(pane);
  pane.querySelector("[data-query-pane-action='close']")?.focus({ preventScroll: true });
}



function queriesConnectionsPaneHtml() {
  const queries = workbookQueries();
  const activeTab = state.queryPane?.activeTab === "connections" ? "connections" : "queries";
  const rows = activeTab === "connections"
    ? queryPaneConnectionsHtml(queries)
    : queryPaneQueriesHtml(queries);
  return `
    <div class="query-pane-header">
      <strong>クエリと接続</strong>
      <button type="button" data-query-pane-action="close" aria-label="閉じる">×</button>
    </div>
    <div class="query-pane-toolbar">
      <button type="button" data-query-pane-action="refresh-all"><i data-lucide="refresh-cw"></i><span>すべて更新</span></button>
    </div>
    <div class="query-pane-tabs" role="tablist" aria-label="クエリと接続">
      <button class="${activeTab === "queries" ? "is-active" : ""}" type="button" role="tab" aria-selected="${activeTab === "queries"}" data-query-pane-action="tab" data-query-pane-tab="queries">クエリ <span>${queries.length}</span></button>
      <button class="${activeTab === "connections" ? "is-active" : ""}" type="button" role="tab" aria-selected="${activeTab === "connections"}" data-query-pane-action="tab" data-query-pane-tab="connections">接続 <span>${queries.length}</span></button>
    </div>
    <div class="query-pane-list">${rows}</div>`;
}



function queryPaneQueriesHtml(queries) {
  return queries.length
    ? queries.map(({ sheet, sheetIndex, query }) => queryPaneRowHtml(sheet, sheetIndex, query)).join("")
    : '<div class="query-pane-empty">このブックにはクエリがありません。</div>';
}



function queryPaneConnectionsHtml(queries) {
  return queries.length
    ? queries.map(({ sheet, sheetIndex, query }) => queryPaneConnectionRowHtml(sheet, sheetIndex, query)).join("")
    : '<div class="query-pane-empty">このブックには接続がありません。</div>';
}



function queryPaneRowHtml(sheet, sheetIndex, query) {
  const range = query.range ? rangeToLabel(query.range) : "";
  const updated = query.lastRefreshed || query.loadedAt || "";
  const loadMode = queryPaneLoadModeLabel(query);
  const authStatus = queryAuthenticationStatusLabel(query);
  const refreshStatus = queryRefreshStatusShortLabel(query);
  return `
    <section class="query-pane-item" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">
      <div class="query-pane-item-main">
        <strong>${escapeHtml(query.name || "Query")}</strong>
        <span>${escapeHtml(sheet.name)}${range ? `!${escapeHtml(range)}` : ""}</span>
        ${loadMode ? `<small>${escapeHtml(loadMode)}</small>` : ""}
        ${authStatus ? `<small>${escapeHtml(authStatus)}</small>` : ""}
        ${refreshStatus ? `<small>${escapeHtml(refreshStatus)}</small>` : ""}
        ${updated ? `<small>最終更新: ${escapeHtml(formatQueryTimestamp(updated))}</small>` : ""}
      </div>
      <div class="query-pane-item-actions">
        <button type="button" data-query-pane-action="select" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">移動</button>
        <button type="button" data-query-pane-action="edit" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">編集</button>
        <button type="button" data-query-pane-action="refresh" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">更新</button>
        <button type="button" data-query-pane-action="append" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">追加</button>
        <button type="button" data-query-pane-action="merge" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">マージ</button>
        <button type="button" data-query-pane-action="change-source" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">接続先</button>
        <button type="button" data-query-pane-action="load-to" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">読み込み先...</button>
        <button type="button" data-query-pane-action="properties" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">プロパティ</button>
        <button type="button" data-query-pane-action="delete" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">削除</button>
      </div>
    </section>`;
}



function queryPaneConnectionRowHtml(sheet, sheetIndex, query) {
  const range = query.range ? rangeToLabel(query.range) : "";
  const updated = query.lastRefreshed || query.loadedAt || "";
  const refreshState = queryConnectionRefreshLabel(query);
  const loadMode = queryPaneLoadModeLabel(query);
  const authStatus = queryAuthenticationStatusLabel(query);
  const refreshStatus = queryRefreshStatusShortLabel(query);
  return `
    <section class="query-pane-item is-connection" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">
      <div class="query-pane-item-main">
        <strong>${escapeHtml(query.name || "Query")}</strong>
        <span>${escapeHtml(queryConnectionTypeLabel(query))}</span>
        <small>${escapeHtml(sheet.name)}${range ? `!${escapeHtml(range)}` : ""}</small>
        ${loadMode ? `<small>${escapeHtml(loadMode)}</small>` : ""}
        ${authStatus ? `<small>${escapeHtml(authStatus)}</small>` : ""}
        ${refreshStatus ? `<small>${escapeHtml(refreshStatus)}</small>` : ""}
        <small>${escapeHtml(refreshState)}${updated ? ` / 最終更新: ${escapeHtml(formatQueryTimestamp(updated))}` : ""}</small>
      </div>
      <div class="query-pane-item-actions">
        <button type="button" data-query-pane-action="select" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">移動</button>
        <button type="button" data-query-pane-action="edit" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">編集</button>
        <button type="button" data-query-pane-action="refresh" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">更新</button>
        <button type="button" data-query-pane-action="append" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">追加</button>
        <button type="button" data-query-pane-action="merge" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">マージ</button>
        <button type="button" data-query-pane-action="change-source" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">接続先</button>
        <button type="button" data-query-pane-action="load-to" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">読み込み先...</button>
        <button type="button" data-query-pane-action="properties" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">プロパティ</button>
        <button type="button" data-query-pane-action="delete" data-query-id="${escapeAttr(query.id)}" data-query-sheet-index="${sheetIndex}">削除</button>
      </div>
    </section>`;
}



function queryAuthenticationStatusLabel(query) {
  if (query?.sourceAuthMode !== "basic") return "";
  const hasCredential = Boolean(query.sourceCredentialKey && state.queryCredentials?.get?.(query.sourceCredentialKey));
  return hasCredential ? "認証: 基本認証" : "認証: 基本認証 (資格情報なし)";
}



function queryRefreshStatusShortLabel(query) {
  if (query?.lastRefreshStatus === "refreshing") return "状態: 更新中";
  if (query?.lastRefreshStatus === "failed") return `状態: 失敗${query.refreshError ? ` - ${query.refreshError}` : ""}`;
  if (query?.lastRefreshStatus === "cancelled") return "状態: キャンセル済み";
  if (query?.lastRefreshStatus === "success") return "状態: 完了";
  return "";
}



function openQueryPaneContextMenuFromEvent(event) {
  event.preventDefault();
  event.stopPropagation();
  const item = event.currentTarget;
  const queryId = item?.dataset.queryId || "";
  const sheetIndex = Number(item?.dataset.querySheetIndex);
  const ref = workbookQueryRef(sheetIndex, queryId);
  if (!ref) return;
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeDataCommandMenu();
  closeQueryPaneContextMenu();
  const menu = document.createElement("div");
  menu.id = "queryPaneContextMenu";
  menu.className = "cell-context-menu query-pane-context-menu";
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", `${ref.query.name || "Query"} のメニュー`);
  menu.innerHTML = queryPaneContextMenuHtml(ref);
  enhanceCommandMenuAccessKeys(menu);
  document.body.appendChild(menu);
  state.queryPaneContextMenu = { sheetIndex, queryId };
  positionPopupMenu(menu, event.clientX, event.clientY);
  menu.querySelector("button:not([disabled])")?.focus({ preventScroll: true });
}



function queryPaneContextMenuHtml(ref) {
  const hasRange = Boolean(ref?.query?.range);
  const items = [
    { action: "edit", label: "編集(E)" },
    { action: "load-to", label: "読み込み先(L)..." },
    { action: "refresh", label: "更新(R)" },
    { action: "select", label: "移動先へジャンプ(G)", disabled: !hasRange },
    { type: "separator" },
    { action: "duplicate", label: "複製(D)" },
    { action: "reference", label: "参照(F)" },
    { action: "append", label: "クエリの追加(A)..." },
    { action: "merge", label: "クエリのマージ(M)..." },
    { action: "rename", label: "名前の変更(M)..." },
    { action: "change-source", label: "接続先の変更(C)..." },
    { type: "separator" },
    { action: "delete", label: "削除(D)" },
    { action: "properties", label: "プロパティ(P)..." },
  ];
  return items.map((item) => {
    if (item.type === "separator") return '<div class="cell-context-menu-separator" role="separator"></div>';
    const disabled = item.disabled ? " disabled" : "";
    return `<button class="cell-context-menu-item" type="button" role="menuitem" data-query-pane-menu-action="${escapeAttr(item.action)}"${disabled}>${accessKeyLabelHtml(item.label)}</button>`;
  }).join("");
}



function closeQueryPaneContextMenuOnOutsidePointer(event) {
  if (!state.queryPaneContextMenu) return;
  if (event.target.closest?.("#queryPaneContextMenu")) return;
  closeQueryPaneContextMenu();
}



function closeQueryPaneContextMenu() {
  document.getElementById("queryPaneContextMenu")?.remove();
  state.queryPaneContextMenu = null;
}



async function handleQueryPaneContextMenuAction(event) {
  const action = event.currentTarget.dataset.queryPaneMenuAction;
  const context = state.queryPaneContextMenu;
  if (!context) return;
  const { sheetIndex, queryId } = context;
  closeQueryPaneContextMenu();
  if (action === "select") {
    selectQueryRange(sheetIndex, queryId);
    return;
  }
  if (action === "edit") {
    openQueryTransformEditorForQuery(sheetIndex, queryId);
    return;
  }
  if (action === "load-to") {
    openImportDataDialog(sheetIndex, queryId);
    return;
  }
  if (action === "refresh") {
    await refreshSingleQuery(sheetIndex, queryId);
    refreshQueriesConnectionsPane();
    return;
  }
  if (action === "duplicate") {
    duplicateWorkbookQuery(sheetIndex, queryId);
    return;
  }
  if (action === "reference") {
    referenceWorkbookQuery(sheetIndex, queryId);
    return;
  }
  if (action === "append") {
    openAppendWorkbookQueryDialog(sheetIndex, queryId);
    return;
  }
  if (action === "merge") {
    openMergeWorkbookQueryDialog(sheetIndex, queryId);
    return;
  }
  if (action === "rename") {
    openQueryRenameDialog(sheetIndex, queryId);
    return;
  }
  if (action === "change-source") {
    openQuerySourceChangeDialog(sheetIndex, queryId);
    return;
  }
  if (action === "properties") {
    openQueryPropertiesDialog(sheetIndex, queryId);
    return;
  }
  if (action === "delete") {
    deleteWorkbookQuery(sheetIndex, queryId);
    refreshQueriesConnectionsPane();
  }
}



function queryPaneLoadModeLabel(query) {
  if (query?.loadMode === "connectionOnly") return "接続のみ";
  if (query?.loadMode === "pivotTable") return "ピボットテーブル";
  if (query?.loadMode === "pivotChart") return "ピボットグラフ";
  return "";
}



function queryLoadDestinationLabel(query, sheet = null) {
  const mode = queryPaneLoadModeLabel(query);
  if (query?.loadMode === "connectionOnly") return mode;
  const location = query?.range
    ? `${sheet?.name || ""}!${rangeToLabel(query.range)}`
    : "未読み込み";
  return mode ? `${mode} (${location})` : `テーブル (${location})`;
}



function queryLoadedDimensions(query) {
  if (query?.range) {
    return {
      rows: Math.max(0, Number(query.range.bottom) - Number(query.range.top) + 1),
      columns: Math.max(0, Number(query.range.right) - Number(query.range.left) + 1),
    };
  }
  if (Array.isArray(query?.sourceRows)) {
    return {
      rows: query.sourceRows.length,
      columns: query.sourceRows.reduce((max, row) => Math.max(max, Array.isArray(row) ? row.length : 0), 0),
    };
  }
  return { rows: 0, columns: 0 };
}



function queryConnectionTypeLabel(query) {
  if (query?.sourceType === "blankQuery") return "空のクエリ接続";
  if (query?.sourceType === "queryReference") return "クエリ参照接続";
  if (query?.sourceType === "queryAppend") return "クエリ追加接続";
  if (query?.sourceType === "queryMerge") return "クエリ マージ接続";
  if (query?.sourceType === "tableRange") return "ワークシート接続";
  if (query?.sourceType === "textCsv") return "テキスト ファイル接続";
  if (query?.sourceType === "workbook") return "Excel ブック接続";
  if (query?.sourceType === "folder") return "フォルダー接続";
  if (query?.sourceType === "odata") return "OData フィード接続";
  if (query?.sourceType === "sqlServer") return "SQL Server 接続";
  if (query?.sourceType === "access") return "Microsoft Access 接続";
  if (query?.sourceType === "json") return "JSON ファイル接続";
  if (query?.sourceType === "xml") return "XML ファイル接続";
  if (query?.sourceType === "web") return "Web クエリ接続";
  return query?.sourceType || "接続";
}



function queryConnectionRefreshLabel(query) {
  const parts = [query?.refreshWithAll === false ? "すべて更新の対象外" : "すべて更新の対象"];
  if (query?.refreshEveryEnabled) parts.push(`${queryRefreshIntervalMinutes(query)} 分ごとに更新`);
  if (query?.refreshOnOpen) parts.push("ファイルを開くときに更新");
  if (query?.backgroundRefresh) parts.push("バックグラウンド更新");
  if (query?.removeDataBeforeSave) parts.push("保存前にデータ削除");
  if (query?.addToDataModel) parts.push("データ モデル");
  return parts.join(" / ");
}



function queryRefreshIntervalMinutes(query) {
  const value = Number(query?.refreshIntervalMinutes);
  if (!Number.isFinite(value)) return 60;
  return Math.max(1, Math.min(32767, Math.round(value)));
}



function queryConnectionString(query, sheet) {
  if (query?.connectionString) return query.connectionString;
  if (query?.sourceType === "blankQuery") return `POWERQUERY;${query.name || query.sourceName || "空のクエリ"}`;
  if (query?.sourceType === "queryReference") {
    const sourceRef = queryReferenceSourceRef(query);
    return `POWERQUERY;Reference=${sourceRef?.query?.name || query.sourceQueryName || query.sourceName || "Query"}`;
  }
  if (query?.sourceType === "queryAppend") {
    return `POWERQUERY;Append=${queryAppendSourceRefs(query).map((ref) => ref.query.name || "Query").join(",")}`;
  }
  if (query?.sourceType === "queryMerge") {
    const refs = queryMergeSourceRefs(query);
    return `POWERQUERY;Merge=${refs.left?.query?.name || "Left"}[${query.mergeLeftColumn || ""}]=${refs.right?.query?.name || "Right"}[${query.mergeRightColumn || ""}]`;
  }
  if (query?.sourceType === "textCsv") {
    const delimiter = query.delimiter === "\t" ? "\\t" : query.delimiter || ",";
    return `TEXT;${query.sourceName || ""};Origin=${query.origin || "65001"};Delimiter=${delimiter}`;
  }
  if (query?.sourceType === "web") return `WEB;${query.sourceUrl || ""}`;
  if (query?.sourceType === "odata") return `ODATA;${query.sourceUrl || query.sourceName || ""}`;
  if (query?.sourceType === "sqlServer") return `OLEDB;Provider=MSOLEDBSQL;Data Source=${query.sourceUrl || query.sourceName || ""}`;
  if (query?.sourceType === "access") return `OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=${query.sourceUrl || query.sourceName || ""}`;
  if (query?.sourceType === "json") return `JSON;${query.sourceUrl || query.sourceName || ""}`;
  if (query?.sourceType === "xml") return `XML;${query.sourceUrl || query.sourceName || ""}`;
  if (query?.sourceType === "workbook") return `WORKBOOK;${query.sourceName || query.sourceUrl || ""}`;
  if (query?.sourceType === "folder") return `FOLDER;${query.sourceUrl || query.sourceName || ""}`;
  if (query?.sourceType === "tableRange") {
    const sourceName = query.sourceSheetName || sheet?.name || "";
    const sourceRange = query.sourceRange || query.range;
    return `WORKSHEET;${sourceName}!${sourceRange ? rangeToLabel(sourceRange) : ""}`;
  }
  return query?.range ? `${sheet?.name || ""}!${rangeToLabel(query.range)}` : "";
}



function queryCommandText(query, sheet) {
  if (query?.commandText) return query.commandText;
  if (query?.sourceType === "blankQuery") return "let\n    Source = null\nin\n    Source";
  if (query?.sourceType === "queryReference") {
    const sourceRef = queryReferenceSourceRef(query);
    return `let\n    Source = #"${sourceRef?.query?.name || query.sourceQueryName || query.sourceName || "Query"}"\nin\n    Source`;
  }
  if (query?.sourceType === "queryAppend") {
    const names = queryAppendSourceRefs(query).map((ref) => `#"${ref.query.name || "Query"}"`).join(", ");
    return `let\n    Source = Table.Combine({${names}})\nin\n    Source`;
  }
  if (query?.sourceType === "queryMerge") {
    const refs = queryMergeSourceRefs(query);
    const joinKind = query.mergeJoinKind || "leftOuter";
    return `let\n    Source = Table.NestedJoin(#"${refs.left?.query?.name || "Left"}", {"${query.mergeLeftColumn || ""}"}, #"${refs.right?.query?.name || "Right"}", {"${query.mergeRightColumn || ""}"}, "Merged", JoinKind.${queryMergeCommandJoinKind(joinKind)})\nin\n    Source`;
  }
  if (query?.sourceType === "textCsv") return query.sourceName || "";
  if (query?.sourceType === "web") return [query.sourceUrl || "", query.sourceItemName || ""].filter(Boolean).join(" - ");
  if (query?.sourceType === "odata") return [query.sourceUrl || query.sourceName || "", query.sourceItemName || ""].filter(Boolean).join(" - ");
  if (query?.sourceType === "sqlServer" || query?.sourceType === "access") return query.commandText || query.sourceItemName || query.sourceName || "";
  if (query?.sourceType === "json" || query?.sourceType === "xml") return [query.sourceUrl || query.sourceName || "", query.sourceItemName || ""].filter(Boolean).join(" - ");
  if (query?.sourceType === "workbook") return [query.sourceName || query.sourceUrl || "", query.sourceItemName || ""].filter(Boolean).join(" - ");
  if (query?.sourceType === "folder") return querySourceCommandText(query);
  if (query?.sourceType === "tableRange") {
    const sourceName = query.sourceSheetName || sheet?.name || "";
    const sourceRange = query.sourceRange || query.range;
    return `${sourceName}!${sourceRange ? rangeToLabel(sourceRange) : ""}`;
  }
  return query?.range ? rangeToLabel(query.range) : "";
}



function querySourceCommandText(query) {
  const sourceName = query?.sourceName || query?.sourceUrl || "";
  const sourceItemName = query?.sourceItemName || "";
  if (!sourceItemName) return sourceName;
  if (sourceName.includes(sourceItemName)) return sourceName;
  return [sourceName, sourceItemName].filter(Boolean).join(" - ");
}



function queryCommandType(query) {
  const value = String(query?.commandType || "").toLowerCase();
  if (["default", "sql", "table", "cube", "text"].includes(value)) return value;
  if (query?.sourceType === "blankQuery") return "default";
  if (query?.sourceType === "queryReference") return "default";
  if (query?.sourceType === "queryAppend") return "default";
  if (query?.sourceType === "queryMerge") return "default";
  if (query?.sourceType === "tableRange") return "table";
  if (query?.sourceType === "textCsv") return "text";
  if (query?.sourceType === "workbook") return "table";
  if (query?.sourceType === "folder") return "table";
  if (query?.sourceType === "odata") return "table";
  if (query?.sourceType === "sqlServer" || query?.sourceType === "access") return query.commandType || "table";
  if (query?.sourceType === "json" || query?.sourceType === "xml") return "table";
  if (query?.sourceType === "web") return "text";
  return "default";
}



function queryCommandTypeOptions(query) {
  const selected = queryCommandType(query);
  const options = [
    ["default", "既定"],
    ["sql", "SQL"],
    ["table", "テーブル"],
    ["cube", "キューブ"],
    ["text", "テキスト"],
  ];
  return options
    .map(([value, label]) => `<option value="${escapeAttr(value)}"${value === selected ? " selected" : ""}>${escapeHtml(label)}</option>`)
    .join("");
}



function queryCommandTypeLabel(query) {
  const selected = queryCommandType(query);
  return {
    default: "既定",
    sql: "SQL",
    table: "テーブル",
    cube: "キューブ",
    text: "テキスト",
  }[selected] || "既定";
}



function formatQueryTimestamp(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value || "");
  return date.toLocaleString("ja-JP", { hour12: false });
}



async function handleQueryPaneAction(event) {
  const action = event.currentTarget.dataset.queryPaneAction;
  if (action === "close") {
    closeQueriesConnectionsPane();
    return;
  }
  if (action === "refresh-all") {
    await refreshWorkbookData();
    refreshQueriesConnectionsPane();
    return;
  }
  if (action === "tab") {
    state.queryPane = state.queryPane && typeof state.queryPane === "object" ? state.queryPane : {};
    state.queryPane.activeTab = event.currentTarget.dataset.queryPaneTab === "connections" ? "connections" : "queries";
    refreshQueriesConnectionsPane();
    return;
  }
  const queryId = event.currentTarget.dataset.queryId;
  const sheetIndex = Number(event.currentTarget.dataset.querySheetIndex);
  if (action === "select") {
    selectQueryRange(sheetIndex, queryId);
    return;
  }
  if (action === "refresh") {
    await refreshSingleQuery(sheetIndex, queryId);
    refreshQueriesConnectionsPane();
    return;
  }
  if (action === "append") {
    openAppendWorkbookQueryDialog(sheetIndex, queryId);
    return;
  }
  if (action === "merge") {
    openMergeWorkbookQueryDialog(sheetIndex, queryId);
    return;
  }
  if (action === "edit") {
    openQueryTransformEditorForQuery(sheetIndex, queryId);
    return;
  }
  if (action === "properties") {
    openQueryPropertiesDialog(sheetIndex, queryId);
    return;
  }
  if (action === "change-source") {
    openQuerySourceChangeDialog(sheetIndex, queryId);
    return;
  }
  if (action === "load-to") {
    openImportDataDialog(sheetIndex, queryId);
    return;
  }
  if (action === "delete") {
    deleteWorkbookQuery(sheetIndex, queryId);
    refreshQueriesConnectionsPane();
  }
}



function refreshQueriesConnectionsPane() {
  const pane = document.getElementById("queryPane");
  if (!pane) return;
  pane.innerHTML = queriesConnectionsPaneHtml();
  if (window.lucide) window.lucide.createIcons({ icons });
}



function closeQueriesConnectionsPane() {
  closeQueryPaneContextMenu();
  document.getElementById("queryPane")?.remove();
  state.queryPane = null;
}



function closeQueryPaneOnEscape(event) {
  if (event.key !== "Escape") return;
  if (state.queryPaneContextMenu) {
    consumeCommandMenuKey(event);
    closeQueryPaneContextMenu();
    return;
  }
  if (state.queryPane) closeQueriesConnectionsPane();
}



function selectQueryRange(sheetIndex, queryId) {
  const sheet = state.model?.sheets?.[sheetIndex];
  const query = sheet?.queries?.find((item) => item.id === queryId);
  if (!sheet || !query?.range) return;
  state.activeSheetIndex = sheetIndex;
  state.selected = { sheetIndex, row: query.range.top, col: query.range.left };
  state.selectionRange = { ...query.range, sheetIndex };
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  renderWorkbook();
  scrollRangeIntoView(state.selectionRange);
  setStatus(`${query.name || "Query"} の範囲へ移動しました。`);
}



async function refreshSingleQuery(sheetIndex, queryId) {
  const sheet = state.model?.sheets?.[sheetIndex];
  const query = sheet?.queries?.find((item) => item.id === queryId);
  if (!sheet || !query) return;
  const sheetBefore = captureSheetForHistory(sheet);
  const dataModelBefore = query.addToDataModel ? cloneWorkbookDataModel() : null;
  try {
    await refreshWorkbookQueryWithRemoteSource(sheet, query);
  } catch (error) {
    console.error(error);
    if (sheetIndex === state.activeSheetIndex) renderSheet();
    refreshQueriesConnectionsPane();
    setStatus(`${query.name || "Query"} を更新できませんでした。`);
    showMessageBox(`${query.name || "Query"} を更新できませんでした。\n${readableRefreshError(error)}`, { title: "Cht WebSheet" });
    scheduleWorkbookPersistence();
    return;
  }
  if (query.addToDataModel) registerQueryInDataModel(sheetIndex, query);
  state.hf = buildFormulaEngine(state.model);
  reapplyAutoFilter({ silent: true, includeSort: true, sheet, sheetIndex });
  state.hf = buildFormulaEngine(state.model);
  pushHistory({
    sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    ...(dataModelBefore ? { dataModelBefore, dataModelAfter: cloneWorkbookDataModel() } : {}),
    selectionBefore: query.range || null,
    selectionAfter: query.range || null,
    label: `${query.name || "Query"} 更新`,
  });
  if (sheetIndex === state.activeSheetIndex) renderSheet();
  setStatus(`${query.name || "Query"} を更新しました。`);
  scheduleWorkbookPersistence();
}



function deleteWorkbookQuery(sheetIndex, queryId, options = {}) {
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet?.queries?.length) return;
  const query = sheet.queries.find((item) => item.id === queryId);
  if (!query) return;
  if (!options.skipConfirm) {
    showMessageBox(workbookQueryDeleteMessage({ sheet, sheetIndex, query }), {
      title: "Cht WebSheet",
      tone: "warning",
      buttons: [
        { label: "削除", result: "delete", callback: () => {
          deleteWorkbookQuery(sheetIndex, queryId, { skipConfirm: true });
          refreshQueriesConnectionsPane();
        } },
        { label: "キャンセル", result: "cancel" },
      ],
      cancelResult: "cancel",
    });
    return;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  const dataModelBefore = query.addToDataModel ? cloneWorkbookDataModel() : null;
  sheet.queries = sheet.queries.filter((item) => item.id !== queryId);
  if (query.addToDataModel && state.model?.dataModel?.tables) {
    state.model.dataModel.tables = state.model.dataModel.tables.filter((table) => table.queryId !== queryId);
  }
  pushHistory({
    sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    ...(dataModelBefore ? { dataModelBefore, dataModelAfter: cloneWorkbookDataModel() } : {}),
    selectionBefore: query.range || null,
    selectionAfter: query.range || null,
    label: `${query.name || "Query"} 削除`,
  });
  if (sheetIndex === state.activeSheetIndex) renderSheet();
  setStatus(`${query.name || "Query"} を削除しました。`);
  scheduleQueryRefreshTimers();
  scheduleWorkbookPersistence();
}



function duplicateWorkbookQuery(sheetIndex, queryId) {
  const ref = workbookQueryRef(sheetIndex, queryId);
  if (!ref) return;
  const { sheet, query } = ref;
  const sheetBefore = captureSheetForHistory(sheet);
  const copy = cloneWorkbookQueryForSheet(query, sheet, {
    range: null,
    loadMode: "connectionOnly",
    addToDataModel: false,
  });
  copy.range = null;
  copy.loadMode = "connectionOnly";
  copy.addToDataModel = false;
  sheet.queries = Array.isArray(sheet.queries) ? sheet.queries : [];
  sheet.queries.push(copy);
  pushHistory({
    sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: query.range || null,
    selectionAfter: copy.range || query.range || null,
    label: `${copy.name || "Query"} 複製`,
  });
  refreshQueriesConnectionsPane();
  setStatus(`${query.name || "Query"} を複製しました。`);
  scheduleQueryRefreshTimers();
  scheduleWorkbookPersistence();
}



function referenceWorkbookQuery(sheetIndex, queryId) {
  const ref = workbookQueryRef(sheetIndex, queryId);
  if (!ref) return;
  const { sheet, query } = ref;
  const sheetBefore = captureSheetForHistory(sheet);
  const now = new Date().toISOString();
  const sourceName = query.name || "Query";
  const reference = {
    id: uniqueWorksheetObjectId("query"),
    name: nextWorkbookQueryName(sourceName),
    sourceType: "queryReference",
    sourceName,
    sourceQueryId: query.id,
    sourceQueryName: sourceName,
    sourceQuerySheetIndex: sheetIndex,
    sourceSheetIndex: sheetIndex,
    sourceSheetName: sheet.name,
    sourceRows: null,
    range: null,
    hasHeader: true,
    typeDetection: query.typeDetection || "first200",
    transformSteps: [],
    loadMode: "connectionOnly",
    addToDataModel: false,
    loadedAt: now,
    lastRefreshed: now,
    refreshWithAll: true,
  };
  sheet.queries = Array.isArray(sheet.queries) ? sheet.queries : [];
  sheet.queries.push(reference);
  pushHistory({
    sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: query.range || null,
    selectionAfter: query.range || null,
    label: `${reference.name || "Query"} 参照`,
  });
  refreshQueriesConnectionsPane();
  setStatus(`${sourceName} を参照するクエリを作成しました。`);
  scheduleQueryRefreshTimers();
  scheduleWorkbookPersistence();
}



function openAppendWorkbookQueryDialog(sheetIndex, queryId) {
  const baseRef = workbookQueryRef(sheetIndex, queryId);
  const queries = workbookQueries();
  if (!baseRef || queries.length < 2) {
    showMessageBox("追加できる別のクエリがありません。", { title: "Cht WebSheet" });
    return;
  }
  state.dataDialog = { type: "query-append", sheetIndex, queryId };
  closeDataDialog({ keepState: true });
  const baseKey = workbookQueryKey(sheetIndex, queryId);
  const rows = queries.map((item) => {
    const key = workbookQueryKey(item.sheetIndex, item.query.id);
    const checked = key === baseKey ? " checked disabled" : "";
    const label = `${item.query.name || "Query"} (${item.sheet.name})`;
    return `<label class="data-dialog-check"><input type="checkbox" data-query-append-key="${escapeAttr(key)}"${checked} />${escapeHtml(label)}</label>`;
  }).join("");
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog query-append-dialog" role="dialog" aria-modal="true" aria-labelledby="queryAppendTitle" data-data-dialog="query-append">
      <div class="excel-mini-titlebar" id="queryAppendTitle">追加</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="data-dialog-note">追加するテーブルを選択してください。列名が一致する列は同じ列として追加し、一致しない列は末尾に追加します。</div>
        <div class="query-append-list">${rows}</div>
        <div class="data-dialog-note" data-query-append-preview>${queryAppendDialogPreviewHtml([baseKey])}</div>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">新しいクエリ名</span><input type="text" data-data-field="appendQueryName" value="${escapeAttr(nextWorkbookQueryName(`${baseRef.query.name || "Query"} の追加`))}" /></label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector("[data-data-field='appendQueryName']")?.focus({ preventScroll: true });
}



function handleAppendQuerySelectionChange(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='query-append']");
  if (!dialog) return;
  const preview = dialog.querySelector("[data-query-append-preview]");
  if (preview) preview.innerHTML = queryAppendDialogPreviewHtml(selectedAppendQueryKeysFromDialog(dialog));
}



function selectedAppendQueryKeysFromDialog(dialog) {
  return Array.from(dialog?.querySelectorAll("[data-query-append-key]:checked") || [])
    .map((input) => input.dataset.queryAppendKey || "")
    .filter(Boolean);
}



function queryAppendDialogPreviewHtml(keys) {
  const refs = [...new Set(keys || [])].map((key) => workbookQueryRefFromKey(key)).filter(Boolean);
  if (!refs.length) return "追加するクエリを選択してください。";
  const rows = queryAppendSourceRows({
    appendQueryRefs: refs.map((ref) => ({ sheetIndex: ref.sheetIndex, queryId: ref.query.id })),
  });
  const rowCount = Math.max(0, rows.length - 1);
  const colCount = rows[0]?.length || 0;
  return `${refs.length} 個のクエリを追加: ${rowCount} 行、${colCount} 列`;
}



function applyAppendWorkbookQueryDialog(dialog) {
  const baseRef = workbookQueryRef(Number(state.dataDialog?.sheetIndex), state.dataDialog?.queryId);
  if (!baseRef) return;
  const selectedKeys = [
    workbookQueryKey(baseRef.sheetIndex, baseRef.query.id),
    ...Array.from(dialog.querySelectorAll("[data-query-append-key]:checked:not(:disabled)")).map((input) => input.dataset.queryAppendKey || ""),
  ].filter(Boolean);
  const refs = [...new Set(selectedKeys)].map((key) => workbookQueryRefFromKey(key)).filter(Boolean);
  if (refs.length < 2) {
    showMessageBox("追加するクエリを 2 つ以上選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const name = String(dialog.querySelector("[data-data-field='appendQueryName']")?.value || "").trim() || nextWorkbookQueryName(`${baseRef.query.name || "Query"} の追加`);
  const sheet = baseRef.sheet;
  const sheetBefore = captureSheetForHistory(sheet);
  const now = new Date().toISOString();
  const query = {
    id: uniqueWorksheetObjectId("query"),
    name: nextWorkbookQueryName(name),
    sourceType: "queryAppend",
    sourceName: refs.map((ref) => ref.query.name || "Query").join(", "),
    appendQueryRefs: refs.map((ref) => ({
      sheetIndex: ref.sheetIndex,
      queryId: ref.query.id,
      name: ref.query.name || "Query",
      sheetName: ref.sheet.name,
    })),
    sourceRows: null,
    range: null,
    hasHeader: true,
    typeDetection: baseRef.query.typeDetection || "first200",
    transformSteps: [],
    loadMode: "connectionOnly",
    addToDataModel: false,
    loadedAt: now,
    lastRefreshed: now,
    refreshWithAll: true,
  };
  sheet.queries = Array.isArray(sheet.queries) ? sheet.queries : [];
  sheet.queries.push(query);
  closeDataDialog();
  pushHistory({
    sheetIndex: baseRef.sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: baseRef.query.range || null,
    selectionAfter: baseRef.query.range || null,
    label: `${query.name || "Query"} 追加`,
  });
  refreshQueriesConnectionsPane();
  setStatus(`${refs.length} 個のクエリを追加するクエリを作成しました。`);
  scheduleQueryRefreshTimers();
  scheduleWorkbookPersistence();
}



function querySourceColumnsForRef(ref) {
  const rows = workbookQueryRowsToSourceValues(workbookQueryRowsForImport(ref));
  return (rows[0] || []).map((value, index) => String(value || `Column${index + 1}`).trim() || `Column${index + 1}`);
}



function openMergeWorkbookQueryDialog(sheetIndex, queryId) {
  const leftRef = workbookQueryRef(sheetIndex, queryId);
  const queries = workbookQueries().filter((item) => workbookQueryKey(item.sheetIndex, item.query.id) !== workbookQueryKey(sheetIndex, queryId));
  if (!leftRef || !queries.length) {
    showMessageBox("マージできる別のクエリがありません。", { title: "Cht WebSheet" });
    return;
  }
  const leftColumns = querySourceColumnsForRef(leftRef);
  const rightDefault = queries[0];
  const rightColumns = querySourceColumnsForRef(rightDefault);
  state.dataDialog = { type: "query-merge", sheetIndex, queryId };
  closeDataDialog({ keepState: true });
  const queryOptions = queries.map((item) => {
    const key = workbookQueryKey(item.sheetIndex, item.query.id);
    return `<option value="${escapeAttr(key)}">${escapeHtml(`${item.query.name || "Query"} (${item.sheet.name})`)}</option>`;
  }).join("");
  const leftColumnOptions = leftColumns.map((column) => `<option value="${escapeAttr(column)}">${escapeHtml(column)}</option>`).join("");
  const rightColumnOptions = rightColumns.map((column) => `<option value="${escapeAttr(column)}">${escapeHtml(column)}</option>`).join("");
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog query-merge-dialog" role="dialog" aria-modal="true" aria-labelledby="queryMergeTitle" data-data-dialog="query-merge">
      <div class="excel-mini-titlebar" id="queryMergeTitle">マージ</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="data-dialog-note">2 つのクエリをキー列で結合します。右側のキー列以外の列は展開されます。</div>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">左テーブル</span><input type="text" value="${escapeAttr(leftRef.query.name || "Query")}" disabled /></label>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">左キー列</span><input type="text" data-data-field="mergeLeftColumn" data-query-merge-preview-sync list="mergeLeftColumns" value="${escapeAttr(leftColumns[0] || "")}" /></label>
        <datalist id="mergeLeftColumns">${leftColumnOptions}</datalist>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">右テーブル</span><select data-data-field="mergeRightQuery">${queryOptions}</select></label>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">右キー列</span><input type="text" data-data-field="mergeRightColumn" data-query-merge-preview-sync list="mergeRightColumns" value="${escapeAttr(rightColumns[0] || "")}" /></label>
        <datalist id="mergeRightColumns">${rightColumnOptions}</datalist>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">結合の種類</span><select data-data-field="mergeJoinKind" data-query-merge-preview-sync>
          <option value="leftOuter">左外部: 最初の行すべて、2 番目の一致行</option>
          <option value="rightOuter">右外部: 2 番目の行すべて、最初の一致行</option>
          <option value="fullOuter">完全外部: 両方の行すべて</option>
          <option value="inner">内部: 一致する行のみ</option>
          <option value="leftAnti">左反: 最初のみに存在する行</option>
          <option value="rightAnti">右反: 2 番目のみに存在する行</option>
        </select></label>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">右列の接頭辞</span><input type="text" data-data-field="mergeRightPrefix" data-query-merge-preview-sync value="${escapeAttr(`${rightDefault.query.name || "Right"}.`)}" /></label>
        <div class="data-dialog-note" data-query-merge-preview>${queryMergeDialogPreviewHtml(leftRef, rightDefault, leftColumns[0] || "", rightColumns[0] || "", "leftOuter", `${rightDefault.query.name || "Right"}.`)}</div>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">新しいクエリ名</span><input type="text" data-data-field="mergeQueryName" value="${escapeAttr(nextWorkbookQueryName(`${leftRef.query.name || "Query"} のマージ`))}" /></label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector("[data-data-field='mergeLeftColumn']")?.focus({ preventScroll: true });
}



function handleMergeRightQueryChange(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='query-merge']");
  if (!dialog) return;
  const rightRef = workbookQueryRefFromKey(event.currentTarget.value || "");
  if (!rightRef) return;
  const columns = querySourceColumnsForRef(rightRef);
  const list = dialog.querySelector("#mergeRightColumns");
  if (list) {
    list.innerHTML = columns.map((column) => `<option value="${escapeAttr(column)}">${escapeHtml(column)}</option>`).join("");
  }
  const rightColumn = dialog.querySelector("[data-data-field='mergeRightColumn']");
  if (rightColumn) rightColumn.value = columns[0] || "";
  const prefix = dialog.querySelector("[data-data-field='mergeRightPrefix']");
  if (prefix) prefix.value = `${rightRef.query.name || "Right"}.`;
  updateMergeQueryDialogPreview(dialog);
}



function handleMergeQueryPreviewChange(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='query-merge']");
  if (dialog) updateMergeQueryDialogPreview(dialog);
}



function updateMergeQueryDialogPreview(dialog) {
  const preview = dialog?.querySelector("[data-query-merge-preview]");
  if (!preview) return;
  const leftRef = workbookQueryRef(Number(state.dataDialog?.sheetIndex), state.dataDialog?.queryId);
  const rightRef = workbookQueryRefFromKey(dialog.querySelector("[data-data-field='mergeRightQuery']")?.value || "");
  preview.innerHTML = queryMergeDialogPreviewHtml(
    leftRef,
    rightRef,
    dialog.querySelector("[data-data-field='mergeLeftColumn']")?.value || "",
    dialog.querySelector("[data-data-field='mergeRightColumn']")?.value || "",
    dialog.querySelector("[data-data-field='mergeJoinKind']")?.value || "leftOuter",
    dialog.querySelector("[data-data-field='mergeRightPrefix']")?.value || "",
  );
}



function queryMergeDialogPreviewHtml(leftRef, rightRef, leftColumn, rightColumn, joinKind, rightPrefix) {
  if (!leftRef || !rightRef) return "マージするクエリを選択してください。";
  const rows = queryMergeSourceRows({
    mergeLeftRef: { sheetIndex: leftRef.sheetIndex, queryId: leftRef.query.id },
    mergeRightRef: { sheetIndex: rightRef.sheetIndex, queryId: rightRef.query.id },
    mergeLeftColumn: leftColumn,
    mergeRightColumn: rightColumn,
    mergeJoinKind: joinKind,
    mergeRightPrefix: rightPrefix,
  });
  const rowCount = Math.max(0, rows.length - 1);
  const colCount = rows[0]?.length || 0;
  return `${leftRef.query.name || "Query"} と ${rightRef.query.name || "Query"} をマージ: ${rowCount} 行、${colCount} 列`;
}



function applyMergeWorkbookQueryDialog(dialog) {
  const leftRef = workbookQueryRef(Number(state.dataDialog?.sheetIndex), state.dataDialog?.queryId);
  const rightRef = workbookQueryRefFromKey(dialog.querySelector("[data-data-field='mergeRightQuery']")?.value || "");
  if (!leftRef || !rightRef) return;
  const leftColumn = String(dialog.querySelector("[data-data-field='mergeLeftColumn']")?.value || "").trim();
  const rightColumn = String(dialog.querySelector("[data-data-field='mergeRightColumn']")?.value || "").trim();
  if (!leftColumn || !rightColumn) {
    showMessageBox("マージするキー列を指定してください。", { title: "Cht WebSheet" });
    return;
  }
  const leftColumns = querySourceColumnsForRef(leftRef);
  const rightColumns = querySourceColumnsForRef(rightRef);
  if (queryHeaderIndex(leftColumns, leftColumn) < 0 || queryHeaderIndex(rightColumns, rightColumn) < 0) {
    showMessageBox("指定したキー列が見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  const name = String(dialog.querySelector("[data-data-field='mergeQueryName']")?.value || "").trim() || nextWorkbookQueryName(`${leftRef.query.name || "Query"} のマージ`);
  const sheet = leftRef.sheet;
  const sheetBefore = captureSheetForHistory(sheet);
  const now = new Date().toISOString();
  const query = {
    id: uniqueWorksheetObjectId("query"),
    name: nextWorkbookQueryName(name),
    sourceType: "queryMerge",
    sourceName: `${leftRef.query.name || "Left"} + ${rightRef.query.name || "Right"}`,
    mergeLeftRef: { sheetIndex: leftRef.sheetIndex, queryId: leftRef.query.id, name: leftRef.query.name || "Query", sheetName: leftRef.sheet.name },
    mergeRightRef: { sheetIndex: rightRef.sheetIndex, queryId: rightRef.query.id, name: rightRef.query.name || "Query", sheetName: rightRef.sheet.name },
    mergeLeftColumn: leftColumn,
    mergeRightColumn: rightColumn,
    mergeJoinKind: dialog.querySelector("[data-data-field='mergeJoinKind']")?.value || "leftOuter",
    mergeRightPrefix: String(dialog.querySelector("[data-data-field='mergeRightPrefix']")?.value || ""),
    sourceRows: null,
    range: null,
    hasHeader: true,
    typeDetection: leftRef.query.typeDetection || "first200",
    transformSteps: [],
    loadMode: "connectionOnly",
    addToDataModel: false,
    loadedAt: now,
    lastRefreshed: now,
    refreshWithAll: true,
  };
  sheet.queries = Array.isArray(sheet.queries) ? sheet.queries : [];
  sheet.queries.push(query);
  closeDataDialog();
  pushHistory({
    sheetIndex: leftRef.sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: leftRef.query.range || null,
    selectionAfter: leftRef.query.range || null,
    label: `${query.name || "Query"} マージ`,
  });
  refreshQueriesConnectionsPane();
  setStatus(`${leftRef.query.name || "Query"} と ${rightRef.query.name || "Query"} をマージするクエリを作成しました。`);
  scheduleQueryRefreshTimers();
  scheduleWorkbookPersistence();
}



function openQueryRenameDialog(sheetIndex, queryId) {
  const ref = workbookQueryRef(sheetIndex, queryId);
  if (!ref) return;
  state.dataDialog = { type: "query-rename", sheetIndex, queryId };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog query-rename-dialog" role="dialog" aria-modal="true" aria-labelledby="queryRenameTitle" data-data-dialog="query-rename">
      <div class="excel-mini-titlebar" id="queryRenameTitle">名前の変更</div>
      <div class="excel-mini-body data-dialog-body">
        <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("新しい名前(N):")}</span><input type="text" data-data-field="queryRenameName" value="${escapeAttr(ref.query.name || "Query")}" /></label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  const input = overlay.querySelector("[data-data-field='queryRenameName']");
  input?.focus({ preventScroll: true });
  input?.select?.();
}



function applyQueryRenameDialog(dialog) {
  const { sheetIndex, queryId } = state.dataDialog || {};
  const ref = workbookQueryRef(sheetIndex, queryId);
  if (!ref) return;
  const nextName = String(dialog.querySelector("[data-data-field='queryRenameName']")?.value || "").trim();
  if (!nextName) {
    showMessageBox("名前を入力してください。", { title: "Cht WebSheet" });
    return;
  }
  const duplicate = workbookQueries().some(({ query }) =>
    query.id !== queryId && String(query.name || "").trim().toLocaleLowerCase("ja-JP") === nextName.toLocaleLowerCase("ja-JP"),
  );
  if (duplicate) {
    showMessageBox("同じ名前のクエリが既に存在します。別の名前を入力してください。", { title: "Cht WebSheet" });
    return;
  }
  const affectedSheetIndexes = [...new Set(workbookQueries()
    .filter(({ sheetIndex: itemSheetIndex, query }) =>
      itemSheetIndex === sheetIndex ||
      (query.sourceType === "queryReference" && query.sourceQueryId === queryId) ||
      (query.sourceType === "queryAppend" && (query.appendQueryRefs || []).some((item) => Number(item.sheetIndex) === sheetIndex && item.queryId === queryId)) ||
      (query.sourceType === "queryMerge" &&
        ((Number(query.mergeLeftRef?.sheetIndex) === sheetIndex && query.mergeLeftRef?.queryId === queryId) ||
          (Number(query.mergeRightRef?.sheetIndex) === sheetIndex && query.mergeRightRef?.queryId === queryId))),
    )
    .map(({ sheetIndex: itemSheetIndex }) => itemSheetIndex))];
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  const dataModelBefore = ref.query.addToDataModel ? cloneWorkbookDataModel() : null;
  const previousName = ref.query.name || "Query";
  ref.query.name = nextName;
  workbookQueries().forEach(({ query }) => {
    if (query.sourceType === "queryReference" && query.sourceQueryId === queryId) {
      query.sourceQueryName = nextName;
      query.sourceName = nextName;
    }
    if (query.sourceType === "queryAppend" && Array.isArray(query.appendQueryRefs)) {
      query.appendQueryRefs.forEach((item) => {
        if (Number(item.sheetIndex) === sheetIndex && item.queryId === queryId) item.name = nextName;
      });
      query.sourceName = query.appendQueryRefs.map((item) => item.name || "Query").join(", ");
    }
    if (query.sourceType === "queryMerge") {
      if (Number(query.mergeLeftRef?.sheetIndex) === sheetIndex && query.mergeLeftRef?.queryId === queryId) query.mergeLeftRef.name = nextName;
      if (Number(query.mergeRightRef?.sheetIndex) === sheetIndex && query.mergeRightRef?.queryId === queryId) query.mergeRightRef.name = nextName;
      query.sourceName = `${query.mergeLeftRef?.name || "Left"} + ${query.mergeRightRef?.name || "Right"}`;
    }
  });
  if (ref.query.addToDataModel) registerQueryInDataModel(sheetIndex, ref.query);
  closeDataDialog();
  pushHistory({
    sheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(affectedSheetIndexes),
    ...(dataModelBefore ? { dataModelBefore, dataModelAfter: cloneWorkbookDataModel() } : {}),
    selectionBefore: ref.query.range || null,
    selectionAfter: ref.query.range || null,
    label: `${previousName} 名前の変更`,
  });
  refreshQueriesConnectionsPane();
  setStatus(`${previousName} の名前を ${nextName} に変更しました。`);
  scheduleQueryRefreshTimers();
  scheduleWorkbookPersistence();
}



function openImportDataDialog(sheetIndex, queryId, options = {}) {
  const ref = workbookQueryRef(sheetIndex, queryId);
  if (!ref) return;
  const activePoint = state.selected?.sheetIndex === state.activeSheetIndex
    ? state.selected
    : { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  const destination = workbookCellReference(activePoint.sheetIndex, activePoint.row, activePoint.col);
  const requestedMode = options.defaultMode || ref.query.loadMode;
  const initialMode = ["table", "pivotTable", "pivotChart", "connectionOnly"].includes(requestedMode) ? requestedMode : "table";
  const initialNewSheet = initialMode === "pivotTable" || initialMode === "pivotChart";
  const initialConnectionOnly = initialMode === "connectionOnly";
  state.dataDialog = {
    type: "import-data",
    sourceSheetIndex: sheetIndex,
    queryId,
    returnToExistingConnectionsKey: options.returnToExistingConnections ? workbookQueryKey(sheetIndex, queryId) : "",
  };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog import-data-dialog" role="dialog" aria-modal="true" aria-labelledby="importDataTitle" data-data-dialog="import-data">
      <div class="excel-mini-titlebar" id="importDataTitle">データのインポート</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="data-dialog-note">接続: ${escapeHtml(ref.query.name || "Query")} / ${escapeHtml(queryConnectionTypeLabel(ref.query))}</div>
        <fieldset class="data-dialog-fieldset import-data-mode">
          <legend>このデータをブックでどのように表示するか選択してください。</legend>
          <label class="data-dialog-check"><input type="radio" name="importDataMode" value="table" data-import-sync ${initialMode === "table" ? "checked" : ""} />テーブル</label>
          <label class="data-dialog-check"><input type="radio" name="importDataMode" value="pivotTable" data-import-sync ${initialMode === "pivotTable" ? "checked" : ""} />ピボットテーブル レポート</label>
          <label class="data-dialog-check"><input type="radio" name="importDataMode" value="pivotChart" data-import-sync ${initialMode === "pivotChart" ? "checked" : ""} />ピボットグラフとピボットテーブル レポート</label>
          <label class="data-dialog-check"><input type="radio" name="importDataMode" value="connectionOnly" data-import-sync ${initialConnectionOnly ? "checked" : ""} />接続の作成のみ</label>
        </fieldset>
        <fieldset class="data-dialog-fieldset import-data-location">
          <legend>データを返す先を選択してください。</legend>
          <label class="data-dialog-check"><input type="radio" name="importDataLocation" value="existing" data-import-sync ${initialNewSheet ? "" : "checked"} ${initialConnectionOnly ? "disabled" : ""} />既存のワークシート</label>
          <label class="excel-dialog-field import-data-destination"><span class="excel-dialog-caption">場所</span><input type="text" data-data-field="destination" value="${escapeAttr(destination)}" ${initialConnectionOnly || initialNewSheet ? "disabled" : ""} /></label>
          <label class="data-dialog-check"><input type="radio" name="importDataLocation" value="new" data-import-sync ${initialNewSheet ? "checked" : ""} ${initialConnectionOnly ? "disabled" : ""} />新規ワークシート</label>
        </fieldset>
        <label class="data-dialog-check"><input type="checkbox" data-data-field="addToDataModel" />このデータをデータ モデルに追加する</label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="import-properties">プロパティ(P)...</button>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector("[data-data-field='destination']")?.focus({ preventScroll: true });
}



function handleImportDataDialogSync(event) {
  const dialog = event.target.closest("[data-data-dialog='import-data']");
  if (!dialog) return;
  const connectionOnly = dialog.querySelector("input[name='importDataMode']:checked")?.value === "connectionOnly";
  const newWorksheet = dialog.querySelector("input[name='importDataLocation']:checked")?.value === "new";
  dialog.querySelector("[data-data-field='destination']")?.toggleAttribute("disabled", connectionOnly || newWorksheet);
  dialog.querySelectorAll("input[name='importDataLocation']").forEach((input) => {
    input.disabled = connectionOnly;
  });
}



function applyImportDataDialog(dialog) {
  const { sourceSheetIndex, queryId } = state.dataDialog || {};
  const ref = workbookQueryRef(sourceSheetIndex, queryId);
  if (!ref) return;
  const mode = dialog.querySelector("input[name='importDataMode']:checked")?.value || "table";
  const addToDataModel = Boolean(dialog.querySelector("[data-data-field='addToDataModel']")?.checked);
  if (mode === "connectionOnly") {
    applyConnectionOnlyImport(ref, { addToDataModel });
    return;
  }
  const location = dialog.querySelector("input[name='importDataLocation']:checked")?.value || "existing";
  if (mode === "pivotTable" || mode === "pivotChart") {
    createPivotReportFromQuery(ref, {
      chart: mode === "pivotChart",
      newWorksheet: location === "new",
      destinationText: dialog.querySelector("[data-data-field='destination']")?.value || "",
      addToDataModel,
    });
    return;
  }
  if (location === "new") {
    loadWorkbookQueryToNewSheet(ref, { addToDataModel });
    return;
  }
  const destinationText = dialog.querySelector("[data-data-field='destination']")?.value || "";
  const destination = decodeWorkbookRange(destinationText);
  if (!destination) {
    showMessageBox("データを返す場所を正しく入力してください。", { title: "Cht WebSheet" });
    return;
  }
  loadWorkbookQueryToRange(ref, {
    targetSheetIndex: destination.sheetIndex ?? state.activeSheetIndex,
    targetRow: destination.top,
    targetCol: destination.left,
    addToDataModel,
  });
}



function applyConnectionOnlyImport(ref, options = {}) {
  const sheet = ref.sheet;
  const sheetBefore = captureSheetForHistory(sheet);
  const dataModelBefore = cloneWorkbookDataModel();
  ref.query.loadMode = "connectionOnly";
  ref.query.addToDataModel = Boolean(options.addToDataModel);
  ref.query.loadedAt = new Date().toISOString();
  if (ref.query.addToDataModel) registerQueryInDataModel(ref.sheetIndex, ref.query);
  closeDataDialog();
  pushHistory({
    sheetIndex: ref.sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    dataModelBefore,
    dataModelAfter: cloneWorkbookDataModel(),
    selectionBefore: state.selectionRange || null,
    selectionAfter: state.selectionRange || null,
    label: `${ref.query.name || "Query"} 接続のみ`,
  });
  refreshQueriesConnectionsPane();
  setStatus(`${ref.query.name || "Query"} を接続のみとして保持しました。`);
  scheduleWorkbookPersistence();
}



function loadWorkbookQueryToRange(ref, options = {}) {
  const rows = workbookQueryRowsForImport(ref);
  if (!rows.height || !rows.width) {
    showMessageBox("読み込むデータがありません。", { title: "Cht WebSheet" });
    return;
  }
  const targetSheetIndex = options.targetSheetIndex ?? state.activeSheetIndex;
  const targetSheet = state.model?.sheets?.[targetSheetIndex];
  if (!targetSheet) return;
  const targetRange = normalizeRange(
    { sheetIndex: targetSheetIndex, row: options.targetRow, col: options.targetCol },
    { sheetIndex: targetSheetIndex, row: options.targetRow + rows.height - 1, col: options.targetCol + rows.width - 1 },
  );
  const ignoreRange = targetSheetIndex === ref.sheetIndex && ref.query.range ? ref.query.range : null;
  if (!options.overwriteConfirmed && rangeHasUserData(targetSheet, targetRange, ignoreRange)) {
    showDataOverwriteConfirmation(() => loadWorkbookQueryToRange(ref, { ...options, overwriteConfirmed: true }));
    return;
  }
  const sheetIndexes = [ref.sheetIndex, targetSheetIndex];
  const sheetsBefore = captureSheetsForHistory(sheetIndexes);
  const dataModelBefore = cloneWorkbookDataModel();
  const beforeSelection = state.selectionRange || null;
  let targetQuery = ref.query;
  if (targetSheetIndex !== ref.sheetIndex) {
    ref.sheet.queries = (ref.sheet.queries || []).filter((query) => query.id !== ref.query.id);
    targetSheet.queries = Array.isArray(targetSheet.queries) ? targetSheet.queries : [];
    targetSheet.queries.push(targetQuery);
  }
  ensureSheetSize(targetSheet, targetRange.bottom, targetRange.right);
  writeWorkbookQueryRows(targetSheet, rows, targetRange.top, targetRange.left);
  targetQuery.range = { ...targetRange };
  targetQuery.loadMode = "table";
  targetQuery.addToDataModel = Boolean(options.addToDataModel);
  targetQuery.loadedAt = new Date().toISOString();
  targetQuery.lastRefreshed = targetQuery.lastRefreshed || targetQuery.loadedAt;
  registerQueryTableLoad(targetSheet, targetRange, targetQuery);
  if (targetQuery.addToDataModel) registerQueryInDataModel(targetSheetIndex, targetQuery);
  internCellsInRange(targetSheet, targetRange);
  state.activeSheetIndex = targetSheetIndex;
  state.hf = buildFormulaEngine(state.model);
  closeDataDialog();
  pushHistory({
    sheetIndex: targetSheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(sheetIndexes),
    dataModelBefore,
    dataModelAfter: cloneWorkbookDataModel(),
    selectionBefore: beforeSelection,
    selectionAfter: targetRange,
    label: `${targetQuery.name || "Query"} 読み込み先`,
  });
  state.selected = { sheetIndex: targetSheetIndex, row: targetRange.top, col: targetRange.left };
  state.selectionRange = { ...targetRange };
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  renderWorkbook();
  scrollRangeIntoView(targetRange);
  refreshQueriesConnectionsPane();
  setStatus(`${targetQuery.name || "Query"} を ${rangeToLabel(targetRange)} に読み込みました。`);
  scheduleWorkbookPersistence();
}



function loadWorkbookQueryToNewSheet(ref, options = {}) {
  const rows = workbookQueryRowsForImport(ref);
  if (!rows.height || !rows.width) {
    showMessageBox("読み込むデータがありません。", { title: "Cht WebSheet" });
    return;
  }
  const dataModelBefore = cloneWorkbookDataModel();
  const activeSheetBefore = state.activeSheetIndex;
  const sheetName = uniqueSheetName(state.model.sheets.map((sheet) => sheet.name));
  const sheet = createBlankSheet(sheetName);
  const targetRange = normalizeRange(
    { sheetIndex: state.model.sheets.length, row: 1, col: 1 },
    { sheetIndex: state.model.sheets.length, row: rows.height, col: rows.width },
  );
  ensureSheetSize(sheet, targetRange.bottom, targetRange.right);
  writeWorkbookQueryRows(sheet, rows, 1, 1);
  const query = cloneWorkbookQueryForSheet(ref.query, sheet, {
    range: targetRange,
    addToDataModel: options.addToDataModel,
    loadMode: "table",
  });
  sheet.queries = [query];
  registerQueryTableLoad(sheet, targetRange, query);
  internCellsInRange(sheet, targetRange);
  state.model.sheets.push(sheet);
  const sheetIndex = state.model.sheets.length - 1;
  query.range = { ...targetRange, sheetIndex };
  if (query.addToDataModel) registerQueryInDataModel(sheetIndex, query);
  state.activeSheetIndex = sheetIndex;
  state.selectedSheetIndexes = [];
  state.hf = buildFormulaEngine(state.model);
  closeDataDialog();
  state.selected = { sheetIndex, row: 1, col: 1 };
  state.selectionRange = { ...query.range };
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = state.selected;
  state.selectionAnchor = state.selected;
  state.selectedImage = null;
  $sheetHost.scrollTop(0);
  $sheetHost.scrollLeft(0);
  $horizontalScrollbar.scrollLeft(0);
  renderWorkbook();
  setStatus(`${query.name || "Query"} を新しいワークシートに読み込みました。`);
  pushSheetInsertHistory(sheetIndex, sheet, activeSheetBefore, sheetIndex, `${sheetName} 読み込み`, {
    dataModelBefore,
    dataModelAfter: cloneWorkbookDataModel(),
  });
  scheduleWorkbookPersistence();
}



function workbookQueryRowsForImport(ref, visited = new Set()) {
  const { sheet, query } = ref;
  if (query?.id && visited.has(query.id)) {
    return { type: "values", rows: [], height: 0, width: 0, typeDetection: "first200" };
  }
  const nextVisited = new Set(visited);
  if (query?.id) nextVisited.add(query.id);
  if (query.sourceType === "blankQuery") {
    return { type: "values", rows: [], height: 0, width: 0, typeDetection: "first200" };
  }
  if (query.sourceType === "queryReference") {
    const sourceRef = queryReferenceSourceRef(query);
    if (!sourceRef) return { type: "values", rows: [], height: 0, width: 0, typeDetection: query.typeDetection || "first200" };
    const sourceRows = workbookQueryRowsForImport(sourceRef, nextVisited);
    return workbookQueryRowsWithTransforms(query, workbookQueryRowsToSourceValues(sourceRows));
  }
  if (query.sourceType === "queryAppend") {
    return workbookQueryRowsWithTransforms(query, queryAppendSourceRows(query, nextVisited));
  }
  if (query.sourceType === "queryMerge") {
    return workbookQueryRowsWithTransforms(query, queryMergeSourceRows(query, nextVisited));
  }
  if (query.sourceType === "web" && query.sourceText != null) {
    const values = webImportRowsFromSource(query.sourceText, query.contentType || "", query.sourceUrl || "Web", query.sourceItemIndex ?? 0);
    return workbookQueryRowsWithTransforms(query, values);
  }
  if ((query.sourceType === "json" || query.sourceType === "xml") && query.sourceText != null) {
    const contentType = query.contentType || (query.sourceType === "xml" ? "application/xml" : "application/json");
    const values = webImportRowsFromSource(query.sourceText, contentType, query.sourceUrl || query.sourceName || query.sourceType.toUpperCase(), query.sourceItemIndex ?? 0);
    return workbookQueryRowsWithTransforms(query, values);
  }
  if (query.sourceType === "odata" && query.sourceText != null) {
    const values = odataImportRowsFromSource(query.sourceText, query.contentType || "", query.sourceUrl || query.sourceName || "OData", query.sourceItemIndex ?? 0);
    return workbookQueryRowsWithTransforms(query, values);
  }
  if (query.sourceType === "textCsv" && query.sourceText != null) {
    const delimiter = query.delimiter || detectDelimitedTextDelimiter(query.sourceText, query.sourceName || "CSV");
    const values = parseDelimitedRows(query.sourceText, delimiter).filter((row, index, list) => index < list.length - 1 || row.some((value) => value !== ""));
    return workbookQueryRowsWithTransforms(query, values);
  }
  if ((query.sourceType === "workbook" || query.sourceType === "folder") && Array.isArray(query.sourceRows)) {
    return workbookQueryRowsWithTransforms(query, query.sourceRows);
  }
  if (query.sourceType === "tableRange" && query.transformSteps?.length) {
    const sourceSheet = state.model?.sheets?.[query.sourceSheetIndex] ||
      state.model?.sheets?.find((candidate) => candidate.name === query.sourceSheetName) ||
      sheet;
    const values = sourceSheet && query.sourceRange
      ? sheetDisplayRowsForRange(sourceSheet, query.sourceRange)
      : Array.isArray(query.sourceRows)
        ? query.sourceRows
        : [];
    return workbookQueryRowsWithTransforms(query, values);
  }
  if (query.range) {
    const rows = [];
    for (let row = query.range.top; row <= query.range.bottom; row += 1) {
      const cells = [];
      for (let col = query.range.left; col <= query.range.right; col += 1) {
        cells.push(cloneCellModel(sheet.cells[cellKey(row, col)]));
      }
      rows.push(cells);
    }
    return { type: "cells", rows, height: rows.length, width: rows[0]?.length || 0 };
  }
  return { type: "values", rows: [], height: 0, width: 0, typeDetection: "first200" };
}



function workbookQueryRowsToSourceValues(rows) {
  if (!rows?.rows?.length) return [];
  if (rows.type === "cells") {
    return rows.rows.map((row) => row.map((cell) => cellDisplayText(cell)));
  }
  return rows.rows.map((row) => row.map((value) => String(value ?? "")));
}



function cloneWorkbookQueryForSheet(query, sheet, overrides = {}) {
  const loadedAt = new Date().toISOString();
  return {
    ...cloneJsonLikeValue(query),
    id: uniqueWorksheetObjectId("query"),
    name: nextWorkbookQueryName(query?.name || "Query"),
    range: overrides.range ? { ...overrides.range } : query?.range ? { ...query.range } : null,
    loadMode: overrides.loadMode || query?.loadMode || "table",
    addToDataModel: Boolean(overrides.addToDataModel),
    loadedAt,
    lastRefreshed: query?.lastRefreshed || loadedAt,
  };
}



function nextWorkbookQueryName(baseName = "Query") {
  const base = String(baseName || "Query").replace(/\s+\(\d+\)$/, "");
  const existing = new Set(workbookQueries().map(({ query }) => String(query.name || "").toLocaleLowerCase("ja-JP")));
  if (!existing.has(base.toLocaleLowerCase("ja-JP"))) return base;
  let index = 2;
  let name = `${base} (${index})`;
  while (existing.has(name.toLocaleLowerCase("ja-JP"))) {
    index += 1;
    name = `${base} (${index})`;
  }
  return name;
}



function workbookQueryValueRows(values, typeDetection = "first200", options = {}) {
  const preserveTrailingBlankRows = Boolean(options.preserveTrailingBlankRows);
  const rows = preserveTrailingBlankRows
    ? (values || [])
    : (values || []).filter((row, index, list) => index < list.length - 1 || row.some((value) => value !== ""));
  return {
    type: "values",
    rows,
    height: rows.length,
    width: rows.length ? Math.max(...rows.map((row) => row.length), 1) : 0,
    typeDetection,
    columnTypes: Array.isArray(options.columnTypes) ? [...options.columnTypes] : null,
    hasHeader: Boolean(options.hasHeader),
  };
}



function writeWorkbookQueryRows(sheet, rows, startRow, startCol) {
  const sheetId = getSheetId(sheet.name);
  for (let rowIndex = 0; rowIndex < rows.height; rowIndex += 1) {
    for (let colIndex = 0; colIndex < rows.width; colIndex += 1) {
      const targetRow = startRow + rowIndex;
      const targetCol = startCol + colIndex;
      if (rows.type === "cells") {
        writeClonedCellContent(sheet, targetRow, targetCol, rows.rows[rowIndex]?.[colIndex] || null);
      } else {
        const value = coerceQueryImportValue(
          rows.rows[rowIndex]?.[colIndex] ?? "",
          rows.typeDetection,
          rows.hasHeader && rowIndex === 0 ? "" : rows.columnTypes?.[colIndex],
        );
        if (sheetId == null) updateCellModel(sheet, targetRow, targetCol, value);
        else writeCellContent(sheet, sheetId, targetRow, targetCol, value);
      }
    }
  }
}



function coerceQueryImportValue(value, typeDetection = "first200", columnType = "") {
  const text = String(value ?? "");
  switch (queryTransformNormalizeDataType(columnType)) {
    case "text":
      return text;
    case "number": {
      const numeric = Number(text.replace(/,/g, ""));
      return Number.isFinite(numeric) ? numeric : text;
    }
    case "whole": {
      const numeric = Number(text.replace(/,/g, ""));
      return Number.isFinite(numeric) ? Math.trunc(numeric) : text;
    }
    case "date":
    case "datetime":
      return coerceUserInput(text);
    case "any":
    default:
      return coerceDelimitedImportValue(text, typeDetection);
  }
}



function registerQueryTableLoad(sheet, range, query) {
  applyExcelTableStyle(sheet, range);
  sheet.tables = Array.isArray(sheet.tables) ? sheet.tables : [];
  const table = sheet.tables.find((item) => item.range && rangesEqual(item.range, range));
  if (table) {
    table.showHeaderRow = query.hasHeader !== false;
    table.showFilterButton = true;
  } else {
    sheet.tables.push({
      id: uniqueWorksheetObjectId("table"),
      name: nextTableName(sheet),
      range: { ...range },
      showHeaderRow: query.hasHeader !== false,
      showFilterButton: true,
      style: "TableStyleMedium2",
    });
  }
}



function ensureWorkbookDataModel() {
  if (!state.model) return null;
  const model = state.model.dataModel && typeof state.model.dataModel === "object"
    ? state.model.dataModel
    : { tables: [] };
  model.tables = Array.isArray(model.tables) ? model.tables : [];
  state.model.dataModel = model;
  return model;
}



function cloneWorkbookDataModel() {
  return cloneJsonLikeValue(state.model?.dataModel || { tables: [] });
}



function restoreWorkbookDataModel(snapshot) {
  if (!state.model || !snapshot) return;
  state.model.dataModel = cloneJsonLikeValue(snapshot);
}



function registerQueryInDataModel(sheetIndex, query) {
  const dataModel = ensureWorkbookDataModel();
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!dataModel || !sheet || !query) return null;
  const range = query.range || null;
  const now = new Date().toISOString();
  const existingIndex = dataModel.tables.findIndex((item) => item.queryId === query.id || item.name === query.name);
  const entry = {
    id: existingIndex >= 0 ? dataModel.tables[existingIndex].id : uniqueWorksheetObjectId("dataModel"),
    queryId: query.id,
    name: query.name || "Query",
    sourceType: query.sourceType || "",
    sheetName: sheet.name,
    range: range ? rangeToLabel(range) : "",
    connection: queryConnectionString(query, sheet),
    columns: range ? dataModelColumnNames(sheet, range) : [],
    createdAt: existingIndex >= 0 ? dataModel.tables[existingIndex].createdAt || now : now,
    updatedAt: now,
  };
  if (existingIndex >= 0) dataModel.tables[existingIndex] = entry;
  else dataModel.tables.push(entry);
  return entry;
}



function dataModelColumnNames(sheet, range) {
  const sheetId = getSheetId(sheet.name);
  const names = [];
  for (let col = range.left; col <= range.right; col += 1) {
    names.push(String(getDisplayForCell(sheet, sheetId, range.top, col) || columnName(col)));
  }
  return names;
}



function syncWorkbookDataModelTables() {
  const dataModel = ensureWorkbookDataModel();
  if (!dataModel) return { tables: [], refs: [] };
  const refs = workbookQueries().filter(({ query }) => query?.addToDataModel);
  refs.forEach(({ sheetIndex, query }) => registerQueryInDataModel(sheetIndex, query));
  const validQueryIds = new Set(refs.map(({ query }) => query.id));
  dataModel.tables = (dataModel.tables || []).filter((table) => validQueryIds.has(table.queryId));
  return { tables: dataModel.tables, refs };
}



function openDataModelDialog(selectedId = "") {
  if (!state.model) {
    setStatus("Excel ファイルを開いてからデータ モデルを管理してください。");
    return;
  }
  const { tables } = syncWorkbookDataModelTables();
  const canManageRelationships = canCreateWorkbookRelationship();
  const selected = tables.some((table) => table.id === selectedId) ? selectedId : tables[0]?.id || "";
  state.dataDialog = { type: "data-model", selectedId: selected };
  closeDataDialog({ keepState: true });
  const rows = tables.length
    ? tables.map((table) => `<option value="${escapeAttr(table.id)}"${table.id === selected ? " selected" : ""}>${escapeHtml(table.name || "Table")}</option>`).join("")
    : '<option value="">データ モデルにテーブルはありません</option>';
  const table = dataModelTableById(selected);
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog data-model-dialog" role="dialog" aria-modal="true" aria-labelledby="dataModelTitle" data-data-dialog="data-model">
      <div class="excel-mini-titlebar" id="dataModelTitle">データ モデルの管理</div>
      <div class="excel-mini-body data-dialog-body data-model-body">
        <label class="data-model-list-wrap">
          <span class="excel-dialog-caption">モデル内のテーブル</span>
          <select class="data-model-table-list" size="8" data-data-model-table-list data-data-field="dataModelTableList">${rows}</select>
        </label>
        <div class="data-model-details">
          <span class="excel-dialog-caption">詳細</span>
          <div class="data-model-details-box" data-data-model-details>${dataModelDetailsHtml(table)}</div>
	          <div class="data-model-relationships">
	            <span class="excel-dialog-caption">リレーションシップ</span>
	            <div data-data-model-relationships>${dataModelRelationshipsHtml()}</div>
	          </div>
	        </div>
	        <div class="data-model-actions">
	          <button class="excel-mini-button" type="button" data-data-dialog-action="data-model-pivot" ${table ? "" : "disabled"}>ピボットテーブル(P)</button>
	          <button class="excel-mini-button" type="button" data-data-dialog-action="data-model-select" ${table ? "" : "disabled"}>場所へ移動(G)</button>
	          <button class="excel-mini-button" type="button" data-data-dialog-action="data-model-load-to" ${table ? "" : "disabled"}>読み込み先(L)...</button>
	          <button class="excel-mini-button" type="button" data-data-dialog-action="data-model-refresh" ${table ? "" : "disabled"}>更新(R)</button>
	          <button class="excel-mini-button" type="button" data-data-dialog-action="data-model-refresh-all" ${tables.length ? "" : "disabled"}>すべて更新(A)</button>
	          <button class="excel-mini-button" type="button" data-data-dialog-action="data-model-properties" ${table ? "" : "disabled"}>プロパティ(O)...</button>
	          <button class="excel-mini-button" type="button" data-data-dialog-action="data-model-relationships" ${canManageRelationships ? "" : "disabled"}>リレーションシップ(S)...</button>
	          <button class="excel-mini-button" type="button" data-data-dialog-action="data-model-remove" ${table ? "" : "disabled"}>モデルから削除(D)</button>
	        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">閉じる</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector("[data-data-model-table-list]")?.focus({ preventScroll: true });
}



function dataModelTableById(id) {
  const tables = ensureWorkbookDataModel()?.tables || [];
  return tables.find((table) => table.id === id) || null;
}



function dataModelQueryRef(table) {
  if (!table) return null;
  return workbookQueries().find(({ query }) => query.id === table.queryId || query.name === table.name) || null;
}



function dataModelDetailsHtml(table) {
  if (!table) return "データ モデルには、まだテーブルが追加されていません。";
  const ref = dataModelQueryRef(table);
  const stats = dataModelTableStats(table, ref);
  const status = dataModelTableStatus(table, ref);
  const lines = [
    `名前: ${table.name || ""}`,
    `状態: ${status}`,
    `接続: ${table.connection || ""}`,
    `種類: ${ref?.query ? queryConnectionTypeLabel(ref.query) : "接続が見つかりません"}`,
    `読み込み先: ${ref?.query ? queryLoadDestinationLabel(ref.query, ref.sheet) : "未読み込み"}`,
    `場所: ${table.sheetName || ""}${table.range ? `!${table.range}` : ""}`,
    `行数: ${stats.rows}`,
    `列数: ${stats.columns}`,
    `列: ${(table.columns || []).join(", ") || "(なし)"}`,
  ];
  if (table.updatedAt || table.createdAt) lines.push(`更新日時: ${formatQueryTimestamp(table.updatedAt || table.createdAt)}`);
  return escapeHtml(lines.join("\n"));
}



function dataModelTableStatus(table, ref = dataModelQueryRef(table)) {
  if (!ref?.query) return "接続が見つかりません";
  if (ref.query.lastRefreshStatus === "error") return "前回の更新でエラー";
  if (ref.query.lastRefreshStatus === "refreshing") return "更新中";
  if (table.updatedAt || ref.query.lastRefreshed) return "最新";
  return "未更新";
}



function dataModelTableStats(table, ref = dataModelQueryRef(table)) {
  const fallbackColumns = Array.isArray(table?.columns) ? table.columns.length : 0;
  if (!ref?.query?.range) return { rows: 0, columns: fallbackColumns };
  const range = ref.query.range;
  return {
    rows: Math.max(0, range.bottom - range.top),
    columns: Math.max(0, range.right - range.left + 1),
  };
}



function dataModelRelationshipsHtml() {
  const relationships = workbookRelationships();
  if (!relationships.length) return '<div class="data-model-empty">リレーションシップはありません。</div>';
  return relationships
    .map((relationship) => `<div class="data-model-relationship-row">${escapeHtml(relationshipSummaryLabel(relationship))}</div>`)
    .join("");
}



function selectedDataModelTableFromDialog(dialog) {
  const id = dialog?.querySelector("[data-data-model-table-list]")?.value || state.dataDialog?.selectedId || "";
  state.dataDialog = state.dataDialog && typeof state.dataDialog === "object" ? state.dataDialog : {};
  state.dataDialog.selectedId = id;
  return dataModelTableById(id);
}



function handleDataModelDialogSync(event) {
  const dialog = event.target.closest("[data-data-dialog='data-model']");
  if (!dialog) return;
  const table = selectedDataModelTableFromDialog(dialog);
  const details = dialog.querySelector("[data-data-model-details]");
  if (details) details.innerHTML = dataModelDetailsHtml(table);
  const hasTables = Boolean((ensureWorkbookDataModel()?.tables || []).length);
  const canManageRelationships = canCreateWorkbookRelationship();
  dialog.querySelectorAll("[data-data-dialog-action^='data-model-']").forEach((button) => {
    if (button.dataset.dataDialogAction === "data-model-relationships") {
      button.disabled = !canManageRelationships;
      return;
    }
    if (button.dataset.dataDialogAction === "data-model-refresh-all") {
      button.disabled = !hasTables;
      return;
    }
    button.disabled = !table;
  });
}



async function refreshDataModelTable(table) {
  const ref = dataModelQueryRef(table);
  if (!ref) {
    showMessageBox("更新する接続が見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  await refreshSingleQuery(ref.sheetIndex, ref.query.id);
  registerQueryInDataModel(ref.sheetIndex, ref.query);
  openDataModelDialog(dataModelTableById(table.id)?.id || "");
}



async function refreshAllDataModelTables(selectedId = "") {
  const tables = syncWorkbookDataModelTables().tables || [];
  const refs = tables
    .map((table) => ({ table, ref: dataModelQueryRef(table) }))
    .filter((item) => item.ref);
  if (!refs.length) {
    showMessageBox("更新するデータ モデル接続が見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  setStatus(`${refs.length} 件のデータ モデル テーブルを更新しています...`);
  for (const { ref } of refs) {
    await refreshSingleQuery(ref.sheetIndex, ref.query.id);
    registerQueryInDataModel(ref.sheetIndex, ref.query);
  }
  openDataModelDialog(selectedId || refs[0].table.id);
  setStatus(`${refs.length} 件のデータ モデル テーブルを更新しました。`);
}



function createPivotFromDataModelTable(table) {
  const ref = dataModelQueryRef(table);
  if (!ref) {
    showMessageBox("ピボットテーブルに使用する接続が見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  createPivotReportFromQuery(ref, { newWorksheet: true, addToDataModel: true });
}



function removeDataModelTable(table) {
  const ref = dataModelQueryRef(table);
  const dataModel = ensureWorkbookDataModel();
  if (!dataModel || !table) return;
  const dataModelBefore = cloneWorkbookDataModel();
  const relationshipsBefore = cloneRelationships(workbookRelationships());
  const sheetBefore = ref ? captureSheetForHistory(ref.sheet) : null;
  const removedRelationshipCount = removeRelationshipsForDataModelTable(table);
  dataModel.tables = dataModel.tables.filter((item) => item.id !== table.id);
  if (ref?.query) ref.query.addToDataModel = false;
  pushHistory({
    sheetIndex: ref?.sheetIndex ?? state.activeSheetIndex,
    ...(sheetBefore ? { sheetBefore, sheetAfter: captureSheetForHistory(ref.sheet) } : {}),
    dataModelBefore,
    dataModelAfter: cloneWorkbookDataModel(),
    relationshipsBefore,
    relationshipsAfter: cloneRelationships(workbookRelationships()),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label: `${table.name || "テーブル"} データ モデル削除`,
  });
  scheduleWorkbookPersistence();
  openDataModelDialog();
  setStatus(`${table.name || "テーブル"} をデータ モデルから削除しました。${removedRelationshipCount ? ` 関連するリレーションシップ ${removedRelationshipCount} 件も削除しました。` : ""}`);
}



function dataModelTableRelationshipKeys(table) {
  const ref = dataModelQueryRef(table);
  if (!ref) return [];
  return workbookTableRefs()
    .filter((tableRef) =>
      tableRef.sheetIndex === ref.sheetIndex &&
      (
        tableRef.table?.id === ref.query.tableId ||
        tableRef.table?.name === ref.query.tableName ||
        rangesEqualLoose(tableRef.table?.range, ref.query.range)
      ),
    )
    .map((tableRef) => tableRef.key);
}



function removeRelationshipsForDataModelTable(table) {
  const keys = new Set(dataModelTableRelationshipKeys(table));
  if (!keys.size) return 0;
  const before = workbookRelationships().length;
  state.model.relationships = workbookRelationships().filter((relationship) =>
    !keys.has(relationship.sourceTableKey) && !keys.has(relationship.targetTableKey),
  );
  return before - workbookRelationships().length;
}



function createPivotReportFromQuery(ref, options = {}) {
  const querySheet = queryRowsToTemporarySheet(ref);
  if (!querySheet.range) {
    showMessageBox("ピボットテーブルを作成するためのデータが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  const summary = pivotSummaryFromRange(querySheet.sheet, querySheet.range);
  if (!summary.rows.length) {
    showMessageBox("ピボットテーブルを作成するためのデータが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  if (options.newWorksheet) {
    createPivotReportSheetFromSummary(ref, summary, options);
    return;
  }
  const destination = decodeWorkbookRange(options.destinationText || "");
  if (!destination) {
    showMessageBox("ピボットテーブル レポートを配置する場所を正しく入力してください。", { title: "Cht WebSheet" });
    return;
  }
  writePivotReportToExistingSheet(ref, summary, {
    ...options,
    targetSheetIndex: destination.sheetIndex ?? state.activeSheetIndex,
    targetRow: destination.top,
    targetCol: destination.left,
  });
}



function queryRowsToTemporarySheet(ref) {
  const rows = workbookQueryRowsForImport(ref);
  if (!rows.height || !rows.width) return { sheet: null, range: null };
  const sheet = createBlankSheet("__QueryPreview");
  ensureSheetSize(sheet, rows.height, rows.width);
  writeWorkbookQueryRows(sheet, rows, 1, 1);
  const range = normalizeRange(
    { sheetIndex: state.activeSheetIndex, row: 1, col: 1 },
    { sheetIndex: state.activeSheetIndex, row: rows.height, col: rows.width },
  );
  return { sheet, range };
}



function createPivotReportSheetFromSummary(ref, summary, options = {}) {
  const activeSheetBefore = state.activeSheetIndex;
  const sourceSnapshotsBefore = options.addToDataModel ? captureSheetsForHistory([ref.sheetIndex]) : [];
  const dataModelBefore = cloneWorkbookDataModel();
  const finalName = uniqueSheetNameWithBase(options.chart ? "PivotChart" : "PivotTable", state.model.sheets.map((item) => item.name));
  const sheet = createBlankSheet(finalName);
  const insertIndex = state.model.sheets.length;
  ensureSheetSize(sheet, summary.rows.length + 18, Math.max(summary.headers.length + 8, 10));
  const pivotRange = writePivotSummarySheet(sheet, summary, insertIndex);
  sheet.pivotSources = [{ queryName: ref.query.name || "Query", connection: queryConnectionString(ref.query, ref.sheet), range: rangeToLabel(ref.query.range || pivotRange) }];
  if (options.addToDataModel) {
    ref.query.addToDataModel = true;
    registerQueryInDataModel(ref.sheetIndex, ref.query);
  }
  if (options.chart) {
    appendPivotChartToSheet(sheet, summary, pivotRange, {
      row: 1,
      col: Math.max(summary.headers.length + 2, 5),
      sourceSheetName: finalName,
    });
  }
  state.model.sheets.push(sheet);
  state.activeSheetIndex = insertIndex;
  state.selectedSheetIndexes = [];
  state.hf = buildFormulaEngine(state.model);
  closeDataDialog();
  const point = { sheetIndex: insertIndex, row: 1, col: 1 };
  state.selected = point;
  state.selectionRange = selectionRangeFromPoints(point, point);
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = point;
  state.selectionAnchor = point;
  state.selectedImage = null;
  $sheetHost.scrollTop(0);
  $sheetHost.scrollLeft(0);
  $horizontalScrollbar.scrollLeft(0);
  renderWorkbook();
  setStatus(`${finalName} に${options.chart ? "ピボットグラフ" : "ピボットテーブル"}を作成しました。`);
  pushSheetInsertHistory(insertIndex, sheet, activeSheetBefore, insertIndex, `${finalName} ${options.chart ? "ピボットグラフ" : "ピボットテーブル"}`, {
    sheetsBefore: sourceSnapshotsBefore,
    sheetsAfter: options.addToDataModel ? captureSheetsForHistory([ref.sheetIndex]) : [],
    dataModelBefore,
    dataModelAfter: cloneWorkbookDataModel(),
  });
  scheduleWorkbookPersistence();
}



function writePivotReportToExistingSheet(ref, summary, options = {}) {
  const sheet = state.model?.sheets?.[options.targetSheetIndex];
  if (!sheet) return;
  const width = summary.headers.length;
  const height = summary.rows.length + 1;
  const targetRange = normalizeRange(
    { sheetIndex: options.targetSheetIndex, row: options.targetRow, col: options.targetCol },
    { sheetIndex: options.targetSheetIndex, row: options.targetRow + height - 1, col: options.targetCol + width - 1 },
  );
  if (!options.overwriteConfirmed && rangeHasUserData(sheet, targetRange)) {
    showDataOverwriteConfirmation(() => writePivotReportToExistingSheet(ref, summary, { ...options, overwriteConfirmed: true }));
    return;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  const dataModelBefore = cloneWorkbookDataModel();
  writePivotSummaryAt(sheet, summary, options.targetRow, options.targetCol, options.targetSheetIndex);
  if (options.chart) {
    appendPivotChartToSheet(sheet, summary, targetRange, {
      row: options.targetRow,
      col: options.targetCol + width + 2,
      sourceSheetName: sheet.name,
    });
  }
  if (options.addToDataModel) {
    ref.query.addToDataModel = true;
    registerQueryInDataModel(ref.sheetIndex, ref.query);
  }
  state.activeSheetIndex = options.targetSheetIndex;
  state.hf = buildFormulaEngine(state.model);
  closeDataDialog();
  pushHistory({
    sheetIndex: options.targetSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    dataModelBefore,
    dataModelAfter: cloneWorkbookDataModel(),
    selectionBefore: state.selectionRange || null,
    selectionAfter: targetRange,
    label: `${ref.query.name || "Query"} ${options.chart ? "ピボットグラフ" : "ピボットテーブル"}`,
  });
  state.selected = { sheetIndex: options.targetSheetIndex, row: targetRange.top, col: targetRange.left };
  state.selectionRange = { ...targetRange };
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = state.selected;
  renderWorkbook();
  setStatus(`${rangeToLabel(targetRange)} に${options.chart ? "ピボットグラフ" : "ピボットテーブル"}を作成しました。`);
  scheduleWorkbookPersistence();
}



function appendPivotChartToSheet(sheet, summary, pivotRange, options = {}) {
  const chart = chartDataFromRange(sheet, pivotRange, "column");
  if (!chart.series.length) return null;
  const descriptor = chartDescriptor(chart, "column", pivotRange, {
    pivot: true,
    title: "ピボットグラフ",
    sourceSheetName: options.sourceSheetName || sheet.name,
  });
  return appendImageDescriptorToSheet(sheet, descriptor, { row: options.row || 1, col: options.col || Math.max(4, pivotRange.right + 2) });
}



function appendImageDescriptorToSheet(sheet, descriptor, anchor) {
  if (!descriptor?.src) return null;
  const size = fitObjectSize(descriptor.width, descriptor.height);
  ensureSheetSize(sheet, anchor.row, anchor.col);
  const end = endingCellForPixelSize(sheet, anchor.row, anchor.col, size.width, size.height);
  ensureSheetSize(sheet, end.row, end.col);
  const image = {
    id: uniqueWorksheetObjectId("image"),
    name: descriptor.name || "グラフ",
    src: descriptor.src,
    mimeType: descriptor.mimeType || "image/svg+xml",
    extension: descriptor.extension || "svg",
    editAs: "oneCell",
    tl: {
      col: anchor.col - 1,
      row: anchor.row - 1,
      nativeCol: anchor.col - 1,
      nativeRow: anchor.row - 1,
      nativeColOff: 0,
      nativeRowOff: 0,
    },
    br: null,
    ext: { width: size.width, height: size.height },
    hyperlink: "",
  };
  if (descriptor.chart) image.chart = cloneJsonLikeValue(descriptor.chart);
  if (descriptor.shape) image.shape = cloneJsonLikeValue(descriptor.shape);
  sheet.images = Array.isArray(sheet.images) ? sheet.images : [];
  sheet.images.push(image);
  return image;
}



function openQueryPropertiesDialog(sheetIndex, queryId, options = {}) {
  const sheet = state.model?.sheets?.[sheetIndex];
  const query = sheet?.queries?.find((item) => item.id === queryId);
  if (!sheet || !query) return;
  state.dataDialog = {
    type: "query-properties",
    sheetIndex,
    queryId,
    returnToExistingConnectionsKey: options.returnToExistingConnections ? workbookQueryKey(sheetIndex, queryId) : "",
  };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog query-properties-dialog" role="dialog" aria-modal="true" aria-labelledby="queryPropertiesTitle" data-data-dialog="query-properties">
      <div class="excel-mini-titlebar" id="queryPropertiesTitle">接続のプロパティ</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="data-dialog-tabs query-properties-tabs" role="tablist" aria-label="接続のプロパティ">
          <button class="is-active" type="button" role="tab" aria-selected="true" data-query-properties-tab="usage" data-data-dialog-action="query-properties-tab">使用状況</button>
          <button type="button" role="tab" aria-selected="false" data-query-properties-tab="definition" data-data-dialog-action="query-properties-tab">定義</button>
        </div>
        <div class="data-dialog-tab-panel query-properties-panel is-active" role="tabpanel" data-query-properties-tab-panel="usage">
          <fieldset class="data-dialog-fieldset query-properties-fieldset">
            <legend>更新の制御</legend>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="backgroundRefresh" ${query.backgroundRefresh ? "checked" : ""} />バックグラウンドで更新する</label>
            <label class="data-dialog-check query-properties-interval"><input type="checkbox" data-data-field="refreshEveryEnabled" data-query-properties-sync ${query.refreshEveryEnabled ? "checked" : ""} />次の間隔で更新する
              <input type="number" min="1" max="32767" step="1" data-data-field="refreshIntervalMinutes" value="${escapeAttr(String(queryRefreshIntervalMinutes(query)))}" ${query.refreshEveryEnabled ? "" : "disabled"} />
              <span>分</span>
            </label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="refreshOnOpen" ${query.refreshOnOpen ? "checked" : ""} />ファイルを開くときにデータを更新する</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="removeDataBeforeSave" ${query.removeDataBeforeSave ? "checked" : ""} />ブックの保存前に外部データ範囲からデータを削除する</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="refreshWithAll" ${query.refreshWithAll !== false ? "checked" : ""} />すべて更新でこの接続を更新する</label>
          </fieldset>
          <div class="data-dialog-note">種類: ${escapeHtml(queryConnectionTypeLabel(query))}${query.range ? ` / 範囲: ${escapeHtml(rangeToLabel(query.range))}` : ""}</div>
        </div>
        <div class="data-dialog-tab-panel query-properties-panel" role="tabpanel" data-query-properties-tab-panel="definition">
          <fieldset class="data-dialog-fieldset query-properties-fieldset">
            <legend>接続名</legend>
            <label class="excel-dialog-field"><span class="excel-dialog-caption">名前</span><input type="text" data-data-field="name" value="${escapeAttr(query.name || "Query")}" /></label>
            <label class="excel-dialog-field"><span class="excel-dialog-caption">説明</span><textarea data-data-field="description" rows="3">${escapeHtml(query.description || "")}</textarea></label>
          </fieldset>
          <fieldset class="data-dialog-fieldset query-properties-fieldset">
            <legend>接続定義</legend>
            <label class="excel-dialog-field"><span class="excel-dialog-caption">接続ファイル</span><input type="text" data-data-field="connectionFile" value="${escapeAttr(query.connectionFile || query.sourceName || "")}" /></label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="alwaysUseConnectionFile" ${query.alwaysUseConnectionFile ? "checked" : ""} />常に接続ファイルを使用する</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="failIfConnectionFileMissing" ${query.failIfConnectionFileMissing ? "checked" : ""} />接続ファイルを使用できない場合は更新しない</label>
            <label class="excel-dialog-field"><span class="excel-dialog-caption">アイテム名</span><input type="text" data-data-field="sourceItemName" value="${escapeAttr(query.sourceItemName || "")}" /></label>
            <label class="excel-dialog-field"><span class="excel-dialog-caption">接続文字列</span><textarea data-data-field="connectionString" rows="3">${escapeHtml(queryConnectionString(query, sheet))}</textarea></label>
            <label class="excel-dialog-field"><span class="excel-dialog-caption">コマンドの種類</span><select data-data-field="commandType">${queryCommandTypeOptions(query)}</select></label>
            <label class="excel-dialog-field"><span class="excel-dialog-caption">コマンド テキスト</span><textarea data-data-field="commandText" rows="3">${escapeHtml(queryCommandText(query, sheet))}</textarea></label>
          </fieldset>
          ${queryAuthenticationPropertiesHtml(query)}
        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector("[data-data-field='backgroundRefresh']")?.focus({ preventScroll: true });
}



function queryAuthenticationPropertiesHtml(query) {
  if (!["web", "odata", "json", "xml"].includes(query?.sourceType)) return "";
  const isBasic = query.sourceAuthMode === "basic";
  const credential = query.sourceCredentialKey ? state.queryCredentials?.get?.(query.sourceCredentialKey) : null;
  const hasSessionCredential = Boolean(isBasic && credential);
  const username = query.sourceAuthUsername || credential?.username || "";
  const passwordHint = hasSessionCredential
    ? "空欄のまま OK すると現在のパスワードを保持します。"
    : "更新を有効にするにはパスワードを入力してください。";
  return `
    <fieldset class="data-dialog-fieldset query-properties-fieldset query-auth-fieldset">
      <legend>認証</legend>
      <div class="query-auth-mode-row">
        <label class="data-dialog-check"><input type="radio" name="queryAuthMode" value="anonymous" data-query-properties-sync ${isBasic ? "" : "checked"} />匿名</label>
        <label class="data-dialog-check"><input type="radio" name="queryAuthMode" value="basic" data-query-properties-sync ${isBasic ? "checked" : ""} />基本認証</label>
        ${query.sourceType === "odata" ? '<label class="data-dialog-check"><input type="radio" name="queryAuthMode" value="organization" disabled />組織アカウント</label>' : ""}
      </div>
      <div class="query-auth-basic-fields${isBasic ? "" : " hidden"}" data-query-auth-basic-fields>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">ユーザー名</span><input type="text" data-data-field="queryAuthUsername" data-query-properties-sync value="${escapeAttr(username)}" autocomplete="username" /></label>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">パスワード</span><input type="password" data-data-field="queryAuthPassword" data-query-properties-sync value="" autocomplete="current-password" placeholder="${escapeAttr(hasSessionCredential ? "********" : "")}" /></label>
        <div class="data-dialog-note" data-query-auth-session-status>${hasSessionCredential ? `現在のセッションで資格情報を保持しています。${passwordHint}` : `資格情報は保存されていません。${passwordHint}`}</div>
        <button class="excel-mini-button" type="button" data-data-dialog-action="query-properties-clear-credential" data-query-auth-can-clear="${hasSessionCredential ? "true" : "false"}" ${hasSessionCredential ? "" : "disabled"}>資格情報のクリア</button>
      </div>
    </fieldset>`;
}

export {
  activeWorkbookQueryRef,
  appendImageDescriptorToSheet,
  appendPivotChartToSheet,
  applyAppendWorkbookQueryDialog,
  applyConnectionOnlyImport,
  applyImportDataDialog,
  applyMergeWorkbookQueryDialog,
  applyQueryRenameDialog,
  applyQuerySourceChange,
  applyQuerySourceChangeDialog,
  cloneWorkbookDataModel,
  cloneWorkbookQueryForSheet,
  closeQueriesConnectionsPane,
  closeQueryPaneContextMenu,
  closeQueryPaneContextMenuOnOutsidePointer,
  closeQueryPaneOnEscape,
  coerceQueryImportValue,
  createPivotFromDataModelTable,
  createPivotReportFromQuery,
  createPivotReportSheetFromSummary,
  dataModelColumnNames,
  dataModelDetailsHtml,
  dataModelQueryRef,
  dataModelRelationshipsHtml,
  dataModelTableById,
  dataModelTableRelationshipKeys,
  dataModelTableStats,
  dataModelTableStatus,
  deleteWorkbookQuery,
  dependentWorkbookQueriesForQuery,
  duplicateWorkbookQuery,
  ensureWorkbookDataModel,
  existingConnectionDetailsHtml,
  formatQueryTimestamp,
  handleAppendQuerySelectionChange,
  handleDataModelDialogSync,
  handleExistingConnectionDialogSync,
  handleExistingConnectionsDialogAction,
  handleImportDataDialogSync,
  handleMergeQueryPreviewChange,
  handleMergeRightQueryChange,
  handleQueryPaneAction,
  handleQueryPaneContextMenuAction,
  htmlTableName,
  htmlTableRows,
  htmlTableSources,
  loadWebNavigatorSelection,
  loadWebRowsIntoSelection,
  loadWorkbookQueryToNewSheet,
  loadWorkbookQueryToRange,
  nextQueryName,
  nextWebQueryName,
  nextWorkbookQueryName,
  normalizeQuerySourceChangeValue,
  openActiveQueryPropertiesDialog,
  openAppendWorkbookQueryDialog,
  openDataModelDialog,
  openExistingConnectionsDialog,
  openImportDataDialog,
  openMergeWorkbookQueryDialog,
  openQueriesConnectionsPane,
  openQueryPaneContextMenuFromEvent,
  openQueryPropertiesDialog,
  openQueryRenameDialog,
  openQuerySourceChangeDialog,
  openQueryTransformEditorForQuery,
  openWebNavigatorDialog,
  openWebNavigatorTransformEditor,
  previousHtmlHeadingText,
  queriesConnectionsPaneHtml,
  queryAppendDialogPreviewHtml,
  queryAppendSourceRefs,
  queryAppendSourceRows,
  queryAuthenticationPropertiesHtml,
  queryAuthenticationStatusLabel,
  queryChangeSourceCaption,
  queryChangeSourceCurrentValue,
  queryCommandText,
  queryCommandType,
  queryCommandTypeLabel,
  queryCommandTypeOptions,
  queryConnectionRefreshLabel,
  queryConnectionString,
  queryConnectionStringAfterSourceChange,
  queryConnectionTypeLabel,
  queryHeaderIndex,
  queryLoadDestinationLabel,
  queryLoadedDimensions,
  queryMergeAntiRows,
  queryMergeCommandJoinKind,
  queryMergeDialogPreviewHtml,
  queryMergeKey,
  queryMergeRightColumns,
  queryMergeSourceRefs,
  queryMergeSourceRows,
  queryPaneConnectionRowHtml,
  queryPaneConnectionsHtml,
  queryPaneContextMenuHtml,
  queryPaneLoadModeLabel,
  queryPaneQueriesHtml,
  queryPaneRowHtml,
  queryReferenceSourceRef,
  queryReferenceSourceRows,
  queryRefreshIntervalMinutes,
  queryRefreshStatusShortLabel,
  queryRowsToTemporarySheet,
  querySortTime,
  querySourceChangeShouldClearCache,
  querySourceColumnsForRef,
  querySourceCommandText,
  querySourceIconName,
  querySourceTypeUsesStoredRows,
  queryStoredRef,
  referenceWorkbookQuery,
  refreshActiveWorkbookConnection,
  refreshAllDataModelTables,
  refreshDataModelTable,
  refreshQueriesConnectionsPane,
  refreshSingleQuery,
  registerQueryInDataModel,
  registerQueryTableLoad,
  removeDataModelTable,
  removeRelationshipsForDataModelTable,
  restoreWorkbookDataModel,
  selectQueryRange,
  selectWebNavigatorSource,
  selectedAppendQueryKeysFromDialog,
  selectedDataModelTableFromDialog,
  selectedExistingConnectionRef,
  syncWorkbookDataModelTables,
  updateMergeQueryDialogPreview,
  webNavigatorDialogHtml,
  webNavigatorPreviewHtml,
  webNavigatorSelectedSource,
  webNavigatorSourceButtonsHtml,
  webNavigatorSourceIconName,
  webNavigatorSummaryHtml,
  workbookQueries,
  workbookQueryDeleteMessage,
  workbookQueryDependsOnQuery,
  workbookQueryKey,
  workbookQueryRef,
  workbookQueryRefFromKey,
  workbookQueryRowsForImport,
  workbookQueryRowsToSourceValues,
  workbookQueryValueRows,
  writePivotReportToExistingSheet,
  writeRowsToSheet,
  writeWorkbookQueryRows,
  xmlElementText,
  xmlElementToRowObject,
  xmlRowsFromElements,
};
