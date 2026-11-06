// Feature module; shared bindings are imported explicitly.
import {
  $sheetHost,
  APP_ICON_DATA_URL,
  APP_ICON_VERSION,
  AUTO_STANDALONE_RUNTIME_BASE_URL,
  CELL_BORDER_SIDES,
  DEFAULT_SHAPE_TYPE,
  DEFAULT_STANDALONE_REMOTE_RUNTIME_BASE_URL,
  DEFAULT_STANDALONE_RUNTIME_BASE_URL,
  EMU_PER_PIXEL,
  JSZip,
  STANDALONE_REMOTE_RUNTIME_BASE_URL_STORAGE_KEY,
  STANDALONE_RUNTIME_BASE_URL_STORAGE_KEY,
  STANDALONE_RUNTIME_CHANNEL,
  STANDALONE_RUNTIME_VERSION,
  loadJsZipModule,
  state,
} from "../../app-core.js";
import {
  autoFilterDynamicDatePeriodType,
  escapeRegExp,
  normalizeAutoFilterCriterion,
} from "../data/filter-and-text-tools.js";
import {
  sortCustomListById,
  tableSortStateForFilterRange,
  visibleTableSortStates,
} from "../data/sort-and-autofilter.js";
import {
  workbookRelationships,
} from "../data/validation-and-consolidation.js";
import {
  chartSvg,
  escapeXml,
} from "../drawing/charts-and-insertion.js";
import {
  normalizedHeaderFooterSettings,
  shapeDescriptorForType,
} from "../drawing/shape-rendering.js";
import {
  normalizeColorValue,
  svgDataUrl,
} from "../editing/ink-and-clipboard.js";
import {
  activeSelectionRanges,
  autoFitCellDisplayText,
  autoFitCellTextHeight,
  getDisplayForCell,
  getCellRawInput,
  materializeSelectionRangeForSheet,
  rangeToLabel,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  canvasFontSize,
  imageBoxToPixels,
  roundCssNumber,
  shapeTextLayerHtml,
  shapeTextPlainValue,
  worksheetImageRenderSource,
} from "../editing/selection-and-cell-text.js";
import {
  cellEffectiveCss,
} from "../editing/structure-and-history.js";
import {
  createPixelPositionCache,
  normalizeRichTextLineEndings,
  paintCanvasCells,
  rangePixelRect,
} from "../rendering/grid-and-canvas.js";
import {
  annotationAuthor,
  normalizeCommentReplies,
} from "../ui/annotations-and-cell-format.js";
import {
  cellKey,
  columnName,
  columnWidthToPixels,
  currentRuntimeVersion,
  dateToExcelSerial,
  decodeCell,
  escapeAttr,
  escapeHtml,
  excelSerial1900To1904,
  excelSerialToDate,
  isDateFormat,
  normalizeVersionInfo,
  refreshServerVersionInfo,
  renderVersionDialogContent,
  rowHeightToPixels,
  runtimeExecutionSource,
  safeJson,
  serverVersionInfoUrl,
  styleObjectToString,
  syncVersionRibbonBadge,
  versionInfoIsMismatch,
  versionInfoRuntimeUpdatedAt,
  versionInfoRuntimeVersion,
} from "../ui/help-and-utilities.js";
import {
  printRangePixelBounds,
} from "../ui/printing-and-home-tools.js";
import {
  stripInternalHyperlinkPrefix,
} from "../ui/ribbon-and-links.js";
import {
  clampZoom,
  scaleSheetDisplay,
} from "../ui/sheet-protection-and-translation.js";
import {
  clampWorkbookSheetIndex,
  normalizeCustomTopLeftCell,
  normalizeCustomViewPoint,
  normalizeCustomViewRange,
  normalizeCustomViewScroll,
  normalizePaneState,
  normalizePreservedOoxmlPackage,
  normalizeWorkbookCalcProperties,
  normalizeWorkbookCustomView,
  normalizeWorkbookProperties,
  normalizeWorkbookViewModel,
  normalizedSheetShowRuler,
  persistedDateDisplayValue,
} from "./import-model.js";
import {
  parseWorkbookStyleDxfsXml,
  workbookSheetPaths,
} from "./ooxml-drawings.js";
import {
  buildMergeMaps,
  excelHeaderFooterText,
  excelSplitActivePane,
  excelSplitUnitsForPaneIndex,
  excelWorkbookViewEntries,
  excelWorksheetTabSelected,
  getSheetId,
  paintCanvasBorders,
  standaloneModelScriptPayload,
} from "./saving-and-export.js";
import {
  cwsAiInstructionsScript,
  cwsAiNoticeComment,
  cwsHtmlMetaTags,
  cwsHtmlRootAttributes,
  cwsModelScriptTag,
} from "./startup-and-loading.js";
import {
  autoFilterStdDevCount,
  base64ToBytes,
  normalizeNote,
  normalizeOutlineSettings,
  topLeftCellFromScrollRecord,
} from "./tables-and-validation.js";
import {
  normalizedPageLayoutSettings,
  printRenderScale,
  sheetUsedRangeForPrint,
} from "./vba-and-printing.js";


async function postProcessExcelWorkbookBuffer(buffer, model, type) {
  await loadJsZipModule();
  const zip = await JSZip.loadAsync(buffer);
  await patchExcelWorkbookDefinedNames(zip, model);
  await patchExcelPreservedOoxmlPackage(zip, model, type);
  await patchExcelWorkbookMacroContentType(zip, type);
  await patchExcelWorkbookMetadata(zip, model);
  await patchExcelWorkbookWindowViews(zip, model);
  await patchExcelWorkbookCustomViews(zip, model);
  await patchExcelWorksheetCustomViews(zip, model);
  await patchExcelWorksheetViewDetails(zip, model);
  await patchExcelWorksheetPageBreaks(zip, model);
  await patchExcelWorksheetOutlineDetails(zip, model);
  await patchExcelWorksheetDataValidations(zip, model);
  await patchExcelWorksheetInternalHyperlinks(zip, model);
  const autoFilterExportContext = await createExcelAutoFilterExportContext(zip);
  await patchExcelWorksheetAutoFilterDetails(zip, model, autoFilterExportContext);
  await patchExcelTableDefinitions(zip, model, autoFilterExportContext);
  await finalizeExcelAutoFilterExportContext(zip, autoFilterExportContext);
  return zip.generateAsync({ type: "arraybuffer" });
}



function hasPreservedVbaProject(model) {
  return normalizePreservedOoxmlPackage(model?.ooxmlPackage)?.hasVbaProject === true;
}



async function patchExcelWorkbookDefinedNames(zip, model) {
  const file = zip.file("xl/workbook.xml");
  if (!file) return;
  const entries = excelWorkbookDefinedNameXmlEntries(model);
  const workbookXml = await file.async("text");
  const nextXml = replaceExcelWorkbookDefinedNamesXml(workbookXml, entries);
  if (nextXml !== workbookXml) zip.file("xl/workbook.xml", nextXml);
}



function replaceExcelWorkbookDefinedNamesXml(workbookXml, entries = []) {
  const xml = String(workbookXml || "");
  const preserved = excelWorkbookPreservedDefinedNameXmlEntries(xml);
  const body = `${preserved.join("")}${entries.join("")}`;
  if (/<definedNames\b[\s\S]*?<\/definedNames>/.test(xml)) {
    return body
      ? xml.replace(/<definedNames\b[^>]*>[\s\S]*?<\/definedNames>/, `<definedNames>${body}</definedNames>`)
      : xml.replace(/<definedNames\b[^>]*>[\s\S]*?<\/definedNames>/, "");
  }
  if (!body) return xml;
  return xml.replace("</sheets>", `</sheets><definedNames>${body}</definedNames>`);
}



function excelWorkbookPreservedDefinedNameXmlEntries(workbookXml) {
  const block = String(workbookXml || "").match(/<definedNames\b[^>]*>([\s\S]*?)<\/definedNames>/)?.[1] || "";
  return [...block.matchAll(/<definedName\b[^>]*(?:\/>|>[\s\S]*?<\/definedName>)/g)]
    .map((match) => match[0])
    .filter((entry) => {
      const name = xmlAttributeValue(entry, "name");
      return /^_xlnm\./i.test(name);
    });
}



function excelWorkbookDefinedNameXmlEntries(model) {
  const seen = new Set();
  return (model?.definedNames || [])
    .map((definition) => {
      const name = String(definition?.name || "").trim();
      const refersTo = String(definition?.refersTo || "").trim().replace(/^=/, "");
      if (!name || !refersTo) return "";
      if (name === "_xlnm.Print_Area" || name === "_xlnm.Print_Titles") return "";
      const sheetIndex = Number(definition.sheetIndex);
      const isSheetScoped = definition.scope === "sheet" && Number.isInteger(sheetIndex) && sheetIndex >= 0;
      const key = `${isSheetScoped ? sheetIndex : "book"}:${name}`;
      if (seen.has(key)) return "";
      seen.add(key);
      const attributes = [
        `name="${escapeAttr(name)}"`,
        isSheetScoped ? `localSheetId="${sheetIndex}"` : "",
        definition.hidden ? 'hidden="1"' : "",
        definition.comment ? `comment="${escapeAttr(definition.comment)}"` : "",
      ].filter(Boolean).join(" ");
      return `<definedName ${attributes}>${escapeHtml(refersTo)}</definedName>`;
    })
    .filter(Boolean);
}



async function patchExcelWorkbookMacroContentType(zip, type) {
  if (type?.value !== "xlsm") return;
  const file = zip.file("[Content_Types].xml");
  if (!file) return;
  const xml = await file.async("text");
  const nextXml = xml.replace(
    /application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet\.main\+xml/g,
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
  );
  zip.file("[Content_Types].xml", nextXml);
}



async function patchExcelWorkbookMetadata(zip, model) {
  const file = zip.file("xl/workbook.xml");
  if (!file) return;
  const workbookXml = await file.async("text");
  let nextXml = upsertWorkbookPrXml(workbookXml, excelWorkbookPrXml(model));
  nextXml = upsertWorkbookCalcPrXml(nextXml, excelWorkbookCalcPrXml(model));
  zip.file("xl/workbook.xml", nextXml);
}



function excelWorkbookPrXml(model) {
  const properties = normalizeWorkbookProperties(model?.workbookProperties);
  const attributes = [
    properties.date1904 ? 'date1904="1"' : "",
    properties.codeName ? `codeName="${escapeAttr(properties.codeName)}"` : "",
    properties.filterPrivacy ? 'filterPrivacy="1"' : "",
    properties.showObjects !== "all" ? `showObjects="${escapeAttr(properties.showObjects)}"` : "",
    properties.defaultThemeVersion ? `defaultThemeVersion="${properties.defaultThemeVersion}"` : 'defaultThemeVersion="164011"',
  ].filter(Boolean);
  return `<workbookPr ${attributes.join(" ")}/>`;
}



function excelWorkbookCalcPrXml(model) {
  const properties = normalizeWorkbookCalcProperties(model?.calcProperties);
  const attributes = [
    properties.calcId != null ? `calcId="${Math.max(0, Math.round(Number(properties.calcId) || 0))}"` : "",
    properties.calcMode !== "auto" ? `calcMode="${escapeAttr(properties.calcMode)}"` : "",
    properties.refMode === "R1C1" ? 'refMode="R1C1"' : "",
    `fullCalcOnLoad="${properties.fullCalcOnLoad ? "1" : "0"}"`,
    properties.forceFullCalc ? 'forceFullCalc="1"' : "",
    `calcOnSave="${properties.calcOnSave ? "1" : "0"}"`,
    `concurrentCalc="${properties.concurrentCalc ? "1" : "0"}"`,
    properties.fullPrecision === false ? 'fullPrecision="0"' : "",
    properties.iterate ? 'iterate="1"' : "",
    properties.iterate ? `iterateCount="${Math.max(1, Math.round(Number(properties.iterateCount) || 100))}"` : "",
    properties.iterate ? `iterateDelta="${Number(properties.iterateDelta) || 0.001}"` : "",
  ].filter(Boolean);
  return `<calcPr ${attributes.join(" ")}/>`;
}



function upsertWorkbookPrXml(workbookXml, workbookPrXml) {
  const withoutExisting = String(workbookXml || "").replace(/<workbookPr\b[^>]*(?:\/>|>[\s\S]*?<\/workbookPr>)/, "");
  if (/<fileVersion\b[^>]*(?:\/>|>[\s\S]*?<\/fileVersion>)/.test(withoutExisting)) {
    return withoutExisting.replace(/<fileVersion\b[^>]*(?:\/>|>[\s\S]*?<\/fileVersion>)/, (match) => `${match}${workbookPrXml}`);
  }
  const beforeWorkbookPrFollowers = /(<workbookProtection\b|<bookViews\b|<sheets\b|<\/workbook>)/;
  return withoutExisting.replace(beforeWorkbookPrFollowers, `${workbookPrXml}$1`);
}



function upsertWorkbookCalcPrXml(workbookXml, calcPrXml) {
  const withoutExisting = String(workbookXml || "").replace(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/, "");
  const beforeCalcFollowers = /(<customWorkbookViews\b|<pivotCaches\b|<smartTagPr\b|<smartTagTypes\b|<webPublishing\b|<fileRecoveryPr\b|<webPublishObjects\b|<extLst\b|<\/workbook>)/;
  return withoutExisting.replace(beforeCalcFollowers, `${calcPrXml}$1`);
}



async function patchExcelWorkbookWindowViews(zip, model) {
  const file = zip.file("xl/workbook.xml");
  if (!file) return;
  const sheetCount = model?.sheets?.length || 1;
  const workbookView = normalizeWorkbookViewModel(model?.workbookView, sheetCount);
  const activeTab = clampWorkbookSheetIndex(
    state.model === model ? state.activeSheetIndex : model?.activeSheetIndex ?? workbookView.activeSheetIndex,
    sheetCount,
  );
  const entries = excelWorkbookViewEntries(model, workbookView, activeTab).map(excelWorkbookViewXml).join("");
  if (!entries) return;
  const workbookXml = await file.async("text");
  zip.file("xl/workbook.xml", upsertWorkbookBookViewsXml(workbookXml, `<bookViews>${entries}</bookViews>`));
}



function excelWorkbookViewXml(view = {}) {
  const visibility = ["hidden", "veryHidden"].includes(String(view.visibility || "visible"))
    ? String(view.visibility || "visible")
    : "visible";
  const attributes = [
    `xWindow="${Math.round(Number(view.x) || 0)}"`,
    `yWindow="${Math.round(Number(view.y) || 0)}"`,
    `windowWidth="${Math.max(1, Math.round(Number(view.width) || 24000))}"`,
    `windowHeight="${Math.max(1, Math.round(Number(view.height) || 14000))}"`,
    `firstSheet="${Math.max(0, Math.round(Number(view.firstSheet) || 0))}"`,
    `activeTab="${Math.max(0, Math.round(Number(view.activeTab) || 0))}"`,
    `visibility="${escapeAttr(visibility)}"`,
    view.maximized ? 'maximized="1"' : "",
    view.minimized ? 'minimized="1"' : "",
    `showHorizontalScroll="${view.showHorizontalScroll === false ? "0" : "1"}"`,
    `showVerticalScroll="${view.showVerticalScroll === false ? "0" : "1"}"`,
    `showSheetTabs="${view.showSheetTabs === false ? "0" : "1"}"`,
  ].filter(Boolean);
  return `<workbookView ${attributes.join(" ")}/>`;
}



function upsertWorkbookBookViewsXml(workbookXml, block) {
  const withoutExisting = String(workbookXml || "").replace(/<bookViews\b[\s\S]*?<\/bookViews>/, "");
  if (/<sheets\b/.test(withoutExisting)) {
    return withoutExisting.replace(/<sheets\b/, `${block}<sheets`);
  }
  return withoutExisting.replace("</workbook>", `${block}</workbook>`);
}



async function patchExcelWorkbookCustomViews(zip, model) {
  const file = zip.file("xl/workbook.xml");
  if (!file) return;
  const entries = excelCustomWorkbookViewXmlEntries(model);
  const workbookXml = await file.async("text");
  const block = entries.length ? `<customWorkbookViews>${entries.join("")}</customWorkbookViews>` : "";
  zip.file("xl/workbook.xml", upsertWorkbookCustomViewsXml(workbookXml, block));
}



function excelCustomWorkbookViewXmlEntries(model) {
  const sheetCount = model?.sheets?.length || 1;
  const workbookView = normalizeWorkbookViewModel(model?.workbookView, sheetCount);
  return (model?.customViews || [])
    .map(normalizeWorkbookCustomView)
    .filter(Boolean)
    .map((view) => {
      const activeSheetIndex = clampWorkbookSheetIndex(view.activeSheetIndex, sheetCount);
      const windowView = view.mainWindowView || workbookView;
      const attributes = [
        `name="${escapeAttr(view.name)}"`,
        `guid="${escapeAttr(view.guid)}"`,
        `autoUpdate="${view.autoUpdate === true ? "1" : "0"}"`,
        `mergeInterval="${Math.max(0, Math.round(Number(view.mergeInterval) || 0))}"`,
        `changesSavedWin="${view.changesSavedWin === true ? "1" : "0"}"`,
        `onlySync="${view.onlySync === true ? "1" : "0"}"`,
        `personalView="${view.personalView === true ? "1" : "0"}"`,
        `includePrintSettings="${view.includePrintSettings !== false ? "1" : "0"}"`,
        `includeHiddenRowCol="${view.includeHiddenRowsColumnsAndFilters !== false ? "1" : "0"}"`,
        `maximized="${windowView?.maximized ? "1" : "0"}"`,
        `minimized="${view.mainWindowHidden ? "1" : "0"}"`,
        `showHorizontalScroll="${windowView?.showHorizontalScroll !== false ? "1" : "0"}"`,
        `showVerticalScroll="${windowView?.showVerticalScroll !== false ? "1" : "0"}"`,
        `showSheetTabs="${windowView?.showSheetTabs !== false ? "1" : "0"}"`,
        `showFormulaBar="${view.showFormulaBar !== false ? "1" : "0"}"`,
        `firstSheet="${clampWorkbookSheetIndex(windowView?.firstSheet ?? workbookView.firstSheet, sheetCount)}"`,
        `xWindow="${Math.round(Number(windowView?.x ?? workbookView.x) || 0)}"`,
        `yWindow="${Math.round(Number(windowView?.y ?? workbookView.y) || 0)}"`,
        `windowWidth="${Math.round(Number(windowView?.width ?? workbookView.width) || 18000)}"`,
        `windowHeight="${Math.round(Number(windowView?.height ?? workbookView.height) || 12000)}"`,
        `activeSheetId="${activeSheetIndex}"`,
      ];
      return `<customWorkbookView ${attributes.join(" ")}/>`;
    });
}



function upsertWorkbookCustomViewsXml(workbookXml, block) {
  const withoutExisting = String(workbookXml || "").replace(/<customWorkbookViews\b[\s\S]*?<\/customWorkbookViews>/, "");
  if (/<calcPr\b[^>]*\/>/.test(withoutExisting)) {
    return withoutExisting.replace(/<calcPr\b[^>]*\/>/, (match) => `${match}${block}`);
  }
  if (/<calcPr\b[\s\S]*?<\/calcPr>/.test(withoutExisting)) {
    return withoutExisting.replace(/<calcPr\b[\s\S]*?<\/calcPr>/, (match) => `${match}${block}`);
  }
  const beforeLateWorkbookChildren = /(<pivotCaches\b|<smartTagPr\b|<smartTagTypes\b|<webPublishing\b|<fileRecoveryPr\b|<webPublishObjects\b|<extLst\b|<\/workbook>)/;
  if (beforeLateWorkbookChildren.test(withoutExisting)) {
    return withoutExisting.replace(beforeLateWorkbookChildren, `${block}$1`);
  }
  return withoutExisting;
}



async function patchExcelWorksheetCustomViews(zip, model) {
  const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
  const workbookRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("text");
  if (!workbookXml || !workbookRelsXml) return;
  const paths = workbookSheetPaths(workbookXml, workbookRelsXml);
  await Promise.all(paths.map(async (sheetInfo, sheetIndex) => {
    const sheet = model?.sheets?.[sheetIndex];
    if (!sheet) return;
    const entries = excelCustomSheetViewXmlEntries(model, sheet, sheetIndex);
    const file = zip.file(sheetInfo.path);
    if (!file) return;
    const xml = await file.async("text");
    const block = entries.length ? `<customSheetViews>${entries.join("")}</customSheetViews>` : "";
    zip.file(sheetInfo.path, upsertWorksheetCustomSheetViewsXml(xml, block));
  }));
}



function excelCustomSheetViewXmlEntries(model, sheet, sheetIndex) {
  return (model?.customViews || [])
    .map(normalizeWorkbookCustomView)
    .filter(Boolean)
    .map((view) => {
      const savedSheet = customViewSheetSnapshotForSheet(view, sheet, sheetIndex);
      return savedSheet ? excelCustomSheetViewXml(view, savedSheet, sheet, sheetIndex) : "";
    })
    .filter(Boolean);
}



function customViewSheetSnapshotForSheet(view, sheet, sheetIndex) {
  const saved = (view.sheets || []).find((candidate) => candidate.name === sheet?.name) || view.sheets?.[sheetIndex] || null;
  return {
    name: sheet?.name || saved?.name || `Sheet${sheetIndex + 1}`,
    viewMode: saved?.viewMode || sheet?.viewMode || "normal",
    zoomScale: saved?.zoomScale || sheet?.zoomScale || 100,
    showGridLines: saved?.showGridLines !== false,
    showHeadings: saved?.showHeadings !== false,
    showRuler: normalizedSheetShowRuler(saved?.showRuler, saved?.viewMode || sheet?.viewMode),
    showOutlineSymbols: saved?.showOutlineSymbols !== false,
    showFormulas: saved?.showFormulas === true,
    showZeros: saved?.showZeros !== false,
    showWhiteSpace: saved?.showWhiteSpace !== false,
    rightToLeft: saved?.rightToLeft === true,
    freezePanes: normalizePaneState(saved?.freezePanes || sheet?.freezePanes),
    splitPane: normalizePaneState(saved?.splitPane || sheet?.splitPane),
    activeScroll: normalizeCustomViewScroll(saved?.activeScroll || (view.activeSheetIndex === sheetIndex ? view.activeScroll : null)),
    topLeftCell: normalizeCustomTopLeftCell(saved?.topLeftCell)
      || (view.activeSheetIndex === sheetIndex ? normalizeCustomTopLeftCell(view.topLeftCell) : null)
      || topLeftCellFromScrollRecord(sheet, saved?.activeScroll || (view.activeSheetIndex === sheetIndex ? view.activeScroll : null)),
    selectedCell: saved?.selectedCell || (view.activeSheetIndex === sheetIndex ? view.selectedCell : null),
    selectionRange: saved?.selectionRange || (view.activeSheetIndex === sheetIndex ? view.selectionRange : null),
    selectionRanges: saved?.selectionRanges?.length
      ? saved.selectionRanges
      : view.activeSheetIndex === sheetIndex
        ? view.selectionRanges || []
        : [],
    pageLayout: saved?.pageLayout || sheet?.pageLayout || null,
    headerFooter: saved?.headerFooter || sheet?.headerFooter || null,
  };
}



function excelCustomSheetViewXml(view, sheetSnapshot, currentSheet, sheetIndex) {
  const topLeftCellAttr = excelCustomSheetViewTopLeftCellAttribute(sheetSnapshot, currentSheet);
  const sheetSnapshotLayout = normalizedPageLayoutSettings(sheetSnapshot?.pageLayout || currentSheet?.pageLayout || {}, currentSheet);
  const attributes = [
    `guid="${escapeAttr(view.guid)}"`,
    `scale="${Math.round(clampZoom(sheetSnapshot.zoomScale || 100))}"`,
    'colorId="64"',
    `showPageBreaks="${sheetSnapshot.viewMode === "page-break" ? "1" : "0"}"`,
    `showFormulas="${sheetSnapshot.showFormulas === true ? "1" : "0"}"`,
    `showGridLines="${sheetSnapshot.showGridLines !== false ? "1" : "0"}"`,
    `showRowCol="${sheetSnapshot.showHeadings !== false ? "1" : "0"}"`,
    `showRuler="${normalizedSheetShowRuler(sheetSnapshot.showRuler, sheetSnapshot.viewMode) ? "1" : "0"}"`,
    `outlineSymbols="${sheetSnapshot.showOutlineSymbols !== false ? "1" : "0"}"`,
    `zeroValues="${sheetSnapshot.showZeros !== false ? "1" : "0"}"`,
    `showWhiteSpace="${sheetSnapshot.showWhiteSpace !== false ? "1" : "0"}"`,
    `rightToLeft="${sheetSnapshot.rightToLeft === true ? "1" : "0"}"`,
    `fitToPage="${sheetSnapshotLayout.fitToPage ? "1" : "0"}"`,
    topLeftCellAttr,
    excelCustomSheetViewModeAttribute(sheetSnapshot.viewMode),
  ].filter(Boolean);
  const paneXml = excelCustomSheetViewPaneXml(sheetSnapshot, currentSheet);
  const selectionXml = excelCustomSheetViewSelectionXml(sheetSnapshot, sheetIndex);
  const printXml = excelCustomSheetViewPrintSettingsXml(view, sheetSnapshot, currentSheet);
  return `<customSheetView ${attributes.join(" ")}>${paneXml}${selectionXml}${printXml}</customSheetView>`;
}



function excelCustomSheetViewTopLeftCellAttribute(sheetSnapshot, currentSheet) {
  const cell = normalizeCustomTopLeftCell(sheetSnapshot?.topLeftCell)
    || topLeftCellFromScrollRecord(currentSheet, sheetSnapshot?.activeScroll);
  if (!cell || (cell.row <= 1 && cell.col <= 1)) return "";
  return `topLeftCell="${excelCellAddress(cell.row, cell.col)}"`;
}



function excelCustomSheetViewModeAttribute(viewMode) {
  if (viewMode === "page-layout") return 'view="pageLayout"';
  if (viewMode === "page-break") return 'view="pageBreakPreview"';
  return "";
}



function excelCustomSheetViewPaneXml(sheetSnapshot, currentSheet) {
  const pane = normalizePaneState(sheetSnapshot.freezePanes);
  if (pane && (pane.rows > 0 || pane.columns > 0)) {
    const rows = Math.max(0, Number(pane.rows) || 0);
    const columns = Math.max(0, Number(pane.columns) || 0);
    const activePane = pane.activePane || excelSplitActivePane(rows, columns) || "topLeft";
    return `<pane xSplit="${columns}" ySplit="${rows}" topLeftCell="${excelCellAddress(rows + 1, columns + 1)}" activePane="${activePane}" state="frozen"/>`;
  }
  const split = normalizePaneState(sheetSnapshot.splitPane);
  if (split && (split.rows > 0 || split.columns > 0)) {
    const rows = Math.max(0, Number(split.rows) || 0);
    const columns = Math.max(0, Number(split.columns) || 0);
    const ySplit = rows > 0 ? excelSplitUnitsForPaneIndex(currentSheet, "row", rows) : undefined;
    const xSplit = columns > 0 ? excelSplitUnitsForPaneIndex(currentSheet, "col", columns) : undefined;
    const attrs = [
      xSplit ? `xSplit="${xSplit}"` : "",
      ySplit ? `ySplit="${ySplit}"` : "",
      `topLeftCell="${excelCellAddress(rows + 1, columns + 1)}"`,
      `activePane="${split.activePane || excelSplitActivePane(rows, columns) || "topLeft"}"`,
      'state="split"',
    ].filter(Boolean);
    return `<pane ${attrs.join(" ")}/>`;
  }
  return "";
}



function excelCustomSheetViewSelectionXml(sheetSnapshot, sheetIndex) {
  const ranges = (Array.isArray(sheetSnapshot.selectionRanges) && sheetSnapshot.selectionRanges.length
    ? sheetSnapshot.selectionRanges
    : sheetSnapshot.selectionRange ? [sheetSnapshot.selectionRange] : [])
    .map((range) => normalizeCustomViewRange(range, sheetIndex))
    .filter((range) => range && range.sheetIndex === sheetIndex);
  const selected = normalizeCustomViewPoint(sheetSnapshot.selectedCell, sheetIndex)
    || (ranges[0] ? { sheetIndex, row: ranges[0].top, col: ranges[0].left } : { sheetIndex, row: 1, col: 1 });
  const sqref = ranges.length
    ? ranges.map((range) => rangeToLabel(range)).join(" ")
    : excelCellAddress(selected.row, selected.col);
  const pane = normalizePaneState(sheetSnapshot.freezePanes) || normalizePaneState(sheetSnapshot.splitPane);
  const paneAttr = pane?.activePane && pane.activePane !== "topLeft" ? ` pane="${escapeAttr(pane.activePane)}"` : "";
  return `<selection${paneAttr} activeCell="${excelCellAddress(selected.row, selected.col)}" sqref="${escapeAttr(sqref)}"/>`;
}



function excelCustomSheetViewPrintSettingsXml(view, sheetSnapshot, currentSheet) {
  if (view?.includePrintSettings === false) return "";
  const layout = normalizedPageLayoutSettings(sheetSnapshot?.pageLayout || currentSheet?.pageLayout || {}, currentSheet);
  const headerFooter = normalizedHeaderFooterSettings(sheetSnapshot?.headerFooter || currentSheet?.headerFooter, layout);
  return [
    excelCustomSheetViewPrintOptionsXml(layout),
    excelCustomSheetViewPageMarginsXml(layout),
    excelCustomSheetViewPageSetupXml(layout),
    excelCustomSheetViewHeaderFooterXml(headerFooter),
  ].filter(Boolean).join("");
}



function excelCustomSheetViewPrintOptionsXml(layout = {}) {
  const attrs = [
    layout.printHeadings ? 'headings="1"' : "",
    layout.printGridLines ? 'gridLines="1"' : "",
    layout.centerHorizontal ? 'horizontalCentered="1"' : "",
    layout.centerVertical ? 'verticalCentered="1"' : "",
  ].filter(Boolean);
  return attrs.length ? `<printOptions ${attrs.join(" ")}/>` : "";
}



function excelCustomSheetViewPageMarginsXml(layout = {}) {
  const margins = layout.margins || {};
  const numeric = (value, fallback) => {
    const number = Number(value);
    return Number.isFinite(number) ? number : fallback;
  };
  const attrs = [
    `left="${numeric(margins.left, 0.7)}"`,
    `right="${numeric(margins.right, 0.7)}"`,
    `top="${numeric(margins.top, 0.75)}"`,
    `bottom="${numeric(margins.bottom, 0.75)}"`,
    `header="${numeric(margins.header, 0.3)}"`,
    `footer="${numeric(margins.footer, 0.3)}"`,
  ];
  return `<pageMargins ${attrs.join(" ")}/>`;
}



function excelCustomSheetViewPageSetupXml(layout = {}) {
  const fitToPage = layout.fitToPage === true;
  const positiveIntegerOrZero = (value) => Math.max(0, Math.round(Number(value) || 0));
  const attrs = [
    `paperSize="${excelPaperSizeCode(layout.paperSize)}"`,
    `orientation="${escapeAttr(layout.orientation || "portrait")}"`,
    `pageOrder="${escapeAttr(layout.pageOrder || "downThenOver")}"`,
    layout.blackAndWhite ? 'blackAndWhite="1"' : "",
    layout.draftQuality ? 'draft="1"' : "",
    `cellComments="${escapeAttr(excelPageCommentsMode(layout.comments))}"`,
    `errors="${escapeAttr(excelPageErrorsMode(layout.cellErrors))}"`,
    fitToPage
      ? `fitToWidth="${layout.fitToWidth === "auto" ? 0 : positiveIntegerOrZero(layout.fitToWidth)}"`
      : `scale="${Math.max(10, Math.min(400, Math.round(Number(layout.scale) || 100)))}"`,
    fitToPage ? `fitToHeight="${layout.fitToHeight === "auto" ? 0 : positiveIntegerOrZero(layout.fitToHeight)}"` : "",
    Number.isFinite(Number(layout.firstPageNumber)) ? `firstPageNumber="${Math.max(1, Math.round(Number(layout.firstPageNumber)))}" useFirstPageNumber="1"` : "",
    'usePrinterDefaults="1"',
    'horizontalDpi="4294967295"',
    'verticalDpi="4294967295"',
  ].filter(Boolean);
  return `<pageSetup ${attrs.join(" ")}/>`;
}



function excelCustomSheetViewHeaderFooterXml(settings = {}) {
  const attrs = [
    settings.differentOddEven ? 'differentOddEven="1"' : "",
    settings.differentFirst ? 'differentFirst="1"' : "",
    settings.scaleWithDoc === false ? 'scaleWithDoc="0"' : "",
    settings.alignWithMargins === false ? 'alignWithMargins="0"' : "",
  ].filter(Boolean);
  const children = [
    excelCustomSheetViewHeaderFooterTextXml("oddHeader", excelHeaderFooterText(settings.oddHeader)),
    excelCustomSheetViewHeaderFooterTextXml("oddFooter", excelHeaderFooterText(settings.oddFooter)),
    excelCustomSheetViewHeaderFooterTextXml("evenHeader", excelHeaderFooterText(settings.evenHeader)),
    excelCustomSheetViewHeaderFooterTextXml("evenFooter", excelHeaderFooterText(settings.evenFooter)),
    excelCustomSheetViewHeaderFooterTextXml("firstHeader", excelHeaderFooterText(settings.firstHeader)),
    excelCustomSheetViewHeaderFooterTextXml("firstFooter", excelHeaderFooterText(settings.firstFooter)),
  ].join("");
  if (!attrs.length && !children) return "";
  return `<headerFooter${attrs.length ? ` ${attrs.join(" ")}` : ""}>${children}</headerFooter>`;
}



function excelCustomSheetViewHeaderFooterTextXml(tag, text) {
  const value = String(text || "");
  return value ? `<${tag}>${escapeXml(value)}</${tag}>` : "";
}



function upsertWorksheetCustomSheetViewsXml(worksheetXml, block) {
  const withoutExisting = String(worksheetXml || "").replace(/<customSheetViews\b[\s\S]*?<\/customSheetViews>/, "");
  const beforeLaterWorksheetChildren = /(<mergeCells\b|<phoneticPr\b|<conditionalFormatting\b|<dataValidations\b|<hyperlinks\b|<printOptions\b|<pageMargins\b|<pageSetup\b|<headerFooter\b|<rowBreaks\b|<colBreaks\b|<customProperties\b|<cellWatches\b|<ignoredErrors\b|<smartTags\b|<drawing\b|<legacyDrawing\b|<legacyDrawingHF\b|<picture\b|<oleObjects\b|<controls\b|<webPublishItems\b|<tableParts\b|<extLst\b|<\/worksheet>)/;
  if (beforeLaterWorksheetChildren.test(withoutExisting)) {
    return withoutExisting.replace(beforeLaterWorksheetChildren, `${block}$1`);
  }
  return withoutExisting;
}



async function patchExcelWorksheetViewDetails(zip, model) {
  const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
  const workbookRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("text");
  if (!workbookXml || !workbookRelsXml) return;
  const paths = workbookSheetPaths(workbookXml, workbookRelsXml);
  for (const [index, sheetInfo] of paths.entries()) {
    const sheet = model?.sheets?.[index];
    if (!sheet) continue;
    const file = zip.file(sheetInfo.path);
    if (!file) continue;
    let xml = await file.async("text");
    let nextXml = xml.replace(/<sheetView\b[^>]*>/, (tag) => {
      let next = tag;
      next = patchXmlTagAttribute(next, "showGridLines", sheet.showGridLines === false ? "0" : null);
      next = patchXmlTagAttribute(next, "showRowColHeaders", sheet.showHeadings === false ? "0" : null);
      next = patchXmlTagAttribute(next, "showRuler", normalizedSheetShowRuler(sheet.showRuler, sheet.viewMode) ? "1" : null);
      next = patchXmlTagAttribute(next, "showFormulas", sheet.showFormulas === true ? "1" : null);
      next = patchXmlTagAttribute(next, "showPageBreaks", sheet.viewMode === "page-break" ? "1" : null);
      next = patchXmlTagAttribute(next, "showOutlineSymbols", sheet.showOutlineSymbols === false ? "0" : null);
      next = patchXmlTagAttribute(next, "showZeros", sheet.showZeros === false ? "0" : null);
      next = patchXmlTagAttribute(next, "showWhiteSpace", sheet.showWhiteSpace === false ? "0" : null);
      next = patchXmlTagAttribute(next, "rightToLeft", sheet.rightToLeft === true ? "1" : null);
      next = patchXmlTagAttribute(next, "tabSelected", excelWorksheetTabSelected(sheet, index) ? "1" : null);
      return next;
    });
    nextXml = patchWorksheetSheetViewSelectionDetails(nextXml, sheet, index);
    if (nextXml !== xml) zip.file(sheetInfo.path, nextXml);
  }
}



function patchWorksheetSheetViewSelectionDetails(worksheetXml, sheet, sheetIndex) {
  const xml = String(worksheetXml || "");
  const selectionXml = excelWorksheetViewSelectionXml(sheet, sheetIndex);
  const selectionState = excelWorksheetViewSelectionState(sheet, sheetIndex);
  const topLeftCell = selectionState.topLeftCell;
  const topLeftCellValue = topLeftCell && (topLeftCell.row > 1 || topLeftCell.col > 1)
    ? excelCellAddress(topLeftCell.row, topLeftCell.col)
    : null;
  const patchSheetViewTag = (tag) => patchXmlTagAttribute(tag, "topLeftCell", topLeftCellValue);
  const patchSheetViewXml = (sheetViewXml) => {
    const selfClosing = /\/>\s*$/.test(sheetViewXml);
    if (selfClosing) {
      const openTag = patchSheetViewTag(sheetViewXml).replace(/\s*\/>\s*$/, ">");
      return `${openTag}${selectionXml}</sheetView>`;
    }
    let next = sheetViewXml.replace(/<sheetView\b[^>]*>/, patchSheetViewTag);
    next = next.replace(/<selection\b[^>]*(?:\/>|>[\s\S]*?<\/selection>)/g, "");
    if (/<pane\b[^>]*(?:\/>|>[\s\S]*?<\/pane>)/.test(next)) {
      return next.replace(/(<pane\b[^>]*(?:\/>|>[\s\S]*?<\/pane>))/, `$1${selectionXml}`);
    }
    return next.replace(/<\/sheetView>$/, `${selectionXml}</sheetView>`);
  };
  if (/<sheetView\b[^>]*>[\s\S]*?<\/sheetView>|<sheetView\b[^>]*\/>/.test(xml)) {
    return xml.replace(/<sheetView\b[^>]*>[\s\S]*?<\/sheetView>|<sheetView\b[^>]*\/>/, patchSheetViewXml);
  }
  const sheetViewXml = `<sheetView workbookViewId="0">${selectionXml}</sheetView>`;
  if (/<sheetViews\b[^>]*>[\s\S]*?<\/sheetViews>/.test(xml)) {
    return xml.replace(/<sheetViews\b[^>]*>[\s\S]*?<\/sheetViews>/, `<sheetViews>${sheetViewXml}</sheetViews>`);
  }
  return xml.replace(/<worksheet\b[^>]*>/, (tag) => `${tag}<sheetViews>${sheetViewXml}</sheetViews>`);
}



function excelWorksheetViewSelectionXml(sheet, sheetIndex) {
  const selection = excelWorksheetViewSelectionState(sheet, sheetIndex);
  const activeCell = excelCellAddress(selection.selected.row, selection.selected.col);
  const sqref = selection.ranges.length
    ? selection.ranges.map((range) => rangeToLabel(range)).join(" ")
    : activeCell;
  const paneAttr = selection.pane?.activePane && selection.pane.activePane !== "topLeft"
    ? ` pane="${escapeAttr(selection.pane.activePane)}"`
    : "";
  return `<selection${paneAttr} activeCell="${activeCell}" sqref="${escapeAttr(sqref)}"/>`;
}



function excelWorksheetViewSelectionState(sheet, sheetIndex) {
  const isLiveSheet = state.model?.sheets?.[sheetIndex] === sheet;
  const isActiveLiveSheet = isLiveSheet && state.activeSheetIndex === sheetIndex;
  const liveSelected = isActiveLiveSheet && state.selected?.sheetIndex === sheetIndex ? { ...state.selected } : null;
  const storedSelected = normalizeCustomViewPoint(sheet?.selectedCell, sheetIndex);
  const pane = normalizePaneState(sheet?.freezePanes) || normalizePaneState(sheet?.splitPane);
  const ranges = excelWorksheetViewSelectionRanges(sheet, sheetIndex, isActiveLiveSheet);
  const selected = normalizeCustomViewPoint(liveSelected || storedSelected, sheetIndex)
    || (ranges[0] ? { sheetIndex, row: ranges[0].top, col: ranges[0].left } : null)
    || { sheetIndex, row: Math.max(1, pane?.row || 1), col: Math.max(1, pane?.col || 1) };
  const host = isActiveLiveSheet ? $sheetHost[0] : null;
  const liveTopLeftCell = host
    ? topLeftCellFromScrollRecord(sheet, { top: host.scrollTop || 0, left: host.scrollLeft || 0, scrollTop: host.scrollTop || 0, scrollLeft: host.scrollLeft || 0 })
    : null;
  const topLeftCell = liveTopLeftCell || normalizeCustomTopLeftCell(sheet?.topLeftCell);
  return {
    selected,
    ranges,
    pane,
    topLeftCell,
  };
}



function excelWorksheetViewSelectionRanges(sheet, sheetIndex, isActiveLiveSheet = false) {
  const source = isActiveLiveSheet
    ? activeSelectionRanges()
    : (Array.isArray(sheet?.selectionRanges) && sheet.selectionRanges.length ? sheet.selectionRanges : [sheet?.selectionRange]);
  const ranges = (source || [])
    .map((range) => normalizeCustomViewRange(range, sheetIndex))
    .filter((range) => range && range.sheetIndex === sheetIndex)
    .map((range) => materializeSelectionRangeForSheet(sheet, range, sheetIndex))
    .filter(Boolean);
  if (ranges.length) return ranges;
  const selected = normalizeCustomViewPoint(sheet?.selectedCell, sheetIndex);
  return selected ? [{ sheetIndex, top: selected.row, left: selected.col, bottom: selected.row, right: selected.col }] : [];
}



function patchXmlTagAttribute(tag, name, value) {
  let next = String(tag || "");
  next = next.replace(new RegExp(`\\s${escapeRegExp(name)}="[^"]*"`, "g"), "");
  if (value == null) return next;
  const close = next.endsWith("/>") ? "/>" : ">";
  return next.replace(/\s*\/?>$/, ` ${name}="${escapeAttr(value)}"${close}`);
}



async function patchExcelPreservedOoxmlPackage(zip, model, type) {
  if (type?.value !== "xlsm") return;
  const preserved = normalizePreservedOoxmlPackage(model?.ooxmlPackage);
  if (!preserved?.hasVbaProject) return;
  preserved.entries.forEach((entry) => {
    zip.file(entry.path, base64ToBytes(entry.base64));
  });
  await patchExcelPreservedContentTypes(zip, preserved);
  await patchExcelPreservedWorkbookRelationships(zip, preserved);
}



async function patchExcelPreservedContentTypes(zip, preserved) {
  const file = zip.file("[Content_Types].xml");
  if (!file) return;
  let xml = await file.async("text");
  const overrides = preserved.contentTypeOverrides?.length
    ? preserved.contentTypeOverrides
    : [{ partName: "/xl/vbaProject.bin", contentType: "application/vnd.ms-office.vbaProject" }];
  overrides.forEach((override) => {
    const partName = normalizeOoxmlPartName(override.partName);
    const contentType = override.contentType || "application/vnd.ms-office.vbaProject";
    if (!partName || xml.includes(`PartName="${escapeAttr(partName)}"`)) return;
    xml = xml.replace("</Types>", `<Override PartName="${escapeAttr(partName)}" ContentType="${escapeAttr(contentType)}"/></Types>`);
  });
  zip.file("[Content_Types].xml", xml);
}



async function patchExcelPreservedWorkbookRelationships(zip, preserved) {
  const path = "xl/_rels/workbook.xml.rels";
  const file = zip.file(path);
  let xml = file
    ? await file.async("text")
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  if (/relationships\/vbaProject/i.test(xml) || /Target="vbaProject\.bin"/i.test(xml)) {
    zip.file(path, xml);
    return;
  }
  const original = preserved.workbookRelationships?.find((relationship) => /vbaProject/i.test(relationship.type) || /vbaProject/i.test(relationship.target));
  const relationship = {
    id: nextWorkbookRelationshipId(xml),
    type: original?.type || "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
    target: original?.target || "vbaProject.bin",
    targetMode: original?.targetMode || "",
  };
  const targetMode = relationship.targetMode ? ` TargetMode="${escapeAttr(relationship.targetMode)}"` : "";
  const node = `<Relationship Id="${escapeAttr(relationship.id)}" Type="${escapeAttr(relationship.type)}" Target="${escapeAttr(relationship.target)}"${targetMode}/>`;
  xml = xml.replace("</Relationships>", `${node}</Relationships>`);
  zip.file(path, xml);
}



function normalizeOoxmlPartName(partName) {
  const normalized = String(partName || "").trim();
  if (!normalized) return "";
  return normalized.startsWith("/") ? normalized : `/${normalized}`;
}



function nextWorkbookRelationshipId(xml) {
  const ids = [...String(xml || "").matchAll(/\bId="rId(\d+)"/g)].map((match) => Number(match[1])).filter(Number.isFinite);
  return `rId${Math.max(0, ...ids) + 1}`;
}



async function patchExcelWorksheetPageBreaks(zip, model) {
  await Promise.all((model?.sheets || []).map(async (sheet, index) => {
    const layout = normalizedPageLayoutSettings(sheet.pageLayout || {}, sheet);
    const path = `xl/worksheets/sheet${index + 1}.xml`;
    const file = zip.file(path);
    if (!file) return;
    const xml = await file.async("text");
    const rowBreaks = excelWorksheetPageBreaksXml("row", layout.pageBreaks?.rows || []);
    const colBreaks = excelWorksheetPageBreaksXml("col", layout.pageBreaks?.cols || []);
    const nextXml = replaceWorksheetPageBreaksXml(xml, rowBreaks, colBreaks);
    if (nextXml !== xml) zip.file(path, nextXml);
  }));
}



async function patchExcelWorksheetOutlineDetails(zip, model) {
  const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
  const workbookRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("text");
  if (!workbookXml || !workbookRelsXml) return;
  const paths = workbookSheetPaths(workbookXml, workbookRelsXml);
  await Promise.all(paths.map(async (sheetInfo, index) => {
    const sheet = model?.sheets?.[index];
    const file = zip.file(sheetInfo.path);
    if (!sheet || !file) return;
    const xml = await file.async("text");
    const nextXml = patchWorksheetOutlineXml(xml, sheet);
    if (nextXml !== xml) zip.file(sheetInfo.path, nextXml);
  }));
}



function patchWorksheetOutlineXml(xml, sheet) {
  let next = patchWorksheetOutlineCollapsedXml(xml, sheet);
  next = patchWorksheetOutlinePropertiesXml(next, sheet);
  return next;
}



function patchWorksheetOutlineCollapsedXml(xml, sheet) {
  const rowCollapsed = outlineCollapsedIndexSet(sheet?.rows);
  const columnCollapsed = outlineCollapsedIndexSet(sheet?.columns);
  let next = String(xml || "").replace(/<row\b[^>]*>/g, (tag) => {
    const row = Math.round(Number(xmlAttributeValue(tag, "r")) || 0);
    if (row < 1) return tag;
    return patchXmlTagAttribute(tag, "collapsed", rowCollapsed.has(row) ? "1" : null);
  });
  next = patchWorksheetColumnCollapsedXml(next, columnCollapsed);
  return next;
}



function patchWorksheetOutlinePropertiesXml(xml, sheet) {
  const settings = normalizeOutlineSettings(sheet?.outlineSettings);
  const hasNonDefault = settings.summaryBelow === false || settings.summaryRight === false || settings.applyStyles === true;
  const hasOutlinePr = /<outlinePr\b/.test(String(xml || ""));
  if (!hasOutlinePr && !hasNonDefault) return xml;
  const patchTag = (tag) => {
    let next = tag;
    next = patchXmlTagAttribute(next, "summaryBelow", settings.summaryBelow === false ? "0" : null);
    next = patchXmlTagAttribute(next, "summaryRight", settings.summaryRight === false ? "0" : null);
    next = patchXmlTagAttribute(next, "applyStyles", settings.applyStyles === true ? "1" : null);
    return next;
  };
  if (hasOutlinePr) {
    return String(xml || "")
      .replace(/<outlinePr\b[^>]*\/>/, patchTag)
      .replace(/<outlinePr\b[^>]*>(?:\s*)<\/outlinePr>/, (tag) => patchTag(tag.replace(/>(?:\s*)<\/outlinePr>$/, "/>")));
  }
  const outlinePr = patchTag("<outlinePr/>");
  if (/<sheetPr\b[^>]*\/>/.test(String(xml || ""))) {
    return String(xml || "").replace(/<sheetPr\b([^>]*)\/>/, `<sheetPr$1>${outlinePr}</sheetPr>`);
  }
  if (/<sheetPr\b[^>]*>/.test(String(xml || ""))) {
    return String(xml || "").replace(/<sheetPr\b[^>]*>/, (tag) => `${tag}${outlinePr}`);
  }
  return String(xml || "").replace(/<worksheet\b[^>]*>/, (tag) => `${tag}<sheetPr>${outlinePr}</sheetPr>`);
}



function outlineCollapsedIndexSet(collection = []) {
  const result = new Set();
  (Array.isArray(collection) ? collection : []).forEach((item, index) => {
    if (item?.outlineCollapsed === true) result.add(index + 1);
  });
  return result;
}



function patchWorksheetColumnCollapsedXml(xml, collapsedColumns) {
  return String(xml || "").replace(/<col\b[^>]*\/>/g, (tag) => {
    const min = Math.max(1, Math.round(Number(xmlAttributeValue(tag, "min")) || 1));
    const max = Math.max(min, Math.round(Number(xmlAttributeValue(tag, "max")) || min));
    const segments = worksheetColumnCollapsedSegments(min, max, collapsedColumns);
    if (segments.length <= 1) {
      return patchXmlTagAttribute(tag, "collapsed", segments[0]?.collapsed ? "1" : null);
    }
    return segments.map((segment) => {
      let next = patchXmlTagAttribute(tag, "min", String(segment.start));
      next = patchXmlTagAttribute(next, "max", String(segment.end));
      next = patchXmlTagAttribute(next, "collapsed", segment.collapsed ? "1" : null);
      return next;
    }).join("");
  });
}



function worksheetColumnCollapsedSegments(min, max, collapsedColumns) {
  const segments = [];
  let start = min;
  let collapsed = collapsedColumns.has(min);
  for (let column = min + 1; column <= max; column += 1) {
    const nextCollapsed = collapsedColumns.has(column);
    if (nextCollapsed === collapsed) continue;
    segments.push({ start, end: column - 1, collapsed });
    start = column;
    collapsed = nextCollapsed;
  }
  segments.push({ start, end: max, collapsed });
  return segments;
}



function excelWorksheetPageBreaksXml(axis, breaks = []) {
  const isColumn = axis === "col";
  const tag = isColumn ? "colBreaks" : "rowBreaks";
  const max = isColumn ? 1048575 : 16383;
  const normalized = [...new Set((Array.isArray(breaks) ? breaks : [])
    .map((value) => Math.max(2, Math.round(Number(value) || 0)))
    .filter((value) => value >= 2))]
    .sort((a, b) => a - b);
  if (!normalized.length) return "";
  const entries = normalized
    .map((value) => `<brk id="${Math.max(1, value - 1)}" max="${max}" man="1"/>`)
    .join("");
  return `<${tag} count="${normalized.length}" manualBreakCount="${normalized.length}">${entries}</${tag}>`;
}



function replaceWorksheetPageBreaksXml(xml, rowBreaksXml, colBreaksXml) {
  const block = `${rowBreaksXml || ""}${colBreaksXml || ""}`;
  const withoutExisting = String(xml || "")
    .replace(/<rowBreaks\b[\s\S]*?<\/rowBreaks>|<rowBreaks\b[^>]*\/>/g, "")
    .replace(/<colBreaks\b[\s\S]*?<\/colBreaks>|<colBreaks\b[^>]*\/>/g, "");
  if (!block) return withoutExisting;
  const anchor = /(<customProperties\b|<cellWatches\b|<ignoredErrors\b|<smartTags\b|<drawing\b|<legacyDrawing\b|<legacyDrawingHF\b|<picture\b|<oleObjects\b|<controls\b|<webPublishItems\b|<tableParts\b|<extLst\b|<\/worksheet>)/;
  return anchor.test(withoutExisting) ? withoutExisting.replace(anchor, `${block}$1`) : withoutExisting;
}



async function patchExcelWorksheetDataValidations(zip, model) {
  await Promise.all((model?.sheets || []).map(async (sheet, index) => {
    const path = `xl/worksheets/sheet${index + 1}.xml`;
    const file = zip.file(path);
    if (!file) return;
    const xml = await file.async("text");
    const block = excelWorksheetDataValidationsXml(sheet);
    const nextXml = replaceWorksheetDataValidationsXml(xml, block);
    if (nextXml !== xml) zip.file(path, nextXml);
  }));
}



function excelWorksheetDataValidationsXml(sheet) {
  const groups = excelWorksheetDataValidationGroups(sheet);
  if (!groups.length) return "";
  const items = groups
    .map((group) => excelDataValidationXml(group.descriptor, excelSqrefFromCells(group.cells)))
    .filter(Boolean);
  return items.length ? `<dataValidations count="${items.length}">${items.join("")}</dataValidations>` : "";
}



function excelWorksheetDataValidationGroups(sheet) {
  const groups = new Map();
  Object.values(sheet?.cells || {}).forEach((cell) => {
    if (!cell?.validation || !Number.isFinite(Number(cell.row)) || !Number.isFinite(Number(cell.col))) return;
    const descriptor = excelDataValidationXmlDescriptor(cell.validation);
    if (!descriptor) return;
    const key = JSON.stringify(descriptor);
    if (!groups.has(key)) groups.set(key, { descriptor, cells: [] });
    groups.get(key).cells.push({ row: Number(cell.row), col: Number(cell.col) });
  });
  return [...groups.values()].sort((left, right) => {
    const a = left.cells[0] || {};
    const b = right.cells[0] || {};
    return (a.row - b.row) || (a.col - b.col);
  });
}



function excelDataValidationXmlDescriptor(validation = {}) {
  const type = validation.type || "custom";
  if (!type || type === "any" || type === "none") return null;
  const formulas = type === "list"
    ? [excelDataValidationListFormula(validation)].filter((formula) => formula !== "")
    : excelDataValidationFormulae(validation).map((formula) => excelDataValidationFormulaXmlText(formula));
  const attrs = {
    type,
    ...(type !== "list" && type !== "custom" && validation.operator ? { operator: validation.operator } : {}),
    allowBlank: validation.allowBlank === true ? "1" : "",
    showInputMessage: validation.showInputMessage === true ? "1" : "",
    showErrorMessage: validation.showErrorMessage !== false ? "1" : "0",
    errorStyle: validation.errorStyle || "stop",
    promptTitle: validation.inputTitle || "",
    prompt: validation.inputMessage || "",
    errorTitle: validation.errorTitle || "",
    error: validation.errorMessage || "",
    ...(type === "list" && validation.showDropDown === false ? { showDropDown: "1" } : {}),
    ...(validation.imeMode && validation.imeMode !== "noControl" ? { imeMode: validation.imeMode } : {}),
  };
  if ((type === "list" || type === "custom") && !formulas[0]) return null;
  return { attrs, formulas };
}



function excelDataValidationFormulaXmlText(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return String(dateToExcelSerial(value));
  return String(value ?? "");
}



function excelDataValidationXml(descriptor, sqref) {
  if (!descriptor || !sqref) return "";
  const attrs = Object.entries({ ...descriptor.attrs, sqref })
    .filter(([_name, value]) => value != null && value !== "")
    .map(([name, value]) => `${name}="${escapeAttr(value)}"`)
    .join(" ");
  const formulas = (descriptor.formulas || [])
    .slice(0, 2)
    .map((formula, index) => `<formula${index + 1}>${escapeXml(formula)}</formula${index + 1}>`)
    .join("");
  return `<dataValidation ${attrs}>${formulas}</dataValidation>`;
}



function excelSqrefFromCells(cells = []) {
  return dataValidationRectanglesFromCells(cells)
    .map((range) => {
      const start = excelCellAddress(range.top, range.left);
      const end = excelCellAddress(range.bottom, range.right);
      return start === end ? start : `${start}:${end}`;
    })
    .join(" ");
}



function dataValidationRectanglesFromCells(cells = []) {
  const rows = new Map();
  cells.forEach((cell) => {
    const row = Number(cell.row);
    const col = Number(cell.col);
    if (!Number.isInteger(row) || !Number.isInteger(col) || row < 1 || col < 1) return;
    if (!rows.has(row)) rows.set(row, new Set());
    rows.get(row).add(col);
  });
  const runs = [];
  [...rows.entries()].sort((a, b) => a[0] - b[0]).forEach(([row, cols]) => {
    const sorted = [...cols].sort((a, b) => a - b);
    let left = null;
    let right = null;
    sorted.forEach((col) => {
      if (left == null) {
        left = col;
        right = col;
      } else if (col === right + 1) {
        right = col;
      } else {
        runs.push({ row, left, right });
        left = col;
        right = col;
      }
    });
    if (left != null) runs.push({ row, left, right });
  });
  const active = new Map();
  const rectangles = [];
  runs.forEach((run) => {
    const key = `${run.left}:${run.right}`;
    const current = active.get(key);
    if (current && current.bottom === run.row - 1) {
      current.bottom = run.row;
      return;
    }
    const rect = { top: run.row, left: run.left, bottom: run.row, right: run.right };
    rectangles.push(rect);
    active.set(key, rect);
  });
  return rectangles.sort((left, right) => (left.top - right.top) || (left.left - right.left));
}



function replaceWorksheetDataValidationsXml(xml, block) {
  const withoutExisting = String(xml || "").replace(/<dataValidations\b[\s\S]*?<\/dataValidations>/, "");
  if (!block) return withoutExisting;
  const anchor = /(<hyperlinks\b|<printOptions\b|<pageMargins\b|<pageSetup\b|<headerFooter\b|<rowBreaks\b|<colBreaks\b|<customProperties\b|<cellWatches\b|<ignoredErrors\b|<smartTags\b|<drawing\b|<legacyDrawing\b|<legacyDrawingHF\b|<picture\b|<oleObjects\b|<controls\b|<webPublishItems\b|<tableParts\b|<extLst\b|<\/worksheet>)/;
  return anchor.test(withoutExisting) ? withoutExisting.replace(anchor, (match) => `${block}${match}`) : withoutExisting;
}



function rangeFromCellLabel(ref) {
  const point = decodeCell(String(ref || "").replace(/\$/g, ""));
  return point ? { top: point.row, left: point.col, bottom: point.row, right: point.col } : null;
}



async function patchExcelWorksheetInternalHyperlinks(zip, model) {
  await Promise.all((model?.sheets || []).map(async (sheet, index) => {
    const internalLinks = excelInternalCellHyperlinkMap(sheet);
    if (!internalLinks.size) return;
    const path = `xl/worksheets/sheet${index + 1}.xml`;
    const file = zip.file(path);
    if (!file) return;
    const removedRelationshipIds = new Set();
    const xml = await file.async("text");
    const nextXml = xml.replace(/<hyperlink\b[^>]*\/>/g, (tag) => {
      const ref = xmlAttributeValue(tag, "ref");
      const target = internalLinks.get(ref);
      if (!target) return tag;
      const rId = xmlAttributeValue(tag, "r:id");
      if (rId) removedRelationshipIds.add(rId);
      const tooltip = xmlAttributeValue(tag, "tooltip") || target;
      return `<hyperlink ref="${escapeAttr(ref)}" location="${escapeAttr(target)}" tooltip="${escapeAttr(tooltip)}"/>`;
    });
    if (nextXml !== xml) zip.file(path, nextXml);
    if (removedRelationshipIds.size) await removeWorksheetRelationships(zip, index, removedRelationshipIds);
  }));
}



function excelInternalCellHyperlinkMap(sheet) {
  const map = new Map();
  Object.values(sheet?.cells || {}).forEach((cell) => {
    const target = excelInternalHyperlinkTarget(cell?.hyperlink);
    if (!target || !Number.isFinite(Number(cell.row)) || !Number.isFinite(Number(cell.col))) return;
    map.set(excelCellAddress(cell.row, cell.col), target);
  });
  return map;
}



function excelInternalHyperlinkTarget(hyperlink) {
  const text = String(hyperlink || "").trim();
  if (!text.startsWith("#")) return "";
  return stripInternalHyperlinkPrefix(text);
}



async function removeWorksheetRelationships(zip, sheetIndex, relationshipIds) {
  const path = `xl/worksheets/_rels/sheet${sheetIndex + 1}.xml.rels`;
  const file = zip.file(path);
  if (!file) return;
  let xml = await file.async("text");
  relationshipIds.forEach((id) => {
    xml = xml.replace(new RegExp(`<Relationship\\b(?=[^>]*\\bId="${escapeRegExp(id)}")[^>]*/>`, "g"), "");
  });
  zip.file(path, xml);
}



function xmlAttributeValue(tag, name) {
  const match = String(tag || "").match(new RegExp(`\\b${escapeRegExp(name)}="([^"]*)"`));
  return match ? match[1] : "";
}



function addXmlAttributesToTag(tag, attrs = {}) {
  let next = String(tag || "");
  Object.entries(attrs).forEach(([name, value]) => {
    if (value == null || value === "" || xmlAttributeValue(next, name)) return;
    next = next.replace(/>$/, ` ${name}="${escapeAttr(value)}">`);
  });
  return next;
}



async function createExcelAutoFilterExportContext(zip) {
  const stylesXml = await zip.file("xl/styles.xml")?.async("text");
  const existingDxfs = parseWorkbookStyleDxfsXml(stylesXml);
  const keyToId = new Map();
  existingDxfs.forEach((dxf, index) => {
    const fillColor = normalizeColorValue(dxf.fillColor);
    const fontColor = normalizeColorValue(dxf.fontColor);
    if (fillColor) keyToId.set(excelColorFilterDxfKey("fill", fillColor), index);
    if (fontColor) keyToId.set(excelColorFilterDxfKey("font", fontColor), index);
  });
  const context = {
    stylesXml: stylesXml || "",
    keyToId,
    initialDxfCount: existingDxfs.length,
    newDxfXml: [],
    dxfIdForColorFilter(target, color) {
      const normalizedTarget = target === "font" ? "font" : "fill";
      const normalizedColor = normalizeColorValue(color);
      if (!normalizedColor) return null;
      const key = excelColorFilterDxfKey(normalizedTarget, normalizedColor);
      if (this.keyToId.has(key)) return this.keyToId.get(key);
      const id = this.initialDxfCount + this.newDxfXml.length;
      this.keyToId.set(key, id);
      this.newDxfXml.push(excelColorFilterDxfXml(normalizedTarget, normalizedColor));
      return id;
    },
  };
  return context;
}



async function finalizeExcelAutoFilterExportContext(zip, context) {
  if (!context?.newDxfXml?.length || !context.stylesXml) return;
  const totalCount = context.initialDxfCount + context.newDxfXml.length;
  const nextXml = patchExcelAutoFilterDxfStylesXml(context.stylesXml, context.newDxfXml.join(""), totalCount);
  if (nextXml && nextXml !== context.stylesXml) zip.file("xl/styles.xml", nextXml);
}



function excelColorFilterDxfKey(target, color) {
  return `${target === "font" ? "font" : "fill"}:${normalizeColorValue(color)}`;
}



function excelColorFilterDxfXml(target, color) {
  const argb = cssColorToExcelArgb(color);
  if (!argb) return "";
  if (target === "font") return `<dxf><font><color rgb="${argb}"/></font></dxf>`;
  return `<dxf><fill><patternFill patternType="solid"><fgColor rgb="${argb}"/><bgColor indexed="64"/></patternFill></fill></dxf>`;
}



function patchExcelAutoFilterDxfStylesXml(stylesXml, appendedDxfs, totalCount) {
  const xml = String(stylesXml || "");
  const countAttr = `count="${Math.max(0, Number(totalCount) || 0)}"`;
  if (/<dxfs\b[^>]*\/>/.test(xml)) {
    return xml.replace(/<dxfs\b([^>]*)\/>/, (match, attrs) => {
      const cleanAttrs = String(attrs || "").replace(/\s+count="[^"]*"/, "");
      return `<dxfs${cleanAttrs} ${countAttr}>${appendedDxfs}</dxfs>`;
    });
  }
  if (/<dxfs\b[\s\S]*?<\/dxfs>/.test(xml)) {
    return xml.replace(/<dxfs\b[^>]*>[\s\S]*?<\/dxfs>/, (block) => {
      const open = block.match(/^<dxfs\b[^>]*>/)?.[0] || "<dxfs>";
      const body = block.replace(/^<dxfs\b[^>]*>/, "").replace(/<\/dxfs>$/, "");
      const nextOpen = /\bcount=/.test(open)
        ? open.replace(/\bcount="[^"]*"/, countAttr)
        : open.replace(/>$/, ` ${countAttr}>`);
      return `${nextOpen}${body}${appendedDxfs}</dxfs>`;
    });
  }
  const dxfsXml = `<dxfs ${countAttr}>${appendedDxfs}</dxfs>`;
  if (/<tableStyles\b/.test(xml)) return xml.replace(/<tableStyles\b/, `${dxfsXml}<tableStyles`);
  return xml.replace(/<\/styleSheet>/, `${dxfsXml}</styleSheet>`);
}



async function patchExcelWorksheetAutoFilterDetails(zip, model, exportContext = null) {
  await Promise.all((model?.sheets || []).map(async (sheet, index) => {
    const autoFilterXml = excelAutoFilterXml(sheet, exportContext);
    const exportSortState = excelSheetSortStateForExport(sheet);
    const sortOnlyXml = autoFilterXml ? "" : excelSortStateXml(exportSortState, null, exportContext);
    const path = `xl/worksheets/sheet${index + 1}.xml`;
    const file = zip.file(path);
    if (!file) return;
    const xml = await file.async("text");
    let nextXml = removeWorksheetAutoFilterAndSortStateXml(xml);
    if (autoFilterXml) {
      nextXml = insertWorksheetChildXml(nextXml, autoFilterXml);
    } else if (sortOnlyXml) {
      nextXml = insertWorksheetChildXml(nextXml, sortOnlyXml);
    }
    if (nextXml !== xml) zip.file(path, nextXml);
  }));
}



function removeWorksheetAutoFilterAndSortStateXml(xml) {
  return String(xml || "")
    .replace(/<autoFilter\b[\s\S]*?<\/autoFilter>|<autoFilter\b[^>]*\/>/g, "")
    .replace(/<sortState\b[\s\S]*?<\/sortState>|<sortState\b[^>]*\/>/g, "");
}



function excelAutoFilterXml(sheet = {}, exportContext = null) {
  const sortState = sheet.sortState || tableSortStateForFilterRange(sheet, sheet.autoFilter?.range);
  return excelAutoFilterXmlForState(sheet.autoFilter, sortState, exportContext);
}



function excelSheetSortStateForExport(sheet = {}) {
  return sheet.sortState || visibleTableSortStates(sheet)[0]?.sortState || null;
}



function excelAutoFilterXmlForState(filter = null, sortState = null, exportContext = null) {
  const range = filter?.range;
  if (!range) return "";
  const filterColumns = excelAutoFilterColumnXml(filter, range, exportContext);
  const sortXml = excelSortStateXml(filter.sortState || sortState, range, exportContext);
  const body = `${filterColumns}${sortXml}`;
  const ref = escapeAttr(rangeToLabel(range));
  return body ? `<autoFilter ref="${ref}">${body}</autoFilter>` : `<autoFilter ref="${ref}"/>`;
}



function excelAutoFilterColumnXml(filter, range, exportContext = null) {
  const criteria = filter?.criteria || {};
  const hiddenButtons = filter?.hiddenButtons || {};
  const columns = new Set([
    ...Object.keys(criteria).map(Number).filter(Number.isFinite),
    ...Object.keys(hiddenButtons).map(Number).filter(Number.isFinite),
  ]);
  return [...columns]
    .filter((col) => col >= range.left && col <= range.right)
    .sort((a, b) => a - b)
    .map((col) => {
      const criterionXml = excelAutoFilterCriterionXml(criteria[col], exportContext);
      const attrs = [
        `colId="${col - range.left}"`,
        hiddenButtons[col] ? 'hiddenButton="1" showButton="0"' : "",
      ].filter(Boolean).join(" ");
      return criterionXml ? `<filterColumn ${attrs}>${criterionXml}</filterColumn>` : `<filterColumn ${attrs}/>`;
    })
    .join("");
}



function excelAutoFilterCriterionXml(criterion, exportContext = null) {
  const normalized = normalizeAutoFilterCriterion(criterion);
  if (normalized.kind === "values") return excelAutoFilterValuesXml(normalized.values || []);
  if (normalized.kind === "dateGroups") return excelAutoFilterDateGroupsXml(normalized);
  if (normalized.kind === "custom") return excelAutoFilterCustomCriterionXml(normalized);
  if (normalized.kind === "color") return excelAutoFilterColorCriterionXml(normalized, exportContext);
  return "";
}



function excelAutoFilterColorCriterionXml(criterion, exportContext = null) {
  const target = criterion?.target === "font" ? "font" : "fill";
  const color = normalizeColorValue(criterion?.color);
  const dxfId = exportContext?.dxfIdForColorFilter?.(target, color);
  if (!Number.isInteger(dxfId)) return "";
  const cellColor = target === "font" ? "0" : "1";
  return `<colorFilter dxfId="${dxfId}" cellColor="${cellColor}"/>`;
}



function excelAutoFilterValuesXml(values = []) {
  const unique = [];
  const seen = new Set();
  let includeBlank = false;
  values.forEach((item) => {
    const value = autoFilterExportValue(item);
    if (value === "") {
      includeBlank = true;
      return;
    }
    if (seen.has(value)) return;
    seen.add(value);
    unique.push(value);
  });
  const filters = unique.map((value) => `<filter val="${escapeAttr(value)}"/>`).join("");
  return `<filters${includeBlank ? ' blank="1"' : ""}>${filters}</filters>`;
}



function autoFilterExportValue(item) {
  if (item && typeof item === "object" && "value" in item) return String(item.value ?? "");
  return String(item ?? "");
}



function excelAutoFilterDateGroupsXml(criterion = {}) {
  const filterValues = Array.isArray(criterion.values) ? criterion.values : [];
  const filters = filterValues
    .map(autoFilterExportValue)
    .filter((value) => value !== "")
    .map((value) => `<filter val="${escapeAttr(value)}"/>`);
  const groups = (criterion.dateGroups || []).map(excelAutoFilterDateGroupItemXml).filter(Boolean);
  const includeBlank = Boolean(criterion.includeBlank) || filterValues.some((value) => autoFilterExportValue(value) === "");
  const body = `${filters.join("")}${groups.join("")}`;
  return body || includeBlank ? `<filters${includeBlank ? ' blank="1"' : ""}>${body}</filters>` : "";
}



function excelAutoFilterDateGroupItemXml(dateGroup = {}) {
  const year = Math.round(Number(dateGroup.year));
  if (!Number.isFinite(year)) return "";
  const grouping = dateGroup.grouping || "year";
  const includeMonth = ["month", "day", "hour", "minute", "second"].includes(grouping);
  const includeDay = ["day", "hour", "minute", "second"].includes(grouping);
  const includeHour = ["hour", "minute", "second"].includes(grouping);
  const includeMinute = ["minute", "second"].includes(grouping);
  const includeSecond = grouping === "second";
  const attrs = [
    `year="${year}"`,
    includeMonth && Number.isFinite(Number(dateGroup.month)) ? `month="${Math.round(Number(dateGroup.month))}"` : "",
    includeDay && Number.isFinite(Number(dateGroup.day)) ? `day="${Math.round(Number(dateGroup.day))}"` : "",
    includeHour && Number.isFinite(Number(dateGroup.hour)) ? `hour="${Math.round(Number(dateGroup.hour))}"` : "",
    includeMinute && Number.isFinite(Number(dateGroup.minute)) ? `minute="${Math.round(Number(dateGroup.minute))}"` : "",
    includeSecond && Number.isFinite(Number(dateGroup.second)) ? `second="${Math.round(Number(dateGroup.second))}"` : "",
    `dateTimeGrouping="${escapeAttr(grouping)}"`,
  ].filter(Boolean).join(" ");
  return `<dateGroupItem ${attrs}/>`;
}



function excelAutoFilterCustomCriterionXml(criterion = {}) {
  const dynamicXml = excelDynamicFilterXml(criterion);
  if (dynamicXml) return dynamicXml;
  const dateGroup = criterion.dateGroup;
  if (dateGroup?.year) {
    const groupXml = excelAutoFilterDateGroupItemXml(dateGroup);
    return groupXml ? `<filters>${groupXml}</filters>` : "";
  }
  const top10 = criterion.top10 || criterion.top10Source;
  if (top10) {
    const count = Math.max(1, Math.round(Number(top10.count) || 10));
    const attrs = [
      top10.direction === "bottom" ? 'top="0"' : "",
      top10.mode === "percent" ? 'percent="1"' : "",
      `val="${count}"`,
      criterion.value1 != null && criterion.value1 !== "" ? `filterVal="${escapeAttr(criterion.value1)}"` : "",
    ].filter(Boolean).join(" ");
    return `<top10 ${attrs}/>`;
  }
  if (criterion.blankOperator === "blanks" || criterion.operator1 === "blank") return '<filters blank="1"/>';
  if (criterion.blankOperator === "nonBlanks" || criterion.operator1 === "nonBlank") return '<customFilters><customFilter operator="notEqual" val=""/></customFilters>';
  const first = excelCustomFilterXml(criterion.operator1, criterion.value1);
  const hasSecond = String(criterion.value2 ?? "").trim() !== "";
  const second = hasSecond ? excelCustomFilterXml(criterion.operator2, criterion.value2) : "";
  if (!first) return "";
  return `<customFilters${criterion.logic === "and" && second ? ' and="1"' : ""}>${first}${second}</customFilters>`;
}



function excelDynamicFilterType(criterion = {}) {
  if (criterion.dynamicSource?.type) return criterion.dynamicSource.type;
  if (criterion.datePeriod) return autoFilterDynamicDatePeriodType(criterion.datePeriod);
  if (criterion.averageOperator === "aboveAverage" || criterion.averageOperator === "belowAverage") return criterion.averageOperator;
  return {
    dateToday: "today",
    dateYesterday: "yesterday",
    dateTomorrow: "tomorrow",
    dateThisWeek: "thisWeek",
    dateLastWeek: "lastWeek",
    dateNextWeek: "nextWeek",
    dateThisMonth: "thisMonth",
    dateLastMonth: "lastMonth",
    dateNextMonth: "nextMonth",
    dateThisQuarter: "thisQuarter",
    dateLastQuarter: "lastQuarter",
    dateNextQuarter: "nextQuarter",
    dateThisYear: "thisYear",
    dateLastYear: "lastYear",
    dateNextYear: "nextYear",
    dateYearToDate: "yearToDate",
  }[criterion.datePreset] || "";
}



function excelDynamicFilterXml(criterion = {}) {
  const dynamicType = excelDynamicFilterType(criterion);
  if (!dynamicType) return "";
  const attrs = [`type="${escapeAttr(dynamicType)}"`];
  if (dynamicType === "aboveStdDev" || dynamicType === "belowStdDev") {
    attrs.push(`val="${autoFilterStdDevCount(criterion.dynamicSource?.val ?? criterion.stdDevCount)}"`);
  }
  return `<dynamicFilter ${attrs.join(" ")}/>`;
}



function excelCustomFilterXml(operator, value) {
  const xmlOperator = excelAutoFilterOperatorXml(operator);
  if (!xmlOperator) return "";
  return `<customFilter operator="${xmlOperator}" val="${escapeAttr(excelCustomFilterValue(operator, value))}"/>`;
}



function excelCustomFilterValue(operator, value) {
  const text = String(value ?? "");
  const escaped = escapeExcelCustomFilterLiteral(text);
  if (operator === "contains" || operator === "notContains" || operator === "doesNotContain") return `*${escaped}*`;
  if (operator === "beginsWith" || operator === "notBeginsWith") return `${escaped}*`;
  if (operator === "endsWith" || operator === "notEndsWith") return `*${escaped}`;
  return text;
}



function escapeExcelCustomFilterLiteral(value) {
  return String(value ?? "").replace(/[~*?]/g, (char) => `~${char}`);
}



function excelAutoFilterOperatorXml(operator) {
  return {
    equal: "equal",
    notEqual: "notEqual",
    greaterThan: "greaterThan",
    greaterThanOrEqual: "greaterThanOrEqual",
    lessThan: "lessThan",
    lessThanOrEqual: "lessThanOrEqual",
    contains: "equal",
    notContains: "notEqual",
    doesNotContain: "notEqual",
    beginsWith: "equal",
    notBeginsWith: "notEqual",
    endsWith: "equal",
    notEndsWith: "notEqual",
  }[operator] || "equal";
}



function excelSortStateXml(sortState, autoFilterRange = null, exportContext = null) {
  if (!sortState?.range || !Array.isArray(sortState.keys) || !sortState.keys.length) return "";
  const range = sortState.range;
  const refRange = sortState.hasHeader && sortState.orientation !== "columns"
    ? { ...range, top: Math.min(range.bottom, range.top + 1) }
    : sortState.hasHeader && sortState.orientation === "columns"
      ? { ...range, left: Math.min(range.right, range.left + 1) }
      : range;
  const ref = rangeToLabel(refRange);
  const conditions = sortState.keys.map((key) => excelSortConditionXml(key, refRange, sortState.orientation, exportContext)).filter(Boolean).join("");
  if (!conditions) return "";
  const attrs = [
    `ref="${escapeAttr(ref)}"`,
    sortState.orientation === "columns" ? 'columnSort="1"' : "",
    sortState.caseSensitive ? 'caseSensitive="1"' : "",
  ].filter(Boolean).join(" ");
  return `<sortState ${attrs}>${conditions}</sortState>`;
}



function excelSortConditionXml(key, range, orientation = "rows", exportContext = null) {
  const descending = key.direction === "desc" ? ' descending="1"' : "";
  const sortBy = key.sortOn === "cellColor" ? ' sortBy="cellColor"' : key.sortOn === "fontColor" ? ' sortBy="fontColor"' : "";
  const colorAttrs = excelSortConditionColorAttrs(key, exportContext);
  const customList = excelSortConditionCustomListAttr(key);
  const ref = orientation === "columns"
    ? rangeToLabel({ top: key.row || range.top, bottom: key.row || range.top, left: range.left, right: range.right })
    : rangeToLabel({ top: range.top, bottom: range.bottom, left: key.col || range.left, right: key.col || range.left });
  return `<sortCondition ref="${escapeAttr(ref)}"${descending}${sortBy}${colorAttrs}${customList}/>`;
}



function excelSortConditionColorAttrs(key, exportContext = null) {
  const target = key?.sortOn === "fontColor" ? "font" : key?.sortOn === "cellColor" ? "fill" : "";
  const color = normalizeColorValue(key?.color || "");
  if (!target || !color) return "";
  const dxfId = exportContext?.dxfIdForColorFilter?.(target, color);
  if (Number.isInteger(dxfId)) return ` dxfId="${dxfId}"`;
  const argb = cssColorToExcelArgb(color);
  return argb ? ` color="${escapeAttr(argb)}"` : ` color="${escapeAttr(color)}"`;
}



function excelSortConditionCustomListAttr(key) {
  if (key?.sortOn === "cellColor" || key?.sortOn === "fontColor") return "";
  const list = sortCustomListById(key?.customList);
  return list ? ` customList="${escapeAttr(list.values.join(","))}"` : "";
}



function insertWorksheetChildXml(xml, childXml) {
  return String(xml || "").replace(/(<\/sheetData>)/, `$1${childXml}`);
}



async function patchExcelTableDefinitions(zip, model, exportContext = null) {
  const tableLookup = excelModelTableLookup(model);
  if (!tableLookup.entries.length) return;
  const tableFiles = Object.keys(zip.files)
    .filter((path) => /^xl\/tables\/table\d+\.xml$/i.test(path))
    .sort((left, right) => left.localeCompare(right, "en", { numeric: true }));
  await Promise.all(tableFiles.map(async (path) => {
    const file = zip.file(path);
    if (!file) return;
    const xml = await file.async("text");
    const table = excelModelTableForXml(xml, tableLookup);
    if (!table) return;
    const nextXml = patchExcelTableXml(xml, table, exportContext);
    if (nextXml !== xml) zip.file(path, nextXml);
  }));
}



function excelModelTableLookup(model) {
  const entries = [];
  (model?.sheets || []).forEach((sheet, sheetIndex) => {
    (sheet.tables || []).forEach((table, tableIndex) => {
      if (!table?.range) return;
      const ref = rangeToLabel(table.range);
      const names = [table.name, table.displayName, `Table${tableIndex + 1}`]
        .map((name) => String(name || "").trim())
        .filter(Boolean);
      entries.push({ sheetIndex, sheet, table, tableIndex, ref, names });
    });
  });
  return { entries };
}



function excelModelTableForXml(xml, lookup) {
  const tableTag = String(xml || "").match(/<table\b[^>]*>/)?.[0] || "";
  const name = xmlAttributeValue(tableTag, "name").toLocaleLowerCase("ja-JP");
  const displayName = xmlAttributeValue(tableTag, "displayName").toLocaleLowerCase("ja-JP");
  const ref = String(xmlAttributeValue(tableTag, "ref") || "").replace(/\$/g, "");
  const byName = lookup.entries.find((entry) =>
    entry.names.some((item) => {
      const normalized = item.toLocaleLowerCase("ja-JP");
      return normalized === name || normalized === displayName;
    }),
  );
  if (byName) return byName.table;
  return lookup.entries.find((entry) => entry.ref === ref)?.table || null;
}



function patchExcelTableXml(xml, table, exportContext = null) {
  let next = String(xml || "");
  next = patchExcelTableAutoFilterXml(next, table, exportContext);
  next = patchExcelTableColumnFormulasXml(next, table);
  return next;
}



function patchExcelTableAutoFilterXml(xml, table, exportContext = null) {
  if (table.showFilterButton === false) {
    return xml.replace(/<autoFilter\b[\s\S]*?<\/autoFilter>|<autoFilter\b[^>]*\/>/, "");
  }
  const filter = table.autoFilterState?.range
    ? table.autoFilterState
    : { range: table.range, criteria: {}, hiddenButtons: {} };
  if (!filter?.range) return xml;
  const tableFilter = { ...filter };
  delete tableFilter.sortState;
  const autoFilterXml = excelAutoFilterXmlForState(tableFilter, null, exportContext);
  if (!autoFilterXml) return xml;
  if (/<autoFilter\b[\s\S]*?<\/autoFilter>|<autoFilter\b[^>]*\/>/.test(xml)) {
    return xml.replace(/<autoFilter\b[\s\S]*?<\/autoFilter>|<autoFilter\b[^>]*\/>/, autoFilterXml);
  }
  return xml.replace(/(<tableColumns\b)/, `${autoFilterXml}$1`);
}



function patchExcelTableColumnFormulasXml(xml, table) {
  const columns = Array.isArray(table.columns) ? table.columns : [];
  if (!columns.some((column) => column?.calculatedColumnFormula || column?.totalsRowFormula)) return xml;
  let index = 0;
  return String(xml || "").replace(/<tableColumn\b[^>]*(?:\/>|>[\s\S]*?<\/tableColumn>)/g, (columnXml) => {
    const model = columns[index++] || {};
    let next = columnXml;
    if (model.calculatedColumnFormula) {
      next = upsertTableColumnChildXml(next, "calculatedColumnFormula", model.calculatedColumnFormula);
    }
    if (model.totalsRowFormula) {
      next = upsertTableColumnChildXml(next, "totalsRowFormula", model.totalsRowFormula);
    }
    return next;
  });
}



function upsertTableColumnChildXml(columnXml, childName, formula) {
  const text = String(formula || "").replace(/^=/, "");
  if (!text) return columnXml;
  const namePattern = escapeRegExp(childName);
  const child = `<${childName}>${escapeHtml(text)}</${childName}>`;
  const existing = new RegExp(`<${namePattern}\\b[\\s\\S]*?<\\/${namePattern}>`);
  if (existing.test(columnXml)) return columnXml.replace(existing, child);
  if (/\/>$/.test(columnXml)) return columnXml.replace(/\/>$/, `>${child}</tableColumn>`);
  return columnXml.replace(/<\/tableColumn>$/, `${child}</tableColumn>`);
}



function excelCellAddress(row, col) {
  return `${columnName(Math.max(1, Math.round(Number(col) || 1)))}${Math.max(1, Math.round(Number(row) || 1))}`;
}



function excelColorObjectFromCss(value) {
  const argb = cssColorToExcelArgb(value);
  return argb ? { argb } : null;
}



function excelPaperSizeCode(value) {
  return { Letter: 1, A3: 8, A4: 9, B5: 13 }[value] || 9;
}



function excelPageCommentsMode(value) {
  return { atEnd: "atEnd", asDisplayed: "asDisplayed" }[value] || "None";
}



function excelPageErrorsMode(value) {
  return { blank: "blank", dash: "dash", na: "NA" }[value] || "displayed";
}



function sanitizeWorksheetName(name) {
  const text = String(name || "Sheet").replace(/[:\\/?*\[\]]/g, "_").slice(0, 31).trim();
  return text || "Sheet";
}



function exportSheetDimensionsToExcel(sheet, worksheet) {
  const columnCount = Math.max(1, Number(sheet.colCount) || sheet.columns?.length || 1);
  for (let col = 1; col <= columnCount; col += 1) {
    const column = sheet.columns?.[col - 1] || {};
    const target = worksheet.getColumn(col);
    target.width = pixelsToExcelColumnWidth(column.width || sheet.defaultColumnWidth || columnWidthToPixels(8.43));
    target.hidden = column.hidden === true;
    if (Number.isFinite(Number(column.outlineLevel)) && Number(column.outlineLevel) > 0) {
      target.outlineLevel = Math.max(0, Math.min(7, Math.round(Number(column.outlineLevel))));
    }
  }
  const rowCount = Math.max(1, Number(sheet.rowCount) || sheet.rows?.length || 1);
  for (let row = 1; row <= rowCount; row += 1) {
    const rowModel = sheet.rows?.[row - 1] || {};
    const target = worksheet.getRow(row);
    target.height = pixelsToExcelRowHeight(rowModel.height || sheet.defaultRowHeight || rowHeightToPixels(15));
    target.hidden = rowModel.hidden === true;
    if (Number.isFinite(Number(rowModel.outlineLevel)) && Number(rowModel.outlineLevel) > 0) {
      target.outlineLevel = Math.max(0, Math.min(7, Math.round(Number(rowModel.outlineLevel))));
    }
  }
}



function exportSheetTablesToExcel(sheet, worksheet, tableNames = new Set(), options = {}) {
  (sheet.tables || []).forEach((table, index) => {
    if (!table?.range) return;
    const range = table.range;
    const columnCount = Math.max(1, range.right - range.left + 1);
    const headerRow = table.showHeaderRow !== false;
    const totalsRow = table.showTotalsRow === true;
    const dataTop = range.top + (headerRow ? 1 : 0);
    const dataBottom = range.bottom - (totalsRow ? 1 : 0);
    const rows = [];
    for (let row = dataTop; row <= dataBottom; row += 1) {
      const values = [];
      for (let col = range.left; col <= range.right; col += 1) {
        values.push(excelTableCellValue(sheet, row, col, options));
      }
      rows.push(values);
    }
    const columns = Array.from({ length: columnCount }, (_item, offset) => excelTableColumnModel(sheet, table, offset));
    const name = uniqueExcelTableName(table.name || table.displayName || `Table${index + 1}`, tableNames);
    try {
      worksheet.addTable({
        name,
        displayName: name,
        ref: excelCellAddress(range.top, range.left),
        headerRow,
        totalsRow,
        style: excelTableStyleModel(table),
        columns,
        rows,
      });
    } catch (error) {
      console.warn("Failed to export table.", table?.name || index, error);
    }
  });
}



function excelTableCellValue(sheet, row, col, options = {}) {
  const cell = sheet.cells?.[cellKey(row, col)];
  if (!cell) return null;
  const formula = cell.formula || (typeof cell.raw === "string" && cell.raw.startsWith("=") ? cell.raw : "");
  if (formula) return { formula: formula.replace(/^=/, ""), result: excelExportScalarValue(cell.cached ?? null, cell, options) };
  if (cell.raw != null) return excelExportScalarValue(cell.raw, cell, options);
  if (cell.cached != null) return excelExportScalarValue(cell.cached, cell, options);
  return cell.display || "";
}



function excelTableColumnModel(sheet, table, offset) {
  const range = table.range;
  const col = range.left + offset;
  const source = table.columns?.[offset] || {};
  const headerText = source.name || (table.showHeaderRow !== false ? getDisplayForCell(sheet, getSheetId(sheet.name), range.top, col) : "") || `Column${offset + 1}`;
  const hiddenButtons = table.autoFilterState?.hiddenButtons || {};
  return {
    name: String(headerText || `Column${offset + 1}`),
    filterButton: table.showFilterButton !== false && hiddenButtons[col] !== true && hiddenButtons[offset] !== true,
    totalsRowLabel: source.totalsRowLabel || (offset === 0 ? "Total" : ""),
    totalsRowFunction: excelTableTotalsFunction(source),
    ...(source.totalsRowFormula ? { totalsRowFormula: String(source.totalsRowFormula).replace(/^=/, "") } : {}),
  };
}



function excelTableTotalsFunction(column = {}) {
  const fn = String(column.totalsRowFunction || "").trim();
  const allowed = new Set(["none", "average", "countNums", "count", "max", "min", "stdDev", "var", "sum", "custom"]);
  if (allowed.has(fn)) return fn;
  return column.totalsRowFormula ? "custom" : "none";
}



function excelTableStyleModel(table = {}) {
  const styleInfo = table.styleInfo || {};
  return {
    theme: styleInfo.name || table.style || "TableStyleMedium2",
    showFirstColumn: styleInfo.showFirstColumn === true,
    showLastColumn: styleInfo.showLastColumn === true,
    showRowStripes: styleInfo.showRowStripes !== false,
    showColumnStripes: styleInfo.showColumnStripes === true,
  };
}



function excelSafeTableName(value) {
  const text = String(value || "Table1").replace(/[^\p{L}\p{N}_\\.]/gu, "_").replace(/^[^A-Za-z_]/, "_").slice(0, 240);
  return text && !/^[A-Z]{1,3}\d+$/i.test(text) ? text : `Table_${text || "1"}`;
}



function uniqueExcelTableName(value, usedNames = new Set()) {
  const base = excelSafeTableName(value);
  let name = base;
  let index = 1;
  while (usedNames.has(name.toLocaleLowerCase("ja-JP"))) {
    index += 1;
    name = `${base}_${index}`;
  }
  usedNames.add(name.toLocaleLowerCase("ja-JP"));
  return name;
}



function exportSheetCellsToExcel(sheet, worksheet, options = {}) {
  const sheetId = getSheetId(sheet.name);
  Object.values(sheet.cells || {}).forEach((cell) => {
    if (!cell || !Number.isFinite(Number(cell.row)) || !Number.isFinite(Number(cell.col))) return;
    const target = worksheet.getCell(Number(cell.row), Number(cell.col));
    target.value = excelExportCellValue(sheet, sheetId, cell, options);
    applyExcelExportCellStyle(sheet, cell, target);
    applyExcelExportCellValidation(cell, target);
    applyExcelExportCellNote(cell, target);
  });
}



function excelExportCellValue(sheet, sheetId, cell, options = {}) {
  const formula = cell.formula || (typeof cell.raw === "string" && cell.raw.startsWith("=") ? cell.raw : "");
  if (formula) {
    return {
      formula: formula.replace(/^=/, ""),
      result: excelExportScalarValue(cell.cached ?? null, cell, options),
    };
  }
  if (cell.hyperlink) {
    const text = cell.display || cell.raw || cell.hyperlink;
    return { text: String(text ?? ""), hyperlink: cell.hyperlink };
  }
  const richText = excelExportRichTextValue(cell);
  if (richText) return richText;
  if (cell.raw != null) return excelExportScalarValue(cell.raw, cell, options);
  if (cell.cached != null) return excelExportScalarValue(cell.cached, cell, options);
  const display = getDisplayForCell(sheet, sheetId, cell.row, cell.col);
  return display === "" ? null : display;
}



function excelExportScalarValue(value, cell = {}, options = {}) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const serial = dateToExcelSerial(value);
    return options.date1904 && isDateFormat(cell.numFmt) ? excelSerial1900To1904(serial) : serial;
  }
  if (typeof value === "number" && options.date1904 && isDateFormat(cell.numFmt)) {
    return excelSerial1900To1904(value);
  }
  return value;
}



function excelExportRichTextValue(cell) {
  const runs = Array.isArray(cell?.richText) ? cell.richText : [];
  if (!runs.length) return null;
  const richText = runs
    .map((run) => {
      const text = normalizeRichTextLineEndings(run?.text || "");
      if (!text) return null;
      const font = excelFontFromCss(run?.css || {});
      return font ? { text, font } : { text };
    })
    .filter(Boolean);
  return richText.length ? { richText } : null;
}



function applyExcelExportCellStyle(sheet, cell, target) {
  const css = cellEffectiveCss(sheet, cell);
  const font = excelFontFromCss(css);
  const alignment = excelAlignmentFromCss(css);
  const fill = excelFillFromCss(css, cell);
  const border = excelBorderFromModel(cell.borders);
  const protection = excelProtectionFromModel(cell.protection);
  if (cell.numFmt) target.numFmt = cell.numFmt;
  if (font) target.font = font;
  if (alignment) target.alignment = alignment;
  if (fill) target.fill = fill;
  if (border) target.border = border;
  if (protection) target.protection = protection;
}



function excelFontFromCss(css = {}) {
  const font = {};
  const family = String(css.fontFamily || "").split(",")[0].replace(/^['"]|['"]$/g, "").trim();
  const size = parseCssFontSizePt(css.fontSize);
  const color = cssColorToExcelArgb(css.color);
  if (family) font.name = family;
  if (size) font.size = size;
  if (String(css.fontWeight || "") === "700" || String(css.fontWeight || "").toLowerCase() === "bold") font.bold = true;
  if (String(css.fontStyle || "").toLowerCase() === "italic") font.italic = true;
  const decoration = String(css.textDecoration || "").toLowerCase();
  if (decoration.includes("underline")) font.underline = decoration.includes("double") ? "double" : true;
  if (decoration.includes("line-through")) font.strike = true;
  if (color) font.color = { argb: color };
  return Object.keys(font).length ? font : null;
}



function excelAlignmentFromCss(css = {}) {
  const alignment = {};
  const horizontal = { left: "left", center: "center", right: "right", justify: "justify" }[css.textAlign];
  const vertical = { top: "top", middle: "middle", bottom: "bottom" }[css.verticalAlign] || (css.alignItems === "center" ? "middle" : "");
  if (horizontal) alignment.horizontal = horizontal;
  if (vertical) alignment.vertical = vertical;
  if (css.whiteSpace === "pre-wrap") alignment.wrapText = true;
  if (css.textRotation === "vertical") alignment.textRotation = "vertical";
  else if (Number.isFinite(Number(css.textRotation))) alignment.textRotation = Number(css.textRotation);
  return Object.keys(alignment).length ? alignment : null;
}



function excelFillFromCss(css = {}, cell = {}) {
  const color = cssColorToExcelArgb(css.backgroundColor);
  if (!color && !cell.hasFill) return null;
  return {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: color || "FFFFFFFF" },
  };
}



function excelBorderFromModel(borders = {}) {
  const result = {};
  CELL_BORDER_SIDES.forEach((side) => {
    const border = borders?.[side];
    if (!border) return;
    const color = cssColorToExcelArgb(border.color);
    result[side] = {
      style: excelBorderStyleFromModel(border),
      ...(color ? { color: { argb: color } } : {}),
    };
  });
  return Object.keys(result).length ? result : null;
}



function excelBorderStyleFromModel(border = {}) {
  if (border.style === "double") return "double";
  if (border.style === "dashed") return Number(border.weight || border.width) >= 2 ? "mediumDashed" : "dashed";
  if (border.style === "dotted") return "dotted";
  if (Number(border.weight || border.width) >= 3) return "thick";
  if (Number(border.weight || border.width) >= 2) return "medium";
  return "thin";
}



function excelProtectionFromModel(protection) {
  if (!protection || typeof protection !== "object") return null;
  const result = {};
  if (protection.locked === false) result.locked = false;
  if (protection.locked === true) result.locked = true;
  if (protection.hidden === true) result.hidden = true;
  if (protection.hidden === false) result.hidden = false;
  return Object.keys(result).length ? result : null;
}



function applyExcelExportCellValidation(cell, target) {
  const validation = cell.validation;
  if (!validation) return;
  const rule = {
    type: validation.type || "custom",
    allowBlank: validation.allowBlank === true,
    showInputMessage: validation.showInputMessage === true,
    promptTitle: validation.inputTitle || "",
    prompt: validation.inputMessage || "",
    showErrorMessage: validation.showErrorMessage !== false,
    errorStyle: validation.errorStyle || "stop",
    errorTitle: validation.errorTitle || "",
    error: validation.errorMessage || "",
  };
  if (validation.type === "list") {
    rule.formulae = [excelDataValidationListFormula(validation)];
  } else {
    rule.operator = validation.operator || undefined;
    rule.formulae = excelDataValidationFormulae(validation);
  }
  target.dataValidation = rule;
}



function excelDataValidationListFormula(validation = {}) {
  const formula = String(validation.formula || "").trim();
  if (formula) return formula.replace(/^=/, "");
  const options = Array.isArray(validation.options) ? validation.options : [];
  if (!options.length) return "";
  return `"${options.map((option) => String(option ?? "").replace(/"/g, '""')).join(",")}"`;
}



function excelDataValidationFormulae(validation = {}) {
  return (validation.formulae || [])
    .map((formula) => excelDataValidationFormula(validation.type, formula))
    .filter((formula) => formula !== "");
}



function excelDataValidationFormula(type, formula) {
  if (formula == null || formula === "") return "";
  if (formula instanceof Date && !Number.isNaN(formula.getTime())) return formula;
  const text = String(formula).trim();
  if (type === "custom") return text.replace(/^=/, "");
  if (type === "date" || type === "time") {
    if (/^-?\d+(?:\.\d+)?$/.test(text) && type === "date") return excelSerialToDate(Number(text));
    const parsed = persistedDateDisplayValue(text);
    if (parsed) return parsed;
  }
  return text.replace(/^=/, "");
}



function applyExcelExportCellNote(cell, target) {
  const note = excelNoteFromCell(cell);
  if (note) target.note = note;
}



function excelNoteFromCell(cell = {}) {
  const commentText = excelCommentThreadExportText(cell.comment);
  const noteText = normalizeNote(cell.note);
  const body = [commentText, noteText ? `${annotationAuthor()}:\n${noteText}` : ""].filter(Boolean).join("\n\n");
  if (!body) return null;
  return {
    texts: [
      {
        font: { size: 9, name: "Tahoma", color: { argb: "FF000000" } },
        text: body,
      },
    ],
    margins: { insetmode: "auto", inset: [0.13, 0.13, 0.25, 0.25] },
    protection: { locked: "True", lockText: "True" },
    editAs: cell.noteVisible ? "oneCells" : "absolute",
  };
}



function excelCommentThreadExportText(comment) {
  const replies = normalizeCommentReplies(comment);
  return replies
    .map((reply) => {
      const author = String(reply.author || annotationAuthor()).trim() || annotationAuthor();
      const createdAt = String(reply.createdAt || "").trim();
      const header = createdAt ? `${author} (${createdAt})` : author;
      return `${header}:\n${String(reply.text || "")}`;
    })
    .filter((text) => text.trim())
    .join("\n\n");
}



function exportSheetMergesToExcel(sheet, worksheet) {
  (sheet.merges || []).forEach((merge) => {
    try {
      worksheet.mergeCells(merge.top, merge.left, merge.bottom, merge.right);
    } catch {
      // Ignore overlapping/invalid merges rather than aborting the whole save.
    }
  });
}



async function exportSheetImagesToExcel(sheet, workbook, worksheet) {
  for (const image of sheet.images || []) {
    if (!image?.src || image.hidden) continue;
    let payload = null;
    try {
      payload = await excelImagePayload(image);
    } catch {
      payload = null;
    }
    if (!payload) continue;
    try {
      const imageId = workbook.addImage(payload);
      worksheet.addImage(imageId, excelImagePlacement(sheet, image));
    } catch {
      // ExcelJS/browser canvas cannot encode every source. Keep saving the sheet data.
    }
  }
}



async function excelImagePayload(image) {
  const src = excelImageExportSource(image);
  const direct = excelDirectImagePayload(src);
  if (direct) return direct;
  return rasterizedImagePayload(src, image.ext?.width, image.ext?.height);
}



// Uses the worksheet's own cell painter and shape SVG renderer. This is an
// offscreen render of the current preview, not a model's prediction of layout.
function cwsAgentVisualDetailRegions(width, height, hasCanvas = false) {
  // A wireframe already has a complete native canvas crop. Partial crops of
  // that canvas introduce misleading coordinate origins for source comparison.
  if (hasCanvas || (width <= 1600 && height <= 1200)) return [];
  // Native-pixel crops supplement the whole sheet; their overlapping cut
  // edges are explicitly identified for the visual reviewer.
  const spans = (length, threshold) => length <= threshold ? [[0,length]]
    : [[0,Math.ceil(length/2)+32],[Math.floor(length/2)-32,length]];
  return spans(height,1200).flatMap(([top,bottom]) => spans(width,1600).map(([left,right]) => ({left,top,width:right-left,height:bottom-top})));
}



async function captureCwsAgentVisualEvidence(sheetNames, cellFocus = null) {
  await document.fonts.ready;
  const results = [];
  for (const name of sheetNames) {
    const original = state.model?.sheets?.find(sheet => sheet.name === name);
    if (!original) throw new Error(`視覚確認の対象シートがありません: ${name}`);
    const sheet = structuredClone(original);
    scaleSheetDisplay(sheet,1/Math.max(.01,Number(sheet.displayScale)||1));
    sheet.zoomScale = 100; sheet.baseZoomScale = 100; sheet.displayScale = 1;
    const focus = cellFocus?.sheetName === name ? cellFocus : null;
    const detailRange = focus && {left:focus.col,right:focus.col,top:focus.row,bottom:focus.row};
    const target = focus && rangePixelRect(sheet,detailRange);
    if (focus && (getCellRawInput(sheet,focus.row,focus.col) !== focus.value ||
        !(target.width>0 && target.height>0) || sheet.columns?.[focus.col-1]?.hidden || sheet.rows?.[focus.row-1]?.hidden)) {
      throw new Error('対象セルの値または表示範囲が確定していません。');
    }
    const range = focus ? {left:Math.max(1,focus.col-1),right:Math.min(sheet.colCount,focus.col+1),
      top:Math.max(1,focus.row-1),bottom:Math.min(sheet.rowCount,focus.row+1)} : sheetUsedRangeForPrint(sheet);
    if (!range) throw new Error(`視覚確認の対象シートが空です: ${name}`);
    for (const merge of focus ? [] : (sheet.merges || [])) {
      range.left=Math.min(range.left,merge.left);range.top=Math.min(range.top,merge.top);
      range.right=Math.max(range.right,merge.right);range.bottom=Math.max(range.bottom,merge.bottom);
    }
    const origin = printRangePixelBounds(sheet, range);
    const boxes = focus ? [] : (sheet.images || []).filter(image => !image.hidden).map(image => imageBoxToPixels(sheet,image)).filter(Boolean);
    // A small border preserves spill, overlap and adjacency evidence without
    // making nearby cells independent work items. The sheet painter is shared.
    const left = focus ? Math.max(0,Math.floor(target.left)-12) : Math.min(origin.left,...boxes.map(b=>b.left));
    const top = focus ? Math.max(0,Math.floor(target.top)-8) : Math.min(origin.top,...boxes.map(b=>b.top));
    const width = focus ? Math.max(32,Math.ceil(target.right-left)+12) : Math.ceil(Math.max(origin.right,...boxes.map(b=>b.left+b.width))-left+8);
    const height = focus ? Math.max(20,Math.ceil(target.bottom-top)+8) : Math.ceil(Math.max(origin.bottom,...boxes.map(b=>b.top+b.height))-top+8);
    if (width > 8192 || height > 8192 || width*height > 25000000) throw new Error(`視覚確認の描画範囲が大きすぎます: ${name}`);
    const canvas = document.createElement('canvas'); canvas.width=width; canvas.height=height;
    const textLayout={version:1,checkedShapes:0,checkedCells:0,issues:[]};
    const ctx=canvas.getContext('2d'); if (!ctx) throw new Error('視覚確認用の描画に失敗しました。');
    ctx.fillStyle='#fff';ctx.fillRect(0,0,width,height);ctx.translate(-left,-top);
    const maps=buildMergeMaps(sheet.merges), cache=createPixelPositionCache(sheet), id=getSheetId(name);
    paintCanvasCells(ctx,sheet,maps,id,range,cache,'base');
    paintCanvasBorders(ctx,sheet,maps,range,cache);
    paintCanvasCells(ctx,sheet,maps,id,range,cache,'text');
    measureCwsAgentCellTextLayout(sheet,maps,id,ctx,textLayout,focus);
    for (const image of sheet.images || []) {
      if (image.hidden) continue;
      const box=imageBoxToPixels(sheet,image); if (!box) throw new Error(`図形の描画位置を確認できません: ${name}`);
      if (focus) {
        // Enclose even rotated shapes: anything that can occlude the target
        // must be painted into its crop.
        const radius=Math.hypot(box.width,box.height)/2, x=box.left+box.width/2, y=box.top+box.height/2;
        if (x+radius<left || x-radius>left+width || y+radius<top || y-radius>top+height) continue;
      }
      const element=await loadImageElement(worksheetImageRenderSource(image));
      ctx.save();
      const rotation=Number(image.rotation || image.shape?.rotation || 0);
      ctx.translate(box.left+box.width/2,box.top+box.height/2);ctx.rotate(rotation*Math.PI/180);
      ctx.drawImage(element,-box.width/2,-box.height/2,box.width,box.height);
      const textLayer=await renderedShapeTextEvidence(image,box.width,box.height,focus ? null : textLayout);
      const measuredRun=textLayout.inlineRuns?.at(-1);
      if(measuredRun?.name===(image.name || image.id)) {measuredRun.left=box.left;measuredRun.top=box.top;}
      if(textLayer) ctx.drawImage(textLayer,-box.width/2,-box.height/2,box.width,box.height);
      ctx.restore();
    }
    if (focus) {
      if (textLayout.checkedCells!==1) throw new Error('対象セルの実測結果がありません。');
      const scale=width<=512 && height<=256 ? 2 : 1;
      const pixels=document.createElement('canvas');
      pixels.width=width*scale; pixels.height=height*scale;
      const detailCtx=pixels.getContext('2d');
      if (!detailCtx) throw new Error('対象セルの描画結果を取得できません。');
      detailCtx.imageSmoothingEnabled=false;
      detailCtx.drawImage(canvas,0,0,pixels.width,pixels.height);
      results.push({sheetName:name,kind:'sheet',name:`render-${results.length+1}.png`,mime:'image/png',
        width:pixels.width,height:pixels.height,dataUrl:pixels.toDataURL('image/png'),details:[],textLayout,
        cellFocus:{row:focus.row,col:focus.col,value:getCellRawInput(sheet,focus.row,focus.col), displayValue:getDisplayForCell(sheet,id,focus.row,focus.col),
          bounds:{left:(target.left-left)*scale,top:(target.top-top)*scale,width:target.width*scale,height:target.height*scale}}});
      continue;
    }
    appendCwsInlineTextSpacingIssues(textLayout);
    const details=[];
    for (const crop of cwsAgentVisualDetailRegions(width,height,(sheet.images || []).some(image=>!image.hidden && image.cwsLayout?.role==='canvas'))) {
      const detail=document.createElement('canvas');detail.width=crop.width;detail.height=crop.height;
      detail.getContext('2d').drawImage(canvas,crop.left,crop.top,crop.width,crop.height,0,0,crop.width,crop.height);
      details.push({sheetName:name,kind:'sheet_detail',crop,name:`sheet-detail-${results.length+1}-${details.length+1}.png`,mime:'image/png',width:crop.width,height:crop.height,dataUrl:detail.toDataURL('image/png')});
    }
    for(const image of sheet.images || []) {
      if(image.hidden || image.cwsLayout?.role!=='canvas') continue;
      const box=imageBoxToPixels(sheet,image);if(!box) continue;
      const detail=document.createElement('canvas');detail.width=Math.ceil(box.width);detail.height=Math.ceil(box.height);
      detail.getContext('2d').drawImage(canvas,box.left-left,box.top-top,box.width,box.height,0,0,detail.width,detail.height);
      const componentBounds=(sheet.images || []).filter(part=>!part.hidden).flatMap(part=>{
        const partBox=imageBoxToPixels(sheet,part);if(!partBox || !part.name) return [];
        const angle=Number(part.rotation || part.shape?.rotation || 0)*Math.PI/180;
        const w=Math.abs(partBox.width*Math.cos(angle))+Math.abs(partBox.height*Math.sin(angle));
        const h=Math.abs(partBox.width*Math.sin(angle))+Math.abs(partBox.height*Math.cos(angle));
        return [{name:part.name,left:(partBox.left+partBox.width/2-w/2-box.left)*detail.width/box.width,
          top:(partBox.top+partBox.height/2-h/2-box.top)*detail.height/box.height,width:w*detail.width/box.width,height:h*detail.height/box.height}];
      });
      details.push({sheetName:name,kind:'wireframe_detail',name:`detail-${results.length+1}.png`,mime:'image/png',width:detail.width,height:detail.height,componentBounds,dataUrl:detail.toDataURL('image/png')});
    }
    results.push({sheetName:name,kind:'sheet',name:`render-${results.length+1}.png`,mime:'image/png',width,height,dataUrl:canvas.toDataURL('image/png'),details,textLayout});
  }
  return results;
}



function appendCwsInlineTextSpacingIssues(report) {
  const groups=new Map();
  for(const run of report.inlineRuns || []) {
    const key=JSON.stringify([run.parent,run.top,run.font]);
    if(!groups.has(key))groups.set(key,[]);
    groups.get(key).push(run);
  }
  delete report.inlineRuns;
  for(const runs of groups.values()) {
    if(runs.length<2)continue;
    runs.sort((a,b)=>a.left-b.left);
    let next=runs[0].left, mismatch=false;
    for(const run of runs) {
      if(Math.abs(run.left-next)>1.5)mismatch=true;
      next+=run.advanceWidth;
    }
    // Tighten only generated inline runs, never an intentionally wider source
    // region or overflowing text requiring a new line-break decision.
    const right=Math.max(...runs.map(r=>r.left+r.width));
    if(mismatch && next<=right+.75)report.issues.push({kind:'inline',name:runs[0].parent,
      text:runs.map(r=>r.text).join(''),runs,reason:'Generated inline runs have gaps inconsistent with actual browser text advance.'});
  }
}

function measureShapeTextLayout(host,width,height) {
  const layer=host.querySelector('.sheet-shape-text-layer'),block=host.querySelector('.sheet-shape-text-block');
  if(!layer || !block) throw new Error('Text layout unavailable');
  const css=getComputedStyle(layer),bounds=layer.getBoundingClientRect();
  const padding={left:parseFloat(css.paddingLeft)||0,right:parseFloat(css.paddingRight)||0,top:parseFloat(css.paddingTop)||0,bottom:parseFloat(css.paddingBottom)||0};
  const range=document.createRange(),rectangles=[];
  // pre-wrap deliberately hangs trailing spaces outside the line box. They
  // contain no ink and must not turn a readable caption into a clipping error.
  const walker=document.createTreeWalker(block,NodeFilter.SHOW_TEXT);
  for(let node=walker.nextNode();node;node=walker.nextNode()) {
    for(const match of node.textContent.matchAll(/\S+/gu)) {
      range.setStart(node,match.index);range.setEnd(node,match.index+match[0].length);
      rectangles.push(...Array.from(range.getClientRects()).filter(r=>r.width>0 && r.height>0));
    }
  }
  range.selectNodeContents(block);
  const advanceRects=Array.from(range.getClientRects()).filter(r=>r.width>0 && r.height>0);
  const singleLine=advanceRects.length>0 && advanceRects.every(r=>Math.abs(r.top-advanceRects[0].top)<.75);
  const advanceWidth=singleLine ? Math.max(...advanceRects.map(r=>r.right))-Math.min(...advanceRects.map(r=>r.left)) : null;
  const requiredHeight=block.getBoundingClientRect().height+padding.top+padding.bottom;
  // overflow:hidden clips at the padding box, not the inner content box.
  // Fallback glyphs may extend into padding without being cut off.
  const overflowSides={left:0,right:0,top:0,bottom:0};
  for(const rect of rectangles) {
    overflowSides.left=Math.max(overflowSides.left,bounds.left-rect.left);
    overflowSides.right=Math.max(overflowSides.right,rect.right-bounds.right);
    overflowSides.top=Math.max(overflowSides.top,bounds.top-rect.top);
    overflowSides.bottom=Math.max(overflowSides.bottom,rect.bottom-bounds.bottom);
  }
  const overflow=Object.values(overflowSides).some(value=>value>.75);
  const fits=!overflow && requiredHeight<=height+.75;
  let preferredWidth=null;
  if (!fits) {
    const natural=block.cloneNode(true);
    Object.assign(natural.style,{position:"absolute",width:"max-content",maxWidth:"none",whiteSpace:"pre",overflowWrap:"normal",wordBreak:"normal",visibility:"hidden"});
    layer.appendChild(natural);
    try {
      preferredWidth=Math.ceil(natural.getBoundingClientRect().width + padding.left + padding.right) + 1;
    } finally {
      natural.remove();
    }
  }
  return {fits,width,height,preferredWidth,requiredHeight:Math.ceil(requiredHeight*100)/100,
    recommendedHeight:Math.ceil(Math.max(requiredHeight,height+overflowSides.top+overflowSides.bottom))+1,
    overflowSides,font:css.font,padding,singleLine,advanceWidth};
}



function measureCwsAgentCellTextLayout(sheet,maps,sheetId,measure,report,focus=null) {
  const entries=focus ? [[focus.row+':'+focus.col,sheet.cells?.[focus.row+':'+focus.col]]] : Object.entries(sheet.cells || {});
  for(const [key,cell] of entries) {
    if (!cell) continue;
    if(maps.covered.has(key) || sheet.rows?.[cell.row-1]?.hidden || sheet.columns?.[cell.col-1]?.hidden) continue;
    const text=autoFitCellDisplayText(sheet,{sheetId},cell.row,cell.col,cell);
    if(!text) continue;
    report.checkedCells++;
    const merge=maps.master.get(key),rect=rangePixelRect(sheet,merge || {top:cell.row,bottom:cell.row,left:cell.col,right:cell.col});
    const css=cellEffectiveCss(sheet,cell);
    const requiredHeight=autoFitCellTextHeight(sheet,{sheetId,measure},cell.row,cell.col,cell,merge)-4+(['top','bottom'].includes(css.verticalAlign)?2:0);
    if(requiredHeight>rect.height+.75) report.issues.push({kind:'cell',name:'R'+cell.row+'C'+cell.col,text,width:rect.width,height:rect.height,requiredHeight});
  }
}



async function renderedShapeTextEvidence(image,width,height,textLayout = null) {
  const html=shapeTextLayerHtml(image,1);
  if(!html) return null;
  // Measure the same HTML layout used on screen and by SVG export, including
  // point-to-pixel conversion, rich text, padding and clipping.
  const host=document.createElement('div');
  host.style.cssText=`position:fixed;left:-100000px;top:0;width:${width}px;height:${height}px;pointer-events:none`;
  host.innerHTML=html;document.body.appendChild(host);
  try {
    const layout=measureShapeTextLayout(host,width,height);
    if(textLayout) {
      textLayout.checkedShapes++;
      const parent=image.cwsLayout?.parentShape || '', name=image.name || image.id;
      if(parent && name.startsWith(parent+'-text-') && /^\d+$/.test(name.slice(parent.length+6))
          && layout.fits && layout.singleLine && layout.advanceWidth>0 && Object.values(layout.padding).every(v=>v===0)
          && (image.shape?.textAlign || 'left')==='left' && !Number(image.rotation || image.shape?.rotation || 0)) {
        (textLayout.inlineRuns ||= []).push({name,parent,text:shapeTextPlainValue(image.shape),width,height,
          advanceWidth:Math.ceil((layout.advanceWidth+.25)*4)/4,font:layout.font});
      }

      if(!layout.fits) textLayout.issues.push({kind:'shape',name:image.name || image.id,text:shapeTextPlainValue(image.shape),...layout});
    }
    const clone=host.cloneNode(true);
    const properties=['position','inset','box-sizing','display','flex-direction','width','height','min-width','min-height','flex','padding','overflow','line-height','white-space','overflow-wrap','word-break','color','font-family','font-size','font-weight','font-style','text-align','justify-content','text-decoration-line','text-decoration-style','-webkit-text-stroke','transform','transform-origin'];
    const originals=host.querySelectorAll('*'),copies=clone.querySelectorAll('*');
    originals.forEach((node,index)=>{
      const css=getComputedStyle(node);
      copies[index].removeAttribute('class');
      copies[index].style.cssText=properties.map(key=>`${key}:${css.getPropertyValue(key)}`).join(';');
    });
    clone.style.cssText=`position:relative;width:${width}px;height:${height}px`;
    clone.setAttribute('xmlns','http://www.w3.org/1999/xhtml');
    const markup=new XMLSerializer().serializeToString(clone);
    const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}"><foreignObject width="100%" height="100%">${markup}</foreignObject></svg>`;
    return await loadImageElement(svgDataUrl(svg));
  } finally { host.remove(); }
}



function excelImageExportSource(image) {
  if (image?.chart?.series?.length && (!image.shape || image.shape.sourcePreset === "graphicFrame")) {
    const width = Math.max(1, Math.round(Number(image.ext?.width) || 520));
    const height = Math.max(1, Math.round(Number(image.ext?.height) || 320));
    return svgDataUrl(chartSvg(image.chart, image.chart.type || "column", { width, height, title: image.chart.title || image.name || "グラフ" }));
  }
  if (image?.shape) {
    const descriptor = shapeDescriptorForType(image.shape.type || DEFAULT_SHAPE_TYPE, {
      ...image.shape,
      width: image.ext?.width,
      height: image.ext?.height,
    });
    return descriptor.src || image.src;
  }
  return image?.src || "";
}



function excelDirectImagePayload(src) {
  const match = String(src || "").match(/^data:image\/(png|jpeg|jpg);base64,/i);
  if (!match) return null;
  const extension = match[1].toLowerCase() === "jpg" ? "jpeg" : match[1].toLowerCase();
  return { base64: src, extension };
}



async function rasterizedImagePayload(src, width, height) {
  if (!/^data:image\//i.test(String(src || ""))) return null;
  const image = await loadImageElement(src);
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(Number(width) || image.naturalWidth || image.width || 1));
  canvas.height = Math.max(1, Math.round(Number(height) || image.naturalHeight || image.height || 1));
  const context = canvas.getContext("2d");
  if (!context) return null;
  context.clearRect(0, 0, canvas.width, canvas.height);
  context.drawImage(image, 0, 0, canvas.width, canvas.height);
  return { base64: canvas.toDataURL("image/png"), extension: "png" };
}



function loadImageElement(src) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("画像を変換できません。"));
    image.src = src;
  });
}



function excelImagePlacement(sheet, image) {
  const placement = {
    tl: excelImageAnchor(sheet, image.tl),
    editAs: image.editAs || "oneCell",
  };
  if (image.br && (!image.ext?.width || !image.ext?.height)) {
    placement.br = excelImageAnchor(sheet, image.br);
  } else {
    placement.ext = {
      width: Math.max(1, Math.round(Number(image.ext?.width) || 1)),
      height: Math.max(1, Math.round(Number(image.ext?.height) || 1)),
    };
  }
  if (image.hyperlink) {
    placement.hyperlinks = {
      hyperlink: String(image.hyperlink),
      tooltip: String(image.name || image.hyperlink),
    };
  }
  return placement;
}



function excelImageAnchor(sheet, anchor = {}) {
  return {
    col: excelImageAxisAnchor(sheet, "col", anchor),
    row: excelImageAxisAnchor(sheet, "row", anchor),
  };
}



function excelImageAxisAnchor(sheet, axis, anchor = {}) {
  const isColumn = axis === "col";
  const valueKey = isColumn ? "col" : "row";
  const nativeKey = isColumn ? "nativeCol" : "nativeRow";
  const offsetKey = isColumn ? "nativeColOff" : "nativeRowOff";
  const base = Number.isFinite(Number(anchor[nativeKey])) ? Number(anchor[nativeKey]) : Number(anchor[valueKey]) || 0;
  const offsetEmu = Number(anchor[offsetKey]) || 0;
  if (!offsetEmu) return Math.round(base * 10000) / 10000;
  const size = excelImageAxisSize(sheet, axis, base);
  const offsetPixels = (offsetEmu / EMU_PER_PIXEL) * (sheet?.displayScale || 1);
  return Math.round((base + offsetPixels / size) * 10000) / 10000;
}



function excelImageAxisSize(sheet, axis, index) {
  const isColumn = axis === "col";
  const items = isColumn ? sheet?.columns || [] : sheet?.rows || [];
  const fallback = (isColumn ? sheet?.defaultColumnWidth : sheet?.defaultRowHeight) || (isColumn ? columnWidthToPixels(8.43) : rowHeightToPixels(15));
  const item = items[Math.max(0, Math.floor(Number(index) || 0))];
  return Math.max(1, Number(item?.hidden ? 0 : item?.width || item?.height || fallback) || fallback || 1);
}



function buildLegacyExcelHtml(model) {
  const sheets = (model?.sheets || []).map((sheet) => legacyExcelSheetHtml(sheet)).join("");
  return `<!doctype html><html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="UTF-8"><style>table{border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt}td{box-sizing:border-box;padding:0 3px;white-space:pre-wrap;font-family:Arial,sans-serif;font-size:10pt;vertical-align:middle}.sheet-title{font:bold 13pt Arial,sans-serif;margin:12px 0 6px}</style></head><body>${sheets}</body></html>`;
}



function legacyExcelSheetHtml(sheet) {
  const sheetId = getSheetId(sheet.name);
  const mergeMaps = buildMergeMaps(sheet.merges || []);
  const rows = [];
  const rowCount = Math.max(1, Number(sheet.rowCount) || sheet.rows?.length || 1);
  const colCount = Math.max(1, Number(sheet.colCount) || sheet.columns?.length || 1);
  const colgroup = Array.from({ length: colCount }, (_item, index) => {
    const width = sheet.columns?.[index]?.width || sheet.defaultColumnWidth || columnWidthToPixels(8.43);
    return `<col style="width:${Math.max(1, Math.round(width))}px">`;
  }).join("");
  for (let row = 1; row <= rowCount; row += 1) {
    const cells = [];
    const rowHeight = sheet.rows?.[row - 1]?.height || sheet.defaultRowHeight || rowHeightToPixels(15);
    for (let col = 1; col <= colCount; col += 1) {
      const key = cellKey(row, col);
      if (mergeMaps.covered.has(key)) continue;
      const merge = mergeMaps.master.get(key);
      const rowSpan = merge ? merge.bottom - merge.top + 1 : 1;
      const colSpan = merge ? merge.right - merge.left + 1 : 1;
      const cell = sheet.cells?.[cellKey(row, col)] || {};
      const spanAttrs = `${rowSpan > 1 ? ` rowspan="${rowSpan}"` : ""}${colSpan > 1 ? ` colspan="${colSpan}"` : ""}`;
      cells.push(`<td${spanAttrs} style="${escapeAttr(legacyExcelCellStyle(sheet, cell, rowHeight))}">${escapeHtml(getDisplayForCell(sheet, sheetId, row, col) ?? cell.display ?? "")}</td>`);
    }
    rows.push(`<tr style="height:${Math.max(1, Math.round(rowHeight))}px">${cells.join("")}</tr>`);
  }
  return `<div class="sheet-title">${escapeHtml(sheet.name || "Sheet")}</div><table>${colgroup}${rows.join("")}</table>`;
}



function legacyExcelCellStyle(sheet, cell, rowHeight, renderScale = 1) {
  const css = cellEffectiveCss(sheet, cell);
  const scaleValue = printRenderScale(renderScale);
  const fontSize = scaleValue === 1
    ? (css.fontSize || "10pt")
    : `${roundCssNumber(canvasFontSize(css.fontSize || "10pt", scaleValue))}px`;
  const paddingX = roundCssNumber(3 * scaleValue);
  const style = {
    height: `${Math.max(1, Math.round(rowHeight))}px`,
    fontFamily: css.fontFamily || "Arial, sans-serif",
    fontSize,
    fontWeight: css.fontWeight || "400",
    fontStyle: css.fontStyle || "normal",
    textDecoration: css.textDecoration || "none",
    lineHeight: "1.2",
    padding: `0 ${paddingX}px`,
    color: css.color || "#000000",
    backgroundColor: css.backgroundColor || "transparent",
    textAlign: css.textAlign || "left",
    verticalAlign: css.verticalAlign === "bottom" ? "bottom" : css.verticalAlign === "top" ? "top" : "middle",
    whiteSpace: css.whiteSpace === "pre-wrap" ? "pre-wrap" : "nowrap",
  };
  CELL_BORDER_SIDES.forEach((side) => {
    const border = cell.borders?.[side];
    if (!border) return;
    const cssSide = side === "top" ? "borderTop" : side === "right" ? "borderRight" : side === "bottom" ? "borderBottom" : "borderLeft";
    style[cssSide] = `${roundCssNumber(Math.max(0.5, (Number(border.weight || border.width) || 1) * scaleValue))}px ${border.style || "solid"} ${border.color || "#000000"}`;
  });
  return styleObjectToString(style);
}



function pixelsToExcelColumnWidth(pixels) {
  const value = (Number(pixels) - 5) / 7;
  return Math.max(0.1, Math.round((Number.isFinite(value) ? value : 8.43) * 100) / 100);
}



function pixelsToExcelRowHeight(pixels) {
  const value = Number(pixels) * 72 / 96;
  return Math.max(0.1, Math.round((Number.isFinite(value) ? value : 15) * 100) / 100);
}



function parseCssFontSizePt(value) {
  const text = String(value || "").trim();
  const number = Number.parseFloat(text);
  if (!Number.isFinite(number)) return null;
  if (text.endsWith("px")) return Math.round((number * 72 / 96) * 10) / 10;
  return Math.round(number * 10) / 10;
}



function cssColorToExcelArgb(value) {
  const text = String(value || "").trim();
  const hex = text.match(/^#([0-9a-f]{6})$/i);
  if (hex) return `FF${hex[1].toUpperCase()}`;
  const rgb = text.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*(\d*\.?\d+))?\s*\)$/i);
  if (!rgb) return null;
  const alpha = rgb[4] == null ? 1 : Math.max(0, Math.min(1, Number(rgb[4])));
  const channel = (index) => Math.max(0, Math.min(255, Number(rgb[index]) || 0)).toString(16).padStart(2, "0").toUpperCase();
  const alphaHex = Math.round(alpha * 255).toString(16).padStart(2, "0").toUpperCase();
  return `${alphaHex}${channel(1)}${channel(2)}${channel(3)}`;
}



async function buildStandaloneHtml(model) {
  const versionInfo = await standaloneRuntimeVersionInfo();
  const runtimeVersion = versionInfoRuntimeVersion(versionInfo) || STANDALONE_RUNTIME_VERSION;
  const runtimeBaseUrl = standaloneRuntimeBaseUrl();
  const versionedRuntimeBaseUrl = standaloneRuntimeVersionedBaseUrl(runtimeVersion);
  const remoteRuntimeBaseUrl = standaloneRemoteRuntimeBaseUrl();
  const versionedRemoteRuntimeBaseUrl = standaloneVersionedRemoteRuntimeBaseUrl(runtimeVersion);
  const runtimeUpdatedAt = versionInfoRuntimeUpdatedAt(versionInfo) || await standaloneRuntimeRequiredUpdatedAt();
  const modelPayload = await standaloneModelScriptPayload(model);
  return `<!doctype html>
<html ${cwsHtmlRootAttributes()}>
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
${cwsHtmlMetaTags(versionInfo)}
<link rel="icon" type="image/png" href="${escapeAttr(standaloneDocumentIconUrl())}" />
<title>${escapeHtml(model.sourceName)} - WebSheet</title>
</head>
<body>
<div id="websheetStandaloneBoot" style="font:14px system-ui,sans-serif;padding:16px;color:#1f2937">
  <strong>${escapeHtml(model.sourceName)}</strong>
  <div>WebSheet 共通ランタイムを読み込み中...</div>
</div>
${cwsAiNoticeComment()}
${cwsModelScriptTag(safeJson(modelPayload))}
${cwsAiInstructionsScript()}
<script>
${standaloneLoaderScript({
  version: runtimeVersion,
  runtimeUpdatedAt,
  localBaseUrl: runtimeBaseUrl,
  versionedLocalBaseUrl: versionedRuntimeBaseUrl,
  remoteBaseUrl: remoteRuntimeBaseUrl,
  versionedRemoteBaseUrl: versionedRemoteRuntimeBaseUrl,
  localSaveEndpoint: standaloneLocalSaveEndpointUrl(),
})}
</script>
</body>
</html>`;
}



async function standaloneRuntimeRequiredUpdatedAt() {
  const direct = [
    versionInfoRuntimeUpdatedAt(window.__WEBSHEET_STANDALONE_RUNTIME_VERSION__),
    versionInfoRuntimeUpdatedAt(window.__WEBSHEET_LOCAL_RUNTIME_INFO__),
    state.serverVersionInfo?.runtimeUpdatedAt,
    state.serverVersionInfo?.updatedAt,
  ].map((value) => String(value || "").trim()).find(Boolean);
  if (direct) return direct;
  try {
    const response = await fetch(`${serverVersionInfoUrl()}?t=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) return "";
    const info = normalizeVersionInfo(await response.json());
    state.serverVersionInfo = info;
    state.serverVersionStatus = "loaded";
    state.versionMismatch = versionInfoIsMismatch(info);
    syncVersionRibbonBadge();
    renderVersionDialogContent();
    return info.runtimeUpdatedAt || info.updatedAt || "";
  } catch {
    return "";
  }
}



async function standaloneRuntimeVersionInfo() {
  const source = runtimeExecutionSource();
  if (source === "local" && window.__WEBSHEET_LOCAL_RUNTIME_INFO__) {
    return normalizeVersionInfo(window.__WEBSHEET_LOCAL_RUNTIME_INFO__);
  }
  if (state.serverVersionInfo) return state.serverVersionInfo;
  const serverInfo = await refreshServerVersionInfo();
  if (serverInfo) return serverInfo;
  const runtimeInfo = window.__WEBSHEET_STANDALONE_RUNTIME_VERSION__ || window.__WEBSHEET_LOCAL_RUNTIME_INFO__;
  return normalizeVersionInfo(runtimeInfo || {});
}



function standaloneRuntimeBaseUrl() {
  try {
    const customUrl = localStorage.getItem(STANDALONE_RUNTIME_BASE_URL_STORAGE_KEY);
    if (customUrl) return normalizeRuntimeBaseUrl(customUrl);
  } catch {
    // Ignore localStorage failures and use the build-time default runtime path.
  }
  try {
    const configuredUrl = window.__WEBSHEET_STANDALONE_CONFIG__?.localBaseUrl;
    if (configuredUrl) return normalizeRuntimeBaseUrl(configuredUrl);
  } catch {
    // Ignore standalone config failures and use the platform default.
  }
  return normalizeRuntimeBaseUrl(defaultStandaloneRuntimeBaseUrl());
}



function defaultStandaloneRuntimeBaseUrl() {
  const configured = String(DEFAULT_STANDALONE_RUNTIME_BASE_URL || "").trim();
  if (configured && configured !== AUTO_STANDALONE_RUNTIME_BASE_URL) return configured;
  const platform = String(navigator.userAgentData?.platform || navigator.platform || "").toLowerCase();
  if (platform.includes("win")) return `file:///C:/ProgramData/WebSheet/runtime/${STANDALONE_RUNTIME_CHANNEL}`;
  if (platform.includes("mac")) return `file:///Library/Application%20Support/WebSheet/runtime/${STANDALONE_RUNTIME_CHANNEL}`;
  if (platform.includes("linux")) return `file:///usr/local/share/WebSheet/runtime/${STANDALONE_RUNTIME_CHANNEL}`;
  return `./websheet-runtime/${STANDALONE_RUNTIME_CHANNEL}`;
}



function standaloneRuntimeVersionedBaseUrl(version = currentRuntimeVersion()) {
  const current = standaloneRuntimeBaseUrl();
  return runtimeVersionedBaseUrlFromCurrent(current, version);
}



function runtimeVersionedBaseUrlFromCurrent(baseUrl, version = currentRuntimeVersion()) {
  const normalized = normalizeRuntimeBaseUrl(baseUrl);
  if (new RegExp(`/${STANDALONE_RUNTIME_CHANNEL}$`).test(normalized)) {
    return normalized.replace(new RegExp(`/${STANDALONE_RUNTIME_CHANNEL}$`), `/${version}`);
  }
  if (new RegExp(`/${STANDALONE_RUNTIME_CHANNEL.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}$`).test(normalized)) {
    return normalized.replace(new RegExp(`/${STANDALONE_RUNTIME_CHANNEL.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}$`), `/${version}`);
  }
  return normalized;
}



function standaloneRemoteRuntimeBaseUrl() {
  try {
    const configuredUrl = window.__WEBSHEET_STANDALONE_CONFIG__?.remoteBaseUrl;
    if (isRemoteRuntimeBaseUrl(configuredUrl)) return normalizeRuntimeBaseUrl(configuredUrl);
    const customUrl = localStorage.getItem(STANDALONE_REMOTE_RUNTIME_BASE_URL_STORAGE_KEY);
    if (isRemoteRuntimeBaseUrl(customUrl)) return normalizeRuntimeBaseUrl(customUrl);
  } catch {
    // Ignore localStorage failures and use the current WebSheet origin.
  }
  const hostedUrl = appHostedAssetUrl(`download/runtime/${STANDALONE_RUNTIME_CHANNEL}`);
  return normalizeRuntimeBaseUrl(
    isRemoteRuntimeBaseUrl(hostedUrl)
      ? hostedUrl
      : DEFAULT_STANDALONE_REMOTE_RUNTIME_BASE_URL,
  );
}



function standaloneVersionedRemoteRuntimeBaseUrl(version = currentRuntimeVersion()) {
  const runtimeVersion = String(version || STANDALONE_RUNTIME_VERSION).trim() || STANDALONE_RUNTIME_VERSION;
  const hostedUrl = appHostedAssetUrl(`download/runtime/${runtimeVersion}`);
  if (isRemoteRuntimeBaseUrl(hostedUrl)) return normalizeRuntimeBaseUrl(hostedUrl);
  return `https://chtec.co.jp/cws/download/runtime/${runtimeVersion}`;
}



function standaloneLocalSaveEndpointUrl() {
  try {
    if (/^https?:$/i.test(window.location.protocol) && /^(127\.0\.0\.1|localhost)$/i.test(window.location.hostname || "")) {
      return new URL("/__websheet/local-save", window.location.origin).href;
    }
  } catch {
    // Ignore URL construction failures and try the runtime source.
  }
  try {
    const remoteUrl = new URL(standaloneRemoteRuntimeBaseUrl());
    if (/^(127\.0\.0\.1|localhost)$/i.test(remoteUrl.hostname || "")) {
      return new URL("/__websheet/local-save", remoteUrl.origin).href;
    }
  } catch {
    // A local save bridge is optional; standalone HTML can still fall back to browser saving.
  }
  return "";
}



function normalizeRuntimeBaseUrl(value) {
  return String(value || "").trim().replace(/\/+$/, "");
}



function uniqueRuntimeBaseUrls(values) {
  const result = [];
  values.forEach((value) => {
    const normalized = normalizeRuntimeBaseUrl(value);
    if (normalized && !result.includes(normalized)) result.push(normalized);
  });
  return result;
}



function isRemoteRuntimeBaseUrl(value) {
  try {
    return /^https?:$/i.test(new URL(String(value || ""), window.location.href).protocol);
  } catch {
    return false;
  }
}



function appHostedAssetUrl(relativePath = "") {
  const normalized = String(relativePath || "").replace(/^\/+/, "");
  const baseUrl = appHostedBaseUrl();
  try {
    return new URL(normalized, baseUrl).href;
  } catch {
    return normalized;
  }
}



function appHostedBaseUrl() {
  if (typeof document !== "undefined") {
    const scriptUrls = [
      document.currentScript?.src || "",
      ...Array.from(document.scripts || []).map((script) => script.src || "").reverse(),
    ].filter(Boolean);
    for (const scriptUrl of scriptUrls) {
      const baseUrl = appBaseUrlFromScriptUrl(scriptUrl);
      if (baseUrl) return baseUrl;
    }
  }
  try {
    return new URL("./", document.baseURI || window.location.href).href;
  } catch {
    return "./";
  }
}



function appBaseUrlFromScriptUrl(scriptUrl) {
  try {
    const url = new URL(scriptUrl, document.baseURI || window.location.href);
    const path = url.pathname || "";
    const assetIndex = path.lastIndexOf("/assets/");
    if (assetIndex >= 0) {
      url.pathname = `${path.slice(0, assetIndex)}/`;
      url.search = "";
      url.hash = "";
      return url.href;
    }
    const runtimeIndex = path.lastIndexOf("/download/runtime/");
    if (runtimeIndex >= 0) {
      url.pathname = `${path.slice(0, runtimeIndex)}/`;
      url.search = "";
      url.hash = "";
      return url.href;
    }
  } catch {
    // Ignore script URLs that cannot be parsed.
  }
  return "";
}



function downloadAssetUrlFromRuntimeBase(relativePath, runtimeBaseUrl) {
  const normalized = String(relativePath || "").replace(/^\/+/, "");
  try {
    const url = new URL(runtimeBaseUrl, window.location.href);
    const path = url.pathname.replace(/\/+$/, "");
    const downloadRuntimeIndex = path.lastIndexOf("/download/runtime/");
    const runtimeIndex = path.lastIndexOf("/runtime/");
    url.pathname = downloadRuntimeIndex >= 0
      ? `${path.slice(0, downloadRuntimeIndex)}/${normalized}`
      : runtimeIndex >= 0
        ? `${path.slice(0, runtimeIndex)}/${normalized}`
      : `/${normalized}`;
    url.search = "";
    url.hash = "";
    return url.href;
  } catch {
    return appHostedAssetUrl(normalized);
  }
}



function standaloneRuntimeAssetUrl(fileName, runtimeBaseUrl = standaloneRuntimeBaseUrl()) {
  return `${runtimeBaseUrl}/${fileName}`;
}



function standaloneRuntimeIconUrl(runtimeBaseUrl = standaloneRuntimeBaseUrl()) {
  const url = standaloneRuntimeAssetUrl("icon.png", runtimeBaseUrl);
  return `${url}${url.includes("?") ? "&" : "?"}v=${encodeURIComponent(APP_ICON_VERSION)}`;
}



function standaloneDocumentIconUrl(runtimeBaseUrl = standaloneRuntimeBaseUrl()) {
  return APP_ICON_DATA_URL || standaloneRuntimeIconUrl(runtimeBaseUrl);
}



function standaloneLoaderScript(config) {
  const json = safeJson(config);
  return `(function(){
  "use strict";
  var config = ${json};
  var remoteRuntimeCacheBuster = Date.now().toString(36);
  window.__WEBSHEET_STANDALONE_CONFIG__ = config;
  window.__WEBSHEET_LOCAL_SAVE_ENDPOINT__ = config.localSaveEndpoint || "";
  function baseUrl(value){ return String(value || "").replace(/\\/+$/, ""); }
  function asset(base, file){ return baseUrl(base) + "/" + file; }
  function runtimeAsset(base, file, source){
    var url = asset(base, file);
    if (source !== "remote") return url;
    return url + (url.indexOf("?") >= 0 ? "&" : "?") + "v=" + encodeURIComponent(config.version || "") + "&b=" + remoteRuntimeCacheBuster;
  }
  function status(message){
    var boot = document.getElementById("websheetStandaloneBoot");
    var statusNode = document.getElementById("statusText");
    if (boot) {
      var detail = boot.querySelector("div") || boot;
      detail.textContent = message;
    }
    if (statusNode) statusNode.textContent = message;
  }
  function loadCss(href){
    return new Promise(function(resolve, reject){
      var link = document.createElement("link");
      link.rel = "stylesheet";
      link.href = href;
      link.onload = resolve;
      link.onerror = function(){ reject(new Error("スタイルシートを読み込めませんでした: " + href)); };
      document.head.appendChild(link);
    });
  }
  function loadScript(src){
    return new Promise(function(resolve, reject){
      var script = document.createElement("script");
      script.src = src;
      script.async = false;
      script.onload = resolve;
      script.onerror = function(){ reject(new Error("スクリプトを読み込めませんでした: " + src)); };
      document.head.appendChild(script);
    });
  }
  async function fetchRuntimeVersionJson(base){
    try {
      var response = await fetch(asset(base, "version.json") + "?t=" + Date.now(), { cache: "no-store" });
      if (!response.ok) return null;
      var info = await response.json();
      return info && typeof info === "object" ? info : null;
    } catch (error) {
      return null;
    }
  }
  function probeRuntimeVersion(base){
    return new Promise(function(resolve){
      if (!base) {
        resolve(null);
        return;
      }
      var done = false;
      var previousProbe = window.__WEBSHEET_RUNTIME_VERSION_PROBE__;
      var timer = setTimeout(function(){ finish(null); }, 1200);
      function finish(info){
        if (done) return;
        done = true;
        clearTimeout(timer);
        window.__WEBSHEET_RUNTIME_VERSION_PROBE__ = previousProbe;
        resolve(info || window.__WEBSHEET_STANDALONE_RUNTIME_VERSION__ || null);
      }
      window.__WEBSHEET_RUNTIME_VERSION_PROBE__ = finish;
      var script = document.createElement("script");
      script.src = asset(base, "version.js");
      script.async = true;
      script.onload = function(){ finish(window.__WEBSHEET_STANDALONE_RUNTIME_VERSION__ || null); };
      script.onerror = function(){ finish(null); };
      document.head.appendChild(script);
    });
  }
  async function readRuntimeVersion(base){
    return await fetchRuntimeVersionJson(base) || await probeRuntimeVersion(base);
  }
  function unique(values){
    var result = [];
    values.forEach(function(value){
      value = baseUrl(value);
      if (value && result.indexOf(value) < 0) result.push(value);
    });
    return result;
  }
  function inferredUserRuntimeBaseUrls(){
    var bases = [];
    try {
      var currentUrl = new URL(window.location.href);
      if (currentUrl.protocol !== "file:") return bases;
      var path = currentUrl.pathname || "";
      var version = encodeURIComponent(config.version || "");
      function add(root){
        bases.push(root + "/current");
        if (version) bases.push(root + "/" + version);
      }
      var macMatch = path.match(/^\\/Users\\/([^/]+)\\//);
      if (macMatch) add("file:///Users/" + macMatch[1] + "/Library/Application%20Support/WebSheet/runtime");
      var windowsMatch = path.match(/^\\/([A-Za-z]:)\\/Users\\/([^/]+)\\//);
      if (windowsMatch) add("file:///" + windowsMatch[1] + "/Users/" + windowsMatch[2] + "/AppData/Roaming/WebSheet/runtime");
      var linuxMatch = path.match(/^\\/home\\/([^/]+)\\//);
      if (linuxMatch) add("file:///home/" + linuxMatch[1] + "/.local/share/WebSheet/runtime");
    } catch {
      // If the current file URL cannot be parsed, keep the configured candidates.
    }
    return bases;
  }
  function localRuntimeCandidates(){
    return unique(inferredUserRuntimeBaseUrls().concat([config.localBaseUrl, config.versionedLocalBaseUrl]));
  }
  function versionParts(value){
    return String(value || "").split(/[.-]/).map(function(part){
      var number = parseInt(part, 10);
      return Number.isFinite(number) ? number : 0;
    });
  }
  function compareVersions(left, right){
    var a = versionParts(left);
    var b = versionParts(right);
    var length = Math.max(a.length, b.length, 3);
    for (var index = 0; index < length; index += 1) {
      var delta = (a[index] || 0) - (b[index] || 0);
      if (delta) return delta;
    }
    return 0;
  }
  function timestamp(value){
    var text = String(value || "").trim();
    if (!text) return 0;
    var time = Date.parse(text);
    return Number.isFinite(time) ? time : 0;
  }
  function runtimeIsUsable(info){
    var version = info && (info.runtimeVersion || info.version);
    if (!version) return false;
    var versionCompare = compareVersions(version, config.version);
    if (versionCompare < 0) return false;
    if (versionCompare > 0) return true;
    var requiredUpdatedAt = timestamp(config.runtimeUpdatedAt || config.updatedAt || "");
    if (!requiredUpdatedAt) return true;
    var actualUpdatedAt = timestamp(info.runtimeUpdatedAt || info.updatedAt || "");
    return !!actualUpdatedAt && actualUpdatedAt >= requiredUpdatedAt;
  }
  async function firstUsableLocalRuntime(){
    var firstInfo = null;
    var firstBase = "";
    var candidates = localRuntimeCandidates();
    window.__WEBSHEET_LOCAL_RUNTIME_CANDIDATES__ = candidates.slice();
    for (var index = 0; index < candidates.length; index += 1) {
      var base = candidates[index];
      var info = await readRuntimeVersion(base);
      if (info && !firstInfo) {
        firstInfo = info;
        firstBase = base;
      }
      if (runtimeIsUsable(info)) {
        return { base: base, info: info, status: "current" };
      }
    }
    return { base: firstBase, info: firstInfo, status: firstInfo ? "outdated" : "missing" };
  }
  async function loadRuntime(base, source){
    window.__WEBSHEET_RUNTIME_SOURCE__ = source;
    status(source === "local" ? "ローカル共通ランタイムを読み込み中..." : "共通ランタイムをダウンロード中...");
    await loadCss(runtimeAsset(base, "websheet-standalone.css", source));
    await loadScript(runtimeAsset(base, "websheet-standalone.js", source));
  }
  async function boot(){
    var localRuntime = await firstUsableLocalRuntime();
    window.__WEBSHEET_LOCAL_RUNTIME_INFO__ = localRuntime.info || null;
    window.__WEBSHEET_LOCAL_RUNTIME_STATUS__ = localRuntime.status;
    if (localRuntime.status === "current") {
      await loadRuntime(localRuntime.base, "local");
      return;
    }
    var remoteCandidates = unique([config.remoteBaseUrl, config.versionedRemoteBaseUrl]);
    var lastError = null;
    for (var index = 0; index < remoteCandidates.length; index += 1) {
      try {
        await loadRuntime(remoteCandidates[index], "remote");
        return;
      } catch (error) {
        lastError = error;
      }
    }
    throw lastError || new Error("共通ランタイムを読み込めませんでした。");
  }
  boot().catch(function(error){
    console.error(error);
    status("WebSheet 共通ランタイムを読み込めませんでした。");
    var errorBox = document.createElement("pre");
    errorBox.style.cssText = "margin:16px;padding:12px;border:1px solid #ef4444;background:#fff1f2;color:#991b1b;white-space:pre-wrap";
    errorBox.textContent = "WebSheet 共通ランタイムを読み込めませんでした。\\n" + (error && error.message ? error.message : String(error));
    document.body.appendChild(errorBox);
  });
})();`;
}

export {
  addXmlAttributesToTag,
  appBaseUrlFromScriptUrl,
  appHostedAssetUrl,
  appHostedBaseUrl,
  applyExcelExportCellNote,
  applyExcelExportCellStyle,
  applyExcelExportCellValidation,
  autoFilterExportValue,
  buildLegacyExcelHtml,
  buildStandaloneHtml,
  captureCwsAgentVisualEvidence,
  createExcelAutoFilterExportContext,
  cssColorToExcelArgb,
  customViewSheetSnapshotForSheet,
  cwsAgentVisualDetailRegions,
  dataValidationRectanglesFromCells,
  defaultStandaloneRuntimeBaseUrl,
  downloadAssetUrlFromRuntimeBase,
  escapeExcelCustomFilterLiteral,
  excelAlignmentFromCss,
  excelAutoFilterColorCriterionXml,
  excelAutoFilterColumnXml,
  excelAutoFilterCriterionXml,
  excelAutoFilterCustomCriterionXml,
  excelAutoFilterDateGroupItemXml,
  excelAutoFilterDateGroupsXml,
  excelAutoFilterOperatorXml,
  excelAutoFilterValuesXml,
  excelAutoFilterXml,
  excelAutoFilterXmlForState,
  excelBorderFromModel,
  excelBorderStyleFromModel,
  excelCellAddress,
  excelColorFilterDxfKey,
  excelColorFilterDxfXml,
  excelColorObjectFromCss,
  excelCommentThreadExportText,
  excelCustomFilterValue,
  excelCustomFilterXml,
  excelCustomSheetViewHeaderFooterTextXml,
  excelCustomSheetViewHeaderFooterXml,
  excelCustomSheetViewModeAttribute,
  excelCustomSheetViewPageMarginsXml,
  excelCustomSheetViewPageSetupXml,
  excelCustomSheetViewPaneXml,
  excelCustomSheetViewPrintOptionsXml,
  excelCustomSheetViewPrintSettingsXml,
  excelCustomSheetViewSelectionXml,
  excelCustomSheetViewTopLeftCellAttribute,
  excelCustomSheetViewXml,
  excelCustomSheetViewXmlEntries,
  excelCustomWorkbookViewXmlEntries,
  excelDataValidationFormula,
  excelDataValidationFormulaXmlText,
  excelDataValidationFormulae,
  excelDataValidationListFormula,
  excelDataValidationXml,
  excelDataValidationXmlDescriptor,
  excelDirectImagePayload,
  excelDynamicFilterType,
  excelDynamicFilterXml,
  excelExportCellValue,
  excelExportRichTextValue,
  excelExportScalarValue,
  excelFillFromCss,
  excelFontFromCss,
  excelImageAnchor,
  excelImageAxisAnchor,
  excelImageAxisSize,
  excelImageExportSource,
  excelImagePayload,
  excelImagePlacement,
  excelInternalCellHyperlinkMap,
  excelInternalHyperlinkTarget,
  excelModelTableForXml,
  excelModelTableLookup,
  excelNoteFromCell,
  excelPageCommentsMode,
  excelPageErrorsMode,
  excelPaperSizeCode,
  excelProtectionFromModel,
  excelSafeTableName,
  excelSheetSortStateForExport,
  excelSortConditionColorAttrs,
  excelSortConditionCustomListAttr,
  excelSortConditionXml,
  excelSortStateXml,
  excelSqrefFromCells,
  excelTableCellValue,
  excelTableColumnModel,
  excelTableStyleModel,
  excelTableTotalsFunction,
  excelWorkbookCalcPrXml,
  excelWorkbookDefinedNameXmlEntries,
  excelWorkbookPrXml,
  excelWorkbookPreservedDefinedNameXmlEntries,
  excelWorkbookViewXml,
  excelWorksheetDataValidationGroups,
  excelWorksheetDataValidationsXml,
  excelWorksheetPageBreaksXml,
  excelWorksheetViewSelectionRanges,
  excelWorksheetViewSelectionState,
  excelWorksheetViewSelectionXml,
  exportSheetCellsToExcel,
  exportSheetDimensionsToExcel,
  exportSheetImagesToExcel,
  exportSheetMergesToExcel,
  exportSheetTablesToExcel,
  finalizeExcelAutoFilterExportContext,
  hasPreservedVbaProject,
  insertWorksheetChildXml,
  isRemoteRuntimeBaseUrl,
  legacyExcelCellStyle,
  legacyExcelSheetHtml,
  loadImageElement,
  measureCwsAgentCellTextLayout,
  measureShapeTextLayout,
  nextWorkbookRelationshipId,
  normalizeOoxmlPartName,
  normalizeRuntimeBaseUrl,
  outlineCollapsedIndexSet,
  parseCssFontSizePt,
  patchExcelAutoFilterDxfStylesXml,
  patchExcelPreservedContentTypes,
  patchExcelPreservedOoxmlPackage,
  patchExcelPreservedWorkbookRelationships,
  patchExcelTableAutoFilterXml,
  patchExcelTableColumnFormulasXml,
  patchExcelTableDefinitions,
  patchExcelTableXml,
  patchExcelWorkbookCustomViews,
  patchExcelWorkbookDefinedNames,
  patchExcelWorkbookMacroContentType,
  patchExcelWorkbookMetadata,
  patchExcelWorkbookWindowViews,
  patchExcelWorksheetAutoFilterDetails,
  patchExcelWorksheetCustomViews,
  patchExcelWorksheetDataValidations,
  patchExcelWorksheetInternalHyperlinks,
  patchExcelWorksheetOutlineDetails,
  patchExcelWorksheetPageBreaks,
  patchExcelWorksheetViewDetails,
  patchWorksheetColumnCollapsedXml,
  patchWorksheetOutlineCollapsedXml,
  patchWorksheetOutlinePropertiesXml,
  patchWorksheetOutlineXml,
  patchWorksheetSheetViewSelectionDetails,
  patchXmlTagAttribute,
  pixelsToExcelColumnWidth,
  pixelsToExcelRowHeight,
  postProcessExcelWorkbookBuffer,
  rangeFromCellLabel,
  rasterizedImagePayload,
  removeWorksheetAutoFilterAndSortStateXml,
  removeWorksheetRelationships,
  renderedShapeTextEvidence,
  replaceExcelWorkbookDefinedNamesXml,
  replaceWorksheetDataValidationsXml,
  replaceWorksheetPageBreaksXml,
  runtimeVersionedBaseUrlFromCurrent,
  sanitizeWorksheetName,
  standaloneDocumentIconUrl,
  standaloneLoaderScript,
  standaloneLocalSaveEndpointUrl,
  standaloneRemoteRuntimeBaseUrl,
  standaloneRuntimeAssetUrl,
  standaloneRuntimeBaseUrl,
  standaloneRuntimeIconUrl,
  standaloneRuntimeRequiredUpdatedAt,
  standaloneRuntimeVersionInfo,
  standaloneRuntimeVersionedBaseUrl,
  standaloneVersionedRemoteRuntimeBaseUrl,
  uniqueExcelTableName,
  uniqueRuntimeBaseUrls,
  upsertTableColumnChildXml,
  upsertWorkbookBookViewsXml,
  upsertWorkbookCalcPrXml,
  upsertWorkbookCustomViewsXml,
  upsertWorkbookPrXml,
  upsertWorksheetCustomSheetViewsXml,
  worksheetColumnCollapsedSegments,
  xmlAttributeValue,
};
