// Sequential E2E feature module.
import {
  appUrl,
  richNativeShapeWorkbookPath,
} from "../../e2e-real-upload.mjs";
import {
  exerciseCustomViewImportExport,
} from "../helpers/formula-smoke.mjs";
import {
  exerciseAutoFilterConditionMenu,
  exerciseDataValidationMenu,
  exerciseHomeEditingCommands,
  exerciseImportedAutoFilterCriteria,
} from "../helpers/editing-and-validation.mjs";
import {
  exerciseAdvancedFilterCopyOverwrite,
  exerciseAdvancedFilterUniqueDates,
  exerciseAdvancedFilterUniqueSubsetCopy,
  exerciseAutoFilterBlankAndKeyboardMenu,
  exerciseAutoFilterDateMenu,
  exerciseAutoFilterReapplyAndClear,
  exerciseDelimitedTextImportOverwrite,
} from "../helpers/views-and-macros.mjs";
import {
  exerciseDataCurrentRegion,
  exerciseFlashFillDataPatterns,
  exercisePowerQueryTransformEditor,
  exerciseQueryConnectionsPane,
  exerciseScenarioManagerSummary,
  exerciseWhatIfGoalSeekDataTable,
} from "../helpers/query-transform.mjs";
import {
  exerciseConsolidateDialog,
  exerciseDataAnalysisToolpak,
  exerciseDataGetConnectionMenus,
  exerciseForecastSheet,
  exerciseOutlineGrouping,
  exerciseQuickSortActiveColumn,
  exerciseQuickSortPartialSelection,
  exerciseRefreshAllUndo,
  exerciseRelationshipsDialog,
  exerciseRemoveDuplicatesDialog,
  exerciseSortDialogBlankAndCase,
  exerciseSortDialogColorOrder,
  exerciseSortDialogCustomList,
  exerciseSortDialogDateOrder,
  exerciseSortDialogLeftToRight,
  exerciseSortDialogLevels,
  exerciseSortFormulaResults,
  exerciseSubtotalDialog,
  exerciseTextToColumnsAdvanced,
  exerciseTextToColumnsDialog,
  exerciseTextToColumnsFixedWidth,
} from "../helpers/connections-and-analysis.mjs";
async function runExportAndResults(scenarioState) {
  let { context, page, messages, result, dataRibbonLayout, coordinateHitTest, delayedLazyLoad, stickyHeaders, sheetMenu, sheetTabColor, sheetRename, sheetTabDirectActions, sheetMoveCopy, sheetCodeDialog, sheetSelectAll, sheetHidden, sheetUnhideDialog, sheetUnhidden, sheetUnhideUndo, sheetUnhideRedo, sheetGroupActions, sheetGroupMoveCopy, sheetGroupCellEditing, sheetGroupStructure, sheetGroupFormatDialog, sheetGroupDataValidation, sheetGroupDataFilterSort, sheetGroupDataTools, sheetGroupHyperlink, sheetGroupHomeCommands, sheetGroupTranslation, sheetGroupQuickAnalysis, sheetGroupSparklines, sheetGroupCharts, sheetGroupTables, sheetGroupPivot, sheetGroupObjects, insertRibbon, formulaRibbon, formulaDefineNameMain, formulaDefineNameMenu, formulaDefineNameCreate, createNamesDialog, formulaNames, formulaNameManager, formulaUseNameMenu, formulaUseNameManage, formulaUseNameInserted, formulaPasteNameDialog, formulaPasteNameInserted, formulaPasteNameListDialog, formulaPasteNameList, formulaPasteNameListUndo, formulaPasteNameListRedo, formulaFunctionDialog, formulaFunctionSearch, formulaInserted, formulaFunctionArguments, formulaLogicalMenu, formulaMenuArguments, formulaMenuInserted, formulaRecentMenu, formulaMathMenu, formulaVariableArguments, formulaZeroArgumentFunction, formulaFinancialMenu, formulaDateMenu, formulaTextMenu, formulaMoreMenu, formulaShowState, formulaTrace, formulaRemoveArrowsMenu, formulaRemovePrecedents, formulaEvaluateInitial, formulaEvaluateStepIn, formulaEvaluateStep, formulaEvaluateStepOut, formulaEvaluateReferenceInitial, formulaEvaluateReferenceStep, formulaEvaluateReferenceStepIn, formulaEvaluateReferenceStepOut, formulaWatchWindow, formulaWatchSelect, formulaWatchDelete, formulaErrorMenu, formulaErrorChecking, formulaErrorIgnore, formulaErrorTrace, formulaErrorResetIgnored, formulaCalculation, viewRibbon, viewToggles, viewWindowCommands, layoutRibbon, layoutThemeMenuInitial, layoutThemeApplied, layoutColorMenuInitial, layoutFontMenuInitial, layoutEffectMenuInitial, layoutThemeAfterMenus, layoutThemeEffectUndoRedo, layoutPrintAreaMenuBefore, layoutPrintAreaMenuAfterSet, layoutPrintAreaMenuAfterClear, layoutMarginsMenuInitial, layoutOrientationMenuInitial, layoutSizeMenuInitial, layoutBackgroundApplied, layoutBackgroundUndo, layoutBackgroundRedo, layoutBackgroundDeleted, layoutBackgroundDeleteUndo, pageSetupDialog, printTitlesDialogInitial, printTitlesDialogFilled, pageLayout, pageLayoutPreview, pageBreakMenuKeyboard, pageBreakDrag, drawRibbon, drawInk, drawConverted, headerFooterDialog, headerFooterApplied, headerFooterUndo, headerFooterRedo, insertMapChart, insertShapeMenu, insertShapeDraw, insertIconDialog, signatureLineDialog, insertObjectDialog, insertAddinsStoreDialog, insertAddinsMineDialog, insertTextObjects, fileSideMenu, fileSideMenuClosedByOutsideClick, fileSideMenuClosedByCommand, shortcutSave, exportedHtmlChecks, saveAsDialog, expectedReloadWorkbookName, standaloneBlockedRequests, standaloneOffline, standaloneRemoteFallback, standaloneEmptyModel, standaloneMalformedModel, standaloneWrappedModel, launcherBlockedRequests, standaloneLauncher, ar2Gridline, imageAndFont, imageSelection, footerLayout, realCtrlWheelZoom, zoomPopup, footerBeforeScroll, tabScroll, singleClickEdit, doubleClickEdit, editAfterDeferredRender, editAlignment, editClickCaret, editTextSelection, editNoEllipsis, cellEditAltEnter, cellEditAltEnterDuringIme, editDragOutsideSelection, cellEditImeEnter, editScrollbar, editExitRender, multilineCommitDisplay, formulaBarInitial, formulaBarSizing, formulaBarDuringEdit, formulaBarEscape, formulaBarEnter, formulaAutocomplete, formulaReferenceInsert, directTyping, definedNameAndLink, statusBarStats, autoFillSeries, insertDeleteDialog, shortcutUndo, formatUndoRedo, toolbarSync, selectionFrame, copyStateAfterCopy, shortcutClipboard, copyStateWithTarget, richPaste, matchDestinationPaste, pasteSpecialFormats, pasteSpecialAdvanced, formulaPasteReferences, cutAndRepeatPaste, cellMenu, cellAnnotations, furigana, formatCellsDialog, cellMenuClear, cellMenuUndo, deleteKeyClear, deleteKeyUndo, backspaceClearEdit, backspaceUndo, popupDismissal, newSheet, colorPalette, headerSelection, selectionAutoScroll, keyboardNavigation, pressWorkbookShortcut, pressCtrlShiftShortcut } = scenarioState;

  keyboardNavigation.afterHiddenColRight = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.evaluate(() => window.__webSheetDevSelectRange?.(70, 5, 70, 5));


  await page.keyboard.press("ArrowDown");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "E72", null, { timeout: 10000 });


  keyboardNavigation.afterHiddenRowDown = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.evaluate(() => {
    window.__webSheetDevSetHeaderHidden?.("column", 6, false);
    window.__webSheetDevSetHeaderHidden?.("row", 71, false);
  });


  keyboardNavigation.sheetSwitch = await page.evaluate(() => ({
    before: window.__webSheetDevWorkbookState?.().activeSheetName || "",
  }));


  await page.keyboard.press("Control+PageDown");


  await page.waitForFunction((before) => window.__webSheetDevWorkbookState?.().activeSheetName !== before, keyboardNavigation.sheetSwitch.before, { timeout: 10000 });


  keyboardNavigation.sheetSwitch.afterNext = await page.evaluate(() => window.__webSheetDevWorkbookState?.().activeSheetName || "");


  await page.keyboard.press("Control+PageUp");


  await page.waitForFunction((before) => window.__webSheetDevWorkbookState?.().activeSheetName === before, keyboardNavigation.sheetSwitch.before, { timeout: 10000 });


  keyboardNavigation.sheetSwitch.afterPrevious = await page.evaluate(() => window.__webSheetDevWorkbookState?.().activeSheetName || "");


  keyboardNavigation.sheetGroupShortcut = await page.evaluate(() => {
    const workbookState = window.__webSheetDevWorkbookState?.();
    const visibleIndexes = [...document.querySelectorAll(".sheet-tab")].map((tab) => Number(tab.dataset.sheetIndex));
    const currentPosition = visibleIndexes.indexOf(Number(workbookState?.activeSheetIndex));
    const previousPosition = (currentPosition - 1 + visibleIndexes.length) % visibleIndexes.length;
    const previousIndex = visibleIndexes[previousPosition];
    const nextIndex = visibleIndexes[(previousPosition + 1) % visibleIndexes.length];
    return {
      beforeIndex: Number(workbookState?.activeSheetIndex),
      beforeName: workbookState?.activeSheetName || "",
      visibleCount: visibleIndexes.length,
      expectedPreviousIndex: previousIndex,
      expectedNextIndex: nextIndex,
    };
  });


  if (keyboardNavigation.sheetGroupShortcut.visibleCount >= 2) {
    await page.keyboard.press("Control+Shift+PageUp");
    await page.waitForFunction(() => (window.__webSheetDevWorkbookState?.().selectedSheetIndexes || []).length >= 2, null, { timeout: 10000 });
    Object.assign(keyboardNavigation.sheetGroupShortcut, await page.evaluate(() => {
      const workbookState = window.__webSheetDevWorkbookState?.();
      return {
        afterPreviousActiveIndex: Number(workbookState?.activeSheetIndex),
        afterPreviousSelected: workbookState?.selectedSheetIndexes || [],
        afterPreviousStatus: document.querySelector("#statusText")?.textContent || "",
      };
    }));
    await page.keyboard.press("Control+Shift+PageDown");
    await page.waitForFunction(
      (expected) => Number(window.__webSheetDevWorkbookState?.().activeSheetIndex) === Number(expected),
      keyboardNavigation.sheetGroupShortcut.expectedNextIndex,
      { timeout: 10000 },
    );
    Object.assign(keyboardNavigation.sheetGroupShortcut, await page.evaluate(() => {
      const workbookState = window.__webSheetDevWorkbookState?.();
      return {
        afterNextActiveIndex: Number(workbookState?.activeSheetIndex),
        afterNextSelected: workbookState?.selectedSheetIndexes || [],
        afterNextStatus: document.querySelector("#statusText")?.textContent || "",
      };
    }));
    await page.evaluate((sheetName) => window.__webSheetDevSetActiveSheet?.(sheetName), keyboardNavigation.sheetGroupShortcut.beforeName);
    await page.waitForFunction(
      (beforeName) => window.__webSheetDevWorkbookState?.().activeSheetName === beforeName,
      keyboardNavigation.sheetGroupShortcut.beforeName,
      { timeout: 10000 },
    );
    Object.assign(keyboardNavigation.sheetGroupShortcut, await page.evaluate(() => {
      const workbookState = window.__webSheetDevWorkbookState?.();
      return {
        restoredName: workbookState?.activeSheetName || "",
        restoredSelectedCount: (workbookState?.selectedSheetIndexes || []).length,
      };
    }));
  }



  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(73, 8, "FormatShortcut");
    window.__webSheetDevSelectRange?.(73, 8, 73, 8);
  });


  await page.keyboard.press("Control+B");


  await page.keyboard.press("Control+I");


  await page.keyboard.press("Control+U");


  await page.keyboard.press("Control+5");


  keyboardNavigation.shortcutFormat = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    const css = window.__webSheetDevCellModel?.(sheetName, 73, 8)?.css || {};
    return {
      bold: css.fontWeight,
      italic: css.fontStyle,
      decoration: css.textDecoration,
    };
  });


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(79, 8, "FormatAliasShortcut");
    window.__webSheetDevSelectRange?.(79, 8, 79, 8);
  });


  await page.keyboard.press("Control+2");


  await page.keyboard.press("Control+3");


  await page.keyboard.press("Control+4");


  keyboardNavigation.formatAliasShortcuts = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    const css = window.__webSheetDevCellModel?.(sheetName, 79, 8)?.css || {};
    return {
      bold: css.fontWeight,
      italic: css.fontStyle,
      decoration: css.textDecoration,
    };
  });


  await page.keyboard.press("Control+1");


  await page.waitForSelector("#formatCellsDialogOverlay", { timeout: 10000 });


  keyboardNavigation.formatDialogTitle = await page.evaluate(() => document.querySelector("#formatCellsDialogOverlay .format-dialog-titlebar")?.textContent?.trim() || "");


  await page.click("#formatCellsDialogOverlay [data-format-dialog-action='cancel']");


  await page.evaluate(() => {
    for (let col = 21; col <= 27; col += 1) window.__webSheetDevSetCell?.(79, col, "1234.56");
  });


  for (const [keyName, col] of [
    ["Backquote", 21],
    ["Digit1", 22],
    ["Digit2", 23],
    ["Digit3", 24],
    ["Digit4", 25],
    ["Digit5", 26],
    ["Digit6", 27],
  ]) {
    await page.evaluate((targetCol) => window.__webSheetDevSelectRange?.(79, targetCol, 79, targetCol), col);
    await pressCtrlShiftShortcut(keyName);
  }


  keyboardNavigation.numberFormatShortcuts = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    const read = (col) => window.__webSheetDevCellModel?.(sheetName, 79, col)?.numFmt ?? null;
    return {
      general: read(21),
      number: read(22),
      time: read(23),
      date: read(24),
      currency: read(25),
      percent: read(26),
      scientific: read(27),
    };
  });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(80, 21, 81, 22));


  await pressCtrlShiftShortcut("Digit7");


  keyboardNavigation.borderShortcutApplied = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    const read = (row, col) => {
      const borders = window.__webSheetDevCellModel?.(sheetName, row, col)?.borders || {};
      return {
        top: Boolean(borders.top),
        right: Boolean(borders.right),
        bottom: Boolean(borders.bottom),
        left: Boolean(borders.left),
      };
    };
    return {
      topLeft: read(80, 21),
      topRight: read(80, 22),
      bottomLeft: read(81, 21),
      bottomRight: read(81, 22),
    };
  });


  await pressCtrlShiftShortcut("Minus");


  keyboardNavigation.borderShortcutRemoved = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    const read = (row, col) => {
      const borders = window.__webSheetDevCellModel?.(sheetName, row, col)?.borders || {};
      return {
        top: Boolean(borders.top),
        right: Boolean(borders.right),
        bottom: Boolean(borders.bottom),
        left: Boolean(borders.left),
      };
    };
    return {
      topLeft: read(80, 21),
      topRight: read(80, 22),
      bottomLeft: read(81, 21),
      bottomRight: read(81, 22),
    };
  });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(79, 24, 79, 24));


  await pressCtrlShiftShortcut("Equal");


  await page.waitForSelector("#cellStructureDialog", { timeout: 5000 });


  keyboardNavigation.insertShortcutDialogTitle = await page.evaluate(() => document.querySelector("#cellStructureDialogTitle")?.textContent || "");


  await page.click("#cellStructureDialog [data-cell-structure-action='cancel']");


  await page.waitForSelector("#cellStructureDialogOverlay", { state: "detached", timeout: 5000 });


  await pressWorkbookShortcut(["Control"], "Minus");


  await page.waitForSelector("#cellStructureDialog", { timeout: 5000 });


  keyboardNavigation.deleteShortcutDialogTitle = await page.evaluate(() => document.querySelector("#cellStructureDialogTitle")?.textContent || "");


  await page.click("#cellStructureDialog [data-cell-structure-action='cancel']");


  await page.waitForSelector("#cellStructureDialogOverlay", { state: "detached", timeout: 5000 });


  keyboardNavigation.formulaBarShortcut = await page.evaluate(() => ({
    beforeRows: Number(document.querySelector("#formulaInput")?.getAttribute("rows") || 0),
  }));


  await pressCtrlShiftShortcut("u");


  await page.waitForFunction((beforeRows) => Number(document.querySelector("#formulaInput")?.getAttribute("rows") || 0) !== beforeRows, keyboardNavigation.formulaBarShortcut.beforeRows, { timeout: 5000 });


  Object.assign(keyboardNavigation.formulaBarShortcut, await page.evaluate(() => ({
    afterRows: Number(document.querySelector("#formulaInput")?.getAttribute("rows") || 0),
  })));


  await pressCtrlShiftShortcut("u");


  await page.waitForFunction((beforeRows) => Number(document.querySelector("#formulaInput")?.getAttribute("rows") || 0) === beforeRows, keyboardNavigation.formulaBarShortcut.beforeRows, { timeout: 5000 });


  Object.assign(keyboardNavigation.formulaBarShortcut, await page.evaluate(() => ({
    restoredRows: Number(document.querySelector("#formulaInput")?.getAttribute("rows") || 0),
  })));


  await pressWorkbookShortcut(["Control"], "Digit6");


  keyboardNavigation.objectsShortcutHidden = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return window.__webSheetDevSheetModel?.(sheetName)?.showObjects;
  });


  await pressWorkbookShortcut(["Control"], "Digit6");


  keyboardNavigation.objectsShortcutVisible = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return window.__webSheetDevSheetModel?.(sheetName)?.showObjects;
  });


  await pressWorkbookShortcut(["Control"], "Digit8");


  keyboardNavigation.outlineShortcutHidden = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return window.__webSheetDevSheetModel?.(sheetName)?.showOutlineSymbols;
  });


  await pressWorkbookShortcut(["Control"], "Digit8");


  keyboardNavigation.outlineShortcutVisible = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return window.__webSheetDevSheetModel?.(sheetName)?.showOutlineSymbols;
  });


  await page.evaluate(() => {
    const values = [
      ["月", "売上", "利益"],
      ["4月", 120, 32],
      ["5月", 180, 48],
      ["6月", 150, 41],
    ];
    values.forEach((rowValues, rowOffset) => {
      rowValues.forEach((value, colOffset) => window.__webSheetDevSetCell?.(86 + rowOffset, 21 + colOffset, value));
    });
    window.__webSheetDevSelectRange?.(86, 21, 89, 23);
  });


  keyboardNavigation.altF1ChartBefore = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    const sheet = window.__webSheetDevSheetModel?.(sheetName) || {};
    return {
      sheetName,
      imageCount: (sheet.images || []).length,
    };
  });


  await pressWorkbookShortcut(["Alt"], "F1");


  await page.waitForFunction(
    (before) => {
      const sheet = window.__webSheetDevSheetModel?.(before.sheetName) || {};
      return (sheet.images || []).length === before.imageCount + 1;
    },
    keyboardNavigation.altF1ChartBefore,
    { timeout: 5000 },
  );


  keyboardNavigation.altF1Chart = await page.evaluate((before) => {
    const state = window.__webSheetDevWorkbookState?.() || {};
    const sheet = window.__webSheetDevSheetModel?.(before.sheetName) || {};
    const latest = (sheet.images || []).at(-1) || {};
    return {
      activeSheetName: state.activeSheetName || "",
      countDelta: (sheet.images || []).length - before.imageCount,
      chartType: latest.chart?.type || "",
      sourceRange: latest.chart?.sourceRange || "",
      selectedImageFrame: Boolean(document.querySelector(".sheet-image-selection-frame")),
    };
  }, keyboardNavigation.altF1ChartBefore);


  await page.evaluate(() => window.__webSheetDevSelectRange?.(86, 21, 89, 23));


  const f11Before = await page.evaluate(() => ({
    activeSheetName: window.__webSheetDevWorkbookState?.().activeSheetName || "",
    tabCount: document.querySelectorAll(".sheet-tab").length,
  }));


  await page.keyboard.press("F11");


  await page.waitForFunction((before) => document.querySelectorAll(".sheet-tab").length === before.tabCount + 1, f11Before, { timeout: 5000 });


  keyboardNavigation.f11ChartSheet = await page.evaluate((before) => {
    const state = window.__webSheetDevWorkbookState?.() || {};
    const sheet = window.__webSheetDevSheetModel?.(state.activeSheetName) || {};
    const chart = (sheet.images || [])[0]?.chart || {};
    return {
      previousSheetName: before.activeSheetName,
      activeSheetName: state.activeSheetName || "",
      tabCountDelta: document.querySelectorAll(".sheet-tab").length - before.tabCount,
      chartSheet: sheet.chartSheet === true,
      showGridLines: sheet.showGridLines,
      showHeadings: sheet.showHeadings,
      imageCount: (sheet.images || []).length,
      chartType: chart.type || "",
      sourceSheet: chart.sourceSheet || "",
      sourceRange: chart.sourceRange || "",
      selectedImageFrame: Boolean(document.querySelector(".sheet-image-selection-frame")),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, f11Before);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction((before) => document.querySelectorAll(".sheet-tab").length === before.tabCount, f11Before, { timeout: 5000 });


  keyboardNavigation.f11ChartSheetUndo = await page.evaluate((before) => ({
    activeSheetName: window.__webSheetDevWorkbookState?.().activeSheetName || "",
    tabCount: document.querySelectorAll(".sheet-tab").length,
    status: document.querySelector("#statusText")?.textContent || "",
    chartSheetGone: !window.__webSheetDevSheetModel?.("グラフ1"),
    expectedTabCount: before.tabCount,
  }), f11Before);


  const shiftF11Before = await page.evaluate(() => ({
    activeSheetName: window.__webSheetDevWorkbookState?.().activeSheetName || "",
    tabCount: document.querySelectorAll(".sheet-tab").length,
  }));


  await pressWorkbookShortcut(["Shift"], "F11");


  await page.waitForFunction((before) => document.querySelectorAll(".sheet-tab").length === before.tabCount + 1, shiftF11Before, { timeout: 5000 });


  keyboardNavigation.shiftF11Worksheet = await page.evaluate((before) => ({
    tabCountDelta: document.querySelectorAll(".sheet-tab").length - before.tabCount,
    activeSheetName: window.__webSheetDevWorkbookState?.().activeSheetName || "",
    selectedLabel: document.querySelector("#selectedCellLabel")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), shiftF11Before);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction((before) => document.querySelectorAll(".sheet-tab").length === before.tabCount, shiftF11Before, { timeout: 5000 });


  keyboardNavigation.shiftF11WorksheetUndo = await page.evaluate((before) => ({
    activeSheetName: window.__webSheetDevWorkbookState?.().activeSheetName || "",
    tabCount: document.querySelectorAll(".sheet-tab").length,
    expectedTabCount: before.tabCount,
    status: document.querySelector("#statusText")?.textContent || "",
  }), shiftF11Before);


  await page.click('[data-ribbon-tab="view"]');


  const windowShortcutBefore = await page.evaluate(() => ({
    activeWindowId: window.__webSheetDevWorkbookState?.().activeWindowId || "main",
    visibleWindowCount: 1 + (window.__webSheetDevWorkbookState?.().windowViews || []).filter((view) => !view.hidden).length,
  }));


  await page.click('[data-ribbon-command="view-new-window"]');


  await page.waitForFunction(
    (before) => 1 + (window.__webSheetDevWorkbookState?.().windowViews || []).filter((view) => !view.hidden).length >= before.visibleWindowCount + 1,
    windowShortcutBefore,
    { timeout: 5000 },
  );


  const windowShortcutReady = await page.evaluate(() => ({
    activeWindowId: window.__webSheetDevWorkbookState?.().activeWindowId || "main",
    visibleWindowCount: 1 + (window.__webSheetDevWorkbookState?.().windowViews || []).filter((view) => !view.hidden).length,
  }));


  await pressWorkbookShortcut(["Control"], "F6");


  await page.waitForFunction((before) => window.__webSheetDevWorkbookState?.().activeWindowId !== before.activeWindowId, windowShortcutReady, { timeout: 5000 });


  const windowAfterNext = await page.evaluate(() => ({
    activeWindowId: window.__webSheetDevWorkbookState?.().activeWindowId || "main",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await pressWorkbookShortcut(["Control", "Shift"], "F6");


  await page.waitForFunction((before) => window.__webSheetDevWorkbookState?.().activeWindowId === before.activeWindowId, windowShortcutReady, { timeout: 5000 });


  keyboardNavigation.ctrlF6WindowSwitch = await page.evaluate(({ ready, previousStep }) => ({
    ready,
    afterNext: window.__webSheetDevWorkbookState?.().activeWindowId || "main",
    previousStep,
    visibleWindowCount: 1 + (window.__webSheetDevWorkbookState?.().windowViews || []).filter((view) => !view.hidden).length,
    status: document.querySelector("#statusText")?.textContent || "",
  }), { ready: windowShortcutReady, previousStep: windowAfterNext });


  await pressWorkbookShortcut(["Control"], "F4");


  await page.waitForFunction(
    (before) => 1 + (window.__webSheetDevWorkbookState?.().windowViews || []).filter((view) => !view.hidden).length === before.visibleWindowCount - 1,
    windowShortcutReady,
    { timeout: 5000 },
  );


  keyboardNavigation.ctrlF4CloseWindow = await page.evaluate((before) => ({
    visibleWindowCount: 1 + (window.__webSheetDevWorkbookState?.().windowViews || []).filter((view) => !view.hidden).length,
    expectedCount: before.visibleWindowCount - 1,
    status: document.querySelector("#statusText")?.textContent || "",
  }), windowShortcutReady);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    (before) => 1 + (window.__webSheetDevWorkbookState?.().windowViews || []).filter((view) => !view.hidden).length === before.visibleWindowCount,
    windowShortcutReady,
    { timeout: 5000 },
  );


  keyboardNavigation.ctrlF4CloseWindowUndo = await page.evaluate((before) => ({
    visibleWindowCount: 1 + (window.__webSheetDevWorkbookState?.().windowViews || []).filter((view) => !view.hidden).length,
    expectedCount: before.visibleWindowCount,
    status: document.querySelector("#statusText")?.textContent || "",
  }), windowShortcutReady);


  await pressWorkbookShortcut(["Control"], "W");


  await page.waitForFunction(
    (before) => 1 + (window.__webSheetDevWorkbookState?.().windowViews || []).filter((view) => !view.hidden).length === before.visibleWindowCount - 1,
    windowShortcutReady,
    { timeout: 5000 },
  );


  keyboardNavigation.ctrlWCloseWindow = await page.evaluate((before) => ({
    visibleWindowCount: 1 + (window.__webSheetDevWorkbookState?.().windowViews || []).filter((view) => !view.hidden).length,
    expectedCount: before.visibleWindowCount - 1,
    status: document.querySelector("#statusText")?.textContent || "",
  }), windowShortcutReady);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    (before) => 1 + (window.__webSheetDevWorkbookState?.().windowViews || []).filter((view) => !view.hidden).length === before.visibleWindowCount,
    windowShortcutReady,
    { timeout: 5000 },
  );


  keyboardNavigation.ctrlWCloseWindowUndo = await page.evaluate((before) => ({
    visibleWindowCount: 1 + (window.__webSheetDevWorkbookState?.().windowViews || []).filter((view) => !view.hidden).length,
    expectedCount: before.visibleWindowCount,
    status: document.querySelector("#statusText")?.textContent || "",
  }), windowShortcutReady);


  await page.keyboard.press("Control+F1");


  await page.waitForFunction(() => document.querySelector(".ribbon")?.classList.contains("is-collapsed"), null, { timeout: 5000 });


  keyboardNavigation.ribbonCollapsed = await page.evaluate(() => ({
    collapsed: document.querySelector(".ribbon")?.classList.contains("is-collapsed") || false,
    panelsHidden: getComputedStyle(document.querySelector("#ribbonPanels")).display === "none",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+F1");


  await page.waitForFunction(() => !document.querySelector(".ribbon")?.classList.contains("is-collapsed"), null, { timeout: 5000 });


  keyboardNavigation.ribbonExpanded = await page.evaluate(() => ({
    collapsed: document.querySelector(".ribbon")?.classList.contains("is-collapsed") || false,
    panelsHidden: getComputedStyle(document.querySelector("#ribbonPanels")).display === "none",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await pressWorkbookShortcut(["Alt"], "F8");


  await page.waitForSelector('[data-data-dialog="macros"]', { timeout: 5000 });


  keyboardNavigation.altF8MacrosTitle = await page.evaluate(() => document.querySelector("#macrosTitle")?.textContent || "");


  await page.locator('[data-data-dialog="macros"] [data-data-dialog-action="cancel"]').click();


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  await pressWorkbookShortcut(["Alt"], "F11");


  await page.waitForSelector('[data-data-dialog="sheet-code"]', { timeout: 5000 });


  keyboardNavigation.altF11VbaTitle = await page.evaluate(() => document.querySelector("#sheetCodeTitle")?.textContent || "");


  await page.locator('[data-data-dialog="sheet-code"] [data-data-dialog-action="sheet-code-ok"]').click();


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  await page.keyboard.press("F9");


  keyboardNavigation.f9Calculate = await page.evaluate(() => document.querySelector("#statusText")?.textContent || "");


  await page.keyboard.press("Shift+F9");


  keyboardNavigation.shiftF9Calculate = await page.evaluate(() => document.querySelector("#statusText")?.textContent || "");


  await pressWorkbookShortcut(["Alt"], "A");


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent?.trim() === "データ", null, { timeout: 5000 });


  await pressWorkbookShortcut(["Alt"], "W");


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent?.trim() === "表示", null, { timeout: 5000 });


  await pressWorkbookShortcut(["Alt"], "H");


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent?.trim() === "ホーム", null, { timeout: 5000 });


  keyboardNavigation.ribbonAltTabs = await page.evaluate(() => ({
    active: document.querySelector(".ribbon-tab.is-active")?.textContent?.trim() || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("F10");


  keyboardNavigation.ribbonKeyboard = await page.evaluate(() => ({
    focusedAfterF10: document.activeElement?.dataset?.ribbonTab || "",
    statusAfterF10: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.locator('[data-ribbon-tab="home"]').focus();


  await page.keyboard.press("ArrowRight");


  Object.assign(keyboardNavigation.ribbonKeyboard, await page.evaluate(() => ({
    focusedAfterRight: document.activeElement?.dataset?.ribbonTab || "",
    activeAfterRight: document.querySelector(".ribbon-tab.is-active")?.dataset?.ribbonTab || "",
  })));


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.dataset?.ribbonTab === "insert", null, { timeout: 5000 });


  Object.assign(keyboardNavigation.ribbonKeyboard, await page.evaluate(() => ({
    activeAfterEnter: document.querySelector(".ribbon-tab.is-active")?.dataset?.ribbonTab || "",
    statusAfterEnter: document.querySelector("#statusText")?.textContent || "",
  })));


  await page.keyboard.press("ArrowLeft");


  await page.keyboard.press("Space");


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.dataset?.ribbonTab === "home", null, { timeout: 5000 });


  Object.assign(keyboardNavigation.ribbonKeyboard, await page.evaluate(() => ({
    focusedBeforePanel: document.activeElement?.dataset?.ribbonTab || "",
    activeAfterSpace: document.querySelector(".ribbon-tab.is-active")?.dataset?.ribbonTab || "",
    statusAfterSpace: document.querySelector("#statusText")?.textContent || "",
  })));


  await page.keyboard.press("ArrowDown");


  Object.assign(keyboardNavigation.ribbonKeyboard, await page.evaluate(() => ({
    focusedCommandAfterDown: document.activeElement?.dataset?.ribbonCommand || "",
    focusedTagAfterDown: document.activeElement?.tagName || "",
  })));


  await page.keyboard.press("ArrowRight");


  Object.assign(keyboardNavigation.ribbonKeyboard, await page.evaluate(() => ({
    focusedCommandAfterControlRight: document.activeElement?.dataset?.ribbonCommand || "",
    focusedClassAfterControlRight: document.activeElement?.className || "",
  })));


  await page.keyboard.press("Escape");


  Object.assign(keyboardNavigation.ribbonKeyboard, await page.evaluate(() => ({
    focusedAfterEscape: document.activeElement?.dataset?.ribbonTab || "",
  })));


  await page.click('[data-ribbon-tab="home"]');


  await page.evaluate(() => window.__webSheetDevSelectRange?.(73, 9, 73, 9));


  await page.keyboard.down("Control");


  await page.keyboard.press(";");


  await page.keyboard.up("Control");


  keyboardNavigation.shortcutDate = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return window.__webSheetDevCellModel?.(sheetName, 73, 9)?.raw || "";
  });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(73, 10, 73, 10));


  await page.keyboard.down("Control");


  await page.keyboard.down("Shift");


  await page.keyboard.press(";");


  await page.keyboard.up("Shift");


  await page.keyboard.up("Control");


  keyboardNavigation.shortcutTime = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return window.__webSheetDevCellModel?.(sheetName, 73, 10)?.raw || "";
  });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(73, 25, 74, 26));


  await page.locator("#formulaInput").click();


  await page.locator("#formulaInput").fill("CtrlEnterFill");


  await page.keyboard.press("Control+Enter");


  keyboardNavigation.ctrlEnterFormulaFill = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return {
      values: [
        window.__webSheetDevCellModel?.(sheetName, 73, 25)?.raw || "",
        window.__webSheetDevCellModel?.(sheetName, 73, 26)?.raw || "",
        window.__webSheetDevCellModel?.(sheetName, 74, 25)?.raw || "",
        window.__webSheetDevCellModel?.(sheetName, 74, 26)?.raw || "",
      ],
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    };
  });


  await page.keyboard.press("Control+Z");


  keyboardNavigation.ctrlEnterFormulaFillUndo = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return {
      values: [
        window.__webSheetDevCellModel?.(sheetName, 73, 25)?.raw || "",
        window.__webSheetDevCellModel?.(sheetName, 73, 26)?.raw || "",
        window.__webSheetDevCellModel?.(sheetName, 74, 25)?.raw || "",
        window.__webSheetDevCellModel?.(sheetName, 74, 26)?.raw || "",
      ],
      status: document.querySelector("#statusText")?.textContent || "",
    };
  });


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(73, 25, "ActiveCtrlEnter");
    window.__webSheetDevSelectRange?.(73, 25, 74, 26);
  });


  await page.keyboard.press("Control+Enter");


  keyboardNavigation.ctrlEnterActiveFill = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return {
      values: [
        window.__webSheetDevCellModel?.(sheetName, 73, 25)?.raw || "",
        window.__webSheetDevCellModel?.(sheetName, 73, 26)?.raw || "",
        window.__webSheetDevCellModel?.(sheetName, 74, 25)?.raw || "",
        window.__webSheetDevCellModel?.(sheetName, 74, 26)?.raw || "",
      ],
      status: document.querySelector("#statusText")?.textContent || "",
    };
  });


  await page.keyboard.press("Control+Z");


  await page.evaluate(() => window.__webSheetDevSelectRange?.(88, 27, 88, 27));


  await page.locator("#selectedCellLabel").click();


  await page.locator("#selectedCellLabel").fill("AA88");


  await page.locator("#selectedCellLabel").press("Enter");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent?.trim() === "AA88", null, { timeout: 5000 });


  keyboardNavigation.nameBoxJump = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent?.trim() || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(88, 28, 89, 29));


  await page.locator("#selectedCellLabel").click();


  await page.locator("#selectedCellLabel").fill("NameBoxRange");


  await page.locator("#selectedCellLabel").press("Enter");


  keyboardNavigation.nameBoxDefine = await page.evaluate(() => ({
    exists: window.__webSheetDevDefinedNames?.().some((definition) => definition.name === "NameBoxRange") || false,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.locator("#selectedCellLabel").click();


  await page.locator("#selectedCellLabel").fill("NameBoxRange");


  await page.locator("#selectedCellLabel").press("Enter");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent?.trim() === "AB88:AC89", null, { timeout: 5000 });


  keyboardNavigation.nameBoxDefinedNameJump = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent?.trim() || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.locator("#selectedCellLabel").click();


  await page.keyboard.press("Alt+ArrowDown");


  await page.waitForSelector('#nameBoxMenu [data-name-box-menu-item="NameBoxRange"]', { timeout: 5000 });


  const nameBoxMenuKeyboard = await page.evaluate(() => ({
    focusedBefore: document.activeElement?.dataset?.nameBoxMenuItem || "",
    itemCount: document.querySelectorAll("#nameBoxMenu [data-name-box-menu-item]").length,
  }));


  await page.keyboard.press("Escape");


  await page.waitForSelector("#nameBoxMenu", { state: "detached", timeout: 5000 });


  Object.assign(nameBoxMenuKeyboard, await page.evaluate(() => ({
    closedByEscape: !document.querySelector("#nameBoxMenu"),
    focusAfterEscape: document.activeElement?.id || "",
  })));


  await page.keyboard.press("Alt+ArrowDown");


  await page.waitForSelector('#nameBoxMenu [data-name-box-menu-item="NameBoxRange"]', { timeout: 5000 });


  await page.locator('#nameBoxMenu [data-name-box-menu-item="NameBoxRange"]').focus();


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent?.trim() === "AB88:AC89", null, { timeout: 5000 });


  keyboardNavigation.nameBoxMenuJump = {
    ...(await page.evaluate(() => ({
      label: document.querySelector("#selectedCellLabel")?.textContent?.trim() || "",
      status: document.querySelector("#statusText")?.textContent || "",
      menuClosed: !document.querySelector("#nameBoxMenu"),
    }))),
    keyboard: nameBoxMenuKeyboard,
  };


  await page.evaluate(() => window.__webSheetDevSelectRange?.(88, 30, 88, 30));


  await page.locator("#formulaInput").click();


  await page.locator("#formulaInput").fill("=a1:b2");


  await page.locator("#formulaInput").evaluate((input) => input.setSelectionRange(1, 3));


  await page.keyboard.press("F4");


  const formulaF4First = await page.locator("#formulaInput").inputValue();


  await page.keyboard.press("F4");


  const formulaF4Second = await page.locator("#formulaInput").inputValue();


  await page.keyboard.press("F4");


  const formulaF4Third = await page.locator("#formulaInput").inputValue();


  await page.keyboard.press("F4");


  const formulaF4Fourth = await page.locator("#formulaInput").inputValue();


  keyboardNavigation.formulaF4ReferenceCycle = {
    first: formulaF4First,
    second: formulaF4Second,
    third: formulaF4Third,
    fourth: formulaF4Fourth,
    status: await page.locator("#statusText").innerText(),
  };


  await page.keyboard.press("Escape");


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(73, 12, "FillDownShortcut");
    window.__webSheetDevSelectRange?.(73, 12, 75, 12);
  });


  await page.keyboard.press("Control+D");


  keyboardNavigation.fillDown = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return [73, 74, 75].map((row) => window.__webSheetDevCellModel?.(sheetName, row, 12)?.raw || "");
  });


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(73, 13, "FillRightShortcut");
    window.__webSheetDevSelectRange?.(73, 13, 73, 15);
  });


  await page.keyboard.press("Control+R");


  keyboardNavigation.fillRight = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return [13, 14, 15].map((col) => window.__webSheetDevCellModel?.(sheetName, 73, col)?.raw || "");
  });


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(72, 16, "SingleFillDown");
    window.__webSheetDevSetCell?.(73, 17, "SingleFillRight");
    window.__webSheetDevSelectRange?.(73, 16, 73, 16);
  });


  await page.keyboard.press("Control+D");


  await page.evaluate(() => window.__webSheetDevSelectRange?.(73, 18, 73, 18));


  await page.keyboard.press("Control+R");


  keyboardNavigation.singleCellFill = await page.evaluate(() => ({
    down: window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 73, 16)?.raw || "",
    right: window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 73, 18)?.raw || "",
  }));


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(73, 19, "=P73");
    window.__webSheetDevSelectRange?.(73, 19, 74, 19);
  });


  await page.keyboard.press("Control+D");


  keyboardNavigation.fillFormula = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return window.__webSheetDevCellModel?.(sheetName, 74, 19)?.formula || "";
  });


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(74, 20, "1");
    window.__webSheetDevSetCell?.(75, 20, "2");
    window.__webSheetDevSelectRange?.(76, 20, 76, 20);
  });


  await page.keyboard.down("Alt");


  await page.keyboard.press("=");


  await page.keyboard.up("Alt");


  keyboardNavigation.autoSumFormula = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return window.__webSheetDevCellModel?.(sheetName, 76, 20)?.formula || "";
  });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(74, 8, 74, 8));


  await page.keyboard.press("Control+9");


  keyboardNavigation.hiddenRowByShortcut = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return Boolean(window.__webSheetDevSheetModel?.(sheetName)?.rows?.[73]?.hidden);
  });


  await page.keyboard.press("Control+Shift+9");


  keyboardNavigation.unhiddenRowByShortcut = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return Boolean(window.__webSheetDevSheetModel?.(sheetName)?.rows?.[73]?.hidden);
  });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(73, 8, 73, 8));


  await page.keyboard.press("Control+0");


  keyboardNavigation.hiddenColumnByShortcut = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return Boolean(window.__webSheetDevSheetModel?.(sheetName)?.columns?.[7]?.hidden);
  });


  await page.keyboard.press("Control+Shift+0");


  keyboardNavigation.unhiddenColumnByShortcut = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return Boolean(window.__webSheetDevSheetModel?.(sheetName)?.columns?.[7]?.hidden);
  });


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(74, 21, "FilterA");
    window.__webSheetDevSetCell?.(74, 22, "FilterB");
    window.__webSheetDevSetCell?.(75, 21, "A1");
    window.__webSheetDevSetCell?.(75, 22, "B1");
    window.__webSheetDevSetCell?.(76, 21, "A2");
    window.__webSheetDevSetCell?.(76, 22, "B2");
    window.__webSheetDevSelectRange?.(74, 21, 74, 22);
  });


  await page.keyboard.press("Control+Shift+L");


  keyboardNavigation.filterShortcutOn = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return window.__webSheetDevSheetModel?.(sheetName)?.autoFilter?.range || null;
  });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(74, 21, 74, 21));


  await page.keyboard.press("Alt+ArrowDown");


  await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });


  keyboardNavigation.autoFilterAltDown = await page.evaluate(() => ({
    open: Boolean(document.querySelector("#autoFilterMenu")),
    col: document.querySelector("#autoFilterMenu")?.dataset.filterCol || "",
    focused: document.activeElement?.textContent?.trim() || document.activeElement?.getAttribute("aria-label") || "",
  }));


  await page.keyboard.press("Escape");


  await page.waitForSelector("#autoFilterMenu", { state: "detached", timeout: 5000 });


  await page.keyboard.press("Control+Shift+L");


  keyboardNavigation.filterShortcutOff = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return window.__webSheetDevSheetModel?.(sheetName)?.autoFilter?.range || null;
  });


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(74, 23, "=U75");
    window.__webSheetDevSelectRange?.(74, 23, 74, 23);
  });


  await page.keyboard.down("Control");


  await page.keyboard.press("Backquote");


  await page.keyboard.up("Control");


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().showFormulas === true, null, { timeout: 5000 });


  keyboardNavigation.showFormulasOn = await page.evaluate(() => ({
    showFormulas: window.__webSheetDevWorkbookState?.().showFormulas,
    text: document.querySelector('.sheet-cell[data-address="W74"] .cell-content')?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().showFormulas === false, null, { timeout: 5000 });


  keyboardNavigation.showFormulasUndo = await page.evaluate(() => ({
    showFormulas: window.__webSheetDevWorkbookState?.().showFormulas,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().showFormulas === true, null, { timeout: 5000 });


  keyboardNavigation.showFormulasRedo = await page.evaluate(() => ({
    showFormulas: window.__webSheetDevWorkbookState?.().showFormulas,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.down("Control");


  await page.keyboard.press("Backquote");


  await page.keyboard.up("Control");


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().showFormulas === false, null, { timeout: 5000 });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(74, 23, 74, 23));


  await page.keyboard.press("Control+[");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "U75", null, { timeout: 5000 });


  keyboardNavigation.precedentShortcut = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent || "");


  await page.keyboard.press("Control+]");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "W74", null, { timeout: 5000 });


  keyboardNavigation.dependentShortcut = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent || "");


  await page.evaluate(() => {
    window.__webSheetDevSetHeaderHidden?.("row", 75, true);
    window.__webSheetDevSetHeaderHidden?.("column", 22, true);
    window.__webSheetDevSelectRange?.(74, 21, 76, 23);
  });


  await page.keyboard.down("Alt");


  await page.keyboard.press(";");


  await page.keyboard.up("Alt");


  keyboardNavigation.visibleCellsShortcut = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent || "",
    ranges: window.__webSheetDevWorkbookState?.().selectionRanges || [],
  }));


  await page.evaluate(() => {
    window.__webSheetDevSetHeaderHidden?.("row", 75, false);
    window.__webSheetDevSetHeaderHidden?.("column", 22, false);
  });


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(77, 24, "FindAgain");
    window.__webSheetDevSetCell?.(78, 24, "FindAgain");
    window.__webSheetDevSelectRange?.(77, 24, 77, 24);
  });


  await page.keyboard.press("Control+F");


  await page.waitForSelector("#findDialog", { timeout: 5000 });


  await page.fill("#findDialog [data-find-field='query']", "FindAgain");


  await page.click("#findDialog [data-find-dialog-action='find-next']");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "X78", null, { timeout: 5000 });


  keyboardNavigation.findDialogFirst = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click("#findDialog [data-find-dialog-action='close']");


  await page.waitForSelector("#findDialog", { state: "detached", timeout: 5000 });


  await pressWorkbookShortcut(["Shift"], "F4");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "X77", null, { timeout: 5000 });


  keyboardNavigation.shiftF4FindNext = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await pressWorkbookShortcut(["Shift"], "F5");


  await page.waitForSelector("#findDialog", { timeout: 5000 });


  keyboardNavigation.shiftF5FindDialogTitle = await page.evaluate(() => document.querySelector("#findDialogTitle")?.textContent || "");


  await page.click("#findDialog [data-find-dialog-action='close']");


  await page.waitForSelector("#findDialog", { state: "detached", timeout: 5000 });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(74, 21, 74, 21));


  await page.keyboard.press("Control+K");


  await page.waitForSelector("#linkDialogOverlay", { timeout: 5000 });


  keyboardNavigation.linkShortcutTitle = await page.evaluate(() => document.querySelector("#linkDialogTitle")?.textContent || "");


  await page.click("#linkDialogOverlay [data-link-dialog-action='cancel']");


  await page.evaluate(() => window.__webSheetDevSelectRange?.(80, 22, 80, 22));


  await pressWorkbookShortcut(["Shift"], "F3");


  await page.waitForSelector('[data-data-dialog="insert-function"]', { timeout: 5000 });


  keyboardNavigation.insertFunctionShortcutTitle = await page.evaluate(() => document.querySelector("#insertFunctionTitle")?.textContent || "");


  await page.click('[data-data-dialog="insert-function"] [data-data-dialog-action="cancel"]');


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  await page.keyboard.press("Control+F3");


  await page.waitForSelector('[data-data-dialog="name-manager"]', { timeout: 5000 });


  keyboardNavigation.nameManagerShortcutTitle = await page.evaluate(() => document.querySelector("#nameManagerTitle")?.textContent || "");


  await page.locator('[data-data-dialog="name-manager"] [data-data-dialog-action="cancel"]').click();


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(78, 21, "NameHeader");
    window.__webSheetDevSetCell?.(79, 21, "NameValue");
    window.__webSheetDevSelectRange?.(78, 21, 79, 21);
  });


  await page.keyboard.press("Control+Shift+F3");


  await page.waitForSelector('[data-data-dialog="create-names"]', { timeout: 5000 });


  keyboardNavigation.createNamesShortcutTitle = await page.evaluate(() => document.querySelector("#createNamesTitle")?.textContent || "");


  await page.click('[data-data-dialog="create-names"] [data-data-dialog-action="create-names-ok"]');


  await page.waitForFunction(() => window.__webSheetDevDefinedNames?.().some((definition) => definition.name === "NameHeader"), null, { timeout: 5000 });


  keyboardNavigation.createdNameExists = await page.evaluate(() => window.__webSheetDevDefinedNames?.().some((definition) => definition.name === "NameHeader"));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(80, 21, 80, 21));


  await page.keyboard.press("F3");


  await page.waitForSelector('[data-data-dialog="paste-name"]', { timeout: 5000 });


  await page.evaluate(() => {
    const list = document.querySelector("[data-formula-paste-name-list]");
    const option = [...(list?.options || [])].find((item) => item.textContent.includes("NameHeader"));
    if (list && option) {
      list.value = option.value;
      list.dispatchEvent(new Event("change", { bubbles: true }));
    }
  });


  keyboardNavigation.pasteNameShortcut = await page.evaluate(() => ({
    title: document.querySelector("#pasteNameTitle")?.textContent || "",
    hasName: [...document.querySelectorAll("[data-formula-paste-name-list] option")].some((option) => option.textContent.includes("NameHeader")),
    reference: document.querySelector("[data-formula-paste-name-reference]")?.textContent || "",
  }));


  await page.click('[data-data-dialog="paste-name"] [data-data-dialog-action="paste-name-ok"]');


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  keyboardNavigation.pasteNameFormula = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    return window.__webSheetDevCellModel?.(sheetName, 80, 21)?.formula || "";
  });


  await page.keyboard.press("Control+P");


  await page.waitForSelector('[data-data-dialog="print"]', { timeout: 5000 });


  keyboardNavigation.printShortcut = await page.evaluate(() => ({
    title: document.querySelector("#printDialogTitle")?.textContent || "",
    fields: [...document.querySelectorAll("[data-data-dialog='print'] [data-data-field]")].map((field) => field.dataset.dataField),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.selectOption("[data-data-dialog='print'] [data-data-field='printOrientation']", "landscape");


  await page.fill("[data-data-dialog='print'] [data-data-field='printScale']", "125");


  await page.click('[data-data-dialog="print"] [data-data-dialog-action="print-page-setup"]');


  await page.waitForSelector('[data-data-dialog="page-setup"]', { timeout: 5000 });


  keyboardNavigation.printPageSetup = await page.evaluate(() => ({
    title: document.querySelector("#pageSetupTitle")?.textContent || "",
    landscape: document.querySelector("[data-page-setup-field='orientation'][value='landscape']")?.checked || false,
    scale: document.querySelector("[data-page-setup-field='scale']")?.value || "",
  }));


  await page.keyboard.press("Escape");


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  await page.keyboard.press("Control+F2");


  await page.waitForSelector('[data-data-dialog="print"]', { timeout: 5000 });


  keyboardNavigation.ctrlF2PrintPreview = await page.evaluate(() => ({
    title: document.querySelector("#printDialogTitle")?.textContent || "",
    orientation: document.querySelector("[data-data-field='printOrientation']")?.value || "",
    scale: document.querySelector("[data-data-field='printScale']")?.value || "",
  }));


  await page.keyboard.press("Escape");


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(82, 21, "TableName");
    window.__webSheetDevSetCell?.(82, 22, "TableScore");
    window.__webSheetDevSetCell?.(83, 21, "Ada");
    window.__webSheetDevSetCell?.(83, 22, "10");
    window.__webSheetDevSetCell?.(84, 21, "Grace");
    window.__webSheetDevSetCell?.(84, 22, "20");
    window.__webSheetDevSelectRange?.(82, 21, 84, 22);
  });


  await page.keyboard.press("Control+T");


  await page.waitForSelector('[data-data-dialog="create-table"]', { timeout: 5000 });


  keyboardNavigation.createTableShortcutDialog = await page.evaluate(() => ({
    title: document.querySelector("#createTableTitle")?.textContent || "",
    range: document.querySelector("[data-data-field='createTableRange']")?.value || "",
    hasHeaders: document.querySelector("[data-data-field='createTableHasHeaders']")?.checked || false,
  }));


  await page.click('[data-data-dialog="create-table"] [data-data-dialog-action="create-table-ok"]');


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  keyboardNavigation.createTableShortcut = await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    const tables = window.__webSheetDevSheetModel?.(sheetName)?.tables || [];
    const table = tables.find((item) => item.range?.top === 82 && item.range?.left === 21);
    return {
      tableCount: tables.length,
      range: table?.range || null,
      showFilterButton: table?.showFilterButton,
      status: document.querySelector("#statusText")?.textContent || "",
    };
  });


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => !(window.__webSheetDevSheetModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "")?.tables || []).some((item) => item.range?.top === 82 && item.range?.left === 21), null, { timeout: 5000 });


  keyboardNavigation.createTableShortcutUndo = await page.evaluate(() => ({
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(73, 5, 74, 7));


  await page.keyboard.press("Delete");


  await page.evaluate(() => window.__webSheetDevSelectRange?.(79, 8, 79, 8));


  await page.keyboard.press("Delete");


  await page.evaluate(() => window.__webSheetDevSelectRange?.(74, 21, 84, 27));


  await page.keyboard.press("Delete");


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 1, 1, 1));



  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    host.scrollTop = 0;
    host.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll"));
  });


  const handles = await page.waitForFunction(
    () => {
      const rectFor = (selector) => {
        const element = document.querySelector(selector);
        const rect = element?.getBoundingClientRect();
        if (!rect || rect.width <= 0 || rect.height <= 0) return null;
        return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
      };
      const column = rectFor('.column-resize-handle[data-col="1"]');
      const row = rectFor('.row-resize-handle[data-row="1"]');
      return column && row ? { column, row } : null;
    },
    null,
    { timeout: 10000 },
  );


  const { column: columnHandle, row: rowHandle } = await handles.jsonValue();


  const resizedHeaders = { beforeColWidth: newSheet.borderRightWidth };


  const beforeResize = await page.evaluate(() => {
    const col = document.querySelector('.column-header[data-col="1"]')?.getBoundingClientRect();
    const row = document.querySelector('.row-header[data-row="1"]')?.getBoundingClientRect();
    return { colWidth: Math.round(col?.width || 0), rowHeight: Math.round(row?.height || 0) };
  });


  await page.mouse.move(columnHandle.x + columnHandle.width / 2, columnHandle.y + columnHandle.height / 2);


  await page.mouse.down();


  await page.mouse.move(columnHandle.x + columnHandle.width / 2 + 40, columnHandle.y + columnHandle.height / 2, { steps: 6 });


  const resizeHandleVisual = await page.evaluate(() => {
    const visibleBackgroundCount = (selector) =>
      [...document.querySelectorAll(selector)].filter((handle) => {
        const color = getComputedStyle(handle).backgroundColor;
        return color && color !== "rgba(0, 0, 0, 0)" && color !== "transparent";
      }).length;
    return {
      bodyResizing: document.body.classList.contains("is-resizing-sheet"),
      highlightedColumnHandles: visibleBackgroundCount(".column-resize-handle"),
      highlightedRowHandles: visibleBackgroundCount(".row-resize-handle"),
      columnHandleCount: document.querySelectorAll(".column-resize-handle").length,
      rowHandleCount: document.querySelectorAll(".row-resize-handle").length,
    };
  });


  await page.mouse.up();


  await page.mouse.move(rowHandle.x + rowHandle.width / 2, rowHandle.y + rowHandle.height / 2);


  await page.mouse.down();


  await page.mouse.move(rowHandle.x + rowHandle.width / 2, rowHandle.y + rowHandle.height / 2 + 18, { steps: 5 });


  await page.mouse.up();


  await page.waitForTimeout(200);


  Object.assign(resizedHeaders, await page.evaluate(() => {
    const col = document.querySelector('.column-header[data-col="1"]')?.getBoundingClientRect();
    const row = document.querySelector('.row-header[data-row="1"]')?.getBoundingClientRect();
    return { afterColWidth: Math.round(col?.width || 0), afterRowHeight: Math.round(row?.height || 0) };
  }), beforeResize);


  const headerResizeUndoRedo = {};


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    ({ rowHeight }) => Math.abs(Math.round(document.querySelector('.row-header[data-row="1"]')?.getBoundingClientRect().height || 0) - rowHeight) <= 2,
    { rowHeight: resizedHeaders.rowHeight },
    { timeout: 10000 },
  );


  Object.assign(headerResizeUndoRedo, await page.evaluate(() => ({
    undoRowHeight: Math.round(document.querySelector('.row-header[data-row="1"]')?.getBoundingClientRect().height || 0),
    undoStatus: document.querySelector("#statusText")?.textContent,
  })));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(
    ({ rowHeight }) => Math.round(document.querySelector('.row-header[data-row="1"]')?.getBoundingClientRect().height || 0) > rowHeight + 10,
    { rowHeight: resizedHeaders.rowHeight },
    { timeout: 10000 },
  );


  Object.assign(headerResizeUndoRedo, await page.evaluate(() => ({
    redoRowHeight: Math.round(document.querySelector('.row-header[data-row="1"]')?.getBoundingClientRect().height || 0),
    redoStatus: document.querySelector("#statusText")?.textContent,
  })));



  await page.locator('.column-header[data-col="2"]').click();


  await page.locator('.column-header[data-col="4"]').click({ modifiers: ["Shift"] });


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "B:D", null, { timeout: 10000 });


  const multiColumnHandle = await page.waitForFunction(
    () => {
      const rect = document.querySelector('.column-resize-handle[data-col="3"]')?.getBoundingClientRect();
      return rect && rect.width > 0 && rect.height > 0 ? { x: rect.x, y: rect.y, width: rect.width, height: rect.height } : null;
    },
    null,
    { timeout: 10000 },
  );


  const multiColumnBox = await multiColumnHandle.jsonValue();


  await page.mouse.move(multiColumnBox.x + multiColumnBox.width / 2, multiColumnBox.y + multiColumnBox.height / 2);


  await page.mouse.down();


  await page.mouse.move(multiColumnBox.x + multiColumnBox.width / 2 + 28, multiColumnBox.y + multiColumnBox.height / 2, { steps: 5 });


  await page.mouse.up();


  await page.waitForTimeout(160);



  await page.locator('.row-header[data-row="4"]').click();


  await page.locator('.row-header[data-row="6"]').click({ modifiers: ["Shift"] });


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "4:6", null, { timeout: 10000 });


  const multiRowHandle = await page.waitForFunction(
    () => {
      const rect = document.querySelector('.row-resize-handle[data-row="5"]')?.getBoundingClientRect();
      return rect && rect.width > 0 && rect.height > 0 ? { x: rect.x, y: rect.y, width: rect.width, height: rect.height } : null;
    },
    null,
    { timeout: 10000 },
  );


  const multiRowBox = await multiRowHandle.jsonValue();


  await page.mouse.move(multiRowBox.x + multiRowBox.width / 2, multiRowBox.y + multiRowBox.height / 2);


  await page.mouse.down();


  await page.mouse.move(multiRowBox.x + multiRowBox.width / 2, multiRowBox.y + multiRowBox.height / 2 + 16, { steps: 5 });


  await page.mouse.up();


  await page.waitForTimeout(160);


  const multiHeaderResize = await page.evaluate(() => {
    const widthOf = (col) => Math.round(document.querySelector(`.column-header[data-col="${col}"]`)?.getBoundingClientRect().width || 0);
    const heightOf = (row) => Math.round(document.querySelector(`.row-header[data-row="${row}"]`)?.getBoundingClientRect().height || 0);
    return {
      columnLabel: "B:D",
      columnWidths: [widthOf(2), widthOf(3), widthOf(4)],
      rowLabel: document.querySelector("#selectedCellLabel")?.textContent,
      rowHeights: [heightOf(4), heightOf(5), heightOf(6)],
    };
  });



  const lazyLoadAnchor = await page.evaluate(async () => {
    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const columnNumber = (name) =>
      String(name || "A")
        .toUpperCase()
        .split("")
        .reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
    const maxRenderedRow = () => Math.max(...[...document.querySelectorAll("#sheetHost .row-header")].map((node) => Number(node.textContent) || 0));
    const maxRenderedCol = () => Math.max(...[...document.querySelectorAll("#sheetHost .column-header")].map((node) => columnNumber(node.textContent)));
    const minRenderedRow = () => Math.min(...[...document.querySelectorAll("#sheetHost .row-header")].map((node) => Number(node.textContent) || 0));
    const minRenderedCol = () => Math.min(...[...document.querySelectorAll("#sheetHost .column-header")].map((node) => columnNumber(node.textContent)));
    const host = document.querySelector("#sheetHost");
    const offsetFromEdge = 40;
    host.scrollTop = Math.max(0, host.scrollHeight - host.clientHeight - offsetFromEdge);
    host.scrollLeft = Math.max(0, host.scrollWidth - host.clientWidth - offsetFromEdge);
    const before = {
      scrollTop: host.scrollTop,
      scrollLeft: host.scrollLeft,
      scrollHeight: host.scrollHeight,
      scrollWidth: host.scrollWidth,
      bottomRemaining: host.scrollHeight - host.clientHeight - host.scrollTop,
      rightRemaining: host.scrollWidth - host.clientWidth - host.scrollLeft,
      minRow: minRenderedRow(),
      minCol: minRenderedCol(),
      maxRow: maxRenderedRow(),
      maxCol: maxRenderedCol(),
      rowHeaderWidth: Math.round(document.querySelector("#sheetHost .row-header")?.getBoundingClientRect().width || 0),
    };
    host.dispatchEvent(new Event("scroll"));
    await frame();
    await frame();
    const after = {
      scrollTop: host.scrollTop,
      scrollLeft: host.scrollLeft,
      scrollHeight: host.scrollHeight,
      scrollWidth: host.scrollWidth,
      bottomRemaining: host.scrollHeight - host.clientHeight - host.scrollTop,
      rightRemaining: host.scrollWidth - host.clientWidth - host.scrollLeft,
      minRow: minRenderedRow(),
      minCol: minRenderedCol(),
      maxRow: maxRenderedRow(),
      maxCol: maxRenderedCol(),
      rowHeaderWidth: Math.round(document.querySelector("#sheetHost .row-header")?.getBoundingClientRect().width || 0),
    };
    host.dispatchEvent(new Event("scroll"));
    await frame();
    await frame();
    const repeat = {
      scrollTop: host.scrollTop,
      scrollLeft: host.scrollLeft,
      scrollHeight: host.scrollHeight,
      scrollWidth: host.scrollWidth,
      bottomRemaining: host.scrollHeight - host.clientHeight - host.scrollTop,
      rightRemaining: host.scrollWidth - host.clientWidth - host.scrollLeft,
      maxRow: maxRenderedRow(),
      maxCol: maxRenderedCol(),
    };
    const dragTargetTop = Math.max(0, host.scrollTop - 600);
    const dragTargetLeft = Math.max(0, host.scrollLeft - 600);
    host.scrollTop = dragTargetTop;
    host.scrollLeft = dragTargetLeft;
    host.dispatchEvent(new Event("scroll"));
    await frame();
    await frame();
    return {
      before,
      after,
      repeat,
      dragBack: {
        targetTop: dragTargetTop,
      targetLeft: dragTargetLeft,
      scrollTop: host.scrollTop,
      scrollLeft: host.scrollLeft,
      scrollHeight: host.scrollHeight,
      scrollWidth: host.scrollWidth,
      bottomRemaining: host.scrollHeight - host.clientHeight - host.scrollTop,
      rightRemaining: host.scrollWidth - host.clientWidth - host.scrollLeft,
      maxRow: maxRenderedRow(),
        maxCol: maxRenderedCol(),
      },
    };
  });



  for (let i = 0; i < 8; i += 1) {
    await page.evaluate(() => {
      const host = document.querySelector("#sheetHost");
      host.scrollTop = host.scrollHeight;
      host.scrollLeft = host.scrollWidth;
      host.dispatchEvent(new Event("scroll"));
    });
    await page.evaluate(() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))));
    const reachedAq = await page.evaluate(() => Boolean(document.querySelector('.column-header[data-col="43"]')));
    if (reachedAq) break;
  }



  await page.waitForFunction(
    ({ row, col }) => {
      const maxRow = Math.max(...[...document.querySelectorAll("#sheetHost .row-header")].map((node) => Number(node.textContent) || 0));
      const columnNumber = (name) =>
        String(name)
          .toUpperCase()
          .split("")
          .reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
      const maxCol = Math.max(...[...document.querySelectorAll("#sheetHost .column-header")].map((node) => columnNumber(node.textContent || "A")));
      return maxRow > row && maxCol > col;
    },
    { row: newSheet.beforeMaxRow, col: newSheet.beforeMaxCol },
    { timeout: 10000 },
  );


  await page.evaluate(async () => {
    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const host = document.querySelector("#sheetHost");
    const grid = document.querySelector(".sheet-grid");
    const tracks = getComputedStyle(grid)
      .gridTemplateColumns.split(" ")
      .map((track) => Number.parseFloat(track) || 0);
    const aqLeft = tracks.slice(0, 43).reduce((total, width) => total + width, 0);
    host.scrollLeft = Math.max(0, aqLeft - 80);
    host.dispatchEvent(new Event("scroll"));
    await frame();
  });


  await page.waitForFunction(() => document.querySelector('.column-header[data-col="43"]'), null, { timeout: 10000 });


  const extension = await page.evaluate(() => ({
    rowsAfter: document.querySelectorAll("#sheetHost .row-header").length,
    colsAfter: document.querySelectorAll("#sheetHost .column-header").length,
    maxRowAfter: Math.max(...[...document.querySelectorAll("#sheetHost .row-header")].map((node) => Number(node.textContent) || 0)),
    clientWidthAfter: document.querySelector("#sheetHost")?.clientWidth || 0,
    scrollHeightAfter: document.querySelector("#sheetHost")?.scrollHeight || 0,
    scrollWidthAfter: document.querySelector("#sheetHost")?.scrollWidth || 0,
    rowHeaderWidthAfter: Math.round(document.querySelector("#sheetHost .row-header")?.getBoundingClientRect().width || 0),
    aqHeader: (() => {
      const header = document.querySelector('.column-header[data-col="43"]');
      const label = header?.querySelector(".header-label");
      return {
        text: header?.textContent?.trim(),
        textOverflow: header ? getComputedStyle(header).textOverflow : null,
        labelTextOverflow: label ? getComputedStyle(label).textOverflow : null,
      };
    })(),
    maxColAfter: Math.max(
      ...[...document.querySelectorAll("#sheetHost .column-header")].map((node) =>
        String(node.textContent || "A")
          .toUpperCase()
          .split("")
          .reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0),
      ),
    ),
  }));



  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    host.scrollTop = 0;
    host.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll"));
  });


  await page.waitForTimeout(260);


  await page.evaluate(() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))));


  const trimmedExtension = await page.evaluate(() => ({
    rowsAfterReturn: document.querySelectorAll("#sheetHost .row-header").length,
    colsAfterReturn: document.querySelectorAll("#sheetHost .column-header").length,
    maxRowAfterReturn: Math.max(...[...document.querySelectorAll("#sheetHost .row-header")].map((node) => Number(node.textContent) || 0)),
    maxColAfterReturn: Math.max(
      ...[...document.querySelectorAll("#sheetHost .column-header")].map((node) =>
        String(node.textContent || "A")
          .toUpperCase()
          .split("")
          .reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0),
      ),
    ),
    clientWidthAfterReturn: document.querySelector("#sheetHost")?.clientWidth || 0,
    scrollHeightAfterReturn: document.querySelector("#sheetHost")?.scrollHeight || 0,
    scrollWidthAfterReturn: document.querySelector("#sheetHost")?.scrollWidth || 0,
    rowHeaderWidthAfterReturn: Math.round(document.querySelector("#sheetHost .row-header")?.getBoundingClientRect().width || 0),
  }));


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="B2"]'), null, { timeout: 10000 });


  await page.locator('.sheet-cell[data-address="B2"]').click();


  await page.locator('[data-ribbon-style="fillColor"]').evaluate((input) => {
    input.value = "#ffffff";
    input.dispatchEvent(new Event("input", { bubbles: true }));
  });


  await page.waitForTimeout(200);


  const whiteFillGridMask = await page.evaluate(() => {
    const inspect = (address) => {
      const cell = document.querySelector(`.sheet-cell[data-address="${address}"]`);
      const style = cell ? getComputedStyle(cell) : null;
      return {
        className: cell?.className,
        fillColor: style?.getPropertyValue("--cell-fill-color").trim().toLowerCase(),
        backgroundColor: style?.backgroundColor,
        borderRightColor: style?.borderRightColor,
        borderBottomColor: style?.borderBottomColor,
      };
    };
    return {
      filled: inspect("B2"),
      leftNeighbor: inspect("A2"),
      topNeighbor: inspect("B1"),
      untouched: inspect("A1"),
    };
  });



  await page.click('.sheet-tab[data-sheet-index="0"]');


  await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent === "CHT", null, { timeout: 10000 });



  await page.click('[data-ribbon-tab="home"]');


  await page.locator('.sheet-cell[data-address="B14"]').click();


  await page.click('[data-ribbon-command="bold"]');


  await page.selectOption('[data-ribbon-style="fontSize"]', "16");


  await page.locator('.sheet-cell[data-address="D14"]').click({ modifiers: ["Shift"] });


  await page.click('[data-ribbon-command="merge-center"]');


  await page.waitForFunction(() => /span\s+3/.test(document.querySelector('.sheet-cell[data-address="B14"]')?.style.gridColumn || ""), null, { timeout: 10000 });


  const mergeUndoRedo = {};


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="D14"]') && !/span\s+3/.test(document.querySelector('.sheet-cell[data-address="B14"]')?.style.gridColumn || ""), null, { timeout: 10000 });


  Object.assign(mergeUndoRedo, await page.evaluate(() => ({
    undoLabel: document.querySelector("#selectedCellLabel")?.textContent,
    undoB14GridColumn: document.querySelector('.sheet-cell[data-address="B14"]')?.style.gridColumn,
    undoD14Rendered: Boolean(document.querySelector('.sheet-cell[data-address="D14"]')),
    undoStatus: document.querySelector("#statusText")?.textContent,
  })));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => /span\s+3/.test(document.querySelector('.sheet-cell[data-address="B14"]')?.style.gridColumn || "") && !document.querySelector('.sheet-cell[data-address="D14"]'), null, { timeout: 10000 });


  Object.assign(mergeUndoRedo, await page.evaluate(() => ({
    redoLabel: document.querySelector("#selectedCellLabel")?.textContent,
    redoB14GridColumn: document.querySelector('.sheet-cell[data-address="B14"]')?.style.gridColumn,
    redoD14Rendered: Boolean(document.querySelector('.sheet-cell[data-address="D14"]')),
    redoMergeButtonActive: document.querySelector('[data-ribbon-command="merge-center"]')?.classList.contains("is-active"),
    redoStatus: document.querySelector("#statusText")?.textContent,
  })));


  await page.click('[data-ribbon-command="merge-center"]');


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="D14"]') && !/span\s+3/.test(document.querySelector('.sheet-cell[data-address="B14"]')?.style.gridColumn || ""), null, { timeout: 10000 });


  Object.assign(mergeUndoRedo, await page.evaluate(() => ({
    toggleUnmergeLabel: document.querySelector("#selectedCellLabel")?.textContent,
    toggleUnmergeButtonActive: document.querySelector('[data-ribbon-command="merge-center"]')?.classList.contains("is-active"),
    toggleUnmergeD14Rendered: Boolean(document.querySelector('.sheet-cell[data-address="D14"]')),
    toggleUnmergeStatus: document.querySelector("#statusText")?.textContent,
  })));


  await page.click('[data-ribbon-command="merge-center"]');


  await page.waitForFunction(() => /span\s+3/.test(document.querySelector('.sheet-cell[data-address="B14"]')?.style.gridColumn || "") && !document.querySelector('.sheet-cell[data-address="D14"]'), null, { timeout: 10000 });


  Object.assign(mergeUndoRedo, await page.evaluate(() => ({
    toggleRemergeButtonActive: document.querySelector('[data-ribbon-command="merge-center"]')?.classList.contains("is-active"),
    toggleRemergeD14Rendered: Boolean(document.querySelector('.sheet-cell[data-address="D14"]')),
  })));


  const mergeAutoFillMismatch = {};


  const mergedFillHandleBox = await page.locator(".sheet-selection-frame .sheet-fill-handle").boundingBox();


  const e14Box = await page.locator('.sheet-cell[data-address="E14"]').boundingBox();


  if (!mergedFillHandleBox || !e14Box) throw new Error("Merged auto-fill handle or mismatch target cell was not visible.");


  await page.mouse.move(mergedFillHandleBox.x + mergedFillHandleBox.width / 2, mergedFillHandleBox.y + mergedFillHandleBox.height / 2);


  await page.mouse.down();


  await page.mouse.move(e14Box.x + e14Box.width / 2, e14Box.y + e14Box.height / 2, { steps: 6 });


  await page.mouse.up();


  await page.waitForFunction(
    () =>
      document.querySelector("#messageBox")?.textContent?.includes("結合セル") ||
      document.querySelector("#statusText")?.textContent?.includes("結合セル"),
    null,
    { timeout: 10000 },
  );


  Object.assign(mergeAutoFillMismatch, await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    topAlert: document.querySelector("#appAlert")?.textContent,
    messageTitle: document.querySelector("#messageBoxTitle")?.textContent,
    messageText: document.querySelector("#messageBoxText")?.textContent,
    okButtonText: document.querySelector("[data-message-box-close]")?.textContent,
    status: document.querySelector("#statusText")?.textContent,
    e14GridColumn: document.querySelector('.sheet-cell[data-address="E14"]')?.style.gridColumn,
    f14Rendered: Boolean(document.querySelector('.sheet-cell[data-address="F14"]')),
  })));


  await page.locator("[data-message-box-close]").click();


  await page.waitForFunction(() => !document.querySelector("#messageBoxOverlay"), null, { timeout: 10000 });


  const underlineToolbar = {};


  await page.click("[data-ribbon-underline-menu]");


  await page.waitForSelector("#ribbonUnderlineMenu");


  Object.assign(underlineToolbar, await page.evaluate(() => ({
    menuLabels: [...document.querySelectorAll("#ribbonUnderlineMenu [data-ribbon-underline-choice]")].map((button) => button.textContent.trim()),
    accessKeyTargets: [...document.querySelectorAll("#ribbonUnderlineMenu [data-access-key]")].map((button) => button.dataset.accessKey || ""),
    focusedBefore: document.activeElement?.dataset?.ribbonUnderlineChoice || "",
    marks: [...document.querySelectorAll("#ribbonUnderlineMenu [data-ribbon-underline-mark]")].map((mark) => ({
      letter: mark.dataset.ribbonUnderlineLetter,
      type: mark.dataset.ribbonUnderlineMark,
      decorationLine: getComputedStyle(mark).textDecorationLine,
      svgTextDecoration: getComputedStyle(mark.querySelector("text")).textDecorationLine,
      lineCount: mark.querySelectorAll("line").length,
    })),
    activeBefore: document.querySelector("#ribbonUnderlineMenu [data-ribbon-underline-choice].is-active")?.dataset.ribbonUnderlineChoice,
    accessKeysUnderlined: [...document.querySelectorAll("#ribbonUnderlineMenu .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
  })));


  await page.keyboard.press("D");


  await page.waitForFunction(() => !document.querySelector("#ribbonUnderlineMenu"), null, { timeout: 10000 });


  Object.assign(underlineToolbar, await page.evaluate(() => {
    const merged = document.querySelector('.sheet-cell[data-address="B14"]');
    const style = merged ? getComputedStyle(merged) : null;
    return {
      currentAfter: document.querySelector("[data-ribbon-underline-apply]")?.dataset.ribbonUnderlineCurrent,
      markAfter: document.querySelector("[data-ribbon-underline-mark]")?.dataset.ribbonUnderlineMark,
      markLetterAfter: document.querySelector("[data-ribbon-underline-mark]")?.dataset.ribbonUnderlineLetter,
      titleAfter: document.querySelector("[data-ribbon-underline-apply]")?.getAttribute("title"),
      decorationLine: style?.textDecorationLine,
      decorationStyle: style?.textDecorationStyle,
      statusAfter: document.querySelector("#statusText")?.textContent,
    };
  }));


  const borderToolbar = {};


  await page.click("[data-ribbon-border-menu]");


  await page.waitForSelector("#ribbonBorderMenu");


  Object.assign(borderToolbar, await page.evaluate(() => ({
    menuLabels: [...document.querySelectorAll("#ribbonBorderMenu [data-ribbon-border-choice]")].map((button) => button.textContent.trim()),
    activeBefore: document.querySelector("#ribbonBorderMenu [data-ribbon-border-choice].is-active")?.dataset.ribbonBorderChoice,
    focusedBefore: document.activeElement?.dataset?.ribbonBorderChoice || "",
    markBefore: document.querySelector("[data-ribbon-border-mark]")?.dataset.ribbonBorderMark,
    hasTextSelect: Boolean(document.querySelector('[data-ribbon-style="border"]')),
  })));


  await page.keyboard.press("ArrowDown");


  await page.keyboard.press("ArrowDown");


  Object.assign(borderToolbar, await page.evaluate(() => ({
    focusedAfterArrows: document.activeElement?.dataset?.ribbonBorderChoice || "",
  })));


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => !document.querySelector("#ribbonBorderMenu"), null, { timeout: 10000 });


  Object.assign(borderToolbar, await page.evaluate(() => ({
    currentAfter: document.querySelector("[data-ribbon-border-apply]")?.dataset.ribbonBorderCurrent,
    markAfter: document.querySelector("[data-ribbon-border-mark]")?.dataset.ribbonBorderMark,
    titleAfter: document.querySelector("[data-ribbon-border-apply]")?.getAttribute("title"),
    statusAfter: document.querySelector("#statusText")?.textContent,
  })));


  await page.locator('[data-ribbon-style="fillColor"]').evaluate((input) => {
    input.value = "#d9ead3";
    input.dispatchEvent(new Event("input", { bubbles: true }));
  });


  await page.locator('[data-ribbon-style="fontColor"]').evaluate((input) => {
    input.value = "#990000";
    input.dispatchEvent(new Event("input", { bubbles: true }));
  });


  await page.waitForTimeout(300);



  await page.evaluate(() => {
    const scroller = document.querySelector("#horizontalScrollbar");
    scroller.scrollLeft = 300;
    scroller.dispatchEvent(new Event("scroll"));
  });


  await page.waitForTimeout(200);



  const bottomBar = await page.evaluate(() => ({
    hostScrollLeft: document.querySelector("#sheetHost")?.scrollLeft,
    scrollbarScrollLeft: document.querySelector("#horizontalScrollbar")?.scrollLeft,
  }));



  const ribbon = await page.evaluate(() => {
    const homePanel = document.querySelector('[data-ribbon-panel="home"]');
    const merged = document.querySelector('.sheet-cell[data-address="B14"]');
    const mergedStyle = merged ? getComputedStyle(merged) : null;
    return {
      title: document.title,
      activeRibbonTab: document.querySelector(".ribbon-tab.is-active")?.textContent,
      homeVisible: Boolean(homePanel && getComputedStyle(homePanel).display !== "none"),
      selectionLabel: document.querySelector("#selectedCellLabel")?.textContent,
      mergedGridColumn: merged?.style.gridColumn,
      mergedTextAlign: mergedStyle?.textAlign,
      mergedFontWeight: mergedStyle?.fontWeight,
      mergedFontSize: mergedStyle?.fontSize,
      mergedBackground: mergedStyle?.backgroundColor,
      mergedColor: mergedStyle?.color,
      borderLines: document.querySelectorAll(".sheet-border-line").length,
      status: document.querySelector("#statusText")?.textContent,
    };
  });



  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    const scroller = document.querySelector("#horizontalScrollbar");
    host.scrollLeft = 0;
    scroller.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll"));
  });


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="B14"]'), null, { timeout: 10000 });


  await page.locator('.sheet-cell[data-address="B14"]').click();


  const mergedSingleSelection = await page.evaluate(() => {
    const rectFor = (selector) => {
      const element = document.querySelector(selector);
      const rect = element?.getBoundingClientRect();
      return rect
        ? {
            left: Math.round(rect.left),
            top: Math.round(rect.top),
            right: Math.round(rect.right),
            bottom: Math.round(rect.bottom),
          }
        : null;
    };
    return {
      label: document.querySelector("#selectedCellLabel")?.textContent,
      mergeButtonActive: document.querySelector('[data-ribbon-command="merge-center"]')?.classList.contains("is-active"),
      frame: rectFor(".sheet-selection-frame"),
      merged: rectFor('.sheet-cell[data-address="B14"]'),
    };
  });


  const mergedKeyboardNavigation = {};


  await page.keyboard.press("ArrowRight");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "E14", null, { timeout: 10000 });


  mergedKeyboardNavigation.afterRight = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("ArrowLeft");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "B14:D14", null, { timeout: 10000 });


  mergedKeyboardNavigation.afterLeftIntoMerge = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "B15:B43", null, { timeout: 10000 });


  mergedKeyboardNavigation.afterEnter = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("ArrowUp");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "B14:D14", null, { timeout: 10000 });


  mergedKeyboardNavigation.afterUpIntoMerge = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("Control+C");


  const mergedCopyState = await page.evaluate(() => {
    const frame = document.querySelector(".sheet-selection-frame");
    const beforeStyle = frame ? getComputedStyle(frame, "::before") : null;
    return {
      label: document.querySelector("#selectedCellLabel")?.textContent,
      frameClass: frame?.className,
      copyFrameCount: document.querySelectorAll(".sheet-copy-frame").length,
      animationName: beforeStyle?.animationName,
      backgroundImage: beforeStyle?.backgroundImage,
      clipboardStats: window.__webSheetMemoryStats?.() || null,
    };
  });


  await page.locator('.sheet-cell[data-address="E14"]').click();


  await page.keyboard.press("Control+V");


  await page.waitForFunction(
    () =>
      /span\s+3/.test(document.querySelector('.sheet-cell[data-address="E14"]')?.style.gridColumn || "") &&
      !document.querySelector('.sheet-cell[data-address="F14"]'),
    null,
    { timeout: 10000 },
  );


  const mergedPaste = await page.evaluate(() => {
    const pasted = document.querySelector('.sheet-cell[data-address="E14"]');
    const sourceCopyFrame = document.querySelector(".sheet-copy-frame");
    return {
      label: document.querySelector("#selectedCellLabel")?.textContent,
      pastedGridColumn: pasted?.style.gridColumn,
      pastedText: pasted?.querySelector(".cell-content")?.textContent,
      f14Rendered: Boolean(document.querySelector('.sheet-cell[data-address="F14"]')),
      g14Rendered: Boolean(document.querySelector('.sheet-cell[data-address="G14"]')),
      sourceCopyFrameClass: sourceCopyFrame?.className,
    };
  });


  await page.locator('.sheet-cell[data-address="A14"]').click();


  await page.locator('.sheet-cell[data-address="B14"]').click({ modifiers: ["Shift"] });


  const mergedCrossSelection = await page.evaluate(() => {
    const rectFor = (selector) => {
      const element = document.querySelector(selector);
      const rect = element?.getBoundingClientRect();
      return rect
        ? {
            left: Math.round(rect.left),
            top: Math.round(rect.top),
            right: Math.round(rect.right),
            bottom: Math.round(rect.bottom),
          }
        : null;
    };
    return {
      label: document.querySelector("#selectedCellLabel")?.textContent,
      mergeButtonActive: document.querySelector('[data-ribbon-command="merge-center"]')?.classList.contains("is-active"),
      frame: rectFor(".sheet-selection-frame"),
      leftCell: rectFor('.sheet-cell[data-address="A14"]'),
      merged: rectFor('.sheet-cell[data-address="B14"]'),
    };
  });



  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    const scroller = document.querySelector("#horizontalScrollbar");
    host.scrollTop = 0;
    host.scrollLeft = 0;
    scroller.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll"));
  });


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="A1"]'), null, { timeout: 10000 });


  await page.locator('.sheet-cell[data-address="A1"]').click();


  await page.keyboard.type("ReloadKeep");


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="A1"] .cell-content')?.textContent === "ReloadKeep", null, { timeout: 10000 });


  await page.locator('.column-header[data-col="1"]').click();


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "A:A", null, { timeout: 10000 });


  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    const scroller = document.querySelector("#horizontalScrollbar");
    host.scrollTop = 700;
    host.scrollLeft = 120;
    scroller.scrollLeft = 120;
    host.dispatchEvent(new Event("scroll"));
  });


  await page.evaluate(() => window.__webSheetPersistenceFlush?.());


  await page.evaluate(async () => {
    const db = await new Promise((resolve, reject) => {
      const request = indexedDB.open("cht-websheet", 1);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
      request.onblocked = () => reject(new Error("IndexedDB open was blocked."));
    });
    try {
      const snapshot = await new Promise((resolve, reject) => {
        const request = db.transaction("snapshots", "readonly").objectStore("snapshots").get("current-workbook");
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const dateCell = snapshot?.model?.sheets?.[0]?.cells?.["50:3"];
      if (dateCell) {
        const pollutedDate = new Date(2005, 1, 1, 9, 0, 0);
        dateCell.raw = pollutedDate.toISOString();
        dateCell.cached = pollutedDate.toString();
        dateCell.display = pollutedDate.toString();
        await new Promise((resolve, reject) => {
          const tx = db.transaction("snapshots", "readwrite");
          tx.objectStore("snapshots").put(snapshot, "current-workbook");
          tx.oncomplete = resolve;
          tx.onerror = () => reject(tx.error);
          tx.onabort = () => reject(tx.error || new Error("IndexedDB transaction aborted."));
        });
      }
    } finally {
      db.close();
    }
  });


  await page.reload({ waitUntil: "load" });


  await page.waitForFunction(
    (expectedName) =>
      document.querySelector("#workbookName")?.textContent?.includes(expectedName) &&
      document.querySelector('.sheet-cell[data-address="A1"] .cell-content')?.textContent === "ReloadKeep" &&
      document.querySelector("#selectedCellLabel")?.textContent === "A1" &&
      (document.querySelector("#sheetHost")?.scrollTop || 0) === 0 &&
      (document.querySelector("#sheetHost")?.scrollLeft || 0) === 0,
    expectedReloadWorkbookName,
    { timeout: 15000 },
  );


  const persistenceReload = await page.evaluate(() => ({
    workbookName: document.querySelector("#workbookName")?.textContent,
    title: document.title,
    a1: document.querySelector('.sheet-cell[data-address="A1"] .cell-content')?.textContent,
    status: document.querySelector("#statusText")?.textContent,
    exportDisabled: document.querySelector("#exportHtmlButton")?.disabled,
    workspaceVisible: !document.querySelector("#sheetWorkspace")?.classList.contains("hidden"),
    selectionLabel: document.querySelector("#selectedCellLabel")?.textContent,
    scrollTop: document.querySelector("#sheetHost")?.scrollTop || 0,
    scrollLeft: document.querySelector("#sheetHost")?.scrollLeft || 0,
    firstVisibleRows: [...document.querySelectorAll(".row-header")].slice(0, 3).map((row) => row.textContent),
    dateC50: window.__webSheetDevCellDisplay?.("CHT", 50, 3),
    dateCellModel: window.__webSheetDevCellModel?.("CHT", 50, 3),
    dateTimeSamples: {
      hms: window.__webSheetDevDisplayValue?.(new Date(2005, 1, 1, 9, 5, 7), "h:mm:ss"),
      ampm: window.__webSheetDevDisplayValue?.(new Date(2005, 1, 1, 21, 5, 7), "h:mm AM/PM"),
      dateTime: window.__webSheetDevDisplayValue?.(new Date(2005, 1, 1, 9, 5, 17), "yyyy/m/d h:mm:ss"),
    },
  }));



  await page.setViewportSize({ width: 2048, height: 720 });


  await page.reload({ waitUntil: "load" });


  await page.waitForFunction(
    (expectedName) =>
      document.querySelector("#workbookName")?.textContent?.includes(expectedName) &&
      document.querySelector("#selectedCellLabel")?.textContent === "A1" &&
      (document.querySelector("#sheetHost")?.scrollTop || 0) === 0 &&
      (document.querySelector("#sheetHost")?.scrollLeft || 0) === 0,
    expectedReloadWorkbookName,
    { timeout: 15000 },
  );


  const wideViewportCoverage = await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    const grid = document.querySelector(".sheet-grid");
    const headers = [...document.querySelectorAll(".column-header")];
    const hostRect = host?.getBoundingClientRect();
    const gridRect = grid?.getBoundingClientRect();
    const lastHeader = headers.at(-1);
    const lastHeaderRect = lastHeader?.getBoundingClientRect();
    return {
      viewportWidth: window.innerWidth,
      hostRight: Math.round(hostRect?.right || 0),
      gridRight: Math.round(gridRect?.right || 0),
      lastHeaderRight: Math.round(lastHeaderRect?.right || 0),
      lastHeaderText: lastHeader?.textContent || "",
      lastHeaderCol: Number(lastHeader?.dataset.col || 0),
      headerCount: headers.length,
    };
  });



  const nativeShapePage = await context.newPage();


  await nativeShapePage.goto(`${appUrl}?native-shape-test=${Date.now()}`, { waitUntil: "load" });


  await nativeShapePage.setInputFiles("#fileInput", richNativeShapeWorkbookPath);


  await nativeShapePage.waitForFunction(
    () => {
      const status = document.querySelector("#statusText")?.textContent || "";
      return document.querySelectorAll(".sheet-tab").length >= 3 && /読み込みました|読み込み済み|已加载/.test(status);
    },
    null,
    { timeout: 20000 },
  );


  await nativeShapePage.getByRole("button", { name: "Sheet1" }).click();


  await nativeShapePage.waitForFunction(() => document.querySelectorAll(".sheet-image.sheet-shape").length >= 6, null, { timeout: 10000 });


  const importedNativeShapes = await nativeShapePage.evaluate(() => {
    const objects = [...document.querySelectorAll(".sheet-image")].map((image) => {
      const rect = image.getBoundingClientRect();
      const source = image.getAttribute("src") || image.querySelector(".sheet-picture-crop-content")?.getAttribute("src") || "";
      const cropContent = image.querySelector(".sheet-picture-crop-content");
      return {
        alt: image.getAttribute("alt") || image.getAttribute("aria-label") || "",
        isShape: image.classList.contains("sheet-shape"),
        isCropped: image.classList.contains("sheet-picture-crop"),
        width: Math.round(rect.width),
        height: Math.round(rect.height),
        source: decodeURIComponent(source),
        transform: getComputedStyle(image).transform,
        cropContentWidth: cropContent ? Math.round(cropContent.getBoundingClientRect().width) : 0,
      };
    });
    const shapes = objects.filter((shape) => shape.isShape);
    const presetShapes = shapes.filter((shape) => !shape.alt.includes("グラフ"));
    const singlePictureElement = [...document.querySelectorAll(".sheet-image")].find((image) => (image.getAttribute("alt") || "").includes("単独画像"));
    const singlePictureStyle = singlePictureElement ? getComputedStyle(singlePictureElement) : null;
    const chartModel = window.__webSheetDevSheetModel?.("Sheet1")?.images?.find((image) => String(image.name || "").includes("グラフ 11"))?.chart || null;
    return {
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent.trim()),
      active: document.querySelector(".sheet-tab.is-active")?.textContent.trim() || "",
      objectCount: objects.length,
      shapeCount: shapes.length,
      names: objects.map((shape) => shape.alt),
      hasSvgShapes: presetShapes.every((shape) => shape.source.includes("image/svg+xml") && shape.source.includes('preserveAspectRatio="none"')),
      hasMeaningfulBounds: shapes.every((shape) => shape.width > 20 && shape.height > 20),
      hasRichTextRuns: shapes.some((shape) => {
        const source = shape.source.toLowerCase();
        return shape.source.includes("赤Bold") && shape.source.includes("青Italic") && source.includes('font-weight="700"') && source.includes("#ff0000");
      }),
      hasImportedShapeEffect: shapes.some((shape) => shape.source.includes('id="shapeEffects"') && shape.source.includes("feDropShadow")),
      hasGroupedPicture: objects.some((shape) => !shape.isShape && shape.alt.includes("グループ画像") && shape.source.startsWith("data:image/png;base64,")),
      hasGroupedPictureTransformAndCrop: objects.some((shape) => !shape.isShape && shape.alt.includes("グループ画像") && shape.isCropped && shape.transform !== "none" && shape.cropContentWidth > shape.width),
      hasGroupedShape: shapes.some((shape) => shape.alt.includes("グループ四角形")),
      hasSinglePicture: objects.some((shape) => !shape.isShape && shape.alt.includes("単独画像") && shape.source.startsWith("data:image/png;base64,")),
      hasImportedPictureBorder: singlePictureStyle?.borderTopColor === "rgb(0, 176, 80)" && singlePictureStyle?.borderTopWidth === "2px" && singlePictureStyle?.borderTopStyle === "dashed",
      hasImportedPictureEffect: singlePictureStyle?.filter?.includes("drop-shadow"),
      hasGraphicFrameChartRender: shapes.some((shape) => shape.alt.includes("グラフ 11") && shape.source.includes("売上推移") && shape.source.includes("Q1") && shape.source.includes("<rect")),
      hasGraphicFrameChartSeriesColors: shapes.some((shape) => shape.alt.includes("グラフ 11") && shape.source.includes("#70AD47") && shape.source.includes("#FFC000")),
      hasGraphicFrameChartAreaFormat: shapes.some((shape) => shape.alt.includes("グラフ 11") && shape.source.includes('fill="#FFF2CC"') && shape.source.includes('stroke="#C65911"') && shape.source.includes("stroke-dasharray")),
      hasGraphicFrameChartPlotFormat: shapes.some((shape) => shape.alt.includes("グラフ 11") && shape.source.includes('fill="#E2F0D9"') && shape.source.includes('stroke="#70AD47"')),
      hasGraphicFrameChartRightLegend: shapes.some((shape) => shape.alt.includes("グラフ 11") && shape.source.includes('data-chart-legend-position="right"')),
      hasGraphicFrameChartAxisFormat: shapes.some((shape) => shape.alt.includes("グラフ 11") && shape.source.includes('data-chart-axis="category"') && shape.source.includes('stroke="#5B9BD5"') && shape.source.includes('stroke="#ED7D31"')),
      hasGraphicFrameChartGridlines: shapes.some((shape) => shape.alt.includes("グラフ 11") && shape.source.includes('data-chart-gridlines="value"') && shape.source.includes('stroke="#A5A5A5"')),
      hasGraphicFrameChartDataLabels: shapes.some((shape) => shape.alt.includes("グラフ 11") && shape.source.includes('data-chart-data-label="value"') && shape.source.includes(">180<")),
      hasGraphicFrameChartHorizontalBar: shapes.some((shape) => shape.alt.includes("グラフ 11") && shape.source.includes('data-chart-bar-orientation="horizontal"')),
      hasGraphicFrameChartStacked: shapes.some((shape) => shape.alt.includes("グラフ 11") && shape.source.includes('data-chart-grouping="stacked"')),
      hasGraphicFrameChartAxisTitles: shapes.some((shape) => shape.alt.includes("グラフ 11") && shape.source.includes('data-chart-axis-title="category"') && shape.source.includes("四半期") && shape.source.includes('data-chart-axis-title="value"') && shape.source.includes("金額")),
      hasGraphicFrameChartTextStyles: shapes.some((shape) => shape.alt.includes("グラフ 11") && shape.source.includes('data-chart-title="chart"') && shape.source.includes('fill="#990000"') && shape.source.includes('font-size="18"') && shape.source.includes('data-chart-axis-title="category"') && shape.source.includes('fill="#00B050"') && shape.source.includes('font-style="italic"') && shape.source.includes('data-chart-axis-title="value"') && shape.source.includes('fill="#0070C0"') && shape.source.includes('data-chart-data-label="value"') && shape.source.includes('fill="#C00000"') && shape.source.includes('font-weight="700"')),
      hasGraphicFrameLineChartMarker: shapes.some((shape) => shape.alt.includes("グラフ 12") && shape.source.includes("マーカー確認") && shape.source.includes('data-chart-marker="diamond"') && shape.source.includes('fill="#FF0000"') && shape.source.includes('stroke="#0000FF"')),
      hasGraphicFramePieChartLabels: shapes.some((shape) => shape.alt.includes("グラフ 13") && shape.source.includes("構成比") && shape.source.includes('data-chart-data-label-position="bestFit"') && shape.source.includes("A部門, 40, 40%")),
      chartTitleRef: chartModel?.titleRef || "",
      chartLabelsRef: chartModel?.labelsRef || "",
      chartFirstNameRef: chartModel?.series?.[0]?.nameRef || "",
      chartFirstValuesRef: chartModel?.series?.[0]?.valuesRef || "",
      chartSecondNameRef: chartModel?.series?.[1]?.nameRef || "",
      chartSecondValuesRef: chartModel?.series?.[1]?.valuesRef || "",
    };
  });


  await nativeShapePage.locator('.sheet-image[alt*="グラフ 11"]').click();


  await nativeShapePage.waitForSelector(".sheet-image-selection-frame:not([data-image-group-id])", { timeout: 5000 });


  Object.assign(importedNativeShapes, await nativeShapePage.evaluate(() => ({
    chartFormatTabText: document.querySelector('[data-ribbon-tab="shape-format"]')?.textContent?.trim() || "",
    chartActiveRibbonTab: document.querySelector(".ribbon-tab.is-active")?.textContent?.trim() || "",
    chartFormatPanelMode: document.querySelector('[data-ribbon-panel="shape-format"]')?.classList.contains("is-chart-format") || false,
    chartInsertGroupHidden: document.querySelector(".ribbon-shape-insert-group")?.classList.contains("hidden") || false,
    chartTextGroupHidden: document.querySelector(".ribbon-shape-text-group")?.classList.contains("hidden") || false,
    chartFillButtonDisabled: document.querySelector("[data-shape-fill-menu]")?.disabled === true,
    chartStyleGroupLabel: document.querySelector(".ribbon-shape-style-group .ribbon-group-label")?.textContent?.trim() || "",
    chartOutlineButtonTitle: document.querySelector("[data-shape-outline-menu]")?.getAttribute("title") || "",
  })));


  await nativeShapePage.click(".sheet-image-selection-frame:not([data-image-group-id])", { button: "right" });


  Object.assign(importedNativeShapes, await nativeShapePage.evaluate(() => ({
    chartContextMenuFormatLabel: document.querySelector('#cellContextMenu [data-cell-menu-action="object-format"]')?.textContent?.trim() || "",
    chartEditTextDisabled: Boolean(document.querySelector('#cellContextMenu [data-cell-menu-action="object-edit-text"]:disabled')),
  })));


  await nativeShapePage.click('#cellContextMenu [data-cell-menu-action="object-format"]');


  await nativeShapePage.waitForSelector("#shapeFormatPane", { timeout: 5000 });


  Object.assign(importedNativeShapes, await nativeShapePage.evaluate(() => ({
    chartPaneTitle: document.querySelector("#shapeFormatPane .shape-format-pane-header strong")?.textContent?.trim() || "",
    chartShapeOnlyFieldsDisabled: ["fillColor", "textFontFamily"].every((field) => document.querySelector(`[data-shape-format-pane-field="${field}"]`)?.disabled),
    chartVisualFieldsEnabled: ["outlineColor", "outlineWidth", "lineDash", "shadow"].every((field) => !document.querySelector(`[data-shape-format-pane-field="${field}"]`)?.disabled),
  })));


  await nativeShapePage.evaluate(() => {
    const setField = (field, value) => {
      const input = document.querySelector(`[data-shape-format-pane-field="${field}"]`);
      if (!input) return false;
      input.value = value;
      input.dispatchEvent(new Event("change", { bubbles: true }));
      return true;
    };
    setField("outlineColor", "#ff0000");
    setField("outlineWidth", "2");
    setField("lineDash", "dash");
    setField("shadow", "outer1");
  });


  await nativeShapePage.waitForFunction(() => {
    const chart = document.querySelector('.sheet-image[alt*="グラフ 11"]');
    if (!chart) return false;
    const style = getComputedStyle(chart);
    return style.borderTopColor === "rgb(255, 0, 0)" &&
      style.borderTopWidth === "2px" &&
      style.borderTopStyle === "dashed" &&
      style.filter.includes("drop-shadow");
  }, null, { timeout: 5000 });


  Object.assign(importedNativeShapes, await nativeShapePage.evaluate(() => {
    const chart = document.querySelector('.sheet-image[alt*="グラフ 11"]');
    const style = chart ? getComputedStyle(chart) : null;
    return {
      chartBorderColor: style?.borderTopColor || "",
      chartBorderWidth: style?.borderTopWidth || "",
      chartBorderStyle: style?.borderTopStyle || "",
      chartFilter: style?.filter || "",
    };
  }));


  await nativeShapePage.keyboard.press("Escape");


  await nativeShapePage.locator('.sheet-image[aria-label*="グループ画像"]').click();


  await nativeShapePage.waitForSelector(".sheet-image-selection-frame[data-image-group-id]", { timeout: 5000 });


  Object.assign(importedNativeShapes, await nativeShapePage.evaluate(() => ({
    groupedObjectFormatTabVisible: !document.querySelector('[data-ribbon-tab="shape-format"]')?.classList.contains("hidden"),
    groupedObjectActiveRibbonTab: document.querySelector(".ribbon-tab.is-active")?.textContent?.trim() || "",
    groupedObjectSizeInputs: [...document.querySelectorAll("[data-shape-size]")].map((input) => input.value),
    groupedObjectRotationInput: document.querySelector("[data-shape-rotation]")?.value || "",
  })));


  await nativeShapePage.click(".sheet-image-selection-frame[data-image-group-id]", { button: "right" });


  Object.assign(importedNativeShapes, await nativeShapePage.evaluate(() => ({
    groupedObjectSelectionCount: document.querySelectorAll(".sheet-image.is-selected, .sheet-image-link.is-selected").length,
    groupedObjectUngroupEnabled: Boolean(document.querySelector('#cellContextMenu [data-cell-menu-action="object-ungroup"]:not(:disabled)')),
  })));


  await nativeShapePage.keyboard.press("Escape");


  await nativeShapePage.locator('.sheet-image[alt*="単独画像"]').click();


  await nativeShapePage.waitForSelector(".sheet-image-selection-frame:not([data-image-group-id])", { timeout: 5000 });


  Object.assign(importedNativeShapes, await nativeShapePage.evaluate(() => ({
    singlePictureFormatTabText: document.querySelector('[data-ribbon-tab="shape-format"]')?.textContent?.trim() || "",
    singlePictureActiveRibbonTab: document.querySelector(".ribbon-tab.is-active")?.textContent?.trim() || "",
    singlePictureFormatPanelPictureMode: document.querySelector('[data-ribbon-panel="shape-format"]')?.classList.contains("is-picture-format") || false,
    singlePictureInsertGroupHidden: document.querySelector(".ribbon-shape-insert-group")?.classList.contains("hidden") || false,
    singlePictureTextGroupHidden: document.querySelector(".ribbon-shape-text-group")?.classList.contains("hidden") || false,
    singlePictureFillButtonDisabled: document.querySelector("[data-shape-fill-menu]")?.disabled === true,
    singlePictureStyleGroupLabel: document.querySelector(".ribbon-shape-style-group .ribbon-group-label")?.textContent?.trim() || "",
    singlePictureOutlineButtonTitle: document.querySelector("[data-shape-outline-menu]")?.getAttribute("title") || "",
    singlePictureEffectsButtonTitle: document.querySelector("[data-shape-effects-menu]")?.getAttribute("title") || "",
  })));


  await nativeShapePage.click(".sheet-image-selection-frame:not([data-image-group-id])", { button: "right" });


  Object.assign(importedNativeShapes, await nativeShapePage.evaluate(() => ({
    singlePictureContextMenuFormatLabel: document.querySelector('#cellContextMenu [data-cell-menu-action="object-format"]')?.textContent?.trim() || "",
  })));


  await nativeShapePage.click('#cellContextMenu [data-cell-menu-action="object-format"]');


  await nativeShapePage.waitForSelector("#shapeFormatPane", { timeout: 5000 });


  Object.assign(importedNativeShapes, await nativeShapePage.evaluate(() => ({
    singlePicturePaneTitle: document.querySelector("#shapeFormatPane .shape-format-pane-header strong")?.textContent?.trim() || "",
    singlePicturePaneSizeValues: ["height", "width", "rotation"].map((field) => document.querySelector(`[data-shape-format-pane-field="${field}"]`)?.value || ""),
    singlePictureShapeOnlyFieldsDisabled: ["fillColor", "textFontFamily"].every((field) => document.querySelector(`[data-shape-format-pane-field="${field}"]`)?.disabled),
    singlePictureVisualFieldsEnabled: ["outlineColor", "outlineWidth", "lineDash", "shadow"].every((field) => !document.querySelector(`[data-shape-format-pane-field="${field}"]`)?.disabled),
    singlePictureSizeFieldsEnabled: ["height", "width", "rotation"].every((field) => !document.querySelector(`[data-shape-format-pane-field="${field}"]`)?.disabled),
  })));


  await nativeShapePage.locator('[data-shape-format-pane-field="outlineColor"]').evaluate((input) => {
    input.value = "#ff0000";
    input.dispatchEvent(new Event("change", { bubbles: true }));
  });


  await nativeShapePage.locator('[data-shape-format-pane-field="outlineWidth"]').evaluate((input) => {
    input.value = "2";
    input.dispatchEvent(new Event("change", { bubbles: true }));
  });


  await nativeShapePage.click('[data-shape-format-pane-tab="effects"]');


  await nativeShapePage.selectOption('[data-shape-format-pane-field="shadow"]', "outer1");


  await nativeShapePage.waitForTimeout(120);


  Object.assign(importedNativeShapes, await nativeShapePage.evaluate(() => {
    const image = document.querySelector('.sheet-image[alt*="単独画像"]');
    const style = image ? getComputedStyle(image) : null;
    return {
      singlePictureBorderColor: style?.borderTopColor || "",
      singlePictureBorderWidth: style?.borderTopWidth || "",
      singlePictureFilter: style?.filter || "",
    };
  }));


  await nativeShapePage.keyboard.press("Escape");


  const nativeConnectorBefore = await nativeShapePage.locator('.sheet-image.sheet-shape[alt*="直線矢印"]').boundingBox();


  await nativeShapePage.locator('.sheet-image.sheet-shape[alt*="正方形"]').click();


  await nativeShapePage.waitForSelector(".sheet-image-selection-frame:not([data-image-group-id])", { timeout: 5000 });


  const nativeRectFrame = await nativeShapePage.locator(".sheet-image-selection-frame:not([data-image-group-id])").boundingBox();


  await nativeShapePage.mouse.move(nativeRectFrame.x + 18, nativeRectFrame.y + 18);


  await nativeShapePage.mouse.down();


  await nativeShapePage.mouse.move(nativeRectFrame.x + 62, nativeRectFrame.y + 30, { steps: 6 });


  await nativeShapePage.mouse.up();


  await nativeShapePage.waitForTimeout(120);


  const nativeConnectorAfter = await nativeShapePage.locator('.sheet-image.sheet-shape[alt*="直線矢印"]').boundingBox();


  Object.assign(importedNativeShapes, {
    connectorBefore: nativeConnectorBefore
      ? {
          x: Math.round(nativeConnectorBefore.x),
          y: Math.round(nativeConnectorBefore.y),
          width: Math.round(nativeConnectorBefore.width),
          height: Math.round(nativeConnectorBefore.height),
        }
      : null,
    connectorAfter: nativeConnectorAfter
      ? {
          x: Math.round(nativeConnectorAfter.x),
          y: Math.round(nativeConnectorAfter.y),
          width: Math.round(nativeConnectorAfter.width),
          height: Math.round(nativeConnectorAfter.height),
        }
      : null,
    connectorFollowedMove: Boolean(
      nativeConnectorBefore &&
        nativeConnectorAfter &&
        (Math.abs(nativeConnectorAfter.x - nativeConnectorBefore.x) > 2 ||
          Math.abs(nativeConnectorAfter.y - nativeConnectorBefore.y) > 2 ||
          Math.abs(nativeConnectorAfter.width - nativeConnectorBefore.width) > 2 ||
          Math.abs(nativeConnectorAfter.height - nativeConnectorBefore.height) > 2),
    ),
  });


  Object.assign(importedNativeShapes, await nativeShapePage.evaluate(() => {
    [
      [1, 1, "R1"],
      [2, 1, "R2"],
      [3, 1, "R3"],
      [1, 2, 111],
      [2, 2, 222],
      [3, 2, "=B1+B2"],
      [1, 3, 11],
      [2, 3, 22],
      [3, 3, "=C1+C2"],
      [1, 4, "更新グラフ"],
      [1, 5, "更新売上"],
      [1, 6, "更新利益"],
    ].forEach(([row, col, value]) => {
      window.__webSheetDevSetCell?.(row, col, value);
    });
    const explicitRefreshChanged = window.__webSheetDevRefreshCharts?.() === true;
    const chartImage = window.__webSheetDevSheetModel?.("Sheet1")?.images?.find((image) => String(image.name || "").includes("グラフ 11"));
    const chart = chartImage?.chart || null;
    const source = decodeURIComponent(chartImage?.src || "");
    return {
      chartDataRefresh: {
        refreshed: typeof window.__webSheetDevRefreshCharts === "function",
        explicitRefreshChanged,
        title: chart?.title || "",
        firstName: chart?.series?.[0]?.name || "",
        secondName: chart?.series?.[1]?.name || "",
        labels: chart?.labels || [],
        firstValues: chart?.series?.[0]?.values || [],
        secondValues: chart?.series?.[1]?.values || [],
        svgHasUpdatedTitle: source.includes("更新グラフ"),
        svgHasUpdatedSeriesName: source.includes("更新売上") && source.includes("更新利益"),
        svgHasUpdatedLabel: source.includes(">R3<"),
        svgHasUpdatedValue: source.includes(">333<"),
      },
    };
  }));


  Object.assign(importedNativeShapes, await nativeShapePage.evaluate(() => {
    const chartRefs = () => {
      const chart = window.__webSheetDevSheetModel?.("Sheet1")?.images?.find((image) => String(image.name || "").includes("グラフ 11"))?.chart || null;
      return {
        title: chart?.titleRef || "",
        labels: chart?.labelsRef || "",
        firstName: chart?.series?.[0]?.nameRef || "",
        first: chart?.series?.[0]?.valuesRef || "",
        secondName: chart?.series?.[1]?.nameRef || "",
        second: chart?.series?.[1]?.valuesRef || "",
      };
    };
    const before = chartRefs();
    const rowApplied = window.__webSheetDevApplyCellStructure?.("insert", "entire-row", 1, 1);
    const afterRow = chartRefs();
    const columnApplied = window.__webSheetDevApplyCellStructure?.("insert", "entire-column", 1, 2);
    return {
      chartRowStructureApplied: rowApplied === true,
      chartColumnStructureApplied: columnApplied === true,
      chartRefsBeforeStructure: before,
      chartRefsAfterRowInsert: afterRow,
      chartRefsAfterRowAndColumnInsert: chartRefs(),
    };
  }));


  await nativeShapePage.close();



  const customViewImportExport = await exerciseCustomViewImportExport(context);


  const homeEditingCommands = await exerciseHomeEditingCommands(context);


  const dataValidationMenu = await exerciseDataValidationMenu(context);


  const autoFilterConditionMenu = await exerciseAutoFilterConditionMenu(context);


  const importedAutoFilterCriteria = await exerciseImportedAutoFilterCriteria(context);


  const autoFilterBlankKeyboardMenu = await exerciseAutoFilterBlankAndKeyboardMenu(context);


  const autoFilterReapplyClear = await exerciseAutoFilterReapplyAndClear(context);


  const advancedFilterCopyOverwrite = await exerciseAdvancedFilterCopyOverwrite(context);


  const autoFilterDateMenu = await exerciseAutoFilterDateMenu(context);


  const advancedFilterUniqueDates = await exerciseAdvancedFilterUniqueDates(context);


  const advancedFilterUniqueSubsetCopy = await exerciseAdvancedFilterUniqueSubsetCopy(context);


  const delimitedTextImport = await exerciseDelimitedTextImportOverwrite(context);


  const powerQueryTransformEditor = await exercisePowerQueryTransformEditor(context);


  const whatIfGoalSeekDataTable = await exerciseWhatIfGoalSeekDataTable(context);


  const scenarioManagerSummary = await exerciseScenarioManagerSummary(context);


  const flashFillDataPatterns = await exerciseFlashFillDataPatterns(context);


  const dataCurrentRegion = await exerciseDataCurrentRegion(context);


  const queryConnectionsPane = await exerciseQueryConnectionsPane(context);


  const dataGetConnectionMenus = await exerciseDataGetConnectionMenus(context);


  const relationshipsDialog = await exerciseRelationshipsDialog(context);


  const refreshAllUndo = await exerciseRefreshAllUndo(context);


  const forecastSheet = await exerciseForecastSheet(context);


  const dataAnalysisToolpak = await exerciseDataAnalysisToolpak(context);


  const consolidateDialog = await exerciseConsolidateDialog(context);


  const subtotalDialog = await exerciseSubtotalDialog(context);


  const outlineGrouping = await exerciseOutlineGrouping(context);


  const textToColumnsDialog = await exerciseTextToColumnsDialog(context);


  const textToColumnsFixedWidth = await exerciseTextToColumnsFixedWidth(context);


  const textToColumnsAdvanced = await exerciseTextToColumnsAdvanced(context);


  const sortDialogLevels = await exerciseSortDialogLevels(context);


  const sortDialogColorOrder = await exerciseSortDialogColorOrder(context);


  const sortDialogLeftToRight = await exerciseSortDialogLeftToRight(context);


  const sortDialogBlankAndCase = await exerciseSortDialogBlankAndCase(context);


  const sortDialogCustomList = await exerciseSortDialogCustomList(context);


  const sortDialogDateOrder = await exerciseSortDialogDateOrder(context);


  const sortFormulaResults = await exerciseSortFormulaResults(context);


  const quickSortActiveColumn = await exerciseQuickSortActiveColumn(context);


  const quickSortPartialSelection = await exerciseQuickSortPartialSelection(context);


  const removeDuplicatesDialog = await exerciseRemoveDuplicatesDialog(context);



	  console.log(JSON.stringify({ result, dataRibbonLayout, coordinateHitTest, delayedLazyLoad, stickyHeaders, sheetMenu, sheetTabColor, sheetRename, sheetTabDirectActions, sheetMoveCopy, sheetCodeDialog, sheetSelectAll, sheetHidden, sheetUnhideDialog, sheetUnhidden, sheetUnhideUndo, sheetUnhideRedo, sheetGroupActions, sheetGroupMoveCopy, sheetGroupCellEditing, sheetGroupStructure, sheetGroupFormatDialog, sheetGroupDataValidation, sheetGroupDataFilterSort, sheetGroupDataTools, sheetGroupHyperlink, sheetGroupHomeCommands, sheetGroupTranslation, sheetGroupQuickAnalysis, sheetGroupSparklines, sheetGroupCharts, sheetGroupTables, sheetGroupPivot, sheetGroupObjects, insertRibbon, formulaRibbon, formulaDefineNameMain, formulaDefineNameMenu, formulaDefineNameCreate, createNamesDialog, formulaNames, formulaNameManager, formulaUseNameMenu, formulaUseNameManage, formulaUseNameInserted, formulaPasteNameDialog, formulaPasteNameInserted, formulaPasteNameListDialog, formulaPasteNameList, formulaPasteNameListUndo, formulaPasteNameListRedo, formulaFunctionDialog, formulaFunctionSearch, formulaFunctionArguments, formulaInserted, formulaLogicalMenu, formulaMenuArguments, formulaMenuInserted, formulaRecentMenu, formulaVariableArguments, formulaZeroArgumentFunction, formulaMathMenu, formulaFinancialMenu, formulaDateMenu, formulaTextMenu, formulaMoreMenu, formulaShowState, formulaTrace, formulaRemoveArrowsMenu, formulaRemovePrecedents, formulaEvaluateInitial, formulaEvaluateStepIn, formulaEvaluateStep, formulaEvaluateStepOut, formulaEvaluateReferenceInitial, formulaEvaluateReferenceStep, formulaEvaluateReferenceStepIn, formulaEvaluateReferenceStepOut, formulaWatchWindow, formulaWatchSelect, formulaWatchDelete, formulaErrorMenu, formulaErrorChecking, formulaErrorIgnore, formulaErrorTrace, formulaErrorResetIgnored, formulaCalculation, viewRibbon, viewToggles, viewWindowCommands, customViewImportExport, layoutRibbon, layoutThemeMenuInitial, layoutThemeApplied, layoutColorMenuInitial, layoutFontMenuInitial, layoutEffectMenuInitial, layoutThemeAfterMenus, layoutThemeEffectUndoRedo, layoutBackgroundApplied, layoutBackgroundUndo, layoutBackgroundRedo, layoutBackgroundDeleted, layoutBackgroundDeleteUndo, layoutPrintAreaMenuBefore, layoutPrintAreaMenuAfterSet, layoutPrintAreaMenuAfterClear, layoutMarginsMenuInitial, layoutOrientationMenuInitial, layoutSizeMenuInitial, pageSetupDialog, printTitlesDialogInitial, printTitlesDialogFilled, pageLayout, pageLayoutPreview, pageBreakMenuKeyboard, pageBreakDrag, drawRibbon, drawInk, drawConverted, headerFooterDialog, headerFooterApplied, headerFooterUndo, headerFooterRedo, insertMapChart, insertIconDialog, signatureLineDialog, insertObjectDialog, insertAddinsStoreDialog, insertAddinsMineDialog, insertTextObjects, insertShapeMenu, insertShapeDraw, fileSideMenu, fileSideMenuClosedByOutsideClick, fileSideMenuClosedByCommand, shortcutSave, saveAsDialog, exportedHtmlChecks, standaloneOffline, standaloneRemoteFallback, standaloneEmptyModel, standaloneMalformedModel, standaloneWrappedModel, standaloneLauncher, standaloneBlockedRequests, launcherBlockedRequests, ar2Gridline, imageAndFont, imageSelection, footerLayout, realCtrlWheelZoom, zoomPopup, bottomBar, footerBeforeScroll, tabScroll, singleClickEdit, doubleClickEdit, editAfterDeferredRender, editAlignment, editClickCaret, editTextSelection, editNoEllipsis, cellEditAltEnter, cellEditAltEnterDuringIme, editDragOutsideSelection, cellEditImeEnter, editScrollbar, editExitRender, multilineCommitDisplay, formulaBarInitial, formulaBarSizing, formulaBarDuringEdit, formulaBarEscape, formulaBarEnter, formulaAutocomplete, formulaReferenceInsert, directTyping, definedNameAndLink, statusBarStats, autoFillSeries, insertDeleteDialog, keyboardNavigation, shortcutUndo, shortcutClipboard, formatUndoRedo, toolbarSync, selectionFrame, copyStateAfterCopy, copyStateWithTarget, richPaste, matchDestinationPaste, pasteSpecialFormats, pasteSpecialAdvanced, formulaPasteReferences, cutAndRepeatPaste, cellMenu, cellAnnotations, furigana, formatCellsDialog, cellMenuClear, cellMenuUndo, deleteKeyClear, deleteKeyUndo, backspaceClearEdit, backspaceUndo, popupDismissal, newSheet, colorPalette, headerSelection, selectionAutoScroll, resizedHeaders, resizeHandleVisual, headerResizeUndoRedo, multiHeaderResize, lazyLoadAnchor, extension, trimmedExtension, whiteFillGridMask, underlineToolbar, borderToolbar, ribbon, mergeUndoRedo, mergeAutoFillMismatch, mergedSingleSelection, mergedKeyboardNavigation, mergedCopyState, mergedPaste, mergedCrossSelection, persistenceReload, wideViewportCoverage, importedNativeShapes, homeEditingCommands, dataValidationMenu, autoFilterConditionMenu, importedAutoFilterCriteria, autoFilterBlankKeyboardMenu, autoFilterReapplyClear, advancedFilterCopyOverwrite, autoFilterDateMenu, advancedFilterUniqueDates, advancedFilterUniqueSubsetCopy, delimitedTextImport, powerQueryTransformEditor, whatIfGoalSeekDataTable, scenarioManagerSummary, flashFillDataPatterns, dataCurrentRegion, queryConnectionsPane, dataGetConnectionMenus, relationshipsDialog, refreshAllUndo, forecastSheet, dataAnalysisToolpak, consolidateDialog, subtotalDialog, outlineGrouping, textToColumnsDialog, textToColumnsFixedWidth, textToColumnsAdvanced, sortDialogLevels, sortDialogColorOrder, sortDialogLeftToRight, sortDialogBlankAndCase, sortDialogCustomList, sortDialogDateOrder, sortFormulaResults, quickSortActiveColumn, quickSortPartialSelection, removeDuplicatesDialog, messages: messages.slice(-30) }, null, 2));



  if (result.hasTopbar || result.hasHeaderExport || result.hasHeaderToggleFormulas) {
    throw new Error("Removed header controls are still present.");
  }


  if (
    dataRibbonLayout.active !== "データ" ||
    dataRibbonLayout.sortOverlaps?.length ||
    dataRibbonLayout.panelOverflow > 0 ||
    !["データの取得と変換", "並べ替えとフィルター", "データ ツール", "予測", "アウトライン"].every((label) => dataRibbonLayout.groupLabels.includes(label))
  ) {
    throw new Error("Data ribbon should be compact and should not overlap controls.");
  }


  if (
    homeEditingCommands.autoSum?.label !== "D5" ||
    homeEditingCommands.autoSum?.formula !== "=SUM(D2:D4)" ||
    homeEditingCommands.autoSum?.display !== "6" ||
    !homeEditingCommands.autoSum?.status?.includes("オートSUM") ||
    homeEditingCommands.autoSumUndo?.display !== "" ||
    !homeEditingCommands.autoSumUndo?.status?.includes("元に戻しました") ||
    homeEditingCommands.autoSumRedo?.display !== "6" ||
    !homeEditingCommands.autoSumRedo?.status?.includes("やり直しました") ||
    !["合計(S)", "平均(A)", "数値の個数(C)", "最大値(M)", "最小値(I)"].every((label) =>
      homeEditingCommands.autoSumMenu?.labels?.includes(label),
    ) ||
    !["S", "A", "C", "M", "I"].every((key) => homeEditingCommands.autoSumMenu?.accessKeys?.includes(key)) ||
    homeEditingCommands.autoSumMenu?.focused !== "合計(S)" ||
    homeEditingCommands.autoSumAverage?.label !== "H5" ||
    homeEditingCommands.autoSumAverage?.formula !== "=AVERAGE(H2:H4)" ||
    homeEditingCommands.autoSumAverage?.display !== "8" ||
    !homeEditingCommands.autoSumAverage?.status?.includes("平均") ||
	    !["下方向へコピー(D)", "右方向へコピー(R)", "上方向へコピー(U)", "左方向へコピー(L)", "作業グループへコピー(W)...", "連続データ(S)...", "両端揃え(J)"].every((label) =>
	      homeEditingCommands.fillMenu?.labels?.includes(label),
	    ) ||
	    !["D", "R", "U", "L", "W", "S", "J"].every((key) => homeEditingCommands.fillMenu?.accessKeys?.includes(key)) ||
	    homeEditingCommands.fillDown?.values?.join("|") !== "Seed|Seed|Seed" ||
	    homeEditingCommands.fillDown?.sourceFill !== true ||
	    homeEditingCommands.fillDown?.copiedFill !== true ||
	    !homeEditingCommands.fillDown?.status?.includes("フィル") ||
	    homeEditingCommands.fillFormula?.k3 !== "=H3*2" ||
	    homeEditingCommands.fillFormula?.k4 !== "=H4*2" ||
	    homeEditingCommands.fillFormula?.displays?.join("|") !== "8|16|24" ||
	    homeEditingCommands.fillUp?.values?.join("|") !== "UpSeed|UpSeed|UpSeed" ||
	    !homeEditingCommands.fillUp?.status?.includes("フィル") ||
	    homeEditingCommands.fillLeft?.values?.join("|") !== "LeftSeed|LeftSeed|LeftSeed" ||
	    !homeEditingCommands.fillLeft?.status?.includes("フィル") ||
	    homeEditingCommands.seriesDialogInitial?.title !== "連続データ" ||
	    homeEditingCommands.seriesDialogInitial?.direction !== "columns" ||
	    homeEditingCommands.seriesDialogInitial?.type !== "linear" ||
	    homeEditingCommands.seriesDialogInitial?.step !== "1" ||
	    homeEditingCommands.seriesDialogInitial?.dateUnitDisabled !== true ||
	    !["行(R)", "列(C)", "加算(L)", "乗算(G)", "日付(D)", "オートフィル(A)"].every((label) => homeEditingCommands.seriesDialogInitial?.labels?.includes(label)) ||
	    homeEditingCommands.seriesLinear?.values?.join("|") !== "2|3|4|5" ||
	    !homeEditingCommands.seriesLinear?.status?.includes("連続データ") ||
	    homeEditingCommands.seriesGrowth?.values?.join("|") !== "2|4|8|16" ||
	    !homeEditingCommands.seriesGrowth?.status?.includes("連続データ") ||
	    !homeEditingCommands.fillJustify?.values?.join(" ")?.includes("Alpha Beta Gamma") ||
	    !homeEditingCommands.fillJustify?.values?.join(" ")?.includes("Lambda Mu") ||
	    !homeEditingCommands.fillJustify?.status?.includes("両端揃え") ||
	    !homeEditingCommands.worksheetFillNoGroup?.message?.includes("複数のシート") ||
	    !["すべてクリア(A)", "書式のクリア(F)", "数式と値のクリア(C)", "コメントとメモのクリア(M)", "ハイパーリンクのクリア(H)"].every((label) =>
      homeEditingCommands.clearMenu?.labels?.includes(label),
    ) ||
    !["A", "F", "C", "M", "H"].every((key) => homeEditingCommands.clearMenu?.accessKeys?.includes(key)) ||
    homeEditingCommands.clearContents?.raw !== null ||
    homeEditingCommands.clearContents?.display !== "" ||
    homeEditingCommands.clearContents?.hasFill !== true ||
    homeEditingCommands.clearUndo?.display !== "Seed" ||
    homeEditingCommands.clearUndo?.hasFill !== true ||
    !homeEditingCommands.clearUndo?.status?.includes("元に戻しました") ||
    homeEditingCommands.clearFormats?.raw !== "Format" ||
    homeEditingCommands.clearFormats?.hasFill !== false ||
    homeEditingCommands.clearFormats?.borders !== 0 ||
    homeEditingCommands.clearNotes?.note !== null ||
    !homeEditingCommands.clearNotes?.status?.includes("コメントとメモ") ||
    homeEditingCommands.clearHyperlinks?.raw !== "Link" ||
    homeEditingCommands.clearHyperlinks?.hyperlink !== null ||
    !["昇順(S)", "降順(O)", "ユーザー設定の並べ替え(U)...", "フィルター(F)", "クリア(C)", "再適用(R)"].every((label) =>
      homeEditingCommands.sortFilterMenu?.labels?.includes(label),
    ) ||
    !["S", "O", "U", "F", "C", "R"].every((key) => homeEditingCommands.sortFilterMenu?.accessKeys?.includes(key)) ||
    homeEditingCommands.sortFilterMenu?.firstFocused !== "昇順(S)" ||
    homeEditingCommands.sorted?.names?.join("|") !== "Alpha|Beta|Gamma" ||
    homeEditingCommands.sorted?.scores?.join("|") !== "1|2|3" ||
    homeEditingCommands.filter?.buttonCount < 2 ||
    homeEditingCommands.filter?.menuClosed !== true ||
    !["検索(F)...", "置換(R)...", "ジャンプ(G)...", "条件を選択(S)...", "数式(R)", "コメント(M)", "メモ(N)", "条件付き書式(C)", "定数(O)", "データの入力規則(V)", "オブジェクト(B)"].every((label) =>
      homeEditingCommands.findSelectMenu?.labels?.includes(label),
    ) ||
    !["F", "R", "G", "S", "M", "N", "C", "O", "V", "B"].every((key) => homeEditingCommands.findSelectMenu?.accessKeys?.includes(key)) ||
    homeEditingCommands.findDialogInitial?.title !== "検索と置換" ||
    homeEditingCommands.findDialogInitial?.activeInput !== true ||
    homeEditingCommands.findDialogInitial?.nonBlocking !== true ||
    homeEditingCommands.findResult?.label !== "A4" ||
    !homeEditingCommands.findResult?.status?.includes("A4") ||
    homeEditingCommands.findResult?.dialogStillOpen !== true ||
    homeEditingCommands.shortcutFind?.title !== "検索と置換" ||
    homeEditingCommands.shortcutFind?.activeTab !== "検索" ||
    homeEditingCommands.shortcutFind?.activeInput !== true ||
    homeEditingCommands.shortcutReplace?.activeTab !== "置換" ||
    homeEditingCommands.shortcutReplace?.replaceVisible !== true ||
    homeEditingCommands.shortcutGoTo?.title !== "ジャンプ" ||
    homeEditingCommands.shortcutGoTo?.activeInput !== true ||
    homeEditingCommands.goToDialogInitial?.title !== "ジャンプ" ||
    !homeEditingCommands.goToDialogInitial?.referenceLabel?.includes("参照先") ||
    homeEditingCommands.goToDialogInitial?.activeInput !== true ||
    homeEditingCommands.goToResult?.label !== "K4" ||
    !homeEditingCommands.goToResult?.status?.includes("K4") ||
    homeEditingCommands.specialDialogInitial?.title !== "選択オプション" ||
    !["コメント(M)", "定数(O)", "数式(F)", "空白セル(K)", "行内の相違(W)", "列内の相違(C)", "参照元(P)", "参照先(D)", "アクティブセル領域(R)", "可視セル(Y)", "最後のセル(S)", "条件付き書式(T)", "データの入力規則(V)", "同じ入力規則(A)", "オブジェクト(B)"].every((label) =>
      homeEditingCommands.specialDialogInitial?.labels?.includes(label),
    ) ||
    !["M", "O", "F", "K", "W", "C", "P", "D", "R", "Y", "S", "T", "V", "A", "B"].every((key) => homeEditingCommands.specialDialogInitial?.accessKeys?.includes(key)) ||
    !homeEditingCommands.specialBlanks?.label?.includes("I1") ||
    !homeEditingCommands.specialBlanks?.status?.includes("空白セル") ||
    (homeEditingCommands.specialBlanks?.frameCount || 0) < 2 ||
    !homeEditingCommands.specialRowDifferences?.label?.includes("R2") ||
    !homeEditingCommands.specialRowDifferences?.label?.includes("Q3") ||
    !homeEditingCommands.specialRowDifferences?.status?.includes("行内の相違") ||
    (homeEditingCommands.specialRowDifferences?.frameCount || 0) < 2 ||
    !homeEditingCommands.specialColumnDifferences?.label?.includes("Q6") ||
    !homeEditingCommands.specialColumnDifferences?.label?.includes("P7") ||
    !homeEditingCommands.specialColumnDifferences?.status?.includes("列内の相違") ||
    (homeEditingCommands.specialColumnDifferences?.frameCount || 0) < 2 ||
    homeEditingCommands.specialPrecedents?.label !== "S2:S3" ||
    !homeEditingCommands.specialPrecedents?.status?.includes("参照元") ||
    homeEditingCommands.specialDependents?.label !== "S4" ||
    !homeEditingCommands.specialDependents?.status?.includes("参照先") ||
    !homeEditingCommands.specialSameValidation?.label?.includes("N2") ||
    !homeEditingCommands.specialSameValidation?.label?.includes("N3") ||
    !homeEditingCommands.specialSameValidation?.status?.includes("同じ入力規則") ||
    homeEditingCommands.replaceDialogInitial?.activeTab !== "置換" ||
    homeEditingCommands.replaceDialogInitial?.replaceVisible !== true ||
    !homeEditingCommands.replaceDialogInitial?.labels?.some((label) => label.includes("置換後の文字列")) ||
    homeEditingCommands.replaceOne?.a4 !== "Delta" ||
    !homeEditingCommands.replaceOne?.status?.includes("置換") ||
    homeEditingCommands.replaceUndo?.a4 !== "Gamma" ||
    !homeEditingCommands.replaceUndo?.status?.includes("元に戻しました") ||
    homeEditingCommands.replaceAll?.o2 !== "find baz" ||
    homeEditingCommands.replaceAll?.o3 !== "baz baz" ||
    !homeEditingCommands.replaceAll?.status?.includes("2 件") ||
    !homeEditingCommands.selectFormulas?.label?.includes("F1") ||
    !homeEditingCommands.selectFormulas?.label?.includes("D5") ||
    homeEditingCommands.selectFormulas?.formula !== "=SUM(B2:B4)" ||
    !homeEditingCommands.selectFormulas?.status?.includes("数式") ||
    !(homeEditingCommands.selectNotes?.label === "G2" || homeEditingCommands.selectNotes?.message?.includes("見つかりません")) ||
    !homeEditingCommands.selectValidation?.label?.includes("N2") ||
    !homeEditingCommands.selectValidation?.label?.includes("N3") ||
    !homeEditingCommands.selectValidation?.status?.includes("データの入力規則")
  ) {
    throw new Error("Home editing commands should provide working AutoSUM, sort/filter menu, find/select menu, and undo/redo.");
  }


  if (
    !dataValidationMenu.labels.some((label) => label.includes("データの入力規則(V)...")) ||
    !dataValidationMenu.labels.some((label) => label.includes("無効データのマーク(C)")) ||
    !dataValidationMenu.labels.some((label) => label.includes("入力規則マークのクリア(R)")) ||
    !["V", "C", "R"].every((key) => dataValidationMenu.accessKeys.includes(key)) ||
    dataValidationMenu.focusedBefore !== "open" ||
    dataValidationMenu.focusedAfterDown !== "clear-validation" ||
    !dataValidationMenu.marked?.b1Class?.includes("has-validation-invalid") ||
    dataValidationMenu.marked?.a1Class?.includes("has-validation-invalid") ||
    !dataValidationMenu.marked?.status?.includes("2 個") ||
    dataValidationMenu.cleared?.b1Class?.includes("has-validation-invalid") ||
    !dataValidationMenu.cleared?.status?.includes("クリア") ||
    dataValidationMenu.dialogInitial?.allow !== "whole" ||
    dataValidationMenu.dialogInitial?.operatorHidden ||
    dataValidationMenu.dialogInitial?.sourceCaption !== "最小値" ||
    dataValidationMenu.dialogListMode?.operatorHidden !== true ||
    dataValidationMenu.dialogListMode?.sourceCaption !== "元の値" ||
    dataValidationMenu.dialogListMode?.source2Hidden !== true ||
    dataValidationMenu.dialogListMode?.dropdownHidden !== false ||
    dataValidationMenu.dialogAfterClearAll?.visible !== true ||
    dataValidationMenu.dialogAfterClearAll?.allow !== "any" ||
    dataValidationMenu.dialogAfterClearAll?.operatorHidden !== true ||
    dataValidationMenu.dialogAfterClearAll?.sourceHidden !== true ||
    dataValidationMenu.dialogAfterClearAll?.source !== "" ||
    dataValidationMenu.dialogAfterClearAll?.errorTitle !== "" ||
    dataValidationMenu.afterClearAllMark?.b1Class?.includes("has-validation-invalid") ||
    !dataValidationMenu.afterClearAllMark?.status?.includes("1 個") ||
    dataValidationMenu.customInvalid?.title !== "入力エラー" ||
    !dataValidationMenu.customInvalid?.message?.includes("3文字以内") ||
    dataValidationMenu.customInvalid?.buttons?.join("|") !== "再試行(R)|キャンセル" ||
    dataValidationMenu.customInvalid?.b2 !== "abcd" ||
    dataValidationMenu.customRetry?.formula !== "abcd" ||
    dataValidationMenu.customRetry?.editing !== true ||
    !dataValidationMenu.customRetry?.status?.includes("修正") ||
    dataValidationMenu.customValid?.b2 !== "abc" ||
    dataValidationMenu.listReferenceBefore?.options?.join("|") !== "Alpha|Beta|Gamma" ||
    dataValidationMenu.listReferenceAfter?.value !== "Beta" ||
    dataValidationMenu.listCaseInsensitive?.value !== "beta" ||
    dataValidationMenu.listCaseInsensitive?.messageVisible !== false ||
    dataValidationMenu.contextDropdownMenu?.disabled !== false ||
    !dataValidationMenu.contextDropdownMenu?.label?.includes("ドロップダウン") ||
    dataValidationMenu.contextDropdownPopup?.row !== "3" ||
    dataValidationMenu.contextDropdownPopup?.col !== "2" ||
    dataValidationMenu.contextDropdownPopup?.options?.join("|") !== "Alpha|Beta|Gamma" ||
    dataValidationMenu.contextDropdownPopup?.selected !== "Beta" ||
    dataValidationMenu.contextDropdownAfter?.value !== "Gamma" ||
    dataValidationMenu.contextDropdownAfter?.popupClosed !== true ||
    !dataValidationMenu.contextDropdownAfter?.status?.includes("Gamma") ||
    dataValidationMenu.contextDropdownKeyboardPopup?.selected !== "Gamma" ||
    dataValidationMenu.contextDropdownKeyboardPopup?.focused !== "Beta" ||
    dataValidationMenu.contextDropdownKeyboardPopup?.optionCount !== 3 ||
    dataValidationMenu.contextDropdownKeyboard?.value !== "Beta" ||
    dataValidationMenu.contextDropdownKeyboard?.popupClosed !== true ||
    dataValidationMenu.contextDropdownKeyboard?.selectionLabel !== "B3" ||
    !dataValidationMenu.contextDropdownKeyboard?.status?.includes("Beta") ||
    dataValidationMenu.contextDropdownDisabled !== true ||
    dataValidationMenu.warningInvalid?.title !== "注意" ||
    !dataValidationMenu.warningInvalid?.message?.includes("続行") ||
    dataValidationMenu.warningInvalid?.buttons?.join("|") !== "はい(Y)|いいえ(N)|キャンセル" ||
    dataValidationMenu.warningAccepted?.b4 !== "99" ||
    dataValidationMenu.informationInvalid?.title !== "情報" ||
    !dataValidationMenu.informationInvalid?.message?.includes("確認") ||
    dataValidationMenu.informationInvalid?.buttons?.join("|") !== "OK|キャンセル" ||
    dataValidationMenu.informationAccepted?.b5 !== "99" ||
    dataValidationMenu.noAlertAccepted?.b6 !== "99" ||
    !dataValidationMenu.noAlertAccepted?.b6Class?.includes("has-validation-invalid") ||
    dataValidationMenu.inputMessagePopup?.visible !== true ||
    dataValidationMenu.inputMessagePopup?.row !== "7" ||
    dataValidationMenu.inputMessagePopup?.col !== "2" ||
    dataValidationMenu.inputMessagePopup?.title !== "入力案内" ||
    !dataValidationMenu.inputMessagePopup?.text?.includes("1 から 10") ||
    !dataValidationMenu.inputMessagePopup?.text?.includes("範囲外") ||
    !dataValidationMenu.inputMessagePopup?.rightOfCell ||
    !dataValidationMenu.inputMessagePopup?.insideViewport ||
    dataValidationMenu.inputMessagePopup?.closedAfterOtherCell !== true ||
    dataValidationMenu.dateDialog?.sourceCaption !== "開始日" ||
    dataValidationMenu.dateDialog?.source2Caption !== "終了日" ||
    dataValidationMenu.dateInvalid?.buttons?.join("|") !== "再試行(R)|キャンセル" ||
    dataValidationMenu.dateInvalid?.b8 !== "2026/06/01" ||
    dataValidationMenu.dateValid?.b8 !== "2026/05/20" ||
    dataValidationMenu.timeDialog?.sourceCaption !== "開始時刻" ||
    dataValidationMenu.timeDialog?.source2Caption !== "終了時刻" ||
    dataValidationMenu.timeInvalid?.buttons?.join("|") !== "再試行(R)|キャンセル" ||
    dataValidationMenu.timeInvalid?.b9 !== "08:30" ||
    dataValidationMenu.timeValid?.b9 !== "09:30" ||
    dataValidationMenu.timeAmPmValid?.b9 !== "6 PM" ||
    dataValidationMenu.timeAmPmValid?.messageVisible !== false ||
    dataValidationMenu.textLengthDialog?.sourceCaption !== "値" ||
    dataValidationMenu.textLengthDialog?.source2Hidden !== true ||
    dataValidationMenu.textLengthInvalid?.buttons?.join("|") !== "再試行(R)|キャンセル" ||
    dataValidationMenu.textLengthInvalid?.b10 !== "123456" ||
    dataValidationMenu.textLengthValid?.b10 !== "12345" ||
    dataValidationMenu.blankDialog?.allowBlank !== false ||
    dataValidationMenu.blankInvalid?.buttons?.join("|") !== "再試行(R)|キャンセル" ||
    dataValidationMenu.blankInvalid?.b11 !== "" ||
    dataValidationMenu.blankRejected?.b11 !== "5" ||
    dataValidationMenu.formulaBoundValid?.b13 !== "9" ||
    dataValidationMenu.formulaBoundValid?.messageVisible !== false ||
    dataValidationMenu.formulaBoundInvalid?.buttons?.join("|") !== "再試行(R)|キャンセル" ||
    dataValidationMenu.formulaBoundInvalid?.b12 !== "6" ||
    dataValidationMenu.sameValidationInitial?.source2 !== "10" ||
    dataValidationMenu.sameValidationInitial?.disabled !== false ||
    !dataValidationMenu.sameValidationInitial?.label?.includes("同じ入力規則") ||
    !dataValidationMenu.sameValidationApplied?.status?.includes("同じ入力規則") ||
    dataValidationMenu.sameValidationApplied?.b15 !== "25" ||
    dataValidationMenu.sameValidationApplied?.buttons?.join("|") !== "再試行(R)|キャンセル" ||
    dataValidationMenu.listSemicolon?.options?.join("|") !== "Red|Green|Blue" ||
    dataValidationMenu.listSemicolon?.value !== "green" ||
    dataValidationMenu.listSemicolon?.messageVisible !== false ||
    dataValidationMenu.listLocalNameReference?.source !== "='入力規則'!候補リスト" ||
    dataValidationMenu.listLocalNameReference?.options?.join("|") !== "Alpha|Beta|Gamma" ||
    dataValidationMenu.listLocalNameReference?.localName?.scope !== "sheet" ||
    dataValidationMenu.listLocalNameReference?.localName?.sheetIndex !== 0 ||
    dataValidationMenu.listLocalNameReference?.value !== "Gamma" ||
    dataValidationMenu.listLocalNameReference?.popupClosed !== true ||
    !dataValidationMenu.listLocalNameReference?.status?.includes("Gamma") ||
    dataValidationMenu.listDynamicNameReference?.source !== "=動的候補" ||
    dataValidationMenu.listDynamicNameReference?.options?.join("|") !== "Alpha|Beta|Gamma" ||
    !dataValidationMenu.listDynamicNameReference?.dynamicName?.refersTo?.includes("OFFSET") ||
    dataValidationMenu.listDynamicNameReference?.value !== "Beta" ||
    dataValidationMenu.listDynamicNameReference?.popupClosed !== true ||
    !dataValidationMenu.listDynamicNameReference?.status?.includes("Beta") ||
    dataValidationMenu.listIndirectNameReference?.source !== "=INDIRECT($E$1)" ||
    dataValidationMenu.listIndirectNameReference?.driverValue !== "候補リスト" ||
    dataValidationMenu.listIndirectNameReference?.options?.join("|") !== "Alpha|Beta|Gamma" ||
    dataValidationMenu.listIndirectNameReference?.value !== "Alpha" ||
    dataValidationMenu.listIndirectNameReference?.popupClosed !== true ||
    !dataValidationMenu.listIndirectNameReference?.status?.includes("Alpha") ||
    dataValidationMenu.listIndirectCellRange?.source !== "=INDIRECT($E$2)" ||
    dataValidationMenu.listIndirectCellRange?.driverValue !== "'入力規則'!$D$1:$D$3" ||
    dataValidationMenu.listIndirectCellRange?.options?.join("|") !== "Alpha|Beta|Gamma" ||
    dataValidationMenu.listIndirectCellRange?.value !== "Beta" ||
    dataValidationMenu.listIndirectCellRange?.popupClosed !== true ||
    !dataValidationMenu.listIndirectCellRange?.status?.includes("Beta") ||
    dataValidationMenu.listIndexRangeReference?.source !== "=$D$1:INDEX($D:$D,COUNTA($D:$D))" ||
    dataValidationMenu.listIndexRangeReference?.options?.join("|") !== "Alpha|Beta|Gamma" ||
    dataValidationMenu.listIndexRangeReference?.value !== "Gamma" ||
    dataValidationMenu.listIndexRangeReference?.popupClosed !== true ||
    !dataValidationMenu.listIndexRangeReference?.status?.includes("Gamma") ||
    dataValidationMenu.listIndirectLiteralRange?.source !== "=INDIRECT(\"'入力規則'!$D$1:$D$3\")" ||
    dataValidationMenu.listIndirectLiteralRange?.options?.join("|") !== "Alpha|Beta|Gamma" ||
    dataValidationMenu.listIndirectLiteralRange?.value !== "Alpha" ||
    dataValidationMenu.listIndirectLiteralRange?.popupClosed !== true ||
    !dataValidationMenu.listIndirectLiteralRange?.status?.includes("Alpha") ||
    dataValidationMenu.listIndexColumnReference?.source !== "=INDEX($D$1:$D$3,0,1)" ||
    dataValidationMenu.listIndexColumnReference?.options?.join("|") !== "Alpha|Beta|Gamma" ||
    dataValidationMenu.listIndexColumnReference?.value !== "Beta" ||
    dataValidationMenu.listIndexColumnReference?.popupClosed !== true ||
    !dataValidationMenu.listIndexColumnReference?.status?.includes("Beta") ||
    dataValidationMenu.listIndexMatchReference?.source !== "=INDEX($F$2:$G$4,0,MATCH($E$3,$F$1:$G$1,0))" ||
    dataValidationMenu.listIndexMatchReference?.driverValue !== "部署2" ||
    dataValidationMenu.listIndexMatchReference?.options?.join("|") !== "Red|Blue|Yellow" ||
    dataValidationMenu.listIndexMatchReference?.value !== "Blue" ||
    dataValidationMenu.listIndexMatchReference?.popupClosed !== true ||
    !dataValidationMenu.listIndexMatchReference?.status?.includes("Blue") ||
    dataValidationMenu.importedHiddenDropdown?.value !== "Alpha" ||
    dataValidationMenu.importedHiddenDropdown?.hasSelect !== false ||
    dataValidationMenu.importedHiddenDropdown?.validation?.showDropDown !== false ||
    dataValidationMenu.importedHiddenDropdown?.validation?.imeMode !== "disabled" ||
    dataValidationMenu.importedHiddenDropdownDialog?.allow !== "list" ||
    dataValidationMenu.importedHiddenDropdownDialog?.showDropDown !== false ||
    dataValidationMenu.importedHiddenDropdownDialog?.dropdownHidden !== false ||
    dataValidationMenu.importedHiddenDropdownDialog?.imeMode !== "disabled" ||
    dataValidationMenu.exportedHiddenDropdown?.showDropDown !== true ||
    dataValidationMenu.exportedHiddenDropdown?.imeMode !== true ||
    dataValidationMenu.importedBlankRangeValidation?.rowCount < 31 ||
    dataValidationMenu.importedBlankRangeValidation?.colCount < 13 ||
    dataValidationMenu.importedBlankRangeValidation?.c30?.type !== "list" ||
    dataValidationMenu.importedBlankRangeValidation?.c30?.options?.join("|") !== "One|Two" ||
    dataValidationMenu.importedBlankRangeValidation?.d31?.type !== "list" ||
    dataValidationMenu.importedBlankRangeValidation?.e30?.type !== "whole" ||
    dataValidationMenu.importedBlankRangeValidation?.e30?.formulae?.[0] !== "=A12" ||
    dataValidationMenu.importedBlankRangeValidation?.f30?.type !== "custom" ||
    dataValidationMenu.importedBlankRangeValidation?.f30?.operator !== "" ||
    dataValidationMenu.importedBlankRangeValidation?.f30?.formulae?.[0] !== "=LEN(F30)<=3" ||
    dataValidationMenu.importedBlankRangeValidation?.h30?.options?.join("|") !== "East|West" ||
    dataValidationMenu.importedBlankRangeValidation?.h32?.options?.join("|") !== "East|West" ||
    dataValidationMenu.importedBlankRangeValidation?.i33?.options?.join("|") !== "East|West" ||
    dataValidationMenu.importedBlankRangeValidation?.j30?.type !== "decimal" ||
    dataValidationMenu.importedBlankRangeValidation?.j30?.operator !== "greaterThan" ||
    dataValidationMenu.importedBlankRangeValidation?.j30?.formulae?.[0] !== "1.5" ||
    dataValidationMenu.importedBlankRangeValidation?.j30?.showInputMessage !== true ||
    dataValidationMenu.importedBlankRangeValidation?.j30?.inputTitle !== "小数案内" ||
    dataValidationMenu.importedBlankRangeValidation?.j30?.errorStyle !== "warning" ||
    dataValidationMenu.importedBlankRangeValidation?.j30?.errorTitle !== "小数エラー" ||
    dataValidationMenu.importedBlankRangeValidation?.k30?.type !== "date" ||
    dataValidationMenu.importedBlankRangeValidation?.k30?.formulae?.join("|") !== "45413|45443" ||
    dataValidationMenu.importedBlankRangeValidation?.l30?.type !== "time" ||
    dataValidationMenu.importedBlankRangeValidation?.l30?.formulae?.join("|") !== "0.375|0.75" ||
    dataValidationMenu.importedBlankRangeValidation?.m30?.type !== "textLength" ||
    dataValidationMenu.importedBlankRangeValidation?.m30?.formulae?.[0] !== "5" ||
    dataValidationMenu.exportedBlankRangeValidation?.hasRange !== true ||
    dataValidationMenu.exportedBlankRangeValidation?.hasFormula !== true ||
    dataValidationMenu.exportedBlankRangeValidation?.hasReferenceFormula !== true ||
    dataValidationMenu.exportedBlankRangeValidation?.hasCustomFormula !== true ||
    !dataValidationMenu.exportedBlankRangeValidation?.multiRangeTag?.includes('sqref="H30 H32:I33"') ||
    !dataValidationMenu.exportedBlankRangeValidation?.decimalTag?.includes('promptTitle="小数案内"') ||
    !dataValidationMenu.exportedBlankRangeValidation?.decimalTag?.includes('errorStyle="warning"') ||
    !dataValidationMenu.exportedBlankRangeValidation?.dateTag ||
    !dataValidationMenu.exportedBlankRangeValidation?.timeTag ||
    !dataValidationMenu.exportedBlankRangeValidation?.textLengthTag
  ) {
    throw new Error("Data validation should match Excel for menu commands, formulas, list references, input messages, alert styles, and imported OOXML details.");
  }
  Object.assign(scenarioState, { result, coordinateHitTest, delayedLazyLoad, stickyHeaders, sheetMenu, sheetTabColor, sheetRename, sheetTabDirectActions, sheetMoveCopy, sheetCodeDialog, sheetSelectAll, sheetHidden, sheetUnhideDialog, sheetUnhidden, sheetUnhideUndo, sheetUnhideRedo, sheetGroupActions, sheetGroupMoveCopy, sheetGroupCellEditing, sheetGroupStructure, sheetGroupFormatDialog, sheetGroupDataValidation, sheetGroupDataFilterSort, sheetGroupDataTools, sheetGroupHyperlink, sheetGroupHomeCommands, sheetGroupTranslation, sheetGroupQuickAnalysis, sheetGroupSparklines, sheetGroupCharts, sheetGroupTables, sheetGroupPivot, sheetGroupObjects, insertRibbon, formulaRibbon, formulaDefineNameMain, formulaDefineNameMenu, formulaDefineNameCreate, createNamesDialog, formulaNames, formulaNameManager, formulaUseNameMenu, formulaUseNameManage, formulaUseNameInserted, formulaPasteNameDialog, formulaPasteNameInserted, formulaPasteNameListDialog, formulaPasteNameList, formulaPasteNameListUndo, formulaPasteNameListRedo, formulaFunctionDialog, formulaFunctionSearch, formulaInserted, formulaFunctionArguments, formulaLogicalMenu, formulaMenuArguments, formulaMenuInserted, formulaRecentMenu, formulaMathMenu, formulaVariableArguments, formulaZeroArgumentFunction, formulaFinancialMenu, formulaDateMenu, formulaTextMenu, formulaMoreMenu, formulaShowState, formulaTrace, formulaRemoveArrowsMenu, formulaRemovePrecedents, formulaEvaluateInitial, formulaEvaluateStepIn, formulaEvaluateStep, formulaEvaluateStepOut, formulaEvaluateReferenceInitial, formulaEvaluateReferenceStep, formulaEvaluateReferenceStepIn, formulaEvaluateReferenceStepOut, formulaWatchWindow, formulaWatchSelect, formulaWatchDelete, formulaErrorMenu, formulaErrorChecking, formulaErrorIgnore, formulaErrorTrace, formulaErrorResetIgnored, formulaCalculation, viewRibbon, viewToggles, viewWindowCommands, layoutRibbon, layoutThemeMenuInitial, layoutThemeApplied, layoutColorMenuInitial, layoutFontMenuInitial, layoutEffectMenuInitial, layoutThemeAfterMenus, layoutThemeEffectUndoRedo, layoutPrintAreaMenuBefore, layoutPrintAreaMenuAfterSet, layoutPrintAreaMenuAfterClear, layoutMarginsMenuInitial, layoutOrientationMenuInitial, layoutSizeMenuInitial, layoutBackgroundApplied, layoutBackgroundUndo, layoutBackgroundRedo, layoutBackgroundDeleted, layoutBackgroundDeleteUndo, pageSetupDialog, printTitlesDialogInitial, printTitlesDialogFilled, pageLayout, pageLayoutPreview, pageBreakMenuKeyboard, pageBreakDrag, drawRibbon, drawInk, drawConverted, headerFooterDialog, headerFooterApplied, headerFooterUndo, headerFooterRedo, insertMapChart, insertShapeMenu, insertShapeDraw, insertIconDialog, signatureLineDialog, insertObjectDialog, insertAddinsStoreDialog, insertAddinsMineDialog, insertTextObjects, fileSideMenu, fileSideMenuClosedByOutsideClick, fileSideMenuClosedByCommand, shortcutSave, exportedHtmlChecks, saveAsDialog, expectedReloadWorkbookName, standaloneBlockedRequests, standaloneOffline, standaloneRemoteFallback, standaloneEmptyModel, standaloneMalformedModel, standaloneWrappedModel, launcherBlockedRequests, standaloneLauncher, ar2Gridline, imageAndFont, imageSelection, footerLayout, realCtrlWheelZoom, zoomPopup, footerBeforeScroll, tabScroll, singleClickEdit, doubleClickEdit, editAfterDeferredRender, editAlignment, editClickCaret, editTextSelection, editNoEllipsis, cellEditAltEnter, cellEditAltEnterDuringIme, editDragOutsideSelection, cellEditImeEnter, editScrollbar, editExitRender, multilineCommitDisplay, formulaBarInitial, formulaBarSizing, formulaBarDuringEdit, formulaBarEscape, formulaBarEnter, formulaAutocomplete, formulaReferenceInsert, directTyping, definedNameAndLink, statusBarStats, autoFillSeries, insertDeleteDialog, shortcutUndo, formatUndoRedo, toolbarSync, selectionFrame, copyStateAfterCopy, shortcutClipboard, copyStateWithTarget, richPaste, matchDestinationPaste, pasteSpecialFormats, pasteSpecialAdvanced, formulaPasteReferences, cutAndRepeatPaste, cellMenu, cellAnnotations, furigana, formatCellsDialog, cellMenuClear, cellMenuUndo, deleteKeyClear, deleteKeyUndo, backspaceClearEdit, backspaceUndo, popupDismissal, newSheet, colorPalette, headerSelection, selectionAutoScroll, keyboardNavigation, column, row, resizedHeaders, resizeHandleVisual, headerResizeUndoRedo, multiHeaderResize, lazyLoadAnchor, extension, trimmedExtension, whiteFillGridMask, mergeUndoRedo, mergeAutoFillMismatch, underlineToolbar, borderToolbar, bottomBar, ribbon, mergedSingleSelection, mergedKeyboardNavigation, mergedCopyState, mergedPaste, mergedCrossSelection, persistenceReload, wideViewportCoverage, importedNativeShapes, customViewImportExport, autoFilterConditionMenu, importedAutoFilterCriteria, autoFilterBlankKeyboardMenu, autoFilterReapplyClear, advancedFilterCopyOverwrite, autoFilterDateMenu, advancedFilterUniqueDates, advancedFilterUniqueSubsetCopy, delimitedTextImport, powerQueryTransformEditor, whatIfGoalSeekDataTable, scenarioManagerSummary, flashFillDataPatterns, dataCurrentRegion, queryConnectionsPane, dataGetConnectionMenus, relationshipsDialog, refreshAllUndo, forecastSheet, dataAnalysisToolpak, consolidateDialog, subtotalDialog, outlineGrouping, textToColumnsDialog, textToColumnsFixedWidth, textToColumnsAdvanced, sortDialogLevels, sortDialogColorOrder, sortDialogLeftToRight, sortDialogBlankAndCase, sortDialogCustomList, sortDialogDateOrder, sortFormulaResults, quickSortActiveColumn, quickSortPartialSelection, removeDuplicatesDialog });
}

export {
  runExportAndResults,
};
