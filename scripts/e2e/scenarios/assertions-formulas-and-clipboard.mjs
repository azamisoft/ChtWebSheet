// Sequential E2E feature module.
import {
  path,
} from "../../e2e-real-upload.mjs";
async function runAssertionsFormulasAndClipboard(scenarioState) {
  let { result, runtimeSource, exportedHtmlChecks, expectedReloadWorkbookName, standaloneBlockedRequests, standaloneOffline, standaloneRemoteFallback, standaloneEmptyModel, standaloneMalformedModel, standaloneWrappedModel, launcherBlockedRequests, standaloneLauncher, ar2Gridline, imageAndFont, imageSelection, footerLayout, realCtrlWheelZoom, zoomPopup, footerBeforeScroll, tabScroll, singleCellStatusBar, statusBarStats, autoFillSeries, insertDeleteDialog, shortcutUndo, wrapToolbar, formatUndoRedo, painterActiveAfterStart, toolbarSync, selectionFrame, copyStateAfterCopy, shortcutClipboard, copyStateWithTarget, richPaste, matchDestinationPaste, pasteSpecialFormats, pasteSpecialAdvanced, formulaPasteReferences, cutAndRepeatPaste, cellMenu, translationPaneInitial, translationPaneHeader, translationReplace, translationUndo, commentContextMenu, commentReplyFromMenu, memoContextMenu, memoEdit, memoToggleShown, memoToggleHidden, memoDelete, memoDeleteUndo, memoDeleteRedo, commentDeleteContextMenu, commentDelete, commentDeleteUndo, commentDeleteRedo, cellAnnotations, furigana, formatCellsDialog, cellMenuClear, cellMenuUndo, deleteKeyClear, deleteKeyUndo, backspaceClearEdit, backspaceUndo, popupDismissal, newSheet, colorPalette, headerSelection, selectionAutoScroll, keyboardNavigation, column, row, resizedHeaders, resizeHandleVisual, headerResizeUndoRedo, multiHeaderResize, lazyLoadAnchor, extension, trimmedExtension, whiteFillGridMask, mergeUndoRedo, mergeAutoFillMismatch, bottomBar, ribbon, mergedSingleSelection, mergedKeyboardNavigation, mergedCopyState, mergedPaste, mergedCrossSelection, importedNativeShapes } = scenarioState;

  if (
    autoFillSeries.label !== "A4:A7" ||
    autoFillSeries.a6 !== "3" ||
    autoFillSeries.a7 !== "4" ||
    !autoFillSeries.handleColor?.includes("34, 197, 94") ||
    !autoFillSeries.status?.includes("オートフィル") ||
    autoFillSeries.undoA6 === "3" ||
    autoFillSeries.undoA7 === "4" ||
    !autoFillSeries.undoStatus?.includes("元に戻しました") ||
    autoFillSeries.redoA6 !== "3" ||
    autoFillSeries.redoA7 !== "4" ||
    !autoFillSeries.redoStatus?.includes("やり直しました") ||
    !["セルのコピー(C)", "連続データ(S)", "書式のみコピー(F)", "書式なしコピー(O)", "線形近似(L)", "成長曲線(G)"].every((label) =>
      autoFillSeries.optionsMenu?.labels?.includes(label),
    ) ||
    autoFillSeries.optionsMenu?.checked !== "series" ||
    !["C", "S", "F", "O", "L", "G"].every((key) => autoFillSeries.optionsMenu?.accessKeys?.includes(key)) ||
    autoFillSeries.optionsMenu?.focusedBefore !== "series" ||
    autoFillSeries.optionsMenu?.focusedAfterDown !== "formatting" ||
    autoFillSeries.copyOption?.a6 !== "1" ||
    autoFillSeries.copyOption?.a7 !== "2" ||
    !autoFillSeries.copyOption?.status?.includes("セルのコピー") ||
    autoFillSeries.formattingOption?.a6 !== "" ||
    autoFillSeries.formattingOption?.a7 !== "" ||
    !autoFillSeries.formattingOption?.status?.includes("書式のみコピー") ||
    autoFillSeries.noFormattingOption?.a6 !== "3" ||
    autoFillSeries.noFormattingOption?.a7 !== "4" ||
    !autoFillSeries.noFormattingOption?.status?.includes("書式なしコピー") ||
    autoFillSeries.seriesOption?.a6 !== "3" ||
    autoFillSeries.seriesOption?.a7 !== "4" ||
    !autoFillSeries.seriesOption?.status?.includes("連続データ") ||
    autoFillSeries.formulaShiftRight !== "=SUM(C:C)+SUM(2:4)+D3" ||
    autoFillSeries.formulaShiftDown !== "=SUM(B:B)+SUM(3:5)+C4" ||
    autoFillSeries.absoluteFormulaShift !== "=SUM($B:$B)+SUM($2:$4)+$C$3" ||
    autoFillSeries.textNumber?.a9 !== "項目002" ||
    autoFillSeries.textNumber?.a10 !== "項目003" ||
    autoFillSeries.textNumber?.a11 !== "項目004" ||
    autoFillSeries.textNumber?.c13 !== "R-002" ||
    autoFillSeries.textNumber?.d13 !== "R-003" ||
    autoFillSeries.textNumber?.e13 !== "R-004" ||
    autoFillSeries.weekdays?.a13 !== "火曜日" ||
    autoFillSeries.weekdays?.a14 !== "水曜日" ||
    autoFillSeries.weekdays?.a15 !== "木曜日" ||
    autoFillSeries.months?.a17 !== "2月" ||
    autoFillSeries.months?.a18 !== "3月" ||
    autoFillSeries.months?.a19 !== "4月" ||
    autoFillSeries.dateText?.a2 !== "2026/6/2" ||
    autoFillSeries.dateText?.a3 !== "2026/6/3" ||
    autoFillSeries.dateStep?.d6 !== "2026/6/5" ||
    autoFillSeries.dateStep?.e6 !== "2026/6/7" ||
    autoFillSeries.timeText?.c6 !== "10:00" ||
    autoFillSeries.timeText?.d6 !== "11:00" ||
    autoFillSeries.japaneseDate?.c6 !== "2026年6月2日" ||
    autoFillSeries.japaneseDate?.d6 !== "2026年6月3日" ||
    autoFillSeries.singleNumberDefault?.aq8 !== "10" ||
    autoFillSeries.singleNumberDefault?.aq9 !== "10" ||
    autoFillSeries.singleNumberDefault?.checked !== "copy" ||
    autoFillSeries.singleNumberDefault?.focused !== "copy" ||
    autoFillSeries.singleNumberSeries?.aq8 !== "11" ||
    autoFillSeries.singleNumberSeries?.aq9 !== "12" ||
    !autoFillSeries.singleNumberSeries?.status?.includes("連続データ") ||
    autoFillSeries.controlDrag?.ar8 !== "11" ||
    autoFillSeries.controlDrag?.ar9 !== "12" ||
    autoFillSeries.doubleClick?.label !== "AS16:AS22" ||
    autoFillSeries.doubleClick?.as18 !== "3" ||
    autoFillSeries.doubleClick?.as22 !== "7" ||
    !autoFillSeries.dateOptionsMenu?.labels?.includes("日単位(D)") ||
    !autoFillSeries.dateOptionsMenu?.labels?.includes("週日単位(W)") ||
    !autoFillSeries.dateOptionsMenu?.labels?.includes("月単位(M)") ||
    !autoFillSeries.dateOptionsMenu?.labels?.includes("年単位(Y)") ||
    !["D", "W", "M", "Y"].every((key) => autoFillSeries.dateOptionsMenu?.accessKeys?.includes(key)) ||
    autoFillSeries.dateOptionsMenu?.focused !== "series" ||
    autoFillSeries.dateMonth?.at8 !== "2026/2/28" ||
    autoFillSeries.dateMonth?.at9 !== "2026/3/31" ||
    !autoFillSeries.dateMonth?.status?.includes("月単位") ||
    autoFillSeries.dateWeekday?.au8 !== "2026/6/8" ||
    autoFillSeries.dateWeekday?.au9 !== "2026/6/9" ||
    !autoFillSeries.dateWeekday?.status?.includes("週日単位") ||
    autoFillSeries.dateYear?.av8 !== "2027/6/1" ||
    autoFillSeries.dateYear?.av9 !== "2028/6/1" ||
    !autoFillSeries.dateYear?.status?.includes("年単位") ||
    autoFillSeries.rightDragMenu?.aw8Before !== "" ||
    autoFillSeries.rightDragMenu?.aw9Before !== "" ||
    autoFillSeries.rightDragMenu?.checked !== "copy" ||
    autoFillSeries.rightDragMenu?.focused !== "copy" ||
    autoFillSeries.rightDragMenu?.label !== "AW7:AW9" ||
    !autoFillSeries.rightDragMenu?.labels?.includes("セルのコピー(C)") ||
    !autoFillSeries.rightDragMenu?.labels?.includes("連続データ(S)") ||
	    autoFillSeries.rightDragSeries?.aw8 !== "21" ||
	    autoFillSeries.rightDragSeries?.aw9 !== "22" ||
	    !autoFillSeries.rightDragSeries?.menuClosed ||
	    !autoFillSeries.rightDragSeries?.status?.includes("連続データ") ||
	    !autoFillSeries.linearTrendMenu?.labels?.includes("線形近似(L)") ||
	    !autoFillSeries.linearTrendMenu?.accessKeys?.includes("L") ||
	    autoFillSeries.linearTrend?.ax9 !== "6" ||
	    autoFillSeries.linearTrend?.ax10 !== "8" ||
	    !autoFillSeries.linearTrend?.menuClosed ||
	    !autoFillSeries.linearTrend?.status?.includes("線形近似") ||
	    !autoFillSeries.growthTrendMenu?.labels?.includes("成長曲線(G)") ||
	    !autoFillSeries.growthTrendMenu?.accessKeys?.includes("G") ||
	    autoFillSeries.growthTrend?.ay9 !== "8" ||
	    autoFillSeries.growthTrend?.ay10 !== "16" ||
	    !autoFillSeries.growthTrend?.menuClosed ||
	    !autoFillSeries.growthTrend?.status?.includes("成長曲線") ||
	    autoFillSeries.quarter?.n13 !== "Q2" ||
    autoFillSeries.quarter?.o13 !== "Q3" ||
    autoFillSeries.quarter?.p13 !== "Q4" ||
    autoFillSeries.reverseVertical?.aq10 !== "3" ||
    autoFillSeries.reverseVertical?.aq11 !== "4" ||
    autoFillSeries.reverseHorizontal?.k14 !== "3" ||
    autoFillSeries.reverseHorizontal?.l14 !== "4"
  ) {
    throw new Error("The Excel-like fill handle should extend numeric, date, list, quarter, reverse, and text-number series with undo/redo.");
  }


  const insertOptions = insertDeleteDialog.insert.options.map((item) => `${item.value}:${item.label}`).join("|");


  const deleteOptions = insertDeleteDialog.delete.options.map((item) => `${item.value}:${item.label}`).join("|");


  if (
    insertDeleteDialog.insert.title !== "挿入" ||
    insertOptions !== "shift-right:セルを右にシフト|shift-down:セルを下にシフト|entire-row:行全体|entire-column:列全体" ||
    !insertDeleteDialog.insert.options.find((item) => item.value === "shift-right")?.checked ||
    insertDeleteDialog.insert.ok !== "OK" ||
    insertDeleteDialog.insert.cancel !== "キャンセル" ||
    insertDeleteDialog.afterInsert.k6 !== "" ||
    insertDeleteDialog.afterInsert.l6 !== "InsertA" ||
    insertDeleteDialog.afterInsert.m6 !== "InsertB" ||
    insertDeleteDialog.delete.title !== "削除" ||
    deleteOptions !== "shift-left:セルを左にシフト|shift-up:セルを上にシフト|entire-row:行全体|entire-column:列全体" ||
    !insertDeleteDialog.delete.options.find((item) => item.value === "shift-left")?.checked ||
    insertDeleteDialog.afterDelete.k6 !== "InsertA" ||
    insertDeleteDialog.afterDelete.l6 !== "InsertB" ||
    insertDeleteDialog.undoDelete.k6 !== "" ||
    insertDeleteDialog.undoDelete.l6 !== "InsertA" ||
    insertDeleteDialog.undoDelete.m6 !== "InsertB" ||
    insertDeleteDialog.redoDelete.k6 !== "InsertA" ||
    insertDeleteDialog.redoDelete.l6 !== "InsertB"
  ) {
    throw new Error("Cell context menu insert/delete dialogs should match Excel options, shift cells, and participate in undo/redo.");
  }


  if (
    keyboardNavigation.afterRight !== "C2" ||
    keyboardNavigation.afterDown !== "C3" ||
    keyboardNavigation.afterLeft !== "B3" ||
    keyboardNavigation.afterUp !== "B2" ||
    keyboardNavigation.afterEnter !== "B3" ||
    keyboardNavigation.afterShiftEnter !== "B2" ||
    keyboardNavigation.afterF2?.label !== "B3" ||
    keyboardNavigation.afterF2?.contenteditable !== "true" ||
    keyboardNavigation.afterEscape?.contenteditable !== null ||
    keyboardNavigation.contextMenuShortcut?.label !== "B2" ||
    !keyboardNavigation.contextMenuShortcut?.ariaLabel?.includes("B2") ||
    keyboardNavigation.contextMenuShortcut?.focusedAction !== "cut" ||
    keyboardNavigation.contextMenuShortcut?.hasFormat !== true ||
    keyboardNavigation.contextMenuShortcut?.accessKeysUnderlined !== true ||
    keyboardNavigation.f6Cycle?.afterWorksheet !== "selectedCellLabel" ||
    !keyboardNavigation.f6Cycle?.statusAfterWorksheet?.includes("名前ボックス") ||
    keyboardNavigation.f6Cycle?.afterName !== "formulaInput" ||
    !keyboardNavigation.f6Cycle?.statusAfterName?.includes("数式バー") ||
    keyboardNavigation.f6Cycle?.afterShiftFromWorksheet !== "home" ||
    !keyboardNavigation.f6Cycle?.statusAfterShiftFromWorksheet?.includes("ホーム") ||
    keyboardNavigation.f6Cycle?.afterRibbonForward !== "sheetHost" ||
    !keyboardNavigation.f6Cycle?.statusAfterRibbonForward?.includes("ワークシート") ||
    keyboardNavigation.sheetTabKeyboard?.visibleCount < 2 ||
    keyboardNavigation.sheetTabKeyboard?.focusedAfterArrow !== keyboardNavigation.sheetTabKeyboard?.expectedFocus ||
    keyboardNavigation.sheetTabKeyboard?.activeBeforeEnter !== keyboardNavigation.sheetTabKeyboard?.beforeActiveIndex ||
    keyboardNavigation.sheetTabKeyboard?.activeAfterEnter !== keyboardNavigation.sheetTabKeyboard?.expectedFocus ||
    keyboardNavigation.sheetTabKeyboard?.focusedAfterEnter !== keyboardNavigation.sheetTabKeyboard?.expectedFocus ||
    keyboardNavigation.sheetTabKeyboard?.menuOpen !== true ||
    keyboardNavigation.sheetTabKeyboard?.focusedAction !== "insert" ||
    !keyboardNavigation.sheetTabKeyboard?.accessKeysUnderlined ||
    keyboardNavigation.sheetTabKeyboard?.menuSheetIndex !== keyboardNavigation.sheetTabKeyboard?.expectedFocus ||
    keyboardNavigation.sheetTabKeyboard?.restoredActiveIndex !== keyboardNavigation.sheetTabKeyboard?.beforeActiveIndex ||
    keyboardNavigation.afterTab !== "F70" ||
    keyboardNavigation.afterShiftTab !== "E70" ||
    keyboardNavigation.afterCtrlRight !== "G70" ||
    keyboardNavigation.afterCtrlLeft !== "E70" ||
    keyboardNavigation.afterCtrlDown !== "E71" ||
    keyboardNavigation.afterHome !== "A71" ||
    keyboardNavigation.afterEnd !== "G71" ||
    keyboardNavigation.afterCtrlHome !== "A1" ||
    !/^A\d+$/.test(keyboardNavigation.afterPageDown || "") ||
    keyboardNavigation.afterPageDownRow <= 70 ||
    !/^A\d+$/.test(keyboardNavigation.afterPageUp || "") ||
    keyboardNavigation.afterPageUpRow >= keyboardNavigation.afterPageDownRow ||
    keyboardNavigation.afterShiftRight !== "E70:F70" ||
    keyboardNavigation.afterShiftDown !== "E70:F71" ||
    keyboardNavigation.afterCtrlShiftRight !== "E70:G70" ||
    keyboardNavigation.afterCtrlARegion !== "E70:G71" ||
    keyboardNavigation.afterCtrlAWorksheet?.top !== 1 ||
    keyboardNavigation.afterCtrlAWorksheet?.left !== 1 ||
    keyboardNavigation.afterCtrlAWorksheet?.bottom < 71 ||
    keyboardNavigation.afterCtrlAWorksheet?.right < 7 ||
    keyboardNavigation.afterCtrlShiftSpace !== "E70:G71" ||
    keyboardNavigation.afterCtrlAsterisk !== "E70:G71" ||
    keyboardNavigation.rowDifferencesShortcut?.ranges?.length !== 2 ||
    !keyboardNavigation.rowDifferencesShortcut?.ranges?.some((range) => range.top === 73 && range.left === 6 && range.bottom === 73 && range.right === 6) ||
    !keyboardNavigation.rowDifferencesShortcut?.ranges?.some((range) => range.top === 74 && range.left === 7 && range.bottom === 74 && range.right === 7) ||
    keyboardNavigation.columnDifferencesShortcut?.ranges?.length !== 3 ||
    keyboardNavigation.columnDifferencesShortcut?.ranges?.some((range) => range.top !== 74 || range.bottom !== 74 || range.left !== range.right || range.left < 5 || range.left > 7) ||
    keyboardNavigation.afterCtrlSpace !== "F:F" ||
    keyboardNavigation.afterShiftSpace !== "70:70" ||
    keyboardNavigation.afterHiddenColRight !== "G70" ||
    keyboardNavigation.afterHiddenRowDown !== "E72" ||
    keyboardNavigation.sheetSwitch?.afterNext === keyboardNavigation.sheetSwitch?.before ||
    keyboardNavigation.sheetSwitch?.afterPrevious !== keyboardNavigation.sheetSwitch?.before ||
    keyboardNavigation.sheetGroupShortcut?.visibleCount < 2 ||
    keyboardNavigation.sheetGroupShortcut?.afterPreviousActiveIndex !== keyboardNavigation.sheetGroupShortcut?.expectedPreviousIndex ||
    !keyboardNavigation.sheetGroupShortcut?.afterPreviousSelected?.includes(keyboardNavigation.sheetGroupShortcut?.beforeIndex) ||
    !keyboardNavigation.sheetGroupShortcut?.afterPreviousSelected?.includes(keyboardNavigation.sheetGroupShortcut?.expectedPreviousIndex) ||
    !keyboardNavigation.sheetGroupShortcut?.afterPreviousStatus?.includes("表示ワークシート") ||
    keyboardNavigation.sheetGroupShortcut?.afterNextActiveIndex !== keyboardNavigation.sheetGroupShortcut?.expectedNextIndex ||
    !keyboardNavigation.sheetGroupShortcut?.afterNextSelected?.includes(keyboardNavigation.sheetGroupShortcut?.beforeIndex) ||
    !keyboardNavigation.sheetGroupShortcut?.afterNextSelected?.includes(keyboardNavigation.sheetGroupShortcut?.expectedPreviousIndex) ||
    !keyboardNavigation.sheetGroupShortcut?.afterNextStatus?.includes("表示ワークシート") ||
    keyboardNavigation.sheetGroupShortcut?.restoredName !== keyboardNavigation.sheetGroupShortcut?.beforeName ||
    keyboardNavigation.sheetGroupShortcut?.restoredSelectedCount !== 0 ||
    keyboardNavigation.shortcutFormat?.bold !== "700" ||
    keyboardNavigation.shortcutFormat?.italic !== "italic" ||
    !String(keyboardNavigation.shortcutFormat?.decoration || "").includes("underline") ||
    !String(keyboardNavigation.shortcutFormat?.decoration || "").includes("line-through") ||
    keyboardNavigation.formatAliasShortcuts?.bold !== "700" ||
    keyboardNavigation.formatAliasShortcuts?.italic !== "italic" ||
    !String(keyboardNavigation.formatAliasShortcuts?.decoration || "").includes("underline") ||
    keyboardNavigation.formatDialogTitle !== "セルの書式設定" ||
    !["", null].includes(keyboardNavigation.numberFormatShortcuts?.general) ||
    keyboardNavigation.numberFormatShortcuts?.number !== "#,##0.00" ||
    keyboardNavigation.numberFormatShortcuts?.time !== "h:mm" ||
    keyboardNavigation.numberFormatShortcuts?.date !== "yyyy/m/d" ||
    keyboardNavigation.numberFormatShortcuts?.currency !== "¥#,##0.00" ||
    keyboardNavigation.numberFormatShortcuts?.percent !== "0%" ||
    keyboardNavigation.numberFormatShortcuts?.scientific !== "0.00E+00" ||
    keyboardNavigation.borderShortcutApplied?.topLeft?.top !== true ||
    keyboardNavigation.borderShortcutApplied?.topLeft?.left !== true ||
    keyboardNavigation.borderShortcutApplied?.topRight?.top !== true ||
    keyboardNavigation.borderShortcutApplied?.topRight?.right !== true ||
    keyboardNavigation.borderShortcutApplied?.bottomLeft?.bottom !== true ||
    keyboardNavigation.borderShortcutApplied?.bottomLeft?.left !== true ||
    keyboardNavigation.borderShortcutApplied?.bottomRight?.bottom !== true ||
    keyboardNavigation.borderShortcutApplied?.bottomRight?.right !== true ||
    [keyboardNavigation.borderShortcutRemoved?.topLeft, keyboardNavigation.borderShortcutRemoved?.topRight, keyboardNavigation.borderShortcutRemoved?.bottomLeft, keyboardNavigation.borderShortcutRemoved?.bottomRight].some((borders) => Object.values(borders || {}).some(Boolean)) ||
    keyboardNavigation.insertShortcutDialogTitle !== "挿入" ||
    keyboardNavigation.deleteShortcutDialogTitle !== "削除" ||
    !keyboardNavigation.formulaBarShortcut?.beforeRows ||
    keyboardNavigation.formulaBarShortcut?.afterRows === keyboardNavigation.formulaBarShortcut?.beforeRows ||
    keyboardNavigation.formulaBarShortcut?.restoredRows !== keyboardNavigation.formulaBarShortcut?.beforeRows ||
    keyboardNavigation.objectsShortcutHidden !== false ||
    keyboardNavigation.objectsShortcutVisible !== true ||
    keyboardNavigation.outlineShortcutHidden !== false ||
    keyboardNavigation.outlineShortcutVisible !== true ||
    keyboardNavigation.altF1Chart?.countDelta !== 1 ||
    keyboardNavigation.altF1Chart?.activeSheetName !== keyboardNavigation.altF1ChartBefore?.sheetName ||
    keyboardNavigation.altF1Chart?.chartType !== "column" ||
    keyboardNavigation.altF1Chart?.sourceRange !== "U86:W89" ||
    keyboardNavigation.altF1Chart?.selectedImageFrame !== true ||
    keyboardNavigation.f11ChartSheet?.tabCountDelta !== 1 ||
    keyboardNavigation.f11ChartSheet?.chartSheet !== true ||
    keyboardNavigation.f11ChartSheet?.showGridLines !== false ||
    keyboardNavigation.f11ChartSheet?.showHeadings !== false ||
    keyboardNavigation.f11ChartSheet?.imageCount !== 1 ||
    keyboardNavigation.f11ChartSheet?.chartType !== "column" ||
    keyboardNavigation.f11ChartSheet?.sourceSheet !== keyboardNavigation.f11ChartSheet?.previousSheetName ||
    keyboardNavigation.f11ChartSheet?.sourceRange !== "U86:W89" ||
    keyboardNavigation.f11ChartSheet?.selectedImageFrame !== true ||
    !keyboardNavigation.f11ChartSheet?.status?.includes("グラフ シート") ||
    keyboardNavigation.f11ChartSheetUndo?.activeSheetName !== keyboardNavigation.f11ChartSheet?.previousSheetName ||
    keyboardNavigation.f11ChartSheetUndo?.tabCount !== keyboardNavigation.f11ChartSheetUndo?.expectedTabCount ||
    !keyboardNavigation.f11ChartSheetUndo?.status?.includes("元に戻しました") ||
    keyboardNavigation.shiftF11Worksheet?.tabCountDelta !== 1 ||
    keyboardNavigation.shiftF11Worksheet?.selectedLabel !== "A1" ||
    !keyboardNavigation.shiftF11Worksheet?.status?.includes("新しいワークシート") ||
    keyboardNavigation.shiftF11WorksheetUndo?.tabCount !== keyboardNavigation.shiftF11WorksheetUndo?.expectedTabCount ||
    !keyboardNavigation.shiftF11WorksheetUndo?.status?.includes("元に戻しました") ||
    keyboardNavigation.ctrlF6WindowSwitch?.ready?.visibleWindowCount < 2 ||
    keyboardNavigation.ctrlF6WindowSwitch?.previousStep?.activeWindowId === keyboardNavigation.ctrlF6WindowSwitch?.ready?.activeWindowId ||
    keyboardNavigation.ctrlF6WindowSwitch?.afterNext !== keyboardNavigation.ctrlF6WindowSwitch?.ready?.activeWindowId ||
    !keyboardNavigation.ctrlF6WindowSwitch?.previousStep?.status?.includes("ウィンドウ") ||
    keyboardNavigation.ctrlF4CloseWindow?.visibleWindowCount !== keyboardNavigation.ctrlF4CloseWindow?.expectedCount ||
    !keyboardNavigation.ctrlF4CloseWindow?.status?.includes("ウィンドウ") ||
    keyboardNavigation.ctrlF4CloseWindowUndo?.visibleWindowCount !== keyboardNavigation.ctrlF4CloseWindowUndo?.expectedCount ||
    !keyboardNavigation.ctrlF4CloseWindowUndo?.status?.includes("元に戻しました") ||
    keyboardNavigation.ctrlWCloseWindow?.visibleWindowCount !== keyboardNavigation.ctrlWCloseWindow?.expectedCount ||
    !keyboardNavigation.ctrlWCloseWindow?.status?.includes("ウィンドウ") ||
    keyboardNavigation.ctrlWCloseWindowUndo?.visibleWindowCount !== keyboardNavigation.ctrlWCloseWindowUndo?.expectedCount ||
    !keyboardNavigation.ctrlWCloseWindowUndo?.status?.includes("元に戻しました") ||
    keyboardNavigation.ribbonCollapsed?.collapsed !== true ||
    keyboardNavigation.ribbonCollapsed?.panelsHidden !== true ||
    !keyboardNavigation.ribbonCollapsed?.status?.includes("リボン") ||
    keyboardNavigation.ribbonExpanded?.collapsed !== false ||
    keyboardNavigation.ribbonExpanded?.panelsHidden !== false ||
    keyboardNavigation.altF8MacrosTitle !== "マクロ" ||
    !String(keyboardNavigation.altF11VbaTitle || "").includes("Microsoft Visual Basic for Applications") ||
    !String(keyboardNavigation.f9Calculate || "").includes("ブック全体") ||
    !String(keyboardNavigation.shiftF9Calculate || "").includes("再計算") ||
    keyboardNavigation.ribbonAltTabs?.active !== "ホーム" ||
    !keyboardNavigation.ribbonAltTabs?.status?.includes("ホーム") ||
    keyboardNavigation.ribbonKeyboard?.focusedAfterF10 !== "home" ||
    !keyboardNavigation.ribbonKeyboard?.statusAfterF10?.includes("ホーム") ||
    keyboardNavigation.ribbonKeyboard?.focusedAfterRight !== "insert" ||
    keyboardNavigation.ribbonKeyboard?.activeAfterRight !== "home" ||
    keyboardNavigation.ribbonKeyboard?.activeAfterEnter !== "insert" ||
    !keyboardNavigation.ribbonKeyboard?.statusAfterEnter?.includes("挿入") ||
    keyboardNavigation.ribbonKeyboard?.focusedBeforePanel !== "home" ||
    keyboardNavigation.ribbonKeyboard?.activeAfterSpace !== "home" ||
    !keyboardNavigation.ribbonKeyboard?.statusAfterSpace?.includes("ホーム") ||
    keyboardNavigation.ribbonKeyboard?.focusedCommandAfterDown !== "paste" ||
    !keyboardNavigation.ribbonKeyboard?.focusedCommandAfterControlRight ||
    keyboardNavigation.ribbonKeyboard?.focusedCommandAfterControlRight === keyboardNavigation.ribbonKeyboard?.focusedCommandAfterDown ||
    keyboardNavigation.ribbonKeyboard?.focusedAfterEscape !== "home" ||
    !/^\d{4}\/\d{2}\/\d{2}$/.test(keyboardNavigation.shortcutDate || "") ||
    !/^\d{2}:\d{2}$/.test(keyboardNavigation.shortcutTime || "") ||
    keyboardNavigation.ctrlEnterFormulaFill?.values?.join("|") !== "CtrlEnterFill|CtrlEnterFill|CtrlEnterFill|CtrlEnterFill" ||
    !keyboardNavigation.ctrlEnterFormulaFill?.status?.includes("4 個") ||
    keyboardNavigation.ctrlEnterFormulaFillUndo?.values?.some((value) => value === "CtrlEnterFill") ||
    !keyboardNavigation.ctrlEnterFormulaFillUndo?.status?.includes("元に戻しました") ||
    keyboardNavigation.ctrlEnterActiveFill?.values?.join("|") !== "ActiveCtrlEnter|ActiveCtrlEnter|ActiveCtrlEnter|ActiveCtrlEnter" ||
    keyboardNavigation.nameBoxJump?.label !== "AA88" ||
    !keyboardNavigation.nameBoxJump?.status?.includes("AA88") ||
    keyboardNavigation.nameBoxDefine?.exists !== true ||
    !keyboardNavigation.nameBoxDefine?.status?.includes("NameBoxRange") ||
    keyboardNavigation.nameBoxDefinedNameJump?.label !== "AB88:AC89" ||
    keyboardNavigation.nameBoxMenuJump?.label !== "AB88:AC89" ||
    keyboardNavigation.nameBoxMenuJump?.menuClosed !== true ||
    keyboardNavigation.nameBoxMenuJump?.keyboard?.closedByEscape !== true ||
    keyboardNavigation.nameBoxMenuJump?.keyboard?.focusAfterEscape !== "selectedCellLabel" ||
    keyboardNavigation.nameBoxMenuJump?.keyboard?.itemCount < 1 ||
    keyboardNavigation.formulaF4ReferenceCycle?.first !== "=$A$1:$B$2" ||
    keyboardNavigation.formulaF4ReferenceCycle?.second !== "=A$1:B$2" ||
    keyboardNavigation.formulaF4ReferenceCycle?.third !== "=$A1:$B2" ||
    keyboardNavigation.formulaF4ReferenceCycle?.fourth !== "=A1:B2" ||
    keyboardNavigation.fillDown?.join("|") !== "FillDownShortcut|FillDownShortcut|FillDownShortcut" ||
    keyboardNavigation.fillRight?.join("|") !== "FillRightShortcut|FillRightShortcut|FillRightShortcut" ||
    keyboardNavigation.singleCellFill?.down !== "SingleFillDown" ||
    keyboardNavigation.singleCellFill?.right !== "SingleFillRight" ||
    keyboardNavigation.fillFormula !== "=P74" ||
    keyboardNavigation.autoSumFormula !== "=SUM(T74:T75)" ||
    keyboardNavigation.hiddenRowByShortcut !== true ||
    keyboardNavigation.unhiddenRowByShortcut !== false ||
    keyboardNavigation.hiddenColumnByShortcut !== true ||
    keyboardNavigation.unhiddenColumnByShortcut !== false ||
    keyboardNavigation.filterShortcutOn?.top !== 74 ||
    keyboardNavigation.filterShortcutOn?.left !== 21 ||
    keyboardNavigation.filterShortcutOn?.bottom !== 76 ||
    keyboardNavigation.filterShortcutOn?.right !== 22 ||
    keyboardNavigation.autoFilterAltDown?.open !== true ||
    keyboardNavigation.autoFilterAltDown?.col !== "21" ||
    keyboardNavigation.autoFilterAltDown?.focused !== "昇順" ||
    keyboardNavigation.filterShortcutOff !== null ||
    keyboardNavigation.showFormulasOn?.showFormulas !== true ||
    keyboardNavigation.showFormulasOn?.text !== "=U75" ||
    keyboardNavigation.showFormulasUndo?.showFormulas !== false ||
    !keyboardNavigation.showFormulasUndo?.status?.includes("元に戻しました") ||
    keyboardNavigation.showFormulasRedo?.showFormulas !== true ||
    !keyboardNavigation.showFormulasRedo?.status?.includes("やり直しました") ||
    keyboardNavigation.precedentShortcut !== "U75" ||
    keyboardNavigation.dependentShortcut !== "W74" ||
    keyboardNavigation.visibleCellsShortcut?.ranges?.length !== 4 ||
    keyboardNavigation.visibleCellsShortcut?.ranges?.some((range) => range.top === 75 || range.left === 22) ||
    keyboardNavigation.findDialogFirst?.label !== "X78" ||
    keyboardNavigation.shiftF4FindNext?.label !== "X77" ||
    keyboardNavigation.shiftF5FindDialogTitle !== "検索と置換" ||
    keyboardNavigation.linkShortcutTitle !== "ハイパーリンクの挿入" ||
    keyboardNavigation.insertFunctionShortcutTitle !== "関数の挿入" ||
    keyboardNavigation.nameManagerShortcutTitle !== "名前の管理" ||
    keyboardNavigation.createNamesShortcutTitle !== "選択範囲から名前を作成" ||
    keyboardNavigation.createdNameExists !== true ||
    keyboardNavigation.pasteNameShortcut?.title !== "名前の貼り付け" ||
    keyboardNavigation.pasteNameShortcut?.hasName !== true ||
    !keyboardNavigation.pasteNameShortcut?.reference?.includes("$U$79") ||
    keyboardNavigation.pasteNameFormula !== "=NameHeader" ||
    keyboardNavigation.printShortcut?.title !== "印刷" ||
    !["printCopies", "printPrinter", "printScope", "printOrientation", "printPaper", "printScale"].every((field) => keyboardNavigation.printShortcut?.fields?.includes(field)) ||
    keyboardNavigation.printPageSetup?.title !== "ページ設定" ||
    keyboardNavigation.printPageSetup?.landscape !== true ||
    keyboardNavigation.printPageSetup?.scale !== "125" ||
    keyboardNavigation.ctrlF2PrintPreview?.title !== "印刷" ||
    keyboardNavigation.ctrlF2PrintPreview?.orientation !== "landscape" ||
    keyboardNavigation.ctrlF2PrintPreview?.scale !== "125" ||
    keyboardNavigation.createTableShortcutDialog?.title !== "テーブルの作成" ||
    keyboardNavigation.createTableShortcutDialog?.range !== "U82:V84" ||
    keyboardNavigation.createTableShortcutDialog?.hasHeaders !== true ||
    keyboardNavigation.createTableShortcut?.range?.top !== 82 ||
    keyboardNavigation.createTableShortcut?.range?.left !== 21 ||
    keyboardNavigation.createTableShortcut?.range?.bottom !== 84 ||
    keyboardNavigation.createTableShortcut?.range?.right !== 22 ||
    keyboardNavigation.createTableShortcut?.showFilterButton !== true ||
    !keyboardNavigation.createTableShortcut?.status?.includes("テーブル") ||
    !keyboardNavigation.createTableShortcutUndo?.status?.includes("元に戻しました")
  ) {
    throw new Error("Enter, Tab, Home/End, PageUp/PageDown, Ctrl+arrow keys, Shift range extension, Ctrl+A, row/column shortcuts, sheet/window switching, hidden row/column skipping, F2/F3/F11, print, and common Excel editing shortcuts should behave like Excel.");
  }


  if (shortcutUndo.g6 === "貼り付けA" || shortcutUndo.h6 === "貼り付けB" || !shortcutUndo.status?.includes("元に戻しました")) {
    throw new Error("Ctrl+Z should undo keyboard paste.");
  }


  if (!shortcutClipboard.includes("貼り付けA") || !shortcutClipboard.includes("貼り付けB")) {
    throw new Error("Ctrl+C should copy the selected range to the clipboard.");
  }


  if (
    formatUndoRedo.undoFill === "#d9ead3" ||
    !formatUndoRedo.undoStatus?.includes("元に戻しました") ||
    formatUndoRedo.redoFill !== "#d9ead3" ||
    !formatUndoRedo.redoStatus?.includes("やり直しました") ||
    formatUndoRedo.painterActiveAfterStart?.active !== true ||
    !formatUndoRedo.painterActiveAfterStart?.status?.includes("書式をコピー") ||
    formatUndoRedo.painterActiveAfterApply !== false ||
    formatUndoRedo.painterText !== "target value" ||
    formatUndoRedo.painterFill !== "#d9ead3" ||
    formatUndoRedo.painterBold !== true ||
    formatUndoRedo.painterAlign !== "center" ||
    !formatUndoRedo.painterStatus?.includes("書式を貼り付け") ||
    formatUndoRedo.painterUndoText !== "target value" ||
    formatUndoRedo.painterUndoFill === "#d9ead3" ||
    !formatUndoRedo.painterUndoStatus?.includes("元に戻しました") ||
    formatUndoRedo.painterRedoFill !== "#d9ead3" ||
    !formatUndoRedo.painterRedoStatus?.includes("やり直しました")
  ) {
    throw new Error("Toolbar formatting actions and format painter should be undoable, redoable, and preserve target values.");
  }


  if (
    !wrapToolbar.activeAfterWrap ||
    wrapToolbar.whiteSpaceAfterWrap !== "pre-wrap" ||
    wrapToolbar.overflowWrapAfterWrap !== "anywhere" ||
    wrapToolbar.activeAfterUnwrap ||
    wrapToolbar.whiteSpaceAfterUnwrap === "pre-wrap"
  ) {
    throw new Error("Wrap Text should toggle cell wrapping and mirror the selected cell state in the ribbon.");
  }


  if (toolbarSync.fontSize !== "16" || !toolbarSync.boldActive || !toolbarSync.alignCenterActive || toolbarSync.selectionLabel !== "G6:H6") {
    throw new Error("Ribbon controls should reflect the first cell in a multi-cell selection.");
  }


  if (toolbarSync.fillColor !== "#123456") {
    throw new Error("Ribbon fill color should not be overwritten when syncing selected cell attributes.");
  }


  if (
    !selectionFrame.exists ||
    selectionFrame.selectionLabel !== "G6:H6" ||
    Math.abs(selectionFrame.frame.left - selectionFrame.expectedLeft) > 2 ||
    Math.abs(selectionFrame.frame.top - selectionFrame.expectedTop) > 2 ||
    Math.abs(selectionFrame.frame.right - selectionFrame.expectedRight) > 2 ||
    Math.abs(selectionFrame.frame.bottom - selectionFrame.expectedBottom) > 2 ||
    selectionFrame.g6Outline !== "none" ||
    selectionFrame.h6Outline !== "none" ||
    selectionFrame.g6UserSelect !== "none"
  ) {
    throw new Error("Multi-cell selection should render as one green range frame without text-selection-like outlines.");
  }


  if (
    copyStateAfterCopy.selectionLabel !== "G6:H6" ||
    !copyStateAfterCopy.selectionFrameClass?.includes("is-copy-source") ||
    copyStateAfterCopy.copyFrameCount !== 0 ||
    copyStateAfterCopy.animationName !== "marching-ants" ||
    !copyStateAfterCopy.backgroundImage?.includes("34, 197, 94")
  ) {
    throw new Error("Copied cell selection should switch to a green Excel-like animated dashed copy frame.");
  }


  if (
    copyStateWithTarget.selectionLabel !== "I6" ||
    copyStateWithTarget.copyFrameClass !== "sheet-copy-frame" ||
    Math.abs(copyStateWithTarget.copyFrame.left - copyStateWithTarget.sourceLeft) > 2 ||
    Math.abs(copyStateWithTarget.copyFrame.right - copyStateWithTarget.sourceRight) > 2 ||
    Math.abs(copyStateWithTarget.targetSelection.left - copyStateWithTarget.targetCell.left) > 2
  ) {
    throw new Error("Copied source range should keep a marching border while another paste target is selected.");
  }


  if (
    richPaste.pasted.text !== "貼り付けA" ||
    richPaste.neighborPasted.text !== "貼り付けB" ||
    Number.parseFloat(richPaste.pasted.fontSize) < 20 ||
    Number.parseInt(richPaste.pasted.fontWeight, 10) < 600 ||
    richPaste.pasted.textAlign !== "center" ||
    richPaste.pasted.fillColor !== "#d9ead3"
  ) {
    throw new Error("Ctrl+V should paste the copied cell attributes, not only text.");
  }


  if (
    matchDestinationPaste.pasted.text !== "貼り付けA" ||
    matchDestinationPaste.neighborPasted.text !== "貼り付けB" ||
    matchDestinationPaste.pasted.fillColor === "#d9ead3" ||
    Number.parseFloat(matchDestinationPaste.pasted.fontSize) >= 20
  ) {
    throw new Error("Paste matching destination formatting should use WebSheet's internal clipboard content without importing source cell formatting or stale system clipboard text.");
  }


  if (
    pasteSpecialFormats.title !== "形式を選択して貼り付け" ||
    ![
      "すべて(A)",
      "コピー元のテーマを使用してすべて(H)",
      "数式(F)",
      "値(V)",
      "書式(T)",
      "コメントとメモ(C)",
      "入力規則(N)",
      "罫線を除くすべて(X)",
      "列幅(W)",
      "数式と数値の書式(R)",
      "値と数値の書式(U)",
    ].every((label) => pasteSpecialFormats.options?.includes(label)) ||
    pasteSpecialFormats.text === "貼り付けA" ||
    pasteSpecialFormats.text === "貼り付けB" ||
    pasteSpecialFormats.fillColor !== "#d9ead3" ||
    Number.parseFloat(pasteSpecialFormats.fontSize) < 20 ||
    !pasteSpecialFormats.status?.includes("書式")
  ) {
    throw new Error("Ctrl+Alt+V should open Paste Special and paste formats without overwriting the target value.");
  }


  if (
    String(pasteSpecialAdvanced.transposeAddSkipBlanks?.q) !== "12" ||
    String(pasteSpecialAdvanced.transposeAddSkipBlanks?.r) !== "99" ||
    String(pasteSpecialAdvanced.transposeAddSkipBlanks?.qNext) !== "33" ||
    String(pasteSpecialAdvanced.transposeAddSkipBlanks?.rNext) !== "44" ||
    !String(pasteSpecialAdvanced.link?.s || "").includes(`!N${pasteSpecialAdvanced.baseRow}`) ||
    !String(pasteSpecialAdvanced.link?.t || "").includes(`!O${pasteSpecialAdvanced.baseRow}`) ||
    !String(pasteSpecialAdvanced.link?.sNext || "").includes(`!N${pasteSpecialAdvanced.baseRow + 1}`) ||
    !String(pasteSpecialAdvanced.link?.tNext || "").includes(`!O${pasteSpecialAdvanced.baseRow + 1}`) ||
    !pasteSpecialAdvanced.status?.includes("リンク貼り付け")
  ) {
    throw new Error("Paste Special should support operation, skip blanks, transpose, and paste link.");
  }


  if (
    formulaPasteReferences.relative !== "=AC186" ||
    formulaPasteReferences.absolute !== "=$AA$184" ||
    !formulaPasteReferences.status?.includes("貼り付け")
  ) {
    throw new Error("Cell paste should shift relative formula references and keep absolute references fixed.");
  }


  if (
    cutAndRepeatPaste.cutFormula !== "=AA188" ||
    String(cutAndRepeatPaste.cutSourceRaw) !== "" ||
    cutAndRepeatPaste.repeat?.ah !== "A" ||
    cutAndRepeatPaste.repeat?.ai !== "B" ||
    cutAndRepeatPaste.repeat?.aj !== "A" ||
    cutAndRepeatPaste.repeat?.ahNext !== "C" ||
    cutAndRepeatPaste.repeat?.ahThird !== "A" ||
    cutAndRepeatPaste.repeat?.akLast !== "D" ||
    !cutAndRepeatPaste.status?.includes("貼り付け")
  ) {
    throw new Error("Cut paste should keep formula references unchanged, and larger paste selections should repeat the copied cell block like Excel.");
  }


  const expectedCellMenuLabels = [
    "切り取り(T)",
    "コピー(C)",
    "翻訳",
    "挿入(I)...",
    "削除(D)...",
    "数式と値のクリア(N)",
    "クイック分析(Q)",
    "フィルター(E)›",
    "並び替え(O)›",
    "テーブルまたは範囲からデータを取得(G)...",
    "新しいコメント(M)",
    "新しいメモ(N)",
    "セルの書式設定(F)...",
    "ドロップダウンリストから選択(K)...",
    "ふりがなの表示(S)",
    "名前の定義(A)...",
    "リンク(I)",
    "ハイパーリンクを開く(O)",
  ].map((label) => label.replace(/\s+/g, ""));


  if (
    expectedCellMenuLabels.some((label) => !cellMenu.labels.includes(label)) ||
    cellMenu.separatorCount !== 7 ||
    !cellMenu.pasteOptions.includes("元の書式を保持(K)") ||
    !cellMenu.pasteOptions.includes("貼り付け先の書式に合わせる(M)") ||
    !cellMenu.disabledActions.includes("furigana") ||
    !["T", "C", "I", "D", "N", "Q", "E", "O", "G", "M", "F", "K", "S", "A"].every((key) => cellMenu.accessKeys.includes(key)) ||
    !cellMenu.accessKeysUnderlined ||
    cellMenu.focusedBefore !== "cut" ||
    cellMenu.focusedAfterDown !== "copy" ||
    cellMenu.formatOpenedByAccessKey !== "セルの書式設定" ||
    cellMenu.rect.left < 0 ||
    cellMenu.rect.top < 0 ||
    cellMenu.rect.right > 1280 ||
    cellMenu.rect.bottom > 720
  ) {
    throw new Error("Cell right-click context menu did not render the expected Excel-like items.");
  }


  if (
    !translationPaneInitial.visible ||
    translationPaneInitial.title !== "翻訳" ||
    !translationPaneInitial.range.includes("G6:H6") ||
    !translationPaneHeader.range.includes("B4:G4") ||
    translationPaneHeader.sourceText !== "氏名" ||
    translationPaneHeader.resultText !== "姓名" ||
    !translationPaneHeader.detected.includes("日本語") ||
    translationReplace.b4 !== "姓名" ||
    !translationReplace.status.includes("翻訳") ||
    !translationReplace.selectionLabel.includes("B4:G4") ||
    translationUndo.b4 !== "氏名" ||
    !translationUndo.status.includes("元に戻しました")
  ) {
    throw new Error("Cell Translate command should open the translation pane, replace the selected cell, and support Undo.");
  }


  if (
    !cellAnnotations.commentEditor ||
    cellAnnotations.commentAria !== "新しいコメント" ||
    cellAnnotations.commentAddress !== "G6" ||
    cellAnnotations.commentPost !== "投稿" ||
    cellAnnotations.commentCancel !== "キャンセル" ||
    !cellAnnotations.comment.hasCommentMarker ||
    cellAnnotations.comment.hasNoteMarker ||
    !cellAnnotations.comment.title?.includes("コメント本文") ||
    !cellAnnotations.comment.status?.includes("コメント") ||
    cellAnnotations.commentPopupAfterCreate.visible ||
    !cellAnnotations.commentPopupInitial.visible ||
    cellAnnotations.commentPopupInitial.title !== "コメント" ||
    cellAnnotations.commentPopupInitial.address !== "G6" ||
    !cellAnnotations.commentPopupInitial.posts.includes("コメント本文") ||
    !cellAnnotations.commentPopupInitial.hasReplyBox ||
    cellAnnotations.commentPopupInitial.postButton !== "投稿" ||
    cellAnnotations.commentPopupInitial.pointerDelta > 2 ||
    !cellAnnotations.commentContextMenu.labels.includes("コメントに返信する(R)") ||
    !cellAnnotations.commentContextMenu.labels.includes("コメントの削除(M)") ||
    cellAnnotations.commentContextMenu.labels.includes("新しいコメント(M)") ||
    !cellAnnotations.commentReplyFromMenu.visible ||
    cellAnnotations.commentReplyFromMenu.activeElement !== "TEXTAREA" ||
    !cellAnnotations.commentReplyFromMenu.activeInPopup ||
    cellAnnotations.commentPopupReply.visible ||
    !cellAnnotations.commentPopupReply.title?.includes("返信本文") ||
    !cellAnnotations.commentPopupReply.status?.includes("コメントを投稿") ||
    !cellAnnotations.commentPopupHover.visible ||
    cellAnnotations.commentPopupHover.postCount !== 2 ||
    !cellAnnotations.commentPopupHover.posts.includes("コメント本文") ||
    !cellAnnotations.commentPopupHover.posts.includes("返信本文") ||
    cellAnnotations.commentPopupHover.pointerDelta > 2 ||
    !cellAnnotations.noteInitial.noteEditor ||
    cellAnnotations.noteInitial.noteAria !== "新しいメモ" ||
    !cellAnnotations.noteInitial.noteAuthor?.includes("ユーザー") ||
    !cellAnnotations.noteInitial.noteBackground?.includes("255") ||
    !cellAnnotations.note.hasCommentMarker ||
    !cellAnnotations.note.hasNoteMarker ||
    !cellAnnotations.note.title?.includes("コメント本文") ||
    !cellAnnotations.note.title?.includes("メモ本文") ||
    cellAnnotations.note.popupVisible ||
    !cellAnnotations.note.status?.includes("メモ") ||
    !cellAnnotations.noteHover.visible ||
    !cellAnnotations.noteHover.popupText?.includes("メモ本文") ||
    cellAnnotations.noteHover.pointerDelta > 2 ||
    !cellAnnotations.undo.hasCommentMarker ||
    cellAnnotations.undo.hasNoteMarker ||
    !cellAnnotations.undo.status?.includes("元に戻しました") ||
    !cellAnnotations.redo.hasCommentMarker ||
    !cellAnnotations.redo.hasNoteMarker ||
    cellAnnotations.redo.visible ||
    !cellAnnotations.redo.status?.includes("やり直しました") ||
    !cellAnnotations.memoContextMenu.labels.includes("メモの編集(E)") ||
    !cellAnnotations.memoContextMenu.labels.includes("メモの削除(T)") ||
    !cellAnnotations.memoContextMenu.labels.includes("メモの表示/非表示(O)") ||
    cellAnnotations.memoContextMenu.labels.includes("新しいメモ(N)") ||
    !cellAnnotations.memoEdit.title?.includes("メモ本文編集") ||
    cellAnnotations.memoEdit.visible ||
    !cellAnnotations.memoToggleHidden.status?.includes("非表示") ||
    cellAnnotations.memoToggleHidden.visible ||
    !cellAnnotations.memoToggleShown.visible ||
    !cellAnnotations.memoToggleShown.popupText?.includes("メモ本文編集") ||
    !cellAnnotations.memoToggleShown.status?.includes("表示") ||
    cellAnnotations.memoDelete.hasNoteMarker ||
    cellAnnotations.memoDelete.visible ||
    !cellAnnotations.memoDelete.status?.includes("削除") ||
    !cellAnnotations.memoDeleteUndo.hasNoteMarker ||
    cellAnnotations.memoDeleteUndo.visible ||
    !cellAnnotations.memoDeleteUndo.title?.includes("メモ本文編集") ||
    !cellAnnotations.memoDeleteUndo.status?.includes("元に戻しました") ||
    cellAnnotations.memoDeleteRedo.hasNoteMarker ||
    cellAnnotations.memoDeleteRedo.visible ||
    !cellAnnotations.memoDeleteRedo.status?.includes("やり直しました") ||
    !cellAnnotations.commentDeleteContextMenu.labels.includes("コメントに返信する(R)") ||
    !cellAnnotations.commentDeleteContextMenu.labels.includes("コメントの削除(M)") ||
    cellAnnotations.commentDeleteContextMenu.labels.includes("新しいコメント(M)") ||
    cellAnnotations.commentDelete.hasCommentMarker ||
    !cellAnnotations.commentDelete.status?.includes("削除") ||
    !cellAnnotations.commentDeleteUndo.hasCommentMarker ||
    !cellAnnotations.commentDeleteUndo.title?.includes("返信本文") ||
    !cellAnnotations.commentDeleteUndo.status?.includes("元に戻しました") ||
    cellAnnotations.commentDeleteRedo.hasCommentMarker ||
    !cellAnnotations.commentDeleteRedo.status?.includes("やり直しました")
  ) {
    throw new Error("New Comment and New Note should open Excel-like editors, persist indicators, and participate in undo/redo.");
  }


  if (
    !furigana.initial.exists ||
    furigana.initial.text !== "東十条" ||
    !furigana.initial.hasPhonetic ||
    furigana.initial.showPhonetic ||
    !furigana.contextMenu.labels.includes("ふりがなの表示(S)") ||
    furigana.contextMenu.disabledActions.includes("furigana") ||
    !furigana.contextMenu.accessKeys.includes("S") ||
    !furigana.contextMenu.accessKeysUnderlined ||
    !furigana.show.showPhonetic ||
    !furigana.show.status?.includes("ふりがなを表示") ||
    !furigana.show.paints.includes("ヒガシジュウジョウ") ||
    furigana.undo.showPhonetic ||
    !furigana.undo.status?.includes("元に戻しました") ||
    !furigana.redo.showPhonetic ||
    !furigana.redo.status?.includes("やり直しました") ||
    !furigana.redo.paints.includes("ヒガシジュウジョウ")
  ) {
    throw new Error("ふりがなの表示 should read Excel phonetic runs, render them on canvas, and support undo/redo.");
  }


  if (
    formatCellsDialog.title !== "セルの書式設定" ||
    formatCellsDialog.tabs.join("|") !== "表示形式|配置|フォント|罫線|塗りつぶし|保護" ||
    formatCellsDialog.activePanel !== "number" ||
    formatCellsDialog.ok !== "OK" ||
    formatCellsDialog.cancel !== "キャンセル" ||
    formatCellsDialog.apply !== "適用" ||
    !["C", "D", "S", "T", "H", "V", "F", "P"].every((key) => formatCellsDialog.accessKeys.includes(key)) ||
    !formatCellsDialog.accessKeysUnderlined ||
    !formatCellsDialog.closed ||
    formatCellsDialog.fillColor !== "#ddebf7" ||
    formatCellsDialog.color !== "rgb(112, 48, 160)" ||
    Number.parseFloat(formatCellsDialog.fontSize) < 23 ||
    formatCellsDialog.textAlign !== "right" ||
    formatCellsDialog.alignItems !== "flex-end" ||
    formatCellsDialog.whiteSpace !== "pre-wrap" ||
    formatCellsDialog.disjointFillColor !== "#ddebf7" ||
    Number.parseFloat(formatCellsDialog.disjointFontSize) < 23 ||
    !formatCellsDialog.status?.includes("セルの書式設定")
  ) {
    throw new Error("Cell format dialog should expose Excel-like tabs and apply number, alignment, font, border, fill, and protection settings.");
  }


  if (cellMenuClear.g6 !== "" || cellMenuClear.h6 !== "" || !cellMenuClear.status?.includes("クリア")) {
    throw new Error("Cell context menu clear did not clear formulas and values.");
  }


  if (cellMenuUndo.g6 !== "貼り付けA" || cellMenuUndo.h6 !== "貼り付けB" || !cellMenuUndo.status?.includes("元に戻しました")) {
    throw new Error("Ctrl+Z should undo cell context menu clear.");
  }


  if (
    deleteKeyClear.g6 !== "" ||
    deleteKeyClear.h6 !== "" ||
    deleteKeyClear.g6Fill !== "#ddebf7" ||
    deleteKeyClear.h6Fill !== "#ddebf7" ||
    !deleteKeyClear.status?.includes("クリア")
  ) {
    throw new Error("Delete key should clear selected cell contents while preserving cell formatting.");
  }


  if (deleteKeyUndo.g6 !== "貼り付けA" || deleteKeyUndo.h6 !== "貼り付けB" || !deleteKeyUndo.status?.includes("元に戻しました")) {
    throw new Error("Ctrl+Z should undo Delete-key cell clearing.");
  }


  if (
    backspaceClearEdit.activeAddress !== "G6" ||
    backspaceClearEdit.g6Editable !== "true" ||
    backspaceClearEdit.g6Text !== "" ||
    backspaceClearEdit.h6Text !== "" ||
    backspaceClearEdit.label !== "G6" ||
    !backspaceClearEdit.status?.includes("クリア") ||
    backspaceUndo.g6 !== "貼り付けA" ||
    backspaceUndo.h6 !== "貼り付けB" ||
    !backspaceUndo.status?.includes("元に戻しました")
  ) {
    throw new Error("Backspace should clear selected cells, enter edit mode on the first selected cell, and remain undoable.");
  }


  if (!popupDismissal.leftClickClosedCellMenu || !popupDismissal.rightClickOpenedSheetMenu || !popupDismissal.rightClickClosedCellMenu) {
    throw new Error("Floating menus should close on outside left click, and outside right click should allow the target menu to open.");
  }


  if (!result.g11.text?.includes("\n") || result.g11.whiteSpace !== "pre-wrap") {
    throw new Error("Multiline/self PR cell line breaks are not preserved.");
  }


  if (result.gridEdge.hostTop !== result.gridEdge.gridTop || result.gridEdge.hostLeft !== result.gridEdge.gridLeft) {
    throw new Error("Sheet grid still has top/left padding outside row and column headers.");
  }


  if (!result.gridEdge.hidesGridlines || result.a1Gridline.hasFillClass || result.a1Gridline.backgroundColor !== "rgba(0, 0, 0, 0)" || result.a1Gridline.borderRightColor !== "rgba(0, 0, 0, 0)" || result.a1Gridline.borderBottomColor !== "rgba(0, 0, 0, 0)") {
    throw new Error("Imported sheets with showGridLines=false should hide default gridlines.");
  }


  if (!result.whiteFill.hasFillClass || result.whiteFill.fillColor !== "#ffffff" || result.whiteFill.backgroundColor !== "rgba(0, 0, 0, 0)" || result.whiteFill.borderRightColor !== "rgba(0, 0, 0, 0)" || result.whiteFill.borderBottomColor !== "rgba(0, 0, 0, 0)") {
    throw new Error("Explicit white fills should render as real fills, not transparent default grid cells.");
  }


  if (!ar2Gridline.gridClassName?.includes("hide-gridlines") || ar2Gridline.ar2.backgroundColor !== "rgba(0, 0, 0, 0)" || ar2Gridline.ar2.borderRightColor !== "rgba(0, 0, 0, 0)" || ar2Gridline.ar2.borderBottomColor !== "rgba(0, 0, 0, 0)" || ar2Gridline.aq2.borderRightColor !== "rgba(0, 0, 0, 0)" || ar2Gridline.ar1.borderBottomColor !== "rgba(0, 0, 0, 0)") {
    throw new Error("AR2 should inherit the imported sheet's hidden-gridline white canvas.");
  }


  if (!imageAndFont.imageCount || imageAndFont.imageCount !== 2 || imageAndFont.images.some((image) => !image.src?.startsWith("data:image/png;base64,") || image.width < 500 || image.height < 270 || image.naturalWidth <= 0)) {
    throw new Error("Worksheet images did not render from embedded Excel media.");
  }


  if (
    importedNativeShapes.shapeCount < 6 ||
    importedNativeShapes.active !== "Sheet1" ||
    !importedNativeShapes.names.some((name) => name.includes("正方形")) ||
    !importedNativeShapes.names.some((name) => name.includes("二等辺三角形")) ||
    !importedNativeShapes.names.some((name) => name.includes("直線矢印コネクタ")) ||
    !importedNativeShapes.hasSvgShapes ||
	    !importedNativeShapes.hasMeaningfulBounds ||
	    !importedNativeShapes.hasRichTextRuns ||
	    !importedNativeShapes.hasImportedShapeEffect ||
	    !importedNativeShapes.hasGroupedPicture ||
	    !importedNativeShapes.hasGroupedPictureTransformAndCrop ||
	    !importedNativeShapes.hasGroupedShape ||
	    !importedNativeShapes.hasSinglePicture ||
	    !importedNativeShapes.hasImportedPictureBorder ||
	    !importedNativeShapes.hasImportedPictureEffect ||
	    !importedNativeShapes.hasGraphicFrameChartRender ||
	    !importedNativeShapes.hasGraphicFrameChartSeriesColors ||
	    !importedNativeShapes.hasGraphicFrameChartAreaFormat ||
	    !importedNativeShapes.hasGraphicFrameChartPlotFormat ||
	    !importedNativeShapes.hasGraphicFrameChartRightLegend ||
	    !importedNativeShapes.hasGraphicFrameChartAxisFormat ||
	    !importedNativeShapes.hasGraphicFrameChartGridlines ||
	    !importedNativeShapes.hasGraphicFrameChartDataLabels ||
	    !importedNativeShapes.hasGraphicFrameChartHorizontalBar ||
	    !importedNativeShapes.hasGraphicFrameChartStacked ||
	    !importedNativeShapes.hasGraphicFrameChartAxisTitles ||
	    !importedNativeShapes.hasGraphicFrameChartTextStyles ||
	    !importedNativeShapes.hasGraphicFrameLineChartMarker ||
	    !importedNativeShapes.hasGraphicFramePieChartLabels ||
	    importedNativeShapes.chartTitleRef !== "Sheet1!$D$1" ||
	    importedNativeShapes.chartLabelsRef !== "Sheet1!$A$1:$A$3" ||
	    importedNativeShapes.chartFirstNameRef !== "Sheet1!$E$1" ||
	    importedNativeShapes.chartFirstValuesRef !== "Sheet1!$B$1:$B$3" ||
	    importedNativeShapes.chartSecondNameRef !== "Sheet1!$F$1" ||
	    importedNativeShapes.chartSecondValuesRef !== "Sheet1!$C$1:$C$3" ||
	    !importedNativeShapes.chartDataRefresh?.refreshed ||
	    importedNativeShapes.chartDataRefresh?.title !== "更新グラフ" ||
	    importedNativeShapes.chartDataRefresh?.firstName !== "更新売上" ||
	    importedNativeShapes.chartDataRefresh?.secondName !== "更新利益" ||
	    importedNativeShapes.chartDataRefresh?.labels?.join("|") !== "R1|R2|R3" ||
	    importedNativeShapes.chartDataRefresh?.firstValues?.join("|") !== "111|222|333" ||
	    importedNativeShapes.chartDataRefresh?.secondValues?.join("|") !== "11|22|33" ||
	    !importedNativeShapes.chartDataRefresh?.svgHasUpdatedTitle ||
	    !importedNativeShapes.chartDataRefresh?.svgHasUpdatedSeriesName ||
	    !importedNativeShapes.chartDataRefresh?.svgHasUpdatedLabel ||
	    !importedNativeShapes.chartDataRefresh?.svgHasUpdatedValue ||
	    !importedNativeShapes.chartRowStructureApplied ||
	    !importedNativeShapes.chartColumnStructureApplied ||
	    importedNativeShapes.chartRefsBeforeStructure?.title !== "Sheet1!$D$1" ||
	    importedNativeShapes.chartRefsBeforeStructure?.labels !== "Sheet1!$A$1:$A$3" ||
	    importedNativeShapes.chartRefsBeforeStructure?.firstName !== "Sheet1!$E$1" ||
	    importedNativeShapes.chartRefsBeforeStructure?.first !== "Sheet1!$B$1:$B$3" ||
	    importedNativeShapes.chartRefsBeforeStructure?.secondName !== "Sheet1!$F$1" ||
	    importedNativeShapes.chartRefsBeforeStructure?.second !== "Sheet1!$C$1:$C$3" ||
	    importedNativeShapes.chartRefsAfterRowInsert?.title !== "Sheet1!$D$2" ||
	    importedNativeShapes.chartRefsAfterRowInsert?.labels !== "Sheet1!$A$2:$A$4" ||
	    importedNativeShapes.chartRefsAfterRowInsert?.firstName !== "Sheet1!$E$2" ||
	    importedNativeShapes.chartRefsAfterRowInsert?.first !== "Sheet1!$B$2:$B$4" ||
	    importedNativeShapes.chartRefsAfterRowInsert?.secondName !== "Sheet1!$F$2" ||
	    importedNativeShapes.chartRefsAfterRowInsert?.second !== "Sheet1!$C$2:$C$4" ||
	    importedNativeShapes.chartRefsAfterRowAndColumnInsert?.title !== "Sheet1!$E$2" ||
	    importedNativeShapes.chartRefsAfterRowAndColumnInsert?.labels !== "Sheet1!$A$2:$A$4" ||
	    importedNativeShapes.chartRefsAfterRowAndColumnInsert?.firstName !== "Sheet1!$F$2" ||
	    importedNativeShapes.chartRefsAfterRowAndColumnInsert?.first !== "Sheet1!$C$2:$C$4" ||
	    importedNativeShapes.chartRefsAfterRowAndColumnInsert?.secondName !== "Sheet1!$G$2" ||
	    importedNativeShapes.chartRefsAfterRowAndColumnInsert?.second !== "Sheet1!$D$2:$D$4" ||
	    importedNativeShapes.chartFormatTabText !== "グラフの書式" ||
	    importedNativeShapes.chartActiveRibbonTab !== "グラフの書式" ||
	    !importedNativeShapes.chartFormatPanelMode ||
	    !importedNativeShapes.chartInsertGroupHidden ||
	    !importedNativeShapes.chartTextGroupHidden ||
	    !importedNativeShapes.chartFillButtonDisabled ||
	    importedNativeShapes.chartStyleGroupLabel !== "グラフのスタイル" ||
	    importedNativeShapes.chartOutlineButtonTitle !== "グラフの枠線" ||
	    !importedNativeShapes.chartContextMenuFormatLabel.includes("グラフ エリアの書式設定") ||
	    !importedNativeShapes.chartEditTextDisabled ||
	    importedNativeShapes.chartPaneTitle !== "グラフ エリアの書式設定" ||
	    !importedNativeShapes.chartShapeOnlyFieldsDisabled ||
	    !importedNativeShapes.chartVisualFieldsEnabled ||
	    importedNativeShapes.chartBorderColor !== "rgb(255, 0, 0)" ||
	    importedNativeShapes.chartBorderWidth !== "2px" ||
	    importedNativeShapes.chartBorderStyle !== "dashed" ||
	    !importedNativeShapes.chartFilter.includes("drop-shadow") ||
	    !importedNativeShapes.groupedObjectFormatTabVisible ||
	    importedNativeShapes.groupedObjectActiveRibbonTab !== "図形の書式" ||
	    importedNativeShapes.groupedObjectSizeInputs.some((value) => !Number(value)) ||
	    !importedNativeShapes.groupedObjectRotationInput ||
	    importedNativeShapes.groupedObjectSelectionCount < 2 ||
	    !importedNativeShapes.groupedObjectUngroupEnabled ||
	    importedNativeShapes.singlePictureFormatTabText !== "図の形式" ||
	    importedNativeShapes.singlePictureActiveRibbonTab !== "図の形式" ||
	    !importedNativeShapes.singlePictureFormatPanelPictureMode ||
	    !importedNativeShapes.singlePictureInsertGroupHidden ||
	    !importedNativeShapes.singlePictureTextGroupHidden ||
	    !importedNativeShapes.singlePictureFillButtonDisabled ||
	    importedNativeShapes.singlePictureStyleGroupLabel !== "図のスタイル" ||
	    importedNativeShapes.singlePictureOutlineButtonTitle !== "図の枠線" ||
	    importedNativeShapes.singlePictureEffectsButtonTitle !== "図の効果" ||
	    !importedNativeShapes.singlePictureContextMenuFormatLabel.includes("図の書式設定") ||
	    importedNativeShapes.singlePicturePaneTitle !== "図の書式設定" ||
	    importedNativeShapes.singlePicturePaneSizeValues.some((value) => !Number(value)) ||
	    !importedNativeShapes.singlePictureShapeOnlyFieldsDisabled ||
	    !importedNativeShapes.singlePictureVisualFieldsEnabled ||
	    !importedNativeShapes.singlePictureSizeFieldsEnabled ||
	    importedNativeShapes.singlePictureBorderColor !== "rgb(255, 0, 0)" ||
	    importedNativeShapes.singlePictureBorderWidth !== "2px" ||
	    !importedNativeShapes.singlePictureFilter.includes("drop-shadow") ||
	    !importedNativeShapes.connectorFollowedMove
  ) {
    throw new Error("Native Excel DrawingML shapes should import as selectable SVG sheet shapes.");
  }


  if (imageAndFont.renderedAtTop >= imageAndFont.imageCount || imageAndFont.renderedAfterScroll >= imageAndFont.imageCount) {
    throw new Error("Offscreen worksheet images should be unloaded from the active DOM and restored when scrolled into view.");
  }


  if (imageAndFont.images.some((image) => Math.abs(image.left - (imageAndFont.rowHeaderWidth + 5)) > 1)) {
    throw new Error("Worksheet images should use native Excel drawing offsets instead of fractional column offsets.");
  }


  if (!imageAndFont.frame || imageAndFont.images.some((image) => image.left + image.width > imageAndFont.frame.right + 1)) {
    throw new Error("Worksheet column widths should keep anchored images inside the original Excel frame.");
  }


  if (imageAndFont.a2WhiteSpace !== "pre-wrap" || imageAndFont.a2OverflowWrap !== "anywhere") {
    throw new Error("Merged cells on the AI sheet should preserve imported Excel wrap-text formatting.");
  }


  if (imageSelection.selectedImages !== 1 || !imageSelection.selectedLabel?.includes("image") || imageSelection.formulaDisabled !== true) {
    throw new Error("Worksheet images should be selectable independently from cells.");
  }


  if (!imageAndFont.a1FontFamily.includes("Noto Sans CJK JP") || !imageAndFont.a1FontFamily.includes("Hiragino Sans") || !imageAndFont.b6FontFamily.includes("Noto Sans CJK JP")) {
    throw new Error("AI sheet font family stack does not preserve Japanese Excel font choices.");
  }


  if (!exportedHtmlChecks.hasRemoteDownloadRuntime || !exportedHtmlChecks.hasNoInlineRuntimeDependencies || !exportedHtmlChecks.hasFullAppBootstrap || !exportedHtmlChecks.hasCanvasLayerRule || !exportedHtmlChecks.hasTransparentDomRule || !exportedHtmlChecks.hasCanvasRuntime || !exportedHtmlChecks.hasCanvasBorderRuntime || !exportedHtmlChecks.hasOverlayCssRuntime || !exportedHtmlChecks.hasFillRuntime || !exportedHtmlChecks.hasAdjacentFillRuntime || !exportedHtmlChecks.hasFillModel || !exportedHtmlChecks.hasHiddenGridlineModel || !exportedHtmlChecks.hasHiddenGridlineRule || !exportedHtmlChecks.hasImages || !exportedHtmlChecks.hasImageCss || !exportedHtmlChecks.hasFontFallback || !exportedHtmlChecks.hasCtrlWheelZoom || !exportedHtmlChecks.hasZoomDialog || !exportedHtmlChecks.hasPhoneticModel || !exportedHtmlChecks.hasPhoneticRuntime || !exportedHtmlChecks.hasCanvasTextBaselineRuntime || !exportedHtmlChecks.hasCanvasTextLineHeightRuntime || !exportedHtmlChecks.hasAccessKeyUnderline || !exportedHtmlChecks.hasDefinedNames || !exportedHtmlChecks.hasImeCompositionRuntime || !exportedHtmlChecks.hasLiveChartReferences || !exportedHtmlChecks.hasAdvancedShapeRuntime || !exportedHtmlChecks.hasLocalSaveEndpoint || !exportedHtmlChecks.hasNoExternalDependencyScripts) {
    throw new Error("Standalone HTML export does not preserve the full WebSheet runtime, fills, images, fonts, hidden gridline settings, Ctrl+wheel zoom, the zoom popup, furigana data, defined names, IME composition handling, live chart references, advanced shape formatting, local save endpoint, menu access-key underlines, or inline runtime dependencies.");
  }


  if (
    standaloneOffline.firstTab !== "CHT" ||
    standaloneOffline.cellCount <= 0 ||
    !standaloneOffline.hasRibbon ||
    !standaloneOffline.hasHomePanel ||
    !standaloneOffline.cellMenuVisible ||
    !standaloneOffline.standaloneModelLoaded ||
    (standaloneOffline.selectionFrameCount < 2 && !standaloneOffline.hasMultiRangeFrame) ||
    standaloneOffline.runtimeSource !== "local" ||
    !standaloneOffline.loadingStatusCleared ||
		    standaloneOffline.quickSave?.pickerCalls?.length !== 0 ||
		    standaloneOffline.quickSave?.localSaveRequests?.length !== 1 ||
		    !standaloneOffline.quickSave?.localSaveRequests?.[0]?.path?.endsWith(".html") ||
		    !standaloneOffline.quickSave?.localSaveRequests?.[0]?.hasWorkbookModel ||
		    !standaloneOffline.quickSave?.status?.includes("上書き保存") ||
	    !standaloneOffline.quickSave?.workbookName?.endsWith(".html") ||
	    standaloneOffline.clipboard?.calls?.writeText !== 0 ||
	    standaloneOffline.clipboard?.calls?.readText !== 0 ||
	    standaloneOffline.clipboard?.toolbar?.calls?.writeText !== 0 ||
	    standaloneOffline.clipboard?.toolbar?.calls?.readText !== 0 ||
	    standaloneOffline.clipboard?.copyEvent?.defaultPrevented !== true ||
	    standaloneOffline.clipboard?.copyEvent?.setDataCalls !== 0 ||
	    !standaloneOffline.clipboard?.toolbar?.status?.includes("貼り付け") ||
	    standaloneOffline.zoomDialog?.dialog?.title !== "ズーム" ||
	    !standaloneOffline.zoomDialog?.dialog?.accessKeys?.includes("7") ||
	    !standaloneOffline.zoomDialog?.dialog?.accessKeys?.includes("F") ||
	    !standaloneOffline.zoomDialog?.dialog?.accessKeys?.includes("C") ||
	    !standaloneOffline.zoomDialog?.dialog?.accessKeysUnderlined ||
	    standaloneOffline.zoomDialog?.applied?.zoom !== 75 ||
	    !standaloneOffline.zoomDialog?.applied?.label?.includes("75") ||
	    !standaloneOffline.zoomDialog?.applied?.status?.includes("ズーム") ||
	    !standaloneOffline.liveChart?.initialHasChart ||
	    !standaloneOffline.liveChart?.initialIncludesLabel ||
	    !standaloneOffline.liveChart?.updated?.hasChart ||
	    !standaloneOffline.liveChart?.updated?.includesUpdatedLabel ||
	    !standaloneOffline.liveChart?.updated?.valueSourceChanged ||
	    standaloneBlockedRequests.length
	  ) {
    throw new Error("Standalone HTML should load the full WebSheet UI from a local file without external JS/CSS requests, async Clipboard API permission prompts, or dead chart references.");
  }


  if (
    standaloneRemoteFallback.firstTab !== "CHT" ||
    standaloneRemoteFallback.cellCount <= 0 ||
    !standaloneRemoteFallback.hasRibbon ||
    !standaloneRemoteFallback.hasHomePanel ||
    !standaloneRemoteFallback.standaloneModelLoaded ||
    standaloneRemoteFallback.runtimeSource !== "remote" ||
    !standaloneRemoteFallback.loadingStatusCleared
  ) {
    throw new Error("Standalone HTML should fall back to /download/runtime when the local common runtime is missing or outdated.");
  }


  if (
    !standaloneEmptyModel.emptyVisible ||
    !standaloneEmptyModel.workspaceHidden ||
    standaloneEmptyModel.tabsBeforeLoad !== 0 ||
    standaloneEmptyModel.runtimeSource !== "local" ||
    !standaloneEmptyModel.standaloneFlag ||
    standaloneEmptyModel.modelLoaded ||
    !standaloneEmptyModel.rawModelPresent ||
    !standaloneEmptyModel.status?.includes("Excel ファイルを選択") ||
    standaloneEmptyModel.alert ||
    standaloneEmptyModel.blockedRequests?.length
  ) {
    throw new Error("Standalone HTML with an empty workbook payload should open the local launcher instead of failing with a malformed workbook error.");
  }


  if (
    !standaloneMalformedModel.emptyVisible ||
    !standaloneMalformedModel.workspaceHidden ||
    standaloneMalformedModel.runtimeSource !== "local" ||
    !standaloneMalformedModel.standaloneFlag ||
    standaloneMalformedModel.modelLoaded ||
    standaloneMalformedModel.rawModelPresent ||
    !standaloneMalformedModel.modelError ||
    !standaloneMalformedModel.status?.includes("HTML ワークブックのデータ") ||
    !standaloneMalformedModel.alert?.includes("HTML ワークブックのデータ") ||
    standaloneMalformedModel.alert?.includes("Unexpected end of JSON input") ||
    standaloneMalformedModel.blockedRequests?.length
  ) {
    throw new Error("Standalone HTML with malformed workbook JSON should still open the local launcher and show a Japanese workbook-data error.");
  }


  if (
    standaloneWrappedModel.firstTab !== "CHT" ||
    standaloneWrappedModel.cellCount <= 0 ||
    !standaloneWrappedModel.hasRibbon ||
    !standaloneWrappedModel.hasHomePanel ||
    !standaloneWrappedModel.workspaceVisible ||
    standaloneWrappedModel.runtimeSource !== "local" ||
    !standaloneWrappedModel.standaloneFlag ||
    !standaloneWrappedModel.modelLoaded ||
    !standaloneWrappedModel.rawModelPresent ||
    !standaloneWrappedModel.sourceName?.endsWith("websheet-wrapped-model.html") ||
    !standaloneWrappedModel.status?.includes("HTML ワークブック") ||
    standaloneWrappedModel.blockedRequests?.length
  ) {
    throw new Error("Standalone HTML should open wrapped legacy workbook payloads such as { workbook: model } without losing the workbook.");
  }


  if (
    !standaloneLauncher.beforeLoad.emptyVisible ||
    !standaloneLauncher.beforeLoad.workspaceHidden ||
    standaloneLauncher.beforeLoad.tabsBeforeLoad !== 0 ||
    standaloneLauncher.beforeLoad.runtimeSource !== "local" ||
    !standaloneLauncher.beforeLoad.standaloneFlag ||
    standaloneLauncher.beforeLoad.modelLoaded ||
    standaloneLauncher.firstTab !== "CHT" ||
    !standaloneLauncher.workspaceVisible ||
    standaloneLauncher.cellCount <= 0 ||
    standaloneLauncher.runtimeSource !== "local" ||
    !standaloneLauncher.standaloneFlag ||
    standaloneLauncher.modelLoaded ||
    standaloneLauncher.saveAsDefault?.type !== "xlsx" ||
    standaloneLauncher.saveAsDefault?.activeType !== "xlsx" ||
    !standaloneLauncher.saveAsDefault?.fileName?.endsWith(".xlsx") ||
    !standaloneLauncher.saveAsDefault?.activeFileName?.endsWith(".xlsx") ||
    !standaloneLauncher.saveAsDefault?.note?.includes("Excel Open XML") ||
    !standaloneLauncher.quickSave?.fileName?.endsWith(".xlsx") ||
    standaloneLauncher.quickSave?.zipHeader !== "504b0304" ||
    standaloneLauncher.quickSave?.isHtml === true ||
    !standaloneLauncher.quickSave?.status?.includes("保存しました") ||
    !standaloneLauncher.quickSave?.workbookName?.endsWith(".xlsx") ||
    launcherBlockedRequests.length
  ) {
    throw new Error("Standalone launcher HTML should boot without embedded workbook data, open and convert an Excel file locally, default Save As to .xlsx, and quick-save that local Excel workbook as .xlsx.");
  }


  if (ribbon.title !== expectedReloadWorkbookName) {
    throw new Error("Document title does not match the workbook file name.");
  }


  if (result.addButton !== "+" || result.footerTabs.length < 1) {
    throw new Error("Bottom sheet tabs or add-sheet button did not render.");
  }


  if (result.scrollMax > 0 && bottomBar.hostScrollLeft < 250) {
    throw new Error("Bottom horizontal scrollbar did not sync with the sheet viewport.");
  }


  if (Math.abs(footerBeforeScroll.beforeBottom - footerBeforeScroll.viewportHeight) > 2) {
    throw new Error("Footer is not pinned to the bottom of the viewport.");
  }


  if (Math.abs(footerBeforeScroll.beforeTop - footerBeforeScroll.afterTop) > 2 || footerBeforeScroll.pageScrollY !== 0) {
    throw new Error("Footer moved during vertical sheet scrolling.");
  }


  if (
    Math.abs(footerLayout.ratio - 0.6) > 0.04 ||
    !footerLayout.scrollbarAboveToolbar ||
    footerLayout.footer.height > 34 ||
    footerLayout.toolbar.height > 18 ||
    footerLayout.visibleViewText !== "" ||
    footerLayout.markCount !== 3 ||
    footerLayout.initialZoomLabel !== "118%" ||
    footerLayout.zoom100Label !== "100%" ||
    !footerLayout.wheelGrowDefaultPrevented ||
    !footerLayout.wheelShrinkDefaultPrevented ||
    footerLayout.zoomAfterWheelGrow !== "128%" ||
    footerLayout.zoomAfterWheelShrink !== "118%" ||
    footerLayout.restoredZoomLabel !== "118%" ||
    footerLayout.b14At100 >= footerLayout.b14Before ||
    footerLayout.b14AfterWheelGrow <= footerLayout.b14Before ||
    Math.abs(footerLayout.b14AfterWheelShrink - footerLayout.b14Before) > 1 ||
    Math.abs(footerLayout.b14Restored - footerLayout.b14Before) > 1 ||
    !realCtrlWheelZoom.hasHost ||
    realCtrlWheelZoom.growLabels.join("|") !== "128%|138%|148%|158%" ||
    realCtrlWheelZoom.shrinkLabels.join("|") !== "148%|138%|128%|118%" ||
    realCtrlWheelZoom.undoLabel !== "128%" ||
    realCtrlWheelZoom.redoLabel !== "118%" ||
    !realCtrlWheelZoom.undoStatus?.includes("元に戻しました") ||
    !realCtrlWheelZoom.redoStatus?.includes("やり直しました") ||
    footerLayout.activeAfterPageLayout !== "ページレイアウト" ||
    footerLayout.activeAfterNormal !== "標準" ||
    footerLayout.viewLabels.join("|") !== "標準|ページレイアウト|改ページプレビュー"
  ) {
    throw new Error("Footer should initialize at a 6:4 tab/scroll split, place the horizontal scrollbar above view controls, and expose continuous per-sheet zoom controls.");
  }


  if (
    statusBarStats.singleCellStatusBar ||
    !statusBarStats.before.includes("平均：1.5") ||
    !statusBarStats.before.includes("データの個数：2") ||
    !statusBarStats.before.includes("合計：3") ||
    !statusBarStats.menuLabels.includes("平均") ||
    !statusBarStats.menuLabels.includes("データの個数") ||
    !statusBarStats.menuLabels.includes("数値の個数") ||
    !statusBarStats.menuLabels.includes("最小値") ||
    !statusBarStats.menuLabels.includes("最大値") ||
    !statusBarStats.menuLabels.includes("合計") ||
    !statusBarStats.menuLabels.includes("表示ショートカット") ||
    !statusBarStats.menuLabels.includes("ズーム スライダー") ||
    !statusBarStats.menuLabels.includes("ズーム") ||
    statusBarStats.focusedBefore !== "average" ||
    statusBarStats.focusedAfterDown !== "count" ||
    !statusBarStats.checkedBefore.includes("average") ||
    !statusBarStats.checkedBefore.includes("count") ||
    !statusBarStats.checkedBefore.includes("sum") ||
    statusBarStats.checkedBefore.includes("numericCount") ||
    !statusBarStats.afterNumericCount.includes("数値の個数：2") ||
    statusBarStats.afterHideSum.includes("合計：3") ||
    !statusBarStats.checkedAfter.includes("numericCount") ||
    statusBarStats.checkedAfter.includes("sum") ||
    !statusBarStats.menuClosed ||
    !statusBarStats.viewButtonsVisible ||
    !statusBarStats.zoomSliderVisible ||
    !statusBarStats.zoomLabelVisible
  ) {
    throw new Error("Footer status bar should show Excel-like selection statistics and a right-click customization menu.");
  }


  if (
    !zoomPopup.visible ||
    zoomPopup.title !== "ズーム" ||
    zoomPopup.sectionTitle !== "倍率" ||
    zoomPopup.selectedOption !== "custom" ||
    zoomPopup.customValue !== "118" ||
    !zoomPopup.popupAboveLabel ||
    !zoomPopup.options.includes("200%(0)") ||
    !zoomPopup.options.includes("100%(1)") ||
    !zoomPopup.options.includes("75%(7)") ||
    !zoomPopup.options.includes("50%(5)") ||
    !zoomPopup.options.includes("25%(2)") ||
    !zoomPopup.options.includes("選択範囲をズーム(F)") ||
    !zoomPopup.options.some((option) => option.startsWith("指定(C):")) ||
    !zoomPopup.accessKeys.includes("F") ||
    !zoomPopup.accessKeys.includes("C") ||
    !zoomPopup.accessKeysUnderlined ||
    zoomPopup.afterPreset !== "75%" ||
    !zoomPopup.dialogClosedAfterPreset ||
    zoomPopup.afterCustom !== "118%" ||
    zoomPopup.afterCancel !== "118%" ||
    !zoomPopup.closedAfterCancel
  ) {
    throw new Error("Clicking the zoom percentage should open an Excel-like zoom popup with presets, selection zoom, custom input, OK, and cancel.");
  }


  if (tabScroll.maxScrollLeft <= 0 || tabScroll.afterRight <= 0 || tabScroll.afterLeft >= tabScroll.afterRight) {
    throw new Error("Sheet tab < > buttons did not scroll the tab strip.");
  }


  if (newSheet.beforeRows > 120 || newSheet.beforeCols > 80 || newSheet.borderRightWidth !== "1px" || newSheet.borderBottomWidth !== "1px") {
    throw new Error("New blank sheet does not look like an Excel grid.");
  }


  if (
    !colorPalette.visible ||
    colorPalette.choiceCount < 60 ||
    !colorPalette.titles.includes("テーマ色") ||
    !colorPalette.titles.includes("標準の色") ||
    !colorPalette.hasNoColor ||
    !colorPalette.hasMoreColor ||
    !colorPalette.accessKeys.includes("N") ||
    !colorPalette.accessKeys.includes("M") ||
    !colorPalette.accessKeyTargets.includes("N") ||
    !colorPalette.accessKeyTargets.includes("M") ||
    !colorPalette.accessKeysUnderlined ||
    !colorPalette.focusedBefore ||
    !colorPalette.focusedAfterRight ||
    colorPalette.focusedBefore === colorPalette.focusedAfterRight ||
    colorPalette.fillInput !== "#123456" ||
    colorPalette.afterPaletteFill !== "#ff0000" ||
    colorPalette.fillInputAfterPalette !== "#ff0000" ||
    colorPalette.afterNoFill === "#ff0000" ||
    colorPalette.afterFontMain !== "rgb(0, 112, 192)"
  ) {
    throw new Error("Ribbon color split buttons should apply current colors and expose an Excel-like palette with theme, standard, no-color, and more-color actions.");
  }


  if (
    headerSelection.column.label !== "B:B" ||
    headerSelection.column.selectedAddress !== "B1" ||
    headerSelection.column.selectedCount < newSheet.beforeMaxRow ||
    headerSelection.row.label !== "3:3" ||
    headerSelection.row.selectedAddress !== "A3" ||
    headerSelection.row.selectedCount < newSheet.beforeMaxCol ||
    headerSelection.columnShift.label !== "B:D" ||
    headerSelection.columnCtrl.label !== "B:B,D:D" ||
    headerSelection.columnCtrl.frameCount < 2 ||
    headerSelection.columnDrag.label !== "B:D" ||
    headerSelection.rowDrag.label !== "3:5" ||
    headerSelection.cellCtrl.label !== "B2,D2" ||
    headerSelection.cellCtrl.frameCount < 2 ||
    !headerSelection.cellCtrl.frameClasses.every((className) => className.includes("is-multi-range")) ||
    headerSelection.cellCtrl.fillHandleCount !== 0 ||
    !headerSelection.cellCtrl.frameBorderWidths.every((width) => width === "1px") ||
    !headerSelection.cellCtrl.selectedAddresses.includes("B2") ||
    !headerSelection.cellCtrl.selectedAddresses.includes("D2") ||
    headerSelection.columnContextMenu.label !== "B:B" ||
    !headerSelection.columnContextMenu.labels.includes("列の幅(C)...") ||
    !headerSelection.columnContextMenu.actions.includes("insert-entire-column") ||
    !headerSelection.columnContextMenu.actions.includes("delete-entire-column") ||
    !headerSelection.columnContextMenu.accessKeysUnderlined ||
    headerSelection.columnKeyboardContextMenu?.label !== "B:B" ||
    !headerSelection.columnKeyboardContextMenu?.ariaLabel?.includes("B column") ||
    headerSelection.columnKeyboardContextMenu?.focusedAction !== "cut" ||
    !headerSelection.columnKeyboardContextMenu?.actions?.includes("column-width") ||
    !headerSelection.columnKeyboardContextMenu?.actions?.includes("insert-entire-column") ||
    !headerSelection.columnWidthDialog.visible ||
    headerSelection.columnWidthDialog.title !== "列の幅" ||
    headerSelection.columnWidthDialog.caption !== "列の幅(C):" ||
    headerSelection.columnWidthDialog.inputType !== "text" ||
    headerSelection.columnWidthDialog.ok !== "OK" ||
    headerSelection.columnWidthDialog.cancel !== "キャンセル" ||
    !headerSelection.columnWidthDialog.accessKeys.includes("C") ||
    !headerSelection.columnWidthDialog.accessKeysUnderlined ||
    headerSelection.columnWidthBrowserDialogs !== 0 ||
    headerSelection.columnWidthDialogApply.width < headerSelection.columnWidthDialogApply.target - 2 ||
    !headerSelection.columnWidthDialogApply.status?.includes("列の幅") ||
    headerSelection.rowContextMenu.label !== "3:3" ||
    !headerSelection.rowContextMenu.labels.includes("行の高さ(R)...") ||
    !headerSelection.rowContextMenu.actions.includes("insert-entire-row") ||
    !headerSelection.rowContextMenu.actions.includes("delete-entire-row") ||
    !headerSelection.rowContextMenu.accessKeysUnderlined ||
    headerSelection.rowKeyboardContextMenu?.label !== "3:3" ||
    !headerSelection.rowKeyboardContextMenu?.ariaLabel?.includes("3 row") ||
    headerSelection.rowKeyboardContextMenu?.focusedAction !== "cut" ||
    !headerSelection.rowKeyboardContextMenu?.actions?.includes("row-height") ||
    !headerSelection.rowKeyboardContextMenu?.actions?.includes("insert-entire-row") ||
    !headerSelection.rowHeightDialog.visible ||
    headerSelection.rowHeightDialog.title !== "行の高さ" ||
    headerSelection.rowHeightDialog.caption !== "行の高さ(R):" ||
    headerSelection.rowHeightDialog.inputType !== "text" ||
    headerSelection.rowHeightDialog.ok !== "OK" ||
    headerSelection.rowHeightDialog.cancel !== "キャンセル" ||
    !headerSelection.rowHeightDialog.accessKeys.includes("R") ||
    !headerSelection.rowHeightDialog.accessKeysUnderlined ||
    headerSelection.rowHeightBrowserDialogs !== 0 ||
    headerSelection.rowHeightDialogApply.height < headerSelection.rowHeightDialogApply.target - 2 ||
    !headerSelection.rowHeightDialogApply.status?.includes("行の高さ") ||
    headerSelection.partialMergeRowInsert.sheet !== "CHT" ||
    headerSelection.partialMergeRowInsert.label !== "14:14" ||
    !headerSelection.partialMergeRowInsert.status?.includes("コピーしたセル") ||
    !headerSelection.partialMergeRowInsert.top?.b15?.exists ||
    !headerSelection.partialMergeRowInsert.top?.b15?.text?.includes("■評価") ||
    !headerSelection.partialMergeRowInsert.top?.b16?.exists ||
    headerSelection.partialMergeRowInsert.top?.b16?.text !== "保有技術" ||
    !headerSelection.partialMergeRowInsert.down?.b16?.exists ||
    headerSelection.partialMergeRowInsert.down?.b16?.text !== "保有技術" ||
    !headerSelection.partialMergeRowInsert.up?.b16?.exists ||
    headerSelection.partialMergeRowInsert.up?.b16?.text !== "保有技術" ||
    headerSelection.rowCopiedContextMenu.label !== "4:4" ||
    !headerSelection.rowCopiedContextMenu.labels.includes("コピーしたセルの挿入(E)") ||
    headerSelection.rowCopiedContextMenu.labels.includes("挿入(I)") ||
    !headerSelection.rowCopiedContextMenu.actions.includes("insert-copied-row") ||
    headerSelection.rowCopiedContextMenu.actions.includes("insert-entire-row") ||
    !headerSelection.rowCopiedContextMenu.accessKeysUnderlined ||
    headerSelection.rowCopiedInsert.label !== "4:4" ||
    headerSelection.rowCopiedInsert.a2 !== "CopiedRow" ||
    headerSelection.rowCopiedInsert.a4 !== "CopiedRow" ||
    headerSelection.rowCopiedInsert.a5 !== "RowKeep" ||
    !headerSelection.rowCopiedInsert.status?.includes("コピーしたセル") ||
    headerSelection.columnCutContextMenu.label !== "D:D" ||
    !headerSelection.columnCutContextMenu.labels.includes("コピーしたセルの挿入(E)") ||
    headerSelection.columnCutContextMenu.labels.includes("挿入(I)") ||
    !headerSelection.columnCutContextMenu.actions.includes("insert-copied-column") ||
    headerSelection.columnCutContextMenu.actions.includes("insert-entire-column") ||
    !headerSelection.columnCutContextMenu.accessKeysUnderlined ||
    headerSelection.columnCutBeforeInsert.label !== "B:B" ||
    headerSelection.columnCutBeforeInsert.b1 !== "CopiedCol" ||
    !headerSelection.columnCutBeforeInsert.status?.includes("切り取り") ||
    headerSelection.columnCutInsert.label !== "C:C" ||
    headerSelection.columnCutInsert.b1 === "CopiedCol" ||
    headerSelection.columnCutInsert.c1 !== "CopiedCol" ||
    headerSelection.columnCutInsert.d1 !== "ColKeep" ||
    !headerSelection.columnCutInsert.status?.includes("切り取ったセル")
  ) {
    throw new Error("Column, row, and Ctrl cell selection should support click, drag, Shift range, Ctrl multi-select visuals, and Excel-like header context menus.");
  }


  if (!selectionAutoScroll.right.changed || !selectionAutoScroll.down.changed || !selectionAutoScroll.left.changed || !selectionAutoScroll.up.changed) {
    throw new Error("Dragging a cell selection toward any viewport edge should auto-scroll the visible sheet range.");
  }


  if (extension.maxRowAfter <= newSheet.beforeMaxRow || extension.maxColAfter <= newSheet.beforeMaxCol || extension.rowsAfter > 120 || extension.colsAfter > 80) {
    throw new Error("Rows/columns did not extend when scrolling to the sheet edge.");
  }


  if (
    lazyLoadAnchor.after.maxRow <= lazyLoadAnchor.before.maxRow ||
    lazyLoadAnchor.after.maxCol <= lazyLoadAnchor.before.maxCol ||
    Math.abs(lazyLoadAnchor.after.scrollTop - lazyLoadAnchor.before.scrollTop) > 2 ||
    Math.abs(lazyLoadAnchor.after.scrollLeft - (lazyLoadAnchor.before.scrollLeft + lazyLoadAnchor.after.rowHeaderWidth - lazyLoadAnchor.before.rowHeaderWidth)) > 2 ||
    lazyLoadAnchor.after.bottomRemaining <= lazyLoadAnchor.before.bottomRemaining + 200 ||
    lazyLoadAnchor.after.rightRemaining <= lazyLoadAnchor.before.rightRemaining + 200
  ) {
    throw new Error("Lazy-loading should extend the sheet while keeping the current visible range anchored.");
  }


  if (
    lazyLoadAnchor.repeat.maxRow !== lazyLoadAnchor.after.maxRow ||
    lazyLoadAnchor.repeat.maxCol !== lazyLoadAnchor.after.maxCol ||
    Math.abs(lazyLoadAnchor.dragBack.scrollTop - lazyLoadAnchor.dragBack.targetTop) > 2 ||
    Math.abs(lazyLoadAnchor.dragBack.scrollLeft - lazyLoadAnchor.dragBack.targetLeft) > 2 ||
    Math.abs(lazyLoadAnchor.dragBack.scrollHeight - lazyLoadAnchor.repeat.scrollHeight) > 4 ||
    Math.abs(lazyLoadAnchor.dragBack.scrollWidth - lazyLoadAnchor.repeat.scrollWidth) > 4 ||
    lazyLoadAnchor.dragBack.maxRow > lazyLoadAnchor.repeat.maxRow ||
    lazyLoadAnchor.dragBack.maxCol > lazyLoadAnchor.repeat.maxCol
  ) {
    throw new Error("Lazy-loading should not repeatedly fire while still locked at the edge, and active drag-back should keep both the requested scrollbar position and the current scrollbar length.");
  }


  const horizontalScrollStillExtended =
    trimmedExtension.scrollWidthAfterReturn >= extension.scrollWidthAfter &&
    trimmedExtension.scrollWidthAfterReturn > (trimmedExtension.clientWidthAfterReturn || 0) + 4;


  if (
    trimmedExtension.maxRowAfterReturn >= extension.maxRowAfter ||
    trimmedExtension.maxColAfterReturn >= extension.maxColAfter ||
    trimmedExtension.scrollHeightAfterReturn >= extension.scrollHeightAfter ||
    horizontalScrollStillExtended ||
    trimmedExtension.rowsAfterReturn > 120 ||
    trimmedExtension.colsAfterReturn > 80 ||
    trimmedExtension.rowHeaderWidthAfterReturn >= 46 ||
    trimmedExtension.rowHeaderWidthAfterReturn > extension.rowHeaderWidthAfter
  ) {
    throw new Error("Rows/columns should lazy-unload empty extended ranges after scrolling back from the sheet edge.");
  }


  if (extension.aqHeader.text !== "AQ" || extension.aqHeader.textOverflow !== "clip" || extension.aqHeader.labelTextOverflow !== "clip") {
    throw new Error("Column headers should clip narrow labels such as AQ instead of rendering an ellipsis.");
  }


  if (resizedHeaders.afterColWidth <= resizedHeaders.colWidth + 25 || resizedHeaders.afterRowHeight <= resizedHeaders.rowHeight + 10) {
    throw new Error("Column and row headers should resize column width and row height by dragging.");
  }


  if (
    Math.abs(headerResizeUndoRedo.undoRowHeight - resizedHeaders.rowHeight) > 2 ||
    !headerResizeUndoRedo.undoStatus?.includes("元に戻しました") ||
    headerResizeUndoRedo.redoRowHeight <= resizedHeaders.rowHeight + 10 ||
    !headerResizeUndoRedo.redoStatus?.includes("やり直しました")
  ) {
    throw new Error("Row height resize should be undoable and redoable.");
  }


  if (
    multiHeaderResize.columnWidths.length !== 3 ||
    Math.max(...multiHeaderResize.columnWidths) - Math.min(...multiHeaderResize.columnWidths) > 1 ||
    multiHeaderResize.rowLabel !== "4:6" ||
    multiHeaderResize.rowHeights.length !== 3 ||
    Math.max(...multiHeaderResize.rowHeights) - Math.min(...multiHeaderResize.rowHeights) > 1
  ) {
    throw new Error("Resizing a header inside a multi-row or multi-column selection should apply the same size to every selected row or column.");
  }


  if (
    !resizeHandleVisual.bodyResizing ||
    resizeHandleVisual.highlightedColumnHandles > 1 ||
    resizeHandleVisual.highlightedRowHandles > 1 ||
    resizeHandleVisual.highlightedColumnHandles === resizeHandleVisual.columnHandleCount ||
    resizeHandleVisual.highlightedRowHandles === resizeHandleVisual.rowHandleCount
  ) {
    throw new Error("Resizing one header should not visually thicken every row or column resize interval.");
  }


  if (
    !whiteFillGridMask.filled.className?.includes("has-fill") ||
    whiteFillGridMask.filled.fillColor !== "#ffffff" ||
    whiteFillGridMask.filled.backgroundColor !== "rgba(0, 0, 0, 0)" ||
    whiteFillGridMask.filled.borderRightColor !== "rgba(0, 0, 0, 0)" ||
    whiteFillGridMask.filled.borderBottomColor !== "rgba(0, 0, 0, 0)" ||
    whiteFillGridMask.leftNeighbor.borderRightColor !== "rgba(0, 0, 0, 0)" ||
    whiteFillGridMask.topNeighbor.borderBottomColor !== "rgba(0, 0, 0, 0)" ||
    whiteFillGridMask.untouched.borderRightColor !== "rgba(0, 0, 0, 0)" ||
    whiteFillGridMask.untouched.borderBottomColor !== "rgba(0, 0, 0, 0)"
  ) {
    throw new Error("Explicit white fill should hide adjacent default gridlines while leaving untouched blank gridlines visible.");
  }


  if (!ribbon.homeVisible || ribbon.activeRibbonTab !== "ホーム") {
    throw new Error("Home ribbon panel did not activate.");
  }


  if (!/span\s+3/.test(ribbon.mergedGridColumn || "") || ribbon.selectionLabel !== "B14:D14") {
    throw new Error("Merge selection did not create a merged B14:D14 cell.");
  }


  if (
    /span\s+3/.test(mergeUndoRedo.undoB14GridColumn || "") ||
    !mergeUndoRedo.undoD14Rendered ||
    !mergeUndoRedo.undoStatus?.includes("元に戻しました") ||
    !/span\s+3/.test(mergeUndoRedo.redoB14GridColumn || "") ||
    mergeUndoRedo.redoD14Rendered ||
    mergeUndoRedo.redoLabel !== "B14:D14" ||
    !mergeUndoRedo.redoStatus?.includes("やり直しました")
  ) {
    throw new Error("Merge cell actions should be undoable and redoable.");
  }


  if (!mergeUndoRedo.redoMergeButtonActive) {
    throw new Error("Merge button should show selected state when the selected cell is already merged.");
  }


  if (
    mergeUndoRedo.toggleUnmergeButtonActive ||
    !mergeUndoRedo.toggleUnmergeD14Rendered ||
    !mergeUndoRedo.toggleUnmergeStatus?.includes("解除") ||
    !mergeUndoRedo.toggleRemergeButtonActive ||
    mergeUndoRedo.toggleRemergeD14Rendered
  ) {
    throw new Error("Clicking the active Merge & Center button should unmerge, then click again should merge the selected range.");
  }


  if (
    mergeAutoFillMismatch.label !== "B14:D14" ||
    mergeAutoFillMismatch.topAlert !== "" ||
    mergeAutoFillMismatch.messageTitle !== "Cht WebSheet" ||
    mergeAutoFillMismatch.okButtonText !== "OK" ||
    !mergeAutoFillMismatch.messageText?.includes("結合セル") ||
    !mergeAutoFillMismatch.messageText?.includes("オートフィルできません") ||
    !mergeAutoFillMismatch.status?.includes("結合セル") ||
    /span\s+3/.test(mergeAutoFillMismatch.e14GridColumn || "") ||
    !mergeAutoFillMismatch.f14Rendered
  ) {
    throw new Error("Auto-fill should reject merged-cell targets whose shape does not match the source selection.");
  }


  if (
    mergedSingleSelection.label !== "B14:D14" ||
    !mergedSingleSelection.mergeButtonActive ||
    Math.abs(mergedSingleSelection.frame.left - mergedSingleSelection.merged.left) > 2 ||
    Math.abs(mergedSingleSelection.frame.top - mergedSingleSelection.merged.top) > 2 ||
    Math.abs(mergedSingleSelection.frame.right - mergedSingleSelection.merged.right) > 2 ||
    Math.abs(mergedSingleSelection.frame.bottom - mergedSingleSelection.merged.bottom) > 2
  ) {
    throw new Error("Selecting a merged cell should frame the full merged range.");
  }


  if (
    mergedKeyboardNavigation.afterRight !== "E14" ||
    mergedKeyboardNavigation.afterLeftIntoMerge !== "B14:D14" ||
    mergedKeyboardNavigation.afterEnter !== "B15:B43" ||
    mergedKeyboardNavigation.afterUpIntoMerge !== "B14:D14"
  ) {
    throw new Error("Keyboard navigation should enter and leave merged cells using the full merged selection range.");
  }


  if (
    mergedCopyState.label !== "B14:D14" ||
    !mergedCopyState.frameClass?.includes("is-copy-source") ||
    mergedCopyState.copyFrameCount !== 0 ||
    mergedCopyState.animationName !== "marching-ants" ||
    !mergedCopyState.backgroundImage?.includes("34, 197, 94") ||
    mergedCopyState.clipboardStats?.clipboardSlots !== 3 ||
    mergedCopyState.clipboardStats?.clipboardStoredCells > mergedCopyState.clipboardStats?.clipboardSlots ||
    mergedCopyState.clipboardStats?.clipboardDenseSlots !== 0
  ) {
    throw new Error("Copying a merged cell should show the green animated copy frame and store the internal clipboard sparsely.");
  }


  if (
    mergedPaste.label !== "E14:G14" ||
    !/span\s+3/.test(mergedPaste.pastedGridColumn || "") ||
    !mergedPaste.pastedText?.includes("D：知見あり") ||
    mergedPaste.f14Rendered ||
    mergedPaste.g14Rendered ||
    mergedPaste.sourceCopyFrameClass !== "sheet-copy-frame"
  ) {
    throw new Error("Pasting a copied merged cell should preserve the merge range and keep the copied-source frame.");
  }


  if (
    mergedCrossSelection.label !== "A14:D14" ||
    mergedCrossSelection.mergeButtonActive ||
    Math.abs(mergedCrossSelection.frame.left - mergedCrossSelection.leftCell.left) > 2 ||
    Math.abs(mergedCrossSelection.frame.top - mergedCrossSelection.leftCell.top) > 2 ||
    Math.abs(mergedCrossSelection.frame.right - mergedCrossSelection.merged.right) > 2 ||
    Math.abs(mergedCrossSelection.frame.bottom - mergedCrossSelection.merged.bottom) > 2
  ) {
    throw new Error("A multi-cell selection crossing a merged cell should expand to the merged cell's full width and height.");
  }


  if (ribbon.mergedTextAlign !== "center" || Number.parseFloat(ribbon.mergedFontSize) < 20 || !["700", "bold"].includes(String(ribbon.mergedFontWeight))) {
    throw new Error("Home ribbon formatting commands did not apply to the merged selection.");
  }
}

export {
  runAssertionsFormulasAndClipboard,
};
