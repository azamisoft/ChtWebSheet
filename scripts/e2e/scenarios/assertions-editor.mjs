// Sequential E2E feature module.

async function runAssertionsEditor(scenarioState) {
  let { result, coordinateHitTest, insertRibbon, formulaRibbon, formulaDefineNameMain, formulaDefineNameMenu, formulaDefineNameCreate, createNamesDialog, formulaNames, formulaNameManager, formulaUseNameMenu, formulaUseNameManage, formulaUseNameInserted, formulaPasteNameDialog, formulaPasteNameInserted, formulaPasteNameListDialog, formulaPasteNameList, formulaPasteNameListUndo, formulaPasteNameListRedo, formulaFunctionDialog, formulaFunctionSearch, formulaInserted, formulaFunctionArguments, formulaLogicalMenu, formulaMenuArguments, formulaMenuInserted, formulaRecentMenu, formulaMathMenu, formulaVariableArguments, formulaZeroArgumentFunction, formulaFinancialMenu, formulaDateMenu, formulaTextMenu, formulaMoreMenu, formulaAutoSumMenu, formulaAutoSumInserted, formulaShowState, formulaTrace, formulaRemoveArrowsMenu, formulaRemovePrecedents, formulaTraceDependents, formulaRemoveDependents, formulaEvaluateInitial, formulaEvaluateStepIn, formulaEvaluateStep, formulaEvaluateStepOut, formulaEvaluateReferenceInitial, formulaEvaluateReferenceStep, formulaEvaluateReferenceStepIn, formulaEvaluateReferenceStepOut, formulaWatchWindow, formulaWatchSelect, formulaWatchDelete, formulaErrorMenu, formulaCircularSelection, formulaErrorChecking, formulaErrorIgnore, formulaErrorTrace, formulaErrorResetIgnored, formulaCalculation, viewRibbon, viewFullScreenOn, fullScreenUndoRedo, viewFullScreenOff, sheetViewNew, sheetViewKept, sheetViewMenuBeforeRename, sheetViewRenamed, sheetViewExited, sheetViewReactivated, sheetViewUpdated, sheetViewMenuBeforeCopy, sheetViewDuplicated, viewZoomDialog, viewZoomApplied, viewZoomHundred, viewZoomSelection, viewZoomSelectionUndoRedo, viewRulerOn, pageLayoutHeaderFooterDialog, pageBreakPreview, pageBreakContextMenu, customViewCaptureState, customViewsMainBeforeAdd, customViewAddDialog, customViewsDialog, customViewSaveUndoRedo, customViewApplied, customViewApplyUndoRedo, customViewDeleteFlow, freezeMenuLabels, freezeMenuInitialMeta, freezeMenuTopRow, freezeMenuTopRowStatus, freezeMenuAccessKey, formulaBarUndoRedo, viewToggles, rtlVisualHitTest, gridlinesUndoRedo, headingsUndoRedo, zeroValuesUndoRedo, objectsUndoRedo, outlineUndoRedo, rightToLeftUndoRedo, freezePanesAfterMain, frozenPaneSticky, splitPaneBeforeDrag, splitPaneDragged, sheetGroupView, viewInitialWindowBar, viewNewWindowState, viewNewWindowBar, viewSecondNewWindowState, viewSecondNewWindowBar, workbookWindowBarSwitch, customWindowSource, customWindowMutated, customWindowRestored, arrangeWindowDialog, arrangeWindowState, arrangeWindowBar, arrangeWindowUndo, arrangeWindowRedo, sideBySideWindowDialog, sideBySideDialogApplied, sideBySideWindowBar, syncScrollState, syncScrollUndoRedo, resetWindowState, resetWindowBar, switchWindowMenu, switchWindowKeyboardState, switchWindowState, switchWindowDialog, switchWindowDialogState, hideWindowState, hideWindowBar, unhideWindowDialog, unhideWindowState, unhideWindowBar, unhideWindowUndoRedo, closeWindowState, closeWindowUndoRedo, macrosDialog, macroMenuBeforeRecord, macroRecordDialog, macroRecordingStarted, macroMenuDuringRecord, macroMenuRelativeRefs, macroRelativeRefsUndoRedo, macroRecordingStopped, macroRecordingStopUndoRedo, macroSecurityDialog, macroSecurityApplied, macroSecurityUndoRedo, viewWindowCommands, pageLayout, expectedReloadWorkbookName, column, row, persistenceReload, wideViewportCoverage, autoFilterConditionMenu, importedAutoFilterCriteria, autoFilterBlankKeyboardMenu, autoFilterReapplyClear, advancedFilterCopyOverwrite, autoFilterDateMenu, advancedFilterUniqueDates, advancedFilterUniqueSubsetCopy, delimitedTextImport, powerQueryTransformEditor, whatIfGoalSeekDataTable, scenarioManagerSummary, flashFillDataPatterns, dataCurrentRegion, queryConnectionsPane, dataGetConnectionMenus, relationshipsDialog, refreshAllUndo, forecastSheet, dataAnalysisToolpak, consolidateDialog, subtotalDialog, outlineGrouping, textToColumnsDialog, textToColumnsFixedWidth, textToColumnsAdvanced, sortDialogLevels, sortDialogColorOrder, sortDialogLeftToRight, sortDialogBlankAndCase, sortDialogCustomList, sortDialogDateOrder, sortFormulaResults, quickSortActiveColumn, quickSortPartialSelection, removeDuplicatesDialog } = scenarioState;

  if (
    !autoFilterConditionMenu.textMenu?.includes("テキスト フィルター") ||
    !autoFilterConditionMenu.textSubmenu?.includes("空白セル") ||
    autoFilterConditionMenu.textCaseRows?.find((row) => row.row === 2)?.hidden !== false ||
    autoFilterConditionMenu.textCaseRows?.find((row) => row.row === 3)?.hidden !== false ||
    autoFilterConditionMenu.textCaseRows?.find((row) => row.row === 4)?.hidden !== true ||
    autoFilterConditionMenu.textCaseRows?.find((row) => row.row === 5)?.hidden !== true ||
    autoFilterConditionMenu.textCaseExported?.customFilters?.[0]?.operator !== "equal" ||
    autoFilterConditionMenu.textCaseExported?.customFilters?.[0]?.value !== "AL*" ||
    autoFilterConditionMenu.textWildcardRows?.find((row) => row.row === 2)?.hidden !== true ||
    autoFilterConditionMenu.textWildcardRows?.find((row) => row.row === 3)?.hidden !== true ||
    autoFilterConditionMenu.textWildcardRows?.find((row) => row.row === 4)?.hidden !== true ||
    autoFilterConditionMenu.textWildcardRows?.find((row) => row.row === 5)?.hidden !== false ||
    autoFilterConditionMenu.textWildcardExported?.customFilters?.[0]?.operator !== "equal" ||
    autoFilterConditionMenu.textWildcardExported?.customFilters?.[0]?.value !== "d*" ||
    autoFilterConditionMenu.textNotContainsRows?.find((row) => row.row === 2)?.hidden !== false ||
    autoFilterConditionMenu.textNotContainsRows?.find((row) => row.row === 3)?.hidden !== false ||
    autoFilterConditionMenu.textNotContainsRows?.find((row) => row.row === 4)?.hidden !== true ||
    autoFilterConditionMenu.textNotContainsRows?.find((row) => row.row === 5)?.hidden !== true ||
    autoFilterConditionMenu.textNotContainsExported?.customFilters?.[0]?.operator !== "notEqual" ||
    autoFilterConditionMenu.textNotContainsExported?.customFilters?.[0]?.value !== "*ta*" ||
    !autoFilterConditionMenu.numberMenu?.includes("数値フィルター") ||
    !autoFilterConditionMenu.numberSubmenu?.includes("指定の値より大きい") ||
    !autoFilterConditionMenu.numberSubmenu?.includes("トップ テン") ||
    autoFilterConditionMenu.preset !== "greaterThan" ||
    !autoFilterConditionMenu.rows?.find((row) => row.row === 2)?.hidden ||
    !autoFilterConditionMenu.rows?.find((row) => row.row === 5)?.hidden ||
    autoFilterConditionMenu.rows?.find((row) => row.row === 3)?.hidden ||
    autoFilterConditionMenu.rows?.find((row) => row.row === 4)?.hidden ||
    autoFilterConditionMenu.rows?.find((row) => row.row === 6)?.hidden !== true ||
    autoFilterConditionMenu.top10Defaults?.title !== "トップ テン オートフィルター" ||
    autoFilterConditionMenu.top10Defaults?.direction !== "top" ||
    autoFilterConditionMenu.top10Defaults?.mode !== "items" ||
    autoFilterConditionMenu.top10Rows?.find((row) => row.row === 2)?.hidden !== true ||
    autoFilterConditionMenu.top10Rows?.find((row) => row.row === 3)?.hidden !== true ||
    autoFilterConditionMenu.top10Rows?.find((row) => row.row === 4)?.hidden !== false ||
    autoFilterConditionMenu.top10Rows?.find((row) => row.row === 5)?.hidden !== true ||
    autoFilterConditionMenu.top10Rows?.find((row) => row.row === 6)?.hidden !== true ||
    autoFilterConditionMenu.bottomPercentRows?.find((row) => row.row === 2)?.hidden !== false ||
    autoFilterConditionMenu.bottomPercentRows?.find((row) => row.row === 3)?.hidden !== true ||
    autoFilterConditionMenu.bottomPercentRows?.find((row) => row.row === 4)?.hidden !== true ||
    autoFilterConditionMenu.bottomPercentRows?.find((row) => row.row === 5)?.hidden !== false ||
    autoFilterConditionMenu.bottomPercentRows?.find((row) => row.row === 6)?.hidden !== true ||
    !autoFilterConditionMenu.numberSubmenu?.includes("指定の範囲外") ||
    autoFilterConditionMenu.notBetweenRows?.find((row) => row.row === 2)?.hidden !== false ||
    autoFilterConditionMenu.notBetweenRows?.find((row) => row.row === 3)?.hidden !== true ||
    autoFilterConditionMenu.notBetweenRows?.find((row) => row.row === 4)?.hidden !== true ||
    autoFilterConditionMenu.notBetweenRows?.find((row) => row.row === 5)?.hidden !== false ||
    autoFilterConditionMenu.notBetweenRows?.find((row) => row.row === 6)?.hidden !== true ||
    autoFilterConditionMenu.notBetweenExported?.customFilters?.map((item) => `${item.operator}:${item.value}`).join("|") !== "lessThan:10|greaterThan:20" ||
    !autoFilterConditionMenu.numberSubmenu?.includes("標準偏差より上") ||
    !autoFilterConditionMenu.numberSubmenu?.includes("標準偏差より下") ||
    autoFilterConditionMenu.aboveStdDevRows?.find((row) => row.row === 2)?.hidden !== true ||
    autoFilterConditionMenu.aboveStdDevRows?.find((row) => row.row === 3)?.hidden !== true ||
    autoFilterConditionMenu.aboveStdDevRows?.find((row) => row.row === 4)?.hidden !== false ||
    autoFilterConditionMenu.aboveStdDevRows?.find((row) => row.row === 5)?.hidden !== true ||
    autoFilterConditionMenu.aboveStdDevRows?.find((row) => row.row === 6)?.hidden !== true ||
    !autoFilterConditionMenu.aboveStdDevExported?.dynamicFilters?.some((item) => item.type === "aboveStdDev" && item.val === "1") ||
    autoFilterConditionMenu.belowStdDevRows?.some((row) => row.hidden !== true) ||
    !autoFilterConditionMenu.belowStdDevExported?.dynamicFilters?.some((item) => item.type === "belowStdDev" && item.val === "1") ||
    !autoFilterConditionMenu.colorMenuBefore?.includes("色フィルター") ||
    !autoFilterConditionMenu.colorSubmenu?.includes("セルの色で並べ替え") ||
    !autoFilterConditionMenu.colorSubmenu?.includes("セルの色でフィルター") ||
    !autoFilterConditionMenu.colorSubmenu?.includes("フォントの色で並べ替え") ||
    !autoFilterConditionMenu.colorSubmenu?.includes("フォントの色でフィルター") ||
    !autoFilterConditionMenu.colorChoice ||
    autoFilterConditionMenu.colorRows?.find((row) => row.row === 2)?.hidden !== true ||
    autoFilterConditionMenu.colorRows?.find((row) => row.row === 3)?.hidden !== true ||
    autoFilterConditionMenu.colorRows?.find((row) => row.row === 4)?.hidden !== false ||
    autoFilterConditionMenu.colorRows?.find((row) => row.row === 5)?.hidden !== true ||
    !autoFilterConditionMenu.fontColorChoice ||
    autoFilterConditionMenu.fontColorRows?.find((row) => row.row === 2)?.hidden !== true ||
    autoFilterConditionMenu.fontColorRows?.find((row) => row.row === 3)?.hidden !== false ||
    autoFilterConditionMenu.fontColorRows?.find((row) => row.row === 4)?.hidden !== true ||
    autoFilterConditionMenu.fontColorRows?.find((row) => row.row === 5)?.hidden !== true ||
    autoFilterConditionMenu.advancedFormulaRows?.find((row) => row.row === 2)?.hidden !== true ||
    autoFilterConditionMenu.advancedFormulaRows?.find((row) => row.row === 3)?.hidden !== false ||
    autoFilterConditionMenu.advancedFormulaRows?.find((row) => row.row === 4)?.hidden !== false ||
    autoFilterConditionMenu.advancedFormulaRows?.find((row) => row.row === 5)?.hidden !== true ||
    autoFilterConditionMenu.colorSortRows?.[0]?.name !== "Beta" ||
    autoFilterConditionMenu.colorSortRows?.[0]?.score !== "20" ||
    !autoFilterConditionMenu.colorSortStatus?.includes("セルの色") ||
    autoFilterConditionMenu.fontColorSortRows?.[0]?.name !== "Alpine" ||
    autoFilterConditionMenu.fontColorSortRows?.[0]?.score !== "12" ||
    !autoFilterConditionMenu.fontColorSortStatus?.includes("フォントの色")
  ) {
    throw new Error("AutoFilter should expose condition menus, color sort/filter commands, and advanced formula criteria like Excel.");
  }


  if (
    importedAutoFilterCriteria.valueFilter?.buttonFiltered !== true ||
    !Array.isArray(importedAutoFilterCriteria.valueFilter?.criteria) ||
    importedAutoFilterCriteria.valueFilter?.criteria?.join("|") !== "Alpha" ||
    importedAutoFilterCriteria.valueFilter?.rows?.find((row) => row.row === 2)?.hidden !== false ||
    importedAutoFilterCriteria.valueFilter?.rows?.find((row) => row.row === 3)?.hidden !== true ||
    importedAutoFilterCriteria.valueFilter?.rows?.find((row) => row.row === 4)?.hidden !== false ||
    importedAutoFilterCriteria.valueFilter?.rows?.find((row) => row.row === 5)?.hidden !== true ||
    importedAutoFilterCriteria.customFilter?.buttonFiltered !== true ||
    importedAutoFilterCriteria.customFilter?.criteria?.operator1 !== "greaterThan" ||
    importedAutoFilterCriteria.customFilter?.criteria?.value1 !== "10" ||
    importedAutoFilterCriteria.customFilter?.rows?.find((row) => row.row === 2)?.hidden !== true ||
    importedAutoFilterCriteria.customFilter?.rows?.find((row) => row.row === 3)?.hidden !== false ||
    importedAutoFilterCriteria.customFilter?.rows?.find((row) => row.row === 4)?.hidden !== false ||
    importedAutoFilterCriteria.customFilter?.rows?.find((row) => row.row === 5)?.hidden !== true ||
    importedAutoFilterCriteria.top10Filter?.buttonFiltered !== true ||
    importedAutoFilterCriteria.top10Filter?.criteria?.top10?.count !== 2 ||
    importedAutoFilterCriteria.top10Filter?.criteria?.top10?.direction !== "top" ||
    importedAutoFilterCriteria.top10Filter?.rows?.find((row) => row.row === 2)?.hidden !== true ||
    importedAutoFilterCriteria.top10Filter?.rows?.find((row) => row.row === 3)?.hidden !== false ||
    importedAutoFilterCriteria.top10Filter?.rows?.find((row) => row.row === 4)?.hidden !== false ||
    importedAutoFilterCriteria.top10Filter?.rows?.find((row) => row.row === 5)?.hidden !== true ||
    importedAutoFilterCriteria.averageFilter?.buttonFiltered !== true ||
    importedAutoFilterCriteria.averageFilter?.criteria?.averageOperator !== "aboveAverage" ||
    importedAutoFilterCriteria.averageFilter?.rows?.find((row) => row.row === 2)?.hidden !== true ||
    importedAutoFilterCriteria.averageFilter?.rows?.find((row) => row.row === 3)?.hidden !== true ||
    importedAutoFilterCriteria.averageFilter?.rows?.find((row) => row.row === 4)?.hidden !== false ||
    importedAutoFilterCriteria.averageFilter?.rows?.find((row) => row.row === 5)?.hidden !== false ||
    importedAutoFilterCriteria.dateGroupFilter?.buttonFiltered !== true ||
    importedAutoFilterCriteria.dateGroupFilter?.criteria?.dateGroup?.grouping !== "month" ||
    importedAutoFilterCriteria.dateGroupFilter?.criteria?.dateGroup?.month !== 5 ||
    importedAutoFilterCriteria.dateGroupFilter?.rows?.find((row) => row.row === 2)?.hidden !== false ||
    importedAutoFilterCriteria.dateGroupFilter?.rows?.find((row) => row.row === 3)?.hidden !== false ||
    importedAutoFilterCriteria.dateGroupFilter?.rows?.find((row) => row.row === 4)?.hidden !== true ||
    importedAutoFilterCriteria.dateGroupFilter?.rows?.find((row) => row.row === 5)?.hidden !== true ||
    importedAutoFilterCriteria.blankFilter?.buttonFiltered !== true ||
    importedAutoFilterCriteria.blankFilter?.criteria?.join("|") !== "" ||
    importedAutoFilterCriteria.blankFilter?.rows?.find((row) => row.row === 2)?.hidden !== true ||
    importedAutoFilterCriteria.blankFilter?.rows?.find((row) => row.row === 3)?.hidden !== false ||
    importedAutoFilterCriteria.blankFilter?.rows?.find((row) => row.row === 4)?.hidden !== false ||
    importedAutoFilterCriteria.blankFilter?.rows?.find((row) => row.row === 5)?.hidden !== true ||
    importedAutoFilterCriteria.savedSort?.hasFilterButton !== true ||
    importedAutoFilterCriteria.savedSort?.buttonSorted !== true ||
    importedAutoFilterCriteria.savedSort?.buttonSortDirection !== "desc" ||
    importedAutoFilterCriteria.savedSort?.sortState?.hasHeader !== true ||
    importedAutoFilterCriteria.savedSort?.sortState?.keys?.[0]?.col !== 2 ||
    importedAutoFilterCriteria.savedSort?.sortState?.keys?.[0]?.direction !== "desc" ||
    importedAutoFilterCriteria.savedSort?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "C:20|B:12|D:8|A:5" ||
    importedAutoFilterCriteria.standaloneSort?.hasFilterButton !== false ||
    importedAutoFilterCriteria.standaloneSort?.sortState?.hasHeader !== false ||
    importedAutoFilterCriteria.standaloneSort?.sortState?.keys?.[0]?.col !== 1 ||
    importedAutoFilterCriteria.standaloneSort?.sortState?.keys?.[0]?.direction !== "asc" ||
    importedAutoFilterCriteria.standaloneSort?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "Alpha:1|Bravo:2|Charlie:3|Delta:4" ||
    importedAutoFilterCriteria.tableFilter?.hasFilterButton !== true ||
    !Array.isArray(importedAutoFilterCriteria.tableFilter?.criteria) ||
    importedAutoFilterCriteria.tableFilter?.criteria?.join("|") !== "Alpha" ||
    importedAutoFilterCriteria.tableFilter?.rows?.find((row) => row.row === 2)?.hidden !== false ||
    importedAutoFilterCriteria.tableFilter?.rows?.find((row) => row.row === 3)?.hidden !== true ||
    importedAutoFilterCriteria.tableFilter?.rows?.find((row) => row.row === 4)?.hidden !== false ||
    importedAutoFilterCriteria.tableFilter?.rows?.find((row) => row.row === 5)?.hidden !== true ||
    importedAutoFilterCriteria.tableSortOnly?.before?.buttonSorted !== true ||
    importedAutoFilterCriteria.tableSortOnly?.before?.buttonSortDirection !== "desc" ||
    importedAutoFilterCriteria.tableSortOnly?.before?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "C:20|B:12|D:8|A:5" ||
    importedAutoFilterCriteria.tableSortOnly?.ribbonBeforeClear?.clear?.disabled !== false ||
    importedAutoFilterCriteria.tableSortOnly?.ribbonBeforeClear?.reapply?.disabled !== false ||
    importedAutoFilterCriteria.tableSortOnly?.exportedBeforeClear?.excelJsLoaded !== true ||
    importedAutoFilterCriteria.tableSortOnly?.exportedBeforeClear?.tableXmlHasSortState ||
    !importedAutoFilterCriteria.tableSortOnly?.exportedBeforeClear?.worksheetSortStateCount ||
    importedAutoFilterCriteria.tableSortOnly?.afterClear?.buttonSorted !== false ||
    importedAutoFilterCriteria.tableSortOnly?.afterClear?.tableSortState ||
    importedAutoFilterCriteria.importedCellColorFilter?.criteria?.kind !== "color" ||
    importedAutoFilterCriteria.importedCellColorFilter?.criteria?.target !== "fill" ||
    importedAutoFilterCriteria.importedCellColorFilter?.buttonFiltered !== true ||
    importedAutoFilterCriteria.importedCellColorFilter?.rows?.find((row) => row.row === 4)?.hidden !== false ||
    importedAutoFilterCriteria.importedCellColorFilter?.rows?.filter((row) => row.row !== 4).some((row) => row.hidden !== true) ||
    importedAutoFilterCriteria.importedFontColorFilter?.criteria?.kind !== "color" ||
    importedAutoFilterCriteria.importedFontColorFilter?.criteria?.target !== "font" ||
    importedAutoFilterCriteria.importedFontColorFilter?.buttonFiltered !== true ||
    importedAutoFilterCriteria.importedFontColorFilter?.rows?.find((row) => row.row === 3)?.hidden !== false ||
    importedAutoFilterCriteria.importedFontColorFilter?.rows?.filter((row) => row.row !== 3).some((row) => row.hidden !== true) ||
    importedAutoFilterCriteria.importedTableColorFilter?.criteria?.kind !== "color" ||
    importedAutoFilterCriteria.importedTableColorFilter?.criteria?.target !== "fill" ||
    importedAutoFilterCriteria.importedTableColorFilter?.buttonFiltered !== true ||
    importedAutoFilterCriteria.importedTableColorFilter?.rows?.find((row) => row.row === 4)?.hidden !== false ||
    importedAutoFilterCriteria.importedTableColorFilter?.rows?.filter((row) => row.row !== 4).some((row) => row.hidden !== true) ||
    importedAutoFilterCriteria.importedCellColorSort?.buttonSorted !== true ||
    importedAutoFilterCriteria.importedCellColorSort?.buttonSortDirection !== "asc" ||
    importedAutoFilterCriteria.importedCellColorSort?.sortState?.keys?.[0]?.sortOn !== "cellColor" ||
    importedAutoFilterCriteria.importedCellColorSort?.sortState?.keys?.[0]?.color !== "#ff0000" ||
    importedAutoFilterCriteria.importedCellColorSort?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "B:12|A:5|C:20|D:8" ||
    importedAutoFilterCriteria.importedFontColorSort?.buttonSorted !== true ||
    importedAutoFilterCriteria.importedFontColorSort?.buttonSortDirection !== "asc" ||
    importedAutoFilterCriteria.importedFontColorSort?.sortState?.keys?.[0]?.sortOn !== "fontColor" ||
    importedAutoFilterCriteria.importedFontColorSort?.sortState?.keys?.[0]?.color !== "#ff0000" ||
    importedAutoFilterCriteria.importedFontColorSort?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "D:8|A:5|B:12|C:20" ||
    importedAutoFilterCriteria.importedCustomListSort?.buttonSorted !== true ||
    importedAutoFilterCriteria.importedCustomListSort?.buttonSortDirection !== "asc" ||
    importedAutoFilterCriteria.importedCustomListSort?.sortState?.keys?.[0]?.customList !== "month-ja" ||
    importedAutoFilterCriteria.importedCustomListSort?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "1月:M1|2月:M2|3月:M3|12月:M12" ||
    importedAutoFilterCriteria.importedTextContainsFilter?.criteria?.operator1 !== "contains" ||
    importedAutoFilterCriteria.importedTextContainsFilter?.criteria?.value1 !== "ta" ||
    importedAutoFilterCriteria.importedTextContainsFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "alpha:hidden|Alpine:hidden|Beta:visible|delta:visible" ||
    importedAutoFilterCriteria.importedTextBeginsFilter?.criteria?.operator1 !== "beginsWith" ||
    importedAutoFilterCriteria.importedTextBeginsFilter?.criteria?.value1 !== "Al" ||
    importedAutoFilterCriteria.importedTextBeginsFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "alpha:visible|Alpine:visible|Beta:hidden|delta:hidden" ||
    importedAutoFilterCriteria.importedTextNotEndsFilter?.criteria?.operator1 !== "notEndsWith" ||
    importedAutoFilterCriteria.importedTextNotEndsFilter?.criteria?.value1 !== "ta" ||
    importedAutoFilterCriteria.importedTextNotEndsFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "alpha:visible|Alpine:visible|Beta:hidden|delta:hidden" ||
    importedAutoFilterCriteria.importedLiteralWildcardFilter?.criteria?.operator1 !== "equal" ||
    importedAutoFilterCriteria.importedLiteralWildcardFilter?.criteria?.value1 !== "A~*" ||
    importedAutoFilterCriteria.importedLiteralWildcardFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "A*:visible|A?:hidden|Alpha:hidden|Star:hidden" ||
    importedAutoFilterCriteria.importedMultiDateGroupFilter?.criteria?.kind !== "dateGroups" ||
    importedAutoFilterCriteria.importedMultiDateGroupFilter?.criteria?.dateGroups?.length !== 2 ||
    importedAutoFilterCriteria.importedMultiDateGroupFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "2026/05/20:visible|2026/06/01:visible|2025/05/10:hidden|2026/07/01:hidden" ||
    (importedAutoFilterCriteria.importedDateGroupExported?.dateGroupItemCount || 0) < 3 ||
    importedAutoFilterCriteria.importedBetweenFilter?.criteria?.logic !== "and" ||
    importedAutoFilterCriteria.importedBetweenFilter?.criteria?.operator1 !== "greaterThanOrEqual" ||
    importedAutoFilterCriteria.importedBetweenFilter?.criteria?.operator2 !== "lessThanOrEqual" ||
    importedAutoFilterCriteria.importedBetweenFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "A:hidden|B:visible|C:visible|D:hidden" ||
    importedAutoFilterCriteria.importedNotBetweenFilter?.criteria?.logic !== "or" ||
    importedAutoFilterCriteria.importedNotBetweenFilter?.criteria?.operator1 !== "lessThan" ||
    importedAutoFilterCriteria.importedNotBetweenFilter?.criteria?.operator2 !== "greaterThan" ||
    importedAutoFilterCriteria.importedNotBetweenFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "A:visible|B:hidden|C:hidden|D:visible" ||
    importedAutoFilterCriteria.importedMonthPeriodFilter?.criteria?.datePeriod !== "M5" ||
    importedAutoFilterCriteria.importedMonthPeriodFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "2026/05/20:visible|2025/05/10:visible|2026/06/01:hidden|2026/04/30:hidden" ||
    importedAutoFilterCriteria.importedQuarterPeriodFilter?.criteria?.datePeriod !== "Q2" ||
    importedAutoFilterCriteria.importedQuarterPeriodFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "2026/04/01:visible|2025/05/10:visible|2026/07/01:hidden|2026/03/31:hidden" ||
    importedAutoFilterCriteria.importedAboveStdDevFilter?.criteria?.stdDevOperator !== "aboveStdDev" ||
    importedAutoFilterCriteria.importedAboveStdDevFilter?.criteria?.stdDevCount !== 1 ||
    importedAutoFilterCriteria.importedAboveStdDevFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "A:hidden|B:hidden|C:hidden|D:visible" ||
    importedAutoFilterCriteria.importedBelowStdDevFilter?.criteria?.stdDevOperator !== "belowStdDev" ||
    importedAutoFilterCriteria.importedBelowStdDevFilter?.criteria?.stdDevCount !== 1 ||
    importedAutoFilterCriteria.importedBelowStdDevFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "A:visible|B:hidden|C:hidden|D:hidden" ||
    importedAutoFilterCriteria.importedHiddenButtonFilter?.hasNameButton !== true ||
    importedAutoFilterCriteria.importedHiddenButtonFilter?.hasScoreButton !== false ||
    importedAutoFilterCriteria.importedHiddenButtonFilter?.hiddenButtons?.["2"] !== true ||
    importedAutoFilterCriteria.importedHiddenButtonFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "Alpha:visible|Beta:hidden|Alpha:visible|Gamma:hidden" ||
    importedAutoFilterCriteria.importedDateHourGroupFilter?.criteria?.dateGroup?.grouping !== "hour" ||
    importedAutoFilterCriteria.importedDateHourGroupFilter?.criteria?.dateGroup?.hour !== 13 ||
    importedAutoFilterCriteria.importedDateHourGroupFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "2026/05/20 13:10:visible|2026/05/20 14:10:hidden|2026/05/21 13:10:hidden|2026/05/20 13:55:visible" ||
    importedAutoFilterCriteria.importedCustomNonBlankFilter?.criteria?.operator1 !== "nonBlank" ||
    importedAutoFilterCriteria.importedCustomNonBlankFilter?.rows?.map((row) => `${row.name || "(空白)"}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "Filled:visible|(空白):hidden|Zero:visible|(空白):hidden" ||
    !importedAutoFilterCriteria.importedDateGroupExported?.dynamicFilters?.some((item) => item.type === "M5") ||
    !importedAutoFilterCriteria.importedDateGroupExported?.dynamicFilters?.some((item) => item.type === "Q2") ||
    !importedAutoFilterCriteria.importedDateGroupExported?.dynamicFilters?.some((item) => item.type === "aboveStdDev" && item.val === "1") ||
    !importedAutoFilterCriteria.importedDateGroupExported?.dynamicFilters?.some((item) => item.type === "belowStdDev" && item.val === "1") ||
    !importedAutoFilterCriteria.importedDateGroupExported?.filterColumns?.some((item) => item.colId === "1" && item.hiddenButton === "1" && item.showButton === "0") ||
    !importedAutoFilterCriteria.importedDateGroupExported?.dateGroupItems?.some((item) => item.grouping === "hour" && item.hour === "13") ||
    !importedAutoFilterCriteria.importedCustomFilterExported?.customFilters?.some((item) => item.operator === "notEqual" && item.value === "") ||
    importedAutoFilterCriteria.tableMetadata?.tableCount !== 1 ||
    importedAutoFilterCriteria.tableMetadata?.table?.name !== "SalesTable" ||
    importedAutoFilterCriteria.tableMetadata?.table?.range?.top !== 1 ||
    importedAutoFilterCriteria.tableMetadata?.table?.range?.left !== 1 ||
    importedAutoFilterCriteria.tableMetadata?.table?.range?.bottom !== 5 ||
    importedAutoFilterCriteria.tableMetadata?.table?.range?.right !== 4 ||
    importedAutoFilterCriteria.tableMetadata?.table?.showHeaderRow !== true ||
    importedAutoFilterCriteria.tableMetadata?.table?.showTotalsRow !== true ||
    importedAutoFilterCriteria.tableMetadata?.table?.showFilterButton !== true ||
    importedAutoFilterCriteria.tableMetadata?.table?.style !== "TableStyleMedium9" ||
    importedAutoFilterCriteria.tableMetadata?.table?.styleInfo?.showFirstColumn !== true ||
    importedAutoFilterCriteria.tableMetadata?.table?.styleInfo?.showRowStripes !== true ||
    importedAutoFilterCriteria.tableMetadata?.table?.columns?.map((column) => column.name).join("|") !== "地域|売上|担当|倍額" ||
    importedAutoFilterCriteria.tableMetadata?.table?.columns?.[1]?.totalsRowFunction !== "sum" ||
    importedAutoFilterCriteria.tableMetadata?.table?.columns?.[3]?.totalsRowFunction !== "sum" ||
    importedAutoFilterCriteria.tableMetadata?.hasFilterButton !== true ||
    importedAutoFilterCriteria.tableMetadata?.doubledRows?.join("|") !== "20|40|60" ||
    importedAutoFilterCriteria.tableMetadata?.totals?.sales !== "60" ||
    importedAutoFilterCriteria.tableMetadata?.totals?.doubled !== "120" ||
    importedAutoFilterCriteria.structuredReferences?.values?.join("|") !== "60|60|売上|60|180" ||
    importedAutoFilterCriteria.structuredReferences?.formulas?.[0] !== "=SUM(SalesTable[売上])" ||
	    importedAutoFilterCriteria.structuredReferences?.formulas?.[3] !== "=SUM(SalesTable[[#Totals],[売上]])" ||
	    importedAutoFilterCriteria.structuredReferences?.renamed?.values?.join("|") !== "60|60|金額|60|180" ||
	    importedAutoFilterCriteria.structuredReferences?.renamed?.formulas?.[0] !== "=SUM(SalesTable[[金額]])" ||
	    importedAutoFilterCriteria.structuredReferences?.renamed?.formulas?.[1] !== "=SUM(SalesTable[[#Data],[金額]])" ||
	    importedAutoFilterCriteria.structuredReferences?.renamed?.formulas?.[3] !== "=SUM(SalesTable[[#Totals],[金額]])" ||
	    importedAutoFilterCriteria.structuredReferences?.renamed?.formulas?.[4] !== "=SUM(SalesTable[[金額]:[倍額]])" ||
	    importedAutoFilterCriteria.structuredReferences?.renamed?.tableColumns?.join("|") !== "地域|金額|担当|倍額" ||
	    importedAutoFilterCriteria.structuredReferences?.renamed?.tableFormula !== "=[@[金額]]*2" ||
	    importedAutoFilterCriteria.tableExpansion?.table?.name !== "ExpandTable" ||
	    importedAutoFilterCriteria.tableExpansion?.table?.range?.bottom !== 6 ||
	    importedAutoFilterCriteria.tableExpansion?.table?.range?.right !== 4 ||
	    importedAutoFilterCriteria.tableExpansion?.autoFilterRange?.bottom !== 6 ||
	    importedAutoFilterCriteria.tableExpansion?.autoFilterRange?.right !== 4 ||
	    importedAutoFilterCriteria.tableExpansion?.columns?.join("|") !== "Name|Amount|Twice|Memo" ||
	    importedAutoFilterCriteria.tableExpansion?.c4 !== "9" ||
	    importedAutoFilterCriteria.tableExpansion?.d1 !== "Memo" ||
	    importedAutoFilterCriteria.tableExpansion?.c4Formula !== "=[@Amount]*3" ||
	    importedAutoFilterCriteria.tableExpansion?.c5 !== "12" ||
	    importedAutoFilterCriteria.tableExpansion?.c5Formula !== "=[@Amount]*3" ||
	    importedAutoFilterCriteria.tableExpansion?.c6 !== "15" ||
	    importedAutoFilterCriteria.tableExpansion?.c6Formula !== "=[@Amount]*3" ||
	    importedAutoFilterCriteria.tableExpansion?.calculatedFormula !== "[@Amount]*3" ||
    importedAutoFilterCriteria.multiTableFilter?.before?.buttons?.map((button) => button.col).join("|") !== "1|2|4|5" ||
	    importedAutoFilterCriteria.multiTableFilter?.before?.buttons?.filter((button) => button.filtered).map((button) => button.col).join("|") !== "1|4" ||
	    importedAutoFilterCriteria.multiTableFilter?.before?.rows?.map((row) => `${row.left}:${row.right}:${row.hidden}`).join("|") !== "Alpha:Red:false|Beta:Blue:true|Gamma:Green:true" ||
	    importedAutoFilterCriteria.multiTableFilter?.before?.tableRanges?.length !== 2 ||
	    importedAutoFilterCriteria.multiTableFilter?.before?.tableRanges?.some((table) => table.showFilterButton !== true || !table.autoFilterRange) ||
	    importedAutoFilterCriteria.multiTableFilter?.afterRightMenu?.sheetRange?.left !== 4 ||
	    importedAutoFilterCriteria.multiTableFilter?.afterRightMenu?.sheetRange?.right !== 5 ||
	    !importedAutoFilterCriteria.multiTableFilter?.afterRightMenu?.menuValues?.join("|").includes("Red") ||
    importedAutoFilterCriteria.multiTableFilter?.afterRightBlue?.rows?.some((row) => row.hidden !== true) ||
    importedAutoFilterCriteria.multiTableFilter?.afterRightBlue?.tableCriteria?.find((table) => table.name === "LeftFilterTable")?.criteria?.["1"]?.join("|") !== "Alpha" ||
    importedAutoFilterCriteria.multiTableFilter?.afterRightBlue?.tableCriteria?.find((table) => table.name === "RightFilterTable")?.criteria?.["4"]?.join("|") !== "Blue" ||
    importedAutoFilterCriteria.multiTableFilter?.afterRightBlue?.buttons?.filter((button) => button.filtered).map((button) => button.col).join("|") !== "1|4" ||
    importedAutoFilterCriteria.multiTableFilter?.afterRightRibbonClear?.rows?.map((row) => row.hidden).join("|") !== "false|true|true" ||
    importedAutoFilterCriteria.multiTableFilter?.afterRightRibbonClear?.tableCriteria?.find((table) => table.name === "LeftFilterTable")?.criteria?.["1"]?.join("|") !== "Alpha" ||
    Object.keys(importedAutoFilterCriteria.multiTableFilter?.afterRightRibbonClear?.tableCriteria?.find((table) => table.name === "RightFilterTable")?.criteria || {}).length !== 0 ||
    importedAutoFilterCriteria.multiTableFilter?.afterRightRibbonClear?.buttons?.filter((button) => button.filtered).map((button) => button.col).join("|") !== "1" ||
    !importedAutoFilterCriteria.multiTableFilter?.afterRightRibbonClear?.status?.includes("クリア") ||
    importedAutoFilterCriteria.multiTableFilter?.afterRightClear?.rows?.map((row) => row.hidden).join("|") !== "false|true|true" ||
    importedAutoFilterCriteria.multiTableFilter?.afterRightClear?.tableCriteria?.find((table) => table.name === "LeftFilterTable")?.criteria?.["1"]?.join("|") !== "Alpha" ||
    Object.keys(importedAutoFilterCriteria.multiTableFilter?.afterRightClear?.tableCriteria?.find((table) => table.name === "RightFilterTable")?.criteria || {}).length !== 0 ||
    importedAutoFilterCriteria.multiTableFilter?.afterRightClear?.buttons?.filter((button) => button.filtered).map((button) => button.col).join("|") !== "1" ||
    importedAutoFilterCriteria.multiTableFilter?.afterRibbonReapply?.rows?.some((row) => row.hidden !== true) ||
    importedAutoFilterCriteria.multiTableFilter?.afterRibbonReapply?.rows?.find((row) => row.row === 2)?.right !== "Blue" ||
    importedAutoFilterCriteria.multiTableFilter?.afterRibbonReapply?.tableCriteria?.find((table) => table.name === "LeftFilterTable")?.criteria?.["1"]?.join("|") !== "Alpha" ||
    importedAutoFilterCriteria.multiTableFilter?.afterRibbonReapply?.tableCriteria?.find((table) => table.name === "RightFilterTable")?.criteria?.["4"]?.join("|") !== "Red" ||
    importedAutoFilterCriteria.multiTableFilter?.afterRibbonReapply?.buttons?.filter((button) => button.filtered).map((button) => button.col).join("|") !== "1|4" ||
    !importedAutoFilterCriteria.multiTableFilter?.afterRibbonReapply?.status?.includes("再適用") ||
    importedAutoFilterCriteria.filterStructureShift?.afterColumnInsert?.rightTable?.range?.left !== 5 ||
    importedAutoFilterCriteria.filterStructureShift?.afterColumnInsert?.rightTable?.range?.right !== 6 ||
    importedAutoFilterCriteria.filterStructureShift?.afterColumnInsert?.rightTable?.autoFilterState?.criteria?.["5"]?.join("|") !== "Red" ||
    importedAutoFilterCriteria.filterStructureShift?.afterColumnInsert?.buttons?.filter((button) => button.filtered).map((button) => button.col).join("|") !== "1|5" ||
    importedAutoFilterCriteria.filterStructureShift?.afterColumnInsert?.rows?.map((row) => `${row.shiftedRight}:${row.hidden}`).join("|") !== "Red:false|Blue:true|Green:true" ||
    importedAutoFilterCriteria.filterStructureShift?.afterFilteredColumnDelete?.rightTable?.range?.left !== 5 ||
    importedAutoFilterCriteria.filterStructureShift?.afterFilteredColumnDelete?.rightTable?.range?.right !== 5 ||
    Object.keys(importedAutoFilterCriteria.filterStructureShift?.afterFilteredColumnDelete?.rightTable?.autoFilterState?.criteria || {}).length !== 0 ||
    importedAutoFilterCriteria.filterStructureShift?.afterFilteredColumnDelete?.buttons?.filter((button) => button.filtered).map((button) => button.col).join("|") !== "1" ||
    importedAutoFilterCriteria.filterStructureShift?.afterSortColumnInsert?.sortState?.range?.right !== 3 ||
    importedAutoFilterCriteria.filterStructureShift?.afterSortColumnInsert?.sortState?.keys?.[0]?.col !== 3 ||
    importedAutoFilterCriteria.filterStructureShift?.afterSortColumnInsert?.buttons?.map((button) => button.col).join("|") !== "1|2|3" ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.table?.name !== "ExpandTable" ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.table?.range?.bottom !== 4 ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.autoFilterRange?.bottom !== 4 ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.buttons?.map((button) => `${button.row}:${button.col}`).join("|") !== "1:1|1:2|1:3" ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.values?.a3 !== "" ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.values?.b3 !== "" ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.values?.a4 !== "B" ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.values?.b4 !== "2" ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.values?.c4 !== "4" ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.formulas?.c3 !== "=[@Amount]*2" ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.formulas?.c4 !== "=[@Amount]*2" ||
    !importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.style?.insertedFill ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.style?.insertedFill === importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.style?.shiftedFill ||
    !importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.style?.insertedBorders?.includes("top") ||
    !importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.style?.insertedBorders?.includes("bottom") ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.style?.insertedHeight !== importedAutoFilterCriteria.tableRowInsert?.afterRowInsert?.style?.shiftedHeight ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowDelete?.table?.range?.bottom !== 3 ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowDelete?.autoFilterRange?.bottom !== 3 ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowDelete?.values?.a3 !== "B" ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowDelete?.values?.b3 !== "2" ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowDelete?.values?.c3 !== "4" ||
    importedAutoFilterCriteria.tableRowInsert?.afterRowDelete?.formulas?.c3 !== "=[@Amount]*2" ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnInsert?.table?.range?.right !== 4 ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnInsert?.autoFilterRange?.right !== 4 ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnInsert?.table?.columns?.map((column) => column.name).join("|") !== "Name|Column2|Amount|Twice" ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnInsert?.buttons?.map((button) => `${button.row}:${button.col}`).join("|") !== "1:1|1:2|1:3|1:4" ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnInsert?.values?.b1 !== "Column2" ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnInsert?.values?.c1 !== "Amount" ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnInsert?.values?.d1 !== "Twice" ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnInsert?.formulas?.d2 !== "=[@Amount]*2" ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnInsert?.formulas?.d3 !== "=[@Amount]*2" ||
    !importedAutoFilterCriteria.tableRowInsert?.afterColumnInsert?.style?.insertedColumnBorders?.includes("left") ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnInsert?.style?.insertedColumnWidth !== importedAutoFilterCriteria.tableRowInsert?.afterColumnInsert?.style?.shiftedColumnWidth ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnDelete?.table?.range?.right !== 3 ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnDelete?.autoFilterRange?.right !== 3 ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnDelete?.table?.columns?.map((column) => column.name).join("|") !== "Name|Amount|Twice" ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnDelete?.buttons?.map((button) => `${button.row}:${button.col}`).join("|") !== "1:1|1:2|1:3" ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnDelete?.values?.b1 !== "Amount" ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnDelete?.values?.c1 !== "Twice" ||
    importedAutoFilterCriteria.tableRowInsert?.afterColumnDelete?.formulas?.c3 !== "=[@Amount]*2" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.table?.name !== "SalesTable" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.table?.range?.bottom !== 6 ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.tableAutoFilter?.range?.bottom !== 5 ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.sheetAutoFilter?.range?.bottom !== 5 ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.buttons?.map((button) => `${button.row}:${button.col}`).join("|") !== "1:1|1:2|1:3|1:4" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.values?.a5 !== "" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.formulas?.d5 !== "=[@売上]*2" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.values?.a6 !== "合計" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.values?.b6 !== "60" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.values?.d6 !== "120" ||
    !importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.formulas?.b6?.includes("SUBTOTAL") ||
    !importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.formulas?.b6?.includes("売上") ||
    !importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.formulas?.d6?.includes("倍額") ||
    !importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.style?.insertedDataFill ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.style?.insertedDataFill === importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowInsert?.style?.totalFillRow6 ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowDelete?.table?.range?.bottom !== 5 ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowDelete?.tableAutoFilter?.range?.bottom !== 4 ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowDelete?.sheetAutoFilter?.range?.bottom !== 4 ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowDelete?.values?.a5 !== "合計" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowDelete?.values?.b5 !== "60" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowDelete?.values?.d5 !== "120" ||
    !importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsRowDelete?.formulas?.d5?.includes("SUBTOTAL") ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnInsert?.table?.range?.right !== 5 ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnInsert?.tableAutoFilter?.range?.right !== 5 ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnInsert?.sheetAutoFilter?.range?.right !== 5 ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnInsert?.table?.columns?.map((column) => column.name).join("|") !== "地域|Column2|売上|担当|倍額" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnInsert?.buttons?.map((button) => `${button.row}:${button.col}`).join("|") !== "1:1|1:2|1:3|1:4|1:5" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnInsert?.values?.b1 !== "Column2" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnInsert?.values?.c1 !== "売上" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnInsert?.values?.e1 !== "倍額" ||
    !importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnInsert?.formulas?.c5?.includes("売上") ||
    !importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnInsert?.formulas?.e5?.includes("倍額") ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnInsert?.formulas?.e4 !== "=[@売上]*2" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnInsert?.style?.insertedColumnWidth !== importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnInsert?.style?.shiftedColumnWidth ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnDelete?.table?.range?.right !== 4 ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnDelete?.tableAutoFilter?.range?.right !== 4 ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnDelete?.sheetAutoFilter?.range?.right !== 4 ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnDelete?.table?.columns?.map((column) => column.name).join("|") !== "地域|売上|担当|倍額" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnDelete?.buttons?.map((button) => `${button.row}:${button.col}`).join("|") !== "1:1|1:2|1:3|1:4" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnDelete?.values?.b1 !== "売上" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnDelete?.values?.d1 !== "倍額" ||
    importedAutoFilterCriteria.tableTotalsStructure?.afterTotalsColumnDelete?.formulas?.d5?.includes("倍額") !== true ||
    importedAutoFilterCriteria.structureReferenceShift?.initial?.formulas?.dataSum !== "=SUM('テーブルメタ'!B2:B4)" ||
    importedAutoFilterCriteria.structureReferenceShift?.initial?.formulas?.totalCell !== "='テーブルメタ'!B5" ||
    importedAutoFilterCriteria.structureReferenceShift?.initial?.formulas?.lastCalculated !== "='テーブルメタ'!D4" ||
    importedAutoFilterCriteria.structureReferenceShift?.initial?.formulas?.stringLiteral !== '=TEXT("B2:B4","@")' ||
    importedAutoFilterCriteria.structureReferenceShift?.initial?.formulas?.wholeColumn !== "=SUM('テーブルメタ'!B:B)" ||
    importedAutoFilterCriteria.structureReferenceShift?.initial?.formulas?.wholeRows !== "=SUM('テーブルメタ'!2:4)" ||
    importedAutoFilterCriteria.structureReferenceShift?.initial?.formulas?.columnRange !== "=SUM('テーブルメタ'!B:D)" ||
    importedAutoFilterCriteria.structureReferenceShift?.initial?.validationList !== "'テーブルメタ'!$B$2:$B$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.initial?.definedNames?.SalesDataAmounts !== "='テーブルメタ'!$B$2:$B$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.initial?.definedNames?.SalesTotalCell !== "='テーブルメタ'!$B$5" ||
    importedAutoFilterCriteria.structureReferenceShift?.initial?.definedNames?.SalesWholeAmountColumn !== "='テーブルメタ'!$B:$B" ||
    importedAutoFilterCriteria.structureReferenceShift?.initial?.definedNames?.SalesDataRows !== "='テーブルメタ'!$2:$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowInsert?.formulas?.dataSum !== "=SUM('テーブルメタ'!B2:B5)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowInsert?.formulas?.totalCell !== "='テーブルメタ'!B6" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowInsert?.formulas?.lastCalculated !== "='テーブルメタ'!D5" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowInsert?.formulas?.stringLiteral !== '=TEXT("B2:B4","@")' ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowInsert?.formulas?.wholeColumn !== "=SUM('テーブルメタ'!B:B)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowInsert?.formulas?.wholeRows !== "=SUM('テーブルメタ'!2:5)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowInsert?.formulas?.columnRange !== "=SUM('テーブルメタ'!B:D)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowInsert?.validationList !== "'テーブルメタ'!$B$2:$B$5" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowInsert?.definedNames?.SalesDataAmounts !== "='テーブルメタ'!$B$2:$B$5" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowInsert?.definedNames?.SalesTotalCell !== "='テーブルメタ'!$B$6" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowInsert?.definedNames?.SalesWholeAmountColumn !== "='テーブルメタ'!$B:$B" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowInsert?.definedNames?.SalesDataRows !== "='テーブルメタ'!$2:$5" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowDelete?.formulas?.dataSum !== "=SUM('テーブルメタ'!B2:B4)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowDelete?.formulas?.totalCell !== "='テーブルメタ'!B5" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowDelete?.formulas?.lastCalculated !== "='テーブルメタ'!D4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowDelete?.formulas?.wholeColumn !== "=SUM('テーブルメタ'!B:B)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowDelete?.formulas?.wholeRows !== "=SUM('テーブルメタ'!2:4)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowDelete?.formulas?.columnRange !== "=SUM('テーブルメタ'!B:D)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowDelete?.validationList !== "'テーブルメタ'!$B$2:$B$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowDelete?.definedNames?.SalesDataAmounts !== "='テーブルメタ'!$B$2:$B$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowDelete?.definedNames?.SalesTotalCell !== "='テーブルメタ'!$B$5" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowDelete?.definedNames?.SalesWholeAmountColumn !== "='テーブルメタ'!$B:$B" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRowDelete?.definedNames?.SalesDataRows !== "='テーブルメタ'!$2:$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnInsert?.formulas?.dataSum !== "=SUM('テーブルメタ'!C2:C4)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnInsert?.formulas?.totalCell !== "='テーブルメタ'!C5" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnInsert?.formulas?.lastCalculated !== "='テーブルメタ'!E4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnInsert?.formulas?.wholeColumn !== "=SUM('テーブルメタ'!C:C)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnInsert?.formulas?.wholeRows !== "=SUM('テーブルメタ'!2:4)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnInsert?.formulas?.columnRange !== "=SUM('テーブルメタ'!C:E)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnInsert?.validationList !== "'テーブルメタ'!$C$2:$C$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnInsert?.definedNames?.SalesDataAmounts !== "='テーブルメタ'!$C$2:$C$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnInsert?.definedNames?.SalesTotalCell !== "='テーブルメタ'!$C$5" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnInsert?.definedNames?.SalesWholeAmountColumn !== "='テーブルメタ'!$C:$C" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnInsert?.definedNames?.SalesDataRows !== "='テーブルメタ'!$2:$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnDelete?.formulas?.dataSum !== "=SUM('テーブルメタ'!B2:B4)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnDelete?.formulas?.totalCell !== "='テーブルメタ'!B5" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnDelete?.formulas?.lastCalculated !== "='テーブルメタ'!D4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnDelete?.formulas?.wholeColumn !== "=SUM('テーブルメタ'!B:B)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnDelete?.formulas?.wholeRows !== "=SUM('テーブルメタ'!2:4)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnDelete?.formulas?.columnRange !== "=SUM('テーブルメタ'!B:D)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnDelete?.validationList !== "'テーブルメタ'!$B$2:$B$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnDelete?.definedNames?.SalesDataAmounts !== "='テーブルメタ'!$B$2:$B$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnDelete?.definedNames?.SalesTotalCell !== "='テーブルメタ'!$B$5" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnDelete?.definedNames?.SalesWholeAmountColumn !== "='テーブルメタ'!$B:$B" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterColumnDelete?.definedNames?.SalesDataRows !== "='テーブルメタ'!$2:$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterUndoColumnDelete?.formulas?.dataSum !== "=SUM('テーブルメタ'!C2:C4)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterUndoColumnDelete?.formulas?.wholeColumn !== "=SUM('テーブルメタ'!C:C)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterUndoColumnDelete?.formulas?.columnRange !== "=SUM('テーブルメタ'!C:E)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterUndoColumnDelete?.validationList !== "'テーブルメタ'!$C$2:$C$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterUndoColumnDelete?.definedNames?.SalesDataAmounts !== "='テーブルメタ'!$C$2:$C$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterUndoColumnDelete?.definedNames?.SalesWholeAmountColumn !== "='テーブルメタ'!$C:$C" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRedoColumnDelete?.formulas?.dataSum !== "=SUM('テーブルメタ'!B2:B4)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRedoColumnDelete?.formulas?.wholeColumn !== "=SUM('テーブルメタ'!B:B)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRedoColumnDelete?.formulas?.columnRange !== "=SUM('テーブルメタ'!B:D)" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRedoColumnDelete?.validationList !== "'テーブルメタ'!$B$2:$B$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRedoColumnDelete?.definedNames?.SalesDataAmounts !== "='テーブルメタ'!$B$2:$B$4" ||
    importedAutoFilterCriteria.structureReferenceShift?.afterRedoColumnDelete?.definedNames?.SalesWholeAmountColumn !== "='テーブルメタ'!$B:$B" ||
    !importedAutoFilterCriteria.tableMetadata?.headerFill ||
    !importedAutoFilterCriteria.tableMetadata?.stripeFill ||
    !importedAutoFilterCriteria.tableMetadata?.totalFill
  ) {
    throw new Error("Imported AutoFilter criteria and Excel table metadata should be preserved, applied, and expanded like Excel.");
  }


  if (
    autoFilterBlankKeyboardMenu.keyboardMenuInitialFocus !== "昇順" ||
    autoFilterBlankKeyboardMenu.keyboardMenuAfterArrow !== "降順" ||
    autoFilterBlankKeyboardMenu.keyboardMenuClosedByEscape !== true ||
    autoFilterBlankKeyboardMenu.searchSelectStateBefore?.labels?.join("|") !== "Note" ||
    autoFilterBlankKeyboardMenu.searchSelectStateBefore?.checked !== true ||
    autoFilterBlankKeyboardMenu.searchSelectStateBefore?.indeterminate !== false ||
    autoFilterBlankKeyboardMenu.searchSelectStateAfter?.states?.find((item) => item.label === "Note")?.checked !== false ||
    autoFilterBlankKeyboardMenu.searchSelectStateAfter?.states?.filter((item) => item.hidden === true && item.checked === true).length < 2 ||
    autoFilterBlankKeyboardMenu.valueSearchRows?.find((row) => row.row === 5)?.hidden !== true ||
    autoFilterBlankKeyboardMenu.valueSearchRows?.find((row) => row.row === 2)?.hidden !== false ||
    autoFilterBlankKeyboardMenu.valueSearchRows?.find((row) => row.row === 3)?.hidden !== false ||
    autoFilterBlankKeyboardMenu.valueSearchRows?.find((row) => row.row === 4)?.hidden !== false ||
    autoFilterBlankKeyboardMenu.blankRows?.find((row) => row.row === 2)?.hidden !== true ||
    autoFilterBlankKeyboardMenu.blankRows?.find((row) => row.row === 3)?.hidden !== false ||
    autoFilterBlankKeyboardMenu.blankRows?.find((row) => row.row === 4)?.hidden !== false ||
    autoFilterBlankKeyboardMenu.blankRows?.find((row) => row.row === 5)?.hidden !== true ||
    autoFilterBlankKeyboardMenu.nonBlankRows?.find((row) => row.row === 2)?.hidden !== false ||
    autoFilterBlankKeyboardMenu.nonBlankRows?.find((row) => row.row === 3)?.hidden !== true ||
    autoFilterBlankKeyboardMenu.nonBlankRows?.find((row) => row.row === 4)?.hidden !== true ||
    autoFilterBlankKeyboardMenu.nonBlankRows?.find((row) => row.row === 5)?.hidden !== false
  ) {
    throw new Error("AutoFilter value search, keyboard navigation, and blank/nonblank filters should behave like Excel.");
  }


  if (
    autoFilterReapplyClear.initial?.find((row) => row.row === 2)?.hidden !== true ||
    autoFilterReapplyClear.initial?.find((row) => row.row === 3)?.hidden !== false ||
    autoFilterReapplyClear.initial?.find((row) => row.row === 4)?.hidden !== false ||
    autoFilterReapplyClear.initial?.find((row) => row.row === 5)?.hidden !== true ||
    autoFilterReapplyClear.ribbonAfterFilter?.filter?.active !== true ||
    autoFilterReapplyClear.ribbonAfterFilter?.clear?.disabled !== false ||
    autoFilterReapplyClear.ribbonAfterFilter?.reapply?.disabled !== false ||
    autoFilterReapplyClear.afterEditBeforeReapply?.find((row) => row.row === 3)?.score !== "1" ||
    autoFilterReapplyClear.afterEditBeforeReapply?.find((row) => row.row === 3)?.hidden !== false ||
    autoFilterReapplyClear.afterReapply?.rows?.find((row) => row.row === 3)?.hidden !== true ||
    autoFilterReapplyClear.afterReapply?.rows?.find((row) => row.row === 4)?.hidden !== false ||
    autoFilterReapplyClear.afterReapply?.filtered !== true ||
    !autoFilterReapplyClear.afterReapply?.status?.includes("再適用") ||
    autoFilterReapplyClear.afterClear?.rows?.some((row) => row.hidden) ||
    autoFilterReapplyClear.afterClear?.filteredButtons !== 0 ||
    autoFilterReapplyClear.ribbonAfterClear?.filter?.active !== true ||
    autoFilterReapplyClear.ribbonAfterClear?.clear?.disabled !== true ||
    autoFilterReapplyClear.ribbonAfterClear?.reapply?.disabled !== false ||
    !autoFilterReapplyClear.afterClear?.status?.includes("クリア") ||
    autoFilterReapplyClear.sortInitial?.map((row) => row.name).join("|") !== "Beta|delta|alpha|Alpine" ||
    autoFilterReapplyClear.ribbonAfterSort?.clear?.disabled !== false ||
    autoFilterReapplyClear.ribbonAfterSort?.reapply?.disabled !== false ||
    autoFilterReapplyClear.sortHeaderState?.sorted !== true ||
    autoFilterReapplyClear.sortHeaderState?.desc !== true ||
    autoFilterReapplyClear.sortHeaderState?.direction !== "desc" ||
    autoFilterReapplyClear.sortHeaderState?.sortState?.keys?.[0]?.col !== 2 ||
    autoFilterReapplyClear.sortMenuState?.direction !== "desc" ||
    autoFilterReapplyClear.sortMenuState?.descPressed !== "true" ||
    !autoFilterReapplyClear.sortMenuState?.current?.includes("降順") ||
    autoFilterReapplyClear.sortAfterEditBeforeReapply?.map((row) => `${row.name}:${row.score}`).join("|") !== "Beta:20|delta:8|alpha:5|Alpine:30" ||
    autoFilterReapplyClear.sortAfterReapply?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "Alpine:30|Beta:20|delta:8|alpha:5" ||
    !autoFilterReapplyClear.sortAfterReapply?.status?.includes("再適用") ||
    autoFilterReapplyClear.sortAfterUndo?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "Beta:20|delta:8|alpha:5|Alpine:30" ||
    !autoFilterReapplyClear.sortAfterUndo?.status?.includes("元に戻しました") ||
    autoFilterReapplyClear.sortClear?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "Beta:20|delta:8|alpha:5|Alpine:30" ||
    autoFilterReapplyClear.sortClear?.header?.sorted !== false ||
    autoFilterReapplyClear.sortClear?.header?.direction !== "" ||
    !autoFilterReapplyClear.sortClear?.status?.includes("並べ替え状態") ||
    autoFilterReapplyClear.sortAfterClearReapply?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "Beta:20|delta:8|alpha:5|Alpine:40" ||
    !autoFilterReapplyClear.sortAfterClearReapply?.status?.includes("フィルターを再適用")
  ) {
    throw new Error("AutoFilter reapply and clear should refresh visible rows, reapply stored sort order, and clear sort state like Excel.");
  }


  if (
    advancedFilterCopyOverwrite.overwritePrompt?.title !== "Cht WebSheet" ||
    advancedFilterCopyOverwrite.namedDefaults?.action !== "in-place" ||
    advancedFilterCopyOverwrite.namedDefaults?.listRange !== "Filter!A1:B5" ||
    advancedFilterCopyOverwrite.namedDefaults?.criteriaRange !== "Criteria!D1:D2" ||
    advancedFilterCopyOverwrite.namedDefaults?.copyToRange !== "Extract!G1" ||
    !advancedFilterCopyOverwrite.overwritePrompt?.message?.includes("既にデータ") ||
    advancedFilterCopyOverwrite.overwritePrompt?.buttons?.join("|") !== "OK|キャンセル" ||
    advancedFilterCopyOverwrite.afterCancel?.dialogVisible !== true ||
    advancedFilterCopyOverwrite.afterCancel?.h1 !== "既存" ||
    advancedFilterCopyOverwrite.values?.join("|") !== "Name|Score|Alpine|12|Beta|20" ||
    !advancedFilterCopyOverwrite.status?.includes("2件を抽出") ||
    !advancedFilterCopyOverwrite.qualifiedCopy?.activeSheet?.includes("Extract") ||
    advancedFilterCopyOverwrite.qualifiedCopy?.values?.join("|") !== "Name|Score|Alpine|12|Beta|20" ||
    !advancedFilterCopyOverwrite.qualifiedCopy?.status?.includes("2件を抽出") ||
    advancedFilterCopyOverwrite.wildcardAndOrCopy?.values?.join("|") !== "Name|Score|Alpine|12|delta|8||" ||
    !advancedFilterCopyOverwrite.wildcardAndOrCopy?.status?.includes("2件を抽出") ||
    advancedFilterCopyOverwrite.namedPrefixCopy?.values?.join("|") !== "Name|Score|alpha|5|Alpine|12||" ||
    !advancedFilterCopyOverwrite.namedPrefixCopy?.status?.includes("2件を抽出") ||
    advancedFilterCopyOverwrite.inPlace?.rows?.map((row) => `${row.row}:${row.hidden}`).join("|") !== "2:true|3:false|4:false|5:true" ||
    advancedFilterCopyOverwrite.inPlace?.criteriaRange?.top !== 1 ||
    advancedFilterCopyOverwrite.inPlace?.criteriaRange?.left !== 1 ||
    !advancedFilterCopyOverwrite.inPlace?.status?.includes("詳細設定フィルター") ||
    advancedFilterCopyOverwrite.inPlaceUndo?.rows?.some((row) => row.hidden) ||
    !advancedFilterCopyOverwrite.inPlaceUndo?.status?.includes("元に戻しました") ||
    advancedFilterCopyOverwrite.inPlaceRedo?.rows?.map((row) => `${row.row}:${row.hidden}`).join("|") !== "2:true|3:false|4:false|5:true" ||
    !advancedFilterCopyOverwrite.inPlaceRedo?.status?.includes("やり直しました")
  ) {
    throw new Error("Advanced Filter should support Excel-style copy ranges, in-place filtering, and undo/redo.");
  }


  const dateFilterRowsMatch = (rows, expectedHiddenByRow) =>
    [2, 3, 4, 5, 6].every((row) => rows?.find((item) => item.row === row)?.hidden === expectedHiddenByRow.get(row));


  const expectedDateFilterHiddenRows = (predicate) => new Map([
    [2, new Date(2026, 4, 20)],
    [3, new Date(2026, 4, 20)],
    [4, new Date(2026, 4, 23)],
    [5, new Date(2026, 5, 1)],
    [6, new Date(2025, 11, 31)],
  ].map(([row, date]) => [row, !predicate(date)]));


  const dateFilterReferenceDate = new Date();


  const dateFilterReferenceStartOfDay = new Date(
    dateFilterReferenceDate.getFullYear(),
    dateFilterReferenceDate.getMonth(),
    dateFilterReferenceDate.getDate(),
  );


  const dateFilterNextMonthReference = new Date(dateFilterReferenceDate.getFullYear(), dateFilterReferenceDate.getMonth() + 1, 1);


  const dateFilterThisMonthExpected = expectedDateFilterHiddenRows((date) =>
    date.getFullYear() === dateFilterReferenceDate.getFullYear() &&
    date.getMonth() === dateFilterReferenceDate.getMonth());


  const dateFilterNextMonthExpected = expectedDateFilterHiddenRows((date) =>
    date.getFullYear() === dateFilterNextMonthReference.getFullYear() &&
    date.getMonth() === dateFilterNextMonthReference.getMonth());


  const dateFilterThisYearExpected = expectedDateFilterHiddenRows((date) =>
    date.getFullYear() === dateFilterReferenceDate.getFullYear());


  const dateFilterYearToDateExpected = expectedDateFilterHiddenRows((date) =>
    date.getFullYear() === dateFilterReferenceDate.getFullYear() &&
    date >= new Date(dateFilterReferenceDate.getFullYear(), 0, 1) &&
    date <= dateFilterReferenceStartOfDay);


  const dateFilterAllMayExpected = expectedDateFilterHiddenRows((date) => date.getMonth() === 4);


  const dateFilterSecondQuarterExpected = expectedDateFilterHiddenRows((date) => date.getMonth() >= 3 && date.getMonth() <= 5);


  if (
    !autoFilterDateMenu.menu?.includes("日付フィルター") ||
    !autoFilterDateMenu.submenu?.includes("指定の日付より後") ||
    !autoFilterDateMenu.submenu?.includes("今月") ||
    !autoFilterDateMenu.submenu?.includes("年初来") ||
    !autoFilterDateMenu.submenu?.includes("第2四半期") ||
    !autoFilterDateMenu.submenu?.includes("5月") ||
    autoFilterDateMenu.valueLabels?.join("|") !== "2025/12/31|May 20, 2026|05/23/2026|06/01/2026" ||
    autoFilterDateMenu.valueSelectionRows?.find((row) => row.row === 2)?.hidden !== false ||
    autoFilterDateMenu.valueSelectionRows?.find((row) => row.row === 3)?.hidden !== false ||
    autoFilterDateMenu.valueSelectionRows?.find((row) => row.row === 4)?.hidden !== true ||
    autoFilterDateMenu.valueSelectionRows?.find((row) => row.row === 5)?.hidden !== true ||
    autoFilterDateMenu.valueSelectionRows?.find((row) => row.row === 6)?.hidden !== true ||
    autoFilterDateMenu.rows?.find((row) => row.row === 2)?.hidden !== true ||
    autoFilterDateMenu.rows?.find((row) => row.row === 3)?.hidden !== true ||
    autoFilterDateMenu.rows?.find((row) => row.row === 4)?.hidden !== false ||
    autoFilterDateMenu.rows?.find((row) => row.row === 5)?.hidden !== false ||
    autoFilterDateMenu.rows?.find((row) => row.row === 6)?.hidden !== true ||
    !dateFilterRowsMatch(autoFilterDateMenu.thisMonthRows, dateFilterThisMonthExpected) ||
    !dateFilterRowsMatch(autoFilterDateMenu.nextMonthRows, dateFilterNextMonthExpected) ||
    !dateFilterRowsMatch(autoFilterDateMenu.thisYearRows, dateFilterThisYearExpected) ||
    !dateFilterRowsMatch(autoFilterDateMenu.yearToDateRows, dateFilterYearToDateExpected) ||
    !dateFilterRowsMatch(autoFilterDateMenu.allMayRows, dateFilterAllMayExpected) ||
    !dateFilterRowsMatch(autoFilterDateMenu.secondQuarterRows, dateFilterSecondQuarterExpected) ||
    !autoFilterDateMenu.periodExported?.dynamicFilters?.some((item) => item.type === "Q2")
  ) {
    throw new Error("AutoFilter should recognize date-like columns and apply Excel-style date presets and period filters.");
  }


  if (
    JSON.stringify(advancedFilterUniqueDates.values) !== JSON.stringify([
      "Name",
      "Score",
      "Date",
      "Alpha",
      "10",
      "May 20, 2026",
      "Beta",
      "20",
      "05/23/2026",
      "",
      "",
      "",
    ]) ||
    !advancedFilterUniqueDates.status?.includes("2件を抽出")
  ) {
    throw new Error("Advanced Filter should compare date criteria and ignore case when returning unique records.");
  }


  if (
    advancedFilterUniqueSubsetCopy.applied?.values?.join("|") !== "Name|Alpha|Beta|" ||
    !advancedFilterUniqueSubsetCopy.applied?.status?.includes("2件を抽出") ||
    advancedFilterUniqueSubsetCopy.undo?.values?.join("|") !== "Name|||" ||
    !advancedFilterUniqueSubsetCopy.undo?.status?.includes("元に戻しました") ||
    advancedFilterUniqueSubsetCopy.redo?.values?.join("|") !== "Name|Alpha|Beta|" ||
    !advancedFilterUniqueSubsetCopy.redo?.status?.includes("やり直しました")
  ) {
    throw new Error("Advanced Filter unique copy should deduplicate by the extracted columns and support undo/redo like Excel.");
  }


  if (
    !delimitedTextImport.initialDialog?.title?.endsWith(".csv") ||
    delimitedTextImport.initialDialog?.delimiter !== "auto" ||
    delimitedTextImport.initialDialog?.origin !== "65001" ||
    !delimitedTextImport.initialDialog?.preview?.includes("NameScoreAlpha10Beta20") ||
    !delimitedTextImport.initialDialog?.transformVisible ||
    delimitedTextImport.overwritePrompt?.title !== "Cht WebSheet" ||
    !delimitedTextImport.overwritePrompt?.message?.includes("既にデータ") ||
    delimitedTextImport.overwritePrompt?.buttons?.join("|") !== "OK|キャンセル" ||
    delimitedTextImport.afterCancel?.a1 !== "既存" ||
    delimitedTextImport.afterCancel?.b2 !== "保持" ||
    delimitedTextImport.afterCancel?.dialogVisible !== true ||
    !delimitedTextImport.commaPreview?.includes("NameScoreAlpha10Beta20") ||
    delimitedTextImport.values?.join("|") !== "Name|Score|Alpha|10|Beta|20" ||
    !delimitedTextImport.status?.includes("A1:B3") ||
    delimitedTextImport.typeDetectionNone?.values?.join("|") !== "Code|Qty|0012|09" ||
    delimitedTextImport.typeDetectionNone?.raw?.d2 !== "0012" ||
    delimitedTextImport.typeDetectionNone?.raw?.e2 !== "09" ||
    !delimitedTextImport.typeDetectionNone?.status?.includes("D1:E2") ||
    delimitedTextImport.utf16Initial?.origin !== "1200" ||
    !delimitedTextImport.utf16Initial?.preview?.includes("名前点数東京10大阪20") ||
    !delimitedTextImport.utf16Initial?.rowCount?.includes("3") ||
    delimitedTextImport.utf8Preview === delimitedTextImport.utf16Decoded?.preview ||
    delimitedTextImport.utf16Decoded?.origin !== "1200" ||
    !delimitedTextImport.utf16Decoded?.delimiterLabel?.includes("コンマ") ||
    delimitedTextImport.utf16Import?.values?.join("|") !== "名前|点数|東京|10|大阪|20" ||
    !delimitedTextImport.utf16Import?.status?.includes("G1:H3") ||
    delimitedTextImport.shiftJisInitial?.origin !== "65001" ||
    delimitedTextImport.shiftJisInitial?.preview === delimitedTextImport.shiftJisDecoded?.preview ||
    delimitedTextImport.shiftJisDecoded?.origin !== "932" ||
    !delimitedTextImport.shiftJisDecoded?.preview?.includes("名前点数東京10大阪20") ||
    !delimitedTextImport.shiftJisDecoded?.delimiterLabel?.includes("コンマ") ||
    delimitedTextImport.shiftJisImport?.values?.join("|") !== "名前|点数|東京|10|大阪|20" ||
    !delimitedTextImport.shiftJisImport?.status?.includes("J1:K3") ||
    !delimitedTextImport.transformEditorInitial?.title?.includes("Power Query エディター") ||
    !delimitedTextImport.transformEditorInitial?.preview?.includes("NameScoreCityAlpha10TokyoBeta20OsakaGamma30Tokyo") ||
    !delimitedTextImport.transformEditorInitial?.columns?.includes("Column1") ||
    !delimitedTextImport.transformEditorAfterSteps?.steps?.some((step) => step.includes("昇格されたヘッダー")) ||
    !delimitedTextImport.transformEditorAfterSteps?.steps?.some((step) => step.includes("削除された列: Score")) ||
    !delimitedTextImport.transformEditorAfterSteps?.steps?.some((step) => step.includes("フィルターされた行: City = Tokyo")) ||
    delimitedTextImport.transformEditorAfterSteps?.columns?.includes("Score") ||
    delimitedTextImport.transformEditorAfterSteps?.preview?.includes("Beta") ||
    delimitedTextImport.transformImport?.values?.join("|") !== "Name|City|Alpha|Tokyo|Gamma|Tokyo" ||
    delimitedTextImport.transformImport?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|removeColumns|filterRows" ||
    delimitedTextImport.transformImport?.query?.range?.left !== 13 ||
    !delimitedTextImport.transformImport?.status?.includes("M1:N3") ||
    delimitedTextImport.transformRefresh?.values?.join("|") !== "Name|City|Alpha|Tokyo|Gamma|Tokyo" ||
    (delimitedTextImport.transformRefresh?.query?.refreshCount || 0) < 1 ||
    !delimitedTextImport.transformRefresh?.status?.includes("更新") ||
    delimitedTextImport.jsonFileImport?.values?.join("|") !== "Name|Score|Detail.City|JsonAlpha|11|Tokyo|JsonBeta|22|Osaka" ||
    delimitedTextImport.jsonFileImport?.query?.sourceItemKind !== "json" ||
    delimitedTextImport.jsonFileImport?.query?.sourceItemName !== "Document" ||
    Number(delimitedTextImport.jsonFileImport?.rawScore) !== 22 ||
    !delimitedTextImport.jsonFileImport?.status?.includes("P1:R3") ||
    delimitedTextImport.xmlFileImport?.values?.join("|") !== "@code|Name|Qty|City|X1|XmlAlpha|4|Tokyo|X2|XmlBeta|8|Kyoto" ||
    delimitedTextImport.xmlFileImport?.query?.sourceItemKind !== "xml" ||
    delimitedTextImport.xmlFileImport?.query?.sourceItemName !== "catalog/entry" ||
    Number(delimitedTextImport.xmlFileImport?.rawQty) !== 8 ||
    !delimitedTextImport.xmlFileImport?.status?.includes("T1:W3") ||
    delimitedTextImport.workbookNavigatorInitial?.sources?.join("|") !== "Summary|Rows" ||
    delimitedTextImport.workbookNavigatorInitial?.active !== "Summary" ||
    !delimitedTextImport.workbookNavigatorInitial?.preview?.includes("MetricValueCount2") ||
    delimitedTextImport.workbookNavigatorSelected?.active !== "Rows" ||
    !delimitedTextImport.workbookNavigatorSelected?.summary?.includes("3行x3列") ||
    !delimitedTextImport.workbookNavigatorSelected?.preview?.includes("NameScoreCityBookAlpha31TokyoBookBeta42Osaka") ||
    !delimitedTextImport.workbookTransformAfterSteps?.steps?.some((step) => step.includes("フィルターされた行: City = Tokyo")) ||
    delimitedTextImport.workbookTransformAfterSteps?.columns?.join("|") !== "Name|Score|City" ||
    delimitedTextImport.workbookTransformAfterSteps?.preview !== "NameScoreCityBookAlpha31Tokyo" ||
    delimitedTextImport.workbookFileImport?.values?.join("|") !== "Name|Score|City|BookAlpha|31|Tokyo" ||
    delimitedTextImport.workbookFileImport?.query?.sourceType !== "workbook" ||
    delimitedTextImport.workbookFileImport?.query?.sourceName !== "import-workbook-source.xlsx - Rows" ||
    delimitedTextImport.workbookFileImport?.query?.sourceItemKind !== "workbook" ||
    delimitedTextImport.workbookFileImport?.query?.sourceItemName !== "Rows" ||
    delimitedTextImport.workbookFileImport?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|filterRows" ||
    Number(delimitedTextImport.workbookFileImport?.rawScore) !== 31 ||
    !delimitedTextImport.workbookFileImport?.status?.includes("A5:C6") ||
    !delimitedTextImport.folderMenuLabels?.some((label) => label.includes("フォルダーから")) ||
    !delimitedTextImport.folderNavigatorInitial?.sources?.includes("フォルダー ファイル一覧") ||
    !delimitedTextImport.folderNavigatorInitial?.sources?.includes("folder-source.csv") ||
    !delimitedTextImport.folderNavigatorInitial?.sources?.includes("import-workbook-source.xlsx - Rows") ||
    delimitedTextImport.folderNavigatorInitial?.active !== "フォルダー ファイル一覧" ||
    !delimitedTextImport.folderNavigatorInitial?.preview?.includes("ContentNameExtensionDateaccessedDatemodifiedDatecreatedAttributesFolderPath") ||
    !delimitedTextImport.folderNavigatorInitial?.preview?.includes("Binaryfolder-source.csv.csv") ||
    delimitedTextImport.folderNavigatorSelected?.active !== "folder-source.csv" ||
    !delimitedTextImport.folderNavigatorSelected?.summary?.includes("3行x3列") ||
    !delimitedTextImport.folderNavigatorSelected?.preview?.includes("NameScoreCityFolderAlpha51TokyoFolderBeta62Osaka") ||
    !delimitedTextImport.folderTransformAfterSteps?.steps?.some((step) => step.includes("昇格されたヘッダー")) ||
    !delimitedTextImport.folderTransformAfterSteps?.steps?.some((step) => step.includes("フィルターされた行: City = Tokyo")) ||
    delimitedTextImport.folderTransformAfterSteps?.columns?.join("|") !== "Name|Score|City" ||
    delimitedTextImport.folderTransformAfterSteps?.preview !== "NameScoreCityFolderAlpha51Tokyo" ||
    delimitedTextImport.folderFileImport?.values?.join("|") !== "Name|Score|City|FolderAlpha|51|Tokyo" ||
    delimitedTextImport.folderFileImport?.query?.sourceType !== "folder" ||
    delimitedTextImport.folderFileImport?.query?.sourceItemKind !== "folder" ||
    delimitedTextImport.folderFileImport?.query?.sourceItemName !== "folder-source.csv" ||
    delimitedTextImport.folderFileImport?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|filterRows" ||
    Number(delimitedTextImport.folderFileImport?.rawScore) !== 51 ||
    !delimitedTextImport.folderFileImport?.status?.includes("D8:F9") ||
    delimitedTextImport.folderRefresh?.values?.join("|") !== "Name|Score|City|FolderAlpha|51|Tokyo" ||
    (delimitedTextImport.folderRefresh?.query?.refreshCount || 0) < 1 ||
    !delimitedTextImport.folderRefresh?.status?.includes("更新") ||
    !delimitedTextImport.folderConnectionDetails?.details?.includes("フォルダー接続") ||
    !delimitedTextImport.folderConnectionDetails?.details?.includes("FOLDER;") ||
    !delimitedTextImport.folderConnectionDetails?.details?.includes("folder-source.csv") ||
    !delimitedTextImport.odataMenuLabels?.some((label) => label.includes("OData フィードから")) ||
    !delimitedTextImport.odataMenuLabels?.some((label) => label.includes("その他のデータ ソース")) ||
    !delimitedTextImport.otherDataSourcesMenuLabels?.some((label) => label.includes("OData フィードから")) ||
    !delimitedTextImport.otherDataSourcesMenuLabels?.some((label) => label.includes("SQL Server データベースから")) ||
    !delimitedTextImport.otherDataSourcesMenuLabels?.some((label) => label.includes("Microsoft Access データベースから")) ||
    delimitedTextImport.odataDialogInitial?.title !== "OData フィードから" ||
    !delimitedTextImport.odataDialogInitial?.placeholder?.includes("Northwind") ||
    !delimitedTextImport.odataDialogInitial?.destination?.includes("H8") ||
    !delimitedTextImport.odataDialogInitial?.actions?.includes("データの変換") ||
    delimitedTextImport.odataNavigatorInitial?.sources?.join("|") !== "Products|Categories" ||
    delimitedTextImport.odataNavigatorInitial?.active !== "Products" ||
    !delimitedTextImport.odataNavigatorInitial?.summary?.includes("3行x3列") ||
    !delimitedTextImport.odataNavigatorInitial?.preview?.includes("ProductScoreCityODataAlpha71TokyoODataBeta82Osaka") ||
    !delimitedTextImport.odataTransformAfterSteps?.steps?.some((step) => step.includes("昇格されたヘッダー")) ||
    !delimitedTextImport.odataTransformAfterSteps?.steps?.some((step) => step.includes("フィルターされた行: City = Tokyo")) ||
    delimitedTextImport.odataTransformAfterSteps?.columns?.join("|") !== "Product|Score|City" ||
    delimitedTextImport.odataTransformAfterSteps?.preview !== "ProductScoreCityODataAlpha71Tokyo" ||
    delimitedTextImport.odataImport?.values?.join("|") !== "Product|Score|City|ODataAlpha|71|Tokyo" ||
    delimitedTextImport.odataImport?.query?.sourceType !== "odata" ||
    delimitedTextImport.odataImport?.query?.sourceItemKind !== "odata" ||
    delimitedTextImport.odataImport?.query?.sourceItemName !== "Products" ||
    delimitedTextImport.odataImport?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|filterRows" ||
    Number(delimitedTextImport.odataImport?.rawScore) !== 71 ||
    !delimitedTextImport.odataImport?.status?.includes("H8:J9") ||
    delimitedTextImport.odataRefresh?.values?.join("|") !== "Product|Score|City|ODataAlpha|71|Tokyo" ||
    (delimitedTextImport.odataRefresh?.query?.refreshCount || 0) < 1 ||
    !delimitedTextImport.odataRefresh?.status?.includes("更新") ||
    !delimitedTextImport.odataConnectionDetails?.details?.includes("OData フィード接続") ||
    !delimitedTextImport.odataConnectionDetails?.details?.includes("ODATA;") ||
    !delimitedTextImport.odataConnectionDetails?.details?.includes("Products")
  ) {
    throw new Error("Text/CSV, folder, OData, and structured JSON/XML imports should preview, transform, warn before overwriting, honor encodings/type detection, persist query metadata, and keep transform steps during refresh.");
  }


  if (
    !powerQueryTransformEditor.tableInitial?.title?.includes("Power Query エディター") ||
    !powerQueryTransformEditor.tableInitial?.preview?.includes("NameScoreCityAlpha10TokyoBeta20OsakaAlpha10TokyoGammaTokyo") ||
    !powerQueryTransformEditor.tableInitial?.source?.includes("PQ!A1:C6") ||
    !powerQueryTransformEditor.tableAfterSteps?.steps?.some((step) => step.includes("名前が変更された列: City → Area")) ||
    !powerQueryTransformEditor.tableAfterSteps?.steps?.some((step) => step.includes("変更された型: Score = 10 進数")) ||
    !powerQueryTransformEditor.tableAfterSteps?.steps?.some((step) => step.includes("並べ替えられた行: Score 降順")) ||
    !powerQueryTransformEditor.tableAfterSteps?.steps?.some((step) => step.includes("削除された重複")) ||
    !powerQueryTransformEditor.tableAfterSteps?.steps?.some((step) => step.includes("削除された空白行")) ||
    !powerQueryTransformEditor.tableAfterSteps?.steps?.some((step) => step.includes("保持された上位行: 2")) ||
    powerQueryTransformEditor.tableAfterSteps?.columns?.join("|") !== "Name|Score|Area" ||
    powerQueryTransformEditor.tableAfterSteps?.preview !== "NameScoreAreaBeta20OsakaAlpha10Tokyo" ||
    powerQueryTransformEditor.tableLoadTo?.values?.join("|") !== "Name|Score|Area|Beta|20|Osaka|Alpha|10|Tokyo" ||
    powerQueryTransformEditor.tableLoadTo?.query?.sourceType !== "tableRange" ||
    powerQueryTransformEditor.tableLoadTo?.query?.sourceRange?.right !== 3 ||
    powerQueryTransformEditor.tableLoadTo?.query?.range?.left !== 5 ||
    powerQueryTransformEditor.tableLoadTo?.query?.transformSteps?.length < 7 ||
    Number(powerQueryTransformEditor.tableLoadTo?.rawScore) !== 20 ||
    !powerQueryTransformEditor.tableLoadTo?.status?.includes("E1:G3") ||
    !powerQueryTransformEditor.keepDuplicatesAfterSteps?.steps?.some((step) => step.includes("保持された重複")) ||
    powerQueryTransformEditor.keepDuplicatesAfterSteps?.columns?.join("|") !== "Name|Score|City" ||
    powerQueryTransformEditor.keepDuplicatesAfterSteps?.preview !== "NameScoreCityAlpha10TokyoAlpha10Tokyo" ||
    powerQueryTransformEditor.keepDuplicatesLoad?.values?.join("|") !== "Name|Score|City|Alpha|10|Tokyo|Alpha|10|Tokyo" ||
    powerQueryTransformEditor.keepDuplicatesLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|keepDuplicates" ||
    !powerQueryTransformEditor.keepDuplicatesLoad?.status?.includes("BY1:CA3") ||
    !powerQueryTransformEditor.keepBlankRowsAfterSteps?.steps?.some((step) => step.includes("保持された空白行")) ||
    powerQueryTransformEditor.keepBlankRowsAfterSteps?.columns?.join("|") !== "Name|Score|City" ||
    powerQueryTransformEditor.keepBlankRowsAfterSteps?.preview !== "NameScoreCity" ||
    powerQueryTransformEditor.keepBlankRowsLoad?.values?.join("|") !== "Name|Score|City|||" ||
    powerQueryTransformEditor.keepBlankRowsLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|keepBlankRows" ||
    !powerQueryTransformEditor.keepBlankRowsLoad?.status?.includes("CC1:CE2") ||
    !powerQueryTransformEditor.webAfterSteps?.steps?.some((step) => step.includes("削除された列: Score")) ||
    !powerQueryTransformEditor.webAfterSteps?.steps?.some((step) => step.includes("フィルターされた行: City = Tokyo")) ||
    powerQueryTransformEditor.webAfterSteps?.columns?.join("|") !== "Name|City" ||
    powerQueryTransformEditor.webAfterSteps?.preview !== "NameCityWebAlphaTokyoWebGammaTokyo" ||
    powerQueryTransformEditor.webLoad?.values?.join("|") !== "Name|City|WebAlpha|Tokyo|WebGamma|Tokyo" ||
    powerQueryTransformEditor.webLoad?.query?.sourceType !== "web" ||
    powerQueryTransformEditor.webLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|removeColumns|filterRows" ||
    !powerQueryTransformEditor.webLoad?.status?.includes("J1:K3") ||
    powerQueryTransformEditor.webNavigatorInitial?.sources?.join("|") !== "Summary|Work Items" ||
    powerQueryTransformEditor.webNavigatorInitial?.active !== "Summary" ||
    !powerQueryTransformEditor.webNavigatorInitial?.preview?.includes("MetricTotalCount3") ||
    powerQueryTransformEditor.webNavigatorSelected?.active !== "Work Items" ||
    !powerQueryTransformEditor.webNavigatorSelected?.summary?.includes("4行x3列") ||
    !powerQueryTransformEditor.webNavigatorSelected?.preview?.includes("TaskOwnerStatusBuildMikaOpenShipKenClosedReviewRenOpen") ||
    !powerQueryTransformEditor.webNavigatorAfterSteps?.steps?.some((step) => step.includes("フィルターされた行: Status = Open")) ||
    powerQueryTransformEditor.webNavigatorAfterSteps?.columns?.join("|") !== "Task|Owner|Status" ||
    powerQueryTransformEditor.webNavigatorAfterSteps?.preview !== "TaskOwnerStatusBuildMikaOpenReviewRenOpen" ||
    powerQueryTransformEditor.webNavigatorLoad?.values?.join("|") !== "Task|Owner|Status|Build|Mika|Open|Review|Ren|Open" ||
    powerQueryTransformEditor.webNavigatorLoad?.query?.sourceType !== "web" ||
    powerQueryTransformEditor.webNavigatorLoad?.query?.sourceItemIndex !== 1 ||
    powerQueryTransformEditor.webNavigatorLoad?.query?.sourceItemName !== "Work Items" ||
    powerQueryTransformEditor.webNavigatorLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|filterRows" ||
    powerQueryTransformEditor.webNavigatorLoad?.rawTask !== "Build" ||
    powerQueryTransformEditor.webNavigatorLoad?.rawStatus !== "Open" ||
    !powerQueryTransformEditor.webNavigatorLoad?.status?.includes("BM1:BO3") ||
    powerQueryTransformEditor.jsonNavigatorInitial?.sources?.join("|") !== "meta|items" ||
    powerQueryTransformEditor.jsonNavigatorInitial?.active !== "meta" ||
    !powerQueryTransformEditor.jsonNavigatorInitial?.preview?.includes("MetricValueTotal3") ||
    !powerQueryTransformEditor.jsonNavigatorAfterSteps?.steps?.some((step) => step.includes("フィルターされた行: Status = Open")) ||
    powerQueryTransformEditor.jsonNavigatorAfterSteps?.columns?.join("|") !== "Task|Owner|Status|Detail.Score" ||
    powerQueryTransformEditor.jsonNavigatorAfterSteps?.preview !== "TaskOwnerStatusDetail.ScoreBuildMikaOpen10ReviewRenOpen30" ||
    powerQueryTransformEditor.jsonNavigatorLoad?.values?.join("|") !== "Task|Owner|Status|Detail.Score|Build|Mika|Open|10|Review|Ren|Open|30" ||
    powerQueryTransformEditor.jsonNavigatorLoad?.query?.sourceItemIndex !== 1 ||
    powerQueryTransformEditor.jsonNavigatorLoad?.query?.sourceItemName !== "items" ||
    powerQueryTransformEditor.jsonNavigatorLoad?.query?.sourceItemKind !== "json" ||
    powerQueryTransformEditor.jsonNavigatorLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|filterRows" ||
    Number(powerQueryTransformEditor.jsonNavigatorLoad?.rawScore) !== 30 ||
    !powerQueryTransformEditor.jsonNavigatorLoad?.status?.includes("BP1:BS3") ||
    !powerQueryTransformEditor.xmlInitial?.preview?.includes("@idNameQtyCityA1Alpha5TokyoB2Beta9Osaka") ||
    powerQueryTransformEditor.xmlLoad?.values?.join("|") !== "@id|Name|Qty|City|A1|Alpha|5|Tokyo|B2|Beta|9|Osaka" ||
    powerQueryTransformEditor.xmlLoad?.query?.sourceItemKind !== "xml" ||
    powerQueryTransformEditor.xmlLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|changeType" ||
    Number(powerQueryTransformEditor.xmlLoad?.rawQty) !== 9 ||
    !powerQueryTransformEditor.xmlLoad?.status?.includes("BU1:BX3") ||
    !powerQueryTransformEditor.textAfterSteps?.steps?.some((step) => step.includes("トリムされたテキスト: FullName")) ||
    !powerQueryTransformEditor.textAfterSteps?.steps?.some((step) => step.includes("クリーンされたテキスト: FullName")) ||
    !powerQueryTransformEditor.textAfterSteps?.steps?.some((step) => step.includes("置換された値: FullName")) ||
    !powerQueryTransformEditor.textAfterSteps?.steps?.some((step) => step.includes("区切り記号による列の分割: FullName")) ||
    !powerQueryTransformEditor.textAfterSteps?.steps?.some((step) => step.includes("下方向へコピー: FullName")) ||
    !powerQueryTransformEditor.textAfterSteps?.steps?.some((step) => step.includes("追加されたカスタム列: Label")) ||
    !powerQueryTransformEditor.textAfterSteps?.steps?.some((step) => step.includes("グループ化された行: Category / 合計 Amount")) ||
    powerQueryTransformEditor.textAfterSteps?.columns?.join("|") !== "Category|Total" ||
    powerQueryTransformEditor.textAfterSteps?.preview !== "CategoryTotalA3B7" ||
    powerQueryTransformEditor.textLoad?.values?.join("|") !== "Category|Total|A|3|B|7" ||
    powerQueryTransformEditor.textLoad?.query?.sourceRange?.top !== 8 ||
    powerQueryTransformEditor.textLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|trim|clean|replaceValues|splitColumnByDelimiter|fillDown|changeType|addCustomColumn|groupBy" ||
    Number(powerQueryTransformEditor.textLoad?.rawTotal) !== 3 ||
    !powerQueryTransformEditor.textLoad?.status?.includes("M1:N3") ||
    !powerQueryTransformEditor.columnAfterSteps?.steps?.some((step) => step.includes("抽出されたテキスト: Code 区切り記号の前")) ||
    !powerQueryTransformEditor.columnAfterSteps?.steps?.some((step) => step.includes("複製された列: Region")) ||
    !powerQueryTransformEditor.columnAfterSteps?.steps?.some((step) => step.includes("移動された列: Region - コピー 先頭")) ||
    !powerQueryTransformEditor.columnAfterSteps?.steps?.some((step) => step.includes("追加されたインデックス列: インデックス")) ||
    !powerQueryTransformEditor.columnAfterSteps?.steps?.some((step) => step.includes("追加された条件列: Size")) ||
    !powerQueryTransformEditor.columnAfterSteps?.steps?.some((step) => step.includes("フィルターされた行: Qty >= 8")) ||
    !powerQueryTransformEditor.columnAfterSteps?.steps?.some((step) => step.includes("保持された列: Region - コピー, Item, Code, インデックス, Size")) ||
    powerQueryTransformEditor.columnAfterSteps?.columns?.join("|") !== "Region - コピー|Item|Code|インデックス|Size" ||
    powerQueryTransformEditor.columnAfterSteps?.preview !== "Region-コピーItemCodeインデックスSizeOsakaBetaB2大TokyoGammaC3小KyotoDeltaD4大" ||
    powerQueryTransformEditor.columnLoad?.values?.join("|") !== "Region - コピー|Item|Code|インデックス|Size|Osaka|Beta|B|2|大|Tokyo|Gamma|C|3|小|Kyoto|Delta|D|4|大" ||
    powerQueryTransformEditor.columnLoad?.query?.sourceRange?.top !== 15 ||
    powerQueryTransformEditor.columnLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|extractBeforeDelimiter|duplicateColumn|moveColumn|addIndexColumn|addConditionalColumn|filterRows|keepColumns" ||
    Number(powerQueryTransformEditor.columnLoad?.rawIndex) !== 2 ||
    !powerQueryTransformEditor.columnLoad?.status?.includes("P1:T4") ||
    !powerQueryTransformEditor.reshapeAfterSteps?.steps?.some((step) => step.includes("変更されたテキスト: City 大文字")) ||
    !powerQueryTransformEditor.reshapeAfterSteps?.steps?.some((step) => step.includes("抽出された最後の文字: Code 3")) ||
    !powerQueryTransformEditor.reshapeAfterSteps?.steps?.some((step) => step.includes("計算された長さ: Code")) ||
    !powerQueryTransformEditor.reshapeAfterSteps?.steps?.some((step) => step.includes("ピボット解除された列: Q1, Q2")) ||
    !powerQueryTransformEditor.reshapeAfterSteps?.steps?.some((step) => step.includes("マージされた列: Dept, 属性")) ||
    !powerQueryTransformEditor.reshapeAfterSteps?.steps?.some((step) => step.includes("反転された行")) ||
    !powerQueryTransformEditor.reshapeAfterSteps?.steps?.some((step) => step.includes("削除された代替行: 開始 2")) ||
    powerQueryTransformEditor.reshapeAfterSteps?.columns?.join("|") !== "Key|City|Code|値" ||
    powerQueryTransformEditor.reshapeAfterSteps?.preview !== "KeyCityCode値South:Q2NAGOYA380North:Q2KYOTO360West:Q2OSAKA340East:Q2TOKYO320" ||
    powerQueryTransformEditor.reshapeLoad?.values?.join("|") !== "Key|City|Code|値|South:Q2|NAGOYA|3|80|North:Q2|KYOTO|3|60|West:Q2|OSAKA|3|40|East:Q2|TOKYO|3|20" ||
    powerQueryTransformEditor.reshapeLoad?.query?.sourceRange?.top !== 22 ||
    powerQueryTransformEditor.reshapeLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|transformTextCase|extractLastCharacters|textLength|unpivotColumns|mergeColumns|reverseRows|removeAlternateRows" ||
    Number(powerQueryTransformEditor.reshapeLoad?.rawCodeLength) !== 3 ||
    Number(powerQueryTransformEditor.reshapeLoad?.rawValue) !== 80 ||
    !powerQueryTransformEditor.reshapeLoad?.status?.includes("V1:Y5") ||
    !powerQueryTransformEditor.numericAfterSteps?.steps?.some((step) => step.includes("削除された空の列")) ||
    !powerQueryTransformEditor.numericAfterSteps?.steps?.some((step) => step.includes("置換されたエラー: Err")) ||
    !powerQueryTransformEditor.numericAfterSteps?.steps?.some((step) => step.includes("変換された数値: Amount 乗算")) ||
    !powerQueryTransformEditor.numericAfterSteps?.steps?.some((step) => step.includes("変換された数値: Amount 丸め")) ||
    !powerQueryTransformEditor.numericAfterSteps?.steps?.some((step) => step.includes("抽出された日付: Date 年")) ||
    powerQueryTransformEditor.numericAfterSteps?.columns?.join("|") !== "Date|Amount|Err" ||
    powerQueryTransformEditor.numericAfterSteps?.preview !== "DateAmountErr202621修正202641OK202561修正" ||
    powerQueryTransformEditor.numericLoad?.values?.join("|") !== "Date|Amount|Err|2026|21|修正|2026|41|OK|2025|61|修正" ||
    powerQueryTransformEditor.numericLoad?.query?.sourceRange?.top !== 29 ||
    powerQueryTransformEditor.numericLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|removeEmptyColumns|replaceErrors|numberOperation|numberOperation|datePart" ||
    Number(powerQueryTransformEditor.numericLoad?.rawYear) !== 2026 ||
    Number(powerQueryTransformEditor.numericLoad?.rawAmount) !== 21 ||
    !powerQueryTransformEditor.numericLoad?.status?.includes("AA1:AC4") ||
    !powerQueryTransformEditor.transposeAfterSteps?.steps?.some((step) => step.includes("降格されたヘッダー")) ||
    !powerQueryTransformEditor.transposeAfterSteps?.steps?.some((step) => step.includes("入れ替えられたテーブル")) ||
    powerQueryTransformEditor.transposeAfterSteps?.columns?.join("|") !== "Column1|Column2|Column3|Column4" ||
    powerQueryTransformEditor.transposeAfterSteps?.preview !== "Column1Column2Column3Column4Column1MetricSalesCostColumn2Q1103Column3Q2204" ||
    powerQueryTransformEditor.transposeLoad?.values?.join("|") !== "Column1|Column2|Column3|Column4|Column1|Metric|Sales|Cost|Column2|Q1|10|3|Column3|Q2|20|4" ||
    powerQueryTransformEditor.transposeLoad?.query?.sourceRange?.top !== 35 ||
    powerQueryTransformEditor.transposeLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|demoteHeaders|transpose" ||
    !powerQueryTransformEditor.transposeLoad?.status?.includes("AE1:AH4") ||
    !powerQueryTransformEditor.rowAfterSteps?.steps?.some((step) => step.includes("削除された上位行: 1")) ||
    !powerQueryTransformEditor.rowAfterSteps?.steps?.some((step) => step.includes("削除された下位行: 1")) ||
    !powerQueryTransformEditor.rowAfterSteps?.steps?.some((step) => step.includes("保持された行範囲: 2 から 2 行")) ||
    !powerQueryTransformEditor.rowAfterSteps?.steps?.some((step) => step.includes("保持された下位行: 1")) ||
    powerQueryTransformEditor.rowAfterSteps?.columns?.join("|") !== "RowID|Value|Note" ||
    powerQueryTransformEditor.rowAfterSteps?.preview !== "RowIDValueNote440tail" ||
    powerQueryTransformEditor.rowLoad?.values?.join("|") !== "RowID|Value|Note|4|40|tail" ||
    powerQueryTransformEditor.rowLoad?.query?.sourceRange?.top !== 40 ||
    powerQueryTransformEditor.rowLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|removeTopRows|removeBottomRows|keepRangeRows|keepBottomRows" ||
    Number(powerQueryTransformEditor.rowLoad?.rawRowId) !== 4 ||
    Number(powerQueryTransformEditor.rowLoad?.rawValue) !== 40 ||
    !powerQueryTransformEditor.rowLoad?.status?.includes("AJ1:AL2") ||
    !powerQueryTransformEditor.keepErrorsAfterSteps?.steps?.some((step) => step.includes("保持されたエラー: Err")) ||
    powerQueryTransformEditor.keepErrorsAfterSteps?.columns?.join("|") !== "Date|Amount|Err|Empty" ||
    powerQueryTransformEditor.keepErrorsAfterSteps?.preview !== "DateAmountErrEmpty2026-05-2410.4#N/A2025-12-3130.6#DIV/0!" ||
    powerQueryTransformEditor.keepErrorsLoad?.values?.join("|") !== "Date|Amount|Err|Empty|2026-05-24|10.4|#N/A||2025-12-31|30.6|#DIV/0!|" ||
    powerQueryTransformEditor.keepErrorsLoad?.query?.sourceRange?.top !== 29 ||
    powerQueryTransformEditor.keepErrorsLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|keepErrors" ||
    powerQueryTransformEditor.keepErrorsLoad?.rawError !== "#N/A" ||
    !powerQueryTransformEditor.keepErrorsLoad?.status?.includes("AM1:AP3") ||
    !powerQueryTransformEditor.pivotAfterSteps?.steps?.some((step) => step.includes("ピボットされた列: Plan / Amount")) ||
    powerQueryTransformEditor.pivotAfterSteps?.columns?.join("|") !== "Region|Basic|Pro" ||
    powerQueryTransformEditor.pivotAfterSteps?.preview !== "RegionBasicProEast1025West712North59" ||
    powerQueryTransformEditor.pivotLoad?.values?.join("|") !== "Region|Basic|Pro|East|10|25|West|7|12|North|5|9" ||
    powerQueryTransformEditor.pivotLoad?.query?.sourceRange?.top !== 48 ||
    powerQueryTransformEditor.pivotLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|pivotColumn" ||
    Number(powerQueryTransformEditor.pivotLoad?.rawBasic) !== 10 ||
    Number(powerQueryTransformEditor.pivotLoad?.rawPro) !== 9 ||
    !powerQueryTransformEditor.pivotLoad?.status?.includes("AQ1:AS4") ||
    !powerQueryTransformEditor.advancedNumberAfterSteps?.steps?.some((step) => step.includes("変換された数値: Value 絶対値")) ||
    !powerQueryTransformEditor.advancedNumberAfterSteps?.steps?.some((step) => step.includes("変換された数値: Value 整数除算")) ||
    !powerQueryTransformEditor.advancedNumberAfterSteps?.steps?.some((step) => step.includes("変換された数値: Value 剰余")) ||
    !powerQueryTransformEditor.advancedNumberAfterSteps?.steps?.some((step) => step.includes("変換された数値: Value パーセント")) ||
    !powerQueryTransformEditor.advancedNumberAfterSteps?.steps?.some((step) => step.includes("変換された数値: Value パーセント比")) ||
    !powerQueryTransformEditor.advancedNumberAfterSteps?.steps?.some((step) => step.includes("変換された数値: Value べき乗")) ||
    powerQueryTransformEditor.advancedNumberAfterSteps?.columns?.join("|") !== "Item|Value" ||
    powerQueryTransformEditor.advancedNumberAfterSteps?.preview !== "ItemValueN14N21N39" ||
    powerQueryTransformEditor.advancedNumberLoad?.values?.join("|") !== "Item|Value|N1|4|N2|1|N3|9" ||
    powerQueryTransformEditor.advancedNumberLoad?.query?.sourceRange?.top !== 57 ||
    powerQueryTransformEditor.advancedNumberLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|numberOperation|numberOperation|numberOperation|numberOperation|numberOperation|numberOperation" ||
    Number(powerQueryTransformEditor.advancedNumberLoad?.rawFirst) !== 4 ||
    Number(powerQueryTransformEditor.advancedNumberLoad?.rawLast) !== 9 ||
    !powerQueryTransformEditor.advancedNumberLoad?.status?.includes("AT1:AU4") ||
    !powerQueryTransformEditor.advancedDateAfterSteps?.steps?.some((step) => step.includes("抽出された日付: Quarter 四半期")) ||
    !powerQueryTransformEditor.advancedDateAfterSteps?.steps?.some((step) => step.includes("抽出された日付: MonthName 月の名前")) ||
    !powerQueryTransformEditor.advancedDateAfterSteps?.steps?.some((step) => step.includes("抽出された日付: DayName 曜日の名前")) ||
    !powerQueryTransformEditor.advancedDateAfterSteps?.steps?.some((step) => step.includes("抽出された日付: StartQuarter 四半期の開始")) ||
    !powerQueryTransformEditor.advancedDateAfterSteps?.steps?.some((step) => step.includes("抽出された日付: EndYear 年の終了")) ||
    powerQueryTransformEditor.advancedDateAfterSteps?.columns?.join("|") !== "Quarter|MonthName|DayName|DayOfWeek|DayOfYear|StartMonth|EndMonth|StartQuarter|EndQuarter|StartYear|EndYear" ||
    !powerQueryTransformEditor.advancedDateAfterSteps?.preview?.includes("25月日曜日01442026-05-012026-05-312026-04-012026-06-302026-01-012026-12-31") ||
    !powerQueryTransformEditor.advancedDateAfterSteps?.preview?.includes("412月木曜日43652026-12-012026-12-312026-10-012026-12-312026-01-012026-12-31") ||
    powerQueryTransformEditor.advancedDateLoad?.headers?.join("|") !== "Quarter|MonthName|DayName|DayOfWeek|DayOfYear|StartMonth|EndMonth|StartQuarter|EndQuarter|StartYear|EndYear" ||
    powerQueryTransformEditor.advancedDateLoad?.firstRow?.join("|") !== "2|5月|日曜日|0|144|2026-05-01|2026-05-31|2026-04-01|2026-06-30|2026-01-01|2026-12-31" ||
    powerQueryTransformEditor.advancedDateLoad?.secondRow?.join("|") !== "4|12月|木曜日|4|365|2026-12-01|2026-12-31|2026-10-01|2026-12-31|2026-01-01|2026-12-31" ||
    powerQueryTransformEditor.advancedDateLoad?.query?.sourceRange?.top !== 63 ||
    powerQueryTransformEditor.advancedDateLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|datePart|datePart|datePart|datePart|datePart|datePart|datePart|datePart|datePart|datePart|datePart" ||
    Number(powerQueryTransformEditor.advancedDateLoad?.rawQuarter) !== 2 ||
    Number(powerQueryTransformEditor.advancedDateLoad?.rawDayOfYear) !== 365 ||
    powerQueryTransformEditor.advancedDateLoad?.rawEndYear !== "2026-12-31" ||
    !powerQueryTransformEditor.advancedDateLoad?.status?.includes("AW1:BG3") ||
    !powerQueryTransformEditor.advancedTextAfterSteps?.steps?.some((step) => step.includes("抽出されたテキスト: Text 区切り記号の間")) ||
    !powerQueryTransformEditor.advancedTextAfterSteps?.steps?.some((step) => step.includes("保持された文字: Text")) ||
    !powerQueryTransformEditor.advancedTextAfterSteps?.steps?.some((step) => step.includes("追加されたプレフィックス: Text")) ||
    !powerQueryTransformEditor.advancedTextAfterSteps?.steps?.some((step) => step.includes("追加されたサフィックス: Text")) ||
    !powerQueryTransformEditor.advancedTextAfterSteps?.steps?.some((step) => step.includes("抽出された範囲: Text 3, 4")) ||
    !powerQueryTransformEditor.advancedTextAfterSteps?.steps?.some((step) => step.includes("削除された文字: Text")) ||
    powerQueryTransformEditor.advancedTextAfterSteps?.columns?.join("|") !== "Text" ||
    powerQueryTransformEditor.advancedTextAfterSteps?.preview !== "Texta12c34e56" ||
    powerQueryTransformEditor.advancedTextLoad?.values?.join("|") !== "Text|a12|c34|e56" ||
    powerQueryTransformEditor.advancedTextLoad?.query?.sourceRange?.top !== 68 ||
    powerQueryTransformEditor.advancedTextLoad?.query?.transformSteps?.map((step) => step.type).join("|") !== "promoteHeaders|extractBetweenDelimiters|keepCharacters|addPrefix|addSuffix|extractRangeCharacters|removeCharacters" ||
    powerQueryTransformEditor.advancedTextLoad?.rawFirst !== "a12" ||
    powerQueryTransformEditor.advancedTextLoad?.rawLast !== "e56" ||
    !powerQueryTransformEditor.advancedTextLoad?.status?.includes("BI1:BI4")
  ) {
    throw new Error("Power Query transform editor should support table/web sources, text/number/date transforms, split/extract/merge/unpivot columns, fill, custom/conditional/index columns, column keep/move/duplicate, grouping, error handling, transpose/demote, row transforms, filtering, load-to, data types, and persisted steps.");
  }


  if (
    !whatIfGoalSeekDataTable.menu?.includes("ゴール シーク") ||
    !whatIfGoalSeekDataTable.menu?.includes("データ テーブル") ||
    !whatIfGoalSeekDataTable.whatIfMenuKeyboard?.accessKeys?.includes("S") ||
    !whatIfGoalSeekDataTable.whatIfMenuKeyboard?.accessKeys?.includes("G") ||
    !whatIfGoalSeekDataTable.whatIfMenuKeyboard?.accessKeys?.includes("D") ||
    whatIfGoalSeekDataTable.whatIfMenuKeyboard?.focusedAction !== "scenario-manager" ||
    whatIfGoalSeekDataTable.whatIfMenuKeyboard?.dialogTitle !== "ゴール シーク" ||
    !whatIfGoalSeekDataTable.whatIfMenuKeyboard?.menuClosed ||
    whatIfGoalSeekDataTable.whatIfMenuKeyboard?.editingCellCount !== 0 ||
    !whatIfGoalSeekDataTable.invalidGoalSeek?.message?.includes("数式入力セル") ||
    whatIfGoalSeekDataTable.invalidGoalSeek?.dialogVisible !== true ||
    whatIfGoalSeekDataTable.invalidGoalSeek?.b1 !== "10" ||
    whatIfGoalSeekDataTable.goalSeekResult?.title !== "ゴール シークの結果" ||
    !whatIfGoalSeekDataTable.goalSeekResult?.message?.includes("解が見つかりました") ||
    !whatIfGoalSeekDataTable.goalSeekResult?.message?.includes("目標値: 50") ||
    !whatIfGoalSeekDataTable.goalSeekResult?.message?.includes("現在の値: 50") ||
    whatIfGoalSeekDataTable.goalSeekResult?.buttons?.join("|") !== "OK" ||
    Math.abs(Number(whatIfGoalSeekDataTable.goalSeek?.b1) - 25) > 0.01 ||
    Math.abs(Number(whatIfGoalSeekDataTable.goalSeek?.b2) - 50) > 0.01 ||
    !whatIfGoalSeekDataTable.goalSeek?.status?.includes("ゴール シーク") ||
    whatIfGoalSeekDataTable.crossGoalSeek?.title !== "ゴール シークの結果" ||
    !whatIfGoalSeekDataTable.crossGoalSeek?.message?.includes("目標値: 60") ||
    !whatIfGoalSeekDataTable.crossGoalSeek?.message?.includes("現在の値: 60") ||
    !whatIfGoalSeekDataTable.crossGoalSeek?.activeSheet?.includes("Inputs") ||
    Math.abs(Number(whatIfGoalSeekDataTable.crossGoalSeek?.inputsB1) - 20) > 0.01 ||
    !whatIfGoalSeekDataTable.crossGoalSeek?.status?.includes("ゴール シーク") ||
    !whatIfGoalSeekDataTable.invalidDataTable?.message?.includes("数式") ||
    whatIfGoalSeekDataTable.invalidDataTable?.dialogVisible !== true ||
    whatIfGoalSeekDataTable.dataTable?.values?.join("|") !== "10|20|30" ||
    Math.abs(Number(whatIfGoalSeekDataTable.dataTable?.b1) - 25) > 0.01 ||
    whatIfGoalSeekDataTable.dataTable?.formula !== "{=TABLE(,B1)}" ||
    whatIfGoalSeekDataTable.twoVariableDataTable?.values?.join("|") !== "6|11|7|12" ||
    Math.abs(Number(whatIfGoalSeekDataTable.twoVariableDataTable?.b1) - 25) > 0.01 ||
    Math.abs(Number(whatIfGoalSeekDataTable.twoVariableDataTable?.b3) - 3) > 0.01 ||
    whatIfGoalSeekDataTable.twoVariableDataTable?.formula !== "{=TABLE(B1,B3)}" ||
    !whatIfGoalSeekDataTable.twoVariableDataTable?.status?.includes("データ テーブル") ||
    whatIfGoalSeekDataTable.crossSheetDataTable?.values?.join("|") !== "6|12|18" ||
    whatIfGoalSeekDataTable.crossSheetDataTable?.formula !== "{=TABLE(,Inputs!B1)}" ||
    !whatIfGoalSeekDataTable.crossSheetDataTable?.activeSheet?.includes("Inputs") ||
    Math.abs(Number(whatIfGoalSeekDataTable.crossSheetDataTable?.inputsB1) - 20) > 0.01 ||
    !whatIfGoalSeekDataTable.crossSheetDataTable?.status?.includes("データ テーブル") ||
    !whatIfGoalSeekDataTable.dataTable?.status?.includes("データ テーブル") ||
    whatIfGoalSeekDataTable.groupedGoalSeekMessage?.title !== "ゴール シークの結果" ||
    !whatIfGoalSeekDataTable.groupedGoalSeekMessage?.message?.includes("対象シート数: 2") ||
    Math.abs(Number(whatIfGoalSeekDataTable.groupedGoalSeek?.firstB1) - 20) > 0.01 ||
    Math.abs(Number(whatIfGoalSeekDataTable.groupedGoalSeek?.secondB1) - 20) > 0.01 ||
    !whatIfGoalSeekDataTable.groupedGoalSeek?.status?.includes("2 個のシート") ||
    Math.abs(Number(whatIfGoalSeekDataTable.groupedGoalSeekUndo?.firstB1) - 5) > 0.01 ||
    Math.abs(Number(whatIfGoalSeekDataTable.groupedGoalSeekUndo?.secondB1) - 10) > 0.01 ||
    !whatIfGoalSeekDataTable.groupedGoalSeekUndo?.status?.includes("元に戻しました") ||
    Math.abs(Number(whatIfGoalSeekDataTable.groupedGoalSeekRedo?.firstB1) - 20) > 0.01 ||
    Math.abs(Number(whatIfGoalSeekDataTable.groupedGoalSeekRedo?.secondB1) - 20) > 0.01 ||
    !whatIfGoalSeekDataTable.groupedGoalSeekRedo?.status?.includes("やり直しました") ||
    whatIfGoalSeekDataTable.groupedDataTable?.first?.values?.join("|") !== "2|4|6" ||
    whatIfGoalSeekDataTable.groupedDataTable?.second?.values?.join("|") !== "3|6|9" ||
    whatIfGoalSeekDataTable.groupedDataTable?.first?.formulas?.some((formula) => formula !== "{=TABLE(,B1)}") ||
    whatIfGoalSeekDataTable.groupedDataTable?.second?.formulas?.some((formula) => formula !== "{=TABLE(,B1)}") ||
    !whatIfGoalSeekDataTable.groupedDataTable?.status?.includes("2 個のシート") ||
    whatIfGoalSeekDataTable.groupedDataTableUndo?.first?.values?.join("|") !== "||" ||
    whatIfGoalSeekDataTable.groupedDataTableUndo?.second?.values?.join("|") !== "||" ||
    !whatIfGoalSeekDataTable.groupedDataTableUndo?.status?.includes("元に戻しました") ||
    whatIfGoalSeekDataTable.groupedDataTableRedo?.first?.values?.join("|") !== "2|4|6" ||
    whatIfGoalSeekDataTable.groupedDataTableRedo?.second?.values?.join("|") !== "3|6|9" ||
    !whatIfGoalSeekDataTable.groupedDataTableRedo?.status?.includes("やり直しました") ||
    whatIfGoalSeekDataTable.groupedWhatIfCleanup?.tabs?.join("|") !== "WhatIf|Inputs"
  ) {
    throw new Error("What-If Analysis should keep dialogs open on invalid input, solve Goal Seek, calculate one/two-variable and sheet-qualified Data Tables, and synchronize grouped sheets with undo/redo like Excel.");
  }


  if (
    scenarioManagerSummary.duplicateName?.title !== "Cht WebSheet" ||
    !scenarioManagerSummary.duplicateName?.message?.includes("同じ名前") ||
    scenarioManagerSummary.duplicateName?.dialogVisible !== true ||
    scenarioManagerSummary.deletePrompt?.title !== "Cht WebSheet" ||
    !scenarioManagerSummary.deletePrompt?.message?.includes("削除") ||
    scenarioManagerSummary.deletePrompt?.buttons?.join("|") !== "削除|キャンセル" ||
    scenarioManagerSummary.deleteUndoBefore?.join("|") !== "Low" ||
    scenarioManagerSummary.deleteUndoAfter?.join("|") !== "Low|High" ||
    scenarioManagerSummary.editUndoBefore?.join("|") !== "Low|High2" ||
    scenarioManagerSummary.editUndoAfter?.join("|") !== "Low|High" ||
    scenarioManagerSummary.manager?.options?.join("|") !== "Low|High" ||
    scenarioManagerSummary.manager?.summaryEnabled !== true ||
    scenarioManagerSummary.scenarioShow?.b1 !== "20" ||
    scenarioManagerSummary.scenarioShow?.b2 !== "40" ||
    !scenarioManagerSummary.scenarioShow?.status?.includes("High") ||
    scenarioManagerSummary.scenarioShowUndo?.b1 !== "10" ||
    scenarioManagerSummary.scenarioShowUndo?.b2 !== "20" ||
    !scenarioManagerSummary.scenarioShowUndo?.status?.includes("元に戻しました") ||
    scenarioManagerSummary.scenarioShowRedo?.b1 !== "20" ||
    scenarioManagerSummary.scenarioShowRedo?.b2 !== "40" ||
    !scenarioManagerSummary.scenarioShowRedo?.status?.includes("やり直しました") ||
    scenarioManagerSummary.summaryDialog?.title !== "シナリオの概要" ||
    !scenarioManagerSummary.summaryDialog?.reportTypes?.some((item) => item.value === "summary" && item.checked) ||
    scenarioManagerSummary.summary?.headers?.join("|") !== "変化させるセル|現在値|Low|High" ||
    scenarioManagerSummary.summary?.changing?.join("|") !== "WhatIf!B1|10|5|20" ||
    scenarioManagerSummary.summary?.resultHeader?.join("|") !== "結果セル|現在値|Low|High" ||
    scenarioManagerSummary.summary?.result?.join("|") !== "WhatIf!B2|20|10|40" ||
    !scenarioManagerSummary.summary?.status?.includes("シナリオの概要") ||
    scenarioManagerSummary.summaryUndo?.activeSheet !== "WhatIf" ||
    scenarioManagerSummary.summaryUndo?.sheetNames?.some((name) => name.startsWith("シナリオの概要")) ||
    !scenarioManagerSummary.summaryUndo?.status?.includes("元に戻しました") ||
    !scenarioManagerSummary.summaryRedo?.activeSheet?.startsWith("シナリオの概要") ||
    scenarioManagerSummary.summaryRedo?.title !== "シナリオの概要" ||
    !scenarioManagerSummary.summaryRedo?.status?.includes("やり直しました") ||
    !scenarioManagerSummary.pivotSummary?.activeSheet?.startsWith("シナリオ ピボット") ||
    scenarioManagerSummary.pivotSummary?.title !== "シナリオ ピボットテーブル レポート" ||
    scenarioManagerSummary.pivotSummary?.groups?.join("|") !== "値|変化させるセル|結果セル" ||
    scenarioManagerSummary.pivotSummary?.headers?.join("|") !== "シナリオ|WhatIf!B1|WhatIf!B2" ||
    scenarioManagerSummary.pivotSummary?.current?.join("|") !== "現在値|10|20" ||
    scenarioManagerSummary.pivotSummary?.low?.join("|") !== "Low|5|10" ||
    scenarioManagerSummary.pivotSummary?.high?.join("|") !== "High|20|40" ||
    !scenarioManagerSummary.pivotSummary?.status?.includes("シナリオ ピボット")
  ) {
    throw new Error("Scenario Manager should add, show, undo/redo, and create summary/pivot reports with current/scenario result values.");
  }


  if (
    flashFillDataPatterns.values?.join("|") !== "JS|AL|GH|alpha.one|beta.long|gamma|Smith, John|Lovelace, Ada|Hopper, Grace|001|120|045|AB|CD|EF|JP|OS|KY|2026-05-24|2026-06-01|2025-12-31|090-1234-5678|080-9988-7766|070-1111-2222|2026/05/24|2026/06/01|2025/12/31" ||
    !flashFillDataPatterns.status?.includes("フラッシュ フィル")
  ) {
    throw new Error("Flash Fill should infer Excel-like initials, delimiter extraction, token reordering, character extraction, bracket extraction, token rejoining, and fixed slice formatting patterns.");
  }


  if (
    dataCurrentRegion.filterCols?.join("|") !== "1|2" ||
    !dataCurrentRegion.filterStatus?.includes("A1:B5") ||
    dataCurrentRegion.rowHeaderFilter?.cols?.join("|") !== "1|2" ||
    dataCurrentRegion.rowHeaderFilter?.rows?.join("|") !== "1|1" ||
    dataCurrentRegion.rowHeaderFilter?.label !== "1:1" ||
    !dataCurrentRegion.rowHeaderFilter?.status?.includes("A1:B5") ||
    dataCurrentRegion.chtHeaderFilter?.label !== "15:15" ||
    dataCurrentRegion.chtHeaderFilter?.rows?.length < 2 ||
    dataCurrentRegion.chtHeaderFilter?.rows?.some((row) => row !== 15) ||
    dataCurrentRegion.chtHeaderFilter?.modelRange?.top !== 15 ||
    dataCurrentRegion.chtHeaderFilter?.modelRange?.left !== 2 ||
    dataCurrentRegion.chtHeaderFilter?.modelRange?.right < 40 ||
    dataCurrentRegion.chtHeaderFilter?.modelRange?.bottom <= 21 ||
    dataCurrentRegion.chtHeaderFilter?.cols?.length < 6 ||
    dataCurrentRegion.chtHeaderFilter?.status?.includes("14") ||
    dataCurrentRegion.blankHeaderFilter?.cols?.join("|") !== "1|2|3" ||
    dataCurrentRegion.blankHeaderFilter?.rows?.some((row) => row !== 1) ||
    dataCurrentRegion.blankHeaderFilter?.range?.top !== 1 ||
    dataCurrentRegion.blankHeaderFilter?.range?.left !== 1 ||
    dataCurrentRegion.blankHeaderFilter?.range?.right !== 3 ||
    dataCurrentRegion.blankHeaderFilter?.range?.bottom !== 4 ||
    dataCurrentRegion.tablePriorityFilter?.cols?.join("|") !== "2|3|4" ||
    dataCurrentRegion.tablePriorityFilter?.range?.top !== 2 ||
    dataCurrentRegion.tablePriorityFilter?.range?.left !== 2 ||
    dataCurrentRegion.tablePriorityFilter?.range?.right !== 4 ||
    dataCurrentRegion.tablePriorityFilter?.range?.bottom !== 5 ||
    dataCurrentRegion.tablePriorityFilter?.table?.showFilterButton !== true ||
    dataCurrentRegion.tablePriorityFilter?.table?.autoFilterState?.range?.left !== 2 ||
    dataCurrentRegion.tablePriorityFilter?.table?.autoFilterState?.range?.right !== 4 ||
    dataCurrentRegion.toggleFilter?.hasAutoFilter !== false ||
    dataCurrentRegion.toggleFilter?.buttonCount !== 0 ||
    dataCurrentRegion.sortKeys?.join("|") !== "Name|Score" ||
    dataCurrentRegion.duplicateLabels?.join("|") !== "Name|Score"
  ) {
    throw new Error("Data tab commands should expand a single active cell to the Excel current region without leaking into separated data blocks.");
  }


  if (
    queryConnectionsPane.loadDialog?.range !== "A1:B5" ||
    queryConnectionsPane.loadDialog?.hasHeader !== true ||
    !queryConnectionsPane.initialPane?.tabs?.some((tab) => tab.text === "クエリ 1" && tab.active) ||
    !queryConnectionsPane.initialPane?.rows?.[0]?.includes("Query1") ||
    !queryConnectionsPane.initialPane?.rows?.[0]?.includes("読み込み先") ||
    queryConnectionsPane.loadToDialog?.title !== "データのインポート" ||
    !queryConnectionsPane.loadToDialog?.modes?.some((mode) => mode.value === "table" && mode.checked && mode.label.includes("テーブル")) ||
    !queryConnectionsPane.loadToDialog?.modes?.some((mode) => mode.value === "pivotTable" && !mode.disabled && mode.label.includes("ピボットテーブル")) ||
    !queryConnectionsPane.loadToDialog?.modes?.some((mode) => mode.value === "pivotChart" && !mode.disabled && mode.label.includes("ピボットグラフ")) ||
    !queryConnectionsPane.loadToDialog?.modes?.some((mode) => mode.value === "connectionOnly" && mode.label.includes("接続の作成のみ")) ||
    !queryConnectionsPane.loadToDialog?.locations?.some((location) => location.value === "new" && !location.disabled && location.label.includes("新規ワークシート")) ||
    !queryConnectionsPane.loadToDialog?.destination?.includes("A1") ||
    !queryConnectionsPane.loadToDialog?.propertyButton?.includes("プロパティ") ||
    !queryConnectionsPane.propertiesInitial?.tabs?.some((tab) => tab.text === "使用状況" && tab.active) ||
    !queryConnectionsPane.propertiesInitial?.tabs?.some((tab) => tab.text === "定義") ||
    !queryConnectionsPane.propertiesInitial?.usageLabels?.some((label) => label.includes("次の間隔で更新する")) ||
    queryConnectionsPane.propertiesInitial?.interval !== "60" ||
    queryConnectionsPane.propertiesInitial?.intervalDisabled !== true ||
    !queryConnectionsPane.propertiesInitial?.typeNote?.includes("ワークシート接続") ||
    !queryConnectionsPane.definitionInitial?.labels?.includes("接続文字列") ||
    !queryConnectionsPane.definitionInitial?.labels?.includes("コマンドの種類") ||
    !queryConnectionsPane.definitionInitial?.connectionString?.includes("WORKSHEET;") ||
    !queryConnectionsPane.definitionInitial?.commandText?.includes("Filter!A1:B5") ||
    queryConnectionsPane.definitionInitial?.commandType !== "table" ||
    !queryConnectionsPane.definitionInitial?.commandTypeOptions?.includes("SQL") ||
    !queryConnectionsPane.definitionInitial?.definitionChecks?.some((label) => label.includes("常に接続ファイル")) ||
    queryConnectionsPane.intervalEnabledAfterCheck !== true ||
    !queryConnectionsPane.renamed?.rows?.[0]?.includes("SalesQuery") ||
    !queryConnectionsPane.renamed?.status?.includes("接続プロパティ") ||
    queryConnectionsPane.renamed?.query?.refreshIntervalMinutes !== 15 ||
    queryConnectionsPane.renamed?.query?.refreshEveryEnabled !== true ||
    queryConnectionsPane.renamed?.query?.alwaysUseConnectionFile !== true ||
    queryConnectionsPane.renamed?.query?.failIfConnectionFileMissing !== true ||
    queryConnectionsPane.renamed?.query?.commandType !== "table" ||
    queryConnectionsPane.renamed?.refreshTimerCount !== 1 ||
    !queryConnectionsPane.propertyUndo?.rows?.[0]?.includes("Query1") ||
    !queryConnectionsPane.propertyUndo?.status?.includes("元に戻しました") ||
    queryConnectionsPane.propertyUndo?.refreshTimerCount !== 0 ||
    !queryConnectionsPane.propertyRedo?.rows?.[0]?.includes("SalesQuery") ||
    !queryConnectionsPane.propertyRedo?.status?.includes("やり直しました") ||
    queryConnectionsPane.propertyRedo?.refreshTimerCount !== 1 ||
    !queryConnectionsPane.refreshAllSkipped?.includes("更新する接続はありません") ||
    !queryConnectionsPane.connectionsTab?.tabs?.some((tab) => tab.text === "接続 1" && tab.active) ||
    !queryConnectionsPane.connectionsTab?.rows?.[0]?.includes("すべて更新の対象外") ||
    !queryConnectionsPane.connectionsTab?.rows?.[0]?.includes("15 分ごとに更新") ||
    !queryConnectionsPane.connectionsTab?.rows?.[0]?.includes("ファイルを開くときに更新") ||
    !queryConnectionsPane.connectionsTab?.rows?.[0]?.includes("バックグラウンド更新") ||
    !queryConnectionsPane.connectionsTab?.rows?.[0]?.includes("保存前にデータ削除") ||
    !queryConnectionsPane.connectionsTab?.rows?.[0]?.includes("削除") ||
    !queryConnectionsPane.singleRefresh?.rows?.[0]?.includes("最終更新") ||
    !queryConnectionsPane.singleRefresh?.status?.includes("SalesQuery を更新") ||
    queryConnectionsPane.connectionDeletePrompt?.title !== "Cht WebSheet" ||
    !queryConnectionsPane.connectionDeletePrompt?.message?.includes("SalesQuery") ||
    queryConnectionsPane.connectionDeletePrompt?.buttons?.join("|") !== "削除|キャンセル" ||
    queryConnectionsPane.connectionDeleteCancelCount !== 1 ||
    queryConnectionsPane.deletePrompt?.title !== "Cht WebSheet" ||
    !queryConnectionsPane.deletePrompt?.message?.includes("SalesQuery") ||
    queryConnectionsPane.deletePrompt?.buttons?.join("|") !== "削除|キャンセル" ||
    queryConnectionsPane.afterCancelDeleteCount !== 1 ||
    queryConnectionsPane.afterDelete?.itemCount !== 0 ||
    !queryConnectionsPane.afterDelete?.empty?.includes("クエリがありません") ||
    !queryConnectionsPane.afterDelete?.status?.includes("削除") ||
    queryConnectionsPane.afterDeleteUndo?.itemCount !== 1 ||
    !queryConnectionsPane.afterDeleteUndo?.rows?.[0]?.includes("SalesQuery") ||
    !queryConnectionsPane.afterDeleteUndo?.status?.includes("元に戻しました")
  ) {
    throw new Error("Queries & Connections should expose query/connection tabs, properties, refresh, and delete confirmation like Excel.");
  }


  if (
    !dataGetConnectionMenus.getMenu?.labels?.some((label) => label.includes("Excel ブックから")) ||
    !dataGetConnectionMenus.getMenu?.labels?.some((label) => label.includes("テキストまたは CSV から")) ||
    !dataGetConnectionMenus.getMenu?.labels?.some((label) => label.includes("Web から")) ||
    !dataGetConnectionMenus.getMenu?.labels?.some((label) => label.includes("JSON から")) ||
    !dataGetConnectionMenus.getMenu?.labels?.some((label) => label.includes("XML から")) ||
    !dataGetConnectionMenus.getMenu?.labels?.some((label) => label.includes("既存の接続")) ||
    !dataGetConnectionMenus.getMenu?.accessKeys?.includes("W") ||
    dataGetConnectionMenus.getMenu?.focusedAction !== "get-workbook" ||
    dataGetConnectionMenus.getMenu?.j1BeforeShortcut !== "" ||
    dataGetConnectionMenus.dataGetShortcut?.dialogTitle !== "Web から" ||
    dataGetConnectionMenus.dataGetShortcut?.j1AfterShortcut !== "" ||
    dataGetConnectionMenus.dataGetShortcut?.editingCellCount !== 0 ||
    dataGetConnectionMenus.dataGetShortcut?.focusedDialogField !== true ||
    !dataGetConnectionMenus.dataGetShortcut?.menuClosed ||
    dataGetConnectionMenus.webDialog?.title !== "Web から" ||
    !dataGetConnectionMenus.webDialog?.captions?.includes("URL") ||
    dataGetConnectionMenus.webImport?.j1 !== "Name" ||
    dataGetConnectionMenus.webImport?.k2 !== "9" ||
    dataGetConnectionMenus.webImport?.query?.sourceType !== "web" ||
    !dataGetConnectionMenus.webImport?.query?.sourceUrl?.startsWith("data:text/csv") ||
    !dataGetConnectionMenus.webImport?.status?.includes("Web データ") ||
    !dataGetConnectionMenus.recentMenu?.labels?.some((label) => label.includes("WebQuery1")) ||
    dataGetConnectionMenus.recentSelect?.selected !== "J1" ||
    !dataGetConnectionMenus.recentSelect?.status?.includes("WebQuery1 の範囲へ移動") ||
    !dataGetConnectionMenus.recentSelect?.frameStyle?.includes("width:") ||
    !dataGetConnectionMenus.existingDialog?.options?.[0]?.includes("WebQuery1") ||
    !dataGetConnectionMenus.existingDialog?.details?.includes("Web クエリ接続") ||
    !dataGetConnectionMenus.existingDialog?.details?.includes("WEB;data:text/csv") ||
    dataGetConnectionMenus.importDataDialog?.title !== "データのインポート" ||
    !dataGetConnectionMenus.importDataDialog?.note?.includes("WebQuery1") ||
    !dataGetConnectionMenus.importDataDialog?.modes?.some((mode) => mode.value === "table" && mode.checked && mode.label.includes("テーブル")) ||
    !dataGetConnectionMenus.importDataDialog?.modes?.some((mode) => mode.value === "pivotTable" && !mode.disabled && mode.label.includes("ピボットテーブル")) ||
    !dataGetConnectionMenus.importDataDialog?.modes?.some((mode) => mode.value === "pivotChart" && !mode.disabled && mode.label.includes("ピボットグラフ")) ||
    !dataGetConnectionMenus.importDataDialog?.modes?.some((mode) => mode.value === "connectionOnly" && mode.label.includes("接続の作成のみ")) ||
    !dataGetConnectionMenus.importDataDialog?.locations?.some((location) => location.value === "new" && !location.disabled && location.label.includes("新規ワークシート")) ||
    !dataGetConnectionMenus.importDataDialog?.destination?.includes("M1") ||
    !dataGetConnectionMenus.importDataDialog?.buttons?.some((button) => button.includes("プロパティ")) ||
    dataGetConnectionMenus.connectionOnlySync?.destinationDisabled !== true ||
    dataGetConnectionMenus.connectionOnlySync?.locationsDisabled !== true ||
    dataGetConnectionMenus.connectionOnlyImport?.m1 !== "" ||
    dataGetConnectionMenus.connectionOnlyImport?.query?.loadMode !== "connectionOnly" ||
    dataGetConnectionMenus.connectionOnlyImport?.query?.addToDataModel !== true ||
    !dataGetConnectionMenus.connectionOnlyImport?.dataModel?.tables?.some((table) => table.name === "WebQuery1") ||
    dataGetConnectionMenus.dataModelDialog?.title !== "データ モデルの管理" ||
    !dataGetConnectionMenus.dataModelDialog?.options?.some((option) => option.includes("WebQuery1")) ||
    !dataGetConnectionMenus.dataModelDialog?.details?.includes("WEB;data:text/csv") ||
    !dataGetConnectionMenus.dataModelDialog?.details?.includes("列: Name, Score") ||
    !dataGetConnectionMenus.dataModelDialog?.relationships?.includes("ありません") ||
    !dataGetConnectionMenus.dataModelDialog?.actions?.some((action) => action.text.includes("ピボットテーブル") && !action.disabled) ||
    dataGetConnectionMenus.importedConnection?.m1 !== "Name" ||
    dataGetConnectionMenus.importedConnection?.n2 !== "9" ||
    !dataGetConnectionMenus.importedConnection?.selected?.includes("M1") ||
    !dataGetConnectionMenus.importedConnection?.status?.includes("WebQuery1") ||
    dataGetConnectionMenus.importedConnection?.query?.range?.left !== 13 ||
    dataGetConnectionMenus.pivotImportDialog?.destinationDisabled !== true ||
    dataGetConnectionMenus.pivotImportDialog?.pivotChecked !== true ||
    dataGetConnectionMenus.pivotImportDialog?.newWorksheetChecked !== true ||
    !dataGetConnectionMenus.pivotImport?.active?.includes("PivotTable") ||
    !dataGetConnectionMenus.pivotImport?.headers?.some((header) => header.includes("データの個数")) ||
    !dataGetConnectionMenus.pivotImport?.status?.includes("ピボットテーブル") ||
    !dataGetConnectionMenus.pivotChartImport?.active?.includes("PivotChart") ||
    dataGetConnectionMenus.pivotChartImport?.imageCount < 1 ||
    !dataGetConnectionMenus.pivotChartImport?.chartName?.includes("ピボットグラフ") ||
    !dataGetConnectionMenus.existingProperties?.connectionString?.startsWith("WEB;data:text/csv") ||
    !dataGetConnectionMenus.existingProperties?.commandText?.startsWith("data:text/csv") ||
    dataGetConnectionMenus.existingAfterPropertiesCancel?.title !== "既存の接続" ||
    !dataGetConnectionMenus.existingAfterPropertiesCancel?.options?.[0]?.includes("WebQuery1") ||
    dataGetConnectionMenus.existingDeletePrompt?.title !== "Cht WebSheet" ||
    !dataGetConnectionMenus.existingDeletePrompt?.message?.includes("WebQuery1") ||
    dataGetConnectionMenus.existingDeletePrompt?.buttons?.join("|") !== "削除|キャンセル" ||
    dataGetConnectionMenus.existingDeletePrompt?.dialogVisible !== true ||
    dataGetConnectionMenus.existingDeleteCancelCount !== 1 ||
    !dataGetConnectionMenus.refreshMenu?.labels?.some((label) => label.includes("すべて更新")) ||
    !dataGetConnectionMenus.refreshMenu?.labels?.some((label) => label.includes("接続のプロパティ")) ||
    !dataGetConnectionMenus.refreshAll?.status?.includes("1 件の接続を更新") ||
    dataGetConnectionMenus.refreshAll?.query?.refreshCount < 1 ||
    !dataGetConnectionMenus.webQueryPaneBeforeEdit?.rows?.some((row) => row.includes("WebQuery1") && row.includes("編集")) ||
    dataGetConnectionMenus.webQueryPaneBeforeEdit?.editButtons < 1 ||
    !dataGetConnectionMenus.webQueryEditDialog?.title?.includes("Power Query エディター") ||
    dataGetConnectionMenus.webQueryEditDialog?.queryName !== "WebQuery1" ||
    !dataGetConnectionMenus.webQueryEditDialog?.preview?.includes("Name") ||
    !dataGetConnectionMenus.webQueryEditDialog?.preview?.includes("Web") ||
    dataGetConnectionMenus.webQueryEditDialog?.loadToVisible !== true ||
    !dataGetConnectionMenus.webQueryContextMenu?.labels?.some((label) => label.includes("複製")) ||
    !dataGetConnectionMenus.webQueryContextMenu?.labels?.some((label) => label.includes("参照")) ||
    !dataGetConnectionMenus.webQueryContextMenu?.labels?.some((label) => label.includes("名前の変更")) ||
    dataGetConnectionMenus.webQueryContextMenu?.focusedBefore !== "edit" ||
    dataGetConnectionMenus.webQueryContextMenu?.focusedAfterDown !== "load-to" ||
    dataGetConnectionMenus.duplicatedQuery?.query?.name !== "WebQuery1 (2)" ||
    dataGetConnectionMenus.duplicatedQuery?.query?.loadMode !== "connectionOnly" ||
    !dataGetConnectionMenus.duplicatedQuery?.status?.includes("複製") ||
    dataGetConnectionMenus.referenceCreated?.query?.sourceType !== "queryReference" ||
    dataGetConnectionMenus.referenceCreated?.query?.sourceQueryName !== "WebQuery1" ||
    dataGetConnectionMenus.referenceCreated?.query?.loadMode !== "connectionOnly" ||
    !dataGetConnectionMenus.referenceCreated?.status?.includes("参照") ||
    dataGetConnectionMenus.renameDialog?.title !== "名前の変更" ||
    !dataGetConnectionMenus.renameDialog?.value?.includes("WebQuery1") ||
    dataGetConnectionMenus.renamedReference?.query?.name !== "WebQueryRef" ||
    dataGetConnectionMenus.renamedReference?.query?.sourceType !== "queryReference" ||
    !dataGetConnectionMenus.renamedReference?.status?.includes("WebQueryRef") ||
    dataGetConnectionMenus.referenceProperties?.title !== "接続のプロパティ" ||
    dataGetConnectionMenus.referenceProperties?.name !== "WebQueryRef" ||
    !dataGetConnectionMenus.referenceProperties?.typeNote?.includes("クエリ参照接続") ||
    !dataGetConnectionMenus.referenceProperties?.connectionString?.includes("Reference=WebQuery1") ||
    !dataGetConnectionMenus.referenceProperties?.commandText?.includes('Source = #"WebQuery1"') ||
    dataGetConnectionMenus.referenceProperties?.commandType !== "default" ||
    dataGetConnectionMenus.referenceLoadTo?.p1 !== "Name" ||
    dataGetConnectionMenus.referenceLoadTo?.q2 !== "9" ||
    dataGetConnectionMenus.referenceLoadTo?.query?.loadMode !== "table" ||
    !dataGetConnectionMenus.referenceLoadTo?.status?.includes("WebQueryRef") ||
    dataGetConnectionMenus.referenceRefresh?.p1 !== "Name" ||
    dataGetConnectionMenus.referenceRefresh?.q2 !== "9" ||
    dataGetConnectionMenus.referenceRefresh?.refreshCount < 1 ||
    !dataGetConnectionMenus.referenceRefresh?.status?.includes("WebQueryRef を更新") ||
    !dataGetConnectionMenus.blankQueryMenu?.labels?.some((label) => label.includes("空のクエリ")) ||
    dataGetConnectionMenus.blankQueryMenu?.blankDisabled !== false ||
    !dataGetConnectionMenus.blankQueryEditor?.title?.includes("Power Query エディター") ||
    !dataGetConnectionMenus.blankQueryEditor?.title?.includes("空のクエリ") ||
    !dataGetConnectionMenus.blankQueryEditor?.preview?.includes("プレビューできるデータがありません") ||
    !dataGetConnectionMenus.blankQueryEditor?.formula?.includes("Source") ||
    !dataGetConnectionMenus.blankQueryEditor?.queryName?.includes("空のクエリ") ||
    dataGetConnectionMenus.blankQueryEditor?.loadToVisible !== true ||
    dataGetConnectionMenus.blankQueryImportDialog?.title !== "データのインポート" ||
    !dataGetConnectionMenus.blankQueryImportDialog?.note?.includes("空のクエリ") ||
    !dataGetConnectionMenus.blankQueryImportDialog?.modes?.some((mode) => mode.value === "connectionOnly" && mode.label.includes("接続の作成のみ")) ||
    dataGetConnectionMenus.blankQueryConnection?.query?.sourceType !== "blankQuery" ||
    dataGetConnectionMenus.blankQueryConnection?.query?.loadMode !== "connectionOnly" ||
    !dataGetConnectionMenus.blankQueryConnection?.status?.includes("接続のみ") ||
    dataGetConnectionMenus.blankQueryExistingSelected !== true ||
    !dataGetConnectionMenus.blankQueryExistingDetails?.selected?.includes("空のクエリ") ||
    !dataGetConnectionMenus.blankQueryExistingDetails?.details?.includes("空のクエリ接続") ||
    !dataGetConnectionMenus.blankQueryExistingDetails?.details?.includes("POWERQUERY;") ||
    dataGetConnectionMenus.blankQueryExistingDetails?.canOpen !== true ||
    dataGetConnectionMenus.blankQueryExistingDetails?.canProperties !== true ||
    !dataGetConnectionMenus.blankQueryPane?.rows?.some((row) => row.includes("空のクエリ") && row.includes("接続のみ")) ||
    dataGetConnectionMenus.blankQueryPaneProperties?.title !== "接続のプロパティ" ||
    !dataGetConnectionMenus.blankQueryPaneProperties?.name?.includes("空のクエリ") ||
    !dataGetConnectionMenus.blankQueryPaneProperties?.typeNote?.includes("空のクエリ接続") ||
    !dataGetConnectionMenus.blankQueryPaneProperties?.connectionString?.startsWith("POWERQUERY;") ||
    dataGetConnectionMenus.blankQueryPaneProperties?.commandType !== "default" ||
    !dataGetConnectionMenus.blankQueryPaneProperties?.commandText?.includes("Source = null") ||
    dataGetConnectionMenus.blankQueryLoadToInitial?.connectionOnlyChecked !== true ||
    dataGetConnectionMenus.blankQueryLoadToInitial?.destinationDisabled !== true ||
    dataGetConnectionMenus.blankQueryLoadToInitial?.locationsDisabled !== true
  ) {
    throw new Error("Data Get/Refresh menus should import Web data, create blank Power Query connections, list existing connections, and refresh registered queries.");
  }


  if (
    relationshipsDialog.empty?.title !== "リレーションシップの管理" ||
    !relationshipsDialog.empty?.details?.includes("ありません") ||
    relationshipsDialog.empty?.editDisabled !== true ||
    relationshipsDialog.empty?.deleteDisabled !== true ||
    relationshipsDialog.editInitial?.title !== "リレーションシップの作成" ||
    !relationshipsDialog.editInitial?.sourceTables?.some((label) => label.includes("Table2")) ||
    !relationshipsDialog.editInitial?.targetTables?.some((label) => label.includes("Table1")) ||
    relationshipsDialog.editInitial?.sourceColumns?.join("|") !== "OrderID|CustomerID|Amount" ||
    relationshipsDialog.editInitial?.targetColumns?.join("|") !== "CustomerID|Name" ||
    !relationshipsDialog.noteAfterColumnChange?.includes("Table2[CustomerID]") ||
    !relationshipsDialog.noteAfterColumnChange?.includes("Table1[CustomerID]") ||
    !relationshipsDialog.created?.options?.[0]?.includes("Table2[CustomerID]") ||
    !relationshipsDialog.created?.details?.includes("関連テーブル: Table1") ||
    !relationshipsDialog.created?.status?.includes("作成") ||
    relationshipsDialog.duplicatePrompt?.title !== "Cht WebSheet" ||
    !relationshipsDialog.duplicatePrompt?.message?.includes("既に存在") ||
    relationshipsDialog.duplicatePrompt?.buttons?.join("|") !== "OK" ||
    relationshipsDialog.duplicatePrompt?.editDialogVisible !== true ||
    relationshipsDialog.deletePrompt?.title !== "Cht WebSheet" ||
    !relationshipsDialog.deletePrompt?.message?.includes("Table2[CustomerID]") ||
    relationshipsDialog.deletePrompt?.buttons?.join("|") !== "削除|キャンセル" ||
    !relationshipsDialog.afterDelete?.details?.includes("ありません") ||
    !relationshipsDialog.afterDelete?.status?.includes("削除") ||
    !relationshipsDialog.undo?.details?.includes("Table2") ||
    !relationshipsDialog.undo?.status?.includes("元に戻しました") ||
    !relationshipsDialog.redo?.details?.includes("ありません") ||
    !relationshipsDialog.redo?.status?.includes("やり直しました") ||
    !relationshipsDialog.autoDetected?.options?.[0]?.includes("Table2[CustomerID]") ||
    !relationshipsDialog.autoDetected?.status?.includes("1 件のリレーションシップを検出") ||
    !relationshipsDialog.autoDetectUndo?.details?.includes("ありません") ||
    !relationshipsDialog.autoDetectUndo?.status?.includes("元に戻しました")
  ) {
    throw new Error("Relationships should create table relationships, prevent duplicates, delete with confirmation, and support undo/redo like Excel.");
  }


  if (
    !refreshAllUndo.afterRefresh?.rows?.[0]?.includes("Query1") ||
    !refreshAllUndo.afterRefresh?.status?.includes("1 件の接続を更新") ||
    !refreshAllUndo.afterUndo?.rows?.[0]?.includes("Query1") ||
    !refreshAllUndo.afterUndo?.status?.includes("元に戻しました：すべて更新") ||
    !refreshAllUndo.afterRedo?.rows?.[0]?.includes("Query1") ||
    !refreshAllUndo.afterRedo?.status?.includes("やり直しました：すべて更新")
  ) {
    throw new Error("Refresh All should update query connections and remain undoable while the Queries & Connections pane is open.");
  }


  if (
    !forecastSheet.activeSheet?.startsWith("Forecast") ||
    forecastSheet.chartCount !== 1 ||
    forecastSheet.headers?.join("|") !== "期間|実績|予測|信頼下限|信頼上限" ||
    forecastSheet.dialogDefaults?.forecastEnd !== "9" ||
    forecastSheet.dialogDefaults?.confidenceLevel !== "95" ||
    forecastSheet.finalPeriod !== "6" ||
    forecastSheet.confidenceLower === "" ||
    forecastSheet.confidenceUpper === "" ||
    !forecastSheet.status?.includes("予測シートを作成") ||
    forecastSheet.afterUndo?.activeSheet !== "ForecastSource" ||
    forecastSheet.afterUndo?.tabs?.some((tab) => tab !== "ForecastSource" && tab.startsWith("Forecast")) ||
    !forecastSheet.afterUndo?.status?.includes("元に戻しました") ||
    !forecastSheet.afterRedo?.activeSheet?.startsWith("Forecast") ||
    forecastSheet.afterRedo?.chartCount !== 1 ||
    !forecastSheet.afterRedo?.status?.includes("やり直しました")
  ) {
    throw new Error("Forecast Sheet should create an Excel-like forecast worksheet and support undo/redo of the inserted sheet.");
  }


  if (
    dataAnalysisToolpak.picker?.title !== "データ分析" ||
    dataAnalysisToolpak.picker?.selected !== "descriptive" ||
    dataAnalysisToolpak.picker?.methods?.slice(0, 6).join("|") !== "分散分析: 一元配置|分散分析: 繰り返しのある二元配置|分散分析: 繰り返しのない二元配置|相関|共分散|基本統計量" ||
    dataAnalysisToolpak.picker?.methods?.includes("記述統計") ||
    !["回帰分析", "乱数発生", "フーリエ解析", "基本統計量"].every((label) => dataAnalysisToolpak.picker?.methods?.includes(label)) ||
    dataAnalysisToolpak.pickerHelp?.title !== "データ分析 ヘルプ" ||
    !dataAnalysisToolpak.pickerHelp?.text?.includes("回帰分析") ||
    !dataAnalysisToolpak.optionsHelp?.text?.includes("残差") ||
    dataAnalysisToolpak.pickerAfterBack?.selected !== "regression" ||
    dataAnalysisToolpak.regressionDialog?.title !== "回帰分析" ||
    dataAnalysisToolpak.regressionDialog?.hasMethodList !== false ||
    !dataAnalysisToolpak.regressionDialog?.visibleFields?.includes("入力 Y 範囲") ||
    !dataAnalysisToolpak.regressionDialog?.visibleFields?.includes("入力 X 範囲") ||
    dataAnalysisToolpak.regressionStatePreserved?.outputMode !== "newWorksheet" ||
    dataAnalysisToolpak.regressionStatePreserved?.outputSheetName !== "保持確認" ||
    dataAnalysisToolpak.regressionStatePreserved?.labels !== true ||
    dataAnalysisToolpak.regressionStatePreserved?.residuals !== true ||
    dataAnalysisToolpak.regressionStatePreserved?.standardizedResiduals !== true ||
    dataAnalysisToolpak.regressionStatePreserved?.residualPlot !== true ||
    dataAnalysisToolpak.regressionStatePreserved?.lineFitPlot !== true ||
    dataAnalysisToolpak.regression?.activeSheet !== "回帰結果" ||
    dataAnalysisToolpak.regression?.title !== "概要" ||
    dataAnalysisToolpak.regression?.statLabel !== "回帰統計" ||
    dataAnalysisToolpak.regression?.coefficientHeader !== "係数" ||
    dataAnalysisToolpak.regression?.coefficientFixedLowerHeader !== "下限 95.0%" ||
    dataAnalysisToolpak.regression?.coefficientFixedUpperHeader !== "上限 95.0%" ||
    dataAnalysisToolpak.regression?.residualHeader !== "残差出力" ||
    dataAnalysisToolpak.regression?.standardizedResidualHeader !== "標準化残差" ||
    dataAnalysisToolpak.regression?.chartCount < 2 ||
    !dataAnalysisToolpak.regression?.chartTitles?.includes("X 線形近似プロット") ||
    !dataAnalysisToolpak.regression?.chartTitles?.includes("X 残差プロット") ||
    !dataAnalysisToolpak.regression?.status?.includes("回帰分析") ||
    dataAnalysisToolpak.randomDialog?.title !== "乱数発生" ||
    !dataAnalysisToolpak.randomDialog?.visibleFields?.includes("乱数の行数") ||
    !dataAnalysisToolpak.randomDialog?.visibleFields?.includes("乱数の列数") ||
    dataAnalysisToolpak.random?.activeSheet !== "乱数結果" ||
    !dataAnalysisToolpak.random?.header?.includes("乱数") ||
    dataAnalysisToolpak.random?.values?.length !== 6 ||
    dataAnalysisToolpak.random?.values?.some((value) => !Number.isFinite(value) || value < 10 || value > 20) ||
    !dataAnalysisToolpak.random?.status?.includes("乱数発生") ||
    dataAnalysisToolpak.samplingPeriodic?.activeSheet !== "周期標本" ||
    dataAnalysisToolpak.samplingPeriodic?.values?.join("|") !== "周期サンプリング|標本番号|GroupA|GroupB|2|12|9|4|18|14|||" ||
    !dataAnalysisToolpak.samplingPeriodic?.status?.includes("サンプリング") ||
    dataAnalysisToolpak.samplingRandom?.activeSheet !== "ランダム標本" ||
    dataAnalysisToolpak.samplingRandom?.values?.join("|") !== "ランダム サンプリング|標本番号|GroupA|GroupB|1|12|14|2|18|9|3|20|17" ||
    !dataAnalysisToolpak.samplingRandom?.status?.includes("サンプリング") ||
    dataAnalysisToolpak.rankPercentile?.activeSheet !== "順位結果" ||
    dataAnalysisToolpak.rankPercentile?.values?.join("|") !== "順位と百分位数|点|GroupA|順位|百分率|1|10|5|0.00%|5|20|1|100.00%|点|GroupB|順位|百分率" ||
    !dataAnalysisToolpak.rankPercentile?.status?.includes("順位と百分位数") ||
    dataAnalysisToolpak.histogram?.activeSheet !== "ヒストグラム結果" ||
    dataAnalysisToolpak.histogram?.values?.join("|") !== "ビン|頻度|累積 %|10|3|30.00%|15|4|70.00%|20|3|100.00%|その他|0|100.00%" ||
    dataAnalysisToolpak.histogram?.chartCount !== 1 ||
    !dataAnalysisToolpak.histogram?.status?.includes("ヒストグラム") ||
    dataAnalysisToolpak.correlation?.activeSheet !== "相関結果" ||
    dataAnalysisToolpak.correlation?.values?.join("|") !== "|GroupA|GroupB|GroupA|1||GroupB|0.982894|1" ||
    !dataAnalysisToolpak.correlation?.status?.includes("相関") ||
    dataAnalysisToolpak.covariance?.activeSheet !== "共分散結果" ||
    dataAnalysisToolpak.covariance?.values?.join("|") !== "|GroupA|GroupB|GroupA|13.600000||GroupB|12|10.960000" ||
    !dataAnalysisToolpak.covariance?.status?.includes("共分散") ||
    dataAnalysisToolpak.correlationPairwise?.activeSheet !== "相関ペア" ||
    dataAnalysisToolpak.correlationPairwise?.values?.join("|") !== "|PairA|PairB|PairA|1||PairB|1|1" ||
    !dataAnalysisToolpak.correlationPairwise?.status?.includes("相関") ||
    dataAnalysisToolpak.covariancePairwise?.activeSheet !== "共分散ペア" ||
    dataAnalysisToolpak.covariancePairwise?.values?.join("|") !== "|PairA|PairB|PairA|2.500000||PairB|28.888889|218.750000" ||
    !dataAnalysisToolpak.covariancePairwise?.status?.includes("共分散") ||
    dataAnalysisToolpak.regressionNormalProbability?.activeSheet !== "正規確率回帰" ||
    !dataAnalysisToolpak.regressionNormalProbability?.values?.includes("正規確率出力") ||
    dataAnalysisToolpak.regressionNormalProbability?.values?.join("|") !== "概要||係数|下限 95.0%|上限 95.0%|正規確率出力|百分位|Y|10|3|90|11" ||
    dataAnalysisToolpak.regressionNormalProbability?.chartCount !== 1 ||
    !dataAnalysisToolpak.regressionNormalProbability?.status?.includes("回帰分析") ||
    dataAnalysisToolpak.histogramPareto?.activeSheet !== "パレートヒストグラム" ||
    dataAnalysisToolpak.histogramPareto?.values?.join("|") !== "ビン|頻度|累積 %|15|4|40.00%|10|3|70.00%|20|3|100.00%|その他|0|100.00%" ||
    dataAnalysisToolpak.histogramPareto?.chartCount !== 1 ||
    !dataAnalysisToolpak.histogramPareto?.status?.includes("ヒストグラム") ||
    dataAnalysisToolpak.movingAverage?.activeSheet !== "移動平均結果" ||
    !dataAnalysisToolpak.movingAverage?.dialogCaptions?.includes("区間") ||
    dataAnalysisToolpak.movingAverage?.values?.join("|") !== "期|GroupA 実績|GroupA 移動平均|GroupA 標準誤差|1|10|||2|12|||3|15|12.333333|2.666667|4|18|15|2.838231|5|20|17.666667|2.680520" ||
    dataAnalysisToolpak.movingAverage?.chartCount !== 1 ||
    !dataAnalysisToolpak.movingAverage?.status?.includes("移動平均") ||
    dataAnalysisToolpak.exponentialSmoothing?.activeSheet !== "指数平滑結果" ||
    !dataAnalysisToolpak.exponentialSmoothing?.dialogCaptions?.includes("減衰係数") ||
    dataAnalysisToolpak.exponentialSmoothing?.dialogCaptions?.includes("有意水準 α") ||
    dataAnalysisToolpak.exponentialSmoothing?.values?.join("|") !== "期|GroupA 実績|GroupA 予測値|GroupA 標準誤差|1|10|||2|12|10|2|3|15|11.600000|2.789265|4|18|14.320000|3.114611|5|20|17.264000|3.024405" ||
    dataAnalysisToolpak.exponentialSmoothing?.chartCount !== 1 ||
    !dataAnalysisToolpak.exponentialSmoothing?.status?.includes("指数平滑") ||
    dataAnalysisToolpak.fTest?.activeSheet !== "F検定結果" ||
    !dataAnalysisToolpak.fTest?.values?.join("|").startsWith("F 検定: 2 標本を使った分散の検定|GroupA|GroupB|平均|15.000000|11.800000|分散|17.000000|13.700000") ||
    !dataAnalysisToolpak.fTest?.status?.includes("F 検定") ||
    dataAnalysisToolpak.tTestEqual?.activeSheet !== "t検定結果" ||
    !dataAnalysisToolpak.tTestEqual?.values?.join("|").startsWith("t 検定: 等分散を仮定した 2 標本による検定|GroupA|GroupB|平均|15.000000|11.800000") ||
    !dataAnalysisToolpak.tTestEqual?.values?.includes("t") ||
    dataAnalysisToolpak.tTestEqual?.values?.includes("有意水準 α") ||
    !dataAnalysisToolpak.tTestEqual?.status?.includes("t 検定") ||
    dataAnalysisToolpak.zTest?.activeSheet !== "z検定結果" ||
    !dataAnalysisToolpak.zTest?.dialogCaptions?.includes("変数 1 の分散 (既知)") ||
    !dataAnalysisToolpak.zTest?.dialogCaptions?.includes("変数 2 の分散 (既知)") ||
    !dataAnalysisToolpak.zTest?.values?.join("|").startsWith("z 検定: 2 標本による平均の検定|GroupA|GroupB|平均|15.000000|11.800000") ||
    !dataAnalysisToolpak.zTest?.values?.[0]?.includes("z 検定") ||
    dataAnalysisToolpak.zTest?.values?.includes("有意水準 α") ||
    !dataAnalysisToolpak.zTest?.status?.includes("z 検定") ||
    dataAnalysisToolpak.tTestPaired?.activeSheet !== "対応t検定結果" ||
    !dataAnalysisToolpak.tTestPaired?.values?.join("|").startsWith("t 検定: 一対の標本による平均の検定|GroupA|GroupB|平均|15.000000|11.800000") ||
    !dataAnalysisToolpak.tTestPaired?.values?.includes("ピアソン相関") ||
    !dataAnalysisToolpak.tTestPaired?.values?.includes("t") ||
    dataAnalysisToolpak.tTestPaired?.values?.includes("有意水準 α") ||
    !dataAnalysisToolpak.tTestPaired?.status?.includes("t 検定") ||
    dataAnalysisToolpak.tTestUnequal?.activeSheet !== "不等分散t検定結果" ||
    !dataAnalysisToolpak.tTestUnequal?.values?.join("|").startsWith("t 検定: 分散が等しくないと仮定した 2 標本による検定|GroupA|GroupB|平均|15.000000|11.800000") ||
    !dataAnalysisToolpak.tTestUnequal?.values?.includes("t") ||
    dataAnalysisToolpak.tTestUnequal?.values?.includes("有意水準 α") ||
    !dataAnalysisToolpak.tTestUnequal?.status?.includes("t 検定") ||
    dataAnalysisToolpak.randomPattern?.activeSheet !== "乱数パターン" ||
    dataAnalysisToolpak.randomPattern?.values?.join("|") !== "乱数 (パターン)|1|3|5|1|3|5" ||
    !dataAnalysisToolpak.randomPattern?.status?.includes("乱数発生") ||
    dataAnalysisToolpak.randomDiscrete?.activeSheet !== "乱数離散" ||
    !dataAnalysisToolpak.randomDiscrete?.header?.includes("離散") ||
    !dataAnalysisToolpak.randomDiscrete?.values?.slice(1).every((value) => ["100", "200", "300"].includes(String(value))) ||
    !dataAnalysisToolpak.randomDiscrete?.status?.includes("乱数発生") ||
    dataAnalysisToolpak.randomBinomial?.activeSheet !== "乱数二項" ||
    !dataAnalysisToolpak.randomBinomial?.header?.includes("二項") ||
    !dataAnalysisToolpak.randomBinomial?.values?.slice(1).every((value) => Number.isInteger(Number(value)) && Number(value) >= 0 && Number(value) <= 4) ||
    !dataAnalysisToolpak.randomBinomial?.status?.includes("乱数発生") ||
    dataAnalysisToolpak.randomPoisson?.activeSheet !== "乱数ポアソン" ||
    !dataAnalysisToolpak.randomPoisson?.header?.includes("ポアソン") ||
    !dataAnalysisToolpak.randomPoisson?.values?.slice(1).every((value) => Number.isInteger(Number(value)) && Number(value) >= 0) ||
    !dataAnalysisToolpak.randomPoisson?.status?.includes("乱数発生") ||
    dataAnalysisToolpak.randomNormal?.activeSheet !== "乱数正規" ||
    !dataAnalysisToolpak.randomNormal?.header?.includes("正規") ||
    !dataAnalysisToolpak.randomNormal?.values?.slice(1).every((value) => Number.isFinite(Number(value))) ||
    !dataAnalysisToolpak.randomNormal?.status?.includes("乱数発生") ||
    dataAnalysisToolpak.randomBernoulli?.activeSheet !== "乱数ベルヌーイ" ||
    !dataAnalysisToolpak.randomBernoulli?.header?.includes("ベルヌーイ") ||
    !dataAnalysisToolpak.randomBernoulli?.values?.slice(1).every((value) => ["0", "1"].includes(String(value))) ||
    !dataAnalysisToolpak.randomBernoulli?.status?.includes("乱数発生") ||
    dataAnalysisToolpak.descriptive?.activeSheet !== "記述統計結果" ||
    !dataAnalysisToolpak.descriptive?.values?.join("|").startsWith("|GroupA|GroupB|平均|15|11.800000|標準誤差") ||
    !dataAnalysisToolpak.descriptive?.values?.includes("標本分散") ||
    !dataAnalysisToolpak.descriptive?.values?.includes("最大(2)") ||
    !dataAnalysisToolpak.descriptive?.values?.includes("信頼区間(95.0%)") ||
    !dataAnalysisToolpak.descriptive?.status?.includes("基本統計量") ||
    dataAnalysisToolpak.anovaOneWay?.activeSheet !== "一元分散結果" ||
    !dataAnalysisToolpak.anovaOneWay?.values?.includes("標本数") ||
    !dataAnalysisToolpak.anovaOneWay?.values?.includes("分散分析表") ||
    !dataAnalysisToolpak.anovaOneWay?.values?.includes("P-値") ||
    !dataAnalysisToolpak.anovaOneWay?.status?.includes("分散分析") ||
    dataAnalysisToolpak.anovaTwoFactor?.activeSheet !== "二元分散結果" ||
    !dataAnalysisToolpak.anovaTwoFactor?.values?.includes("装置A") ||
    !dataAnalysisToolpak.anovaTwoFactor?.values?.includes("分散分析表") ||
    !dataAnalysisToolpak.anovaTwoFactor?.values?.includes("P-値") ||
    !dataAnalysisToolpak.anovaTwoFactor?.status?.includes("繰り返しのない二元配置") ||
    dataAnalysisToolpak.anovaReplication?.activeSheet !== "繰返分散結果" ||
    !dataAnalysisToolpak.anovaReplication?.dialogCaptions?.includes("標本あたりの行数") ||
    !dataAnalysisToolpak.anovaReplication?.values?.includes("装置A") ||
    !dataAnalysisToolpak.anovaReplication?.values?.includes("データの個数") ||
    !dataAnalysisToolpak.anovaReplication?.values?.includes("分散分析表") ||
    !dataAnalysisToolpak.anovaReplication?.status?.includes("繰り返しのある二元配置") ||
    dataAnalysisToolpak.fourierReal?.activeSheet !== "フーリエ結果" ||
    dataAnalysisToolpak.fourierReal?.values?.join("|") !== "k|出力|実数部|虚数部|0|0|0|0|1|2|2|0|2|0|0|0|3|2|2|0" ||
    !dataAnalysisToolpak.fourierReal?.status?.includes("フーリエ解析") ||
    dataAnalysisToolpak.fourierComplex?.activeSheet !== "複素フーリエ" ||
    dataAnalysisToolpak.fourierComplex?.values?.join("|") !== "k|出力|実数部|虚数部|0|4|4|0|1|0|0|0|2|0|0|0|3|4i|0|4" ||
    !dataAnalysisToolpak.fourierComplex?.status?.includes("フーリエ解析") ||
    dataAnalysisToolpak.fourierInverse?.activeSheet !== "逆フーリエ" ||
    dataAnalysisToolpak.fourierInverse?.values?.join("|") !== "k|出力|実数部|虚数部|0|1+1i|1|1|1|2|2|0|2|1-1i|1|-1|3|0|0|0" ||
    !dataAnalysisToolpak.fourierInverse?.status?.includes("フーリエ解析")
  ) {
    throw new Error("Data Analysis should use the Excel-like picker/options flow and generate regression, random-number, sampling, rank/percentile, histogram, matrix, moving-average, smoothing, test, and Fourier worksheets.");
  }


  if (
    consolidateDialog.static?.headers?.join("|") !== "統合|Q1|Q2" ||
    consolidateDialog.static?.rows?.map((row) => `${row.label}:${row.q1}:${row.q2}`).join("|") !== "East:12:23|West:5:7|North:9:8" ||
    !consolidateDialog.static?.status?.includes("統合しました") ||
    consolidateDialog.linked?.initialReference !== "Consolidate!A1:C3" ||
    !consolidateDialog.linked?.references?.includes("ConsolidateExtra") ||
    !consolidateDialog.linked?.functionOptions?.includes("標準偏差P") ||
    !consolidateDialog.linked?.functionOptions?.includes("分散P") ||
    consolidateDialog.linked?.rows?.map((row) => `${row.label}:${row.q1}:${row.q2}`).join("|") !== "East:21:33|West:5:7|South:4:6" ||
    !consolidateDialog.linked?.formula?.includes("=SUM(") ||
    !consolidateDialog.linked?.formula?.includes("Consolidate!B2") ||
    !consolidateDialog.linked?.formula?.includes("'追加'!B2") ||
    !consolidateDialog.linked?.status?.includes("リンク")
  ) {
    throw new Error("Consolidate should align labels, accept sheet references, and create linked formulas like Excel.");
  }


  if (
    subtotalDialog.defaults?.groupBy !== "Group" ||
    subtotalDialog.defaults?.functionLabel !== "合計" ||
    subtotalDialog.defaults?.fields?.join("|") !== "Amount" ||
    subtotalDialog.rows?.map((row) => `${row.group}:${row.amount}`).join("|") !== "Group:Amount|A:5|A:20|A 合計:25|B:12|B:8|B 合計:20|総計:45" ||
    !subtotalDialog.status?.includes("小計を挿入") ||
    subtotalDialog.removeAll?.rows?.map((row) => `${row.group}:${row.amount}`).join("|") !== "Group:Amount|A:5|A:20|B:12|B:8" ||
    !subtotalDialog.removeAll?.status?.includes("小計を削除") ||
    subtotalDialog.summaryAbove?.rows?.map((row) => `${row.group}:${row.amount}`).join("|") !== "Group:Amount|総計:45|A 合計:25|A:5|A:20|B 合計:20|B:12|B:8" ||
    !subtotalDialog.summaryAbove?.status?.includes("小計を挿入")
  ) {
    throw new Error("Subtotal should insert, remove, and place Excel-like group subtotal rows and a grand total row.");
  }


  if (
    !outlineGrouping.groupMenuActions?.includes("outline-group") ||
    !outlineGrouping.groupMenuActions?.includes("outline-auto") ||
    !outlineGrouping.groupMenuActions?.includes("outline-settings") ||
    !outlineGrouping.ungroupMenuActions?.includes("outline-ungroup") ||
    !outlineGrouping.ungroupMenuActions?.includes("outline-clear") ||
    outlineGrouping.rowGroupDefault !== "rows" ||
    outlineGrouping.outlineRibbonAfterGroup?.show?.disabled !== false ||
    outlineGrouping.outlineRibbonAfterGroup?.hide?.disabled !== false ||
    outlineGrouping.outlineRibbonAfterGroup?.show?.active !== true ||
    outlineGrouping.outlineRibbonAfterGroup?.hide?.active !== true ||
    outlineGrouping.rowToggleBefore !== "−" ||
    outlineGrouping.rowsCollapsed?.toggleText !== "+" ||
    !outlineGrouping.rowsCollapsed?.row2 ||
    !outlineGrouping.rowsCollapsed?.row3 ||
    !outlineGrouping.rowsCollapsed?.row4 ||
    outlineGrouping.rowsCollapsed?.row5 ||
    outlineGrouping.rowsExpanded?.row2 ||
    outlineGrouping.rowsExpanded?.row3 ||
    outlineGrouping.rowsExpanded?.row4 ||
    outlineGrouping.rowUngroupDefault !== "rows" ||
    !outlineGrouping.rowUngroupNote?.includes("グループ解除") ||
    outlineGrouping.rowTogglesAfterUngroup !== 0 ||
    outlineGrouping.columnGroupDefault !== "columns" ||
    outlineGrouping.columnToggleBefore !== "−" ||
    outlineGrouping.columnsCollapsed?.toggleText !== "+" ||
    !outlineGrouping.columnsCollapsed?.col2 ||
    !outlineGrouping.columnsCollapsed?.col3 ||
    !outlineGrouping.columnsCollapsed?.col4 ||
    outlineGrouping.columnsCollapsed?.col5 ||
    outlineGrouping.columnsExpanded?.col2 ||
    outlineGrouping.columnsExpanded?.col3 ||
    outlineGrouping.columnsExpanded?.col4 ||
    outlineGrouping.outlineSettingsDefaults?.summaryBelow !== true ||
    outlineGrouping.outlineSettingsDefaults?.summaryRight !== true ||
    outlineGrouping.outlineSettingsDefaults?.autoStyle !== false ||
    outlineGrouping.summaryAboveRowToggleHost !== "1" ||
    outlineGrouping.outlineCleared !== 0 ||
    outlineGrouping.outlineRibbonAfterClear?.show?.disabled !== true ||
    outlineGrouping.outlineRibbonAfterClear?.hide?.disabled !== true ||
    outlineGrouping.autoOutlineRows < 1
  ) {
    throw new Error("Outline grouping should expose Excel-like menus, settings, auto outline, clear outline, and row/column collapse behavior.");
  }


  if (
    textToColumnsDialog.delimiterDefaults?.comma !== true ||
    textToColumnsDialog.delimiterDefaults?.tab !== false ||
    textToColumnsDialog.delimiterDefaults?.semicolon !== false ||
    textToColumnsDialog.dateOptions?.join("|") !== "MDY|DMY|YMD|MYD|DYM|YDM" ||
    !textToColumnsDialog.selectedHeader?.includes("日付(MDY)") ||
    textToColumnsDialog.overwritePrompt?.title !== "Cht WebSheet" ||
    !textToColumnsDialog.overwritePrompt?.message?.includes("既にデータ") ||
    textToColumnsDialog.overwritePrompt?.buttons?.join("|") !== "OK|キャンセル" ||
    textToColumnsDialog.afterCancel?.dialogVisible !== true ||
    textToColumnsDialog.afterCancel?.b1 !== "既存" ||
    textToColumnsDialog.afterCancel?.c2 !== "保持" ||
    textToColumnsDialog.values?.join("|") !== "2026/05/24|ZJ|100|2026/06/01|AA|200|2026/05/20|BB|300" ||
    !textToColumnsDialog.status?.includes("A1:A3 を分割") ||
    textToColumnsDialog.afterUndo?.b1 !== "既存" ||
    textToColumnsDialog.afterUndo?.c2 !== "保持" ||
    !textToColumnsDialog.afterUndo?.status?.includes("元に戻しました")
  ) {
    throw new Error("Text to Columns should expose Excel-like date order formats and convert selected date columns.");
  }


  if (
    textToColumnsFixedWidth.inferredBreaks !== "5, 8" ||
    textToColumnsFixedWidth.breakMarkersBefore !== 2 ||
    textToColumnsFixedWidth.breaksAfterRemove !== "8" ||
    textToColumnsFixedWidth.breaksAfterAdd !== "5, 8" ||
    !textToColumnsFixedWidth.fixedRowsText?.includes("ALPHA001TOKYO") ||
    !textToColumnsFixedWidth.fixedRowsText?.includes("GAMMA010OSAKA") ||
    !textToColumnsFixedWidth.previewStep3?.includes("文字列") ||
    !textToColumnsFixedWidth.previewStep3?.includes("削除") ||
    textToColumnsFixedWidth.values?.join("|") !== "ALPHA|001||GAMMA|010|" ||
    !textToColumnsFixedWidth.status?.includes("A1:A2 を分割")
  ) {
    throw new Error("Fixed-width Text to Columns should infer break positions, preserve text-formatted leading zeros, and skip omitted columns.");
  }


  if (
    textToColumnsAdvanced.delimiterDefaults?.semicolon !== true ||
    textToColumnsAdvanced.delimiterDefaults?.comma !== false ||
    textToColumnsAdvanced.display?.join("|") !== "1234.5|-123|2000|-45" ||
    Math.abs(Number(textToColumnsAdvanced.raw?.b1) - 1234.5) > 0.0001 ||
    textToColumnsAdvanced.raw?.c1 !== -123 ||
    textToColumnsAdvanced.raw?.b2 !== 2000 ||
    textToColumnsAdvanced.raw?.c2 !== -45 ||
    !textToColumnsAdvanced.status?.includes("A1:A2 を分割")
  ) {
    throw new Error("Text to Columns advanced settings should honor decimal/thousands separators and trailing minus signs like Excel.");
  }


  if (
    sortDialogLevels.initialCount !== 1 ||
    sortDialogLevels.afterAddCount !== 2 ||
    sortDialogLevels.afterCopyCount !== 3 ||
    sortDialogLevels.afterDeleteCount !== 2 ||
    !sortDialogLevels.toolbarLabels?.some((label) => label.includes("レベルの追加")) ||
    sortDialogLevels.rows?.map((row) => `${row.group}:${row.score}`).join("|") !== "A:20|A:5|B:12|B:8" ||
    sortDialogLevels.reopened?.levelCount !== 2 ||
    sortDialogLevels.reopened?.keys?.join("|") !== "1|2" ||
    sortDialogLevels.reopened?.orders?.join("|") !== "asc|desc" ||
    sortDialogLevels.reopened?.header !== true ||
    sortDialogLevels.reopened?.orientation !== "rows"
  ) {
    throw new Error("Sort dialog should support Excel-like level add/copy/delete, multi-key sorting, and reopening with the saved sort state.");
  }


  if (
    !sortDialogColorOrder.colorOptions?.includes("#ff0000") ||
    sortDialogColorOrder.position !== "top" ||
    sortDialogColorOrder.rows?.[0]?.score !== "8" ||
    !sortDialogColorOrder.colorStatus?.includes("並べ替え") ||
    !sortDialogColorOrder.colorExported?.cellColorSortCount ||
    !sortDialogColorOrder.colorExported?.dxfSortCount ||
    sortDialogColorOrder.colorExported?.inlineColorSortCount !== 0 ||
    !sortDialogColorOrder.fontColorOptions?.includes("#ff0000") ||
    sortDialogColorOrder.fontPosition !== "top" ||
    sortDialogColorOrder.fontRows?.[0]?.score !== "5" ||
    !sortDialogColorOrder.fontStatus?.includes("並べ替え") ||
    !sortDialogColorOrder.fontExported?.fontColorSortCount ||
    !sortDialogColorOrder.fontExported?.dxfSortCount ||
    sortDialogColorOrder.fontExported?.inlineColorSortCount !== 0
  ) {
    throw new Error("Sort dialog should support Excel-like cell-color and font-color ordering.");
  }


  if (
    sortDialogLeftToRight.optionState?.axisCaption !== "行" ||
    !sortDialogLeftToRight.optionState?.headerLabel?.includes("先頭列") ||
    sortDialogLeftToRight.optionState?.keys?.join("|") !== "Field|Score|Code" ||
    sortDialogLeftToRight.rows?.map((row) => row.join(":")).join("|") !== "Field:East:North:West|Score:10:20:30|Code:E:N:W" ||
    !sortDialogLeftToRight.status?.includes("左から右")
  ) {
    throw new Error("Sort dialog should support Excel-like left-to-right column sorting with first-column headers.");
  }


  if (
    sortDialogBlankAndCase.blankAsc?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "Low:5|Mid:10|High:20|Blank:" ||
    sortDialogBlankAndCase.blankDesc?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "High:20|Mid:10|Low:5|Blank:" ||
    sortDialogBlankAndCase.caseSensitive?.rows?.map((row) => `${row.key}:${row.value}`).join("|") !== "a:2|A:1|b:4|B:3"
  ) {
    throw new Error("Sort dialog should keep blank cells at the bottom and honor case-sensitive ordering like Excel.");
  }


  if (
    !sortDialogCustomList.months?.orderOptions?.some((option) => option.value === "custom:month-ja" && option.label.includes("1月")) ||
    !sortDialogCustomList.months?.orderOptions?.some((option) => option.value === "custom:weekday-ja" && option.label.includes("日曜日")) ||
    sortDialogCustomList.months?.rows?.map((row) => `${row.month}:${row.value}`).join("|") !== "1月:M1|2月:M2|3月:M3|12月:M12" ||
    sortDialogCustomList.weekdays?.rows?.map((row) => `${row.weekday}:${row.code}`).join("|") !== "日:S|月:M|金:F|土:Sat" ||
    !sortDialogCustomList.months?.status?.includes("並べ替え") ||
    !sortDialogCustomList.weekdays?.status?.includes("並べ替え") ||
    sortDialogCustomList.months?.exported?.customListTexts?.[0] !== "1月,2月,3月,4月,5月,6月,7月,8月,9月,10月,11月,12月" ||
    sortDialogCustomList.weekdays?.exported?.customListTexts?.[0] !== "日曜日,月曜日,火曜日,水曜日,木曜日,金曜日,土曜日"
  ) {
    throw new Error("Sort dialog should support Excel-like custom list ordering for Japanese months and weekdays.");
  }


  if (
    sortDialogDateOrder.rows?.map((row) => `${row.date}:${row.label}`).join("|") !== "2025/12/31:Prev|2026/05/20:May20|05/23/2026:May23|06/01/2026:Jun|:Blank" ||
    !sortDialogDateOrder.status?.includes("並べ替え") ||
    sortDialogDateOrder.rawDateRows?.map((row) => `${row.date}:${row.label}`).join("|") !== "2025/12/31:Prev|2026/05/20:May20|06/01/2026:Jun|:Blank" ||
    !sortDialogDateOrder.rawDateStatus?.includes("並べ替え")
  ) {
    throw new Error("Sort dialog should sort date-like strings and real date values by chronological date instead of plain text.");
  }


  if (
    sortFormulaResults.rows?.map((row) => `${row.name}:${row.calc}`).join("|") !== "One:1|Three:3|TenA:10|TenB:10" ||
    !sortFormulaResults.status?.includes("並べ替え") ||
    sortFormulaResults.quickRows?.map((row) => `${row.name}:${row.calc}`).join("|") !== "TenA:10|TenB:10|Three:3|One:1" ||
    !sortFormulaResults.quickStatus?.includes("並べ替え")
  ) {
    throw new Error("Sort dialog and quick sort should order formula cells by their calculated displayed values like Excel.");
  }


  if (
    quickSortActiveColumn.rows?.map((row) => `${row.group}:${row.score}`).join("|") !== "A:20|B:12|B:8|A:5" ||
    quickSortActiveColumn.selectionLabel !== "B3" ||
    !quickSortActiveColumn.status?.includes("A1:B5")
  ) {
    throw new Error("Quick sort should expand a single active cell to the current data region and sort by the active column.");
  }


  if (
    quickSortPartialSelection.rows?.map((row) => `${row.group}:${row.score}`).join("|") !== "A:20|B:12|B:8|A:5" ||
    !quickSortPartialSelection.status?.includes("A1:B5")
  ) {
    throw new Error("Quick sort should expand a partial column selection to the surrounding current region instead of scrambling adjacent data.");
  }


  if (
    removeDuplicatesDialog.labelsWithHeader?.join("|") !== "Name|Score" ||
    removeDuplicatesDialog.labelsWithoutHeader?.join("|") !== "A|B" ||
    removeDuplicatesDialog.noColumnTitle !== "Cht WebSheet" ||
    !removeDuplicatesDialog.noColumnMessage?.includes("列を1つ以上選択") ||
    !removeDuplicatesDialog.noColumnDialogStillOpen ||
    removeDuplicatesDialog.uncheckedCount !== 2 ||
    removeDuplicatesDialog.recheckedCount !== 2 ||
    removeDuplicatesDialog.title !== "Cht WebSheet" ||
    !removeDuplicatesDialog.message?.includes("重複する値が 2 個") ||
    !removeDuplicatesDialog.message?.includes("一意の値が 3 個") ||
    removeDuplicatesDialog.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "A:1|B:2|C:3|:|:" ||
    removeDuplicatesDialog.afterUndo?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "A:1|B:2|A:1|C:3|a:1" ||
    !removeDuplicatesDialog.afterUndo?.status?.includes("元に戻しました") ||
    removeDuplicatesDialog.formulaDisplayDuplicates?.title !== "Cht WebSheet" ||
    !removeDuplicatesDialog.formulaDisplayDuplicates?.message?.includes("重複する値が 1 個") ||
    removeDuplicatesDialog.formulaDisplayDuplicates?.rows?.map((row) => `${row.calc}:${row.group}`).join("|") !== "2:Same|5:Other|:" ||
    !removeDuplicatesDialog.formulaDisplayDuplicates?.status?.includes("重複") ||
    removeDuplicatesDialog.dateValueDuplicates?.title !== "Cht WebSheet" ||
    !removeDuplicatesDialog.dateValueDuplicates?.message?.includes("重複する値が 1 個") ||
    removeDuplicatesDialog.dateValueDuplicates?.rows?.map((row) => `${row.date}:${row.group}`).join("|") !== "2026/05/20:Same|2026/05/21:Other|:" ||
    !removeDuplicatesDialog.dateValueDuplicates?.status?.includes("重複")
  ) {
    throw new Error("Remove Duplicates should update column labels, compare calculated/date values, ignore case like Excel, and show Excel-like result message.");
  }


  if (
    persistenceReload.a1 !== "ReloadKeep" ||
    !persistenceReload.workbookName?.includes(expectedReloadWorkbookName) ||
    persistenceReload.title !== expectedReloadWorkbookName ||
    persistenceReload.exportDisabled ||
    !persistenceReload.workspaceVisible ||
    !persistenceReload.status?.includes("復元") ||
    persistenceReload.selectionLabel !== "A1" ||
    persistenceReload.scrollTop !== 0 ||
    persistenceReload.scrollLeft !== 0 ||
    !persistenceReload.firstVisibleRows.includes("1") ||
    persistenceReload.dateC50 !== "2005/2" ||
    /GMT|標準時|Standard Time/.test(String(persistenceReload.dateC50 || "")) ||
    !Number.isFinite(Number(persistenceReload.dateCellModel?.raw)) ||
    (persistenceReload.dateCellModel?.cached != null && !Number.isFinite(Number(persistenceReload.dateCellModel.cached))) ||
    /GMT|標準時|Standard Time/.test(String(persistenceReload.dateCellModel?.cached || "")) ||
    /GMT|標準時|Standard Time/.test(String(persistenceReload.dateCellModel?.display || "")) ||
    persistenceReload.dateTimeSamples?.hms !== "9:05:07" ||
    persistenceReload.dateTimeSamples?.ampm !== "9:05 PM" ||
    persistenceReload.dateTimeSamples?.dateTime !== "2005/2/1 9:05:17"
  ) {
    throw new Error("Reload should restore the workbook contents while resetting selection and viewport to A1.");
  }


  if (
    wideViewportCoverage.viewportWidth !== 2048 ||
    wideViewportCoverage.gridRight < wideViewportCoverage.hostRight - 1 ||
    wideViewportCoverage.lastHeaderRight < wideViewportCoverage.hostRight - 1 ||
    wideViewportCoverage.lastHeaderCol <= 42
  ) {
    throw new Error("Reloaded wide view should extend lazy columns far enough to cover the visible sheet area past AP.");
  }


  if (!result.canvasLayer.exists || result.canvasLayer.width <= 0 || result.canvasLayer.height <= 0 || result.canvasLayer.position !== "absolute" || result.canvasLayer.pointerEvents !== "none") {
    throw new Error("Sheet cells should be painted by a non-interactive canvas layer below the DOM overlay.");
  }


  if (coordinateHitTest.label !== "G6" || coordinateHitTest.activeAddress !== "G6" || coordinateHitTest.eventTargetDisabled !== "none") {
    throw new Error("Cell hit-testing should resolve from sheet coordinates, even when DOM cell pointer events are disabled.");
  }


  if (result.activeRibbonTab !== "ホーム" || !result.homePanelVisible || !result.fileSideMenuHidden || result.fileRibbonPanelExists) {
    throw new Error("The ribbon should default to Home, with File represented as a hidden side menu instead of a toolbar panel.");
  }


  if (
    result.ribbonHomeLayout.groupCount < 6 ||
    ["クリップボード", "フォント", "配置", "数値", "セル", "編集"].some((label) => !result.ribbonHomeLayout.groupLabels.includes(label)) ||
    result.ribbonHomeLayout.clipboardStackCount !== 3 ||
    !result.ribbonHomeLayout.pasteButton ||
    !result.ribbonHomeLayout.cutButton ||
    result.ribbonHomeLayout.pasteButton.height < result.ribbonHomeLayout.cutButton.height * 2 ||
    result.ribbonHomeLayout.fontRows !== 2 ||
    result.ribbonHomeLayout.alignmentIconCount !== 6 ||
    result.ribbonHomeLayout.alignmentActionCommands.join("|") !== "wrap-text|merge-center" ||
    !result.ribbonHomeLayout.mergeInsideAlignment ||
    !result.ribbonHomeLayout.wrapInsideAlignment ||
    result.ribbonHomeLayout.alignJustifyVisible ||
    !result.ribbonHomeLayout.numberFormatVisible ||
    !result.ribbonHomeLayout.compactFontFamily?.includes("Arial Narrow") ||
    result.ribbonHomeLayout.tabFontSize !== "11px" ||
    result.ribbonHomeLayout.toolFontSize !== "10px" ||
    result.ribbonHomeLayout.selectFontSize !== "10px" ||
    (result.ribbonHomeLayout.ribbonTabRow?.height || 999) > 23 ||
    (result.ribbonHomeLayout.ribbonPanels?.height || 999) > 72 ||
    (result.ribbonHomeLayout.fontField?.width || 999) > 118 ||
    (result.ribbonHomeLayout.sizeField?.width || 999) > 46 ||
    (result.ribbonHomeLayout.groupWidths?.clipboard || 0) < 192 ||
    (result.ribbonHomeLayout.groupWidths?.clipboard || 999) > 204 ||
    (result.ribbonHomeLayout.groupWidths?.font || 999) > 214 ||
    (result.ribbonHomeLayout.groupWidths?.alignment || 999) > 220 ||
    (result.ribbonHomeLayout.groupWidths?.number || 999) > 152 ||
    !result.ribbonHomeLayout.mergeIconClass?.includes("table-cells-merge") ||
    result.ribbonHomeLayout.borderSelectExists ||
    !result.ribbonHomeLayout.borderSplitVisible ||
    !result.ribbonHomeLayout.underlineSplitVisible ||
    !["挿入", "削除", "書式"].every((label) => result.ribbonHomeLayout.cellsCommands.includes(label)) ||
    !["オートSUM", "フィル", "クリア", "並べ替えとフィルター", "検索と選択"].every((label) => result.ribbonHomeLayout.editingCommands.includes(label))
  ) {
    throw new Error("Home ribbon layout should use compact fonts and Excel-like grouped rows with icon split controls.");
  }


  if (
    insertRibbon.activeRibbonTab !== "挿入" ||
    !insertRibbon.panelVisible ||
    ["テーブル", "図", "アドイン", "グラフ", "スパークライン", "フィルター", "リンク", "テキスト", "記号と特殊文字"].some(
      (label) => !insertRibbon.groupLabels.includes(label),
    ) ||
    [
      "insert-pivot-table",
      "insert-recommended-pivot",
      "insert-table",
      "insert-picture",
      "insert-icons",
      "insert-3d-model",
      "insert-smartart",
      "insert-screenshot",
      "insert-get-addins",
      "insert-my-addins",
      "insert-shape",
      "insert-recommended-chart",
      "insert-column-chart",
      "insert-line-chart",
      "insert-pie-chart",
      "insert-scatter-chart",
      "insert-map-chart",
      "insert-pivot-chart",
      "insert-sparkline-line",
      "insert-sparkline-column",
      "insert-sparkline-winloss",
      "insert-slicer",
      "insert-timeline",
      "insert-link",
      "insert-text-box",
      "insert-header-footer",
      "insert-wordart",
      "insert-signature-line",
      "insert-object",
      "insert-equation",
      "insert-symbol",
    ].some(
      (command) => !insertRibbon.enabledCommands.includes(command),
    ) ||
    insertRibbon.disabledCommandCount > 9 ||
    insertRibbon.largeButtonCount < 8 ||
    insertRibbon.compactButtonCount < 12 ||
    insertRibbon.chartIconCount !== 6 ||
    insertRibbon.missingIcons.length ||
    insertRibbon.panelHeight > 72
  ) {
    throw new Error("Insert ribbon should expose Excel-like groups with real enabled actions and disabled placeholders for unsupported object features.");
  }


  const pageBreakContext = viewWindowCommands.pageBreakContextMenu || {};


  const pageBreakContextAxis = pageBreakContext.target?.axis === "row" ? "row" : "col";


  const pageBreakContextIndex = pageBreakContext.target?.index;


  const pageBreakContextUndoList = pageBreakContextAxis === "row"
    ? pageBreakContext.undo?.pageBreaks?.rows || []
    : pageBreakContext.undo?.pageBreaks?.cols || [];


  const pageBreakContextRedoList = pageBreakContextAxis === "row"
    ? pageBreakContext.redo?.pageBreaks?.rows || []
    : pageBreakContext.redo?.pageBreaks?.cols || [];


  if (
    formulaRibbon.activeRibbonTab !== "数式" ||
    !formulaRibbon.panelVisible ||
    ["関数ライブラリ", "定義された名前", "ワークシート分析", "計算方法"].some((label) => !formulaRibbon.groupLabels.includes(label)) ||
    [
      "formula-insert-function",
      "formula-autosum",
      "formula-function-recent",
      "formula-function-financial",
      "formula-function-if",
      "formula-function-more",
      "formula-name-manager",
      "formula-use-in-formula",
      "formula-paste-name",
      "formula-create-from-selection",
      "formula-trace-precedents",
      "formula-show-formulas",
      "formula-watch-window",
      "formula-calculate-now",
    ].some((command) => !formulaRibbon.enabledCommands.includes(command)) ||
    formulaRibbon.missingIcons.length ||
    formulaRibbon.panelHeight > 72 ||
    formulaDefineNameMain.title !== "新しい名前" ||
    !/^=CHT!\$[A-Z]+\$1:\$[A-Z]+\$3$/.test(formulaDefineNameMain.refersTo) ||
    formulaDefineNameMenu.label !== "名前の定義" ||
    !formulaDefineNameMenu.accessKeys?.includes("C") ||
    formulaDefineNameMenu.seedBefore !== "FormulaSeed" ||
    !formulaDefineNameMenu.actions.some((item) => item.action === "define" && item.text.includes("名前の定義")) ||
    !formulaDefineNameMenu.actions.some((item) => item.action === "manage" && item.text.includes("名前の管理")) ||
    !formulaDefineNameMenu.actions.some((item) => item.action === "create-from-selection" && item.text.includes("選択範囲から作成")) ||
    formulaDefineNameCreate.title !== "選択範囲から名前を作成" ||
    formulaDefineNameCreate.seedAfterShortcut !== "FormulaSeed" ||
    !formulaDefineNameCreate.menuClosed ||
    createNamesDialog.title !== "選択範囲から名前を作成" ||
    !createNamesDialog.checks.some((item) => item.source === "top" && item.checked) ||
    !formulaNames.names.some((definition) => definition.name === "FormulaSeed" && definition.refersTo.includes("$H$2:$H$3")) ||
    formulaNameManager.initial.title !== "名前の管理" ||
    !["名前(N)", "値", "参照範囲", "範囲", "コメント"].every((header) => formulaNameManager.initial.headers.includes(header)) ||
    !formulaNameManager.initial.filterOptions.includes("現在のワークシート") ||
    !formulaNameManager.initial.filterOptions.includes("エラーがある名前") ||
    !formulaNameManager.initial.rows.some((row) => row.includes("FormulaSeed")) ||
    !formulaNameManager.initial.hasEdit ||
    !formulaNameManager.currentSheet.rows.some((row) => row.includes("LocalMetric") && row.includes("ローカル名")) ||
    !formulaNameManager.errors.rows.some((row) => row.includes("BrokenName") && row.includes("#REF!")) ||
    formulaNameManager.edit.comment !== "Name Manager から編集" ||
    !formulaNameManager.edit.managerOpen ||
    !formulaNameManager.newName.exists ||
    !formulaNameManager.newName.rows.some((row) => row.includes("ManagerNewName")) ||
    formulaNameManager.delete.brokenExists ||
    !formulaNameManager.delete.managerOpen ||
    !formulaNameManager.undo.brokenExists ||
    formulaNameManager.redo.brokenExists ||
    !formulaFunctionDialog.functions.includes("IF") ||
    !formulaFunctionDialog.functions.includes("SUMPRODUCT") ||
    !formulaFunctionDialog.functions.includes("NETWORKDAYS.INTL") ||
    formulaFunctionDialog.functionCount < 350 ||
    !formulaFunctionDialog.visibleFunctions.includes("SUM") ||
    formulaUseNameMenu.label !== "数式で使用" ||
    !["名前", "値", "範囲", "参照範囲"].every((header) => formulaUseNameMenu.headers.includes(header)) ||
    !formulaUseNameMenu.names.includes("FormulaSeed") ||
    !formulaUseNameMenu.rows.some((row) => row.name === "Metric" && row.scope === "ブック" && row.reference.includes("$I$2:$I$3")) ||
    !formulaUseNameMenu.rows.some((row) => row.name === "LocalMetric" && row.scope === "CHT" && row.reference.includes("$H$2")) ||
    !formulaUseNameMenu.rows.some((row) => row.name === "UseMenuBroken" && row.reference === "#REF!" && row.hasError) ||
    !formulaUseNameMenu.hasPasteName ||
    !formulaUseNameMenu.hasManage ||
    formulaUseNameManage.title !== "名前の管理" ||
    !formulaUseNameManage.hasUseMenuBroken ||
    formulaUseNameInserted.formula !== "=FormulaSeed" ||
    formulaPasteNameDialog.title !== "名前の貼り付け" ||
    !["名前", "参照範囲"].every((header) => formulaPasteNameDialog.headers.includes(header)) ||
    !formulaPasteNameDialog.hasMetric ||
    !formulaPasteNameDialog.reference.includes("$I$2:$I$3") ||
    formulaPasteNameInserted.formula !== "=Metric" ||
    formulaPasteNameListDialog.title !== "名前の貼り付け" ||
    !formulaPasteNameListDialog.hasPasteList ||
    formulaPasteNameListDialog.optionCount < 4 ||
    !formulaPasteNameList.selection.includes("CD210") ||
    !formulaPasteNameList.status.includes("名前一覧") ||
    !formulaPasteNameList.names.some((row) =>
      row.name === "FormulaSeed" &&
      row.referenceDisplay.includes("$H$2:$H$3") &&
      !row.referenceFormula &&
      String(row.referenceRaw).startsWith("'=")
    ) ||
    !formulaPasteNameList.names.some((row) =>
      row.name === "Metric" &&
      row.referenceDisplay.includes("$I$2:$I$3") &&
      !row.referenceFormula &&
      String(row.referenceRaw).startsWith("'=")
    ) ||
    formulaPasteNameListUndo.firstName !== "" ||
    formulaPasteNameListRedo.firstName !== "FormulaSeed" ||
    !formulaFunctionDialog.searchLabel.includes("関数の検索") ||
    !formulaFunctionDialog.description.includes("SUM(number1") ||
    !formulaFunctionDialog.categories.includes("論理") ||
    !formulaFunctionDialog.categories.includes("互換性") ||
    !formulaFunctionDialog.functions.includes("NORMDIST") ||
    !formulaFunctionDialog.functions.includes("BETADIST") ||
    !formulaFunctionDialog.categories.includes("すべて表示") ||
    !formulaFunctionSearch.visibleFunctions.includes("SUMPRODUCT") ||
    formulaFunctionSearch.active !== "SUMPRODUCT" ||
    !formulaFunctionSearch.description.includes("SUMPRODUCT(array1") ||
    formulaFunctionArguments.initial.title !== "関数の引数" ||
    formulaFunctionArguments.initial.functionName !== "IF" ||
    !["logical_test", "value_if_true", "value_if_false"].every((label) => formulaFunctionArguments.initial.labels.includes(label)) ||
    formulaFunctionArguments.initial.preview !== "=IF(1=1,1,0)" ||
    !formulaFunctionArguments.initial.argumentResults.includes("= TRUE") ||
    formulaFunctionArguments.initial.result !== "1" ||
    !formulaFunctionArguments.initial.description.includes("logical_test") ||
    formulaFunctionArguments.initial.referenceButtonCount < 3 ||
    !formulaFunctionArguments.references.pickerCollapsed.collapsed ||
    formulaFunctionArguments.references.pickerCollapsed.visibleRows.join("|") !== "1" ||
    !formulaFunctionArguments.references.pickerCollapsed.actionsHidden ||
    formulaFunctionArguments.references.pickerCollapsed.buttonText !== "↙" ||
    formulaFunctionArguments.references.pickerExpanded.collapsed ||
    formulaFunctionArguments.references.pickerExpanded.visibleRows.length < 3 ||
    formulaFunctionArguments.references.pickerExpanded.activeIndex !== "1" ||
    formulaFunctionArguments.references.values[0] !== "A4" ||
    formulaFunctionArguments.references.values[1] !== "A4:A5" ||
    !formulaFunctionArguments.references.values[2]?.includes("!A4") ||
    !formulaFunctionArguments.references.preview.startsWith("=IF(A4,A4:A5,") ||
    !formulaFunctionArguments.references.preview.endsWith("!A4)") ||
    formulaFunctionArguments.back.title !== "関数の挿入" ||
    formulaFunctionArguments.back.selected !== "IF" ||
    formulaFunctionArguments.back.category !== "論理" ||
    formulaFunctionArguments.edited.preview !== "=IF(1=0,1,0)" ||
    !formulaFunctionArguments.edited.argumentResults.includes("= FALSE") ||
    formulaFunctionArguments.edited.result !== "0" ||
    formulaInserted.formula !== "=IF(1=0,1,0)" ||
    formulaInserted.display !== "0" ||
    formulaInserted.modelDisplay !== "" ||
    (formulaInserted.modelCached !== null && formulaInserted.modelCached !== 0) ||
    formulaLogicalMenu.label !== "論理" ||
    !formulaLogicalMenu.functions.includes("IFERROR") ||
    !formulaLogicalMenu.functions.includes("SWITCH") ||
    !formulaLogicalMenu.syntaxText.includes("IF(logical_test") ||
    formulaMenuArguments.title !== "関数の引数" ||
    formulaMenuArguments.functionName !== "AND" ||
    !formulaMenuArguments.labels.includes("logical1") ||
    formulaMenuArguments.preview !== "=AND(1=1)" ||
    !formulaMenuArguments.argumentResults.includes("= TRUE") ||
    formulaMenuArguments.result !== "TRUE" ||
    formulaMenuInserted.formula !== "=AND(1=1)" ||
    formulaMenuInserted.display !== "TRUE" ||
    formulaMenuInserted.modelDisplay !== "" ||
    (formulaMenuInserted.modelCached !== null && formulaMenuInserted.modelCached !== true) ||
    formulaRecentMenu.label !== "最近使用した関数" ||
    formulaRecentMenu.functions[0] !== "AND" ||
    !formulaRecentMenu.functions.includes("IF") ||
    formulaMathMenu.label !== "数学/三角" ||
    !["CEILING.MATH", "COMBIN", "SERIESSUM", "CSCH", "ATAN2"].every((name) => formulaMathMenu.functions.includes(name)) ||
    !formulaMathMenu.hasCategoryDialog ||
    formulaVariableArguments.initial.title !== "関数の引数" ||
    formulaVariableArguments.initial.functionName !== "SUM" ||
    !["number1", "number2"].every((label) => formulaVariableArguments.initial.labels.includes(label)) ||
    !formulaVariableArguments.expanded.labels.includes("number3") ||
    formulaVariableArguments.expanded.activeIndex !== "2" ||
    formulaVariableArguments.expanded.preview !== "=SUM(2,3)" ||
    formulaVariableArguments.expanded.result !== "5" ||
    !formulaVariableArguments.expanded.argumentResults.includes("= 2") ||
    !formulaVariableArguments.expanded.argumentResults.includes("= 3") ||
    !formulaVariableArguments.expanded.description.includes("number3") ||
    formulaVariableArguments.inserted.formula !== "=SUM(2,3)" ||
    formulaVariableArguments.inserted.display !== "5" ||
    formulaVariableArguments.inserted.modelDisplay !== "" ||
    (formulaVariableArguments.inserted.modelCached !== null && formulaVariableArguments.inserted.modelCached !== 5) ||
    formulaZeroArgumentFunction.dialog.title !== "関数の引数" ||
    formulaZeroArgumentFunction.dialog.functionName !== "TRUE" ||
    !formulaZeroArgumentFunction.dialog.emptyText.includes("引数はありません") ||
    formulaZeroArgumentFunction.dialog.labels.length !== 0 ||
    formulaZeroArgumentFunction.dialog.preview !== "=TRUE()" ||
    formulaZeroArgumentFunction.dialog.result !== "TRUE" ||
    formulaZeroArgumentFunction.inserted.formula !== "=TRUE()" ||
    formulaZeroArgumentFunction.inserted.display !== "TRUE" ||
    formulaZeroArgumentFunction.inserted.modelDisplay !== "" ||
    (formulaZeroArgumentFunction.inserted.modelCached !== null && formulaZeroArgumentFunction.inserted.modelCached !== true) ||
    formulaDateMenu.label !== "日付/時刻" ||
    !["DATEVALUE", "TIMEVALUE", "ISOWEEKNUM", "DAYS", "DAYS360"].every((name) => formulaDateMenu.functions.includes(name)) ||
    !formulaDateMenu.hasCategoryDialog ||
    formulaTextMenu.label !== "文字列" ||
    !["CHAR", "CODE", "CLEAN", "REPT", "NUMBERVALUE", "UNICODE"].every((name) => formulaTextMenu.functions.includes(name)) ||
    !formulaTextMenu.hasCategoryDialog ||
    formulaFinancialMenu.label !== "財務" ||
    !formulaFinancialMenu.functions.includes("PMT") ||
    !formulaFinancialMenu.functions.includes("IRR") ||
    !formulaFinancialMenu.hasCategoryDialog ||
    formulaMoreMenu.label !== "その他の関数" ||
    !formulaMoreMenu.functions.includes("COUNTBLANK") ||
    !formulaMoreMenu.functions.includes("IMABS") ||
    !formulaMoreMenu.functions.includes("NORMDIST") ||
    !formulaMoreMenu.functions.includes("BETADIST") ||
    !["DSUM", "IMCSCH", "WEIBULL.DIST", "BETA.INV", "ERROR.TYPE"].every((name) => formulaMoreMenu.functions.includes(name)) ||
    !formulaMoreMenu.hasAllFunctionsDialog ||
    formulaAutoSumMenu.label !== "オートSUM" ||
    !formulaAutoSumMenu.functions.includes("AVERAGE") ||
    !formulaAutoSumMenu.hasMoreFunctions ||
    formulaAutoSumInserted.formula !== "=AVERAGE(CB200:CB201)" ||
    formulaShowState.showFormulas !== true ||
    formulaShowState.buttonActive !== true ||
    formulaTrace.precedents.length < 1 ||
    formulaTrace.highlighted < 1 ||
    formulaTrace.arrows.length < 1 ||
    formulaTrace.overlayArrows < 1 ||
    formulaRemoveArrowsMenu.label !== "トレース矢印の削除" ||
    !formulaRemoveArrowsMenu.actions.includes("remove-precedent-arrows") ||
    formulaRemovePrecedents.precedents.length !== 0 ||
    formulaRemovePrecedents.arrows.length !== 0 ||
    formulaTraceDependents.dependents.length < 1 ||
    formulaTraceDependents.arrows.length < 1 ||
    formulaTraceDependents.overlayArrows < 1 ||
    formulaRemoveDependents.dependents.length !== 0 ||
    formulaRemoveDependents.arrows.some((arrow) => arrow.kind === "dependents") ||
    formulaEvaluateInitial.title !== "数式の検証" ||
    !formulaEvaluateInitial.location.includes("H7") ||
    !formulaEvaluateInitial.expression.includes("=ROUND(SUM(H2:H3),2)") ||
    formulaEvaluateInitial.stepCount < 3 ||
    formulaEvaluateInitial.stepInDisabled !== false ||
    formulaEvaluateInitial.stepOutDisabled !== true ||
    !formulaEvaluateStepIn.location.includes("SUM にステップ イン") ||
    !formulaEvaluateStepIn.expression.includes("=SUM(H2:H3)") ||
    formulaEvaluateStepIn.stepInDisabled !== true ||
    formulaEvaluateStepIn.stepOutDisabled !== false ||
    !formulaEvaluateStep.activeStep.includes("H2:H3") ||
    !formulaEvaluateStep.expression.includes("{") ||
    formulaEvaluateStep.stepInDisabled !== true ||
    formulaEvaluateStep.stepOutDisabled !== false ||
    !formulaEvaluateStepOut.expression.includes("ROUND(SUM") ||
    formulaEvaluateStepOut.stepOutDisabled !== true ||
    !formulaEvaluateReferenceInitial.location.includes("CB205") ||
    !formulaEvaluateReferenceInitial.expression.includes("=CB204+1") ||
    formulaEvaluateReferenceInitial.stepInDisabled !== true ||
    formulaEvaluateReferenceInitial.stepOutDisabled !== true ||
    !formulaEvaluateReferenceStep.activeStep.includes("CB204") ||
    !formulaEvaluateReferenceStep.expression.includes("15") ||
    formulaEvaluateReferenceStep.stepInDisabled !== false ||
    formulaEvaluateReferenceStep.stepOutDisabled !== true ||
    !formulaEvaluateReferenceStepIn.location.includes("CB204 にステップ イン") ||
    !formulaEvaluateReferenceStepIn.expression.includes("=10+5") ||
    !formulaEvaluateReferenceStepIn.result.includes("15") ||
    formulaEvaluateReferenceStepIn.stepOutDisabled !== false ||
    !formulaEvaluateReferenceStepOut.location.includes("CB205") ||
    !formulaEvaluateReferenceStepOut.expression.includes("15+1") ||
    !formulaEvaluateReferenceStepOut.activeStep.includes("CB204") ||
    formulaEvaluateReferenceStepOut.stepOutDisabled !== true ||
    !formulaWatchWindow.open ||
    !formulaWatchWindow.title.includes("ウォッチ ウィンドウ") ||
    !formulaWatchWindow.rows.some((row) => row.includes("CB203") && row.includes("=SUM(H2:H3")) ||
    !formulaWatchWindow.watches.some((watch) => watch.row === 203 && watch.col === 80) ||
    !formulaWatchSelect.selected.includes("CB203") ||
    formulaWatchSelect.formula !== "=SUM(H2:H3)" ||
    formulaWatchDelete.rows !== 0 ||
    !formulaWatchDelete.emptyText.includes("ウォッチ式はありません") ||
    formulaErrorMenu.label !== "エラー チェック" ||
    !formulaErrorMenu.actions.includes("trace-errors") ||
    !formulaErrorMenu.actions.includes("reset-ignored-errors") ||
    !formulaErrorMenu.circularItems.some((item) => item.text.includes("CB4") || item.text.includes("CB5")) ||
    !formulaCircularSelection.selected.match(/CB[45]/) ||
    formulaCircularSelection.cell4 !== "=CB5" ||
    formulaCircularSelection.cell5 !== "=CB4" ||
    formulaErrorChecking.title !== "エラー チェック" ||
    !formulaErrorChecking.location.includes("CB1") ||
    !formulaErrorChecking.value.startsWith("#DIV/0!") ||
    !formulaErrorChecking.formula.includes("=1/0") ||
    !formulaErrorChecking.entries.some((entry) => entry.includes("CB1")) ||
    !formulaErrorIgnore.ignored.some((entry) => entry.includes("1:80:#DIV/0!")) ||
    formulaErrorTrace.arrows.length < 1 ||
    formulaErrorTrace.overlayArrows < 1 ||
    formulaErrorResetIgnored.ignored.length !== 0 ||
    formulaCalculation.menuInitial.label !== "計算方法の設定" ||
    !formulaCalculation.menuInitial.actions.includes("manual") ||
    !formulaCalculation.menuInitial.actions.includes("except-tables") ||
    formulaCalculation.manual.calculationMode !== "manual" ||
    formulaCalculation.manual.calcProperties?.calcMode !== "manual" ||
    !formulaCalculation.manual.status.includes("手動") ||
    !formulaCalculation.menuManual.checked.includes("manual") ||
    formulaCalculation.exceptTables.calculationMode !== "except-tables" ||
    formulaCalculation.exceptTables.calcProperties?.calcMode !== "autoNoTable" ||
    formulaCalculation.automatic.calculationMode !== "automatic" ||
    formulaCalculation.automatic.calcProperties?.calcMode !== "auto" ||
    !formulaCalculation.status.includes("再計算") ||
    formulaCalculation.calculationMode !== "automatic" ||
    formulaCalculation.calcProperties?.calcMode !== "auto" ||
    formulaCalculation.undo.calculationMode !== "except-tables" ||
    formulaCalculation.undo.calcProperties?.calcMode !== "autoNoTable" ||
    formulaCalculation.redo.calculationMode !== "automatic" ||
    formulaCalculation.redo.calcProperties?.calcMode !== "auto" ||
    !formulaCalculation.sheet.status.includes("CHT を再計算")
  ) {
    throw new Error("Formula ribbon should provide Excel-like function insertion, name creation, formula display, tracing, and calculation actions.");
  }


  const windowSwitch = viewWindowCommands.workbookWindowBarSwitch || {};


  const rangeSignature = (range) => range
    ? `${range.sheetIndex}:${range.top}:${range.left}:${range.bottom}:${range.right}`
    : "none";


  const normalizePane = (pane) => {
    if (!pane) return null;
    const rows = Number(pane.rows ?? pane.row ?? 0);
    const columns = Number(pane.columns ?? pane.col ?? 0);
    return {
      rows: Number.isFinite(rows) ? rows : 0,
      columns: Number.isFinite(columns) ? columns : 0,
    };
  };


  const paneSignature = (pane) => {
    const normalized = normalizePane(pane);
    return normalized ? `${normalized.rows}:${normalized.columns}` : "none";
  };


	  const windowSheetViewSignature = (view) => [
	    view?.showGridLines !== false,
	    view?.showHeadings !== false,
	    view?.showFormulas === true,
	    paneSignature(view?.freezePanes),
	    paneSignature(view?.splitPane),
	  ].join("|");


  const windowSwitchSelectionsDistinct =
    rangeSignature(windowSwitch.beforeStoredSelection) !== rangeSignature(windowSwitch.targetStoredSelection);


  const windowSwitchViewsDistinct =
    windowSheetViewSignature(windowSwitch.restoredOriginalSheetView) !== windowSheetViewSignature(windowSwitch.targetSheetView);


  const windowSwitchStoredMatchesVisible =
    windowSheetViewSignature(windowSwitch.beforeStoredView) === windowSheetViewSignature(windowSwitch.restoredOriginalSheetView) &&
    windowSheetViewSignature(windowSwitch.targetStoredView) === windowSheetViewSignature(windowSwitch.targetSheetView);


  const rowHeightSignature = (row) => Number(row?.height || 0);


  const rowHeightRestored = (before, after) => Math.abs(rowHeightSignature(before) - rowHeightSignature(after)) <= 0.5;


  if (
    viewRibbon.activeRibbonTab !== "表示" ||
    !viewRibbon.panelVisible ||
    ["ブックの表示", "表示", "ズーム", "ウィンドウ"].some((label) => !viewRibbon.groupLabels.includes(label)) ||
    [
      "view-normal",
      "view-page-layout",
      "view-page-break",
      "view-custom-views",
      "view-full-screen",
      "view-sheet-view-new",
      "view-sheet-view-keep",
      "view-sheet-view-exit",
      "view-sheet-view-options",
      "view-ruler",
      "view-gridlines",
      "view-headings",
      "view-formula-bar",
      "view-zero-values",
      "view-objects",
      "view-outline-symbols",
      "view-right-to-left",
      "view-zoom",
      "view-zoom-selection",
      "view-freeze-panes",
      "view-freeze-menu",
      "view-split",
      "view-hide-window",
      "view-unhide-window",
      "view-arrange-all",
      "view-side-by-side",
      "view-sync-scrolling",
      "view-reset-window-position",
      "view-switch-windows",
      "view-macros",
      "view-macros-menu",
    ].some((command) => !viewRibbon.allCommands.includes(command)) ||
    viewRibbon.showCheckBoxCount !== 8 ||
    !viewRibbon.hasFreezeSplit ||
    !viewRibbon.windowCommandTitles?.["view-freeze-panes"]?.includes("選択") ||
    !viewRibbon.windowCommandTitles?.["view-new-window"]?.includes("表示中") ||
    !viewRibbon.windowCommandTitles?.["view-arrange-all"]?.includes("表示中") ||
    !viewRibbon.windowCommandTitles?.["view-sync-scrolling"]?.includes("並べて比較") ||
    viewRibbon.panelHeight > 74 ||
    viewWindowCommands.viewFullScreenOn?.modelState !== true ||
    viewWindowCommands.viewFullScreenOn?.bodyClass !== true ||
    viewWindowCommands.viewFullScreenOn?.buttonActive !== true ||
    viewWindowCommands.viewFullScreenOn?.panelsDisplay !== "none" ||
    viewWindowCommands.viewFullScreenOn?.formulaBarDisplay !== "none" ||
    viewWindowCommands.fullScreenUndoRedo?.undo?.modelState !== false ||
    viewWindowCommands.fullScreenUndoRedo?.undo?.bodyClass !== false ||
    viewWindowCommands.fullScreenUndoRedo?.undo?.buttonActive !== false ||
    !viewWindowCommands.fullScreenUndoRedo?.undo?.status?.includes("元に戻しました") ||
    viewWindowCommands.fullScreenUndoRedo?.redo?.modelState !== true ||
    viewWindowCommands.fullScreenUndoRedo?.redo?.bodyClass !== true ||
    viewWindowCommands.fullScreenUndoRedo?.redo?.buttonActive !== true ||
    !viewWindowCommands.fullScreenUndoRedo?.redo?.status?.includes("やり直しました") ||
    viewWindowCommands.viewFullScreenOff?.modelState !== false ||
    viewWindowCommands.viewFullScreenOff?.bodyClass !== false ||
    viewWindowCommands.viewFullScreenOff?.buttonActive !== false ||
    !viewWindowCommands.viewFullScreenOff?.status?.includes("全画面表示を解除") ||
	    viewWindowCommands.viewZoomDialog?.title !== "ズーム" ||
	    !viewWindowCommands.viewZoomDialog?.nearRibbonButton ||
	    !viewWindowCommands.viewZoomDialog?.accessKeys?.includes("7") ||
	    !viewWindowCommands.viewZoomDialog?.accessKeys?.includes("F") ||
	    !viewWindowCommands.viewZoomDialog?.accessKeys?.includes("C") ||
	    !viewWindowCommands.viewZoomDialog?.accessKeysUnderlined ||
	    viewWindowCommands.viewZoomApplied?.zoom !== 75 ||
    !viewWindowCommands.viewZoomApplied?.label?.includes("75") ||
    viewWindowCommands.viewZoomHundred?.zoom !== 100 ||
    viewWindowCommands.viewZoomHundred?.active !== true ||
    !viewWindowCommands.viewZoomHundred?.title?.includes("100%") ||
    !viewWindowCommands.viewZoomSelection?.rangeLabel?.includes(":") ||
    viewWindowCommands.viewZoomSelection?.zoom === 100 ||
    !viewWindowCommands.viewZoomSelection?.label?.includes(String(viewWindowCommands.viewZoomSelection?.zoom)) ||
    !viewWindowCommands.viewZoomSelection?.status?.includes("ズーム") ||
    viewWindowCommands.viewZoomSelection?.activeHundred !== false ||
    viewWindowCommands.viewZoomSelectionUndoRedo?.undo?.zoom !== 100 ||
    viewWindowCommands.viewZoomSelectionUndoRedo?.undo?.activeHundred !== true ||
    !viewWindowCommands.viewZoomSelectionUndoRedo?.undo?.status?.includes("元に戻しました") ||
    viewWindowCommands.viewZoomSelectionUndoRedo?.redo?.zoom !== viewWindowCommands.viewZoomSelection?.zoom ||
    viewWindowCommands.viewZoomSelectionUndoRedo?.redo?.activeHundred !== false ||
    !viewWindowCommands.viewZoomSelectionUndoRedo?.redo?.status?.includes("やり直しました") ||
    viewWindowCommands.viewRulerOn?.showRuler !== true ||
    viewWindowCommands.viewRulerOn?.viewMode !== "page-layout" ||
    !viewWindowCommands.viewRulerOn?.gridHasRuler ||
    !viewWindowCommands.viewRulerOn?.rulerButtonActive ||
    viewWindowCommands.viewRulerOn?.horizontalRulerTicks < 4 ||
    viewWindowCommands.viewRulerOn?.verticalRulerTicks < 4 ||
    !viewWindowCommands.viewRulerOn?.majorRulerLabels?.length ||
    !viewWindowCommands.viewRulerOn?.headerFooterPlaceholders?.some((text) => text.includes("ヘッダー")) ||
    viewWindowCommands.viewRulerOn?.headerPointerEvents !== "auto" ||
    viewWindowCommands.pageLayoutHeaderFooterDialog?.title !== "ヘッダーとフッター" ||
    viewWindowCommands.pageLayoutHeaderFooterDialog?.firstFieldFocused !== "oddHeader.center" ||
    viewWindowCommands.pageBreakPreview?.viewMode !== "page-break" ||
    viewWindowCommands.pageBreakPreview?.guideCount < 1 ||
    !viewWindowCommands.pageBreakPreview?.label?.includes("ページ") ||
	    !pageBreakContext.initial?.labels?.some((label) => label.includes("改ページの挿入")) ||
	    !pageBreakContext.initial?.labels?.some((label) => label.includes("改ページの解除")) ||
	    !pageBreakContext.initial?.labels?.some((label) => label.includes("すべての改ページ")) ||
	    !pageBreakContext.initial?.accessKeys?.includes("I") ||
	    !pageBreakContext.initial?.accessKeys?.includes("R") ||
	    !pageBreakContext.initial?.accessKeys?.includes("A") ||
	    !pageBreakContext.initial?.accessKeys?.includes("V") ||
	    !pageBreakContext.initial?.accessKeysUnderlined ||
	    !pageBreakContext.initial?.disabledActions?.some((action) => action.startsWith("page-break:remove:")) ||
	    !pageBreakContext.initial?.descriptions?.some((text) => text.includes("手動改ページ")) ||
	    pageBreakContext.inserted?.inserted !== true ||
	    !pageBreakContext.inserted?.status?.includes("挿入") ||
	    pageBreakContext.inserted?.manualGuideCount < 1 ||
	    pageBreakContext.manual?.checkedRemoveEnabled !== true ||
	    pageBreakContext.manual?.disabledActions?.some((action) => action.startsWith("page-break:remove:")) ||
	    pageBreakContext.removed?.removed !== true ||
	    !pageBreakContext.removed?.status?.includes("解除") ||
	    !pageBreakContextUndoList.includes(pageBreakContextIndex) ||
	    !pageBreakContext.undo?.status?.includes("元に戻しました") ||
	    pageBreakContextRedoList.includes(pageBreakContextIndex) ||
	    !pageBreakContext.redo?.status?.includes("やり直しました") ||
	    pageBreakContext.closePreview?.viewMode !== "normal" ||
	    pageBreakContext.closePreview?.activeNormal !== true ||
	    viewWindowCommands.customViewsMainBeforeAdd?.title !== "ユーザー設定のビュー" ||
	    viewWindowCommands.customViewsMainBeforeAdd?.hasInlineNameField !== false ||
	    !viewWindowCommands.customViewsMainBeforeAdd?.addButtonText?.includes("追加") ||
	    viewWindowCommands.customViewsMainBeforeAdd?.addAccessKey !== "A" ||
	    !viewWindowCommands.customViewsMainBeforeAdd?.accessKeysUnderlined ||
	    viewWindowCommands.customViewAddDialog?.title !== "ビューの追加" ||
	    !viewWindowCommands.customViewAddDialog?.defaultName?.includes("ビュー") ||
	    viewWindowCommands.customViewAddDialog?.printChecked !== true ||
	    viewWindowCommands.customViewAddDialog?.hiddenRowsChecked !== true ||
	    !viewWindowCommands.customViewAddDialog?.fieldsetLegend?.includes("含める") ||
	    !viewWindowCommands.customViewAddDialog?.checkboxAccessKeys?.includes("P") ||
	    !viewWindowCommands.customViewAddDialog?.checkboxAccessKeys?.includes("R") ||
	    !viewWindowCommands.customViewAddDialog?.accessKeysUnderlined ||
    !viewWindowCommands.customViewsDialog?.names?.includes("表示テスト") ||
    viewWindowCommands.customViewsDialog?.selectedText !== "表示テスト" ||
    viewWindowCommands.customViewsDialog?.savedView?.includePrintSettings !== true ||
    viewWindowCommands.customViewsDialog?.savedView?.includeHiddenRowsColumnsAndFilters !== true ||
    !viewWindowCommands.customViewsDialog?.savedView?.sheets?.some((sheet) => sheet.name === "CHT" && sheet.pageLayout) ||
	    viewWindowCommands.customViewsDialog?.savedView?.selectionRange?.top !== 30 ||
	    viewWindowCommands.customViewsDialog?.savedView?.selectionRange?.left > 30 ||
	    viewWindowCommands.customViewsDialog?.savedView?.selectionRange?.bottom !== 31 ||
	    viewWindowCommands.customViewsDialog?.savedView?.selectionRange?.right < 31 ||
	    viewWindowCommands.customViewsDialog?.savedView?.selectionRanges?.[0]?.top !== 30 ||
	    viewWindowCommands.customViewsDialog?.savedView?.selectionRanges?.[0]?.left > 30 ||
	    viewWindowCommands.customViewsDialog?.savedView?.selectionRanges?.[0]?.bottom !== 31 ||
	    viewWindowCommands.customViewsDialog?.savedView?.selectionRanges?.[0]?.right < 31 ||
    viewWindowCommands.customViewsDialog?.savedView?.activeScroll?.top < 1 ||
	    viewWindowCommands.customViewsDialog?.savedView?.activeScroll?.left < 1 ||
	    viewWindowCommands.customViewsDialog?.savedView?.mainWindowView?.scrollTop == null ||
	    viewWindowCommands.customViewSaveUndoRedo?.undo?.names?.includes("表示テスト") ||
	    viewWindowCommands.customViewSaveUndoRedo?.undo?.dialogNames?.includes("表示テスト") ||
	    !viewWindowCommands.customViewSaveUndoRedo?.undo?.status?.includes("元に戻しました") ||
	    !viewWindowCommands.customViewSaveUndoRedo?.redo?.names?.includes("表示テスト") ||
	    !viewWindowCommands.customViewSaveUndoRedo?.redo?.dialogNames?.includes("表示テスト") ||
	    !viewWindowCommands.customViewSaveUndoRedo?.redo?.status?.includes("やり直しました") ||
	    !viewWindowCommands.customViewCaptureState?.selectionLabel?.includes("30") ||
	    viewWindowCommands.customViewApplied?.selectionLabel !== viewWindowCommands.customViewCaptureState?.selectionLabel ||
		    viewWindowCommands.customViewApplied?.scrollTop < 1 ||
		    viewWindowCommands.customViewApplied?.scrollLeft < 1 ||
		    viewWindowCommands.customViewApplied?.viewMode !== "page-layout" ||
		    viewWindowCommands.customViewApplyUndoRedo?.undo?.selectionLabel !== "A1" ||
		    viewWindowCommands.customViewApplyUndoRedo?.undo?.scrollTop > 5 ||
		    viewWindowCommands.customViewApplyUndoRedo?.undo?.scrollLeft > 5 ||
		    !viewWindowCommands.customViewApplyUndoRedo?.undo?.status?.includes("元に戻しました") ||
	    viewWindowCommands.customViewApplyUndoRedo?.redo?.selectionLabel !== viewWindowCommands.customViewCaptureState?.selectionLabel ||
	    viewWindowCommands.customViewApplyUndoRedo?.redo?.scrollTop < 1 ||
	    viewWindowCommands.customViewApplyUndoRedo?.redo?.scrollLeft < 1 ||
	    !viewWindowCommands.customViewApplyUndoRedo?.redo?.status?.includes("やり直しました") ||
		    viewWindowCommands.customViewsDialog?.showDisabled !== false ||
	    !viewWindowCommands.customViewDeleteFlow?.before?.names?.includes("表示テスト") ||
	    !viewWindowCommands.customViewDeleteFlow?.before?.names?.includes("削除テスト") ||
	    viewWindowCommands.customViewDeleteFlow?.before?.selectedText !== "削除テスト" ||
	    viewWindowCommands.customViewDeleteFlow?.before?.deleteDisabled !== false ||
	    viewWindowCommands.customViewDeleteFlow?.prompt?.title !== "Cht WebSheet" ||
	    !viewWindowCommands.customViewDeleteFlow?.prompt?.message?.includes("削除テスト") ||
	    !viewWindowCommands.customViewDeleteFlow?.prompt?.buttons?.some((label) => label.includes("削除")) ||
	    !viewWindowCommands.customViewDeleteFlow?.prompt?.accessKeys?.includes("D") ||
		    viewWindowCommands.customViewDeleteFlow?.after?.names?.includes("削除テスト") ||
		    viewWindowCommands.customViewDeleteFlow?.after?.dialogNames?.includes("削除テスト") ||
		    !viewWindowCommands.customViewDeleteFlow?.after?.names?.includes("表示テスト") ||
		    viewWindowCommands.customViewDeleteFlow?.after?.dialogVisible !== true ||
		    !viewWindowCommands.customViewDeleteFlow?.after?.status?.includes("削除しました") ||
	    !viewWindowCommands.customViewDeleteFlow?.undoRedo?.undo?.names?.includes("削除テスト") ||
	    !viewWindowCommands.customViewDeleteFlow?.undoRedo?.undo?.dialogNames?.includes("削除テスト") ||
	    !viewWindowCommands.customViewDeleteFlow?.undoRedo?.undo?.status?.includes("元に戻しました") ||
	    viewWindowCommands.customViewDeleteFlow?.undoRedo?.redo?.names?.includes("削除テスト") ||
	    viewWindowCommands.customViewDeleteFlow?.undoRedo?.redo?.dialogNames?.includes("削除テスト") ||
	    !viewWindowCommands.customViewDeleteFlow?.undoRedo?.redo?.status?.includes("やり直しました") ||
		    viewWindowCommands.sheetViewNew?.count < 1 ||
    viewWindowCommands.sheetViewNew?.temporary !== true ||
    !viewWindowCommands.sheetViewNew?.gridMarked ||
    !viewWindowCommands.sheetViewNew?.newActive ||
    viewWindowCommands.sheetViewNew?.keepDisabled !== false ||
    viewWindowCommands.sheetViewNew?.exitDisabled !== false ||
    viewWindowCommands.sheetViewKept?.temporary !== false ||
    !viewWindowCommands.sheetViewMenuBeforeRename?.labels?.some((label) => label.includes("既定")) ||
    !viewWindowCommands.sheetViewMenuBeforeRename?.labels?.some((label) => label.includes("ビュー")) ||
    viewWindowCommands.sheetViewMenuBeforeRename?.descriptions?.length < 3 ||
    !viewWindowCommands.sheetViewMenuBeforeRename?.descriptions?.some((text) => text.includes("更新")) ||
    !viewWindowCommands.sheetViewMenuBeforeRename?.accessKeysUnderlined ||
    viewWindowCommands.sheetViewRenamed?.name !== "表示ビュー" ||
    viewWindowCommands.sheetViewRenamed?.temporary !== false ||
    viewWindowCommands.sheetViewUpdated?.activeName !== "表示ビュー" ||
    viewWindowCommands.sheetViewUpdated?.temporary !== false ||
    viewWindowCommands.sheetViewUpdated?.snapshotGridlines !== false ||
    !viewWindowCommands.sheetViewUpdated?.status?.includes("更新") ||
    !viewWindowCommands.sheetViewMenuBeforeCopy?.actions?.includes("sheet-view:update") ||
    !viewWindowCommands.sheetViewMenuBeforeCopy?.actions?.includes("sheet-view:duplicate") ||
    !viewWindowCommands.sheetViewMenuBeforeCopy?.accessKeys?.includes("U") ||
    !viewWindowCommands.sheetViewMenuBeforeCopy?.accessKeys?.includes("C") ||
    !viewWindowCommands.sheetViewDuplicated?.activeName?.includes("コピー") ||
    viewWindowCommands.sheetViewDuplicated?.temporary !== false ||
    viewWindowCommands.sheetViewDuplicated?.snapshotGridlines !== false ||
    viewWindowCommands.sheetViewDuplicated?.count < 2 ||
    !viewWindowCommands.sheetViewDuplicated?.status?.includes("コピー") ||
    viewWindowCommands.sheetViewExited?.activeId ||
    viewWindowCommands.sheetViewExited?.count < 1 ||
    viewWindowCommands.sheetViewExited?.gridMarked !== false ||
    viewWindowCommands.sheetViewExited?.keepDisabled !== true ||
    viewWindowCommands.sheetViewExited?.exitDisabled !== true ||
    viewWindowCommands.sheetViewReactivated?.activeName !== "表示ビュー" ||
    !viewWindowCommands.sheetViewReactivated?.gridMarked ||
    !viewWindowCommands.freezeMenuLabels?.includes("先頭行の固定") ||
    viewWindowCommands.freezeMenuInitialMeta?.focusedAfterArrowDown !== "freeze-top-row" ||
    viewWindowCommands.freezeMenuInitialMeta?.focusedAfterArrowRight !== "freeze-first-column" ||
    viewWindowCommands.freezeMenuInitialMeta?.descriptionCount < 3 ||
    viewWindowCommands.freezeMenuInitialMeta?.width < 300 ||
    !viewWindowCommands.freezeMenuInitialMeta?.descriptions?.some((text) => text.includes("選択範囲")) ||
    !viewWindowCommands.freezeMenuInitialMeta?.descriptions?.some((text) => text.includes("先頭行")) ||
    viewWindowCommands.freezeMenuTopRow?.rows !== 1 ||
    viewWindowCommands.freezeMenuTopRow?.columns !== 0 ||
    viewWindowCommands.freezeMenuTopRowStatus?.checkedAction !== "freeze-top-row" ||
    viewWindowCommands.freezeMenuTopRowStatus?.topRowChecked !== "true" ||
    viewWindowCommands.freezeMenuTopRowStatus?.selectionChecked !== "" ||
    viewWindowCommands.freezeMenuTopRowStatus?.checkMarkCount !== 1 ||
    !viewWindowCommands.freezeMenuTopRowStatus?.mainButtonLabel?.includes("解除") ||
    viewWindowCommands.freezeMenuTopRowStatus?.mainButtonIcon !== "unlock-keyhole" ||
    viewWindowCommands.freezeMenuTopRowStatus?.unfreezeFirst !== true ||
    viewWindowCommands.freezeMenuAccessKey?.pane?.columns !== 1 ||
    viewWindowCommands.freezeMenuAccessKey?.pane?.rows !== 0 ||
    viewWindowCommands.freezeMenuAccessKey?.menuClosed !== true ||
    !viewWindowCommands.freezeMenuAccessKey?.mainButtonLabel?.includes("解除") ||
    !viewWindowCommands.freezeMenuAccessKey?.mainButtonTitle?.includes("解除") ||
    !viewWindowCommands.freezeMenuAccessKey?.status?.includes("固定") ||
    viewToggles.showGridLines !== false ||
    viewToggles.showHeadings !== false ||
    viewToggles.showZeros !== false ||
    viewToggles.showObjects !== false ||
    viewToggles.showOutlineSymbols !== false ||
    viewToggles.rightToLeft !== true ||
    !viewToggles.gridHasHideGridlines ||
    !viewToggles.gridHasHideHeadings ||
    !viewToggles.gridHasHideOutlineSymbols ||
    !viewToggles.gridHasRightToLeft ||
    viewToggles.zeroCellText !== "" ||
    viewToggles.visibleObjectLayerItems !== 0 ||
	    viewToggles.formulaBarVisible !== false ||
	    viewToggles.activeCommands.includes("view-gridlines") ||
	    viewToggles.activeCommands.includes("view-headings") ||
	    viewToggles.activeCommands.includes("view-zero-values") ||
	    viewToggles.activeCommands.includes("view-objects") ||
	    viewToggles.activeCommands.includes("view-outline-symbols") ||
	    viewToggles.activeCommands.includes("view-formula-bar") ||
	    !viewToggles.activeCommands.includes("view-right-to-left") ||
	    viewWindowCommands.formulaBarUndoRedo?.undo?.formulaBarVisible !== true ||
	    viewWindowCommands.formulaBarUndoRedo?.undo?.commandActive !== true ||
	    !viewWindowCommands.formulaBarUndoRedo?.undo?.status?.includes("元に戻しました") ||
	    viewWindowCommands.formulaBarUndoRedo?.redo?.formulaBarVisible !== false ||
	    viewWindowCommands.formulaBarUndoRedo?.redo?.commandActive !== false ||
	    !viewWindowCommands.formulaBarUndoRedo?.redo?.status?.includes("やり直しました") ||
	    viewWindowCommands.gridlinesUndoRedo?.undo?.showGridLines !== false ||
	    viewWindowCommands.gridlinesUndoRedo?.undo?.gridHasHideGridlines !== true ||
	    viewWindowCommands.gridlinesUndoRedo?.undo?.commandActive !== false ||
	    viewWindowCommands.gridlinesUndoRedo?.undo?.ariaChecked !== "false" ||
	    !viewWindowCommands.gridlinesUndoRedo?.undo?.status?.includes("元に戻しました") ||
	    viewWindowCommands.gridlinesUndoRedo?.redo?.showGridLines !== true ||
	    viewWindowCommands.gridlinesUndoRedo?.redo?.gridHasHideGridlines !== false ||
	    viewWindowCommands.gridlinesUndoRedo?.redo?.commandActive !== true ||
	    viewWindowCommands.gridlinesUndoRedo?.redo?.ariaChecked !== "true" ||
	    !viewWindowCommands.gridlinesUndoRedo?.redo?.status?.includes("やり直しました") ||
	    viewWindowCommands.headingsUndoRedo?.undo?.showHeadings !== false ||
	    viewWindowCommands.headingsUndoRedo?.undo?.gridHasHideHeadings !== true ||
	    viewWindowCommands.headingsUndoRedo?.undo?.commandActive !== false ||
	    viewWindowCommands.headingsUndoRedo?.undo?.ariaChecked !== "false" ||
	    !viewWindowCommands.headingsUndoRedo?.undo?.status?.includes("元に戻しました") ||
	    viewWindowCommands.headingsUndoRedo?.redo?.showHeadings !== true ||
	    viewWindowCommands.headingsUndoRedo?.redo?.gridHasHideHeadings !== false ||
	    viewWindowCommands.headingsUndoRedo?.redo?.commandActive !== true ||
	    viewWindowCommands.headingsUndoRedo?.redo?.ariaChecked !== "true" ||
	    !viewWindowCommands.headingsUndoRedo?.redo?.status?.includes("やり直しました") ||
	    viewWindowCommands.zeroValuesUndoRedo?.undo?.showZeros !== false ||
	    viewWindowCommands.zeroValuesUndoRedo?.undo?.commandActive !== false ||
	    viewWindowCommands.zeroValuesUndoRedo?.undo?.ariaChecked !== "false" ||
	    !viewWindowCommands.zeroValuesUndoRedo?.undo?.status?.includes("元に戻しました") ||
	    viewWindowCommands.zeroValuesUndoRedo?.redo?.showZeros !== true ||
	    viewWindowCommands.zeroValuesUndoRedo?.redo?.commandActive !== true ||
	    viewWindowCommands.zeroValuesUndoRedo?.redo?.ariaChecked !== "true" ||
	    !viewWindowCommands.zeroValuesUndoRedo?.redo?.status?.includes("やり直しました") ||
	    viewWindowCommands.objectsUndoRedo?.undo?.showObjects !== false ||
	    viewWindowCommands.objectsUndoRedo?.undo?.commandActive !== false ||
	    viewWindowCommands.objectsUndoRedo?.undo?.ariaChecked !== "false" ||
	    !viewWindowCommands.objectsUndoRedo?.undo?.status?.includes("元に戻しました") ||
	    viewWindowCommands.objectsUndoRedo?.redo?.showObjects !== true ||
	    viewWindowCommands.objectsUndoRedo?.redo?.commandActive !== true ||
	    viewWindowCommands.objectsUndoRedo?.redo?.ariaChecked !== "true" ||
	    !viewWindowCommands.objectsUndoRedo?.redo?.status?.includes("やり直しました") ||
	    viewWindowCommands.outlineUndoRedo?.undo?.showOutlineSymbols !== false ||
	    viewWindowCommands.outlineUndoRedo?.undo?.gridHasHideOutlineSymbols !== true ||
	    viewWindowCommands.outlineUndoRedo?.undo?.commandActive !== false ||
	    viewWindowCommands.outlineUndoRedo?.undo?.ariaChecked !== "false" ||
	    !viewWindowCommands.outlineUndoRedo?.undo?.status?.includes("元に戻しました") ||
	    viewWindowCommands.outlineUndoRedo?.redo?.showOutlineSymbols !== true ||
	    viewWindowCommands.outlineUndoRedo?.redo?.gridHasHideOutlineSymbols !== false ||
	    viewWindowCommands.outlineUndoRedo?.redo?.commandActive !== true ||
	    viewWindowCommands.outlineUndoRedo?.redo?.ariaChecked !== "true" ||
	    !viewWindowCommands.outlineUndoRedo?.redo?.status?.includes("やり直しました") ||
	    viewWindowCommands.rightToLeftUndoRedo?.undo?.rightToLeft !== true ||
	    viewWindowCommands.rightToLeftUndoRedo?.undo?.gridHasRightToLeft !== true ||
	    viewWindowCommands.rightToLeftUndoRedo?.undo?.commandActive !== true ||
	    viewWindowCommands.rightToLeftUndoRedo?.undo?.ariaChecked !== "true" ||
	    !viewWindowCommands.rightToLeftUndoRedo?.undo?.status?.includes("元に戻しました") ||
	    viewWindowCommands.rightToLeftUndoRedo?.redo?.rightToLeft !== false ||
	    viewWindowCommands.rightToLeftUndoRedo?.redo?.gridHasRightToLeft !== false ||
	    viewWindowCommands.rightToLeftUndoRedo?.redo?.commandActive !== false ||
	    viewWindowCommands.rightToLeftUndoRedo?.redo?.ariaChecked !== "false" ||
	    !viewWindowCommands.rightToLeftUndoRedo?.redo?.status?.includes("やり直しました") ||
	    viewWindowCommands.rtlVisualHitTest?.visibleCellCount < 2 ||
	    viewWindowCommands.rtlVisualHitTest?.firstVisualCol !== viewWindowCommands.rtlVisualHitTest?.highestRendered ||
	    viewWindowCommands.rtlVisualHitTest?.selectedCol !== viewWindowCommands.rtlVisualHitTest?.firstVisualCol ||
	    !viewWindowCommands.rtlVisualHitTest?.selectionLabel?.includes(viewWindowCommands.rtlVisualHitTest?.firstVisualAddress) ||
	    viewToggles.showChecks?.length !== 8 ||
	    viewToggles.showChecks.some((item) => !item.hasCheckBox) ||
	    viewToggles.showChecks.some((item) => item.role !== "checkbox") ||
    viewToggles.showChecks.some((item) => item.active !== (item.ariaChecked === "true")) ||
    viewWindowCommands.freezePanesAfterMain?.rows !== 3 ||
    viewWindowCommands.freezePanesAfterMain?.columns !== 3 ||
    viewWindowCommands.frozenPaneSticky?.cornerPosition !== "sticky" ||
    !viewWindowCommands.frozenPaneSticky?.cornerClasses?.includes("is-frozen-corner") ||
    viewWindowCommands.frozenPaneSticky?.cornerTop == null ||
    viewWindowCommands.frozenPaneSticky.cornerTop > 90 ||
    viewWindowCommands.frozenPaneSticky?.cornerLeft == null ||
    viewWindowCommands.frozenPaneSticky.cornerLeft > 100 ||
    viewWindowCommands.frozenPaneSticky?.rowCellPosition !== "sticky" ||
    viewWindowCommands.frozenPaneSticky?.columnCellPosition !== "sticky" ||
    !viewWindowCommands.frozenPaneSticky?.cornerHeaderSticky ||
    !viewWindowCommands.frozenPaneSticky?.rowHeaderSticky ||
    viewWindowCommands.splitPaneBeforeDrag?.row !== 4 ||
    viewWindowCommands.splitPaneBeforeDrag?.col !== 4 ||
    viewWindowCommands.splitPaneDragged?.pane?.columns <= 3 ||
    viewWindowCommands.splitPaneDragged?.pane?.rows <= 3 ||
    viewWindowCommands.splitPaneDragged?.verticalGuideCount !== 1 ||
    viewWindowCommands.splitPaneDragged?.horizontalGuideCount !== 1 ||
    viewWindowCommands.sheetGroupView?.gridlines?.first !== false ||
    viewWindowCommands.sheetGroupView?.gridlines?.second !== false ||
    !viewWindowCommands.sheetGroupView?.gridlines?.status?.includes("2 個のシート") ||
    viewWindowCommands.sheetGroupView?.gridlinesUndo?.first === false ||
    viewWindowCommands.sheetGroupView?.gridlinesUndo?.second === false ||
    !viewWindowCommands.sheetGroupView?.gridlinesUndo?.status?.includes("元に戻しました") ||
    viewWindowCommands.sheetGroupView?.gridlinesRedo?.first !== false ||
    viewWindowCommands.sheetGroupView?.gridlinesRedo?.second !== false ||
    !viewWindowCommands.sheetGroupView?.gridlinesRedo?.status?.includes("やり直しました") ||
    viewWindowCommands.sheetGroupView?.showOptions?.first?.showHeadings !== false ||
    viewWindowCommands.sheetGroupView?.showOptions?.second?.showHeadings !== false ||
    viewWindowCommands.sheetGroupView?.showOptions?.first?.showZeros !== false ||
    viewWindowCommands.sheetGroupView?.showOptions?.second?.showZeros !== false ||
    viewWindowCommands.sheetGroupView?.showOptions?.first?.showObjects !== false ||
    viewWindowCommands.sheetGroupView?.showOptions?.second?.showObjects !== false ||
    viewWindowCommands.sheetGroupView?.showOptions?.first?.showOutlineSymbols !== false ||
    viewWindowCommands.sheetGroupView?.showOptions?.second?.showOutlineSymbols !== false ||
    viewWindowCommands.sheetGroupView?.showOptions?.first?.rightToLeft !== true ||
    viewWindowCommands.sheetGroupView?.showOptions?.second?.rightToLeft !== true ||
    viewWindowCommands.sheetGroupView?.freeze?.first?.rows !== 3 ||
    viewWindowCommands.sheetGroupView?.freeze?.second?.rows !== 3 ||
    viewWindowCommands.sheetGroupView?.freeze?.first?.columns !== 3 ||
    viewWindowCommands.sheetGroupView?.freeze?.second?.columns !== 3 ||
    !viewWindowCommands.sheetGroupView?.freeze?.status?.includes("2 個のシート") ||
    viewWindowCommands.sheetGroupView?.split?.first?.row !== 4 ||
    viewWindowCommands.sheetGroupView?.split?.second?.row !== 4 ||
    viewWindowCommands.sheetGroupView?.split?.first?.col !== 4 ||
    viewWindowCommands.sheetGroupView?.split?.second?.col !== 4 ||
    !viewWindowCommands.sheetGroupView?.split?.status?.includes("2 個のシート") ||
    viewWindowCommands.sheetGroupView?.zoom?.first !== 75 ||
    viewWindowCommands.sheetGroupView?.zoom?.second !== 75 ||
    !viewWindowCommands.sheetGroupView?.zoom?.status?.includes("2 個のシート") ||
    viewWindowCommands.sheetGroupView?.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介" ||
    viewWindowCommands.sheetGroupView?.cleanup?.active !== "CHT" ||
    !viewWindowCommands.windowViews.length ||
    viewWindowCommands.viewNewWindowState?.activeWindowId === "main" ||
    viewWindowCommands.viewNewWindowState?.windowCount < 2 ||
    viewWindowCommands.viewNewWindowState?.activeWindow?.scrollTop !== viewWindowCommands.viewNewWindowState?.hostScrollTop ||
    viewWindowCommands.viewNewWindowState?.activeWindow?.scrollLeft !== viewWindowCommands.viewNewWindowState?.hostScrollLeft ||
    viewWindowCommands.viewSecondNewWindowState?.windowCount < 3 ||
    viewWindowCommands.viewSecondNewWindowState?.visibleCount < 3 ||
    /:\d+$/.test(viewWindowCommands.viewSecondNewWindowState?.title || "") ||
    viewWindowCommands.viewSecondNewWindowState?.title !== viewWindowCommands.viewSecondNewWindowState?.workbookName ||
    viewWindowCommands.viewInitialWindowBar?.visible !== false ||
    viewWindowCommands.viewNewWindowBar?.visible !== true ||
    viewWindowCommands.viewNewWindowBar?.count !== 2 ||
    viewWindowCommands.viewNewWindowBar?.chipIds?.length !== 2 ||
    viewWindowCommands.viewNewWindowBar?.previewSlotIds?.length !== 2 ||
    viewWindowCommands.viewNewWindowBar?.activeChipId !== viewWindowCommands.viewNewWindowState?.activeWindowId ||
    viewWindowCommands.viewNewWindowBar?.activePreviewId !== viewWindowCommands.viewNewWindowState?.activeWindowId ||
    !viewWindowCommands.viewNewWindowBar?.summary?.includes("ウィンドウ") ||
    viewWindowCommands.viewSecondNewWindowBar?.visible !== true ||
    viewWindowCommands.viewSecondNewWindowBar?.count !== 3 ||
    viewWindowCommands.viewSecondNewWindowBar?.chipIds?.length !== 3 ||
    viewWindowCommands.viewSecondNewWindowBar?.previewSlotIds?.length !== 3 ||
    viewWindowCommands.viewSecondNewWindowBar?.chipLabels?.some((label) => !/:\d+\b/.test(label || "")) ||
    !windowSwitch.targetId ||
    windowSwitch.after !== windowSwitch.targetId ||
    windowSwitch.activeChipId !== windowSwitch.targetId ||
    windowSwitch.previewActiveId !== windowSwitch.targetId ||
    !windowSwitch.restoredOriginalLabel ||
    !windowSwitch.targetSelectionLabel ||
    !windowSwitchSelectionsDistinct ||
	    !windowSwitchViewsDistinct ||
	    !windowSwitchStoredMatchesVisible ||
	    !Object.prototype.hasOwnProperty.call(windowSwitch.beforeStoredView || {}, "showFormulas") ||
	    !Object.prototype.hasOwnProperty.call(windowSwitch.targetStoredView || {}, "showFormulas") ||
	    !windowSwitch.status?.includes("切り替え") ||
	    viewWindowCommands.customWindowSource?.skipped === true ||
	    viewWindowCommands.customWindowSource?.activeWindowId === "main" ||
	    !viewWindowCommands.customWindowSource?.selectionLabel ||
	    viewWindowCommands.customWindowSource?.sheet?.showGridLines === false ||
	    viewWindowCommands.customWindowSource?.sheet?.freezePanes?.rows !== 5 ||
	    viewWindowCommands.customWindowSource?.sheet?.freezePanes?.columns !== 5 ||
	    viewWindowCommands.customWindowMutated?.activeWindowId !== "main" ||
	    viewWindowCommands.customWindowMutated?.selectionLabel !== "A1" ||
	    viewWindowCommands.customWindowMutated?.sheet?.showGridLines !== false ||
	    viewWindowCommands.customWindowRestored?.activeWindowId !== viewWindowCommands.customWindowSource?.activeWindowId ||
	    viewWindowCommands.customWindowRestored?.selectionLabel !== viewWindowCommands.customWindowSource?.selectionLabel ||
	    viewWindowCommands.customWindowRestored?.sheet?.showGridLines === false ||
	    viewWindowCommands.customWindowRestored?.sheet?.freezePanes?.rows !== 5 ||
	    viewWindowCommands.customWindowRestored?.sheet?.freezePanes?.columns !== 5 ||
	    !viewWindowCommands.customWindowRestored?.status?.includes("ユーザー設定のビュー") ||
	    viewWindowCommands.arrangeWindowState?.mode !== "vertical" ||
    viewWindowCommands.arrangeWindowState?.windows?.length < 3 ||
    viewWindowCommands.arrangeWindowState?.windows?.some((windowView) => !/:\d+\b/.test(windowView.name || "")) ||
    viewWindowCommands.arrangeWindowBar?.mode !== "vertical" ||
    viewWindowCommands.arrangeWindowBar?.count !== viewWindowCommands.arrangeWindowState?.windows?.length ||
    viewWindowCommands.arrangeWindowBar?.previewSlotIds?.length !== viewWindowCommands.arrangeWindowState?.windows?.length ||
    viewWindowCommands.arrangeWindowUndo?.arrangement?.mode === "vertical" ||
    viewWindowCommands.arrangeWindowUndo?.bar?.mode === "vertical" ||
    !viewWindowCommands.arrangeWindowUndo?.status?.includes("元に戻しました") ||
    viewWindowCommands.arrangeWindowRedo?.arrangement?.mode !== "vertical" ||
    viewWindowCommands.arrangeWindowRedo?.bar?.mode !== "vertical" ||
    !viewWindowCommands.arrangeWindowRedo?.status?.includes("やり直しました") ||
    viewWindowCommands.sideBySideWindowDialog?.title !== "並べて比較" ||
    viewWindowCommands.sideBySideWindowDialog?.options?.length < 2 ||
    !viewWindowCommands.sideBySideWindowDialog?.accessKeys?.includes("C") ||
    !viewWindowCommands.sideBySideWindowDialog?.accessKeysUnderlined ||
    viewWindowCommands.sideBySideWindowDialog?.options?.some((option) => option.value === viewWindowCommands.sideBySideWindowDialog?.activeWindowId) ||
    viewWindowCommands.sideBySideWindowDialog?.options?.some((option) => !/:\d+\b/.test(option.text || "")) ||
    viewWindowCommands.sideBySideDialogApplied?.pair?.left !== viewWindowCommands.sideBySideDialogApplied?.activeWindowId ||
    viewWindowCommands.sideBySideDialogApplied?.pair?.right !== viewWindowCommands.sideBySideDialogApplied?.selectedId ||
    viewWindowCommands.sideBySideDialogApplied?.arrangement?.mode !== "side-by-side" ||
    viewWindowCommands.sideBySideDialogApplied?.arrangement?.windows?.length !== 2 ||
    viewWindowCommands.sideBySideWindowBar?.mode !== "side-by-side" ||
    viewWindowCommands.sideBySideWindowBar?.previewSlotIds?.length !== 2 ||
    viewWindowCommands.sideBySideWindowBar?.pairLeft !== viewWindowCommands.sideBySideDialogApplied?.pair?.left ||
    viewWindowCommands.sideBySideWindowBar?.pairRight !== viewWindowCommands.sideBySideDialogApplied?.pair?.right ||
    viewWindowCommands.sideBySideWindowBar?.syncScrolling !== "false" ||
    viewWindowCommands.sideBySideWindowBar?.comparedChipIds?.length !== 2 ||
    viewWindowCommands.sideBySideWindowBar?.comparedPreviewIds?.length !== 2 ||
    !viewWindowCommands.sideBySideWindowBar?.comparedChipIds?.includes(viewWindowCommands.sideBySideDialogApplied?.pair?.left) ||
    !viewWindowCommands.sideBySideWindowBar?.comparedChipIds?.includes(viewWindowCommands.sideBySideDialogApplied?.pair?.right) ||
    !viewWindowCommands.customViews?.some((view) => view.name === "表示テスト") ||
    !viewWindowCommands.viewSideBySide ||
    viewWindowCommands.synchronousScrolling !== true ||
    viewWindowCommands.syncScrollState?.windowArrangement?.mode !== "side-by-side" ||
    viewWindowCommands.syncScrollState?.sameScroll !== true ||
    viewWindowCommands.syncScrollUndoRedo?.undo?.synchronousScrolling !== false ||
    !viewWindowCommands.syncScrollUndoRedo?.undo?.status?.includes("元に戻しました") ||
    viewWindowCommands.syncScrollUndoRedo?.redo?.synchronousScrolling !== true ||
    viewWindowCommands.syncScrollUndoRedo?.redo?.bar?.syncScrolling !== "true" ||
    !viewWindowCommands.syncScrollUndoRedo?.redo?.bar?.badges?.some((badge) => badge.includes("同期スクロール：オン")) ||
    !viewWindowCommands.syncScrollUndoRedo?.redo?.status?.includes("やり直しました") ||
    viewWindowCommands.resetWindowState?.mode !== "side-by-side" ||
    viewWindowCommands.resetWindowState?.windows?.length !== 2 ||
    viewWindowCommands.resetWindowBar?.mode !== "side-by-side" ||
    viewWindowCommands.resetWindowBar?.previewSlotIds?.length !== 2 ||
    viewWindowCommands.switchWindowState?.activeWindowId !== viewWindowCommands.switchWindowState?.expectedWindowId ||
    viewWindowCommands.switchWindowState?.hostScrollTop !== viewWindowCommands.switchWindowState?.mainWindowView?.scrollTop ||
    viewWindowCommands.switchWindowState?.hostScrollLeft !== viewWindowCommands.switchWindowState?.mainWindowView?.scrollLeft ||
    /:\d+$/.test(viewWindowCommands.switchWindowState?.title || "") ||
    viewWindowCommands.switchWindowState?.title !== viewWindowCommands.switchWindowState?.workbookName ||
    (Number(viewWindowCommands.hideWindowState?.hiddenIds?.length || 0) + (viewWindowCommands.hideWindowState?.mainHidden ? 1 : 0)) < 1 ||
    viewWindowCommands.hideWindowBar?.hiddenCount < 1 ||
    viewWindowCommands.hideWindowBar?.count < 2 ||
    (viewWindowCommands.hideWindowState?.viewSideBySide && viewWindowCommands.hideWindowBar?.comparedChipIds?.length !== 2) ||
    (!viewWindowCommands.hideWindowState?.viewSideBySide && viewWindowCommands.hideWindowState?.synchronousScrolling !== false) ||
	    viewWindowCommands.unhideWindowState?.mainHidden !== false ||
	    viewWindowCommands.unhideWindowState?.hiddenIds?.length !== 0 ||
	    viewWindowCommands.unhideWindowBar?.hiddenCount !== 0 ||
	    viewWindowCommands.unhideWindowBar?.count < 3 ||
	    (Number(viewWindowCommands.unhideWindowUndoRedo?.undo?.hiddenIds?.length || 0) + (viewWindowCommands.unhideWindowUndoRedo?.undo?.mainHidden ? 1 : 0)) < 1 ||
	    viewWindowCommands.unhideWindowUndoRedo?.undo?.bar?.hiddenCount < 1 ||
	    !viewWindowCommands.unhideWindowUndoRedo?.undo?.status?.includes("元に戻しました") ||
	    viewWindowCommands.unhideWindowUndoRedo?.redo?.mainHidden !== false ||
	    viewWindowCommands.unhideWindowUndoRedo?.redo?.hiddenIds?.length !== 0 ||
	    viewWindowCommands.unhideWindowUndoRedo?.redo?.bar?.hiddenCount !== 0 ||
	    viewWindowCommands.unhideWindowUndoRedo?.redo?.bar?.activePreviewId !== viewWindowCommands.unhideWindowUndoRedo?.redo?.activeWindowId ||
	    !viewWindowCommands.unhideWindowUndoRedo?.redo?.status?.includes("やり直しました") ||
	    viewWindowCommands.closeWindowState?.exists !== false ||
	    viewWindowCommands.closeWindowState?.activeWindowId === viewWindowCommands.closeWindowState?.closedId ||
	    viewWindowCommands.closeWindowState?.bar?.count !== viewWindowCommands.closeWindowState?.windowCount ||
	    viewWindowCommands.closeWindowState?.closeButtons !== viewWindowCommands.closeWindowState?.windowCount ||
	    !viewWindowCommands.closeWindowState?.status?.includes("閉じました") ||
	    viewWindowCommands.closeWindowUndoRedo?.undo?.exists !== true ||
	    viewWindowCommands.closeWindowUndoRedo?.undo?.activeWindowId !== viewWindowCommands.closeWindowState?.closedId ||
	    viewWindowCommands.closeWindowUndoRedo?.undo?.bar?.count !== viewWindowCommands.closeWindowUndoRedo?.undo?.windowCount ||
	    !viewWindowCommands.closeWindowUndoRedo?.undo?.status?.includes("元に戻しました") ||
	    viewWindowCommands.closeWindowUndoRedo?.redo?.exists !== false ||
	    viewWindowCommands.closeWindowUndoRedo?.redo?.activeWindowId === viewWindowCommands.closeWindowState?.closedId ||
	    viewWindowCommands.closeWindowUndoRedo?.redo?.bar?.count !== viewWindowCommands.closeWindowUndoRedo?.redo?.windowCount ||
	    !viewWindowCommands.closeWindowUndoRedo?.redo?.status?.includes("やり直しました") ||
	    viewWindowCommands.hiddenWindows !== 0 ||
    viewWindowCommands.switchWindowMenu?.switchLabels?.length < 2 ||
    viewWindowCommands.switchWindowMenu?.switchLabels?.some((label) => !/:\d+(?!\d)/.test(label || "")) ||
    viewWindowCommands.switchWindowMenu?.switchDescriptions?.length < 2 ||
    !viewWindowCommands.switchWindowMenu?.switchDescriptions?.every((text) => text.includes("%")) ||
    !viewWindowCommands.switchWindowMenu?.actions?.includes("switch-dialog") ||
    !viewWindowCommands.switchWindowMenu?.dialogLabel?.includes("その他のウィンドウ") ||
    !viewWindowCommands.switchWindowMenu?.dialogDescription?.includes("一覧") ||
    !viewWindowCommands.switchWindowMenu?.checkedAction?.startsWith("switch:") ||
	    !viewWindowCommands.switchWindowMenu?.accessKeys?.includes("1") ||
	    !viewWindowCommands.switchWindowMenu?.accessKeys?.includes("2") ||
	    !viewWindowCommands.switchWindowMenu?.accessKeys?.includes("M") ||
	    !viewWindowCommands.switchWindowMenu?.accessKeysUnderlined ||
	    viewWindowCommands.switchWindowKeyboardState?.skipped !== false ||
	    viewWindowCommands.switchWindowKeyboardState?.activeWindowId !== viewWindowCommands.switchWindowKeyboardState?.expectedWindowId ||
	    viewWindowCommands.switchWindowKeyboardState?.menuClosed !== true ||
	    !viewWindowCommands.switchWindowKeyboardState?.status?.includes("切り替え") ||
	    viewWindowCommands.switchWindowDialog?.title !== "ウィンドウの切り替え" ||
	    !viewWindowCommands.switchWindowDialog?.caption?.includes("ウィンドウ") ||
	    viewWindowCommands.switchWindowDialog?.options?.length < 2 ||
	    viewWindowCommands.switchWindowDialog?.selected !== viewWindowCommands.switchWindowDialog?.target ||
	    Number(viewWindowCommands.switchWindowDialog?.size || 0) < 3 ||
	    !viewWindowCommands.switchWindowDialog?.summary?.includes("切り替える") ||
	    !viewWindowCommands.switchWindowDialog?.accessKeys?.includes("W") ||
	    !viewWindowCommands.switchWindowDialog?.accessKeysUnderlined ||
	    viewWindowCommands.switchWindowDialogState?.activeWindowId !== viewWindowCommands.switchWindowDialogState?.expectedWindowId ||
	    viewWindowCommands.switchWindowDialogState?.dialogClosed !== true ||
	    !viewWindowCommands.switchWindowDialogState?.status?.includes("切り替え") ||
	    viewWindowCommands.arrangeWindowDialog?.title !== "ウィンドウの整列" ||
	    !viewWindowCommands.arrangeWindowDialog?.options?.includes("vertical") ||
	    !viewWindowCommands.arrangeWindowDialog?.labels?.some((label) => label.includes("左右に並べて表示")) ||
	    viewWindowCommands.arrangeWindowDialog?.activeWorkbookOnlyChecked !== true ||
	    viewWindowCommands.arrangeWindowDialog?.activeWorkbookOnlyDisabled !== true ||
	    !viewWindowCommands.arrangeWindowDialog?.accessKeys?.includes("V") ||
	    !viewWindowCommands.arrangeWindowDialog?.accessKeys?.includes("W") ||
	    !viewWindowCommands.arrangeWindowDialog?.accessKeysUnderlined ||
	    !viewWindowCommands.unhideWindowDialog?.options?.length ||
    !viewWindowCommands.unhideWindowDialog?.accessKeys?.includes("U") ||
    !viewWindowCommands.unhideWindowDialog?.accessKeysUnderlined ||
    viewWindowCommands.unhideWindowDialog?.options?.some((label) => !/:\d+\b/.test(label || "")) ||
    viewWindowCommands.macrosDialog?.title !== "マクロ" ||
    viewWindowCommands.macrosDialog?.initial?.title !== "マクロ" ||
    viewWindowCommands.macrosDialog?.initial?.runDisabled !== true ||
    viewWindowCommands.macrosDialog?.initial?.createDisabled !== false ||
    viewWindowCommands.macrosDialog?.afterCreate?.selectedName !== "ViewMacro" ||
    viewWindowCommands.macrosDialog?.afterCreate?.runDisabled !== false ||
    viewWindowCommands.macrosDialog?.afterCreate?.editDisabled !== false ||
    viewWindowCommands.macrosDialog?.afterCreate?.optionsDisabled !== false ||
    viewWindowCommands.macrosDialog?.createUndoRedo?.undo?.exists !== false ||
    viewWindowCommands.macrosDialog?.createUndoRedo?.undo?.runDisabled !== true ||
    viewWindowCommands.macrosDialog?.createUndoRedo?.redo?.exists !== true ||
    viewWindowCommands.macrosDialog?.createUndoRedo?.redo?.selectedName !== "ViewMacro" ||
    viewWindowCommands.macrosDialog?.options?.title !== "マクロ オプション" ||
    viewWindowCommands.macrosDialog?.options?.name !== "ViewMacro" ||
    viewWindowCommands.macrosDialog?.optionsUndoRedo?.undo?.macro?.shortcut !== "" ||
    viewWindowCommands.macrosDialog?.optionsUndoRedo?.undo?.macro?.description !== "" ||
    viewWindowCommands.macrosDialog?.optionsUndoRedo?.undo?.dialogVisible !== true ||
    viewWindowCommands.macrosDialog?.optionsUndoRedo?.redo?.macro?.shortcut !== "q" ||
    !viewWindowCommands.macrosDialog?.optionsUndoRedo?.redo?.macro?.description?.includes("表示タブ") ||
    viewWindowCommands.macrosDialog?.optionsUndoRedo?.redo?.dialogVisible !== true ||
    !viewWindowCommands.macrosDialog?.editor?.title?.includes("ViewMacro") ||
    viewWindowCommands.macrosDialog?.editorUndoRedo?.undo?.macro?.code !== "" ||
    viewWindowCommands.macrosDialog?.editorUndoRedo?.undo?.dialogVisible !== true ||
    !viewWindowCommands.macrosDialog?.editorUndoRedo?.redo?.macro?.code?.includes("Sub ViewMacro") ||
    viewWindowCommands.macrosDialog?.editorUndoRedo?.redo?.dialogVisible !== true ||
	    viewWindowCommands.macrosDialog?.created?.shortcut !== "q" ||
		    !viewWindowCommands.macrosDialog?.created?.description?.includes("表示タブ") ||
		    !viewWindowCommands.macrosDialog?.created?.code?.includes("Sub ViewMacro") ||
    !viewWindowCommands.macrosDialog?.created?.lastRunAt ||
		    viewWindowCommands.macrosDialog?.runMessage?.text !== "View" ||
    viewWindowCommands.macrosDialog?.runMessage?.i6 !== "MacroRun" ||
    viewWindowCommands.macrosDialog?.runMessage?.j6 !== "42" ||
    viewWindowCommands.macrosDialog?.runMessage?.k6 !== "" ||
    viewWindowCommands.macrosDialog?.runMessage?.i7 !== "Offset" ||
    viewWindowCommands.macrosDialog?.runMessage?.i8 !== "Resize" ||
    viewWindowCommands.macrosDialog?.runMessage?.j8 !== "Resize" ||
    viewWindowCommands.macrosDialog?.runMessage?.i9 !== "ActiveOffset" ||
    viewWindowCommands.macrosDialog?.runMessage?.i10 !== "SelectionOffset" ||
    viewWindowCommands.macrosDialog?.runMessage?.v6 !== "CellsValue" ||
    viewWindowCommands.macrosDialog?.runMessage?.w6 !== "Array1" ||
    viewWindowCommands.macrosDialog?.runMessage?.x6 !== "Array2" ||
    viewWindowCommands.macrosDialog?.runMessage?.y6Model?.formula !== "=J6+1" ||
    viewWindowCommands.macrosDialog?.runMessage?.z6 !== "50" ||
    viewWindowCommands.macrosDialog?.runMessage?.ab6 !== "44" ||
    viewWindowCommands.macrosDialog?.runMessage?.aa6Model?.raw !== "Indexed" ||
    viewWindowCommands.macrosDialog?.runMessage?.aa6Model?.css?.color !== "#FF0000" ||
    viewWindowCommands.macrosDialog?.runMessage?.aa6Model?.css?.backgroundColor !== "#FFFF00" ||
    viewWindowCommands.macrosDialog?.runMessage?.aa6Model?.css?.textAlign !== "center" ||
    viewWindowCommands.macrosDialog?.runMessage?.aa6Model?.css?.verticalAlign !== "bottom" ||
    !String(viewWindowCommands.macrosDialog?.runMessage?.aa6Model?.css?.textDecoration || "").includes("underline") ||
    viewWindowCommands.macrosDialog?.runMessage?.i6Model?.css?.fontWeight !== "700" ||
    viewWindowCommands.macrosDialog?.runMessage?.i6Model?.css?.backgroundColor !== "#FFFF00" ||
    viewWindowCommands.macrosDialog?.runMessage?.i6Model?.css?.textAlign !== "right" ||
    viewWindowCommands.macrosDialog?.runMessage?.i6Model?.css?.whiteSpace !== "pre-wrap" ||
    viewWindowCommands.macrosDialog?.runMessage?.j6Model?.css?.backgroundColor !== "#FFFF00" ||
    viewWindowCommands.macrosDialog?.runMessage?.k6Model?.hasFill === true ||
    viewWindowCommands.macrosDialog?.runMessage?.l6Model !== null ||
    viewWindowCommands.macrosDialog?.runMessage?.m6Model?.raw !== "CopySrc" ||
    viewWindowCommands.macrosDialog?.runMessage?.m6Model?.css?.fontStyle !== "italic" ||
    viewWindowCommands.macrosDialog?.runMessage?.n6Model !== null ||
    viewWindowCommands.macrosDialog?.runMessage?.o6Model?.raw !== "CutSrc" ||
    viewWindowCommands.macrosDialog?.runMessage?.p6Model?.raw !== "Merged" ||
    viewWindowCommands.macrosDialog?.runMessage?.r6Model?.raw !== "Font" ||
    !String(viewWindowCommands.macrosDialog?.runMessage?.r6Model?.css?.fontFamily || "").includes("MS PGothic") ||
    viewWindowCommands.macrosDialog?.runMessage?.r6Model?.css?.fontSize !== "14pt" ||
    !String(viewWindowCommands.macrosDialog?.runMessage?.r6Model?.css?.textDecoration || "").includes("line-through") ||
    viewWindowCommands.macrosDialog?.runMessage?.r6Model?.borders?.top?.width !== 3 ||
    viewWindowCommands.macrosDialog?.runMessage?.r6Model?.borders?.top?.color !== "#FF0000" ||
    viewWindowCommands.macrosDialog?.runMessage?.s6Model?.borders?.left?.style !== "solid" ||
    Number(viewWindowCommands.macrosDialog?.runMessage?.tColumn?.width || 0) <= 70 ||
    viewWindowCommands.macrosDialog?.runMessage?.uColumn?.hidden !== false ||
    Number(viewWindowCommands.macrosDialog?.runMessage?.row11?.height || 0) <= 25 ||
    viewWindowCommands.macrosDialog?.runMessage?.row12?.hidden !== false ||
    viewWindowCommands.macrosDialog?.runMessage?.pqMerged !== false ||
    viewWindowCommands.macrosDialog?.runMessage?.selectionLabel !== "I6:J6" ||
    viewWindowCommands.macrosDialog?.runMessage?.dialogClosed !== true ||
    !viewWindowCommands.macrosDialog?.runMessage?.status?.includes("実行") ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.undo?.i6 === "MacroRun" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.undo?.i6Model?.css?.backgroundColor === "#FFFF00" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.undo?.r6Model?.raw === "Font" ||
    Number(viewWindowCommands.macrosDialog?.runUndoRedo?.undo?.tColumn?.width || 0) > 70 ||
    !rowHeightRestored(viewWindowCommands.macrosDialog?.runBefore?.row11, viewWindowCommands.macrosDialog?.runUndoRedo?.undo?.row11) ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.undo?.m6Model?.raw === "CopySrc" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.undo?.o6Model?.raw === "CutSrc" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.undo?.p6Model?.raw === "Merged" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.undo?.v6Model?.raw === "CellsValue" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.undo?.y6Model?.formula === "=J6+1" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.undo?.aa6Model?.raw === "Indexed" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.undo?.lastRunAt ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.i6 !== "MacroRun" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.j6 !== "42" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.k6 !== "" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.i7 !== "Offset" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.i8 !== "Resize" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.j8 !== "Resize" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.i9 !== "ActiveOffset" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.i10 !== "SelectionOffset" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.v6 !== "CellsValue" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.w6 !== "Array1" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.x6 !== "Array2" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.y6Model?.formula !== "=J6+1" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.z6 !== "50" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.ab6 !== "44" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.aa6Model?.raw !== "Indexed" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.aa6Model?.css?.color !== "#FF0000" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.aa6Model?.css?.backgroundColor !== "#FFFF00" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.aa6Model?.css?.textAlign !== "center" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.aa6Model?.css?.verticalAlign !== "bottom" ||
    !String(viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.aa6Model?.css?.textDecoration || "").includes("underline") ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.i6Model?.css?.fontWeight !== "700" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.i6Model?.css?.backgroundColor !== "#FFFF00" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.i6Model?.css?.textAlign !== "right" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.i6Model?.css?.whiteSpace !== "pre-wrap" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.k6Model?.hasFill === true ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.l6Model !== null ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.m6Model?.raw !== "CopySrc" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.m6Model?.css?.fontStyle !== "italic" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.n6Model !== null ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.o6Model?.raw !== "CutSrc" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.p6Model?.raw !== "Merged" ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.r6Model?.raw !== "Font" ||
    !String(viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.r6Model?.css?.fontFamily || "").includes("MS PGothic") ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.r6Model?.borders?.top?.width !== 3 ||
    Number(viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.tColumn?.width || 0) <= 70 ||
    Number(viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.row11?.height || 0) <= 25 ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.pqMerged !== false ||
    !viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.lastRunAt ||
    viewWindowCommands.macrosDialog?.runUndoRedo?.redo?.selectionLabel !== "I6:J6" ||
	    !viewWindowCommands.macrosDialog?.stepDialog?.title?.includes("ViewMacro") ||
    !viewWindowCommands.macrosDialog?.stepDialog?.summary?.includes("1 /") ||
    !viewWindowCommands.macrosDialog?.stepDialog?.current?.includes("Range(\"I6\")") ||
    viewWindowCommands.macrosDialog?.stepDialog?.count < 10 ||
    viewWindowCommands.macrosDialog?.stepDialog?.nextDisabled !== false ||
    !viewWindowCommands.macrosDialog?.stepDialog?.afterFirst?.summary?.includes("2 /") ||
    viewWindowCommands.macrosDialog?.stepDialog?.afterFirst?.completeCount !== 1 ||
    !viewWindowCommands.macrosDialog?.stepDialog?.afterFirst?.current?.includes("Range(\"J6\")") ||
    viewWindowCommands.macrosDialog?.stepDialog?.afterContinue?.nextDisabled !== true ||
    viewWindowCommands.macrosDialog?.stepDialog?.afterContinue?.continueDisabled !== true ||
    !viewWindowCommands.macrosDialog?.stepDialog?.afterContinue?.log?.includes("View") ||
    !viewWindowCommands.macrosDialog?.stepDialog?.afterContinue?.summary?.includes("ステートメントを実行") ||
	    !viewWindowCommands.macroMenuBeforeRecord?.labels?.some((label) => label.includes("マクロの記録")) ||
    !viewWindowCommands.macroMenuBeforeRecord?.descriptions?.some((text) => text.includes("操作を記録")) ||
    !viewWindowCommands.macroMenuBeforeRecord?.descriptions?.some((text) => text.includes("現在:") && text.includes("警告")) ||
    viewWindowCommands.macroRecordDialog?.title !== "マクロの記録" ||
    !viewWindowCommands.macroRecordDialog?.defaultName?.includes("Macro") ||
    !viewWindowCommands.macroRecordDialog?.scopeOptions?.includes("作業中のブック") ||
    viewWindowCommands.macroRecordingStarted?.recording?.name !== "RecordedViewMacro" ||
    viewWindowCommands.macroRecordingStarted?.macro?.shortcut !== "r" ||
    !viewWindowCommands.macroRecordingStarted?.macro?.description?.includes("表示タブ") ||
    viewWindowCommands.macroRecordingStarted?.macroButtonActive !== true ||
    viewWindowCommands.macroRecordingStarted?.menuButtonActive !== true ||
    !viewWindowCommands.macroMenuDuringRecord?.labels?.some((label) => label.includes("記録終了")) ||
    viewWindowCommands.macroMenuDuringRecord?.relativeChecked !== "false" ||
    viewWindowCommands.macroMenuRelativeRefs?.relativeChecked !== "true" ||
    viewWindowCommands.macroRelativeRefsUndoRedo?.undo?.relativeReferences !== false ||
    viewWindowCommands.macroRelativeRefsUndoRedo?.undo?.recordingRelativeReferences !== false ||
    viewWindowCommands.macroRelativeRefsUndoRedo?.redo?.relativeReferences !== true ||
    viewWindowCommands.macroRelativeRefsUndoRedo?.redo?.recordingRelativeReferences !== true ||
    viewWindowCommands.macroRecordingStopped?.recording !== null ||
    viewWindowCommands.macroRecordingStopped?.relativeReferences !== true ||
    !viewWindowCommands.macroRecordingStopped?.status?.includes("記録を終了") ||
    viewWindowCommands.macroRecordingStopped?.macroCount < 2 ||
    viewWindowCommands.macroRecordingStopUndoRedo?.undo?.recording?.name !== "RecordedViewMacro" ||
    viewWindowCommands.macroRecordingStopUndoRedo?.undo?.macroButtonActive !== true ||
    viewWindowCommands.macroRecordingStopUndoRedo?.redo?.recording !== null ||
    viewWindowCommands.macroRecordingStopUndoRedo?.redo?.macroButtonActive !== false ||
    viewWindowCommands.macroSecurityDialog?.title !== "セキュリティ センター" ||
    viewWindowCommands.macroSecurityDialog?.navText !== "マクロの設定" ||
    viewWindowCommands.macroSecurityDialog?.navItems?.length < 8 ||
    !viewWindowCommands.macroSecurityDialog?.navItems?.includes("信頼できる場所") ||
    !viewWindowCommands.macroSecurityDialog?.navItems?.includes("外部コンテンツ") ||
    viewWindowCommands.macroSecurityDialog?.options?.length !== 4 ||
    !viewWindowCommands.macroSecurityDialog?.options?.some((option) => option.value === "disableWithNotification" && option.checked) ||
    !viewWindowCommands.macroSecurityDialog?.options?.some((option) => option.label.includes("すべての VBA マクロを有効")) ||
    !viewWindowCommands.macroSecurityDialog?.options?.every((option) => option.description) ||
    !viewWindowCommands.macroSecurityDialog?.warning?.includes("VBA マクロ") ||
    viewWindowCommands.macroSecurityApplied?.security?.level !== "enableAll" ||
    viewWindowCommands.macroSecurityApplied?.security?.trustVbaObjectModel !== true ||
    !viewWindowCommands.macroSecurityApplied?.security?.updatedAt ||
    !viewWindowCommands.macroSecurityApplied?.status?.includes("マクロのセキュリティ") ||
    viewWindowCommands.macroSecurityUndoRedo?.undo?.security?.level !== "disableWithNotification" ||
    viewWindowCommands.macroSecurityUndoRedo?.undo?.security?.trustVbaObjectModel !== false ||
    viewWindowCommands.macroSecurityUndoRedo?.redo?.security?.level !== "enableAll" ||
    viewWindowCommands.macroSecurityUndoRedo?.redo?.security?.trustVbaObjectModel !== true
  ) {
    throw new Error("View ribbon should expose Excel-like display toggles, workbook views, zoom controls, and window commands.");
  }
  Object.assign(scenarioState, { result, pageLayout, expectedReloadWorkbookName, column, row });
}

export {
  runAssertionsEditor,
};
