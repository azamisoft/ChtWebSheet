// Sequential E2E feature module.
import {
  path,
} from "../../e2e-real-upload.mjs";
async function runAssertionsWorkbook(scenarioState) {
  let { result, delayedLazyLoad, stickyHeaders, sheetMenu, sheetTabColor, sheetRename, sheetTabDirectActions, sheetMoveCopy, sheetCodeDialog, sheetSelectAll, sheetHidden, sheetUnhideDialog, sheetUnhidden, sheetUnhideUndo, sheetUnhideRedo, sheetGroupActions, sheetGroupMoveCopy, sheetGroupCellEditing, sheetGroupStructure, sheetGroupFormatDialog, sheetGroupDataValidation, sheetGroupDataFilterSort, sheetGroupDataTools, sheetGroupHyperlink, sheetGroupHomeCommands, sheetGroupTranslation, sheetGroupQuickAnalysis, sheetGroupSparklines, sheetGroupCharts, sheetGroupTables, sheetGroupPivot, sheetGroupObjects, layoutRibbon, layoutThemeMenuInitial, layoutThemeApplied, layoutColorMenuInitial, layoutFontMenuInitial, layoutEffectMenuInitial, layoutThemeAfterMenus, layoutThemeEffectUndoRedo, layoutPrintAreaMenuBefore, layoutPrintAreaMenuAfterSet, layoutPrintAreaMenuAfterClear, layoutMarginsMenuInitial, layoutOrientationMenuInitial, layoutSizeMenuInitial, layoutBackgroundApplied, layoutBackgroundUndo, layoutBackgroundRedo, layoutBackgroundDeleted, layoutBackgroundDeleteUndo, pageSetupDialog, printTitlesDialogInitial, printTitlesDialogFilled, pageLayout, pageLayoutPreview, pageBreakMenuKeyboard, pageBreakDrag, drawRibbon, drawInk, drawConverted, headerFooterDialog, headerFooterApplied, headerFooterUndo, headerFooterRedo, insertMapChart, insertShapeMenu, insertShapeDraw, insertIconDialog, signatureLineDialog, insertObjectDialog, insertAddinsStoreDialog, insertAddinsMineDialog, insertTextObjects, fileSideMenu, fileSideMenuClosedByOutsideClick, fileSideMenuClosedByCommand, shortcutSave, saveAsDialog, singleClickEdit, doubleClickEdit, editAfterDeferredRender, editAlignment, editClickCaret, editTextSelection, editNoEllipsis, cellEditAltEnter, cellEditAltEnterDuringIme, editDragOutsideSelection, cellEditImeEnter, editScrollbar, editExitRender, multilineCommitDisplay, formulaEmptyAddress, formulaBarInitial, formulaBarSizing, formulaBarDuringEdit, formulaBarEscape, formulaBarEnter, formulaAutocomplete, formulaCustomFunctions, formulaReferenceInsert, directTyping, nameLinkAddress, absoluteNameLinkAddress, nameDialogInitial, definedNameFormula, unicodeDefinedNameFormula, duplicateDefinedName, localDefinedNameFormula, linkDialogInitial, linkDocumentKind, linkDocumentSheetChoice, linkEmailKind, linkWebKind, linkApply, linkContextMenu, openedLink, linkUndo, linkRedo, definedNameAndLink, row, underlineToolbar, borderToolbar, customViewImportExport } = scenarioState;

  if (
    customViewImportExport.imported?.name !== "ImportedView" ||
    customViewImportExport.imported?.guid !== "{11111111-2222-3333-4444-555555555555}" ||
    customViewImportExport.imported?.autoUpdate !== true ||
    customViewImportExport.imported?.mergeInterval !== 17 ||
    customViewImportExport.imported?.changesSavedWin !== true ||
    customViewImportExport.imported?.onlySync !== true ||
    customViewImportExport.imported?.personalView !== true ||
    customViewImportExport.imported?.showFormulaBar !== false ||
    customViewImportExport.imported?.activeSheetIndex !== 1 ||
    customViewImportExport.imported?.activeScroll?.top !== 7 ||
    customViewImportExport.imported?.activeScroll?.left !== 5 ||
    customViewImportExport.imported?.workbookProperties?.date1904 !== true ||
    customViewImportExport.imported?.workbookProperties?.codeName !== "ThisWorkbook" ||
    customViewImportExport.imported?.workbookProperties?.filterPrivacy !== true ||
    customViewImportExport.imported?.calcProperties?.calcMode !== "manual" ||
    customViewImportExport.imported?.calcProperties?.refMode !== "R1C1" ||
    customViewImportExport.imported?.calcProperties?.iterate !== true ||
    customViewImportExport.imported?.calcProperties?.iterateCount !== 42 ||
    Math.abs((customViewImportExport.imported?.calcProperties?.iterateDelta || 0) - 0.0005) > 0.000001 ||
    Math.abs((customViewImportExport.imported?.date1904Cell?.raw || 0) - 38384.375) > 0.000001 ||
    customViewImportExport.imported?.workbookView?.activeSheetIndex !== 0 ||
    customViewImportExport.imported?.mainWindowView?.sheetIndex !== 0 ||
    customViewImportExport.imported?.mainWindowView?.x !== 10 ||
    customViewImportExport.imported?.mainWindowView?.width !== 24000 ||
    customViewImportExport.imported?.customMainWindowView?.showHorizontalScroll !== false ||
    customViewImportExport.imported?.customMainWindowView?.showVerticalScroll !== false ||
    customViewImportExport.imported?.customMainWindowView?.showSheetTabs !== false ||
    customViewImportExport.imported?.windowViews?.length < 1 ||
    customViewImportExport.imported?.windowViews?.[0]?.sheetIndex !== 1 ||
    customViewImportExport.imported?.windowViews?.[0]?.hidden !== true ||
    customViewImportExport.imported?.windowViews?.[0]?.x !== 410 ||
    customViewImportExport.imported?.windowViews?.[0]?.width !== 18000 ||
    customViewImportExport.imported?.activeWindowId !== "main" ||
    customViewImportExport.imported?.sheetA?.showGridLines !== false ||
    customViewImportExport.imported?.sheetA?.showHeadings !== false ||
    customViewImportExport.imported?.sheetA?.showFormulas !== true ||
    customViewImportExport.imported?.sheetA?.showZeros !== false ||
    customViewImportExport.imported?.sheetA?.showWhiteSpace !== true ||
    customViewImportExport.imported?.sheetA?.rightToLeft !== false ||
    customViewImportExport.imported?.sheetA?.activeScroll?.top !== 4 ||
    customViewImportExport.imported?.sheetA?.activeScroll?.left !== 3 ||
    customViewImportExport.imported?.sheetA?.selectionRange?.top !== 2 ||
    customViewImportExport.imported?.sheetA?.selectionRange?.right !== 3 ||
    customViewImportExport.imported?.sheetB?.viewMode !== "page-layout" ||
    customViewImportExport.imported?.sheetB?.zoomScale !== 125 ||
    customViewImportExport.imported?.sheetB?.showWhiteSpace !== false ||
    customViewImportExport.imported?.sheetB?.rightToLeft !== true ||
    customViewImportExport.imported?.sheetB?.freezePanes?.rows !== 2 ||
    customViewImportExport.imported?.sheetB?.freezePanes?.columns !== 1 ||
    customViewImportExport.imported?.sheetB?.selectedCell?.row !== 4 ||
    customViewImportExport.imported?.sheetB?.selectedCell?.col !== 4 ||
    customViewImportExport.dialog?.title !== "ユーザー設定のビュー" ||
    !customViewImportExport.dialog?.names?.includes("ImportedView") ||
    customViewImportExport.dialog?.showDisabled !== false ||
    customViewImportExport.applied?.activeSheetName !== "CustomB" ||
    customViewImportExport.applied?.selectionLabel !== "D4" ||
    customViewImportExport.applied?.showFormulaBar !== false ||
    customViewImportExport.applied?.viewMode !== "page-layout" ||
    customViewImportExport.applied?.zoom !== 125 ||
    customViewImportExport.applied?.showWhiteSpace !== false ||
    customViewImportExport.applied?.rightToLeft !== true ||
    customViewImportExport.applied?.freezePanes?.rows !== 2 ||
    customViewImportExport.applied?.freezePanes?.columns !== 1 ||
    customViewImportExport.applied?.mainWindowView?.showHorizontalScroll !== false ||
    customViewImportExport.applied?.mainWindowView?.showVerticalScroll !== false ||
    customViewImportExport.applied?.mainWindowView?.showSheetTabs !== false ||
    customViewImportExport.applied?.hideHorizontalScrollbar !== true ||
    customViewImportExport.applied?.hideVerticalScrollbar !== true ||
    customViewImportExport.applied?.hideSheetTabs !== true ||
    customViewImportExport.applied?.scrollTop < 1 ||
    customViewImportExport.applied?.scrollLeft < 1 ||
    !customViewImportExport.applied?.status?.includes("ユーザー設定のビュー") ||
    Object.values(customViewImportExport.exported || {}).some((value) => value !== true) ||
    customViewImportExport.deletedExport?.stateCount !== 0 ||
    customViewImportExport.deletedExport?.stateNames?.length !== 0 ||
    !customViewImportExport.deletedExport?.status?.includes("削除") ||
    customViewImportExport.deletedExport?.workbookCustomViewsRemoved !== true ||
    customViewImportExport.deletedExport?.sheetACustomViewsRemoved !== true ||
    customViewImportExport.deletedExport?.sheetBCustomViewsRemoved !== true ||
    customViewImportExport.deletedExport?.workbookPrKept !== true ||
    customViewImportExport.deletedExport?.calcPrKept !== true
  ) {
    throw new Error("Imported Excel custom views should restore selection, scroll, sheet display state, and survive Save As XLSX.");
  }


  if (
    layoutThemeMenuInitial.label !== "テーマ" ||
    layoutThemeMenuInitial.items?.length < 8 ||
    layoutThemeMenuInitial.descriptionCount < 8 ||
    !layoutThemeMenuInitial.items?.some((item) => item.action === "theme:office" && item.checked && item.role === "menuitemradio") ||
    layoutThemeApplied.model?.theme !== "parcel" ||
    layoutThemeApplied.model?.colorScheme !== "parcel" ||
    layoutThemeApplied.model?.fontScheme !== "parcel" ||
    layoutThemeApplied.model?.effectScheme !== "parcel" ||
    layoutThemeApplied.themeColors?.length < 10 ||
    layoutThemeApplied.themeFonts?.body !== "MS Pゴシック" ||
    layoutThemeApplied.themeEffects?.scheme !== "parcel" ||
    !layoutThemeApplied.themeEffects?.shape?.shadow ||
    !layoutThemeApplied.themeEffects?.shape?.reflection ||
    !layoutThemeApplied.status?.includes("テーマ") ||
    layoutColorMenuInitial.label !== "配色" ||
    layoutColorMenuInitial.descriptionCount < 8 ||
    !layoutColorMenuInitial.items?.some((item) => item.action === "colors:parcel" && item.checked && item.role === "menuitemradio") ||
    layoutFontMenuInitial.label !== "フォント" ||
    layoutFontMenuInitial.descriptionCount < 8 ||
    !layoutFontMenuInitial.items?.some((item) => item.action === "fonts:parcel" && item.checked && item.role === "menuitemradio") ||
    layoutEffectMenuInitial.label !== "効果" ||
    layoutEffectMenuInitial.descriptionCount < 8 ||
    !layoutEffectMenuInitial.items?.some((item) => item.action === "effects:parcel" && item.checked && item.role === "menuitemradio") ||
    !layoutEffectMenuInitial.items?.some((item) => item.action === "effects:flat" && item.hasDescription) ||
    layoutThemeAfterMenus.model?.theme !== "parcel" ||
    layoutThemeAfterMenus.model?.colorScheme !== "facet" ||
    layoutThemeAfterMenus.model?.fontScheme !== "wisp" ||
    layoutThemeAfterMenus.model?.effectScheme !== "retrospect" ||
    layoutThemeAfterMenus.themeFonts?.body !== "Calibri" ||
    layoutThemeAfterMenus.themeEffects?.scheme !== "retrospect" ||
    !layoutThemeAfterMenus.themeEffects?.shape?.glow ||
    layoutThemeEffectUndoRedo.undo?.model?.effectScheme !== "parcel" ||
    layoutThemeEffectUndoRedo.undo?.themeEffects?.scheme !== "parcel" ||
    !layoutThemeEffectUndoRedo.undo?.themeEffects?.shape?.reflection ||
    layoutThemeEffectUndoRedo.redo?.model?.effectScheme !== "retrospect" ||
    layoutThemeEffectUndoRedo.redo?.themeEffects?.scheme !== "retrospect" ||
    !layoutThemeEffectUndoRedo.redo?.themeEffects?.shape?.glow ||
    !layoutBackgroundApplied.background?.src?.startsWith("data:image/png") ||
    layoutBackgroundApplied.background?.name !== "page-background.png" ||
    layoutBackgroundApplied.buttonText !== "背景の削除" ||
    layoutBackgroundApplied.buttonPressed !== "true" ||
    layoutBackgroundApplied.layerName !== "page-background.png" ||
    !layoutBackgroundApplied.layerImage?.includes("data:image/png") ||
    layoutBackgroundUndo.background !== null ||
    layoutBackgroundUndo.buttonText !== "背景" ||
    layoutBackgroundUndo.layerCount !== 0 ||
    !layoutBackgroundRedo.background?.src?.startsWith("data:image/png") ||
    layoutBackgroundRedo.buttonText !== "背景の削除" ||
    layoutBackgroundRedo.layerCount < 1 ||
    layoutBackgroundDeleted.background !== null ||
    layoutBackgroundDeleted.buttonText !== "背景" ||
    layoutBackgroundDeleted.layerCount !== 0 ||
    !layoutBackgroundDeleteUndo.background?.src?.startsWith("data:image/png") ||
    layoutBackgroundDeleteUndo.buttonText !== "背景の削除" ||
    layoutBackgroundDeleteUndo.layerCount < 1 ||
    String(layoutThemeAfterMenus.themeColors?.[0] || "").toLowerCase() !== "#4472c4" ||
    !layoutPrintAreaMenuBefore.items?.some((item) => item.action === "print-area:set" && item.hasDescription && !item.disabled) ||
    !layoutPrintAreaMenuAfterSet.items?.some((item) => item.action === "print-area:clear" && item.hasDescription && !item.disabled && item.text.includes("B1:AP4")) ||
    !layoutPrintAreaMenuAfterClear.items?.some((item) => item.action === "print-area:clear" && item.disabled && item.hasDescription) ||
    layoutMarginsMenuInitial.descriptionCount < 3 ||
    !layoutMarginsMenuInitial.items?.some((item) => item.action === "margins:normal" && item.checked && item.role === "menuitemradio") ||
    layoutOrientationMenuInitial.descriptionCount < 2 ||
    !layoutOrientationMenuInitial.items?.some((item) => item.checked && item.role === "menuitemradio") ||
    layoutSizeMenuInitial.descriptionCount < 4 ||
    !layoutSizeMenuInitial.items?.some((item) => item.checked && item.role === "menuitemradio")
  ) {
    throw new Error("Page Layout preset menus should show Excel-like checked states, descriptions, and disabled print-area actions.");
  }


  if (
    layoutRibbon.activeRibbonTab !== "ページレイアウト" ||
    !layoutRibbon.panelVisible ||
    ["テーマ", "ページ設定", "拡大縮小印刷", "シートのオプション", "配置"].some((label) => !layoutRibbon.groupLabels.includes(label)) ||
    [
      "layout-themes",
      "layout-effects",
      "layout-margins",
      "layout-orientation",
      "layout-size",
      "layout-print-area",
      "layout-breaks",
      "layout-background",
      "layout-print-titles",
      "layout-page-setup",
    ].some((command) => !layoutRibbon.enabledCommands.includes(command)) ||
    layoutRibbon.panelHeight > 72 ||
    pageSetupDialog.title !== "ページ設定" ||
    !["ページ", "余白", "ヘッダー/フッター", "シート"].every((tab) => pageSetupDialog.tabs.includes(tab)) ||
    pageSetupDialog.activePanel !== "page" ||
    !["orientation", "paperSize", "scale", "printArea", "printTitleRows", "printTitleColumns", "firstPageNumber", "pageOrder", "comments", "cellErrors"].every((field) => pageSetupDialog.fields.includes(field)) ||
    !pageSetupDialog.rangePickers?.some((picker) => picker.target === "printArea" && picker.mode === "range") ||
    !pageSetupDialog.rangePickers?.some((picker) => picker.target === "printTitleRows" && picker.mode === "rows") ||
    !pageSetupDialog.rangePickers?.some((picker) => picker.target === "printTitleColumns" && picker.mode === "columns") ||
    !["printGridLines", "printHeadings", "centerHorizontal", "centerVertical", "blackAndWhite", "draftQuality"].every((check) => pageSetupDialog.checks.includes(check)) ||
    printTitlesDialogInitial.activePanel !== "sheet" ||
    printTitlesDialogInitial.focusedField !== "printTitleRows" ||
    printTitlesDialogInitial.printArea !== "B1:AP4" ||
    printTitlesDialogFilled.titleRows !== "$6:$6" ||
    printTitlesDialogFilled.titleColumns !== "$D:$E" ||
    pageLayout.model?.orientation !== "landscape" ||
    pageLayout.model?.paperSize !== "Letter" ||
    pageLayout.model?.marginPreset !== "narrow" ||
    pageLayout.model?.scale !== 95 ||
    pageLayout.model?.fitToWidth !== "1" ||
    pageLayout.model?.firstPageNumber !== 3 ||
    pageLayout.model?.pageOrder !== "overThenDown" ||
    pageLayout.model?.comments !== "atEnd" ||
    pageLayout.model?.cellErrors !== "blank" ||
    pageLayout.model?.centerHorizontal !== true ||
    pageLayout.model?.centerVertical !== true ||
    pageLayout.model?.blackAndWhite !== true ||
    pageLayout.model?.draftQuality !== true ||
    pageLayout.model?.headerFooter?.header !== "Layout Header" ||
    pageLayout.model?.headerFooter?.footer !== "Layout Footer" ||
    pageLayout.headerFooter?.oddHeader?.center !== "Layout Header" ||
    pageLayout.headerFooter?.oddFooter?.center !== "Layout Footer" ||
    pageLayout.model?.printGridLines !== true ||
    pageLayout.model?.printArea?.label !== "B1:AP4" ||
    pageLayout.model?.printTitles?.rows?.top !== 6 ||
    pageLayout.model?.printTitles?.rows?.bottom !== 6 ||
    pageLayout.model?.printTitles?.rows?.label !== "$6:$6" ||
    pageLayout.model?.printTitles?.cols?.left !== 4 ||
    pageLayout.model?.printTitles?.cols?.right !== 5 ||
    pageLayout.model?.printTitles?.cols?.label !== "$D:$E" ||
    pageLayout.gridlinesPrintChecked !== true ||
    pageLayout.scaleControl !== "95" ||
    pageLayoutPreview.viewMode !== "page-layout" ||
    pageLayoutPreview.pageCount < 1 ||
    pageLayoutPreview.marginCount < 1 ||
    pageLayoutPreview.headerFooterCount < 1 ||
    pageLayoutPreview.printAreaCount < 1 ||
    pageLayoutPreview.titleRowCount < 1 ||
    pageLayoutPreview.titleColumnCount < 1 ||
    pageBreakMenuKeyboard.label !== "改ページ" ||
    !pageBreakMenuKeyboard.accessKeys?.includes("I") ||
    !pageBreakMenuKeyboard.accessKeys?.includes("R") ||
    !pageBreakMenuKeyboard.accessKeys?.includes("A") ||
    pageBreakMenuKeyboard.h8BeforeShortcut !== "BreakMenuSeed" ||
    pageBreakMenuKeyboard.h8AfterShortcut !== "BreakMenuSeed" ||
    !pageBreakMenuKeyboard.menuClosed ||
    !pageBreakDrag.before?.cols?.includes(8) ||
    pageBreakDrag.after?.cols?.includes(8) ||
    !pageBreakDrag.after?.cols?.some((col) => col !== 8) ||
    !pageBreakDrag.movedColumn ||
    pageBreakDrag.afterRemove?.cols?.includes(pageBreakDrag.movedColumn) ||
    !pageBreakDrag.statusAfterRemove?.includes("解除") ||
    !pageBreakDrag.undoRemove?.pageBreaks?.cols?.includes(pageBreakDrag.movedColumn) ||
    !pageBreakDrag.undoRemove?.status?.includes("元に戻しました") ||
    !pageBreakDrag.undoRemove?.status?.includes("改ページの解除") ||
    pageBreakDrag.redoRemove?.pageBreaks?.cols?.includes(pageBreakDrag.movedColumn) ||
    !pageBreakDrag.redoRemove?.status?.includes("やり直しました") ||
    !pageBreakDrag.redoRemove?.status?.includes("改ページの解除") ||
    pageBreakDrag.manualGuideCount < 1 ||
    !pageBreakDrag.status?.includes("改ページ")
  ) {
    throw new Error("Page Layout ribbon should expose Excel-like page setup, print area, scaling, and sheet option controls.");
  }


  if (
    drawRibbon.activeRibbonTab !== "描画" ||
    !drawRibbon.panelVisible ||
    ["描画ツール", "ペン", "変換"].some((label) => !drawRibbon.groupLabels.includes(label)) ||
    ["draw-select", "draw-pen", "draw-pencil", "draw-highlighter", "draw-eraser", "draw-lasso", "draw-convert-shape", "draw-clear-ink"].some((command) => !drawRibbon.enabledCommands.includes(command)) ||
    drawRibbon.panelHeight > 72 ||
    drawInk.inkCount < 1 ||
    drawInk.selectedFrames < 1 ||
    drawInk.outlineColor !== "#2563eb" ||
    drawInk.customPathPoints < 3 ||
    drawInk.activePen !== true ||
    drawInk.activeColor !== true ||
    drawConverted.customPaths < 1 ||
    drawConverted.remainingInk !== 0 ||
    !drawConverted.status.includes("インク")
  ) {
    throw new Error("Draw ribbon should create ink, preserve active pen/color state, convert ink to a freeform shape, and clear remaining ink.");
  }


  if (
    headerFooterDialog.title !== "ヘッダーとフッター" ||
    !["ヘッダー", "フッター", "オプション"].every((label) => headerFooterDialog.legends.includes(label)) ||
    !["oddHeader.center", "oddFooter.right"].every((field) => headerFooterDialog.fields.includes(field)) ||
    headerFooterDialog.values?.["oddHeader.center"] !== "Layout Header" ||
    headerFooterDialog.values?.["oddFooter.center"] !== "Layout Footer" ||
    !headerFooterDialog.checks.includes("differentFirst") ||
    headerFooterApplied.model?.oddHeader?.center !== "WebSheet Header" ||
    headerFooterApplied.model?.oddFooter?.right !== "Page &[Page]" ||
    headerFooterApplied.pageLayoutHeaderFooter?.header !== "WebSheet Header" ||
    headerFooterApplied.pageLayoutHeaderFooter?.footer !== "Layout Footer Page &[Page]" ||
    headerFooterApplied.model?.differentFirst !== true ||
    !headerFooterApplied.status.includes("ヘッダーとフッター") ||
    headerFooterUndo.model?.oddHeader?.center !== "Layout Header" ||
    headerFooterUndo.model?.oddFooter?.center !== "Layout Footer" ||
    headerFooterUndo.pageLayoutHeaderFooter?.header !== "Layout Header" ||
    headerFooterUndo.pageLayoutHeaderFooter?.footer !== "Layout Footer" ||
    !headerFooterUndo.status.includes("元に戻しました") ||
    headerFooterRedo.model?.oddHeader?.center !== "WebSheet Header" ||
    headerFooterRedo.pageLayoutHeaderFooter?.header !== "WebSheet Header" ||
    headerFooterRedo.pageLayoutHeaderFooter?.footer !== "Layout Footer Page &[Page]" ||
    !headerFooterRedo.status.includes("やり直しました")
  ) {
    throw new Error("Insert Header/Footer should open an Excel-like dialog, save sheet settings, and participate in undo/redo.");
  }


  if (
    insertMapChart.countDelta !== 1 ||
    insertMapChart.chartType !== "map" ||
    !String(insertMapChart.chartRange || "").startsWith("A1:") ||
    !insertMapChart.hasMapSvg ||
    !insertMapChart.hasRegion
  ) {
    throw new Error("Insert map chart should create a distinct map chart SVG from the selected data range.");
  }


  if (
    insertIconDialog?.title !== "アイコン" ||
    insertIconDialog?.searchPlaceholder !== "アイコンを検索" ||
    insertIconDialog?.choiceCount < 6 ||
    insertIconDialog?.visibleAfterSearch?.join("|") !== "チェック" ||
    insertIconDialog?.emptyHidden !== true ||
    !insertIconDialog?.labels?.some((label) => label.includes("チェック")) ||
    signatureLineDialog?.title !== "署名の設定" ||
    !signatureLineDialog?.captions?.some((label) => label.includes("署名候補者")) ||
    signatureLineDialog?.signer !== "署名者" ||
    signatureLineDialog?.allowComments !== true ||
    signatureLineDialog?.showDate !== true ||
    insertObjectDialog?.title !== "オブジェクトの挿入" ||
    !insertObjectDialog?.types?.some((label) => label.includes("Cht WebSheet")) ||
    !String(insertObjectDialog?.activeTab || "").includes("新規作成") ||
    !String(insertObjectDialog?.activeTabAfterFile || "").includes("ファイルから") ||
    insertObjectDialog?.filePanelActive !== true ||
    insertObjectDialog?.createSelectDisabled !== true ||
    insertObjectDialog?.fileName !== "sample-report.xlsx" ||
    insertObjectDialog?.displayAsIconAfter !== true ||
    insertObjectDialog?.iconLabelEnabled !== true ||
    insertObjectDialog?.preview !== "sample-report.xlsx" ||
    insertAddinsStoreDialog?.title !== "Office アドイン" ||
    insertAddinsStoreDialog?.activeTab !== "ストア" ||
    !insertAddinsStoreDialog?.cards?.includes("websheet-mini-dashboard") ||
    insertAddinsMineDialog?.title !== "個人用アドイン" ||
    insertAddinsMineDialog?.activeTab !== "個人用アドイン" ||
    !insertAddinsMineDialog?.cards?.includes("websheet-bing-maps")
  ) {
    throw new Error("Insert icons and Office add-ins should open Excel-like pickers with selectable content.");
  }


  if (
    insertTextObjects.countDelta < 12 ||
    !["アイコン", "3D モデル", "SmartArt", "スクリーンショット", "ワードアート", "署名欄", "オブジェクト", "数式", "スライサー", "タイムライン", "Mini Dashboard", "Bing Maps"].every((name) => insertTextObjects.names.includes(name)) ||
    insertTextObjects.icon?.objectKind !== "icon" ||
    insertTextObjects.icon?.iconKind !== "check" ||
    insertTextObjects.model3d?.objectKind !== "3dModel" ||
    insertTextObjects.smartArt?.objectKind !== "smartArt" ||
    insertTextObjects.screenshot?.objectKind !== "screenshot" ||
    insertTextObjects.wordArt?.text !== "ワードアート" ||
    insertTextObjects.wordArt?.textBold !== true ||
    insertTextObjects.wordArt?.textItalic !== true ||
    !String(insertTextObjects.signature?.text || "").includes("ZJ") ||
    insertTextObjects.signature?.objectKind !== "signatureLine" ||
    insertTextObjects.signature?.objectMetadata?.email !== "zj@example.com" ||
    insertTextObjects.object?.objectKind !== "embeddedObject" ||
    !String(insertTextObjects.object?.text || "").includes("埋め込み") ||
    insertTextObjects.object?.objectMetadata?.mode !== "file" ||
    insertTextObjects.object?.objectMetadata?.fileName !== "sample-report.xlsx" ||
    insertTextObjects.object?.objectMetadata?.displayAsIcon !== true ||
    insertTextObjects.equation?.text !== "y = mx + b" ||
    !String(insertTextObjects.equation?.textFontFamily || "").includes("Cambria") ||
    insertTextObjects.slicer?.objectKind !== "slicer" ||
    !String(insertTextObjects.slicer?.text || "").includes("スライサー") ||
    insertTextObjects.timeline?.objectKind !== "timeline" ||
    !String(insertTextObjects.timeline?.text || "").includes("タイムライン") ||
    insertTextObjects.miniDashboardAddin?.objectKind !== "officeAddin" ||
    insertTextObjects.miniDashboardAddin?.addinId !== "websheet-mini-dashboard" ||
    insertTextObjects.bingMapsAddin?.objectKind !== "officeAddin" ||
    insertTextObjects.bingMapsAddin?.addinId !== "websheet-bing-maps" ||
    !insertTextObjects.contextualTabVisible
  ) {
    throw new Error("Insert text/object commands should create selectable worksheet objects with preserved shape metadata.");
  }


  if (
    ["最近使用した図形", "線", "四角形", "基本図形", "ブロック矢印", "数式図形", "フローチャート", "星とリボン", "吹き出し"].some(
      (label) => !insertShapeMenu.sections.includes(label),
    ) ||
    insertShapeMenu.choiceCount < 45 ||
    !insertShapeMenu.focusedBefore ||
    !insertShapeMenu.focusedAfterRight ||
    insertShapeMenu.focusedBefore === insertShapeMenu.focusedAfterRight ||
    !insertShapeMenu.hasLine ||
    !insertShapeMenu.hasRectangle ||
    !insertShapeMenu.hasFlowchart ||
    !insertShapeMenu.hasCallout
  ) {
    throw new Error("Shape insert dropdown should expose Excel-like grouped shape palettes instead of inserting only a rectangle.");
  }


  if (
    insertShapeDraw.countAfterChoice !== insertShapeDraw.beforeCount ||
    !insertShapeDraw.drawingMode ||
    !insertShapeDraw.statusAfterChoice.includes("ドラッグ") ||
    !insertShapeDraw.draftVisible ||
    insertShapeDraw.draftWidth < 160 ||
    insertShapeDraw.draftHeight < 80 ||
    insertShapeDraw.countAfterDraw <= insertShapeDraw.beforeCount ||
    !insertShapeDraw.selected ||
    insertShapeDraw.alt !== "判断" ||
    insertShapeDraw.width < 160 ||
    insertShapeDraw.height < 80 ||
    !insertShapeDraw.shapeClassApplied ||
    insertShapeDraw.shapeObjectFit !== "fill" ||
    !insertShapeDraw.shapePreserveAspectRatioNone ||
    insertShapeDraw.draftAfterDraw ||
    insertShapeDraw.drawingModeAfterDraw ||
    !insertShapeDraw.contextualTabVisible ||
    insertShapeDraw.contextualTabActive !== "図形の書式" ||
    insertShapeDraw.resizeHandleCount !== 8 ||
    insertShapeDraw.sizeInputsAfterDraw.join("|") !== `${insertShapeDraw.height}|${insertShapeDraw.width}` ||
    !insertShapeDraw.shapeStyleTextOverflowHidden ||
    !insertShapeDraw.shapeStyleTextFits ||
    !insertShapeDraw.shapeTextToolbarPresent ||
    insertShapeDraw.objectLinkDialogTitle !== "ハイパーリンクの挿入" ||
    !insertShapeDraw.objectLinkKinds?.some((item) => item.kind === "web" && item.pressed === "true") ||
    !insertShapeDraw.objectLinkKinds?.some((item) => item.kind === "document" && item.text === "このドキュメント内") ||
    !insertShapeDraw.objectLinkKinds?.some((item) => item.kind === "email" && item.text === "電子メール アドレス") ||
    insertShapeDraw.objectLinkInitialCaption !== "アドレス(E):" ||
    insertShapeDraw.objectLinkDocumentPanelHiddenInitial !== true ||
    insertShapeDraw.objectLinkDocumentCaption !== "セル参照を入力してください(E):" ||
    insertShapeDraw.objectLinkDocumentPanelHidden !== false ||
    insertShapeDraw.objectLinkDocumentPressed !== "true" ||
    insertShapeDraw.objectLinkSelectedSheet !== "AI製品紹介" ||
    insertShapeDraw.objectLinkHyperlink !== "#'AI製品紹介'!B2" ||
    !insertShapeDraw.objectLinkStatus?.includes("リンクを設定") ||
    !insertShapeDraw.moveStatus.includes("移動") ||
    !insertShapeDraw.resizeStatus.includes("サイズ変更") ||
    insertShapeDraw.resizedWidth <= insertShapeDraw.width ||
    insertShapeDraw.resizedHeight <= insertShapeDraw.height ||
    insertShapeDraw.sizeInputsAfterResize.join("|") !== `${insertShapeDraw.resizedHeight}|${insertShapeDraw.resizedWidth}` ||
    Math.abs((insertShapeDraw.lockedResizeWidth / insertShapeDraw.lockedResizeHeight) - (insertShapeDraw.resizedWidth / insertShapeDraw.resizedHeight)) > 0.08 ||
    !insertShapeDraw.directTextEditorVisible ||
    insertShapeDraw.directTextEditorText !== "T" ||
    insertShapeDraw.directTextEditorCursor !== "text" ||
    insertShapeDraw.directTextEditorDecoration !== "none" ||
    !insertShapeDraw.contextualHiddenAfterCell ||
    insertShapeDraw.activeTabAfterCell !== "ホーム" ||
    insertShapeDraw.frameAfterCell ||
    !insertShapeDraw.contextualVisibleAfterReselect ||
    !insertShapeDraw.rotateMenuItems.includes("右へ 90 度回転") ||
    !insertShapeDraw.rotateMenuItems.includes("上下反転") ||
    insertShapeDraw.rotateMenuFocusedBefore !== "right90" ||
    insertShapeDraw.rotationInputAfterMenu !== "90" ||
    insertShapeDraw.shapeTransformAfterMenu === "none" ||
    insertShapeDraw.fillPatternChoiceCount < 10 ||
    insertShapeDraw.lineAdvancedChoiceCount < 11 ||
    insertShapeDraw.shapeEffectChoiceCount < 12 ||
    !insertShapeDraw.shapeEffectSourceHasFilter ||
    insertShapeDraw.arrowTypeChoiceCount < 12 ||
    !insertShapeDraw.arrowTypeLabels.includes("開いた矢印") ||
    !insertShapeDraw.arrowTypeLabels.includes("ひし形矢印") ||
    insertShapeDraw.arrowSizeChoiceCount < 6 ||
    !insertShapeDraw.arrowSizeLabels.includes("小") ||
    !insertShapeDraw.arrowSizeLabels.includes("大") ||
    insertShapeDraw.textFillMenuChoices < 60 ||
    insertShapeDraw.textOutlineMenuChoices < 60 ||
    !insertShapeDraw.textOutlineNoneVisible ||
    !insertShapeDraw.shapeTextPaneVisible ||
    insertShapeDraw.shapeTextPaneFields.length !== 14 ||
    insertShapeDraw.shapeLinePaneFields.length !== 9 ||
    insertShapeDraw.shapeAdvancedPaneFields.length !== 4 ||
    insertShapeDraw.shapePaneTabs.join("|") !== "塗りつぶしと線|効果|サイズとプロパティ" ||
    insertShapeDraw.shapePaneActiveInitial.join("|") !== "塗りつぶし|線" ||
    insertShapeDraw.shapePaneActiveEffects.join("|") !== "効果" ||
    insertShapeDraw.shapePaneActiveSizeProps.join("|") !== "文字|サイズ|プロパティ" ||
    !insertShapeDraw.selectionPaneVisible ||
    insertShapeDraw.selectionPaneRows < 1 ||
    insertShapeDraw.selectionPaneSelectedRows < 1 ||
    !insertShapeDraw.selectionPaneHasFormatOffset ||
    !insertShapeDraw.selectionPaneNames.some((name) => name.includes("判断")) ||
    insertShapeDraw.countAfterCtrlDragCopy <= insertShapeDraw.countAfterDraw ||
    insertShapeDraw.selectionPaneRowsAfterCopy < 2 ||
    insertShapeDraw.selectedAfterCtrlDragCopy !== 1 ||
    !insertShapeDraw.statusAfterCtrlDragCopy.includes("コピー") ||
    (Math.abs(insertShapeDraw.shiftMoveDeltaX) > 3 && Math.abs(insertShapeDraw.shiftMoveDeltaY) > 3) ||
    !insertShapeDraw.selectionPaneRenamed.includes("判断コピー") ||
    insertShapeDraw.selectedLabelAfterRename !== "判断コピー" ||
    !insertShapeDraw.selectionPaneKeyboardSelected ||
    insertShapeDraw.selectionPaneHomeSelected !== insertShapeDraw.selectionPaneFirstName ||
    insertShapeDraw.selectionPaneEndSelected !== insertShapeDraw.selectionPaneLastName ||
    insertShapeDraw.selectionPaneCtrlASelectedRows < 2 ||
    insertShapeDraw.selectionPaneSpaceHiddenRows < 1 ||
    insertShapeDraw.countAfterSelectionPaneSpaceHide >= insertShapeDraw.countAfterCtrlDragCopy ||
    !insertShapeDraw.statusAfterSelectionPaneDeleteKey.includes("削除") ||
    insertShapeDraw.selectionPaneRowsAfterDeleteKey !== insertShapeDraw.selectionPaneRowsBeforeDeleteKey - 1 ||
    !insertShapeDraw.statusAfterSelectionPaneForward.includes("前面") ||
    insertShapeDraw.selectionPaneOrderAfter === insertShapeDraw.selectionPaneOrderBefore ||
    !insertShapeDraw.statusAfterSelectionPaneOrder.includes("順序") ||
    insertShapeDraw.selectionPaneHiddenRows < 2 ||
    insertShapeDraw.countAfterSelectionPaneHide !== 0 ||
    !insertShapeDraw.statusAfterSelectionPaneHide.includes("非表示") ||
    insertShapeDraw.selectionPaneHiddenRowsAfterShow !== 0 ||
    insertShapeDraw.countAfterSelectionPaneShow <= insertShapeDraw.beforeCount ||
    !insertShapeDraw.statusAfterSelectionPaneShow.includes("表示") ||
    !insertShapeDraw.paneTextEditorVisible ||
    insertShapeDraw.paneTextEditorFontSize !== "32px" ||
    insertShapeDraw.countAfterUndo !== insertShapeDraw.beforeCount
  ) {
    throw new Error("Choosing a shape should arm drawing mode, insert by drag, then allow Excel-like move/resize with a contextual Shape Format tab.");
  }


  if (
    !underlineToolbar.menuLabels.some((label) => label.includes("下線(U)")) ||
    !underlineToolbar.menuLabels.some((label) => label.includes("二重下線(D)")) ||
    underlineToolbar.marks.find((mark) => mark.type === "underline")?.letter !== "U" ||
    underlineToolbar.marks.find((mark) => mark.type === "underline double")?.letter !== "D" ||
    underlineToolbar.marks.find((mark) => mark.type === "underline")?.lineCount !== 1 ||
    underlineToolbar.marks.find((mark) => mark.type === "underline double")?.lineCount !== 2 ||
    underlineToolbar.marks.some((mark) => mark.decorationLine !== "none") ||
    underlineToolbar.marks.some((mark) => mark.svgTextDecoration !== "none") ||
    !underlineToolbar.accessKeysUnderlined ||
    !underlineToolbar.accessKeyTargets.includes("U") ||
    !underlineToolbar.accessKeyTargets.includes("D") ||
    underlineToolbar.focusedBefore !== "underline" ||
    underlineToolbar.currentAfter !== "underline double" ||
    underlineToolbar.markAfter !== "underline double" ||
    underlineToolbar.markLetterAfter !== "D" ||
    underlineToolbar.titleAfter !== "二重下線" ||
    !underlineToolbar.decorationLine?.includes("underline") ||
    underlineToolbar.decorationStyle !== "double" ||
    !underlineToolbar.statusAfter?.includes("二重下線")
  ) {
    throw new Error("Ribbon underline control should be a split button with single and double underline choices.");
  }


  if (
    borderToolbar.hasTextSelect ||
    !borderToolbar.menuLabels.includes("外枠") ||
    !borderToolbar.menuLabels.includes("外枠太罫線") ||
    !borderToolbar.menuLabels.includes("格子") ||
    borderToolbar.focusedBefore !== "none" ||
    borderToolbar.focusedAfterArrows !== "thick-outside" ||
    borderToolbar.currentAfter !== "thick-outside" ||
    borderToolbar.markAfter !== "thick-outside" ||
    borderToolbar.titleAfter !== "外枠太罫線" ||
    !borderToolbar.statusAfter?.includes("罫線を適用")
  ) {
    throw new Error("Ribbon border control should be an icon split button whose mark updates only when choosing and applies the chosen border preset.");
  }


  if (!result.hasSheetTabScrollLeft || !result.hasSheetTabScrollRight) {
    throw new Error("Sheet tab scroll buttons did not render.");
  }


  if (
    (delayedLazyLoad.before.scrollWidth - delayedLazyLoad.before.clientWidth > 260 && delayedLazyLoad.after.rightRemaining < 200) ||
    delayedLazyLoad.after.bottomRemaining < 200 ||
    delayedLazyLoad.resetTop !== 0 ||
    delayedLazyLoad.resetLeft !== 0
  ) {
    throw new Error("Viewport-fill extension should not behave like early edge lazy-load or disturb the scroll position.");
  }


  if (
    Math.abs(stickyHeaders.columnTop - stickyHeaders.hostTop) > 1 ||
    Math.abs(stickyHeaders.rowLeft - stickyHeaders.hostLeft) > 1 ||
    Math.abs(stickyHeaders.cornerTop - stickyHeaders.hostTop) > 1 ||
    Math.abs(stickyHeaders.cornerLeft - stickyHeaders.hostLeft) > 1 ||
    stickyHeaders.rowHeaderWidth >= 46
  ) {
    throw new Error("Column and row headers should stay pinned while scrolling, and the row header gutter should shrink to the row-number digits it needs.");
  }


  const expectedSheetMenuLabels = [
    "挿入(I)...",
    "削除(D)",
    "名前の変更(R)",
    "移動またはコピー(M)...",
    "コード表示(V)",
    "シートの保護(P)...",
    "非表示(H)",
    "再表示(U)...",
    "すべてのシートを選択(S)",
  ].map((label) => label.replace(/\s+/g, ""));


  if (
    expectedSheetMenuLabels.some((label) => !sheetMenu.labels.includes(label)) ||
    !sheetMenu.labels.some((label) => label.startsWith("シート見出しの色(T)")) ||
    !["I", "D", "R", "M", "V", "P", "T", "H", "U", "S"].every((key) => sheetMenu.accessKeys.includes(key)) ||
    !sheetMenu.accessKeysUnderlined ||
    sheetMenu.focusedBefore !== "insert" ||
    sheetMenu.focusedAfterDown !== "delete" ||
    sheetMenu.separatorCount !== 2 ||
    !sheetMenu.hasColorChoices ||
    sheetMenu.rect.left < 0 ||
    sheetMenu.rect.top < 0 ||
    sheetMenu.rect.right > 1280 ||
    sheetMenu.rect.bottom > 720
  ) {
    throw new Error("Sheet tab right-click menu did not render like Excel.");
  }


  if (!sheetTabColor.hasTabColor || sheetTabColor.borderTopColor !== "rgb(255, 192, 0)" || !sheetTabColor.activeUnderline.includes("255, 192, 0")) {
    throw new Error("Sheet tab color menu did not apply the selected tab color.");
  }


  if (
    sheetRename.inline?.value !== "CHT" ||
    sheetRename.inline?.active !== true ||
    sheetRename.inline?.selected !== true ||
    sheetRename.inline?.tabRenaming !== true ||
    sheetRename.inline?.dialogOpen !== false ||
    sheetRename.applied?.tab !== "CHT Rename Test" ||
    sheetRename.undo?.tab !== "CHT" ||
    sheetRename.redo?.tab !== "CHT Rename Test" ||
    !sheetRename.applied?.status?.includes("シート名")
  ) {
    throw new Error("Sheet rename should edit inline on the tab and participate in undo/redo.");
  }


  if (
    sheetTabDirectActions.doubleClickRename?.inline?.value !== "CHT" ||
    sheetTabDirectActions.doubleClickRename?.inline?.active !== true ||
    sheetTabDirectActions.doubleClickRename?.inline?.selected !== true ||
    sheetTabDirectActions.doubleClickRename?.inline?.tabRenaming !== true ||
    sheetTabDirectActions.doubleClickRename?.applied?.tab !== "CHT Double Rename" ||
    sheetTabDirectActions.doubleClickRename?.undo?.tab !== "CHT" ||
    sheetTabDirectActions.doubleClickRename?.cancel?.tab !== "CHT" ||
    sheetTabDirectActions.doubleClickRename?.cancel?.renaming !== false ||
    !sheetTabDirectActions.doubleClickRename?.cancel?.status?.includes("キャンセル") ||
    !sheetTabDirectActions.doubleClickRename?.applied?.status?.includes("シート名") ||
    sheetTabDirectActions.dragMove?.visual?.hasIndicator !== true ||
    sheetTabDirectActions.dragMove?.visual?.copyIndicator !== false ||
    sheetTabDirectActions.dragMove?.visual?.dragSourceCount < 1 ||
    sheetTabDirectActions.dragMove?.applied?.tabs?.join("|") !== "AI製品紹介|CHT" ||
    sheetTabDirectActions.dragMove?.undo?.tabs?.join("|") !== "CHT|AI製品紹介" ||
    !sheetTabDirectActions.dragMove?.applied?.status?.includes("移動") ||
    !sheetTabDirectActions.dragMove?.undo?.status?.includes("元に戻しました") ||
    sheetTabDirectActions.dragCopy?.visual?.hasIndicator !== true ||
    sheetTabDirectActions.dragCopy?.visual?.copyIndicator !== true ||
    sheetTabDirectActions.dragCopy?.visual?.dragSourceCount < 1 ||
    sheetTabDirectActions.dragCopy?.applied?.tabs?.join("|") !== "CHT|AI製品紹介|CHT (2)" ||
    sheetTabDirectActions.dragCopy?.applied?.active !== "CHT (2)" ||
    !sheetTabDirectActions.dragCopy?.applied?.status?.includes("コピー")
  ) {
    throw new Error("Sheet tabs should support Excel-like double-click rename, drag move, and Ctrl-drag copy.");
  }


  if (
    sheetMoveCopy.dialog?.title !== "移動またはコピー" ||
    !sheetMoveCopy.dialog?.workbookCaption?.includes("移動先ブック名") ||
    !sheetMoveCopy.dialog?.beforeCaption?.includes("挿入先") ||
    !sheetMoveCopy.dialog?.options?.includes("（末尾へ移動）") ||
    sheetMoveCopy.dialog?.activeField !== "sheetMoveBeforeIndex" ||
    !sheetMoveCopy.dialog?.hasCopyCheckbox ||
    !sheetMoveCopy.dialog?.isExcelDialog ||
    sheetMoveCopy.moved?.tabs?.join("|") !== "AI製品紹介|CHT" ||
    sheetMoveCopy.moveUndo?.tabs?.join("|") !== "CHT|AI製品紹介" ||
    sheetMoveCopy.moveRedo?.tabs?.join("|") !== "AI製品紹介|CHT" ||
    !sheetMoveCopy.moveUndo?.status?.includes("元に戻しました") ||
    !sheetMoveCopy.moveRedo?.status?.includes("やり直しました") ||
    sheetMoveCopy.copied?.tabs?.join("|") !== "CHT|AI製品紹介|CHT (2)" ||
    sheetMoveCopy.copyUndo?.tabs?.join("|") !== "CHT|AI製品紹介" ||
    sheetMoveCopy.copyRedo?.tabs?.join("|") !== "CHT|AI製品紹介|CHT (2)" ||
    !sheetMoveCopy.copyUndo?.status?.includes("元に戻しました") ||
    !sheetMoveCopy.copyRedo?.status?.includes("やり直しました") ||
    sheetMoveCopy.insertDialog?.title !== "挿入" ||
    sheetMoveCopy.insertDialog?.tab !== "全般" ||
    sheetMoveCopy.insertDialog?.activeField !== "worksheet" ||
    !sheetMoveCopy.insertDialog?.isExcelDialog ||
    !sheetMoveCopy.insertDialog?.choices?.some((choice) => choice.type === "worksheet" && choice.selected && !choice.disabled) ||
    !sheetMoveCopy.insertDialog?.choices?.some((choice) => choice.type === "chart" && !choice.disabled) ||
    !["macro", "intlMacro", "dialog"].every((type) => sheetMoveCopy.insertDialog?.choices?.some((choice) => choice.type === type && !choice.disabled)) ||
    sheetMoveCopy.inserted?.tabs?.join("|") !== "Sheet3|CHT|AI製品紹介" ||
    sheetMoveCopy.deleteConfirm?.title !== "Cht WebSheet" ||
    !sheetMoveCopy.deleteConfirm?.message?.includes("完全に削除") ||
    !sheetMoveCopy.deleteConfirm?.buttons?.some((label) => label.includes("削除(D)")) ||
    !sheetMoveCopy.deleteConfirm?.deleteAccessKeyUnderlined ||
    sheetMoveCopy.deleted?.tabs?.join("|") !== "CHT|AI製品紹介" ||
    sheetMoveCopy.deleteUndo?.tabs?.join("|") !== "Sheet3|CHT|AI製品紹介" ||
    sheetMoveCopy.deleteRedo?.tabs?.join("|") !== "CHT|AI製品紹介" ||
    !sheetMoveCopy.deleteUndo?.status?.includes("元に戻しました") ||
    !sheetMoveCopy.deleteRedo?.status?.includes("やり直しました") ||
    sheetMoveCopy.chartDialog?.title !== "挿入" ||
    sheetMoveCopy.chartDialog?.chartDisabled !== false ||
    sheetMoveCopy.chartDialog?.chartSelectedBefore !== false ||
    sheetMoveCopy.chart?.tabCountDelta !== 1 ||
    sheetMoveCopy.chart?.chartSheet !== true ||
    sheetMoveCopy.chart?.showGridLines !== false ||
    sheetMoveCopy.chart?.showHeadings !== false ||
    sheetMoveCopy.chart?.imageCount !== 1 ||
    sheetMoveCopy.chart?.chartType !== "column" ||
    sheetMoveCopy.chart?.sourceSheet !== sheetMoveCopy.chart?.previousActive ||
    !String(sheetMoveCopy.chart?.sourceRange || "").startsWith("U86:W") ||
    sheetMoveCopy.chart?.selectedImageFrame !== true ||
    !sheetMoveCopy.chart?.status?.includes("グラフ シート") ||
    sheetMoveCopy.chartUndo?.active !== sheetMoveCopy.chart?.previousActive ||
    sheetMoveCopy.chartUndo?.tabCount !== sheetMoveCopy.chartUndo?.expectedTabCount ||
    !sheetMoveCopy.chartUndo?.status?.includes("元に戻しました") ||
    !Array.isArray(sheetMoveCopy.specialTemplates) ||
    sheetMoveCopy.specialTemplates.length !== 3 ||
    !sheetMoveCopy.specialTemplates.every((item) => item.dialog?.selected && item.dialog?.disabled === false && item.created?.tabCount === item.undo?.expectedTabCount + 1 && item.undo?.tabCount === item.undo?.expectedTabCount && item.undo?.status?.includes("元に戻しました")) ||
    sheetMoveCopy.specialTemplates.find((item) => item.type === "macro")?.created?.sheetKind !== "macro" ||
    sheetMoveCopy.specialTemplates.find((item) => item.type === "macro")?.created?.xl4MacroSheet !== true ||
    sheetMoveCopy.specialTemplates.find((item) => item.type === "macro")?.created?.codeObjectName !== "Macro" ||
    sheetMoveCopy.specialTemplates.find((item) => item.type === "intlMacro")?.created?.sheetKind !== "intlMacro" ||
    sheetMoveCopy.specialTemplates.find((item) => item.type === "intlMacro")?.created?.internationalMacroSheet !== true ||
    sheetMoveCopy.specialTemplates.find((item) => item.type === "intlMacro")?.created?.codeObjectName !== "InternationalMacro" ||
    sheetMoveCopy.specialTemplates.find((item) => item.type === "dialog")?.created?.sheetKind !== "dialog" ||
    sheetMoveCopy.specialTemplates.find((item) => item.type === "dialog")?.created?.dialogSheet !== true ||
    sheetMoveCopy.specialTemplates.find((item) => item.type === "dialog")?.created?.showGridLines !== false ||
    sheetMoveCopy.specialTemplates.find((item) => item.type === "dialog")?.created?.showHeadings !== false ||
    sheetMoveCopy.specialTemplates.find((item) => item.type === "dialog")?.created?.codeObjectName !== "DialogSheet"
  ) {
    throw new Error("Move or Copy sheet should open an Excel-like dialog and support undo/redo for move and copy.");
  }


  if (
    !sheetCodeDialog.initial?.title?.includes("Microsoft Visual Basic for Applications") ||
    !sheetCodeDialog.initial?.activeProjectItem?.includes("CHT") ||
    !sheetCodeDialog.initial?.objectOptions?.includes("Worksheet") ||
    !sheetCodeDialog.initial?.procedureOptions?.includes("Change") ||
    !sheetCodeDialog.initial?.status?.includes("コード ウィンドウ") ||
    !sheetCodeDialog.procedureInserted?.includes("Private Sub Worksheet_Change") ||
    !sheetCodeDialog.saved?.module?.code?.includes("Changed") ||
    sheetCodeDialog.saved?.module?.objectName !== "Worksheet" ||
    !sheetCodeDialog.saved?.module?.updatedAt ||
    !sheetCodeDialog.saved?.status?.includes("コードを保存") ||
    sheetCodeDialog.undo?.module?.code ||
    !sheetCodeDialog.redo?.module?.code?.includes("Changed") ||
    !sheetCodeDialog.reopened?.code?.includes("Changed") ||
    sheetCodeDialog.reopened?.procedure !== "Change" ||
    !sheetCodeDialog.reopened?.title?.includes("CHT")
  ) {
    throw new Error("Sheet code window should open like Excel VBA, insert worksheet stubs, persist code, and participate in undo/redo.");
  }


  if (sheetSelectAll.selectedTabs < 2 || !sheetSelectAll.status?.includes("表示ワークシート")) {
    throw new Error("Select all sheets action did not mark visible sheet tabs.");
  }


  if (sheetHidden.tabs.includes("CHT") || sheetHidden.active !== "AI製品紹介" || !sheetHidden.status?.includes("非表示")) {
    throw new Error("Hide sheet action did not hide the selected sheet tab.");
  }


  if (
    sheetUnhideDialog.title !== "再表示" ||
    !sheetUnhideDialog.caption?.includes("再表示するシート") ||
    !sheetUnhideDialog.options?.includes("CHT") ||
    sheetUnhideDialog.selected !== "CHT" ||
    sheetUnhideDialog.activeField !== "unhideSheetIndex" ||
    !sheetUnhideDialog.isExcelDialog
  ) {
    throw new Error("Unhide sheet should open an Excel-like sheet picker dialog.");
  }


  if (!sheetUnhidden.tabs.includes("CHT") || sheetUnhidden.active !== "CHT" || !sheetUnhidden.status?.includes("再表示")) {
    throw new Error("Unhide sheet action did not restore the hidden sheet tab.");
  }


  if (sheetUnhideUndo.tabs.includes("CHT") || !sheetUnhideUndo.status?.includes("元に戻しました") || !sheetUnhideRedo.tabs.includes("CHT") || sheetUnhideRedo.active !== "CHT" || !sheetUnhideRedo.status?.includes("やり直しました")) {
    throw new Error("Sheet hide/unhide should participate in undo/redo without clearing history.");
  }


  if (
    sheetGroupActions.selection?.selectedTabs?.length !== 2 ||
    !sheetGroupActions.selection?.selectedTabs?.includes("Sheet3") ||
    !sheetGroupActions.selection?.selectedTabs?.includes("Sheet4") ||
    !sheetGroupActions.selection?.status?.includes("2 個") ||
    !sheetGroupActions.color?.coloredTabs?.includes("Sheet3") ||
    !sheetGroupActions.color?.coloredTabs?.includes("Sheet4") ||
    !sheetGroupActions.color?.status?.includes("2 個") ||
    sheetGroupActions.colorUndo?.coloredTabs?.includes("Sheet3") ||
    sheetGroupActions.colorUndo?.coloredTabs?.includes("Sheet4") ||
    !sheetGroupActions.colorUndo?.status?.includes("元に戻しました") ||
    sheetGroupActions.hidden?.tabs?.includes("Sheet3") ||
    sheetGroupActions.hidden?.tabs?.includes("Sheet4") ||
    !sheetGroupActions.hidden?.status?.includes("2 個") ||
    !sheetGroupActions.hideUndo?.tabs?.includes("Sheet3") ||
    !sheetGroupActions.hideUndo?.tabs?.includes("Sheet4") ||
    !sheetGroupActions.hideUndo?.status?.includes("元に戻しました") ||
    !sheetGroupActions.deleteConfirm?.message?.includes("複数のシート") ||
    !sheetGroupActions.deleted?.status?.includes("2 個") ||
    sheetGroupActions.deleted?.tabs?.includes("Sheet3") ||
    sheetGroupActions.deleted?.tabs?.includes("Sheet4") ||
    !sheetGroupActions.deleteUndo?.tabs?.includes("Sheet3") ||
    !sheetGroupActions.deleteUndo?.tabs?.includes("Sheet4") ||
    !sheetGroupActions.deleteUndo?.status?.includes("元に戻しました") ||
    sheetGroupActions.deleteRedo?.tabs?.includes("Sheet3") ||
    sheetGroupActions.deleteRedo?.tabs?.includes("Sheet4") ||
    !sheetGroupActions.deleteRedo?.status?.includes("やり直しました")
  ) {
    throw new Error("Grouped sheet tabs should support Ctrl/Cmd selection, shared tab color, hide, delete, and single-step undo/redo.");
  }


  if (
    sheetGroupMoveCopy.dragMove?.visual?.hasIndicator !== true ||
    sheetGroupMoveCopy.dragMove?.visual?.copyIndicator !== false ||
    sheetGroupMoveCopy.dragMove?.visual?.dragSourceCount !== 2 ||
    !sheetGroupMoveCopy.dragMove?.visual?.status?.includes("2 個") ||
    sheetGroupMoveCopy.dragMove?.applied?.tabs?.join("|") !== "Sheet3|Sheet4|CHT|AI製品紹介" ||
    sheetGroupMoveCopy.dragMove?.applied?.selectedTabs?.join("|") !== "Sheet3|Sheet4" ||
    !sheetGroupMoveCopy.dragMove?.applied?.status?.includes("2 個") ||
    sheetGroupMoveCopy.dragMove?.undo?.tabs?.join("|") !== "CHT|AI製品紹介|Sheet3|Sheet4" ||
    !sheetGroupMoveCopy.dragMove?.undo?.status?.includes("元に戻しました") ||
    sheetGroupMoveCopy.dragCopy?.visual?.hasIndicator !== true ||
    sheetGroupMoveCopy.dragCopy?.visual?.copyIndicator !== true ||
    sheetGroupMoveCopy.dragCopy?.visual?.dragSourceCount !== 2 ||
    !sheetGroupMoveCopy.dragCopy?.visual?.status?.includes("2 個") ||
    sheetGroupMoveCopy.dragCopy?.applied?.tabs?.join("|") !== "Sheet3 (2)|Sheet4 (2)|CHT|AI製品紹介|Sheet3|Sheet4" ||
    sheetGroupMoveCopy.dragCopy?.applied?.selectedTabs?.join("|") !== "Sheet3 (2)|Sheet4 (2)" ||
    !sheetGroupMoveCopy.dragCopy?.applied?.status?.includes("2 個") ||
    sheetGroupMoveCopy.dragCopy?.undo?.tabs?.join("|") !== "CHT|AI製品紹介|Sheet3|Sheet4" ||
    !sheetGroupMoveCopy.dragCopy?.undo?.status?.includes("元に戻しました") ||
    sheetGroupMoveCopy.dialog?.title !== "移動またはコピー" ||
    sheetGroupMoveCopy.dialog?.activeField !== "sheetMoveBeforeIndex" ||
    !sheetGroupMoveCopy.dialog?.hasCopyCheckbox ||
    !sheetGroupMoveCopy.dialog?.status?.includes("2 個") ||
    sheetGroupMoveCopy.moved?.tabs?.slice(0, 2).join("|") !== "Sheet3|Sheet4" ||
    sheetGroupMoveCopy.moved?.selectedTabs?.join("|") !== "Sheet3|Sheet4" ||
    !sheetGroupMoveCopy.moved?.status?.includes("2 個") ||
    sheetGroupMoveCopy.moveUndo?.tabs?.join("|") !== "CHT|AI製品紹介|Sheet3|Sheet4" ||
    !sheetGroupMoveCopy.moveUndo?.status?.includes("元に戻しました") ||
    sheetGroupMoveCopy.moveRedo?.tabs?.slice(0, 2).join("|") !== "Sheet3|Sheet4" ||
    sheetGroupMoveCopy.moveRedo?.selectedTabs?.join("|") !== "Sheet3|Sheet4" ||
    !sheetGroupMoveCopy.moveRedo?.status?.includes("やり直しました") ||
    !sheetGroupMoveCopy.copied?.tabs?.includes("Sheet3 (2)") ||
    !sheetGroupMoveCopy.copied?.tabs?.includes("Sheet4 (2)") ||
    sheetGroupMoveCopy.copied?.selectedTabs?.join("|") !== "Sheet3 (2)|Sheet4 (2)" ||
    !sheetGroupMoveCopy.copied?.status?.includes("2 個") ||
    sheetGroupMoveCopy.copyUndo?.tabs?.includes("Sheet3 (2)") ||
    sheetGroupMoveCopy.copyUndo?.tabs?.includes("Sheet4 (2)") ||
    !sheetGroupMoveCopy.copyUndo?.status?.includes("元に戻しました") ||
    sheetGroupMoveCopy.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("Move or Copy should support grouped sheet tabs, preserve order, select copied sheets, and undo/redo in one step.");
  }


  if (
    sheetGroupCellEditing.input?.sheet3 !== "GroupValue" ||
    sheetGroupCellEditing.input?.sheet4 !== "GroupValue" ||
    sheetGroupCellEditing.format?.sheet3 !== "700" ||
    sheetGroupCellEditing.format?.sheet4 !== "700" ||
    !sheetGroupCellEditing.merge?.sheet3?.some((merge) => merge.top === 1 && merge.left === 3 && merge.bottom === 1 && merge.right === 4) ||
    !sheetGroupCellEditing.merge?.sheet4?.some((merge) => merge.top === 1 && merge.left === 3 && merge.bottom === 1 && merge.right === 4) ||
    sheetGroupCellEditing.mergeUndo?.sheet3?.some((merge) => merge.top === 1 && merge.left === 3 && merge.bottom === 1 && merge.right === 4) ||
    sheetGroupCellEditing.mergeUndo?.sheet4?.some((merge) => merge.top === 1 && merge.left === 3 && merge.bottom === 1 && merge.right === 4) ||
    sheetGroupCellEditing.clear?.sheet3 ||
    sheetGroupCellEditing.clear?.sheet4 ||
    sheetGroupCellEditing.clearUndo?.sheet3 !== "GroupValue" ||
    sheetGroupCellEditing.clearUndo?.sheet4 !== "GroupValue" ||
    sheetGroupCellEditing.clearRedo?.sheet3 ||
    sheetGroupCellEditing.clearRedo?.sheet4 ||
    sheetGroupCellEditing.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("Grouped sheets should synchronize cell input, formatting, merge/unmerge, clear, and undo/redo.");
  }


  if (
    sheetGroupStructure.rowInsert?.sheet3A3 !== "GroupRow" ||
    sheetGroupStructure.rowInsert?.sheet4A3 !== "GroupRow" ||
    sheetGroupStructure.rowInsertUndo?.sheet3A2 !== "GroupRow" ||
    sheetGroupStructure.rowInsertUndo?.sheet4A2 !== "GroupRow" ||
    sheetGroupStructure.rowDelete?.sheet3A2 !== "GroupRow" ||
    sheetGroupStructure.rowDelete?.sheet4A2 !== "GroupRow" ||
    sheetGroupStructure.columnInsert?.sheet3C1 !== "GroupCol" ||
    sheetGroupStructure.columnInsert?.sheet4C1 !== "GroupCol" ||
    sheetGroupStructure.columnDelete?.sheet3B1 !== "GroupCol" ||
    sheetGroupStructure.columnDelete?.sheet4B1 !== "GroupCol" ||
    sheetGroupStructure.columnWidth?.sheet3?.join("|") !== "91|91" ||
    sheetGroupStructure.columnWidth?.sheet4?.join("|") !== "91|91" ||
    sheetGroupStructure.rowHeight?.sheet3?.join("|") !== "33|33" ||
    sheetGroupStructure.rowHeight?.sheet4?.join("|") !== "33|33" ||
    !sheetGroupStructure.columnHidden?.sheet3?.every(Boolean) ||
    !sheetGroupStructure.columnHidden?.sheet4?.every(Boolean) ||
    sheetGroupStructure.columnHiddenUndo?.sheet3?.some(Boolean) ||
    sheetGroupStructure.columnHiddenUndo?.sheet4?.some(Boolean) ||
    sheetGroupStructure.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("Grouped sheets should synchronize row/column structure, dimensions, visibility, and undo.");
  }


  if (
    sheetGroupFormatDialog.applied?.firstB2?.numFmt !== "¥0.0" ||
    sheetGroupFormatDialog.applied?.firstC2?.numFmt !== "¥0.0" ||
    sheetGroupFormatDialog.applied?.secondB2?.numFmt !== "¥0.0" ||
    sheetGroupFormatDialog.applied?.secondC2?.numFmt !== "¥0.0" ||
    sheetGroupFormatDialog.applied?.firstB2?.backgroundColor !== "#FCE4D6" ||
    sheetGroupFormatDialog.applied?.secondB2?.backgroundColor !== "#FCE4D6" ||
    sheetGroupFormatDialog.applied?.firstB2?.color !== "#C00000" ||
    sheetGroupFormatDialog.applied?.secondB2?.color !== "#C00000" ||
    sheetGroupFormatDialog.applied?.firstB2?.fontSize !== "16pt" ||
    sheetGroupFormatDialog.applied?.secondB2?.fontSize !== "16pt" ||
    sheetGroupFormatDialog.applied?.firstB2?.textAlign !== "center" ||
    sheetGroupFormatDialog.applied?.secondB2?.textAlign !== "center" ||
    sheetGroupFormatDialog.applied?.firstB2?.verticalAlign !== "top" ||
    sheetGroupFormatDialog.applied?.secondB2?.verticalAlign !== "top" ||
    sheetGroupFormatDialog.applied?.firstB2?.whiteSpace !== "pre-wrap" ||
    sheetGroupFormatDialog.applied?.secondB2?.whiteSpace !== "pre-wrap" ||
    sheetGroupFormatDialog.applied?.firstB2?.borderTop?.style !== "double" ||
    sheetGroupFormatDialog.applied?.secondB2?.borderTop?.style !== "double" ||
    sheetGroupFormatDialog.applied?.firstB2?.borderTop?.color !== "#00B050" ||
    sheetGroupFormatDialog.applied?.secondB2?.borderTop?.color !== "#00B050" ||
    sheetGroupFormatDialog.applied?.firstB2?.protection?.hidden !== true ||
    sheetGroupFormatDialog.applied?.secondB2?.protection?.hidden !== true ||
    sheetGroupFormatDialog.undo?.firstB2Fill ||
    sheetGroupFormatDialog.undo?.secondB2Fill ||
    sheetGroupFormatDialog.undo?.firstB2Format ||
    sheetGroupFormatDialog.undo?.secondB2Format ||
    sheetGroupFormatDialog.redo?.firstB2Fill !== "#FCE4D6" ||
    sheetGroupFormatDialog.redo?.secondB2Fill !== "#FCE4D6" ||
    sheetGroupFormatDialog.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("Format Cells dialog should synchronize all cell formatting across grouped sheets with single-step undo/redo.");
  }


  if (
    sheetGroupDataValidation.applied?.firstB2?.type !== "list" ||
    sheetGroupDataValidation.applied?.firstC2?.type !== "list" ||
    sheetGroupDataValidation.applied?.secondB2?.type !== "list" ||
    sheetGroupDataValidation.applied?.secondC2?.type !== "list" ||
    sheetGroupDataValidation.applied?.firstB2?.formula !== "Red,Blue,Green" ||
    sheetGroupDataValidation.applied?.secondB2?.formula !== "Red,Blue,Green" ||
    sheetGroupDataValidation.applied?.firstB2?.showInputMessage !== true ||
    sheetGroupDataValidation.applied?.secondB2?.showInputMessage !== true ||
    sheetGroupDataValidation.applied?.firstB2?.inputTitle !== "選択" ||
    sheetGroupDataValidation.applied?.secondB2?.inputTitle !== "選択" ||
    sheetGroupDataValidation.applied?.firstB2?.errorStyle !== "warning" ||
    sheetGroupDataValidation.applied?.secondB2?.errorStyle !== "warning" ||
    sheetGroupDataValidation.undo?.firstB2 ||
    sheetGroupDataValidation.undo?.secondB2 ||
    sheetGroupDataValidation.redo?.firstB2?.type !== "list" ||
    sheetGroupDataValidation.redo?.secondB2?.type !== "list" ||
    sheetGroupDataValidation.clear?.firstB2 ||
    sheetGroupDataValidation.clear?.firstC2 ||
    sheetGroupDataValidation.clear?.secondB2 ||
    sheetGroupDataValidation.clear?.secondC2 ||
    sheetGroupDataValidation.clearUndo?.firstB2?.type !== "list" ||
    sheetGroupDataValidation.clearUndo?.secondB2?.type !== "list" ||
    sheetGroupDataValidation.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("Data Validation should synchronize dialog apply, clear, and undo/redo across grouped sheets.");
  }


  if (
    sheetGroupDataFilterSort.filter?.first?.autoFilter?.range?.top !== 1 ||
    sheetGroupDataFilterSort.filter?.first?.autoFilter?.range?.right !== 2 ||
    sheetGroupDataFilterSort.filter?.second?.autoFilter?.range?.top !== 1 ||
    sheetGroupDataFilterSort.filter?.second?.autoFilter?.range?.right !== 2 ||
    !sheetGroupDataFilterSort.filter?.status?.includes("2 個のシート") ||
    sheetGroupDataFilterSort.sort?.first?.b?.join("|") !== "1|2|3" ||
    sheetGroupDataFilterSort.sort?.first?.a?.join("|") !== "One|Two|Three" ||
    sheetGroupDataFilterSort.sort?.second?.b?.join("|") !== "10|20|30" ||
    sheetGroupDataFilterSort.sort?.second?.a?.join("|") !== "Ten|Twenty|Thirty" ||
    sheetGroupDataFilterSort.sort?.first?.sortState?.keys?.[0]?.col !== 2 ||
    sheetGroupDataFilterSort.sort?.second?.sortState?.keys?.[0]?.col !== 2 ||
    !sheetGroupDataFilterSort.sort?.status?.includes("2 個のシート") ||
    sheetGroupDataFilterSort.sortUndo?.first?.b?.join("|") !== "3|1|2" ||
    sheetGroupDataFilterSort.sortUndo?.second?.b?.join("|") !== "30|10|20" ||
    !sheetGroupDataFilterSort.sortUndo?.status?.includes("元に戻しました") ||
    sheetGroupDataFilterSort.sortRedo?.first?.b?.join("|") !== "1|2|3" ||
    sheetGroupDataFilterSort.sortRedo?.second?.b?.join("|") !== "10|20|30" ||
    !sheetGroupDataFilterSort.sortRedo?.status?.includes("やり直しました") ||
    sheetGroupDataFilterSort.reapply?.first?.sortState?.keys?.[0]?.col !== 2 ||
    sheetGroupDataFilterSort.reapply?.second?.sortState?.keys?.[0]?.col !== 2 ||
    !sheetGroupDataFilterSort.reapply?.status?.includes("2 個のシート") ||
    sheetGroupDataFilterSort.clear?.first?.sortState ||
    sheetGroupDataFilterSort.clear?.second?.sortState ||
    sheetGroupDataFilterSort.clear?.first?.autoFilter?.range?.right !== 2 ||
    sheetGroupDataFilterSort.clear?.second?.autoFilter?.range?.right !== 2 ||
    !sheetGroupDataFilterSort.clear?.status?.includes("クリア") ||
    sheetGroupDataFilterSort.removed?.first?.autoFilter ||
    sheetGroupDataFilterSort.removed?.second?.autoFilter ||
    !sheetGroupDataFilterSort.removed?.status?.includes("2 個のシート") ||
    sheetGroupDataFilterSort.removeUndo?.first?.autoFilter?.range?.right !== 2 ||
    sheetGroupDataFilterSort.removeUndo?.second?.autoFilter?.range?.right !== 2 ||
    !sheetGroupDataFilterSort.removeUndo?.status?.includes("元に戻しました") ||
    sheetGroupDataFilterSort.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("Data filter, reapply, clear, and quick sort should synchronize across grouped sheets with undo/redo.");
  }


  if (
    sheetGroupDataTools.flash?.firstB?.join("|") !== "Ada|Grace|Alan" ||
    sheetGroupDataTools.flash?.secondB?.join("|") !== "Marie|Niels|Rosalind" ||
    !sheetGroupDataTools.flash?.status?.includes("2 個のシート") ||
    sheetGroupDataTools.textColumns?.first?.join("|") !== "x|y|p|q" ||
    sheetGroupDataTools.textColumns?.second?.join("|") !== "r|s|t|u" ||
    !sheetGroupDataTools.textColumns?.status?.includes("2 個のシート") ||
    sheetGroupDataTools.removeDuplicates?.first?.d?.join("|") !== "A|B|C|" ||
    sheetGroupDataTools.removeDuplicates?.second?.d?.join("|") !== "X|Y|Z|" ||
    !sheetGroupDataTools.removeDuplicates?.status?.includes("2 個のシート") ||
    sheetGroupDataTools.removeUndo?.first?.d?.join("|") !== "A|B|A|C" ||
    sheetGroupDataTools.removeUndo?.second?.d?.join("|") !== "X|Y|X|Z" ||
    !sheetGroupDataTools.removeUndo?.status?.includes("元に戻しました") ||
    sheetGroupDataTools.removeRedo?.first?.d?.join("|") !== "A|B|C|" ||
    sheetGroupDataTools.removeRedo?.second?.d?.join("|") !== "X|Y|Z|" ||
    !sheetGroupDataTools.removeRedo?.status?.includes("やり直しました") ||
    sheetGroupDataTools.subtotal?.first?.d?.join("|") !== "Name|A|A 合計|B|B 合計|C|C 合計|総計" ||
    sheetGroupDataTools.subtotal?.second?.d?.join("|") !== "Name|X|X 合計|Y|Y 合計|Z|Z 合計|総計" ||
    !sheetGroupDataTools.subtotal?.status?.includes("2 個のシート") ||
    sheetGroupDataTools.subtotalUndo?.first?.d?.join("|") !== "Name|A|B|C||||" ||
    sheetGroupDataTools.subtotalUndo?.second?.d?.join("|") !== "Name|X|Y|Z||||" ||
    !sheetGroupDataTools.subtotalUndo?.status?.includes("元に戻しました") ||
    sheetGroupDataTools.subtotalRedo?.first?.d?.join("|") !== "Name|A|A 合計|B|B 合計|C|C 合計|総計" ||
    sheetGroupDataTools.subtotalRedo?.second?.d?.join("|") !== "Name|X|X 合計|Y|Y 合計|Z|Z 合計|総計" ||
    !sheetGroupDataTools.subtotalRedo?.status?.includes("やり直しました") ||
    sheetGroupDataTools.outlineHidden?.first?.hiddenRows?.[0] !== true ||
    sheetGroupDataTools.outlineHidden?.second?.hiddenRows?.[0] !== true ||
    !sheetGroupDataTools.outlineHidden?.status?.includes("2 個のシート") ||
    sheetGroupDataTools.outlineShown?.first?.hiddenRows?.[0] !== false ||
    sheetGroupDataTools.outlineShown?.second?.hiddenRows?.[0] !== false ||
    !sheetGroupDataTools.outlineShown?.status?.includes("2 個のシート") ||
    sheetGroupDataTools.outlineSettings?.first?.settings?.summaryBelow !== false ||
    sheetGroupDataTools.outlineSettings?.second?.settings?.summaryBelow !== false ||
    sheetGroupDataTools.outlineSettings?.first?.settings?.summaryRight !== false ||
    sheetGroupDataTools.outlineSettings?.second?.settings?.summaryRight !== false ||
    !sheetGroupDataTools.outlineSettings?.status?.includes("2 個のシート") ||
    sheetGroupDataTools.consolidate?.first?.h?.join("|") !== "統合|Alpha|Beta" ||
    sheetGroupDataTools.consolidate?.first?.i?.join("|") !== "Q1|2|5" ||
    sheetGroupDataTools.consolidate?.first?.j?.join("|") !== "Q2|3|7" ||
    sheetGroupDataTools.consolidate?.second?.h?.join("|") !== "統合|Alpha|Beta" ||
    sheetGroupDataTools.consolidate?.second?.i?.join("|") !== "Q1|20|50" ||
    sheetGroupDataTools.consolidate?.second?.j?.join("|") !== "Q2|30|70" ||
    !sheetGroupDataTools.consolidate?.status?.includes("2 個のシート") ||
    sheetGroupDataTools.consolidateUndo?.first?.h?.join("|") !== "||" ||
    sheetGroupDataTools.consolidateUndo?.second?.h?.join("|") !== "||" ||
    !sheetGroupDataTools.consolidateUndo?.status?.includes("元に戻しました") ||
    sheetGroupDataTools.consolidateRedo?.first?.i?.join("|") !== "Q1|2|5" ||
    sheetGroupDataTools.consolidateRedo?.second?.i?.join("|") !== "Q1|20|50" ||
    !sheetGroupDataTools.consolidateRedo?.status?.includes("やり直しました") ||
    sheetGroupDataTools.analysis?.first?.n?.join("|") !== "|平均|合計" ||
    sheetGroupDataTools.analysis?.first?.o?.join("|") !== "Value|4|12" ||
    sheetGroupDataTools.analysis?.second?.n?.join("|") !== "|平均|合計" ||
    sheetGroupDataTools.analysis?.second?.o?.join("|") !== "Value|40|120" ||
    !sheetGroupDataTools.analysis?.status?.includes("2 個のシート") ||
    sheetGroupDataTools.analysisUndo?.first?.n?.join("|") !== "||" ||
    sheetGroupDataTools.analysisUndo?.second?.n?.join("|") !== "||" ||
    !sheetGroupDataTools.analysisUndo?.status?.includes("元に戻しました") ||
    sheetGroupDataTools.analysisRedo?.first?.o?.join("|") !== "Value|4|12" ||
    sheetGroupDataTools.analysisRedo?.second?.o?.join("|") !== "Value|40|120" ||
    !sheetGroupDataTools.analysisRedo?.status?.includes("やり直しました") ||
    sheetGroupDataTools.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("Flash Fill, Text to Columns, Remove Duplicates, Subtotal, Outline, Consolidate, and Data Analysis should synchronize across grouped sheets with undo/redo.");
  }


  if (
    sheetGroupHyperlink.applied?.firstD4?.raw !== "Grouped Link" ||
    sheetGroupHyperlink.applied?.secondD4?.raw !== "Grouped Link" ||
    sheetGroupHyperlink.applied?.firstD4?.hyperlink !== "https://example.com/grouped" ||
    sheetGroupHyperlink.applied?.secondD4?.hyperlink !== "https://example.com/grouped" ||
    sheetGroupHyperlink.applied?.firstD4?.textDecoration !== "underline" ||
    sheetGroupHyperlink.applied?.secondD4?.textDecoration !== "underline" ||
    sheetGroupHyperlink.undo?.firstD4?.raw !== "Before" ||
    sheetGroupHyperlink.undo?.secondD4?.raw !== "Before" ||
    sheetGroupHyperlink.undo?.firstD4?.hyperlink ||
    sheetGroupHyperlink.undo?.secondD4?.hyperlink ||
    sheetGroupHyperlink.redo?.firstD4 !== "https://example.com/grouped" ||
    sheetGroupHyperlink.redo?.secondD4 !== "https://example.com/grouped" ||
    sheetGroupHyperlink.remove?.firstD4?.raw !== "Grouped Link" ||
    sheetGroupHyperlink.remove?.secondD4?.raw !== "Grouped Link" ||
    sheetGroupHyperlink.remove?.firstD4?.hyperlink ||
    sheetGroupHyperlink.remove?.secondD4?.hyperlink ||
    sheetGroupHyperlink.removeUndo?.firstD4 !== "https://example.com/grouped" ||
    sheetGroupHyperlink.removeUndo?.secondD4 !== "https://example.com/grouped" ||
    sheetGroupHyperlink.homeClear?.firstD4?.raw !== "Grouped Link" ||
    sheetGroupHyperlink.homeClear?.secondD4?.raw !== "Grouped Link" ||
    sheetGroupHyperlink.homeClear?.firstD4?.hyperlink ||
    sheetGroupHyperlink.homeClear?.secondD4?.hyperlink ||
    sheetGroupHyperlink.homeClearUndo?.firstD4 !== "https://example.com/grouped" ||
    sheetGroupHyperlink.homeClearUndo?.secondD4 !== "https://example.com/grouped" ||
    sheetGroupHyperlink.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("Hyperlink insert/remove and Home clear should synchronize across grouped sheets with single-step undo/redo.");
  }


  if (
    sheetGroupHomeCommands.fill?.firstB?.join("|") !== "FirstSeed|FirstSeed|FirstSeed" ||
    sheetGroupHomeCommands.fill?.secondB?.join("|") !== "SecondSeed|SecondSeed|SecondSeed" ||
    sheetGroupHomeCommands.fillUndo?.firstB?.join("|") !== "FirstSeed||" ||
    sheetGroupHomeCommands.fillUndo?.secondB?.join("|") !== "SecondSeed||" ||
    sheetGroupHomeCommands.fillRedo?.firstB?.join("|") !== "FirstSeed|FirstSeed|FirstSeed" ||
    sheetGroupHomeCommands.fillRedo?.secondB?.join("|") !== "SecondSeed|SecondSeed|SecondSeed" ||
    sheetGroupHomeCommands.autoSum?.firstA3?.formula !== "=SUM(A1:A2)" ||
    sheetGroupHomeCommands.autoSum?.secondA3?.formula !== "=SUM(A1:A2)" ||
    sheetGroupHomeCommands.autoSumUndo?.firstA3 ||
    sheetGroupHomeCommands.autoSumUndo?.secondA3 ||
    sheetGroupHomeCommands.autoSumRedo?.firstA3 !== "=SUM(A1:A2)" ||
    sheetGroupHomeCommands.autoSumRedo?.secondA3 !== "=SUM(A1:A2)" ||
    sheetGroupHomeCommands.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("Home fill and AutoSum should synchronize across grouped sheets with single-step undo/redo.");
  }


  if (
    sheetGroupTranslation.pane?.title !== "翻訳" ||
    !sheetGroupTranslation.pane?.range?.includes("B2") ||
    !(
      (sheetGroupTranslation.pane?.sourceText === "氏名" && sheetGroupTranslation.pane?.resultText === "姓名") ||
      (sheetGroupTranslation.pane?.sourceText === "年齢" && sheetGroupTranslation.pane?.resultText === "年龄")
    ) ||
    sheetGroupTranslation.insertRight?.firstC2 !== "姓名" ||
    sheetGroupTranslation.insertRight?.secondC2 !== "年龄" ||
    sheetGroupTranslation.insertUndo?.firstC2 ||
    sheetGroupTranslation.insertUndo?.secondC2 ||
    sheetGroupTranslation.insertRedo?.firstC2 !== "姓名" ||
    sheetGroupTranslation.insertRedo?.secondC2 !== "年龄" ||
    sheetGroupTranslation.replace?.firstB2 !== "姓名" ||
    sheetGroupTranslation.replace?.secondB2 !== "年龄" ||
    sheetGroupTranslation.replaceUndo?.firstB2 !== "氏名" ||
    sheetGroupTranslation.replaceUndo?.secondB2 !== "年齢" ||
    sheetGroupTranslation.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("Translation pane insert-right and replace should synchronize across grouped sheets with undo/redo.");
  }


  if (
    sheetGroupQuickAnalysis.keyboard?.focusedBefore !== "format-data-bars" ||
    !sheetGroupQuickAnalysis.keyboard?.focusedAfterRight ||
    sheetGroupQuickAnalysis.keyboard?.focusedAfterRight === sheetGroupQuickAnalysis.keyboard?.focusedBefore ||
    Math.abs((sheetGroupQuickAnalysis.bars?.firstA3?.ratio ?? 0) - 1) > 0.001 ||
    Math.abs((sheetGroupQuickAnalysis.bars?.secondA3?.ratio ?? 0) - 1) > 0.001 ||
    sheetGroupQuickAnalysis.barsUndo?.firstA3 ||
    sheetGroupQuickAnalysis.barsUndo?.secondA3 ||
    Math.abs((sheetGroupQuickAnalysis.barsRedo?.firstA3?.ratio ?? 0) - 1) > 0.001 ||
    Math.abs((sheetGroupQuickAnalysis.barsRedo?.secondA3?.ratio ?? 0) - 1) > 0.001 ||
    sheetGroupQuickAnalysis.clear?.firstA3 ||
    sheetGroupQuickAnalysis.clear?.secondA3 ||
    Math.abs((sheetGroupQuickAnalysis.clearUndo?.firstA3?.ratio ?? 0) - 1) > 0.001 ||
    Math.abs((sheetGroupQuickAnalysis.clearUndo?.secondA3?.ratio ?? 0) - 1) > 0.001 ||
    sheetGroupQuickAnalysis.colorScale?.firstA3Background !== "#63BE7B" ||
    sheetGroupQuickAnalysis.colorScale?.secondA3Background !== "#63BE7B" ||
    sheetGroupQuickAnalysis.colorScale?.firstA3Bar ||
    sheetGroupQuickAnalysis.colorScale?.secondA3Bar ||
    sheetGroupQuickAnalysis.colorScaleUndo?.firstA3Background ||
    sheetGroupQuickAnalysis.colorScaleUndo?.secondA3Background ||
    Math.abs((sheetGroupQuickAnalysis.colorScaleUndo?.firstA3Bar?.ratio ?? 0) - 1) > 0.001 ||
    Math.abs((sheetGroupQuickAnalysis.colorScaleUndo?.secondA3Bar?.ratio ?? 0) - 1) > 0.001 ||
    sheetGroupQuickAnalysis.colorScaleRedo?.firstA3Background !== "#63BE7B" ||
    sheetGroupQuickAnalysis.colorScaleRedo?.secondA3Background !== "#63BE7B" ||
    sheetGroupQuickAnalysis.iconSet?.firstA3Icon?.shape !== "up" ||
    sheetGroupQuickAnalysis.iconSet?.secondA3Icon?.shape !== "up" ||
    sheetGroupQuickAnalysis.iconUndo?.firstA3Icon ||
    sheetGroupQuickAnalysis.iconUndo?.secondA3Icon ||
    sheetGroupQuickAnalysis.iconRedo?.firstA3Icon?.shape !== "up" ||
    sheetGroupQuickAnalysis.iconRedo?.secondA3Icon?.shape !== "up" ||
    sheetGroupQuickAnalysis.duplicates?.first?.join("|") !== "true|true|false" ||
    sheetGroupQuickAnalysis.duplicates?.second?.join("|") !== "true|false|true" ||
    sheetGroupQuickAnalysis.duplicatesUndo?.first?.join("|") !== "false|false|false" ||
    sheetGroupQuickAnalysis.duplicatesUndo?.second?.join("|") !== "false|false|false" ||
    sheetGroupQuickAnalysis.duplicatesRedo?.first?.join("|") !== "true|true|false" ||
    sheetGroupQuickAnalysis.duplicatesRedo?.second?.join("|") !== "true|false|true" ||
    sheetGroupQuickAnalysis.total?.firstA4 !== "=SUM($A$1:$A$3)" ||
    sheetGroupQuickAnalysis.total?.secondA4 !== "=SUM($A$1:$A$3)" ||
    sheetGroupQuickAnalysis.totalUndo?.firstA4 ||
    sheetGroupQuickAnalysis.totalUndo?.secondA4 ||
    sheetGroupQuickAnalysis.totalRedo?.firstA4 !== "=SUM($A$1:$A$3)" ||
    sheetGroupQuickAnalysis.totalRedo?.secondA4 !== "=SUM($A$1:$A$3)" ||
    sheetGroupQuickAnalysis.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("Quick Analysis format and totals should synchronize across grouped sheets with undo/redo.");
  }


  if (
    sheetGroupSparklines.create?.firstD1?.type !== "line" ||
    sheetGroupSparklines.create?.firstD1?.values?.join("|") !== "1|2|3" ||
    sheetGroupSparklines.create?.firstD2?.values?.join("|") !== "4|5|6" ||
    sheetGroupSparklines.create?.secondD1?.values?.join("|") !== "10|20|30" ||
    sheetGroupSparklines.create?.secondD2?.values?.join("|") !== "40|50|60" ||
    !sheetGroupSparklines.create?.status?.includes("2 個のシート") ||
    sheetGroupSparklines.undo?.firstD1 ||
    sheetGroupSparklines.undo?.secondD1 ||
    !sheetGroupSparklines.undo?.status?.includes("元に戻しました") ||
    sheetGroupSparklines.redo?.firstD1?.values?.join("|") !== "1|2|3" ||
    sheetGroupSparklines.redo?.secondD1?.values?.join("|") !== "10|20|30" ||
    !sheetGroupSparklines.redo?.status?.includes("やり直しました") ||
    sheetGroupSparklines.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("Sparkline insertion should synchronize across grouped sheets with each sheet's own data and undo/redo.");
  }


  if (
    sheetGroupCharts.create?.first?.type !== "column" ||
    sheetGroupCharts.create?.second?.type !== "column" ||
    sheetGroupCharts.create?.first?.values?.join("|") !== "1|2" ||
    sheetGroupCharts.create?.second?.values?.join("|") !== "10|20" ||
    !sheetGroupCharts.create?.first?.sourceSheet ||
    !sheetGroupCharts.create?.second?.sourceSheet ||
    sheetGroupCharts.create?.first?.sourceSheet === sheetGroupCharts.create?.second?.sourceSheet ||
    sheetGroupCharts.create?.first?.sourceRange !== "A1:B3" ||
    sheetGroupCharts.create?.second?.sourceRange !== "A1:B3" ||
    !sheetGroupCharts.create?.status?.includes("2 個のシート") ||
    sheetGroupCharts.undo?.firstCount !== 0 ||
    sheetGroupCharts.undo?.secondCount !== 0 ||
    !sheetGroupCharts.undo?.status?.includes("元に戻しました") ||
    sheetGroupCharts.redo?.first?.values?.join("|") !== "1|2" ||
    sheetGroupCharts.redo?.second?.values?.join("|") !== "10|20" ||
    !sheetGroupCharts.redo?.status?.includes("やり直しました") ||
    sheetGroupCharts.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("Chart insertion should synchronize across grouped sheets with each sheet's own source data and undo/redo.");
  }


  if (
    sheetGroupTables.create?.first?.count !== 1 ||
    sheetGroupTables.create?.second?.count !== 1 ||
    sheetGroupTables.create?.first?.range?.top !== 1 ||
    sheetGroupTables.create?.first?.range?.right !== 2 ||
    sheetGroupTables.create?.second?.b2 !== 10 ||
    sheetGroupTables.create?.first?.headerFill !== "#4472c4" ||
    sheetGroupTables.create?.second?.headerFill !== "#4472c4" ||
    !sheetGroupTables.create?.status?.includes("2 個のシート") ||
    sheetGroupTables.undo?.firstCount !== 0 ||
    sheetGroupTables.undo?.secondCount !== 0 ||
    sheetGroupTables.undo?.firstHeaderFill === "#4472c4" ||
    sheetGroupTables.undo?.secondHeaderFill === "#4472c4" ||
    !sheetGroupTables.undo?.status?.includes("元に戻しました") ||
    sheetGroupTables.redo?.firstCount !== 1 ||
    sheetGroupTables.redo?.secondCount !== 1 ||
    sheetGroupTables.redo?.firstB2 !== 1 ||
    sheetGroupTables.redo?.secondB2 !== 10 ||
    !sheetGroupTables.redo?.status?.includes("やり直しました") ||
    sheetGroupTables.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("Table creation should synchronize across grouped sheets with each sheet's own data and undo/redo.");
  }


  if (
    sheetGroupPivot.create?.pivotNames?.length !== 2 ||
    !sheetGroupPivot.create?.pivotNames?.every((name) => name.includes("PivotTable_")) ||
    sheetGroupPivot.create?.pivotData?.[0]?.tables?.length !== 1 ||
    sheetGroupPivot.create?.pivotData?.[1]?.tables?.length !== 1 ||
    sheetGroupPivot.create?.pivotData?.[0]?.rows?.[0]?.join("|") !== "A|2|3" ||
    sheetGroupPivot.create?.pivotData?.[1]?.rows?.[0]?.join("|") !== "A|1|10" ||
    !sheetGroupPivot.create?.status?.includes("2 個のピボットテーブル") ||
    sheetGroupPivot.undo?.tabCount !== sheetGroupPivot.undo?.expectedCount ||
    !sheetGroupPivot.undo?.status?.includes("元に戻しました") ||
    sheetGroupPivot.redo?.pivotNames?.length !== 2 ||
    sheetGroupPivot.redo?.firstRows?.[0]?.join("|") !== "A|3" ||
    sheetGroupPivot.redo?.secondRows?.[0]?.join("|") !== "A|10" ||
    !sheetGroupPivot.redo?.status?.includes("やり直しました") ||
    sheetGroupPivot.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("PivotTable summaries should be created for grouped sheets with each sheet's own source data and undo/redo.");
  }


  if (
    !sheetGroupObjects.wordArt?.first?.hasWordArt ||
    !sheetGroupObjects.wordArt?.second?.hasWordArt ||
    sheetGroupObjects.wordArt?.first?.shapeType !== "textBox" ||
    sheetGroupObjects.wordArt?.second?.shapeType !== "textBox" ||
    !sheetGroupObjects.wordArt?.status?.includes("2 個のシート") ||
    sheetGroupObjects.wordArtUndo?.firstCount !== 0 ||
    sheetGroupObjects.wordArtUndo?.secondCount !== 0 ||
    !sheetGroupObjects.wordArtUndo?.status?.includes("元に戻しました") ||
    sheetGroupObjects.wordArtRedo?.firstWordArtCount !== 1 ||
    sheetGroupObjects.wordArtRedo?.secondWordArtCount !== 1 ||
    !sheetGroupObjects.wordArtRedo?.status?.includes("やり直しました") ||
    sheetGroupObjects.shape?.first?.rectCount !== 1 ||
    sheetGroupObjects.shape?.second?.rectCount !== 1 ||
    sheetGroupObjects.shape?.first?.imageCount !== 2 ||
    sheetGroupObjects.shape?.second?.imageCount !== 2 ||
    sheetGroupObjects.shape?.first?.width <= 40 ||
    sheetGroupObjects.shape?.second?.height <= 40 ||
    !sheetGroupObjects.shape?.status?.includes("2 個のシート") ||
    sheetGroupObjects.shapeUndo?.firstWordArtCount !== 1 ||
    sheetGroupObjects.shapeUndo?.secondWordArtCount !== 1 ||
    sheetGroupObjects.shapeUndo?.firstRectCount !== 0 ||
    sheetGroupObjects.shapeUndo?.secondRectCount !== 0 ||
    !sheetGroupObjects.shapeUndo?.status?.includes("元に戻しました") ||
    sheetGroupObjects.shapeRedo?.firstRectCount !== 1 ||
    sheetGroupObjects.shapeRedo?.secondRectCount !== 1 ||
    !sheetGroupObjects.shapeRedo?.status?.includes("やり直しました") ||
    sheetGroupObjects.style?.first?.fillColor !== "#e2f0d9" ||
    sheetGroupObjects.style?.second?.fillColor !== "#e2f0d9" ||
    !sheetGroupObjects.style?.status?.includes("2 個のシート") ||
    sheetGroupObjects.styleUndo?.firstFill === "#e2f0d9" ||
    sheetGroupObjects.styleUndo?.secondFill === "#e2f0d9" ||
    sheetGroupObjects.size?.firstWidth !== 220 ||
    sheetGroupObjects.size?.secondWidth !== 220 ||
    !sheetGroupObjects.size?.status?.includes("2 個のシート") ||
    sheetGroupObjects.sizeUndo?.firstWidth === 220 ||
    sheetGroupObjects.sizeUndo?.secondWidth === 220 ||
    sheetGroupObjects.sizeRedo?.firstWidth !== 220 ||
    sheetGroupObjects.sizeRedo?.secondWidth !== 220 ||
    sheetGroupObjects.group?.first?.groupedCount !== 2 ||
    sheetGroupObjects.group?.second?.groupedCount !== 2 ||
    sheetGroupObjects.group?.first?.groupIdCount !== 1 ||
    sheetGroupObjects.group?.second?.groupIdCount !== 1 ||
    !sheetGroupObjects.group?.status?.includes("2 個のシート") ||
    sheetGroupObjects.groupUndo?.firstGrouped !== 0 ||
    sheetGroupObjects.groupUndo?.secondGrouped !== 0 ||
    sheetGroupObjects.groupRedo?.firstGrouped !== 2 ||
    sheetGroupObjects.groupRedo?.secondGrouped !== 2 ||
    sheetGroupObjects.ungroup?.firstGrouped !== 0 ||
    sheetGroupObjects.ungroup?.secondGrouped !== 0 ||
    sheetGroupObjects.ungroupUndo?.firstGrouped !== 2 ||
    sheetGroupObjects.ungroupUndo?.secondGrouped !== 2 ||
    sheetGroupObjects.delete?.firstRectCount !== 0 ||
    sheetGroupObjects.delete?.secondRectCount !== 0 ||
    sheetGroupObjects.delete?.firstWordArtCount !== 1 ||
    sheetGroupObjects.delete?.secondWordArtCount !== 1 ||
    !sheetGroupObjects.delete?.status?.includes("2 個のシート") ||
    sheetGroupObjects.deleteUndo?.firstRectCount !== 1 ||
    sheetGroupObjects.deleteUndo?.secondRectCount !== 1 ||
    sheetGroupObjects.deleteRedo?.firstRectCount !== 0 ||
    sheetGroupObjects.deleteRedo?.secondRectCount !== 0 ||
    sheetGroupObjects.cleanup?.tabs?.join("|") !== "CHT|AI製品紹介"
  ) {
    throw new Error("Object insertion and editing should synchronize across grouped sheets with single-step undo/redo.");
  }


  if (
    !fileSideMenu.visible ||
    fileSideMenu.activeRibbonTab !== "挿入" ||
    !fileSideMenu.insertPanelVisible ||
    Math.abs(fileSideMenu.left - fileSideMenu.fileTabLeft) > 2 ||
    Math.abs(fileSideMenu.top - fileSideMenu.fileTabBottom) > 2 ||
    fileSideMenu.width < 240 ||
    !fileSideMenu.labels.includes("開く") ||
    !fileSideMenu.labels.includes("保存") ||
    fileSideMenu.focusedBefore !== "open" ||
    fileSideMenu.focusedAfterArrowDown !== "save" ||
    !fileSideMenu.accessKeys.includes("S") ||
    !fileSideMenu.accessKeys.includes("A") ||
    fileSideMenu.keyboardEscClosed !== true ||
    fileSideMenu.activeRibbonTabAfterEsc !== "挿入" ||
    !fileSideMenu.insertPanelVisibleAfterEsc ||
    !fileSideMenuClosedByOutsideClick.hidden ||
    fileSideMenuClosedByOutsideClick.activeRibbonTab !== "挿入" ||
    !fileSideMenuClosedByOutsideClick.insertPanelVisible ||
	    !fileSideMenuClosedByCommand.hidden ||
	    fileSideMenuClosedByCommand.activeRibbonTab !== "挿入" ||
	    !fileSideMenuClosedByCommand.insertPanelVisible ||
    !fileSideMenuClosedByCommand.saveFileName?.endsWith(".html") ||
    !fileSideMenuClosedByCommand.saveIsHtml ||
    !fileSideMenuClosedByCommand.saveHasWorkbook ||
    fileSideMenuClosedByCommand.saveState?.saveType !== "html" ||
    fileSideMenuClosedByCommand.saveState?.currentSaveType !== "html" ||
    fileSideMenuClosedByCommand.saveState?.isStandalone
	  ) {
	    throw new Error("File tab should open next to the File tab without changing the current ribbon panel, and Save should quick-save WebSheet HTML in the web app.");
	  }


  if (
    !shortcutSave.ctrlFileName?.endsWith(".html") ||
    !shortcutSave.metaFileName?.endsWith(".html") ||
    !shortcutSave.ctrlIsHtml ||
    !shortcutSave.metaIsHtml ||
    !shortcutSave.formulaFocusedDuringMetaSave
  ) {
    throw new Error("Ctrl+S/Cmd+S should intercept the browser save shortcut and quick-save the workbook's current Save As type, even from text inputs.");
  }


  if (
    saveAsDialog.title !== "名前を付けて保存" ||
    !saveAsDialog.fileName?.endsWith(".html") ||
    saveAsDialog.location !== "ダウンロード" ||
    saveAsDialog.path !== "PC > ダウンロード" ||
    !saveAsDialog.pathOptions?.includes("PC > ドキュメント") ||
    saveAsDialog.fileHeaders?.join("|") !== "名前|更新日時|種類|サイズ" ||
    !saveAsDialog.fileRows?.some((row) => row.type === "html" && row.active && row.name?.endsWith(".html")) ||
    !saveAsDialog.fileRows?.some((row) => row.type === "xlsx" && row.name?.endsWith(".xlsx")) ||
    !saveAsDialog.fileRows?.some((row) => row.type === "xlsm" && row.name?.endsWith(".xlsm")) ||
    !saveAsDialog.fileRows?.some((row) => row.type === "xls" && row.name?.endsWith(".xls")) ||
    saveAsDialog.newFolderButton !== "新しいフォルダー(N)" ||
    saveAsDialog.xlsmRowSelected?.type !== "xlsm" ||
    saveAsDialog.xlsmRowSelected?.activeType !== "xlsm" ||
    !saveAsDialog.xlsmRowSelected?.fileName?.endsWith(".xlsm") ||
    saveAsDialog.documentsPathSelected?.location !== "ドキュメント" ||
    saveAsDialog.documentsPathSelected?.path !== "PC > ドキュメント" ||
    saveAsDialog.documentsPathSelected?.activePlace !== "ドキュメント" ||
    !saveAsDialog.documentsPathSelected?.note?.includes("ブラウザー制限") ||
    !saveAsDialog.types?.includes("WebSheet HTML (*.html)") ||
    !saveAsDialog.types?.includes("Excel ブック (*.xlsx)") ||
    !saveAsDialog.types?.includes("Excel マクロ有効ブック (*.xlsm)") ||
    !saveAsDialog.types?.includes("Excel 97-2003 ブック (*.xls)") ||
    !saveAsDialog.note?.includes("WebSheet") ||
    !saveAsDialog.xlsxFileName?.endsWith(".xlsx") ||
    !saveAsDialog.xlsxSuggestedFileName?.endsWith(".xlsx") ||
    saveAsDialog.xlsxZipHeader !== "504b0304" ||
    !saveAsDialog.handlePicked?.location?.includes("選択済み") ||
    !saveAsDialog.handlePicked?.fileName?.endsWith("-picked.xlsx") ||
    !saveAsDialog.handlePicked?.note?.includes("直接書き込み") ||
    saveAsDialog.handleQuickSave?.writeCount !== 1 ||
    !saveAsDialog.handleQuickSave?.writes?.every((entry) => !entry.size || entry.type.includes("spreadsheetml")) ||
    !saveAsDialog.handleQuickSave?.saveFileName?.endsWith("-picked.html") ||
    saveAsDialog.handleQuickSave?.saveIsHtml !== true ||
    saveAsDialog.handleQuickSave?.saveHasWorkbook !== true ||
    !saveAsDialog.handleQuickSave?.status?.includes("保存しました") ||
    !saveAsDialog.handleQuickSave?.workbookName?.endsWith("-picked.html")
  ) {
    throw new Error("Save As should show an Excel-like path/name/type dialog, export Excel workbooks, and keep web-service quick save as a small HTML workbook.");
  }


  if (result.b14.contentWidth <= result.b14.cellWidth || !result.b14.text?.includes("D：知見あり")) {
    throw new Error("B14 long text is still clipped to the source cell width.");
  }


  if (Math.abs(result.b14.cellWidth - 31) > 1 || Math.abs(result.b14.cellHeight - 20) > 1 || Math.abs(result.b14.columnHeaderWidth - 31) > 1 || Math.abs(result.b14.rowHeaderHeight - 20) > 1) {
    throw new Error("CHT sheet dimensions should import the workbook's original row heights and column widths without baking the saved view zoom into the model.");
  }


  if (!result.canvasTextCentering?.length || result.canvasTextCentering.some((item) => !item || item.count < 10 || Math.abs(item.delta) > 1.25)) {
    throw new Error("Canvas-rendered vertically centered text should align its ink bounds with the visual center of the Excel cell.");
  }


  if (!result.verticalText?.hasTextRotation || result.verticalText.text !== "保有技術" || !/span\s+29/.test(result.verticalText.gridRow || "")) {
    throw new Error("Imported Excel vertical text cells should keep vertical text rotation across merged ranges.");
  }


  if (
    result.modelIntegrity?.richRuns !== 3 ||
    !result.modelIntegrity.richRunFonts.some((font) => /MS|ＭＳ|ゴシック|Gothic/i.test(font)) ||
    result.modelIntegrity.datedifFormula !== '=DATEDIF(C50,E50,"M")+1' ||
    result.modelIntegrity.datedifCached !== 21 ||
    String(result.modelIntegrity.datedifDisplay) !== "21"
  ) {
    throw new Error("Imported Excel rich text and DATEDIF formula values should survive model conversion and formula evaluation.");
  }


  if (
    !result.memoryStats ||
    result.memoryStats.cells > 5000 ||
    result.memoryStats.denseFormulaCells > 0 ||
    result.memoryStats.redundantDisplayCells !== 0 ||
    result.memoryStats.emptyCssFieldCells !== 0 ||
    result.memoryStats.emptyBorderFieldCells !== 0 ||
    result.memoryStats.styledCells <= result.memoryStats.uniqueCssObjects ||
    result.memoryStats.borderedCells <= result.memoryStats.uniqueBorderObjects
  ) {
    throw new Error("Workbook model should compact default cell fields, share duplicate style objects, and avoid retaining dense formula arrays.");
  }


  if (result.cells > 5000 || result.rows > 120) {
    throw new Error("Sheet DOM is not being virtualized/lazily unloaded.");
  }


  if (result.a1Contenteditable !== null || singleClickEdit !== null || doubleClickEdit !== "true") {
    throw new Error("Cell edit mode should start only on double click.");
  }


	  if (
	    editAfterDeferredRender.editable !== "true" ||
	    editAfterDeferredRender.activeAddress !== "B14" ||
	    editAlignment.editable !== "true" ||
	    editAlignment.display !== "block" ||
	    editAlignment.textAlign !== "right" ||
	    editAlignment.justifyContent !== "flex-end" ||
	    editAlignment.alignItems !== "flex-end" ||
	    editAlignment.paddingTopPx <= 0 ||
	    editClickCaret.editable !== "true" ||
	    editClickCaret.activeAddress !== "B14" ||
    !editClickCaret.selectionInCell ||
    editTextSelection.editable !== "true" ||
    !editTextSelection.selectionInCell ||
    !editTextSelection.selectedText
	  ) {
	    throw new Error("Double-click edit mode should survive deferred viewport rendering, preserve cell alignment, and allow mouse caret placement/text selection inside the cell.");
	  }


		  if (
		    !editNoEllipsis ||
		    editNoEllipsis.editable !== "true" ||
		    editNoEllipsis.textOverflow === "ellipsis" ||
		    !editNoEllipsis.scrollsHorizontally ||
		    editNoEllipsis.text.includes("…")
		  ) {
		    throw new Error("Edit-mode cells should not visually abbreviate long input with an ellipsis.");
		  }


	  if (
	    !cellEditAltEnter ||
	    cellEditAltEnter.editable !== "true" ||
	    cellEditAltEnter.text !== "Line1\n" ||
	    cellEditAltEnter.lineCount !== 2 ||
	    cellEditAltEnter.whiteSpace !== "pre" ||
	    cellEditAltEnter.activeAddress !== "B14" ||
	    cellEditAltEnter.label !== "B14"
	  ) {
	    throw new Error("Alt+Enter should insert a line break inside the active cell without leaving edit mode.");
	  }


	  if (
	    !cellEditAltEnterDuringIme ||
	    !cellEditAltEnterDuringIme.defaultPrevented ||
	    cellEditAltEnterDuringIme.editable !== "true" ||
	    cellEditAltEnterDuringIme.text !== "IME\n" ||
	    cellEditAltEnterDuringIme.activeAddress !== "B14" ||
	    !cellEditAltEnterDuringIme.activeElementInCell ||
	    cellEditAltEnterDuringIme.label !== "B14"
	  ) {
	    throw new Error("Alt+Enter should insert a line break even when the editable cell is in IME composition mode.");
	  }


	  if (
	    !editDragOutsideSelection ||
	    !editDragOutsideSelection.targetAddress ||
	    editDragOutsideSelection.targetAddress === "B14" ||
	    editDragOutsideSelection.editable !== "true" ||
	    editDragOutsideSelection.activeAddress !== "B14" ||
	    editDragOutsideSelection.label !== "B14"
	  ) {
	    throw new Error("Dragging from an edited cell and releasing outside should not select the release target cell.");
	  }


		  if (
		    !cellEditImeEnter ||
	    cellEditImeEnter.defaultPrevented ||
    cellEditImeEnter.editable !== "true" ||
    cellEditImeEnter.afterEndEditable !== "true" ||
    cellEditImeEnter.activeAddress !== "B14" ||
    cellEditImeEnter.afterEndActiveAddress !== "B14" ||
    !cellEditImeEnter.activeElementInCell ||
    cellEditImeEnter.label !== "B14" ||
    cellEditImeEnter.text !== "変換中"
  ) {
    throw new Error("IME conversion Enter should keep the current cell in edit mode instead of committing or moving selection.");
  }


	  if (
	    editExitRender.editable !== null ||
	    !editExitRender.hasContentWrapper ||
	    editExitRender.directTextNode ||
    editExitRender.contentOpacity !== "0" ||
    !editExitRender.text?.includes("D：知見あり")
	  ) {
	    throw new Error("Leaving edit mode without changes should restore the normal hidden DOM overlay and canvas-rendered cell text.");
	  }


	  if (
	    !multilineCommitDisplay ||
	    multilineCommitDisplay.raw !== "上\n下" ||
	    multilineCommitDisplay.display !== "上\n下" ||
	    multilineCommitDisplay.whiteSpace !== "pre-wrap" ||
	    multilineCommitDisplay.overflowWrap !== "anywhere" ||
	    multilineCommitDisplay.afterHeight <= multilineCommitDisplay.beforeHeight
	  ) {
	    throw new Error("Committed cells containing Alt+Enter line breaks should enable wrap text and grow row height so all lines can render.");
	  }


	  if (!result.fontSamples.b14?.includes("MS PGothic") || !result.fontSamples.b14?.includes("Hiragino Sans")) {
    throw new Error("CHT Japanese Excel font fallback stack did not render.");
  }


  if (!result.fontSamples.fontOptions.includes("MS Pゴシック") || !result.fontSamples.fontOptions.includes("Noto Sans CJK JP") || !result.fontSamples.fontOptions.includes("Hiragino Kaku Gothic ProN")) {
    throw new Error("Ribbon font list is missing expected Japanese font choices.");
  }


  if (editScrollbar?.overflowY !== "auto" || editScrollbar?.scrollbarWidth !== "none" || editScrollbar?.webkitDisplay !== "none" || editScrollbar?.webkitWidth !== "0px") {
    throw new Error("Edit-mode cell scrollbars should stay hidden while internal scrolling remains available.");
  }


  if (formulaBarInitial.placeholder !== null || formulaBarInitial.hasApplyButton || formulaBarInitial.value !== "") {
    throw new Error("Formula bar should not show placeholder text or an apply button, and empty cells should leave it blank.");
  }


  if (
    formulaBarSizing.initial.rows !== 1 ||
    formulaBarSizing.initial.rowHeight >= 40 ||
    Number.parseFloat(formulaBarSizing.initial.fontSize) > 12.5 ||
    formulaBarSizing.initial.expanded !== "false" ||
    formulaBarSizing.expanded3.rows !== 3 ||
    formulaBarSizing.expanded3.expanded !== "true" ||
    formulaBarSizing.expanded3.rowHeight < formulaBarSizing.initial.rowHeight + formulaBarSizing.initial.lineHeight * 1.5 ||
    formulaBarSizing.collapsedAfter3.rows !== 1 ||
    formulaBarSizing.manual5.rows !== 5 ||
    formulaBarSizing.collapsedAfterManual.rows !== 1 ||
    formulaBarSizing.reopenedManual5.rows !== 5 ||
    formulaBarSizing.restored.rows !== 1 ||
    formulaBarSizing.hiddenScroll.scrollHeight <= formulaBarSizing.hiddenScroll.clientHeight ||
    formulaBarSizing.hiddenScroll.scrollTop <= 0 ||
    formulaBarSizing.hiddenScroll.overflowY !== "auto" ||
    formulaBarSizing.hiddenScroll.scrollbarWidth !== "none" ||
    formulaBarSizing.hiddenScroll.webkitDisplay !== "none" ||
    formulaBarSizing.hiddenScroll.buttonBorderTopWidth !== "0px" ||
    formulaBarSizing.hiddenScroll.buttonBackground !== "rgba(0, 0, 0, 0)"
  ) {
    throw new Error("Formula bar should be compact, resize by line-height steps, remember manual rows, scroll hidden overflow while dragging text selection, and use an integrated arrow.");
  }


  if (
    formulaBarDuringEdit.contenteditable !== "true" ||
    formulaBarDuringEdit.previewText !== "FormulaBarEdit" ||
    formulaBarDuringEdit.inputValue !== "FormulaBarEdit"
  ) {
    throw new Error("Clicking the formula bar should put the selected cell into edit mode and preview the input.");
  }


  if (
    formulaBarEscape.contenteditable !== null ||
    formulaBarEscape.text !== "" ||
    formulaBarEscape.inputValue !== "" ||
    !formulaBarEscape.status?.includes("キャンセル")
  ) {
    throw new Error("Escape in the formula bar should cancel the current input and restore the cell.");
  }


  if (
    formulaBarEnter.contenteditable !== null ||
    formulaBarEnter.text !== "FormulaBarEdit" ||
    formulaBarEnter.inputValue !== "FormulaBarEdit" ||
    !formulaBarEnter.status?.includes("更新")
  ) {
    throw new Error("Enter in the formula bar should commit the input and return the cell to selection mode.");
  }


  if (
    !formulaAutocomplete.barOpen?.names?.includes("SUM") ||
    !formulaAutocomplete.barOpen?.names?.includes("SUMIF") ||
    formulaAutocomplete.barOpen?.active !== "SUM" ||
    formulaAutocomplete.barOpen?.role !== "listbox" ||
    formulaAutocomplete.barOpen?.ariaLabel !== "数式オートコンプリート" ||
    formulaAutocomplete.barOpen?.activeDescendant !== formulaAutocomplete.barOpen?.activeId ||
    formulaAutocomplete.barOpen?.selectedAria !== "true" ||
    formulaAutocomplete.barOpen?.inputValue !== "=su" ||
    !formulaAutocomplete.barOpen?.syntaxes?.some((syntax) => String(syntax || "").startsWith("SUM(")) ||
    formulaAutocomplete.imeEnter?.defaultPrevented ||
    formulaAutocomplete.imeEnter?.inputValue !== "=su" ||
    !formulaAutocomplete.imeEnter?.popupVisible ||
    formulaAutocomplete.imeEnter?.active !== "SUM" ||
    (formulaAutocomplete.barOpen?.count > 1 && formulaAutocomplete.keyboard?.end?.active === "SUM") ||
    formulaAutocomplete.keyboard?.end?.activeDescendant !== formulaAutocomplete.keyboard?.end?.activeId ||
    formulaAutocomplete.keyboard?.home?.active !== "SUM" ||
    formulaAutocomplete.keyboard?.home?.activeDescendant !== formulaAutocomplete.keyboard?.home?.activeId ||
    (formulaAutocomplete.barOpen?.count > 6 && formulaAutocomplete.keyboard?.pageDown?.active === "SUM") ||
    formulaAutocomplete.keyboard?.pageDown?.activeDescendant !== formulaAutocomplete.keyboard?.pageDown?.activeId ||
    formulaAutocomplete.keyboard?.pageUp?.active !== "SUM" ||
    formulaAutocomplete.keyboard?.pageUp?.activeDescendant !== formulaAutocomplete.keyboard?.pageUp?.activeId ||
    formulaAutocomplete.barAccepted?.inputValue !== "=SUM(" ||
    !formulaAutocomplete.barAccepted?.hint?.includes("SUM(number1") ||
    !String(formulaAutocomplete.barAccepted?.activeArgument || "").includes("number1") ||
    !String(formulaAutocomplete.barAccepted?.activeDescription || "").includes("number1") ||
    formulaAutocomplete.barAccepted?.listVisible ||
    formulaAutocomplete.barRepeatHint?.inputValue !== "=SUM(2,3," ||
    !String(formulaAutocomplete.barRepeatHint?.hint || "").includes("number3") ||
    !String(formulaAutocomplete.barRepeatHint?.activeArgument || "").includes("number3") ||
    !String(formulaAutocomplete.barRepeatHint?.activeDescription || "").includes("number3") ||
    formulaAutocomplete.barRepeatHint?.listVisible ||
    formulaAutocomplete.barArrayHint?.inputValue !== "=SUM({1,2}," ||
    !String(formulaAutocomplete.barArrayHint?.activeArgument || "").includes("number2") ||
    String(formulaAutocomplete.barArrayHint?.activeArgument || "").includes("number3") ||
    formulaAutocomplete.barArrayHint?.listVisible ||
    formulaAutocomplete.argumentExpansion?.bar?.inputValue !== "=IF(logical_test, value_if_true, value_if_false)" ||
    !String(formulaAutocomplete.argumentExpansion?.bar?.activeDescription || "").includes("logical_test") ||
    formulaAutocomplete.functionArgumentsShortcut?.barInitial?.title !== "関数の引数" ||
    formulaAutocomplete.functionArgumentsShortcut?.barInitial?.functionName !== "IF" ||
    formulaAutocomplete.functionArgumentsShortcut?.barInitial?.hasBack !== false ||
    formulaAutocomplete.functionArgumentsShortcut?.barInitial?.values?.join("|") !== "logical_test|value_if_true|value_if_false" ||
    formulaAutocomplete.functionArgumentsShortcut?.barApplied?.inputValue !== "=IF(1=1,9,3)" ||
    formulaAutocomplete.functionArgumentsShortcut?.barApplied?.dialogClosed !== true ||
    formulaAutocomplete.functionArgumentsShortcut?.barApplied?.activeElement !== "formulaInput" ||
    formulaAutocomplete.functionArgumentsShortcut?.barApplied?.previewText !== "=IF(1=1,9,3)" ||
    !formulaAutocomplete.nameOpen?.names?.includes("SalesDataAmounts") ||
    !formulaAutocomplete.nameOpen?.categories?.includes("名前") ||
    formulaAutocomplete.nameAccepted?.inputValue !== "=SalesDataAmounts" ||
    formulaAutocomplete.nameAccepted?.popupVisible ||
    !formulaAutocomplete.localNameScope?.names?.includes("LocalMetric") ||
    !formulaAutocomplete.localNameScope?.categories?.includes("名前 (シート)") ||
    !formulaAutocomplete.localNameScope?.descriptions?.some((description) => description.includes("CHT:")) ||
    formulaAutocomplete.otherSheetLocalName?.names?.includes("OtherSheetOnly") ||
    !formulaAutocomplete.tableOpen?.names?.includes("SalesTable") ||
    !formulaAutocomplete.tableOpen?.categories?.includes("テーブル") ||
    formulaAutocomplete.tableAccepted?.inputValue !== "=SalesTable" ||
    formulaAutocomplete.tableAccepted?.popupVisible ||
    !formulaAutocomplete.structuredOpen?.names?.includes("#Data") ||
    !formulaAutocomplete.structuredOpen?.names?.includes("売上") ||
    !formulaAutocomplete.structuredOpen?.categories?.includes("テーブル列") ||
    formulaAutocomplete.structuredAccepted?.inputValue !== "=SalesTable[売上]" ||
    formulaAutocomplete.structuredAccepted?.popupVisible ||
    !formulaAutocomplete.structuredNestedOpen?.names?.includes("売上") ||
    formulaAutocomplete.structuredNestedOpen?.names?.includes("#Data") ||
    formulaAutocomplete.structuredNestedAccepted?.inputValue !== "=SalesTable[[#Data],[売上]]" ||
    formulaAutocomplete.structuredNestedAccepted?.popupVisible ||
    formulaAutocomplete.engineDefinitions?.SUMPRODUCT?.engineSupported !== true ||
    formulaAutocomplete.engineDefinitions?.COUNTBLANK?.syntax !== "COUNTBLANK(range)" ||
    formulaAutocomplete.engineDefinitions?.["NETWORKDAYS.INTL"]?.category !== "日付/時刻" ||
    formulaAutocomplete.engineDefinitions?.IMABS?.category !== "エンジニアリング" ||
    formulaAutocomplete.engineDefinitions?.["CEILING.MATH"]?.category !== "数学/三角" ||
    formulaAutocomplete.engineDefinitions?.AGGREGATE?.syntax !== "AGGREGATE(function_num, options, ref1, [ref2], ...)" ||
    formulaAutocomplete.engineDefinitions?.LET?.syntax !== "LET(name1, name_value1, calculation_or_name2, [name_value2], ...)" ||
    formulaAutocomplete.engineDefinitions?.MAKEARRAY?.syntax !== "MAKEARRAY(rows, cols, lambda)" ||
    formulaAutocomplete.engineDefinitions?.BYROW?.category !== "検索/行列" ||
    formulaAutocomplete.engineDefinitions?.MAP?.category !== "検索/行列" ||
    formulaAutocomplete.engineDefinitions?.ROMAN?.syntax !== "ROMAN(number, [form])" ||
    formulaAutocomplete.engineDefinitions?.POWER?.category !== "数学/三角" ||
    formulaAutocomplete.engineDefinitions?.SQRT?.syntax !== "SQRT(number)" ||
    formulaAutocomplete.engineDefinitions?.["F.DIST.RT"]?.syntax !== "F.DIST.RT(x, deg_freedom1, deg_freedom2)" ||
    formulaAutocomplete.engineDefinitions?.["F.INV.RT"]?.category !== "統計" ||
    formulaAutocomplete.engineDefinitions?.["LOGNORM.DIST"]?.syntax !== "LOGNORM.DIST(x, mean, standard_dev, cumulative)" ||
    formulaAutocomplete.engineDefinitions?.["LOGNORM.INV"]?.category !== "統計" ||
    formulaAutocomplete.engineDefinitions?.["NEGBINOM.DIST"]?.syntax !== "NEGBINOM.DIST(number_f, number_s, probability_s, cumulative)" ||
    formulaAutocomplete.engineDefinitions?.DELTA?.category !== "エンジニアリング" ||
    formulaAutocomplete.engineDefinitions?.GESTEP?.syntax !== "GESTEP(number, [step])" ||
    formulaAutocomplete.engineDefinitions?.["ERFC.PRECISE"]?.category !== "エンジニアリング" ||
    formulaAutocomplete.engineDefinitions?.TBILLEQ?.category !== "財務" ||
    formulaAutocomplete.engineDefinitions?.ISPMT?.syntax !== "ISPMT(rate, per, nper, pv)" ||
    !formulaAutocomplete.engineOpen?.names?.includes("SUMPRODUCT") ||
    formulaAutocomplete.engineOpen?.active !== "SUMPRODUCT" ||
    !formulaAutocomplete.engineOpen?.syntaxes?.some((syntax) => String(syntax || "").startsWith("SUMPRODUCT(")) ||
    formulaAutocomplete.engineAccepted?.inputValue !== "=SUMPRODUCT(" ||
    !formulaAutocomplete.engineAccepted?.hint?.includes("SUMPRODUCT(array1") ||
    !formulaAutocomplete.advancedOpen?.names?.includes("LOGNORM.DIST") ||
    !formulaAutocomplete.advancedOpen?.names?.includes("LOGNORM.INV") ||
    !formulaAutocomplete.advancedOpen?.categories?.includes("統計") ||
    !formulaAutocomplete.advancedOpen?.syntaxes?.includes("LOGNORM.DIST(x, mean, standard_dev, cumulative)") ||
    !formulaAutocomplete.advancedOpen?.descriptions?.some((description) => String(description || "").includes("対数正規分布")) ||
    formulaAutocomplete.engineeringOpen?.active !== "DELTA" ||
    !formulaAutocomplete.engineeringOpen?.categories?.includes("エンジニアリング") ||
    !formulaAutocomplete.engineeringOpen?.syntaxes?.includes("DELTA(number1, [number2])") ||
    formulaAutocomplete.cellOpen?.editable !== "true" ||
    formulaAutocomplete.cellOpen?.active !== "AVERAGE" ||
    !formulaAutocomplete.cellOpen?.names?.includes("AVERAGEIF") ||
    formulaAutocomplete.cellAccepted?.editable !== "true" ||
    formulaAutocomplete.cellAccepted?.text !== "=AVERAGE(" ||
    !formulaAutocomplete.cellAccepted?.hint?.includes("AVERAGE(number1") ||
    formulaAutocomplete.argumentExpansion?.cell?.editable !== "true" ||
    formulaAutocomplete.argumentExpansion?.cell?.text !== "=AVERAGE(number1, number2, ...)" ||
    !String(formulaAutocomplete.argumentExpansion?.cell?.hint || "").includes("AVERAGE(number1") ||
    formulaAutocomplete.functionArgumentsShortcut?.cellInitial?.title !== "関数の引数" ||
    formulaAutocomplete.functionArgumentsShortcut?.cellInitial?.functionName !== "AVERAGE" ||
    formulaAutocomplete.functionArgumentsShortcut?.cellInitial?.hasBack !== false ||
    formulaAutocomplete.functionArgumentsShortcut?.cellInitial?.values?.slice(0, 2).join("|") !== "2|4" ||
    formulaAutocomplete.functionArgumentsShortcut?.cellApplied?.editable !== "true" ||
    formulaAutocomplete.functionArgumentsShortcut?.cellApplied?.text !== "=AVERAGE(2,8)" ||
    formulaAutocomplete.functionArgumentsShortcut?.cellApplied?.dialogClosed !== true ||
    formulaAutocomplete.functionArgumentsShortcut?.cellApplied?.activeElement !== formulaAutocomplete.cellAddress
  ) {
    throw new Error("Formula autocomplete should appear while typing formulas in the formula bar and cells, support keyboard completion, show argument hints/descriptions, expand argument placeholders with Ctrl+Shift+A, and open Function Arguments from formula editing with Ctrl+A.");
  }


  const customFormulaValue = (key) => formulaCustomFunctions.values?.[key]?.value;


  const customFormulaDisplay = (key) => formulaCustomFunctions.values?.[key]?.display;


  const customFormulaDefinitions = Object.values(formulaCustomFunctions.engineDefinitions || {});


  if (
    customFormulaDefinitions.some((definition) => !definition?.engineSupported) ||
    customFormulaDisplay("textjoin") !== "x-y" ||
    customFormulaDisplay("concat") !== "Axy" ||
    Math.abs(Number(customFormulaValue("averageifs")) - 26.6666666667) > 0.00001 ||
    Number(customFormulaValue("modeSngl")) !== 2 ||
    Number(customFormulaValue("rankEq")) !== 2 ||
    Number(customFormulaValue("percentileInc")) !== 2 ||
    Number(customFormulaValue("percentileExc")) !== 2 ||
    Number(customFormulaValue("quartileInc")) !== 4 ||
    Number(customFormulaValue("quartileExc")) !== 2 ||
    Math.abs(Number(customFormulaValue("intercept")) - 4.0740740741) > 0.00001 ||
    Math.abs(Number(customFormulaValue("forecastLinear")) - 31.8518518519) > 0.00001 ||
    customFormulaDisplay("dollar") !== "$1,234.57" ||
    customFormulaDisplay("fixed") !== "1,234.6" ||
    Number(customFormulaValue("numbervalue")) !== 1234.5 ||
    Number(customFormulaValue("errorType")) !== 2 ||
    Number(customFormulaValue("typeText")) !== 2 ||
    Number(customFormulaValue("typeNumber")) !== 1 ||
    Number(customFormulaValue("typeBoolean")) !== 4 ||
    Number(customFormulaValue("lookup")) !== 40 ||
    Number(customFormulaValue("xmatch")) !== 4 ||
    Number(customFormulaValue("aggregateSum")) !== 14 ||
    Number(customFormulaValue("aggregateLarge")) !== 4 ||
    Number(customFormulaValue("aggregateMedian")) !== 2 ||
    Number(customFormulaValue("aggregateIgnoreError")) !== 14 ||
    Number(customFormulaValue("ifs")) !== 2 ||
    customFormulaDisplay("xor") !== "TRUE" ||
    Number(customFormulaValue("letSimple")) !== 5 ||
    Number(customFormulaValue("letMulti")) !== 10 ||
    Number(customFormulaValue("letRange")) !== 14 ||
    formulaCustomFunctions.spillDisplays?.makearray?.join("|") !== "2|3|3|4" ||
    formulaCustomFunctions.spillDisplays?.byrow?.join("|") !== "11|22" ||
    formulaCustomFunctions.spillDisplays?.bycol?.join("|") !== "3|30" ||
    formulaCustomFunctions.spillDisplays?.map?.join("|") !== "2|4" ||
    Number(customFormulaValue("reduce")) !== 14 ||
    formulaCustomFunctions.spillDisplays?.scan?.join("|") !== "1|3|5" ||
    formulaCustomFunctions.spillDisplays?.sortby?.join("|") !== "5|50|4|40" ||
    formulaCustomFunctions.spillDisplays?.textsplitMultiDelimiter?.join("|") !== "A|B|C|D" ||
    Number(customFormulaValue("xmatchWildcard")) !== 1 ||
    Number(customFormulaValue("xmatchWildcardReverse")) !== 5 ||
    formulaCustomFunctions.spillDisplays?.takeLastRows?.join("|") !== "4|40|5|50" ||
    formulaCustomFunctions.spillDisplays?.dropLastRows?.join("|") !== "1|10|2|20|2|30" ||
    Math.abs(Number(customFormulaValue("convertLength")) - 3.280839895) > 0.00001 ||
    Math.abs(Number(customFormulaValue("convertTemp"))) > 0.00001 ||
    Number(customFormulaValue("dsum")) !== 250 ||
    Number(customFormulaValue("dcount")) !== 2 ||
    Number(customFormulaValue("dcounta")) !== 3 ||
    Number(customFormulaValue("daverage")) !== 125 ||
    Number(customFormulaValue("dmax")) !== 150 ||
    Number(customFormulaValue("dmin")) !== 100 ||
    Number(customFormulaValue("dproduct")) !== 15000 ||
    Math.abs(Number(customFormulaValue("dstdev")) - Math.sqrt(1250)) > 0.00001 ||
    Math.abs(Number(customFormulaValue("dstdevp")) - 25) > 0.00001 ||
    Number(customFormulaValue("dvar")) !== 1250 ||
    Number(customFormulaValue("dvarp")) !== 625 ||
    formulaCustomFunctions.spillDisplays?.unique?.join("|") !== "A|B" ||
    formulaCustomFunctions.spillDisplays?.sort?.join("|") !== "5|50" ||
    formulaCustomFunctions.spillDisplays?.filter?.join("|") !== "4|40|5" ||
    formulaCustomFunctions.spillDisplays?.sequence?.join("|") !== "10|15|20|25" ||
    customFormulaDisplay("indirect") !== "20" ||
    customFormulaDisplay("offsetStatic") !== "40" ||
    formulaCustomFunctions.spillDisplays?.offsetRange?.join("|") !== "2|20|2" ||
    formulaCustomFunctions.spillDisplays?.take?.join("|") !== "1|10|2" ||
    formulaCustomFunctions.spillDisplays?.drop?.join("|") !== "4|40|5" ||
    formulaCustomFunctions.spillDisplays?.vstack?.join("|") !== "1|5|50" ||
    formulaCustomFunctions.spillDisplays?.hstack?.join("|") !== "1|10|2" ||
    formulaCustomFunctions.spillDisplays?.chooseRows?.join("|") !== "1|5|50" ||
    formulaCustomFunctions.spillDisplays?.chooseCols?.join("|") !== "10|20" ||
    formulaCustomFunctions.spillDisplays?.tocol?.join("|") !== "1|10|2|20" ||
    formulaCustomFunctions.spillDisplays?.torow?.join("|") !== "1|10|2|20" ||
    formulaCustomFunctions.spillDisplays?.wraprows?.join("|") !== "1|2|5|" ||
    formulaCustomFunctions.spillDisplays?.wrapcols?.join("|") !== "1|2|5|" ||
    formulaCustomFunctions.spillDisplays?.expand?.join("|") !== "1|x|x|x" ||
    formulaCustomFunctions.spillDisplays?.indexColumn?.join("|") !== "10|20" ||
    formulaCustomFunctions.spillDisplays?.directRange?.join("|") !== "1|10|2" ||
    formulaCustomFunctions.spillDisplays?.indirectR1C1?.join("|") !== "1|10|2" ||
    formulaCustomFunctions.spillDisplays?.namedRange?.join("|") !== "1|10|2" ||
    formulaCustomFunctions.spillDisplays?.offsetCellArgs?.join("|") !== "2|20|2" ||
    formulaCustomFunctions.spillDisplays?.transpose?.join("|") !== "1|2|10" ||
    formulaCustomFunctions.spillDisplays?.chooseRange?.join("|") !== "10|20" ||
    formulaCustomFunctions.spillDisplays?.textsplit?.join("|") !== "A|B|C|D" ||
    formulaCustomFunctions.spillDisplays?.textsplitIgnoreEmpty?.join("|") !== "A|B" ||
    customFormulaDisplay("textbefore") !== "A-B" ||
    customFormulaDisplay("textafter") !== "C" ||
    customFormulaDisplay("textafterLast") !== "C" ||
    customFormulaDisplay("textbeforeFallback") !== "NF"
  ) {
    throw new Error(`Custom Excel formula compatibility functions should be registered in autocomplete and calculate the same basic results as Excel: ${JSON.stringify(formulaCustomFunctions)}`);
  }


  if (
    formulaReferenceInsert.barClick?.inputValue !== `=SUM(${formulaReferenceInsert.targets?.click}` ||
    formulaReferenceInsert.barClick?.selectedLabel !== formulaEmptyAddress ||
    !formulaReferenceInsert.barClick?.frameVisible ||
    !formulaReferenceInsert.barClick?.highlighted ||
    formulaReferenceInsert.barDrag?.inputValue !== `=SUM(${formulaReferenceInsert.targets?.dragExpected}` ||
    !formulaReferenceInsert.barDrag?.frameVisible ||
    formulaReferenceInsert.barDrag?.highlightedCount < 1 ||
    formulaReferenceInsert.replace?.inputValue !== `=SUM(${formulaReferenceInsert.targets?.dragEnd})` ||
    !formulaReferenceInsert.replace?.frameVisible ||
    !formulaReferenceInsert.replace?.highlighted ||
    formulaReferenceInsert.columnClick?.inputValue !== `=SUM(${formulaReferenceInsert.targets?.clickColumnName}:${formulaReferenceInsert.targets?.clickColumnName}` ||
    !formulaReferenceInsert.columnClick?.frameVisible ||
    formulaReferenceInsert.columnClick?.highlightedCount < 1 ||
    formulaReferenceInsert.columnDrag?.inputValue !== `=SUM(${formulaReferenceInsert.targets?.clickColumnName}:${formulaReferenceInsert.targets?.dragEndColumnName}` ||
    !formulaReferenceInsert.columnDrag?.frameVisible ||
    formulaReferenceInsert.columnDrag?.highlightedCount < 1 ||
    formulaReferenceInsert.rowClick?.inputValue !== `=SUM(${formulaReferenceInsert.targets?.clickRow}:${formulaReferenceInsert.targets?.clickRow}` ||
    !formulaReferenceInsert.rowClick?.frameVisible ||
    formulaReferenceInsert.rowClick?.highlightedCount < 1 ||
    formulaReferenceInsert.rowDrag?.inputValue !== `=SUM(${formulaReferenceInsert.targets?.clickRow}:${formulaReferenceInsert.targets?.dragEndRow}` ||
    !formulaReferenceInsert.rowDrag?.frameVisible ||
    formulaReferenceInsert.rowDrag?.highlightedCount < 1 ||
    formulaReferenceInsert.columnF4 !== `=$${formulaReferenceInsert.targets?.clickColumnName}:$${formulaReferenceInsert.targets?.dragEndColumnName}` ||
    formulaReferenceInsert.rowF4 !== `=$${formulaReferenceInsert.targets?.clickRow}:$${formulaReferenceInsert.targets?.dragEndRow}` ||
    formulaReferenceInsert.cellClick?.editable !== "true" ||
    formulaReferenceInsert.cellClick?.text !== `=SUM(${formulaReferenceInsert.targets?.click}` ||
    !formulaReferenceInsert.cellClick?.activeEditor ||
    !formulaReferenceInsert.cellClick?.frameVisible ||
    !formulaReferenceInsert.cellClick?.highlighted ||
    formulaReferenceInsert.colors?.frameCount < 2 ||
    !formulaReferenceInsert.colors?.firstCellClass ||
    !formulaReferenceInsert.colors?.secondCellClass ||
    formulaReferenceInsert.colors?.firstCellClass === formulaReferenceInsert.colors?.secondCellClass ||
    formulaReferenceInsert.f4Expected?.length !== 4 ||
    formulaReferenceInsert.f4?.join("|") !== formulaReferenceInsert.f4Expected?.join("|") ||
    (!formulaReferenceInsert.crossSheet?.skipped &&
      (!formulaReferenceInsert.crossSheet?.inputValue?.includes("!") ||
        !formulaReferenceInsert.crossSheet?.inputValue?.includes("A1") ||
        formulaReferenceInsert.crossSheet?.committed !== formulaReferenceInsert.crossSheet?.inputValue ||
        formulaReferenceInsert.crossSheet?.activeSheet !== formulaReferenceInsert.crossSheet?.source))
  ) {
    throw new Error("Formula editing should support Excel-like click/drag references, colored multi-reference frames, F4 absolute-reference cycling, and cross-sheet reference insertion.");
  }


  if (
    directTyping.firstContenteditable !== null ||
    directTyping.secondContenteditable !== null ||
    directTyping.firstText !== "X" ||
    directTyping.secondText !== "Y" ||
    !directTyping.status?.includes(`CHT!${directTyping.secondAddress} を更新`)
  ) {
    throw new Error("Typing printable keys while cells are selected should commit single-character edits and keep prior cell values.");
  }


  if (
    definedNameAndLink.nameDialogInitial.title !== "新しい名前" ||
    !definedNameAndLink.nameDialogInitial.refersTo?.includes(definedNameAndLink.absoluteNameLinkAddress) ||
    !["N", "S", "O", "R"].every((key) => definedNameAndLink.nameDialogInitial.accessKeys.includes(key)) ||
    !definedNameAndLink.nameDialogInitial.accessKeysUnderlined ||
    definedNameAndLink.definedNameFormula.value !== "50" ||
    definedNameAndLink.unicodeDefinedNameFormula.value !== "50" ||
    definedNameAndLink.unicodeDefinedNameFormula.formula !== "=スキル値+8" ||
    !definedNameAndLink.unicodeDefinedNameFormula.names.includes("スキル値") ||
    definedNameAndLink.duplicateDefinedName.title !== "Cht WebSheet" ||
    !definedNameAndLink.duplicateDefinedName.message?.includes("既に存在") ||
    definedNameAndLink.duplicateDefinedName.dialogStillOpen !== true ||
    definedNameAndLink.duplicateDefinedName.count !== 1 ||
    definedNameAndLink.localDefinedNameFormula.value !== "51" ||
    definedNameAndLink.localDefinedNameFormula.formula !== "='CHT'!ローカル値+9" ||
    definedNameAndLink.localDefinedNameFormula.localName?.scope !== "sheet" ||
    definedNameAndLink.localDefinedNameFormula.localName?.sheetIndex !== 0 ||
    definedNameAndLink.linkDialogInitial.title !== "ハイパーリンクの挿入" ||
    definedNameAndLink.linkDialogInitial.textValue !== "42" ||
    definedNameAndLink.linkDialogInitial.hasRemove ||
    !definedNameAndLink.linkDialogInitial.captions.includes("表示文字列(T):") ||
    !definedNameAndLink.linkDialogInitial.captions.includes("アドレス(E):") ||
    !definedNameAndLink.linkDialogInitial.captionNowrap ||
    !definedNameAndLink.linkDialogInitial.captionAboveInput ||
    definedNameAndLink.linkDialogInitial.kindButtons.find((button) => button.kind === "web")?.pressed !== "true" ||
    definedNameAndLink.linkDocumentKind.activeKind !== "document" ||
    definedNameAndLink.linkDocumentKind.pressed !== "true" ||
    definedNameAndLink.linkDocumentKind.addressCaption !== "セル参照を入力してください(E):" ||
    definedNameAndLink.linkDocumentKind.urlValue !== definedNameAndLink.nameLinkAddress ||
    definedNameAndLink.linkDocumentKind.placeholder !== "A1" ||
    definedNameAndLink.linkDocumentKind.documentPanelHidden ||
    definedNameAndLink.linkDocumentKind.documentPanelCaption !== "またはドキュメント内の場所を選択してください(C):" ||
    !definedNameAndLink.linkDocumentKind.sheetNames.includes("CHT") ||
    !definedNameAndLink.linkDocumentKind.sheetNames.includes("AI製品紹介") ||
    definedNameAndLink.linkDocumentKind.selectedSheetName !== "CHT" ||
    definedNameAndLink.linkDocumentKind.selectedSheetAria !== "true" ||
    definedNameAndLink.linkDocumentSheetChoice.selectedSheetIndex !== "1" ||
    definedNameAndLink.linkDocumentSheetChoice.selectedSheetName !== "AI製品紹介" ||
    definedNameAndLink.linkDocumentSheetChoice.selectedSheetAria !== "true" ||
    definedNameAndLink.linkDocumentSheetChoice.urlValue !== definedNameAndLink.nameLinkAddress ||
    definedNameAndLink.linkEmailKind.activeKind !== "email" ||
    definedNameAndLink.linkEmailKind.pressed !== "true" ||
    definedNameAndLink.linkEmailKind.addressCaption !== "電子メール アドレス(E):" ||
    definedNameAndLink.linkEmailKind.placeholder !== "name@example.com" ||
    !definedNameAndLink.linkEmailKind.documentPanelHidden ||
    definedNameAndLink.linkWebKind.activeKind !== "web" ||
    definedNameAndLink.linkWebKind.pressed !== "true" ||
    definedNameAndLink.linkWebKind.addressCaption !== "アドレス(E):" ||
    definedNameAndLink.linkWebKind.urlValue !== "" ||
    !["T", "E"].every((key) => definedNameAndLink.linkDialogInitial.accessKeys.includes(key)) ||
    !definedNameAndLink.linkDialogInitial.accessKeysUnderlined ||
    definedNameAndLink.linkApply.text !== "WebSheet" ||
    definedNameAndLink.linkApply.href !== "https://example.com/docs" ||
    !definedNameAndLink.linkApply.decoration?.includes("underline") ||
    definedNameAndLink.linkContextMenu.disabledActions.includes("open-hyperlink") ||
    definedNameAndLink.openedLink.opened?.url !== "https://example.com/docs" ||
    definedNameAndLink.openedLink.opened?.target !== "_blank" ||
    definedNameAndLink.linkUndo.hasLink ||
    definedNameAndLink.linkUndo.text !== "42" ||
    definedNameAndLink.linkRedo.text !== "WebSheet" ||
    definedNameAndLink.linkRedo.href !== "https://example.com/docs"
  ) {
    throw new Error("Name definition and link dialogs should create usable defined names and hyperlinks with undo/redo.");
  }
  Object.assign(scenarioState, { result, row });
}

export {
  runAssertionsWorkbook,
};
