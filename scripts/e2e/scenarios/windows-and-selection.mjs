// Sequential E2E feature module.
import {
  STANDALONE_RUNTIME_CHANNEL,
  browser,
  pageBackgroundImagePath,
  path,
  readFile,
  standaloneRuntimeCurrentBaseUrl,
  standaloneRuntimeCurrentDir,
  viewDiagnosticsOnly,
} from "../../e2e-real-upload.mjs";
import {
  looksLikeStandaloneWorkbookHtml,
} from "../helpers/fixtures-and-model.mjs";
async function runWindowsAndSelection(scenarioState) {
  let { context, page, result, viewRibbon, viewFullScreenOn, fullScreenUndoRedo, viewFullScreenOff, sheetViewNew, sheetViewKept, sheetViewMenuBeforeRename, sheetViewRenamed, sheetViewExited, sheetViewReactivated, sheetViewUpdated, sheetViewMenuBeforeCopy, sheetViewDuplicated, viewZoomDialog, viewZoomApplied, viewZoomHundred, viewZoomSelection, viewZoomSelectionUndoRedo, viewRulerOn, pageLayoutHeaderFooterDialog, pageBreakPreview, pageBreakContextMenu, customViewCaptureState, customViewsMainBeforeAdd, customViewAddDialog, customViewsDialog, customViewSaveUndoRedo, customViewApplied, customViewApplyUndoRedo, customViewDeleteFlow, freezeMenuLabels, freezeMenuInitialMeta, freezeMenuTopRow, freezeMenuTopRowStatus, freezeMenuAccessKey, formulaBarUndoRedo, viewToggles, rtlVisualHitTest, gridlinesUndoRedo, headingsUndoRedo, zeroValuesUndoRedo, objectsUndoRedo, outlineUndoRedo, rightToLeftUndoRedo } = scenarioState;

  await page.evaluate(() => window.__webSheetDevSelectRange?.(4, 4, 4, 4));


  await page.click('[data-ribbon-command="view-freeze-panes"]');


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.freezePanes?.rows === 3 && window.__webSheetDevSheetModel?.("CHT")?.freezePanes?.columns === 3, null, { timeout: 5000 });


  const freezePanesAfterMain = await page.evaluate(() => window.__webSheetDevSheetModel?.("CHT")?.freezePanes || null);


  const frozenPaneSticky = await page.evaluate(async () => {
    const host = document.querySelector("#sheetHost");
    if (!host) return null;
    host.scrollTop = 520;
    host.scrollLeft = 520;
    host.dispatchEvent(new Event("scroll", { bubbles: true }));
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const hostRect = host.getBoundingClientRect();
    const cornerCell = document.querySelector('.sheet-cell[data-row="1"][data-col="1"]');
    const frozenRowCell = document.querySelector(".sheet-cell.is-frozen-row:not(.is-frozen-column)");
    const frozenColumnCell = document.querySelector(".sheet-cell.is-frozen-column:not(.is-frozen-row)");
    const cornerHeader = document.querySelector('.column-header[data-col="1"]');
    const rowHeader = document.querySelector('.row-header[data-row="1"]');
    const cellRect = cornerCell?.getBoundingClientRect();
    const rowCellRect = frozenRowCell?.getBoundingClientRect();
    const colCellRect = frozenColumnCell?.getBoundingClientRect();
    const result = {
      cornerPosition: cornerCell ? getComputedStyle(cornerCell).position : "",
      cornerClasses: cornerCell ? [...cornerCell.classList] : [],
      cornerTop: cellRect ? Math.round(cellRect.top - hostRect.top) : null,
      cornerLeft: cellRect ? Math.round(cellRect.left - hostRect.left) : null,
      rowCellPosition: frozenRowCell ? getComputedStyle(frozenRowCell).position : "",
      rowCellTop: rowCellRect ? Math.round(rowCellRect.top - hostRect.top) : null,
      columnCellPosition: frozenColumnCell ? getComputedStyle(frozenColumnCell).position : "",
      columnCellLeft: colCellRect ? Math.round(colCellRect.left - hostRect.left) : null,
      cornerHeaderSticky: cornerHeader?.classList.contains("is-frozen-column") || false,
      rowHeaderSticky: rowHeader?.classList.contains("is-frozen-row") || false,
    };
    host.scrollTop = 0;
    host.scrollLeft = 0;
    document.querySelector("#horizontalScrollbar")?.scrollTo({ left: 0, behavior: "auto" });
    host.dispatchEvent(new Event("scroll", { bubbles: true }));
    return result;
  });


  await page.click('[data-ribbon-command="view-split"]');


  await page.waitForSelector(".sheet-view-guide.is-split-pane.is-vertical", { timeout: 5000 });


  const splitPaneBeforeDrag = await page.evaluate(() => window.__webSheetDevSheetModel?.("CHT")?.splitPane || null);


  const verticalSplitBox = await page.locator(".sheet-view-guide.is-split-pane.is-vertical").boundingBox();


  if (verticalSplitBox) {
    await page.mouse.move(verticalSplitBox.x + 2, verticalSplitBox.y + Math.min(30, Math.max(2, verticalSplitBox.height / 2)));
    await page.mouse.down();
    await page.mouse.move(verticalSplitBox.x + 150, verticalSplitBox.y + Math.min(30, Math.max(2, verticalSplitBox.height / 2)), { steps: 4 });
    await page.mouse.up();
  }


  await page.waitForTimeout(150);


  if (!((await page.evaluate(() => window.__webSheetDevSheetModel?.("CHT")?.splitPane?.columns || 0)) > 3)) {
    await page.evaluate(() => {
      const guide = document.querySelector(".sheet-view-guide.is-split-pane.is-vertical");
      const rect = guide?.getBoundingClientRect();
      if (!guide || !rect) return;
      const y = rect.top + Math.min(30, Math.max(2, rect.height / 2));
      const startX = rect.left + Math.max(2, rect.width / 2);
      guide.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0, buttons: 1, pointerId: 31, clientX: startX, clientY: y }));
      window.dispatchEvent(new PointerEvent("pointermove", { bubbles: true, button: 0, buttons: 1, pointerId: 31, clientX: startX + 150, clientY: y }));
      window.dispatchEvent(new PointerEvent("pointerup", { bubbles: true, button: 0, pointerId: 31, clientX: startX + 150, clientY: y }));
    });
  }


  await page.waitForFunction(() => (window.__webSheetDevSheetModel?.("CHT")?.splitPane?.columns || 0) > 3, null, { timeout: 5000 });


  const horizontalSplitBox = await page.locator(".sheet-view-guide.is-split-pane.is-horizontal").boundingBox();


  if (horizontalSplitBox) {
    await page.mouse.move(horizontalSplitBox.x + Math.min(30, Math.max(2, horizontalSplitBox.width / 2)), horizontalSplitBox.y + 2);
    await page.mouse.down();
    await page.mouse.move(horizontalSplitBox.x + Math.min(30, Math.max(2, horizontalSplitBox.width / 2)), horizontalSplitBox.y + 82, { steps: 4 });
    await page.mouse.up();
  }


  await page.waitForFunction(() => (window.__webSheetDevSheetModel?.("CHT")?.splitPane?.rows || 0) > 3, null, { timeout: 5000 });


  const splitPaneDragged = await page.evaluate(() => ({
    pane: window.__webSheetDevSheetModel?.("CHT")?.splitPane || null,
    status: document.querySelector("#statusText")?.textContent || "",
    verticalGuideCount: document.querySelectorAll(".sheet-view-guide.is-split-pane.is-vertical").length,
    horizontalGuideCount: document.querySelectorAll(".sheet-view-guide.is-split-pane.is-horizontal").length,
  }));


  const sheetGroupViewBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupViewBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupViewBaseCount, { timeout: 5000 });


  const sheetGroupViewSheets = await page.evaluate((baseCount) =>
    [...document.querySelectorAll(".sheet-tab")]
      .slice(baseCount)
      .map((button) => ({ index: Number(button.dataset.sheetIndex), name: button.textContent.trim() })),
  sheetGroupViewBaseCount);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupViewSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupViewSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.click('[data-ribbon-tab="view"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "表示", null, { timeout: 10000 });


  await page.click('[data-ribbon-command="view-page-break"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevSheetModel?.(first)?.viewMode === "page-break" &&
    window.__webSheetDevSheetModel?.(second)?.viewMode === "page-break",
    sheetGroupViewSheets.map((sheet) => sheet.name),
    { timeout: 5000 },
  );


  await page.click('[data-ribbon-command="view-gridlines"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevSheetModel?.(first)?.showGridLines === false &&
    window.__webSheetDevSheetModel?.(second)?.showGridLines === false,
    sheetGroupViewSheets.map((sheet) => sheet.name),
    { timeout: 5000 },
  );


  const sheetGroupViewGridlines = await page.evaluate(([first, second]) => ({
    first: window.__webSheetDevSheetModel?.(first)?.showGridLines,
    second: window.__webSheetDevSheetModel?.(second)?.showGridLines,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupViewSheets.map((sheet) => sheet.name));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevSheetModel?.(first)?.showGridLines !== false &&
    window.__webSheetDevSheetModel?.(second)?.showGridLines !== false,
    sheetGroupViewSheets.map((sheet) => sheet.name),
    { timeout: 5000 },
  );


  const sheetGroupViewGridlinesUndo = await page.evaluate(([first, second]) => ({
    first: window.__webSheetDevSheetModel?.(first)?.showGridLines,
    second: window.__webSheetDevSheetModel?.(second)?.showGridLines,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupViewSheets.map((sheet) => sheet.name));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevSheetModel?.(first)?.showGridLines === false &&
    window.__webSheetDevSheetModel?.(second)?.showGridLines === false,
    sheetGroupViewSheets.map((sheet) => sheet.name),
    { timeout: 5000 },
  );


  const sheetGroupViewGridlinesRedo = await page.evaluate(([first, second]) => ({
    first: window.__webSheetDevSheetModel?.(first)?.showGridLines,
    second: window.__webSheetDevSheetModel?.(second)?.showGridLines,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupViewSheets.map((sheet) => sheet.name));


  await page.click('[data-ribbon-command="view-headings"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevSheetModel?.(first)?.showHeadings === false &&
    window.__webSheetDevSheetModel?.(second)?.showHeadings === false,
    sheetGroupViewSheets.map((sheet) => sheet.name),
    { timeout: 5000 },
  );


  await page.click('[data-ribbon-command="view-objects"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevSheetModel?.(first)?.showObjects === false &&
    window.__webSheetDevSheetModel?.(second)?.showObjects === false,
    sheetGroupViewSheets.map((sheet) => sheet.name),
    { timeout: 5000 },
  );


  await page.click('[data-ribbon-command="view-outline-symbols"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevSheetModel?.(first)?.showOutlineSymbols === false &&
    window.__webSheetDevSheetModel?.(second)?.showOutlineSymbols === false,
    sheetGroupViewSheets.map((sheet) => sheet.name),
    { timeout: 5000 },
  );


  await page.click('[data-ribbon-command="view-zero-values"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevSheetModel?.(first)?.showZeros === false &&
    window.__webSheetDevSheetModel?.(second)?.showZeros === false,
    sheetGroupViewSheets.map((sheet) => sheet.name),
    { timeout: 5000 },
  );


  await page.click('[data-ribbon-command="view-right-to-left"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevSheetModel?.(first)?.rightToLeft === true &&
    window.__webSheetDevSheetModel?.(second)?.rightToLeft === true,
    sheetGroupViewSheets.map((sheet) => sheet.name),
    { timeout: 5000 },
  );


  const sheetGroupViewShowOptions = await page.evaluate(([first, second]) => ({
    first: {
      showHeadings: window.__webSheetDevSheetModel?.(first)?.showHeadings,
      showZeros: window.__webSheetDevSheetModel?.(first)?.showZeros,
      showObjects: window.__webSheetDevSheetModel?.(first)?.showObjects,
      showOutlineSymbols: window.__webSheetDevSheetModel?.(first)?.showOutlineSymbols,
      rightToLeft: window.__webSheetDevSheetModel?.(first)?.rightToLeft,
    },
    second: {
      showHeadings: window.__webSheetDevSheetModel?.(second)?.showHeadings,
      showZeros: window.__webSheetDevSheetModel?.(second)?.showZeros,
      showObjects: window.__webSheetDevSheetModel?.(second)?.showObjects,
      showOutlineSymbols: window.__webSheetDevSheetModel?.(second)?.showOutlineSymbols,
      rightToLeft: window.__webSheetDevSheetModel?.(second)?.rightToLeft,
    },
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupViewSheets.map((sheet) => sheet.name));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(4, 4, 4, 4));


  await page.click('[data-ribbon-command="view-freeze-panes"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevSheetModel?.(first)?.freezePanes?.rows === 3 &&
    window.__webSheetDevSheetModel?.(first)?.freezePanes?.columns === 3 &&
    window.__webSheetDevSheetModel?.(second)?.freezePanes?.rows === 3 &&
    window.__webSheetDevSheetModel?.(second)?.freezePanes?.columns === 3,
    sheetGroupViewSheets.map((sheet) => sheet.name),
    { timeout: 5000 },
  );


  const sheetGroupViewFreeze = await page.evaluate(([first, second]) => ({
    first: window.__webSheetDevSheetModel?.(first)?.freezePanes || null,
    second: window.__webSheetDevSheetModel?.(second)?.freezePanes || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupViewSheets.map((sheet) => sheet.name));


  await page.click('[data-ribbon-command="view-split"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevSheetModel?.(first)?.splitPane?.row === 4 &&
    window.__webSheetDevSheetModel?.(first)?.splitPane?.col === 4 &&
    window.__webSheetDevSheetModel?.(second)?.splitPane?.row === 4 &&
    window.__webSheetDevSheetModel?.(second)?.splitPane?.col === 4,
    sheetGroupViewSheets.map((sheet) => sheet.name),
    { timeout: 5000 },
  );


  const sheetGroupViewSplit = await page.evaluate(([first, second]) => ({
    first: window.__webSheetDevSheetModel?.(first)?.splitPane || null,
    second: window.__webSheetDevSheetModel?.(second)?.splitPane || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupViewSheets.map((sheet) => sheet.name));


  await page.click('[data-ribbon-command="view-zoom"]');


  await page.waitForSelector("#zoomDialog", { timeout: 5000 });


  await page.keyboard.press("Alt+7");


  await page.waitForFunction(() => document.querySelector('#zoomDialog input[name="zoomDialogOption"][value="75"]')?.checked === true, null, { timeout: 5000 });


  await page.locator("#zoomDialog").press("Enter");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevSheetModel?.(first)?.zoomScale === 75 &&
    window.__webSheetDevSheetModel?.(second)?.zoomScale === 75,
    sheetGroupViewSheets.map((sheet) => sheet.name),
    { timeout: 5000 },
  );


  const sheetGroupViewZoom = await page.evaluate(([first, second]) => ({
    first: window.__webSheetDevSheetModel?.(first)?.zoomScale,
    second: window.__webSheetDevSheetModel?.(second)?.zoomScale,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupViewSheets.map((sheet) => sheet.name));


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupViewSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupViewSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupViewSheets[0].index}"]`, { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupViewBaseCount, { timeout: 5000 });


  await page.click('.sheet-tab[data-sheet-index="0"]');


  await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "CHT", null, { timeout: 5000 });


  await page.click('[data-ribbon-tab="view"]');


  const sheetGroupView = {
    gridlines: sheetGroupViewGridlines,
    gridlinesUndo: sheetGroupViewGridlinesUndo,
    gridlinesRedo: sheetGroupViewGridlinesRedo,
    showOptions: sheetGroupViewShowOptions,
    freeze: sheetGroupViewFreeze,
    split: sheetGroupViewSplit,
    zoom: sheetGroupViewZoom,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent.trim()),
      active: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
    })),
  };


  const workbookWindowBarSnapshot = async () => page.evaluate(() => {
    const bar = document.querySelector("#workbookWindowBar");
    if (!bar) return { exists: false, visible: false };
    const chipButtons = [...bar.querySelectorAll(".workbook-window-chip")];
    const previewSlots = [...bar.querySelectorAll(".workbook-window-preview-slot")];
    return {
      exists: true,
      visible: !bar.classList.contains("hidden") && getComputedStyle(bar).display !== "none",
      count: Number(bar.dataset.windowCount || 0),
      hiddenCount: Number(bar.dataset.hiddenCount || 0),
      mode: bar.dataset.windowArrangementMode || "",
      pairLeft: bar.dataset.sideBySideLeft || "",
      pairRight: bar.dataset.sideBySideRight || "",
      syncScrolling: bar.dataset.syncScrolling || "",
      chipIds: chipButtons.map((button) => button.dataset.workbookWindowId || ""),
      chipLabels: chipButtons.map((button) => button.textContent.replace(/\s+/g, " ").trim()),
      activeChipId: bar.querySelector(".workbook-window-chip.is-active")?.dataset.workbookWindowId || "",
      comparedChipIds: chipButtons.filter((button) => button.classList.contains("is-compared")).map((button) => button.dataset.workbookWindowId || ""),
      previewSlotIds: previewSlots.map((button) => button.dataset.workbookWindowId || ""),
      activePreviewId: bar.querySelector(".workbook-window-preview-slot.is-active")?.dataset.workbookWindowId || "",
      comparedPreviewIds: previewSlots.filter((button) => button.classList.contains("is-compared")).map((button) => button.dataset.workbookWindowId || ""),
      summary: bar.querySelector(".workbook-window-summary")?.textContent.replace(/\s+/g, " ").trim() || "",
      badges: [...bar.querySelectorAll(".workbook-window-badge")].map((node) => node.textContent.replace(/\s+/g, " ").trim()),
      height: Math.round(bar.getBoundingClientRect().height),
    };
  });


  const viewInitialWindowBar = await workbookWindowBarSnapshot();


  await page.click('[data-ribbon-command="view-new-window"]');


  await page.waitForFunction(() => (window.__webSheetDevWorkbookState?.().windowViews || []).length >= 1 && window.__webSheetDevWorkbookState?.().activeWindowId !== "main", null, { timeout: 5000 });


  const viewNewWindowState = await page.evaluate(async () => {
    const host = document.querySelector("#sheetHost");
    if (host) {
      host.scrollTop = 280;
      host.scrollLeft = 220;
      host.dispatchEvent(new Event("scroll", { bubbles: true }));
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    }
    const state = window.__webSheetDevWorkbookState?.() || {};
    const active = (state.windowViews || []).find((view) => view.id === state.activeWindowId) || null;
    return {
      activeWindowId: state.activeWindowId,
      mainWindowView: state.mainWindowView,
      activeWindow: active,
      windowCount: (state.windowViews || []).length + 1,
      hostScrollTop: host?.scrollTop || 0,
      hostScrollLeft: host?.scrollLeft || 0,
    };
  });


  const viewNewWindowBar = await workbookWindowBarSnapshot();


  await page.click('[data-ribbon-command="view-new-window"]');


  await page.waitForFunction(() => (window.__webSheetDevWorkbookState?.().windowViews || []).length >= 2, null, { timeout: 5000 });


  const viewSecondNewWindowState = await page.evaluate(() => {
    const state = window.__webSheetDevWorkbookState?.() || {};
    return {
      activeWindowId: state.activeWindowId,
      windowCount: (state.windowViews || []).length + 1,
      visibleCount: 1 + (state.windowViews || []).filter((view) => !view.hidden).length,
      names: [state.sourceName || "", ...((state.windowViews || []).map((view) => view.name || ""))].filter(Boolean),
      title: document.title,
      workbookName: document.querySelector("#workbookName")?.textContent || "",
    };
  });


  const viewSecondNewWindowBar = await workbookWindowBarSnapshot();


	  const workbookWindowBarSwitch = await page.evaluate(async () => {
	    const stateBefore = window.__webSheetDevWorkbookState?.() || {};
    const before = stateBefore.activeWindowId || "main";
    const target = [...document.querySelectorAll("#workbookWindowBar [data-workbook-window-id]")].find((button) => button.dataset.workbookWindowId !== before);
    if (!target) return { before, targetId: "", after: before, activeChipId: "", status: "" };
    const targetId = target.dataset.workbookWindowId || "";
    const waitFrames = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const clickRibbonCommand = async (command) => {
      document.querySelector(`[data-ribbon-command="${command}"]`)?.click();
      await waitFrames();
    };
    const setViewSheetCommandActive = async (command, active, readActive) => {
      const button = document.querySelector(`[data-ribbon-command="${command}"]`);
      if (!button || button.disabled) return;
      if (readActive() !== active) await clickRibbonCommand(command);
    };
    const setShowFormulasActive = async (active) => {
      if (activeSheetSnapshot().showFormulas === active) return;
      document.dispatchEvent(new KeyboardEvent("keydown", {
        key: "`",
        code: "Backquote",
        ctrlKey: true,
        bubbles: true,
        cancelable: true,
      }));
      await waitFrames();
      if (activeSheetSnapshot().showFormulas === active) return;
      document.querySelector('[data-ribbon-tab="formulas"]')?.click();
      await waitFrames();
      await clickRibbonCommand("formula-show-formulas");
      document.querySelector('[data-ribbon-tab="view"]')?.click();
      await waitFrames();
    };
    const activeSheetSnapshot = () => {
      const workbook = window.__webSheetDevWorkbookState?.() || {};
      const sheetName = workbook.activeSheetName || "CHT";
      return window.__webSheetDevSheetModel?.(sheetName) || {};
    };
	    window.__webSheetDevSelectRange?.(8, 8, 8, 8);
	    await setViewSheetCommandActive("view-gridlines", false, () => activeSheetSnapshot().showGridLines !== false);
	    await setViewSheetCommandActive("view-headings", false, () => activeSheetSnapshot().showHeadings !== false);
	    await setShowFormulasActive(true);
	    if (activeSheetSnapshot().splitPane) await clickRibbonCommand("view-split");
	    if (activeSheetSnapshot().freezePanes) await clickRibbonCommand("view-freeze-panes");
	    await clickRibbonCommand("view-freeze-panes");
	    target.click();
    await waitFrames();
	    window.__webSheetDevSelectRange?.(5, 3, 5, 3);
	    await setViewSheetCommandActive("view-gridlines", true, () => activeSheetSnapshot().showGridLines !== false);
	    await setViewSheetCommandActive("view-headings", true, () => activeSheetSnapshot().showHeadings !== false);
	    await setShowFormulasActive(false);
	    if (activeSheetSnapshot().freezePanes) await clickRibbonCommand("view-freeze-panes");
	    if (activeSheetSnapshot().splitPane) await clickRibbonCommand("view-split");
	    await clickRibbonCommand("view-split");
    const originalTarget = document.querySelector(`#workbookWindowBar [data-workbook-window-id="${CSS.escape(before)}"]`);
    originalTarget?.click();
    await waitFrames();
    const restoredOriginal = {
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      state: window.__webSheetDevWorkbookState?.() || {},
      sheet: activeSheetSnapshot(),
    };
    document.querySelector(`#workbookWindowBar [data-workbook-window-id="${CSS.escape(targetId)}"]`)?.click();
    await waitFrames();
    const stateAfter = window.__webSheetDevWorkbookState?.() || {};
    const sheetAfter = activeSheetSnapshot();
    const selectionFor = (state, id) => {
      if (id === "main") return state.mainWindowView?.selectionRange || null;
      return state.windowViews?.find?.((view) => view.id === id)?.selectionRange || null;
    };
    const viewFor = (state, id) => {
      if (id === "main") return state.mainWindowView || null;
      return state.windowViews?.find?.((view) => view.id === id) || null;
    };
    return {
      before,
      targetId,
      after: stateAfter.activeWindowId || "main",
      activeChipId: document.querySelector("#workbookWindowBar .workbook-window-chip.is-active")?.dataset.workbookWindowId || "",
      previewActiveId: document.querySelector("#workbookWindowBar .workbook-window-preview-slot.is-active")?.dataset.workbookWindowId || "",
      restoredOriginalLabel: restoredOriginal.label,
      targetSelectionLabel: document.querySelector("#selectedCellLabel")?.textContent || "",
      beforeStoredSelection: selectionFor(restoredOriginal.state, before),
      targetStoredSelection: selectionFor(stateAfter, targetId),
	      restoredOriginalSheetView: {
	        showGridLines: restoredOriginal.sheet.showGridLines,
	        showHeadings: restoredOriginal.sheet.showHeadings,
	        showFormulas: restoredOriginal.sheet.showFormulas,
	        freezePanes: restoredOriginal.sheet.freezePanes || null,
	        splitPane: restoredOriginal.sheet.splitPane || null,
	      },
	      targetSheetView: {
	        showGridLines: sheetAfter.showGridLines,
	        showHeadings: sheetAfter.showHeadings,
	        showFormulas: sheetAfter.showFormulas,
	        freezePanes: sheetAfter.freezePanes || null,
	        splitPane: sheetAfter.splitPane || null,
	      },
      beforeStoredView: viewFor(restoredOriginal.state, before),
      targetStoredView: viewFor(stateAfter, targetId),
	      status: document.querySelector("#statusText")?.textContent || "",
	    };
	  });


	  const customWindowSource = await page.evaluate(async () => {
	    const waitFrames = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
	    const workbook = window.__webSheetDevWorkbookState?.() || {};
	    const targetId = (workbook.windowViews || []).find((view) => !view.hidden)?.id || "";
	    if (!targetId) return { skipped: true };
	    document.querySelector(`#workbookWindowBar [data-workbook-window-id="${CSS.escape(targetId)}"]`)?.click();
	    await waitFrames();
	    const clickCommand = async (command) => {
	      document.querySelector(`[data-ribbon-command="${command}"]`)?.click();
	      await waitFrames();
	    };
	    const activeSheetSnapshot = () => {
	      const state = window.__webSheetDevWorkbookState?.() || {};
	      return window.__webSheetDevSheetModel?.(state.activeSheetName || "CHT") || {};
	    };
	    let sheet = activeSheetSnapshot();
	    if (sheet.splitPane) await clickCommand("view-split");
	    sheet = activeSheetSnapshot();
	    if (sheet.freezePanes) await clickCommand("view-freeze-panes");
	    sheet = activeSheetSnapshot();
	    if (sheet.showGridLines === false) await clickCommand("view-gridlines");
	    window.__webSheetDevSelectRange?.(6, 6, 6, 6);
	    await waitFrames();
	    await clickCommand("view-freeze-panes");
	    window.__webSheetDevSelectRange?.(12, 12, 12, 12);
	    const host = document.querySelector("#sheetHost");
	    if (host) {
	      host.scrollTop = 180;
	      host.scrollLeft = 160;
	      host.dispatchEvent(new Event("scroll", { bubbles: true }));
	      await waitFrames();
	    }
	    const state = window.__webSheetDevWorkbookState?.() || {};
	    const activeWindow = (state.windowViews || []).find((view) => view.id === state.activeWindowId) || state.mainWindowView || null;
	    return {
	      skipped: false,
	      activeWindowId: state.activeWindowId || "main",
	      selectionLabel: document.querySelector("#selectedCellLabel")?.textContent?.trim() || "",
	      activeWindow,
	      sheet: activeSheetSnapshot(),
	      scrollTop: host?.scrollTop || 0,
	      scrollLeft: host?.scrollLeft || 0,
	    };
	  });


	  await page.click('[data-ribbon-command="view-custom-views"]');


	  await page.waitForSelector('[data-data-dialog="custom-views"]', { timeout: 5000 });


	  await page.locator("[data-data-dialog-action='custom-view-add']").click();


	  await page.waitForSelector('[data-data-dialog="custom-view-add"]', { timeout: 5000 });


	  await page.locator("[data-data-field='customViewName']").fill("ウィンドウ復元テスト");


	  await page.locator("[data-data-dialog-action='custom-view-add-ok']").click();


	  await page.waitForFunction(() => (window.__webSheetDevWorkbookState?.().customViews || []).some((view) => view.name === "ウィンドウ復元テスト"), null, { timeout: 5000 });


	  await page.waitForSelector('[data-data-dialog="custom-views"]', { timeout: 5000 });


	  await page.locator("[data-data-dialog-action='cancel']").click();


	  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


	  const customWindowMutated = await page.evaluate(async () => {
	    const waitFrames = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
	    const clickCommand = async (command) => {
	      document.querySelector(`[data-ribbon-command="${command}"]`)?.click();
	      await waitFrames();
	    };
	    document.querySelector('#workbookWindowBar [data-workbook-window-id="main"]')?.click();
	    await waitFrames();
	    let sheet = window.__webSheetDevSheetModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "CHT") || {};
	    if (sheet.freezePanes) await clickCommand("view-freeze-panes");
	    sheet = window.__webSheetDevSheetModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "CHT") || {};
	    if (sheet.showGridLines !== false) await clickCommand("view-gridlines");
	    window.__webSheetDevSelectRange?.(1, 1, 1, 1);
	    const host = document.querySelector("#sheetHost");
	    if (host) {
	      host.scrollTop = 0;
	      host.scrollLeft = 0;
	      host.dispatchEvent(new Event("scroll", { bubbles: true }));
	      await waitFrames();
	    }
	    const state = window.__webSheetDevWorkbookState?.() || {};
	    return {
	      activeWindowId: state.activeWindowId || "main",
	      selectionLabel: document.querySelector("#selectedCellLabel")?.textContent?.trim() || "",
	      sheet: window.__webSheetDevSheetModel?.(state.activeSheetName || "CHT") || {},
	    };
	  });


	  await page.click('[data-ribbon-command="view-custom-views"]');


	  await page.waitForSelector('[data-data-dialog="custom-views"]', { timeout: 5000 });


	  await page.locator("[data-data-field='customViewIndex']").selectOption({ label: "ウィンドウ復元テスト" });


	  await page.locator("[data-data-dialog-action='custom-view-show']").click();


	  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


	  const customWindowRestored = await page.evaluate(async () => {
	    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
	    const state = window.__webSheetDevWorkbookState?.() || {};
	    const sheet = window.__webSheetDevSheetModel?.(state.activeSheetName || "CHT") || {};
	    const activeWindow = state.activeWindowId === "main"
	      ? state.mainWindowView
	      : (state.windowViews || []).find((view) => view.id === state.activeWindowId) || null;
	    const host = document.querySelector("#sheetHost");
	    return {
	      activeWindowId: state.activeWindowId || "main",
	      activeWindow,
	      selectionLabel: document.querySelector("#selectedCellLabel")?.textContent?.trim() || "",
	      sheet: {
	        showGridLines: sheet.showGridLines,
	        freezePanes: sheet.freezePanes || null,
	        splitPane: sheet.splitPane || null,
	      },
	      scrollTop: host?.scrollTop || 0,
	      scrollLeft: host?.scrollLeft || 0,
	      status: document.querySelector("#statusText")?.textContent || "",
	    };
	  });


	  await page.click('[data-ribbon-command="view-arrange-all"]');


  await page.waitForSelector('[data-data-dialog="view-arrange-window"]', { timeout: 5000 });


  await page.waitForSelector("[data-data-dialog='view-arrange-window'] [data-access-key='V']", { timeout: 5000 });


  const arrangeWindowDialog = await page.evaluate(() => ({
	    title: document.querySelector("#arrangeWindowTitle")?.textContent || "",
	    options: [...document.querySelectorAll("input[name='viewArrangeMode']")].map((input) => input.value),
	    selected: document.querySelector("input[name='viewArrangeMode']:checked")?.value || "",
	    labels: [...document.querySelectorAll("[data-data-dialog='view-arrange-window'] .data-dialog-check")].map((label) => label.textContent.replace(/\s+/g, " ").trim()),
	    activeWorkbookOnlyChecked: document.querySelector("[data-data-field='activeWorkbookOnly']")?.checked,
	    activeWorkbookOnlyDisabled: document.querySelector("[data-data-field='activeWorkbookOnly']")?.disabled,
	    accessKeys: [...document.querySelectorAll("[data-data-dialog='view-arrange-window'] [data-access-key]")].map((item) => item.dataset.accessKey || ""),
	    accessKeysUnderlined: [...document.querySelectorAll("[data-data-dialog='view-arrange-window'] .menu-access-key")].every((item) =>
	      getComputedStyle(item).textDecorationLine.includes("underline"),
	    ),
  }));


  await page.keyboard.press("Alt+V");


  await page.waitForFunction(() => document.querySelector('input[name="viewArrangeMode"][value="vertical"]')?.checked === true, null, { timeout: 5000 });


  await page.locator('[data-data-dialog="view-arrange-window"]').press("Enter");


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  const arrangeWindowState = await page.evaluate(() => window.__webSheetDevWorkbookState?.().windowArrangement || null);


  const arrangeWindowBar = await workbookWindowBarSnapshot();


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().windowArrangement?.mode !== "vertical", null, { timeout: 5000 });


  const arrangeWindowUndo = await page.evaluate(() => ({
    arrangement: window.__webSheetDevWorkbookState?.().windowArrangement || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  arrangeWindowUndo.bar = await workbookWindowBarSnapshot();


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().windowArrangement?.mode === "vertical", null, { timeout: 5000 });


  const arrangeWindowRedo = await page.evaluate(() => ({
    arrangement: window.__webSheetDevWorkbookState?.().windowArrangement || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  arrangeWindowRedo.bar = await workbookWindowBarSnapshot();


  await page.click('[data-ribbon-command="view-side-by-side"]');


  await page.waitForSelector('[data-data-dialog="view-side-by-side-window"]', { timeout: 5000 });


  await page.waitForSelector("[data-data-dialog='view-side-by-side-window'] [data-access-key='C']", { timeout: 5000 });


  const sideBySideWindowDialog = await page.evaluate(() => ({
    title: document.querySelector("#sideBySideWindowTitle")?.textContent || "",
    activeWindowId: window.__webSheetDevWorkbookState?.().activeWindowId || "",
    options: [...document.querySelectorAll("[data-data-field='sideBySideWindowId'] option")].map((option) => ({ value: option.value, text: option.textContent.trim() })),
    accessKeys: [...document.querySelectorAll("[data-data-dialog='view-side-by-side-window'] [data-access-key]")].map((item) => item.dataset.accessKey || ""),
    accessKeysUnderlined: [...document.querySelectorAll("[data-data-dialog='view-side-by-side-window'] .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
  }));


  await page.locator("[data-data-field='sideBySideWindowId']").selectOption({ index: 0 });


  const selectedSideBySideWindowId = await page.locator("[data-data-field='sideBySideWindowId']").inputValue();


  await page.locator("[data-data-field='sideBySideWindowId']").dispatchEvent("keydown", { key: "Enter", bubbles: true });


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  await page.waitForFunction(() => Boolean(window.__webSheetDevWorkbookState?.().viewSideBySide), null, { timeout: 5000 });


  const sideBySideDialogApplied = await page.evaluate((selectedId) => {
    const state = window.__webSheetDevWorkbookState?.() || {};
    return {
      selectedId,
      activeWindowId: state.activeWindowId,
      pair: state.viewSideBySide || null,
      arrangement: state.windowArrangement || null,
    };
  }, selectedSideBySideWindowId);


  const sideBySideWindowBar = await workbookWindowBarSnapshot();


  await page.click('[data-ribbon-command="view-sync-scrolling"]');


  const syncScrollState = await page.evaluate(async () => {
    const host = document.querySelector("#sheetHost");
    if (host) {
      host.scrollTop = 360;
      host.scrollLeft = 260;
      host.dispatchEvent(new Event("scroll", { bubbles: true }));
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    }
    const state = window.__webSheetDevWorkbookState?.() || {};
    const records = [
      { id: "main", ...(state.mainWindowView || {}) },
      ...((state.windowViews || []).map((view) => ({ ...view }))),
    ];
    const pair = state.viewSideBySide || {};
    const left = records.find((view) => view.id === pair.left) || null;
    const right = records.find((view) => view.id === pair.right) || null;
    return {
      activeWindowId: state.activeWindowId,
      sideBySide: state.viewSideBySide,
      synchronousScrolling: state.synchronousScrolling,
      windowArrangement: state.windowArrangement,
      left,
      right,
      sameScroll: Boolean(left && right && left.scrollTop === right.scrollTop && left.scrollLeft === right.scrollLeft),
    };
  });


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().synchronousScrolling === false, null, { timeout: 5000 });


  const syncScrollUndoRedo = {
    undo: await page.evaluate(() => ({
      synchronousScrolling: window.__webSheetDevWorkbookState?.().synchronousScrolling,
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().synchronousScrolling === true, null, { timeout: 5000 });


  syncScrollUndoRedo.redo = await page.evaluate(() => ({
    synchronousScrolling: window.__webSheetDevWorkbookState?.().synchronousScrolling,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  syncScrollUndoRedo.redo.bar = await workbookWindowBarSnapshot();


  await page.click('[data-ribbon-command="view-reset-window-position"]');


  const resetWindowState = await page.evaluate(() => window.__webSheetDevWorkbookState?.().windowArrangement || null);


  const resetWindowBar = await workbookWindowBarSnapshot();


  await page.click('[data-ribbon-command="view-switch-windows"]');


  await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });


  const firstSwitchAction = await page.locator("#viewCommandMenu [data-view-menu-action]").first().getAttribute("data-view-menu-action");


  const secondSwitchAction = await page.locator("#viewCommandMenu [data-view-menu-action]").nth(1).getAttribute("data-view-menu-action");


  const switchWindowMenu = await page.evaluate(() => {
    const buttons = [...document.querySelectorAll("#viewCommandMenu [data-view-menu-action]")];
    const switchButtons = buttons.filter((button) => String(button.dataset.viewMenuAction || "").startsWith("switch:"));
    const dialogButton = buttons.find((button) => button.dataset.viewMenuAction === "switch-dialog");
    return {
      actions: buttons.map((button) => button.dataset.viewMenuAction || ""),
      labels: buttons.map((button) => button.textContent.replace(/\s+/g, " ").trim()),
      switchLabels: switchButtons.map((button) => button.textContent.replace(/\s+/g, " ").trim()),
      switchDescriptions: switchButtons.map((button) => button.querySelector(".home-command-menu-text small")?.textContent.trim() || ""),
      dialogLabel: dialogButton?.textContent.replace(/\s+/g, " ").trim() || "",
      dialogDescription: dialogButton?.querySelector(".home-command-menu-text small")?.textContent.trim() || "",
      checkedAction: document.querySelector("#viewCommandMenu [aria-checked='true']")?.dataset.viewMenuAction || "",
      accessKeys: [...document.querySelectorAll("#viewCommandMenu [data-access-key]")].map((button) => button.dataset.accessKey || ""),
      accessKeysUnderlined: [...document.querySelectorAll("#viewCommandMenu .menu-access-key")].every((item) =>
        getComputedStyle(item).textDecorationLine.includes("underline"),
      ),
    };
  });


  let switchWindowKeyboardState = { skipped: true };


  if (secondSwitchAction) {
    await page.keyboard.press("2");
    await page.waitForFunction(
      (expectedId) => window.__webSheetDevWorkbookState?.().activeWindowId === expectedId,
      String(secondSwitchAction).replace("switch:", ""),
      { timeout: 5000 },
    );
    switchWindowKeyboardState = await page.evaluate((action) => {
      const state = window.__webSheetDevWorkbookState?.() || {};
      return {
        skipped: false,
        expectedWindowId: String(action || "").replace("switch:", ""),
        activeWindowId: state.activeWindowId || "main",
        menuClosed: !document.querySelector("#viewCommandMenu"),
        status: document.querySelector("#statusText")?.textContent || "",
      };
    }, secondSwitchAction);
    await page.click('[data-ribbon-command="view-switch-windows"]');
    await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });
  }


  await page.locator("#viewCommandMenu [data-view-menu-action]").first().click();


  const switchWindowState = await page.evaluate(async (action) => {
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const state = window.__webSheetDevWorkbookState?.() || {};
    const host = document.querySelector("#sheetHost");
    return {
      expectedWindowId: String(action || "").replace("switch:", ""),
      activeWindowId: state.activeWindowId,
      hostScrollTop: host?.scrollTop || 0,
      hostScrollLeft: host?.scrollLeft || 0,
      mainWindowView: state.mainWindowView,
      title: document.title,
      workbookName: document.querySelector("#workbookName")?.textContent || "",
    };
  }, firstSwitchAction);


  await page.click('[data-ribbon-command="view-switch-windows"]');


  await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });


  await page.locator('#viewCommandMenu [data-view-menu-action="switch-dialog"]').click();


  await page.waitForSelector('[data-data-dialog="view-switch-window"]', { timeout: 5000 });


  await page.waitForSelector("[data-data-dialog='view-switch-window'] [data-access-key='W']", { timeout: 5000 });


  const switchWindowDialog = await page.evaluate(() => {
    const state = window.__webSheetDevWorkbookState?.() || {};
    const current = state.activeWindowId || "main";
    const select = document.querySelector("[data-data-field='switchWindowId']");
    const options = [...(select?.querySelectorAll("option") || [])].map((option) => ({ value: option.value, text: option.textContent.trim(), selected: option.selected }));
    const target = options.find((option) => option.value !== current)?.value || options[0]?.value || "";
    if (select && target) {
      select.value = target;
      select.dispatchEvent(new Event("change", { bubbles: true }));
    }
    return {
      title: document.querySelector("#switchWindowTitle")?.textContent || "",
      caption: document.querySelector("[data-data-dialog='view-switch-window'] .excel-dialog-caption")?.textContent.replace(/\s+/g, " ").trim() || "",
      options,
      current,
      target,
      selected: select?.value || "",
      size: select?.getAttribute("size") || "",
      summary: document.querySelector("[data-data-dialog='view-switch-window'] .data-dialog-summary")?.textContent.trim() || "",
      accessKeys: [...document.querySelectorAll("[data-data-dialog='view-switch-window'] [data-access-key]")].map((item) => item.dataset.accessKey || ""),
      accessKeysUnderlined: [...document.querySelectorAll("[data-data-dialog='view-switch-window'] .menu-access-key")].every((item) =>
        getComputedStyle(item).textDecorationLine.includes("underline"),
      ),
    };
  });


  await page.locator("[data-data-field='switchWindowId']").press("Enter");


  await page.waitForFunction(
    (expectedId) => window.__webSheetDevWorkbookState?.().activeWindowId === expectedId,
    switchWindowDialog.target,
    { timeout: 5000 },
  );


  const switchWindowDialogState = await page.evaluate((expectedId) => ({
    expectedWindowId: expectedId,
    activeWindowId: window.__webSheetDevWorkbookState?.().activeWindowId || "main",
    dialogClosed: !document.querySelector("[data-data-dialog='view-switch-window']"),
    status: document.querySelector("#statusText")?.textContent || "",
  }), switchWindowDialog.target);


  await page.click('[data-ribbon-command="view-hide-window"]');


  await page.waitForFunction(() => {
    const state = window.__webSheetDevWorkbookState?.() || {};
    return Boolean(state.mainWindowHidden) || (state.windowViews || []).some((view) => view.hidden);
  }, null, { timeout: 5000 });


  const hideWindowState = await page.evaluate(() => {
    const state = window.__webSheetDevWorkbookState?.() || {};
    return {
      activeWindowId: state.activeWindowId,
      mainHidden: state.mainWindowHidden === true,
      hiddenIds: (state.windowViews || []).filter((view) => view.hidden).map((view) => view.id),
      viewSideBySide: state.viewSideBySide || null,
      synchronousScrolling: state.synchronousScrolling,
      arrangement: state.windowArrangement || null,
    };
  });


  const hideWindowBar = await workbookWindowBarSnapshot();


  await page.click('[data-ribbon-command="view-unhide-window"]');


  await page.waitForSelector('[data-data-dialog="view-unhide-window"]', { timeout: 5000 });


  await page.waitForSelector("[data-data-dialog='view-unhide-window'] [data-access-key='U']", { timeout: 5000 });


  const unhideWindowDialog = await page.evaluate(() => ({
    options: [...document.querySelectorAll("[data-data-field='windowId'] option")].map((option) => option.textContent.trim()),
    accessKeys: [...document.querySelectorAll("[data-data-dialog='view-unhide-window'] [data-access-key]")].map((item) => item.dataset.accessKey || ""),
    accessKeysUnderlined: [...document.querySelectorAll("[data-data-dialog='view-unhide-window'] .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
  }));


  await page.locator("[data-data-field='windowId']").press("Enter");


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  const unhideWindowState = await page.evaluate(() => {
    const state = window.__webSheetDevWorkbookState?.() || {};
    return {
      activeWindowId: state.activeWindowId,
      mainHidden: state.mainWindowHidden === true,
      hiddenIds: (state.windowViews || []).filter((view) => view.hidden).map((view) => view.id),
      viewSideBySide: state.viewSideBySide || null,
      synchronousScrolling: state.synchronousScrolling,
      arrangement: state.windowArrangement || null,
    };
	  });


	  const unhideWindowBar = await workbookWindowBarSnapshot();


	  await page.keyboard.press("Control+Z");


	  await page.waitForFunction(() => {
	    const state = window.__webSheetDevWorkbookState?.() || {};
	    return Boolean(state.mainWindowHidden) || (state.windowViews || []).some((view) => view.hidden);
	  }, null, { timeout: 5000 });


	  const unhideWindowUndoRedo = {
	    undo: await page.evaluate(() => {
	      const state = window.__webSheetDevWorkbookState?.() || {};
	      return {
	        activeWindowId: state.activeWindowId,
	        mainHidden: state.mainWindowHidden === true,
	        hiddenIds: (state.windowViews || []).filter((view) => view.hidden).map((view) => view.id),
	        status: document.querySelector("#statusText")?.textContent || "",
	      };
	    }),
	  };


	  unhideWindowUndoRedo.undo.bar = await workbookWindowBarSnapshot();


	  await page.keyboard.press("Control+Y");


	  await page.waitForFunction(() => {
	    const state = window.__webSheetDevWorkbookState?.() || {};
	    return state.mainWindowHidden !== true && !(state.windowViews || []).some((view) => view.hidden);
	  }, null, { timeout: 5000 });


  unhideWindowUndoRedo.redo = await page.evaluate(() => {
    const state = window.__webSheetDevWorkbookState?.() || {};
    return {
      activeWindowId: state.activeWindowId,
	      mainHidden: state.mainWindowHidden === true,
	      hiddenIds: (state.windowViews || []).filter((view) => view.hidden).map((view) => view.id),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  });


  unhideWindowUndoRedo.redo.bar = await workbookWindowBarSnapshot();


  await page.click('[data-ribbon-command="view-new-window"]');


  await page.waitForFunction(() => {
    const state = window.__webSheetDevWorkbookState?.() || {};
    return (state.windowViews || []).filter((view) => !view.hidden).length + (state.mainWindowHidden ? 0 : 1) >= 4;
  }, null, { timeout: 5000 });


  const closeWindowTarget = await page.evaluate(() => window.__webSheetDevWorkbookState?.().activeWindowId || "");


  await page.locator(`#workbookWindowBar [data-workbook-window-close="${closeWindowTarget}"]`).click();


  await page.waitForFunction(
    (closedId) => {
      const state = window.__webSheetDevWorkbookState?.() || {};
      return state.activeWindowId !== closedId && !(state.windowViews || []).some((view) => view.id === closedId);
    },
    closeWindowTarget,
    { timeout: 5000 },
  );


  const closeWindowState = await page.evaluate((closedId) => {
    const state = window.__webSheetDevWorkbookState?.() || {};
    return {
      closedId,
      activeWindowId: state.activeWindowId || "main",
      exists: (state.windowViews || []).some((view) => view.id === closedId),
      windowCount: (state.windowViews || []).length + (state.mainWindowHidden ? 0 : 1),
      status: document.querySelector("#statusText")?.textContent || "",
      closeButtons: document.querySelectorAll("#workbookWindowBar [data-workbook-window-close]").length,
    };
  }, closeWindowTarget);


  closeWindowState.bar = await workbookWindowBarSnapshot();


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    (closedId) => (window.__webSheetDevWorkbookState?.().windowViews || []).some((view) => view.id === closedId),
    closeWindowTarget,
    { timeout: 5000 },
  );


  const closeWindowUndoRedo = {
    undo: await page.evaluate((closedId) => {
      const state = window.__webSheetDevWorkbookState?.() || {};
      return {
        activeWindowId: state.activeWindowId || "main",
        exists: (state.windowViews || []).some((view) => view.id === closedId),
        windowCount: (state.windowViews || []).length + (state.mainWindowHidden ? 0 : 1),
        status: document.querySelector("#statusText")?.textContent || "",
      };
    }, closeWindowTarget),
  };


  closeWindowUndoRedo.undo.bar = await workbookWindowBarSnapshot();


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(
    (closedId) => !(window.__webSheetDevWorkbookState?.().windowViews || []).some((view) => view.id === closedId),
    closeWindowTarget,
    { timeout: 5000 },
  );


  closeWindowUndoRedo.redo = await page.evaluate((closedId) => {
    const state = window.__webSheetDevWorkbookState?.() || {};
    return {
      activeWindowId: state.activeWindowId || "main",
      exists: (state.windowViews || []).some((view) => view.id === closedId),
      windowCount: (state.windowViews || []).length + (state.mainWindowHidden ? 0 : 1),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, closeWindowTarget);


  closeWindowUndoRedo.redo.bar = await workbookWindowBarSnapshot();


  await page.click('[data-ribbon-command="view-macros"]');


  await page.waitForSelector('[data-data-dialog="macros"]', { timeout: 5000 });


  const macrosInitialDialog = await page.evaluate(() => ({
    title: document.querySelector("#macrosTitle")?.textContent || "",
    macroName: document.querySelector("[data-data-field='macroName']")?.value || "",
    runDisabled: document.querySelector("[data-data-dialog-action='macro-run']")?.disabled,
    editDisabled: document.querySelector("[data-data-dialog-action='macro-edit']")?.disabled,
    optionsDisabled: document.querySelector("[data-data-dialog-action='macro-options']")?.disabled,
    createDisabled: document.querySelector("[data-data-dialog-action='macro-create']")?.disabled,
  }));


  await page.locator("[data-data-field='macroName']").fill("ViewMacro");


  await page.locator("[data-data-dialog-action='macro-create']").click();


  await page.waitForFunction(() => (window.__webSheetDevWorkbookState?.().macros || []).some((macro) => macro.name === "ViewMacro"), null, { timeout: 5000 });


  const macrosAfterCreateDialog = await page.evaluate(() => ({
    selectedName: document.querySelector("[data-data-field='macroName']")?.value || "",
    optionCount: document.querySelectorAll("[data-data-field='macroIndex'] option").length,
    runDisabled: document.querySelector("[data-data-dialog-action='macro-run']")?.disabled,
    editDisabled: document.querySelector("[data-data-dialog-action='macro-edit']")?.disabled,
    optionsDisabled: document.querySelector("[data-data-dialog-action='macro-options']")?.disabled,
  }));


  await page.locator("[data-data-field='macroIndex']").focus();


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => !(window.__webSheetDevWorkbookState?.().macros || []).some((macro) => macro.name === "ViewMacro"), null, { timeout: 5000 });


  const macroCreateUndoRedo = {
    undo: await page.evaluate(() => ({
      exists: (window.__webSheetDevWorkbookState?.().macros || []).some((macro) => macro.name === "ViewMacro"),
      selectedName: document.querySelector("[data-data-field='macroName']")?.value || "",
      runDisabled: document.querySelector("[data-data-dialog-action='macro-run']")?.disabled,
    })),
  };


  await page.locator("[data-data-dialog-action='macro-create']").focus();


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => (window.__webSheetDevWorkbookState?.().macros || []).some((macro) => macro.name === "ViewMacro"), null, { timeout: 5000 });


  macroCreateUndoRedo.redo = await page.evaluate(() => ({
    exists: (window.__webSheetDevWorkbookState?.().macros || []).some((macro) => macro.name === "ViewMacro"),
    selectedName: document.querySelector("[data-data-field='macroName']")?.value || "",
    runDisabled: document.querySelector("[data-data-dialog-action='macro-run']")?.disabled,
  }));


  await page.locator("[data-data-dialog-action='macro-options']").click();


  await page.waitForSelector('[data-data-dialog="macro-options"]', { timeout: 5000 });


  const macroOptionsDialog = await page.evaluate(() => ({
    title: document.querySelector("#macroOptionsTitle")?.textContent || "",
    name: document.querySelector(".macro-options-name")?.textContent || "",
  }));


  await page.locator("[data-data-field='macroShortcut']").fill("q");


  await page.locator("[data-data-field='macroDescription']").fill("表示タブのテスト用マクロ");


  await page.locator("[data-data-dialog-action='macro-options-ok']").click();


  await page.waitForSelector('[data-data-dialog="macros"]', { timeout: 5000 });


  await page.locator("[data-data-field='macroIndex']").focus();


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => {
    const macro = (window.__webSheetDevWorkbookState?.().macros || []).find((item) => item.name === "ViewMacro");
    return macro && !macro.shortcut && !macro.description;
  }, null, { timeout: 5000 });


  const macroOptionsUndoRedo = {
    undo: await page.evaluate(() => ({
      macro: (window.__webSheetDevWorkbookState?.().macros || []).find((item) => item.name === "ViewMacro") || null,
      dialogVisible: Boolean(document.querySelector("[data-data-dialog='macros']")),
    })),
  };


  await page.locator("[data-data-field='macroIndex']").focus();


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => {
    const macro = (window.__webSheetDevWorkbookState?.().macros || []).find((item) => item.name === "ViewMacro");
    return macro?.shortcut === "q" && /表示タブ/.test(macro.description || "");
  }, null, { timeout: 5000 });


  macroOptionsUndoRedo.redo = await page.evaluate(() => ({
    macro: (window.__webSheetDevWorkbookState?.().macros || []).find((item) => item.name === "ViewMacro") || null,
    dialogVisible: Boolean(document.querySelector("[data-data-dialog='macros']")),
  }));


  await page.locator("[data-data-dialog-action='macro-edit']").click();


  await page.waitForSelector('[data-data-dialog="macro-editor"]', { timeout: 5000 });


  const macroEditorDialog = await page.evaluate(() => ({
    title: document.querySelector("#macroEditorTitle")?.textContent || "",
  }));


  await page.locator("[data-data-field='macroCode']").fill("Sub ViewMacro()\n  Range(\"I6\").Value = \"MacroRun\"\n  Range(\"J6\").Value = 42\n  ActiveSheet.Range(\"I6:J6\").Font.Bold = True\n  ActiveSheet.Range(\"I6:J6\").Interior.Color = RGB(255, 255, 0)\n  ActiveSheet.Range(\"I6:J6\").HorizontalAlignment = xlRight\n  ActiveSheet.Range(\"I6:J6\").WrapText = True\n  Range(\"K6\").Value = \"消去\"\n  Range(\"K6\").Interior.Color = RGB(0, 255, 0)\n  Range(\"K6\").ClearContents\n  Range(\"K6\").ClearFormats\n  Range(\"L6\").Value = \"CopySrc\"\n  Range(\"L6\").Font.Italic = True\n  Range(\"L6\").Copy Destination:=Range(\"M6\")\n  Range(\"L6\").Clear\n  Range(\"N6\").Value = \"CutSrc\"\n  Range(\"N6\").Cut Destination:=Range(\"O6\")\n  Range(\"P6:Q6\").Merge\n  Range(\"P6\").Value = \"Merged\"\n  Range(\"P6:Q6\").UnMerge\n  Range(\"I6\").Offset(1, 0).Value = \"Offset\"\n  Range(\"I6\").Offset(2, 0).Resize(1, 2).Value = \"Resize\"\n  Range(\"R6\").Value = \"Font\"\n  Range(\"R6\").Font.Name = \"MS PGothic\"\n  Range(\"R6\").Font.Size = 14\n  Range(\"R6\").Font.Strikethrough = True\n  Range(\"R6:S6\").Borders.LineStyle = xlContinuous\n  Range(\"R6:S6\").Borders.Weight = xlThick\n  Range(\"R6:S6\").Borders.Color = RGB(255, 0, 0)\n  Columns(\"T:T\").ColumnWidth = 12\n  Rows(\"11:11\").RowHeight = 22\n  Columns(\"U:U\").Hidden = True\n  Columns(\"U:U\").Hidden = False\n  Rows(\"12:12\").Hidden = True\n  Rows(\"12:12\").Hidden = False\n  Cells(6, 22).Value = \"CellsValue\"\n  Range(Cells(6, 23), Cells(6, 24)).Value = Array(\"Array1\", \"Array2\")\n  Range(\"Y6\").FormulaR1C1 = \"=RC[-15]+1\"\n  Range(\"Z6\").Value = WorksheetFunction.Sum(Range(\"J6:J6\"), 8)\n  Range(\"AB6\").Value = Application.WorksheetFunction.Average(Range(\"J6:J6\"), 46)\n  Range(\"AA6\").Value = \"Indexed\"\n  Range(\"AA6\").Font.ColorIndex = 2\n  Range(\"AA6\").Interior.ColorIndex = 5\n  Range(\"AA6\").HorizontalAlignment = xlHAlignCenter\n  Range(\"AA6\").VerticalAlignment = xlVAlignBottom\n  Range(\"AA6\").Font.Underline = xlUnderlineStyleDouble\n  Range(\"I6:J6\").Select\n  ActiveCell.Offset(3, 0).Value = \"ActiveOffset\"\n  Selection.Offset(4, 0).Resize(1, 1).Value = \"SelectionOffset\"\n  MsgBox \"View\"\nEnd Sub");


  await page.locator("[data-data-dialog-action='macro-editor-ok']").click();


  await page.waitForSelector('[data-data-dialog="macros"]', { timeout: 5000 });


  await page.locator("[data-data-field='macroIndex']").focus();


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => {
    const macro = (window.__webSheetDevWorkbookState?.().macros || []).find((item) => item.name === "ViewMacro");
    return macro && !macro.code;
  }, null, { timeout: 5000 });


  const macroEditorUndoRedo = {
    undo: await page.evaluate(() => ({
      macro: (window.__webSheetDevWorkbookState?.().macros || []).find((item) => item.name === "ViewMacro") || null,
      dialogVisible: Boolean(document.querySelector("[data-data-dialog='macros']")),
    })),
  };


  await page.locator("[data-data-field='macroIndex']").focus();


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => {
    const macro = (window.__webSheetDevWorkbookState?.().macros || []).find((item) => item.name === "ViewMacro");
    return /Sub ViewMacro/.test(macro?.code || "");
  }, null, { timeout: 5000 });


  macroEditorUndoRedo.redo = await page.evaluate(() => ({
    macro: (window.__webSheetDevWorkbookState?.().macros || []).find((item) => item.name === "ViewMacro") || null,
    dialogVisible: Boolean(document.querySelector("[data-data-dialog='macros']")),
  }));


	  const macroRunBefore = await page.evaluate(() => {
	    const sheet = window.__webSheetDevSheetModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "CHT");
	    return {
	      tColumn: sheet?.columns?.[19] || null,
	      row11: sheet?.rows?.[10] || null,
	      row12: sheet?.rows?.[11] || null,
	    };
	  });


	  await page.locator("[data-data-dialog-action='macro-run']").click();


	  await page.waitForSelector("#messageBox", { timeout: 5000 });


	  const macroRunMessage = await page.evaluate(() => {
	    const sheet = window.__webSheetDevSheetModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "CHT");
	    const cellText = (address, row, col) => {
	      const domText = document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent;
	      if (domText != null && domText !== "") return domText;
	      const cell = sheet?.cells?.[`${row}:${col}`];
	      const value = cell?.raw ?? cell?.cached ?? cell?.display ?? "";
	      return value == null ? "" : String(value);
	    };
	    return {
	      title: document.querySelector("#messageBoxTitle")?.textContent || "",
	      text: document.querySelector("#messageBoxText")?.textContent || "",
	      i6: cellText("I6", 6, 9),
	      j6: cellText("J6", 6, 10),
	      k6: cellText("K6", 6, 11),
	      i7: cellText("I7", 7, 9),
	      i8: cellText("I8", 8, 9),
	      j8: cellText("J8", 8, 10),
	      i9: cellText("I9", 9, 9),
	      i10: cellText("I10", 10, 9),
	      v6: cellText("V6", 6, 22),
	      w6: cellText("W6", 6, 23),
	      x6: cellText("X6", 6, 24),
	      z6: cellText("Z6", 6, 26),
	      ab6: cellText("AB6", 6, 28),
	      i6Model: sheet?.cells?.["6:9"] || null,
	      j6Model: sheet?.cells?.["6:10"] || null,
	      k6Model: sheet?.cells?.["6:11"] || null,
	      l6Model: sheet?.cells?.["6:12"] || null,
	      m6Model: sheet?.cells?.["6:13"] || null,
	      n6Model: sheet?.cells?.["6:14"] || null,
	      o6Model: sheet?.cells?.["6:15"] || null,
	      p6Model: sheet?.cells?.["6:16"] || null,
	      r6Model: sheet?.cells?.["6:18"] || null,
	      s6Model: sheet?.cells?.["6:19"] || null,
	      v6Model: sheet?.cells?.["6:22"] || null,
	      w6Model: sheet?.cells?.["6:23"] || null,
	      x6Model: sheet?.cells?.["6:24"] || null,
	      y6Model: sheet?.cells?.["6:25"] || null,
	      z6Model: sheet?.cells?.["6:26"] || null,
	      aa6Model: sheet?.cells?.["6:27"] || null,
	      ab6Model: sheet?.cells?.["6:28"] || null,
	      tColumn: sheet?.columns?.[19] || null,
	      uColumn: sheet?.columns?.[20] || null,
	      row11: sheet?.rows?.[10] || null,
	      row12: sheet?.rows?.[11] || null,
	      pqMerged: Boolean(sheet?.merges?.some((merge) => merge.top === 6 && merge.left === 16 && merge.bottom === 6 && merge.right === 17)),
	      selectionLabel: document.querySelector("#selectedCellLabel")?.textContent || "",
	      status: document.querySelector("#statusText")?.textContent || "",
	      dialogClosed: !document.querySelector("[data-data-dialog='macros']"),
	    };
	  });


	  await page.locator("[data-message-box-close]").click();


	  await page.waitForSelector("#messageBox", { state: "detached", timeout: 5000 });


	  await page.keyboard.press("Control+Z");


	  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="I6"] .cell-content')?.textContent !== "MacroRun", null, { timeout: 5000 });


	  const macroRunUndoRedo = {
	    undo: await page.evaluate(() => {
	      const sheet = window.__webSheetDevSheetModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "CHT");
	      const cellText = (address, row, col) => {
	        const domText = document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent;
	        if (domText != null && domText !== "") return domText;
	        const cell = sheet?.cells?.[`${row}:${col}`];
	        const value = cell?.raw ?? cell?.cached ?? cell?.display ?? "";
	        return value == null ? "" : String(value);
	      };
	      return {
	        i6: cellText("I6", 6, 9),
	        j6: cellText("J6", 6, 10),
	        v6: cellText("V6", 6, 22),
	        i6Model: sheet?.cells?.["6:9"] || null,
	        r6Model: sheet?.cells?.["6:18"] || null,
	        tColumn: sheet?.columns?.[19] || null,
	        row11: sheet?.rows?.[10] || null,
	        m6Model: sheet?.cells?.["6:13"] || null,
	        o6Model: sheet?.cells?.["6:15"] || null,
	        p6Model: sheet?.cells?.["6:16"] || null,
	        v6Model: sheet?.cells?.["6:22"] || null,
	        y6Model: sheet?.cells?.["6:25"] || null,
	        aa6Model: sheet?.cells?.["6:27"] || null,
	        lastRunAt: (window.__webSheetDevWorkbookState?.().macros || []).find((macro) => macro.name === "ViewMacro")?.lastRunAt || "",
	        status: document.querySelector("#statusText")?.textContent || "",
	      };
	    }),
	  };


	  await page.keyboard.press("Control+Y");


	  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="I6"] .cell-content')?.textContent === "MacroRun", null, { timeout: 5000 });


	  macroRunUndoRedo.redo = await page.evaluate(() => {
	    const sheet = window.__webSheetDevSheetModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "CHT");
	    const cellText = (address, row, col) => {
	      const domText = document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent;
	      if (domText != null && domText !== "") return domText;
	      const cell = sheet?.cells?.[`${row}:${col}`];
	      const value = cell?.raw ?? cell?.cached ?? cell?.display ?? "";
	      return value == null ? "" : String(value);
	    };
	    return {
	      i6: cellText("I6", 6, 9),
	      j6: cellText("J6", 6, 10),
	      k6: cellText("K6", 6, 11),
	      i7: cellText("I7", 7, 9),
	      i8: cellText("I8", 8, 9),
	      j8: cellText("J8", 8, 10),
	      i9: cellText("I9", 9, 9),
	      i10: cellText("I10", 10, 9),
	      v6: cellText("V6", 6, 22),
	      w6: cellText("W6", 6, 23),
	      x6: cellText("X6", 6, 24),
	      z6: cellText("Z6", 6, 26),
	      ab6: cellText("AB6", 6, 28),
	      i6Model: sheet?.cells?.["6:9"] || null,
	      j6Model: sheet?.cells?.["6:10"] || null,
	      k6Model: sheet?.cells?.["6:11"] || null,
	      l6Model: sheet?.cells?.["6:12"] || null,
	      m6Model: sheet?.cells?.["6:13"] || null,
	      n6Model: sheet?.cells?.["6:14"] || null,
	      o6Model: sheet?.cells?.["6:15"] || null,
	      p6Model: sheet?.cells?.["6:16"] || null,
	      r6Model: sheet?.cells?.["6:18"] || null,
	      s6Model: sheet?.cells?.["6:19"] || null,
	      v6Model: sheet?.cells?.["6:22"] || null,
	      w6Model: sheet?.cells?.["6:23"] || null,
	      x6Model: sheet?.cells?.["6:24"] || null,
	      y6Model: sheet?.cells?.["6:25"] || null,
	      z6Model: sheet?.cells?.["6:26"] || null,
	      aa6Model: sheet?.cells?.["6:27"] || null,
	      ab6Model: sheet?.cells?.["6:28"] || null,
	      tColumn: sheet?.columns?.[19] || null,
	      uColumn: sheet?.columns?.[20] || null,
	      row11: sheet?.rows?.[10] || null,
	      row12: sheet?.rows?.[11] || null,
	      pqMerged: Boolean(sheet?.merges?.some((merge) => merge.top === 6 && merge.left === 16 && merge.bottom === 6 && merge.right === 17)),
	      lastRunAt: (window.__webSheetDevWorkbookState?.().macros || []).find((macro) => macro.name === "ViewMacro")?.lastRunAt || "",
	      selectionLabel: document.querySelector("#selectedCellLabel")?.textContent || "",
	      status: document.querySelector("#statusText")?.textContent || "",
	    };
	  });


	  await page.click('[data-ribbon-command="view-macros"]');


	  await page.waitForSelector('[data-data-dialog="macros"]', { timeout: 5000 });


	  const macrosDialog = await page.evaluate(() => ({
	    title: document.querySelector("#macrosTitle")?.textContent || "",
	    initial: window.__webSheetMacroInitialForTest || null,
	    created: (window.__webSheetDevWorkbookState?.().macros || []).find((macro) => macro.name === "ViewMacro") || null,
	  }));


  macrosDialog.initial = macrosInitialDialog;


  macrosDialog.afterCreate = macrosAfterCreateDialog;


  macrosDialog.createUndoRedo = macroCreateUndoRedo;


  macrosDialog.options = macroOptionsDialog;


  macrosDialog.optionsUndoRedo = macroOptionsUndoRedo;


  macrosDialog.editor = macroEditorDialog;


  macrosDialog.editorUndoRedo = macroEditorUndoRedo;


  macrosDialog.runBefore = macroRunBefore;


  macrosDialog.runMessage = macroRunMessage;


  macrosDialog.runUndoRedo = macroRunUndoRedo;


  await page.locator("[data-data-dialog-action='macro-step']").click();


  await page.waitForSelector('[data-data-dialog="macro-step"]', { timeout: 5000 });


  macrosDialog.stepDialog = await page.evaluate(() => ({
    title: document.querySelector("#macroStepTitle")?.textContent || "",
    summary: document.querySelector("[data-macro-step-summary]")?.textContent || "",
    current: document.querySelector(".macro-step-item.is-current code")?.textContent || "",
    count: document.querySelectorAll(".macro-step-item").length,
    nextDisabled: document.querySelector("[data-data-dialog-action='macro-step-next']")?.disabled || false,
    continueDisabled: document.querySelector("[data-data-dialog-action='macro-step-continue']")?.disabled || false,
  }));


  await page.keyboard.press("F8");


  await page.waitForFunction(() => document.querySelector("[data-macro-step-summary]")?.textContent?.includes("2 /"), null, { timeout: 5000 });


  macrosDialog.stepDialog.afterFirst = await page.evaluate(() => ({
    summary: document.querySelector("[data-macro-step-summary]")?.textContent || "",
    log: document.querySelector("[data-macro-step-log]")?.textContent || "",
    completeCount: document.querySelectorAll(".macro-step-item.is-complete").length,
    current: document.querySelector(".macro-step-item.is-current code")?.textContent || "",
  }));


  await page.keyboard.press("F5");


  await page.waitForFunction(() => document.querySelector("[data-data-dialog-action='macro-step-next']")?.disabled, null, { timeout: 5000 });


  macrosDialog.stepDialog.afterContinue = await page.evaluate(() => ({
    summary: document.querySelector("[data-macro-step-summary]")?.textContent || "",
    log: document.querySelector("[data-macro-step-log]")?.textContent || "",
    nextDisabled: document.querySelector("[data-data-dialog-action='macro-step-next']")?.disabled || false,
    continueDisabled: document.querySelector("[data-data-dialog-action='macro-step-continue']")?.disabled || false,
    completeCount: document.querySelectorAll(".macro-step-item.is-complete").length,
  }));


  await page.locator("[data-data-dialog-action='cancel']").click();


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  await page.click('[data-ribbon-command="view-macros-menu"]');


  await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });


  const macroMenuBeforeRecord = await page.evaluate(() => ({
    labels: [...document.querySelectorAll("#viewCommandMenu [data-view-menu-action]")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
    descriptions: [...document.querySelectorAll("#viewCommandMenu .home-command-menu-text small")].map((node) => node.textContent.trim()),
    checkedAction: document.querySelector("#viewCommandMenu [aria-checked='true']")?.dataset.viewMenuAction || "",
  }));


  await page.locator('#viewCommandMenu [data-view-menu-action="macro-record"]').click();


  await page.waitForSelector('[data-data-dialog="macro-record"]', { timeout: 5000 });


  const macroRecordDialog = await page.evaluate(() => ({
    title: document.querySelector("#macroRecordTitle")?.textContent || "",
    defaultName: document.querySelector("[data-data-field='macroRecordName']")?.value || "",
    scopeOptions: [...document.querySelectorAll("[data-data-field='macroRecordScope'] option")].map((option) => option.textContent.trim()),
  }));


  await page.locator("[data-data-field='macroRecordName']").fill("RecordedViewMacro");


  await page.locator("[data-data-field='macroRecordShortcut']").fill("r");


  await page.locator("[data-data-field='macroRecordDescription']").fill("表示タブから記録したマクロ");


  await page.locator("[data-data-dialog-action='macro-record-ok']").click();


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  await page.waitForFunction(() => Boolean(window.__webSheetDevWorkbookState?.().macroRecording), null, { timeout: 5000 });


  const macroRecordingStarted = await page.evaluate(() => ({
    recording: window.__webSheetDevWorkbookState?.().macroRecording || null,
    macro: (window.__webSheetDevWorkbookState?.().macros || []).find((macro) => macro.name === "RecordedViewMacro") || null,
    macroButtonActive: document.querySelector('[data-ribbon-command="view-macros"]')?.classList.contains("is-active") || false,
    menuButtonActive: document.querySelector('[data-ribbon-command="view-macros-menu"]')?.classList.contains("is-active") || false,
  }));


  await page.click('[data-ribbon-command="view-macros-menu"]');


  await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });


  const macroMenuDuringRecord = await page.evaluate(() => ({
    labels: [...document.querySelectorAll("#viewCommandMenu [data-view-menu-action]")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
    relativeChecked: document.querySelector('#viewCommandMenu [data-view-menu-action="macro-use-relative-refs"]')?.getAttribute("aria-checked") || "",
  }));


  await page.locator('#viewCommandMenu [data-view-menu-action="macro-use-relative-refs"]').click();


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().macroRelativeReferences === true, null, { timeout: 5000 });


  await page.click('[data-ribbon-command="view-macros-menu"]');


  await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });


  const macroMenuRelativeRefs = await page.evaluate(() => ({
    relativeChecked: document.querySelector('#viewCommandMenu [data-view-menu-action="macro-use-relative-refs"]')?.getAttribute("aria-checked") || "",
  }));


  await page.evaluate(() => window.__webSheetDevCloseFloatingMenus?.());


  await page.waitForSelector("#viewCommandMenu", { state: "detached", timeout: 5000 });


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().macroRelativeReferences === false, null, { timeout: 5000 });


  const macroRelativeRefsUndoRedo = {
    undo: await page.evaluate(() => ({
      relativeReferences: window.__webSheetDevWorkbookState?.().macroRelativeReferences,
      recordingRelativeReferences: window.__webSheetDevWorkbookState?.().macroRecording?.relativeReferences,
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().macroRelativeReferences === true, null, { timeout: 5000 });


  macroRelativeRefsUndoRedo.redo = await page.evaluate(() => ({
    relativeReferences: window.__webSheetDevWorkbookState?.().macroRelativeReferences,
    recordingRelativeReferences: window.__webSheetDevWorkbookState?.().macroRecording?.relativeReferences,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('[data-ribbon-command="view-macros-menu"]');


  await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });


  await page.locator('#viewCommandMenu [data-view-menu-action="macro-stop-recording"]').click();


  await page.waitForFunction(() => !window.__webSheetDevWorkbookState?.().macroRecording, null, { timeout: 5000 });


  const macroRecordingStopped = await page.evaluate(() => ({
    recording: window.__webSheetDevWorkbookState?.().macroRecording || null,
    relativeReferences: window.__webSheetDevWorkbookState?.().macroRelativeReferences,
    status: document.querySelector("#statusText")?.textContent || "",
    macroCount: (window.__webSheetDevWorkbookState?.().macros || []).length,
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => Boolean(window.__webSheetDevWorkbookState?.().macroRecording), null, { timeout: 5000 });


  const macroRecordingStopUndoRedo = {
    undo: await page.evaluate(() => ({
      recording: window.__webSheetDevWorkbookState?.().macroRecording || null,
      macroButtonActive: document.querySelector('[data-ribbon-command="view-macros"]')?.classList.contains("is-active") || false,
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => !window.__webSheetDevWorkbookState?.().macroRecording, null, { timeout: 5000 });


  macroRecordingStopUndoRedo.redo = await page.evaluate(() => ({
    recording: window.__webSheetDevWorkbookState?.().macroRecording || null,
    macroButtonActive: document.querySelector('[data-ribbon-command="view-macros"]')?.classList.contains("is-active") || false,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('[data-ribbon-command="view-macros-menu"]');


  await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });


  await page.locator('#viewCommandMenu [data-view-menu-action="macro-security"]').click();


  await page.waitForSelector('[data-data-dialog="macro-security"]', { timeout: 5000 });


  const macroSecurityDialog = await page.evaluate(() => ({
    title: document.querySelector("#macroSecurityTitle")?.textContent || "",
    navText: document.querySelector(".macro-security-nav .is-active")?.textContent.trim() || "",
    navItems: [...document.querySelectorAll(".macro-security-nav button")].map((button) => button.textContent.trim()),
    options: [...document.querySelectorAll("[data-data-field='macroSecurityLevel']")].map((input) => ({
      value: input.value,
      checked: input.checked,
      label: input.closest(".macro-security-option")?.querySelector("strong")?.textContent.trim() || "",
      description: input.closest(".macro-security-option")?.querySelector("small")?.textContent.trim() || "",
    })),
    trustChecked: document.querySelector("[data-data-field='macroTrustVbaObjectModel']")?.checked || false,
    warning: document.querySelector("[data-data-dialog='macro-security'] .data-dialog-summary")?.textContent.trim() || "",
  }));


  await page.locator("[data-data-field='macroSecurityLevel'][value='enableAll']").check();


  await page.locator("[data-data-field='macroTrustVbaObjectModel']").check();


  await page.locator('[data-data-dialog="macro-security"]').press("Enter");


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().macroSecurity?.level === "enableAll", null, { timeout: 5000 });


  const macroSecurityApplied = await page.evaluate(() => ({
    security: window.__webSheetDevWorkbookState?.().macroSecurity || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().macroSecurity?.level !== "enableAll", null, { timeout: 5000 });


  const macroSecurityUndoRedo = {
    undo: await page.evaluate(() => ({
      security: window.__webSheetDevWorkbookState?.().macroSecurity || null,
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().macroSecurity?.level === "enableAll", null, { timeout: 5000 });


  macroSecurityUndoRedo.redo = await page.evaluate(() => ({
    security: window.__webSheetDevWorkbookState?.().macroSecurity || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  const viewWindowCommands = {
    ...(await page.evaluate(() => ({
    freezePanes: window.__webSheetDevSheetModel?.("CHT")?.freezePanes || null,
    splitPane: window.__webSheetDevSheetModel?.("CHT")?.splitPane || null,
    windowViews: window.__webSheetDevWorkbookState?.().windowViews || [],
    customViews: window.__webSheetDevWorkbookState?.().customViews || [],
    viewSideBySide: window.__webSheetDevWorkbookState?.().viewSideBySide || null,
    synchronousScrolling: window.__webSheetDevWorkbookState?.().synchronousScrolling,
    hiddenWindows: (window.__webSheetDevWorkbookState?.().windowViews || []).filter((view) => view.hidden).length,
    status: document.querySelector("#statusText")?.textContent || "",
    }))),
    switchWindowMenu,
    switchWindowKeyboardState,
    viewInitialWindowBar,
    viewNewWindowState,
    viewNewWindowBar,
	    viewSecondNewWindowState,
	    viewSecondNewWindowBar,
	    workbookWindowBarSwitch,
	    customWindowSource,
	    customWindowMutated,
	    customWindowRestored,
	    arrangeWindowState,
    arrangeWindowBar,
    arrangeWindowUndo,
    arrangeWindowRedo,
    sideBySideWindowDialog,
    sideBySideDialogApplied,
    sideBySideWindowBar,
    syncScrollState,
    syncScrollUndoRedo,
    resetWindowState,
    resetWindowBar,
    switchWindowState,
    switchWindowDialog,
    switchWindowDialogState,
    hideWindowState,
    hideWindowBar,
	    unhideWindowDialog,
	    unhideWindowState,
	    unhideWindowBar,
	    unhideWindowUndoRedo,
	    closeWindowState,
	    closeWindowUndoRedo,
	    macrosDialog,
    macroMenuBeforeRecord,
    macroRecordDialog,
    macroRecordingStarted,
    macroMenuDuringRecord,
    macroMenuRelativeRefs,
    macroRelativeRefsUndoRedo,
    macroRecordingStopped,
    macroRecordingStopUndoRedo,
    macroSecurityDialog,
    macroSecurityApplied,
    macroSecurityUndoRedo,
    arrangeWindowDialog,
    sheetViewNew,
    sheetViewKept,
    sheetViewMenuBeforeRename,
    sheetViewRenamed,
    sheetViewUpdated,
    sheetViewMenuBeforeCopy,
    sheetViewDuplicated,
    sheetViewExited,
    sheetViewReactivated,
    freezeMenuLabels,
    freezeMenuInitialMeta,
    freezeMenuTopRow,
    freezeMenuTopRowStatus,
    freezeMenuAccessKey,
    viewFullScreenOn,
    fullScreenUndoRedo,
    viewFullScreenOff,
    freezePanesAfterMain,
    frozenPaneSticky,
    splitPaneBeforeDrag,
    splitPaneDragged,
    sheetGroupView,
    viewZoomDialog,
    viewZoomApplied,
    viewZoomHundred,
    viewZoomSelection,
    viewZoomSelectionUndoRedo,
    viewRulerOn,
    pageLayoutHeaderFooterDialog,
	    pageBreakPreview,
	    pageBreakContextMenu,
	    gridlinesUndoRedo,
	    headingsUndoRedo,
	    zeroValuesUndoRedo,
	    objectsUndoRedo,
	    outlineUndoRedo,
	    rightToLeftUndoRedo,
	    rtlVisualHitTest,
	    formulaBarUndoRedo,
	    customViewCaptureState,
		    customViewsMainBeforeAdd,
		    customViewAddDialog,
		    customViewsDialog,
		    customViewSaveUndoRedo,
		    customViewApplied,
		    customViewApplyUndoRedo,
		    customViewDeleteFlow,
		  };


  if (viewDiagnosticsOnly) {
    console.log(JSON.stringify({
      viewRibbon,
      viewToggles,
      viewWindowCommands,
    }, null, 2));
    await context.close();
    await browser.close();
    process.exit(0);
  }


  await page.click('[data-ribbon-command="view-split"]');


  await page.click('[data-ribbon-command="view-freeze-panes"]');



  await page.click('[data-ribbon-tab="layout"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "ページレイアウト", null, { timeout: 10000 });


  const layoutRibbon = await page.evaluate(() => {
    const panel = document.querySelector('[data-ribbon-panel="layout"]');
    return {
      activeRibbonTab: document.querySelector(".ribbon-tab.is-active")?.textContent,
      panelVisible: Boolean(panel?.classList.contains("is-active")),
      groupLabels: [...(panel?.querySelectorAll(".ribbon-group-label") || [])].map((label) => label.textContent.trim()),
      enabledCommands: [...(panel?.querySelectorAll("[data-ribbon-command]:not(:disabled)") || [])].map((button) => button.dataset.ribbonCommand),
      panelHeight: Math.round(document.querySelector(".ribbon-panels")?.getBoundingClientRect().height || 0),
    };
  });


  const collectLayoutMenu = async (command, selector = "#homeCommandMenu") => {
    await page.click(`[data-ribbon-command="${command}"]`);
    await page.waitForSelector(selector, { timeout: 5000 });
    return await page.evaluate(() => ({
      label: document.querySelector("#homeCommandMenu")?.getAttribute("aria-label") || "",
      kind: document.querySelector("#homeCommandMenu")?.dataset.homeCommandKind || "",
      items: [...document.querySelectorAll("#homeCommandMenu button")].map((button) => ({
        action: button.dataset.layoutMenuAction || "",
        text: button.textContent.replace(/\s+/g, " ").trim(),
        checked: button.getAttribute("aria-checked") === "true",
        disabled: button.disabled,
        role: button.getAttribute("role") || "",
        hasDescription: button.classList.contains("has-description"),
        accessKey: button.dataset.accessKey || "",
      })),
      descriptionCount: document.querySelectorAll("#homeCommandMenu button.has-description small").length,
    }));
  };


  const layoutThemeMenuInitial = await collectLayoutMenu("layout-themes", "#homeCommandMenu.is-layout-themes");


  await page.click('[data-layout-menu-action="theme:parcel"]');


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.theme === "parcel", null, { timeout: 5000 });


  const layoutThemeApplied = await page.evaluate(() => ({
    model: window.__webSheetDevSheetModel?.("CHT")?.pageLayout || null,
    themeColors: window.__webSheetDevWorkbookState?.().themeColors || [],
    themeFonts: window.__webSheetDevWorkbookState?.().themeFonts || {},
    themeEffects: window.__webSheetDevWorkbookState?.().themeEffects || {},
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  const layoutColorMenuInitial = await collectLayoutMenu("layout-colors", "#homeCommandMenu.is-layout-colors");


  await page.click('[data-layout-menu-action="colors:facet"]');


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.colorScheme === "facet", null, { timeout: 5000 });


  const layoutFontMenuInitial = await collectLayoutMenu("layout-fonts", "#homeCommandMenu.is-layout-fonts");


  await page.click('[data-layout-menu-action="fonts:wisp"]');


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.fontScheme === "wisp", null, { timeout: 5000 });


  const layoutEffectMenuInitial = await collectLayoutMenu("layout-effects", "#homeCommandMenu.is-layout-effects");


  await page.click('[data-layout-menu-action="effects:retrospect"]');


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.effectScheme === "retrospect", null, { timeout: 5000 });


  const layoutThemeAfterMenus = await page.evaluate(() => ({
    model: window.__webSheetDevSheetModel?.("CHT")?.pageLayout || null,
    themeColors: window.__webSheetDevWorkbookState?.().themeColors || [],
    themeFonts: window.__webSheetDevWorkbookState?.().themeFonts || {},
    themeEffects: window.__webSheetDevWorkbookState?.().themeEffects || {},
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.effectScheme === "parcel", null, { timeout: 5000 });


  const layoutThemeEffectUndoRedo = {
    undo: await page.evaluate(() => ({
      model: window.__webSheetDevSheetModel?.("CHT")?.pageLayout || null,
      themeEffects: window.__webSheetDevWorkbookState?.().themeEffects || {},
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.effectScheme === "retrospect", null, { timeout: 5000 });


  layoutThemeEffectUndoRedo.redo = await page.evaluate(() => ({
    model: window.__webSheetDevSheetModel?.("CHT")?.pageLayout || null,
    themeEffects: window.__webSheetDevWorkbookState?.().themeEffects || {},
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(2, 2, 4, 3));


  const layoutPrintAreaMenuBefore = await collectLayoutMenu("layout-print-area", "#homeCommandMenu.is-layout-print-area");


  await page.click('[data-layout-menu-action="print-area:set"]');


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.printArea?.label === "B1:AP4", null, { timeout: 5000 });


  const layoutPrintAreaMenuAfterSet = await collectLayoutMenu("layout-print-area", "#homeCommandMenu.is-layout-print-area");


  await page.click('[data-layout-menu-action="print-area:clear"]');


  await page.waitForFunction(() => !window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.printArea, null, { timeout: 5000 });


  const layoutPrintAreaMenuAfterClear = await collectLayoutMenu("layout-print-area", "#homeCommandMenu.is-layout-print-area");


  await page.click('[data-layout-menu-action="print-area:set"]');


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.printArea?.label === "B1:AP4", null, { timeout: 5000 });


  const layoutMarginsMenuInitial = await collectLayoutMenu("layout-margins", "#homeCommandMenu.is-layout-margins");


  await page.click('[data-ribbon-command="layout-margins"]');


  await page.click('[data-layout-menu-action="margins:narrow"]');


  const layoutOrientationMenuInitial = await collectLayoutMenu("layout-orientation", "#homeCommandMenu.is-layout-orientation");


  await page.click('[data-ribbon-command="layout-orientation"]');


  await page.click('[data-layout-menu-action="orientation:landscape"]');


  const layoutSizeMenuInitial = await collectLayoutMenu("layout-size", "#homeCommandMenu.is-layout-size");


  await page.click('[data-ribbon-command="layout-size"]');


  await page.click('[data-layout-menu-action="paper:Letter"]');


  await page.locator('[data-layout-sheet-option="gridlines-print"]').check();


  await page.evaluate(() => {
    const scale = document.querySelector('[data-layout-field="scale"]');
    if (scale) {
      scale.value = "85";
      scale.dispatchEvent(new Event("change", { bubbles: true }));
    }
    const width = document.querySelector('[data-layout-field="fitToWidth"]');
    if (width) {
      width.value = "1";
      width.dispatchEvent(new Event("change", { bubbles: true }));
    }
  });


  await page.click('[data-ribbon-command="layout-background"]');


  await page.locator("#pageBackgroundInput").setInputFiles(pageBackgroundImagePath);


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.background?.src?.startsWith("data:image/png"), null, { timeout: 5000 });


  const layoutBackgroundApplied = await page.evaluate(() => ({
    background: window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.background || null,
    buttonText: document.querySelector('[data-ribbon-command="layout-background"] span')?.textContent?.trim() || "",
    buttonPressed: document.querySelector('[data-ribbon-command="layout-background"]')?.getAttribute("aria-pressed") || "",
    layerName: document.querySelector(".sheet-background-layer")?.dataset.backgroundName || "",
    layerImage: document.querySelector(".sheet-background-layer")?.style.backgroundImage || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => !window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.background, null, { timeout: 5000 });


  const layoutBackgroundUndo = await page.evaluate(() => ({
    background: window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.background || null,
    buttonText: document.querySelector('[data-ribbon-command="layout-background"] span')?.textContent?.trim() || "",
    layerCount: document.querySelectorAll(".sheet-background-layer").length,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.background?.src?.startsWith("data:image/png"), null, { timeout: 5000 });


  const layoutBackgroundRedo = await page.evaluate(() => ({
    background: window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.background || null,
    buttonText: document.querySelector('[data-ribbon-command="layout-background"] span')?.textContent?.trim() || "",
    layerCount: document.querySelectorAll(".sheet-background-layer").length,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('[data-ribbon-command="layout-background"]');


  await page.waitForFunction(() => !window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.background, null, { timeout: 5000 });


  const layoutBackgroundDeleted = await page.evaluate(() => ({
    background: window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.background || null,
    buttonText: document.querySelector('[data-ribbon-command="layout-background"] span')?.textContent?.trim() || "",
    layerCount: document.querySelectorAll(".sheet-background-layer").length,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.background?.src?.startsWith("data:image/png"), null, { timeout: 5000 });


  const layoutBackgroundDeleteUndo = await page.evaluate(() => ({
    background: window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.background || null,
    buttonText: document.querySelector('[data-ribbon-command="layout-background"] span')?.textContent?.trim() || "",
    layerCount: document.querySelectorAll(".sheet-background-layer").length,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('[data-ribbon-command="layout-page-setup"]');


  await page.waitForSelector('[data-data-dialog="page-setup"]', { timeout: 5000 });


  const pageSetupDialog = await page.evaluate(() => ({
    title: document.querySelector("#pageSetupTitle")?.textContent || "",
    fields: [...document.querySelectorAll("[data-page-setup-field]")].map((input) => input.dataset.pageSetupField),
    checks: [...document.querySelectorAll("[data-page-setup-check]")].map((input) => input.dataset.pageSetupCheck),
    rangePickers: [...document.querySelectorAll('[data-data-dialog="page-setup"] [data-data-range-target]')].map((button) => ({
      target: button.dataset.dataRangeTarget || "",
      mode: button.dataset.dataRangeMode || "",
    })),
    tabs: [...document.querySelectorAll("[data-page-setup-tab]")].map((tab) => tab.textContent.trim()),
    activePanel: document.querySelector(".page-setup-panel.is-active")?.dataset.pageSetupPanel || "",
  }));


  await page.evaluate(() => {
    const setValue = (selector, value) => {
      const input = document.querySelector(selector);
      if (!input) return;
      input.value = value;
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new Event("change", { bubbles: true }));
    };
    const setChecked = (name, checked) => {
      const input = document.querySelector(`[data-page-setup-check="${name}"]`);
      if (!input) return;
      input.checked = checked;
      input.dispatchEvent(new Event("change", { bubbles: true }));
    };
    setValue('[data-page-setup-field="scale"]', "95");
    setValue('[data-page-setup-field="firstPageNumber"]', "3");
    setValue('[data-page-setup-field="pageOrder"]', "overThenDown");
    setValue('[data-page-setup-field="comments"]', "atEnd");
    setValue('[data-page-setup-field="cellErrors"]', "blank");
    setValue('[data-page-setup-field="headerText"]', "Layout Header");
    setValue('[data-page-setup-field="footerText"]', "Layout Footer");
    setChecked("centerHorizontal", true);
    setChecked("centerVertical", true);
    setChecked("blackAndWhite", true);
    setChecked("draftQuality", true);
  });


  await page.locator('[data-data-dialog="page-setup"] [data-data-dialog-action="ok"]').click();


  await page.waitForSelector('[data-data-dialog="page-setup"]', { state: "detached", timeout: 5000 });


  await page.click('[data-ribbon-command="layout-print-titles"]');


  await page.waitForSelector('[data-data-dialog="page-setup"] .page-setup-panel.is-active[data-page-setup-panel="sheet"]', { timeout: 5000 });


  const printTitlesDialogInitial = await page.evaluate(() => ({
    activePanel: document.querySelector(".page-setup-panel.is-active")?.dataset.pageSetupPanel || "",
    focusedField: document.activeElement?.dataset?.pageSetupField || "",
    printArea: document.querySelector('[data-page-setup-field="printArea"]')?.value || "",
    titleRows: document.querySelector('[data-page-setup-field="printTitleRows"]')?.value || "",
    titleColumns: document.querySelector('[data-page-setup-field="printTitleColumns"]')?.value || "",
    rangePickers: [...document.querySelectorAll('[data-data-dialog="page-setup"] [data-data-range-target]')].map((button) => ({
      target: button.dataset.dataRangeTarget || "",
      mode: button.dataset.dataRangeMode || "",
    })),
  }));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(6, 1, 6, 3));


  await page.click('[data-data-dialog="page-setup"] [data-data-range-target="printTitleRows"]');


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 4, 3, 5));


  await page.click('[data-data-dialog="page-setup"] [data-data-range-target="printTitleColumns"]');


  const printTitlesDialogFilled = await page.evaluate(() => ({
    titleRows: document.querySelector('[data-page-setup-field="printTitleRows"]')?.value || "",
    titleColumns: document.querySelector('[data-page-setup-field="printTitleColumns"]')?.value || "",
  }));


  await page.locator('[data-data-dialog="page-setup"] [data-data-dialog-action="ok"]').click();


  await page.waitForSelector('[data-data-dialog="page-setup"]', { state: "detached", timeout: 5000 });


  const pageLayout = await page.evaluate(() => ({
    model: window.__webSheetDevSheetModel?.("CHT")?.pageLayout || null,
    headerFooter: window.__webSheetDevSheetModel?.("CHT")?.headerFooter || null,
    gridlinesPrintChecked: document.querySelector('[data-layout-sheet-option="gridlines-print"]')?.checked,
    scaleControl: document.querySelector('[data-layout-field="scale"]')?.value || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.evaluate(() => document.querySelector('[data-sheet-view="page-layout"]')?.click());


  await page.waitForFunction(() => document.querySelector(".sheet-grid.view-page-layout .sheet-page-layout-page"), null, { timeout: 5000 });


  const pageLayoutPreview = await page.evaluate(() => ({
    viewMode: window.__webSheetDevSheetModel?.("CHT")?.viewMode,
    pageCount: document.querySelectorAll(".sheet-page-layout-page").length,
    marginCount: document.querySelectorAll(".sheet-page-layout-margin").length,
    headerFooterCount: document.querySelectorAll(".sheet-page-layout-header, .sheet-page-layout-footer").length,
    printAreaCount: document.querySelectorAll(".sheet-page-layout-print-area").length,
    titleRowCount: document.querySelectorAll(".sheet-page-layout-title-row").length,
    titleColumnCount: document.querySelectorAll(".sheet-page-layout-title-col").length,
  }));


  await page.evaluate(() => document.querySelector('[data-sheet-view="normal"]')?.click());


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(8, 8, "BreakMenuSeed");
    window.__webSheetDevSelectRange?.(8, 8, 8, 8);
  });


  await page.click('[data-ribbon-command="layout-breaks"]');


  await page.waitForSelector("#homeCommandMenu.is-layout-breaks", { timeout: 5000 });


  const pageBreakMenuKeyboard = await page.evaluate(() => ({
    label: document.querySelector("#homeCommandMenu")?.getAttribute("aria-label") || "",
    accessKeys: [...document.querySelectorAll("#homeCommandMenu button")].map((button) => button.dataset.accessKey || ""),
    h8BeforeShortcut: window.__webSheetDevCellModel?.("CHT", 8, 8)?.raw || "",
  }));


  await page.keyboard.press("I");


  await page.waitForFunction(() => {
    const breaks = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks;
    return breaks?.rows?.includes(8) && breaks?.cols?.includes(8);
  }, null, { timeout: 5000 });


  Object.assign(pageBreakMenuKeyboard, await page.evaluate(() => ({
    h8AfterShortcut: window.__webSheetDevCellModel?.("CHT", 8, 8)?.raw || "",
    menuClosed: !document.querySelector("#homeCommandMenu"),
    status: document.querySelector("#statusText")?.textContent || "",
  })));


  await page.evaluate(() => document.querySelector('[data-sheet-view="page-break"]')?.click());


  await page.waitForSelector('.sheet-view-guide.is-page-break.is-manual[data-page-break-axis="col"]', { timeout: 5000 });


  const pageBreakDrag = await page.evaluate(async () => {
    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const guide = document.querySelector('.sheet-view-guide.is-page-break.is-manual[data-page-break-axis="col"]');
    const before = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {};
    const rect = guide?.getBoundingClientRect();
    if (!guide || !rect) return { before, after: null, status: "missing" };
    const startX = rect.left + 1;
    const y = rect.top + Math.max(4, rect.height / 2);
    guide.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0, pointerId: 41, clientX: startX, clientY: y }));
    window.dispatchEvent(new PointerEvent("pointermove", { bubbles: true, button: 0, pointerId: 41, clientX: startX + 90, clientY: y }));
    await frame();
    window.dispatchEvent(new PointerEvent("pointerup", { bubbles: true, button: 0, pointerId: 41, clientX: startX + 90, clientY: y }));
    await frame();
    const after = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {};
    const statusAfterMove = document.querySelector("#statusText")?.textContent || "";
    const movedColumn = (after.cols || []).find((col) => !(before.cols || []).includes(col)) || (after.cols || []).find((col) => col !== 8) || null;
    const guideAfterMove = document.querySelector(`.sheet-view-guide.is-page-break.is-manual[data-page-break-axis="col"]${movedColumn ? `[data-page-break-index="${movedColumn}"]` : ""}`);
    const removeRect = guideAfterMove?.getBoundingClientRect();
    const hostRect = document.querySelector("#sheetHost")?.getBoundingClientRect();
    if (guideAfterMove && removeRect && hostRect) {
      const removeStartX = removeRect.left + 1;
      const removeY = removeRect.top + Math.max(4, removeRect.height / 2);
      guideAfterMove.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0, pointerId: 42, clientX: removeStartX, clientY: removeY }));
      window.dispatchEvent(new PointerEvent("pointermove", { bubbles: true, button: 0, pointerId: 42, clientX: hostRect.left + 2, clientY: removeY }));
      await frame();
      window.dispatchEvent(new PointerEvent("pointerup", { bubbles: true, button: 0, pointerId: 42, clientX: hostRect.left + 2, clientY: removeY }));
      await frame();
    }
    const afterRemove = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {};
    return {
      before,
      after,
      afterRemove,
      movedColumn,
      manualGuideCount: document.querySelectorAll(".sheet-view-guide.is-page-break.is-manual").length,
      status: statusAfterMove,
      statusAfterRemove: document.querySelector("#statusText")?.textContent || "",
    };
  });


  await page.keyboard.press("Control+Z");


  await page.waitForFunction((expectedCols) => {
    const cols = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks?.cols || [];
    return Array.isArray(expectedCols) && expectedCols.every((col) => cols.includes(col));
  }, pageBreakDrag.after?.cols || [], { timeout: 5000 });


  pageBreakDrag.undoRemove = await page.evaluate(() => ({
    pageBreaks: window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {},
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction((expectedCols) => {
    const cols = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks?.cols || [];
    return JSON.stringify(cols) === JSON.stringify(expectedCols || []);
  }, pageBreakDrag.afterRemove?.cols || [], { timeout: 5000 });


  pageBreakDrag.redoRemove = await page.evaluate(() => ({
    pageBreaks: window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {},
    status: document.querySelector("#statusText")?.textContent || "",
  }));



  await page.click('[data-ribbon-tab="draw"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "描画", null, { timeout: 10000 });


  const drawShapeCountBefore = await page.locator(".sheet-image, .sheet-image-link").count();


  const drawRibbon = await page.evaluate(() => {
    const panel = document.querySelector('[data-ribbon-panel="draw"]');
    return {
      activeRibbonTab: document.querySelector(".ribbon-tab.is-active")?.textContent,
      panelVisible: Boolean(panel?.classList.contains("is-active")),
      groupLabels: [...(panel?.querySelectorAll(".ribbon-group-label") || [])].map((label) => label.textContent.trim()),
      enabledCommands: [...(panel?.querySelectorAll("[data-ribbon-command]:not(:disabled)") || [])].map((button) => button.dataset.ribbonCommand),
      panelHeight: Math.round(document.querySelector(".ribbon-panels")?.getBoundingClientRect().height || 0),
    };
  });


  await page.click('[data-ribbon-command="draw-color-blue"]');


  await page.click('[data-ribbon-command="draw-pen"]');


  const drawSheetHostBox = await page.locator("#sheetHost").boundingBox();


  if (!drawSheetHostBox) throw new Error("Sheet host is not visible for draw ribbon test.");


  await page.mouse.move(drawSheetHostBox.x + 260, drawSheetHostBox.y + 190);


  await page.mouse.down();


  await page.mouse.move(drawSheetHostBox.x + 310, drawSheetHostBox.y + 220, { steps: 4 });


  await page.mouse.move(drawSheetHostBox.x + 360, drawSheetHostBox.y + 190, { steps: 4 });


  await page.mouse.up();


  await page.waitForFunction(() => (window.__webSheetDevSheetModel?.("CHT")?.images || []).some((image) => image.shape?.ink), null, { timeout: 5000 });


  const drawInk = await page.evaluate(() => {
    const images = window.__webSheetDevSheetModel?.("CHT")?.images || [];
    const ink = images.find((image) => image.shape?.ink);
    return {
      inkCount: images.filter((image) => image.shape?.ink).length,
      selectedFrames: document.querySelectorAll(".sheet-image-selection-frame").length,
      outlineColor: ink?.shape?.outlineColor || "",
      customPathPoints: ink?.shape?.customPath?.commands?.length || 0,
      activePen: document.querySelector('[data-ribbon-command="draw-pen"]')?.classList.contains("is-active"),
      activeColor: document.querySelector('[data-ribbon-command="draw-color-blue"]')?.classList.contains("is-active"),
    };
  });


  await page.click('[data-ribbon-tab="draw"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "描画", null, { timeout: 10000 });


  await page.click('[data-ribbon-command="draw-convert-shape"]');


  await page.waitForFunction(() => (window.__webSheetDevSheetModel?.("CHT")?.images || []).some((image) => image.shape?.type === "customPath" && image.shape?.ink === false), null, { timeout: 5000 });


  await page.click('[data-ribbon-tab="draw"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "描画", null, { timeout: 10000 });


  await page.click('[data-ribbon-command="draw-highlighter"]');


  await page.mouse.move(drawSheetHostBox.x + 280, drawSheetHostBox.y + 245);


  await page.mouse.down();


  await page.mouse.move(drawSheetHostBox.x + 390, drawSheetHostBox.y + 245, { steps: 5 });


  await page.mouse.up();


  await page.waitForFunction(() => (window.__webSheetDevSheetModel?.("CHT")?.images || []).some((image) => image.shape?.ink), null, { timeout: 5000 });


  await page.click('[data-ribbon-tab="draw"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "描画", null, { timeout: 10000 });


  await page.click('[data-ribbon-command="draw-clear-ink"]');


  await page.waitForFunction(() => !(window.__webSheetDevSheetModel?.("CHT")?.images || []).some((image) => image.shape?.ink), null, { timeout: 5000 });


  const drawConverted = await page.evaluate(() => {
    const images = window.__webSheetDevSheetModel?.("CHT")?.images || [];
    return {
      customPaths: images.filter((image) => image.shape?.type === "customPath").length,
      remainingInk: images.filter((image) => image.shape?.ink).length,
      status: document.querySelector("#statusText")?.textContent || "",
    };
  });


  for (let index = 0; index < 8; index += 1) {
    const count = await page.locator(".sheet-image, .sheet-image-link").count();
    if (count === drawShapeCountBefore) break;
    await page.keyboard.press("Control+Z");
    await page.waitForTimeout(80);
  }


  await page.waitForFunction((beforeCount) => document.querySelectorAll(".sheet-image, .sheet-image-link").length === beforeCount, drawShapeCountBefore, { timeout: 5000 });



  await page.click('[data-ribbon-tab="insert"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "挿入", null, { timeout: 10000 });



  await page.click('[data-ribbon-command="insert-header-footer"]');


  await page.waitForSelector('[data-data-dialog="header-footer"]', { timeout: 5000 });


  const headerFooterDialog = await page.evaluate(() => ({
    title: document.querySelector("#headerFooterTitle")?.textContent || "",
    legends: [...document.querySelectorAll("[data-data-dialog='header-footer'] legend")].map((legend) => legend.textContent.trim()),
    fields: [...document.querySelectorAll("[data-header-footer-field]")].map((input) => input.dataset.headerFooterField),
    values: Object.fromEntries([...document.querySelectorAll("[data-header-footer-field]")].map((input) => [input.dataset.headerFooterField, input.value])),
    checks: [...document.querySelectorAll("[data-header-footer-check]")].map((input) => input.dataset.headerFooterCheck),
  }));


  await page.locator("[data-header-footer-field='oddHeader.center']").fill("WebSheet Header");


  await page.locator("[data-header-footer-field='oddFooter.right']").fill("Page &[Page]");


  await page.locator("[data-header-footer-check='differentFirst']").check();


  await page.locator("[data-data-dialog='header-footer'] [data-data-dialog-action='ok']").click();


  await page.waitForSelector('[data-data-dialog="header-footer"]', { state: "detached", timeout: 5000 });


  const headerFooterApplied = await page.evaluate(() => ({
    model: window.__webSheetDevSheetModel?.("CHT")?.headerFooter || null,
    pageLayoutHeaderFooter: window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.headerFooter || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.headerFooter?.oddHeader?.center === "Layout Header", null, { timeout: 5000 });


  const headerFooterUndo = await page.evaluate(() => ({
    model: window.__webSheetDevSheetModel?.("CHT")?.headerFooter || null,
    pageLayoutHeaderFooter: window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.headerFooter || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.headerFooter?.oddHeader?.center === "WebSheet Header", null, { timeout: 5000 });


  const headerFooterRedo = await page.evaluate(() => ({
    model: window.__webSheetDevSheetModel?.("CHT")?.headerFooter || null,
    pageLayoutHeaderFooter: window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.headerFooter || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }));



  await page.evaluate(() => {
    window.__webSheetDevSetActiveSheet?.("CHT");
    [
      [1, 1, "地域"],
      [1, 2, "数値"],
      [2, 1, "東京"],
      [2, 2, 120],
      [3, 1, "大阪"],
      [3, 2, 85],
      [4, 1, "福岡"],
      [4, 2, 65],
    ].forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
    window.__webSheetDevSelectRange?.(1, 1, 4, 2);
  });


  await page.waitForFunction(() => String(window.__webSheetDevCellModel?.("CHT", 4, 2)?.raw ?? "") === "65", null, { timeout: 5000 });


  const mapChartStartCount = await page.locator(".sheet-image, .sheet-image-link").count();


  await page.click('[data-ribbon-tab="insert"]');


  await page.click('[data-ribbon-command="insert-map-chart"]');


  await page.waitForFunction((startCount) => document.querySelectorAll(".sheet-image, .sheet-image-link").length > startCount, mapChartStartCount, { timeout: 5000 });


  const insertMapChart = await page.evaluate((startCount) => {
    const images = window.__webSheetDevSheetModel?.("CHT")?.images || [];
    const image = images[images.length - 1] || null;
    const domImage = [...document.querySelectorAll(".sheet-image, .sheet-image-link .sheet-image")].at(-1);
    const source = decodeURIComponent(domImage?.getAttribute("src") || image?.src || "");
    return {
      countDelta: document.querySelectorAll(".sheet-image, .sheet-image-link").length - startCount,
      chartType: image?.chart?.type || "",
      chartRange: image?.chart?.sourceRange || "",
      alt: domImage?.getAttribute("alt") || "",
      hasMapSvg: source.includes('data-chart-map="true"'),
      hasRegion: source.includes('data-map-region="東京"'),
    };
  }, mapChartStartCount);


  for (let index = 0; index < 9; index += 1) {
    await page.keyboard.press("Control+Z");
    await page.waitForTimeout(50);
  }


  await page.waitForFunction((startCount) => document.querySelectorAll(".sheet-image, .sheet-image-link").length === startCount, mapChartStartCount, { timeout: 5000 });



  await page.click('[data-ribbon-tab="insert"]');


  await page.click("[data-ribbon-shape-menu]");


  await page.waitForSelector("#ribbonShapeMenu", { timeout: 5000 });


  const insertShapeMenu = await page.evaluate(() => ({
    sections: [...document.querySelectorAll("#ribbonShapeMenu .ribbon-shape-section-title")].map((label) => label.textContent.trim()),
    choiceCount: document.querySelectorAll("#ribbonShapeMenu [data-ribbon-shape-choice]").length,
    focusedBefore: document.activeElement?.dataset?.ribbonShapeChoice || "",
    hasLine: Boolean(document.querySelector('#ribbonShapeMenu [data-ribbon-shape-choice="line"]')),
    hasRectangle: Boolean(document.querySelector('#ribbonShapeMenu [data-ribbon-shape-choice="rect"]')),
    hasFlowchart: Boolean(document.querySelector('#ribbonShapeMenu [data-ribbon-shape-choice="flowDecision"]')),
    hasCallout: Boolean(document.querySelector('#ribbonShapeMenu [data-ribbon-shape-choice="calloutCloud"]')),
  }));


  await page.keyboard.press("ArrowRight");


  Object.assign(insertShapeMenu, await page.evaluate(() => ({
    focusedAfterRight: document.activeElement?.dataset?.ribbonShapeChoice || "",
  })));


  const shapeDomCountBefore = await page.locator(".sheet-image, .sheet-image-link").count();


  await page.click('#ribbonShapeMenu [data-ribbon-shape-choice="flowDecision"]');


  await page.waitForFunction(() => !document.querySelector("#ribbonShapeMenu"), null, { timeout: 5000 });


  const insertShapeDraw = await page.evaluate((beforeCount) => ({
    beforeCount,
    countAfterChoice: document.querySelectorAll(".sheet-image, .sheet-image-link").length,
    drawingMode: document.body.classList.contains("is-drawing-shape"),
    statusAfterChoice: document.querySelector("#statusText")?.textContent?.trim() || "",
  }), shapeDomCountBefore);


  const shapeHostBox = await page.locator("#sheetHost").boundingBox();


  await page.mouse.move(shapeHostBox.x + 160, shapeHostBox.y + 120);


  await page.mouse.down();


  await page.waitForSelector(".sheet-shape-draft", { timeout: 5000 });


  await page.mouse.move(shapeHostBox.x + 340, shapeHostBox.y + 210, { steps: 8 });


  Object.assign(insertShapeDraw, await page.evaluate(() => {
    const draft = document.querySelector(".sheet-shape-draft");
    const rect = draft?.getBoundingClientRect();
    return {
      draftVisible: Boolean(draft),
      draftWidth: Math.round(rect?.width || 0),
      draftHeight: Math.round(rect?.height || 0),
    };
  }));


  await page.mouse.up();


  await page.waitForFunction((beforeCount) => document.querySelectorAll(".sheet-image, .sheet-image-link").length > beforeCount, shapeDomCountBefore, { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => {
    const selected = document.querySelector(".sheet-image.is-selected, .sheet-image-link.is-selected");
    const image = selected?.matches?.(".sheet-image") ? selected : selected?.querySelector?.(".sheet-image");
    const rect = image?.getBoundingClientRect();
    const tab = document.querySelector('[data-ribbon-tab="shape-format"]');
    const shapeStyleButtons = [...document.querySelectorAll('[data-ribbon-panel="shape-format"] .ribbon-shape-style-group .ribbon-tool span')];
    const shapePanel = document.querySelector('[data-ribbon-panel="shape-format"]');
    const decodedShapeSource = image?.src ? decodeURIComponent(image.src) : "";
    return {
      countAfterDraw: document.querySelectorAll(".sheet-image, .sheet-image-link").length,
      selected: Boolean(selected),
      alt: image?.getAttribute("alt") || "",
      width: Math.round(rect?.width || 0),
      height: Math.round(rect?.height || 0),
      shapeClassApplied: image?.classList?.contains("sheet-shape") || false,
      shapeObjectFit: image ? getComputedStyle(image).objectFit : "",
      shapePreserveAspectRatioNone: decodedShapeSource.includes('preserveAspectRatio="none"'),
      draftAfterDraw: Boolean(document.querySelector(".sheet-shape-draft")),
      drawingModeAfterDraw: document.body.classList.contains("is-drawing-shape"),
      contextualTabVisible: Boolean(tab && !tab.classList.contains("hidden")),
      contextualTabActive: document.querySelector(".ribbon-tab.is-active")?.textContent?.trim() || "",
      resizeHandleCount: document.querySelectorAll(".sheet-image-resize-handle").length,
      sizeInputsAfterDraw: [...document.querySelectorAll("[data-shape-size]")].map((input) => input.value),
      shapeStyleTextOverflowHidden: shapeStyleButtons.every((span) => getComputedStyle(span).overflow === "hidden"),
      shapeStyleTextFits: shapeStyleButtons.every((span) => span.scrollWidth <= span.clientWidth + 1),
      shapeTextToolbarPresent: Boolean(
        shapePanel?.querySelector("[data-shape-text-fill-menu]") &&
        shapePanel?.querySelector("[data-shape-text-outline-menu]") &&
        shapePanel?.querySelector('[data-ribbon-command="shape-edit-text"]'),
      ),
    };
  }));


  await page.locator(".sheet-image-selection-frame.is-selected, .sheet-image-selection-frame").first().click({ button: "right" });


  await page.waitForSelector("#cellContextMenu", { timeout: 5000 });


  await page.click('#cellContextMenu [data-cell-menu-action="object-link"]');


  await page.waitForSelector("#objectLinkDialog", { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    objectLinkDialogTitle: document.querySelector("#objectLinkDialogTitle")?.textContent || "",
    objectLinkKinds: [...document.querySelectorAll("[data-object-link-kind]")].map((button) => ({
      kind: button.dataset.objectLinkKind || "",
      text: button.textContent.trim(),
      pressed: button.getAttribute("aria-pressed") || "",
    })),
    objectLinkInitialCaption: document.querySelector("[data-object-link-address-caption]")?.textContent || "",
    objectLinkDocumentPanelHiddenInitial: document.querySelector("[data-object-link-document-panel]")?.hidden,
  })));


  await page.click('[data-object-link-kind="document"]');


  await page.locator('[data-object-link-field="url"]').fill("B2");


  await page.click('[data-object-link-document-sheet-index="1"]');


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    objectLinkDocumentCaption: document.querySelector("[data-object-link-address-caption]")?.textContent || "",
    objectLinkDocumentPanelHidden: document.querySelector("[data-object-link-document-panel]")?.hidden,
    objectLinkDocumentPressed: document.querySelector('[data-object-link-kind="document"]')?.getAttribute("aria-pressed") || "",
    objectLinkSelectedSheet: document.querySelector("[data-object-link-document-sheet-index].is-active")?.textContent?.trim() || "",
  })));


  await page.click('[data-object-link-dialog-action="ok"]');


  await page.waitForSelector("#objectLinkDialog", { state: "detached", timeout: 5000 });


  await page.waitForFunction(
    () => (window.__webSheetDevSheetModel?.("CHT")?.images || []).some((image) => image.hyperlink === "#'AI製品紹介'!B2"),
    null,
    { timeout: 5000 },
  );


  Object.assign(insertShapeDraw, await page.evaluate(() => {
    const linked = (window.__webSheetDevSheetModel?.("CHT")?.images || []).find((image) => image.hyperlink === "#'AI製品紹介'!B2");
    return {
      objectLinkHyperlink: linked?.hyperlink || "",
      objectLinkStatus: document.querySelector("#statusText")?.textContent?.trim() || "",
    };
  }));


  const shapeFrameBox = await page.locator(".sheet-image-selection-frame").boundingBox();


  await page.mouse.move(shapeFrameBox.x + 16, shapeFrameBox.y + 16);


  await page.mouse.down();


  await page.mouse.move(shapeFrameBox.x + 58, shapeFrameBox.y + 42, { steps: 6 });


  await page.mouse.up();


  await page.waitForTimeout(120);


  Object.assign(insertShapeDraw, await page.evaluate(() => {
    const frame = document.querySelector(".sheet-image-selection-frame");
    const rect = frame?.getBoundingClientRect();
    return {
      moveLeft: Math.round(rect?.left || 0),
      moveTop: Math.round(rect?.top || 0),
      moveStatus: document.querySelector("#statusText")?.textContent?.trim() || "",
    };
  }));


  const resizeHandleBox = await page.locator(".sheet-image-resize-handle.is-se").boundingBox();


  await page.mouse.move(resizeHandleBox.x + 3, resizeHandleBox.y + 3);


  await page.mouse.down();


  await page.mouse.move(resizeHandleBox.x + 43, resizeHandleBox.y + 33, { steps: 6 });


  await page.mouse.up();


  await page.waitForTimeout(120);


  Object.assign(insertShapeDraw, await page.evaluate(() => {
    const frame = document.querySelector(".sheet-image-selection-frame");
    const rect = frame?.getBoundingClientRect();
    return {
      resizedWidth: Math.round(rect?.width || 0),
      resizedHeight: Math.round(rect?.height || 0),
      sizeInputsAfterResize: [...document.querySelectorAll("[data-shape-size]")].map((input) => input.value),
      resizeStatus: document.querySelector("#statusText")?.textContent?.trim() || "",
    };
  }));


  const lockedResizeHandleBox = await page.locator(".sheet-image-resize-handle.is-se").boundingBox();


  await page.keyboard.down("Shift");


  await page.mouse.move(lockedResizeHandleBox.x + 3, lockedResizeHandleBox.y + 3);


  await page.mouse.down();


  await page.mouse.move(lockedResizeHandleBox.x + 83, lockedResizeHandleBox.y + 13, { steps: 6 });


  await page.mouse.up();


  await page.keyboard.up("Shift");


  await page.waitForTimeout(120);


  Object.assign(insertShapeDraw, await page.evaluate(() => {
    const frame = document.querySelector(".sheet-image-selection-frame");
    const rect = frame?.getBoundingClientRect();
    return {
      lockedResizeWidth: Math.round(rect?.width || 0),
      lockedResizeHeight: Math.round(rect?.height || 0),
    };
  }));


  await page.keyboard.press("T");


  await page.waitForSelector("#shapeTextEditor", { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => {
    const editor = document.querySelector("#shapeTextEditor");
    const style = editor ? getComputedStyle(editor) : null;
    return {
      directTextEditorVisible: Boolean(editor),
      directTextEditorText: editor?.textContent || "",
      directTextEditorCursor: style?.cursor || "",
      directTextEditorDecoration: style?.textDecorationLine || "",
    };
  }));


  await page.keyboard.press("Escape");


  await page.waitForFunction(() => !document.querySelector("#shapeTextEditor"), null, { timeout: 5000 });


  await page.locator('.sheet-cell[data-address="A1"]').click();


  await page.waitForTimeout(120);


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    contextualHiddenAfterCell: document.querySelector('[data-ribbon-tab="shape-format"]')?.classList.contains("hidden"),
    activeTabAfterCell: document.querySelector(".ribbon-tab.is-active")?.textContent?.trim() || "",
    frameAfterCell: Boolean(document.querySelector(".sheet-image-selection-frame")),
  })));


  await page.locator(".sheet-image").last().click();


  await page.waitForSelector(".sheet-image-selection-frame", { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    contextualVisibleAfterReselect: !document.querySelector('[data-ribbon-tab="shape-format"]')?.classList.contains("hidden"),
  })));


  await page.click("[data-shape-rotate-menu]");


  await page.waitForSelector("#shapeFormatMenu", { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    rotateMenuItems: [...document.querySelectorAll('#shapeFormatMenu [data-shape-format-action="rotate"]')].map((button) => button.textContent.trim()),
    rotateMenuFocusedBefore: document.activeElement?.dataset?.shapeFormatValue || "",
  })));


  await page.keyboard.press("Enter");


  await page.waitForTimeout(120);


  Object.assign(insertShapeDraw, await page.evaluate(() => {
    const image = document.querySelector(".sheet-image.is-selected, .sheet-image-link.is-selected .sheet-image");
    return {
      rotationInputAfterMenu: document.querySelector("[data-shape-rotation]")?.value || "",
      shapeTransformAfterMenu: image ? getComputedStyle(image).transform : "",
    };
  }));


  await page.click("[data-shape-fill-menu]");


  await page.waitForSelector("#shapeFormatMenu", { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    fillPatternChoiceCount: document.querySelectorAll('#shapeFormatMenu [data-shape-format-action="fill-pattern"]').length,
  })));


  await page.keyboard.press("Escape");


  await page.click("[data-shape-outline-menu]");


  await page.waitForSelector("#shapeFormatMenu", { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    lineAdvancedChoiceCount: document.querySelectorAll('#shapeFormatMenu [data-shape-format-action="outline-compound"], #shapeFormatMenu [data-shape-format-action="outline-cap"], #shapeFormatMenu [data-shape-format-action="outline-join"]').length,
    arrowTypeChoiceCount: document.querySelectorAll('#shapeFormatMenu [data-shape-format-action="arrow-start-type"], #shapeFormatMenu [data-shape-format-action="arrow-end-type"]').length,
    arrowTypeLabels: [...document.querySelectorAll('#shapeFormatMenu [data-shape-format-action="arrow-start-type"], #shapeFormatMenu [data-shape-format-action="arrow-end-type"]')]
      .map((button) => button.textContent.trim()),
    arrowSizeChoiceCount: document.querySelectorAll('#shapeFormatMenu [data-shape-format-action="arrow-start-size"], #shapeFormatMenu [data-shape-format-action="arrow-end-size"]').length,
    arrowSizeLabels: [...document.querySelectorAll('#shapeFormatMenu [data-shape-format-action="arrow-start-size"], #shapeFormatMenu [data-shape-format-action="arrow-end-size"]')]
      .map((button) => button.textContent.trim()),
  })));


  await page.keyboard.press("Escape");


  await page.click("[data-shape-effects-menu]");


  await page.waitForSelector("#shapeFormatMenu", { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    shapeEffectChoiceCount: document.querySelectorAll('#shapeFormatMenu [data-shape-format-action="glow"], #shapeFormatMenu [data-shape-format-action="soft-edge"], #shapeFormatMenu [data-shape-format-action="reflection"]').length,
  })));


  await page.click('#shapeFormatMenu [data-shape-format-action="glow"][data-shape-format-value="blue"]');


  await page.waitForTimeout(120);


  Object.assign(insertShapeDraw, await page.evaluate(() => {
    const image = document.querySelector(".sheet-image.is-selected, .sheet-image-link.is-selected .sheet-image");
    const source = decodeURIComponent(image?.getAttribute("src") || "");
    return {
      shapeEffectSourceHasFilter: source.includes('id="shapeEffects"') && source.includes("feDropShadow"),
    };
  }));


  await page.click("[data-shape-text-fill-menu]");


  await page.waitForSelector("#shapeFormatMenu", { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    textFillMenuChoices: document.querySelectorAll('#shapeFormatMenu [data-shape-format-action="text-fill-color"]').length,
  })));


  await page.keyboard.press("Escape");


  await page.click("[data-shape-text-outline-menu]");


  await page.waitForSelector("#shapeFormatMenu", { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    textOutlineMenuChoices: document.querySelectorAll('#shapeFormatMenu [data-shape-format-action="text-outline-color"]').length,
    textOutlineNoneVisible: Boolean(document.querySelector('#shapeFormatMenu [data-shape-format-action="text-outline-none"]')),
  })));


  await page.keyboard.press("Escape");


  await page.click(".sheet-image-selection-frame", { button: "right" });


  await page.click('#cellContextMenu [data-cell-menu-action="object-format"]');


  await page.waitForSelector("#shapeFormatPane", { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    shapeTextPaneVisible: Boolean(document.querySelector("#shapeFormatPane")),
    shapeTextPaneFields: [
      "textFontFamily",
      "textFontSize",
      "textColor",
      "textOutlineColor",
      "textOutlineWidth",
      "textBold",
      "textItalic",
      "textUnderline",
      "textAlign",
      "textVAlign",
      "textMarginLeft",
      "textMarginRight",
      "textMarginTop",
      "textMarginBottom",
    ].filter((field) => document.querySelector(`[data-shape-format-pane-field="${field}"]`)),
    shapeLinePaneFields: [
      "lineDash",
      "lineCompound",
      "lineCap",
      "lineJoin",
      "arrow",
      "lineStartArrowType",
      "lineEndArrowType",
      "lineStartArrowSize",
      "lineEndArrowSize",
    ].filter((field) => document.querySelector(`[data-shape-format-pane-field="${field}"]`)),
    shapeAdvancedPaneFields: [
      "fillPattern",
      "glow",
      "softEdges",
      "reflection",
    ].filter((field) => document.querySelector(`[data-shape-format-pane-field="${field}"]`)),
    shapePaneTabs: [...document.querySelectorAll("[data-shape-format-pane-tab]")].map((button) => button.textContent.trim()),
    shapePaneActiveInitial: [...document.querySelectorAll("[data-shape-format-pane-section].is-active h3")].map((heading) => heading.textContent.trim()),
  })));


  await page.click('[data-shape-format-pane-tab="effects"]');


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    shapePaneActiveEffects: [...document.querySelectorAll("[data-shape-format-pane-section].is-active h3")].map((heading) => heading.textContent.trim()),
  })));


  await page.click('[data-shape-format-pane-tab="size-props"]');


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    shapePaneActiveSizeProps: [...document.querySelectorAll("[data-shape-format-pane-section].is-active h3")].map((heading) => heading.textContent.trim()),
  })));


  await page.click('[data-shape-format-pane-action="selection-pane"]');


  await page.waitForSelector("#shapeSelectionPane", { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    selectionPaneVisible: Boolean(document.querySelector("#shapeSelectionPane")),
    selectionPaneRows: document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-row").length,
    selectionPaneSelectedRows: document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-row.is-selected").length,
    selectionPaneHasFormatOffset: document.querySelector("#shapeSelectionPane")?.classList.contains("has-format-pane") || false,
    selectionPaneNames: [...document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-object")].map((button) => button.textContent.trim()),
  })));


  const copyCountBefore = await page.locator(".sheet-image, .sheet-image-link").count();


  const copyFrameBox = await page.locator(".sheet-image-selection-frame").boundingBox();


  await page.keyboard.down("Control");


  await page.mouse.move(copyFrameBox.x + 20, copyFrameBox.y + 20);


  await page.mouse.down();


  await page.mouse.move(copyFrameBox.x + 76, copyFrameBox.y + 20, { steps: 6 });


  await page.mouse.up();


  await page.keyboard.up("Control");


  await page.waitForFunction((beforeCount) => document.querySelectorAll(".sheet-image, .sheet-image-link").length > beforeCount, copyCountBefore, { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    countAfterCtrlDragCopy: document.querySelectorAll(".sheet-image, .sheet-image-link").length,
    selectionPaneRowsAfterCopy: document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-row").length,
    selectedAfterCtrlDragCopy: new Set(
      [...document.querySelectorAll(".sheet-image.is-selected, .sheet-image-link.is-selected")]
        .map((item) => item.dataset.imageId || item.querySelector?.("[data-image-id]")?.dataset.imageId || "")
        .filter(Boolean),
    ).size,
    statusAfterCtrlDragCopy: document.querySelector("#statusText")?.textContent?.trim() || "",
  })));


  const constrainedFrameStart = await page.locator(".sheet-image-selection-frame").boundingBox();


  await page.keyboard.down("Shift");


  await page.mouse.move(constrainedFrameStart.x + 16, constrainedFrameStart.y + 16);


  await page.mouse.down();


  await page.mouse.move(constrainedFrameStart.x + 64, constrainedFrameStart.y + 42, { steps: 6 });


  await page.mouse.up();


  await page.keyboard.up("Shift");


  await page.waitForTimeout(120);


  const constrainedFrameEnd = await page.locator(".sheet-image-selection-frame").boundingBox();


  Object.assign(insertShapeDraw, {
    shiftMoveDeltaX: Math.round(constrainedFrameEnd.x - constrainedFrameStart.x),
    shiftMoveDeltaY: Math.round(constrainedFrameEnd.y - constrainedFrameStart.y),
  });


  await page.click('#shapeSelectionPane .shape-selection-pane-row.is-selected [data-shape-selection-pane-action="rename"]');


  await page.fill("#shapeSelectionPane [data-shape-selection-pane-rename]", "判断コピー");


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => [...document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-object")].some((button) => button.textContent.trim() === "判断コピー"), null, { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    selectionPaneRenamed: [...document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-object")].map((button) => button.textContent.trim()),
    selectedLabelAfterRename: document.querySelector("#selectedCellLabel")?.textContent?.trim() || "",
  })));


  await page.focus("#shapeSelectionPane .shape-selection-pane-row.is-selected [data-shape-selection-pane-object]");


  await page.keyboard.press("ArrowDown");


  await page.waitForTimeout(120);


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    selectionPaneKeyboardSelected: document.querySelector("#shapeSelectionPane .shape-selection-pane-row.is-selected .shape-selection-pane-object")?.textContent?.trim() || "",
  })));


  await page.keyboard.press("Home");


  await page.waitForTimeout(80);


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    selectionPaneHomeSelected: document.querySelector("#shapeSelectionPane .shape-selection-pane-row.is-selected .shape-selection-pane-object")?.textContent?.trim() || "",
    selectionPaneFirstName: document.querySelector("#shapeSelectionPane .shape-selection-pane-row:first-child .shape-selection-pane-object")?.textContent?.trim() || "",
  })));


  await page.keyboard.press("End");


  await page.waitForTimeout(80);


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    selectionPaneEndSelected: document.querySelector("#shapeSelectionPane .shape-selection-pane-row.is-selected .shape-selection-pane-object")?.textContent?.trim() || "",
    selectionPaneLastName: document.querySelector("#shapeSelectionPane .shape-selection-pane-row:last-child .shape-selection-pane-object")?.textContent?.trim() || "",
  })));


  await page.keyboard.press("Control+A");


  await page.waitForTimeout(80);


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    selectionPaneCtrlASelectedRows: document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-row.is-selected").length,
  })));


  await page.click("#shapeSelectionPane .shape-selection-pane-row:first-child .shape-selection-pane-object");


  await page.focus("#shapeSelectionPane .shape-selection-pane-row.is-selected [data-shape-selection-pane-object]");


  await page.keyboard.press("Space");


  await page.waitForFunction(() => document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-row.is-hidden").length >= 1, null, { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    selectionPaneSpaceHiddenRows: document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-row.is-hidden").length,
    countAfterSelectionPaneSpaceHide: document.querySelectorAll(".sheet-image, .sheet-image-link").length,
  })));


  await page.keyboard.press("Space");


  await page.waitForFunction(() => document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-row.is-hidden").length === 0, null, { timeout: 5000 });


  await page.focus("#shapeSelectionPane .shape-selection-pane-row.is-selected [data-shape-selection-pane-object]");


  await page.keyboard.press("End");


  await page.waitForFunction(() => {
    const selected = document.querySelector("#shapeSelectionPane .shape-selection-pane-row.is-selected .shape-selection-pane-object")?.textContent?.trim();
    const last = document.querySelector("#shapeSelectionPane .shape-selection-pane-row:last-child .shape-selection-pane-object")?.textContent?.trim();
    return selected && selected === last;
  }, null, { timeout: 5000 });


  await page.click('#shapeSelectionPane [data-shape-selection-pane-action="forward"]');


  await page.waitForTimeout(120);


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    statusAfterSelectionPaneForward: document.querySelector("#statusText")?.textContent?.trim() || "",
  })));


  const orderBefore = await page.evaluate(() => [...document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-object")].map((button) => button.textContent.trim()).join("|"));


  const firstDragHandleBox = await page.locator("#shapeSelectionPane [data-shape-selection-pane-drag]").first().boundingBox();


  const lastRowBox = await page.locator("#shapeSelectionPane .shape-selection-pane-row").last().boundingBox();


  await page.mouse.move(firstDragHandleBox.x + 4, firstDragHandleBox.y + 10);


  await page.mouse.down();


  await page.mouse.move(lastRowBox.x + 20, lastRowBox.y + lastRowBox.height - 2, { steps: 6 });


  await page.mouse.up();


  await page.waitForTimeout(120);


  Object.assign(insertShapeDraw, await page.evaluate((before) => ({
    selectionPaneOrderBefore: before,
    selectionPaneOrderAfter: [...document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-object")].map((button) => button.textContent.trim()).join("|"),
    statusAfterSelectionPaneOrder: document.querySelector("#statusText")?.textContent?.trim() || "",
  }), orderBefore));


  await page.click('#shapeSelectionPane [data-shape-selection-pane-action="hide-all"]');


  await page.waitForFunction(() => document.querySelectorAll(".sheet-image, .sheet-image-link").length === 0, null, { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    selectionPaneHiddenRows: document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-row.is-hidden").length,
    countAfterSelectionPaneHide: document.querySelectorAll(".sheet-image, .sheet-image-link").length,
    statusAfterSelectionPaneHide: document.querySelector("#statusText")?.textContent?.trim() || "",
  })));


  await page.click('#shapeSelectionPane [data-shape-selection-pane-action="show-all"]');


  await page.waitForFunction((beforeCount) => document.querySelectorAll(".sheet-image, .sheet-image-link").length > beforeCount, shapeDomCountBefore, { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => ({
    selectionPaneHiddenRowsAfterShow: document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-row.is-hidden").length,
    countAfterSelectionPaneShow: document.querySelectorAll(".sheet-image, .sheet-image-link").length,
    statusAfterSelectionPaneShow: document.querySelector("#statusText")?.textContent?.trim() || "",
  })));


  await page.waitForSelector("#shapeFormatPane", { timeout: 5000 });


  await page.selectOption('[data-shape-format-pane-field="textFontSize"]', "24");


  await page.click('[data-shape-format-pane-action="edit-text"]');


  await page.waitForSelector("#shapeTextEditor", { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate(() => {
    const editor = document.querySelector("#shapeTextEditor");
    const style = editor ? getComputedStyle(editor) : null;
    return {
      paneTextEditorVisible: Boolean(editor),
      paneTextEditorFontSize: style?.fontSize || "",
    };
  }));


  await page.keyboard.press("Escape");


  await page.waitForFunction(() => !document.querySelector("#shapeTextEditor"), null, { timeout: 5000 });


  await page.click("#shapeSelectionPane .shape-selection-pane-row:first-child .shape-selection-pane-object");


  await page.focus("#shapeSelectionPane .shape-selection-pane-row.is-selected [data-shape-selection-pane-object]");


  const rowCountBeforePaneDelete = await page.locator("#shapeSelectionPane .shape-selection-pane-row").count();


  await page.keyboard.press("Delete");


  await page.waitForFunction((before) => document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-row").length === before - 1, rowCountBeforePaneDelete, { timeout: 5000 });


  Object.assign(insertShapeDraw, await page.evaluate((before) => ({
    selectionPaneRowsBeforeDeleteKey: before,
    selectionPaneRowsAfterDeleteKey: document.querySelectorAll("#shapeSelectionPane .shape-selection-pane-row").length,
    statusAfterSelectionPaneDeleteKey: document.querySelector("#statusText")?.textContent?.trim() || "",
  }), rowCountBeforePaneDelete));


  await page.evaluate(() => document.activeElement?.blur?.());


  for (let index = 0; index < 60; index += 1) {
    const count = await page.locator(".sheet-image, .sheet-image-link").count();
    if (count <= shapeDomCountBefore) break;
    await page.keyboard.press("Control+Z");
    await page.waitForTimeout(60);
  }


  await page.waitForFunction((beforeCount) => document.querySelectorAll(".sheet-image, .sheet-image-link").length === beforeCount, shapeDomCountBefore, { timeout: 5000 });


  insertShapeDraw.countAfterUndo = await page.locator(".sheet-image, .sheet-image-link").count();


  await page.click('#shapeSelectionPane [data-shape-selection-pane-action="close"]');


  await page.waitForFunction(() => !document.querySelector("#shapeSelectionPane"), null, { timeout: 5000 });



  const insertTextObjectCommands = [
    "insert-icons",
    "insert-3d-model",
    "insert-smartart",
    "insert-screenshot",
    "insert-wordart",
    "insert-signature-line",
    "insert-object",
    "insert-equation",
    "insert-slicer",
    "insert-timeline",
  ];


  const textObjectStartCount = await page.locator(".sheet-image, .sheet-image-link").count();


  let insertIconDialog = null;


  let signatureLineDialog = null;


  let insertObjectDialog = null;


  let expectedInsertedObjectCount = textObjectStartCount;


  for (const command of insertTextObjectCommands) {
    await page.click('[data-ribbon-tab="insert"]');
    await page.click(`[data-ribbon-command="${command}"]`);
    if (command === "insert-icons") {
      await page.waitForSelector('[data-data-dialog="insert-icon"]', { timeout: 5000 });
      await page.locator("[data-insert-icon-search]").fill("チェック");
      await page.waitForFunction(() => [...document.querySelectorAll("[data-insert-icon-kind]")].filter((button) => !button.hidden).length === 1, null, { timeout: 5000 });
      insertIconDialog = await page.evaluate(() => ({
        title: document.querySelector("#insertIconTitle")?.textContent || "",
        searchPlaceholder: document.querySelector("[data-data-dialog='insert-icon'] input")?.getAttribute("placeholder") || "",
        choiceCount: document.querySelectorAll("[data-insert-icon-kind]").length,
        visibleAfterSearch: [...document.querySelectorAll("[data-insert-icon-kind]")].filter((button) => !button.hidden).map((button) => button.textContent.trim()),
        emptyHidden: document.querySelector("[data-insert-icon-empty]")?.hidden === true,
        labels: [...document.querySelectorAll("[data-insert-icon-kind]")].map((button) => button.textContent.trim()),
      }));
      await page.click('[data-insert-icon-kind="check"]');
      await page.waitForSelector('[data-data-dialog="insert-icon"]', { state: "detached", timeout: 5000 });
    } else if (command === "insert-signature-line") {
      await page.waitForSelector('[data-data-dialog="signature-line"]', { timeout: 5000 });
      signatureLineDialog = await page.evaluate(() => ({
        title: document.querySelector("#signatureLineTitle")?.textContent || "",
        captions: [...document.querySelectorAll('[data-data-dialog="signature-line"] .excel-dialog-caption')].map((item) => item.textContent.trim()),
        signer: document.querySelector("[data-signature-field='signer']")?.value || "",
        allowComments: document.querySelector("[data-signature-check='allowComments']")?.checked === true,
        showDate: document.querySelector("[data-signature-check='showDate']")?.checked === true,
      }));
      await page.locator("[data-signature-field='signer']").fill("ZJ");
      await page.locator("[data-signature-field='title']").fill("代表");
      await page.locator("[data-signature-field='email']").fill("zj@example.com");
      await page.locator("[data-signature-field='instructions']").fill("内容を確認して署名してください。");
      await page.click('[data-data-dialog-action="signature-line-ok"]');
      await page.waitForSelector('[data-data-dialog="signature-line"]', { state: "detached", timeout: 5000 });
    } else if (command === "insert-object") {
      await page.waitForSelector('[data-data-dialog="insert-object"]', { timeout: 5000 });
      const initialObjectDialog = await page.evaluate(() => ({
        title: document.querySelector("#insertObjectTitle")?.textContent || "",
        activeTab: document.querySelector("[data-object-insert-tab].is-active")?.textContent.trim() || "",
        types: [...document.querySelectorAll("[data-object-field='classType'] option")].map((option) => option.textContent.trim()),
        displayAsIcon: document.querySelector("[data-object-field='displayAsIcon']")?.checked === true,
      }));
      await page.click('[data-object-insert-tab="file"]');
      await page.locator("[data-object-field='fileName']").fill("sample-report.xlsx");
      await page.locator("[data-object-field='displayAsIcon']").check();
      await page.locator("[data-object-field='iconLabel']").fill("sample-report.xlsx");
      insertObjectDialog = await page.evaluate((initial) => ({
        ...initial,
        activeTabAfterFile: document.querySelector("[data-object-insert-tab].is-active")?.textContent.trim() || "",
        filePanelActive: document.querySelector("[data-object-insert-panel='file']")?.classList.contains("is-active") || false,
        createSelectDisabled: document.querySelector("[data-object-field='classType']")?.disabled === true,
        fileName: document.querySelector("[data-object-field='fileName']")?.value || "",
        displayAsIconAfter: document.querySelector("[data-object-field='displayAsIcon']")?.checked === true,
        iconLabelEnabled: document.querySelector("[data-object-field='iconLabel']")?.disabled === false,
        preview: document.querySelector("[data-object-preview-label]")?.textContent.trim() || "",
      }), initialObjectDialog);
      await page.click('[data-data-dialog-action="object-insert-ok"]');
      await page.waitForSelector('[data-data-dialog="insert-object"]', { state: "detached", timeout: 5000 });
    }
    expectedInsertedObjectCount += 1;
    await page.waitForFunction(
      ({ expected }) => {
        const count = document.querySelectorAll(".sheet-image, .sheet-image-link").length;
        const status = document.querySelector("#statusText")?.textContent || "";
        return count >= expected && (status.includes("挿入") || status.includes("ワードアート") || status.includes("署名欄") || status.includes("オブジェクト") || status.includes("数式") || status.includes("SmartArt") || status.includes("スクリーンショット") || status.includes("スライサー") || status.includes("タイムライン"));
      },
      { expected: expectedInsertedObjectCount },
      { timeout: 5000 },
    );
  }


  const addinObjectStartCount = await page.locator(".sheet-image, .sheet-image-link").count();


  await page.click('[data-ribbon-tab="insert"]');


  await page.click('[data-ribbon-command="insert-get-addins"]');


  await page.waitForSelector('[data-data-dialog="office-addins"]', { timeout: 5000 });


  const insertAddinsStoreDialog = await page.evaluate(() => ({
    title: document.querySelector("#officeAddinsTitle")?.textContent || "",
    activeTab: document.querySelector(".office-addin-tabs .is-active")?.textContent?.trim() || "",
    cards: [...document.querySelectorAll("[data-office-addin-id]")].map((card) => card.dataset.officeAddinId).filter(Boolean),
  }));


  await page.click('[data-office-addin-id="websheet-mini-dashboard"] [data-data-dialog-action="office-addin-insert"]');


  await page.waitForSelector('[data-data-dialog="office-addins"]', { state: "detached", timeout: 5000 });


  await page.waitForFunction((startCount) => document.querySelectorAll(".sheet-image, .sheet-image-link").length === startCount + 1, addinObjectStartCount, { timeout: 5000 });


  await page.click('[data-ribbon-tab="insert"]');


  await page.click('[data-ribbon-command="insert-my-addins"]');


  await page.waitForSelector('[data-data-dialog="office-addins"]', { timeout: 5000 });


  const insertAddinsMineDialog = await page.evaluate(() => ({
    title: document.querySelector("#officeAddinsTitle")?.textContent || "",
    activeTab: document.querySelector(".office-addin-tabs .is-active")?.textContent?.trim() || "",
    cards: [...document.querySelectorAll("[data-office-addin-id]")].map((card) => card.dataset.officeAddinId).filter(Boolean),
  }));


  await page.click('[data-office-addin-id="websheet-bing-maps"] [data-data-dialog-action="office-addin-insert"]');


  await page.waitForFunction((startCount) => document.querySelectorAll(".sheet-image, .sheet-image-link").length === startCount + 2, addinObjectStartCount, { timeout: 5000 });


  const insertTextObjects = await page.evaluate((startCount) => {
    const images = window.__webSheetDevSheetModel?.("CHT")?.images || [];
    const names = images.map((image) => image.name || image.shape?.label || "");
    const byName = (name) => images.find((image) => image.name === name || image.shape?.label === name) || null;
    return {
      countDelta: document.querySelectorAll(".sheet-image, .sheet-image-link").length - startCount,
      names,
      icon: byName("アイコン")?.shape || null,
      model3d: byName("3D モデル")?.shape || null,
      smartArt: byName("SmartArt")?.shape || null,
      screenshot: byName("スクリーンショット")?.shape || null,
      wordArt: byName("ワードアート")?.shape || null,
      signature: byName("署名欄")?.shape || null,
      object: byName("オブジェクト")?.shape || null,
      equation: byName("数式")?.shape || null,
      slicer: byName("スライサー")?.shape || null,
      timeline: byName("タイムライン")?.shape || null,
      miniDashboardAddin: byName("Mini Dashboard")?.shape || null,
      bingMapsAddin: byName("Bing Maps")?.shape || null,
      selectedAlt: document.querySelector(".sheet-image.is-selected, .sheet-image-link.is-selected .sheet-image")?.getAttribute("alt") || "",
      contextualTabVisible: !document.querySelector('[data-ribbon-tab="shape-format"]')?.classList.contains("hidden"),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, textObjectStartCount);


  for (let index = 0; index < insertTextObjectCommands.length + 2; index += 1) {
    await page.keyboard.press("Control+Z");
    await page.waitForTimeout(60);
  }


  await page.waitForFunction((startCount) => document.querySelectorAll(".sheet-image, .sheet-image-link").length === startCount, textObjectStartCount, { timeout: 5000 });



  const standaloneLiveChartUndoCount = await page.evaluate(() => {
    window.__webSheetDevSetActiveSheet?.("CHT");
    const entries = [
      [1, 1, "期"],
      [1, 2, "売上"],
      [2, 1, "上期"],
      [2, 2, 120],
      [3, 1, "下期"],
      [3, 2, 85],
      [4, 1, "通期"],
      [4, 2, 65],
    ];
    entries.forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
    window.__webSheetDevSelectRange?.(1, 1, 4, 2);
    return entries.length;
  });


  await page.waitForFunction(() => String(window.__webSheetDevCellModel?.("CHT", 4, 2)?.raw ?? "") === "65", null, { timeout: 5000 });


  const standaloneLiveChartStartCount = await page.locator(".sheet-image, .sheet-image-link").count();


  await page.click('[data-ribbon-tab="insert"]');


  await page.click('[data-ribbon-command="insert-column-chart"]');


  await page.waitForFunction((startCount) => document.querySelectorAll(".sheet-image, .sheet-image-link").length > startCount, standaloneLiveChartStartCount, { timeout: 5000 });


  const standaloneLiveChartSourceBeforeExport = await page.evaluate(() => {
    const domImage = [...document.querySelectorAll(".sheet-image, .sheet-image-link .sheet-image")]
      .find((image) => decodeURIComponent(image.getAttribute("src") || "").includes("上期"));
    const source = decodeURIComponent(domImage?.getAttribute("src") || "");
    const modelImage = window.__webSheetDevSheetModel?.("CHT")?.images?.find((image) => image.chart?.sourceRange === "A1:B4") || null;
    return {
      hasDomImage: Boolean(domImage),
      sourceIncludesLabel: source.includes("上期"),
      modelSourceRef: modelImage?.chart?.sourceRef || "",
      modelLabelsRef: modelImage?.chart?.labelsRef || "",
      modelFirstValuesRef: modelImage?.chart?.series?.[0]?.valuesRef || "",
    };
  });


  const standaloneLiveChartUndoTotal = standaloneLiveChartUndoCount + 1;



  await page.click('[data-ribbon-tab="insert"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "挿入", null, { timeout: 10000 });


  await page.click('[data-ribbon-tab="file"]');


  await page.waitForFunction(() => !document.querySelector("#fileSideMenu")?.classList.contains("hidden"), null, { timeout: 10000 });


  const fileSideMenu = await page.evaluate(() => {
    const menu = document.querySelector("#fileSideMenu");
    const fileTab = document.querySelector('[data-ribbon-tab="file"]');
    const rect = menu?.getBoundingClientRect();
    const fileRect = fileTab?.getBoundingClientRect();
    return {
      visible: Boolean(menu && !menu.classList.contains("hidden")),
      activeRibbonTab: document.querySelector(".ribbon-tab.is-active")?.textContent,
      insertPanelVisible: Boolean(document.querySelector('[data-ribbon-panel="insert"].is-active')),
      labels: [...document.querySelectorAll("#fileSideMenu [data-ribbon-command]")].map((button) => button.textContent.trim()),
      left: Math.round(rect?.left || 0),
      top: Math.round(rect?.top || 0),
      width: Math.round(rect?.width || 0),
      fileTabLeft: Math.round(fileRect?.left || 0),
      fileTabBottom: Math.round(fileRect?.bottom || 0),
      focusedBefore: document.activeElement?.dataset?.ribbonCommand || "",
      accessKeys: [...document.querySelectorAll("#fileSideMenu [data-ribbon-command]")].map((button) => button.dataset.accessKey || ""),
    };
  });


  await page.keyboard.press("ArrowDown");


  Object.assign(fileSideMenu, await page.evaluate(() => ({
    focusedAfterArrowDown: document.activeElement?.dataset?.ribbonCommand || "",
  })));


  await page.keyboard.press("Escape");


  await page.waitForFunction(() => document.querySelector("#fileSideMenu")?.classList.contains("hidden"), null, { timeout: 10000 });


  Object.assign(fileSideMenu, await page.evaluate(() => ({
    keyboardEscClosed: document.querySelector("#fileSideMenu")?.classList.contains("hidden"),
    activeRibbonTabAfterEsc: document.querySelector(".ribbon-tab.is-active")?.textContent,
    insertPanelVisibleAfterEsc: Boolean(document.querySelector('[data-ribbon-panel="insert"].is-active')),
  })));


  await page.click('[data-ribbon-tab="file"]');


  await page.waitForFunction(() => !document.querySelector("#fileSideMenu")?.classList.contains("hidden"), null, { timeout: 10000 });


  await page.mouse.click(1000, 260);


  await page.waitForFunction(() => document.querySelector("#fileSideMenu")?.classList.contains("hidden"), null, { timeout: 10000 });


  const fileSideMenuClosedByOutsideClick = await page.evaluate(() => ({
    hidden: document.querySelector("#fileSideMenu")?.classList.contains("hidden"),
    activeRibbonTab: document.querySelector(".ribbon-tab.is-active")?.textContent,
    insertPanelVisible: Boolean(document.querySelector('[data-ribbon-panel="insert"].is-active')),
  }));


  await page.click('[data-ribbon-tab="file"]');


  await page.waitForFunction(() => !document.querySelector("#fileSideMenu")?.classList.contains("hidden"), null, { timeout: 10000 });


  const downloadPromise = page.waitForEvent("download");


  await page.click('#fileSideMenu [data-ribbon-command="save"]');


  const download = await downloadPromise;


  const fileSideMenuClosedByCommand = await page.evaluate(() => ({
    hidden: document.querySelector("#fileSideMenu")?.classList.contains("hidden"),
    activeRibbonTab: document.querySelector(".ribbon-tab.is-active")?.textContent,
    insertPanelVisible: Boolean(document.querySelector('[data-ribbon-panel="insert"].is-active')),
  }));


  const quickSavePath = await download.path();


  const quickSaveHtml = await readFile(quickSavePath, "utf8");


  fileSideMenuClosedByCommand.saveFileName = download.suggestedFilename();


  fileSideMenuClosedByCommand.saveIsHtml = looksLikeStandaloneWorkbookHtml(quickSaveHtml);


  fileSideMenuClosedByCommand.saveHasWorkbook = quickSaveHtml.includes("ZJ_スキルシート_20260517.xlsx");


  fileSideMenuClosedByCommand.saveState = await page.evaluate(() => {
    const workbookState = window.__webSheetDevWorkbookState?.() || {};
    return {
      sourceName: workbookState.sourceName || "",
      saveType: workbookState.saveType || "",
      currentSaveType: workbookState.currentSaveType || "",
      isStandalone: workbookState.isStandalone === true,
    };
  });



  await page.click('[data-ribbon-tab="file"]');


  await page.waitForFunction(() => !document.querySelector("#fileSideMenu")?.classList.contains("hidden"), null, { timeout: 10000 });


  await page.click('#fileSideMenu [data-ribbon-command="save-as"]');


  await page.waitForSelector('[data-data-dialog="save-as"]', { timeout: 10000 });


  const standaloneHtmlDownloadPromise = page.waitForEvent("download");


  await page.click('[data-data-dialog-action="save-as-ok"]');


  const standaloneHtmlDownload = await standaloneHtmlDownloadPromise;


  const exportedHtmlPath = await standaloneHtmlDownload.path();


  const exportedHtmlFilePath = path.join(path.dirname(exportedHtmlPath), standaloneHtmlDownload.suggestedFilename());


  await standaloneHtmlDownload.saveAs(exportedHtmlFilePath);


  const exportedHtml = await readFile(exportedHtmlFilePath, "utf8");


  const runtimeDir = standaloneRuntimeCurrentDir();


  const runtimeBaseUrl = standaloneRuntimeCurrentBaseUrl();


  const runtimeJs = await readFile(path.join(runtimeDir, "websheet-standalone.js"), "utf8");


  const runtimeCss = await readFile(path.join(runtimeDir, "websheet-standalone.css"), "utf8");


  const appJs = await readFile(path.join(runtimeDir, "websheet-app.js"), "utf8");


  const runtimeSource = `${exportedHtml}\n${runtimeJs}\n${runtimeCss}\n${appJs}`;


  for (let index = 0; index < standaloneLiveChartUndoTotal; index += 1) {
    await page.keyboard.press("Control+Z");
    await page.waitForTimeout(50);
  }


  await page.waitForFunction((startCount) => document.querySelectorAll(".sheet-image, .sheet-image-link").length === startCount, standaloneLiveChartStartCount, { timeout: 5000 });


  const shortcutSave = {};


  const ctrlSavePromise = page.waitForEvent("download");


  await page.keyboard.press("Control+S");


  const ctrlSaveDownload = await ctrlSavePromise;


  shortcutSave.ctrlFileName = ctrlSaveDownload.suggestedFilename();


  const ctrlSaveHtml = await readFile(await ctrlSaveDownload.path(), "utf8");


  shortcutSave.ctrlHasWorkbook = ctrlSaveHtml.includes("ZJ_スキルシート_20260517.xlsx");


  shortcutSave.ctrlIsHtml = looksLikeStandaloneWorkbookHtml(ctrlSaveHtml);


  await page.locator("#formulaInput").focus();


  const metaSavePromise = page.waitForEvent("download");


  await page.keyboard.press("Meta+S");


  const metaSaveDownload = await metaSavePromise;


  shortcutSave.metaFileName = metaSaveDownload.suggestedFilename();


  const metaSaveHtml = await readFile(await metaSaveDownload.path(), "utf8");


  shortcutSave.metaHasWorkbook = metaSaveHtml.includes("ZJ_スキルシート_20260517.xlsx");


  shortcutSave.metaIsHtml = looksLikeStandaloneWorkbookHtml(metaSaveHtml);


  shortcutSave.formulaFocusedDuringMetaSave = await page.evaluate(() => document.activeElement?.id === "formulaInput");


  const exportedHtmlChecks = {
    hasRuntimeStylesheet: exportedHtml.includes(runtimeBaseUrl) && exportedHtml.includes("websheet-standalone.css"),
    hasRuntimeScripts:
      exportedHtml.includes(runtimeBaseUrl) &&
      exportedHtml.includes("websheet-standalone.js"),
    hasRemoteDownloadRuntime: exportedHtml.includes(`/download/runtime/${STANDALONE_RUNTIME_CHANNEL}`),
    hasNoInlineRuntimeDependencies: !exportedHtml.includes("jQuery v3.7.1") && !exportedHtml.includes("HYPERFORMULA"),
    hasFullAppBootstrap: runtimeJs.includes("__WEBSHEET_STANDALONE_MODEL__") && runtimeJs.includes("websheet-app.js") && runtimeJs.includes("document.body.innerHTML"),
    hasCanvasLayerRule: runtimeSource.includes(".sheet-canvas{") && runtimeSource.includes("pointer-events:none") && runtimeSource.includes("z-index:0"),
    hasTransparentDomRule: runtimeSource.includes(".sheet-cell{") && runtimeSource.includes("background:#0000") && runtimeSource.includes("border-right:1px solid #0000"),
    hasCanvasRuntime: runtimeSource.includes("sheet-canvas") && runtimeSource.includes("getContext") && runtimeSource.includes("fillText"),
    hasCanvasBorderRuntime: runtimeSource.includes("border-right") && runtimeSource.includes("border-bottom") && runtimeSource.includes("strokeStyle"),
    hasOverlayCssRuntime: runtimeSource.includes("--cell-fill-color") && runtimeSource.includes("cell-content"),
    hasFillRuntime: runtimeSource.includes("has-fill"),
    hasAdjacentFillRuntime: runtimeSource.includes("hide-grid-right") && runtimeSource.includes("hide-grid-bottom"),
    hasFillModel: exportedHtml.includes('"hasFill":true'),
    hasHiddenGridlineModel: exportedHtml.includes('"showGridLines":false'),
    hasHiddenGridlineRule:
      runtimeSource.includes(".sheet-grid.hide-gridlines .sheet-cell{") &&
      runtimeSource.includes("border-right-color:#0000") &&
      runtimeSource.includes("border-bottom-color:#0000"),
    hasImages: exportedHtml.includes('"images":[{"id":"image-0-0"') && exportedHtml.includes("data:image/png;base64,"),
    hasImageCss: runtimeSource.includes(".sheet-image,.sheet-image-link{") && runtimeSource.includes("position:absolute") && runtimeSource.includes("z-index:3"),
    hasFontFallback: runtimeSource.includes("MS PGothic") && runtimeSource.includes("Hiragino Sans") && runtimeSource.includes("MS Pゴシック"),
    hasCtrlWheelZoom: runtimeSource.includes("wheel") && runtimeSource.includes("ctrlKey") && runtimeSource.includes("deltaY"),
    hasZoomDialog: runtimeSource.includes("zoom-dialog") && runtimeSource.includes("ズーム"),
    hasPhoneticModel: exportedHtml.includes('"phonetics":[{"start":0,"end":3,"text":"ヒガシジュウジョウ"}'),
    hasPhoneticRuntime: runtimeSource.includes("phonetics") && runtimeSource.includes("showPhonetic"),
    hasCanvasTextBaselineRuntime: runtimeSource.includes("alphabetic"),
    hasCanvasTextLineHeightRuntime: runtimeSource.includes("*1.2") || runtimeSource.includes("1.2"),
    hasAccessKeyUnderline: runtimeSource.includes("menu-access-key") && runtimeSource.includes("text-decoration:underline"),
    hasDefinedNames: exportedHtml.includes('"definedNames"') && runtimeSource.includes("addNamedExpression"),
    hasImeCompositionRuntime: runtimeSource.includes("compositionstart") && runtimeSource.includes("isComposing"),
    hasLiveChartReferences:
      standaloneLiveChartSourceBeforeExport.hasDomImage &&
      standaloneLiveChartSourceBeforeExport.sourceIncludesLabel &&
      exportedHtml.includes('"labelsRef"') &&
      exportedHtml.includes('"valuesRef"') &&
      runtimeSource.includes("labelsRef") &&
      runtimeSource.includes("valuesRef"),
	    hasAdvancedShapeRuntime:
	      runtimeSource.includes("fillPattern") &&
      runtimeSource.includes("lineCompound") &&
      runtimeSource.includes("softEdges") &&
      runtimeSource.includes("reflection") &&
      runtimeSource.includes("customPath") &&
      runtimeSource.includes("custom-point") &&
      runtimeSource.includes("sheet-connection-point") &&
      runtimeSource.includes("object-add-vertex") &&
      runtimeSource.includes("shape-selection-pane") &&
      runtimeSource.includes("object-selection-pane") &&
	      runtimeSource.includes("パターン") &&
	      runtimeSource.includes("光彩"),
	    hasLocalSaveEndpoint: exportedHtml.includes("/__websheet/local-save"),
	    hasNoExternalDependencyScripts: !/<script\s+[^>]*src=["']https?:/i.test(exportedHtml) && !exportedHtml.includes("cdn.jsdelivr.net"),
	  };



  await page.click('[data-ribbon-tab="file"]');


  await page.waitForFunction(() => !document.querySelector("#fileSideMenu")?.classList.contains("hidden"), null, { timeout: 10000 });


  await page.click('#fileSideMenu [data-ribbon-command="save-as"]');


  await page.waitForSelector('[data-data-dialog="save-as"]', { timeout: 10000 });


  const saveAsDialog = await page.evaluate(() => ({
    title: document.querySelector("#saveAsTitle")?.textContent || "",
    fileName: document.querySelector("[data-data-field='saveAsFileName']")?.value || "",
    location: document.querySelector("[data-data-field='saveAsLocation']")?.value || "",
    path: document.querySelector("[data-data-field='saveAsPath']")?.selectedOptions?.[0]?.textContent?.trim() || "",
    pathOptions: [...document.querySelectorAll("[data-data-field='saveAsPath'] option")].map((option) => option.textContent.trim()),
    types: [...document.querySelectorAll("[data-data-field='saveAsType'] option")].map((option) => option.textContent.trim()),
    fileHeaders: [...document.querySelectorAll(".save-as-file-header span")].map((span) => span.textContent.trim()),
    fileRows: [...document.querySelectorAll("[data-save-as-file-row]")].map((row) => ({
      name: row.querySelector("span")?.textContent?.trim() || "",
      type: row.dataset.saveAsFileType || "",
      active: row.classList.contains("is-active"),
    })),
    newFolderButton: document.querySelector('[data-data-dialog-action="save-as-new-folder"]')?.textContent?.trim() || "",
    note: document.querySelector("[data-save-as-note]")?.textContent || "",
  }));


  await page.click('[data-save-as-file-row][data-save-as-file-type="xlsm"]');


  saveAsDialog.xlsmRowSelected = await page.evaluate(() => ({
    fileName: document.querySelector("[data-data-field='saveAsFileName']")?.value || "",
    type: document.querySelector("[data-data-field='saveAsType']")?.value || "",
    activeType: document.querySelector("[data-save-as-file-row].is-active")?.dataset.saveAsFileType || "",
    note: document.querySelector("[data-save-as-note]")?.textContent || "",
  }));


  await page.selectOption("[data-data-field='saveAsPath']", "documents");


  saveAsDialog.documentsPathSelected = await page.evaluate(() => ({
    location: document.querySelector("[data-data-field='saveAsLocation']")?.value || "",
    path: document.querySelector("[data-data-field='saveAsPath']")?.selectedOptions?.[0]?.textContent?.trim() || "",
    activePlace: document.querySelector("[data-save-as-location].is-active span")?.textContent?.trim() || "",
    note: document.querySelector("[data-save-as-note]")?.textContent || "",
  }));


  await page.selectOption("[data-data-field='saveAsType']", "xlsx");


  saveAsDialog.xlsxFileName = await page.locator("[data-data-field='saveAsFileName']").inputValue();


  const saveAsXlsxPromise = page.waitForEvent("download");


  await page.click('[data-data-dialog-action="save-as-ok"]');


  const saveAsXlsxDownload = await saveAsXlsxPromise;


  const saveAsXlsxPath = await saveAsXlsxDownload.path();


  const saveAsXlsxBytes = await readFile(saveAsXlsxPath);


  saveAsDialog.xlsxSuggestedFileName = saveAsXlsxDownload.suggestedFilename();


  saveAsDialog.xlsxZipHeader = saveAsXlsxBytes.subarray(0, 4).toString("hex");


  await page.evaluate(() => {
    window.__webSheetSaveHandleWrites = [];
    window.__webSheetSavePickerCalls = [];
    Object.defineProperty(window, "showSaveFilePicker", {
      configurable: true,
      value: async (options = {}) => {
        window.__webSheetSavePickerCalls.push({
          suggestedName: String(options.suggestedName || ""),
          typeCount: Array.isArray(options.types) ? options.types.length : 0,
        });
        const suggestedName = String(options.suggestedName || "picked.xlsx").replace(/\.[^.]+$/, "");
        return {
          name: `${suggestedName}-picked.xlsx`,
          async createWritable() {
            return {
              async write(chunk) {
                window.__webSheetSaveHandleWrites.push({
                  size: Number(chunk?.size ?? chunk?.byteLength ?? 0),
                  type: String(chunk?.type || ""),
                });
              },
              async close() {
                window.__webSheetSaveHandleWrites.push({ closed: true });
              },
            };
          },
        };
      },
    });
  });
  Object.assign(scenarioState, { context, page, result, viewRibbon, viewFullScreenOn, fullScreenUndoRedo, viewFullScreenOff, sheetViewNew, sheetViewKept, sheetViewMenuBeforeRename, sheetViewRenamed, sheetViewExited, sheetViewReactivated, sheetViewUpdated, sheetViewMenuBeforeCopy, sheetViewDuplicated, viewZoomDialog, viewZoomApplied, viewZoomHundred, viewZoomSelection, viewZoomSelectionUndoRedo, viewRulerOn, pageLayoutHeaderFooterDialog, pageBreakPreview, pageBreakContextMenu, customViewCaptureState, customViewsMainBeforeAdd, customViewAddDialog, customViewsDialog, customViewSaveUndoRedo, customViewApplied, customViewApplyUndoRedo, customViewDeleteFlow, freezeMenuLabels, freezeMenuInitialMeta, freezeMenuTopRow, freezeMenuTopRowStatus, freezeMenuAccessKey, formulaBarUndoRedo, viewToggles, rtlVisualHitTest, gridlinesUndoRedo, headingsUndoRedo, zeroValuesUndoRedo, objectsUndoRedo, outlineUndoRedo, rightToLeftUndoRedo, freezePanesAfterMain, frozenPaneSticky, splitPaneBeforeDrag, splitPaneDragged, sheetGroupView, viewInitialWindowBar, viewNewWindowState, viewNewWindowBar, viewSecondNewWindowState, viewSecondNewWindowBar, workbookWindowBarSwitch, customWindowSource, customWindowMutated, customWindowRestored, arrangeWindowDialog, arrangeWindowState, arrangeWindowBar, arrangeWindowUndo, arrangeWindowRedo, sideBySideWindowDialog, sideBySideDialogApplied, sideBySideWindowBar, syncScrollState, syncScrollUndoRedo, resetWindowState, resetWindowBar, switchWindowMenu, switchWindowKeyboardState, switchWindowState, switchWindowDialog, switchWindowDialogState, hideWindowState, hideWindowBar, unhideWindowDialog, unhideWindowState, unhideWindowBar, unhideWindowUndoRedo, closeWindowState, closeWindowUndoRedo, macrosDialog, macroMenuBeforeRecord, macroRecordDialog, macroRecordingStarted, macroMenuDuringRecord, macroMenuRelativeRefs, macroRelativeRefsUndoRedo, macroRecordingStopped, macroRecordingStopUndoRedo, macroSecurityDialog, macroSecurityApplied, macroSecurityUndoRedo, viewWindowCommands, layoutRibbon, layoutThemeMenuInitial, layoutThemeApplied, layoutColorMenuInitial, layoutFontMenuInitial, layoutEffectMenuInitial, layoutThemeAfterMenus, layoutThemeEffectUndoRedo, layoutPrintAreaMenuBefore, layoutPrintAreaMenuAfterSet, layoutPrintAreaMenuAfterClear, layoutMarginsMenuInitial, layoutOrientationMenuInitial, layoutSizeMenuInitial, layoutBackgroundApplied, layoutBackgroundUndo, layoutBackgroundRedo, layoutBackgroundDeleted, layoutBackgroundDeleteUndo, pageSetupDialog, printTitlesDialogInitial, printTitlesDialogFilled, pageLayout, pageLayoutPreview, pageBreakMenuKeyboard, pageBreakDrag, drawRibbon, drawInk, drawConverted, headerFooterDialog, headerFooterApplied, headerFooterUndo, headerFooterRedo, insertMapChart, insertShapeMenu, insertShapeDraw, insertIconDialog, signatureLineDialog, insertObjectDialog, insertAddinsStoreDialog, insertAddinsMineDialog, insertTextObjects, fileSideMenu, fileSideMenuClosedByOutsideClick, fileSideMenuClosedByCommand, exportedHtmlFilePath, exportedHtml, runtimeBaseUrl, runtimeSource, shortcutSave, exportedHtmlChecks, saveAsDialog });
}

export {
  runWindowsAndSelection,
};
