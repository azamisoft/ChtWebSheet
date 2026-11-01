// Sequential E2E feature module.
import {
  appUrl,
  browser,
  workbookPath,
} from "../../e2e-real-upload.mjs";
import {
  clickRibbonTab,
} from "../helpers/fixtures-and-model.mjs";
async function runSetupAndWorkbook(scenarioState) {

  const context = await browser.newContext({
    viewport: { width: 1280, height: 720 },
    acceptDownloads: true,
    permissions: ["clipboard-read", "clipboard-write"],
  });


  const page = await context.newPage();


  const messages = [];


  const browserDialogs = [];



  page.on("console", (message) => {
    messages.push(`${message.type()}: ${message.text()}`);
  });


  page.on("pageerror", (error) => {
    messages.push(`pageerror: ${error.stack || error.message}`);
  });


  page.on("dialog", async (dialog) => {
    browserDialogs.push({ type: dialog.type(), message: dialog.message() });
    await dialog.dismiss();
  });



  await page.goto(`${appUrl}?t=${Date.now()}`, { waitUntil: "load" });


  await page.setInputFiles("#fileInput", workbookPath);


  await page.waitForFunction(
    () => {
      const status = document.querySelector("#statusText")?.textContent || "";
      const alert = document.querySelector("#appAlert")?.textContent || "";
      return /読み込みました|読み込み済み|已加载/.test(status) || /読み込みに失敗|读取失败/.test(alert);
    },
    null,
    { timeout: 20000 },
  );



  const result = await page.evaluate(() => ({
    hasTopbar: Boolean(document.querySelector(".topbar")),
    hasHeaderExport: Boolean(document.querySelector("#exportHtmlButton")),
    hasHeaderToggleFormulas: Boolean(document.querySelector("#toggleFormulasButton")),
    hasSheetTabScrollLeft: Boolean(document.querySelector("#sheetTabScrollLeft")),
    hasSheetTabScrollRight: Boolean(document.querySelector("#sheetTabScrollRight")),
    exportDisabled: document.querySelector("#exportHtmlButton")?.disabled,
    formulaDisabled: document.querySelector("#toggleFormulasButton")?.disabled,
    activeRibbonTab: document.querySelector(".ribbon-tab.is-active")?.textContent,
    homePanelVisible: Boolean(document.querySelector('[data-ribbon-panel="home"].is-active')),
    ribbonHomeLayout: (() => {
      const panel = document.querySelector('[data-ribbon-panel="home"]');
      const rect = (selector) => {
        const element = panel?.querySelector(selector);
        const box = element?.getBoundingClientRect();
        return box ? { width: Math.round(box.width), height: Math.round(box.height) } : null;
      };
      const pageRect = (selector) => {
        const element = document.querySelector(selector);
        const box = element?.getBoundingClientRect();
        return box ? { width: Math.round(box.width), height: Math.round(box.height) } : null;
      };
      return {
        ribbon: pageRect(".ribbon"),
        ribbonTabRow: pageRect(".ribbon-tab-row"),
        ribbonPanels: {
          width: Math.round(document.querySelector(".ribbon-panels")?.getBoundingClientRect().width || 0),
          height: Math.round(document.querySelector(".ribbon-panels")?.getBoundingClientRect().height || 0),
        },
        groupLabels: [...(panel?.querySelectorAll(".ribbon-group-label") || [])].map((label) => label.textContent),
        groupCount: panel?.querySelectorAll(".ribbon-group").length || 0,
        clipboardStackCount: panel?.querySelectorAll(".ribbon-clipboard-group .ribbon-stack .ribbon-tool").length || 0,
        pasteButton: rect(".ribbon-clipboard-group .ribbon-large-tool"),
        cutButton: rect('[data-ribbon-command="cut"]'),
        fontField: rect(".ribbon-font-field"),
        sizeField: rect(".ribbon-size-field"),
        groupWidths: {
          clipboard: rect(".ribbon-clipboard-group")?.width,
          font: rect(".ribbon-font-group")?.width,
          alignment: rect(".ribbon-alignment-group")?.width,
          number: rect(".ribbon-number-group")?.width,
          cells: rect(".ribbon-cells-group")?.width,
          editing: rect(".ribbon-editing-group")?.width,
        },
        fontRows: panel?.querySelectorAll(".ribbon-font-group .ribbon-row").length || 0,
        alignmentRows: panel?.querySelectorAll(".ribbon-alignment-group .ribbon-row").length || 0,
        alignmentIconCount: panel?.querySelectorAll(".ribbon-align-grid .ribbon-icon-tool").length || 0,
        alignmentActionCommands: [...(panel?.querySelectorAll(".ribbon-alignment-actions [data-ribbon-command]") || [])].map((button) => button.dataset.ribbonCommand || ""),
        mergeInsideAlignment: Boolean(panel?.querySelector('.ribbon-alignment-group [data-ribbon-command="merge-center"]')),
        mergeIconClass: panel?.querySelector('[data-ribbon-command="merge-center"] svg')?.getAttribute("class") || "",
        wrapInsideAlignment: Boolean(panel?.querySelector('.ribbon-alignment-group [data-ribbon-command="wrap-text"]')),
        alignJustifyVisible: Boolean(panel?.querySelector('.ribbon-alignment-group [data-ribbon-command="align-justify"]')),
        numberFormatVisible: Boolean(panel?.querySelector(".number-format-select")),
        compactFontFamily: getComputedStyle(document.querySelector(".ribbon")).fontFamily,
        tabFontSize: getComputedStyle(document.querySelector(".ribbon-tab")).fontSize,
        toolFontSize: getComputedStyle(panel?.querySelector(".ribbon-tool")).fontSize,
        selectFontSize: getComputedStyle(panel?.querySelector(".ribbon-select")).fontSize,
        borderSelectExists: Boolean(panel?.querySelector('[data-ribbon-style="border"]')),
        borderSplitVisible: Boolean(panel?.querySelector(".ribbon-border-split")),
        underlineSplitVisible: Boolean(panel?.querySelector(".ribbon-underline-split")),
        borderMark: panel?.querySelector("[data-ribbon-border-mark]")?.dataset.ribbonBorderMark,
        cellsCommands: [...(panel?.querySelectorAll(".ribbon-cells-group [data-ribbon-command]") || [])].map((button) => button.textContent.trim()),
        editingCommands: [...(panel?.querySelectorAll(".ribbon-editing-group [data-ribbon-command]") || [])].map((button) => button.textContent.trim()),
      };
    })(),
    fileSideMenuHidden: document.querySelector("#fileSideMenu")?.classList.contains("hidden"),
    fileRibbonPanelExists: Boolean(document.querySelector('[data-ribbon-panel="file"]')),
    status: document.querySelector("#statusText")?.textContent,
    workbookName: document.querySelector("#workbookName")?.textContent,
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    alert: document.querySelector("#appAlert")?.textContent,
    memoryStats: window.__webSheetMemoryStats?.() || null,
    modelIntegrity: (() => {
      const richCell = window.__webSheetDevCellModel?.("CHT", 51, 6);
      const formulaCell = window.__webSheetDevCellModel?.("CHT", 52, 3);
      return {
        richRuns: richCell?.richText?.length || 0,
        richRunFonts: (richCell?.richText || []).map((run) => run.css?.fontFamily || ""),
        datedifFormula: formulaCell?.formula || "",
        datedifCached: formulaCell?.cached,
        datedifDisplay: window.__webSheetDevCellDisplay?.("CHT", 52, 3),
      };
    })(),
    rows: document.querySelectorAll("#sheetHost .row-header").length,
    cells: document.querySelectorAll("#sheetHost .sheet-cell").length,
    canvasLayer: (() => {
      const canvas = document.querySelector("#sheetHost .sheet-canvas");
      const style = canvas ? getComputedStyle(canvas) : null;
      return {
        exists: Boolean(canvas),
        width: canvas?.width || 0,
        height: canvas?.height || 0,
        position: style?.position,
        pointerEvents: style?.pointerEvents,
      };
    })(),
    canvasTextCentering: (() => {
      const canvas = document.querySelector("#sheetHost .sheet-canvas");
      const context = canvas?.getContext("2d");
      const canvasRect = canvas?.getBoundingClientRect();
      if (!canvas || !context || !canvasRect) return null;
      const dpr = Number(canvas.dataset.dpr || 1);
      const inkBounds = (address) => {
        const cell = document.querySelector(`.sheet-cell[data-address="${address}"]`);
        if (!cell) return null;
        const rect = cell.getBoundingClientRect();
        const x0 = Math.max(0, Math.floor((rect.left - canvasRect.left + 3) * dpr));
        const x1 = Math.min(canvas.width, Math.ceil((rect.right - canvasRect.left - 3) * dpr));
        const y0 = Math.max(0, Math.floor((rect.top - canvasRect.top + 2) * dpr));
        const y1 = Math.min(canvas.height, Math.ceil((rect.bottom - canvasRect.top - 2) * dpr));
        const image = context.getImageData(x0, y0, Math.max(1, x1 - x0), Math.max(1, y1 - y0));
        let minY = Infinity;
        let maxY = -Infinity;
        let count = 0;
        for (let y = 0; y < image.height; y += 1) {
          for (let x = 0; x < image.width; x += 1) {
            const index = (y * image.width + x) * 4;
            const r = image.data[index];
            const g = image.data[index + 1];
            const b = image.data[index + 2];
            const a = image.data[index + 3];
            if (a > 100 && r < 80 && g < 80 && b < 80) {
              minY = Math.min(minY, y);
              maxY = Math.max(maxY, y);
              count += 1;
            }
          }
        }
        if (!count) return { address, count, delta: Infinity };
        const inkTop = minY / dpr + 2;
        const inkBottom = maxY / dpr + 2;
        return {
          address,
          count,
          cellHeight: rect.height,
          delta: (inkTop + inkBottom) / 2 - rect.height / 2,
        };
      };
      return [inkBounds("B4"), inkBounds("B5")];
    })(),
    verticalText: (() => {
      const cell = document.querySelector('.sheet-cell[data-address="B15"]');
      return {
        text: cell?.textContent?.trim() || "",
        styleAttr: cell?.getAttribute("style") || "",
        hasTextRotation: (cell?.getAttribute("style") || "").includes("text-rotation:vertical"),
        gridRow: cell?.style.gridRow || "",
      };
    })(),
    footerTabs: [...document.querySelectorAll("#sheetFooter .sheet-tab")].map((button) => button.textContent),
    addButton: document.querySelector("#addSheetButton")?.textContent,
    scrollMax: (document.querySelector("#sheetHost")?.scrollWidth || 0) - (document.querySelector("#sheetHost")?.clientWidth || 0),
    b14: (() => {
      const cell = document.querySelector('.sheet-cell[data-address="B14"]');
      const content = cell?.querySelector(".cell-content");
      const header = [...document.querySelectorAll(".column-header")].find((node) => node.textContent.trim() === "B");
      const rowHeader = [...document.querySelectorAll(".row-header")].find((node) => node.textContent.trim() === "14");
      return {
        cellWidth: cell?.getBoundingClientRect().width,
        cellHeight: cell?.getBoundingClientRect().height,
        columnHeaderWidth: header?.getBoundingClientRect().width,
        rowHeaderHeight: rowHeader?.getBoundingClientRect().height,
        contentWidth: content?.getBoundingClientRect().width,
        contentScrollWidth: content?.scrollWidth,
        text: content?.textContent,
      };
    })(),
    g11: (() => {
      const content = document.querySelector('.sheet-cell[data-address="G11"] .cell-content');
      const style = content ? getComputedStyle(content) : null;
      return {
        text: content?.textContent,
        whiteSpace: style?.whiteSpace,
      };
    })(),
    fontSamples: (() => {
      const b14 = document.querySelector('.sheet-cell[data-address="B14"]');
      const b14Style = b14 ? getComputedStyle(b14) : null;
      return {
        b14: b14Style?.fontFamily,
        fontOptions: [...document.querySelectorAll(".font-family-select option")].map((option) => option.value),
      };
    })(),
    a1Contenteditable: document.querySelector('.sheet-cell[data-address="A1"]')?.getAttribute("contenteditable"),
    gridEdge: (() => {
      const host = document.querySelector("#sheetHost");
      const grid = document.querySelector(".sheet-grid");
      return {
        hidesGridlines: grid?.classList.contains("hide-gridlines"),
        hostTop: Math.round(host?.getBoundingClientRect().top || 0),
        gridTop: Math.round(grid?.getBoundingClientRect().top || 0),
        hostLeft: Math.round(host?.getBoundingClientRect().left || 0),
        gridLeft: Math.round(grid?.getBoundingClientRect().left || 0),
      };
    })(),
    a1Gridline: (() => {
      const cell = document.querySelector('.sheet-cell[data-address="A1"]');
      const style = cell ? getComputedStyle(cell) : null;
      return {
        hasFillClass: cell?.classList.contains("has-fill"),
        fillColor: style?.getPropertyValue("--cell-fill-color").trim().toLowerCase(),
        backgroundColor: style?.backgroundColor,
        borderRightWidth: style?.borderRightWidth,
        borderBottomWidth: style?.borderBottomWidth,
        borderRightColor: style?.borderRightColor,
        borderBottomColor: style?.borderBottomColor,
      };
    })(),
    whiteFill: (() => {
      const cell = document.querySelector('.sheet-cell[data-address="G10"]');
      const style = cell ? getComputedStyle(cell) : null;
      return {
        hasFillClass: cell?.classList.contains("has-fill"),
        fillColor: style?.getPropertyValue("--cell-fill-color").trim().toLowerCase(),
        backgroundColor: style?.backgroundColor,
        borderRightColor: style?.borderRightColor,
        borderBottomColor: style?.borderBottomColor,
      };
    })(),
  }));



  await clickRibbonTab(page, "data");


  const dataRibbonLayout = await page.evaluate(() => {
    const panel = document.querySelector('[data-ribbon-panel="data"]');
    const sortButtons = [...(panel?.querySelectorAll(".ribbon-data-sort-filter-group [data-ribbon-command]") || [])].map((button) => {
      const rect = button.getBoundingClientRect();
      return {
        command: button.dataset.ribbonCommand,
        text: button.textContent.trim(),
        left: Math.round(rect.left),
        top: Math.round(rect.top),
        right: Math.round(rect.right),
        bottom: Math.round(rect.bottom),
        width: Math.round(rect.width),
        height: Math.round(rect.height),
      };
    });
    const overlaps = [];
    for (let leftIndex = 0; leftIndex < sortButtons.length; leftIndex += 1) {
      for (let rightIndex = leftIndex + 1; rightIndex < sortButtons.length; rightIndex += 1) {
        const a = sortButtons[leftIndex];
        const b = sortButtons[rightIndex];
        const intersects = a.left < b.right - 1 && a.right > b.left + 1 && a.top < b.bottom - 1 && a.bottom > b.top + 1;
        if (intersects) overlaps.push(`${a.command}/${b.command}`);
      }
    }
    return {
      active: document.querySelector(".ribbon-tab.is-active")?.textContent?.trim() || "",
      panelHeight: Math.round(panel?.getBoundingClientRect().height || 0),
      groupLabels: [...(panel?.querySelectorAll(".ribbon-group-label") || [])].map((label) => label.textContent.trim()),
      sortButtons,
      sortOverlaps: overlaps,
      panelOverflow: Math.round((panel?.scrollWidth || 0) - (panel?.clientWidth || 0)),
      groupWidths: Object.fromEntries(
        [...(panel?.querySelectorAll(".ribbon-group") || [])].map((group) => {
          const label = group.querySelector(".ribbon-group-label")?.textContent.trim() || "";
          return [label, Math.round(group.getBoundingClientRect().width)];
        }),
      ),
    };
  });


  await clickRibbonTab(page, "home");



  const coordinateHitTest = await page.evaluate(() => {
    const style = document.createElement("style");
    style.id = "coordinate-hit-test-style";
    style.textContent = '#sheetHost .sheet-cell:not([contenteditable="true"]) { pointer-events: none !important; }';
    document.head.append(style);
    const rect = document.querySelector('.sheet-cell[data-address="G6"]')?.getBoundingClientRect();
    return {
      x: Math.round((rect?.left || 0) + (rect?.width || 0) / 2),
      y: Math.round((rect?.top || 0) + (rect?.height || 0) / 2),
    };
  });


  await page.mouse.click(coordinateHitTest.x, coordinateHitTest.y);


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "G6", null, { timeout: 10000 });


  Object.assign(coordinateHitTest, await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    activeAddress: document.querySelector(".sheet-cell.is-selected")?.dataset.address,
    eventTargetDisabled: getComputedStyle(document.querySelector('.sheet-cell[data-address="G6"]')).pointerEvents,
  })));


  await page.evaluate(() => document.querySelector("#coordinate-hit-test-style")?.remove());



  const delayedLazyLoad = await page.evaluate(async () => {
    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const host = document.querySelector("#sheetHost");
    const before = {
      scrollHeight: host?.scrollHeight || 0,
      scrollWidth: host?.scrollWidth || 0,
      clientHeight: host?.clientHeight || 0,
      clientWidth: host?.clientWidth || 0,
    };
    const offsetFromEdge = 260;
    host.scrollTop = Math.max(0, host.scrollHeight - host.clientHeight - offsetFromEdge);
    host.scrollLeft = Math.max(0, host.scrollWidth - host.clientWidth - offsetFromEdge);
    host.dispatchEvent(new Event("scroll"));
    await frame();
    const after = {
      scrollHeight: host.scrollHeight,
      scrollWidth: host.scrollWidth,
      rightRemaining: host.scrollWidth - host.clientWidth - host.scrollLeft,
      bottomRemaining: host.scrollHeight - host.clientHeight - host.scrollTop,
    };
    host.scrollTop = 0;
    host.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll"));
    await frame();
    return {
      before,
      after,
      resetTop: host.scrollTop,
      resetLeft: host.scrollLeft,
    };
  });



  const stickyHeaders = await page.evaluate(async () => {
    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const host = document.querySelector("#sheetHost");
    host.scrollTop = Math.min(420, Math.max(0, host.scrollHeight - host.clientHeight));
    host.scrollLeft = Math.min(420, Math.max(0, host.scrollWidth - host.clientWidth));
    host.dispatchEvent(new Event("scroll"));
    await frame();
    const hostRect = host.getBoundingClientRect();
    const columnHeader = document.querySelector(".column-header")?.getBoundingClientRect();
    const rowHeader = document.querySelector(".row-header")?.getBoundingClientRect();
    const cornerHeader = document.querySelector(".corner-header")?.getBoundingClientRect();
    const result = {
      hostTop: Math.round(hostRect.top),
      hostLeft: Math.round(hostRect.left),
      columnTop: Math.round(columnHeader?.top || 0),
      rowLeft: Math.round(rowHeader?.left || 0),
      cornerTop: Math.round(cornerHeader?.top || 0),
      cornerLeft: Math.round(cornerHeader?.left || 0),
      rowHeaderWidth: Math.round(rowHeader?.width || 0),
    };
    host.scrollTop = 0;
    host.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll"));
    await frame();
    return result;
  });



  await page.click('.sheet-tab[data-sheet-index="0"]', { button: "right" });


  await page.waitForSelector("#sheetTabContextMenu");


  const sheetMenu = await page.evaluate(() => ({
    labels: [...document.querySelectorAll("#sheetTabContextMenu .sheet-tab-menu-item")].map((item) =>
      item.textContent.replace(/\s+/g, "").trim(),
    ),
    accessKeys: [...document.querySelectorAll("#sheetTabContextMenu .menu-access-key")].map((item) => item.textContent),
    accessKeysUnderlined: [...document.querySelectorAll("#sheetTabContextMenu .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
    focusedBefore: document.activeElement?.dataset?.sheetMenuAction || "",
    separatorCount: document.querySelectorAll("#sheetTabContextMenu .sheet-tab-menu-separator").length,
    hasColorChoices: document.querySelectorAll("#sheetTabContextMenu [data-sheet-tab-color]").length >= 8,
    rect: (() => {
      const rect = document.querySelector("#sheetTabContextMenu")?.getBoundingClientRect();
      return {
        left: Math.round(rect?.left || 0),
        top: Math.round(rect?.top || 0),
        right: Math.round(rect?.right || 0),
        bottom: Math.round(rect?.bottom || 0),
      };
    })(),
  }));


  await page.keyboard.press("ArrowDown");


  Object.assign(sheetMenu, await page.evaluate(() => ({
    focusedAfterDown: document.activeElement?.dataset?.sheetMenuAction || "",
  })));


  await page.hover("#sheetTabContextMenu .sheet-tab-menu-submenu");


  await page.click('#sheetTabContextMenu [data-sheet-tab-color="#ffc000"]');


  await page.waitForFunction(() => document.querySelector('.sheet-tab[data-sheet-index="0"]')?.dataset.tabColor === "true", null, { timeout: 10000 });


  const sheetTabColor = await page.evaluate(() => {
    const tab = document.querySelector('.sheet-tab[data-sheet-index="0"]');
    const style = tab ? getComputedStyle(tab) : null;
    return {
      hasTabColor: tab?.dataset.tabColor === "true",
      borderTopColor: style?.borderTopColor,
      activeUnderline: style?.boxShadow,
    };
  });


  await page.click('.sheet-tab[data-sheet-index="0"]', { button: "right" });


  await page.keyboard.press("r");


  await page.waitForSelector(".sheet-tab-rename-input", { timeout: 5000 });


  await page.waitForFunction(() => document.activeElement === document.querySelector(".sheet-tab-rename-input"), null, { timeout: 5000 });


  const sheetRenameInline = await page.evaluate(() => {
    const input = document.querySelector(".sheet-tab-rename-input");
    const tab = document.querySelector('.sheet-tab[data-sheet-index="0"]');
    return {
      value: input?.value || "",
      active: document.activeElement === input,
      selected: Boolean(input && input.selectionStart === 0 && input.selectionEnd === input.value.length),
      tabRenaming: Boolean(tab?.classList.contains("is-renaming")),
      dialogOpen: Boolean(document.querySelector('[data-data-dialog="sheet-rename"]')),
    };
  });


  await page.locator(".sheet-tab-rename-input").fill("CHT Rename Test");


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => document.querySelector('.sheet-tab[data-sheet-index="0"]')?.textContent?.trim() === "CHT Rename Test", null, { timeout: 5000 });


  const sheetRenameApplied = await page.evaluate(() => ({
    tab: document.querySelector('.sheet-tab[data-sheet-index="0"]')?.textContent?.trim() || "",
    sheetName: window.__webSheetDevWorkbookState?.().activeSheetName || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => document.querySelector('.sheet-tab[data-sheet-index="0"]')?.textContent?.trim() === "CHT", null, { timeout: 5000 });


  const sheetRenameUndo = await page.evaluate(() => ({
    tab: document.querySelector('.sheet-tab[data-sheet-index="0"]')?.textContent?.trim() || "",
    sheetName: window.__webSheetDevWorkbookState?.().activeSheetName || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => document.querySelector('.sheet-tab[data-sheet-index="0"]')?.textContent?.trim() === "CHT Rename Test", null, { timeout: 5000 });


  const sheetRenameRedo = await page.evaluate(() => ({
    tab: document.querySelector('.sheet-tab[data-sheet-index="0"]')?.textContent?.trim() || "",
    sheetName: window.__webSheetDevWorkbookState?.().activeSheetName || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => document.querySelector('.sheet-tab[data-sheet-index="0"]')?.textContent?.trim() === "CHT", null, { timeout: 5000 });


  const sheetRename = {
    inline: sheetRenameInline,
    applied: sheetRenameApplied,
    undo: sheetRenameUndo,
    redo: sheetRenameRedo,
  };


  await page.dblclick('.sheet-tab[data-sheet-index="0"]');


  await page.waitForSelector(".sheet-tab-rename-input", { timeout: 5000 });


  await page.waitForFunction(() => document.activeElement === document.querySelector(".sheet-tab-rename-input"), null, { timeout: 5000 });


  const sheetDoubleClickRenameInline = await page.evaluate(() => {
    const input = document.querySelector(".sheet-tab-rename-input");
    return {
      value: input?.value || "",
      active: document.activeElement === input,
      selected: Boolean(input && input.selectionStart === 0 && input.selectionEnd === input.value.length),
      tabRenaming: Boolean(document.querySelector('.sheet-tab[data-sheet-index="0"]')?.classList.contains("is-renaming")),
    };
  });


  await page.locator(".sheet-tab-rename-input").fill("CHT Double Rename");


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => document.querySelector('.sheet-tab[data-sheet-index="0"]')?.textContent?.trim() === "CHT Double Rename", null, { timeout: 5000 });


  const sheetDoubleClickRenameApplied = await page.evaluate(() => ({
    tab: document.querySelector('.sheet-tab[data-sheet-index="0"]')?.textContent?.trim() || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => document.querySelector('.sheet-tab[data-sheet-index="0"]')?.textContent?.trim() === "CHT", null, { timeout: 5000 });


  const sheetDoubleClickRenameUndo = await page.evaluate(() => ({
    tab: document.querySelector('.sheet-tab[data-sheet-index="0"]')?.textContent?.trim() || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.dblclick('.sheet-tab[data-sheet-index="0"]');


  await page.waitForSelector(".sheet-tab-rename-input", { timeout: 5000 });


  await page.waitForFunction(() => document.activeElement === document.querySelector(".sheet-tab-rename-input"), null, { timeout: 5000 });


  await page.locator(".sheet-tab-rename-input").fill("CHT Cancel Rename");


  await page.keyboard.press("Escape");


  await page.waitForFunction(() => !document.querySelector(".sheet-tab-rename-input"), null, { timeout: 5000 });


  const sheetDoubleClickRenameCancel = await page.evaluate(() => ({
    tab: document.querySelector('.sheet-tab[data-sheet-index="0"]')?.textContent?.trim() || "",
    renaming: Boolean(document.querySelector('.sheet-tab[data-sheet-index="0"]')?.classList.contains("is-renaming")),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  const dragSheetTab = async ({ copy = false, sourcePosition = 0, drop = "end" } = {}) => {
    const points = await page.evaluate(() => {
      const tabs = [...document.querySelectorAll(".sheet-tab")];
      const rects = tabs.map((tab) => {
        const rect = tab.getBoundingClientRect();
        return {
          left: Math.round(rect.left),
          right: Math.round(rect.right),
          top: Math.round(rect.top),
          width: Math.round(rect.width),
          height: Math.round(rect.height),
        };
      });
      return rects;
    });
    const source = points[Math.max(0, Math.min(sourcePosition, points.length - 1))];
    const next = points[Math.min(points.length - 1, Math.max(0, sourcePosition + 1))] || source;
    const first = points[0] || source;
    const last = points[points.length - 1] || source;
    const dragPoints = {
      startX: Math.round(source.left + source.width / 2),
      startY: Math.round(source.top + source.height / 2),
      midX: Math.round(next.left + next.width / 2),
      endX: drop === "start" ? Math.round(first.left - 12) : Math.round(last.right + 12),
      endY: Math.round((drop === "start" ? first : last).top + (drop === "start" ? first : last).height / 2),
    };
    if (copy) await page.keyboard.down("Control");
    await page.mouse.move(dragPoints.startX, dragPoints.startY);
    await page.mouse.down();
    await page.mouse.move(dragPoints.midX, dragPoints.startY, { steps: 4 });
    await page.mouse.move(dragPoints.endX, dragPoints.endY, { steps: 4 });
    const visual = await page.evaluate(() => ({
      hasIndicator: Boolean(document.querySelector(".sheet-tab-drop-indicator")),
      copyIndicator: Boolean(document.querySelector(".sheet-tab-drop-indicator.is-copy")),
      dragSourceCount: document.querySelectorAll(".sheet-tab.is-drag-source").length,
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.mouse.up();
    if (copy) await page.keyboard.up("Control");
    return visual;
  };


  const dragSheetTabToEnd = (options = {}) => dragSheetTab({ ...options, drop: "end" });


  const sheetTabDragMoveVisual = await dragSheetTabToEnd();


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "AI製品紹介|CHT",
    null,
    { timeout: 5000 },
  );


  const sheetTabDragMoved = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "CHT|AI製品紹介",
    null,
    { timeout: 5000 },
  );


  const sheetTabDragMoveUndo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  const sheetTabDragCopyVisual = await dragSheetTabToEnd({ copy: true });


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "CHT|AI製品紹介|CHT (2)",
    null,
    { timeout: 5000 },
  );


  const sheetTabDragCopied = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "CHT|AI製品紹介",
    null,
    { timeout: 5000 },
  );


  const sheetTabDirectActions = {
    doubleClickRename: {
      inline: sheetDoubleClickRenameInline,
      applied: sheetDoubleClickRenameApplied,
      undo: sheetDoubleClickRenameUndo,
      cancel: sheetDoubleClickRenameCancel,
    },
    dragMove: {
      visual: sheetTabDragMoveVisual,
      applied: sheetTabDragMoved,
      undo: sheetTabDragMoveUndo,
    },
    dragCopy: {
      visual: sheetTabDragCopyVisual,
      applied: sheetTabDragCopied,
    },
  };


  await page.click('.sheet-tab[data-sheet-index="0"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="move-copy"]');


  await page.waitForSelector('[data-data-dialog="sheet-move-copy"]', { timeout: 5000 });


  const sheetMoveDialog = await page.evaluate(() => ({
    title: document.querySelector("#sheetMoveCopyTitle")?.textContent || "",
    workbookCaption: [...document.querySelectorAll('[data-data-dialog="sheet-move-copy"] .excel-dialog-caption')][0]?.textContent || "",
    beforeCaption: [...document.querySelectorAll('[data-data-dialog="sheet-move-copy"] .excel-dialog-caption')][1]?.textContent || "",
    options: [...document.querySelectorAll('[data-data-field="sheetMoveBeforeIndex"] option')].map((option) => option.textContent.trim()),
    activeField: document.activeElement?.dataset?.dataField || "",
    hasCopyCheckbox: Boolean(document.querySelector('[data-data-field="sheetMoveCreateCopy"]')),
    isExcelDialog: Boolean(document.querySelector(".excel-mini-dialog.sheet-move-copy-dialog")),
  }));


  await page.locator('[data-data-field="sheetMoveBeforeIndex"]').selectOption("__end");


  await page.locator('[data-data-dialog="sheet-move-copy"] [data-data-dialog-action="ok"]').click();


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "AI製品紹介|CHT",
    null,
    { timeout: 5000 },
  );


  const sheetMoved = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "CHT|AI製品紹介",
    null,
    { timeout: 5000 },
  );


  const sheetMoveUndo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "AI製品紹介|CHT",
    null,
    { timeout: 5000 },
  );


  const sheetMoveRedo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "CHT|AI製品紹介",
    null,
    { timeout: 5000 },
  );


  await page.click('.sheet-tab[data-sheet-index="0"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="move-copy"]');


  await page.waitForSelector('[data-data-dialog="sheet-move-copy"]', { timeout: 5000 });


  await page.locator('[data-data-field="sheetMoveBeforeIndex"]').selectOption("__end");


  await page.locator('[data-data-field="sheetMoveCreateCopy"]').check();


  await page.locator('[data-data-dialog="sheet-move-copy"] [data-data-dialog-action="ok"]').click();


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "CHT|AI製品紹介|CHT (2)",
    null,
    { timeout: 5000 },
  );


  const sheetCopied = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "CHT|AI製品紹介",
    null,
    { timeout: 5000 },
  );


  const sheetCopyUndo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "CHT|AI製品紹介|CHT (2)",
    null,
    { timeout: 5000 },
  );


  const sheetCopyRedo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "CHT|AI製品紹介",
    null,
    { timeout: 5000 },
  );


  await page.click('.sheet-tab[data-sheet-index="0"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="insert"]');


  await page.waitForSelector('[data-data-dialog="sheet-insert"]', { timeout: 5000 });


  const sheetInsertDialog = await page.evaluate(() => ({
    title: document.querySelector("#sheetInsertTitle")?.textContent || "",
    activeField: document.activeElement?.dataset?.sheetInsertTemplate || "",
    tab: document.querySelector(".sheet-insert-tab.is-active")?.textContent || "",
    choices: [...document.querySelectorAll("[data-sheet-insert-template]")].map((button) => ({
      type: button.dataset.sheetInsertTemplate,
      label: button.textContent.replace(/\s+/g, " ").trim(),
      selected: button.classList.contains("is-selected"),
      disabled: button.disabled,
    })),
    isExcelDialog: Boolean(document.querySelector(".excel-mini-dialog.sheet-insert-dialog")),
  }));


  await page.locator('[data-data-dialog="sheet-insert"] [data-data-dialog-action="ok"]').click();


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "Sheet3|CHT|AI製品紹介",
    null,
    { timeout: 5000 },
  );


  const sheetInsertedFromMenu = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('.sheet-tab[data-sheet-index="0"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  const sheetDeleteConfirm = await page.evaluate(() => ({
    title: document.querySelector("#messageBoxTitle")?.textContent || "",
    message: document.querySelector("#messageBoxText")?.textContent || "",
    buttons: [...document.querySelectorAll("[data-message-box-close]")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
    deleteAccessKeyUnderlined: Boolean([...document.querySelectorAll("#messageBox .menu-access-key")].find((item) => item.textContent === "D" && getComputedStyle(item).textDecorationLine.includes("underline"))),
  }));


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "CHT|AI製品紹介",
    null,
    { timeout: 5000 },
  );


  const sheetDeleted = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "Sheet3|CHT|AI製品紹介",
    null,
    { timeout: 5000 },
  );


  const sheetDeleteUndo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(
    () => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "CHT|AI製品紹介",
    null,
    { timeout: 5000 },
  );


  const sheetDeleteRedo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(86, 21, 89, 23));


  const sheetInsertChartBefore = await page.evaluate(() => ({
    active: window.__webSheetDevWorkbookState?.().activeSheetName || "",
    tabCount: document.querySelectorAll(".sheet-tab").length,
  }));


  await page.click('.sheet-tab[data-sheet-index="0"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="insert"]');


  await page.waitForSelector('[data-data-dialog="sheet-insert"]', { timeout: 5000 });


  const sheetInsertChartDialog = await page.evaluate(() => ({
    title: document.querySelector("#sheetInsertTitle")?.textContent || "",
    chartDisabled: document.querySelector('[data-sheet-insert-template="chart"]')?.disabled === true,
    chartSelectedBefore: document.querySelector('[data-sheet-insert-template="chart"]')?.classList.contains("is-selected") || false,
  }));


  await page.dblclick('[data-sheet-insert-template="chart"]');


  await page.waitForFunction((before) => document.querySelectorAll(".sheet-tab").length === before.tabCount + 1, sheetInsertChartBefore, { timeout: 5000 });


  const sheetInsertChart = await page.evaluate((before) => {
    const state = window.__webSheetDevWorkbookState?.() || {};
    const sheet = window.__webSheetDevSheetModel?.(state.activeSheetName) || {};
    const chart = (sheet.images || [])[0]?.chart || {};
    return {
      previousActive: before.active,
      active: state.activeSheetName || "",
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
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
  }, sheetInsertChartBefore);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction((before) => document.querySelectorAll(".sheet-tab").length === before.tabCount, sheetInsertChartBefore, { timeout: 5000 });


  const sheetInsertChartUndo = await page.evaluate((before) => ({
    active: window.__webSheetDevWorkbookState?.().activeSheetName || "",
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    tabCount: document.querySelectorAll(".sheet-tab").length,
    expectedTabCount: before.tabCount,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetInsertChartBefore);


  const sheetSpecialTemplates = [];


  for (const templateType of ["macro", "intlMacro", "dialog"]) {
    const before = await page.evaluate(() => ({
      active: window.__webSheetDevWorkbookState?.().activeSheetName || "",
      tabCount: document.querySelectorAll(".sheet-tab").length,
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    }));
    await page.click('.sheet-tab[data-sheet-index="0"]', { button: "right" });
    await page.click('#sheetTabContextMenu [data-sheet-menu-action="insert"]');
    await page.waitForSelector('[data-data-dialog="sheet-insert"]', { timeout: 5000 });
    await page.click(`[data-sheet-insert-template="${templateType}"]`);
    const dialogState = await page.evaluate((type) => ({
      selected: document.querySelector(`[data-sheet-insert-template="${type}"]`)?.classList.contains("is-selected") || false,
      disabled: document.querySelector(`[data-sheet-insert-template="${type}"]`)?.disabled === true,
    }), templateType);
    await page.locator('[data-data-dialog="sheet-insert"] [data-data-dialog-action="ok"]').click();
    await page.waitForFunction((expected) => document.querySelectorAll(".sheet-tab").length === expected, before.tabCount + 1, { timeout: 5000 });
    const created = await page.evaluate((type) => {
      const state = window.__webSheetDevWorkbookState?.() || {};
      const sheet = window.__webSheetDevSheetModel?.(state.activeSheetName) || {};
      return {
        type,
        active: state.activeSheetName || "",
        tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
        tabCount: document.querySelectorAll(".sheet-tab").length,
        sheetKind: sheet.sheetKind || "",
        specialType: sheet.specialSheet?.type || "",
        macroSheet: sheet.macroSheet === true,
        xl4MacroSheet: sheet.xl4MacroSheet === true,
        internationalMacroSheet: sheet.internationalMacroSheet === true,
        dialogSheet: sheet.dialogSheet === true,
        showGridLines: sheet.showGridLines,
        showHeadings: sheet.showHeadings,
        codeObjectName: sheet.codeModule?.objectName || "",
        status: document.querySelector("#statusText")?.textContent || "",
      };
    }, templateType);
    await page.keyboard.press("Control+Z");
    await page.waitForFunction((expected) => document.querySelectorAll(".sheet-tab").length === expected, before.tabCount, { timeout: 5000 });
    const undo = await page.evaluate((beforeState) => ({
      active: window.__webSheetDevWorkbookState?.().activeSheetName || "",
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      tabCount: document.querySelectorAll(".sheet-tab").length,
      expectedTabCount: beforeState.tabCount,
      status: document.querySelector("#statusText")?.textContent || "",
    }), before);
    sheetSpecialTemplates.push({ type: templateType, dialog: dialogState, created, undo });
  }


  const sheetMoveCopy = {
    dialog: sheetMoveDialog,
    moved: sheetMoved,
    moveUndo: sheetMoveUndo,
    moveRedo: sheetMoveRedo,
    copied: sheetCopied,
    copyUndo: sheetCopyUndo,
    copyRedo: sheetCopyRedo,
    insertDialog: sheetInsertDialog,
    inserted: sheetInsertedFromMenu,
    deleteConfirm: sheetDeleteConfirm,
    deleted: sheetDeleted,
    deleteUndo: sheetDeleteUndo,
    deleteRedo: sheetDeleteRedo,
    chartDialog: sheetInsertChartDialog,
    chart: sheetInsertChart,
    chartUndo: sheetInsertChartUndo,
    specialTemplates: sheetSpecialTemplates,
  };


  await page.click('.sheet-tab[data-sheet-index="0"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="code"]');


  await page.waitForSelector('[data-data-dialog="sheet-code"]', { timeout: 5000 });


  const sheetCodeInitial = await page.evaluate(() => ({
    title: document.querySelector("#sheetCodeTitle")?.textContent || "",
    activeProjectItem: document.querySelector(".vba-project-tree .is-active")?.textContent || "",
    objectOptions: [...document.querySelectorAll("[data-data-field='sheetCodeObject'] option")].map((option) => option.textContent.trim()),
    procedureOptions: [...document.querySelectorAll("[data-data-field='sheetCodeProcedure'] option")].map((option) => option.textContent.trim()),
    code: document.querySelector("[data-data-field='sheetCode']")?.value || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.locator("[data-data-field='sheetCodeProcedure']").selectOption("Change");


  await page.waitForFunction(() => document.querySelector("[data-data-field='sheetCode']")?.value.includes("Worksheet_Change"), null, { timeout: 5000 });


  const sheetCodeProcedureInserted = await page.evaluate(() => document.querySelector("[data-data-field='sheetCode']")?.value || "");


  await page.locator("[data-data-field='sheetCode']").fill('Private Sub Worksheet_Change(ByVal Target As Range)\n  Range("A1").Value = "Changed"\nEnd Sub');


  await page.locator("[data-data-dialog-action='sheet-code-ok']").click();


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  await page.waitForFunction(() => /Changed/.test(window.__webSheetDevSheetModel?.("CHT")?.codeModule?.code || ""), null, { timeout: 5000 });


  const sheetCodeSaved = await page.evaluate(() => ({
    module: window.__webSheetDevSheetModel?.("CHT")?.codeModule || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => !window.__webSheetDevSheetModel?.("CHT")?.codeModule?.code, null, { timeout: 5000 });


  const sheetCodeUndo = await page.evaluate(() => ({
    module: window.__webSheetDevSheetModel?.("CHT")?.codeModule || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => /Changed/.test(window.__webSheetDevSheetModel?.("CHT")?.codeModule?.code || ""), null, { timeout: 5000 });


  const sheetCodeRedo = await page.evaluate(() => ({
    module: window.__webSheetDevSheetModel?.("CHT")?.codeModule || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('.sheet-tab[data-sheet-index="0"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="code"]');


  await page.waitForSelector('[data-data-dialog="sheet-code"]', { timeout: 5000 });


	  const sheetCodeReopened = await page.evaluate(() => ({
	    code: document.querySelector("[data-data-field='sheetCode']")?.value || "",
	    title: document.querySelector("#sheetCodeTitle")?.textContent || "",
	    procedure: document.querySelector("[data-data-field='sheetCodeProcedure']")?.value || "",
	  }));


  await page.locator("[data-data-dialog-action='cancel']").click();


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  const sheetCodeDialog = {
    initial: sheetCodeInitial,
    procedureInserted: sheetCodeProcedureInserted,
    saved: sheetCodeSaved,
    undo: sheetCodeUndo,
    redo: sheetCodeRedo,
    reopened: sheetCodeReopened,
  };


  await page.click('.sheet-tab[data-sheet-index="0"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="select-all"]');


  const sheetSelectAll = await page.evaluate(() => ({
    selectedTabs: document.querySelectorAll(".sheet-tab.is-group-selected").length,
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.click('.sheet-tab[data-sheet-index="0"]');


  await page.click('.sheet-tab[data-sheet-index="0"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="hide"]');


  await page.waitForFunction(
    () => !document.querySelector('.sheet-tab[data-sheet-index="0"]') && document.querySelector(".sheet-tab.is-active")?.textContent === "AI製品紹介",
    null,
    { timeout: 10000 },
  );


  const sheetHidden = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.click('.sheet-tab[data-sheet-index="1"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="unhide"]');


  await page.waitForSelector('[data-data-dialog="sheet-unhide"]', { timeout: 5000 });


  const sheetUnhideDialog = await page.evaluate(() => ({
    title: document.querySelector("#sheetUnhideTitle")?.textContent || "",
    caption: document.querySelector('[data-data-dialog="sheet-unhide"] .excel-dialog-caption')?.textContent || "",
    options: [...document.querySelectorAll('[data-data-field="unhideSheetIndex"] option')].map((option) => option.textContent.trim()),
    selected: document.querySelector('[data-data-field="unhideSheetIndex"] option:checked')?.textContent.trim() || "",
    activeField: document.activeElement?.dataset?.dataField || "",
    isExcelDialog: Boolean(document.querySelector(".excel-mini-dialog.sheet-unhide-dialog")),
  }));


  await page.locator('[data-data-dialog="sheet-unhide"] [data-data-dialog-action="ok"]').click();


  await page.waitForFunction(() => document.querySelector('.sheet-tab[data-sheet-index="0"]') && document.querySelector(".sheet-tab.is-active")?.textContent === "CHT", null, { timeout: 10000 });


  const sheetUnhidden = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent,
    status: document.querySelector("#statusText")?.textContent,
    dialog: window.__lastNativeDialogMessage || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => !document.querySelector('.sheet-tab[data-sheet-index="0"]'), null, { timeout: 5000 });


  const sheetUnhideUndo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => document.querySelector('.sheet-tab[data-sheet-index="0"]') && document.querySelector(".sheet-tab.is-active")?.textContent === "CHT", null, { timeout: 5000 });


  const sheetUnhideRedo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }));


  const sheetGroupBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupBaseCount, { timeout: 5000 });


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  const sheetGroupSelection = await page.evaluate(() => ({
    selectedTabs: [...document.querySelectorAll(".sheet-tab.is-group-selected")].map((button) => button.textContent),
    active: document.querySelector(".sheet-tab.is-active")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('.sheet-tab[data-sheet-index="2"]', { button: "right" });


  await page.hover("#sheetTabContextMenu .sheet-tab-menu-submenu");


  await page.click('#sheetTabContextMenu [data-sheet-tab-color="#70ad47"]');


  await page.waitForFunction(() =>
    document.querySelector('.sheet-tab[data-sheet-index="2"]')?.dataset.tabColor === "true" &&
    document.querySelector('.sheet-tab[data-sheet-index="3"]')?.dataset.tabColor === "true",
    null,
    { timeout: 5000 },
  );


  const sheetGroupColor = await page.evaluate(() => ({
    coloredTabs: [...document.querySelectorAll(".sheet-tab[data-tab-color='true']")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() =>
    document.querySelector('.sheet-tab[data-sheet-index="2"]')?.dataset.tabColor !== "true" &&
    document.querySelector('.sheet-tab[data-sheet-index="3"]')?.dataset.tabColor !== "true",
    null,
    { timeout: 5000 },
  );


  const sheetGroupColorUndo = await page.evaluate(() => ({
    coloredTabs: [...document.querySelectorAll(".sheet-tab[data-tab-color='true']")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  await page.click('.sheet-tab[data-sheet-index="2"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="hide"]');


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupBaseCount, { timeout: 5000 });


  const sheetGroupHidden = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupBaseCount, { timeout: 5000 });


  const sheetGroupHideUndo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  await page.click('.sheet-tab[data-sheet-index="2"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  const sheetGroupDeleteConfirm = await page.evaluate(() => ({
    message: document.querySelector("#messageBoxText")?.textContent || "",
    buttons: [...document.querySelectorAll("[data-message-box-close]")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
  }));


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupBaseCount, { timeout: 5000 });


  const sheetGroupDeleted = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupBaseCount, { timeout: 5000 });


  const sheetGroupDeleteUndo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupBaseCount, { timeout: 5000 });


  const sheetGroupDeleteRedo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  const sheetGroupActions = {
    selection: sheetGroupSelection,
    color: sheetGroupColor,
    colorUndo: sheetGroupColorUndo,
    hidden: sheetGroupHidden,
    hideUndo: sheetGroupHideUndo,
    deleteConfirm: sheetGroupDeleteConfirm,
    deleted: sheetGroupDeleted,
    deleteUndo: sheetGroupDeleteUndo,
    deleteRedo: sheetGroupDeleteRedo,
  };


  const sheetGroupMoveCopyBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupMoveCopyBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupMoveCopyBaseCount, { timeout: 5000 });


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  const sheetGroupDragMoveVisual = await dragSheetTab({ sourcePosition: 2, drop: "start" });


  await page.waitForFunction(() => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "Sheet3|Sheet4|CHT|AI製品紹介", null, { timeout: 5000 });


  const sheetGroupDragMoved = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    selectedTabs: [...document.querySelectorAll(".sheet-tab.is-group-selected")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "CHT|AI製品紹介|Sheet3|Sheet4", null, { timeout: 5000 });


  const sheetGroupDragMoveUndo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  const sheetGroupDragCopyVisual = await dragSheetTab({ sourcePosition: 2, drop: "start", copy: true });


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 4, sheetGroupMoveCopyBaseCount, { timeout: 5000 });


  const sheetGroupDragCopied = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    selectedTabs: [...document.querySelectorAll(".sheet-tab.is-group-selected")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "CHT|AI製品紹介|Sheet3|Sheet4", null, { timeout: 5000 });


  const sheetGroupDragCopyUndo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  await page.click('.sheet-tab[data-sheet-index="2"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="move-copy"]');


  await page.waitForSelector('[data-data-dialog="sheet-move-copy"]', { timeout: 5000 });


  const sheetGroupMoveDialog = await page.evaluate(() => ({
    title: document.querySelector("#sheetMoveCopyTitle")?.textContent || "",
    activeField: document.activeElement?.dataset?.dataField || "",
    hasCopyCheckbox: Boolean(document.querySelector('[data-data-field="sheetMoveCreateCopy"]')),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.locator('[data-data-field="sheetMoveBeforeIndex"]').selectOption("0");


  await page.locator('[data-data-dialog="sheet-move-copy"] [data-data-dialog-action="ok"]').click();


  await page.waitForFunction(() => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).slice(0, 2).join("|") === "Sheet3|Sheet4", null, { timeout: 5000 });


  const sheetGroupMoved = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    selectedTabs: [...document.querySelectorAll(".sheet-tab.is-group-selected")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "CHT|AI製品紹介|Sheet3|Sheet4", null, { timeout: 5000 });


  const sheetGroupMoveUndo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).slice(0, 2).join("|") === "Sheet3|Sheet4", null, { timeout: 5000 });


  const sheetGroupMoveRedo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    selectedTabs: [...document.querySelectorAll(".sheet-tab.is-group-selected")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent).join("|") === "CHT|AI製品紹介|Sheet3|Sheet4", null, { timeout: 5000 });


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  await page.click('.sheet-tab[data-sheet-index="2"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="move-copy"]');


  await page.waitForSelector('[data-data-dialog="sheet-move-copy"]', { timeout: 5000 });


  await page.locator('[data-data-field="sheetMoveBeforeIndex"]').selectOption("__end");


  await page.locator('[data-data-field="sheetMoveCreateCopy"]').check();


  await page.locator('[data-data-dialog="sheet-move-copy"] [data-data-dialog-action="ok"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 4, sheetGroupMoveCopyBaseCount, { timeout: 5000 });


  const sheetGroupCopied = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    selectedTabs: [...document.querySelectorAll(".sheet-tab.is-group-selected")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupMoveCopyBaseCount, { timeout: 5000 });


  const sheetGroupCopyUndo = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  await page.click('.sheet-tab[data-sheet-index="2"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupMoveCopyBaseCount, { timeout: 5000 });


  const sheetGroupMoveCopyCleanup = await page.evaluate(() => ({
    tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  const sheetGroupMoveCopy = {
    dragMove: {
      visual: sheetGroupDragMoveVisual,
      applied: sheetGroupDragMoved,
      undo: sheetGroupDragMoveUndo,
    },
    dragCopy: {
      visual: sheetGroupDragCopyVisual,
      applied: sheetGroupDragCopied,
      undo: sheetGroupDragCopyUndo,
    },
    dialog: sheetGroupMoveDialog,
    moved: sheetGroupMoved,
    moveUndo: sheetGroupMoveUndo,
    moveRedo: sheetGroupMoveRedo,
    copied: sheetGroupCopied,
    copyUndo: sheetGroupCopyUndo,
    cleanup: sheetGroupMoveCopyCleanup,
  };


  const sheetGroupCellBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupCellBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupCellBaseCount, { timeout: 5000 });


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  await page.evaluate(() => {
    window.__webSheetDevSelectRange?.(1, 1, 1, 1);
    window.__webSheetDevSetCell?.(1, 1, "GroupValue");
  });


  await page.waitForFunction(() =>
    window.__webSheetDevCellModel?.("Sheet3", 1, 1)?.raw === "GroupValue" &&
    window.__webSheetDevCellModel?.("Sheet4", 1, 1)?.raw === "GroupValue",
    null,
    { timeout: 5000 },
  );


  const sheetGroupCellInput = await page.evaluate(() => ({
    sheet3: window.__webSheetDevCellModel?.("Sheet3", 1, 1)?.raw || "",
    sheet4: window.__webSheetDevCellModel?.("Sheet4", 1, 1)?.raw || "",
    selectedTabs: [...document.querySelectorAll(".sheet-tab.is-group-selected")].map((button) => button.textContent),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 2, 1, 2));


  await page.click('[data-ribbon-command="bold"]');


  await page.waitForFunction(() =>
    window.__webSheetDevCellModel?.("Sheet3", 1, 2)?.css?.fontWeight === "700" &&
    window.__webSheetDevCellModel?.("Sheet4", 1, 2)?.css?.fontWeight === "700",
    null,
    { timeout: 5000 },
  );


  const sheetGroupCellFormat = await page.evaluate(() => ({
    sheet3: window.__webSheetDevCellModel?.("Sheet3", 1, 2)?.css?.fontWeight || "",
    sheet4: window.__webSheetDevCellModel?.("Sheet4", 1, 2)?.css?.fontWeight || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 3, 1, 4));


  await page.click('[data-ribbon-command="merge-center"]');


  await page.waitForFunction(() =>
    (window.__webSheetDevSheetModel?.("Sheet3")?.merges || []).some((merge) => merge.top === 1 && merge.left === 3 && merge.bottom === 1 && merge.right === 4) &&
    (window.__webSheetDevSheetModel?.("Sheet4")?.merges || []).some((merge) => merge.top === 1 && merge.left === 3 && merge.bottom === 1 && merge.right === 4),
    null,
    { timeout: 5000 },
  );


  const sheetGroupCellMerge = await page.evaluate(() => ({
    sheet3: window.__webSheetDevSheetModel?.("Sheet3")?.merges || [],
    sheet4: window.__webSheetDevSheetModel?.("Sheet4")?.merges || [],
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() =>
    !(window.__webSheetDevSheetModel?.("Sheet3")?.merges || []).some((merge) => merge.top === 1 && merge.left === 3 && merge.bottom === 1 && merge.right === 4) &&
    !(window.__webSheetDevSheetModel?.("Sheet4")?.merges || []).some((merge) => merge.top === 1 && merge.left === 3 && merge.bottom === 1 && merge.right === 4),
    null,
    { timeout: 5000 },
  );


  const sheetGroupCellMergeUndo = await page.evaluate(() => ({
    sheet3: window.__webSheetDevSheetModel?.("Sheet3")?.merges || [],
    sheet4: window.__webSheetDevSheetModel?.("Sheet4")?.merges || [],
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 1, 1, 1));


  await page.keyboard.press("Delete");


  await page.waitForFunction(() =>
    !window.__webSheetDevCellModel?.("Sheet3", 1, 1)?.raw &&
    !window.__webSheetDevCellModel?.("Sheet4", 1, 1)?.raw,
    null,
    { timeout: 5000 },
  );


  const sheetGroupCellClear = await page.evaluate(() => ({
    sheet3: window.__webSheetDevCellModel?.("Sheet3", 1, 1)?.raw || "",
    sheet4: window.__webSheetDevCellModel?.("Sheet4", 1, 1)?.raw || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() =>
    window.__webSheetDevCellModel?.("Sheet3", 1, 1)?.raw === "GroupValue" &&
    window.__webSheetDevCellModel?.("Sheet4", 1, 1)?.raw === "GroupValue",
    null,
    { timeout: 5000 },
  );


  const sheetGroupCellClearUndo = await page.evaluate(() => ({
    sheet3: window.__webSheetDevCellModel?.("Sheet3", 1, 1)?.raw || "",
    sheet4: window.__webSheetDevCellModel?.("Sheet4", 1, 1)?.raw || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() =>
    !window.__webSheetDevCellModel?.("Sheet3", 1, 1)?.raw &&
    !window.__webSheetDevCellModel?.("Sheet4", 1, 1)?.raw,
    null,
    { timeout: 5000 },
  );


  const sheetGroupCellClearRedo = await page.evaluate(() => ({
    sheet3: window.__webSheetDevCellModel?.("Sheet3", 1, 1)?.raw || "",
    sheet4: window.__webSheetDevCellModel?.("Sheet4", 1, 1)?.raw || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  await page.click('.sheet-tab[data-sheet-index="2"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupCellBaseCount, { timeout: 5000 });


  const sheetGroupCellEditing = {
    input: sheetGroupCellInput,
    format: sheetGroupCellFormat,
    merge: sheetGroupCellMerge,
    mergeUndo: sheetGroupCellMergeUndo,
    clear: sheetGroupCellClear,
    clearUndo: sheetGroupCellClearUndo,
    clearRedo: sheetGroupCellClearRedo,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };


  const sheetGroupStructureBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupStructureBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupStructureBaseCount, { timeout: 5000 });


  const sheetGroupStructureNames = await page.evaluate(() =>
    [...document.querySelectorAll(".sheet-tab")].slice(-2).map((button) => button.textContent),
  );


  await page.evaluate(([firstSheet, secondSheet]) => {
    window.__webSheetDevSetActiveSheet?.(firstSheet);
    window.__webSheetDevSetCell?.(2, 1, "GroupRow");
    window.__webSheetDevSetCell?.(1, 2, "GroupCol");
    window.__webSheetDevSetActiveSheet?.(secondSheet);
    window.__webSheetDevSetCell?.(2, 1, "GroupRow");
    window.__webSheetDevSetCell?.(1, 2, "GroupCol");
  }, sheetGroupStructureNames);


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(2, 1, 2, 1));


  await page.evaluate(() => window.__webSheetDevApplyCellStructure?.("insert", "entire-row", 2, 1, 2, 1));


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    window.__webSheetDevCellModel?.(firstSheet, 3, 1)?.raw === "GroupRow" &&
    window.__webSheetDevCellModel?.(secondSheet, 3, 1)?.raw === "GroupRow",
    sheetGroupStructureNames,
    { timeout: 5000 },
  );


  const sheetGroupRowInsert = await page.evaluate(([firstSheet, secondSheet]) => ({
    sheet3A3: window.__webSheetDevCellModel?.(firstSheet, 3, 1)?.raw || "",
    sheet4A3: window.__webSheetDevCellModel?.(secondSheet, 3, 1)?.raw || "",
    sheet3Rows: window.__webSheetDevSheetModel?.(firstSheet)?.rowCount || 0,
    sheet4Rows: window.__webSheetDevSheetModel?.(secondSheet)?.rowCount || 0,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupStructureNames);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    window.__webSheetDevCellModel?.(firstSheet, 2, 1)?.raw === "GroupRow" &&
    window.__webSheetDevCellModel?.(secondSheet, 2, 1)?.raw === "GroupRow",
    sheetGroupStructureNames,
    { timeout: 5000 },
  );


  const sheetGroupRowInsertUndo = await page.evaluate(([firstSheet, secondSheet]) => ({
    sheet3A2: window.__webSheetDevCellModel?.(firstSheet, 2, 1)?.raw || "",
    sheet4A2: window.__webSheetDevCellModel?.(secondSheet, 2, 1)?.raw || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupStructureNames);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    window.__webSheetDevCellModel?.(firstSheet, 3, 1)?.raw === "GroupRow" &&
    window.__webSheetDevCellModel?.(secondSheet, 3, 1)?.raw === "GroupRow",
    sheetGroupStructureNames,
    { timeout: 5000 },
  );


  await page.evaluate(() => window.__webSheetDevApplyCellStructure?.("delete", "entire-row", 2, 1, 2, 1));


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    window.__webSheetDevCellModel?.(firstSheet, 2, 1)?.raw === "GroupRow" &&
    window.__webSheetDevCellModel?.(secondSheet, 2, 1)?.raw === "GroupRow",
    sheetGroupStructureNames,
    { timeout: 5000 },
  );


  const sheetGroupRowDelete = await page.evaluate(([firstSheet, secondSheet]) => ({
    sheet3A2: window.__webSheetDevCellModel?.(firstSheet, 2, 1)?.raw || "",
    sheet4A2: window.__webSheetDevCellModel?.(secondSheet, 2, 1)?.raw || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupStructureNames);


  await page.evaluate(() => window.__webSheetDevApplyCellStructure?.("insert", "entire-column", 1, 2, 1, 2));


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    window.__webSheetDevCellModel?.(firstSheet, 1, 3)?.raw === "GroupCol" &&
    window.__webSheetDevCellModel?.(secondSheet, 1, 3)?.raw === "GroupCol",
    sheetGroupStructureNames,
    { timeout: 5000 },
  );


  const sheetGroupColumnInsert = await page.evaluate(([firstSheet, secondSheet]) => ({
    sheet3C1: window.__webSheetDevCellModel?.(firstSheet, 1, 3)?.raw || "",
    sheet4C1: window.__webSheetDevCellModel?.(secondSheet, 1, 3)?.raw || "",
    sheet3Cols: window.__webSheetDevSheetModel?.(firstSheet)?.colCount || 0,
    sheet4Cols: window.__webSheetDevSheetModel?.(secondSheet)?.colCount || 0,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupStructureNames);


  await page.evaluate(() => window.__webSheetDevApplyCellStructure?.("delete", "entire-column", 1, 2, 1, 2));


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    window.__webSheetDevCellModel?.(firstSheet, 1, 2)?.raw === "GroupCol" &&
    window.__webSheetDevCellModel?.(secondSheet, 1, 2)?.raw === "GroupCol",
    sheetGroupStructureNames,
    { timeout: 5000 },
  );


  const sheetGroupColumnDelete = await page.evaluate(([firstSheet, secondSheet]) => ({
    sheet3B1: window.__webSheetDevCellModel?.(firstSheet, 1, 2)?.raw || "",
    sheet4B1: window.__webSheetDevCellModel?.(secondSheet, 1, 2)?.raw || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupStructureNames);


  await page.evaluate(() => window.__webSheetDevApplyHeaderSize?.("column", 1, 2, 1, 3, 91));


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    window.__webSheetDevSheetModel?.(firstSheet)?.columns?.[1]?.width === 91 &&
    window.__webSheetDevSheetModel?.(firstSheet)?.columns?.[2]?.width === 91 &&
    window.__webSheetDevSheetModel?.(secondSheet)?.columns?.[1]?.width === 91 &&
    window.__webSheetDevSheetModel?.(secondSheet)?.columns?.[2]?.width === 91,
    sheetGroupStructureNames,
    { timeout: 5000 },
  );


  const sheetGroupColumnWidth = await page.evaluate(([firstSheet, secondSheet]) => ({
    sheet3: [window.__webSheetDevSheetModel?.(firstSheet)?.columns?.[1]?.width, window.__webSheetDevSheetModel?.(firstSheet)?.columns?.[2]?.width],
    sheet4: [window.__webSheetDevSheetModel?.(secondSheet)?.columns?.[1]?.width, window.__webSheetDevSheetModel?.(secondSheet)?.columns?.[2]?.width],
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupStructureNames);


  await page.evaluate(() => window.__webSheetDevApplyHeaderSize?.("row", 4, 1, 5, 1, 33));


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    window.__webSheetDevSheetModel?.(firstSheet)?.rows?.[3]?.height === 33 &&
    window.__webSheetDevSheetModel?.(firstSheet)?.rows?.[4]?.height === 33 &&
    window.__webSheetDevSheetModel?.(secondSheet)?.rows?.[3]?.height === 33 &&
    window.__webSheetDevSheetModel?.(secondSheet)?.rows?.[4]?.height === 33,
    sheetGroupStructureNames,
    { timeout: 5000 },
  );


  const sheetGroupRowHeight = await page.evaluate(([firstSheet, secondSheet]) => ({
    sheet3: [window.__webSheetDevSheetModel?.(firstSheet)?.rows?.[3]?.height, window.__webSheetDevSheetModel?.(firstSheet)?.rows?.[4]?.height],
    sheet4: [window.__webSheetDevSheetModel?.(secondSheet)?.rows?.[3]?.height, window.__webSheetDevSheetModel?.(secondSheet)?.rows?.[4]?.height],
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupStructureNames);


  await page.evaluate(() => window.__webSheetDevSetHeaderHiddenWithHistory?.("column", true, 1, 2, 1, 3));


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    window.__webSheetDevSheetModel?.(firstSheet)?.columns?.[1]?.hidden === true &&
    window.__webSheetDevSheetModel?.(firstSheet)?.columns?.[2]?.hidden === true &&
    window.__webSheetDevSheetModel?.(secondSheet)?.columns?.[1]?.hidden === true &&
    window.__webSheetDevSheetModel?.(secondSheet)?.columns?.[2]?.hidden === true,
    sheetGroupStructureNames,
    { timeout: 5000 },
  );


  const sheetGroupColumnHidden = await page.evaluate(([firstSheet, secondSheet]) => ({
    sheet3: [window.__webSheetDevSheetModel?.(firstSheet)?.columns?.[1]?.hidden, window.__webSheetDevSheetModel?.(firstSheet)?.columns?.[2]?.hidden],
    sheet4: [window.__webSheetDevSheetModel?.(secondSheet)?.columns?.[1]?.hidden, window.__webSheetDevSheetModel?.(secondSheet)?.columns?.[2]?.hidden],
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupStructureNames);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    window.__webSheetDevSheetModel?.(firstSheet)?.columns?.[1]?.hidden !== true &&
    window.__webSheetDevSheetModel?.(firstSheet)?.columns?.[2]?.hidden !== true &&
    window.__webSheetDevSheetModel?.(secondSheet)?.columns?.[1]?.hidden !== true &&
    window.__webSheetDevSheetModel?.(secondSheet)?.columns?.[2]?.hidden !== true,
    sheetGroupStructureNames,
    { timeout: 5000 },
  );


  const sheetGroupColumnHiddenUndo = await page.evaluate(([firstSheet, secondSheet]) => ({
    sheet3: [window.__webSheetDevSheetModel?.(firstSheet)?.columns?.[1]?.hidden === true, window.__webSheetDevSheetModel?.(firstSheet)?.columns?.[2]?.hidden === true],
    sheet4: [window.__webSheetDevSheetModel?.(secondSheet)?.columns?.[1]?.hidden === true, window.__webSheetDevSheetModel?.(secondSheet)?.columns?.[2]?.hidden === true],
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupStructureNames);


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  await page.click('.sheet-tab[data-sheet-index="2"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupStructureBaseCount, { timeout: 5000 });


  const sheetGroupStructure = {
    rowInsert: sheetGroupRowInsert,
    rowInsertUndo: sheetGroupRowInsertUndo,
    rowDelete: sheetGroupRowDelete,
    columnInsert: sheetGroupColumnInsert,
    columnDelete: sheetGroupColumnDelete,
    columnWidth: sheetGroupColumnWidth,
    rowHeight: sheetGroupRowHeight,
    columnHidden: sheetGroupColumnHidden,
    columnHiddenUndo: sheetGroupColumnHiddenUndo,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };


  const sheetGroupFormatBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupFormatBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupFormatBaseCount, { timeout: 5000 });


  const sheetGroupFormatNames = await page.evaluate(() =>
    [...document.querySelectorAll(".sheet-tab")].slice(-2).map((button) => button.textContent),
  );


  await page.evaluate(([firstSheet, secondSheet]) => {
    window.__webSheetDevSetActiveSheet?.(firstSheet);
    window.__webSheetDevSetCell?.(2, 2, 1234);
    window.__webSheetDevSetCell?.(2, 3, 5678);
    window.__webSheetDevSetActiveSheet?.(secondSheet);
    window.__webSheetDevSetCell?.(2, 2, 1234);
    window.__webSheetDevSetCell?.(2, 3, 5678);
  }, sheetGroupFormatNames);


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  await page.evaluate(() => {
    window.__webSheetDevSelectRange?.(2, 2, 2, 3);
    window.__webSheetDevOpenFormatCellsDialog?.(2, 2);
  });


  await page.waitForSelector("#formatCellsDialog", { timeout: 5000 });


  await page.selectOption('[data-format-field="numberCategory"]', "currency");


  await page.locator('[data-format-field="decimals"]').fill("1");


  await page.click('[data-format-dialog-tab="alignment"]');


  await page.selectOption('[data-format-field="horizontal"]', "center");


  await page.selectOption('[data-format-field="vertical"]', "top");


  const groupFormatWrap = page.locator('[data-format-field="wrapText"]');


  if (!(await groupFormatWrap.isChecked())) await groupFormatWrap.check();


  await page.click('[data-format-dialog-tab="font"]');


  await page.selectOption('[data-format-field="fontSize"]', "16");


  await page.locator('[data-format-field="fontColor"]').evaluate((input) => {
    input.value = "#c00000";
    input.dispatchEvent(new Event("input", { bubbles: true }));
  });


  await page.click('[data-format-dialog-tab="border"]');


  await page.selectOption('[data-format-field="borderStyle"]', "double");


  await page.selectOption('[data-format-field="borderPreset"]', "all");


  await page.locator('[data-format-field="borderColor"]').evaluate((input) => {
    input.value = "#00b050";
    input.dispatchEvent(new Event("input", { bubbles: true }));
  });


  await page.click('[data-format-dialog-tab="fill"]');


  const groupFormatNoFill = page.locator('[data-format-field="noFill"]');


  if (await groupFormatNoFill.isChecked()) await groupFormatNoFill.uncheck();


  await page.locator('[data-format-field="fillColor"]').evaluate((input) => {
    input.value = "#fce4d6";
    input.dispatchEvent(new Event("input", { bubbles: true }));
  });


  await page.click('[data-format-dialog-tab="protection"]');


  const groupFormatHiddenFormula = page.locator('[data-format-field="hiddenFormula"]');


  if (!(await groupFormatHiddenFormula.isChecked())) await groupFormatHiddenFormula.check();


  await page.click('[data-format-dialog-action="ok"]');


  await page.waitForSelector("#formatCellsDialog", { state: "detached", timeout: 5000 });


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    window.__webSheetDevCellModel?.(firstSheet, 2, 2)?.css?.backgroundColor === "#FCE4D6" &&
    window.__webSheetDevCellModel?.(secondSheet, 2, 2)?.css?.backgroundColor === "#FCE4D6",
    sheetGroupFormatNames,
    { timeout: 5000 },
  );


  const sheetGroupFormatApplied = await page.evaluate(([firstSheet, secondSheet]) => {
    const read = (sheetName, row, col) => {
      const cell = window.__webSheetDevCellModel?.(sheetName, row, col) || {};
      return {
        numFmt: cell.numFmt || "",
        backgroundColor: cell.css?.backgroundColor || "",
        color: cell.css?.color || "",
        fontSize: cell.css?.fontSize || "",
        textAlign: cell.css?.textAlign || "",
        verticalAlign: cell.css?.verticalAlign || "",
        whiteSpace: cell.css?.whiteSpace || "",
        borderTop: cell.borders?.top || null,
        protection: cell.protection || null,
      };
    };
    return {
      firstB2: read(firstSheet, 2, 2),
      firstC2: read(firstSheet, 2, 3),
      secondB2: read(secondSheet, 2, 2),
      secondC2: read(secondSheet, 2, 3),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupFormatNames);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    !window.__webSheetDevCellModel?.(firstSheet, 2, 2)?.css?.backgroundColor &&
    !window.__webSheetDevCellModel?.(secondSheet, 2, 2)?.css?.backgroundColor,
    sheetGroupFormatNames,
    { timeout: 5000 },
  );


  const sheetGroupFormatUndo = await page.evaluate(([firstSheet, secondSheet]) => ({
    firstB2Fill: window.__webSheetDevCellModel?.(firstSheet, 2, 2)?.css?.backgroundColor || "",
    secondB2Fill: window.__webSheetDevCellModel?.(secondSheet, 2, 2)?.css?.backgroundColor || "",
    firstB2Format: window.__webSheetDevCellModel?.(firstSheet, 2, 2)?.numFmt || "",
    secondB2Format: window.__webSheetDevCellModel?.(secondSheet, 2, 2)?.numFmt || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupFormatNames);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    window.__webSheetDevCellModel?.(firstSheet, 2, 2)?.css?.backgroundColor === "#FCE4D6" &&
    window.__webSheetDevCellModel?.(secondSheet, 2, 2)?.css?.backgroundColor === "#FCE4D6",
    sheetGroupFormatNames,
    { timeout: 5000 },
  );


  const sheetGroupFormatRedo = await page.evaluate(([firstSheet, secondSheet]) => ({
    firstB2Fill: window.__webSheetDevCellModel?.(firstSheet, 2, 2)?.css?.backgroundColor || "",
    secondB2Fill: window.__webSheetDevCellModel?.(secondSheet, 2, 2)?.css?.backgroundColor || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupFormatNames);


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  await page.click('.sheet-tab[data-sheet-index="2"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupFormatBaseCount, { timeout: 5000 });


  const sheetGroupFormatDialog = {
    applied: sheetGroupFormatApplied,
    undo: sheetGroupFormatUndo,
    redo: sheetGroupFormatRedo,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };


  const sheetGroupValidationBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupValidationBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupValidationBaseCount, { timeout: 5000 });


  const sheetGroupValidationNames = await page.evaluate(() =>
    [...document.querySelectorAll(".sheet-tab")].slice(-2).map((button) => button.textContent),
  );


  await page.evaluate(([firstSheet, secondSheet]) => {
    window.__webSheetDevSetActiveSheet?.(firstSheet);
    window.__webSheetDevSetCell?.(2, 2, "Red");
    window.__webSheetDevSetCell?.(2, 3, "Blue");
    window.__webSheetDevSetActiveSheet?.(secondSheet);
    window.__webSheetDevSetCell?.(2, 2, "Red");
    window.__webSheetDevSetCell?.(2, 3, "Blue");
  }, sheetGroupValidationNames);


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(2, 2, 2, 3));


  await page.click('[data-ribbon-tab="data"]');


  await page.click('[data-ribbon-command="data-validation"]');


  await page.waitForSelector("#dataValidationMenu", { timeout: 5000 });


  await page.click('[data-validation-menu-action="open"]');


  await page.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });


  await page.selectOption("[data-data-field='allow']", "list");


  await page.locator("[data-data-field='source']").fill("Red,Blue,Green");


  await page.locator("[data-validation-tab='input']").click();


  const sheetGroupValidationInputMessage = page.locator("[data-data-field='showInputMessage']");


  if (!(await sheetGroupValidationInputMessage.isChecked())) await sheetGroupValidationInputMessage.check();


  await page.locator("[data-data-field='inputTitle']").fill("選択");


  await page.locator("[data-data-field='inputMessage']").fill("リストから選択してください。");


  await page.locator("[data-validation-tab='error']").click();


  await page.selectOption("[data-data-field='errorStyle']", "warning");


  await page.locator("[data-data-field='errorTitle']").fill("入力確認");


  await page.locator("[data-data-field='errorMessage']").fill("リスト内の値を入力してください。");


  await page.click('[data-data-dialog-action="ok"]');


  await page.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    window.__webSheetDevCellModel?.(firstSheet, 2, 2)?.validation?.type === "list" &&
    window.__webSheetDevCellModel?.(secondSheet, 2, 2)?.validation?.type === "list",
    sheetGroupValidationNames,
    { timeout: 5000 },
  );


  const sheetGroupValidationApplied = await page.evaluate(([firstSheet, secondSheet]) => {
    const read = (sheetName, row, col) => window.__webSheetDevCellModel?.(sheetName, row, col)?.validation || null;
    return {
      firstB2: read(firstSheet, 2, 2),
      firstC2: read(firstSheet, 2, 3),
      secondB2: read(secondSheet, 2, 2),
      secondC2: read(secondSheet, 2, 3),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupValidationNames);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    !window.__webSheetDevCellModel?.(firstSheet, 2, 2)?.validation &&
    !window.__webSheetDevCellModel?.(secondSheet, 2, 2)?.validation,
    sheetGroupValidationNames,
    { timeout: 5000 },
  );


  const sheetGroupValidationUndo = await page.evaluate(([firstSheet, secondSheet]) => ({
    firstB2: window.__webSheetDevCellModel?.(firstSheet, 2, 2)?.validation || null,
    secondB2: window.__webSheetDevCellModel?.(secondSheet, 2, 2)?.validation || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupValidationNames);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    window.__webSheetDevCellModel?.(firstSheet, 2, 2)?.validation?.type === "list" &&
    window.__webSheetDevCellModel?.(secondSheet, 2, 2)?.validation?.type === "list",
    sheetGroupValidationNames,
    { timeout: 5000 },
  );


  const sheetGroupValidationRedo = await page.evaluate(([firstSheet, secondSheet]) => ({
    firstB2: window.__webSheetDevCellModel?.(firstSheet, 2, 2)?.validation || null,
    secondB2: window.__webSheetDevCellModel?.(secondSheet, 2, 2)?.validation || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupValidationNames);


  await page.click('[data-ribbon-command="data-validation"]');


  await page.waitForSelector("#dataValidationMenu", { timeout: 5000 });


  await page.click('[data-validation-menu-action="clear-validation"]');


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    !window.__webSheetDevCellModel?.(firstSheet, 2, 2)?.validation &&
    !window.__webSheetDevCellModel?.(secondSheet, 2, 2)?.validation &&
    !window.__webSheetDevCellModel?.(firstSheet, 2, 3)?.validation &&
    !window.__webSheetDevCellModel?.(secondSheet, 2, 3)?.validation,
    sheetGroupValidationNames,
    { timeout: 5000 },
  );


  const sheetGroupValidationClear = await page.evaluate(([firstSheet, secondSheet]) => ({
    firstB2: window.__webSheetDevCellModel?.(firstSheet, 2, 2)?.validation || null,
    firstC2: window.__webSheetDevCellModel?.(firstSheet, 2, 3)?.validation || null,
    secondB2: window.__webSheetDevCellModel?.(secondSheet, 2, 2)?.validation || null,
    secondC2: window.__webSheetDevCellModel?.(secondSheet, 2, 3)?.validation || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupValidationNames);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([firstSheet, secondSheet]) =>
    window.__webSheetDevCellModel?.(firstSheet, 2, 2)?.validation?.type === "list" &&
    window.__webSheetDevCellModel?.(secondSheet, 2, 2)?.validation?.type === "list",
    sheetGroupValidationNames,
    { timeout: 5000 },
  );


  const sheetGroupValidationClearUndo = await page.evaluate(([firstSheet, secondSheet]) => ({
    firstB2: window.__webSheetDevCellModel?.(firstSheet, 2, 2)?.validation || null,
    secondB2: window.__webSheetDevCellModel?.(secondSheet, 2, 2)?.validation || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupValidationNames);


  await page.click('.sheet-tab[data-sheet-index="2"]');


  await page.click('.sheet-tab[data-sheet-index="3"]', { modifiers: ["Meta"] });


  await page.click('.sheet-tab[data-sheet-index="2"]', { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupValidationBaseCount, { timeout: 5000 });


  const sheetGroupDataValidation = {
    applied: sheetGroupValidationApplied,
    undo: sheetGroupValidationUndo,
    redo: sheetGroupValidationRedo,
    clear: sheetGroupValidationClear,
    clearUndo: sheetGroupValidationClearUndo,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };



  const sheetGroupDataFilterBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupDataFilterBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupDataFilterBaseCount, { timeout: 5000 });


  const sheetGroupDataFilterSheets = await page.evaluate(() =>
    [...document.querySelectorAll(".sheet-tab")]
      .slice(-2)
      .map((button) => ({ name: button.textContent.trim(), index: button.dataset.sheetIndex })),
  );


  await page.evaluate(([first, second]) => {
    const fill = (sheetName, rows) => {
      window.__webSheetDevSetActiveSheet?.(sheetName);
      window.__webSheetDevSetCell?.(1, 1, "Name");
      window.__webSheetDevSetCell?.(1, 2, "Score");
      rows.forEach((row, index) => {
        window.__webSheetDevSetCell?.(index + 2, 1, row[0]);
        window.__webSheetDevSetCell?.(index + 2, 2, row[1]);
      });
    };
    fill(first.name, [["Three", 3], ["One", 1], ["Two", 2]]);
    fill(second.name, [["Thirty", 30], ["Ten", 10], ["Twenty", 20]]);
  }, sheetGroupDataFilterSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupDataFilterSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupDataFilterSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 1, 4, 2));


  await page.click('[data-ribbon-tab="data"]');


  await page.click('[data-ribbon-command="data-filter"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevSheetModel?.(first.name)?.autoFilter?.range?.right === 2 &&
    window.__webSheetDevSheetModel?.(second.name)?.autoFilter?.range?.right === 2,
    sheetGroupDataFilterSheets,
    { timeout: 5000 },
  );


  const readGroupedFilterSort = async () => page.evaluate(([first, second]) => {
    const read = (sheetInfo) => {
      const sheet = window.__webSheetDevSheetModel?.(sheetInfo.name) || {};
      const raw = (row, col) => sheet.cells?.[`${row}:${col}`]?.raw ?? "";
      return {
        a: [2, 3, 4].map((row) => raw(row, 1)),
        b: [2, 3, 4].map((row) => raw(row, 2)),
        autoFilter: sheet.autoFilter || null,
        sortState: sheet.sortState || null,
        hiddenRows: [2, 3, 4].map((row) => sheet.rows?.[row - 1]?.hidden === true),
      };
    };
    return {
      first: read(first),
      second: read(second),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupDataFilterSheets);


  const sheetGroupFilterApplied = await readGroupedFilterSort();


  await page.evaluate(() => window.__webSheetDevSelectRange?.(2, 2, 2, 2));


  await page.click('[data-ribbon-command="data-sort-asc"]');


  await page.waitForFunction(([first, second]) => {
    const bValues = (sheetInfo) => {
      const sheet = window.__webSheetDevSheetModel?.(sheetInfo.name) || {};
      return [2, 3, 4].map((row) => sheet.cells?.[`${row}:2`]?.raw ?? "").join("|");
    };
    return bValues(first) === "1|2|3" && bValues(second) === "10|20|30";
  }, sheetGroupDataFilterSheets, { timeout: 5000 });


  const sheetGroupSortApplied = await readGroupedFilterSort();


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) => {
    const bValues = (sheetInfo) => {
      const sheet = window.__webSheetDevSheetModel?.(sheetInfo.name) || {};
      return [2, 3, 4].map((row) => sheet.cells?.[`${row}:2`]?.raw ?? "").join("|");
    };
    return bValues(first) === "3|1|2" && bValues(second) === "30|10|20";
  }, sheetGroupDataFilterSheets, { timeout: 5000 });


  const sheetGroupSortUndo = await readGroupedFilterSort();


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) => {
    const bValues = (sheetInfo) => {
      const sheet = window.__webSheetDevSheetModel?.(sheetInfo.name) || {};
      return [2, 3, 4].map((row) => sheet.cells?.[`${row}:2`]?.raw ?? "").join("|");
    };
    return bValues(first) === "1|2|3" && bValues(second) === "10|20|30";
  }, sheetGroupDataFilterSheets, { timeout: 5000 });


  const sheetGroupSortRedo = await readGroupedFilterSort();


  await page.click('[data-ribbon-command="data-reapply-filter"]');


  await page.waitForFunction(() => (document.querySelector("#statusText")?.textContent || "").includes("2 個のシート"), null, { timeout: 5000 });


  const sheetGroupReapply = await readGroupedFilterSort();


  await page.click('[data-ribbon-command="data-clear-filter"]');


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevSheetModel?.(first.name)?.sortState &&
    !window.__webSheetDevSheetModel?.(second.name)?.sortState,
    sheetGroupDataFilterSheets,
    { timeout: 5000 },
  );


  const sheetGroupClearFilter = await readGroupedFilterSort();


  await page.click('[data-ribbon-command="data-filter"]');


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevSheetModel?.(first.name)?.autoFilter &&
    !window.__webSheetDevSheetModel?.(second.name)?.autoFilter,
    sheetGroupDataFilterSheets,
    { timeout: 5000 },
  );


  const sheetGroupFilterRemoved = await readGroupedFilterSort();


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevSheetModel?.(first.name)?.autoFilter?.range?.right === 2 &&
    window.__webSheetDevSheetModel?.(second.name)?.autoFilter?.range?.right === 2,
    sheetGroupDataFilterSheets,
    { timeout: 5000 },
  );


  const sheetGroupFilterRemoveUndo = await readGroupedFilterSort();


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupDataFilterSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupDataFilterSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupDataFilterSheets[0].index}"]`, { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupDataFilterBaseCount, { timeout: 5000 });


  const sheetGroupDataFilterSort = {
    filter: sheetGroupFilterApplied,
    sort: sheetGroupSortApplied,
    sortUndo: sheetGroupSortUndo,
    sortRedo: sheetGroupSortRedo,
    reapply: sheetGroupReapply,
    clear: sheetGroupClearFilter,
    removed: sheetGroupFilterRemoved,
    removeUndo: sheetGroupFilterRemoveUndo,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };



  const sheetGroupDataToolsBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupDataToolsBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupDataToolsBaseCount, { timeout: 5000 });


  const sheetGroupDataToolsSheets = await page.evaluate(() =>
    [...document.querySelectorAll(".sheet-tab")]
      .slice(-2)
      .map((button) => ({ name: button.textContent.trim(), index: button.dataset.sheetIndex })),
  );


  await page.evaluate(([first, second]) => {
    const fill = (sheetName, rows) => {
      window.__webSheetDevSetActiveSheet?.(sheetName);
      rows.forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
    };
    fill(first.name, [
      [1, 1, "Full"], [1, 2, "First"],
      [2, 1, "Ada Lovelace"], [2, 2, "Ada"],
      [3, 1, "Grace Hopper"], [4, 1, "Alan Turing"],
      [7, 1, "x,y"], [8, 1, "p,q"],
      [1, 4, "Name"], [1, 5, "Score"],
      [2, 4, "A"], [2, 5, 1],
      [3, 4, "B"], [3, 5, 2],
      [4, 4, "A"], [4, 5, 1],
      [5, 4, "C"], [5, 5, 3],
    ]);
    fill(second.name, [
      [1, 1, "Full"], [1, 2, "First"],
      [2, 1, "Marie Curie"], [2, 2, "Marie"],
      [3, 1, "Niels Bohr"], [4, 1, "Rosalind Franklin"],
      [7, 1, "r,s"], [8, 1, "t,u"],
      [1, 4, "Name"], [1, 5, "Score"],
      [2, 4, "X"], [2, 5, 10],
      [3, 4, "Y"], [3, 5, 20],
      [4, 4, "X"], [4, 5, 10],
      [5, 4, "Z"], [5, 5, 30],
    ]);
  }, sheetGroupDataToolsSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupDataToolsSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupDataToolsSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(2, 2, 4, 2));


  await page.click('[data-ribbon-tab="data"]');


  await page.click('[data-ribbon-command="data-flash-fill"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 3, 2)?.raw === "Grace" &&
    window.__webSheetDevCellModel?.(first.name, 4, 2)?.raw === "Alan" &&
    window.__webSheetDevCellModel?.(second.name, 3, 2)?.raw === "Niels" &&
    window.__webSheetDevCellModel?.(second.name, 4, 2)?.raw === "Rosalind",
    sheetGroupDataToolsSheets,
    { timeout: 5000 },
  );


  const sheetGroupDataToolsFlash = await page.evaluate(([first, second]) => ({
    firstB: [2, 3, 4].map((row) => window.__webSheetDevCellModel?.(first.name, row, 2)?.raw ?? ""),
    secondB: [2, 3, 4].map((row) => window.__webSheetDevCellModel?.(second.name, row, 2)?.raw ?? ""),
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupDataToolsSheets);


  await page.evaluate(() => window.__webSheetDevSelectRange?.(7, 1, 8, 1));


  await page.click('[data-ribbon-command="data-text-to-columns"]');


  await page.waitForSelector('[data-data-dialog="text-to-columns"]', { timeout: 5000 });


  await page.click('[data-data-dialog-action="text-columns-finish"]');


  await page.waitForSelector('[data-data-dialog="text-to-columns"]', { state: "detached", timeout: 5000 });


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 7, 2)?.raw === "y" &&
    window.__webSheetDevCellModel?.(first.name, 8, 2)?.raw === "q" &&
    window.__webSheetDevCellModel?.(second.name, 7, 2)?.raw === "s" &&
    window.__webSheetDevCellModel?.(second.name, 8, 2)?.raw === "u",
    sheetGroupDataToolsSheets,
    { timeout: 5000 },
  );


  const sheetGroupDataToolsTextColumns = await page.evaluate(([first, second]) => ({
    first: [[7, 1], [7, 2], [8, 1], [8, 2]].map(([row, col]) => window.__webSheetDevCellModel?.(first.name, row, col)?.raw ?? ""),
    second: [[7, 1], [7, 2], [8, 1], [8, 2]].map(([row, col]) => window.__webSheetDevCellModel?.(second.name, row, col)?.raw ?? ""),
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupDataToolsSheets);


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 4, 5, 5));


  await page.click('[data-ribbon-command="data-remove-duplicates"]');


  await page.waitForSelector('[data-data-dialog="remove-duplicates"]', { timeout: 5000 });


  await page.click('[data-data-dialog-action="ok"]');


  await page.waitForSelector('[data-data-dialog="remove-duplicates"]', { state: "detached", timeout: 5000 });


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator("[data-message-box-close]").first().click();


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 5, 4)?.raw &&
    !window.__webSheetDevCellModel?.(second.name, 5, 4)?.raw,
    sheetGroupDataToolsSheets,
    { timeout: 5000 },
  );


  const readSheetGroupDataTools = async () => page.evaluate(([first, second]) => {
    const read = (sheetInfo) => ({
      d: [2, 3, 4, 5].map((row) => window.__webSheetDevCellModel?.(sheetInfo.name, row, 4)?.raw ?? ""),
      e: [2, 3, 4, 5].map((row) => window.__webSheetDevCellModel?.(sheetInfo.name, row, 5)?.raw ?? ""),
    });
    return {
      first: read(first),
      second: read(second),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupDataToolsSheets);


  const sheetGroupDataToolsRemoveDuplicates = await readSheetGroupDataTools();


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 5, 4)?.raw === "C" &&
    window.__webSheetDevCellModel?.(second.name, 5, 4)?.raw === "Z",
    sheetGroupDataToolsSheets,
    { timeout: 5000 },
  );


  const sheetGroupDataToolsRemoveUndo = await readSheetGroupDataTools();


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 5, 4)?.raw &&
    !window.__webSheetDevCellModel?.(second.name, 5, 4)?.raw,
    sheetGroupDataToolsSheets,
    { timeout: 5000 },
  );


  const sheetGroupDataToolsRemoveRedo = await readSheetGroupDataTools();


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 4, 4, 5));


  await page.click('[data-ribbon-command="data-subtotal"]');


  await page.waitForSelector('[data-data-dialog="subtotal"]', { timeout: 5000 });


  await page.click('[data-data-dialog-action="ok"]');


  await page.waitForSelector('[data-data-dialog="subtotal"]', { state: "detached", timeout: 5000 });


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 8, 4)?.raw === "総計" &&
    window.__webSheetDevCellModel?.(second.name, 8, 4)?.raw === "総計",
    sheetGroupDataToolsSheets,
    { timeout: 5000 },
  );


  const readSheetGroupSubtotalOutline = async () => page.evaluate(([first, second]) => {
    const read = (sheetInfo) => {
      const sheet = window.__webSheetDevSheetModel?.(sheetInfo.name) || {};
      return {
        d: [1, 2, 3, 4, 5, 6, 7, 8].map((row) => sheet.cells?.[`${row}:4`]?.raw ?? ""),
        e: [1, 2, 3, 4, 5, 6, 7, 8].map((row) => sheet.cells?.[`${row}:5`]?.formula || sheet.cells?.[`${row}:5`]?.raw || ""),
        outlineLevels: [2, 3, 4, 5, 6, 7, 8].map((row) => sheet.rows?.[row - 1]?.outlineLevel || 0),
        hiddenRows: [2, 3, 4, 5, 6, 7, 8].map((row) => sheet.rows?.[row - 1]?.hidden === true),
        settings: sheet.outlineSettings || null,
      };
    };
    return {
      first: read(first),
      second: read(second),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupDataToolsSheets);


  const sheetGroupDataToolsSubtotal = await readSheetGroupSubtotalOutline();


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 8, 4)?.raw &&
    !window.__webSheetDevCellModel?.(second.name, 8, 4)?.raw,
    sheetGroupDataToolsSheets,
    { timeout: 5000 },
  );


  const sheetGroupDataToolsSubtotalUndo = await readSheetGroupSubtotalOutline();


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 8, 4)?.raw === "総計" &&
    window.__webSheetDevCellModel?.(second.name, 8, 4)?.raw === "総計",
    sheetGroupDataToolsSheets,
    { timeout: 5000 },
  );


  const sheetGroupDataToolsSubtotalRedo = await readSheetGroupSubtotalOutline();


  await page.locator('.row-header[data-row="2"]').click();


  await page.click('[data-ribbon-command="data-hide-detail"]');


  await page.waitForFunction(([first, second]) => {
    const firstSheet = window.__webSheetDevSheetModel?.(first.name) || {};
    const secondSheet = window.__webSheetDevSheetModel?.(second.name) || {};
    return firstSheet.rows?.[1]?.hidden === true && secondSheet.rows?.[1]?.hidden === true;
  }, sheetGroupDataToolsSheets, { timeout: 5000 });


  const sheetGroupDataToolsOutlineHidden = await readSheetGroupSubtotalOutline();


  await page.click('[data-ribbon-command="data-show-detail"]');


  await page.waitForFunction(([first, second]) => {
    const firstSheet = window.__webSheetDevSheetModel?.(first.name) || {};
    const secondSheet = window.__webSheetDevSheetModel?.(second.name) || {};
    return firstSheet.rows?.[1]?.hidden !== true && secondSheet.rows?.[1]?.hidden !== true;
  }, sheetGroupDataToolsSheets, { timeout: 5000 });


  const sheetGroupDataToolsOutlineShown = await readSheetGroupSubtotalOutline();


  await page.click('[data-ribbon-command="data-group"]');


  await page.waitForSelector("#dataCommandMenu.is-outline-group", { timeout: 5000 });


  await page.click('[data-data-menu-action="outline-settings"]');


  await page.waitForSelector('[data-data-dialog="outline-settings"]', { timeout: 5000 });


  await page.locator("[data-data-field='summaryBelow']").uncheck();


  await page.locator("[data-data-field='summaryRight']").uncheck();


  await page.click('[data-data-dialog-action="ok"]');


  await page.waitForFunction(([first, second]) => {
    const firstSheet = window.__webSheetDevSheetModel?.(first.name) || {};
    const secondSheet = window.__webSheetDevSheetModel?.(second.name) || {};
    return firstSheet.outlineSettings?.summaryBelow === false &&
      secondSheet.outlineSettings?.summaryBelow === false &&
      firstSheet.outlineSettings?.summaryRight === false &&
      secondSheet.outlineSettings?.summaryRight === false;
  }, sheetGroupDataToolsSheets, { timeout: 5000 });


  const sheetGroupDataToolsOutlineSettings = await readSheetGroupSubtotalOutline();


  await page.locator(".sheet-tab").first().click();


  await page.evaluate(([first, second]) => {
    const fill = (sheetName, rows) => {
      window.__webSheetDevSetActiveSheet?.(sheetName);
      rows.forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
    };
    fill(first.name, [
      [1, 8, "Item"], [1, 9, "Q1"], [1, 10, "Q2"],
      [2, 8, "Alpha"], [2, 9, 2], [2, 10, 3],
      [3, 8, "Beta"], [3, 9, 5], [3, 10, 7],
    ]);
    fill(second.name, [
      [1, 8, "Item"], [1, 9, "Q1"], [1, 10, "Q2"],
      [2, 8, "Alpha"], [2, 9, 20], [2, 10, 30],
      [3, 8, "Beta"], [3, 9, 50], [3, 10, 70],
    ]);
  }, sheetGroupDataToolsSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupDataToolsSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupDataToolsSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 8, 3, 10));


  await page.click('[data-ribbon-command="data-consolidate"]');


  await page.waitForSelector('[data-data-dialog="consolidate"]', { timeout: 5000 });


  await page.locator("[data-data-field='leftColumn']").check();


  await page.click('[data-data-dialog-action="ok"]');


  await page.waitForSelector('[data-data-dialog="consolidate"]', { state: "detached", timeout: 5000 });


  await page.waitForFunction(([first, second]) =>
    Number(window.__webSheetDevCellModel?.(first.name, 6, 9)?.raw) === 2 &&
    Number(window.__webSheetDevCellModel?.(first.name, 7, 10)?.raw) === 7 &&
    Number(window.__webSheetDevCellModel?.(second.name, 6, 9)?.raw) === 20 &&
    Number(window.__webSheetDevCellModel?.(second.name, 7, 10)?.raw) === 70,
    sheetGroupDataToolsSheets,
    { timeout: 5000 },
  );


  const readSheetGroupConsolidate = async () => page.evaluate(([first, second]) => {
    const read = (sheetInfo) => ({
      h: [5, 6, 7].map((row) => window.__webSheetDevCellModel?.(sheetInfo.name, row, 8)?.raw ?? ""),
      i: [5, 6, 7].map((row) => window.__webSheetDevCellModel?.(sheetInfo.name, row, 9)?.raw ?? ""),
      j: [5, 6, 7].map((row) => window.__webSheetDevCellModel?.(sheetInfo.name, row, 10)?.raw ?? ""),
    });
    return {
      first: read(first),
      second: read(second),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupDataToolsSheets);


  const sheetGroupDataToolsConsolidate = await readSheetGroupConsolidate();


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 6, 9)?.raw &&
    !window.__webSheetDevCellModel?.(second.name, 6, 9)?.raw,
    sheetGroupDataToolsSheets,
    { timeout: 5000 },
  );


  const sheetGroupDataToolsConsolidateUndo = await readSheetGroupConsolidate();


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    Number(window.__webSheetDevCellModel?.(first.name, 6, 9)?.raw) === 2 &&
    Number(window.__webSheetDevCellModel?.(second.name, 6, 9)?.raw) === 20,
    sheetGroupDataToolsSheets,
    { timeout: 5000 },
  );


  const sheetGroupDataToolsConsolidateRedo = await readSheetGroupConsolidate();


  await page.locator(".sheet-tab").first().click();


  await page.evaluate(([first, second]) => {
    const fill = (sheetName, rows) => {
      window.__webSheetDevSetActiveSheet?.(sheetName);
      rows.forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
    };
    fill(first.name, [[1, 12, "Value"], [2, 12, 2], [3, 12, 4], [4, 12, 6]]);
    fill(second.name, [[1, 12, "Value"], [2, 12, 20], [3, 12, 40], [4, 12, 60]]);
  }, sheetGroupDataToolsSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupDataToolsSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupDataToolsSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 12, 4, 12));


  await page.click('[data-ribbon-command="data-analysis"]');


  await page.waitForSelector('[data-data-dialog="data-analysis-picker"]', { timeout: 5000 });


  await page.locator("[data-data-analysis-picker]").selectOption("descriptive");


  await page.click('[data-data-dialog-action="data-analysis-picker-ok"]');


  await page.waitForSelector('[data-data-dialog="data-analysis"]', { timeout: 5000 });


  await page.locator("[data-data-field='labels']").check();


  await page.click('[data-data-dialog-action="ok"]');


  await page.waitForSelector('[data-data-dialog="data-analysis"]', { state: "detached", timeout: 5000 });


  await page.waitForFunction(([first, second]) =>
    Number(window.__webSheetDevCellModel?.(first.name, 2, 15)?.raw) === 4 &&
    Number(window.__webSheetDevCellModel?.(second.name, 2, 15)?.raw) === 40,
    sheetGroupDataToolsSheets,
    { timeout: 5000 },
  );


  const readSheetGroupDataAnalysis = async () => page.evaluate(([first, second]) => {
    const read = (sheetInfo) => ({
      n: [1, 2, 13].map((row) => window.__webSheetDevCellModel?.(sheetInfo.name, row, 14)?.raw ?? ""),
      o: [1, 2, 13].map((row) => window.__webSheetDevCellModel?.(sheetInfo.name, row, 15)?.raw ?? ""),
    });
    return {
      first: read(first),
      second: read(second),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupDataToolsSheets);


  const sheetGroupDataToolsAnalysis = await readSheetGroupDataAnalysis();


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 2, 15)?.raw &&
    !window.__webSheetDevCellModel?.(second.name, 2, 15)?.raw,
    sheetGroupDataToolsSheets,
    { timeout: 5000 },
  );


  const sheetGroupDataToolsAnalysisUndo = await readSheetGroupDataAnalysis();


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    Number(window.__webSheetDevCellModel?.(first.name, 2, 15)?.raw) === 4 &&
    Number(window.__webSheetDevCellModel?.(second.name, 2, 15)?.raw) === 40,
    sheetGroupDataToolsSheets,
    { timeout: 5000 },
  );


  const sheetGroupDataToolsAnalysisRedo = await readSheetGroupDataAnalysis();


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupDataToolsSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupDataToolsSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupDataToolsSheets[0].index}"]`, { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupDataToolsBaseCount, { timeout: 5000 });


  const sheetGroupDataTools = {
    flash: sheetGroupDataToolsFlash,
    textColumns: sheetGroupDataToolsTextColumns,
    removeDuplicates: sheetGroupDataToolsRemoveDuplicates,
    removeUndo: sheetGroupDataToolsRemoveUndo,
    removeRedo: sheetGroupDataToolsRemoveRedo,
    subtotal: sheetGroupDataToolsSubtotal,
    subtotalUndo: sheetGroupDataToolsSubtotalUndo,
    subtotalRedo: sheetGroupDataToolsSubtotalRedo,
    outlineHidden: sheetGroupDataToolsOutlineHidden,
    outlineShown: sheetGroupDataToolsOutlineShown,
    outlineSettings: sheetGroupDataToolsOutlineSettings,
    consolidate: sheetGroupDataToolsConsolidate,
    consolidateUndo: sheetGroupDataToolsConsolidateUndo,
    consolidateRedo: sheetGroupDataToolsConsolidateRedo,
    analysis: sheetGroupDataToolsAnalysis,
    analysisUndo: sheetGroupDataToolsAnalysisUndo,
    analysisRedo: sheetGroupDataToolsAnalysisRedo,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };



  const sheetGroupHyperlinkBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupHyperlinkBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupHyperlinkBaseCount, { timeout: 5000 });


  const sheetGroupHyperlinkSheets = await page.evaluate(() =>
    [...document.querySelectorAll(".sheet-tab")]
      .slice(-2)
      .map((button) => ({ name: button.textContent, index: button.dataset.sheetIndex })),
  );


  await page.evaluate(([first, second]) => {
    window.__webSheetDevSetActiveSheet?.(first.name);
    window.__webSheetDevSetCell?.(4, 4, "Before");
    window.__webSheetDevSetActiveSheet?.(second.name);
    window.__webSheetDevSetCell?.(4, 4, "Before");
  }, sheetGroupHyperlinkSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupHyperlinkSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupHyperlinkSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(4, 4, 4, 4));


  await page.click('.sheet-cell[data-address="D4"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu", { timeout: 5000 });


  await page.click('#cellContextMenu [data-cell-menu-action="link"]');


  await page.waitForSelector("#linkDialog", { timeout: 5000 });


  await page.locator('[data-link-field="text"]').fill("Grouped Link");


  await page.locator('[data-link-field="url"]').fill("example.com/grouped");


  await page.click('[data-link-dialog-action="ok"]');


  await page.waitForSelector("#linkDialog", { state: "detached", timeout: 5000 });


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 4, 4)?.hyperlink === "https://example.com/grouped" &&
    window.__webSheetDevCellModel?.(second.name, 4, 4)?.hyperlink === "https://example.com/grouped",
    sheetGroupHyperlinkSheets,
    { timeout: 5000 },
  );


  const sheetGroupHyperlinkApplied = await page.evaluate(([first, second]) => {
    const read = (sheetName) => {
      const cell = window.__webSheetDevCellModel?.(sheetName, 4, 4) || {};
      return {
        raw: cell.raw ?? "",
        hyperlink: cell.hyperlink || null,
        color: cell.css?.color || "",
        textDecoration: cell.css?.textDecoration || "",
      };
    };
    return {
      firstD4: read(first.name),
      secondD4: read(second.name),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupHyperlinkSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 4, 4)?.hyperlink &&
    !window.__webSheetDevCellModel?.(second.name, 4, 4)?.hyperlink,
    sheetGroupHyperlinkSheets,
    { timeout: 5000 },
  );


  const sheetGroupHyperlinkUndo = await page.evaluate(([first, second]) => {
    const read = (sheetName) => {
      const cell = window.__webSheetDevCellModel?.(sheetName, 4, 4) || {};
      return { raw: cell.raw ?? "", hyperlink: cell.hyperlink || null };
    };
    return {
      firstD4: read(first.name),
      secondD4: read(second.name),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupHyperlinkSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 4, 4)?.hyperlink === "https://example.com/grouped" &&
    window.__webSheetDevCellModel?.(second.name, 4, 4)?.hyperlink === "https://example.com/grouped",
    sheetGroupHyperlinkSheets,
    { timeout: 5000 },
  );


  const sheetGroupHyperlinkRedo = await page.evaluate(([first, second]) => ({
    firstD4: window.__webSheetDevCellModel?.(first.name, 4, 4)?.hyperlink || null,
    secondD4: window.__webSheetDevCellModel?.(second.name, 4, 4)?.hyperlink || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupHyperlinkSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupHyperlinkSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupHyperlinkSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(4, 4, 4, 4));


  await page.click('.sheet-cell[data-address="D4"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu", { timeout: 5000 });


  await page.click('#cellContextMenu [data-cell-menu-action="link"]');


  await page.waitForSelector("#linkDialog", { timeout: 5000 });


  await page.click('[data-link-dialog-action="remove"]');


  await page.waitForSelector("#linkDialog", { state: "detached", timeout: 5000 });


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 4, 4)?.hyperlink &&
    !window.__webSheetDevCellModel?.(second.name, 4, 4)?.hyperlink,
    sheetGroupHyperlinkSheets,
    { timeout: 5000 },
  );


  const sheetGroupHyperlinkRemove = await page.evaluate(([first, second]) => {
    const read = (sheetName) => {
      const cell = window.__webSheetDevCellModel?.(sheetName, 4, 4) || {};
      return { raw: cell.raw ?? "", hyperlink: cell.hyperlink || null };
    };
    return {
      firstD4: read(first.name),
      secondD4: read(second.name),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupHyperlinkSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 4, 4)?.hyperlink === "https://example.com/grouped" &&
    window.__webSheetDevCellModel?.(second.name, 4, 4)?.hyperlink === "https://example.com/grouped",
    sheetGroupHyperlinkSheets,
    { timeout: 5000 },
  );


  const sheetGroupHyperlinkRemoveUndo = await page.evaluate(([first, second]) => ({
    firstD4: window.__webSheetDevCellModel?.(first.name, 4, 4)?.hyperlink || null,
    secondD4: window.__webSheetDevCellModel?.(second.name, 4, 4)?.hyperlink || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupHyperlinkSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupHyperlinkSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupHyperlinkSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(4, 4, 4, 4));


  await page.click('[data-ribbon-tab="home"]');


  await page.click('[data-ribbon-command="clear-menu"]');


  await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });


  await page.click('[data-home-clear-action="hyperlinks"]');


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 4, 4)?.hyperlink &&
    !window.__webSheetDevCellModel?.(second.name, 4, 4)?.hyperlink,
    sheetGroupHyperlinkSheets,
    { timeout: 5000 },
  );


  const sheetGroupHyperlinkHomeClear = await page.evaluate(([first, second]) => {
    const read = (sheetName) => {
      const cell = window.__webSheetDevCellModel?.(sheetName, 4, 4) || {};
      return { raw: cell.raw ?? "", hyperlink: cell.hyperlink || null };
    };
    return {
      firstD4: read(first.name),
      secondD4: read(second.name),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupHyperlinkSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 4, 4)?.hyperlink === "https://example.com/grouped" &&
    window.__webSheetDevCellModel?.(second.name, 4, 4)?.hyperlink === "https://example.com/grouped",
    sheetGroupHyperlinkSheets,
    { timeout: 5000 },
  );


  const sheetGroupHyperlinkHomeClearUndo = await page.evaluate(([first, second]) => ({
    firstD4: window.__webSheetDevCellModel?.(first.name, 4, 4)?.hyperlink || null,
    secondD4: window.__webSheetDevCellModel?.(second.name, 4, 4)?.hyperlink || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupHyperlinkSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupHyperlinkSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupHyperlinkSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupHyperlinkSheets[0].index}"]`, { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupHyperlinkBaseCount, { timeout: 5000 });


  const sheetGroupHyperlink = {
    applied: sheetGroupHyperlinkApplied,
    undo: sheetGroupHyperlinkUndo,
    redo: sheetGroupHyperlinkRedo,
    remove: sheetGroupHyperlinkRemove,
    removeUndo: sheetGroupHyperlinkRemoveUndo,
    homeClear: sheetGroupHyperlinkHomeClear,
    homeClearUndo: sheetGroupHyperlinkHomeClearUndo,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };



  const sheetGroupHomeBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupHomeBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupHomeBaseCount, { timeout: 5000 });


  const sheetGroupHomeSheets = await page.evaluate(() =>
    [...document.querySelectorAll(".sheet-tab")]
      .slice(-2)
      .map((button) => ({ name: button.textContent, index: button.dataset.sheetIndex })),
  );


  await page.evaluate(([first, second]) => {
    window.__webSheetDevSetActiveSheet?.(first.name);
    window.__webSheetDevSetCell?.(1, 1, 1);
    window.__webSheetDevSetCell?.(2, 1, 2);
    window.__webSheetDevSetCell?.(2, 2, "FirstSeed");
    window.__webSheetDevSetActiveSheet?.(second.name);
    window.__webSheetDevSetCell?.(1, 1, 10);
    window.__webSheetDevSetCell?.(2, 1, 20);
    window.__webSheetDevSetCell?.(2, 2, "SecondSeed");
  }, sheetGroupHomeSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupHomeSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupHomeSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(2, 2, 4, 2));


  await page.click('[data-ribbon-tab="home"]');


  await page.click('[data-ribbon-command="fill-menu"]');


  await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });


  await page.click('[data-home-fill-action="down"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 3, 2)?.raw === "FirstSeed" &&
    window.__webSheetDevCellModel?.(first.name, 4, 2)?.raw === "FirstSeed" &&
    window.__webSheetDevCellModel?.(second.name, 3, 2)?.raw === "SecondSeed" &&
    window.__webSheetDevCellModel?.(second.name, 4, 2)?.raw === "SecondSeed",
    sheetGroupHomeSheets,
    { timeout: 5000 },
  );


  const sheetGroupHomeFill = await page.evaluate(([first, second]) => ({
    firstB: [2, 3, 4].map((row) => window.__webSheetDevCellModel?.(first.name, row, 2)?.raw ?? ""),
    secondB: [2, 3, 4].map((row) => window.__webSheetDevCellModel?.(second.name, row, 2)?.raw ?? ""),
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupHomeSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 3, 2)?.raw &&
    !window.__webSheetDevCellModel?.(second.name, 3, 2)?.raw,
    sheetGroupHomeSheets,
    { timeout: 5000 },
  );


  const sheetGroupHomeFillUndo = await page.evaluate(([first, second]) => ({
    firstB: [2, 3, 4].map((row) => window.__webSheetDevCellModel?.(first.name, row, 2)?.raw ?? ""),
    secondB: [2, 3, 4].map((row) => window.__webSheetDevCellModel?.(second.name, row, 2)?.raw ?? ""),
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupHomeSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 4, 2)?.raw === "FirstSeed" &&
    window.__webSheetDevCellModel?.(second.name, 4, 2)?.raw === "SecondSeed",
    sheetGroupHomeSheets,
    { timeout: 5000 },
  );


  const sheetGroupHomeFillRedo = await page.evaluate(([first, second]) => ({
    firstB: [2, 3, 4].map((row) => window.__webSheetDevCellModel?.(first.name, row, 2)?.raw ?? ""),
    secondB: [2, 3, 4].map((row) => window.__webSheetDevCellModel?.(second.name, row, 2)?.raw ?? ""),
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupHomeSheets);


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 1, 3, 1));


  await page.click('[data-ribbon-command="autosum"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 3, 1)?.formula === "=SUM(A1:A2)" &&
    window.__webSheetDevCellModel?.(second.name, 3, 1)?.formula === "=SUM(A1:A2)",
    sheetGroupHomeSheets,
    { timeout: 5000 },
  );


  const sheetGroupHomeAutoSum = await page.evaluate(([first, second]) => ({
    firstA3: {
      formula: window.__webSheetDevCellModel?.(first.name, 3, 1)?.formula || "",
      raw: window.__webSheetDevCellModel?.(first.name, 3, 1)?.raw ?? "",
    },
    secondA3: {
      formula: window.__webSheetDevCellModel?.(second.name, 3, 1)?.formula || "",
      raw: window.__webSheetDevCellModel?.(second.name, 3, 1)?.raw ?? "",
    },
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupHomeSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 3, 1)?.formula &&
    !window.__webSheetDevCellModel?.(second.name, 3, 1)?.formula,
    sheetGroupHomeSheets,
    { timeout: 5000 },
  );


  const sheetGroupHomeAutoSumUndo = await page.evaluate(([first, second]) => ({
    firstA3: window.__webSheetDevCellModel?.(first.name, 3, 1)?.formula || "",
    secondA3: window.__webSheetDevCellModel?.(second.name, 3, 1)?.formula || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupHomeSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 3, 1)?.formula === "=SUM(A1:A2)" &&
    window.__webSheetDevCellModel?.(second.name, 3, 1)?.formula === "=SUM(A1:A2)",
    sheetGroupHomeSheets,
    { timeout: 5000 },
  );


  const sheetGroupHomeAutoSumRedo = await page.evaluate(([first, second]) => ({
    firstA3: window.__webSheetDevCellModel?.(first.name, 3, 1)?.formula || "",
    secondA3: window.__webSheetDevCellModel?.(second.name, 3, 1)?.formula || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupHomeSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupHomeSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupHomeSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupHomeSheets[0].index}"]`, { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupHomeBaseCount, { timeout: 5000 });


  const sheetGroupHomeCommands = {
    fill: sheetGroupHomeFill,
    fillUndo: sheetGroupHomeFillUndo,
    fillRedo: sheetGroupHomeFillRedo,
    autoSum: sheetGroupHomeAutoSum,
    autoSumUndo: sheetGroupHomeAutoSumUndo,
    autoSumRedo: sheetGroupHomeAutoSumRedo,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };



  const sheetGroupTranslationBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupTranslationBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupTranslationBaseCount, { timeout: 5000 });


  const sheetGroupTranslationSheets = await page.evaluate(() =>
    [...document.querySelectorAll(".sheet-tab")]
      .slice(-2)
      .map((button) => ({ name: button.textContent, index: button.dataset.sheetIndex })),
  );


  await page.evaluate(([first, second]) => {
    window.__webSheetDevSetActiveSheet?.(first.name);
    window.__webSheetDevSetCell?.(2, 2, "氏名");
    window.__webSheetDevSetActiveSheet?.(second.name);
    window.__webSheetDevSetCell?.(2, 2, "年齢");
  }, sheetGroupTranslationSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupTranslationSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupTranslationSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(2, 2, 2, 2));


  await page.click('.sheet-cell[data-address="B2"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu", { timeout: 5000 });


  await page.click('#cellContextMenu [data-cell-menu-action="translate"]');


  await page.waitForSelector("#translationPane", { timeout: 5000 });


  const sheetGroupTranslationPane = await page.evaluate(() => ({
    title: document.querySelector("#translationPane .translation-pane-header strong")?.textContent || "",
    range: document.querySelector("#translationPane .translation-pane-header span")?.textContent || "",
    sourceText: document.querySelector("#translationPane .translation-pane-source")?.textContent || "",
    resultText: document.querySelector("#translationPane .translation-pane-result")?.textContent || "",
  }));


  await page.click('[data-translation-pane-action="insert-right"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 2, 3)?.raw === "姓名" &&
    window.__webSheetDevCellModel?.(second.name, 2, 3)?.raw === "年龄",
    sheetGroupTranslationSheets,
    { timeout: 5000 },
  );


  const sheetGroupTranslationInsertRight = await page.evaluate(([first, second]) => ({
    firstC2: window.__webSheetDevCellModel?.(first.name, 2, 3)?.raw ?? "",
    secondC2: window.__webSheetDevCellModel?.(second.name, 2, 3)?.raw ?? "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupTranslationSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 2, 3)?.raw &&
    !window.__webSheetDevCellModel?.(second.name, 2, 3)?.raw,
    sheetGroupTranslationSheets,
    { timeout: 5000 },
  );


  const sheetGroupTranslationInsertUndo = await page.evaluate(([first, second]) => ({
    firstC2: window.__webSheetDevCellModel?.(first.name, 2, 3)?.raw ?? "",
    secondC2: window.__webSheetDevCellModel?.(second.name, 2, 3)?.raw ?? "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupTranslationSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 2, 3)?.raw === "姓名" &&
    window.__webSheetDevCellModel?.(second.name, 2, 3)?.raw === "年龄",
    sheetGroupTranslationSheets,
    { timeout: 5000 },
  );


  const sheetGroupTranslationInsertRedo = await page.evaluate(([first, second]) => ({
    firstC2: window.__webSheetDevCellModel?.(first.name, 2, 3)?.raw ?? "",
    secondC2: window.__webSheetDevCellModel?.(second.name, 2, 3)?.raw ?? "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupTranslationSheets);


  await page.waitForSelector("#translationPane", { timeout: 5000 });


  await page.click('[data-translation-pane-action="replace"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 2, 2)?.raw === "姓名" &&
    window.__webSheetDevCellModel?.(second.name, 2, 2)?.raw === "年龄",
    sheetGroupTranslationSheets,
    { timeout: 5000 },
  );


  const sheetGroupTranslationReplace = await page.evaluate(([first, second]) => ({
    firstB2: window.__webSheetDevCellModel?.(first.name, 2, 2)?.raw ?? "",
    secondB2: window.__webSheetDevCellModel?.(second.name, 2, 2)?.raw ?? "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupTranslationSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 2, 2)?.raw === "氏名" &&
    window.__webSheetDevCellModel?.(second.name, 2, 2)?.raw === "年齢",
    sheetGroupTranslationSheets,
    { timeout: 5000 },
  );


  const sheetGroupTranslationReplaceUndo = await page.evaluate(([first, second]) => ({
    firstB2: window.__webSheetDevCellModel?.(first.name, 2, 2)?.raw ?? "",
    secondB2: window.__webSheetDevCellModel?.(second.name, 2, 2)?.raw ?? "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupTranslationSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupTranslationSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupTranslationSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupTranslationSheets[0].index}"]`, { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupTranslationBaseCount, { timeout: 5000 });


  const sheetGroupTranslation = {
    pane: sheetGroupTranslationPane,
    insertRight: sheetGroupTranslationInsertRight,
    insertUndo: sheetGroupTranslationInsertUndo,
    insertRedo: sheetGroupTranslationInsertRedo,
    replace: sheetGroupTranslationReplace,
    replaceUndo: sheetGroupTranslationReplaceUndo,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };



  const sheetGroupQuickAnalysisBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupQuickAnalysisBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupQuickAnalysisBaseCount, { timeout: 5000 });


  const sheetGroupQuickAnalysisSheets = await page.evaluate(() =>
    [...document.querySelectorAll(".sheet-tab")]
      .slice(-2)
      .map((button) => ({ name: button.textContent, index: button.dataset.sheetIndex })),
  );


  await page.evaluate(([first, second]) => {
    window.__webSheetDevSetActiveSheet?.(first.name);
    [1, 2, 3].forEach((value, index) => window.__webSheetDevSetCell?.(index + 1, 1, value));
    window.__webSheetDevSetActiveSheet?.(second.name);
    [10, 20, 30].forEach((value, index) => window.__webSheetDevSetCell?.(index + 1, 1, value));
  }, sheetGroupQuickAnalysisSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupQuickAnalysisSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupQuickAnalysisSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 1, 3, 1));


  await page.click('.sheet-cell[data-address="A1"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu", { timeout: 5000 });


  await page.click('#cellContextMenu [data-cell-menu-action="quick-analysis"]');


  await page.waitForSelector("#quickAnalysisPopup", { timeout: 5000 });


  const sheetGroupQuickAnalysisKeyboard = await page.evaluate(() => ({
    focusedBefore: document.activeElement?.dataset?.quickAnalysisAction || document.activeElement?.dataset?.quickAnalysisTab || "",
  }));


  await page.keyboard.press("ArrowRight");


  Object.assign(sheetGroupQuickAnalysisKeyboard, await page.evaluate(() => ({
    focusedAfterRight: document.activeElement?.dataset?.quickAnalysisAction || document.activeElement?.dataset?.quickAnalysisTab || "",
  })));


  await page.click('#quickAnalysisPopup [data-quick-analysis-action="format-data-bars"]');


  await page.waitForFunction(([first, second]) =>
    Boolean(window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisBar) &&
    Boolean(window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisBar),
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisBars = await page.evaluate(([first, second]) => ({
    firstA3: window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisBar || null,
    secondA3: window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisBar || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupQuickAnalysisSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisBar &&
    !window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisBar,
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisBarsUndo = await page.evaluate(([first, second]) => ({
    firstA3: window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisBar || null,
    secondA3: window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisBar || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupQuickAnalysisSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    Boolean(window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisBar) &&
    Boolean(window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisBar),
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisBarsRedo = await page.evaluate(([first, second]) => ({
    firstA3: window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisBar || null,
    secondA3: window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisBar || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupQuickAnalysisSheets);


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 1, 3, 1));


  await page.click('.sheet-cell[data-address="A1"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu", { timeout: 5000 });


  await page.click('#cellContextMenu [data-cell-menu-action="quick-analysis"]');


  await page.waitForSelector("#quickAnalysisPopup", { timeout: 5000 });


  await page.click('#quickAnalysisPopup [data-quick-analysis-action="format-clear"]');


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisBar &&
    !window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisBar,
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisClear = await page.evaluate(([first, second]) => ({
    firstA3: window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisBar || null,
    secondA3: window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisBar || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupQuickAnalysisSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    Boolean(window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisBar) &&
    Boolean(window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisBar),
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisClearUndo = await page.evaluate(([first, second]) => ({
    firstA3: window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisBar || null,
    secondA3: window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisBar || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupQuickAnalysisSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupQuickAnalysisSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupQuickAnalysisSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 1, 3, 1));


  await page.click('.sheet-cell[data-address="A1"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu", { timeout: 5000 });


  await page.click('#cellContextMenu [data-cell-menu-action="quick-analysis"]');


  await page.waitForSelector("#quickAnalysisPopup", { timeout: 5000 });


  await page.click('#quickAnalysisPopup [data-quick-analysis-action="format-color-scale"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 3, 1)?.css?.backgroundColor === "#63BE7B" &&
    window.__webSheetDevCellModel?.(second.name, 3, 1)?.css?.backgroundColor === "#63BE7B",
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisColorScale = await page.evaluate(([first, second]) => ({
    firstA3Background: window.__webSheetDevCellModel?.(first.name, 3, 1)?.css?.backgroundColor || "",
    secondA3Background: window.__webSheetDevCellModel?.(second.name, 3, 1)?.css?.backgroundColor || "",
    firstA3Bar: window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisBar || null,
    secondA3Bar: window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisBar || null,
  }), sheetGroupQuickAnalysisSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    Boolean(window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisBar) &&
    Boolean(window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisBar) &&
    !window.__webSheetDevCellModel?.(first.name, 3, 1)?.css?.backgroundColor &&
    !window.__webSheetDevCellModel?.(second.name, 3, 1)?.css?.backgroundColor,
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisColorScaleUndo = await page.evaluate(([first, second]) => ({
    firstA3Background: window.__webSheetDevCellModel?.(first.name, 3, 1)?.css?.backgroundColor || "",
    secondA3Background: window.__webSheetDevCellModel?.(second.name, 3, 1)?.css?.backgroundColor || "",
    firstA3Bar: window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisBar || null,
    secondA3Bar: window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisBar || null,
  }), sheetGroupQuickAnalysisSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 3, 1)?.css?.backgroundColor === "#63BE7B" &&
    window.__webSheetDevCellModel?.(second.name, 3, 1)?.css?.backgroundColor === "#63BE7B",
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisColorScaleRedo = await page.evaluate(([first, second]) => ({
    firstA3Background: window.__webSheetDevCellModel?.(first.name, 3, 1)?.css?.backgroundColor || "",
    secondA3Background: window.__webSheetDevCellModel?.(second.name, 3, 1)?.css?.backgroundColor || "",
    firstA3Bar: window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisBar || null,
    secondA3Bar: window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisBar || null,
  }), sheetGroupQuickAnalysisSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupQuickAnalysisSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupQuickAnalysisSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 1, 3, 1));


  await page.click('.sheet-cell[data-address="A1"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu", { timeout: 5000 });


  await page.click('#cellContextMenu [data-cell-menu-action="quick-analysis"]');


  await page.waitForSelector("#quickAnalysisPopup", { timeout: 5000 });


  await page.click('#quickAnalysisPopup [data-quick-analysis-action="format-icon-set"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisIcon?.shape === "up" &&
    window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisIcon?.shape === "up",
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisIconSet = await page.evaluate(([first, second]) => ({
    firstA3Icon: window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisIcon || null,
    secondA3Icon: window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisIcon || null,
  }), sheetGroupQuickAnalysisSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisIcon &&
    !window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisIcon,
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisIconUndo = await page.evaluate(([first, second]) => ({
    firstA3Icon: window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisIcon || null,
    secondA3Icon: window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisIcon || null,
  }), sheetGroupQuickAnalysisSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisIcon?.shape === "up" &&
    window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisIcon?.shape === "up",
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisIconRedo = await page.evaluate(([first, second]) => ({
    firstA3Icon: window.__webSheetDevCellModel?.(first.name, 3, 1)?.quickAnalysisIcon || null,
    secondA3Icon: window.__webSheetDevCellModel?.(second.name, 3, 1)?.quickAnalysisIcon || null,
  }), sheetGroupQuickAnalysisSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupQuickAnalysisSheets[0].index}"]`);


  await page.evaluate(([first, second]) => {
    window.__webSheetDevSetActiveSheet?.(first.name);
    ["X", "X", "Y"].forEach((value, index) => window.__webSheetDevSetCell?.(index + 1, 2, value));
    window.__webSheetDevSetActiveSheet?.(second.name);
    ["P", "Q", "P"].forEach((value, index) => window.__webSheetDevSetCell?.(index + 1, 2, value));
  }, sheetGroupQuickAnalysisSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupQuickAnalysisSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupQuickAnalysisSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 2, 3, 2));


  await page.click('.sheet-cell[data-address="B1"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu", { timeout: 5000 });


  await page.click('#cellContextMenu [data-cell-menu-action="quick-analysis"]');


  await page.waitForSelector("#quickAnalysisPopup", { timeout: 5000 });


  await page.click('#quickAnalysisPopup [data-quick-analysis-action="format-duplicates"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 1, 2)?.quickAnalysisStyle?.background &&
    window.__webSheetDevCellModel?.(first.name, 2, 2)?.quickAnalysisStyle?.background &&
    !window.__webSheetDevCellModel?.(first.name, 3, 2)?.quickAnalysisStyle?.background &&
    window.__webSheetDevCellModel?.(second.name, 1, 2)?.quickAnalysisStyle?.background &&
    !window.__webSheetDevCellModel?.(second.name, 2, 2)?.quickAnalysisStyle?.background &&
    window.__webSheetDevCellModel?.(second.name, 3, 2)?.quickAnalysisStyle?.background,
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisDuplicates = await page.evaluate(([first, second]) => ({
    first: [1, 2, 3].map((row) => Boolean(window.__webSheetDevCellModel?.(first.name, row, 2)?.quickAnalysisStyle?.background)),
    second: [1, 2, 3].map((row) => Boolean(window.__webSheetDevCellModel?.(second.name, row, 2)?.quickAnalysisStyle?.background)),
  }), sheetGroupQuickAnalysisSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 1, 2)?.quickAnalysisStyle?.background &&
    !window.__webSheetDevCellModel?.(first.name, 2, 2)?.quickAnalysisStyle?.background &&
    !window.__webSheetDevCellModel?.(second.name, 1, 2)?.quickAnalysisStyle?.background &&
    !window.__webSheetDevCellModel?.(second.name, 3, 2)?.quickAnalysisStyle?.background,
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisDuplicatesUndo = await page.evaluate(([first, second]) => ({
    first: [1, 2, 3].map((row) => Boolean(window.__webSheetDevCellModel?.(first.name, row, 2)?.quickAnalysisStyle?.background)),
    second: [1, 2, 3].map((row) => Boolean(window.__webSheetDevCellModel?.(second.name, row, 2)?.quickAnalysisStyle?.background)),
  }), sheetGroupQuickAnalysisSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 1, 2)?.quickAnalysisStyle?.background &&
    window.__webSheetDevCellModel?.(second.name, 3, 2)?.quickAnalysisStyle?.background,
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisDuplicatesRedo = await page.evaluate(([first, second]) => ({
    first: [1, 2, 3].map((row) => Boolean(window.__webSheetDevCellModel?.(first.name, row, 2)?.quickAnalysisStyle?.background)),
    second: [1, 2, 3].map((row) => Boolean(window.__webSheetDevCellModel?.(second.name, row, 2)?.quickAnalysisStyle?.background)),
  }), sheetGroupQuickAnalysisSheets);
  Object.assign(scenarioState, { context, page, messages, browserDialogs, result, dataRibbonLayout, coordinateHitTest, delayedLazyLoad, stickyHeaders, sheetMenu, sheetTabColor, sheetRename, sheetTabDirectActions, sheetMoveCopy, sheetCodeDialog, sheetSelectAll, sheetHidden, sheetUnhideDialog, sheetUnhidden, sheetUnhideUndo, sheetUnhideRedo, sheetGroupActions, sheetGroupMoveCopy, sheetGroupCellEditing, sheetGroupStructure, sheetGroupFormatDialog, sheetGroupDataValidation, sheetGroupDataFilterSort, sheetGroupDataTools, sheetGroupHyperlink, sheetGroupHomeCommands, sheetGroupTranslation, sheetGroupQuickAnalysisBaseCount, sheetGroupQuickAnalysisSheets, sheetGroupQuickAnalysisKeyboard, sheetGroupQuickAnalysisBars, sheetGroupQuickAnalysisBarsUndo, sheetGroupQuickAnalysisBarsRedo, sheetGroupQuickAnalysisClear, sheetGroupQuickAnalysisClearUndo, sheetGroupQuickAnalysisColorScale, sheetGroupQuickAnalysisColorScaleUndo, sheetGroupQuickAnalysisColorScaleRedo, sheetGroupQuickAnalysisIconSet, sheetGroupQuickAnalysisIconUndo, sheetGroupQuickAnalysisIconRedo, sheetGroupQuickAnalysisDuplicates, sheetGroupQuickAnalysisDuplicatesUndo, sheetGroupQuickAnalysisDuplicatesRedo });
}

export {
  runSetupAndWorkbook,
};
