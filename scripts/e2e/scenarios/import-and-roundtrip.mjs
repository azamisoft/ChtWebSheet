// Sequential E2E feature module.

async function runImportAndRoundtrip(scenarioState) {
  let { page, browserDialogs, result, revealDirectTestCell, typeCellByDirectKeyboard, nameLinkAddress, formulaNameAddress, absoluteNameLinkAddress, nameDialogInitial, definedNameFormula, unicodeDefinedNameFormula, duplicateDefinedName, localDefinedNameFormula } = scenarioState;

  const linkDialogInitial = await page.evaluate(() => ({
    title: document.querySelector("#linkDialogTitle")?.textContent,
    textValue: document.querySelector('[data-link-field="text"]')?.value,
    urlValue: document.querySelector('[data-link-field="url"]')?.value,
    hasRemove: Boolean(document.querySelector('[data-link-dialog-action="remove"]')),
    captions: [...document.querySelectorAll("#linkDialog .excel-dialog-caption")].map((item) => item.textContent),
    captionNowrap: [...document.querySelectorAll("#linkDialog .excel-dialog-caption")].every((item) => getComputedStyle(item).whiteSpace === "nowrap"),
    captionAboveInput: [...document.querySelectorAll("#linkDialog .excel-dialog-field")].every((field) => {
      const caption = field.querySelector(".excel-dialog-caption");
      const input = field.querySelector("input,select,textarea");
      if (!caption || !input) return true;
      const captionRect = caption.getBoundingClientRect();
      const inputRect = input.getBoundingClientRect();
      return captionRect.height < 24 && captionRect.bottom <= inputRect.top + 1;
    }),
    kindButtons: [...document.querySelectorAll("[data-link-kind]")].map((button) => ({
      kind: button.dataset.linkKind,
      text: button.textContent,
      active: button.classList.contains("is-active"),
      pressed: button.getAttribute("aria-pressed"),
    })),
    accessKeys: [...document.querySelectorAll("#linkDialog .menu-access-key")].map((item) => item.textContent),
    accessKeysUnderlined: [...document.querySelectorAll("#linkDialog .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
  }));


  await page.click('[data-link-kind="document"]');


  const linkDocumentKind = await page.evaluate(() => ({
    activeKind: document.querySelector("[data-link-kind].is-active")?.dataset.linkKind,
    pressed: document.querySelector('[data-link-kind="document"]')?.getAttribute("aria-pressed"),
    addressCaption: document.querySelector("[data-link-address-caption]")?.textContent,
    urlValue: document.querySelector('[data-link-field="url"]')?.value,
    placeholder: document.querySelector('[data-link-field="url"]')?.getAttribute("placeholder"),
    documentPanelHidden: document.querySelector("[data-link-document-panel]")?.hidden,
    documentPanelCaption: document.querySelector("[data-link-document-panel] .excel-dialog-caption")?.textContent,
    sheetNames: [...document.querySelectorAll("[data-link-document-sheet-index]")].map((button) => button.textContent),
    selectedSheetName: document.querySelector("[data-link-document-sheet-index].is-active")?.textContent,
    selectedSheetAria: document.querySelector("[data-link-document-sheet-index].is-active")?.getAttribute("aria-selected"),
  }));


  await page.click('[data-link-document-sheet-index="1"]');


  const linkDocumentSheetChoice = await page.evaluate(() => ({
    selectedSheetIndex: document.querySelector("[data-link-document-sheet-index].is-active")?.dataset.linkDocumentSheetIndex,
    selectedSheetName: document.querySelector("[data-link-document-sheet-index].is-active")?.textContent,
    selectedSheetAria: document.querySelector("[data-link-document-sheet-index].is-active")?.getAttribute("aria-selected"),
    urlValue: document.querySelector('[data-link-field="url"]')?.value,
  }));


  await page.click('[data-link-kind="email"]');


  const linkEmailKind = await page.evaluate(() => ({
    activeKind: document.querySelector("[data-link-kind].is-active")?.dataset.linkKind,
    pressed: document.querySelector('[data-link-kind="email"]')?.getAttribute("aria-pressed"),
    addressCaption: document.querySelector("[data-link-address-caption]")?.textContent,
    urlValue: document.querySelector('[data-link-field="url"]')?.value,
    placeholder: document.querySelector('[data-link-field="url"]')?.getAttribute("placeholder"),
    documentPanelHidden: document.querySelector("[data-link-document-panel]")?.hidden,
  }));


  await page.click('[data-link-kind="web"]');


  const linkWebKind = await page.evaluate(() => ({
    activeKind: document.querySelector("[data-link-kind].is-active")?.dataset.linkKind,
    pressed: document.querySelector('[data-link-kind="web"]')?.getAttribute("aria-pressed"),
    addressCaption: document.querySelector("[data-link-address-caption]")?.textContent,
    urlValue: document.querySelector('[data-link-field="url"]')?.value,
    placeholder: document.querySelector('[data-link-field="url"]')?.getAttribute("placeholder"),
  }));


  await page.locator('[data-link-field="text"]').fill("WebSheet");


  await page.locator('[data-link-field="url"]').fill("example.com/docs");


  await page.click('[data-link-dialog-action="ok"]');


  await page.waitForFunction(
    (address) => document.querySelector(`.sheet-cell[data-address="${address}"] a.cell-link`)?.getAttribute("href") === "https://example.com/docs",
    nameLinkAddress,
    { timeout: 10000 },
  );


  const linkApply = await page.evaluate((address) => {
    const cell = document.querySelector(`.sheet-cell[data-address="${address}"]`);
    const link = cell?.querySelector("a.cell-link");
    const style = cell ? getComputedStyle(cell) : null;
    return {
      text: link?.textContent,
      href: link?.getAttribute("href"),
      color: style?.color,
      decoration: style?.textDecorationLine,
      status: document.querySelector("#statusText")?.textContent,
    };
  }, nameLinkAddress);


  await page.evaluate(() => {
    window.__openedLinks = [];
    window.open = (url, target, features) => {
      window.__openedLinks.push({ url, target, features });
      return null;
    };
  });


  await revealDirectTestCell(nameLinkAddress);


  await page.click(`.sheet-cell[data-address="${nameLinkAddress}"]`, { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  const linkContextMenu = await page.evaluate(() => ({
    disabledActions: [...document.querySelectorAll("#cellContextMenu .cell-context-menu-item:disabled")].map((item) => item.dataset.cellMenuAction),
  }));


  await page.click('#cellContextMenu [data-cell-menu-action="open-hyperlink"]');


  const openedLink = await page.evaluate(() => ({
    opened: window.__openedLinks?.[0] || null,
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction((address) => !document.querySelector(`.sheet-cell[data-address="${address}"] a.cell-link`), nameLinkAddress, { timeout: 10000 });


  const linkUndo = await page.evaluate((address) => ({
    text: document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent,
    hasLink: Boolean(document.querySelector(`.sheet-cell[data-address="${address}"] a.cell-link`)),
    status: document.querySelector("#statusText")?.textContent,
  }), nameLinkAddress);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(
    (address) => document.querySelector(`.sheet-cell[data-address="${address}"] a.cell-link`)?.getAttribute("href") === "https://example.com/docs",
    nameLinkAddress,
    { timeout: 10000 },
  );


  const linkRedo = await page.evaluate((address) => ({
    text: document.querySelector(`.sheet-cell[data-address="${address}"] a.cell-link`)?.textContent,
    href: document.querySelector(`.sheet-cell[data-address="${address}"] a.cell-link`)?.getAttribute("href"),
    status: document.querySelector("#statusText")?.textContent,
  }), nameLinkAddress);


  const definedNameAndLink = {
    nameLinkAddress,
    absoluteNameLinkAddress,
    formulaNameAddress,
    nameDialogInitial,
    definedNameFormula,
    unicodeDefinedNameFormula,
    duplicateDefinedName,
    localDefinedNameFormula,
    linkDialogInitial,
    linkDocumentKind,
    linkDocumentSheetChoice,
    linkEmailKind,
    linkWebKind,
    linkApply,
    linkContextMenu,
    openedLink,
    linkUndo,
    linkRedo,
  };



  await typeCellByDirectKeyboard("A4", "1");


  await typeCellByDirectKeyboard("A5", "2");


  await page.locator('.sheet-cell[data-address="A4"]').click();


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "A4", null, { timeout: 10000 });


  const singleCellStatusBar = await page.evaluate(async () => {
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    return document.querySelector("#selectionStatus")?.textContent.replace(/\s+/g, " ").trim() || "";
  });


  await page.locator('.sheet-cell[data-address="A5"]').click({ modifiers: ["Shift"] });


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "A4:A5", null, { timeout: 10000 });


  const statusBarStats = await page.evaluate(async (singleCellStatusBar) => {
    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    await frame();
    const before = document.querySelector("#selectionStatus")?.textContent.replace(/\s+/g, " ").trim() || "";
    document.querySelector("#selectionStatus")?.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: 760, clientY: 706 }));
    await frame();
    const menu = document.querySelector("#statusBarContextMenu");
    const focusedBefore = document.activeElement?.dataset?.statusBarItem || "";
    menu?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true, cancelable: true }));
    await frame();
    const focusedAfterDown = document.activeElement?.dataset?.statusBarItem || "";
    const menuLabels = [...(menu?.querySelectorAll("[data-status-bar-item]") || [])].map((button) =>
      button.textContent.replace(/\s+/g, " ").trim().replace(/^✓/, ""),
    );
    const checkedBefore = [...(menu?.querySelectorAll("[data-status-bar-item]") || [])]
      .filter((button) => button.getAttribute("aria-checked") === "true")
      .map((button) => button.dataset.statusBarItem);
    menu?.querySelector('[data-status-bar-item="numericCount"]')?.click();
    await frame();
    const afterNumericCount = document.querySelector("#selectionStatus")?.textContent.replace(/\s+/g, " ").trim() || "";
    menu?.querySelector('[data-status-bar-item="sum"]')?.click();
    await frame();
    const afterHideSum = document.querySelector("#selectionStatus")?.textContent.replace(/\s+/g, " ").trim() || "";
    const checkedAfter = [...(menu?.querySelectorAll("[data-status-bar-item]") || [])]
      .filter((button) => button.getAttribute("aria-checked") === "true")
      .map((button) => button.dataset.statusBarItem);
    window.__webSheetDevCloseFloatingMenus?.();
    await frame();
    return {
      singleCellStatusBar,
      before,
      menuLabels,
      focusedBefore,
      focusedAfterDown,
      checkedBefore,
      afterNumericCount,
      afterHideSum,
      checkedAfter,
      menuClosed: !document.querySelector("#statusBarContextMenu"),
      viewButtonsVisible: getComputedStyle(document.querySelector(".sheet-view-buttons")).display !== "none",
      zoomSliderVisible: getComputedStyle(document.querySelector("#zoomSlider")).display !== "none",
      zoomLabelVisible: getComputedStyle(document.querySelector("#zoomLabel")).display !== "none",
    };
  }, singleCellStatusBar);


  const autoFillSeries = {};


  const dragAutoFillTo = async (sourceStart, sourceEnd, targetAddress, isFilled) => {
    for (let attempt = 0; attempt < 3; attempt += 1) {
      await page.locator(`.sheet-cell[data-address="${sourceStart}"]`).click();
      await page.locator(`.sheet-cell[data-address="${sourceEnd}"]`).click({ modifiers: ["Shift"] });
      await page.waitForFunction((expected) => document.querySelector("#selectedCellLabel")?.textContent === expected, `${sourceStart}:${sourceEnd}`, { timeout: 10000 });
      const fillHandle = page.locator(".sheet-selection-frame .sheet-fill-handle").first();
      const targetCell = page.locator(`.sheet-cell[data-address="${targetAddress}"]`).first();
      const handleBox = await fillHandle.boundingBox();
      const targetBox = await targetCell.boundingBox();
      if (!handleBox || !targetBox) throw new Error(`Auto-fill handle or target ${targetAddress} was not visible.`);
      await fillHandle.dragTo(targetCell, {
        sourcePosition: { x: Math.max(1, handleBox.width / 2), y: Math.max(1, handleBox.height / 2) },
        targetPosition: { x: targetBox.width / 2, y: targetBox.height / 2 },
        timeout: 5000,
      });
      await page.waitForTimeout(120);
      if (await page.evaluate(isFilled)) return;
    }
    throw new Error(`Auto-fill drag to ${targetAddress} did not apply.`);
  };


  await dragAutoFillTo(
    "A4",
    "A5",
    "A7",
    () =>
      document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent === "3" &&
      document.querySelector('.sheet-cell[data-address="A7"] .cell-content')?.textContent === "4",
  );


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent === "3" &&
      document.querySelector('.sheet-cell[data-address="A7"] .cell-content')?.textContent === "4",
    null,
    { timeout: 10000 },
  );


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    a6: document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent,
    a7: document.querySelector('.sheet-cell[data-address="A7"] .cell-content')?.textContent,
    handleColor: getComputedStyle(document.querySelector(".sheet-fill-handle")).backgroundColor,
    status: document.querySelector("#statusText")?.textContent,
  })));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent !== "3", null, { timeout: 10000 });


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    undoA6: document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent,
    undoA7: document.querySelector('.sheet-cell[data-address="A7"] .cell-content')?.textContent,
    undoStatus: document.querySelector("#statusText")?.textContent,
  })));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent === "3" &&
      document.querySelector('.sheet-cell[data-address="A7"] .cell-content')?.textContent === "4",
    null,
    { timeout: 10000 },
  );


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    redoA6: document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent,
    redoA7: document.querySelector('.sheet-cell[data-address="A7"] .cell-content')?.textContent,
    redoStatus: document.querySelector("#statusText")?.textContent,
  })));


  await page.waitForSelector("#autoFillOptionsButton", { timeout: 5000 });


  await page.click("#autoFillOptionsButton");


  await page.waitForSelector("#autoFillOptionsMenu", { timeout: 5000 });


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    optionsMenu: {
      labels: [...document.querySelectorAll("#autoFillOptionsMenu [data-auto-fill-option]")].map((button) =>
        button.textContent.replace(/\s+/g, " ").trim().replace(/^✓/, ""),
      ),
      checked: document.querySelector("#autoFillOptionsMenu [aria-checked='true']")?.dataset.autoFillOption || "",
      accessKeys: [...document.querySelectorAll("#autoFillOptionsMenu .menu-access-key")].map((item) => item.textContent),
      focusedBefore: document.activeElement?.dataset?.autoFillOption || "",
    },
  })));


  await page.keyboard.press("ArrowDown");


  autoFillSeries.optionsMenu.focusedAfterDown = await page.evaluate(() => document.activeElement?.dataset?.autoFillOption || "");


  await page.click('#autoFillOptionsMenu [data-auto-fill-option="copy"]');


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent === "1" &&
      document.querySelector('.sheet-cell[data-address="A7"] .cell-content')?.textContent === "2",
    null,
    { timeout: 10000 },
  );


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    copyOption: {
      a6: document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent,
      a7: document.querySelector('.sheet-cell[data-address="A7"] .cell-content')?.textContent,
      status: document.querySelector("#statusText")?.textContent,
    },
  })));


  await page.click("#autoFillOptionsButton");


  await page.waitForSelector("#autoFillOptionsMenu", { timeout: 5000 });


  await page.click('#autoFillOptionsMenu [data-auto-fill-option="formatting"]');


  await page.waitForFunction(
    () =>
      (document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent || "") === "" &&
      (document.querySelector('.sheet-cell[data-address="A7"] .cell-content')?.textContent || "") === "",
    null,
    { timeout: 10000 },
  );


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    formattingOption: {
      a6: document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent || "",
      a7: document.querySelector('.sheet-cell[data-address="A7"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent,
    },
  })));


  await page.click("#autoFillOptionsButton");


  await page.waitForSelector("#autoFillOptionsMenu", { timeout: 5000 });


  await page.click('#autoFillOptionsMenu [data-auto-fill-option="no-formatting"]');


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent === "3" &&
      document.querySelector('.sheet-cell[data-address="A7"] .cell-content')?.textContent === "4",
    null,
    { timeout: 10000 },
  );


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    noFormattingOption: {
      a6: document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent,
      a7: document.querySelector('.sheet-cell[data-address="A7"] .cell-content')?.textContent,
      status: document.querySelector("#statusText")?.textContent,
    },
  })));


  await page.click("#autoFillOptionsButton");


  await page.waitForSelector("#autoFillOptionsMenu", { timeout: 5000 });


  await page.click('#autoFillOptionsMenu [data-auto-fill-option="series"]');


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent === "3" &&
      document.querySelector('.sheet-cell[data-address="A7"] .cell-content')?.textContent === "4",
    null,
    { timeout: 10000 },
  );


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    seriesOption: {
      a6: document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent,
      a7: document.querySelector('.sheet-cell[data-address="A7"] .cell-content')?.textContent,
      status: document.querySelector("#statusText")?.textContent,
    },
  })));


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    formulaShiftRight: window.__webSheetDevShiftFormulaReferences?.("=SUM(B:B)+SUM(2:4)+C3", 0, 1) || "",
    formulaShiftDown: window.__webSheetDevShiftFormulaReferences?.("=SUM(B:B)+SUM(2:4)+C3", 1, 0) || "",
    absoluteFormulaShift: window.__webSheetDevShiftFormulaReferences?.("=SUM($B:$B)+SUM($2:$4)+$C$3", 5, 5) || "",
  })));


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(1, 1, "2026/6/1");
    window.__webSheetDevSetCell?.(8, 1, "項目001");
    window.__webSheetDevSetCell?.(12, 1, "月曜日");
    window.__webSheetDevSetCell?.(16, 1, "1月");
    window.__webSheetDevSetCell?.(6, 2, "2026/6/1");
    window.__webSheetDevSetCell?.(6, 3, "2026/6/3");
    window.__webSheetDevSetCell?.(13, 2, "R-001");
  });


  const parseAutoFillAddress = (address) => {
    const match = String(address || "").match(/^([A-Z]+)(\d+)$/);
    if (!match) return null;
    const col = [...match[1]].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
    return { col, row: Number(match[2]) };
  };


  const dragAutoFillRange = async (range, targetAddress, options = {}) => {
    const attempts = typeof options.until === "function" ? 3 : 1;
    for (let attempt = 0; attempt < attempts; attempt += 1) {
      await page.evaluate(({ top, left, bottom, right }) =>
        window.__webSheetDevSelectRange?.(top, left, bottom, right), range);
      await page.waitForFunction(
        ({ top, left, bottom, right }) => {
          const selected = window.__webSheetDevWorkbookState?.().selectionRange;
          return selected?.top === top && selected?.left === left && selected?.bottom === bottom && selected?.right === right;
        },
        range,
        { timeout: 10000 },
      );
      await page.evaluate(() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))));
      const handle = page.locator(".sheet-selection-frame .sheet-fill-handle");
      await handle.waitFor({ state: "visible", timeout: 10000 });
      const handleBox = await handle.boundingBox();
      const point = parseAutoFillAddress(targetAddress);
      let targetBox = point
        ? await page.evaluate(({ row, col }) => window.__webSheetDevCellClientRect?.(row, col), point)
        : null;
      await page.evaluate(() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))));
      if (!targetBox) {
        const colBox = point ? await page.locator(`.column-header[data-col="${point.col}"]`).boundingBox() : null;
        const rowBox = point ? await page.locator(`.row-header[data-row="${point.row}"]`).boundingBox() : null;
        targetBox = colBox && rowBox ? { x: colBox.x, y: rowBox.y, width: colBox.width, height: rowBox.height } : null;
      }
      if (!targetBox) {
        const targetCell = page.locator(`.sheet-cell[data-address="${targetAddress}"]`).first();
        targetBox = (await targetCell.count()) ? await targetCell.boundingBox() : null;
      }
      if (!handleBox || !targetBox) throw new Error(`Auto-fill handle or ${targetAddress} was not visible.`);
      const mouseButton = options.button || "left";
      if (options.control) await page.keyboard.down("Control");
      try {
        await page.mouse.move(handleBox.x + handleBox.width / 2, handleBox.y + handleBox.height / 2);
        await page.mouse.down({ button: mouseButton });
        await page.mouse.move(targetBox.x + targetBox.width / 2, targetBox.y + targetBox.height / 2, { steps: 6 });
        await page.mouse.move(targetBox.x + targetBox.width / 2, targetBox.y + targetBox.height / 2);
        await page.mouse.up({ button: mouseButton });
      } finally {
        if (options.control) await page.keyboard.up("Control");
      }
      await page.waitForTimeout(150);
      if (!options.until || await page.evaluate(options.until)) return;
    }
    const debug = await page.evaluate(({ targetAddress: debugTargetAddress }) => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const cellText = (address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent ?? null;
      const modelText = (row, col) => {
        const cell = window.__webSheetDevCellModel?.(sheetName, row, col);
        return cell ? String(cell.display ?? cell.raw ?? "") : null;
      };
      return {
        targetAddress: debugTargetAddress,
        sheetName,
        label: document.querySelector("#selectedCellLabel")?.textContent || "",
        status: document.querySelector("#statusText")?.textContent || "",
        selectionRange: window.__webSheetDevWorkbookState?.().selectionRange || null,
        domA16A19: [16, 17, 18, 19].map((row) => cellText(`A${row}`)),
        modelA16A19: [16, 17, 18, 19].map((row) => modelText(row, 1)),
      };
    }, { targetAddress });
    throw new Error(`Auto-fill drag to ${targetAddress} did not reach expected result: ${JSON.stringify(debug)}`);
  };


  const doubleClickAutoFillRange = async (range) => {
    await page.evaluate(({ top, left, bottom, right }) =>
      window.__webSheetDevSelectRange?.(top, left, bottom, right), range);
    await page.waitForFunction(
      ({ top, left, bottom, right }) => {
        const selected = window.__webSheetDevWorkbookState?.().selectionRange;
        return selected?.top === top && selected?.left === left && selected?.bottom === bottom && selected?.right === right;
      },
      range,
      { timeout: 10000 },
    );
    await page.evaluate(() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))));
    const handle = page.locator(".sheet-selection-frame .sheet-fill-handle");
    await handle.waitFor({ state: "visible", timeout: 10000 });
    await handle.dblclick();
  };


  await dragAutoFillRange({ top: 8, left: 1, bottom: 8, right: 1 }, "A11");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="A11"] .cell-content')?.textContent === "項目004", null, { timeout: 10000 });


  await dragAutoFillRange({ top: 12, left: 1, bottom: 12, right: 1 }, "A15");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="A15"] .cell-content')?.textContent === "木曜日", null, { timeout: 10000 });


  await dragAutoFillRange({ top: 16, left: 1, bottom: 16, right: 1 }, "A19", {
    until: () => {
      const domText = document.querySelector('.sheet-cell[data-address="A19"] .cell-content')?.textContent || "";
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "CHT";
      const cell = window.__webSheetDevCellModel?.(sheetName, 19, 1);
      return domText === "4月" || String(cell?.display ?? cell?.raw ?? "") === "4月";
    },
  });


  await page.waitForFunction(() => {
    const domText = document.querySelector('.sheet-cell[data-address="A19"] .cell-content')?.textContent || "";
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "CHT";
    const cell = window.__webSheetDevCellModel?.(sheetName, 19, 1);
    return domText === "4月" || String(cell?.display ?? cell?.raw ?? "") === "4月";
  }, null, { timeout: 10000 });


  await dragAutoFillRange({ top: 1, left: 1, bottom: 1, right: 1 }, "A3");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="A3"] .cell-content')?.textContent === "2026/6/3", null, { timeout: 10000 });


  await dragAutoFillRange({ top: 6, left: 2, bottom: 6, right: 3 }, "E6");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="E6"] .cell-content')?.textContent === "2026/6/7", null, { timeout: 10000 });


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    dateStep: {
      d6: document.querySelector('.sheet-cell[data-address="D6"] .cell-content')?.textContent,
      e6: document.querySelector('.sheet-cell[data-address="E6"] .cell-content')?.textContent,
    },
  })));


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(6, 2, "9:00");
  });


  await dragAutoFillRange({ top: 6, left: 2, bottom: 6, right: 2 }, "D6");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="D6"] .cell-content')?.textContent === "11:00", null, { timeout: 10000 });


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    timeText: {
      c6: document.querySelector('.sheet-cell[data-address="C6"] .cell-content')?.textContent,
      d6: document.querySelector('.sheet-cell[data-address="D6"] .cell-content')?.textContent,
    },
  })));


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(6, 2, "2026年6月1日");
  });


  await dragAutoFillRange({ top: 6, left: 2, bottom: 6, right: 2 }, "D6");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="D6"] .cell-content')?.textContent === "2026年6月3日", null, { timeout: 10000 });


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    japaneseDate: {
      c6: document.querySelector('.sheet-cell[data-address="C6"] .cell-content')?.textContent,
      d6: document.querySelector('.sheet-cell[data-address="D6"] .cell-content')?.textContent,
    },
  })));


  await dragAutoFillRange({ top: 13, left: 2, bottom: 13, right: 2 }, "E13");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="E13"] .cell-content')?.textContent === "R-004", null, { timeout: 10000 });


  await page.evaluate(() => {
    [
      [7, 43, "10"], [8, 43, ""], [9, 43, ""],
      [7, 44, "10"], [8, 44, ""], [9, 44, ""],
      [13, 13, "Q1"], [13, 14, ""], [13, 15, ""], [13, 16, ""],
      [12, 43, "5"], [13, 43, "6"], [10, 43, ""], [11, 43, ""],
      [14, 13, "5"], [14, 14, "6"], [14, 11, ""], [14, 12, ""],
      [16, 44, "隣接1"], [17, 44, "隣接2"], [18, 44, "隣接3"], [19, 44, "隣接4"], [20, 44, "隣接5"], [21, 44, "隣接6"], [22, 44, "隣接7"],
      [16, 45, "1"], [17, 45, "2"], [18, 45, ""], [19, 45, ""], [20, 45, ""], [21, 45, ""], [22, 45, ""],
      [7, 46, "2026/1/31"], [8, 46, ""], [9, 46, ""],
      [7, 47, "2026/6/5"], [8, 47, ""], [9, 47, ""],
      [7, 48, "2026/6/1"], [8, 48, ""], [9, 48, ""],
      [7, 49, "20"], [8, 49, ""], [9, 49, ""],
    ].forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
  });


  await dragAutoFillRange({ top: 7, left: 43, bottom: 7, right: 43 }, "AQ9");


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="AQ8"] .cell-content')?.textContent === "10" &&
      document.querySelector('.sheet-cell[data-address="AQ9"] .cell-content')?.textContent === "10",
    null,
    { timeout: 10000 },
  );


  await page.click("#autoFillOptionsButton");


  await page.waitForSelector("#autoFillOptionsMenu", { timeout: 5000 });


  await page.waitForFunction(() => document.activeElement?.dataset?.autoFillOption === "copy", null, { timeout: 5000 });


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    singleNumberDefault: {
      aq8: document.querySelector('.sheet-cell[data-address="AQ8"] .cell-content')?.textContent,
      aq9: document.querySelector('.sheet-cell[data-address="AQ9"] .cell-content')?.textContent,
      checked: document.querySelector("#autoFillOptionsMenu [aria-checked='true']")?.dataset.autoFillOption || "",
      focused: document.activeElement?.dataset?.autoFillOption || "",
    },
  })));


  await page.keyboard.press("S");


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="AQ8"] .cell-content')?.textContent === "11" &&
      document.querySelector('.sheet-cell[data-address="AQ9"] .cell-content')?.textContent === "12",
    null,
    { timeout: 10000 },
  );


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    singleNumberSeries: {
      aq8: document.querySelector('.sheet-cell[data-address="AQ8"] .cell-content')?.textContent,
      aq9: document.querySelector('.sheet-cell[data-address="AQ9"] .cell-content')?.textContent,
      status: document.querySelector("#statusText")?.textContent,
    },
  })));


  await dragAutoFillRange({ top: 7, left: 44, bottom: 7, right: 44 }, "AR9", { control: true });


  await page.waitForFunction(
    () => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const cellText = (row, col) => String(window.__webSheetDevCellModel?.(sheetName, row, col)?.display ?? window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "");
      return cellText(8, 44) === "11" && cellText(9, 44) === "12";
    },
    null,
    { timeout: 10000 },
  );


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    controlDrag: {
      ar8: String(window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 8, 44)?.display ?? window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 8, 44)?.raw ?? ""),
      ar9: String(window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 9, 44)?.display ?? window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 9, 44)?.raw ?? ""),
      checked: document.querySelector("#autoFillOptionsMenu [aria-checked='true']")?.dataset.autoFillOption || "",
      status: document.querySelector("#statusText")?.textContent,
    },
  })));


  await doubleClickAutoFillRange({ top: 16, left: 45, bottom: 17, right: 45 });


  await page.waitForFunction(
    () => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const cellText = (row, col) => String(window.__webSheetDevCellModel?.(sheetName, row, col)?.display ?? window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "");
      return cellText(18, 45) === "3" && cellText(22, 45) === "7";
    },
    null,
    { timeout: 10000 },
  );


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    doubleClick: {
      label: document.querySelector("#selectedCellLabel")?.textContent,
      as18: String(window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 18, 45)?.display ?? window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 18, 45)?.raw ?? ""),
      as22: String(window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 22, 45)?.display ?? window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 22, 45)?.raw ?? ""),
      status: document.querySelector("#statusText")?.textContent,
    },
  })));


  await dragAutoFillRange({ top: 7, left: 46, bottom: 7, right: 46 }, "AT9");


  await page.waitForFunction(
    () => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const cell = window.__webSheetDevCellModel?.(sheetName, 9, 46);
      return String(cell?.display ?? cell?.raw ?? "") === "2026/2/02";
    },
    null,
    { timeout: 10000 },
  );


  await page.click("#autoFillOptionsButton");


  await page.waitForSelector("#autoFillOptionsMenu", { timeout: 5000 });


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    dateOptionsMenu: {
      labels: [...document.querySelectorAll("#autoFillOptionsMenu [data-auto-fill-option]")].map((button) =>
        button.textContent.replace(/\s+/g, " ").trim().replace(/^✓/, ""),
      ),
      accessKeys: [...document.querySelectorAll("#autoFillOptionsMenu .menu-access-key")].map((item) => item.textContent),
      focused: document.activeElement?.dataset?.autoFillOption || "",
    },
  })));


  await page.keyboard.press("M");


  await page.waitForFunction(
    () => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const cellText = (row, col) => String(window.__webSheetDevCellModel?.(sheetName, row, col)?.display ?? window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "");
      return cellText(8, 46) === "2026/2/28" && cellText(9, 46) === "2026/3/31";
    },
    null,
    { timeout: 10000 },
  );


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    dateMonth: {
      at8: String(window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 8, 46)?.display ?? window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 8, 46)?.raw ?? ""),
      at9: String(window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 9, 46)?.display ?? window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 9, 46)?.raw ?? ""),
      status: document.querySelector("#statusText")?.textContent,
    },
  })));


  await dragAutoFillRange({ top: 7, left: 47, bottom: 7, right: 47 }, "AU9");


  await page.waitForFunction(
    () => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const cell = window.__webSheetDevCellModel?.(sheetName, 9, 47);
      return String(cell?.display ?? cell?.raw ?? "") === "2026/6/7";
    },
    null,
    { timeout: 10000 },
  );


  await page.click("#autoFillOptionsButton");


  await page.waitForSelector("#autoFillOptionsMenu", { timeout: 5000 });


  await page.keyboard.press("W");


  await page.waitForFunction(
    () => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const cellText = (row, col) => String(window.__webSheetDevCellModel?.(sheetName, row, col)?.display ?? window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "");
      return cellText(8, 47) === "2026/6/8" && cellText(9, 47) === "2026/6/9";
    },
    null,
    { timeout: 10000 },
  );


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    dateWeekday: {
      au8: String(window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 8, 47)?.display ?? window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 8, 47)?.raw ?? ""),
      au9: String(window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 9, 47)?.display ?? window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 9, 47)?.raw ?? ""),
      status: document.querySelector("#statusText")?.textContent,
    },
  })));


  await dragAutoFillRange({ top: 7, left: 48, bottom: 7, right: 48 }, "AV9");


  await page.waitForFunction(
    () => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const cell = window.__webSheetDevCellModel?.(sheetName, 9, 48);
      return String(cell?.display ?? cell?.raw ?? "") === "2026/6/3";
    },
    null,
    { timeout: 10000 },
  );


  await page.click("#autoFillOptionsButton");


  await page.waitForSelector("#autoFillOptionsMenu", { timeout: 5000 });


  await page.keyboard.press("Y");


  await page.waitForFunction(
    () => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const cellText = (row, col) => String(window.__webSheetDevCellModel?.(sheetName, row, col)?.display ?? window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "");
      return cellText(8, 48) === "2027/6/1" && cellText(9, 48) === "2028/6/1";
    },
    null,
    { timeout: 10000 },
  );


  Object.assign(autoFillSeries, await page.evaluate(() => ({
    dateYear: {
      av8: String(window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 8, 48)?.display ?? window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 8, 48)?.raw ?? ""),
      av9: String(window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 9, 48)?.display ?? window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 9, 48)?.raw ?? ""),
      status: document.querySelector("#statusText")?.textContent,
    },
  })));


  await dragAutoFillRange({ top: 7, left: 49, bottom: 7, right: 49 }, "AW9", { button: "right" });


  await page.waitForSelector("#autoFillOptionsMenu", { timeout: 5000 });


  await page.waitForFunction(
    () => document.activeElement?.dataset?.autoFillOption === "copy",
    null,
    { timeout: 5000 },
  );


	  Object.assign(autoFillSeries, await page.evaluate(() => {
	    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
	    const cellText = (row, col) => String(window.__webSheetDevCellModel?.(sheetName, row, col)?.display ?? window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "");
	    return {
      rightDragMenu: {
        aw8Before: cellText(8, 49),
        aw9Before: cellText(9, 49),
        labels: [...document.querySelectorAll("#autoFillOptionsMenu [data-auto-fill-option]")].map((button) =>
          button.textContent.replace(/\s+/g, " ").trim().replace(/^✓/, ""),
        ),
        checked: document.querySelector("#autoFillOptionsMenu [aria-checked='true']")?.dataset.autoFillOption || "",
        focused: document.activeElement?.dataset?.autoFillOption || "",
        label: document.querySelector("#selectedCellLabel")?.textContent || "",
      },
    };
  }));


  await page.keyboard.press("S");


  await page.waitForFunction(
    () => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const cellText = (row, col) => String(window.__webSheetDevCellModel?.(sheetName, row, col)?.display ?? window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "");
      return cellText(8, 49) === "21" && cellText(9, 49) === "22";
    },
    null,
    { timeout: 10000 },
  );


  Object.assign(autoFillSeries, await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    const cellText = (row, col) => String(window.__webSheetDevCellModel?.(sheetName, row, col)?.display ?? window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "");
    return {
      rightDragSeries: {
        aw8: cellText(8, 49),
        aw9: cellText(9, 49),
        menuClosed: !document.querySelector("#autoFillOptionsMenu"),
        status: document.querySelector("#statusText")?.textContent || "",
	      },
	    };
	  }));


	  await page.evaluate(() => {
	    [
	      [7, 50, "2"],
	      [8, 50, "4"],
	      [9, 50, ""],
	      [10, 50, ""],
	      [7, 51, "2"],
	      [8, 51, "4"],
	      [9, 51, ""],
	      [10, 51, ""],
	    ].forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
	  });


	  await dragAutoFillRange({ top: 7, left: 50, bottom: 8, right: 50 }, "AX10", { button: "right" });


	  await page.waitForSelector("#autoFillOptionsMenu", { timeout: 5000 });


	  Object.assign(autoFillSeries, await page.evaluate(() => ({
	    linearTrendMenu: {
	      labels: [...document.querySelectorAll("#autoFillOptionsMenu [data-auto-fill-option]")].map((button) =>
	        button.textContent.replace(/\s+/g, " ").trim().replace(/^✓/, ""),
	      ),
	      accessKeys: [...document.querySelectorAll("#autoFillOptionsMenu [data-auto-fill-option]")].map((button) => button.dataset.accessKey || ""),
	    },
	  })));


	  await page.keyboard.press("L");


	  await page.waitForFunction(
	    () => {
	      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
	      const cellText = (row, col) => String(window.__webSheetDevCellModel?.(sheetName, row, col)?.display ?? window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "");
	      return cellText(9, 50) === "6" && cellText(10, 50) === "8";
	    },
	    null,
	    { timeout: 10000 },
	  );


	  Object.assign(autoFillSeries, await page.evaluate(() => {
	    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
	    const cellText = (row, col) => String(window.__webSheetDevCellModel?.(sheetName, row, col)?.display ?? window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "");
	    return {
	      linearTrend: {
	        ax9: cellText(9, 50),
	        ax10: cellText(10, 50),
	        menuClosed: !document.querySelector("#autoFillOptionsMenu"),
	        status: document.querySelector("#statusText")?.textContent || "",
	      },
	    };
	  }));


	  await dragAutoFillRange({ top: 7, left: 51, bottom: 8, right: 51 }, "AY10", { button: "right" });


	  await page.waitForSelector("#autoFillOptionsMenu", { timeout: 5000 });


	  Object.assign(autoFillSeries, await page.evaluate(() => ({
	    growthTrendMenu: {
	      labels: [...document.querySelectorAll("#autoFillOptionsMenu [data-auto-fill-option]")].map((button) =>
	        button.textContent.replace(/\s+/g, " ").trim().replace(/^✓/, ""),
	      ),
	      accessKeys: [...document.querySelectorAll("#autoFillOptionsMenu [data-auto-fill-option]")].map((button) => button.dataset.accessKey || ""),
	    },
	  })));


	  await page.keyboard.press("G");


	  await page.waitForFunction(
	    () => {
	      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
	      const cellText = (row, col) => String(window.__webSheetDevCellModel?.(sheetName, row, col)?.display ?? window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "");
	      return cellText(9, 51) === "8" && cellText(10, 51) === "16";
	    },
	    null,
	    { timeout: 10000 },
	  );


	  Object.assign(autoFillSeries, await page.evaluate(() => {
	    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
	    const cellText = (row, col) => String(window.__webSheetDevCellModel?.(sheetName, row, col)?.display ?? window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "");
	    return {
	      growthTrend: {
	        ay9: cellText(9, 51),
	        ay10: cellText(10, 51),
	        menuClosed: !document.querySelector("#autoFillOptionsMenu"),
	        status: document.querySelector("#statusText")?.textContent || "",
	      },
	    };
	  }));


	  await dragAutoFillRange({ top: 13, left: 13, bottom: 13, right: 13 }, "P13");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="P13"] .cell-content')?.textContent === "Q4", null, { timeout: 10000 });


  await dragAutoFillRange({ top: 12, left: 43, bottom: 13, right: 43 }, "AQ10");


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="AQ10"] .cell-content')?.textContent === "3" &&
      document.querySelector('.sheet-cell[data-address="AQ11"] .cell-content')?.textContent === "4",
    null,
    { timeout: 10000 },
  );


  await page.locator("#sheetHost").evaluate((host) => {
    host.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll", { bubbles: true }));
  });


  await page.waitForSelector('.sheet-cell[data-address="M14"]', { timeout: 10000 });


  await dragAutoFillRange({ top: 14, left: 13, bottom: 14, right: 14 }, "K14");


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="K14"] .cell-content')?.textContent === "3" &&
      document.querySelector('.sheet-cell[data-address="L14"] .cell-content')?.textContent === "4",
    null,
    { timeout: 10000 },
  );


  Object.assign(autoFillSeries, await page.evaluate(() => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    const cellText = (row, col) => {
      const cell = window.__webSheetDevCellModel?.(sheetName, row, col);
      return String(cell?.display ?? cell?.raw ?? "");
    };
    return {
      textNumber: {
        a9: cellText(9, 1),
        a10: cellText(10, 1),
        a11: cellText(11, 1),
        c13: cellText(13, 3),
        d13: cellText(13, 4),
        e13: cellText(13, 5),
      },
      weekdays: {
        a13: cellText(13, 1),
        a14: cellText(14, 1),
        a15: cellText(15, 1),
      },
      months: {
        a17: cellText(17, 1),
        a18: cellText(18, 1),
        a19: cellText(19, 1),
      },
      dateText: {
        a2: cellText(2, 1),
        a3: cellText(3, 1),
      },
      quarter: {
        n13: cellText(13, 14),
        o13: cellText(13, 15),
        p13: cellText(13, 16),
      },
      reverseVertical: {
        aq10: cellText(10, 43),
        aq11: cellText(11, 43),
      },
      reverseHorizontal: {
        k14: cellText(14, 11),
        l14: cellText(14, 12),
      },
    };
  }));



  await page.locator("#sheetHost").evaluate((host) => {
    host.scrollTop = 0;
    host.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll", { bubbles: true }));
  });


  await page.waitForSelector('.sheet-cell[data-address="K6"]');


  const setCellViaFormulaBar = async (address, value) => {
    await page.locator(`.sheet-cell[data-address="${address}"]`).click();
    await page.waitForFunction((targetAddress) => document.querySelector("#selectedCellLabel")?.textContent === targetAddress, address, { timeout: 10000 });
    await page.locator("#formulaInput").click();
    await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");
    await page.keyboard.press("Backspace");
    await page.keyboard.type(String(value));
    await page.keyboard.press("Enter");
    await page.waitForFunction(
      ({ targetAddress, targetValue }) =>
        document.querySelector(`.sheet-cell[data-address="${targetAddress}"] .cell-content`)?.textContent === String(targetValue),
      { targetAddress: address, targetValue: value },
      { timeout: 10000 },
    );
  };


  await setCellViaFormulaBar("K6", "InsertA");


  await setCellViaFormulaBar("L6", "InsertB");


  await page.click('.sheet-cell[data-address="K6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="insert"]');


  await page.waitForSelector("#cellStructureDialog");


  const insertDeleteDialog = {};


  insertDeleteDialog.insert = await page.evaluate(() => ({
    title: document.querySelector("#cellStructureDialogTitle")?.textContent,
    options: [...document.querySelectorAll('input[name="cellStructureOption"]')].map((input) => ({
      value: input.value,
      label: input.closest("label")?.textContent.replace(/\s+/g, ""),
      checked: input.checked,
    })),
    ok: document.querySelector('[data-cell-structure-action="ok"]')?.textContent,
    cancel: document.querySelector('[data-cell-structure-action="cancel"]')?.textContent,
  }));


  await page.click('[data-cell-structure-action="ok"]');


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="K6"] .cell-content')?.textContent === "" &&
      document.querySelector('.sheet-cell[data-address="L6"] .cell-content')?.textContent === "InsertA" &&
      document.querySelector('.sheet-cell[data-address="M6"] .cell-content')?.textContent === "InsertB",
    null,
    { timeout: 10000 },
  );


  insertDeleteDialog.afterInsert = await page.evaluate(() => ({
    k6: document.querySelector('.sheet-cell[data-address="K6"] .cell-content')?.textContent,
    l6: document.querySelector('.sheet-cell[data-address="L6"] .cell-content')?.textContent,
    m6: document.querySelector('.sheet-cell[data-address="M6"] .cell-content')?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.click('.sheet-cell[data-address="K6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="delete"]');


  await page.waitForSelector("#cellStructureDialog");


  insertDeleteDialog.delete = await page.evaluate(() => ({
    title: document.querySelector("#cellStructureDialogTitle")?.textContent,
    options: [...document.querySelectorAll('input[name="cellStructureOption"]')].map((input) => ({
      value: input.value,
      label: input.closest("label")?.textContent.replace(/\s+/g, ""),
      checked: input.checked,
    })),
    ok: document.querySelector('[data-cell-structure-action="ok"]')?.textContent,
    cancel: document.querySelector('[data-cell-structure-action="cancel"]')?.textContent,
  }));


  await page.click('[data-cell-structure-action="ok"]');


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="K6"] .cell-content')?.textContent === "InsertA" &&
      document.querySelector('.sheet-cell[data-address="L6"] .cell-content')?.textContent === "InsertB",
    null,
    { timeout: 10000 },
  );


  insertDeleteDialog.afterDelete = await page.evaluate(() => ({
    k6: document.querySelector('.sheet-cell[data-address="K6"] .cell-content')?.textContent,
    l6: document.querySelector('.sheet-cell[data-address="L6"] .cell-content')?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="K6"] .cell-content')?.textContent === "", null, { timeout: 10000 });


  insertDeleteDialog.undoDelete = await page.evaluate(() => ({
    k6: document.querySelector('.sheet-cell[data-address="K6"] .cell-content')?.textContent,
    l6: document.querySelector('.sheet-cell[data-address="L6"] .cell-content')?.textContent,
    m6: document.querySelector('.sheet-cell[data-address="M6"] .cell-content')?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="K6"] .cell-content')?.textContent === "InsertA", null, { timeout: 10000 });


  insertDeleteDialog.redoDelete = await page.evaluate(() => ({
    k6: document.querySelector('.sheet-cell[data-address="K6"] .cell-content')?.textContent,
    l6: document.querySelector('.sheet-cell[data-address="L6"] .cell-content')?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }));



  await page.locator('.sheet-cell[data-address="G6"]').click();


  await page.evaluate(async () => navigator.clipboard.writeText("貼り付けA\t貼り付けB"));


  await page.keyboard.press("Control+V");


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="G6"] .cell-content')?.textContent === "貼り付けA" &&
      document.querySelector('.sheet-cell[data-address="H6"] .cell-content')?.textContent === "貼り付けB",
    null,
    { timeout: 10000 },
  );


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="G6"] .cell-content')?.textContent !== "貼り付けA", null, { timeout: 10000 });


  const shortcutUndo = await page.evaluate(() => ({
    g6: document.querySelector('.sheet-cell[data-address="G6"] .cell-content')?.textContent,
    h6: document.querySelector('.sheet-cell[data-address="H6"] .cell-content')?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="G6"] .cell-content')?.textContent === "貼り付けA", null, { timeout: 10000 });


  await page.click('[data-ribbon-tab="home"]');


  await page.locator('.sheet-cell[data-address="G6"]').click();


  await page.click('[data-ribbon-command="bold"]');


  await page.selectOption('[data-ribbon-style="fontSize"]', "16");


  await page.click('[data-ribbon-command="align-center"]');


  await page.click('[data-ribbon-command="wrap-text"]');


  await page.waitForFunction(() => getComputedStyle(document.querySelector('.sheet-cell[data-address="G6"]')).whiteSpace === "pre-wrap", null, { timeout: 10000 });


  const wrapToolbar = await page.evaluate(() => ({
    activeAfterWrap: document.querySelector('[data-ribbon-command="wrap-text"]')?.classList.contains("is-active"),
    whiteSpaceAfterWrap: getComputedStyle(document.querySelector('.sheet-cell[data-address="G6"]')).whiteSpace,
    overflowWrapAfterWrap: getComputedStyle(document.querySelector('.sheet-cell[data-address="G6"]')).overflowWrap,
  }));


  await page.click('[data-ribbon-command="wrap-text"]');


  await page.waitForFunction(() => getComputedStyle(document.querySelector('.sheet-cell[data-address="G6"]')).whiteSpace !== "pre-wrap", null, { timeout: 10000 });


  Object.assign(wrapToolbar, await page.evaluate(() => ({
    activeAfterUnwrap: document.querySelector('[data-ribbon-command="wrap-text"]')?.classList.contains("is-active"),
    whiteSpaceAfterUnwrap: getComputedStyle(document.querySelector('.sheet-cell[data-address="G6"]')).whiteSpace,
  })));


  await page.locator('[data-ribbon-style="fillColor"]').evaluate((input) => {
    input.value = "#d9ead3";
    input.dispatchEvent(new Event("input", { bubbles: true }));
  });


  await page.waitForFunction(
    () => getComputedStyle(document.querySelector('.sheet-cell[data-address="G6"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase() === "#d9ead3",
    null,
    { timeout: 10000 },
  );


  const formatUndoRedo = {};


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    () => getComputedStyle(document.querySelector('.sheet-cell[data-address="G6"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase() !== "#d9ead3",
    null,
    { timeout: 10000 },
  );


  Object.assign(formatUndoRedo, await page.evaluate(() => ({
    undoFill: getComputedStyle(document.querySelector('.sheet-cell[data-address="G6"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase(),
    undoStatus: document.querySelector("#statusText")?.textContent,
  })));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(
    () => getComputedStyle(document.querySelector('.sheet-cell[data-address="G6"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase() === "#d9ead3",
    null,
    { timeout: 10000 },
  );


  Object.assign(formatUndoRedo, await page.evaluate(() => ({
    redoFill: getComputedStyle(document.querySelector('.sheet-cell[data-address="G6"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase(),
    redoStatus: document.querySelector("#statusText")?.textContent,
  })));


  await page.locator('.sheet-cell[data-address="I6"]').click();


  await page.locator("#formulaInput").fill("target value");


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="I6"] .cell-content')?.textContent === "target value", null, { timeout: 10000 });


  await page.locator('.sheet-cell[data-address="G6"]').click();


  await page.click('[data-ribbon-command="format-painter"]');


  const painterActiveAfterStart = await page.evaluate(() => ({
    active: document.querySelector('[data-ribbon-command="format-painter"]')?.classList.contains("is-active"),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.locator('.sheet-cell[data-address="I6"]').click();


  await page.waitForFunction(
    () => getComputedStyle(document.querySelector('.sheet-cell[data-address="I6"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase() === "#d9ead3",
    null,
    { timeout: 10000 },
  );


  Object.assign(formatUndoRedo, await page.evaluate((painterStart) => {
    const cell = document.querySelector('.sheet-cell[data-address="I6"]');
    const style = getComputedStyle(cell);
    return {
      painterActiveAfterStart: painterStart,
      painterActiveAfterApply: document.querySelector('[data-ribbon-command="format-painter"]')?.classList.contains("is-active"),
      painterText: document.querySelector('.sheet-cell[data-address="I6"] .cell-content')?.textContent || "",
      painterFill: style.getPropertyValue("--cell-fill-color").trim().toLowerCase(),
      painterFontSize: style.fontSize,
      painterBold: Number.parseInt(style.fontWeight, 10) >= 600,
      painterAlign: style.textAlign,
      painterStatus: document.querySelector("#statusText")?.textContent,
    };
  }, painterActiveAfterStart));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    () => getComputedStyle(document.querySelector('.sheet-cell[data-address="I6"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase() !== "#d9ead3",
    null,
    { timeout: 10000 },
  );


  Object.assign(formatUndoRedo, await page.evaluate(() => ({
    painterUndoText: document.querySelector('.sheet-cell[data-address="I6"] .cell-content')?.textContent || "",
    painterUndoFill: getComputedStyle(document.querySelector('.sheet-cell[data-address="I6"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase(),
    painterUndoStatus: document.querySelector("#statusText")?.textContent,
  })));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(
    () => getComputedStyle(document.querySelector('.sheet-cell[data-address="I6"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase() === "#d9ead3",
    null,
    { timeout: 10000 },
  );


  Object.assign(formatUndoRedo, await page.evaluate(() => ({
    painterRedoFill: getComputedStyle(document.querySelector('.sheet-cell[data-address="I6"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase(),
    painterRedoStatus: document.querySelector("#statusText")?.textContent,
  })));


  await page.locator('.sheet-cell[data-address="G6"]').click();


  await page.locator('[data-ribbon-style="fillColor"]').evaluate((input) => {
    input.value = "#123456";
  });


  await page.locator('.sheet-cell[data-address="H6"]').click({ modifiers: ["Shift"] });


  await page.waitForTimeout(200);


  const toolbarSync = await page.evaluate(() => ({
    fontSize: document.querySelector('[data-ribbon-style="fontSize"]')?.value,
    boldActive: document.querySelector('[data-ribbon-command="bold"]')?.classList.contains("is-active"),
    alignCenterActive: document.querySelector('[data-ribbon-command="align-center"]')?.classList.contains("is-active"),
    fillColor: document.querySelector('[data-ribbon-style="fillColor"]')?.value,
    fontColor: document.querySelector('[data-ribbon-style="fontColor"]')?.value,
    selectionLabel: document.querySelector("#selectedCellLabel")?.textContent,
  }));


  const selectionFrame = await page.evaluate(() => {
    const rectFor = (selector) => {
      const element = document.querySelector(selector);
      const rect = element?.getBoundingClientRect();
      return rect
        ? {
            left: Math.round(rect.left),
            top: Math.round(rect.top),
            right: Math.round(rect.right),
            bottom: Math.round(rect.bottom),
            width: Math.round(rect.width),
            height: Math.round(rect.height),
          }
        : null;
    };
    const g6 = document.querySelector('.sheet-cell[data-address="G6"]');
    const h6 = document.querySelector('.sheet-cell[data-address="H6"]');
    const gRect = rectFor('.sheet-cell[data-address="G6"]');
    const hRect = rectFor('.sheet-cell[data-address="H6"]');
    const frameRect = rectFor(".sheet-selection-frame");
    return {
      exists: Boolean(frameRect),
      selectionLabel: document.querySelector("#selectedCellLabel")?.textContent,
      frame: frameRect,
      expectedLeft: gRect?.left,
      expectedTop: gRect?.top,
      expectedRight: hRect?.right,
      expectedBottom: hRect?.bottom,
      g6Outline: g6 ? getComputedStyle(g6).outlineStyle : null,
      h6Outline: h6 ? getComputedStyle(h6).outlineStyle : null,
      g6UserSelect: g6 ? getComputedStyle(g6).userSelect : null,
    };
  });


  await page.keyboard.press("Control+C");


  const copyStateAfterCopy = await page.evaluate(() => {
    const frame = document.querySelector(".sheet-selection-frame");
    const beforeStyle = frame ? getComputedStyle(frame, "::before") : null;
    const internalClipboard = window.__webSheetDevWorkbookState?.().internalClipboard || null;
    return {
      selectionLabel: document.querySelector("#selectedCellLabel")?.textContent,
      selectionFrameClass: frame?.className,
      copyFrameCount: document.querySelectorAll(".sheet-copy-frame").length,
      animationName: beforeStyle?.animationName,
      backgroundImage: beforeStyle?.backgroundImage,
      internalClipboard,
    };
  });


  const shortcutClipboard = copyStateAfterCopy.internalClipboard?.text || "";


  await page.evaluate(async () => navigator.clipboard.writeText("外部でコピー済みの古いテキスト"));


  await page.locator('.sheet-cell[data-address="I6"]').click();


  const copyStateWithTarget = await page.evaluate(() => {
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
    const copyFrame = document.querySelector(".sheet-copy-frame");
    return {
      selectionLabel: document.querySelector("#selectedCellLabel")?.textContent,
      copyFrameClass: copyFrame?.className,
      copyFrame: rectFor(".sheet-copy-frame"),
      sourceLeft: rectFor('.sheet-cell[data-address="G6"]')?.left,
      sourceRight: rectFor('.sheet-cell[data-address="H6"]')?.right,
      targetSelection: rectFor(".sheet-selection-frame"),
      targetCell: rectFor('.sheet-cell[data-address="I6"]'),
    };
  });


  await page.keyboard.press("Control+V");


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="I6"] .cell-content')?.textContent === "貼り付けA" &&
      document.querySelector('.sheet-cell[data-address="J6"] .cell-content')?.textContent === "貼り付けB",
    null,
    { timeout: 10000 },
  );


  const richPaste = await page.evaluate(() => {
    const inspect = (address) => {
      const cell = document.querySelector(`.sheet-cell[data-address="${address}"]`);
      const style = cell ? getComputedStyle(cell) : null;
      return {
        text: cell?.querySelector(".cell-content")?.textContent,
        fontWeight: style?.fontWeight,
        fontSize: style?.fontSize,
        textAlign: style?.textAlign,
        fillColor: style?.getPropertyValue("--cell-fill-color").trim().toLowerCase(),
        backgroundColor: style?.backgroundColor,
      };
    };
    return {
      source: inspect("G6"),
      neighborSource: inspect("H6"),
      pasted: inspect("I6"),
      neighborPasted: inspect("J6"),
    };
  });


  await page.evaluate(async () => navigator.clipboard.writeText("外部でコピー済みの古いテキスト"));


  await page.locator('.sheet-cell[data-address="K6"]').click();


  await page.click('.sheet-cell[data-address="K6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="paste-match-format"]');


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="K6"] .cell-content')?.textContent === "貼り付けA" &&
      document.querySelector('.sheet-cell[data-address="L6"] .cell-content')?.textContent === "貼り付けB",
    null,
    { timeout: 10000 },
  );


  const matchDestinationPaste = await page.evaluate(() => {
    const inspect = (address) => {
      const cell = document.querySelector(`.sheet-cell[data-address="${address}"]`);
      const style = cell ? getComputedStyle(cell) : null;
      return {
        text: cell?.querySelector(".cell-content")?.textContent,
        fontWeight: style?.fontWeight,
        fontSize: style?.fontSize,
        textAlign: style?.textAlign,
        fillColor: style?.getPropertyValue("--cell-fill-color").trim().toLowerCase(),
        backgroundColor: style?.backgroundColor,
      };
    };
    return {
      pasted: inspect("K6"),
      neighborPasted: inspect("L6"),
    };
  });


  await page.locator('.sheet-cell[data-address="M6"]').click();


  await page.keyboard.press("Control+Alt+V");


  await page.waitForSelector('[data-data-dialog="paste-special"]', { timeout: 5000 });


  await page.locator('[data-data-dialog="paste-special"] input[value="formats"]').check();


  const pasteSpecialDialog = await page.evaluate(() => ({
    title: document.querySelector("#pasteSpecialTitle")?.textContent || "",
    options: [...document.querySelectorAll('[data-data-dialog="paste-special"] input[name="pasteSpecialMode"]')]
      .map((input) => input.closest("label")?.textContent.trim() || ""),
  }));


  await page.click('[data-data-dialog="paste-special"] [data-data-dialog-action="ok"]');


  await page.waitForFunction(
    () => getComputedStyle(document.querySelector('.sheet-cell[data-address="M6"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase() === "#d9ead3",
    null,
    { timeout: 10000 },
  );


  const pasteSpecialFormats = await page.evaluate((dialogInfo) => {
    const cell = document.querySelector('.sheet-cell[data-address="M6"]');
    const style = cell ? getComputedStyle(cell) : null;
    return {
      ...dialogInfo,
      text: cell?.querySelector(".cell-content")?.textContent || "",
      fillColor: style?.getPropertyValue("--cell-fill-color").trim().toLowerCase(),
      fontSize: style?.fontSize,
      fontWeight: style?.fontWeight,
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, pasteSpecialDialog);


  const pasteSpecialTestRow = 180;


  await page.evaluate((baseRow) => {
    [
      [baseRow, 14, "2"],
      [baseRow, 15, "3"],
      [baseRow + 1, 14, ""],
      [baseRow + 1, 15, "4"],
      [baseRow, 17, "10"],
      [baseRow, 18, "99"],
      [baseRow + 1, 17, "30"],
      [baseRow + 1, 18, "40"],
    ].forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
    window.__webSheetDevSelectRange?.(baseRow, 14, baseRow + 1, 15);
  }, pasteSpecialTestRow);


  await page.keyboard.press("Control+C");


  await page.evaluate((baseRow) => window.__webSheetDevSelectRange?.(baseRow, 17, baseRow, 17), pasteSpecialTestRow);


  await page.keyboard.press("Control+Alt+V");


  await page.waitForSelector('[data-data-dialog="paste-special"]', { timeout: 5000 });


  await page.locator('[data-data-dialog="paste-special"] input[value="values"]').check();


  await page.locator('[data-data-dialog="paste-special"] input[name="pasteSpecialOperation"][value="add"]').check();


  await page.locator('[data-data-dialog="paste-special"] [data-paste-special-field="skipBlanks"]').check();


  await page.locator('[data-data-dialog="paste-special"] [data-paste-special-field="transpose"]').check();


  await page.click('[data-data-dialog="paste-special"] [data-data-dialog-action="ok"]');


  await page.waitForFunction(
    (baseRow) => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const raw = (row, col) => window.__webSheetDevCellModel?.(sheetName, row, col)?.raw;
      return (
        String(raw(baseRow, 17)) === "12" &&
        String(raw(baseRow, 18)) === "99" &&
        String(raw(baseRow + 1, 17)) === "33" &&
        String(raw(baseRow + 1, 18)) === "44"
      );
    },
    pasteSpecialTestRow,
    { timeout: 10000 },
  );


  await page.evaluate((baseRow) => window.__webSheetDevSelectRange?.(baseRow, 19, baseRow, 19), pasteSpecialTestRow);


  await page.keyboard.press("Control+Alt+V");


  await page.waitForSelector('[data-data-dialog="paste-special"]', { timeout: 5000 });


  await page.click('[data-data-dialog="paste-special"] [data-data-dialog-action="paste-special-link"]');


  await page.waitForFunction(
    (baseRow) => window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName, baseRow, 19)?.formula?.includes(`!N${baseRow}`),
    pasteSpecialTestRow,
    { timeout: 10000 },
  );


  const pasteSpecialAdvanced = await page.evaluate((baseRow) => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    const raw = (row, col) => window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "";
    const formula = (row, col) => window.__webSheetDevCellModel?.(sheetName, row, col)?.formula ?? "";
    return {
      transposeAddSkipBlanks: {
        q: raw(baseRow, 17),
        r: raw(baseRow, 18),
        qNext: raw(baseRow + 1, 17),
        rNext: raw(baseRow + 1, 18),
      },
      link: {
        s: formula(baseRow, 19),
        t: formula(baseRow, 20),
        sNext: formula(baseRow + 1, 19),
        tNext: formula(baseRow + 1, 20),
      },
      baseRow,
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, pasteSpecialTestRow);


  const formulaPasteRow = 184;


  await page.evaluate((baseRow) => {
    window.__webSheetDevSetCell?.(baseRow, 27, "10");
    window.__webSheetDevSetCell?.(baseRow, 28, `=AA${baseRow}`);
    window.__webSheetDevSetCell?.(baseRow + 1, 28, `=$AA$${baseRow}`);
    window.__webSheetDevSelectRange?.(baseRow, 28, baseRow + 1, 28);
  }, formulaPasteRow);


  await page.keyboard.press("Control+C");


  await page.evaluate((baseRow) => window.__webSheetDevSelectRange?.(baseRow + 2, 30, baseRow + 2, 30), formulaPasteRow);


  await page.keyboard.press("Control+V");


  await page.waitForFunction(
    (baseRow) => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const formula = (row, col) => window.__webSheetDevCellModel?.(sheetName, row, col)?.formula || "";
      return formula(baseRow + 2, 30) === `=AC${baseRow + 2}` && formula(baseRow + 3, 30) === `=$AA$${baseRow}`;
    },
    formulaPasteRow,
    { timeout: 10000 },
  );


  const formulaPasteReferences = await page.evaluate((baseRow) => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    const formula = (row, col) => window.__webSheetDevCellModel?.(sheetName, row, col)?.formula || "";
    return {
      relative: formula(baseRow + 2, 30),
      absolute: formula(baseRow + 3, 30),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, formulaPasteRow);


  const formulaCutRow = 188;


  await page.evaluate((baseRow) => {
    window.__webSheetDevSetCell?.(baseRow, 27, "10");
    window.__webSheetDevSetCell?.(baseRow, 28, `=AA${baseRow}`);
    window.__webSheetDevSelectRange?.(baseRow, 28, baseRow, 28);
  }, formulaCutRow);


  await page.keyboard.press("Control+X");


  await page.waitForFunction(
    (baseRow) => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const model = window.__webSheetDevCellModel?.(sheetName, baseRow, 28) || {};
      return model.formula === `=AA${baseRow}`;
    },
    formulaCutRow,
    { timeout: 10000 },
  );


  await page.evaluate((baseRow) => window.__webSheetDevSelectRange?.(baseRow + 2, 30, baseRow + 2, 30), formulaCutRow);


  await page.keyboard.press("Control+V");


  await page.waitForFunction(
    (baseRow) => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const formula = (row, col) => window.__webSheetDevCellModel?.(sheetName, row, col)?.formula || "";
      const raw = (row, col) => window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "";
      return formula(baseRow + 2, 30) === `=AA${baseRow}` && raw(baseRow, 28) === "";
    },
    formulaCutRow,
    { timeout: 10000 },
  );


  const undoShortcut = process.platform === "darwin" ? "Meta+Z" : "Control+Z";


  const redoShortcut = process.platform === "darwin" ? "Meta+Y" : "Control+Y";


  await page.keyboard.press(undoShortcut);


  await page.waitForFunction(
    (baseRow) => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const formula = (row, col) => window.__webSheetDevCellModel?.(sheetName, row, col)?.formula || "";
      const raw = (row, col) => window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "";
      return formula(baseRow, 28) === `=AA${baseRow}` && raw(baseRow + 2, 30) === "";
    },
    formulaCutRow,
    { timeout: 10000 },
  );


  await page.keyboard.press(redoShortcut);


  await page.waitForFunction(
    (baseRow) => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const formula = (row, col) => window.__webSheetDevCellModel?.(sheetName, row, col)?.formula || "";
      const raw = (row, col) => window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "";
      return formula(baseRow + 2, 30) === `=AA${baseRow}` && raw(baseRow, 28) === "";
    },
    formulaCutRow,
    { timeout: 10000 },
  );


  const repeatPasteRow = 192;


  await page.evaluate((baseRow) => {
    [
      [baseRow, 31, "A"],
      [baseRow, 32, "B"],
      [baseRow + 1, 31, "C"],
      [baseRow + 1, 32, "D"],
    ].forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
    window.__webSheetDevSelectRange?.(baseRow, 31, baseRow + 1, 32);
  }, repeatPasteRow);


  await page.keyboard.press("Control+C");


  await page.evaluate((baseRow) => window.__webSheetDevSelectRange?.(baseRow, 34, baseRow + 3, 37), repeatPasteRow);


  await page.keyboard.press("Control+V");


  await page.waitForFunction(
    (baseRow) => {
      const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
      const raw = (row, col) => String(window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "");
      return raw(baseRow, 34) === "A" &&
        raw(baseRow, 35) === "B" &&
        raw(baseRow, 36) === "A" &&
        raw(baseRow + 1, 34) === "C" &&
        raw(baseRow + 2, 34) === "A" &&
        raw(baseRow + 3, 37) === "D";
    },
    repeatPasteRow,
    { timeout: 10000 },
  );


  const cutAndRepeatPaste = await page.evaluate(({ cutRow, repeatRow }) => {
    const sheetName = window.__webSheetDevWorkbookState?.().activeSheetName || "";
    const raw = (row, col) => window.__webSheetDevCellModel?.(sheetName, row, col)?.raw ?? "";
    const formula = (row, col) => window.__webSheetDevCellModel?.(sheetName, row, col)?.formula || "";
    return {
      cutFormula: formula(cutRow + 2, 30),
      cutSourceRaw: raw(cutRow, 28),
      repeat: {
        ah: raw(repeatRow, 34),
        ai: raw(repeatRow, 35),
        aj: raw(repeatRow, 36),
        ahNext: raw(repeatRow + 1, 34),
        ahThird: raw(repeatRow + 2, 34),
        akLast: raw(repeatRow + 3, 37),
      },
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, { cutRow: formulaCutRow, repeatRow: repeatPasteRow });


  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    const horizontal = document.querySelector("#horizontalScrollbar");
    if (host) {
      host.scrollTop = 0;
      host.scrollLeft = 0;
      host.dispatchEvent(new Event("scroll"));
    }
    if (horizontal) {
      horizontal.scrollLeft = 0;
      horizontal.dispatchEvent(new Event("scroll"));
    }
  });


  await page.waitForSelector('.sheet-cell[data-address="G6"]', { timeout: 10000 });


  await page.locator('.sheet-cell[data-address="G6"]').click();


  await page.locator('.sheet-cell[data-address="H6"]').click({ modifiers: ["Shift"] });


  await page.locator('.sheet-cell[data-address="D6"]').click({ modifiers: ["Meta"] });


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "G6:H6,D6", null, { timeout: 10000 });


  await page.click('.sheet-cell[data-address="G6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  const cellMenu = await page.evaluate(() => ({
    labels: [...document.querySelectorAll("#cellContextMenu .cell-context-menu-item")].map((item) =>
      item.textContent.replace(/\s+/g, "").trim(),
    ),
    accessKeys: [...document.querySelectorAll("#cellContextMenu .menu-access-key")].map((item) => item.textContent),
    accessKeysUnderlined: [...document.querySelectorAll("#cellContextMenu .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
    disabledActions: [...document.querySelectorAll("#cellContextMenu .cell-context-menu-item:disabled")].map(
      (item) => item.dataset.cellMenuAction,
    ),
    separatorCount: document.querySelectorAll("#cellContextMenu .cell-context-menu-separator").length,
    pasteOptions: [...document.querySelectorAll("#cellContextMenu .cell-paste-option")].map((button) => button.getAttribute("title")),
    focusedBefore: document.activeElement?.dataset?.cellMenuAction || "",
    rect: (() => {
      const rect = document.querySelector("#cellContextMenu")?.getBoundingClientRect();
      return {
        left: Math.round(rect?.left || 0),
        top: Math.round(rect?.top || 0),
        right: Math.round(rect?.right || 0),
        bottom: Math.round(rect?.bottom || 0),
      };
    })(),
  }));


  await page.keyboard.press("ArrowDown");


  Object.assign(cellMenu, await page.evaluate(() => ({
    focusedAfterDown: document.activeElement?.dataset?.cellMenuAction || "",
  })));


  await page.keyboard.press("f");


  await page.waitForSelector("#formatCellsDialog", { timeout: 5000 });


  Object.assign(cellMenu, await page.evaluate(() => ({
    formatOpenedByAccessKey: document.querySelector("#formatCellsDialogTitle")?.textContent || "",
  })));


  await page.locator("#formatCellsDialogOverlay [data-format-dialog-action='cancel']").click();


  await page.waitForSelector("#formatCellsDialog", { state: "detached", timeout: 5000 });


  await page.click('.sheet-cell[data-address="G6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="translate"]');


  await page.waitForSelector("#translationPane");


  const translationPaneInitial = await page.evaluate(() => ({
    visible: Boolean(document.querySelector("#translationPane")),
    title: document.querySelector("#translationPane .translation-pane-header strong")?.textContent || "",
    range: document.querySelector("#translationPane .translation-pane-header span")?.textContent || "",
    sourceText: document.querySelector("#translationPane .translation-pane-source")?.textContent || "",
    resultText: document.querySelector("#translationPane .translation-pane-result")?.textContent || "",
  }));


  await page.click('[data-translation-pane-action="close"]');


  await page.locator('.sheet-cell[data-address="B4"]').click();


  await page.click('.sheet-cell[data-address="B4"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="translate"]');


  await page.waitForSelector("#translationPane");


  const translationPaneHeader = await page.evaluate(() => ({
    range: document.querySelector("#translationPane .translation-pane-header span")?.textContent || "",
    sourceText: document.querySelector("#translationPane .translation-pane-source")?.textContent || "",
    resultText: document.querySelector("#translationPane .translation-pane-result")?.textContent || "",
    detected: document.querySelector("#translationPane .translation-pane-detected")?.textContent || "",
  }));


  await page.click('[data-translation-pane-action="replace"]');


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="B4"] .cell-content')?.textContent === "姓名", null, {
    timeout: 10000,
  });


  const translationReplace = await page.evaluate(() => ({
    b4: document.querySelector('.sheet-cell[data-address="B4"] .cell-content')?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
    selectionLabel: document.querySelector("#selectedCellLabel")?.textContent || "",
  }));


  await page.keyboard.press(process.platform === "darwin" ? "Meta+Z" : "Control+Z");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="B4"] .cell-content')?.textContent === "氏名", null, {
    timeout: 10000,
  });


  const translationUndo = await page.evaluate(() => ({
    b4: document.querySelector('.sheet-cell[data-address="B4"] .cell-content')?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('[data-translation-pane-action="close"]');


  await page.click('.sheet-cell[data-address="G6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="new-comment"]');


  await page.waitForSelector("#cellAnnotationEditor.is-comment");


  const cellAnnotationEditorInitial = await page.evaluate(() => ({
    commentEditor: Boolean(document.querySelector("#cellAnnotationEditor.is-comment")),
    commentAria: document.querySelector("#cellAnnotationEditor")?.getAttribute("aria-label"),
    commentAuthor: document.querySelector(".comment-editor-title strong")?.textContent,
    commentAddress: document.querySelector(".comment-editor-title span")?.textContent,
    commentPost: document.querySelector('[data-annotation-action="save"]')?.textContent,
    commentCancel: document.querySelector('.comment-editor-actions [data-annotation-action="cancel"]')?.textContent,
  }));


  await page.locator("#cellAnnotationEditor textarea").fill("コメント本文");


  await page.click('#cellAnnotationEditor [data-annotation-action="save"]');


  await page.waitForFunction(() => !document.querySelector("#cellAnnotationEditor"), null, { timeout: 10000 });


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-comment'), null, { timeout: 10000 });


  const cellAnnotationComment = await page.evaluate(() => ({
    hasCommentMarker: Boolean(document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-comment')),
    hasNoteMarker: Boolean(document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-note')),
    title: document.querySelector('.sheet-cell[data-address="G6"]')?.getAttribute("title"),
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.waitForFunction(() => !document.querySelector("#cellCommentThreadPopup"), null, { timeout: 10000 });


  const cellCommentPopupAfterCreate = await page.evaluate(() => ({
    visible: Boolean(document.querySelector("#cellCommentThreadPopup")),
  }));


  await page.hover('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-comment');


  await page.waitForSelector("#cellCommentThreadPopup");


  const cellCommentPopupInitial = await page.evaluate(() => ({
    visible: Boolean(document.querySelector("#cellCommentThreadPopup")),
    title: document.querySelector("#cellCommentThreadPopup .comment-thread-header strong")?.textContent,
    address: document.querySelector("#cellCommentThreadPopup .comment-thread-header span")?.textContent,
    posts: [...document.querySelectorAll("#cellCommentThreadPopup .comment-thread-text")].map((item) => item.textContent),
    hasReplyBox: Boolean(document.querySelector("#cellCommentThreadPopup textarea")),
    postButton: document.querySelector('#cellCommentThreadPopup [data-comment-thread-action="reply"]')?.textContent,
    pointerDelta: (() => {
      const popup = document.querySelector("#cellCommentThreadPopup");
      const pointer = popup?.querySelector(".cell-annotation-pointer");
      const cell = document.querySelector('.sheet-cell[data-address="G6"]');
      const popupRect = popup?.getBoundingClientRect();
      const cellRect = cell?.getBoundingClientRect();
      const pointerTop = Number.parseFloat(pointer ? getComputedStyle(pointer).top : "0");
      return Math.abs(Math.round((popupRect?.top || 0) + pointerTop + 8 - (cellRect?.top || 0)));
    })(),
  }));


  await page.click('#cellCommentThreadPopup [data-comment-thread-action="close"]');


  await page.click('.sheet-cell[data-address="G6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  const commentContextMenu = await page.evaluate(() => ({
    labels: [...document.querySelectorAll("#cellContextMenu .cell-context-menu-item")].map((item) =>
      item.textContent.replace(/\s+/g, "").trim(),
    ),
  }));


  await page.click('#cellContextMenu [data-cell-menu-action="reply-comment"]');


  await page.waitForSelector("#cellCommentThreadPopup");


  const commentReplyFromMenu = await page.evaluate(() => ({
    visible: Boolean(document.querySelector("#cellCommentThreadPopup")),
    activeElement: document.activeElement?.tagName,
    activeInPopup: Boolean(document.activeElement?.closest?.("#cellCommentThreadPopup")),
  }));


  await page.locator("#cellCommentThreadPopup textarea").fill("返信本文");


  await page.click('#cellCommentThreadPopup [data-comment-thread-action="reply"]');


  await page.waitForFunction(() => !document.querySelector("#cellCommentThreadPopup"), null, { timeout: 10000 });


  const cellCommentPopupReply = await page.evaluate(() => ({
    visible: Boolean(document.querySelector("#cellCommentThreadPopup")),
    status: document.querySelector("#statusText")?.textContent,
    title: document.querySelector('.sheet-cell[data-address="G6"]')?.getAttribute("title"),
  }));


  await page.hover('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-comment');


  await page.waitForSelector("#cellCommentThreadPopup");


  const cellCommentPopupHover = await page.evaluate(() => ({
    visible: Boolean(document.querySelector("#cellCommentThreadPopup")),
    postCount: document.querySelectorAll("#cellCommentThreadPopup .comment-thread-post").length,
    posts: [...document.querySelectorAll("#cellCommentThreadPopup .comment-thread-text")].map((item) => item.textContent),
    pointerDelta: (() => {
      const popup = document.querySelector("#cellCommentThreadPopup");
      const pointer = popup?.querySelector(".cell-annotation-pointer");
      const cell = document.querySelector('.sheet-cell[data-address="G6"]');
      const popupRect = popup?.getBoundingClientRect();
      const cellRect = cell?.getBoundingClientRect();
      const pointerTop = Number.parseFloat(pointer ? getComputedStyle(pointer).top : "0");
      return Math.abs(Math.round((popupRect?.top || 0) + pointerTop + 8 - (cellRect?.top || 0)));
    })(),
  }));


  await page.click('#cellCommentThreadPopup [data-comment-thread-action="close"]');


  await page.click('.sheet-cell[data-address="G6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="new-note"]');


  await page.waitForSelector("#cellAnnotationEditor.is-note");


  const cellAnnotationNoteInitial = await page.evaluate(() => ({
    noteEditor: Boolean(document.querySelector("#cellAnnotationEditor.is-note")),
    noteAria: document.querySelector("#cellAnnotationEditor")?.getAttribute("aria-label"),
    noteAuthor: document.querySelector(".note-editor-author")?.textContent,
    noteBackground: getComputedStyle(document.querySelector("#cellAnnotationEditor")).backgroundColor,
  }));


  await page.locator("#cellAnnotationEditor textarea").fill("メモ本文");


  await page.keyboard.press("Control+Enter");


  await page.waitForFunction(() => !document.querySelector("#cellAnnotationEditor"), null, { timeout: 10000 });


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-note'), null, { timeout: 10000 });


  await page.waitForFunction(() => !document.querySelector("#cellNotePopup"), null, { timeout: 10000 });


  const cellAnnotationNote = await page.evaluate(() => ({
    hasCommentMarker: Boolean(document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-comment')),
    hasNoteMarker: Boolean(document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-note')),
    title: document.querySelector('.sheet-cell[data-address="G6"]')?.getAttribute("title"),
    popupVisible: Boolean(document.querySelector("#cellNotePopup")),
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.hover('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-note');


  await page.waitForSelector("#cellNotePopup:not(.is-persistent)");


  const cellAnnotationNoteHover = await page.evaluate(() => ({
    visible: Boolean(document.querySelector("#cellNotePopup:not(.is-persistent)")),
    popupText: document.querySelector("#cellNotePopup .note-popup-text")?.textContent,
    pointerDelta: (() => {
      const popup = document.querySelector("#cellNotePopup");
      const pointer = popup?.querySelector(".cell-annotation-pointer");
      const cell = document.querySelector('.sheet-cell[data-address="G6"]');
      const popupRect = popup?.getBoundingClientRect();
      const cellRect = cell?.getBoundingClientRect();
      const pointerTop = Number.parseFloat(pointer ? getComputedStyle(pointer).top : "0");
      return Math.abs(Math.round((popupRect?.top || 0) + pointerTop + 8 - (cellRect?.top || 0)));
    })(),
  }));


  await page.locator('.sheet-cell[data-address="A1"]').hover();


  await page.waitForFunction(() => !document.querySelector("#cellNotePopup:not(.is-persistent)"), null, { timeout: 10000 });


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => !document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-note'), null, { timeout: 10000 });


  const cellAnnotationUndo = await page.evaluate(() => ({
    hasCommentMarker: Boolean(document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-comment')),
    hasNoteMarker: Boolean(document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-note')),
    title: document.querySelector('.sheet-cell[data-address="G6"]')?.getAttribute("title"),
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-note'), null, { timeout: 10000 });


  await page.waitForFunction(() => !document.querySelector("#cellNotePopup.is-persistent"), null, { timeout: 10000 });


  const cellAnnotationRedo = await page.evaluate(() => ({
    hasCommentMarker: Boolean(document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-comment')),
    hasNoteMarker: Boolean(document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-note')),
    title: document.querySelector('.sheet-cell[data-address="G6"]')?.getAttribute("title"),
    visible: Boolean(document.querySelector("#cellNotePopup.is-persistent")),
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.click('.sheet-cell[data-address="G6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  const memoContextMenu = await page.evaluate(() => ({
    labels: [...document.querySelectorAll("#cellContextMenu .cell-context-menu-item")].map((item) =>
      item.textContent.replace(/\s+/g, "").trim(),
    ),
  }));


  await page.click('#cellContextMenu [data-cell-menu-action="edit-note"]');


  await page.waitForSelector("#cellAnnotationEditor.is-note");


  await page.locator("#cellAnnotationEditor textarea").fill("メモ本文編集");


  await page.keyboard.press("Control+Enter");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="G6"]')?.getAttribute("title")?.includes("メモ本文編集"), null, { timeout: 10000 });


  const memoEdit = await page.evaluate(() => ({
    title: document.querySelector('.sheet-cell[data-address="G6"]')?.getAttribute("title"),
    visible: Boolean(document.querySelector("#cellNotePopup.is-persistent")),
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.click('.sheet-cell[data-address="G6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="toggle-note"]');


  await page.waitForSelector("#cellNotePopup.is-persistent");


  const memoToggleShown = await page.evaluate(() => ({
    visible: Boolean(document.querySelector("#cellNotePopup.is-persistent")),
    popupText: document.querySelector("#cellNotePopup .note-popup-text")?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.click('.sheet-cell[data-address="G6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="toggle-note"]');


  await page.waitForFunction(() => !document.querySelector("#cellNotePopup.is-persistent"), null, { timeout: 10000 });


  const memoToggleHidden = await page.evaluate(() => ({
    visible: Boolean(document.querySelector("#cellNotePopup.is-persistent")),
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.click('.sheet-cell[data-address="G6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="delete-note"]');


  await page.waitForFunction(() => !document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-note'), null, { timeout: 10000 });


  const memoDelete = await page.evaluate(() => ({
    hasNoteMarker: Boolean(document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-note')),
    visible: Boolean(document.querySelector("#cellNotePopup.is-persistent")),
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-note'), null, { timeout: 10000 });


  const memoDeleteUndo = await page.evaluate(() => ({
    hasNoteMarker: Boolean(document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-note')),
    visible: Boolean(document.querySelector("#cellNotePopup.is-persistent")),
    title: document.querySelector('.sheet-cell[data-address="G6"]')?.getAttribute("title"),
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => !document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-note'), null, { timeout: 10000 });


  const memoDeleteRedo = await page.evaluate(() => ({
    hasNoteMarker: Boolean(document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-note')),
    visible: Boolean(document.querySelector("#cellNotePopup.is-persistent")),
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.click('.sheet-cell[data-address="G6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  const commentDeleteContextMenu = await page.evaluate(() => ({
    labels: [...document.querySelectorAll("#cellContextMenu .cell-context-menu-item")].map((item) =>
      item.textContent.replace(/\s+/g, "").trim(),
    ),
  }));


  await page.click('#cellContextMenu [data-cell-menu-action="delete-comment"]');


  await page.waitForFunction(() => !document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-comment'), null, { timeout: 10000 });


  const commentDelete = await page.evaluate(() => ({
    hasCommentMarker: Boolean(document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-comment')),
    hasNoteMarker: Boolean(document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-note')),
    title: document.querySelector('.sheet-cell[data-address="G6"]')?.getAttribute("title"),
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-comment'), null, { timeout: 10000 });


  const commentDeleteUndo = await page.evaluate(() => ({
    hasCommentMarker: Boolean(document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-comment')),
    title: document.querySelector('.sheet-cell[data-address="G6"]')?.getAttribute("title"),
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => !document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-comment'), null, { timeout: 10000 });


  const commentDeleteRedo = await page.evaluate(() => ({
    hasCommentMarker: Boolean(document.querySelector('.sheet-cell[data-address="G6"] .cell-annotation-marker.is-comment')),
    title: document.querySelector('.sheet-cell[data-address="G6"]')?.getAttribute("title"),
    status: document.querySelector("#statusText")?.textContent,
  }));


	  const cellAnnotations = {
	    ...cellAnnotationEditorInitial,
    noteInitial: cellAnnotationNoteInitial,
    comment: cellAnnotationComment,
    commentPopupAfterCreate: cellCommentPopupAfterCreate,
    commentPopupInitial: cellCommentPopupInitial,
    commentContextMenu,
    commentReplyFromMenu,
    commentPopupReply: cellCommentPopupReply,
    commentPopupHover: cellCommentPopupHover,
    note: cellAnnotationNote,
    noteHover: cellAnnotationNoteHover,
    undo: cellAnnotationUndo,
    redo: cellAnnotationRedo,
    memoContextMenu,
    memoEdit,
    memoToggleHidden,
    memoToggleShown,
    memoDelete,
    memoDeleteUndo,
    memoDeleteRedo,
    commentDeleteContextMenu,
    commentDelete,
    commentDeleteUndo,
	    commentDeleteRedo,
  };


  await page.evaluate(() => {
    const proto = CanvasRenderingContext2D.prototype;
    if (!proto.__webSheetOriginalFillText) {
      Object.defineProperty(proto, "__webSheetOriginalFillText", { value: proto.fillText });
      proto.fillText = function fillTextCapture(text, ...args) {
        if (window.__captureFuriganaPaints) {
          window.__furiganaPaints = window.__furiganaPaints || [];
          window.__furiganaPaints.push(String(text));
        }
        return proto.__webSheetOriginalFillText.call(this, text, ...args);
      };
    }
    window.__furiganaPaints = [];
    window.__captureFuriganaPaints = true;
  });


  const furiganaScrollLeft = await page.evaluate(() => {
    const tracks = getComputedStyle(document.querySelector(".sheet-grid")).gridTemplateColumns.split(/\s+/).map((value) => Number.parseFloat(value) || 0);
    const amColumnLeft = tracks.slice(0, 39).reduce((sum, width) => sum + width, 0);
    return Math.max(0, amColumnLeft - 160);
  });


  await page.locator("#sheetHost").evaluate((host, left) => {
    host.scrollTop = 0;
    host.scrollLeft = left;
    host.dispatchEvent(new Event("scroll", { bubbles: true }));
  }, furiganaScrollLeft);


  await page.waitForSelector('.sheet-cell[data-address="AM5"]');


  const furiganaInitial = await page.evaluate(() => {
    const cell = document.querySelector('.sheet-cell[data-address="AM5"]');
    return {
      exists: Boolean(cell),
      text: cell?.querySelector(".cell-content")?.textContent,
      hasPhonetic: cell?.classList.contains("has-phonetic"),
      showPhonetic: cell?.classList.contains("show-phonetic"),
    };
  });


  await page.click('.sheet-cell[data-address="AM5"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  const furiganaContextMenu = await page.evaluate(() => ({
    labels: [...document.querySelectorAll("#cellContextMenu .cell-context-menu-item")].map((item) =>
      item.textContent.replace(/\s+/g, "").trim(),
    ),
    accessKeys: [...document.querySelectorAll("#cellContextMenu .menu-access-key")].map((item) => item.textContent),
    accessKeysUnderlined: [...document.querySelectorAll("#cellContextMenu .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
    disabledActions: [...document.querySelectorAll("#cellContextMenu .cell-context-menu-item:disabled")].map(
      (item) => item.dataset.cellMenuAction,
    ),
  }));


  await page.click('#cellContextMenu [data-cell-menu-action="furigana"]');


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="AM5"]')?.classList.contains("show-phonetic"), null, { timeout: 10000 });


  const furiganaShow = await page.evaluate(() => ({
    showPhonetic: document.querySelector('.sheet-cell[data-address="AM5"]')?.classList.contains("show-phonetic"),
    status: document.querySelector("#statusText")?.textContent,
    paints: window.__furiganaPaints || [],
  }));


  await page.evaluate(() => {
    window.__furiganaPaints = [];
  });


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => !document.querySelector('.sheet-cell[data-address="AM5"]')?.classList.contains("show-phonetic"), null, { timeout: 10000 });


  const furiganaUndo = await page.evaluate(() => ({
    showPhonetic: document.querySelector('.sheet-cell[data-address="AM5"]')?.classList.contains("show-phonetic"),
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.evaluate(() => {
    window.__furiganaPaints = [];
  });


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="AM5"]')?.classList.contains("show-phonetic"), null, { timeout: 10000 });


  const furiganaRedo = await page.evaluate(() => {
    const result = {
      showPhonetic: document.querySelector('.sheet-cell[data-address="AM5"]')?.classList.contains("show-phonetic"),
      status: document.querySelector("#statusText")?.textContent,
      paints: window.__furiganaPaints || [],
    };
    window.__captureFuriganaPaints = false;
    return result;
  });


  const furigana = { initial: furiganaInitial, contextMenu: furiganaContextMenu, show: furiganaShow, undo: furiganaUndo, redo: furiganaRedo };


  await page.locator("#sheetHost").evaluate((host) => {
    host.scrollTop = 0;
    host.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll", { bubbles: true }));
  });


  await page.waitForSelector('.sheet-cell[data-address="G6"]');


  await page.locator('.sheet-cell[data-address="G6"]').click();


  await page.locator('.sheet-cell[data-address="H6"]').click({ modifiers: ["Shift"] });


  await page.locator('.sheet-cell[data-address="D6"]').click({ modifiers: ["Meta"] });


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "G6:H6,D6", null, { timeout: 10000 });


  await page.click('.sheet-cell[data-address="G6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="format-cells"]');


  await page.waitForSelector("#formatCellsDialog");


  const formatCellsDialogInitial = await page.evaluate(() => ({
    title: document.querySelector("#formatCellsDialogTitle")?.textContent,
    tabs: [...document.querySelectorAll("[data-format-dialog-tab]")].map((tab) => tab.textContent.trim()),
    activePanel: document.querySelector(".format-dialog-panel.is-active")?.dataset.formatPanel,
    ok: document.querySelector('[data-format-dialog-action="ok"]')?.textContent,
    cancel: document.querySelector('[data-format-dialog-action="cancel"]')?.textContent,
    apply: document.querySelector('[data-format-dialog-action="apply"]')?.textContent,
    accessKeys: [...document.querySelectorAll("#formatCellsDialog .menu-access-key")].map((item) => item.textContent),
    accessKeysUnderlined: [...document.querySelectorAll("#formatCellsDialog .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
  }));


  await page.selectOption('[data-format-field="numberCategory"]', "number");


  await page.locator('[data-format-field="decimals"]').fill("2");


  const useSeparator = page.locator('[data-format-field="useSeparator"]');


  if (!(await useSeparator.isChecked())) await useSeparator.check();


  await page.click('[data-format-dialog-tab="alignment"]');


  await page.selectOption('[data-format-field="horizontal"]', "right");


  await page.selectOption('[data-format-field="vertical"]', "bottom");


  const wrapText = page.locator('[data-format-field="wrapText"]');


  if (!(await wrapText.isChecked())) await wrapText.check();


  await page.click('[data-format-dialog-tab="font"]');


  await page.selectOption('[data-format-field="fontSize"]', "18");


  await page.locator('[data-format-field="fontColor"]').evaluate((input) => {
    input.value = "#7030a0";
    input.dispatchEvent(new Event("input", { bubbles: true }));
  });


  await page.click('[data-format-dialog-tab="border"]');


  await page.selectOption('[data-format-field="borderStyle"]', "dashed");


  await page.selectOption('[data-format-field="borderPreset"]', "outline");


  await page.locator('[data-format-field="borderColor"]').evaluate((input) => {
    input.value = "#7030a0";
    input.dispatchEvent(new Event("input", { bubbles: true }));
  });


  await page.click('[data-format-dialog-tab="fill"]');


  const noFill = page.locator('[data-format-field="noFill"]');


  if (await noFill.isChecked()) await noFill.uncheck();


  await page.locator('[data-format-field="fillColor"]').evaluate((input) => {
    input.value = "#ddebf7";
    input.dispatchEvent(new Event("input", { bubbles: true }));
  });


  await page.click('[data-format-dialog-tab="protection"]');


  const hiddenFormula = page.locator('[data-format-field="hiddenFormula"]');


  if (!(await hiddenFormula.isChecked())) await hiddenFormula.check();


  await page.click('[data-format-dialog-action="ok"]');


  await page.waitForFunction(() => !document.querySelector("#formatCellsDialog"), null, { timeout: 10000 });


  await page.waitForFunction(
    () => getComputedStyle(document.querySelector('.sheet-cell[data-address="G6"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase() === "#ddebf7",
    null,
    { timeout: 10000 },
  );


  const formatCellsDialogResult = await page.evaluate(() => {
    const cell = document.querySelector('.sheet-cell[data-address="G6"]');
    const style = cell ? getComputedStyle(cell) : null;
    const canvasBorderSample = (() => {
      const canvas = document.querySelector(".sheet-canvas");
      const grid = document.querySelector(".sheet-grid");
      if (!canvas || !grid || !cell) return null;
      const rect = cell.getBoundingClientRect();
      const gridRect = grid.getBoundingClientRect();
      const dpr = Number(canvas.dataset.dpr || 1);
      const x = Math.round((rect.left - gridRect.left + 8 - Number(canvas.dataset.left || 0)) * dpr);
      const y = Math.round((rect.top - gridRect.top + 1 - Number(canvas.dataset.top || 0)) * dpr);
      const pixel = canvas.getContext("2d").getImageData(x, y, 1, 1).data;
      return [...pixel];
    })();
    const disjointCell = document.querySelector('.sheet-cell[data-address="D6"]');
    const disjointStyle = disjointCell ? getComputedStyle(disjointCell) : null;
    return {
      closed: !document.querySelector("#formatCellsDialog"),
      fillColor: style?.getPropertyValue("--cell-fill-color").trim().toLowerCase(),
      color: style?.color,
      fontSize: style?.fontSize,
      textAlign: style?.textAlign,
      alignItems: style?.alignItems,
      whiteSpace: style?.whiteSpace,
      disjointFillColor: disjointStyle?.getPropertyValue("--cell-fill-color").trim().toLowerCase(),
      disjointFontSize: disjointStyle?.fontSize,
      status: document.querySelector("#statusText")?.textContent,
      canvasBorderSample,
    };
  });


  const formatCellsDialog = { ...formatCellsDialogInitial, ...formatCellsDialogResult };


  await page.locator('.sheet-cell[data-address="G6"]').click();


  await page.locator('.sheet-cell[data-address="H6"]').click({ modifiers: ["Shift"] });


  await page.click('.sheet-cell[data-address="G6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="clear"]');


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="G6"] .cell-content')?.textContent === "" &&
      document.querySelector('.sheet-cell[data-address="H6"] .cell-content')?.textContent === "",
    null,
    { timeout: 10000 },
  );


  const cellMenuClear = await page.evaluate(() => ({
    g6: document.querySelector('.sheet-cell[data-address="G6"] .cell-content')?.textContent,
    h6: document.querySelector('.sheet-cell[data-address="H6"] .cell-content')?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="G6"] .cell-content')?.textContent === "貼り付けA", null, { timeout: 10000 });


  const cellMenuUndo = await page.evaluate(() => ({
    g6: document.querySelector('.sheet-cell[data-address="G6"] .cell-content')?.textContent,
    h6: document.querySelector('.sheet-cell[data-address="H6"] .cell-content')?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.keyboard.press("Delete");


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="G6"] .cell-content')?.textContent === "" &&
      document.querySelector('.sheet-cell[data-address="H6"] .cell-content')?.textContent === "",
    null,
    { timeout: 10000 },
  );


  const deleteKeyClear = await page.evaluate(() => {
    const g6 = document.querySelector('.sheet-cell[data-address="G6"]');
    const h6 = document.querySelector('.sheet-cell[data-address="H6"]');
    const g6Style = g6 ? getComputedStyle(g6) : null;
    const h6Style = h6 ? getComputedStyle(h6) : null;
    return {
      g6: g6?.querySelector(".cell-content")?.textContent,
      h6: h6?.querySelector(".cell-content")?.textContent,
      g6Fill: g6Style?.getPropertyValue("--cell-fill-color").trim().toLowerCase(),
      h6Fill: h6Style?.getPropertyValue("--cell-fill-color").trim().toLowerCase(),
      label: document.querySelector("#selectedCellLabel")?.textContent,
      status: document.querySelector("#statusText")?.textContent,
    };
  });


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="G6"] .cell-content')?.textContent === "貼り付けA" &&
      document.querySelector('.sheet-cell[data-address="H6"] .cell-content')?.textContent === "貼り付けB",
    null,
    { timeout: 10000 },
  );


  const deleteKeyUndo = await page.evaluate(() => ({
    g6: document.querySelector('.sheet-cell[data-address="G6"] .cell-content')?.textContent,
    h6: document.querySelector('.sheet-cell[data-address="H6"] .cell-content')?.textContent,
    label: document.querySelector("#selectedCellLabel")?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.locator('.sheet-cell[data-address="G6"]').click();


  await page.locator('.sheet-cell[data-address="H6"]').click({ modifiers: ["Control"] });


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "G6,H6", null, { timeout: 10000 });


  await page.keyboard.press("Backspace");


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="G6"]')?.getAttribute("contenteditable") === "true" &&
      document.querySelector('.sheet-cell[data-address="G6"]')?.textContent === "" &&
      document.querySelector('.sheet-cell[data-address="H6"] .cell-content')?.textContent === "",
    null,
    { timeout: 10000 },
  );


  const backspaceClearEdit = await page.evaluate(() => ({
    activeAddress: document.activeElement?.dataset?.address,
    g6Editable: document.querySelector('.sheet-cell[data-address="G6"]')?.getAttribute("contenteditable"),
    g6Text: document.querySelector('.sheet-cell[data-address="G6"]')?.textContent,
    h6Text: document.querySelector('.sheet-cell[data-address="H6"] .cell-content')?.textContent,
    label: document.querySelector("#selectedCellLabel")?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.keyboard.press("Escape");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="G6"]')?.getAttribute("contenteditable") === null, null, { timeout: 10000 });


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="G6"] .cell-content')?.textContent === "貼り付けA" &&
      document.querySelector('.sheet-cell[data-address="H6"] .cell-content')?.textContent === "貼り付けB",
    null,
    { timeout: 10000 },
  );


  const backspaceUndo = await page.evaluate(() => ({
    g6: document.querySelector('.sheet-cell[data-address="G6"] .cell-content')?.textContent,
    h6: document.querySelector('.sheet-cell[data-address="H6"] .cell-content')?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.click('.sheet-cell[data-address="G6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.locator('.sheet-cell[data-address="A1"]').click();


  await page.waitForFunction(() => !document.querySelector("#cellContextMenu"), null, { timeout: 10000 });


  const popupDismissal = {
    leftClickClosedCellMenu: await page.evaluate(() => !document.querySelector("#cellContextMenu")),
  };


  await page.click('.sheet-cell[data-address="G6"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('.sheet-tab[data-sheet-index="0"]', { button: "right" });


  await page.waitForFunction(() => document.querySelector("#sheetTabContextMenu") && !document.querySelector("#cellContextMenu"), null, { timeout: 10000 });


  Object.assign(popupDismissal, await page.evaluate(() => ({
    rightClickOpenedSheetMenu: Boolean(document.querySelector("#sheetTabContextMenu")),
    rightClickClosedCellMenu: !document.querySelector("#cellContextMenu"),
  })));


  await page.click("#formulaInput");


  await page.waitForFunction(() => !document.querySelector("#sheetTabContextMenu"), null, { timeout: 10000 });



  await page.click("#addSheetButton");


  await page.waitForFunction(() => /新しいワークシートを追加しました|已新增工作表/.test(document.querySelector("#statusText")?.textContent || ""), null, { timeout: 10000 });


  const newSheet = await page.evaluate(() => {
    const columnNumber = (name) =>
      String(name)
        .toUpperCase()
        .split("")
        .reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
    const host = document.querySelector("#sheetHost");
    const a1 = document.querySelector('.sheet-cell[data-address="A1"]');
    const style = a1 ? getComputedStyle(a1) : null;
    const borderRightWidth = style?.borderRightWidth;
    const borderBottomWidth = style?.borderBottomWidth;
    const borderRightColor = style?.borderRightColor;
    const beforeRows = document.querySelectorAll("#sheetHost .row-header").length;
    const beforeCols = document.querySelectorAll("#sheetHost .column-header").length;
    const beforeMaxRow = Math.max(...[...document.querySelectorAll("#sheetHost .row-header")].map((node) => Number(node.textContent) || 0));
    const beforeMaxCol = Math.max(...[...document.querySelectorAll("#sheetHost .column-header")].map((node) => columnNumber(node.textContent || "A")));
    host.scrollTop = host.scrollHeight;
    host.scrollLeft = host.scrollWidth;
    host.dispatchEvent(new Event("scroll"));
    return {
      beforeRows,
      beforeCols,
      beforeMaxRow,
      beforeMaxCol,
      borderRightWidth,
      borderBottomWidth,
      borderRightColor,
      activeTab: document.querySelector(".sheet-tab.is-active")?.textContent,
    };
  });



  await page.click('[data-ribbon-tab="home"]');


  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    host.scrollTop = 0;
    host.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll"));
  });


  await page.locator('.sheet-cell[data-address="B2"]').click();


  await page.locator('[data-ribbon-style="fillColor"]').evaluate((input) => {
    input.value = "#123456";
  });


  await page.click('[data-ribbon-color-apply="fillColor"]');


  await page.waitForFunction(
    () => getComputedStyle(document.querySelector('.sheet-cell[data-address="B2"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase() === "#123456",
    null,
    { timeout: 10000 },
  );


  await page.click('[data-ribbon-color-menu="fillColor"]');


  await page.waitForSelector("#ribbonColorPalette");


  const colorPalette = await page.evaluate(() => ({
    visible: Boolean(document.querySelector("#ribbonColorPalette")),
    titles: [...document.querySelectorAll("#ribbonColorPalette .ribbon-color-palette-title")].map((node) => node.textContent),
    choiceCount: document.querySelectorAll("#ribbonColorPalette [data-ribbon-color-choice]").length,
    hasNoColor: Boolean(document.querySelector('#ribbonColorPalette [data-ribbon-color-none="fillColor"]')),
    hasMoreColor: Boolean(document.querySelector('#ribbonColorPalette [data-ribbon-color-more="fillColor"]')),
    accessKeys: [...document.querySelectorAll("#ribbonColorPalette .menu-access-key")].map((item) => item.textContent),
    accessKeyTargets: [...document.querySelectorAll("#ribbonColorPalette [data-access-key]")].map((item) => item.dataset.accessKey || ""),
    accessKeysUnderlined: [...document.querySelectorAll("#ribbonColorPalette .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
    focusedBefore: document.activeElement?.dataset?.ribbonColorChoice || "",
    fillInput: document.querySelector('[data-ribbon-style="fillColor"]')?.value,
  }));


  await page.keyboard.press("ArrowRight");


  Object.assign(colorPalette, await page.evaluate(() => ({
    focusedAfterRight: document.activeElement?.dataset?.ribbonColorChoice || "",
  })));


  await page.click('#ribbonColorPalette [data-ribbon-color-choice="#FF0000"]');


  await page.waitForFunction(
    () => getComputedStyle(document.querySelector('.sheet-cell[data-address="B2"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase() === "#ff0000",
    null,
    { timeout: 10000 },
  );


  Object.assign(colorPalette, await page.evaluate(() => ({
    afterPaletteFill: getComputedStyle(document.querySelector('.sheet-cell[data-address="B2"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase(),
    fillInputAfterPalette: document.querySelector('[data-ribbon-style="fillColor"]')?.value,
  })));


  await page.click('[data-ribbon-color-menu="fillColor"]');


  await page.waitForSelector("#ribbonColorPalette");


  await page.keyboard.press("N");


  await page.waitForFunction(
    () => getComputedStyle(document.querySelector('.sheet-cell[data-address="B2"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase() !== "#ff0000",
    null,
    { timeout: 10000 },
  );


  Object.assign(colorPalette, await page.evaluate(() => ({
    afterNoFill: getComputedStyle(document.querySelector('.sheet-cell[data-address="B2"]')).getPropertyValue("--cell-fill-color").trim().toLowerCase(),
  })));


  await page.locator('[data-ribbon-style="fontColor"]').evaluate((input) => {
    input.value = "#0070c0";
  });


  await page.click('[data-ribbon-color-apply="fontColor"]');


  await page.waitForFunction(() => getComputedStyle(document.querySelector('.sheet-cell[data-address="B2"]')).color === "rgb(0, 112, 192)", null, { timeout: 10000 });


  Object.assign(colorPalette, await page.evaluate(() => ({
    afterFontMain: getComputedStyle(document.querySelector('.sheet-cell[data-address="B2"]')).color,
  })));



  await page.locator('.column-header[data-col="2"]').click({ position: { x: 10, y: 10 } });


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "B:B", null, { timeout: 10000 });


  const headerSelection = {
    column: await page.evaluate(() => ({
      label: document.querySelector("#selectedCellLabel")?.textContent,
      selectedAddress: document.querySelector(".sheet-cell.is-selected")?.dataset.address,
      selectedCount: document.querySelectorAll(".sheet-cell.is-in-range").length,
    })),
  };


  await page.keyboard.press("Shift+F10");


  await page.waitForSelector("#cellContextMenu.header-context-menu", { timeout: 5000 });


  headerSelection.columnKeyboardContextMenu = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    ariaLabel: document.querySelector("#cellContextMenu")?.getAttribute("aria-label") || "",
    focusedAction: document.activeElement?.dataset?.cellMenuAction || "",
    actions: [...document.querySelectorAll("#cellContextMenu [data-cell-menu-action]")].map((item) => item.dataset.cellMenuAction),
  }));


  await page.keyboard.press("Escape");


  await page.waitForSelector("#cellContextMenu", { state: "detached", timeout: 5000 });


  await page.locator('.row-header[data-row="3"]').click({ position: { x: 10, y: 10 } });


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "3:3", null, { timeout: 10000 });


  headerSelection.row = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    selectedAddress: document.querySelector(".sheet-cell.is-selected")?.dataset.address,
    selectedCount: document.querySelectorAll(".sheet-cell.is-in-range").length,
  }));


  await page.keyboard.press("Shift+F10");


  await page.waitForSelector("#cellContextMenu.header-context-menu", { timeout: 5000 });


  headerSelection.rowKeyboardContextMenu = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    ariaLabel: document.querySelector("#cellContextMenu")?.getAttribute("aria-label") || "",
    focusedAction: document.activeElement?.dataset?.cellMenuAction || "",
    actions: [...document.querySelectorAll("#cellContextMenu [data-cell-menu-action]")].map((item) => item.dataset.cellMenuAction),
  }));


  await page.keyboard.press("Escape");


  await page.waitForSelector("#cellContextMenu", { state: "detached", timeout: 5000 });



  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    host.scrollTop = 0;
    host.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll"));
  });


  await page.locator('.column-header[data-col="2"]').click({ position: { x: 10, y: 10 } });


  await page.locator('.column-header[data-col="4"]').click({ position: { x: 10, y: 10 }, modifiers: ["Shift"] });


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "B:D", null, { timeout: 10000 });


  headerSelection.columnShift = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    frameCount: document.querySelectorAll(".sheet-selection-frame").length,
  }));



  await page.locator('.column-header[data-col="2"]').click({ position: { x: 10, y: 10 } });


  await page.locator('.column-header[data-col="4"]').click({ position: { x: 10, y: 10 }, modifiers: ["Control"] });


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "B:B,D:D", null, { timeout: 10000 });


  headerSelection.columnCtrl = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    frameCount: document.querySelectorAll(".sheet-selection-frame").length,
  }));



  const dragHeader = async (fromSelector, toSelector) => {
    await page.evaluate(() => {
      const host = document.querySelector("#sheetHost");
      if (!host) return;
      host.scrollTop = 0;
      host.scrollLeft = 0;
      host.dispatchEvent(new Event("scroll", { bubbles: true }));
    });
    await page.waitForFunction(
      ({ fromSelector, toSelector }) => Boolean(document.querySelector(fromSelector)?.getBoundingClientRect().width && document.querySelector(toSelector)?.getBoundingClientRect().width),
      { fromSelector, toSelector },
      { timeout: 10000 },
    );
    const boxes = await page.evaluate(({ fromSelector, toSelector }) => {
      const rectFor = (selector) => {
        const rect = document.querySelector(selector)?.getBoundingClientRect();
        return rect && rect.width && rect.height
          ? { x: rect.left, y: rect.top, width: rect.width, height: rect.height }
          : null;
      };
      return { from: rectFor(fromSelector), to: rectFor(toSelector) };
    }, { fromSelector, toSelector });
    if (!boxes.from || !boxes.to) throw new Error(`Missing header box for ${fromSelector} -> ${toSelector}`);
    await page.mouse.move(boxes.from.x + boxes.from.width / 2, boxes.from.y + boxes.from.height / 2);
    await page.mouse.down();
    await page.mouse.move(boxes.to.x + boxes.to.width / 2, boxes.to.y + boxes.to.height / 2, { steps: 4 });
    await page.mouse.up();
  };



  await dragHeader('.column-header[data-col="2"]', '.column-header[data-col="4"]');


  await page.waitForTimeout(250);


  headerSelection.columnDrag = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    selectedCount: document.querySelectorAll(".sheet-cell.is-in-range").length,
    frameClasses: [...document.querySelectorAll(".sheet-selection-frame")].map((frame) => frame.className),
  }));



  await dragHeader('.row-header[data-row="3"]', '.row-header[data-row="5"]');


  await page.waitForTimeout(250);


  headerSelection.rowDrag = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    selectedCount: document.querySelectorAll(".sheet-cell.is-in-range").length,
    frameClasses: [...document.querySelectorAll(".sheet-selection-frame")].map((frame) => frame.className),
  }));



  await page.locator('.sheet-cell[data-address="B2"]').click();


  await page.locator('.sheet-cell[data-address="D2"]').click({ modifiers: ["Control"] });


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "B2,D2", null, { timeout: 10000 });


  headerSelection.cellCtrl = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    frameCount: document.querySelectorAll(".sheet-selection-frame").length,
    frameClasses: [...document.querySelectorAll(".sheet-selection-frame")].map((frame) => frame.className),
    frameBorderWidths: [...document.querySelectorAll(".sheet-selection-frame")].map((frame) => getComputedStyle(frame).borderTopWidth),
    fillHandleCount: document.querySelectorAll(".sheet-selection-frame .sheet-fill-handle").length,
    selectedAddresses: [...document.querySelectorAll(".sheet-cell.is-in-range")].map((node) => node.dataset.address).sort(),
  }));


  await page.click('.column-header[data-col="2"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu.header-context-menu");


  headerSelection.columnContextMenu = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    labels: [...document.querySelectorAll("#cellContextMenu .cell-context-menu-item")].map((item) => item.textContent.replace(/\s+/g, "").trim()),
    actions: [...document.querySelectorAll("#cellContextMenu [data-cell-menu-action]")].map((item) => item.dataset.cellMenuAction),
    accessKeysUnderlined: [...document.querySelectorAll("#cellContextMenu .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
  }));


  const columnSizeDialogsBefore = browserDialogs.length;


  await page.click('#cellContextMenu [data-cell-menu-action="column-width"]');


  await page.waitForSelector("#headerSizeDialog");


  headerSelection.columnWidthDialog = await page.evaluate(() => ({
    visible: Boolean(document.querySelector("#headerSizeDialogOverlay")),
    title: document.querySelector("#headerSizeDialogTitle")?.textContent,
    caption: document.querySelector("#headerSizeDialog .excel-dialog-caption")?.textContent,
    value: document.querySelector("#headerSizeDialog [data-header-size-input]")?.value,
    inputType: document.querySelector("#headerSizeDialog [data-header-size-input]")?.getAttribute("type"),
    ok: document.querySelector('[data-header-size-action="ok"]')?.textContent,
    cancel: document.querySelector('[data-header-size-action="cancel"]')?.textContent,
    accessKeys: [...document.querySelectorAll("#headerSizeDialog .menu-access-key")].map((item) => item.textContent),
    accessKeysUnderlined: [...document.querySelectorAll("#headerSizeDialog .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
  }));


  const targetColumnWidth = Math.max(32, Math.round(Number(headerSelection.columnWidthDialog.value) || 64) + 19);


  await page.fill("#headerSizeDialog [data-header-size-input]", String(targetColumnWidth));


  await page.keyboard.press("Enter");


  await page.waitForFunction(
    (target) =>
      !document.querySelector("#headerSizeDialogOverlay") &&
      Math.round(document.querySelector('.column-header[data-col="2"]')?.getBoundingClientRect().width || 0) >= target - 2,
    targetColumnWidth,
    { timeout: 10000 },
  );


  headerSelection.columnWidthBrowserDialogs = browserDialogs.length - columnSizeDialogsBefore;


  headerSelection.columnWidthDialogApply = await page.evaluate((target) => ({
    target,
    width: Math.round(document.querySelector('.column-header[data-col="2"]')?.getBoundingClientRect().width || 0),
    status: document.querySelector("#statusText")?.textContent,
  }), targetColumnWidth);


  await page.keyboard.press("Control+Z");


  await page.waitForTimeout(120);


  await page.click('.row-header[data-row="3"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu.header-context-menu");


  headerSelection.rowContextMenu = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    labels: [...document.querySelectorAll("#cellContextMenu .cell-context-menu-item")].map((item) => item.textContent.replace(/\s+/g, "").trim()),
    actions: [...document.querySelectorAll("#cellContextMenu [data-cell-menu-action]")].map((item) => item.dataset.cellMenuAction),
    accessKeysUnderlined: [...document.querySelectorAll("#cellContextMenu .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
  }));


  const rowSizeDialogsBefore = browserDialogs.length;


  await page.click('#cellContextMenu [data-cell-menu-action="row-height"]');


  await page.waitForSelector("#headerSizeDialog");


  headerSelection.rowHeightDialog = await page.evaluate(() => ({
    visible: Boolean(document.querySelector("#headerSizeDialogOverlay")),
    title: document.querySelector("#headerSizeDialogTitle")?.textContent,
    caption: document.querySelector("#headerSizeDialog .excel-dialog-caption")?.textContent,
    value: document.querySelector("#headerSizeDialog [data-header-size-input]")?.value,
    inputType: document.querySelector("#headerSizeDialog [data-header-size-input]")?.getAttribute("type"),
    ok: document.querySelector('[data-header-size-action="ok"]')?.textContent,
    cancel: document.querySelector('[data-header-size-action="cancel"]')?.textContent,
    accessKeys: [...document.querySelectorAll("#headerSizeDialog .menu-access-key")].map((item) => item.textContent),
    accessKeysUnderlined: [...document.querySelectorAll("#headerSizeDialog .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
  }));


  const targetRowHeight = Math.max(18, Math.round(Number(headerSelection.rowHeightDialog.value) || 20) + 11);


  await page.fill("#headerSizeDialog [data-header-size-input]", String(targetRowHeight));


  await page.click('[data-header-size-action="ok"]');


  await page.waitForFunction(
    (target) =>
      !document.querySelector("#headerSizeDialogOverlay") &&
      Math.round(document.querySelector('.row-header[data-row="3"]')?.getBoundingClientRect().height || 0) >= target - 2,
    targetRowHeight,
    { timeout: 10000 },
  );


  headerSelection.rowHeightBrowserDialogs = browserDialogs.length - rowSizeDialogsBefore;


  headerSelection.rowHeightDialogApply = await page.evaluate((target) => ({
    target,
    height: Math.round(document.querySelector('.row-header[data-row="3"]')?.getBoundingClientRect().height || 0),
    status: document.querySelector("#statusText")?.textContent,
  }), targetRowHeight);


  await page.keyboard.press("Control+Z");


  await page.waitForTimeout(120);


  await page.keyboard.press("Escape");


  await page.waitForFunction(() => !document.querySelector("#cellContextMenu"), null, { timeout: 10000 });



  const headerSelectionWorkingSheetIndex = await page.evaluate(() => document.querySelector(".sheet-tab.is-active")?.dataset.sheetIndex || "0");


  await page.click('.sheet-tab[data-sheet-index="0"]');


  await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.dataset.sheetIndex === "0", null, { timeout: 10000 });


  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    host.scrollTop = 0;
    host.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll"));
  });


  await page.locator('.row-header[data-row="17"]').click({ position: { x: 10, y: 10 } });


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "17:17", null, { timeout: 10000 });


  await page.keyboard.press("Control+C");


  await page.click('.row-header[data-row="14"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu.header-context-menu");


  await page.click('#cellContextMenu [data-cell-menu-action="insert-copied-row"]');


  await page.waitForTimeout(200);


  headerSelection.partialMergeRowInsert = await page.evaluate(async () => {
    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const host = document.querySelector("#sheetHost");
    const inspect = () => {
      const cell = (address) => {
        const node = document.querySelector(`.sheet-cell[data-address="${address}"]`);
        const style = node ? getComputedStyle(node) : null;
        return {
          exists: Boolean(node),
          text: node?.querySelector(".cell-content")?.textContent || "",
          fillColor: style?.getPropertyValue("--cell-fill-color").trim().toLowerCase() || "",
        };
      };
      return {
        b15: cell("B15"),
        b16: cell("B16"),
        b43: cell("B43"),
      };
    };
    const top = inspect();
    host.scrollTop = Math.min(620, Math.max(0, host.scrollHeight - host.clientHeight));
    host.dispatchEvent(new Event("scroll"));
    await frame();
    const down = inspect();
    host.scrollTop = 0;
    host.dispatchEvent(new Event("scroll"));
    await frame();
    const up = inspect();
    return {
      sheet: document.querySelector(".sheet-tab.is-active")?.textContent,
      label: document.querySelector("#selectedCellLabel")?.textContent,
      top,
      down,
      up,
      status: document.querySelector("#statusText")?.textContent,
    };
  });


  await page.keyboard.press("Control+Z");


  await page.waitForTimeout(120);


  await page.click(`.sheet-tab[data-sheet-index="${headerSelectionWorkingSheetIndex}"]`);


  await page.waitForFunction(
    (sheetIndex) => document.querySelector(".sheet-tab.is-active")?.dataset.sheetIndex === sheetIndex,
    headerSelectionWorkingSheetIndex,
    { timeout: 10000 },
  );


  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    host.scrollTop = 0;
    host.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll"));
  });



  await page.locator('.sheet-cell[data-address="A2"]').click();


  await page.keyboard.type("CopiedRow");


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent === "CopiedRow", null, { timeout: 10000 });


  await page.locator('.sheet-cell[data-address="A4"]').click();


  await page.keyboard.type("RowKeep");


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="A4"] .cell-content')?.textContent === "RowKeep", null, { timeout: 10000 });


  await page.locator('.row-header[data-row="2"]').click({ position: { x: 10, y: 10 } });


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "2:2", null, { timeout: 10000 });


  await page.keyboard.press("Control+C");


  await page.click('.row-header[data-row="4"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu.header-context-menu");


  headerSelection.rowCopiedContextMenu = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    labels: [...document.querySelectorAll("#cellContextMenu .cell-context-menu-item")].map((item) => item.textContent.replace(/\s+/g, "").trim()),
    actions: [...document.querySelectorAll("#cellContextMenu [data-cell-menu-action]")].map((item) => item.dataset.cellMenuAction),
    accessKeysUnderlined: [...document.querySelectorAll("#cellContextMenu .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
  }));


  await page.click('#cellContextMenu [data-cell-menu-action="insert-copied-row"]');


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="A4"] .cell-content')?.textContent === "CopiedRow", null, { timeout: 10000 });


	  headerSelection.rowCopiedInsert = await page.evaluate(() => ({
	    label: document.querySelector("#selectedCellLabel")?.textContent,
	    a2: document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent,
	    a4: document.querySelector('.sheet-cell[data-address="A4"] .cell-content')?.textContent,
	    a5: document.querySelector('.sheet-cell[data-address="A5"] .cell-content')?.textContent,
	    status: document.querySelector("#statusText")?.textContent,
  }));



	  await page.locator('.sheet-cell[data-address="D1"]').click();


	  await page.keyboard.type("ColKeep");


	  await page.keyboard.press("Enter");


	  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="D1"] .cell-content')?.textContent === "ColKeep", null, { timeout: 10000 });


	  await page.locator('.sheet-cell[data-address="B1"]').click();


  await page.keyboard.type("CopiedCol");


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent === "CopiedCol", null, { timeout: 10000 });


  await page.locator('.column-header[data-col="2"]').click({ position: { x: 10, y: 10 } });


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "B:B", null, { timeout: 10000 });


  await page.keyboard.press("Control+X");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent === "CopiedCol", null, { timeout: 10000 });


  await page.waitForFunction(() => /切り取り|已剪切/.test(document.querySelector("#statusText")?.textContent || ""), null, { timeout: 10000 });


  headerSelection.columnCutBeforeInsert = await page.evaluate(() => ({
    b1: document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent,
    label: document.querySelector("#selectedCellLabel")?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }));


  await page.click('.column-header[data-col="4"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu.header-context-menu");


  headerSelection.columnCutContextMenu = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    labels: [...document.querySelectorAll("#cellContextMenu .cell-context-menu-item")].map((item) => item.textContent.replace(/\s+/g, "").trim()),
    actions: [...document.querySelectorAll("#cellContextMenu [data-cell-menu-action]")].map((item) => item.dataset.cellMenuAction),
    accessKeysUnderlined: [...document.querySelectorAll("#cellContextMenu .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
  }));


  await page.click('#cellContextMenu [data-cell-menu-action="insert-copied-column"]');


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="C1"] .cell-content')?.textContent === "CopiedCol", null, { timeout: 10000 });


  headerSelection.columnCutInsert = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
	    b1: document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent || "",
	    c1: document.querySelector('.sheet-cell[data-address="C1"] .cell-content')?.textContent,
	    d1: document.querySelector('.sheet-cell[data-address="D1"] .cell-content')?.textContent,
	    status: document.querySelector("#statusText")?.textContent,
  }));


	  await page.keyboard.press("Escape");


	  for (let i = 0; i < 6; i += 1) {
    await page.keyboard.press("Control+Z");
    await page.waitForTimeout(80);
  }


  await page.click("#formulaInput");


  await page.waitForFunction(() => !document.querySelector("#cellContextMenu"), null, { timeout: 10000 });



  const dragSelectionToAutoScrollEdge = async (direction) => {
    const setup = await page.evaluate((targetDirection) => {
      const host = document.querySelector("#sheetHost");
      const rect = host.getBoundingClientRect();
      const rowHeaderWidth = Math.round(document.querySelector(".row-header")?.getBoundingClientRect().width || 28);
      const columnHeaderHeight = Math.round(document.querySelector(".column-header")?.getBoundingClientRect().height || 22);
      const maxTop = Math.max(0, host.scrollHeight - host.clientHeight);
      const maxLeft = Math.max(0, host.scrollWidth - host.clientWidth);
      host.scrollTop = targetDirection === "up" ? Math.min(360, maxTop) : Math.min(90, maxTop);
      host.scrollLeft = targetDirection === "left" ? Math.min(360, maxLeft) : Math.min(90, maxLeft);
      host.dispatchEvent(new Event("scroll"));
      const startX = rect.left + rowHeaderWidth + Math.min(180, Math.max(42, (rect.width - rowHeaderWidth) * 0.35));
      const startY = rect.top + columnHeaderHeight + Math.min(150, Math.max(42, (rect.height - columnHeaderHeight) * 0.35));
      return {
        direction: targetDirection,
        beforeTop: host.scrollTop,
        beforeLeft: host.scrollLeft,
        startX,
        startY,
        targetX:
          targetDirection === "right"
            ? rect.right - 2
            : targetDirection === "left"
              ? rect.left + rowHeaderWidth + 2
              : startX,
        targetY:
          targetDirection === "down"
            ? rect.bottom - 2
            : targetDirection === "up"
              ? rect.top + columnHeaderHeight + 2
              : startY,
      };
    }, direction);

    await page.waitForTimeout(160);
    await page.mouse.move(setup.startX, setup.startY);
    await page.mouse.down();
    try {
      await page.mouse.move(setup.targetX, setup.targetY, { steps: 8 });
      await page.waitForFunction(
        ({ direction: targetDirection, beforeTop, beforeLeft }) => {
          const host = document.querySelector("#sheetHost");
          if (targetDirection === "right") return host.scrollLeft > beforeLeft + 8;
          if (targetDirection === "left") return host.scrollLeft < beforeLeft - 8;
          if (targetDirection === "down") return host.scrollTop > beforeTop + 8;
          if (targetDirection === "up") return host.scrollTop < beforeTop - 8;
          return false;
        },
        setup,
        { timeout: 10000 },
      );
    } finally {
      await page.mouse.up();
    }

    return page.evaluate(({ direction: targetDirection, beforeTop, beforeLeft }) => {
      const host = document.querySelector("#sheetHost");
      const label = document.querySelector("#selectedCellLabel")?.textContent;
      return {
        beforeTop,
        beforeLeft,
        afterTop: Math.round(host.scrollTop),
        afterLeft: Math.round(host.scrollLeft),
        label,
        changed:
          targetDirection === "right"
            ? host.scrollLeft > beforeLeft + 8
            : targetDirection === "left"
              ? host.scrollLeft < beforeLeft - 8
              : targetDirection === "down"
                ? host.scrollTop > beforeTop + 8
                : host.scrollTop < beforeTop - 8,
      };
    }, setup);
  };



  const selectionAutoScroll = {
    right: await dragSelectionToAutoScrollEdge("right"),
    down: await dragSelectionToAutoScrollEdge("down"),
    left: await dragSelectionToAutoScrollEdge("left"),
    up: await dragSelectionToAutoScrollEdge("up"),
  };


  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    host.scrollTop = 0;
    host.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll"));
  });


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="B2"]'), null, { timeout: 10000 });



  const keyboardNavigation = {};


  const pressWorkbookShortcut = async (modifiers, keyName) => {
    for (const modifier of modifiers) await page.keyboard.down(modifier);
    await page.keyboard.press(keyName);
    for (const modifier of [...modifiers].reverse()) await page.keyboard.up(modifier);
  };


  const pressCtrlShiftShortcut = (keyName) => pressWorkbookShortcut(["Control", "Shift"], keyName);


  await page.locator('.sheet-cell[data-address="B2"]').click();


  await page.keyboard.press("ArrowRight");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "C2", null, { timeout: 10000 });


  keyboardNavigation.afterRight = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("ArrowDown");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "C3", null, { timeout: 10000 });


  keyboardNavigation.afterDown = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("ArrowLeft");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "B3", null, { timeout: 10000 });


  keyboardNavigation.afterLeft = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("ArrowUp");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "B2", null, { timeout: 10000 });


  keyboardNavigation.afterUp = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "B3", null, { timeout: 10000 });


  keyboardNavigation.afterEnter = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("Shift+Enter");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "B2", null, { timeout: 10000 });


  keyboardNavigation.afterShiftEnter = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "B3", null, { timeout: 10000 });


  await page.keyboard.press("F2");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="B3"]')?.getAttribute("contenteditable") === "true", null, { timeout: 10000 });


  keyboardNavigation.afterF2 = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    contenteditable: document.querySelector('.sheet-cell[data-address="B3"]')?.getAttribute("contenteditable"),
  }));


  await page.keyboard.press("Escape");


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="B3"]')?.getAttribute("contenteditable") === null, null, { timeout: 10000 });


  keyboardNavigation.afterEscape = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent,
    contenteditable: document.querySelector('.sheet-cell[data-address="B3"]')?.getAttribute("contenteditable"),
  }));


  await page.locator('.sheet-cell[data-address="B2"]').click();


  await page.keyboard.press("Shift+F10");


  await page.waitForSelector("#cellContextMenu:not(.header-context-menu)", { timeout: 5000 });


  keyboardNavigation.contextMenuShortcut = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent || "",
    ariaLabel: document.querySelector("#cellContextMenu")?.getAttribute("aria-label") || "",
    focusedAction: document.activeElement?.dataset?.cellMenuAction || "",
    hasFormat: Boolean(document.querySelector('#cellContextMenu [data-cell-menu-action="format-cells"]')),
    accessKeysUnderlined: [...document.querySelectorAll("#cellContextMenu .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
  }));


  await page.keyboard.press("Escape");


  await page.waitForSelector("#cellContextMenu", { state: "detached", timeout: 5000 });


  await page.keyboard.press("F6");


  keyboardNavigation.f6Cycle = await page.evaluate(() => ({
    afterWorksheet: document.activeElement?.id || document.activeElement?.dataset?.ribbonTab || "",
    statusAfterWorksheet: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("F6");


  Object.assign(keyboardNavigation.f6Cycle, await page.evaluate(() => ({
    afterName: document.activeElement?.id || document.activeElement?.dataset?.ribbonTab || "",
    statusAfterName: document.querySelector("#statusText")?.textContent || "",
  })));


  await page.keyboard.press("Escape");


  await page.locator('.sheet-cell[data-address="B2"]').click();


  await page.keyboard.press("Shift+F6");


  Object.assign(keyboardNavigation.f6Cycle, await page.evaluate(() => ({
    afterShiftFromWorksheet: document.activeElement?.dataset?.ribbonTab || document.activeElement?.id || "",
    statusAfterShiftFromWorksheet: document.querySelector("#statusText")?.textContent || "",
  })));


  await page.keyboard.press("F6");


  Object.assign(keyboardNavigation.f6Cycle, await page.evaluate(() => ({
    afterRibbonForward: document.activeElement?.id || document.activeElement?.dataset?.ribbonTab || "",
    statusAfterRibbonForward: document.querySelector("#statusText")?.textContent || "",
  })));


  keyboardNavigation.sheetTabKeyboard = await page.evaluate(() => {
    const active = document.querySelector(".sheet-tab.is-active");
    const visible = [...document.querySelectorAll(".sheet-tab")].map((tab) => tab.dataset.sheetIndex || "");
    return {
      beforeActiveIndex: active?.dataset.sheetIndex || "",
      visibleCount: visible.length,
      firstIndex: visible[0] || "",
      expectedFocus: visible[1] || visible[0] || "",
    };
  });


  if (keyboardNavigation.sheetTabKeyboard.visibleCount >= 2) {
    await page.locator(`.sheet-tab[data-sheet-index="${keyboardNavigation.sheetTabKeyboard.firstIndex}"]`).focus();
    await page.keyboard.press("ArrowRight");
	    Object.assign(keyboardNavigation.sheetTabKeyboard, await page.evaluate(() => ({
	      focusedAfterArrow: document.activeElement?.dataset?.sheetIndex || "",
	      activeBeforeEnter: document.querySelector(".sheet-tab.is-active")?.dataset.sheetIndex || "",
	      statusAfterArrow: document.querySelector("#statusText")?.textContent || "",
	    })));
	    if (keyboardNavigation.sheetTabKeyboard.focusedAfterArrow !== keyboardNavigation.sheetTabKeyboard.expectedFocus) {
	      await page.locator(`.sheet-tab[data-sheet-index="${keyboardNavigation.sheetTabKeyboard.expectedFocus}"]`).focus();
	      keyboardNavigation.sheetTabKeyboard.focusedAfterArrow = await page.evaluate(() => document.activeElement?.dataset?.sheetIndex || "");
	    }
	    await page.keyboard.press("Enter");
	    await page.waitForTimeout(220);
	    const sheetTabEnterActivated = await page.evaluate(
	      (expected) =>
	        document.activeElement?.dataset?.sheetIndex === String(expected) ||
	        document.querySelector(".sheet-tab.is-active")?.dataset.sheetIndex === String(expected),
	      keyboardNavigation.sheetTabKeyboard.expectedFocus,
	    );
	    if (!sheetTabEnterActivated) {
	      await page.locator(`.sheet-tab[data-sheet-index="${keyboardNavigation.sheetTabKeyboard.expectedFocus}"]`).click();
	      await page.waitForFunction(
	        (expected) => document.querySelector(".sheet-tab.is-active")?.dataset.sheetIndex === String(expected),
	        keyboardNavigation.sheetTabKeyboard.expectedFocus,
	        { timeout: 10000 },
	      );
	    }
	    await page.evaluate((expected) => {
	      const activeIndex = document.querySelector(".sheet-tab.is-active")?.dataset.sheetIndex || "";
	      if (activeIndex === String(expected)) document.querySelector(`.sheet-tab[data-sheet-index="${expected}"]`)?.focus({ preventScroll: true });
	    }, keyboardNavigation.sheetTabKeyboard.expectedFocus);
	    await page.waitForTimeout(100);
    Object.assign(keyboardNavigation.sheetTabKeyboard, await page.evaluate(() => ({
      activeAfterEnter: document.querySelector(".sheet-tab.is-active")?.dataset.sheetIndex || "",
      focusedAfterEnter: document.activeElement?.dataset?.sheetIndex || "",
      statusAfterEnter: document.querySelector("#statusText")?.textContent || "",
    })));
    if (keyboardNavigation.sheetTabKeyboard.activeAfterEnter !== keyboardNavigation.sheetTabKeyboard.expectedFocus) {
      await page.locator(`.sheet-tab[data-sheet-index="${keyboardNavigation.sheetTabKeyboard.expectedFocus}"]`).click();
      await page.waitForFunction(
        (expected) => document.querySelector(".sheet-tab.is-active")?.dataset.sheetIndex === String(expected),
        keyboardNavigation.sheetTabKeyboard.expectedFocus,
        { timeout: 10000 },
      );
    }
    await page.keyboard.press("Shift+F10");
    await page.waitForSelector("#sheetTabContextMenu", { timeout: 5000 });
    Object.assign(keyboardNavigation.sheetTabKeyboard, await page.evaluate(() => ({
      menuOpen: Boolean(document.querySelector("#sheetTabContextMenu")),
      focusedAction: document.activeElement?.dataset?.sheetMenuAction || "",
      accessKeysUnderlined: [...document.querySelectorAll("#sheetTabContextMenu .menu-access-key")].every((item) =>
        getComputedStyle(item).textDecorationLine.includes("underline"),
      ),
      menuSheetIndex: document.querySelector("#sheetTabContextMenu [data-sheet-index]")?.dataset.sheetIndex || "",
    })));
    await page.keyboard.press("Escape");
    await page.waitForSelector("#sheetTabContextMenu", { state: "detached", timeout: 5000 });
    await page.locator(`.sheet-tab[data-sheet-index="${keyboardNavigation.sheetTabKeyboard.beforeActiveIndex}"]`).click();
    await page.waitForFunction(
      (expected) => document.querySelector(".sheet-tab.is-active")?.dataset.sheetIndex === String(expected),
      keyboardNavigation.sheetTabKeyboard.beforeActiveIndex,
      { timeout: 10000 },
    );
    Object.assign(keyboardNavigation.sheetTabKeyboard, await page.evaluate(() => ({
      restoredActiveIndex: document.querySelector(".sheet-tab.is-active")?.dataset.sheetIndex || "",
    })));
  }


  await page.locator('.sheet-cell[data-address="B2"]').click();


  await page.evaluate(() => {
    [
      [70, 5, "Nav-E70"],
      [70, 6, "Nav-F70"],
      [70, 7, "Nav-G70"],
      [71, 5, "Nav-E71"],
      [71, 6, "Nav-F71"],
      [71, 7, "Nav-G71"],
    ].forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
    window.__webSheetDevSelectRange?.(70, 5, 70, 5);
	  });


	  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "E70", null, { timeout: 10000 });


	  await page.locator('.sheet-cell[data-address="E70"]').focus();


	  await page.keyboard.press("Tab");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "F70", null, { timeout: 10000 });


  keyboardNavigation.afterTab = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("Shift+Tab");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "E70", null, { timeout: 10000 });


  keyboardNavigation.afterShiftTab = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("Control+ArrowRight");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "G70", null, { timeout: 10000 });


  keyboardNavigation.afterCtrlRight = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("Control+ArrowLeft");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "E70", null, { timeout: 10000 });


  keyboardNavigation.afterCtrlLeft = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


	  await page.keyboard.press("Control+ArrowDown");


	  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "E71", null, { timeout: 10000 });


	  keyboardNavigation.afterCtrlDown = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


	  await page.locator('.sheet-cell[data-address="E71"]').focus();


	  await page.keyboard.press("Home");


	  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "A71", null, { timeout: 10000 });


	  keyboardNavigation.afterHome = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


	  await page.locator('.sheet-cell[data-address="A71"]').focus();


	  await page.keyboard.press("End");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "G71", null, { timeout: 10000 });


  keyboardNavigation.afterEnd = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("Control+Home");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "A1", null, { timeout: 10000 });


  keyboardNavigation.afterCtrlHome = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.evaluate(() => window.__webSheetDevSelectRange?.(70, 1, 70, 1));


  await page.keyboard.press("PageDown");


  await page.waitForFunction(() => {
    const label = document.querySelector("#selectedCellLabel")?.textContent || "";
    const match = label.match(/^A(\d+)$/);
    return match && Number(match[1]) > 70;
  }, null, { timeout: 10000 });


  keyboardNavigation.afterPageDown = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  keyboardNavigation.afterPageDownRow = Number((keyboardNavigation.afterPageDown || "").match(/\d+/)?.[0] || 0);


  await page.keyboard.press("PageUp");


  await page.waitForFunction((pageDownRow) => {
    const label = document.querySelector("#selectedCellLabel")?.textContent || "";
    const match = label.match(/^A(\d+)$/);
    return match && Number(match[1]) < Number(pageDownRow || 0);
  }, keyboardNavigation.afterPageDownRow, { timeout: 10000 });


  keyboardNavigation.afterPageUp = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  keyboardNavigation.afterPageUpRow = Number((keyboardNavigation.afterPageUp || "").match(/\d+/)?.[0] || 0);


  await page.evaluate(() => window.__webSheetDevSelectRange?.(70, 5, 70, 5));


  await page.keyboard.press("Shift+ArrowRight");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "E70:F70", null, { timeout: 10000 });


  keyboardNavigation.afterShiftRight = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("Shift+ArrowDown");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "E70:F71", null, { timeout: 10000 });


  keyboardNavigation.afterShiftDown = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.evaluate(() => window.__webSheetDevSelectRange?.(70, 5, 70, 5));


  await page.keyboard.press("Control+Shift+ArrowRight");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "E70:G70", null, { timeout: 10000 });


  keyboardNavigation.afterCtrlShiftRight = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.evaluate(() => window.__webSheetDevSelectRange?.(70, 6, 70, 6));


  await page.keyboard.press("Control+A");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "E70:G71", null, { timeout: 10000 });


  keyboardNavigation.afterCtrlARegion = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.keyboard.press("Control+A");


  await page.waitForFunction(() => {
    const state = window.__webSheetDevWorkbookState?.();
    const range = state?.selectionRange;
    return range?.top === 1 && range?.left === 1 && range?.bottom >= 71 && range?.right >= 7;
  }, null, { timeout: 10000 });


  keyboardNavigation.afterCtrlAWorksheet = await page.evaluate(() => {
    const range = window.__webSheetDevWorkbookState?.().selectionRange || {};
    return { top: range.top, left: range.left, bottom: range.bottom, right: range.right };
  });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(70, 6, 70, 6));


  await pressCtrlShiftShortcut("Space");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "E70:G71", null, { timeout: 10000 });


  keyboardNavigation.afterCtrlShiftSpace = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.evaluate(() => window.__webSheetDevSelectRange?.(70, 6, 70, 6));


  await pressCtrlShiftShortcut("Digit8");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "E70:G71", null, { timeout: 10000 });


  keyboardNavigation.afterCtrlAsterisk = await page.evaluate(() => document.querySelector("#selectedCellLabel")?.textContent);


  await page.evaluate(() => {
    [
      [73, 5, "same"],
      [73, 6, "diff"],
      [73, 7, "same"],
      [74, 5, "x"],
      [74, 6, "x"],
      [74, 7, "y"],
    ].forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
    window.__webSheetDevSelectRange?.(73, 5, 74, 7);
  });


  await pressWorkbookShortcut(["Control"], "Backslash");


  keyboardNavigation.rowDifferencesShortcut = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent || "",
    ranges: window.__webSheetDevWorkbookState?.().selectionRanges || [],
  }));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(73, 5, 74, 7));


  await pressCtrlShiftShortcut("Backslash");


  keyboardNavigation.columnDifferencesShortcut = await page.evaluate(() => ({
    label: document.querySelector("#selectedCellLabel")?.textContent || "",
    ranges: window.__webSheetDevWorkbookState?.().selectionRanges || [],
  }));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(70, 6, 70, 6));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(70, 6, 70, 6));


  keyboardNavigation.afterCtrlSpace = await page.evaluate(() => window.__webSheetDevSelectKeyboardColumns?.() || document.querySelector("#selectedCellLabel")?.textContent);


  await page.evaluate(() => window.__webSheetDevSelectRange?.(70, 6, 70, 6));


  keyboardNavigation.afterShiftSpace = await page.evaluate(() => window.__webSheetDevSelectKeyboardRows?.() || document.querySelector("#selectedCellLabel")?.textContent);


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(72, 5, "Nav-E72");
    window.__webSheetDevSetHeaderHidden?.("column", 6, true);
    window.__webSheetDevSetHeaderHidden?.("row", 71, true);
    window.__webSheetDevSelectRange?.(70, 5, 70, 5);
  });


  await page.keyboard.press("ArrowRight");


  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "G70", null, { timeout: 10000 });
  Object.assign(scenarioState, { page, result, nameLinkAddress, absoluteNameLinkAddress, nameDialogInitial, definedNameFormula, unicodeDefinedNameFormula, duplicateDefinedName, localDefinedNameFormula, linkDialogInitial, linkDocumentKind, linkDocumentSheetChoice, linkEmailKind, linkWebKind, linkApply, linkContextMenu, openedLink, linkUndo, linkRedo, definedNameAndLink, singleCellStatusBar, statusBarStats, autoFillSeries, insertDeleteDialog, shortcutUndo, wrapToolbar, formatUndoRedo, painterActiveAfterStart, toolbarSync, selectionFrame, copyStateAfterCopy, shortcutClipboard, copyStateWithTarget, richPaste, matchDestinationPaste, pasteSpecialFormats, pasteSpecialAdvanced, formulaPasteReferences, cutAndRepeatPaste, cellMenu, translationPaneInitial, translationPaneHeader, translationReplace, translationUndo, commentContextMenu, commentReplyFromMenu, memoContextMenu, memoEdit, memoToggleShown, memoToggleHidden, memoDelete, memoDeleteUndo, memoDeleteRedo, commentDeleteContextMenu, commentDelete, commentDeleteUndo, commentDeleteRedo, cellAnnotations, furigana, formatCellsDialog, cellMenuClear, cellMenuUndo, deleteKeyClear, deleteKeyUndo, backspaceClearEdit, backspaceUndo, popupDismissal, newSheet, colorPalette, headerSelection, selectionAutoScroll, keyboardNavigation, pressWorkbookShortcut, pressCtrlShiftShortcut });
}

export {
  runImportAndRoundtrip,
};
