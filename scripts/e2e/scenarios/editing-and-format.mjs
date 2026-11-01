// Sequential E2E feature module.

async function runEditingAndFormat(scenarioState) {
  let { page, result, sheetGroupQuickAnalysisBaseCount, sheetGroupQuickAnalysisSheets, sheetGroupQuickAnalysisKeyboard, sheetGroupQuickAnalysisBars, sheetGroupQuickAnalysisBarsUndo, sheetGroupQuickAnalysisBarsRedo, sheetGroupQuickAnalysisClear, sheetGroupQuickAnalysisClearUndo, sheetGroupQuickAnalysisColorScale, sheetGroupQuickAnalysisColorScaleUndo, sheetGroupQuickAnalysisColorScaleRedo, sheetGroupQuickAnalysisIconSet, sheetGroupQuickAnalysisIconUndo, sheetGroupQuickAnalysisIconRedo, sheetGroupQuickAnalysisDuplicates, sheetGroupQuickAnalysisDuplicatesUndo, sheetGroupQuickAnalysisDuplicatesRedo } = scenarioState;

  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupQuickAnalysisSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupQuickAnalysisSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 1, 3, 1));


  await page.click('.sheet-cell[data-address="A1"]', { button: "right" });


  await page.waitForSelector("#cellContextMenu", { timeout: 5000 });


  await page.click('#cellContextMenu [data-cell-menu-action="quick-analysis"]');


  await page.waitForSelector("#quickAnalysisPopup", { timeout: 5000 });


  await page.click('#quickAnalysisPopup [data-quick-analysis-tab="totals"]');


  await page.click('#quickAnalysisPopup [data-quick-analysis-action="total-sum"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 4, 1)?.formula === "=SUM($A$1:$A$3)" &&
    window.__webSheetDevCellModel?.(second.name, 4, 1)?.formula === "=SUM($A$1:$A$3)",
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisTotal = await page.evaluate(([first, second]) => ({
    firstA4: window.__webSheetDevCellModel?.(first.name, 4, 1)?.formula || "",
    secondA4: window.__webSheetDevCellModel?.(second.name, 4, 1)?.formula || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupQuickAnalysisSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 4, 1)?.formula &&
    !window.__webSheetDevCellModel?.(second.name, 4, 1)?.formula,
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisTotalUndo = await page.evaluate(([first, second]) => ({
    firstA4: window.__webSheetDevCellModel?.(first.name, 4, 1)?.formula || "",
    secondA4: window.__webSheetDevCellModel?.(second.name, 4, 1)?.formula || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupQuickAnalysisSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 4, 1)?.formula === "=SUM($A$1:$A$3)" &&
    window.__webSheetDevCellModel?.(second.name, 4, 1)?.formula === "=SUM($A$1:$A$3)",
    sheetGroupQuickAnalysisSheets,
    { timeout: 5000 },
  );


  const sheetGroupQuickAnalysisTotalRedo = await page.evaluate(([first, second]) => ({
    firstA4: window.__webSheetDevCellModel?.(first.name, 4, 1)?.formula || "",
    secondA4: window.__webSheetDevCellModel?.(second.name, 4, 1)?.formula || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupQuickAnalysisSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupQuickAnalysisSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupQuickAnalysisSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupQuickAnalysisSheets[0].index}"]`, { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupQuickAnalysisBaseCount, { timeout: 5000 });


  const sheetGroupQuickAnalysis = {
    keyboard: sheetGroupQuickAnalysisKeyboard,
    bars: sheetGroupQuickAnalysisBars,
    barsUndo: sheetGroupQuickAnalysisBarsUndo,
    barsRedo: sheetGroupQuickAnalysisBarsRedo,
    clear: sheetGroupQuickAnalysisClear,
    clearUndo: sheetGroupQuickAnalysisClearUndo,
    colorScale: sheetGroupQuickAnalysisColorScale,
    colorScaleUndo: sheetGroupQuickAnalysisColorScaleUndo,
    colorScaleRedo: sheetGroupQuickAnalysisColorScaleRedo,
    iconSet: sheetGroupQuickAnalysisIconSet,
    iconUndo: sheetGroupQuickAnalysisIconUndo,
    iconRedo: sheetGroupQuickAnalysisIconRedo,
    duplicates: sheetGroupQuickAnalysisDuplicates,
    duplicatesUndo: sheetGroupQuickAnalysisDuplicatesUndo,
    duplicatesRedo: sheetGroupQuickAnalysisDuplicatesRedo,
    total: sheetGroupQuickAnalysisTotal,
    totalUndo: sheetGroupQuickAnalysisTotalUndo,
    totalRedo: sheetGroupQuickAnalysisTotalRedo,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };



  const sheetGroupSparklinesBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupSparklinesBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupSparklinesBaseCount, { timeout: 5000 });


  const sheetGroupSparklinesSheets = await page.evaluate(() =>
    [...document.querySelectorAll(".sheet-tab")]
      .slice(-2)
      .map((button) => ({ name: button.textContent, index: button.dataset.sheetIndex })),
  );


  await page.evaluate(([first, second]) => {
    window.__webSheetDevSetActiveSheet?.(first.name);
    [[1, 2, 3], [4, 5, 6]].forEach((rowValues, rowIndex) => {
      rowValues.forEach((value, colIndex) => window.__webSheetDevSetCell?.(rowIndex + 1, colIndex + 1, value));
    });
    window.__webSheetDevSetActiveSheet?.(second.name);
    [[10, 20, 30], [40, 50, 60]].forEach((rowValues, rowIndex) => {
      rowValues.forEach((value, colIndex) => window.__webSheetDevSetCell?.(rowIndex + 1, colIndex + 1, value));
    });
  }, sheetGroupSparklinesSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupSparklinesSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupSparklinesSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 1, 2, 3));


  await page.click('[data-ribbon-tab="insert"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "挿入", null, { timeout: 5000 });


  await page.click('[data-ribbon-command="insert-sparkline-line"]');


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 1, 4)?.sparkline?.type === "line" &&
    window.__webSheetDevCellModel?.(second.name, 1, 4)?.sparkline?.type === "line",
    sheetGroupSparklinesSheets,
    { timeout: 5000 },
  );


  const sheetGroupSparklinesCreate = await page.evaluate(([first, second]) => ({
    firstD1: window.__webSheetDevCellModel?.(first.name, 1, 4)?.sparkline || null,
    firstD2: window.__webSheetDevCellModel?.(first.name, 2, 4)?.sparkline || null,
    secondD1: window.__webSheetDevCellModel?.(second.name, 1, 4)?.sparkline || null,
    secondD2: window.__webSheetDevCellModel?.(second.name, 2, 4)?.sparkline || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupSparklinesSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !window.__webSheetDevCellModel?.(first.name, 1, 4)?.sparkline &&
    !window.__webSheetDevCellModel?.(second.name, 1, 4)?.sparkline,
    sheetGroupSparklinesSheets,
    { timeout: 5000 },
  );


  const sheetGroupSparklinesUndo = await page.evaluate(([first, second]) => ({
    firstD1: window.__webSheetDevCellModel?.(first.name, 1, 4)?.sparkline || null,
    firstD2: window.__webSheetDevCellModel?.(first.name, 2, 4)?.sparkline || null,
    secondD1: window.__webSheetDevCellModel?.(second.name, 1, 4)?.sparkline || null,
    secondD2: window.__webSheetDevCellModel?.(second.name, 2, 4)?.sparkline || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupSparklinesSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    window.__webSheetDevCellModel?.(first.name, 1, 4)?.sparkline?.type === "line" &&
    window.__webSheetDevCellModel?.(second.name, 1, 4)?.sparkline?.type === "line",
    sheetGroupSparklinesSheets,
    { timeout: 5000 },
  );


  const sheetGroupSparklinesRedo = await page.evaluate(([first, second]) => ({
    firstD1: window.__webSheetDevCellModel?.(first.name, 1, 4)?.sparkline || null,
    firstD2: window.__webSheetDevCellModel?.(first.name, 2, 4)?.sparkline || null,
    secondD1: window.__webSheetDevCellModel?.(second.name, 1, 4)?.sparkline || null,
    secondD2: window.__webSheetDevCellModel?.(second.name, 2, 4)?.sparkline || null,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupSparklinesSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupSparklinesSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupSparklinesSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupSparklinesSheets[0].index}"]`, { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupSparklinesBaseCount, { timeout: 5000 });


  const sheetGroupSparklines = {
    create: sheetGroupSparklinesCreate,
    undo: sheetGroupSparklinesUndo,
    redo: sheetGroupSparklinesRedo,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };



  const sheetGroupChartsBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupChartsBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupChartsBaseCount, { timeout: 5000 });


  const sheetGroupChartsSheets = await page.evaluate(() =>
    [...document.querySelectorAll(".sheet-tab")]
      .slice(-2)
      .map((button) => ({ name: button.textContent, index: button.dataset.sheetIndex })),
  );


  await page.evaluate(([first, second]) => {
    window.__webSheetDevSetActiveSheet?.(first.name);
    [["区分", "売上"], ["A", 1], ["B", 2]].forEach((rowValues, rowIndex) => {
      rowValues.forEach((value, colIndex) => window.__webSheetDevSetCell?.(rowIndex + 1, colIndex + 1, value));
    });
    window.__webSheetDevSetActiveSheet?.(second.name);
    [["区分", "売上"], ["A", 10], ["B", 20]].forEach((rowValues, rowIndex) => {
      rowValues.forEach((value, colIndex) => window.__webSheetDevSetCell?.(rowIndex + 1, colIndex + 1, value));
    });
  }, sheetGroupChartsSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupChartsSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupChartsSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 1, 3, 2));


  await page.click('[data-ribbon-tab="insert"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "挿入", null, { timeout: 5000 });


  await page.click('[data-ribbon-command="insert-column-chart"]');


  await page.waitForFunction(([first, second]) =>
    (window.__webSheetDevSheetModel?.(first.name)?.images || []).some((image) => image.chart?.type === "column") &&
    (window.__webSheetDevSheetModel?.(second.name)?.images || []).some((image) => image.chart?.type === "column"),
    sheetGroupChartsSheets,
    { timeout: 5000 },
  );


  const sheetGroupChartsCreate = await page.evaluate(([first, second]) => {
    const chartFor = (sheetName) => {
      const sheet = window.__webSheetDevSheetModel?.(sheetName) || {};
      const image = (sheet.images || []).find((item) => item.chart?.type === "column") || null;
      return {
        count: (sheet.images || []).filter((item) => item.chart).length,
        type: image?.chart?.type || "",
        sourceSheet: image?.chart?.sourceSheet || "",
        sourceRange: image?.chart?.sourceRange || "",
        values: image?.chart?.series?.[0]?.values || [],
      };
    };
    return {
      first: chartFor(first.name),
      second: chartFor(second.name),
      selectedObjects: document.querySelectorAll(".sheet-image.is-selected, .sheet-image-link.is-selected").length,
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupChartsSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !(window.__webSheetDevSheetModel?.(first.name)?.images || []).some((image) => image.chart?.type === "column") &&
    !(window.__webSheetDevSheetModel?.(second.name)?.images || []).some((image) => image.chart?.type === "column"),
    sheetGroupChartsSheets,
    { timeout: 5000 },
  );


  const sheetGroupChartsUndo = await page.evaluate(([first, second]) => ({
    firstCount: (window.__webSheetDevSheetModel?.(first.name)?.images || []).filter((item) => item.chart).length,
    secondCount: (window.__webSheetDevSheetModel?.(second.name)?.images || []).filter((item) => item.chart).length,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupChartsSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    (window.__webSheetDevSheetModel?.(first.name)?.images || []).some((image) => image.chart?.type === "column") &&
    (window.__webSheetDevSheetModel?.(second.name)?.images || []).some((image) => image.chart?.type === "column"),
    sheetGroupChartsSheets,
    { timeout: 5000 },
  );


  const sheetGroupChartsRedo = await page.evaluate(([first, second]) => {
    const chartFor = (sheetName) => {
      const sheet = window.__webSheetDevSheetModel?.(sheetName) || {};
      const image = (sheet.images || []).find((item) => item.chart?.type === "column") || null;
      return {
        count: (sheet.images || []).filter((item) => item.chart).length,
        type: image?.chart?.type || "",
        sourceSheet: image?.chart?.sourceSheet || "",
        values: image?.chart?.series?.[0]?.values || [],
      };
    };
    return {
      first: chartFor(first.name),
      second: chartFor(second.name),
      selectedObjects: document.querySelectorAll(".sheet-image.is-selected, .sheet-image-link.is-selected").length,
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupChartsSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupChartsSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupChartsSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupChartsSheets[0].index}"]`, { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupChartsBaseCount, { timeout: 5000 });


  const sheetGroupCharts = {
    create: sheetGroupChartsCreate,
    undo: sheetGroupChartsUndo,
    redo: sheetGroupChartsRedo,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };



  const sheetGroupTablesBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupTablesBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupTablesBaseCount, { timeout: 5000 });


  const sheetGroupTablesSheets = await page.evaluate(() =>
    [...document.querySelectorAll(".sheet-tab")]
      .slice(-2)
      .map((button) => ({ name: button.textContent, index: button.dataset.sheetIndex })),
  );


  await page.evaluate(([first, second]) => {
    window.__webSheetDevSetActiveSheet?.(first.name);
    [["名前", "値"], ["A", 1], ["B", 2]].forEach((rowValues, rowIndex) => {
      rowValues.forEach((value, colIndex) => window.__webSheetDevSetCell?.(rowIndex + 1, colIndex + 1, value));
    });
    window.__webSheetDevSetActiveSheet?.(second.name);
    [["名前", "値"], ["A", 10], ["B", 20]].forEach((rowValues, rowIndex) => {
      rowValues.forEach((value, colIndex) => window.__webSheetDevSetCell?.(rowIndex + 1, colIndex + 1, value));
    });
  }, sheetGroupTablesSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupTablesSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupTablesSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 1, 3, 2));


  await page.click('[data-ribbon-tab="insert"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "挿入", null, { timeout: 5000 });


  await page.click('[data-ribbon-command="insert-table"]');


  await page.waitForSelector('[data-data-dialog="create-table"]', { timeout: 5000 });


  await page.click('[data-data-dialog-action="create-table-ok"]');


  await page.waitForFunction(([first, second]) =>
    (window.__webSheetDevSheetModel?.(first.name)?.tables || []).length === 1 &&
    (window.__webSheetDevSheetModel?.(second.name)?.tables || []).length === 1,
    sheetGroupTablesSheets,
    { timeout: 5000 },
  );


  const sheetGroupTablesCreate = await page.evaluate(([first, second]) => {
    const tableFor = (sheetName) => {
      const sheet = window.__webSheetDevSheetModel?.(sheetName) || {};
      const table = (sheet.tables || [])[0] || {};
      return {
        count: (sheet.tables || []).length,
        range: table.range || null,
        hasHeaders: table.hasHeaders,
        showFilterButton: table.showFilterButton,
        a2: window.__webSheetDevCellModel?.(sheetName, 2, 1)?.raw ?? "",
        b2: window.__webSheetDevCellModel?.(sheetName, 2, 2)?.raw ?? "",
        headerFill: window.__webSheetDevCellModel?.(sheetName, 1, 1)?.css?.backgroundColor || "",
      };
    };
    return {
      first: tableFor(first.name),
      second: tableFor(second.name),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupTablesSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !(window.__webSheetDevSheetModel?.(first.name)?.tables || []).length &&
    !(window.__webSheetDevSheetModel?.(second.name)?.tables || []).length,
    sheetGroupTablesSheets,
    { timeout: 5000 },
  );


  const sheetGroupTablesUndo = await page.evaluate(([first, second]) => ({
    firstCount: (window.__webSheetDevSheetModel?.(first.name)?.tables || []).length,
    secondCount: (window.__webSheetDevSheetModel?.(second.name)?.tables || []).length,
    firstHeaderFill: window.__webSheetDevCellModel?.(first.name, 1, 1)?.css?.backgroundColor || "",
    secondHeaderFill: window.__webSheetDevCellModel?.(second.name, 1, 1)?.css?.backgroundColor || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupTablesSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    (window.__webSheetDevSheetModel?.(first.name)?.tables || []).length === 1 &&
    (window.__webSheetDevSheetModel?.(second.name)?.tables || []).length === 1,
    sheetGroupTablesSheets,
    { timeout: 5000 },
  );


  const sheetGroupTablesRedo = await page.evaluate(([first, second]) => ({
    firstCount: (window.__webSheetDevSheetModel?.(first.name)?.tables || []).length,
    secondCount: (window.__webSheetDevSheetModel?.(second.name)?.tables || []).length,
    firstB2: window.__webSheetDevCellModel?.(first.name, 2, 2)?.raw ?? "",
    secondB2: window.__webSheetDevCellModel?.(second.name, 2, 2)?.raw ?? "",
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupTablesSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupTablesSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupTablesSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupTablesSheets[0].index}"]`, { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupTablesBaseCount, { timeout: 5000 });


  const sheetGroupTables = {
    create: sheetGroupTablesCreate,
    undo: sheetGroupTablesUndo,
    redo: sheetGroupTablesRedo,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };



  const sheetGroupPivotBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupPivotBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupPivotBaseCount, { timeout: 5000 });


  const sheetGroupPivotSheets = await page.evaluate(() =>
    [...document.querySelectorAll(".sheet-tab")]
      .slice(-2)
      .map((button) => ({ name: button.textContent, index: button.dataset.sheetIndex })),
  );


  await page.evaluate(([first, second]) => {
    window.__webSheetDevSetActiveSheet?.(first.name);
    [["区分", "金額"], ["A", 1], ["A", 2], ["B", 3]].forEach((rowValues, rowIndex) => {
      rowValues.forEach((value, colIndex) => window.__webSheetDevSetCell?.(rowIndex + 1, colIndex + 1, value));
    });
    window.__webSheetDevSetActiveSheet?.(second.name);
    [["区分", "金額"], ["A", 10], ["B", 20], ["B", 30]].forEach((rowValues, rowIndex) => {
      rowValues.forEach((value, colIndex) => window.__webSheetDevSetCell?.(rowIndex + 1, colIndex + 1, value));
    });
  }, sheetGroupPivotSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupPivotSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupPivotSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 1, 4, 2));


  await page.click('[data-ribbon-tab="insert"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "挿入", null, { timeout: 5000 });


  await page.click('[data-ribbon-command="insert-pivot-table"]');


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 4, sheetGroupPivotBaseCount, { timeout: 5000 });


  const sheetGroupPivotCreate = await page.evaluate((baseCount) => {
    const tabs = [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent);
    const pivotNames = tabs.slice(baseCount + 2);
    const pivotData = pivotNames.map((name) => ({
      name,
      tables: window.__webSheetDevSheetModel?.(name)?.tables || [],
      rows: [2, 3].map((row) => [
        window.__webSheetDevCellModel?.(name, row, 1)?.raw ?? "",
        window.__webSheetDevCellModel?.(name, row, 2)?.raw ?? "",
        window.__webSheetDevCellModel?.(name, row, 3)?.raw ?? "",
      ]),
    }));
    return {
      tabs,
      pivotNames,
      pivotData,
      activeSheetName: window.__webSheetDevWorkbookState?.().activeSheetName,
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupPivotBaseCount);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupPivotBaseCount, { timeout: 5000 });


  const sheetGroupPivotUndo = await page.evaluate((baseCount) => ({
    tabCount: document.querySelectorAll(".sheet-tab").length,
    expectedCount: baseCount + 2,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupPivotBaseCount);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 4, sheetGroupPivotBaseCount, { timeout: 5000 });


  const sheetGroupPivotRedo = await page.evaluate((baseCount) => {
    const tabs = [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent);
    const pivotNames = tabs.slice(baseCount + 2);
    return {
      pivotNames,
      firstRows: [2, 3].map((row) => [
        window.__webSheetDevCellModel?.(pivotNames[0], row, 1)?.raw ?? "",
        window.__webSheetDevCellModel?.(pivotNames[0], row, 3)?.raw ?? "",
      ]),
      secondRows: [2, 3].map((row) => [
        window.__webSheetDevCellModel?.(pivotNames[1], row, 1)?.raw ?? "",
        window.__webSheetDevCellModel?.(pivotNames[1], row, 3)?.raw ?? "",
      ]),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupPivotBaseCount);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupPivotBaseCount, { timeout: 5000 });


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupPivotSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupPivotSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupPivotSheets[0].index}"]`, { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupPivotBaseCount, { timeout: 5000 });


  const sheetGroupPivot = {
    create: sheetGroupPivotCreate,
    undo: sheetGroupPivotUndo,
    redo: sheetGroupPivotRedo,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };



  const sheetGroupObjectsBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, sheetGroupObjectsBaseCount, { timeout: 5000 });


  await page.click("#addSheetButton");


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, sheetGroupObjectsBaseCount, { timeout: 5000 });


  const sheetGroupObjectsSheets = await page.evaluate(() =>
    [...document.querySelectorAll(".sheet-tab")]
      .slice(-2)
      .map((button) => ({ name: button.textContent, index: button.dataset.sheetIndex })),
  );


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupObjectsSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupObjectsSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(2, 2, 2, 2));


  await page.click('[data-ribbon-tab="insert"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "挿入", null, { timeout: 5000 });


  await page.click('[data-ribbon-command="insert-wordart"]');


  await page.waitForFunction(([first, second]) =>
    (window.__webSheetDevSheetModel?.(first.name)?.images || []).some((image) => image.name === "ワードアート") &&
    (window.__webSheetDevSheetModel?.(second.name)?.images || []).some((image) => image.name === "ワードアート"),
    sheetGroupObjectsSheets,
    { timeout: 5000 },
  );


  const sheetGroupObjectsWordArt = await page.evaluate(([first, second]) => {
    const summary = (sheetName) => {
      const images = window.__webSheetDevSheetModel?.(sheetName)?.images || [];
      const wordArt = images.find((image) => image.name === "ワードアート") || null;
      return {
        imageCount: images.length,
        hasWordArt: Boolean(wordArt),
        shapeType: wordArt?.shape?.type || "",
        text: wordArt?.shape?.text || "",
      };
    };
    return {
      first: summary(first.name),
      second: summary(second.name),
      selectedObjects: document.querySelectorAll(".sheet-image.is-selected, .sheet-image-link.is-selected").length,
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupObjectsSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !(window.__webSheetDevSheetModel?.(first.name)?.images || []).some((image) => image.name === "ワードアート") &&
    !(window.__webSheetDevSheetModel?.(second.name)?.images || []).some((image) => image.name === "ワードアート"),
    sheetGroupObjectsSheets,
    { timeout: 5000 },
  );


  const sheetGroupObjectsWordArtUndo = await page.evaluate(([first, second]) => ({
    firstCount: (window.__webSheetDevSheetModel?.(first.name)?.images || []).length,
    secondCount: (window.__webSheetDevSheetModel?.(second.name)?.images || []).length,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupObjectsSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    (window.__webSheetDevSheetModel?.(first.name)?.images || []).some((image) => image.name === "ワードアート") &&
    (window.__webSheetDevSheetModel?.(second.name)?.images || []).some((image) => image.name === "ワードアート"),
    sheetGroupObjectsSheets,
    { timeout: 5000 },
  );


  const sheetGroupObjectsWordArtRedo = await page.evaluate(([first, second]) => ({
    firstWordArtCount: (window.__webSheetDevSheetModel?.(first.name)?.images || []).filter((image) => image.name === "ワードアート").length,
    secondWordArtCount: (window.__webSheetDevSheetModel?.(second.name)?.images || []).filter((image) => image.name === "ワードアート").length,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupObjectsSheets);


  await page.click('[data-ribbon-tab="insert"]');


  await page.click("[data-ribbon-shape-menu]");


  await page.waitForSelector("#ribbonShapeMenu", { timeout: 5000 });


  await page.click('#ribbonShapeMenu [data-ribbon-shape-choice="rect"]');


  await page.waitForFunction(() => !document.querySelector("#ribbonShapeMenu") && document.body.classList.contains("is-drawing-shape"), null, { timeout: 5000 });


  const sheetGroupObjectsHostBox = await page.locator("#sheetHost").boundingBox();


  if (!sheetGroupObjectsHostBox) throw new Error("Sheet host is not visible for grouped object draw test.");


  await page.mouse.move(sheetGroupObjectsHostBox.x + 210, sheetGroupObjectsHostBox.y + 130);


  await page.mouse.down();


  await page.waitForSelector(".sheet-shape-draft", { timeout: 5000 });


  await page.mouse.move(sheetGroupObjectsHostBox.x + 360, sheetGroupObjectsHostBox.y + 220, { steps: 8 });


  await page.mouse.up();


  await page.waitForFunction(([first, second]) =>
    (window.__webSheetDevSheetModel?.(first.name)?.images || []).some((image) => image.shape?.type === "rect") &&
    (window.__webSheetDevSheetModel?.(second.name)?.images || []).some((image) => image.shape?.type === "rect"),
    sheetGroupObjectsSheets,
    { timeout: 5000 },
  );


  const sheetGroupObjectsShape = await page.evaluate(([first, second]) => {
    const summary = (sheetName) => {
      const images = window.__webSheetDevSheetModel?.(sheetName)?.images || [];
      const rect = images.find((image) => image.shape?.type === "rect") || null;
      return {
        imageCount: images.length,
        rectCount: images.filter((image) => image.shape?.type === "rect").length,
        width: Math.round(rect?.ext?.width || 0),
        height: Math.round(rect?.ext?.height || 0),
      };
    };
    return {
      first: summary(first.name),
      second: summary(second.name),
      selectedObjects: document.querySelectorAll(".sheet-image.is-selected, .sheet-image-link.is-selected").length,
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupObjectsSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !(window.__webSheetDevSheetModel?.(first.name)?.images || []).some((image) => image.shape?.type === "rect") &&
    !(window.__webSheetDevSheetModel?.(second.name)?.images || []).some((image) => image.shape?.type === "rect"),
    sheetGroupObjectsSheets,
    { timeout: 5000 },
  );


  const sheetGroupObjectsShapeUndo = await page.evaluate(([first, second]) => ({
    firstWordArtCount: (window.__webSheetDevSheetModel?.(first.name)?.images || []).filter((image) => image.name === "ワードアート").length,
    firstRectCount: (window.__webSheetDevSheetModel?.(first.name)?.images || []).filter((image) => image.shape?.type === "rect").length,
    secondWordArtCount: (window.__webSheetDevSheetModel?.(second.name)?.images || []).filter((image) => image.name === "ワードアート").length,
    secondRectCount: (window.__webSheetDevSheetModel?.(second.name)?.images || []).filter((image) => image.shape?.type === "rect").length,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupObjectsSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    (window.__webSheetDevSheetModel?.(first.name)?.images || []).some((image) => image.shape?.type === "rect") &&
    (window.__webSheetDevSheetModel?.(second.name)?.images || []).some((image) => image.shape?.type === "rect"),
    sheetGroupObjectsSheets,
    { timeout: 5000 },
  );


  const sheetGroupObjectsShapeRedo = await page.evaluate(([first, second]) => ({
    firstRectCount: (window.__webSheetDevSheetModel?.(first.name)?.images || []).filter((image) => image.shape?.type === "rect").length,
    secondRectCount: (window.__webSheetDevSheetModel?.(second.name)?.images || []).filter((image) => image.shape?.type === "rect").length,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupObjectsSheets);


  await page.locator('.sheet-image[alt="四角形"]').first().click();


  await page.waitForSelector('[data-ribbon-panel="shape-format"].is-active', { timeout: 5000 });


  await page.click('[data-shape-style-preset="softGreen"]');


  await page.waitForFunction(([first, second]) => {
    const rect = (name) => (window.__webSheetDevSheetModel?.(name)?.images || []).find((image) => image.shape?.type === "rect");
    return rect(first.name)?.shape?.fillColor === "#e2f0d9" && rect(second.name)?.shape?.fillColor === "#e2f0d9";
  }, sheetGroupObjectsSheets, { timeout: 5000 });


  const sheetGroupObjectsStyle = await page.evaluate(([first, second]) => {
    const summary = (sheetName) => {
      const rect = (window.__webSheetDevSheetModel?.(sheetName)?.images || []).find((image) => image.shape?.type === "rect") || null;
      return {
        fillColor: rect?.shape?.fillColor || "",
        outlineColor: rect?.shape?.outlineColor || "",
      };
    };
    return {
      first: summary(first.name),
      second: summary(second.name),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupObjectsSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) => {
    const rect = (name) => (window.__webSheetDevSheetModel?.(name)?.images || []).find((image) => image.shape?.type === "rect");
    return rect(first.name)?.shape?.fillColor !== "#e2f0d9" && rect(second.name)?.shape?.fillColor !== "#e2f0d9";
  }, sheetGroupObjectsSheets, { timeout: 5000 });


  const sheetGroupObjectsStyleUndo = await page.evaluate(([first, second]) => {
    const rect = (name) => (window.__webSheetDevSheetModel?.(name)?.images || []).find((image) => image.shape?.type === "rect");
    return {
      firstFill: rect(first.name)?.shape?.fillColor || "",
      secondFill: rect(second.name)?.shape?.fillColor || "",
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupObjectsSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) => {
    const rect = (name) => (window.__webSheetDevSheetModel?.(name)?.images || []).find((image) => image.shape?.type === "rect");
    return rect(first.name)?.shape?.fillColor === "#e2f0d9" && rect(second.name)?.shape?.fillColor === "#e2f0d9";
  }, sheetGroupObjectsSheets, { timeout: 5000 });


  await page.locator('.sheet-image[alt="四角形"]').first().click();


  await page.waitForSelector('[data-ribbon-panel="shape-format"].is-active', { timeout: 5000 });


  await page.locator('[data-shape-size="width"]').first().evaluate((input) => {
    input.value = "220";
    input.dispatchEvent(new Event("change", { bubbles: true }));
  });


  await page.waitForFunction(([first, second]) => {
    const rect = (name) => (window.__webSheetDevSheetModel?.(name)?.images || []).find((image) => image.shape?.type === "rect");
    return Math.round(rect(first.name)?.ext?.width || 0) === 220 && Math.round(rect(second.name)?.ext?.width || 0) === 220;
  }, sheetGroupObjectsSheets, { timeout: 5000 });


  const sheetGroupObjectsSize = await page.evaluate(([first, second]) => {
    const rect = (name) => (window.__webSheetDevSheetModel?.(name)?.images || []).find((image) => image.shape?.type === "rect");
    return {
      firstWidth: Math.round(rect(first.name)?.ext?.width || 0),
      secondWidth: Math.round(rect(second.name)?.ext?.width || 0),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupObjectsSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) => {
    const rect = (name) => (window.__webSheetDevSheetModel?.(name)?.images || []).find((image) => image.shape?.type === "rect");
    return Math.round(rect(first.name)?.ext?.width || 0) !== 220 && Math.round(rect(second.name)?.ext?.width || 0) !== 220;
  }, sheetGroupObjectsSheets, { timeout: 5000 });


  const sheetGroupObjectsSizeUndo = await page.evaluate(([first, second]) => {
    const rect = (name) => (window.__webSheetDevSheetModel?.(name)?.images || []).find((image) => image.shape?.type === "rect");
    return {
      firstWidth: Math.round(rect(first.name)?.ext?.width || 0),
      secondWidth: Math.round(rect(second.name)?.ext?.width || 0),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupObjectsSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) => {
    const rect = (name) => (window.__webSheetDevSheetModel?.(name)?.images || []).find((image) => image.shape?.type === "rect");
    return Math.round(rect(first.name)?.ext?.width || 0) === 220 && Math.round(rect(second.name)?.ext?.width || 0) === 220;
  }, sheetGroupObjectsSheets, { timeout: 5000 });


  const sheetGroupObjectsSizeRedo = await page.evaluate(([first, second]) => {
    const rect = (name) => (window.__webSheetDevSheetModel?.(name)?.images || []).find((image) => image.shape?.type === "rect");
    return {
      firstWidth: Math.round(rect(first.name)?.ext?.width || 0),
      secondWidth: Math.round(rect(second.name)?.ext?.width || 0),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupObjectsSheets);


  await page.locator('.sheet-image[alt="ワードアート"]').first().click();


  await page.locator('.sheet-image[alt="四角形"]').first().click({ modifiers: ["Meta"] });


  await page.waitForFunction(() => document.querySelectorAll(".sheet-image.is-selected, .sheet-image-link.is-selected").length >= 2, null, { timeout: 5000 });


  await page.click('[data-ribbon-command="shape-group"]');


  await page.waitForFunction(([first, second]) => {
    const grouped = (name) => (window.__webSheetDevSheetModel?.(name)?.images || []).filter((image) => image.groupId);
    return grouped(first.name).length === 2 && grouped(second.name).length === 2;
  }, sheetGroupObjectsSheets, { timeout: 5000 });


  const sheetGroupObjectsGroup = await page.evaluate(([first, second]) => {
    const summary = (sheetName) => {
      const grouped = (window.__webSheetDevSheetModel?.(sheetName)?.images || []).filter((image) => image.groupId);
      return {
        groupedCount: grouped.length,
        groupIdCount: new Set(grouped.map((image) => image.groupId)).size,
      };
    };
    return {
      first: summary(first.name),
      second: summary(second.name),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetGroupObjectsSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    !(window.__webSheetDevSheetModel?.(first.name)?.images || []).some((image) => image.groupId) &&
    !(window.__webSheetDevSheetModel?.(second.name)?.images || []).some((image) => image.groupId),
    sheetGroupObjectsSheets,
    { timeout: 5000 },
  );


  const sheetGroupObjectsGroupUndo = await page.evaluate(([first, second]) => ({
    firstGrouped: (window.__webSheetDevSheetModel?.(first.name)?.images || []).filter((image) => image.groupId).length,
    secondGrouped: (window.__webSheetDevSheetModel?.(second.name)?.images || []).filter((image) => image.groupId).length,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupObjectsSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) => {
    const grouped = (name) => (window.__webSheetDevSheetModel?.(name)?.images || []).filter((image) => image.groupId);
    return grouped(first.name).length === 2 && grouped(second.name).length === 2;
  }, sheetGroupObjectsSheets, { timeout: 5000 });


  const sheetGroupObjectsGroupRedo = await page.evaluate(([first, second]) => ({
    firstGrouped: (window.__webSheetDevSheetModel?.(first.name)?.images || []).filter((image) => image.groupId).length,
    secondGrouped: (window.__webSheetDevSheetModel?.(second.name)?.images || []).filter((image) => image.groupId).length,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupObjectsSheets);


  await page.locator('.sheet-image[alt="四角形"]').first().click();


  await page.waitForFunction(() => !document.querySelector('[data-ribbon-tab="shape-format"]')?.classList.contains("hidden"), null, { timeout: 5000 });


  await page.click('[data-ribbon-tab="shape-format"]');


  await page.waitForSelector('[data-ribbon-panel="shape-format"].is-active', { timeout: 5000 });


  await page.click('[data-ribbon-command="shape-ungroup"]');


  await page.waitForFunction(([first, second]) =>
    !(window.__webSheetDevSheetModel?.(first.name)?.images || []).some((image) => image.groupId) &&
    !(window.__webSheetDevSheetModel?.(second.name)?.images || []).some((image) => image.groupId),
    sheetGroupObjectsSheets,
    { timeout: 5000 },
  );


  const sheetGroupObjectsUngroup = await page.evaluate(([first, second]) => ({
    firstGrouped: (window.__webSheetDevSheetModel?.(first.name)?.images || []).filter((image) => image.groupId).length,
    secondGrouped: (window.__webSheetDevSheetModel?.(second.name)?.images || []).filter((image) => image.groupId).length,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupObjectsSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) => {
    const grouped = (name) => (window.__webSheetDevSheetModel?.(name)?.images || []).filter((image) => image.groupId);
    return grouped(first.name).length === 2 && grouped(second.name).length === 2;
  }, sheetGroupObjectsSheets, { timeout: 5000 });


  const sheetGroupObjectsUngroupUndo = await page.evaluate(([first, second]) => ({
    firstGrouped: (window.__webSheetDevSheetModel?.(first.name)?.images || []).filter((image) => image.groupId).length,
    secondGrouped: (window.__webSheetDevSheetModel?.(second.name)?.images || []).filter((image) => image.groupId).length,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupObjectsSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    !(window.__webSheetDevSheetModel?.(first.name)?.images || []).some((image) => image.groupId) &&
    !(window.__webSheetDevSheetModel?.(second.name)?.images || []).some((image) => image.groupId),
    sheetGroupObjectsSheets,
    { timeout: 5000 },
  );


  await page.locator('.sheet-image[alt="四角形"]').first().click();


  await page.keyboard.press("Delete");


  await page.waitForFunction(([first, second]) =>
    !(window.__webSheetDevSheetModel?.(first.name)?.images || []).some((image) => image.shape?.type === "rect") &&
    !(window.__webSheetDevSheetModel?.(second.name)?.images || []).some((image) => image.shape?.type === "rect"),
    sheetGroupObjectsSheets,
    { timeout: 5000 },
  );


  const sheetGroupObjectsDelete = await page.evaluate(([first, second]) => ({
    firstRectCount: (window.__webSheetDevSheetModel?.(first.name)?.images || []).filter((image) => image.shape?.type === "rect").length,
    secondRectCount: (window.__webSheetDevSheetModel?.(second.name)?.images || []).filter((image) => image.shape?.type === "rect").length,
    firstWordArtCount: (window.__webSheetDevSheetModel?.(first.name)?.images || []).filter((image) => image.name === "ワードアート").length,
    secondWordArtCount: (window.__webSheetDevSheetModel?.(second.name)?.images || []).filter((image) => image.name === "ワードアート").length,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupObjectsSheets);


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(([first, second]) =>
    (window.__webSheetDevSheetModel?.(first.name)?.images || []).some((image) => image.shape?.type === "rect") &&
    (window.__webSheetDevSheetModel?.(second.name)?.images || []).some((image) => image.shape?.type === "rect"),
    sheetGroupObjectsSheets,
    { timeout: 5000 },
  );


  const sheetGroupObjectsDeleteUndo = await page.evaluate(([first, second]) => ({
    firstRectCount: (window.__webSheetDevSheetModel?.(first.name)?.images || []).filter((image) => image.shape?.type === "rect").length,
    secondRectCount: (window.__webSheetDevSheetModel?.(second.name)?.images || []).filter((image) => image.shape?.type === "rect").length,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupObjectsSheets);


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(([first, second]) =>
    !(window.__webSheetDevSheetModel?.(first.name)?.images || []).some((image) => image.shape?.type === "rect") &&
    !(window.__webSheetDevSheetModel?.(second.name)?.images || []).some((image) => image.shape?.type === "rect"),
    sheetGroupObjectsSheets,
    { timeout: 5000 },
  );


  const sheetGroupObjectsDeleteRedo = await page.evaluate(([first, second]) => ({
    firstRectCount: (window.__webSheetDevSheetModel?.(first.name)?.images || []).filter((image) => image.shape?.type === "rect").length,
    secondRectCount: (window.__webSheetDevSheetModel?.(second.name)?.images || []).filter((image) => image.shape?.type === "rect").length,
    status: document.querySelector("#statusText")?.textContent || "",
  }), sheetGroupObjectsSheets);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupObjectsSheets[0].index}"]`);


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupObjectsSheets[1].index}"]`, { modifiers: ["Meta"] });


  await page.click(`.sheet-tab[data-sheet-index="${sheetGroupObjectsSheets[0].index}"]`, { button: "right" });


  await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, sheetGroupObjectsBaseCount, { timeout: 5000 });


  const sheetGroupObjects = {
    wordArt: sheetGroupObjectsWordArt,
    wordArtUndo: sheetGroupObjectsWordArtUndo,
    wordArtRedo: sheetGroupObjectsWordArtRedo,
    shape: sheetGroupObjectsShape,
    shapeUndo: sheetGroupObjectsShapeUndo,
    shapeRedo: sheetGroupObjectsShapeRedo,
    style: sheetGroupObjectsStyle,
    styleUndo: sheetGroupObjectsStyleUndo,
    size: sheetGroupObjectsSize,
    sizeUndo: sheetGroupObjectsSizeUndo,
    sizeRedo: sheetGroupObjectsSizeRedo,
    group: sheetGroupObjectsGroup,
    groupUndo: sheetGroupObjectsGroupUndo,
    groupRedo: sheetGroupObjectsGroupRedo,
    ungroup: sheetGroupObjectsUngroup,
    ungroupUndo: sheetGroupObjectsUngroupUndo,
    delete: sheetGroupObjectsDelete,
    deleteUndo: sheetGroupObjectsDeleteUndo,
    deleteRedo: sheetGroupObjectsDeleteRedo,
    cleanup: await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };



  await page.click('[data-ribbon-tab="insert"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "挿入", null, { timeout: 10000 });


  const insertRibbon = await page.evaluate(() => {
    const panel = document.querySelector('[data-ribbon-panel="insert"]');
    return {
      activeRibbonTab: document.querySelector(".ribbon-tab.is-active")?.textContent,
      panelVisible: Boolean(panel?.classList.contains("is-active")),
      groupLabels: [...(panel?.querySelectorAll(".ribbon-group-label") || [])].map((label) => label.textContent),
      enabledCommands: [...(panel?.querySelectorAll("[data-ribbon-command]:not(:disabled)") || [])].map((button) => button.dataset.ribbonCommand),
      disabledCommandCount: panel?.querySelectorAll("[data-ribbon-command]:disabled").length || 0,
      largeButtonCount: panel?.querySelectorAll(".ribbon-large-tool").length || 0,
      compactButtonCount: panel?.querySelectorAll(".ribbon-compact-tool").length || 0,
      chartIconCount: panel?.querySelectorAll(".ribbon-insert-charts-group .ribbon-icon-tool").length || 0,
      missingIcons: [...(panel?.querySelectorAll("i[data-lucide]") || [])]
        .filter((icon) => !icon.querySelector("svg"))
        .map((icon) => icon.dataset.lucide),
      panelHeight: Math.round(document.querySelector(".ribbon-panels")?.getBoundingClientRect().height || 0),
    };
  });



  await page.click('[data-ribbon-tab="formula"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "数式", null, { timeout: 10000 });


  const formulaRibbon = await page.evaluate(() => {
    const panel = document.querySelector('[data-ribbon-panel="formula"]');
    return {
      activeRibbonTab: document.querySelector(".ribbon-tab.is-active")?.textContent,
      panelVisible: Boolean(panel?.classList.contains("is-active")),
      groupLabels: [...(panel?.querySelectorAll(".ribbon-group-label") || [])].map((label) => label.textContent.trim()),
      enabledCommands: [...(panel?.querySelectorAll("[data-ribbon-command]:not(:disabled)") || [])].map((button) => button.dataset.ribbonCommand),
      compactButtonCount: panel?.querySelectorAll(".ribbon-compact-tool").length || 0,
      missingIcons: [...(panel?.querySelectorAll("i[data-lucide]") || [])]
        .filter((icon) => !icon.querySelector("svg"))
        .map((icon) => icon.dataset.lucide),
      panelHeight: Math.round(document.querySelector(".ribbon-panels")?.getBoundingClientRect().height || 0),
    };
  });


  await page.evaluate(() => {
    window.__webSheetDevSetActiveSheet?.("CHT");
    [
      [1, 8, "FormulaSeed"],
      [2, 8, 10],
      [3, 8, 20],
      [1, 9, "Metric"],
      [2, 9, 7],
      [3, 9, 11],
    ].forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
    window.__webSheetDevSelectRange?.(1, 8, 3, 9);
  });


  await page.click('[data-ribbon-command="formula-define-name"] span');


  await page.waitForSelector("#nameDialogOverlay", { timeout: 5000 });


  const formulaDefineNameMain = await page.evaluate(() => ({
    title: document.querySelector("#nameDialogTitle")?.textContent || "",
    refersTo: document.querySelector('[data-name-field="refersTo"]')?.value || "",
  }));


  await page.click('#nameDialogOverlay [data-name-dialog-action="cancel"]');


  await page.waitForSelector("#nameDialogOverlay", { state: "detached", timeout: 5000 });


  await page.click('[data-ribbon-command="formula-define-name"] .ribbon-command-caret');


  await page.waitForSelector("#homeCommandMenu.is-formula-define-name", { timeout: 5000 });


  const formulaDefineNameMenu = await page.evaluate(() => ({
    label: document.querySelector("#homeCommandMenu")?.getAttribute("aria-label") || "",
    accessKeys: [...document.querySelectorAll("#homeCommandMenu button")].map((button) => button.dataset.accessKey || ""),
    seedBefore: window.__webSheetDevCellModel?.("CHT", 1, 8)?.raw || "",
    actions: [...document.querySelectorAll("#homeCommandMenu [data-formula-name-menu-action]")].map((button) => ({
      action: button.dataset.formulaNameMenuAction,
      text: button.textContent.trim(),
    })),
  }));


  await page.keyboard.press("C");


  await page.waitForSelector('[data-data-dialog="create-names"]', { timeout: 5000 });


  const formulaDefineNameCreate = await page.evaluate(() => ({
    title: document.querySelector("#createNamesTitle")?.textContent || "",
    seedAfterShortcut: window.__webSheetDevCellModel?.("CHT", 1, 8)?.raw || "",
    menuClosed: !document.querySelector("#homeCommandMenu"),
  }));


  await page.click('[data-data-dialog="create-names"] [data-data-dialog-action="cancel"]');


  await page.waitForSelector('[data-data-dialog="create-names"]', { state: "detached", timeout: 5000 });


  await page.click('[data-ribbon-command="formula-create-from-selection"]');


  await page.waitForSelector('[data-data-dialog="create-names"]', { timeout: 5000 });


  const createNamesDialog = await page.evaluate(() => ({
    title: document.querySelector("#createNamesTitle")?.textContent || "",
    checks: [...document.querySelectorAll("[data-create-name-source]")].map((input) => ({ source: input.dataset.createNameSource, checked: input.checked })),
  }));


  await page.click('[data-data-dialog="create-names"] [data-data-dialog-action="create-names-ok"]');


  await page.waitForSelector('[data-data-dialog="create-names"]', { state: "detached", timeout: 5000 });


  const formulaNames = await page.evaluate(() => ({
    names: window.__webSheetDevDefinedNames?.().map((definition) => ({ name: definition.name, refersTo: definition.refersTo })) || [],
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.evaluate(() => {
    window.__webSheetDevAddDefinedName?.({ name: "LocalMetric", refersTo: "=CHT!$H$2", scope: "sheet", sheetIndex: 0, comment: "ローカル名" });
    window.__webSheetDevAddDefinedName?.({ name: "BrokenName", refersTo: "=#REF!", comment: "壊れた名前" });
  });


  await page.click('[data-ribbon-command="formula-name-manager"]');


  await page.waitForSelector('[data-data-dialog="name-manager"]', { timeout: 5000 });


  const formulaNameManagerInitial = await page.evaluate(() => ({
    title: document.querySelector("#nameManagerTitle")?.textContent || "",
    headers: [...document.querySelectorAll(".name-manager-table th")].map((item) => item.textContent.replace(/\s+/g, " ").trim()),
    filterOptions: [...document.querySelectorAll("[data-formula-name-filter] option")].map((option) => option.textContent.trim()),
    rows: [...document.querySelectorAll("[data-formula-name-row]")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
    summary: document.querySelector("[data-formula-name-summary]")?.textContent || "",
    hasEdit: Boolean(document.querySelector('[data-data-dialog-action="name-manager-edit"]')),
    currentRefersTo: document.querySelector("[data-formula-name-current-refers-to]")?.value || "",
  }));


  await page.selectOption("[data-formula-name-filter]", "current-sheet");


  const formulaNameManagerCurrentSheet = await page.evaluate(() => ({
    rows: [...document.querySelectorAll("[data-formula-name-row]")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
    summary: document.querySelector("[data-formula-name-summary]")?.textContent || "",
  }));


  await page.selectOption("[data-formula-name-filter]", "errors");


  const formulaNameManagerErrors = await page.evaluate(() => ({
    rows: [...document.querySelectorAll("[data-formula-name-row]")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
    summary: document.querySelector("[data-formula-name-summary]")?.textContent || "",
  }));


  await page.selectOption("[data-formula-name-filter]", "all");


  await page.evaluate(() => {
    [...document.querySelectorAll("[data-formula-name-row] [data-formula-name-row-index]")].find((button) => button.textContent.trim() === "FormulaSeed")?.click();
  });


  await page.click('[data-data-dialog-action="name-manager-edit"]');


  await page.waitForSelector("#nameDialogOverlay", { timeout: 5000 });


  await page.fill('#nameDialogOverlay [data-name-field="comment"]', "Name Manager から編集");


  await page.click('#nameDialogOverlay [data-name-dialog-action="ok"]');


  await page.waitForSelector('[data-data-dialog="name-manager"]', { timeout: 5000 });


  const formulaNameManagerEdit = await page.evaluate(() => {
    const definition = window.__webSheetDevDefinedNames?.().find((item) => item.name === "FormulaSeed");
    return {
      managerOpen: Boolean(document.querySelector('[data-data-dialog="name-manager"]')),
      comment: definition?.comment || "",
      status: document.querySelector("#statusText")?.textContent || "",
    };
  });


  await page.click('[data-data-dialog-action="name-manager-new"]');


  await page.waitForSelector("#nameDialogOverlay", { timeout: 5000 });


  await page.fill('#nameDialogOverlay [data-name-field="name"]', "ManagerNewName");


  await page.fill('#nameDialogOverlay [data-name-field="comment"]', "Name Manager から新規作成");


  await page.fill('#nameDialogOverlay [data-name-field="refersTo"]', "=CHT!$H$2");


  await page.click('#nameDialogOverlay [data-name-dialog-action="ok"]');


  await page.waitForSelector('[data-data-dialog="name-manager"]', { timeout: 5000 });


  const formulaNameManagerNew = await page.evaluate(() => ({
    managerOpen: Boolean(document.querySelector('[data-data-dialog="name-manager"]')),
    exists: window.__webSheetDevDefinedNames?.().some((item) => item.name === "ManagerNewName") || false,
    rows: [...document.querySelectorAll("[data-formula-name-row]")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.selectOption("[data-formula-name-filter]", "errors");


  await page.evaluate(() => {
    [...document.querySelectorAll("[data-formula-name-row] [data-formula-name-row-index]")].find((button) => button.textContent.trim() === "BrokenName")?.click();
  });


  await page.click('[data-data-dialog-action="name-manager-delete"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  await page.locator('[data-message-box-result="delete"]').click();


  await page.waitForFunction(() => !(window.__webSheetDevDefinedNames?.() || []).some((item) => item.name === "BrokenName"), null, { timeout: 5000 });


  const formulaNameManagerDelete = await page.evaluate(() => ({
    managerOpen: Boolean(document.querySelector('[data-data-dialog="name-manager"]')),
    brokenExists: window.__webSheetDevDefinedNames?.().some((item) => item.name === "BrokenName") || false,
    summary: document.querySelector("[data-formula-name-summary]")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('[data-data-dialog="name-manager"] [data-data-dialog-action="cancel"]');


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => (window.__webSheetDevDefinedNames?.() || []).some((item) => item.name === "BrokenName"), null, { timeout: 5000 });


  const formulaNameManagerUndo = await page.evaluate(() => ({
    brokenExists: window.__webSheetDevDefinedNames?.().some((item) => item.name === "BrokenName") || false,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => !(window.__webSheetDevDefinedNames?.() || []).some((item) => item.name === "BrokenName"), null, { timeout: 5000 });


  const formulaNameManagerRedo = await page.evaluate(() => ({
    brokenExists: window.__webSheetDevDefinedNames?.().some((item) => item.name === "BrokenName") || false,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  const formulaNameManager = {
    initial: formulaNameManagerInitial,
    currentSheet: formulaNameManagerCurrentSheet,
    errors: formulaNameManagerErrors,
    edit: formulaNameManagerEdit,
    newName: formulaNameManagerNew,
    delete: formulaNameManagerDelete,
    undo: formulaNameManagerUndo,
    redo: formulaNameManagerRedo,
  };


  await page.evaluate(() => {
    window.__webSheetDevAddDefinedName?.({ name: "UseMenuBroken", refersTo: "=#REF!", comment: "数式で使用メニューのエラー表示" });
  });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(9, 8, 9, 8));


  await page.click('[data-ribbon-command="formula-use-in-formula"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-use-name", { timeout: 5000 });


  const formulaUseNameMenu = await page.evaluate(() => ({
    label: document.querySelector("#homeCommandMenu")?.getAttribute("aria-label") || "",
    headers: [...document.querySelectorAll("#homeCommandMenu .formula-name-menu-header span")].map((item) => item.textContent.trim()),
    names: [...document.querySelectorAll("#homeCommandMenu [data-formula-name-menu-index] strong")].map((item) => item.textContent.trim()),
    rows: [...document.querySelectorAll("#homeCommandMenu [data-formula-name-menu-index]")].map((button) => ({
      name: button.querySelector(".formula-name-menu-name strong")?.textContent.trim() || "",
      value: button.querySelector(".formula-name-menu-value")?.textContent.trim() || "",
      scope: button.querySelector(".formula-name-menu-scope")?.textContent.trim() || "",
      reference: button.querySelector(".formula-name-menu-reference")?.textContent.trim() || "",
      hasError: button.dataset.formulaNameMenuError === "true",
    })),
    hasPasteName: Boolean(document.querySelector("#homeCommandMenu [data-formula-name-menu-action='paste']")),
    hasManage: Boolean(document.querySelector("#homeCommandMenu [data-formula-name-menu-action='manage']")),
  }));


  await page.evaluate(() => {
    document.querySelector("#homeCommandMenu [data-formula-name-menu-action='manage']")?.click();
  });


  await page.waitForSelector('[data-data-dialog="name-manager"]', { timeout: 5000 });


  const formulaUseNameManage = await page.evaluate(() => ({
    title: document.querySelector("#nameManagerTitle")?.textContent || "",
    hasUseMenuBroken: [...document.querySelectorAll("[data-formula-name-table-body] tr")].some((row) => row.textContent.includes("UseMenuBroken")),
  }));


  await page.locator('[data-data-dialog="name-manager"] [data-data-dialog-action="cancel"]').click();


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  await page.click('[data-ribbon-command="formula-use-in-formula"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-use-name", { timeout: 5000 });


  await page.evaluate(() => {
    [...document.querySelectorAll("#homeCommandMenu [data-formula-name-menu-index]")].find((button) =>
      button.textContent.includes("FormulaSeed"),
    )?.click();
  });


  await page.waitForFunction(() => window.__webSheetDevCellModel?.("CHT", 9, 8)?.formula === "=FormulaSeed", null, { timeout: 5000 });


  const formulaUseNameInserted = await page.evaluate(() => ({
    formula: window.__webSheetDevCellModel?.("CHT", 9, 8)?.formula || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(10, 8, 10, 8));


  await page.click('[data-ribbon-command="formula-paste-name"]');


  await page.waitForSelector('[data-data-dialog="paste-name"]', { timeout: 5000 });


  const formulaPasteNameDialog = await page.evaluate(() => {
    const list = document.querySelector("[data-formula-paste-name-list]");
    const option = [...(list?.options || [])].find((item) => item.textContent.trim().split(/\s+/)[0] === "Metric");
    if (option) list.value = option.value;
    list?.dispatchEvent(new Event("change", { bubbles: true }));
    return {
      title: document.querySelector("#pasteNameTitle")?.textContent || "",
      hasMetric: Boolean(option),
      headers: [...document.querySelectorAll(".paste-name-list-header span")].map((item) => item.textContent.trim()),
      reference: document.querySelector("[data-formula-paste-name-reference]")?.textContent || "",
    };
  });


  await page.click('[data-data-dialog="paste-name"] [data-data-dialog-action="paste-name-ok"]');


  await page.waitForFunction(() => window.__webSheetDevCellModel?.("CHT", 10, 8)?.formula === "=Metric", null, { timeout: 5000 });


  const formulaPasteNameInserted = await page.evaluate(() => ({
    formula: window.__webSheetDevCellModel?.("CHT", 10, 8)?.formula || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(210, 82, 210, 82));


  await page.click('[data-ribbon-command="formula-paste-name"]');


  await page.waitForSelector('[data-data-dialog="paste-name"]', { timeout: 5000 });


  const formulaPasteNameListDialog = await page.evaluate(() => ({
    title: document.querySelector("#pasteNameTitle")?.textContent || "",
    hasPasteList: Boolean(document.querySelector('[data-data-dialog-action="paste-name-list"]')),
    optionCount: document.querySelectorAll("[data-formula-paste-name-list] option").length,
  }));


  await page.click('[data-data-dialog="paste-name"] [data-data-dialog-action="paste-name-list"]');


  await page.waitForFunction(() => (window.__webSheetDevCellDisplay?.("CHT", 210, 82) || "") === "FormulaSeed", null, { timeout: 5000 });


  const formulaPasteNameList = await page.evaluate(() => {
    const names = [];
    for (let row = 210; row < 218; row += 1) {
      const name = window.__webSheetDevCellDisplay?.("CHT", row, 82) || "";
      const referenceDisplay = window.__webSheetDevCellDisplay?.("CHT", row, 83) || "";
      const referenceCell = window.__webSheetDevCellModel?.("CHT", row, 83) || {};
      if (name || referenceDisplay) names.push({ row, name, referenceDisplay, referenceFormula: referenceCell.formula || "", referenceRaw: referenceCell.raw || "" });
    }
    return {
      selection: document.querySelector("#selectedCellLabel")?.textContent || "",
      names,
      status: document.querySelector("#statusText")?.textContent || "",
    };
  });


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => !(window.__webSheetDevCellDisplay?.("CHT", 210, 82) || ""), null, { timeout: 5000 });


  const formulaPasteNameListUndo = await page.evaluate(() => ({
    firstName: window.__webSheetDevCellDisplay?.("CHT", 210, 82) || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => (window.__webSheetDevCellDisplay?.("CHT", 210, 82) || "") === "FormulaSeed", null, { timeout: 5000 });


  const formulaPasteNameListRedo = await page.evaluate(() => ({
    firstName: window.__webSheetDevCellDisplay?.("CHT", 210, 82) || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(6, 8, 6, 8));


  await page.click('[data-ribbon-command="formula-insert-function"]');


  await page.waitForSelector('[data-data-dialog="insert-function"]', { timeout: 5000 });


  const formulaFunctionDialog = await page.evaluate(() => ({
    title: document.querySelector("#insertFunctionTitle")?.textContent || "",
    functions: [...document.querySelectorAll("[data-formula-function-choice]")].map((button) => button.dataset.formulaFunctionChoice),
    visibleFunctions: [...document.querySelectorAll("[data-formula-function-choice]:not([hidden])")].map((button) => button.dataset.formulaFunctionChoice),
    categories: [...document.querySelectorAll("[data-formula-function-category] option")].map((option) => option.textContent),
    description: document.querySelector(".insert-function-description")?.textContent || "",
    searchLabel: document.querySelector(".insert-function-search-row .excel-dialog-caption")?.textContent || "",
    functionCount: document.querySelectorAll("[data-formula-function-choice]").length,
  }));


  await page.fill("[data-formula-function-search]", "sumproduct");


  await page.click('[data-data-dialog="insert-function"] [data-data-dialog-action="insert-function-search"]');


  await page.waitForFunction(() => [...document.querySelectorAll("[data-formula-function-choice]:not([hidden])")].some((button) => button.dataset.formulaFunctionChoice === "SUMPRODUCT"), null, { timeout: 5000 });


  const formulaFunctionSearch = await page.evaluate(() => ({
    visibleFunctions: [...document.querySelectorAll("[data-formula-function-choice]:not([hidden])")].map((button) => button.dataset.formulaFunctionChoice),
    active: document.activeElement?.dataset?.formulaFunctionChoice || "",
    description: document.querySelector(".insert-function-description")?.textContent || "",
  }));


  await page.fill("[data-formula-function-search]", "");


  await page.click('[data-formula-function-choice="IF"]');


  await page.click('[data-data-dialog="insert-function"] [data-data-dialog-action="insert-function-ok"]');


  await page.waitForSelector('[data-data-dialog="function-arguments"]', { timeout: 5000 });


  const formulaFunctionArgumentsInitial = await page.evaluate(() => ({
    title: document.querySelector("#functionArgumentsTitle")?.textContent || "",
    functionName: document.querySelector(".function-argument-heading strong")?.textContent || "",
    labels: [...document.querySelectorAll(".function-argument-name")].map((item) => item.textContent.trim().replace(/:$/u, "")),
    values: [...document.querySelectorAll("[data-formula-argument-input]")].map((input) => input.value),
    argumentResults: [...document.querySelectorAll(".function-argument-value")].map((item) => item.textContent.trim()),
    preview: document.querySelector("[data-function-argument-formula]")?.textContent || "",
    result: document.querySelector("[data-function-argument-result]")?.textContent || "",
    description: document.querySelector("[data-function-argument-description]")?.textContent || "",
    syntax: document.querySelector(".function-argument-syntax")?.textContent || "",
    referenceButtonCount: document.querySelectorAll("[data-data-dialog-action='function-argument-reference-toggle']").length,
  }));


  const formulaArgumentCellPoint = async (row, col) => {
    const point = await page.evaluate(({ row, col }) => {
      const rowHeader = document.querySelector(`.row-header[data-row="${row}"]`);
      const columnHeader = document.querySelector(`.column-header[data-col="${col}"]`);
      if (!rowHeader || !columnHeader) return null;
      const rowRect = rowHeader.getBoundingClientRect();
      const columnRect = columnHeader.getBoundingClientRect();
      return {
        x: columnRect.left + columnRect.width / 2,
        y: rowRect.top + rowRect.height / 2,
      };
    }, { row, col });
    if (!point) throw new Error(`Cell point ${row}:${col} is not visible`);
    return point;
  };


  const selectFormulaArgumentInput = async (index) => {
    const selector = `[data-formula-argument-input][data-formula-argument-index="${index}"]`;
    await page.focus(selector);
    await page.evaluate((selector) => {
      const input = document.querySelector(selector);
      input?.focus();
      input?.select?.();
      input?.dispatchEvent(new FocusEvent("focusin", { bubbles: true }));
    }, selector);
  };


  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    if (!host) return;
    host.scrollLeft = 0;
    host.scrollTop = 0;
    host.dispatchEvent(new Event("scroll", { bubbles: true }));
  });


  await page.waitForSelector('.column-header[data-col="1"]', { timeout: 5000 });


  await selectFormulaArgumentInput(0);


  const formulaArgumentA4 = await formulaArgumentCellPoint(4, 1);


  await page.mouse.click(formulaArgumentA4.x, formulaArgumentA4.y);


  await page.waitForFunction(() => document.querySelector('[data-formula-argument-input][data-formula-argument-index="0"]')?.value === "A4", null, { timeout: 5000 });


  await page.click('[data-data-dialog="function-arguments"] [data-data-dialog-action="function-argument-reference-toggle"][data-formula-argument-index="1"]');


  await page.waitForFunction(() => document.querySelector('[data-data-dialog="function-arguments"]')?.classList.contains("is-reference-collapsed"), null, { timeout: 5000 });


  const formulaFunctionArgumentsReferencePickerCollapsed = await page.evaluate(() => ({
    collapsed: document.querySelector('[data-data-dialog="function-arguments"]')?.classList.contains("is-reference-collapsed") || false,
    visibleRows: [...document.querySelectorAll("[data-function-argument-row]")]
      .filter((row) => getComputedStyle(row).display !== "none")
      .map((row) => row.dataset.functionArgumentRow || ""),
    actionsHidden: getComputedStyle(document.querySelector('[data-data-dialog="function-arguments"] .excel-mini-actions')).display === "none",
    buttonText: document.querySelector('[data-data-dialog-action="function-argument-reference-toggle"][data-formula-argument-index="1"]')?.textContent.trim() || "",
  }));


  const formulaArgumentA5 = await formulaArgumentCellPoint(5, 1);


  await page.mouse.move(formulaArgumentA4.x, formulaArgumentA4.y);


  await page.mouse.down();


  await page.mouse.move(formulaArgumentA5.x, formulaArgumentA5.y, { steps: 4 });


  await page.mouse.up();


  await page.waitForFunction(() => document.querySelector('[data-formula-argument-input][data-formula-argument-index="1"]')?.value === "A4:A5", null, { timeout: 5000 });


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => !document.querySelector('[data-data-dialog="function-arguments"]')?.classList.contains("is-reference-collapsed"), null, { timeout: 5000 });


  const formulaFunctionArgumentsReferencePickerExpanded = await page.evaluate(() => ({
    collapsed: document.querySelector('[data-data-dialog="function-arguments"]')?.classList.contains("is-reference-collapsed") || false,
    visibleRows: [...document.querySelectorAll("[data-function-argument-row]")]
      .filter((row) => getComputedStyle(row).display !== "none")
      .map((row) => row.dataset.functionArgumentRow || ""),
    activeIndex: document.activeElement?.dataset?.formulaArgumentIndex || "",
  }));


  const formulaArgumentSheets = await page.evaluate(() => {
    const active = document.querySelector(".sheet-tab.is-active");
    const tabs = [...document.querySelectorAll(".sheet-tab")];
    const other = tabs.find((tab) => tab !== active);
    return {
      sourceIndex: Number(active?.dataset.sheetIndex || 0),
      sourceName: active?.textContent?.trim() || "",
      otherIndex: Number(other?.dataset.sheetIndex || 0),
      otherName: other?.textContent?.trim() || "",
    };
  });


  await selectFormulaArgumentInput(2);


  await page.locator(`.sheet-tab[data-sheet-index="${formulaArgumentSheets.otherIndex}"]`).click();


  await page.waitForFunction((index) => document.querySelector(".sheet-tab.is-active")?.dataset.sheetIndex === String(index), formulaArgumentSheets.otherIndex, { timeout: 5000 });


  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    if (!host) return;
    host.scrollLeft = 0;
    host.scrollTop = 0;
    host.dispatchEvent(new Event("scroll", { bubbles: true }));
  });


  await page.waitForSelector('.column-header[data-col="1"]', { timeout: 5000 });


  const formulaArgumentOtherA4 = await formulaArgumentCellPoint(4, 1);


  await page.mouse.click(formulaArgumentOtherA4.x, formulaArgumentOtherA4.y);


  await page.waitForFunction(() => {
    const value = document.querySelector('[data-formula-argument-input][data-formula-argument-index="2"]')?.value || "";
    return value.includes("!A4");
  }, null, { timeout: 5000 });


  const formulaFunctionArgumentsReferencesRaw = await page.evaluate(() => ({
    activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent.trim() || "",
    values: [...document.querySelectorAll("[data-formula-argument-input]")].map((input) => input.value),
    preview: document.querySelector("[data-function-argument-formula]")?.textContent || "",
    result: document.querySelector("[data-function-argument-result]")?.textContent || "",
  }));


  const formulaFunctionArgumentsReferences = {
    ...formulaFunctionArgumentsReferencesRaw,
    pickerCollapsed: formulaFunctionArgumentsReferencePickerCollapsed,
    pickerExpanded: formulaFunctionArgumentsReferencePickerExpanded,
  };


  await page.locator(`.sheet-tab[data-sheet-index="${formulaArgumentSheets.sourceIndex}"]`).click();


  await page.waitForFunction((index) => document.querySelector(".sheet-tab.is-active")?.dataset.sheetIndex === String(index), formulaArgumentSheets.sourceIndex, { timeout: 5000 });


  await page.evaluate(() => window.__webSheetDevSelectRange?.(6, 8, 6, 8));


  await page.click('[data-data-dialog="function-arguments"] [data-data-dialog-action="function-arguments-back"]');


  await page.waitForSelector('[data-data-dialog="insert-function"]', { timeout: 5000 });


  const formulaFunctionArgumentsBack = await page.evaluate(() => ({
    title: document.querySelector("#insertFunctionTitle")?.textContent || "",
    selected: document.querySelector("[data-formula-function-choice].is-active")?.dataset.formulaFunctionChoice || "",
    category: document.querySelector("[data-formula-function-category]")?.value || "",
  }));


  await page.click('[data-data-dialog="insert-function"] [data-data-dialog-action="insert-function-ok"]');


  await page.waitForSelector('[data-data-dialog="function-arguments"]', { timeout: 5000 });


  await page.fill('[data-formula-argument-input][data-formula-argument-index="0"]', "1=0");


  const formulaFunctionArgumentsEdited = await page.evaluate(() => ({
    preview: document.querySelector("[data-function-argument-formula]")?.textContent || "",
    argumentResults: [...document.querySelectorAll(".function-argument-value")].map((item) => item.textContent.trim()),
    result: document.querySelector("[data-function-argument-result]")?.textContent || "",
  }));


  await page.click('[data-data-dialog="function-arguments"] [data-data-dialog-action="function-arguments-ok"]');


  await page.waitForFunction(() => window.__webSheetDevCellModel?.("CHT", 6, 8)?.formula === "=IF(1=0,1,0)", null, { timeout: 5000 });


  const formulaInserted = await page.evaluate(() => ({
    formula: window.__webSheetDevCellModel?.("CHT", 6, 8)?.formula || "",
    display: window.__webSheetDevCellDisplay?.("CHT", 6, 8) || "",
    modelDisplay: window.__webSheetDevCellModel?.("CHT", 6, 8)?.display || "",
    modelCached: window.__webSheetDevCellModel?.("CHT", 6, 8)?.cached ?? null,
  }));


  const formulaFunctionArguments = {
    initial: formulaFunctionArgumentsInitial,
    references: formulaFunctionArgumentsReferences,
    back: formulaFunctionArgumentsBack,
    edited: formulaFunctionArgumentsEdited,
  };


  await page.evaluate(() => window.__webSheetDevSelectRange?.(8, 8, 8, 8));


  await page.click('[data-ribbon-command="formula-function-if"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-function-logical", { timeout: 5000 });


  const formulaLogicalMenu = await page.evaluate(() => ({
    label: document.querySelector("#homeCommandMenu")?.getAttribute("aria-label") || "",
    functions: [...document.querySelectorAll("#homeCommandMenu [data-formula-function-name]")].map((button) => button.dataset.formulaFunctionName),
    syntaxText: document.querySelector("#homeCommandMenu")?.textContent || "",
  }));


  await page.click('#homeCommandMenu [data-formula-function-name="AND"]');


  await page.waitForSelector('[data-data-dialog="function-arguments"]', { timeout: 5000 });


  const formulaMenuArguments = await page.evaluate(() => ({
    title: document.querySelector("#functionArgumentsTitle")?.textContent || "",
    functionName: document.querySelector(".function-argument-heading strong")?.textContent || "",
    labels: [...document.querySelectorAll(".function-argument-name")].map((item) => item.textContent.trim().replace(/:$/u, "")),
    values: [...document.querySelectorAll("[data-formula-argument-input]")].map((input) => input.value),
    argumentResults: [...document.querySelectorAll(".function-argument-value")].map((item) => item.textContent.trim()),
    preview: document.querySelector("[data-function-argument-formula]")?.textContent || "",
    result: document.querySelector("[data-function-argument-result]")?.textContent || "",
    description: document.querySelector("[data-function-argument-description]")?.textContent || "",
  }));


  await page.click('[data-data-dialog="function-arguments"] [data-data-dialog-action="function-arguments-ok"]');


  await page.waitForFunction(() => window.__webSheetDevCellModel?.("CHT", 8, 8)?.formula === "=AND(1=1)", null, { timeout: 5000 });


  const formulaMenuInserted = await page.evaluate(() => ({
    formula: window.__webSheetDevCellModel?.("CHT", 8, 8)?.formula || "",
    display: window.__webSheetDevCellDisplay?.("CHT", 8, 8) || "",
    modelDisplay: window.__webSheetDevCellModel?.("CHT", 8, 8)?.display || "",
    modelCached: window.__webSheetDevCellModel?.("CHT", 8, 8)?.cached ?? null,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('[data-ribbon-command="formula-function-recent"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-function-recent", { timeout: 5000 });


  const formulaRecentMenu = await page.evaluate(() => ({
    label: document.querySelector("#homeCommandMenu")?.getAttribute("aria-label") || "",
    functions: [...document.querySelectorAll("#homeCommandMenu [data-formula-function-name]")].map((button) => button.dataset.formulaFunctionName),
  }));


  await page.keyboard.press("Escape");


  await page.evaluate(() => window.__webSheetDevSelectRange?.(12, 8, 12, 8));


  await page.click('[data-ribbon-command="formula-function-sum"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-function-math", { timeout: 5000 });


  const formulaMathMenu = await page.evaluate(() => ({
    label: document.querySelector("#homeCommandMenu")?.getAttribute("aria-label") || "",
    functions: [...document.querySelectorAll("#homeCommandMenu [data-formula-function-name]")].map((button) => button.dataset.formulaFunctionName),
    hasCategoryDialog: Boolean(document.querySelector("#homeCommandMenu [data-formula-function-category='数学/三角']")),
  }));


  await page.click('#homeCommandMenu [data-formula-function-name="SUM"]');


  await page.waitForSelector('[data-data-dialog="function-arguments"]', { timeout: 5000 });


  const formulaVariableArgumentsInitial = await page.evaluate(() => ({
    title: document.querySelector("#functionArgumentsTitle")?.textContent || "",
    functionName: document.querySelector(".function-argument-heading strong")?.textContent || "",
    labels: [...document.querySelectorAll(".function-argument-name")].map((item) => item.textContent.trim().replace(/:$/u, "")),
    values: [...document.querySelectorAll("[data-formula-argument-input]")].map((input) => input.value),
    description: document.querySelector("[data-function-argument-description]")?.textContent || "",
  }));


  await page.fill('[data-formula-argument-input][data-formula-argument-index="0"]', "2");


  await page.fill('[data-formula-argument-input][data-formula-argument-index="1"]', "3");


  await page.waitForSelector('[data-formula-argument-input][data-formula-argument-index="2"]', { timeout: 5000 });


  await page.keyboard.press("Tab");


  const formulaVariableArgumentsExpanded = await page.evaluate(() => ({
    labels: [...document.querySelectorAll(".function-argument-name")].map((item) => item.textContent.trim().replace(/:$/u, "")),
    values: [...document.querySelectorAll("[data-formula-argument-input]")].map((input) => input.value),
    activeIndex: document.activeElement?.dataset?.formulaArgumentIndex || "",
    argumentResults: [...document.querySelectorAll(".function-argument-value")].map((item) => item.textContent.trim()),
    preview: document.querySelector("[data-function-argument-formula]")?.textContent || "",
    result: document.querySelector("[data-function-argument-result]")?.textContent || "",
    description: document.querySelector("[data-function-argument-description]")?.textContent || "",
  }));


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => window.__webSheetDevCellModel?.("CHT", 12, 8)?.formula === "=SUM(2,3)", null, { timeout: 5000 });


  const formulaVariableArgumentsInserted = await page.evaluate(() => ({
    formula: window.__webSheetDevCellModel?.("CHT", 12, 8)?.formula || "",
    display: window.__webSheetDevCellDisplay?.("CHT", 12, 8) || "",
    modelDisplay: window.__webSheetDevCellModel?.("CHT", 12, 8)?.display || "",
    modelCached: window.__webSheetDevCellModel?.("CHT", 12, 8)?.cached ?? null,
  }));


  const formulaVariableArguments = {
    initial: formulaVariableArgumentsInitial,
    expanded: formulaVariableArgumentsExpanded,
    inserted: formulaVariableArgumentsInserted,
  };


  await page.evaluate(() => window.__webSheetDevSelectRange?.(13, 8, 13, 8));


  await page.click('[data-ribbon-command="formula-function-if"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-function-logical", { timeout: 5000 });


  await page.click('#homeCommandMenu [data-formula-function-name="TRUE"]');


  await page.waitForSelector('[data-data-dialog="function-arguments"]', { timeout: 5000 });


  const formulaZeroArgumentFunctionDialog = await page.evaluate(() => ({
    title: document.querySelector("#functionArgumentsTitle")?.textContent || "",
    functionName: document.querySelector(".function-argument-heading strong")?.textContent || "",
    emptyText: document.querySelector(".function-argument-empty")?.textContent?.trim() || "",
    labels: [...document.querySelectorAll(".function-argument-name")].map((item) => item.textContent.trim().replace(/:$/u, "")),
    preview: document.querySelector("[data-function-argument-formula]")?.textContent || "",
    result: document.querySelector("[data-function-argument-result]")?.textContent || "",
  }));


  await page.click('[data-data-dialog="function-arguments"] [data-data-dialog-action="function-arguments-ok"]');


  await page.waitForFunction(() => window.__webSheetDevCellModel?.("CHT", 13, 8)?.formula === "=TRUE()", null, { timeout: 5000 });


  const formulaZeroArgumentFunctionInserted = await page.evaluate(() => ({
    formula: window.__webSheetDevCellModel?.("CHT", 13, 8)?.formula || "",
    display: window.__webSheetDevCellDisplay?.("CHT", 13, 8) || "",
    modelDisplay: window.__webSheetDevCellModel?.("CHT", 13, 8)?.display || "",
    modelCached: window.__webSheetDevCellModel?.("CHT", 13, 8)?.cached ?? null,
  }));


  const formulaZeroArgumentFunction = {
    dialog: formulaZeroArgumentFunctionDialog,
    inserted: formulaZeroArgumentFunctionInserted,
  };


  await page.click('[data-ribbon-command="formula-function-financial"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-function-financial", { timeout: 5000 });


  const formulaFinancialMenu = await page.evaluate(() => ({
    label: document.querySelector("#homeCommandMenu")?.getAttribute("aria-label") || "",
    functions: [...document.querySelectorAll("#homeCommandMenu [data-formula-function-name]")].map((button) => button.dataset.formulaFunctionName),
    hasCategoryDialog: Boolean(document.querySelector("#homeCommandMenu [data-formula-function-category='財務']")),
  }));


  await page.keyboard.press("Escape");


  await page.click('[data-ribbon-command="formula-function-today"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-function-date", { timeout: 5000 });


  const formulaDateMenu = await page.evaluate(() => ({
    label: document.querySelector("#homeCommandMenu")?.getAttribute("aria-label") || "",
    functions: [...document.querySelectorAll("#homeCommandMenu [data-formula-function-name]")].map((button) => button.dataset.formulaFunctionName),
    hasCategoryDialog: Boolean(document.querySelector("#homeCommandMenu [data-formula-function-category='日付/時刻']")),
  }));


  await page.keyboard.press("Escape");


  await page.click('[data-ribbon-command="formula-function-textjoin"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-function-text", { timeout: 5000 });


  const formulaTextMenu = await page.evaluate(() => ({
    label: document.querySelector("#homeCommandMenu")?.getAttribute("aria-label") || "",
    functions: [...document.querySelectorAll("#homeCommandMenu [data-formula-function-name]")].map((button) => button.dataset.formulaFunctionName),
    hasCategoryDialog: Boolean(document.querySelector("#homeCommandMenu [data-formula-function-category='文字列']")),
  }));


  await page.keyboard.press("Escape");


  await page.click('[data-ribbon-command="formula-function-more"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-function-more", { timeout: 5000 });


  const formulaMoreMenu = await page.evaluate(() => ({
    label: document.querySelector("#homeCommandMenu")?.getAttribute("aria-label") || "",
    functions: [...document.querySelectorAll("#homeCommandMenu [data-formula-function-name]")].map((button) => button.dataset.formulaFunctionName),
    hasAllFunctionsDialog: Boolean(document.querySelector("#homeCommandMenu [data-formula-function-category='すべて表示']")),
  }));


  await page.keyboard.press("Escape");


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(200, 80, 10);
    window.__webSheetDevSetCell?.(201, 80, 20);
    window.__webSheetDevSelectRange?.(202, 80, 202, 80);
  });


  await page.click('[data-ribbon-command="formula-autosum"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-autosum", { timeout: 5000 });


  const formulaAutoSumMenu = await page.evaluate(() => ({
    label: document.querySelector("#homeCommandMenu")?.getAttribute("aria-label") || "",
    functions: [...document.querySelectorAll("#homeCommandMenu [data-formula-autosum-function]")].map((button) => button.dataset.formulaAutosumFunction),
    hasMoreFunctions: Boolean(document.querySelector("#homeCommandMenu [data-formula-function-category='すべて表示']")),
  }));


  await page.click('#homeCommandMenu [data-formula-autosum-function="AVERAGE"]');


  await page.waitForFunction(() => window.__webSheetDevCellModel?.("CHT", 202, 80)?.formula === "=AVERAGE(CB200:CB201)", null, { timeout: 5000 });


  const formulaAutoSumInserted = await page.evaluate(() => ({
    formula: window.__webSheetDevCellModel?.("CHT", 202, 80)?.formula || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(7, 8, "=ROUND(SUM(H2:H3),2)");
    window.__webSheetDevSelectRange?.(7, 8, 7, 8);
  });


  await page.waitForFunction(() => window.__webSheetDevCellModel?.("CHT", 7, 8)?.formula === "=ROUND(SUM(H2:H3),2)", null, { timeout: 5000 });


  await page.click('[data-ribbon-command="formula-show-formulas"]');


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().showFormulas === true, null, { timeout: 5000 });


  const formulaShowState = await page.evaluate(() => ({
    showFormulas: window.__webSheetDevWorkbookState?.().showFormulas,
    buttonActive: document.querySelector('[data-ribbon-command="formula-show-formulas"]')?.classList.contains("is-active"),
  }));


  await page.click('[data-ribbon-command="formula-show-formulas"]');


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().showFormulas === false, null, { timeout: 5000 });


  await page.evaluate(() => {
    window.__webSheetDevSetActiveSheet?.("CHT");
    window.__webSheetDevSelectRange?.(7, 8, 7, 8);
  });


  await page.click('[data-ribbon-command="formula-trace-precedents"]');


  await page.waitForFunction(() => (window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.precedents || []).length >= 1, null, { timeout: 5000 });


  const formulaTrace = await page.evaluate(() => ({
    precedents: window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.precedents || [],
    arrows: window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.arrows || [],
    highlighted: document.querySelectorAll(".sheet-cell.formula-audit-precedent").length,
    overlayArrows: document.querySelectorAll(".formula-audit-overlay .formula-audit-arrow").length,
  }));


  await page.click('[data-ribbon-command="formula-remove-arrows"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-remove-arrows", { timeout: 5000 });


  const formulaRemoveArrowsMenu = await page.evaluate(() => ({
    label: document.querySelector("#homeCommandMenu")?.getAttribute("aria-label") || "",
    actions: [...document.querySelectorAll("#homeCommandMenu [data-formula-audit-menu-action]")].map((button) => button.dataset.formulaAuditMenuAction),
  }));


  await page.click('#homeCommandMenu [data-formula-audit-menu-action="remove-precedent-arrows"]');


  await page.waitForFunction(() => (window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.arrows || []).length === 0, null, { timeout: 5000 });


  const formulaRemovePrecedents = await page.evaluate(() => ({
    precedents: window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.precedents || [],
    arrows: window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.arrows || [],
  }));


  await page.evaluate(() => window.__webSheetDevSelectRange?.(2, 8, 2, 8));


  await page.click('[data-ribbon-command="formula-trace-dependents"]');


  await page.waitForFunction(() => (window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.dependents || []).length >= 1, null, { timeout: 5000 });


  const formulaTraceDependents = await page.evaluate(() => ({
    dependents: window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.dependents || [],
    arrows: window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.arrows || [],
    overlayArrows: document.querySelectorAll(".formula-audit-overlay .formula-audit-arrow.is-dependents").length,
  }));


  await page.click('[data-ribbon-command="formula-remove-arrows"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-remove-arrows", { timeout: 5000 });


  await page.click('#homeCommandMenu [data-formula-audit-menu-action="remove-dependent-arrows"]');


  await page.waitForFunction(() => (window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.dependents || []).length === 0, null, { timeout: 5000 });


  const formulaRemoveDependents = await page.evaluate(() => ({
    dependents: window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.dependents || [],
    arrows: window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.arrows || [],
  }));


  formulaTrace.dependentsTrace = formulaTraceDependents;


  formulaTrace.dependentsRemoval = formulaRemoveDependents;


  await page.evaluate(() => window.__webSheetDevSelectRange?.(7, 8, 7, 8));


  await page.click('[data-ribbon-command="formula-evaluate"]');


  await page.waitForSelector("[data-data-dialog='formula-evaluate']", { timeout: 5000 });


  const formulaEvaluateInitial = await page.evaluate(() => ({
    title: document.querySelector("#formulaEvaluateTitle")?.textContent || "",
    location: document.querySelector("[data-formula-evaluate-location]")?.textContent || "",
    expression: document.querySelector("[data-formula-evaluate-expression]")?.textContent || "",
    stepCount: document.querySelectorAll("[data-formula-evaluate-steps] > div").length,
    activeStep: document.querySelector("[data-formula-evaluate-steps] > div.is-active")?.textContent || "",
    result: document.querySelector("[data-formula-evaluate-result]")?.value || "",
    stepInDisabled: document.querySelector("[data-data-dialog-action='formula-evaluate-step-in']")?.disabled ?? true,
    stepOutDisabled: document.querySelector("[data-data-dialog-action='formula-evaluate-step-out']")?.disabled ?? true,
  }));


  await page.click("[data-data-dialog-action='formula-evaluate-step-in']");


  await page.waitForFunction(() => (document.querySelector("[data-formula-evaluate-location]")?.textContent || "").includes("SUM にステップ イン"), null, { timeout: 5000 });


  const formulaEvaluateStepIn = await page.evaluate(() => ({
    location: document.querySelector("[data-formula-evaluate-location]")?.textContent || "",
    expression: document.querySelector("[data-formula-evaluate-expression]")?.textContent || "",
    activeStep: document.querySelector("[data-formula-evaluate-steps] > div.is-active")?.textContent || "",
    result: document.querySelector("[data-formula-evaluate-result]")?.value || "",
    stepInDisabled: document.querySelector("[data-data-dialog-action='formula-evaluate-step-in']")?.disabled ?? true,
    stepOutDisabled: document.querySelector("[data-data-dialog-action='formula-evaluate-step-out']")?.disabled ?? true,
  }));


  await page.click("[data-data-dialog-action='formula-evaluate-step']");


  await page.waitForFunction(() => (document.querySelector("[data-formula-evaluate-steps] > div.is-active")?.textContent || "").includes("H2:H3"), null, { timeout: 5000 });


  const formulaEvaluateStep = await page.evaluate(() => ({
    expression: document.querySelector("[data-formula-evaluate-expression]")?.textContent || "",
    activeStep: document.querySelector("[data-formula-evaluate-steps] > div.is-active")?.textContent || "",
    selection: document.querySelector("#selectedCellLabel")?.textContent || "",
    stepInDisabled: document.querySelector("[data-data-dialog-action='formula-evaluate-step-in']")?.disabled ?? true,
    stepOutDisabled: document.querySelector("[data-data-dialog-action='formula-evaluate-step-out']")?.disabled ?? true,
  }));


  await page.click("[data-data-dialog-action='formula-evaluate-step-out']");


  await page.waitForFunction(() => (document.querySelector("[data-formula-evaluate-expression]")?.textContent || "").includes("ROUND(SUM"), null, { timeout: 5000 });


  const formulaEvaluateStepOut = await page.evaluate(() => ({
    location: document.querySelector("[data-formula-evaluate-location]")?.textContent || "",
    expression: document.querySelector("[data-formula-evaluate-expression]")?.textContent || "",
    activeStep: document.querySelector("[data-formula-evaluate-steps] > div.is-active")?.textContent || "",
    stepOutDisabled: document.querySelector("[data-data-dialog-action='formula-evaluate-step-out']")?.disabled ?? true,
  }));


  await page.keyboard.press("Escape");


  await page.waitForSelector("[data-data-dialog='formula-evaluate']", { state: "detached", timeout: 5000 });


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(204, 80, "=10+5");
    window.__webSheetDevSetCell?.(205, 80, "=CB204+1");
    window.__webSheetDevSelectRange?.(205, 80, 205, 80);
  });


  await page.click('[data-ribbon-command="formula-evaluate"]');


  await page.waitForSelector("[data-data-dialog='formula-evaluate']", { timeout: 5000 });


  const formulaEvaluateReferenceInitial = await page.evaluate(() => ({
    location: document.querySelector("[data-formula-evaluate-location]")?.textContent || "",
    expression: document.querySelector("[data-formula-evaluate-expression]")?.textContent || "",
    result: document.querySelector("[data-formula-evaluate-result]")?.value || "",
    stepInDisabled: document.querySelector("[data-data-dialog-action='formula-evaluate-step-in']")?.disabled ?? true,
    stepOutDisabled: document.querySelector("[data-data-dialog-action='formula-evaluate-step-out']")?.disabled ?? true,
  }));


  await page.click("[data-data-dialog-action='formula-evaluate-step']");


  await page.waitForFunction(() => (document.querySelector("[data-formula-evaluate-steps] > div.is-active")?.textContent || "").includes("CB204"), null, { timeout: 5000 });


  const formulaEvaluateReferenceStep = await page.evaluate(() => ({
    location: document.querySelector("[data-formula-evaluate-location]")?.textContent || "",
    expression: document.querySelector("[data-formula-evaluate-expression]")?.textContent || "",
    activeStep: document.querySelector("[data-formula-evaluate-steps] > div.is-active")?.textContent || "",
    selection: document.querySelector("#selectedCellLabel")?.textContent || "",
    stepInDisabled: document.querySelector("[data-data-dialog-action='formula-evaluate-step-in']")?.disabled ?? true,
    stepOutDisabled: document.querySelector("[data-data-dialog-action='formula-evaluate-step-out']")?.disabled ?? true,
  }));


  await page.click("[data-data-dialog-action='formula-evaluate-step-in']");


  await page.waitForFunction(() => (document.querySelector("[data-formula-evaluate-location]")?.textContent || "").includes("CB204 にステップ イン"), null, { timeout: 5000 });


  const formulaEvaluateReferenceStepIn = await page.evaluate(() => ({
    location: document.querySelector("[data-formula-evaluate-location]")?.textContent || "",
    expression: document.querySelector("[data-formula-evaluate-expression]")?.textContent || "",
    activeStep: document.querySelector("[data-formula-evaluate-steps] > div.is-active")?.textContent || "",
    result: document.querySelector("[data-formula-evaluate-result]")?.value || "",
    stepOutDisabled: document.querySelector("[data-data-dialog-action='formula-evaluate-step-out']")?.disabled ?? true,
  }));


  await page.click("[data-data-dialog-action='formula-evaluate-step-out']");


  await page.waitForFunction(() => (document.querySelector("[data-formula-evaluate-location]")?.textContent || "").includes("CB205"), null, { timeout: 5000 });


  const formulaEvaluateReferenceStepOut = await page.evaluate(() => ({
    location: document.querySelector("[data-formula-evaluate-location]")?.textContent || "",
    expression: document.querySelector("[data-formula-evaluate-expression]")?.textContent || "",
    activeStep: document.querySelector("[data-formula-evaluate-steps] > div.is-active")?.textContent || "",
    stepOutDisabled: document.querySelector("[data-data-dialog-action='formula-evaluate-step-out']")?.disabled ?? true,
  }));


  await page.keyboard.press("Escape");


  await page.waitForSelector("[data-data-dialog='formula-evaluate']", { state: "detached", timeout: 5000 });


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(203, 80, "=SUM(H2:H3)");
    window.__webSheetDevSelectRange?.(203, 80, 203, 80);
  });


  await page.click('[data-ribbon-command="formula-watch-window"]');


  await page.waitForSelector("#formulaWatchWindow", { timeout: 5000 });


  await page.click('#formulaWatchWindow [data-formula-watch-action="add"]');


  await page.waitForFunction(() => (window.__webSheetDevFormulaWatches?.() || []).length >= 1, null, { timeout: 5000 });


  const formulaWatchWindow = await page.evaluate(() => ({
    open: Boolean(document.querySelector("#formulaWatchWindow")),
    title: document.querySelector("#formulaWatchTitle")?.textContent || "",
    rows: [...document.querySelectorAll("#formulaWatchWindow tbody tr:not(.is-empty)")].map((row) => row.textContent.trim()),
    watches: window.__webSheetDevFormulaWatches?.() || [],
  }));


  await page.locator("#formulaWatchWindow tbody tr", { hasText: "CB203" }).locator('[data-formula-watch-action="select"]').click();


  await page.waitForFunction(() => (document.querySelector("#selectedCellLabel")?.textContent || "").includes("CB203"), null, { timeout: 5000 });


  const formulaWatchSelect = await page.evaluate(() => ({
    selected: document.querySelector("#selectedCellLabel")?.textContent || "",
    formula: window.__webSheetDevCellModel?.("CHT", 203, 80)?.formula || "",
  }));


  await page.check('#formulaWatchWindow [data-formula-watch-id]');


  await page.click('#formulaWatchWindow [data-formula-watch-action="delete"]');


  await page.waitForFunction(() => (window.__webSheetDevFormulaWatches?.() || []).length === 0, null, { timeout: 5000 });


  const formulaWatchDelete = await page.evaluate(() => ({
    rows: document.querySelectorAll("#formulaWatchWindow tbody tr:not(.is-empty)").length,
    emptyText: document.querySelector("#formulaWatchWindow tbody tr.is-empty")?.textContent || "",
  }));


  await page.click('#formulaWatchWindow [data-formula-watch-action="close"]');


  await page.waitForSelector("#formulaWatchWindow", { state: "detached", timeout: 5000 });


  await page.evaluate(() => {
    window.__webSheetDevSetCell?.(1, 80, "=1/0");
    window.__webSheetDevSetCell?.(2, 80, "=H2/0");
    window.__webSheetDevSetCell?.(4, 80, "=CB5");
    window.__webSheetDevSetCell?.(5, 80, "=CB4");
    window.__webSheetDevSelectRange?.(1, 80, 1, 80);
  });


  await page.waitForFunction(() => String(window.__webSheetDevCellDisplay?.("CHT", 1, 80) || "").startsWith("#DIV/0!"), null, { timeout: 5000 });


  await page.click('[data-ribbon-command="formula-error-checking"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-error-checking", { timeout: 5000 });


  const formulaErrorMenu = await page.evaluate(() => ({
    label: document.querySelector("#homeCommandMenu")?.getAttribute("aria-label") || "",
    actions: [...document.querySelectorAll("#homeCommandMenu [data-formula-audit-menu-action]")].map((button) => button.dataset.formulaAuditMenuAction),
    circularItems: [...document.querySelectorAll("#homeCommandMenu [data-formula-audit-menu-action^='circular:']")].map((button) => ({
      action: button.dataset.formulaAuditMenuAction,
      text: button.textContent.replace(/\s+/g, " ").trim(),
    })),
  }));


  await page.click("#homeCommandMenu [data-formula-audit-menu-action^='circular:']");


  await page.waitForFunction(() => {
    const label = document.querySelector("#selectedCellLabel")?.textContent || "";
    return label.includes("CB4") || label.includes("CB5");
  }, null, { timeout: 5000 });


  const formulaCircularSelection = await page.evaluate(() => ({
    selected: document.querySelector("#selectedCellLabel")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
    cell4: window.__webSheetDevCellModel?.("CHT", 4, 80)?.formula || "",
    cell5: window.__webSheetDevCellModel?.("CHT", 5, 80)?.formula || "",
  }));


  formulaErrorMenu.circularSelection = formulaCircularSelection;


  await page.click('[data-ribbon-command="formula-error-checking"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-error-checking", { timeout: 5000 });


  await page.click('#homeCommandMenu [data-formula-audit-menu-action="run-error-checking"]');


  await page.waitForSelector("[data-data-dialog='formula-error-checking']", { timeout: 5000 });


  const formulaErrorChecking = await page.evaluate(() => ({
    title: document.querySelector("#formulaErrorTitle")?.textContent || "",
    summary: document.querySelector("[data-formula-error-summary]")?.textContent || "",
    location: document.querySelector("[data-formula-error-location]")?.textContent || "",
    value: document.querySelector("[data-formula-error-value]")?.textContent || "",
    message: document.querySelector("[data-formula-error-message]")?.textContent || "",
    formula: document.querySelector("[data-formula-error-formula]")?.value || "",
    entries: [...document.querySelectorAll("[data-formula-error-list] button")].map((button) => button.textContent.trim()),
  }));


  await page.click("[data-data-dialog-action='formula-error-ignore']");


  await page.waitForTimeout(150);


  const formulaErrorIgnore = await page.evaluate(() => ({
    dialogOpen: Boolean(document.querySelector("[data-data-dialog='formula-error-checking']")),
    ignored: window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.ignoredErrors || [],
  }));


  await page.keyboard.press("Escape");


  await page.click('[data-ribbon-command="formula-error-checking"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-error-checking", { timeout: 5000 });


  await page.click('#homeCommandMenu [data-formula-audit-menu-action="trace-errors"]');


  await page.waitForFunction(() => (window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.arrows || []).length >= 1, null, { timeout: 5000 });


  const formulaErrorTrace = await page.evaluate(() => ({
    arrows: window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.arrows || [],
    highlighted: document.querySelectorAll(".sheet-cell.formula-audit-precedent").length,
    overlayArrows: document.querySelectorAll(".formula-audit-overlay .formula-audit-arrow").length,
  }));


  await page.click('[data-ribbon-command="formula-error-checking"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-error-checking", { timeout: 5000 });


  await page.click('#homeCommandMenu [data-formula-audit-menu-action="reset-ignored-errors"]');


  await page.waitForFunction(() => (window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.ignoredErrors || []).length === 0, null, { timeout: 5000 });


  const formulaErrorResetIgnored = await page.evaluate(() => ({
    ignored: window.__webSheetDevSheetModel?.("CHT")?.formulaAudits?.ignoredErrors || [],
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('[data-ribbon-command="formula-remove-arrows"]');


  await page.waitForSelector("#homeCommandMenu.is-formula-remove-arrows", { timeout: 5000 });


  await page.click('#homeCommandMenu [data-formula-audit-menu-action="remove-all-arrows"]');


  await page.click('[data-ribbon-command="formula-calculation-options"]');


  await page.waitForSelector("#homeCommandMenu.is-calculation", { timeout: 5000 });


  const formulaCalculationMenuInitial = await page.evaluate(() => ({
    label: document.querySelector("#homeCommandMenu")?.getAttribute("aria-label") || "",
    checked: [...document.querySelectorAll("#homeCommandMenu [data-formula-calculation-action][aria-checked='true']")].map((button) => button.dataset.formulaCalculationAction),
    actions: [...document.querySelectorAll("#homeCommandMenu [data-formula-calculation-action]")].map((button) => button.dataset.formulaCalculationAction),
    currentText: [...document.querySelectorAll("#homeCommandMenu [role='menuitem']")].map((button) => button.textContent.replace(/\s+/g, " ").trim()).find((text) => text.includes("現在:")) || "",
  }));


  await page.click('#homeCommandMenu [data-formula-calculation-action="manual"]');


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().calcProperties?.calcMode === "manual", null, { timeout: 5000 });


  const formulaCalculationManual = await page.evaluate(() => ({
    status: document.querySelector("#statusText")?.textContent || "",
    calculationMode: window.__webSheetDevWorkbookState?.().calculationMode,
    calcProperties: window.__webSheetDevWorkbookState?.().calcProperties,
  }));


  await page.click('[data-ribbon-command="formula-calculation-options"]');


  await page.waitForSelector("#homeCommandMenu.is-calculation", { timeout: 5000 });


  const formulaCalculationMenuManual = await page.evaluate(() => ({
    checked: [...document.querySelectorAll("#homeCommandMenu [data-formula-calculation-action][aria-checked='true']")].map((button) => button.dataset.formulaCalculationAction),
    currentText: [...document.querySelectorAll("#homeCommandMenu [role='menuitem']")].map((button) => button.textContent.replace(/\s+/g, " ").trim()).find((text) => text.includes("現在:")) || "",
  }));


  await page.click('#homeCommandMenu [data-formula-calculation-action="except-tables"]');


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().calcProperties?.calcMode === "autoNoTable", null, { timeout: 5000 });


  const formulaCalculationExceptTables = await page.evaluate(() => ({
    status: document.querySelector("#statusText")?.textContent || "",
    calculationMode: window.__webSheetDevWorkbookState?.().calculationMode,
    calcProperties: window.__webSheetDevWorkbookState?.().calcProperties,
  }));


  await page.click('[data-ribbon-command="formula-calculation-options"]');


  await page.waitForSelector("#homeCommandMenu.is-calculation", { timeout: 5000 });


  await page.click('#homeCommandMenu [data-formula-calculation-action="automatic"]');


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().calcProperties?.calcMode === "auto", null, { timeout: 5000 });


  const formulaCalculationAutomatic = await page.evaluate(() => ({
    status: document.querySelector("#statusText")?.textContent || "",
    calculationMode: window.__webSheetDevWorkbookState?.().calculationMode,
    calcProperties: window.__webSheetDevWorkbookState?.().calcProperties,
  }));


  await page.click('[data-ribbon-command="formula-calculate-now"]');


  const formulaCalculation = await page.evaluate(() => ({
    status: document.querySelector("#statusText")?.textContent || "",
    calculationMode: window.__webSheetDevWorkbookState?.().calculationMode,
    calcProperties: window.__webSheetDevWorkbookState?.().calcProperties,
  }));


  formulaCalculation.menuInitial = formulaCalculationMenuInitial;


  formulaCalculation.manual = formulaCalculationManual;


  formulaCalculation.menuManual = formulaCalculationMenuManual;


  formulaCalculation.exceptTables = formulaCalculationExceptTables;


  formulaCalculation.automatic = formulaCalculationAutomatic;


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().calcProperties?.calcMode === "autoNoTable", null, { timeout: 5000 });


  formulaCalculation.undo = await page.evaluate(() => ({
    status: document.querySelector("#statusText")?.textContent || "",
    calculationMode: window.__webSheetDevWorkbookState?.().calculationMode,
    calcProperties: window.__webSheetDevWorkbookState?.().calcProperties,
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().calcProperties?.calcMode === "auto", null, { timeout: 5000 });


  formulaCalculation.redo = await page.evaluate(() => ({
    status: document.querySelector("#statusText")?.textContent || "",
    calculationMode: window.__webSheetDevWorkbookState?.().calculationMode,
    calcProperties: window.__webSheetDevWorkbookState?.().calcProperties,
  }));


  await page.click('[data-ribbon-command="formula-calculate-sheet"]');


  formulaCalculation.sheet = await page.evaluate(() => ({
    status: document.querySelector("#statusText")?.textContent || "",
    activeSheetName: window.__webSheetDevWorkbookState?.().activeSheetName,
  }));



  await page.click('[data-ribbon-tab="view"]');


  await page.waitForFunction(() => document.querySelector(".ribbon-tab.is-active")?.textContent === "表示", null, { timeout: 10000 });


  const viewRibbon = await page.evaluate(() => {
    const panel = document.querySelector('[data-ribbon-panel="view"]');
    return {
      activeRibbonTab: document.querySelector(".ribbon-tab.is-active")?.textContent,
      panelVisible: Boolean(panel?.classList.contains("is-active")),
      groupLabels: [...(panel?.querySelectorAll(".ribbon-group-label") || [])].map((label) => label.textContent.trim()),
      allCommands: [...(panel?.querySelectorAll("[data-ribbon-command]") || [])].map((button) => button.dataset.ribbonCommand),
      enabledCommands: [...(panel?.querySelectorAll("[data-ribbon-command]:not(:disabled)") || [])].map((button) => button.dataset.ribbonCommand),
      disabledCommands: [...(panel?.querySelectorAll("[data-ribbon-command]:disabled") || [])].map((button) => button.dataset.ribbonCommand),
      activeCommands: [...(panel?.querySelectorAll("[data-ribbon-command].is-active") || [])].map((button) => button.dataset.ribbonCommand),
      showCheckBoxCount: panel?.querySelectorAll(".ribbon-view-show-group .ribbon-check-box").length || 0,
      hasFreezeSplit: Boolean(panel?.querySelector("[data-view-freeze-split]")),
      windowCommandTitles: Object.fromEntries(["view-freeze-panes", "view-split", "view-new-window", "view-arrange-all", "view-hide-window", "view-unhide-window", "view-side-by-side", "view-sync-scrolling", "view-reset-window-position", "view-switch-windows"].map((command) => [command, panel?.querySelector(`[data-ribbon-command="${command}"]`)?.getAttribute("title") || ""])),
      panelHeight: Math.round(document.querySelector(".ribbon-panels")?.getBoundingClientRect().height || 0),
    };
  });


  await page.evaluate(() => {
    let webSheetFullscreenElement = null;
    Object.defineProperty(document, "fullscreenElement", {
      configurable: true,
      get: () => webSheetFullscreenElement,
    });
    document.documentElement.requestFullscreen = async () => {
      webSheetFullscreenElement = document.documentElement;
      document.dispatchEvent(new Event("fullscreenchange"));
    };
    document.exitFullscreen = async () => {
      webSheetFullscreenElement = null;
      document.dispatchEvent(new Event("fullscreenchange"));
    };
  });


  await page.click('[data-ribbon-command="view-full-screen"]');


  await page.waitForFunction(() => document.body.classList.contains("is-websheet-fullscreen"), null, { timeout: 5000 });


  const viewFullScreenOn = await page.evaluate(() => ({
    modelState: window.__webSheetDevWorkbookState?.().fullScreenView,
    bodyClass: document.body.classList.contains("is-websheet-fullscreen"),
    buttonActive: document.querySelector('[data-ribbon-command="view-full-screen"]')?.classList.contains("is-active") || false,
    panelsDisplay: getComputedStyle(document.querySelector(".ribbon-panels")).display,
    formulaBarDisplay: getComputedStyle(document.querySelector(".formula-bar")).display,
  }));


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(() => !document.body.classList.contains("is-websheet-fullscreen"), null, { timeout: 5000 });


  const fullScreenUndoRedo = {
    undo: await page.evaluate(() => ({
      modelState: window.__webSheetDevWorkbookState?.().fullScreenView,
      bodyClass: document.body.classList.contains("is-websheet-fullscreen"),
      buttonActive: document.querySelector('[data-ribbon-command="view-full-screen"]')?.classList.contains("is-active") || false,
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(() => document.body.classList.contains("is-websheet-fullscreen"), null, { timeout: 5000 });


  fullScreenUndoRedo.redo = await page.evaluate(() => ({
    modelState: window.__webSheetDevWorkbookState?.().fullScreenView,
    bodyClass: document.body.classList.contains("is-websheet-fullscreen"),
    buttonActive: document.querySelector('[data-ribbon-command="view-full-screen"]')?.classList.contains("is-active") || false,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.evaluate(() => window.__webSheetDevExitFullScreen?.());


  await page.waitForFunction(() => !document.body.classList.contains("is-websheet-fullscreen"), null, { timeout: 5000 });


  const viewFullScreenOff = await page.evaluate(() => ({
    modelState: window.__webSheetDevWorkbookState?.().fullScreenView,
    bodyClass: document.body.classList.contains("is-websheet-fullscreen"),
    buttonActive: document.querySelector('[data-ribbon-command="view-full-screen"]')?.classList.contains("is-active") || false,
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.click('[data-ribbon-command="view-sheet-view-new"]');


  await page.waitForFunction(() => Boolean(window.__webSheetDevSheetModel?.("CHT")?.activeSheetViewId), null, { timeout: 5000 });


  const sheetViewNew = await page.evaluate(() => {
    const sheet = window.__webSheetDevSheetModel?.("CHT") || {};
    const active = (sheet.sheetViews || []).find((view) => view.id === sheet.activeSheetViewId) || null;
    return {
      activeId: sheet.activeSheetViewId || "",
      count: (sheet.sheetViews || []).length,
      temporary: active?.temporary,
      gridMarked: document.querySelector(".sheet-grid")?.classList.contains("has-sheet-view"),
      newActive: document.querySelector('[data-ribbon-command="view-sheet-view-new"]')?.classList.contains("is-active"),
      keepDisabled: document.querySelector('[data-ribbon-command="view-sheet-view-keep"]')?.disabled,
      exitDisabled: document.querySelector('[data-ribbon-command="view-sheet-view-exit"]')?.disabled,
    };
  });


  await page.click('[data-ribbon-command="view-sheet-view-keep"]');


  await page.waitForFunction(() => {
    const sheet = window.__webSheetDevSheetModel?.("CHT") || {};
    return (sheet.sheetViews || []).some((view) => view.id === sheet.activeSheetViewId && view.temporary === false);
  }, null, { timeout: 5000 });


  const sheetViewKept = await page.evaluate(() => {
    const sheet = window.__webSheetDevSheetModel?.("CHT") || {};
    return (sheet.sheetViews || []).find((view) => view.id === sheet.activeSheetViewId) || null;
  });


  await page.click('[data-ribbon-command="view-sheet-view-options"]');


  await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });


  const sheetViewMenuBeforeRename = await page.evaluate(() => ({
    labels: [...document.querySelectorAll("#viewCommandMenu [data-view-menu-action]")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
    descriptions: [...document.querySelectorAll("#viewCommandMenu .home-command-menu-text small")].map((node) => node.textContent.trim()),
    accessKeysUnderlined: [...document.querySelectorAll("#viewCommandMenu .menu-access-key")].every((item) => getComputedStyle(item).textDecorationLine.includes("underline")),
  }));


  await page.keyboard.press("R");


  await page.waitForSelector('[data-data-dialog="sheet-view-name"]', { timeout: 5000 });


  await page.locator('[data-data-field="sheetViewName"]').fill("表示ビュー");


  await page.locator('[data-data-field="sheetViewName"]').press("Enter");


  await page.waitForSelector('[data-data-dialog="sheet-view-name"]', { state: "detached", timeout: 5000 });


  const sheetViewRenamed = await page.evaluate(() => {
    const sheet = window.__webSheetDevSheetModel?.("CHT") || {};
    return (sheet.sheetViews || []).find((view) => view.id === sheet.activeSheetViewId) || null;
  });


  await page.click('[data-ribbon-command="view-sheet-view-exit"]');


  await page.waitForFunction(() => !window.__webSheetDevSheetModel?.("CHT")?.activeSheetViewId, null, { timeout: 5000 });


  const sheetViewExited = await page.evaluate(() => ({
    activeId: window.__webSheetDevSheetModel?.("CHT")?.activeSheetViewId || "",
    count: window.__webSheetDevSheetModel?.("CHT")?.sheetViews?.length || 0,
    gridMarked: document.querySelector(".sheet-grid")?.classList.contains("has-sheet-view"),
    keepDisabled: document.querySelector('[data-ribbon-command="view-sheet-view-keep"]')?.disabled,
    exitDisabled: document.querySelector('[data-ribbon-command="view-sheet-view-exit"]')?.disabled,
  }));


  await page.click('[data-ribbon-command="view-sheet-view-options"]');


  await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });


  await page.locator('#viewCommandMenu [data-view-menu-action^="sheet-view:activate:"]').first().click();


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.activeSheetViewId, null, { timeout: 5000 });


  const sheetViewReactivated = await page.evaluate(() => {
    const sheet = window.__webSheetDevSheetModel?.("CHT") || {};
    return {
      activeId: sheet.activeSheetViewId || "",
      activeName: (sheet.sheetViews || []).find((view) => view.id === sheet.activeSheetViewId)?.name || "",
      gridMarked: document.querySelector(".sheet-grid")?.classList.contains("has-sheet-view"),
    };
  });


  await page.evaluate(() => {
    const sheet = window.__webSheetDevSheetModel?.("CHT") || {};
    if (sheet.showGridLines !== false) document.querySelector('[data-ribbon-command="view-gridlines"]')?.click();
  });


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showGridLines === false, null, { timeout: 5000 });


  await page.click('[data-ribbon-command="view-sheet-view-options"]');


  await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });


  await page.waitForSelector('#viewCommandMenu [data-view-menu-action="sheet-view:update"]', { timeout: 5000 });


  await page.keyboard.press("U");


  await page.waitForFunction(() => {
    const sheet = window.__webSheetDevSheetModel?.("CHT") || {};
    const view = (sheet.sheetViews || []).find((item) => item.id === sheet.activeSheetViewId);
    return view?.snapshot?.showGridLines === false && view?.temporary === false;
  }, null, { timeout: 5000 });


  const sheetViewUpdated = await page.evaluate(() => {
    const sheet = window.__webSheetDevSheetModel?.("CHT") || {};
    const view = (sheet.sheetViews || []).find((item) => item.id === sheet.activeSheetViewId) || null;
    return {
      activeName: view?.name || "",
      temporary: view?.temporary,
      snapshotGridlines: view?.snapshot?.showGridLines,
      status: document.querySelector("#statusText")?.textContent || "",
    };
  });


  await page.click('[data-ribbon-command="view-sheet-view-options"]');


  await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });


  const sheetViewMenuBeforeCopy = await page.evaluate(() => ({
    actions: [...document.querySelectorAll("#viewCommandMenu [data-view-menu-action]")].map((button) => button.dataset.viewMenuAction || ""),
    accessKeys: [...document.querySelectorAll("#viewCommandMenu [data-access-key]")].map((button) => button.dataset.accessKey || ""),
  }));


  await page.keyboard.press("C");


  await page.waitForFunction(() => {
    const sheet = window.__webSheetDevSheetModel?.("CHT") || {};
    const view = (sheet.sheetViews || []).find((item) => item.id === sheet.activeSheetViewId);
    return Boolean(view?.name?.includes("コピー")) && view?.temporary === false;
  }, null, { timeout: 5000 });


  const sheetViewDuplicated = await page.evaluate(() => {
    const sheet = window.__webSheetDevSheetModel?.("CHT") || {};
    const view = (sheet.sheetViews || []).find((item) => item.id === sheet.activeSheetViewId) || null;
    return {
      activeName: view?.name || "",
      temporary: view?.temporary,
      snapshotGridlines: view?.snapshot?.showGridLines,
      baseGridlines: view?.baseView?.showGridLines,
      count: (sheet.sheetViews || []).length,
      status: document.querySelector("#statusText")?.textContent || "",
    };
  });


  await page.click('[data-ribbon-command="view-sheet-view-exit"]');


  await page.waitForFunction(() => !window.__webSheetDevSheetModel?.("CHT")?.activeSheetViewId, null, { timeout: 5000 });


  await page.evaluate(() => window.__webSheetDevSetActiveSheet?.("CHT"));


  const viewZoomBefore = await page.evaluate(() => window.__webSheetDevSheetModel?.("CHT")?.zoomScale || 100);


  await page.click('[data-ribbon-command="view-zoom"]');


  await page.waitForSelector("#zoomDialog", { timeout: 5000 });


	  const viewZoomDialog = await page.evaluate(() => {
	    const dialog = document.querySelector("#zoomDialog")?.getBoundingClientRect();
	    const anchor = document.querySelector('[data-ribbon-command="view-zoom"]')?.getBoundingClientRect();
	    return {
	      title: document.querySelector(".zoom-dialog-title")?.textContent || "",
	      nearRibbonButton: Boolean(dialog && anchor && Math.abs(dialog.left - anchor.left) < 180 && dialog.top >= anchor.bottom - 4),
	      selected: document.querySelector('#zoomDialog input[name="zoomDialogOption"]:checked')?.value || "",
	      accessKeys: [...document.querySelectorAll("#zoomDialog [data-access-key]")].map((item) => item.dataset.accessKey || ""),
	      accessKeysUnderlined: [...document.querySelectorAll("#zoomDialog .menu-access-key")].every((item) =>
	        getComputedStyle(item).textDecorationLine.includes("underline"),
	      ),
	    };
	  });


	  await page.keyboard.press("Alt+7");


	  await page.waitForFunction(() => document.querySelector('#zoomDialog input[name="zoomDialogOption"][value="75"]')?.checked === true, null, { timeout: 5000 });


	  await page.locator("#zoomDialog").press("Enter");


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.zoomScale === 75, null, { timeout: 5000 });


  const viewZoomApplied = await page.evaluate(() => ({
    zoom: window.__webSheetDevSheetModel?.("CHT")?.zoomScale,
    label: document.querySelector("#zoomLabel")?.textContent || "",
    activeSheetView: document.querySelector(".sheet-view-buttons [data-sheet-view].is-active")?.dataset.sheetView || "",
  }));


  await page.click('[data-ribbon-command="view-100"]');


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.zoomScale === 100, null, { timeout: 5000 });


  const viewZoomHundred = await page.evaluate(() => ({
    zoom: window.__webSheetDevSheetModel?.("CHT")?.zoomScale,
    active: document.querySelector('[data-ribbon-command="view-100"]')?.classList.contains("is-active") || false,
    title: document.querySelector('[data-ribbon-command="view-100"]')?.getAttribute("title") || "",
  }));


  const viewZoomSelectionRangeLabel = await page.evaluate(() => {
    const sheet = window.__webSheetDevSheetModel?.("CHT") || {};
    const bottom = Math.max(40, Number(sheet.rowCount || 0));
    const right = Math.max(40, Number(sheet.colCount || 0));
    return window.__webSheetDevSelectRange?.(1, 1, bottom, right) || "";
  });


  const viewZoomSelectionHook = await page.evaluate(() => window.__webSheetDevZoomSelectionToFit?.() || {
    result: false,
    missingHook: true,
    zoom: window.__webSheetDevSheetModel?.("CHT")?.zoomScale,
    status: document.querySelector("#statusText")?.textContent || "",
  });


  if (!Number.isFinite(Number(viewZoomSelectionHook.after)) || Number(viewZoomSelectionHook.after) === 100) {
    throw new Error(`View zoom selection should change zoom: ${JSON.stringify(viewZoomSelectionHook)}`);
  }


  const viewZoomSelectionSheetName = viewZoomSelectionHook.sheetName || "CHT";


  const viewZoomSelection = await page.evaluate(({ rangeLabel, zoom }) => ({
    rangeLabel,
    zoom,
    label: document.querySelector("#zoomLabel")?.textContent || "",
    selectedLabel: document.querySelector("#selectedCellLabel")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
    activeHundred: document.querySelector('[data-ribbon-command="view-100"]')?.classList.contains("is-active") || false,
  }), { rangeLabel: viewZoomSelectionRangeLabel, zoom: Number(viewZoomSelectionHook.after) });


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(
    (sheetName) => window.__webSheetDevSheetModel?.(sheetName)?.zoomScale === 100,
    viewZoomSelectionSheetName,
    { timeout: 5000 },
  );


  const viewZoomSelectionUndoRedo = {
    undo: await page.evaluate((sheetName) => ({
      zoom: window.__webSheetDevSheetModel?.(sheetName)?.zoomScale,
      label: document.querySelector("#zoomLabel")?.textContent || "",
      activeHundred: document.querySelector('[data-ribbon-command="view-100"]')?.classList.contains("is-active") || false,
      status: document.querySelector("#statusText")?.textContent || "",
    }), viewZoomSelectionSheetName),
  };


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(
    ({ sheetName, zoom }) => window.__webSheetDevSheetModel?.(sheetName)?.zoomScale === zoom,
    { sheetName: viewZoomSelectionSheetName, zoom: viewZoomSelection.zoom },
    { timeout: 5000 },
  );


  viewZoomSelectionUndoRedo.redo = await page.evaluate((sheetName) => ({
    zoom: window.__webSheetDevSheetModel?.(sheetName)?.zoomScale,
    label: document.querySelector("#zoomLabel")?.textContent || "",
    activeHundred: document.querySelector('[data-ribbon-command="view-100"]')?.classList.contains("is-active") || false,
    status: document.querySelector("#statusText")?.textContent || "",
  }), viewZoomSelectionSheetName);


  await page.click('[data-ribbon-command="view-100"]');


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.zoomScale === 100, null, { timeout: 5000 });


  await page.evaluate((zoom) => {
    const slider = document.querySelector("#zoomSlider");
    if (!slider) return;
    slider.value = String(zoom);
    slider.dispatchEvent(new Event("input", { bubbles: true }));
  }, viewZoomBefore);


  await page.waitForFunction((zoom) => window.__webSheetDevSheetModel?.("CHT")?.zoomScale === zoom, viewZoomBefore, { timeout: 5000 });


  await page.click('[data-ribbon-command="view-page-layout"]');


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.viewMode === "page-layout", null, { timeout: 5000 });


  await page.click('[data-ribbon-command="view-ruler"]');


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showRuler === true, null, { timeout: 5000 });


  const viewRulerOn = await page.evaluate(() => ({
    showRuler: window.__webSheetDevSheetModel?.("CHT")?.showRuler,
    viewMode: window.__webSheetDevSheetModel?.("CHT")?.viewMode,
    gridHasRuler: document.querySelector(".sheet-grid")?.classList.contains("show-ruler"),
    rulerButtonActive: document.querySelector('[data-ribbon-command="view-ruler"]')?.classList.contains("is-active"),
    horizontalRulerTicks: document.querySelectorAll(".sheet-ruler-horizontal .sheet-ruler-tick").length,
    verticalRulerTicks: document.querySelectorAll(".sheet-ruler-vertical .sheet-ruler-tick").length,
    majorRulerLabels: [...document.querySelectorAll(".sheet-ruler-tick.is-major span")].map((node) => node.textContent.trim()).filter(Boolean).slice(0, 6),
    headerFooterPlaceholders: [...document.querySelectorAll(".sheet-page-layout-header-footer-content.is-empty")].map((node) => node.textContent.replace(/\s+/g, " ").trim()),
    headerPointerEvents: getComputedStyle(document.querySelector(".sheet-page-layout-header")).pointerEvents,
  }));


  await page.dblclick(".sheet-page-layout-header");


  await page.waitForSelector('[data-data-dialog="header-footer"]', { timeout: 5000 });


  const pageLayoutHeaderFooterDialog = await page.evaluate(() => ({
    title: document.querySelector("#headerFooterTitle")?.textContent || "",
    firstFieldFocused: document.activeElement?.getAttribute("data-header-footer-field") || "",
  }));


  await page.click('[data-data-dialog="header-footer"] [data-data-dialog-action="cancel"]');


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


  await page.click('[data-ribbon-command="view-ruler"]');


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showRuler !== true, null, { timeout: 5000 });


  await page.evaluate(() => {
    const sheet = window.__webSheetDevSheetModel?.("CHT");
    if (!sheet) return;
    sheet.pageLayout = { ...(sheet.pageLayout || {}), pageBreaks: { rows: [], cols: [] } };
  });


  await page.click('[data-ribbon-command="view-page-break"]');


  await page.waitForFunction(() => document.querySelector(".sheet-grid")?.classList.contains("view-page-break"), null, { timeout: 5000 });


  const pageBreakPreview = await page.evaluate(() => ({
    viewMode: window.__webSheetDevSheetModel?.("CHT")?.viewMode,
    guideCount: document.querySelectorAll(".sheet-view-guide.is-page-break").length,
    label: document.querySelector(".sheet-page-break-label")?.textContent || "",
  }));


  await page.waitForSelector(".sheet-view-guide.is-page-break.is-automatic", { timeout: 5000 });


  const pageBreakContextTarget = await page.evaluate(() => {
    const guide = document.querySelector(".sheet-view-guide.is-page-break.is-automatic");
    const rect = guide?.getBoundingClientRect();
    return {
      axis: guide?.dataset.pageBreakAxis || "",
      index: Number(guide?.dataset.pageBreakIndex || 0),
      x: rect ? rect.left + Math.max(2, Math.min(6, rect.width / 2 || 2)) : 0,
      y: rect ? rect.top + Math.max(2, Math.min(12, rect.height / 2 || 2)) : 0,
    };
  });


  await page.mouse.click(pageBreakContextTarget.x, pageBreakContextTarget.y, { button: "right" });


  await page.waitForSelector("#viewCommandMenu.is-page-break", { timeout: 5000 });


  const pageBreakContextInitial = await page.evaluate(() => ({
    labels: [...document.querySelectorAll("#viewCommandMenu [data-view-menu-action]")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
    actions: [...document.querySelectorAll("#viewCommandMenu [data-view-menu-action]")].map((button) => button.dataset.viewMenuAction || ""),
    disabledActions: [...document.querySelectorAll("#viewCommandMenu [data-view-menu-action]:disabled")].map((button) => button.dataset.viewMenuAction || ""),
    accessKeys: [...document.querySelectorAll("#viewCommandMenu [data-access-key]")].map((button) => button.dataset.accessKey || ""),
    descriptions: [...document.querySelectorAll("#viewCommandMenu .home-command-menu-text small")].map((node) => node.textContent.trim()),
    accessKeysUnderlined: [...document.querySelectorAll("#viewCommandMenu .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
  }));


  await page.keyboard.press("I");


  await page.waitForFunction(({ axis, index }) => {
    const breaks = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {};
    const list = axis === "row" ? breaks.rows || [] : breaks.cols || [];
    return list.includes(index);
  }, pageBreakContextTarget, { timeout: 5000 });


  const pageBreakContextInserted = await page.evaluate(({ axis, index }) => {
    const breaks = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {};
    const list = axis === "row" ? breaks.rows || [] : breaks.cols || [];
    return {
      breaks,
      inserted: list.includes(index),
      status: document.querySelector("#statusText")?.textContent || "",
      manualGuideCount: document.querySelectorAll(".sheet-view-guide.is-page-break.is-manual").length,
    };
  }, pageBreakContextTarget);


  await page.locator(`.sheet-view-guide.is-page-break.is-manual[data-page-break-axis="${pageBreakContextTarget.axis}"][data-page-break-index="${pageBreakContextTarget.index}"]`).first().click({ button: "right" });


  await page.waitForSelector("#viewCommandMenu.is-page-break", { timeout: 5000 });


  const pageBreakContextManual = await page.evaluate(() => ({
    disabledActions: [...document.querySelectorAll("#viewCommandMenu [data-view-menu-action]:disabled")].map((button) => button.dataset.viewMenuAction || ""),
    checkedRemoveEnabled: !document.querySelector('#viewCommandMenu [data-view-menu-action^="page-break:remove:"]')?.disabled,
  }));


  await page.keyboard.press("R");


  try {
    await page.waitForFunction(({ axis, index }) => {
      const breaks = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {};
      const list = axis === "row" ? breaks.rows || [] : breaks.cols || [];
      return !list.includes(index);
    }, pageBreakContextTarget, { timeout: 10000 });
  } catch (error) {
    const debug = await page.evaluate(({ axis, index }) => {
      const breaks = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {};
      const list = axis === "row" ? breaks.rows || [] : breaks.cols || [];
      return {
        breaks,
        stillPresent: list.includes(index),
        activeElement: document.activeElement?.outerHTML?.slice(0, 240) || "",
        menuExists: Boolean(document.querySelector("#viewCommandMenu")),
        menuActions: [...document.querySelectorAll("#viewCommandMenu [data-view-menu-action]")].map((button) => ({
          action: button.dataset.viewMenuAction || "",
          disabled: button.disabled,
          accessKey: button.dataset.accessKey || "",
          text: button.textContent.replace(/\s+/g, " ").trim(),
        })),
        status: document.querySelector("#statusText")?.textContent || "",
      };
    }, pageBreakContextTarget);
    throw new Error(`Page break remove access key failed: ${JSON.stringify({ pageBreakContextTarget, pageBreakContextManual, debug })}`);
  }


  const pageBreakContextRemoved = await page.evaluate(({ axis, index }) => {
    const breaks = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {};
    const list = axis === "row" ? breaks.rows || [] : breaks.cols || [];
    return {
      breaks,
      removed: !list.includes(index),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, pageBreakContextTarget);


  await page.waitForFunction(() => document.activeElement?.id === "sheetHost", null, { timeout: 5000 });


  await page.keyboard.press("Control+Z");


  await page.waitForFunction(({ axis, index }) => {
    const breaks = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {};
    const list = axis === "row" ? breaks.rows || [] : breaks.cols || [];
    return list.includes(index);
  }, pageBreakContextTarget, { timeout: 5000 });


  const pageBreakContextUndo = await page.evaluate(() => ({
    pageBreaks: window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {},
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.keyboard.press("Control+Y");


  await page.waitForFunction(({ axis, index }) => {
    const breaks = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {};
    const list = axis === "row" ? breaks.rows || [] : breaks.cols || [];
    return !list.includes(index);
  }, pageBreakContextTarget, { timeout: 5000 });


  const pageBreakContextRedo = await page.evaluate(() => ({
    pageBreaks: window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {},
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await page.waitForSelector(".sheet-view-guide.is-page-break.is-automatic", { timeout: 5000 });


  await page.locator(".sheet-view-guide.is-page-break.is-automatic").first().click({ button: "right", force: true });


  await page.waitForSelector("#viewCommandMenu.is-page-break", { timeout: 5000 });


  await page.keyboard.press("V");


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.viewMode === "normal", null, { timeout: 5000 });


  const pageBreakContextMenu = {
    target: pageBreakContextTarget,
    initial: pageBreakContextInitial,
    inserted: pageBreakContextInserted,
    manual: pageBreakContextManual,
    removed: pageBreakContextRemoved,
    undo: pageBreakContextUndo,
    redo: pageBreakContextRedo,
    closePreview: await page.evaluate(() => ({
      viewMode: window.__webSheetDevSheetModel?.("CHT")?.viewMode,
      activeNormal: document.querySelector('[data-ribbon-command="view-normal"]')?.classList.contains("is-active"),
      status: document.querySelector("#statusText")?.textContent || "",
    })),
  };


  await page.click('[data-ribbon-command="view-page-layout"]');


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.viewMode === "page-layout", null, { timeout: 5000 });


  const customViewCaptureState = await page.evaluate(async () => {
    window.__webSheetDevSelectRange?.(30, 30, 31, 31);
    const host = document.querySelector("#sheetHost");
    if (host) {
      host.scrollTop = 240;
      host.scrollLeft = 180;
      host.dispatchEvent(new Event("scroll", { bubbles: true }));
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    }
    return {
      selectionLabel: document.querySelector("#selectedCellLabel")?.textContent?.trim() || "",
      scrollTop: host?.scrollTop || 0,
      scrollLeft: host?.scrollLeft || 0,
      viewMode: window.__webSheetDevSheetModel?.("CHT")?.viewMode,
    };
  });


  await page.click('[data-ribbon-command="view-custom-views"]');


  await page.waitForSelector('[data-data-dialog="custom-views"]', { timeout: 5000 });


  await page.waitForSelector("[data-data-dialog='custom-views'] [data-access-key='A']", { timeout: 5000 });


  const customViewsMainBeforeAdd = await page.evaluate(() => ({
	    title: document.querySelector("#customViewsTitle")?.textContent || "",
	    hasInlineNameField: Boolean(document.querySelector('[data-data-dialog="custom-views"] [data-data-field="customViewName"]')),
	    addButtonText: document.querySelector('[data-data-dialog-action="custom-view-add"]')?.textContent?.trim() || "",
	    addAccessKey: document.querySelector('[data-data-dialog-action="custom-view-add"]')?.dataset.accessKey || "",
	    accessKeysUnderlined: [...document.querySelectorAll('[data-data-dialog="custom-views"] .menu-access-key')].every((item) =>
	      getComputedStyle(item).textDecorationLine.includes("underline"),
	    ),
  }));


  await page.keyboard.press("Alt+A");


  await page.waitForSelector('[data-data-dialog="custom-view-add"]', { timeout: 5000 });


  await page.waitForSelector("[data-data-dialog='custom-view-add'] [data-access-key='R']", { timeout: 5000 });


  await page.keyboard.press("Alt+R");


  await page.waitForFunction(() => document.querySelector("[data-data-field='customViewHiddenRowsFilters']")?.checked === false, null, { timeout: 5000 });


  await page.keyboard.press("Alt+R");


  await page.waitForFunction(() => document.querySelector("[data-data-field='customViewHiddenRowsFilters']")?.checked === true, null, { timeout: 5000 });


  const customViewAddDialog = await page.evaluate(() => ({
	    title: document.querySelector("#customViewAddTitle")?.textContent || "",
	    defaultName: document.querySelector("[data-data-field='customViewName']")?.value || "",
	    printChecked: document.querySelector("[data-data-field='customViewPrintSettings']")?.checked,
	    hiddenRowsChecked: document.querySelector("[data-data-field='customViewHiddenRowsFilters']")?.checked,
	    fieldsetLegend: document.querySelector(".custom-view-include-options legend")?.textContent || "",
	    checkboxAccessKeys: [...document.querySelectorAll("[data-data-dialog='custom-view-add'] [data-access-key]")].map((item) => item.dataset.accessKey || ""),
	    accessKeysUnderlined: [...document.querySelectorAll("[data-data-dialog='custom-view-add'] .menu-access-key")].every((item) =>
	      getComputedStyle(item).textDecorationLine.includes("underline"),
	    ),
  }));


  await page.locator("[data-data-field='customViewName']").fill("表示テスト");


  await page.locator("[data-data-dialog-action='custom-view-add-ok']").click();


  await page.waitForFunction(() => (window.__webSheetDevWorkbookState?.().customViews || []).some((view) => view.name === "表示テスト"), null, { timeout: 5000 });


  await page.waitForSelector('[data-data-dialog="custom-views"]', { timeout: 5000 });


	  const customViewsDialog = await page.evaluate(() => ({
	    names: (window.__webSheetDevWorkbookState?.().customViews || []).map((view) => view.name),
	    savedView: (window.__webSheetDevWorkbookState?.().customViews || []).find((view) => view.name === "表示テスト") || null,
	    dialogVisible: Boolean(document.querySelector('[data-data-dialog="custom-views"]')),
	    showDisabled: document.querySelector("[data-data-dialog-action='custom-view-show']")?.disabled,
	    selectedText: document.querySelector("[data-data-field='customViewIndex'] option:checked")?.textContent?.trim() || "",
	  }));


	  await page.keyboard.press("Control+Z");


	  await page.waitForFunction(() => !(window.__webSheetDevWorkbookState?.().customViews || []).some((view) => view.name === "表示テスト"), null, { timeout: 5000 });


	  const customViewSaveUndoRedo = {
	    undo: await page.evaluate(() => ({
	      names: (window.__webSheetDevWorkbookState?.().customViews || []).map((view) => view.name),
	      dialogNames: [...document.querySelectorAll("[data-data-field='customViewIndex'] option")].map((option) => option.textContent.trim()),
	      status: document.querySelector("#statusText")?.textContent || "",
	    })),
	  };


	  await page.keyboard.press("Control+Y");


	  await page.waitForFunction(() => (window.__webSheetDevWorkbookState?.().customViews || []).some((view) => view.name === "表示テスト"), null, { timeout: 5000 });


	  customViewSaveUndoRedo.redo = await page.evaluate(() => ({
	    names: (window.__webSheetDevWorkbookState?.().customViews || []).map((view) => view.name),
	    dialogNames: [...document.querySelectorAll("[data-data-field='customViewIndex'] option")].map((option) => option.textContent.trim()),
	    status: document.querySelector("#statusText")?.textContent || "",
	  }));


	  await page.locator("[data-data-field='customViewIndex']").selectOption({ label: "表示テスト" });


	  await page.evaluate(async () => {
	    window.__webSheetDevSelectRange?.(1, 1, 1, 1);
    const host = document.querySelector("#sheetHost");
    if (host) {
      host.scrollTop = 0;
      host.scrollLeft = 0;
      host.dispatchEvent(new Event("scroll", { bubbles: true }));
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    }
  });


  await page.locator("[data-data-dialog-action='custom-view-show']").click();


  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


	  const customViewApplied = await page.evaluate(async () => {
	    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
	    const host = document.querySelector("#sheetHost");
    return {
      selectionLabel: document.querySelector("#selectedCellLabel")?.textContent?.trim() || "",
      scrollTop: host?.scrollTop || 0,
      scrollLeft: host?.scrollLeft || 0,
      viewMode: window.__webSheetDevSheetModel?.("CHT")?.viewMode,
	      formulaBarVisible: window.__webSheetDevWorkbookState?.().formulaBarVisible,
	    };
	  });


	  await page.keyboard.press("Control+Z");


		  await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent?.trim() === "A1", null, { timeout: 5000 });


		  const customViewApplyUndoRedo = {
		    undo: await page.evaluate(async () => {
		      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
		      const host = document.querySelector("#sheetHost");
		      return {
		        selectionLabel: document.querySelector("#selectedCellLabel")?.textContent?.trim() || "",
		        scrollTop: host?.scrollTop || 0,
		        scrollLeft: host?.scrollLeft || 0,
		        status: document.querySelector("#statusText")?.textContent || "",
		      };
		    }),
		  };


	  await page.keyboard.press("Control+Y");


	  await page.waitForFunction(
	    (label) => document.querySelector("#selectedCellLabel")?.textContent?.trim() === label,
	    customViewCaptureState.selectionLabel,
	    { timeout: 5000 },
	  );


	  customViewApplyUndoRedo.redo = await page.evaluate(async () => {
	    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
	    const host = document.querySelector("#sheetHost");
	    return {
	      selectionLabel: document.querySelector("#selectedCellLabel")?.textContent?.trim() || "",
	      scrollTop: host?.scrollTop || 0,
	      scrollLeft: host?.scrollLeft || 0,
	      status: document.querySelector("#statusText")?.textContent || "",
	    };
	  });


	  await page.click('[data-ribbon-command="view-custom-views"]');


  await page.waitForSelector('[data-data-dialog="custom-views"]', { timeout: 5000 });


  await page.waitForSelector("[data-data-dialog='custom-views'] [data-access-key='A']", { timeout: 5000 });


  await page.locator("[data-data-dialog-action='custom-view-add']").click();


  await page.waitForSelector('[data-data-dialog="custom-view-add"]', { timeout: 5000 });


  await page.locator("[data-data-field='customViewName']").fill("削除テスト");


  await page.locator("[data-data-dialog-action='custom-view-add-ok']").click();


  await page.waitForFunction(() => (window.__webSheetDevWorkbookState?.().customViews || []).some((view) => view.name === "削除テスト"), null, { timeout: 5000 });


  await page.waitForSelector('[data-data-dialog="custom-views"]', { timeout: 5000 });


  await page.locator("[data-data-field='customViewIndex']").selectOption({ label: "削除テスト" });


	  const customViewDeleteBefore = await page.evaluate(() => ({
	    names: (window.__webSheetDevWorkbookState?.().customViews || []).map((view) => view.name),
	    selectedText: document.querySelector("[data-data-field='customViewIndex'] option:checked")?.textContent?.trim() || "",
	    deleteDisabled: document.querySelector("[data-data-dialog-action='custom-view-delete']")?.disabled,
	  }));


	  await page.locator("[data-data-field='customViewIndex']").press("Delete");


	  await page.waitForSelector("#messageBox", { timeout: 5000 });


	  const customViewDeletePrompt = await page.evaluate(() => ({
	    title: document.querySelector("#messageBoxTitle")?.textContent || "",
	    message: document.querySelector("#messageBoxText")?.textContent || "",
	    buttons: [...document.querySelectorAll("#messageBox [data-message-box-result]")].map((button) => button.textContent.trim()),
	    accessKeys: [...document.querySelectorAll("#messageBox [data-message-box-result]")].map((button) => button.dataset.accessKey || ""),
	  }));


	  await page.keyboard.press("D");


	  await page.waitForFunction(() => !(window.__webSheetDevWorkbookState?.().customViews || []).some((view) => view.name === "削除テスト"), null, { timeout: 5000 });


	  await page.waitForSelector('[data-data-dialog="custom-views"]', { timeout: 5000 });


		  const customViewDeleteAfter = await page.evaluate(() => ({
		    names: (window.__webSheetDevWorkbookState?.().customViews || []).map((view) => view.name),
		    dialogNames: [...document.querySelectorAll("[data-data-field='customViewIndex'] option")].map((option) => option.textContent.trim()),
		    dialogVisible: Boolean(document.querySelector('[data-data-dialog="custom-views"]')),
		    status: document.querySelector("#statusText")?.textContent || "",
		  }));


		  await page.keyboard.press("Control+Z");


		  await page.waitForFunction(() => (window.__webSheetDevWorkbookState?.().customViews || []).some((view) => view.name === "削除テスト"), null, { timeout: 5000 });


		  const customViewDeleteUndoRedo = {
		    undo: await page.evaluate(() => ({
		      names: (window.__webSheetDevWorkbookState?.().customViews || []).map((view) => view.name),
		      dialogNames: [...document.querySelectorAll("[data-data-field='customViewIndex'] option")].map((option) => option.textContent.trim()),
		      status: document.querySelector("#statusText")?.textContent || "",
		    })),
		  };


		  await page.keyboard.press("Control+Y");


		  await page.waitForFunction(() => !(window.__webSheetDevWorkbookState?.().customViews || []).some((view) => view.name === "削除テスト"), null, { timeout: 5000 });


		  customViewDeleteUndoRedo.redo = await page.evaluate(() => ({
		    names: (window.__webSheetDevWorkbookState?.().customViews || []).map((view) => view.name),
		    dialogNames: [...document.querySelectorAll("[data-data-field='customViewIndex'] option")].map((option) => option.textContent.trim()),
		    status: document.querySelector("#statusText")?.textContent || "",
		  }));


		  const customViewDeleteFlow = {
		    before: customViewDeleteBefore,
		    prompt: customViewDeletePrompt,
		    after: customViewDeleteAfter,
		    undoRedo: customViewDeleteUndoRedo,
		  };


	  await page.locator("[data-data-dialog-action='cancel']").click();


	  await page.waitForSelector("#dataDialog", { state: "detached", timeout: 5000 });


	  await page.click('[data-ribbon-command="view-freeze-menu"]');


  await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });


  const freezeMenuLabels = await page.locator("#viewCommandMenu").innerText();


  await page.keyboard.press("ArrowDown");


  const freezeMenuInitialMeta = await page.evaluate(() => ({
    focusedAfterArrowDown: document.activeElement?.dataset?.viewMenuAction || "",
    actionCount: document.querySelectorAll("#viewCommandMenu [data-view-menu-action]").length,
    descriptionCount: document.querySelectorAll("#viewCommandMenu .home-command-menu-text small").length,
    descriptions: [...document.querySelectorAll("#viewCommandMenu .home-command-menu-text small")].map((node) => node.textContent.trim()),
    width: Math.round(document.querySelector("#viewCommandMenu")?.getBoundingClientRect().width || 0),
  }));


  await page.keyboard.press("ArrowRight");


  Object.assign(freezeMenuInitialMeta, await page.evaluate(() => ({
    focusedAfterArrowRight: document.activeElement?.dataset?.viewMenuAction || "",
  })));


  await page.click('[data-view-menu-action="freeze-top-row"]');


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.freezePanes?.rows === 1, null, { timeout: 5000 });


  const freezeMenuTopRow = await page.evaluate(() => window.__webSheetDevSheetModel?.("CHT")?.freezePanes || null);


  await page.click('[data-ribbon-command="view-freeze-menu"]');


  await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });


  const freezeMenuTopRowStatus = await page.evaluate(() => ({
    checkedAction: document.querySelector("#viewCommandMenu [aria-checked='true']")?.dataset.viewMenuAction || "",
    topRowChecked: document.querySelector('#viewCommandMenu [data-view-menu-action="freeze-top-row"]')?.getAttribute("aria-checked") || "",
    selectionChecked: document.querySelector('#viewCommandMenu [data-view-menu-action="freeze-selection"]')?.getAttribute("aria-checked") || "",
    checkMarkCount: [...document.querySelectorAll("#viewCommandMenu .home-command-menu-check")].filter((node) => node.textContent.trim()).length,
    mainButtonLabel: document.querySelector('[data-ribbon-command="view-freeze-panes"] span:not(.ribbon-check-box):not(.ribbon-command-caret)')?.textContent.trim() || "",
    mainButtonIcon: document.querySelector('[data-ribbon-command="view-freeze-panes"] [data-lucide]')?.getAttribute("data-lucide") || "",
    unfreezeFirst: document.querySelector("#viewCommandMenu [data-view-menu-action]")?.dataset.viewMenuAction === "unfreeze",
  }));


  await page.click('[data-view-menu-action="unfreeze"]');


  await page.waitForFunction(() => !window.__webSheetDevSheetModel?.("CHT")?.freezePanes, null, { timeout: 5000 });


  await page.click('[data-ribbon-command="view-freeze-menu"]');


  await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });


  await page.keyboard.press("Alt+C");


  await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.freezePanes?.columns === 1, null, { timeout: 5000 });


  const freezeMenuAccessKey = await page.evaluate(() => ({
    pane: window.__webSheetDevSheetModel?.("CHT")?.freezePanes || null,
    status: document.querySelector("#statusText")?.textContent || "",
    menuClosed: !document.querySelector("#viewCommandMenu"),
    mainButtonLabel: document.querySelector('[data-ribbon-command="view-freeze-panes"] span:not(.ribbon-check-box):not(.ribbon-command-caret)')?.textContent.trim() || "",
    mainButtonTitle: document.querySelector('[data-ribbon-command="view-freeze-panes"]')?.getAttribute("title") || "",
  }));


  await page.click('[data-ribbon-command="view-freeze-menu"]');


  await page.waitForSelector("#viewCommandMenu", { timeout: 5000 });


  await page.click('[data-view-menu-action="unfreeze"]');


  await page.waitForFunction(() => !window.__webSheetDevSheetModel?.("CHT")?.freezePanes, null, { timeout: 5000 });


  const viewInitialState = await page.evaluate(() => {
    window.__webSheetDevSetCell?.(2, 52, 0);
    return {
      showGridLines: window.__webSheetDevSheetModel?.("CHT")?.showGridLines,
      showHeadings: window.__webSheetDevSheetModel?.("CHT")?.showHeadings,
      showZeros: window.__webSheetDevSheetModel?.("CHT")?.showZeros,
      showObjects: window.__webSheetDevSheetModel?.("CHT")?.showObjects,
      showOutlineSymbols: window.__webSheetDevSheetModel?.("CHT")?.showOutlineSymbols,
      rightToLeft: window.__webSheetDevSheetModel?.("CHT")?.rightToLeft,
      formulaBarVisible: window.__webSheetDevWorkbookState?.().formulaBarVisible,
    };
  });


  const toggledGridlinesForViewTest = viewInitialState.showGridLines !== false;


  const toggledHeadingsForViewTest = viewInitialState.showHeadings !== false;


  const toggledZeroValuesForViewTest = viewInitialState.showZeros !== false;


  const toggledObjectsForViewTest = viewInitialState.showObjects !== false;


  const toggledOutlineForViewTest = viewInitialState.showOutlineSymbols !== false;


  const toggledRightToLeftForViewTest = viewInitialState.rightToLeft !== true;


  const toggledFormulaBarForViewTest = viewInitialState.formulaBarVisible !== false;


  let formulaBarUndoRedo = null;


  if (toggledGridlinesForViewTest) {
    await page.click('[data-ribbon-command="view-gridlines"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showGridLines === false, null, { timeout: 5000 });
  }


  if (toggledHeadingsForViewTest) {
    await page.click('[data-ribbon-command="view-headings"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showHeadings === false, null, { timeout: 5000 });
  }


  if (toggledFormulaBarForViewTest) {
    await page.click('[data-ribbon-command="view-formula-bar"]');
    await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().formulaBarVisible === false, null, { timeout: 5000 });
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().formulaBarVisible === true, null, { timeout: 5000 });
    const undo = await page.evaluate(() => ({
      formulaBarVisible: window.__webSheetDevWorkbookState?.().formulaBarVisible,
      commandActive: document.querySelector('[data-ribbon-command="view-formula-bar"]')?.classList.contains("is-active"),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().formulaBarVisible === false, null, { timeout: 5000 });
    const redo = await page.evaluate(() => ({
      formulaBarVisible: window.__webSheetDevWorkbookState?.().formulaBarVisible,
      commandActive: document.querySelector('[data-ribbon-command="view-formula-bar"]')?.classList.contains("is-active"),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    formulaBarUndoRedo = { undo, redo };
  }


  if (toggledZeroValuesForViewTest) {
    await page.click('[data-ribbon-command="view-zero-values"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showZeros === false, null, { timeout: 5000 });
  }


  if (toggledObjectsForViewTest) {
    await page.click('[data-ribbon-command="view-objects"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showObjects === false, null, { timeout: 5000 });
  }


  if (toggledOutlineForViewTest) {
    await page.click('[data-ribbon-command="view-outline-symbols"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showOutlineSymbols === false, null, { timeout: 5000 });
  }


  if (toggledRightToLeftForViewTest) {
    await page.click('[data-ribbon-command="view-right-to-left"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.rightToLeft === true, null, { timeout: 5000 });
  }


	  const viewToggles = await page.evaluate(() => ({
	    showGridLines: window.__webSheetDevSheetModel?.("CHT")?.showGridLines,
	    showHeadings: window.__webSheetDevSheetModel?.("CHT")?.showHeadings,
	    showZeros: window.__webSheetDevSheetModel?.("CHT")?.showZeros,
	    showObjects: window.__webSheetDevSheetModel?.("CHT")?.showObjects,
	    showOutlineSymbols: window.__webSheetDevSheetModel?.("CHT")?.showOutlineSymbols,
	    rightToLeft: window.__webSheetDevSheetModel?.("CHT")?.rightToLeft,
    gridHasHideGridlines: document.querySelector(".sheet-grid")?.classList.contains("hide-gridlines"),
    gridHasHideHeadings: document.querySelector(".sheet-grid")?.classList.contains("hide-headings"),
    gridHasHideOutlineSymbols: document.querySelector(".sheet-grid")?.classList.contains("hide-outline-symbols"),
    gridHasRightToLeft: document.querySelector(".sheet-grid")?.classList.contains("is-right-to-left"),
    zeroCellText: document.querySelector('.sheet-cell[data-address="AZ2"] .cell-content')?.textContent ?? "",
    visibleObjectLayerItems: document.querySelectorAll(".sheet-image").length,
    formulaBarVisible: window.__webSheetDevWorkbookState?.().formulaBarVisible,
    activeCommands: [...document.querySelectorAll('[data-ribbon-panel="view"] [data-ribbon-command].is-active')].map((button) => button.dataset.ribbonCommand),
    showChecks: [...document.querySelectorAll(".ribbon-view-show-group [data-ribbon-command]")].map((button) => ({
      command: button.dataset.ribbonCommand,
      active: button.classList.contains("is-active"),
      role: button.getAttribute("role"),
      ariaChecked: button.getAttribute("aria-checked"),
	      hasCheckBox: Boolean(button.querySelector(".ribbon-check-box")),
	    })),
	  }));


  const rtlVisualTarget = await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    const hostRect = host?.getBoundingClientRect();
    const rowHeaderRect = document.querySelector(".row-header")?.getBoundingClientRect();
    const rowHeaderRight = rowHeaderRect?.right ?? hostRect?.left ?? 0;
    const cells = [...document.querySelectorAll(".sheet-cell[data-row][data-col]")]
      .map((cell) => {
        const rect = cell.getBoundingClientRect();
        const visibleLeft = Math.max(rect.left, rowHeaderRight, hostRect?.left ?? rect.left);
        const visibleRight = Math.min(rect.right, hostRect?.right ?? rect.right);
        const visibleTop = Math.max(rect.top, hostRect?.top ?? rect.top);
        const visibleBottom = Math.min(rect.bottom, hostRect?.bottom ?? rect.bottom);
        return {
          row: Number(cell.dataset.row),
          col: Number(cell.dataset.col),
          address: cell.dataset.address || "",
          left: rect.left,
          right: rect.right,
          top: rect.top,
          bottom: rect.bottom,
          width: rect.width,
          height: rect.height,
          visibleLeft,
          visibleRight,
          visibleTop,
          visibleBottom,
          visibleWidth: Math.max(0, visibleRight - visibleLeft),
          visibleHeight: Math.max(0, visibleBottom - visibleTop),
        };
      })
      .filter((cell) =>
        Number.isFinite(cell.row) &&
        Number.isFinite(cell.col) &&
        cell.right > rowHeaderRight + 1 &&
        cell.width > 0 &&
        cell.height > 0 &&
        cell.visibleWidth > 2 &&
        cell.visibleHeight > 2 &&
        (!hostRect || (cell.right > hostRect.left && cell.left < hostRect.right && cell.bottom > hostRect.top && cell.top < hostRect.bottom)),
      )
      .sort((a, b) => a.left - b.left || a.top - b.top);
    const first = cells[0] || null;
    const highestRendered = cells.reduce((max, cell) => Math.max(max, cell.col), 0);
    return {
      firstVisualCol: first?.col || 0,
      firstVisualAddress: first?.address || "",
      highestRendered,
      visibleCellCount: cells.length,
      x: first ? first.visibleLeft + Math.min(10, Math.max(2, first.visibleWidth / 2)) : null,
      y: first ? first.visibleTop + Math.min(10, Math.max(2, first.visibleHeight / 2)) : null,
    };
  });


  if (rtlVisualTarget?.x != null && rtlVisualTarget?.y != null) {
    await page.mouse.click(rtlVisualTarget.x, rtlVisualTarget.y);
  }


  const rtlVisualHitTest = await page.evaluate((target) => ({
    ...target,
    selectedCol: window.__webSheetDevWorkbookState?.().selectedCell?.col || 0,
    selectedRow: window.__webSheetDevWorkbookState?.().selectedCell?.row || 0,
    selectionLabel: document.querySelector("#selectedCellLabel")?.textContent || "",
  }), rtlVisualTarget);


	  let gridlinesUndoRedo = null;


	  let headingsUndoRedo = null;


	  let zeroValuesUndoRedo = null;


	  let objectsUndoRedo = null;


	  let outlineUndoRedo = null;


	  let rightToLeftUndoRedo = null;


	  if (!formulaBarUndoRedo && toggledFormulaBarForViewTest && viewToggles.formulaBarVisible === false) {
	    await page.keyboard.press("Control+Z");
	    await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().formulaBarVisible === true, null, { timeout: 5000 });
	    const undo = await page.evaluate(() => ({
	      formulaBarVisible: window.__webSheetDevWorkbookState?.().formulaBarVisible,
	      commandActive: document.querySelector('[data-ribbon-command="view-formula-bar"]')?.classList.contains("is-active"),
	      status: document.querySelector("#statusText")?.textContent || "",
	    }));
	    await page.keyboard.press("Control+Y");
	    await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().formulaBarVisible === false, null, { timeout: 5000 });
	    const redo = await page.evaluate(() => ({
	      formulaBarVisible: window.__webSheetDevWorkbookState?.().formulaBarVisible,
	      commandActive: document.querySelector('[data-ribbon-command="view-formula-bar"]')?.classList.contains("is-active"),
	      status: document.querySelector("#statusText")?.textContent || "",
	    }));
	    formulaBarUndoRedo = { undo, redo };
	  }


	  if (toggledFormulaBarForViewTest && viewToggles.formulaBarVisible === false) {
	    await page.click('[data-ribbon-command="view-formula-bar"]');
	    await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().formulaBarVisible === true, null, { timeout: 5000 });
	  }


  if (toggledZeroValuesForViewTest && viewToggles.showZeros === false) {
    await page.click('[data-ribbon-command="view-zero-values"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showZeros !== false, null, { timeout: 5000 });
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showZeros === false, null, { timeout: 5000 });
    const undo = await page.evaluate(() => ({
      showZeros: window.__webSheetDevSheetModel?.("CHT")?.showZeros !== false,
      commandActive: document.querySelector('[data-ribbon-command="view-zero-values"]')?.classList.contains("is-active"),
      ariaChecked: document.querySelector('[data-ribbon-command="view-zero-values"]')?.getAttribute("aria-checked") || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showZeros !== false, null, { timeout: 5000 });
    const redo = await page.evaluate(() => ({
      showZeros: window.__webSheetDevSheetModel?.("CHT")?.showZeros !== false,
      commandActive: document.querySelector('[data-ribbon-command="view-zero-values"]')?.classList.contains("is-active"),
      ariaChecked: document.querySelector('[data-ribbon-command="view-zero-values"]')?.getAttribute("aria-checked") || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    zeroValuesUndoRedo = { undo, redo };
    await page.click('[data-ribbon-command="view-zero-values"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showZeros === false, null, { timeout: 5000 });
  }


  if (toggledObjectsForViewTest && viewToggles.showObjects === false) {
    await page.click('[data-ribbon-command="view-objects"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showObjects !== false, null, { timeout: 5000 });
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showObjects === false, null, { timeout: 5000 });
    const undo = await page.evaluate(() => ({
      showObjects: window.__webSheetDevSheetModel?.("CHT")?.showObjects !== false,
      commandActive: document.querySelector('[data-ribbon-command="view-objects"]')?.classList.contains("is-active"),
      ariaChecked: document.querySelector('[data-ribbon-command="view-objects"]')?.getAttribute("aria-checked") || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showObjects !== false, null, { timeout: 5000 });
    const redo = await page.evaluate(() => ({
      showObjects: window.__webSheetDevSheetModel?.("CHT")?.showObjects !== false,
      commandActive: document.querySelector('[data-ribbon-command="view-objects"]')?.classList.contains("is-active"),
      ariaChecked: document.querySelector('[data-ribbon-command="view-objects"]')?.getAttribute("aria-checked") || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    objectsUndoRedo = { undo, redo };
  }


  if (toggledOutlineForViewTest && viewToggles.showOutlineSymbols === false) {
    await page.click('[data-ribbon-command="view-outline-symbols"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showOutlineSymbols !== false, null, { timeout: 5000 });
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showOutlineSymbols === false, null, { timeout: 5000 });
    const undo = await page.evaluate(() => ({
      showOutlineSymbols: window.__webSheetDevSheetModel?.("CHT")?.showOutlineSymbols !== false,
      gridHasHideOutlineSymbols: document.querySelector(".sheet-grid")?.classList.contains("hide-outline-symbols"),
      commandActive: document.querySelector('[data-ribbon-command="view-outline-symbols"]')?.classList.contains("is-active"),
      ariaChecked: document.querySelector('[data-ribbon-command="view-outline-symbols"]')?.getAttribute("aria-checked") || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showOutlineSymbols !== false, null, { timeout: 5000 });
    const redo = await page.evaluate(() => ({
      showOutlineSymbols: window.__webSheetDevSheetModel?.("CHT")?.showOutlineSymbols !== false,
      gridHasHideOutlineSymbols: document.querySelector(".sheet-grid")?.classList.contains("hide-outline-symbols"),
      commandActive: document.querySelector('[data-ribbon-command="view-outline-symbols"]')?.classList.contains("is-active"),
      ariaChecked: document.querySelector('[data-ribbon-command="view-outline-symbols"]')?.getAttribute("aria-checked") || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    outlineUndoRedo = { undo, redo };
  }


  if (viewToggles.showHeadings === false) {
    await page.click('[data-ribbon-command="view-headings"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showHeadings !== false, null, { timeout: 5000 });
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showHeadings === false, null, { timeout: 5000 });
    const undo = await page.evaluate(() => ({
      showHeadings: window.__webSheetDevSheetModel?.("CHT")?.showHeadings !== false,
      gridHasHideHeadings: document.querySelector(".sheet-grid")?.classList.contains("hide-headings"),
      commandActive: document.querySelector('[data-ribbon-command="view-headings"]')?.classList.contains("is-active"),
      ariaChecked: document.querySelector('[data-ribbon-command="view-headings"]')?.getAttribute("aria-checked") || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showHeadings !== false, null, { timeout: 5000 });
    const redo = await page.evaluate(() => ({
      showHeadings: window.__webSheetDevSheetModel?.("CHT")?.showHeadings !== false,
      gridHasHideHeadings: document.querySelector(".sheet-grid")?.classList.contains("hide-headings"),
      commandActive: document.querySelector('[data-ribbon-command="view-headings"]')?.classList.contains("is-active"),
      ariaChecked: document.querySelector('[data-ribbon-command="view-headings"]')?.getAttribute("aria-checked") || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    headingsUndoRedo = { undo, redo };
    if (viewInitialState.showHeadings === false) {
      await page.click('[data-ribbon-command="view-headings"]');
      await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showHeadings === false, null, { timeout: 5000 });
    }
  }


  if (viewToggles.showGridLines === false) {
    await page.click('[data-ribbon-command="view-gridlines"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showGridLines !== false, null, { timeout: 5000 });
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showGridLines === false, null, { timeout: 5000 });
    const undo = await page.evaluate(() => ({
      showGridLines: window.__webSheetDevSheetModel?.("CHT")?.showGridLines !== false,
      gridHasHideGridlines: document.querySelector(".sheet-grid")?.classList.contains("hide-gridlines"),
      commandActive: document.querySelector('[data-ribbon-command="view-gridlines"]')?.classList.contains("is-active"),
      ariaChecked: document.querySelector('[data-ribbon-command="view-gridlines"]')?.getAttribute("aria-checked") || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showGridLines !== false, null, { timeout: 5000 });
    const redo = await page.evaluate(() => ({
      showGridLines: window.__webSheetDevSheetModel?.("CHT")?.showGridLines !== false,
      gridHasHideGridlines: document.querySelector(".sheet-grid")?.classList.contains("hide-gridlines"),
      commandActive: document.querySelector('[data-ribbon-command="view-gridlines"]')?.classList.contains("is-active"),
      ariaChecked: document.querySelector('[data-ribbon-command="view-gridlines"]')?.getAttribute("aria-checked") || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    gridlinesUndoRedo = { undo, redo };
    if (viewInitialState.showGridLines === false) {
      await page.click('[data-ribbon-command="view-gridlines"]');
      await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.showGridLines === false, null, { timeout: 5000 });
    }
  }


  if (toggledRightToLeftForViewTest && viewToggles.rightToLeft === true) {
    await page.click('[data-ribbon-command="view-right-to-left"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.rightToLeft !== true, null, { timeout: 5000 });
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.rightToLeft === true, null, { timeout: 5000 });
    const undo = await page.evaluate(() => ({
      rightToLeft: window.__webSheetDevSheetModel?.("CHT")?.rightToLeft === true,
      gridHasRightToLeft: document.querySelector(".sheet-grid")?.classList.contains("is-right-to-left"),
      commandActive: document.querySelector('[data-ribbon-command="view-right-to-left"]')?.classList.contains("is-active"),
      ariaChecked: document.querySelector('[data-ribbon-command="view-right-to-left"]')?.getAttribute("aria-checked") || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.rightToLeft !== true, null, { timeout: 5000 });
    const redo = await page.evaluate(() => ({
      rightToLeft: window.__webSheetDevSheetModel?.("CHT")?.rightToLeft === true,
      gridHasRightToLeft: document.querySelector(".sheet-grid")?.classList.contains("is-right-to-left"),
      commandActive: document.querySelector('[data-ribbon-command="view-right-to-left"]')?.classList.contains("is-active"),
      ariaChecked: document.querySelector('[data-ribbon-command="view-right-to-left"]')?.getAttribute("aria-checked") || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    rightToLeftUndoRedo = { undo, redo };
  }
  Object.assign(scenarioState, { page, result, sheetGroupQuickAnalysis, sheetGroupSparklines, sheetGroupCharts, sheetGroupTables, sheetGroupPivot, sheetGroupObjects, insertRibbon, formulaRibbon, formulaDefineNameMain, formulaDefineNameMenu, formulaDefineNameCreate, createNamesDialog, formulaNames, formulaNameManager, formulaUseNameMenu, formulaUseNameManage, formulaUseNameInserted, formulaPasteNameDialog, formulaPasteNameInserted, formulaPasteNameListDialog, formulaPasteNameList, formulaPasteNameListUndo, formulaPasteNameListRedo, formulaFunctionDialog, formulaFunctionSearch, formulaInserted, formulaFunctionArguments, formulaLogicalMenu, formulaMenuArguments, formulaMenuInserted, formulaRecentMenu, formulaMathMenu, formulaVariableArguments, formulaZeroArgumentFunction, formulaFinancialMenu, formulaDateMenu, formulaTextMenu, formulaMoreMenu, formulaAutoSumMenu, formulaAutoSumInserted, formulaShowState, formulaTrace, formulaRemoveArrowsMenu, formulaRemovePrecedents, formulaTraceDependents, formulaRemoveDependents, formulaEvaluateInitial, formulaEvaluateStepIn, formulaEvaluateStep, formulaEvaluateStepOut, formulaEvaluateReferenceInitial, formulaEvaluateReferenceStep, formulaEvaluateReferenceStepIn, formulaEvaluateReferenceStepOut, formulaWatchWindow, formulaWatchSelect, formulaWatchDelete, formulaErrorMenu, formulaCircularSelection, formulaErrorChecking, formulaErrorIgnore, formulaErrorTrace, formulaErrorResetIgnored, formulaCalculation, viewRibbon, viewFullScreenOn, fullScreenUndoRedo, viewFullScreenOff, sheetViewNew, sheetViewKept, sheetViewMenuBeforeRename, sheetViewRenamed, sheetViewExited, sheetViewReactivated, sheetViewUpdated, sheetViewMenuBeforeCopy, sheetViewDuplicated, viewZoomDialog, viewZoomApplied, viewZoomHundred, viewZoomSelection, viewZoomSelectionUndoRedo, viewRulerOn, pageLayoutHeaderFooterDialog, pageBreakPreview, pageBreakContextMenu, customViewCaptureState, customViewsMainBeforeAdd, customViewAddDialog, customViewsDialog, customViewSaveUndoRedo, customViewApplied, customViewApplyUndoRedo, customViewDeleteFlow, freezeMenuLabels, freezeMenuInitialMeta, freezeMenuTopRow, freezeMenuTopRowStatus, freezeMenuAccessKey, formulaBarUndoRedo, viewToggles, rtlVisualHitTest, gridlinesUndoRedo, headingsUndoRedo, zeroValuesUndoRedo, objectsUndoRedo, outlineUndoRedo, rightToLeftUndoRedo });
}

export {
  runEditingAndFormat,
};
