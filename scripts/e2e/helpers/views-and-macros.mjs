// Sequential E2E feature module.
import {
  appUrl,
  workbookPath,
} from "../../e2e-real-upload.mjs";
import {
  createAdvancedFilterUniqueSubsetWorkbook,
  createAdvancedFilterUniqueWorkbook,
  createAutoFilterBlankWorkbook,
  createAutoFilterDateWorkbook,
  createAutoFilterMenuWorkbook,
  createDelimitedImportWorkbook,
  waitForUploadedWorkbook,
} from "./fixtures-and-model.mjs";
import {
  exportedAutoFilterDateGroupWorkbookState,
} from "./editing-and-validation.mjs";


async function importedFilterStructureShiftState(context, workbookPath) {
  const page = await context.newPage();
  const applyStructure = async (address, command, option) => {
    await page.locator(`.sheet-cell[data-address="${address}"]`).click();
    await page.locator('[data-ribbon-tab="home"]').click();
    await page.locator(`[data-ribbon-command="${command}"]`).click();
    await page.waitForSelector("#cellStructureDialog", { timeout: 5000 });
    await page.locator(`#cellStructureDialog input[value="${option}"]`).check();
    await page.locator("#cellStructureDialog [data-cell-structure-action='ok']").click();
    await page.waitForSelector("#cellStructureDialog", { state: "detached", timeout: 5000 });
  };
  try {
    await page.goto(`${appUrl}?importedFilterStructure=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);

    await page.evaluate((name) => window.__webSheetDevSetActiveSheet?.(name), "複数テーブルフィルター");
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "複数テーブルフィルター",
      null,
      { timeout: 5000 },
    );
    await page.waitForSelector('.cell-filter-button[data-filter-col="4"]', { timeout: 5000 });
    await applyStructure("D1", "insert-cells", "entire-column");
    await page.waitForSelector('.cell-filter-button[data-filter-col="5"].is-filtered', { timeout: 5000 });
    const afterColumnInsert = await page.evaluate(() => {
      const sheet = window.__webSheetDevSheetModel?.("複数テーブルフィルター");
      return {
        buttons: [...document.querySelectorAll(".cell-filter-button")].map((button) => ({
          col: Number(button.getAttribute("data-filter-col")),
          filtered: button.classList.contains("is-filtered"),
        })),
        rightTable: (sheet?.tables || []).find((table) => table.name === "RightFilterTable") || null,
        rows: [2, 3, 4].map((row) => ({
          row,
          hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
          shiftedRight: document.querySelector(`.sheet-cell[data-address="E${row}"] .cell-content`)?.textContent || "",
        })),
      };
    });
    await applyStructure("E1", "delete-cells", "entire-column");
    await page.waitForFunction(
      () => !document.querySelector('.cell-filter-button[data-filter-col="5"].is-filtered'),
      null,
      { timeout: 5000 },
    );
    const afterFilteredColumnDelete = await page.evaluate(() => {
      const sheet = window.__webSheetDevSheetModel?.("複数テーブルフィルター");
      return {
        buttons: [...document.querySelectorAll(".cell-filter-button")].map((button) => ({
          col: Number(button.getAttribute("data-filter-col")),
          filtered: button.classList.contains("is-filtered"),
        })),
        rightTable: (sheet?.tables || []).find((table) => table.name === "RightFilterTable") || null,
      };
    });

    await page.evaluate((name) => window.__webSheetDevSetActiveSheet?.(name), "保存済み並べ替え");
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "保存済み並べ替え",
      null,
      { timeout: 5000 },
    );
    await applyStructure("B1", "insert-cells", "entire-column");
    const afterSortColumnInsert = await page.evaluate(() => {
      const sheet = window.__webSheetDevSheetModel?.("保存済み並べ替え");
      return {
        sortState: sheet?.sortState || null,
        buttons: [...document.querySelectorAll(".cell-filter-button")].map((button) => ({
          col: Number(button.getAttribute("data-filter-col")),
          filtered: button.classList.contains("is-filtered"),
        })),
      };
    });
    return { afterColumnInsert, afterFilteredColumnDelete, afterSortColumnInsert };
  } finally {
    await page.close();
  }
}



async function importedTableRowInsertState(context, workbookPath) {
  const page = await context.newPage();
  const applyStructure = async (address, command, option) => {
    await page.locator(`.sheet-cell[data-address="${address}"]`).click();
    await page.locator('[data-ribbon-tab="home"]').click();
    await page.locator(`[data-ribbon-command="${command}"]`).click();
    await page.waitForSelector("#cellStructureDialog", { timeout: 5000 });
    await page.locator(`#cellStructureDialog input[value="${option}"]`).check();
    await page.locator("#cellStructureDialog [data-cell-structure-action='ok']").click();
    await page.waitForSelector("#cellStructureDialog", { state: "detached", timeout: 5000 });
  };
  try {
    await page.goto(`${appUrl}?importedTableRowInsert=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);

    await page.evaluate((name) => window.__webSheetDevSetActiveSheet?.(name), "テーブル拡張");
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "テーブル拡張",
      null,
      { timeout: 5000 },
    );
    await page.waitForSelector('.cell-filter-button[data-filter-col="3"]', { timeout: 5000 });
    const captureTableState = () => page.evaluate(() => {
      const sheet = window.__webSheetDevSheetModel?.("テーブル拡張");
      const table = sheet?.tables?.[0] || null;
      const cell = (row, col) => sheet?.cells?.[`${row}:${col}`] || null;
      const text = (address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "";
      const fill = (row, col) => cell(row, col)?.css?.backgroundColor || "";
      const borderKeys = (row, col) => Object.keys(cell(row, col)?.borders || {}).sort().join("|");
      return {
        table,
        autoFilterRange: table?.autoFilterState?.range || sheet?.autoFilter?.range || null,
        buttons: [...document.querySelectorAll(".cell-filter-button")].map((button) => ({
          row: Number(button.getAttribute("data-filter-row")),
          col: Number(button.getAttribute("data-filter-col")),
        })),
        values: {
          a3: text("A3"),
          b3: text("B3"),
          c3: text("C3"),
          a4: text("A4"),
          b4: text("B4"),
          c4: text("C4"),
          b1: text("B1"),
          c1: text("C1"),
          d1: text("D1"),
          d2: text("D2"),
          d3: text("D3"),
        },
        formulas: {
          c3: cell(3, 3)?.formula || "",
          c4: cell(4, 3)?.formula || "",
          d2: cell(2, 4)?.formula || "",
          d3: cell(3, 4)?.formula || "",
        },
        style: {
          insertedFill: fill(3, 1),
          shiftedFill: fill(4, 1),
          insertedBorders: borderKeys(3, 1),
          shiftedBorders: borderKeys(4, 1),
          insertedHeight: sheet?.rows?.[2]?.height || 0,
          shiftedHeight: sheet?.rows?.[3]?.height || 0,
          insertedColumnHeaderFill: fill(1, 2),
          shiftedColumnHeaderFill: fill(1, 3),
          insertedColumnBorders: borderKeys(2, 2),
          shiftedColumnBorders: borderKeys(2, 3),
          insertedColumnWidth: sheet?.columns?.[1]?.width || 0,
          shiftedColumnWidth: sheet?.columns?.[2]?.width || 0,
        },
      };
    });
    await applyStructure("A3", "insert-cells", "entire-row");
    await page.waitForFunction(
      () => {
        const sheet = window.__webSheetDevSheetModel?.("テーブル拡張");
        return sheet?.tables?.[0]?.range?.bottom === 4 &&
          sheet?.cells?.["3:3"]?.formula === "=[@Amount]*2" &&
          sheet?.cells?.["4:3"]?.formula === "=[@Amount]*2" &&
          document.querySelector('.sheet-cell[data-address="A4"] .cell-content')?.textContent === "B";
      },
      null,
      { timeout: 5000 },
    );
    const afterRowInsert = await captureTableState();
    await applyStructure("A3", "delete-cells", "entire-row");
    await page.waitForFunction(
      () => {
        const sheet = window.__webSheetDevSheetModel?.("テーブル拡張");
        return sheet?.tables?.[0]?.range?.bottom === 3 &&
          document.querySelector('.sheet-cell[data-address="A3"] .cell-content')?.textContent === "B" &&
          sheet?.cells?.["3:3"]?.formula === "=[@Amount]*2";
      },
      null,
      { timeout: 5000 },
    );
    const afterRowDelete = await captureTableState();
    await applyStructure("B1", "insert-cells", "entire-column");
    await page.waitForFunction(
      () => {
        const sheet = window.__webSheetDevSheetModel?.("テーブル拡張");
        const table = sheet?.tables?.[0];
        return table?.range?.right === 4 &&
          table?.columns?.map((column) => column.name).join("|") === "Name|Column2|Amount|Twice" &&
          document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent === "Column2" &&
          sheet?.cells?.["3:4"]?.formula === "=[@Amount]*2";
      },
      null,
      { timeout: 5000 },
    );
    const afterColumnInsert = await captureTableState();
    await applyStructure("B1", "delete-cells", "entire-column");
    await page.waitForFunction(
      () => {
        const sheet = window.__webSheetDevSheetModel?.("テーブル拡張");
        const table = sheet?.tables?.[0];
        return table?.range?.right === 3 &&
          table?.columns?.map((column) => column.name).join("|") === "Name|Amount|Twice" &&
          document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent === "Amount" &&
          sheet?.cells?.["3:3"]?.formula === "=[@Amount]*2";
      },
      null,
      { timeout: 5000 },
    );
    const afterColumnDelete = await captureTableState();
    return { afterRowInsert, afterRowDelete, afterColumnInsert, afterColumnDelete };
  } finally {
    await page.close();
  }
}



async function importedTableTotalsStructureState(context, workbookPath) {
  const page = await context.newPage();
  const applyStructure = async (address, command, option) => {
    await page.locator(`.sheet-cell[data-address="${address}"]`).click();
    await page.locator('[data-ribbon-tab="home"]').click();
    await page.locator(`[data-ribbon-command="${command}"]`).click();
    await page.waitForSelector("#cellStructureDialog", { timeout: 5000 });
    await page.locator(`#cellStructureDialog input[value="${option}"]`).check();
    await page.locator("#cellStructureDialog [data-cell-structure-action='ok']").click();
    await page.waitForSelector("#cellStructureDialog", { state: "detached", timeout: 5000 });
  };
  try {
    await page.goto(`${appUrl}?importedTableTotalsStructure=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.evaluate((name) => window.__webSheetDevSetActiveSheet?.(name), "テーブルメタ");
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "テーブルメタ",
      null,
      { timeout: 5000 },
    );
    await page.waitForSelector('.cell-filter-button[data-filter-col="4"]', { timeout: 5000 });
    const capture = () => page.evaluate(() => {
      const sheet = window.__webSheetDevSheetModel?.("テーブルメタ");
      const table = sheet?.tables?.[0] || null;
      const cell = (row, col) => sheet?.cells?.[`${row}:${col}`] || null;
      const text = (address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "";
      const fill = (row, col) => cell(row, col)?.css?.backgroundColor || "";
      return {
        table,
        sheetAutoFilter: sheet?.autoFilter || null,
        tableAutoFilter: table?.autoFilterState || null,
        buttons: [...document.querySelectorAll(".cell-filter-button")].map((button) => ({
          row: Number(button.getAttribute("data-filter-row")),
          col: Number(button.getAttribute("data-filter-col")),
        })),
        values: {
          a5: text("A5"),
          b5: text("B5"),
          c5: text("C5"),
          d5: text("D5"),
          a6: text("A6"),
          b6: text("B6"),
          c6: text("C6"),
          d6: text("D6"),
          b1: text("B1"),
          c1: text("C1"),
          d1: text("D1"),
          e1: text("E1"),
          e5: text("E5"),
        },
        formulas: {
          d5: cell(5, 4)?.formula || "",
          b6: cell(6, 2)?.formula || "",
          d6: cell(6, 4)?.formula || "",
          c5: cell(5, 3)?.formula || "",
          e5: cell(5, 5)?.formula || "",
          e4: cell(4, 5)?.formula || "",
        },
        style: {
          insertedDataFill: fill(5, 1),
          totalFillRow6: fill(6, 1),
          totalFillRow5: fill(5, 1),
          insertedColumnHeaderFill: fill(1, 2),
          shiftedHeaderFill: fill(1, 3),
          insertedColumnWidth: sheet?.columns?.[1]?.width || 0,
          shiftedColumnWidth: sheet?.columns?.[2]?.width || 0,
        },
      };
    });
    await applyStructure("A5", "insert-cells", "entire-row");
    await page.waitForFunction(
      () => {
        const sheet = window.__webSheetDevSheetModel?.("テーブルメタ");
        const table = sheet?.tables?.[0];
        return table?.range?.bottom === 6 &&
          table?.autoFilterState?.range?.bottom === 5 &&
          sheet?.autoFilter?.range?.bottom === 5 &&
          sheet?.cells?.["5:4"]?.formula === "=[@売上]*2" &&
          document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent === "合計";
      },
      null,
      { timeout: 5000 },
    );
    const afterTotalsRowInsert = await capture();
    await applyStructure("A5", "delete-cells", "entire-row");
    await page.waitForFunction(
      () => {
        const sheet = window.__webSheetDevSheetModel?.("テーブルメタ");
        const table = sheet?.tables?.[0];
        return table?.range?.bottom === 5 &&
          table?.autoFilterState?.range?.bottom === 4 &&
          sheet?.autoFilter?.range?.bottom === 4 &&
          document.querySelector('.sheet-cell[data-address="A5"] .cell-content')?.textContent === "合計";
      },
      null,
      { timeout: 5000 },
    );
    const afterTotalsRowDelete = await capture();
    await applyStructure("B1", "insert-cells", "entire-column");
    await page.waitForFunction(
      () => {
        const sheet = window.__webSheetDevSheetModel?.("テーブルメタ");
        const table = sheet?.tables?.[0];
        return table?.range?.right === 5 &&
          table?.autoFilterState?.range?.right === 5 &&
          sheet?.autoFilter?.range?.right === 5 &&
          table?.columns?.map((column) => column.name).join("|") === "地域|Column2|売上|担当|倍額" &&
          document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent === "Column2" &&
          sheet?.cells?.["4:5"]?.formula === "=[@売上]*2";
      },
      null,
      { timeout: 5000 },
    );
    const afterTotalsColumnInsert = await capture();
    await applyStructure("B1", "delete-cells", "entire-column");
    await page.waitForFunction(
      () => {
        const sheet = window.__webSheetDevSheetModel?.("テーブルメタ");
        const table = sheet?.tables?.[0];
        return table?.range?.right === 4 &&
          table?.autoFilterState?.range?.right === 4 &&
          sheet?.autoFilter?.range?.right === 4 &&
          table?.columns?.map((column) => column.name).join("|") === "地域|売上|担当|倍額" &&
          document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent === "売上" &&
          sheet?.cells?.["4:4"]?.formula === "=[@売上]*2";
      },
      null,
      { timeout: 5000 },
    );
    const afterTotalsColumnDelete = await capture();
    return { afterTotalsRowInsert, afterTotalsRowDelete, afterTotalsColumnInsert, afterTotalsColumnDelete };
  } finally {
    await page.close();
  }
}



async function importedStructureReferenceShiftState(context, workbookPath) {
  const page = await context.newPage();
  const applyStructure = async (address, command, option) => {
    await page.locator(`.sheet-cell[data-address="${address}"]`).click();
    await page.locator('[data-ribbon-tab="home"]').click();
    await page.locator(`[data-ribbon-command="${command}"]`).click();
    await page.waitForSelector("#cellStructureDialog", { timeout: 5000 });
    await page.locator(`#cellStructureDialog input[value="${option}"]`).check();
    await page.locator("#cellStructureDialog [data-cell-structure-action='ok']").click();
    await page.waitForSelector("#cellStructureDialog", { state: "detached", timeout: 5000 });
  };
  const capture = () => page.evaluate(() => {
    const refSheet = window.__webSheetDevSheetModel?.("構造シフト参照");
    const names = window.__webSheetDevDefinedNames?.() || [];
    const trackedNames = ["SalesDataAmounts", "SalesTotalCell", "SalesWholeAmountColumn", "SalesDataRows"];
    const nameMap = Object.fromEntries(names.filter((item) => trackedNames.includes(item.name)).map((item) => [item.name, item.refersTo]));
    const formula = (row, col = 2) => refSheet?.cells?.[`${row}:${col}`]?.formula || refSheet?.cells?.[`${row}:${col}`]?.raw || "";
    return {
      formulas: {
        dataSum: formula(1),
        totalCell: formula(2),
        lastCalculated: formula(3),
        stringLiteral: formula(4),
        wholeColumn: formula(5),
        wholeRows: formula(6),
        columnRange: formula(7),
      },
      validationList: refSheet?.cells?.["8:2"]?.validation?.formula || refSheet?.cells?.["8:2"]?.validation?.formulae?.[0] || "",
      definedNames: nameMap,
    };
  });
  try {
    await page.goto(`${appUrl}?importedStructureReferenceShift=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.evaluate((name) => window.__webSheetDevSetActiveSheet?.(name), "テーブルメタ");
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "テーブルメタ",
      null,
      { timeout: 5000 },
    );
    const initial = await capture();
    await applyStructure("A3", "insert-cells", "entire-row");
    await page.waitForFunction(
      () => window.__webSheetDevSheetModel?.("構造シフト参照")?.cells?.["1:2"]?.formula === "=SUM('テーブルメタ'!B2:B5)",
      null,
      { timeout: 5000 },
    );
    const afterRowInsert = await capture();
    await applyStructure("A3", "delete-cells", "entire-row");
    await page.waitForFunction(
      () => window.__webSheetDevSheetModel?.("構造シフト参照")?.cells?.["1:2"]?.formula === "=SUM('テーブルメタ'!B2:B4)",
      null,
      { timeout: 5000 },
    );
    const afterRowDelete = await capture();
    await applyStructure("B1", "insert-cells", "entire-column");
    await page.waitForFunction(
      () => window.__webSheetDevSheetModel?.("構造シフト参照")?.cells?.["1:2"]?.formula === "=SUM('テーブルメタ'!C2:C4)",
      null,
      { timeout: 5000 },
    );
    const afterColumnInsert = await capture();
    await applyStructure("B1", "delete-cells", "entire-column");
    await page.waitForFunction(
      () => window.__webSheetDevSheetModel?.("構造シフト参照")?.cells?.["1:2"]?.formula === "=SUM('テーブルメタ'!B2:B4)",
      null,
      { timeout: 5000 },
    );
    const afterColumnDelete = await capture();
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => window.__webSheetDevSheetModel?.("構造シフト参照")?.cells?.["1:2"]?.formula === "=SUM('テーブルメタ'!C2:C4)",
      null,
      { timeout: 5000 },
    );
    const afterUndoColumnDelete = await capture();
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(
      () => window.__webSheetDevSheetModel?.("構造シフト参照")?.cells?.["1:2"]?.formula === "=SUM('テーブルメタ'!B2:B4)",
      null,
      { timeout: 5000 },
    );
    const afterRedoColumnDelete = await capture();
    return { initial, afterRowInsert, afterRowDelete, afterColumnInsert, afterColumnDelete, afterUndoColumnDelete, afterRedoColumnDelete };
  } finally {
    await page.close();
  }
}



async function importedTableMetadataState(page, sheetName) {
  await page.evaluate((name) => window.__webSheetDevSetActiveSheet?.(name), sheetName);
  await page.waitForFunction(
    (name) => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === name,
    sheetName,
    { timeout: 5000 },
  );
  await page.waitForSelector('.sheet-cell[data-address="A1"]', { timeout: 5000 });
  return page.evaluate((name) => {
    const sheet = window.__webSheetDevSheetModel?.(name);
    const table = sheet?.tables?.[0] || null;
    const header = document.querySelector('.sheet-cell[data-address="A1"]');
    const stripe = document.querySelector('.sheet-cell[data-address="A3"]');
    const total = document.querySelector('.sheet-cell[data-address="A5"]');
    return {
      table,
      tableCount: sheet?.tables?.length || 0,
      hasFilterButton: Boolean(document.querySelector('.cell-filter-button[data-filter-col="1"]')),
      doubledRows: ["D2", "D3", "D4"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      totals: {
        sales: document.querySelector('.sheet-cell[data-address="B5"] .cell-content')?.textContent || "",
        doubled: document.querySelector('.sheet-cell[data-address="D5"] .cell-content')?.textContent || "",
      },
      headerFill: header ? getComputedStyle(header).getPropertyValue("--cell-fill-color").trim() || getComputedStyle(header).backgroundColor : "",
      stripeFill: stripe ? getComputedStyle(stripe).getPropertyValue("--cell-fill-color").trim() || getComputedStyle(stripe).backgroundColor : "",
      totalFill: total ? getComputedStyle(total).getPropertyValue("--cell-fill-color").trim() || getComputedStyle(total).backgroundColor : "",
    };
  }, sheetName);
}



async function importedStructuredReferenceState(page, sheetName) {
  await page.evaluate((name) => window.__webSheetDevSetActiveSheet?.(name), sheetName);
  await page.waitForFunction(
    (name) => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === name,
    sheetName,
    { timeout: 5000 },
  );
  await page.waitForSelector('.sheet-cell[data-address="B1"]', { timeout: 5000 });
  const initial = await page.evaluate((name) => {
    const sheet = window.__webSheetDevSheetModel?.(name);
    return {
      values: ["B1", "B2", "B3", "B4", "B5"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      formulas: ["B1", "B2", "B3", "B4", "B5"].map((address) => {
        const point = address.match(/^([A-Z]+)(\d+)$/);
        if (!point) return "";
        const col = point[1].split("").reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
        return sheet?.cells?.[`${Number(point[2])}:${col}`]?.formula || "";
      }),
    };
  }, sheetName);
  await page.evaluate(() => window.__webSheetDevSetActiveSheet?.("テーブルメタ"));
  await page.waitForFunction(
    () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "テーブルメタ",
    null,
    { timeout: 5000 },
  );
  await editCellWithFormulaBar(page, "B1", "金額");
  await page.waitForFunction(
    () =>
      window.__webSheetDevSheetModel?.("テーブルメタ")?.tables?.[0]?.columns?.[1]?.name === "金額" &&
      window.__webSheetDevSheetModel?.("テーブルメタ")?.cells?.["2:4"]?.formula === "=[@[金額]]*2",
    null,
    { timeout: 10000 },
  );
  await page.evaluate((name) => window.__webSheetDevSetActiveSheet?.(name), sheetName);
  await page.waitForFunction(
    (name) =>
      document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === name &&
      document.querySelector('.sheet-cell[data-address="B3"] .cell-content')?.textContent === "金額",
    sheetName,
    { timeout: 5000 },
  );
  const renamed = await page.evaluate((name) => {
    const sheet = window.__webSheetDevSheetModel?.(name);
    const tableSheet = window.__webSheetDevSheetModel?.("テーブルメタ");
    return {
      values: ["B1", "B2", "B3", "B4", "B5"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      formulas: ["B1", "B2", "B3", "B4", "B5"].map((address) => {
        const point = address.match(/^([A-Z]+)(\d+)$/);
        if (!point) return "";
        const col = point[1].split("").reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
        return sheet?.cells?.[`${Number(point[2])}:${col}`]?.formula || "";
      }),
      tableColumns: tableSheet?.tables?.[0]?.columns?.map((column) => column.name) || [],
      tableFormula: tableSheet?.cells?.["2:4"]?.formula || "",
    };
  }, sheetName);
  return { ...initial, renamed };
}



async function importedTableExpansionState(page, sheetName) {
  await page.evaluate((name) => window.__webSheetDevSetActiveSheet?.(name), sheetName);
  await page.waitForFunction(
    (name) => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === name,
    sheetName,
    { timeout: 5000 },
  );
  await page.waitForSelector('.sheet-cell[data-address="A4"]', { timeout: 5000 });
  await editCellWithFormulaBar(page, "A4", "C");
  await page.waitForFunction(
    (name) => window.__webSheetDevSheetModel?.(name)?.tables?.[0]?.range?.bottom === 4,
    sheetName,
    { timeout: 5000 },
  );
  await editCellWithFormulaBar(page, "B4", "3");
  await page.waitForFunction(
    () => document.querySelector('.sheet-cell[data-address="C4"] .cell-content')?.textContent === "6",
    null,
    { timeout: 5000 },
  );
  await editCellWithFormulaBar(page, "D1", "Memo");
  await page.waitForFunction(
    (name) => window.__webSheetDevSheetModel?.(name)?.tables?.[0]?.range?.right === 4,
    sheetName,
    { timeout: 5000 },
  );
  await page.evaluate(async () => navigator.clipboard.writeText("D\t4\nE\t5"));
  await page.locator('.sheet-cell[data-address="A5"]').click();
  await page.keyboard.press("Control+V");
  await page.waitForFunction(
    (name) =>
      window.__webSheetDevSheetModel?.(name)?.tables?.[0]?.range?.bottom === 6 &&
      document.querySelector('.sheet-cell[data-address="C6"] .cell-content')?.textContent === "10",
    sheetName,
    { timeout: 5000 },
  );
  await editCellWithFormulaBar(page, "C5", "=[@Amount]*3");
  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="C4"] .cell-content')?.textContent === "9" &&
      document.querySelector('.sheet-cell[data-address="C6"] .cell-content')?.textContent === "15",
    null,
    { timeout: 5000 },
  );
  return page.evaluate((name) => {
    const sheet = window.__webSheetDevSheetModel?.(name);
    const table = sheet?.tables?.[0] || null;
    return {
      table,
      autoFilterRange: sheet?.autoFilter?.range || null,
      c4: document.querySelector('.sheet-cell[data-address="C4"] .cell-content')?.textContent || "",
      d1: document.querySelector('.sheet-cell[data-address="D1"] .cell-content')?.textContent || "",
      c4Formula: sheet?.cells?.["4:3"]?.formula || "",
      c5: document.querySelector('.sheet-cell[data-address="C5"] .cell-content')?.textContent || "",
      c5Formula: sheet?.cells?.["5:3"]?.formula || "",
      c6: document.querySelector('.sheet-cell[data-address="C6"] .cell-content')?.textContent || "",
      c6Formula: sheet?.cells?.["6:3"]?.formula || "",
      calculatedFormula: table?.columns?.[2]?.calculatedColumnFormula || "",
      columns: table?.columns?.map((column) => column.name) || [],
    };
  }, sheetName);
}



async function editCellWithFormulaBar(page, address, value) {
  await page.locator(`.sheet-cell[data-address="${address}"]`).click();
  await page.locator("#formulaInput").fill(String(value));
  await page.keyboard.press("Enter");
}



async function exerciseAutoFilterBlankAndKeyboardMenu(context) {
  const workbookPath = await createAutoFilterBlankWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?autoFilterBlankKeyboard=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="A5"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="A5"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-filter"]').click();
    await page.waitForSelector('.cell-filter-button[data-filter-col="1"]', { timeout: 5000 });
    const readMemoRows = () =>
      page.evaluate(() =>
        [2, 3, 4, 5].map((row) => ({
          row,
          hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
          memo: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        })),
      );
    await page.locator('.cell-filter-button[data-filter-col="1"]').click();
    await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await page.waitForFunction(
      () => document.activeElement?.matches?.('#autoFilterMenu [data-auto-filter-action="sort-asc"]'),
      null,
      { timeout: 5000 },
    );
    const keyboardMenuInitialFocus = await page.evaluate(() => document.activeElement?.textContent?.trim() || "");
    await page.keyboard.press("ArrowDown");
    const keyboardMenuAfterArrow = await page.evaluate(() => document.activeElement?.textContent?.trim() || "");
    await page.keyboard.press("Escape");
    await page.waitForSelector("#autoFilterMenu", { state: "detached", timeout: 5000 });
    const keyboardMenuClosedByEscape = await page.evaluate(() => !document.querySelector("#autoFilterMenu"));

    await page.locator('.cell-filter-button[data-filter-col="1"]').click();
    await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await page.locator("[data-auto-filter-search]").fill("No");
    await page.waitForFunction(
      () => [...document.querySelectorAll("#autoFilterMenu .auto-filter-values .auto-filter-value:not(.is-hidden)")].length === 1,
      null,
      { timeout: 5000 },
    );
    const searchSelectStateBefore = await page.evaluate(() => {
      const selectAll = document.querySelector("[data-auto-filter-select-all]");
      return {
        labels: [...document.querySelectorAll("#autoFilterMenu .auto-filter-values .auto-filter-value:not(.is-hidden)")].map((label) =>
          label.textContent.replace(/\s+/g, " ").trim(),
        ),
        checked: selectAll?.checked ?? false,
        indeterminate: selectAll?.indeterminate ?? false,
        disabled: selectAll?.disabled ?? true,
      };
    });
    await page.locator("[data-auto-filter-select-all]").uncheck();
    const searchSelectStateAfter = await page.evaluate(() => {
      const selectAll = document.querySelector("[data-auto-filter-select-all]");
      return {
        checked: selectAll?.checked ?? true,
        indeterminate: selectAll?.indeterminate ?? true,
        states: [...document.querySelectorAll("#autoFilterMenu .auto-filter-values .auto-filter-value")].map((label) => ({
          label: label.textContent.replace(/\s+/g, " ").trim(),
          hidden: label.classList.contains("is-hidden"),
          checked: label.querySelector("input")?.checked ?? false,
        })),
      };
    });
    await page.locator('[data-auto-filter-action="ok"]').click();
    await page.waitForFunction(
      () => getComputedStyle(document.querySelector('.row-header[data-row="5"]')).display === "none",
      null,
      { timeout: 5000 },
    );
    const valueSearchRows = await readMemoRows();
    await page.locator('[data-ribbon-command="data-clear-filter"]').click();
    await page.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );

    await page.locator('.cell-filter-button[data-filter-col="1"]').click();
    await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await page.locator('[data-auto-filter-action="condition-menu"]').click();
    await page.waitForSelector(".auto-filter-submenu:not(.hidden)", { timeout: 5000 });
    await page.locator('[data-auto-filter-operator="blanks"]').click();
    await page.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="1"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const blankRows = await readMemoRows();
    await page.locator('[data-ribbon-command="data-clear-filter"]').click();
    await page.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );

    await page.locator('.cell-filter-button[data-filter-col="1"]').click();
    await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await page.locator('[data-auto-filter-action="condition-menu"]').click();
    await page.waitForSelector(".auto-filter-submenu:not(.hidden)", { timeout: 5000 });
    await page.locator('[data-auto-filter-operator="nonBlanks"]').click();
    await page.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="1"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const nonBlankRows = await readMemoRows();

    return { keyboardMenuInitialFocus, keyboardMenuAfterArrow, keyboardMenuClosedByEscape, searchSelectStateBefore, searchSelectStateAfter, valueSearchRows, blankRows, nonBlankRows };
  } finally {
    await page.close();
  }
}



async function autoFilterVisibleRowsSnapshot(page, rows = [2, 3, 4, 5]) {
  return page.evaluate((targetRows) =>
    targetRows.map((row) => ({
      row,
      hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
      name: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
      score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
    })),
  rows);
}



async function exerciseAutoFilterReapplyAndClear(context) {
  const workbookPath = await createAutoFilterMenuWorkbook();
  const page = await context.newPage();
  const ribbonState = () => page.evaluate(() => {
    const stateFor = (command) => {
      const button = document.querySelector(`[data-ribbon-command="${command}"]`);
      return {
        active: button?.classList.contains("is-active") || false,
        disabled: button?.disabled || false,
      };
    };
    return {
      filter: stateFor("data-filter"),
      clear: stateFor("data-clear-filter"),
      reapply: stateFor("data-reapply-filter"),
    };
  });
  try {
    await page.goto(`${appUrl}?autoFilterReapply=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="B6"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="B6"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-filter"]').click();
    await page.waitForSelector('.cell-filter-button[data-filter-col="2"]', { timeout: 5000 });
    await page.locator('.cell-filter-button[data-filter-col="2"]').click();
    await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await page.locator('[data-auto-filter-action="condition-menu"]').click();
    await page.waitForSelector(".auto-filter-submenu:not(.hidden)", { timeout: 5000 });
    await page.locator('[data-auto-filter-operator="greaterThan"]').click();
    await page.waitForSelector('[data-data-dialog="auto-filter-custom"]', { timeout: 5000 });
    await page.locator('[data-data-field="value1"]').fill("10");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="2"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const ribbonAfterFilter = await ribbonState();
    const initial = await autoFilterVisibleRowsSnapshot(page);
    await page.locator('.sheet-cell[data-address="B3"]').click();
    await page.locator("#formulaInput").fill("1");
    await page.keyboard.press("Enter");
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B3"] .cell-content')?.textContent === "1",
      null,
      { timeout: 5000 },
    );
    const afterEditBeforeReapply = await autoFilterVisibleRowsSnapshot(page);
    await page.locator('[data-ribbon-command="data-reapply-filter"]').click();
    await page.waitForFunction(
      () => getComputedStyle(document.querySelector('.row-header[data-row="3"]')).display === "none",
      null,
      { timeout: 5000 },
    );
    const afterReapply = await page.evaluate((targetRows) => ({
      rows: targetRows.map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        name: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
      status: document.querySelector("#statusText")?.textContent || "",
      filtered: document.querySelector('.cell-filter-button[data-filter-col="2"]')?.classList.contains("is-filtered") || false,
    }), [2, 3, 4, 5]);
    await page.locator('[data-ribbon-command="data-clear-filter"]').click();
    await page.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );
    const ribbonAfterClear = await ribbonState();
    const afterClear = await page.evaluate((targetRows) => ({
      rows: targetRows.map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        name: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
      status: document.querySelector("#statusText")?.textContent || "",
      filteredButtons: [...document.querySelectorAll(".cell-filter-button.is-filtered")].length,
    }), [2, 3, 4, 5]);

    await page.locator('.sheet-cell[data-address="B3"]').click();
    await page.locator('[data-ribbon-command="data-sort-desc"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent === "Beta",
      null,
      { timeout: 5000 },
    );
    const ribbonAfterSort = await ribbonState();
    const sortInitial = await autoFilterVisibleRowsSnapshot(page);
    const sortHeaderState = await page.evaluate(() => {
      const button = document.querySelector('.cell-filter-button[data-filter-col="2"]');
      return {
        sorted: button?.classList.contains("is-sorted") || false,
        desc: button?.classList.contains("is-sort-desc") || false,
        direction: button?.getAttribute("data-sort-direction") || "",
        title: button?.getAttribute("title") || "",
        sortState: window.__webSheetDevSheetModel?.("Filter")?.sortState || null,
      };
    });
    await page.locator('.cell-filter-button[data-filter-col="2"]').click();
    await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    const sortMenuState = await page.evaluate(() => ({
      direction: document.querySelector("#autoFilterMenu")?.dataset.sortDirection || "",
      descPressed: document.querySelector('#autoFilterMenu [data-auto-filter-action="sort-desc"]')?.getAttribute("aria-pressed") || "",
      current: document.querySelector("#autoFilterMenu .auto-filter-current")?.textContent || "",
    }));
    await page.keyboard.press("Escape");
    await page.waitForSelector("#autoFilterMenu", { state: "detached", timeout: 5000 });
    await page.locator('.sheet-cell[data-address="B5"]').click();
    await page.locator("#formulaInput").fill("30");
    await page.keyboard.press("Enter");
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B5"] .cell-content')?.textContent === "30",
      null,
      { timeout: 5000 },
    );
    const sortAfterEditBeforeReapply = await autoFilterVisibleRowsSnapshot(page);
    await page.locator('[data-ribbon-command="data-reapply-filter"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent === "Alpine",
      null,
      { timeout: 5000 },
    );
    const sortAfterReapply = {
      rows: await autoFilterVisibleRowsSnapshot(page),
      status: await page.locator("#statusText").innerText(),
    };
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent === "Beta",
      null,
      { timeout: 5000 },
    );
    const sortAfterUndo = {
      rows: await autoFilterVisibleRowsSnapshot(page),
      status: await page.locator("#statusText").innerText(),
    };
    await page.locator('[data-ribbon-command="data-clear-filter"]').click();
    await page.waitForFunction(
      () => document.querySelector("#statusText")?.textContent?.includes("並べ替え状態"),
      null,
      { timeout: 5000 },
    );
    const sortClear = {
      rows: await autoFilterVisibleRowsSnapshot(page),
      status: await page.locator("#statusText").innerText(),
      header: await page.evaluate(() => {
        const button = document.querySelector('.cell-filter-button[data-filter-col="2"]');
        return {
          sorted: button?.classList.contains("is-sorted") || false,
          direction: button?.getAttribute("data-sort-direction") || "",
        };
      }),
    };
    await page.locator('.sheet-cell[data-address="B5"]').click();
    await page.locator("#formulaInput").fill("40");
    await page.keyboard.press("Enter");
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B5"] .cell-content')?.textContent === "40",
      null,
      { timeout: 5000 },
    );
    await page.locator('[data-ribbon-command="data-reapply-filter"]').click();
    await page.waitForFunction(
      () => document.querySelector("#statusText")?.textContent?.includes("フィルターを再適用"),
      null,
      { timeout: 5000 },
    );
    const sortAfterClearReapply = {
      rows: await autoFilterVisibleRowsSnapshot(page),
      status: await page.locator("#statusText").innerText(),
    };
    return { initial, ribbonAfterFilter, afterEditBeforeReapply, afterReapply, afterClear, ribbonAfterClear, sortInitial, ribbonAfterSort, sortHeaderState, sortMenuState, sortAfterEditBeforeReapply, sortAfterReapply, sortAfterUndo, sortClear, sortAfterClearReapply };
  } finally {
    await page.close();
  }
}



async function exerciseAdvancedFilterCopyOverwrite(context) {
  const workbookPath = await createAutoFilterMenuWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?advancedFilterCopy=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="H1"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="B5"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-advanced-filter"]').click();
    await page.waitForSelector('[data-data-dialog="advanced-filter"]', { timeout: 5000 });
    const namedDefaults = await page.evaluate(() => ({
      action: document.querySelector("[data-data-field='action']:checked")?.value || "",
      listRange: document.querySelector("[data-data-field='listRange']")?.value || "",
      criteriaRange: document.querySelector("[data-data-field='criteriaRange']")?.value || "",
      copyToRange: document.querySelector("[data-data-field='copyToRange']")?.value || "",
    }));
    await page.locator("[data-data-field='action'][value='copy']").check();
    await page.locator("[data-data-field='criteriaRange']").fill("G1:G2");
    await page.locator("[data-data-field='copyToRange']").fill("H1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const overwritePrompt = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      buttons: [...document.querySelectorAll("#messageBoxOverlay [data-message-box-close]")].map((button) =>
        button.textContent.replace(/\s+/g, " ").trim(),
      ),
    }));
    await page.locator("#messageBoxOverlay [data-message-box-result='cancel']").click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    const afterCancel = await page.evaluate(() => ({
      dialogVisible: Boolean(document.querySelector('[data-data-dialog="advanced-filter"]')),
      h1: document.querySelector('.sheet-cell[data-address="H1"] .cell-content')?.textContent || "",
    }));
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    await page.locator("#messageBoxOverlay [data-message-box-result='ok']").click();
    await page.waitForSelector('[data-data-dialog="advanced-filter"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="H1"] .cell-content')?.textContent === "Name",
      null,
      { timeout: 5000 },
    );
    const values = await page.evaluate(() =>
      ["H1", "I1", "H2", "I2", "H3", "I3"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
    );
    const status = await page.locator("#statusText").innerText();

    await page.locator('[data-ribbon-command="data-advanced-filter"]').click();
    await page.waitForSelector('[data-data-dialog="advanced-filter"]', { timeout: 5000 });
    await page.locator("[data-data-field='action'][value='copy']").check();
    await page.locator("[data-data-field='listRange']").fill("Filter!A1:B5");
    await page.locator("[data-data-field='criteriaRange']").fill("Criteria!A1:A2");
    await page.locator("[data-data-field='copyToRange']").fill("Extract!A1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="advanced-filter"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.includes("Extract") &&
        document.querySelector('.sheet-cell[data-address="A1"] .cell-content')?.textContent === "Name",
      null,
      { timeout: 5000 },
    );
    const qualifiedCopy = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent.trim() || "",
      values: ["A1", "B1", "A2", "B2", "A3", "B3"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.locator('[data-ribbon-command="data-advanced-filter"]').click();
    await page.waitForSelector('[data-data-dialog="advanced-filter"]', { timeout: 5000 });
    await page.locator("[data-data-field='action'][value='copy']").check();
    await page.locator("[data-data-field='listRange']").fill("Filter!A1:B5");
    await page.locator("[data-data-field='criteriaRange']").fill("Criteria!B1:C3");
    await page.locator("[data-data-field='copyToRange']").fill("Extract!D1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="advanced-filter"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="D2"] .cell-content')?.textContent === "Alpine",
      null,
      { timeout: 5000 },
    );
    const wildcardAndOrCopy = await page.evaluate(() => ({
      values: ["D1", "E1", "D2", "E2", "D3", "E3", "D4", "E4"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.locator('[data-ribbon-command="data-advanced-filter"]').click();
    await page.waitForSelector('[data-data-dialog="advanced-filter"]', { timeout: 5000 });
    await page.locator("[data-data-field='action'][value='copy']").check();
    await page.locator("[data-data-field='listRange']").fill("Database");
    await page.locator("[data-data-field='criteriaRange']").fill("Criteria");
    await page.locator("[data-data-field='copyToRange']").fill("Extract!G1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="advanced-filter"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="G2"] .cell-content')?.textContent === "alpha",
      null,
      { timeout: 5000 },
    );
    const namedPrefixCopy = await page.evaluate(() => ({
      values: ["G1", "H1", "G2", "H2", "G3", "H3", "G4", "H4"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator(".sheet-tab", { hasText: "Filter" }).click();
    await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "Filter", null, { timeout: 5000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="B5"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="data-advanced-filter"]').click();
    await page.waitForSelector('[data-data-dialog="advanced-filter"]', { timeout: 5000 });
    await page.locator("[data-data-field='action'][value='in-place']").check();
    await page.locator("[data-data-field='listRange']").fill("Filter!A1:B5");
    await page.locator("[data-data-field='criteriaRange']").fill("Criteria!A1:A2");
    await page.locator("[data-data-field='uniqueOnly']").uncheck();
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => window.__webSheetDevSheetModel?.("Filter")?.rows?.[1]?.hidden === true &&
        window.__webSheetDevSheetModel?.("Filter")?.rows?.[2]?.hidden !== true,
      null,
      { timeout: 5000 },
    );
    const readInPlaceState = () =>
      page.evaluate(() => {
        const sheet = window.__webSheetDevSheetModel?.("Filter");
        return {
          activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent.trim() || "",
          rows: [2, 3, 4, 5].map((row) => ({ row, hidden: sheet?.rows?.[row - 1]?.hidden === true })),
          filterRange: sheet?.autoFilter?.range || null,
          criteriaRange: sheet?.autoFilter?.advancedCriteria?.criteriaRange || null,
          uniqueOnly: sheet?.autoFilter?.advancedCriteria?.uniqueOnly === true,
          status: document.querySelector("#statusText")?.textContent || "",
        };
      });
    const inPlace = await readInPlaceState();
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => [1, 2, 3, 4].every((index) => window.__webSheetDevSheetModel?.("Filter")?.rows?.[index]?.hidden !== true) &&
        document.querySelector("#statusText")?.textContent?.includes("元に戻しました"),
      null,
      { timeout: 5000 },
    );
    const inPlaceUndo = await readInPlaceState();
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(
      () => window.__webSheetDevSheetModel?.("Filter")?.rows?.[1]?.hidden === true &&
        document.querySelector("#statusText")?.textContent?.includes("やり直しました"),
      null,
      { timeout: 5000 },
    );
    const inPlaceRedo = await readInPlaceState();
    return { overwritePrompt, afterCancel, values, status, qualifiedCopy, wildcardAndOrCopy, namedDefaults, namedPrefixCopy, inPlace, inPlaceUndo, inPlaceRedo };
  } finally {
    await page.close();
  }
}



async function exerciseAutoFilterDateMenu(context) {
  const workbookPath = await createAutoFilterDateWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?dateFilter=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="B6"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="B6"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-filter"]').click();
    await page.waitForSelector('.cell-filter-button[data-filter-col="1"]', { timeout: 5000 });
    const readDateRows = () =>
      page.evaluate(() =>
        [2, 3, 4, 5, 6].map((row) => ({
          row,
          hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
          date: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        })),
      );
    const applyDatePreset = async (operator) => {
      await page.locator('.cell-filter-button[data-filter-col="1"]').click();
      await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
      await page.locator('[data-auto-filter-action="condition-menu"]').click();
      await page.waitForSelector(".auto-filter-submenu:not(.hidden)", { timeout: 5000 });
      await page.locator(`[data-auto-filter-operator="${operator}"]`).click();
      await page.waitForSelector("#autoFilterMenu", { state: "detached", timeout: 5000 });
      return readDateRows();
    };
    await page.locator('.cell-filter-button[data-filter-col="1"]').click();
    await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    const menu = await page.locator("#autoFilterMenu").innerText();
    const valueLabels = await page.locator("#autoFilterMenu .auto-filter-values .auto-filter-value").evaluateAll((labels) =>
      labels.map((label) => label.textContent.replace(/\s+/g, " ").trim()),
    );
    await page.locator("#autoFilterMenu [data-auto-filter-select-all]").uncheck();
    await page.locator("#autoFilterMenu .auto-filter-value", { hasText: "May 20, 2026" }).locator("[data-auto-filter-value]").check();
    await page.locator('#autoFilterMenu [data-auto-filter-action="ok"]').click();
    await page.waitForSelector("#autoFilterMenu", { state: "detached", timeout: 5000 });
    const valueSelectionRows = await readDateRows();
    await page.locator('.cell-filter-button[data-filter-col="1"]').click();
    await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await page.locator('[data-auto-filter-action="condition-menu"]').click();
    await page.waitForSelector(".auto-filter-submenu:not(.hidden)", { timeout: 5000 });
    const submenu = await page.locator(".auto-filter-submenu:not(.hidden)").innerText();
    await page.locator('[data-auto-filter-operator="greaterThan"]').click();
    await page.waitForSelector('[data-data-dialog="auto-filter-custom"]', { timeout: 5000 });
    await page.locator('[data-data-field="value1"]').fill("2026/05/22");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="1"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const rows = await readDateRows();
    const thisMonthRows = await applyDatePreset("dateThisMonth");
    const nextMonthRows = await applyDatePreset("dateNextMonth");
    const thisYearRows = await applyDatePreset("dateThisYear");
    const yearToDateRows = await applyDatePreset("dateYearToDate");
    const allMayRows = await applyDatePreset("datePeriodM5");
    const secondQuarterRows = await applyDatePreset("datePeriodQ2");
    const periodExported = await exportedAutoFilterDateGroupWorkbookState(
      await page.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
    );
    return { menu, submenu, valueLabels, valueSelectionRows, rows, thisMonthRows, nextMonthRows, thisYearRows, yearToDateRows, allMayRows, secondQuarterRows, periodExported };
  } finally {
    await page.close();
  }
}



async function exerciseAdvancedFilterUniqueDates(context) {
  const workbookPath = await createAdvancedFilterUniqueWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?advancedUnique=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="E2"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="C5"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-advanced-filter"]').click();
    await page.waitForSelector('[data-data-dialog="advanced-filter"]', { timeout: 5000 });
    await page.locator("[data-data-field='action'][value='copy']").check();
    await page.locator("[data-data-field='criteriaRange']").fill("E1:E2");
    await page.locator("[data-data-field='copyToRange']").fill("H1");
    await page.locator("[data-data-field='uniqueOnly']").check();
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="advanced-filter"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="H2"] .cell-content')?.textContent === "Alpha",
      null,
      { timeout: 5000 },
    );
    const values = await page.evaluate(() =>
      ["H1", "I1", "J1", "H2", "I2", "J2", "H3", "I3", "J3", "H4", "I4", "J4"].map((address) =>
        document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
      ),
    );
    const status = await page.locator("#statusText").innerText();
    return { values, status };
  } finally {
    await page.close();
  }
}



async function exerciseAdvancedFilterUniqueSubsetCopy(context) {
  const workbookPath = await createAdvancedFilterUniqueSubsetWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?advancedSubset=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="F1"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="B5"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-advanced-filter"]').click();
    await page.waitForSelector('[data-data-dialog="advanced-filter"]', { timeout: 5000 });
    await page.locator("[data-data-field='action'][value='copy']").check();
    await page.locator("[data-data-field='listRange']").fill("Subset!A1:B5");
    await page.locator("[data-data-field='criteriaRange']").fill("");
    await page.locator("[data-data-field='copyToRange']").fill("Subset!F1");
    await page.locator("[data-data-field='uniqueOnly']").check();
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="advanced-filter"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="F2"] .cell-content')?.textContent === "Alpha" &&
        document.querySelector('.sheet-cell[data-address="F3"] .cell-content')?.textContent === "Beta",
      null,
      { timeout: 5000 },
    );
    const readOutput = () =>
      page.evaluate(() => ({
        values: ["F1", "F2", "F3", "F4"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
        activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent.trim() || "",
        status: document.querySelector("#statusText")?.textContent || "",
      }));
    const applied = await readOutput();
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="F2"] .cell-content')?.textContent !== "Alpha" &&
        document.querySelector("#statusText")?.textContent?.includes("元に戻しました"),
      null,
      { timeout: 5000 },
    );
    const undo = await readOutput();
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="F2"] .cell-content')?.textContent === "Alpha" &&
        document.querySelector('.sheet-cell[data-address="F3"] .cell-content')?.textContent === "Beta" &&
        document.querySelector("#statusText")?.textContent?.includes("やり直しました"),
      null,
      { timeout: 5000 },
    );
    const redo = await readOutput();
    return { applied, undo, redo };
  } finally {
    await page.close();
  }
}



async function exerciseDelimitedTextImportOverwrite(context) {
  const { workbookPath, csvPath, textCsvPath, transformCsvPath, utf16CsvPath, shiftJisCsvPath, jsonPath, xmlPath, workbookDataPath, folderCsvPath } = await createDelimitedImportWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?textImport=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="B2"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('[data-ribbon-tab="data"]').click();
    let chooserPromise = page.waitForEvent("filechooser");
    await page.locator('[data-ribbon-command="data-get-text"]').click();
    let chooser = await chooserPromise;
    await chooser.setFiles(csvPath);
    await page.waitForSelector('[data-data-dialog="text-csv-import"]', { timeout: 5000 });
    const initialDialog = await page.evaluate(() => ({
      title: document.querySelector("#textCsvImportTitle")?.textContent || "",
      delimiter: document.querySelector("[data-data-field='delimiter']")?.value || "",
      origin: document.querySelector("[data-data-field='origin']")?.value || "",
      preview: document.querySelector(".text-csv-preview")?.textContent.replace(/\s+/g, "") || "",
      transformVisible: Boolean(document.querySelector("[data-data-dialog-action='transform']")),
    }));
    await page.locator('[data-data-dialog-action="load"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const overwritePrompt = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      buttons: [...document.querySelectorAll("#messageBoxOverlay [data-message-box-close]")].map((button) =>
        button.textContent.replace(/\s+/g, " ").trim(),
      ),
    }));
    await page.locator("#messageBoxOverlay [data-message-box-result='cancel']").click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    const afterCancel = await page.evaluate(() => ({
      a1: document.querySelector('.sheet-cell[data-address="A1"] .cell-content')?.textContent || "",
      b2: document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent || "",
      dialogVisible: Boolean(document.querySelector('[data-data-dialog="text-csv-import"]')),
    }));
    await page.locator('[data-data-dialog-action="cancel"]').click();
    await page.waitForSelector('[data-data-dialog="text-csv-import"]', { state: "detached", timeout: 5000 });
    chooserPromise = page.waitForEvent("filechooser");
    await page.locator('[data-ribbon-command="data-get-text"]').click();
    chooser = await chooserPromise;
    await chooser.setFiles(csvPath);
    await page.waitForSelector('[data-data-dialog="text-csv-import"]', { timeout: 5000 });
    await page.locator("[data-data-field='delimiter']").selectOption(",");
    const commaPreview = await page.evaluate(() => document.querySelector(".text-csv-preview")?.textContent.replace(/\s+/g, "") || "");
    await page.locator('[data-data-dialog-action="load"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    await page.locator("#messageBoxOverlay [data-message-box-result='ok']").click();
    await page.waitForSelector('[data-data-dialog="text-csv-import"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="A1"] .cell-content')?.textContent === "Name",
      null,
      { timeout: 5000 },
    );
    const values = await page.evaluate(() =>
      ["A1", "B1", "A2", "B2", "A3", "B3"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
    );
    const status = await page.locator("#statusText").innerText();
    await page.locator('.sheet-cell[data-address="D1"]').click();
    chooserPromise = page.waitForEvent("filechooser");
    await page.locator('[data-ribbon-command="data-get-text"]').click();
    chooser = await chooserPromise;
    await chooser.setFiles(textCsvPath);
    await page.waitForSelector('[data-data-dialog="text-csv-import"]', { timeout: 5000 });
    await page.locator("[data-data-field='typeDetection']").selectOption("none");
    await page.locator('[data-data-dialog-action="load"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="D2"] .cell-content')?.textContent === "0012",
      null,
      { timeout: 5000 },
    );
    const typeDetectionNone = await page.evaluate(() => ({
      values: ["D1", "E1", "D2", "E2"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      raw: {
        d2: window.__webSheetDevCellModel?.("Import", 2, 4)?.raw,
        e2: window.__webSheetDevCellModel?.("Import", 2, 5)?.raw,
      },
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.locator('.sheet-cell[data-address="G1"]').click();
    chooserPromise = page.waitForEvent("filechooser");
    await page.locator('[data-ribbon-command="data-get-text"]').click();
    chooser = await chooserPromise;
    await chooser.setFiles(utf16CsvPath);
    await page.waitForSelector('[data-data-dialog="text-csv-import"]', { timeout: 5000 });
    const utf16Initial = await page.evaluate(() => ({
      origin: document.querySelector("[data-data-field='origin']")?.value || "",
      preview: document.querySelector(".text-csv-preview")?.textContent.replace(/\s+/g, "") || "",
      rowCount: document.querySelector("[data-text-csv-row-count]")?.textContent || "",
      delimiterLabel: document.querySelector("[data-data-field='delimiter'] option[value='auto']")?.textContent || "",
    }));
    await page.locator("[data-data-field='origin']").selectOption("65001");
    const utf8Preview = await page.evaluate(() => document.querySelector(".text-csv-preview")?.textContent.replace(/\s+/g, "") || "");
    await page.locator("[data-data-field='origin']").selectOption("1200");
    await page.waitForFunction(
      () => document.querySelector(".text-csv-preview")?.textContent.includes("東京"),
      null,
      { timeout: 5000 },
    );
    const utf16Decoded = await page.evaluate(() => ({
      origin: document.querySelector("[data-data-field='origin']")?.value || "",
      preview: document.querySelector(".text-csv-preview")?.textContent.replace(/\s+/g, "") || "",
      rowCount: document.querySelector("[data-text-csv-row-count]")?.textContent || "",
      delimiterLabel: document.querySelector("[data-data-field='delimiter'] option[value='auto']")?.textContent || "",
    }));
    await page.locator('[data-data-dialog-action="load"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="G2"] .cell-content')?.textContent === "東京",
      null,
      { timeout: 5000 },
    );
    const utf16Import = await page.evaluate(() => ({
      values: ["G1", "H1", "G2", "H2", "G3", "H3"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.locator('.sheet-cell[data-address="J1"]').click();
    chooserPromise = page.waitForEvent("filechooser");
    await page.locator('[data-ribbon-command="data-get-text"]').click();
    chooser = await chooserPromise;
    await chooser.setFiles(shiftJisCsvPath);
    await page.waitForSelector('[data-data-dialog="text-csv-import"]', { timeout: 5000 });
    const shiftJisInitial = await page.evaluate(() => ({
      origin: document.querySelector("[data-data-field='origin']")?.value || "",
      preview: document.querySelector(".text-csv-preview")?.textContent.replace(/\s+/g, "") || "",
      delimiterLabel: document.querySelector("[data-data-field='delimiter'] option[value='auto']")?.textContent || "",
    }));
    await page.locator("[data-data-field='origin']").selectOption("932");
    await page.waitForFunction(
      () => document.querySelector(".text-csv-preview")?.textContent.includes("大阪"),
      null,
      { timeout: 5000 },
    );
    const shiftJisDecoded = await page.evaluate(() => ({
      origin: document.querySelector("[data-data-field='origin']")?.value || "",
      preview: document.querySelector(".text-csv-preview")?.textContent.replace(/\s+/g, "") || "",
      delimiterLabel: document.querySelector("[data-data-field='delimiter'] option[value='auto']")?.textContent || "",
    }));
    await page.locator('[data-data-dialog-action="load"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="J2"] .cell-content')?.textContent === "東京",
      null,
      { timeout: 5000 },
    );
    const shiftJisImport = await page.evaluate(() => ({
      values: ["J1", "K1", "J2", "K2", "J3", "K3"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.locator('.sheet-cell[data-address="M1"]').click();
    chooserPromise = page.waitForEvent("filechooser");
    await page.locator('[data-ribbon-command="data-get-text"]').click();
    chooser = await chooserPromise;
    await chooser.setFiles(transformCsvPath);
    await page.waitForSelector('[data-data-dialog="text-csv-import"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    const transformEditorInitial = await page.evaluate(() => ({
      title: document.querySelector("#queryTransformTitle")?.textContent || "",
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Score"));
    await page.locator("[data-query-transform-column][value='Score']").check();
    await page.locator('[data-data-dialog-action="query-transform-remove-columns"]').click();
    await page.locator("[data-data-field='queryFilterColumn']").selectOption("City");
    await page.locator("[data-data-field='queryFilterText']").fill("Tokyo");
    await page.locator('[data-data-dialog-action="query-transform-apply-filter"]').click();
    const transformEditorAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load"]').first().click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="M2"] .cell-content')?.textContent === "Alpha" &&
        document.querySelector('.sheet-cell[data-address="N3"] .cell-content')?.textContent === "Tokyo",
      null,
      { timeout: 5000 },
    );
    const transformImport = await page.evaluate(() => {
      const queries = window.__webSheetDevQueries?.() || [];
      const query = queries.find((entry) => entry.query?.sourceName === "import-transform-source.csv")?.query || null;
      return {
        values: ["M1", "N1", "M2", "N2", "M3", "N3"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
        query,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });
    await page.locator('.sheet-cell[data-address="M2"]').click();
    await page.locator('[data-ribbon-command="data-refresh-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    await page.locator('#dataCommandMenu [data-data-menu-action="refresh-active"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="M2"] .cell-content')?.textContent === "Alpha" &&
        document.querySelector('.sheet-cell[data-address="M3"] .cell-content')?.textContent === "Gamma",
      null,
      { timeout: 5000 },
    );
    const transformRefresh = await page.evaluate(() => ({
      values: ["M1", "N1", "M2", "N2", "M3", "N3"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      query: (window.__webSheetDevQueries?.() || []).find((entry) => entry.query?.sourceName === "import-transform-source.csv")?.query || null,
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.locator('.sheet-cell[data-address="P1"]').click();
    chooserPromise = page.waitForEvent("filechooser");
    await page.locator('[data-ribbon-command="data-get-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    await page.locator('#dataCommandMenu [data-data-menu-action="get-json"]').click();
    chooser = await chooserPromise;
    await chooser.setFiles(jsonPath);
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="P2"] .cell-content')?.textContent === "JsonAlpha" &&
        document.querySelector('.sheet-cell[data-address="R3"] .cell-content')?.textContent === "Osaka",
      null,
      { timeout: 5000 },
    );
    const jsonFileImport = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .find((entry) => entry.query?.sourceUrl === "import-structured-source.json")?.query || null;
      return {
        values: ["P1", "Q1", "R1", "P2", "Q2", "R2", "P3", "Q3", "R3"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        rawScore: window.__webSheetDevCellModel?.("Import", 3, 17)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });
    await page.locator('.sheet-cell[data-address="T1"]').click();
    chooserPromise = page.waitForEvent("filechooser");
    await page.locator('[data-ribbon-command="data-get-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    await page.locator('#dataCommandMenu [data-data-menu-action="get-xml"]').click();
    chooser = await chooserPromise;
    await chooser.setFiles(xmlPath);
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="U2"] .cell-content')?.textContent === "XmlAlpha" &&
        document.querySelector('.sheet-cell[data-address="V3"] .cell-content')?.textContent === "8",
      null,
      { timeout: 5000 },
    );
    const xmlFileImport = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .find((entry) => entry.query?.sourceUrl === "import-structured-source.xml")?.query || null;
      return {
        values: ["T1", "U1", "V1", "W1", "T2", "U2", "V2", "W2", "T3", "U3", "V3", "W3"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        rawQty: window.__webSheetDevCellModel?.("Import", 3, 22)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });
    await page.evaluate(() => {
      const host = document.querySelector("#sheetHost");
      const horizontal = document.querySelector("#horizontalScrollbar");
      if (host) {
        host.scrollLeft = 0;
        host.scrollTop = 0;
        host.dispatchEvent(new Event("scroll"));
      }
      if (horizontal) {
        horizontal.scrollLeft = 0;
        horizontal.dispatchEvent(new Event("scroll"));
      }
    });
    await page.waitForSelector('.sheet-cell[data-address="A5"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A5"]').click();
    chooserPromise = page.waitForEvent("filechooser");
    await page.locator('[data-ribbon-command="data-get-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    await page.locator('#dataCommandMenu [data-data-menu-action="get-workbook"]').click();
    chooser = await chooserPromise;
    await chooser.setFiles(workbookDataPath);
    await page.waitForSelector('[data-data-dialog="web-navigator"]', { timeout: 5000 });
    const workbookNavigatorInitial = await page.evaluate(() => ({
      sources: [...document.querySelectorAll(".web-navigator-list button span")].map((item) => item.textContent.trim()),
      active: document.querySelector(".web-navigator-list button.is-active span")?.textContent.trim() || "",
      preview: document.querySelector(".web-navigator-preview")?.textContent.replace(/\s+/g, "") || "",
      summary: document.querySelector("[data-web-navigator-summary]")?.textContent.replace(/\s+/g, "") || "",
    }));
    await page.locator('[data-web-source-index="1"]').click();
    const workbookNavigatorSelected = await page.evaluate(() => ({
      active: document.querySelector(".web-navigator-list button.is-active span")?.textContent.trim() || "",
      preview: document.querySelector(".web-navigator-preview")?.textContent.replace(/\s+/g, "") || "",
      summary: document.querySelector("[data-web-navigator-summary]")?.textContent.replace(/\s+/g, "") || "",
    }));
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "City"));
    await page.locator("[data-data-field='queryFilterColumn']").selectOption("City");
    await page.locator("[data-data-field='queryFilterText']").fill("Tokyo");
    await page.locator('[data-data-dialog-action="query-transform-apply-filter"]').click();
    const workbookTransformAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load"]').first().click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent === "BookAlpha" &&
        document.querySelector('.sheet-cell[data-address="C6"] .cell-content')?.textContent === "Tokyo",
      null,
      { timeout: 5000 },
    );
    const workbookFileImport = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .find((entry) => entry.query?.sourceType === "workbook")?.query || null;
      return {
        values: ["A5", "B5", "C5", "A6", "B6", "C6"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        rawScore: window.__webSheetDevCellModel?.("Import", 6, 2)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });
    await page.locator('.sheet-cell[data-address="D8"]').click();
    chooserPromise = page.waitForEvent("filechooser");
    await page.locator('[data-ribbon-command="data-get-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    const folderMenuLabels = await page.evaluate(() =>
      [...document.querySelectorAll("#dataCommandMenu [data-data-menu-action]")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
    );
    await page.locator('#dataCommandMenu [data-data-menu-action="get-folder"]').click();
    chooser = await chooserPromise;
    await chooser.setFiles([folderCsvPath, workbookDataPath]);
    await page.waitForSelector('[data-data-dialog="web-navigator"]', { timeout: 5000 });
    const folderNavigatorInitial = await page.evaluate(() => ({
      sources: [...document.querySelectorAll(".web-navigator-list button span")].map((item) => item.textContent.trim()),
      active: document.querySelector(".web-navigator-list button.is-active span")?.textContent.trim() || "",
      preview: document.querySelector(".web-navigator-preview")?.textContent.replace(/\s+/g, "") || "",
      summary: document.querySelector("[data-web-navigator-summary]")?.textContent.replace(/\s+/g, "") || "",
    }));
    await page.locator('[data-web-source-index="1"]').click();
    const folderNavigatorSelected = await page.evaluate(() => ({
      active: document.querySelector(".web-navigator-list button.is-active span")?.textContent.trim() || "",
      preview: document.querySelector(".web-navigator-preview")?.textContent.replace(/\s+/g, "") || "",
      summary: document.querySelector("[data-web-navigator-summary]")?.textContent.replace(/\s+/g, "") || "",
    }));
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "City"));
    await page.locator("[data-data-field='queryFilterColumn']").selectOption("City");
    await page.locator("[data-data-field='queryFilterText']").fill("Tokyo");
    await page.locator('[data-data-dialog-action="query-transform-apply-filter"]').click();
    const folderTransformAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load"]').first().click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="D9"] .cell-content')?.textContent === "FolderAlpha" &&
        document.querySelector('.sheet-cell[data-address="F9"] .cell-content')?.textContent === "Tokyo",
      null,
      { timeout: 5000 },
    );
    const folderFileImport = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .find((entry) => entry.query?.sourceType === "folder")?.query || null;
      return {
        values: ["D8", "E8", "F8", "D9", "E9", "F9"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        rawScore: window.__webSheetDevCellModel?.("Import", 9, 5)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });
    await page.locator('.sheet-cell[data-address="D9"]').click();
    await page.locator('[data-ribbon-command="data-refresh-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    await page.locator('#dataCommandMenu [data-data-menu-action="refresh-active"]').click();
    await page.waitForFunction(
      () => (window.__webSheetDevQueries?.() || []).some((entry) => entry.query?.sourceType === "folder" && (entry.query.refreshCount || 0) >= 1),
      null,
      { timeout: 5000 },
    );
    const folderRefresh = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .find((entry) => entry.query?.sourceType === "folder")?.query || null;
      return {
        values: ["D8", "E8", "F8", "D9", "E9", "F9"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });
    await page.locator('[data-ribbon-command="data-get-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    await page.locator('#dataCommandMenu [data-data-menu-action="existing-connections"]').click();
    await page.waitForSelector('[data-data-dialog="existing-connections"]', { timeout: 5000 });
    const folderConnectionDetails = await page.evaluate(() => ({
      options: [...document.querySelectorAll("[data-existing-connection-list] option")].map((option) => option.textContent.trim()),
      details: document.querySelector("[data-existing-connection-details]")?.textContent || "",
    }));
    await page.locator('[data-data-dialog-action="cancel"]').click();
    await page.waitForSelector('[data-data-dialog="existing-connections"]', { state: "detached", timeout: 5000 });
    const odataUrl = `data:application/json,${encodeURIComponent(JSON.stringify({
      "@odata.context": "https://example.test/$metadata#Products",
      value: [
        { Product: "ODataAlpha", Score: 71, City: "Tokyo" },
        { Product: "ODataBeta", Score: 82, City: "Osaka" },
      ],
      Categories: [
        { Category: "AI", Count: 2 },
      ],
    }))}`;
    await page.locator('.sheet-cell[data-address="H8"]').click();
    await page.locator('[data-ribbon-command="data-get-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    const odataMenuLabels = await page.evaluate(() =>
      [...document.querySelectorAll("#dataCommandMenu [data-data-menu-action]")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
    );
    await page.locator('#dataCommandMenu [data-data-menu-action="other-data-sources"]').click();
    await page.waitForSelector("#dataCommandMenu.is-other-data-sources", { timeout: 5000 });
    const otherDataSourcesMenuLabels = await page.evaluate(() =>
      [...document.querySelectorAll("#dataCommandMenu [role='menuitem']")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
    );
    await page.locator('#dataCommandMenu [data-data-menu-action="get-odata"]').click();
    await page.waitForSelector('[data-data-dialog="odata-import"]', { timeout: 5000 });
    await page.locator("[data-data-dialog='odata-import'] [data-data-field='sourceUrl']").fill(odataUrl);
    const odataDialogInitial = await page.evaluate(() => ({
      title: document.querySelector("[data-data-dialog='odata-import'] .excel-mini-titlebar")?.textContent.trim() || "",
      placeholder: document.querySelector("[data-data-dialog='odata-import'] [data-data-field='sourceUrl']")?.getAttribute("placeholder") || "",
      destination: document.querySelector("[data-data-dialog='odata-import'] .data-dialog-note")?.textContent || "",
      actions: [...document.querySelectorAll("[data-data-dialog='odata-import'] [data-data-dialog-action]")].map((button) => button.textContent.trim()),
    }));
    await page.locator("[data-data-dialog='odata-import'] [data-data-dialog-action='transform']").click();
    await page.waitForSelector('[data-data-dialog="web-navigator"]', { timeout: 5000 });
    const odataNavigatorInitial = await page.evaluate(() => ({
      sources: [...document.querySelectorAll(".web-navigator-list button span")].map((item) => item.textContent.trim()),
      active: document.querySelector(".web-navigator-list button.is-active span")?.textContent.trim() || "",
      preview: document.querySelector(".web-navigator-preview")?.textContent.replace(/\s+/g, "") || "",
      summary: document.querySelector("[data-web-navigator-summary]")?.textContent.replace(/\s+/g, "") || "",
    }));
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "City"));
    await page.locator("[data-data-field='queryFilterColumn']").selectOption("City");
    await page.locator("[data-data-field='queryFilterText']").fill("Tokyo");
    await page.locator('[data-data-dialog-action="query-transform-apply-filter"]').click();
    const odataTransformAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load"]').first().click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="H9"] .cell-content')?.textContent === "ODataAlpha" &&
        document.querySelector('.sheet-cell[data-address="J9"] .cell-content')?.textContent === "Tokyo",
      null,
      { timeout: 5000 },
    );
    const odataImport = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .find((entry) => entry.query?.sourceType === "odata")?.query || null;
      return {
        values: ["H8", "I8", "J8", "H9", "I9", "J9"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        rawScore: window.__webSheetDevCellModel?.("Import", 9, 9)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });
    await page.locator('.sheet-cell[data-address="H9"]').click();
    await page.locator('[data-ribbon-command="data-refresh-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    await page.locator('#dataCommandMenu [data-data-menu-action="refresh-active"]').click();
    await page.waitForFunction(
      () => (window.__webSheetDevQueries?.() || []).some((entry) => entry.query?.sourceType === "odata" && (entry.query.refreshCount || 0) >= 1),
      null,
      { timeout: 5000 },
    );
    const odataRefresh = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .find((entry) => entry.query?.sourceType === "odata")?.query || null;
      return {
        values: ["H8", "I8", "J8", "H9", "I9", "J9"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });
    await page.locator('[data-ribbon-command="data-get-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    await page.locator('#dataCommandMenu [data-data-menu-action="existing-connections"]').click();
    await page.waitForSelector('[data-data-dialog="existing-connections"]', { timeout: 5000 });
    const odataConnectionDetails = await page.evaluate(() => ({
      options: [...document.querySelectorAll("[data-existing-connection-list] option")].map((option) => option.textContent.trim()),
      details: document.querySelector("[data-existing-connection-details]")?.textContent || "",
    }));
    await page.locator('[data-data-dialog-action="cancel"]').click();
    await page.waitForSelector('[data-data-dialog="existing-connections"]', { state: "detached", timeout: 5000 });
    return {
      initialDialog,
      overwritePrompt,
      afterCancel,
      commaPreview,
      values,
      status,
      typeDetectionNone,
      utf16Initial,
      utf8Preview,
      utf16Decoded,
      utf16Import,
      shiftJisInitial,
      shiftJisDecoded,
      shiftJisImport,
      transformEditorInitial,
      transformEditorAfterSteps,
      transformImport,
      transformRefresh,
      jsonFileImport,
      xmlFileImport,
      workbookNavigatorInitial,
      workbookNavigatorSelected,
      workbookTransformAfterSteps,
      workbookFileImport,
      folderMenuLabels,
      folderNavigatorInitial,
      folderNavigatorSelected,
      folderTransformAfterSteps,
      folderFileImport,
      folderRefresh,
      folderConnectionDetails,
      odataMenuLabels,
      otherDataSourcesMenuLabels,
      odataDialogInitial,
      odataNavigatorInitial,
      odataTransformAfterSteps,
      odataImport,
      odataRefresh,
      odataConnectionDetails,
    };
  } finally {
    await page.close();
  }
}

export {
  autoFilterVisibleRowsSnapshot,
  editCellWithFormulaBar,
  exerciseAdvancedFilterCopyOverwrite,
  exerciseAdvancedFilterUniqueDates,
  exerciseAdvancedFilterUniqueSubsetCopy,
  exerciseAutoFilterBlankAndKeyboardMenu,
  exerciseAutoFilterDateMenu,
  exerciseAutoFilterReapplyAndClear,
  exerciseDelimitedTextImportOverwrite,
  importedFilterStructureShiftState,
  importedStructureReferenceShiftState,
  importedStructuredReferenceState,
  importedTableExpansionState,
  importedTableMetadataState,
  importedTableRowInsertState,
  importedTableTotalsStructureState,
};
