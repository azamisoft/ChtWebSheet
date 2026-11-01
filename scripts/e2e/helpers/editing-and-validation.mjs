// Sequential E2E feature module.
import {
  ExcelJS,
  JSZip,
  appUrl,
  mkdtemp,
  os,
  path,
  workbookPath,
  writeFile,
} from "../../e2e-real-upload.mjs";
import {
  createAutoFilterMenuWorkbook,
  createHomeToolbarWorkbook,
  createImportedAutoFilterWorkbook,
  createValidationMenuWorkbook,
  waitForUploadedWorkbook,
} from "./fixtures-and-model.mjs";
import {
  editCellWithFormulaBar,
  importedFilterStructureShiftState,
  importedStructureReferenceShiftState,
  importedStructuredReferenceState,
  importedTableExpansionState,
  importedTableMetadataState,
  importedTableRowInsertState,
  importedTableTotalsStructureState,
} from "./views-and-macros.mjs";


async function createSortDateWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-sort-date-"));
  const outputPath = path.join(dir, "sort-date.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("SortDate");
  sheet.addRow(["Date", "Label"]);
  sheet.addRow(["06/01/2026", "Jun"]);
  sheet.addRow(["2025/12/31", "Prev"]);
  sheet.addRow(["2026/05/20", "May20"]);
  sheet.addRow(["05/23/2026", "May23"]);
  sheet.addRow(["", "Blank"]);
  sheet.getCell("C1").value = "DateValue";
  sheet.getCell("D1").value = "Label";
  sheet.getCell("C2").value = new Date(Date.UTC(2026, 5, 1));
  sheet.getCell("C2").numFmt = 'mm/dd/yyyy;@';
  sheet.getCell("D2").value = "Jun";
  sheet.getCell("C3").value = new Date(Date.UTC(2026, 4, 20));
  sheet.getCell("C3").numFmt = 'yyyy/mm/dd;@';
  sheet.getCell("D3").value = "May20";
  sheet.getCell("C4").value = new Date(Date.UTC(2025, 11, 31));
  sheet.getCell("C4").numFmt = 'yyyy/mm/dd;@';
  sheet.getCell("D4").value = "Prev";
  sheet.getCell("D5").value = "Blank";
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createSortFormulaWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-sort-formula-"));
  const outputPath = path.join(dir, "sort-formula.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("SortFormula");
  sheet.addRow(["Name", "Calc"]);
  sheet.addRow(["TenA", { formula: "5+5" }]);
  sheet.addRow(["One", { formula: "2-1" }]);
  sheet.addRow(["Three", { formula: "1+2" }]);
  sheet.addRow(["TenB", { formula: "8+2" }]);
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createSortCustomListWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-sort-custom-list-"));
  const outputPath = path.join(dir, "sort-custom-list.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("SortCustomList");
  sheet.addRow(["Month", "Value", "Weekday", "Code"]);
  sheet.addRow(["3月", "M3", "金", "F"]);
  sheet.addRow(["1月", "M1", "日", "S"]);
  sheet.addRow(["12月", "M12", "月", "M"]);
  sheet.addRow(["2月", "M2", "土", "Sat"]);
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function exerciseDataValidationMenu(context) {
  const validationWorkbookPath = await createValidationMenuWorkbook();
  const validationPage = await context.newPage();
  try {
    await validationPage.goto(`${appUrl}?validationMenu=${Date.now()}`, { waitUntil: "load" });
    await validationPage.setInputFiles("#fileInput", validationWorkbookPath);
    await waitForUploadedWorkbook(validationPage, validationWorkbookPath);
    await validationPage.waitForSelector('.sheet-cell[data-address="B1"]', { timeout: 10000 });
    await validationPage.locator('[data-ribbon-tab="data"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    const menu = await validationPage.evaluate(() => ({
      labels: [...document.querySelectorAll("#dataValidationMenu [data-validation-menu-action]")].map((item) =>
        item.textContent.replace(/\s+/g, " ").trim(),
      ),
      items: [...document.querySelectorAll("#dataValidationMenu [data-validation-menu-action]")].map((item) => {
        const rect = item.getBoundingClientRect();
        return {
          action: item.dataset.validationMenuAction || "",
          disabled: item.disabled,
          top: Math.round(rect.top),
          left: Math.round(rect.left),
        };
      }),
      accessKeys: [...document.querySelectorAll("#dataValidationMenu .menu-access-key")].map((item) => item.textContent),
      focusedBefore: document.activeElement?.dataset?.validationMenuAction || "",
    }));
    await validationPage.keyboard.press("ArrowDown");
    Object.assign(menu, await validationPage.evaluate(() => ({
      focusedAfterDown: document.activeElement?.dataset?.validationMenuAction || "",
    })));
    await validationPage.locator('[data-validation-menu-action="circle-invalid"]').click();
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B1"]')?.classList.contains("has-validation-invalid"),
      null,
      { timeout: 5000 },
    );
    const marked = await validationPage.evaluate(() => ({
      b1Class: document.querySelector('.sheet-cell[data-address="B1"]')?.className || "",
      a1Class: document.querySelector('.sheet-cell[data-address="A1"]')?.className || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="clear-circles"]').click();
    await validationPage.waitForFunction(
      () => !document.querySelector('.sheet-cell[data-address="B1"]')?.classList.contains("has-validation-invalid"),
      null,
      { timeout: 5000 },
    );
    const cleared = await validationPage.evaluate(() => ({
      b1Class: document.querySelector('.sheet-cell[data-address="B1"]')?.className || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await validationPage.locator('.sheet-cell[data-address="B1"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    const dialogInitial = await validationPage.evaluate(() => ({
      allow: document.querySelector("[data-data-field='allow']")?.value || "",
      operatorHidden: document.querySelector("[data-validation-row='operator']")?.classList.contains("hidden") || false,
      sourceCaption: document.querySelector("[data-validation-caption='source']")?.textContent || "",
      source2Hidden: document.querySelector("[data-validation-row='source2']")?.classList.contains("hidden") || false,
      dropdownHidden: document.querySelector("[data-validation-row='dropdown']")?.classList.contains("hidden") || false,
    }));
    await validationPage.locator("[data-data-field='allow']").selectOption("list");
    const dialogListMode = await validationPage.evaluate(() => ({
      operatorHidden: document.querySelector("[data-validation-row='operator']")?.classList.contains("hidden") || false,
      sourceCaption: document.querySelector("[data-validation-caption='source']")?.textContent || "",
      source2Hidden: document.querySelector("[data-validation-row='source2']")?.classList.contains("hidden") || false,
      dropdownHidden: document.querySelector("[data-validation-row='dropdown']")?.classList.contains("hidden") || false,
    }));
    await validationPage.locator('[data-data-dialog-action="clear"]').click();
    const dialogAfterClearAll = await validationPage.evaluate(() => ({
      visible: Boolean(document.querySelector('[data-data-dialog="validation"]')),
      allow: document.querySelector("[data-data-field='allow']")?.value || "",
      operatorHidden: document.querySelector("[data-validation-row='operator']")?.classList.contains("hidden") || false,
      sourceHidden: document.querySelector("[data-validation-row='source']")?.classList.contains("hidden") || false,
      source: document.querySelector("[data-data-field='source']")?.value || "",
      errorTitle: document.querySelector("[data-data-field='errorTitle']")?.value || "",
    }));
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="circle-invalid"]').click();
    const afterClearAllMark = await validationPage.evaluate(() => ({
      b1Class: document.querySelector('.sheet-cell[data-address="B1"]')?.className || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await validationPage.locator('.sheet-cell[data-address="B1"]').click();
    await validationPage.locator('.sheet-cell[data-address="B2"]').click({ modifiers: ["Shift"] });
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("custom");
    await validationPage.locator("[data-data-field='source']").fill("=LEN(B1)<=3");
    await validationPage.locator("[data-validation-tab='error']").click();
    await validationPage.locator("[data-data-field='errorTitle']").fill("入力エラー");
    await validationPage.locator("[data-data-field='errorMessage']").fill("3文字以内で入力してください。");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.locator('.sheet-cell[data-address="B2"]').click();
    await validationPage.locator("#formulaInput").click();
    await validationPage.keyboard.type("abcd");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForSelector("#messageBox", { timeout: 5000 });
    const customInvalid = await validationPage.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      buttons: [...document.querySelectorAll("[data-message-box-result]")].map((button) => button.textContent.trim()),
      b2: document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent || "",
    }));
    await validationPage.locator('[data-message-box-result="retry"]').click();
    await validationPage.waitForFunction(
      () =>
        document.querySelector("#formulaInput") === document.activeElement &&
        document.querySelector("#formulaInput")?.value === "abcd" &&
        document.querySelector('.sheet-cell[data-address="B2"]')?.classList.contains("is-formula-editing"),
      null,
      { timeout: 5000 },
    );
    const customRetry = await validationPage.evaluate(() => ({
      formula: document.querySelector("#formulaInput")?.value || "",
      editing: document.querySelector('.sheet-cell[data-address="B2"]')?.classList.contains("is-formula-editing") || false,
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await validationPage.locator("#formulaInput").fill("abc");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent === "abc",
      null,
      { timeout: 5000 },
    );
    const customValid = await validationPage.evaluate(() => ({
      b2: document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await validationPage.locator('.sheet-cell[data-address="B3"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("list");
    await validationPage.locator("[data-data-field='source']").fill("=$D$1:$D$3");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.waitForFunction(
      () =>
        [...document.querySelectorAll('.sheet-cell[data-address="B3"] select.cell-select option')]
          .map((option) => option.textContent || "")
          .includes("Beta"),
      null,
      { timeout: 5000 },
    );
    const listReferenceBefore = await validationPage.evaluate(() => ({
      options: [...document.querySelectorAll('.sheet-cell[data-address="B3"] select.cell-select option')]
        .map((option) => option.textContent || "")
        .filter(Boolean),
    }));
    await validationPage.locator('.sheet-cell[data-address="B3"] select.cell-select').selectOption("Beta");
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B3"] select.cell-select')?.value === "Beta",
      null,
      { timeout: 5000 },
    );
    const listReferenceAfter = await validationPage.evaluate(() => ({
      value: document.querySelector('.sheet-cell[data-address="B3"] select.cell-select')?.value || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await validationPage.locator("#formulaInput").click();
    await validationPage.locator("#formulaInput").fill("beta");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B3"] select.cell-select')?.value === "beta" && !document.querySelector("#messageBox"),
      null,
      { timeout: 5000 },
    );
    const listCaseInsensitive = await validationPage.evaluate(() => ({
      value: document.querySelector('.sheet-cell[data-address="B3"] select.cell-select')?.value || "",
      messageVisible: Boolean(document.querySelector("#messageBox")),
    }));
    await validationPage.locator('.sheet-cell[data-address="B3"] select.cell-select').selectOption("Beta");
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B3"] select.cell-select')?.value === "Beta",
      null,
      { timeout: 5000 },
    );
    await validationPage.locator('.sheet-cell[data-address="B3"]').click({ button: "right" });
    await validationPage.waitForSelector("#cellContextMenu", { timeout: 5000 });
    const contextDropdownMenu = await validationPage.evaluate(() => ({
      disabled: Boolean(document.querySelector("[data-cell-menu-action='select-dropdown']")?.disabled),
      label: document.querySelector("[data-cell-menu-action='select-dropdown']")?.textContent.replace(/\s+/g, " ").trim() || "",
    }));
    await validationPage.locator("[data-cell-menu-action='select-dropdown']").click();
    await validationPage.waitForSelector("#validationDropdownPopup", { timeout: 5000 });
    const contextDropdownPopup = await validationPage.evaluate(() => ({
      row: document.querySelector("#validationDropdownPopup")?.dataset.row || "",
      col: document.querySelector("#validationDropdownPopup")?.dataset.col || "",
      options: [...document.querySelectorAll("[data-validation-dropdown-choice]")].map((button) => button.textContent.trim()),
      selected: document.querySelector("[data-validation-dropdown-choice].is-selected")?.textContent.trim() || "",
    }));
    await validationPage.locator('[data-validation-dropdown-choice="Gamma"]').click();
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B3"] select.cell-select')?.value === "Gamma",
      null,
      { timeout: 5000 },
    );
    const contextDropdownAfter = await validationPage.evaluate(() => ({
      value: document.querySelector('.sheet-cell[data-address="B3"] select.cell-select')?.value || "",
      popupClosed: !document.querySelector("#validationDropdownPopup"),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await validationPage.keyboard.press("Alt+ArrowDown");
    await validationPage.waitForSelector("#validationDropdownPopup", { timeout: 5000 });
    await validationPage.keyboard.press("ArrowUp");
    const contextDropdownKeyboardPopup = await validationPage.evaluate(() => ({
      selected: document.querySelector("[data-validation-dropdown-choice].is-selected")?.textContent.trim() || "",
      focused: document.activeElement?.textContent?.trim() || "",
      optionCount: document.querySelectorAll("[data-validation-dropdown-choice]").length,
    }));
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B3"] select.cell-select')?.value === "Beta",
      null,
      { timeout: 5000 },
    );
    const contextDropdownKeyboard = await validationPage.evaluate(() => ({
      value: document.querySelector('.sheet-cell[data-address="B3"] select.cell-select')?.value || "",
      popupClosed: !document.querySelector("#validationDropdownPopup"),
      status: document.querySelector("#statusText")?.textContent || "",
      selectionLabel: document.querySelector("#selectedCellLabel")?.textContent || "",
    }));
    await validationPage.locator('.sheet-cell[data-address="C3"]').click({ button: "right" });
    await validationPage.waitForSelector("#cellContextMenu", { timeout: 5000 });
    const contextDropdownDisabled = await validationPage.evaluate(() => Boolean(document.querySelector("[data-cell-menu-action='select-dropdown']")?.disabled));
    await validationPage.keyboard.press("Escape");
    await validationPage.locator('.sheet-cell[data-address="B4"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("whole");
    await validationPage.locator("[data-data-field='source']").fill("1");
    await validationPage.locator("[data-data-field='source2']").fill("10");
    await validationPage.locator("[data-validation-tab='error']").click();
    await validationPage.locator("[data-data-field='errorStyle']").selectOption("warning");
    await validationPage.locator("[data-data-field='errorTitle']").fill("注意");
    await validationPage.locator("[data-data-field='errorMessage']").fill("この値を続行しますか。");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.locator("#formulaInput").click();
    await validationPage.keyboard.type("99");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForSelector("#messageBox", { timeout: 5000 });
    const warningInvalid = await validationPage.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      buttons: [...document.querySelectorAll("[data-message-box-result]")].map((button) => button.textContent.trim()),
      b4Before: document.querySelector('.sheet-cell[data-address="B4"] .cell-content')?.textContent || "",
    }));
    await validationPage.locator('[data-message-box-result="yes"]').click();
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B4"] .cell-content')?.textContent === "99",
      null,
      { timeout: 5000 },
    );
    const warningAccepted = await validationPage.evaluate(() => ({
      b4: document.querySelector('.sheet-cell[data-address="B4"] .cell-content')?.textContent || "",
    }));
    await validationPage.locator('.sheet-cell[data-address="B5"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("whole");
    await validationPage.locator("[data-data-field='source']").fill("1");
    await validationPage.locator("[data-data-field='source2']").fill("10");
    await validationPage.locator("[data-validation-tab='error']").click();
    await validationPage.locator("[data-data-field='errorStyle']").selectOption("information");
    await validationPage.locator("[data-data-field='errorTitle']").fill("情報");
    await validationPage.locator("[data-data-field='errorMessage']").fill("入力値を確認してください。");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.locator("#formulaInput").click();
    await validationPage.keyboard.type("99");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForSelector("#messageBox", { timeout: 5000 });
    const informationInvalid = await validationPage.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      buttons: [...document.querySelectorAll("[data-message-box-result]")].map((button) => button.textContent.trim()),
      b5Before: document.querySelector('.sheet-cell[data-address="B5"] .cell-content')?.textContent || "",
    }));
    await validationPage.locator('[data-message-box-result="ok"]').click();
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B5"] .cell-content')?.textContent === "99",
      null,
      { timeout: 5000 },
    );
    const informationAccepted = await validationPage.evaluate(() => ({
      b5: document.querySelector('.sheet-cell[data-address="B5"] .cell-content')?.textContent || "",
    }));
    await validationPage.locator('.sheet-cell[data-address="B6"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("whole");
    await validationPage.locator("[data-data-field='source']").fill("1");
    await validationPage.locator("[data-data-field='source2']").fill("10");
    await validationPage.locator("[data-validation-tab='error']").click();
    await validationPage.locator("[data-data-field='showErrorMessage']").uncheck();
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.locator("#formulaInput").click();
    await validationPage.keyboard.type("99");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B6"] .cell-content')?.textContent === "99" && !document.querySelector("#messageBox"),
      null,
      { timeout: 5000 },
    );
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="circle-invalid"]').click();
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B6"]')?.classList.contains("has-validation-invalid"),
      null,
      { timeout: 5000 },
    );
    const noAlertAccepted = await validationPage.evaluate(() => ({
      b6: document.querySelector('.sheet-cell[data-address="B6"] .cell-content')?.textContent || "",
      b6Class: document.querySelector('.sheet-cell[data-address="B6"]')?.className || "",
    }));
    await validationPage.locator('.sheet-cell[data-address="B7"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("whole");
    await validationPage.locator("[data-data-field='source']").fill("1");
    await validationPage.locator("[data-data-field='source2']").fill("10");
    await validationPage.locator("[data-validation-tab='input']").click();
    await validationPage.locator("[data-data-field='showInputMessage']").check();
    await validationPage.locator("[data-data-field='inputTitle']").fill("入力案内");
    await validationPage.locator("[data-data-field='inputMessage']").fill("1 から 10 までの整数を入力してください。\n範囲外の値は確認されます。");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector("#cellValidationInputPopup", { timeout: 5000 });
    const inputMessagePopup = await validationPage.evaluate(() => {
      const popup = document.querySelector("#cellValidationInputPopup");
      const cell = document.querySelector('.sheet-cell[data-address="B7"]');
      const host = document.querySelector("#sheetHost");
      const popupRect = popup?.getBoundingClientRect();
      const cellRect = cell?.getBoundingClientRect();
      const hostRect = host?.getBoundingClientRect();
      return {
        visible: Boolean(popup),
        row: popup?.dataset.row || "",
        col: popup?.dataset.col || "",
        title: popup?.querySelector("strong")?.textContent || "",
        text: popup?.innerText || "",
        rightOfCell: Boolean(popupRect && cellRect && popupRect.left >= cellRect.right - 1),
        insideViewport: Boolean(
          popupRect &&
            hostRect &&
            popupRect.left >= hostRect.left &&
            popupRect.right <= hostRect.right + 1 &&
            popupRect.top >= hostRect.top &&
            popupRect.bottom <= hostRect.bottom + 1,
        ),
      };
    });
    await validationPage.locator('.sheet-cell[data-address="C7"]').click();
    await validationPage.waitForFunction(() => !document.querySelector("#cellValidationInputPopup"), null, { timeout: 5000 });
    inputMessagePopup.closedAfterOtherCell = await validationPage.evaluate(() => !document.querySelector("#cellValidationInputPopup"));

    await validationPage.locator('.sheet-cell[data-address="B8"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("date");
    await validationPage.locator("[data-data-field='source']").fill("2026/05/01");
    await validationPage.locator("[data-data-field='source2']").fill("2026/05/31");
    const dateDialog = await validationPage.evaluate(() => ({
      sourceCaption: document.querySelector("[data-validation-caption='source']")?.textContent || "",
      source2Caption: document.querySelector("[data-validation-caption='source2']")?.textContent || "",
    }));
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.locator("#formulaInput").fill("2026/06/01");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForSelector("#messageBox", { timeout: 5000 });
    const dateInvalid = await validationPage.evaluate(() => ({
      buttons: [...document.querySelectorAll("[data-message-box-result]")].map((button) => button.textContent.trim()),
      b8: document.querySelector('.sheet-cell[data-address="B8"] .cell-content')?.textContent || "",
    }));
    await validationPage.locator('[data-message-box-result="cancel"]').click();
    await validationPage.locator("#formulaInput").fill("2026/05/20");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B8"] .cell-content')?.textContent === "2026/05/20",
      null,
      { timeout: 5000 },
    );
    const dateValid = await validationPage.evaluate(() => ({
      b8: document.querySelector('.sheet-cell[data-address="B8"] .cell-content')?.textContent || "",
    }));

    await validationPage.locator('.sheet-cell[data-address="B9"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("time");
    await validationPage.locator("[data-data-field='source']").fill("09:00");
    await validationPage.locator("[data-data-field='source2']").fill("18:00");
    const timeDialog = await validationPage.evaluate(() => ({
      sourceCaption: document.querySelector("[data-validation-caption='source']")?.textContent || "",
      source2Caption: document.querySelector("[data-validation-caption='source2']")?.textContent || "",
    }));
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.locator("#formulaInput").fill("08:30");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForSelector("#messageBox", { timeout: 5000 });
    const timeInvalid = await validationPage.evaluate(() => ({
      buttons: [...document.querySelectorAll("[data-message-box-result]")].map((button) => button.textContent.trim()),
      b9: document.querySelector('.sheet-cell[data-address="B9"] .cell-content')?.textContent || "",
    }));
    await validationPage.locator('[data-message-box-result="cancel"]').click();
    await validationPage.locator("#formulaInput").fill("09:30");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B9"] .cell-content')?.textContent === "09:30",
      null,
      { timeout: 5000 },
    );
    const timeValid = await validationPage.evaluate(() => ({
      b9: document.querySelector('.sheet-cell[data-address="B9"] .cell-content')?.textContent || "",
    }));
    await validationPage.locator('.sheet-cell[data-address="B9"]').click();
    await validationPage.locator("#formulaInput").fill("6 PM");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B9"] .cell-content')?.textContent === "6 PM" && !document.querySelector("#messageBox"),
      null,
      { timeout: 5000 },
    );
    const timeAmPmValid = await validationPage.evaluate(() => ({
      b9: document.querySelector('.sheet-cell[data-address="B9"] .cell-content')?.textContent || "",
      messageVisible: Boolean(document.querySelector("#messageBox")),
    }));

    await validationPage.locator('.sheet-cell[data-address="B10"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("textLength");
    await validationPage.locator("[data-data-field='operator']").selectOption("lessThanOrEqual");
    await validationPage.locator("[data-data-field='source']").fill("5");
    const textLengthDialog = await validationPage.evaluate(() => ({
      sourceCaption: document.querySelector("[data-validation-caption='source']")?.textContent || "",
      source2Hidden: document.querySelector("[data-validation-row='source2']")?.classList.contains("hidden") || false,
    }));
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.locator("#formulaInput").fill("123456");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForSelector("#messageBox", { timeout: 5000 });
    const textLengthInvalid = await validationPage.evaluate(() => ({
      buttons: [...document.querySelectorAll("[data-message-box-result]")].map((button) => button.textContent.trim()),
      b10: document.querySelector('.sheet-cell[data-address="B10"] .cell-content')?.textContent || "",
    }));
    await validationPage.locator('[data-message-box-result="cancel"]').click();
    await validationPage.locator("#formulaInput").fill("12345");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B10"] .cell-content')?.textContent === "12345",
      null,
      { timeout: 5000 },
    );
    const textLengthValid = await validationPage.evaluate(() => ({
      b10: document.querySelector('.sheet-cell[data-address="B10"] .cell-content')?.textContent || "",
    }));

    await validationPage.locator('.sheet-cell[data-address="B11"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("whole");
    await validationPage.locator("[data-data-field='source']").fill("1");
    await validationPage.locator("[data-data-field='source2']").fill("10");
    await validationPage.locator("[data-data-field='allowBlank']").uncheck();
    const blankDialog = await validationPage.evaluate(() => ({
      allowBlank: document.querySelector("[data-data-field='allowBlank']")?.checked || false,
    }));
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.locator("#formulaInput").fill("5");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B11"] .cell-content')?.textContent === "5",
      null,
      { timeout: 5000 },
    );
    await validationPage.locator("#formulaInput").fill("");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForSelector("#messageBox", { timeout: 5000 });
    const blankInvalid = await validationPage.evaluate(() => ({
      buttons: [...document.querySelectorAll("[data-message-box-result]")].map((button) => button.textContent.trim()),
      b11: document.querySelector('.sheet-cell[data-address="B11"] .cell-content')?.textContent || "",
    }));
    await validationPage.locator('[data-message-box-result="cancel"]').click();
    const blankRejected = await validationPage.evaluate(() => ({
      b11: document.querySelector('.sheet-cell[data-address="B11"] .cell-content')?.textContent || "",
    }));

    await validationPage.locator('.sheet-cell[data-address="B12"]').click();
    await validationPage.locator('.sheet-cell[data-address="B13"]').click({ modifiers: ["Shift"] });
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("whole");
    await validationPage.locator("[data-data-field='operator']").selectOption("lessThanOrEqual");
    await validationPage.locator("[data-data-field='source']").fill("=A12");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.locator('.sheet-cell[data-address="B13"]').click();
    await validationPage.locator("#formulaInput").fill("9");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B13"] .cell-content')?.textContent === "9",
      null,
      { timeout: 5000 },
    );
    const formulaBoundValid = await validationPage.evaluate(() => ({
      b13: document.querySelector('.sheet-cell[data-address="B13"] .cell-content')?.textContent || "",
      messageVisible: Boolean(document.querySelector("#messageBox")),
    }));
    await validationPage.locator('.sheet-cell[data-address="B12"]').click();
    await validationPage.locator("#formulaInput").fill("6");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForSelector("#messageBox", { timeout: 5000 });
    const formulaBoundInvalid = await validationPage.evaluate(() => ({
      buttons: [...document.querySelectorAll("[data-message-box-result]")].map((button) => button.textContent.trim()),
      b12: document.querySelector('.sheet-cell[data-address="B12"] .cell-content')?.textContent || "",
    }));
    await validationPage.locator('[data-message-box-result="cancel"]').click();
    await validationPage.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    await validationPage.locator('.sheet-cell[data-address="B14"]').click();
    await validationPage.locator('.sheet-cell[data-address="B15"]').click({ modifiers: ["Shift"] });
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("whole");
    await validationPage.locator("[data-data-field='source']").fill("1");
    await validationPage.locator("[data-data-field='source2']").fill("10");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.locator('.sheet-cell[data-address="B14"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    const sameValidationInitial = await validationPage.evaluate(() => ({
      source2: document.querySelector("[data-data-field='source2']")?.value || "",
      disabled: Boolean(document.querySelector("[data-data-field='applySameValidation']")?.disabled),
      label: document.querySelector("[data-data-field='applySameValidation']")?.closest("label")?.textContent.replace(/\s+/g, " ").trim() || "",
    }));
    await validationPage.locator("[data-data-field='source2']").fill("20");
    await validationPage.locator("[data-data-field='applySameValidation']").check();
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    const sameValidationStatus = await validationPage.locator("#statusText").innerText();
    await validationPage.locator('.sheet-cell[data-address="B15"]').click();
    await validationPage.locator("#formulaInput").fill("15");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B15"] .cell-content')?.textContent === "15",
      null,
      { timeout: 5000 },
    );
    await validationPage.locator('.sheet-cell[data-address="B15"]').click();
    await validationPage.locator("#formulaInput").fill("25");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForSelector("#messageBox", { timeout: 5000 });
    const sameValidationApplied = await validationPage.evaluate((status) => ({
      status,
      b15: document.querySelector('.sheet-cell[data-address="B15"] .cell-content')?.textContent || "",
      buttons: [...document.querySelectorAll("[data-message-box-result]")].map((button) => button.textContent.trim()),
    }), sameValidationStatus);
    await validationPage.locator('[data-message-box-result="cancel"]').click();
    await validationPage.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    await validationPage.locator('.sheet-cell[data-address="B16"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("list");
    await validationPage.locator("[data-data-field='source']").fill("Red;Green;Blue");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.waitForFunction(
      () =>
        [...document.querySelectorAll('.sheet-cell[data-address="B16"] select.cell-select option')]
          .map((option) => option.textContent || "")
          .includes("Green"),
      null,
      { timeout: 5000 },
    );
    await validationPage.locator("#formulaInput").fill("green");
    await validationPage.keyboard.press("Enter");
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B16"] select.cell-select')?.value === "green" && !document.querySelector("#messageBox"),
      null,
      { timeout: 5000 },
    );
    const listSemicolon = await validationPage.evaluate(() => ({
      options: [...document.querySelectorAll('.sheet-cell[data-address="B16"] select.cell-select option')]
        .filter((option) => !option.hidden)
        .map((option) => option.textContent || "")
        .filter(Boolean),
      value: document.querySelector('.sheet-cell[data-address="B16"] select.cell-select')?.value || "",
      messageVisible: Boolean(document.querySelector("#messageBox")),
    }));
    await validationPage.locator('.sheet-cell[data-address="B17"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("list");
    await validationPage.locator("[data-data-field='source']").fill("='入力規則'!候補リスト");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.waitForFunction(
      () =>
        [...document.querySelectorAll('.sheet-cell[data-address="B17"] select.cell-select option')]
          .map((option) => option.textContent || "")
          .includes("Gamma"),
      null,
      { timeout: 5000 },
    );
    await validationPage.keyboard.press("Alt+ArrowDown");
    await validationPage.waitForSelector("#validationDropdownPopup", { timeout: 5000 });
    const listLocalNameReferencePopup = await validationPage.evaluate(() => ({
      options: [...document.querySelectorAll("[data-validation-dropdown-choice]")].map((button) => button.textContent.trim()),
      source: window.__webSheetDevCellModel?.("入力規則", 17, 2)?.validation?.formula || "",
      localName: window.__webSheetDevDefinedNames?.().find((definition) => definition.name === "候補リスト") || null,
    }));
    await validationPage.locator('[data-validation-dropdown-choice="Gamma"]').click();
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B17"] select.cell-select')?.value === "Gamma",
      null,
      { timeout: 5000 },
    );
    const listLocalNameReference = await validationPage.evaluate((popup) => ({
      ...popup,
      value: document.querySelector('.sheet-cell[data-address="B17"] select.cell-select')?.value || "",
      popupClosed: !document.querySelector("#validationDropdownPopup"),
      status: document.querySelector("#statusText")?.textContent || "",
    }), listLocalNameReferencePopup);
    await validationPage.locator('.sheet-cell[data-address="B18"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("list");
    await validationPage.locator("[data-data-field='source']").fill("=動的候補");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.waitForFunction(
      () =>
        [...document.querySelectorAll('.sheet-cell[data-address="B18"] select.cell-select option')]
          .map((option) => option.textContent || "")
          .includes("Gamma"),
      null,
      { timeout: 5000 },
    );
    await validationPage.keyboard.press("Alt+ArrowDown");
    await validationPage.waitForSelector("#validationDropdownPopup", { timeout: 5000 });
    const listDynamicNameReferencePopup = await validationPage.evaluate(() => ({
      options: [...document.querySelectorAll("[data-validation-dropdown-choice]")].map((button) => button.textContent.trim()),
      source: window.__webSheetDevCellModel?.("入力規則", 18, 2)?.validation?.formula || "",
      dynamicName: window.__webSheetDevDefinedNames?.().find((definition) => definition.name === "動的候補") || null,
    }));
    await validationPage.locator('[data-validation-dropdown-choice="Beta"]').click();
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B18"] select.cell-select')?.value === "Beta",
      null,
      { timeout: 5000 },
    );
    const listDynamicNameReference = await validationPage.evaluate((popup) => ({
      ...popup,
      value: document.querySelector('.sheet-cell[data-address="B18"] select.cell-select')?.value || "",
      popupClosed: !document.querySelector("#validationDropdownPopup"),
      status: document.querySelector("#statusText")?.textContent || "",
    }), listDynamicNameReferencePopup);
    await validationPage.locator('.sheet-cell[data-address="B19"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("list");
    await validationPage.locator("[data-data-field='source']").fill("=INDIRECT($E$1)");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.waitForFunction(
      () =>
        [...document.querySelectorAll('.sheet-cell[data-address="B19"] select.cell-select option')]
          .map((option) => option.textContent || "")
          .includes("Gamma"),
      null,
      { timeout: 5000 },
    );
    await validationPage.keyboard.press("Alt+ArrowDown");
    await validationPage.waitForSelector("#validationDropdownPopup", { timeout: 5000 });
    const listIndirectNameReferencePopup = await validationPage.evaluate(() => ({
      options: [...document.querySelectorAll("[data-validation-dropdown-choice]")].map((button) => button.textContent.trim()),
      source: window.__webSheetDevCellModel?.("入力規則", 19, 2)?.validation?.formula || "",
      driverValue: window.__webSheetDevCellModel?.("入力規則", 1, 5)?.raw || "",
    }));
    await validationPage.locator('[data-validation-dropdown-choice="Alpha"]').click();
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B19"] select.cell-select')?.value === "Alpha",
      null,
      { timeout: 5000 },
    );
    const listIndirectNameReference = await validationPage.evaluate((popup) => ({
      ...popup,
      value: document.querySelector('.sheet-cell[data-address="B19"] select.cell-select')?.value || "",
      popupClosed: !document.querySelector("#validationDropdownPopup"),
      status: document.querySelector("#statusText")?.textContent || "",
    }), listIndirectNameReferencePopup);
    await validationPage.locator('.sheet-cell[data-address="B20"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("list");
    await validationPage.locator("[data-data-field='source']").fill("=INDIRECT($E$2)");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.waitForFunction(
      () =>
        [...document.querySelectorAll('.sheet-cell[data-address="B20"] select.cell-select option')]
          .map((option) => option.textContent || "")
          .includes("Gamma"),
      null,
      { timeout: 5000 },
    );
    await validationPage.keyboard.press("Alt+ArrowDown");
    await validationPage.waitForSelector("#validationDropdownPopup", { timeout: 5000 });
    const listIndirectCellRangePopup = await validationPage.evaluate(() => ({
      options: [...document.querySelectorAll("[data-validation-dropdown-choice]")].map((button) => button.textContent.trim()),
      source: window.__webSheetDevCellModel?.("入力規則", 20, 2)?.validation?.formula || "",
      driverValue: window.__webSheetDevCellModel?.("入力規則", 2, 5)?.raw || "",
    }));
    await validationPage.locator('[data-validation-dropdown-choice="Beta"]').click();
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B20"] select.cell-select')?.value === "Beta",
      null,
      { timeout: 5000 },
    );
    const listIndirectCellRange = await validationPage.evaluate((popup) => ({
      ...popup,
      value: document.querySelector('.sheet-cell[data-address="B20"] select.cell-select')?.value || "",
      popupClosed: !document.querySelector("#validationDropdownPopup"),
      status: document.querySelector("#statusText")?.textContent || "",
    }), listIndirectCellRangePopup);
    await validationPage.locator('.sheet-cell[data-address="B21"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("list");
    await validationPage.locator("[data-data-field='source']").fill("=$D$1:INDEX($D:$D,COUNTA($D:$D))");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.waitForFunction(
      () =>
        [...document.querySelectorAll('.sheet-cell[data-address="B21"] select.cell-select option')]
          .map((option) => option.textContent || "")
          .includes("Gamma"),
      null,
      { timeout: 5000 },
    );
    await validationPage.keyboard.press("Alt+ArrowDown");
    await validationPage.waitForSelector("#validationDropdownPopup", { timeout: 5000 });
    const listIndexRangeReferencePopup = await validationPage.evaluate(() => ({
      options: [...document.querySelectorAll("[data-validation-dropdown-choice]")].map((button) => button.textContent.trim()),
      source: window.__webSheetDevCellModel?.("入力規則", 21, 2)?.validation?.formula || "",
    }));
    await validationPage.locator('[data-validation-dropdown-choice="Gamma"]').click();
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B21"] select.cell-select')?.value === "Gamma",
      null,
      { timeout: 5000 },
    );
    const listIndexRangeReference = await validationPage.evaluate((popup) => ({
      ...popup,
      value: document.querySelector('.sheet-cell[data-address="B21"] select.cell-select')?.value || "",
      popupClosed: !document.querySelector("#validationDropdownPopup"),
      status: document.querySelector("#statusText")?.textContent || "",
    }), listIndexRangeReferencePopup);
    await validationPage.locator('.sheet-cell[data-address="B22"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("list");
    await validationPage.locator("[data-data-field='source']").fill("=INDIRECT(\"'入力規則'!$D$1:$D$3\")");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.waitForFunction(
      () =>
        [...document.querySelectorAll('.sheet-cell[data-address="B22"] select.cell-select option')]
          .map((option) => option.textContent || "")
          .includes("Gamma"),
      null,
      { timeout: 5000 },
    );
    await validationPage.keyboard.press("Alt+ArrowDown");
    await validationPage.waitForSelector("#validationDropdownPopup", { timeout: 5000 });
    const listIndirectLiteralRangePopup = await validationPage.evaluate(() => ({
      options: [...document.querySelectorAll("[data-validation-dropdown-choice]")].map((button) => button.textContent.trim()),
      source: window.__webSheetDevCellModel?.("入力規則", 22, 2)?.validation?.formula || "",
    }));
    await validationPage.locator('[data-validation-dropdown-choice="Alpha"]').click();
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B22"] select.cell-select')?.value === "Alpha",
      null,
      { timeout: 5000 },
    );
    const listIndirectLiteralRange = await validationPage.evaluate((popup) => ({
      ...popup,
      value: document.querySelector('.sheet-cell[data-address="B22"] select.cell-select')?.value || "",
      popupClosed: !document.querySelector("#validationDropdownPopup"),
      status: document.querySelector("#statusText")?.textContent || "",
    }), listIndirectLiteralRangePopup);
    await validationPage.locator('.sheet-cell[data-address="B23"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("list");
    await validationPage.locator("[data-data-field='source']").fill("=INDEX($D$1:$D$3,0,1)");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.waitForFunction(
      () =>
        [...document.querySelectorAll('.sheet-cell[data-address="B23"] select.cell-select option')]
          .map((option) => option.textContent || "")
          .includes("Gamma"),
      null,
      { timeout: 5000 },
    );
    await validationPage.keyboard.press("Alt+ArrowDown");
    await validationPage.waitForSelector("#validationDropdownPopup", { timeout: 5000 });
    const listIndexColumnReferencePopup = await validationPage.evaluate(() => ({
      options: [...document.querySelectorAll("[data-validation-dropdown-choice]")].map((button) => button.textContent.trim()),
      source: window.__webSheetDevCellModel?.("入力規則", 23, 2)?.validation?.formula || "",
    }));
    await validationPage.locator('[data-validation-dropdown-choice="Beta"]').click();
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B23"] select.cell-select')?.value === "Beta",
      null,
      { timeout: 5000 },
    );
    const listIndexColumnReference = await validationPage.evaluate((popup) => ({
      ...popup,
      value: document.querySelector('.sheet-cell[data-address="B23"] select.cell-select')?.value || "",
      popupClosed: !document.querySelector("#validationDropdownPopup"),
      status: document.querySelector("#statusText")?.textContent || "",
    }), listIndexColumnReferencePopup);
    await validationPage.locator('.sheet-cell[data-address="B24"]').click();
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    await validationPage.locator("[data-data-field='allow']").selectOption("list");
    await validationPage.locator("[data-data-field='source']").fill("=INDEX($F$2:$G$4,0,MATCH($E$3,$F$1:$G$1,0))");
    await validationPage.locator('[data-data-dialog-action="ok"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    await validationPage.waitForFunction(
      () =>
        [...document.querySelectorAll('.sheet-cell[data-address="B24"] select.cell-select option')]
          .map((option) => option.textContent || "")
          .includes("Yellow"),
      null,
      { timeout: 5000 },
    );
    await validationPage.keyboard.press("Alt+ArrowDown");
    await validationPage.waitForSelector("#validationDropdownPopup", { timeout: 5000 });
    const listIndexMatchReferencePopup = await validationPage.evaluate(() => ({
      options: [...document.querySelectorAll("[data-validation-dropdown-choice]")].map((button) => button.textContent.trim()),
      source: window.__webSheetDevCellModel?.("入力規則", 24, 2)?.validation?.formula || "",
      driverValue: window.__webSheetDevCellModel?.("入力規則", 3, 5)?.raw || "",
    }));
    await validationPage.locator('[data-validation-dropdown-choice="Blue"]').click();
    await validationPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B24"] select.cell-select')?.value === "Blue",
      null,
      { timeout: 5000 },
    );
    const listIndexMatchReference = await validationPage.evaluate((popup) => ({
      ...popup,
      value: document.querySelector('.sheet-cell[data-address="B24"] select.cell-select')?.value || "",
      popupClosed: !document.querySelector("#validationDropdownPopup"),
      status: document.querySelector("#statusText")?.textContent || "",
    }), listIndexMatchReferencePopup);
    await validationPage.locator('.sheet-cell[data-address="B25"]').click();
    await validationPage.waitForFunction(
      () => window.__webSheetDevCellModel?.("入力規則", 25, 2)?.validation?.type === "list",
      null,
      { timeout: 5000 },
    );
    const importedHiddenDropdown = await validationPage.evaluate(() => {
      const cell = window.__webSheetDevCellModel?.("入力規則", 25, 2);
      return {
        value: document.querySelector('.sheet-cell[data-address="B25"] .cell-content')?.textContent || "",
        hasSelect: Boolean(document.querySelector('.sheet-cell[data-address="B25"] select.cell-select')),
        validation: cell?.validation || null,
      };
    });
    const importedBlankRangeValidation = await validationPage.evaluate(() => {
      const sheet = window.__webSheetDevSheetModel?.("入力規則");
      const c30 = window.__webSheetDevCellModel?.("入力規則", 30, 3);
      const d31 = window.__webSheetDevCellModel?.("入力規則", 31, 4);
      const e30 = window.__webSheetDevCellModel?.("入力規則", 30, 5);
      const f30 = window.__webSheetDevCellModel?.("入力規則", 30, 6);
      const h30 = window.__webSheetDevCellModel?.("入力規則", 30, 8);
      const h32 = window.__webSheetDevCellModel?.("入力規則", 32, 8);
      const i33 = window.__webSheetDevCellModel?.("入力規則", 33, 9);
      const j30 = window.__webSheetDevCellModel?.("入力規則", 30, 10);
      const k30 = window.__webSheetDevCellModel?.("入力規則", 30, 11);
      const l30 = window.__webSheetDevCellModel?.("入力規則", 30, 12);
      const m30 = window.__webSheetDevCellModel?.("入力規則", 30, 13);
      return {
        rowCount: sheet?.rowCount || 0,
        colCount: sheet?.colCount || 0,
        c30: c30?.validation || null,
        d31: d31?.validation || null,
        e30: e30?.validation || null,
        f30: f30?.validation || null,
        h30: h30?.validation || null,
        h32: h32?.validation || null,
        i33: i33?.validation || null,
        j30: j30?.validation || null,
        k30: k30?.validation || null,
        l30: l30?.validation || null,
        m30: m30?.validation || null,
      };
    });
    await validationPage.locator('[data-ribbon-command="data-validation"]').click();
    await validationPage.waitForSelector("#dataValidationMenu", { timeout: 5000 });
    await validationPage.locator('[data-validation-menu-action="open"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { timeout: 5000 });
    const importedHiddenDropdownDialog = await validationPage.evaluate(() => ({
      allow: document.querySelector("[data-data-field='allow']")?.value || "",
      showDropDown: document.querySelector("[data-data-field='showDropDown']")?.checked || false,
      dropdownHidden: document.querySelector("[data-validation-row='dropdown']")?.classList.contains("hidden") || false,
    }));
    await validationPage.locator("[data-validation-tab='ime']").click();
    importedHiddenDropdownDialog.imeMode = await validationPage.locator("[data-data-field='imeMode']").inputValue();
    await validationPage.locator('[data-data-dialog-action="cancel"]').click();
    await validationPage.waitForSelector('[data-data-dialog="validation"]', { state: "detached", timeout: 5000 });
    const exportedValidationBase64 = await validationPage.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || "");
    const exportedValidationZip = await JSZip.loadAsync(Buffer.from(exportedValidationBase64, "base64"));
    const exportedValidationXml = await exportedValidationZip.file("xl/worksheets/sheet1.xml")?.async("text");
    const exportedHiddenDropdownTag =
      exportedValidationXml?.match(/<dataValidation\b(?=[^>]*\bsqref="B25")[^>]*>/)?.[0] || "";
    const exportedHiddenDropdown = {
      tag: exportedHiddenDropdownTag,
      showDropDown: /showDropDown="1"/.test(exportedHiddenDropdownTag),
      imeMode: /imeMode="disabled"/.test(exportedHiddenDropdownTag),
    };
    const exportedBlankRangeValidation = {
      hasRange: /sqref="C30:D31"/.test(exportedValidationXml || "") || /sqref="C30"/.test(exportedValidationXml || ""),
      hasFormula: /<formula1>(?:"|&quot;)One,Two(?:"|&quot;)<\/formula1>/.test(exportedValidationXml || ""),
      hasReferenceFormula: /<formula1>A12<\/formula1>/.test(exportedValidationXml || ""),
      hasCustomFormula: /<formula1>LEN\(F30\)&lt;=3<\/formula1>/.test(exportedValidationXml || ""),
      multiRangeTag: exportedValidationXml?.match(/<dataValidation\b(?=[^>]*\bsqref="[^"]*H30[^"]*H32:I33[^"]*")[^>]*>[\s\S]*?<formula1>(?:"|&quot;)East,West(?:"|&quot;)<\/formula1><\/dataValidation>/)?.[0] || "",
      decimalTag: exportedValidationXml?.match(/<dataValidation\b(?=[^>]*\btype="decimal")(?=[^>]*\boperator="greaterThan")(?=[^>]*\bsqref="J30")[^>]*>[\s\S]*?<formula1>1\.5<\/formula1><\/dataValidation>/)?.[0] || "",
      dateTag: exportedValidationXml?.match(/<dataValidation\b(?=[^>]*\btype="date")(?=[^>]*\boperator="between")(?=[^>]*\bsqref="K30")[^>]*>[\s\S]*?<formula1>45413<\/formula1><formula2>45443<\/formula2><\/dataValidation>/)?.[0] || "",
      timeTag: exportedValidationXml?.match(/<dataValidation\b(?=[^>]*\btype="time")(?=[^>]*\boperator="between")(?=[^>]*\bsqref="L30")[^>]*>[\s\S]*?<formula1>0\.375<\/formula1><formula2>0\.75<\/formula2><\/dataValidation>/)?.[0] || "",
      textLengthTag: exportedValidationXml?.match(/<dataValidation\b(?=[^>]*\btype="textLength")(?=[^>]*\boperator="equal")(?=[^>]*\bsqref="M30")[^>]*>[\s\S]*?<formula1>5<\/formula1><\/dataValidation>/)?.[0] || "",
    };
    return {
      ...menu,
      marked,
      cleared,
      dialogInitial,
      dialogListMode,
      dialogAfterClearAll,
      afterClearAllMark,
      customInvalid,
      customRetry,
      customValid,
      listReferenceBefore,
      listReferenceAfter,
      listCaseInsensitive,
      contextDropdownMenu,
      contextDropdownPopup,
      contextDropdownAfter,
      contextDropdownKeyboardPopup,
      contextDropdownKeyboard,
      contextDropdownDisabled,
      warningInvalid,
      warningAccepted,
      informationInvalid,
      informationAccepted,
      noAlertAccepted,
      inputMessagePopup,
      dateDialog,
      dateInvalid,
      dateValid,
      timeDialog,
      timeInvalid,
      timeValid,
      timeAmPmValid,
      textLengthDialog,
      textLengthInvalid,
      textLengthValid,
      blankDialog,
      blankInvalid,
      blankRejected,
      formulaBoundValid,
      formulaBoundInvalid,
      sameValidationInitial,
      sameValidationApplied,
      listSemicolon,
      listLocalNameReference,
      listDynamicNameReference,
      listIndirectNameReference,
      listIndirectCellRange,
      listIndexRangeReference,
      listIndirectLiteralRange,
      listIndexColumnReference,
      listIndexMatchReference,
      importedHiddenDropdown,
      importedHiddenDropdownDialog,
      exportedHiddenDropdown,
      importedBlankRangeValidation,
      exportedBlankRangeValidation,
    };
  } finally {
    await validationPage.close();
  }
}



async function exerciseAutoFilterConditionMenu(context) {
  const filterWorkbookPath = await createAutoFilterMenuWorkbook({ includeMixedTextScore: true });
  const filterPage = await context.newPage();
  try {
    await filterPage.goto(`${appUrl}?autoFilterMenu=${Date.now()}`, { waitUntil: "load" });
    await filterPage.setInputFiles("#fileInput", filterWorkbookPath);
    await waitForUploadedWorkbook(filterPage, filterWorkbookPath);
    await filterPage.waitForSelector('.sheet-cell[data-address="B6"]', { timeout: 10000 });
    await filterPage.locator('.sheet-cell[data-address="A1"]').click();
    await filterPage.locator('.sheet-cell[data-address="B6"]').click({ modifiers: ["Shift"] });
    await filterPage.locator('[data-ribbon-tab="data"]').click();
    await filterPage.locator('[data-ribbon-command="data-filter"]').click();
    await filterPage.waitForSelector('.cell-filter-button[data-filter-col="2"]', { timeout: 5000 });
    await filterPage.locator('.cell-filter-button[data-filter-col="1"]').click();
    await filterPage.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    const textMenu = await filterPage.locator("#autoFilterMenu").innerText();
    await filterPage.locator('[data-auto-filter-action="condition-menu"]').click();
    await filterPage.waitForSelector(".auto-filter-submenu:not(.hidden)", { timeout: 5000 });
    const textSubmenu = await filterPage.locator(".auto-filter-submenu:not(.hidden)").innerText();
    await filterPage.locator('[data-auto-filter-operator="beginsWith"]').click();
    await filterPage.waitForSelector('[data-data-dialog="auto-filter-custom"]', { timeout: 5000 });
    await filterPage.locator('[data-data-field="value1"]').fill("AL");
    await filterPage.locator('[data-data-dialog-action="ok"]').click();
    await filterPage.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="1"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const textCaseRows = await filterPage.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        name: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const textCaseExported = await exportedAutoFilterCustomWorkbookState(
      await filterPage.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
    );
    await filterPage.locator('[data-ribbon-command="data-clear-filter"]').click();
    await filterPage.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );
    await filterPage.locator('.cell-filter-button[data-filter-col="1"]').click();
    await filterPage.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-action="condition-menu"]').click();
    await filterPage.waitForSelector(".auto-filter-submenu:not(.hidden)", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-operator="equal"]').click();
    await filterPage.waitForSelector('[data-data-dialog="auto-filter-custom"]', { timeout: 5000 });
    await filterPage.locator('[data-data-field="value1"]').fill("d*");
    await filterPage.locator('[data-data-dialog-action="ok"]').click();
    await filterPage.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="1"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const textWildcardRows = await filterPage.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        name: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const textWildcardExported = await exportedAutoFilterCustomWorkbookState(
      await filterPage.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
    );
    await filterPage.locator('[data-ribbon-command="data-clear-filter"]').click();
    await filterPage.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );
    await filterPage.locator('.cell-filter-button[data-filter-col="1"]').click();
    await filterPage.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-action="condition-menu"]').click();
    await filterPage.waitForSelector(".auto-filter-submenu:not(.hidden)", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-operator="notContains"]').click();
    await filterPage.waitForSelector('[data-data-dialog="auto-filter-custom"]', { timeout: 5000 });
    await filterPage.locator('[data-data-field="value1"]').fill("ta");
    await filterPage.locator('[data-data-dialog-action="ok"]').click();
    await filterPage.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="1"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const textNotContainsRows = await filterPage.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        name: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const textNotContainsExported = await exportedAutoFilterCustomWorkbookState(
      await filterPage.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
    );
    await filterPage.locator('[data-ribbon-command="data-clear-filter"]').click();
    await filterPage.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );
    await filterPage.locator('.cell-filter-button[data-filter-col="2"]').click();
    await filterPage.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-action="condition-menu"]').click();
    await filterPage.waitForSelector(".auto-filter-submenu:not(.hidden)", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-operator="notBetween"]').click();
    await filterPage.waitForSelector('[data-data-dialog="auto-filter-custom"]', { timeout: 5000 });
    await filterPage.locator('[data-data-field="value1"]').fill("10");
    await filterPage.locator('[data-data-field="value2"]').fill("20");
    await filterPage.locator('[data-data-dialog-action="ok"]').click();
    await filterPage.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="2"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const notBetweenRows = await filterPage.evaluate(() =>
      [2, 3, 4, 5, 6].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const notBetweenExported = await exportedAutoFilterCustomWorkbookState(
      await filterPage.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
    );
    await filterPage.locator('[data-ribbon-command="data-clear-filter"]').click();
    await filterPage.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );
    await filterPage.locator('.cell-filter-button[data-filter-col="2"]').click();
    await filterPage.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    const numberMenu = await filterPage.locator("#autoFilterMenu").innerText();
    await filterPage.locator('[data-auto-filter-action="condition-menu"]').click();
    await filterPage.waitForSelector(".auto-filter-submenu:not(.hidden)", { timeout: 5000 });
    const numberSubmenu = await filterPage.locator(".auto-filter-submenu:not(.hidden)").innerText();
    await filterPage.locator('[data-auto-filter-operator="greaterThan"]').click();
    await filterPage.waitForSelector('[data-data-dialog="auto-filter-custom"]', { timeout: 5000 });
    const preset = await filterPage.evaluate(() => document.querySelector('[data-data-field="operator1"]')?.value || "");
    await filterPage.locator('[data-data-field="value1"]').fill("10");
    await filterPage.locator('[data-data-dialog-action="ok"]').click();
    await filterPage.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="2"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const rows = await filterPage.evaluate(() =>
      [2, 3, 4, 5, 6].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    await filterPage.locator('[data-ribbon-command="data-clear-filter"]').click();
    await filterPage.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );
    await filterPage.locator('.cell-filter-button[data-filter-col="2"]').click();
    await filterPage.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-action="condition-menu"]').click();
    await filterPage.waitForSelector(".auto-filter-submenu:not(.hidden)", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-operator="top10"]').click();
    await filterPage.waitForSelector('[data-data-dialog="auto-filter-top10"]', { timeout: 5000 });
    const top10Defaults = await filterPage.evaluate(() => ({
      title: document.querySelector("#top10FilterDialogTitle")?.textContent || "",
      direction: document.querySelector("[data-data-field='top10Direction']")?.value || "",
      mode: document.querySelector("[data-data-field='top10Mode']")?.value || "",
      count: document.querySelector("[data-data-field='top10Count']")?.value || "",
    }));
    await filterPage.locator("[data-data-field='top10Count']").fill("1");
    await filterPage.locator('[data-data-dialog-action="ok"]').click();
    await filterPage.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="2"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const top10Rows = await filterPage.evaluate(() =>
      [2, 3, 4, 5, 6].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    await filterPage.locator('[data-ribbon-command="data-clear-filter"]').click();
    await filterPage.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );
    await filterPage.locator('.cell-filter-button[data-filter-col="2"]').click();
    await filterPage.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-action="condition-menu"]').click();
    await filterPage.waitForSelector(".auto-filter-submenu:not(.hidden)", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-operator="top10"]').click();
    await filterPage.waitForSelector('[data-data-dialog="auto-filter-top10"]', { timeout: 5000 });
    await filterPage.locator("[data-data-field='top10Direction']").selectOption("bottom");
    await filterPage.locator("[data-data-field='top10Mode']").selectOption("percent");
    await filterPage.locator("[data-data-field='top10Count']").fill("50");
    await filterPage.locator('[data-data-dialog-action="ok"]').click();
    await filterPage.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="2"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const bottomPercentRows = await filterPage.evaluate(() =>
      [2, 3, 4, 5, 6].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    await filterPage.locator('[data-ribbon-command="data-clear-filter"]').click();
    await filterPage.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );
    await filterPage.locator('.cell-filter-button[data-filter-col="2"]').click();
    await filterPage.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-action="condition-menu"]').click();
    await filterPage.waitForSelector(".auto-filter-submenu:not(.hidden)", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-operator="aboveStdDev"]').click();
    await filterPage.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="2"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const aboveStdDevRows = await filterPage.evaluate(() =>
      [2, 3, 4, 5, 6].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const aboveStdDevExported = await exportedAutoFilterDateGroupWorkbookState(
      await filterPage.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
    );
    await filterPage.locator('[data-ribbon-command="data-clear-filter"]').click();
    await filterPage.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );
    await filterPage.locator('.cell-filter-button[data-filter-col="2"]').click();
    await filterPage.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-action="condition-menu"]').click();
    await filterPage.waitForSelector(".auto-filter-submenu:not(.hidden)", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-operator="belowStdDev"]').click();
    await filterPage.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="2"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const belowStdDevRows = await filterPage.evaluate(() =>
      [2, 3, 4, 5, 6].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const belowStdDevExported = await exportedAutoFilterDateGroupWorkbookState(
      await filterPage.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
    );
    await filterPage.locator('[data-ribbon-command="data-clear-filter"]').click();
    await filterPage.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );
    await filterPage.locator('.cell-filter-button[data-filter-col="2"]').click();
    await filterPage.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    const colorMenuBefore = await filterPage.locator("#autoFilterMenu").innerText();
    await filterPage.locator('[data-auto-filter-action="color-filter-menu"]').click();
    await filterPage.waitForSelector(".auto-filter-submenu:not(.hidden) [data-auto-filter-action='color-filter'][data-auto-filter-color-target='fill']", { timeout: 5000 });
    const colorSubmenu = await filterPage.locator(".auto-filter-submenu:not(.hidden)").innerText();
    const colorChoice = await filterPage.locator("[data-auto-filter-action='color-filter'][data-auto-filter-color-target='fill']").first().getAttribute("data-auto-filter-color");
    await filterPage.locator("[data-auto-filter-action='color-filter'][data-auto-filter-color-target='fill']").first().click();
    await filterPage.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="2"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const colorRows = await filterPage.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    await filterPage.locator('[data-ribbon-command="data-clear-filter"]').click();
    await filterPage.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );
    await filterPage.locator('.cell-filter-button[data-filter-col="2"]').click();
    await filterPage.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-action="color-filter-menu"]').click();
    await filterPage.waitForSelector(".auto-filter-submenu:not(.hidden) [data-auto-filter-action='color-filter'][data-auto-filter-color-target='font']", { timeout: 5000 });
    const fontColorChoice = await filterPage.locator("[data-auto-filter-action='color-filter'][data-auto-filter-color-target='font']").first().getAttribute("data-auto-filter-color");
    await filterPage.locator("[data-auto-filter-action='color-filter'][data-auto-filter-color-target='font']").first().click();
    await filterPage.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="2"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const fontColorRows = await filterPage.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    await filterPage.locator('[data-ribbon-command="data-clear-filter"]').click();
    await filterPage.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );
    await filterPage.locator('.sheet-cell[data-address="A1"]').click();
    await filterPage.locator('.sheet-cell[data-address="B6"]').click({ modifiers: ["Shift"] });
    await filterPage.locator('[data-ribbon-command="data-advanced-filter"]').click();
    await filterPage.waitForSelector('[data-data-dialog="advanced-filter"]', { timeout: 5000 });
    await filterPage.locator("[data-data-field='criteriaRange']").fill("G1:G2");
    await filterPage.locator('[data-data-dialog-action="ok"]').click();
    await filterPage.waitForFunction(
      () => getComputedStyle(document.querySelector('.row-header[data-row="2"]')).display === "none",
      null,
      { timeout: 5000 },
    );
    const advancedFormulaRows = await filterPage.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    await filterPage.locator('[data-ribbon-command="data-clear-filter"]').click();
    await filterPage.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );
    await filterPage.locator('.cell-filter-button[data-filter-col="2"]').click();
    await filterPage.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-action="color-filter-menu"]').click();
    await filterPage.waitForSelector(".auto-filter-submenu:not(.hidden) [data-auto-filter-action='color-sort'][data-auto-filter-color-target='fill']", { timeout: 5000 });
    await filterPage.locator("[data-auto-filter-action='color-sort'][data-auto-filter-color-target='fill']").first().click();
    await filterPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent === "Beta",
      null,
      { timeout: 5000 },
    );
    const colorSortRows = await filterPage.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        row,
        name: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const colorSortStatus = await filterPage.locator("#statusText").innerText();
    await filterPage.locator('.cell-filter-button[data-filter-col="2"]').click();
    await filterPage.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await filterPage.locator('[data-auto-filter-action="color-filter-menu"]').click();
    await filterPage.waitForSelector(".auto-filter-submenu:not(.hidden) [data-auto-filter-action='color-sort'][data-auto-filter-color-target='font']", { timeout: 5000 });
    await filterPage.locator("[data-auto-filter-action='color-sort'][data-auto-filter-color-target='font']").first().click();
    await filterPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent === "Alpine",
      null,
      { timeout: 5000 },
    );
    const fontColorSortRows = await filterPage.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        row,
        name: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const fontColorSortStatus = await filterPage.locator("#statusText").innerText();
    return { textMenu, textSubmenu, textCaseRows, textCaseExported, textWildcardRows, textWildcardExported, textNotContainsRows, textNotContainsExported, numberMenu, numberSubmenu, preset, rows, top10Defaults, top10Rows, bottomPercentRows, notBetweenRows, notBetweenExported, aboveStdDevRows, aboveStdDevExported, belowStdDevRows, belowStdDevExported, colorMenuBefore, colorSubmenu, colorChoice, colorRows, fontColorChoice, fontColorRows, advancedFormulaRows, colorSortRows, colorSortStatus, fontColorSortRows, fontColorSortStatus };
  } finally {
    await filterPage.close();
  }
}



async function exerciseHomeEditingCommands(context) {
  const workbookPath = await createHomeToolbarWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?homeEditing=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="D5"]', { timeout: 10000 });
    await page.locator('[data-ribbon-tab="home"]').click();

    await page.locator('.sheet-cell[data-address="D5"]').click();
    await page.locator('[data-ribbon-command="autosum"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="D5"] .cell-content')?.textContent === "6",
      null,
      { timeout: 10000 },
    );
    const autoSum = await page.evaluate(() => ({
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      formula: document.querySelector("#formulaInput")?.value || "",
      display: document.querySelector('.sheet-cell[data-address="D5"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => (document.querySelector('.sheet-cell[data-address="D5"] .cell-content')?.textContent || "") === "",
      null,
      { timeout: 10000 },
    );
    const autoSumUndo = await page.evaluate(() => ({
      display: document.querySelector('.sheet-cell[data-address="D5"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="D5"] .cell-content')?.textContent === "6",
      null,
      { timeout: 10000 },
    );
    const autoSumRedo = await page.evaluate(() => ({
      display: document.querySelector('.sheet-cell[data-address="D5"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator('.sheet-cell[data-address="H5"]').click();
    await page.locator('[data-ribbon-command="autosum-menu"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    const autoSumMenu = await page.evaluate(() => ({
      labels: [...document.querySelectorAll("#homeCommandMenu button")].map((button) => button.textContent.replace(/\s+/g, "")),
      accessKeys: [...document.querySelectorAll("#homeCommandMenu .menu-access-key")].map((node) => node.textContent),
      focused: document.activeElement?.textContent?.replace(/\s+/g, "") || "",
    }));
    await page.locator('[data-home-autosum-action="AVERAGE"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="H5"] .cell-content')?.textContent === "8",
      null,
      { timeout: 10000 },
    );
    const autoSumAverage = await page.evaluate(() => ({
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      formula: document.querySelector("#formulaInput")?.value || "",
      display: document.querySelector('.sheet-cell[data-address="H5"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator('.sheet-cell[data-address="J2"]').click();
    await page.locator('.sheet-cell[data-address="J4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="fill-menu"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    const fillMenu = await page.evaluate(() => ({
      labels: [...document.querySelectorAll("#homeCommandMenu button")].map((button) => button.textContent.replace(/\s+/g, "")),
      accessKeys: [...document.querySelectorAll("#homeCommandMenu .menu-access-key")].map((node) => node.textContent),
    }));
    await page.locator('[data-home-fill-action="down"]').click();
    await page.waitForFunction(
      () =>
        document.querySelector('.sheet-cell[data-address="J3"] .cell-content')?.textContent === "Seed" &&
        document.querySelector('.sheet-cell[data-address="J4"] .cell-content')?.textContent === "Seed",
      null,
      { timeout: 10000 },
    );
    const fillDown = await page.evaluate(() => ({
      values: [2, 3, 4].map((row) => document.querySelector(`.sheet-cell[data-address="J${row}"] .cell-content`)?.textContent || ""),
      sourceFill: window.__webSheetDevCellModel?.("HomeTools", 2, 10)?.hasFill || false,
      copiedFill: window.__webSheetDevCellModel?.("HomeTools", 3, 10)?.hasFill || false,
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator('.sheet-cell[data-address="K2"]').click();
    await page.locator('.sheet-cell[data-address="K4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="fill-menu"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-home-fill-action="down"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="K4"] .cell-content')?.textContent === "24",
      null,
      { timeout: 10000 },
    );
	    const fillFormula = await page.evaluate(() => ({
	      k3: window.__webSheetDevCellModel?.("HomeTools", 3, 11)?.formula || "",
	      k4: window.__webSheetDevCellModel?.("HomeTools", 4, 11)?.formula || "",
	      displays: [2, 3, 4].map((row) => document.querySelector(`.sheet-cell[data-address="K${row}"] .cell-content`)?.textContent || ""),
	    }));

	    await page.evaluate(() => {
	      window.__webSheetDevSetCell?.(6, 20, "UpSeed");
	      window.__webSheetDevSelectRange?.(4, 20, 6, 20);
	    });
	    await page.locator('[data-ribbon-command="fill-menu"]').click();
	    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
	    await page.locator('[data-home-fill-action="up"]').click();
	    await page.waitForFunction(
	      () =>
	        window.__webSheetDevCellModel?.("HomeTools", 4, 20)?.raw === "UpSeed" &&
	        window.__webSheetDevCellModel?.("HomeTools", 5, 20)?.raw === "UpSeed",
	      null,
	      { timeout: 10000 },
	    );
	    const fillUp = await page.evaluate(() => ({
	      values: [4, 5, 6].map((row) => window.__webSheetDevCellModel?.("HomeTools", row, 20)?.raw || ""),
	      status: document.querySelector("#statusText")?.textContent || "",
	    }));

	    await page.evaluate(() => {
	      window.__webSheetDevSetCell?.(6, 23, "LeftSeed");
	      window.__webSheetDevSelectRange?.(6, 21, 6, 23);
	    });
	    await page.locator('[data-ribbon-command="fill-menu"]').click();
	    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
	    await page.locator('[data-home-fill-action="left"]').click();
	    await page.waitForFunction(
	      () =>
	        window.__webSheetDevCellModel?.("HomeTools", 6, 21)?.raw === "LeftSeed" &&
	        window.__webSheetDevCellModel?.("HomeTools", 6, 22)?.raw === "LeftSeed",
	      null,
	      { timeout: 10000 },
	    );
	    const fillLeft = await page.evaluate(() => ({
	      values: [21, 22, 23].map((col) => window.__webSheetDevCellModel?.("HomeTools", 6, col)?.raw || ""),
	      status: document.querySelector("#statusText")?.textContent || "",
	    }));

	    await page.evaluate(() => {
	      window.__webSheetDevSetCell?.(8, 20, 2);
	      window.__webSheetDevSelectRange?.(8, 20, 11, 20);
	    });
	    await page.locator('[data-ribbon-command="fill-menu"]').click();
	    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
	    await page.locator('[data-home-fill-action="series-dialog"]').click();
	    await page.waitForSelector("#homeSeriesDialog", { timeout: 5000 });
	    const seriesDialogInitial = await page.evaluate(() => ({
	      title: document.querySelector("#homeSeriesDialogTitle")?.textContent || "",
	      labels: [...document.querySelectorAll("#homeSeriesDialog .data-dialog-check")].map((label) => label.textContent.replace(/\s+/g, "")),
	      step: document.querySelector("[data-home-series-step]")?.value || "",
	      direction: document.querySelector('input[name="homeSeriesDirection"]:checked')?.value || "",
	      type: document.querySelector('input[name="homeSeriesType"]:checked')?.value || "",
	      dateUnitDisabled: [...document.querySelectorAll('[data-home-series-date-unit] input')].every((input) => input.disabled),
	    }));
	    await page.locator('[data-home-series-dialog-action="ok"]').click();
	    await page.waitForSelector("#homeSeriesDialog", { state: "detached", timeout: 5000 });
	    await page.waitForFunction(
	      () => Number(window.__webSheetDevCellModel?.("HomeTools", 11, 20)?.raw) === 5,
	      null,
	      { timeout: 10000 },
	    );
	    const seriesLinear = await page.evaluate(() => ({
	      values: [8, 9, 10, 11].map((row) => window.__webSheetDevCellModel?.("HomeTools", row, 20)?.raw ?? ""),
	      status: document.querySelector("#statusText")?.textContent || "",
	    }));

	    await page.evaluate(() => {
	      window.__webSheetDevSetCell?.(8, 21, 2);
	      window.__webSheetDevSelectRange?.(8, 21, 11, 21);
	    });
	    await page.locator('[data-ribbon-command="fill-menu"]').click();
	    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
	    await page.locator('[data-home-fill-action="series-dialog"]').click();
	    await page.waitForSelector("#homeSeriesDialog", { timeout: 5000 });
	    await page.locator('input[name="homeSeriesType"][value="growth"]').check();
	    await page.locator("[data-home-series-step]").fill("2");
	    await page.locator('[data-home-series-dialog-action="ok"]').click();
	    await page.waitForFunction(
	      () => Number(window.__webSheetDevCellModel?.("HomeTools", 11, 21)?.raw) === 16,
	      null,
	      { timeout: 10000 },
	    );
	    const seriesGrowth = await page.evaluate(() => ({
	      values: [8, 9, 10, 11].map((row) => window.__webSheetDevCellModel?.("HomeTools", row, 21)?.raw ?? ""),
	      status: document.querySelector("#statusText")?.textContent || "",
	    }));

	    await page.evaluate(() => {
	      window.__webSheetDevSetCell?.(12, 20, "Alpha Beta Gamma Delta Epsilon Zeta Eta Theta Iota Kappa Lambda Mu");
	      window.__webSheetDevSelectRange?.(12, 20, 14, 20);
	    });
	    await page.locator('[data-ribbon-command="fill-menu"]').click();
	    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
	    await page.locator('[data-home-fill-action="justify"]').click();
	    await page.waitForFunction(
	      () => (document.querySelector("#statusText")?.textContent || "").includes("両端揃え"),
	      null,
	      { timeout: 10000 },
	    );
	    const fillJustify = await page.evaluate(() => ({
	      values: [12, 13, 14].map((row) => window.__webSheetDevCellModel?.("HomeTools", row, 20)?.raw ?? ""),
	      status: document.querySelector("#statusText")?.textContent || "",
	    }));

	    await page.evaluate(() => window.__webSheetDevSelectRange?.(12, 20, 14, 20));
	    await page.locator('[data-ribbon-command="fill-menu"]').click();
	    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
	    await page.locator('[data-home-fill-action="worksheet-fill-dialog"]').click();
	    await page.waitForSelector("#messageBox", { timeout: 5000 });
	    const worksheetFillNoGroup = await page.evaluate(() => ({
	      message: document.querySelector("#messageBoxText")?.textContent || "",
	      title: document.querySelector("#messageBoxTitle")?.textContent || "",
	    }));
	    await page.locator("[data-message-box-close]").first().click();
	    await page.waitForSelector("#messageBox", { state: "detached", timeout: 5000 });

	    await page.locator('.sheet-cell[data-address="J3"]').click();
    await page.locator('[data-ribbon-command="clear-menu"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    const clearMenu = await page.evaluate(() => ({
      labels: [...document.querySelectorAll("#homeCommandMenu button")].map((button) => button.textContent.replace(/\s+/g, "")),
      accessKeys: [...document.querySelectorAll("#homeCommandMenu .menu-access-key")].map((node) => node.textContent),
    }));
    await page.locator('[data-home-clear-action="contents"]').click();
    await page.waitForFunction(
      () => (document.querySelector('.sheet-cell[data-address="J3"] .cell-content')?.textContent || "") === "",
      null,
      { timeout: 10000 },
    );
    const clearContents = await page.evaluate(() => ({
      raw: window.__webSheetDevCellModel?.("HomeTools", 3, 10)?.raw ?? null,
      display: document.querySelector('.sheet-cell[data-address="J3"] .cell-content')?.textContent || "",
      hasFill: window.__webSheetDevCellModel?.("HomeTools", 3, 10)?.hasFill || false,
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="J3"] .cell-content')?.textContent === "Seed",
      null,
      { timeout: 10000 },
    );
    const clearUndo = await page.evaluate(() => ({
      display: document.querySelector('.sheet-cell[data-address="J3"] .cell-content')?.textContent || "",
      hasFill: window.__webSheetDevCellModel?.("HomeTools", 3, 10)?.hasFill || false,
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(
      () => (document.querySelector('.sheet-cell[data-address="J3"] .cell-content')?.textContent || "") === "",
      null,
      { timeout: 10000 },
    );

    await page.locator('.sheet-cell[data-address="L2"]').click();
    await page.locator('[data-ribbon-command="clear-menu"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-home-clear-action="formats"]').click();
    await page.waitForFunction(() => window.__webSheetDevCellModel?.("HomeTools", 2, 12)?.hasFill !== true, null, { timeout: 10000 });
    const clearFormats = await page.evaluate(() => ({
      raw: window.__webSheetDevCellModel?.("HomeTools", 2, 12)?.raw || "",
      hasFill: window.__webSheetDevCellModel?.("HomeTools", 2, 12)?.hasFill || false,
      borders: Object.keys(window.__webSheetDevCellModel?.("HomeTools", 2, 12)?.borders || {}).length,
    }));

    await page.locator('.sheet-cell[data-address="G2"]').click();
    await page.locator('[data-ribbon-command="clear-menu"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-home-clear-action="comments"]').click();
    await page.waitForFunction(() => !window.__webSheetDevCellModel?.("HomeTools", 2, 7)?.note, null, { timeout: 10000 });
    const clearNotes = await page.evaluate(() => ({
      note: window.__webSheetDevCellModel?.("HomeTools", 2, 7)?.note || null,
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator('.sheet-cell[data-address="M2"]').click();
    await page.locator('[data-ribbon-command="clear-menu"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-home-clear-action="hyperlinks"]').click();
    await page.waitForFunction(() => !window.__webSheetDevCellModel?.("HomeTools", 2, 13)?.hyperlink, null, { timeout: 10000 });
    const clearHyperlinks = await page.evaluate(() => ({
      raw: window.__webSheetDevCellModel?.("HomeTools", 2, 13)?.raw || "",
      hyperlink: window.__webSheetDevCellModel?.("HomeTools", 2, 13)?.hyperlink || null,
    }));

    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="B4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="sort-filter"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    const sortFilterMenu = await page.evaluate(() => ({
      labels: [...document.querySelectorAll("#homeCommandMenu button")].map((button) => button.textContent.replace(/\s+/g, "")),
      accessKeys: [...document.querySelectorAll("#homeCommandMenu .menu-access-key")].map((node) => node.textContent),
      firstFocused: document.activeElement?.textContent?.replace(/\s+/g, "") || "",
    }));
    await page.locator('[data-home-sort-filter-action="sort-asc"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent === "Alpha",
      null,
      { timeout: 10000 },
    );
    const sorted = await page.evaluate(() => ({
      names: [2, 3, 4].map((row) => document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || ""),
      scores: [2, 3, 4].map((row) => document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || ""),
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="B4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="sort-filter"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-home-sort-filter-action="filter"]').click();
    await page.waitForSelector('.cell-filter-button[data-filter-col="1"]', { timeout: 5000 });
    const filter = await page.evaluate(() => ({
      buttonCount: document.querySelectorAll(".cell-filter-button").length,
      status: document.querySelector("#statusText")?.textContent || "",
      menuClosed: !document.querySelector("#homeCommandMenu"),
    }));

    await page.locator('[data-ribbon-command="find-select"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    const findSelectMenu = await page.evaluate(() => ({
      labels: [...document.querySelectorAll("#homeCommandMenu button")].map((button) => button.textContent.replace(/\s+/g, "")),
      accessKeys: [...document.querySelectorAll("#homeCommandMenu .menu-access-key")].map((node) => node.textContent),
    }));
    await page.locator('[data-find-select-action="find"]').click();
    await page.waitForSelector("#findDialog", { timeout: 5000 });
    const findDialogInitial = await page.evaluate(() => ({
      title: document.querySelector("#findDialogTitle")?.textContent || "",
      labels: [...document.querySelectorAll("#findDialog .excel-dialog-caption")].map((label) => label.textContent.trim()),
      nonBlocking: getComputedStyle(document.querySelector("#findDialogOverlay")).pointerEvents === "none",
      activeInput: document.activeElement?.matches?.("[data-find-field='query']") || false,
    }));
    await page.locator('[data-find-field="query"]').fill("Gamma");
    await page.locator('[data-find-dialog-action="find-next"]').click();
    await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "A4", null, { timeout: 5000 });
    const findResult = await page.evaluate(() => ({
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
      dialogStillOpen: Boolean(document.querySelector("#findDialog")),
    }));
    await page.locator('[data-find-dialog-action="close"]').click();
    await page.waitForSelector("#findDialog", { state: "detached", timeout: 5000 });

    await page.keyboard.press("Control+F");
    await page.waitForSelector("#findDialog", { timeout: 5000 });
    const shortcutFind = await page.evaluate(() => ({
      title: document.querySelector("#findDialogTitle")?.textContent || "",
      activeTab: document.querySelector(".find-dialog-tab.is-active")?.textContent || "",
      activeInput: document.activeElement?.matches?.("[data-find-field='query']") || false,
    }));
    await page.locator('[data-find-dialog-action="close"]').click();
    await page.waitForSelector("#findDialog", { state: "detached", timeout: 5000 });
    await page.keyboard.press("Control+H");
    await page.waitForSelector("#findDialog", { timeout: 5000 });
    const shortcutReplace = await page.evaluate(() => ({
      activeTab: document.querySelector(".find-dialog-tab.is-active")?.textContent || "",
      replaceVisible: !document.querySelector(".find-dialog-replace")?.classList.contains("hidden"),
      activeInput: document.activeElement?.matches?.("[data-find-field='query']") || false,
    }));
    await page.locator('[data-find-dialog-action="close"]').click();
    await page.waitForSelector("#findDialog", { state: "detached", timeout: 5000 });
    await page.keyboard.press("Control+G");
    await page.waitForSelector("#goToDialog", { timeout: 5000 });
    const shortcutGoTo = await page.evaluate(() => ({
      title: document.querySelector("#goToDialogTitle")?.textContent || "",
      activeInput: document.activeElement?.matches?.("[data-go-to-field='reference']") || false,
    }));
    await page.locator('[data-go-to-dialog-action="cancel"]').click();
    await page.waitForSelector("#goToDialog", { state: "detached", timeout: 5000 });

    await page.locator('[data-ribbon-command="find-select"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-find-select-action="go-to"]').click();
    await page.waitForSelector("#goToDialog", { timeout: 5000 });
    const goToDialogInitial = await page.evaluate(() => ({
      title: document.querySelector("#goToDialogTitle")?.textContent || "",
      referenceLabel: document.querySelector("#goToDialog .excel-dialog-caption")?.textContent || "",
      activeInput: document.activeElement?.matches?.("[data-go-to-field='reference']") || false,
    }));
    await page.locator("[data-go-to-field='reference']").fill("K4");
    await page.locator('[data-go-to-dialog-action="ok"]').click();
    await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "K4", null, { timeout: 5000 });
    const goToResult = await page.evaluate(() => ({
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator('.sheet-cell[data-address="H1"]').click();
    await page.locator('.sheet-cell[data-address="K4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="find-select"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-find-select-action="go-to-special"]').click();
    await page.waitForSelector("#specialDialog", { timeout: 5000 });
    const specialDialogInitial = await page.evaluate(() => ({
      title: document.querySelector("#specialDialogTitle")?.textContent || "",
      labels: [...document.querySelectorAll("#specialDialog .special-dialog-option > span")].map((node) => node.textContent.replace(/\s+/g, "")),
      accessKeys: [...document.querySelectorAll("#specialDialog .menu-access-key")].map((node) => node.textContent),
    }));
    await page.locator('input[name="specialOption"][value="blanks"]').check();
    await page.locator('[data-special-dialog-action="ok"]').click();
    await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent?.includes("I1"), null, { timeout: 5000 });
    const specialBlanks = await page.evaluate(() => ({
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
      frameCount: document.querySelectorAll(".sheet-selection-frame").length,
    }));

    await page.locator('.sheet-cell[data-address="P2"]').click();
    await page.locator('.sheet-cell[data-address="R3"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="find-select"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-find-select-action="go-to-special"]').click();
    await page.waitForSelector("#specialDialog", { timeout: 5000 });
    await page.locator('input[name="specialOption"][value="row-differences"]').check();
    await page.locator('[data-special-dialog-action="ok"]').click();
    await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent?.includes("R2"), null, { timeout: 5000 });
    const specialRowDifferences = await page.evaluate(() => ({
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
      frameCount: document.querySelectorAll(".sheet-selection-frame").length,
    }));

    await page.locator('.sheet-cell[data-address="P5"]').click();
    await page.locator('.sheet-cell[data-address="Q7"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="find-select"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-find-select-action="go-to-special"]').click();
    await page.waitForSelector("#specialDialog", { timeout: 5000 });
    await page.locator('input[name="specialOption"][value="column-differences"]').check();
    await page.locator('[data-special-dialog-action="ok"]').click();
    await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent?.includes("Q6"), null, { timeout: 5000 });
    const specialColumnDifferences = await page.evaluate(() => ({
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
      frameCount: document.querySelectorAll(".sheet-selection-frame").length,
    }));

    await page.locator('.sheet-cell[data-address="S4"]').click();
    await page.locator('[data-ribbon-command="find-select"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-find-select-action="go-to-special"]').click();
    await page.waitForSelector("#specialDialog", { timeout: 5000 });
    await page.locator('input[name="specialOption"][value="precedents"]').check();
    await page.locator('[data-special-dialog-action="ok"]').click();
    await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent?.includes("S2:S3"), null, { timeout: 5000 });
    const specialPrecedents = await page.evaluate(() => ({
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator('.sheet-cell[data-address="S2"]').click();
    await page.locator('.sheet-cell[data-address="S3"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="find-select"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-find-select-action="go-to-special"]').click();
    await page.waitForSelector("#specialDialog", { timeout: 5000 });
    await page.locator('input[name="specialOption"][value="dependents"]').check();
    await page.locator('[data-special-dialog-action="ok"]').click();
    await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "S4", null, { timeout: 5000 });
    const specialDependents = await page.evaluate(() => ({
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.keyboard.press("Control+G");
    await page.waitForSelector("#goToDialog", { timeout: 5000 });
    await page.locator("[data-go-to-field='reference']").fill("N2");
    await page.locator('[data-go-to-dialog-action="ok"]').click();
    await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "N2", null, { timeout: 5000 });
    await page.locator('[data-ribbon-command="find-select"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-find-select-action="go-to-special"]').click();
    await page.waitForSelector("#specialDialog", { timeout: 5000 });
    await page.locator('input[name="specialOption"][value="same-validation"]').check();
    await page.locator('[data-special-dialog-action="ok"]').click();
    await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent?.includes("N3"), null, { timeout: 5000 });
    const specialSameValidation = await page.evaluate(() => ({
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
      frameCount: document.querySelectorAll(".sheet-selection-frame").length,
    }));

    await page.locator('[data-ribbon-command="find-select"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-find-select-action="replace"]').click();
    await page.waitForSelector("#findDialog", { timeout: 5000 });
    const replaceDialogInitial = await page.evaluate(() => ({
      labels: [...document.querySelectorAll("#findDialog .excel-dialog-caption")].map((label) => label.textContent.trim()),
      replaceVisible: !document.querySelector(".find-dialog-replace")?.classList.contains("hidden"),
      activeTab: document.querySelector(".find-dialog-tab.is-active")?.textContent || "",
    }));
    await page.locator('[data-find-field="query"]').fill("Gamma");
    await page.locator('[data-find-field="replacement"]').fill("Delta");
    await page.locator('[data-find-dialog-action="replace-one"]').click();
    await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="A4"] .cell-content')?.textContent === "Delta", null, { timeout: 10000 });
    const replaceOne = await page.evaluate(() => ({
      a4: document.querySelector('.sheet-cell[data-address="A4"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
      formula: document.querySelector("#formulaInput")?.value || "",
    }));
    await page.locator('[data-find-dialog-action="close"]').click();
    await page.waitForSelector("#findDialog", { state: "detached", timeout: 5000 });
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="A4"] .cell-content')?.textContent === "Gamma", null, { timeout: 10000 });
    const replaceUndo = await page.evaluate(() => ({
      a4: document.querySelector('.sheet-cell[data-address="A4"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.keyboard.press("Control+H");
    await page.waitForSelector("#findDialog", { timeout: 5000 });
    await page.locator('[data-find-field="query"]').fill("foo");
    await page.locator('[data-find-field="replacement"]').fill("baz");
    await page.locator('[data-find-dialog-action="replace-all"]').click();
    await page.waitForFunction(
      () =>
        document.querySelector('.sheet-cell[data-address="O2"] .cell-content')?.textContent === "find baz" &&
        document.querySelector('.sheet-cell[data-address="O3"] .cell-content')?.textContent === "baz baz",
      null,
      { timeout: 10000 },
    );
    const replaceAll = await page.evaluate(() => ({
      o2: document.querySelector('.sheet-cell[data-address="O2"] .cell-content')?.textContent || "",
      o3: document.querySelector('.sheet-cell[data-address="O3"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.locator('[data-find-dialog-action="close"]').click();
    await page.waitForSelector("#findDialog", { state: "detached", timeout: 5000 });

    await page.locator('[data-ribbon-command="find-select"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-find-select-action="select-formulas"]').click();
    await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent?.includes("F1"), null, { timeout: 5000 });
    const selectFormulas = await page.evaluate(() => ({
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      formula: document.querySelector("#formulaInput")?.value || "",
      status: document.querySelector("#statusText")?.textContent || "",
      frameCount: document.querySelectorAll(".sheet-selection-frame").length,
    }));

    await page.locator('[data-ribbon-command="find-select"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-find-select-action="select-notes"]').click();
    await page.waitForSelector("#homeCommandMenu", { state: "detached", timeout: 5000 });
    await page.waitForTimeout(120);
    const selectNotes = await page.evaluate(() => ({
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
    }));

    if (await page.locator("#messageBox").count()) {
      await page.locator("[data-message-box-close]").first().click();
      await page.waitForSelector("#messageBox", { state: "detached", timeout: 5000 });
    }
    await page.locator('[data-ribbon-command="find-select"]').click();
    await page.waitForSelector("#homeCommandMenu", { timeout: 5000 });
    await page.locator('[data-find-select-action="select-validation"]').click();
    await page.waitForSelector("#homeCommandMenu", { state: "detached", timeout: 5000 });
    await page.waitForTimeout(120);
    const selectValidation = await page.evaluate(() => ({
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    return {
      autoSum,
      autoSumUndo,
      autoSumRedo,
      autoSumMenu,
      autoSumAverage,
	      fillMenu,
	      fillDown,
	      fillFormula,
	      fillUp,
	      fillLeft,
	      seriesDialogInitial,
	      seriesLinear,
	      seriesGrowth,
	      fillJustify,
	      worksheetFillNoGroup,
	      clearMenu,
      clearContents,
      clearUndo,
      clearFormats,
      clearNotes,
      clearHyperlinks,
      sortFilterMenu,
      sorted,
      filter,
      findSelectMenu,
      findDialogInitial,
      findResult,
      shortcutFind,
      shortcutReplace,
      shortcutGoTo,
      goToDialogInitial,
      goToResult,
      specialDialogInitial,
      specialBlanks,
      specialRowDifferences,
      specialColumnDifferences,
      specialPrecedents,
      specialDependents,
      specialSameValidation,
      replaceDialogInitial,
      replaceOne,
      replaceUndo,
      replaceAll,
      selectFormulas,
      selectNotes,
      selectValidation,
    };
  } finally {
    await page.close();
  }
}



async function exerciseImportedAutoFilterCriteria(context) {
  const workbookPath = await createImportedAutoFilterWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?importedAutoFilter=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector(".cell-filter-button", { timeout: 10000 });
    const valueFilter = await importedAutoFilterSheetState(page, "値フィルター", [2, 3, 4, 5], 1);
    const customFilter = await importedAutoFilterSheetState(page, "条件フィルター", [2, 3, 4, 5], 2);
    const top10Filter = await importedAutoFilterSheetState(page, "トップテン", [2, 3, 4, 5], 2);
    const averageFilter = await importedAutoFilterSheetState(page, "平均", [2, 3, 4, 5], 2);
    const dateGroupFilter = await importedAutoFilterSheetState(page, "日付グループ", [2, 3, 4, 5], 1);
    const blankFilter = await importedAutoFilterSheetState(page, "空白インポート", [2, 3, 4, 5], 1);
    const savedSort = await importedAutoFilterSheetState(page, "保存済み並べ替え", [2, 3, 4, 5], 2);
    const standaloneSort = await importedAutoFilterSheetState(page, "単独並べ替え", [2, 3, 4, 5], 1, { expectFilterButton: false });
    const tableFilter = await importedAutoFilterSheetState(page, "テーブルフィルター", [2, 3, 4, 5], 1);
    const tableSortOnly = await importedTableSortOnlyState(page, "テーブル並べ替え");
    const importedCellColorFilter = await importedAutoFilterSheetState(page, "セル色フィルター", [2, 3, 4, 5], 2);
    const importedFontColorFilter = await importedAutoFilterSheetState(page, "フォント色フィルター", [2, 3, 4, 5], 2);
    const importedTableColorFilter = await importedAutoFilterSheetState(page, "テーブル色フィルター", [2, 3, 4, 5], 2);
    const importedCellColorSort = await importedAutoFilterSheetState(page, "セル色並べ替え", [2, 3, 4, 5], 2);
    const importedFontColorSort = await importedAutoFilterSheetState(page, "フォント色並べ替え", [2, 3, 4, 5], 2);
    const importedCustomListSort = await importedAutoFilterSheetState(page, "リスト並べ替え", [2, 3, 4, 5], 1);
    const importedTextContainsFilter = await importedAutoFilterSheetState(page, "文字含むフィルター", [2, 3, 4, 5], 1);
    const importedTextBeginsFilter = await importedAutoFilterSheetState(page, "文字始まりフィルター", [2, 3, 4, 5], 1);
    const importedTextNotEndsFilter = await importedAutoFilterSheetState(page, "文字終わり除外", [2, 3, 4, 5], 1);
    const importedLiteralWildcardFilter = await importedAutoFilterSheetState(page, "文字リテラルフィルター", [2, 3, 4, 5], 1);
    const importedMultiDateGroupFilter = await importedAutoFilterSheetState(page, "日付複数グループ", [2, 3, 4, 5], 1);
    const importedBetweenFilter = await importedAutoFilterSheetState(page, "範囲内フィルター", [2, 3, 4, 5], 2);
    const importedNotBetweenFilter = await importedAutoFilterSheetState(page, "範囲外フィルター", [2, 3, 4, 5], 2);
    const importedMonthPeriodFilter = await importedAutoFilterSheetState(page, "月期間フィルター", [2, 3, 4, 5], 1);
    const importedQuarterPeriodFilter = await importedAutoFilterSheetState(page, "四半期期間フィルター", [2, 3, 4, 5], 1);
    const importedAboveStdDevFilter = await importedAutoFilterSheetState(page, "標準偏差上フィルター", [2, 3, 4, 5], 2);
    const importedBelowStdDevFilter = await importedAutoFilterSheetState(page, "標準偏差下フィルター", [2, 3, 4, 5], 2);
    const importedHiddenButtonFilter = await importedHiddenButtonAutoFilterState(page, "非表示ボタンフィルター");
    const importedDateHourGroupFilter = await importedAutoFilterSheetState(page, "日付時刻グループ", [2, 3, 4, 5], 1);
    const importedCustomNonBlankFilter = await importedAutoFilterSheetState(page, "空白以外カスタム", [2, 3, 4, 5], 1);
    const exportedWorkbookBase64 = await page.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || "");
    const importedDateGroupExported = await exportedAutoFilterDateGroupWorkbookState(exportedWorkbookBase64);
    const importedCustomFilterExported = await exportedAutoFilterCustomWorkbookState(exportedWorkbookBase64);
    const tableMetadata = await importedTableMetadataState(page, "テーブルメタ");
    const structuredReferences = await importedStructuredReferenceState(page, "構造化参照");
    const tableExpansion = await importedTableExpansionState(page, "テーブル拡張");
    const multiTableFilter = await importedMultiTableFilterState(page, "複数テーブルフィルター");
    const filterStructureShift = await importedFilterStructureShiftState(context, workbookPath);
    const tableRowInsert = await importedTableRowInsertState(context, workbookPath);
    const tableTotalsStructure = await importedTableTotalsStructureState(context, workbookPath);
    const structureReferenceShift = await importedStructureReferenceShiftState(context, workbookPath);
    return { valueFilter, customFilter, top10Filter, averageFilter, dateGroupFilter, blankFilter, savedSort, standaloneSort, tableFilter, tableSortOnly, importedCellColorFilter, importedFontColorFilter, importedTableColorFilter, importedCellColorSort, importedFontColorSort, importedCustomListSort, importedTextContainsFilter, importedTextBeginsFilter, importedTextNotEndsFilter, importedLiteralWildcardFilter, importedMultiDateGroupFilter, importedBetweenFilter, importedNotBetweenFilter, importedMonthPeriodFilter, importedQuarterPeriodFilter, importedAboveStdDevFilter, importedBelowStdDevFilter, importedHiddenButtonFilter, importedDateHourGroupFilter, importedCustomNonBlankFilter, importedDateGroupExported, importedCustomFilterExported, tableMetadata, structuredReferences, tableExpansion, multiTableFilter, filterStructureShift, tableRowInsert, tableTotalsStructure, structureReferenceShift };
  } finally {
    await page.close();
  }
}



async function importedAutoFilterSheetState(page, sheetName, rows, filteredCol, options = {}) {
  await page.evaluate((name) => window.__webSheetDevSetActiveSheet?.(name), sheetName);
  await page.waitForFunction(
    (name) => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === name,
    sheetName,
    { timeout: 5000 },
  );
  if (options.expectFilterButton !== false) {
    await page.waitForSelector(`.cell-filter-button[data-filter-col="${filteredCol}"]`, { timeout: 5000 });
  }
  return page.evaluate(({ name, rows: targetRows, filteredCol: col }) => {
    const sheet = window.__webSheetDevSheetModel?.(name);
    return {
      criteria: sheet?.autoFilter?.criteria?.[col] || null,
      sortState: sheet?.sortState || null,
      buttonFiltered: document.querySelector(`.cell-filter-button[data-filter-col="${col}"]`)?.classList.contains("is-filtered") || false,
      buttonSorted: document.querySelector(`.cell-filter-button[data-filter-col="${col}"]`)?.classList.contains("is-sorted") || false,
      buttonSortDirection: document.querySelector(`.cell-filter-button[data-filter-col="${col}"]`)?.getAttribute("data-sort-direction") || "",
      hasFilterButton: Boolean(document.querySelector(`.cell-filter-button[data-filter-col="${col}"]`)),
      rows: targetRows.map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        name: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    };
  }, { name: sheetName, rows, filteredCol });
}



async function importedHiddenButtonAutoFilterState(page, sheetName) {
  await page.evaluate((name) => window.__webSheetDevSetActiveSheet?.(name), sheetName);
  await page.waitForFunction(
    (name) => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === name,
    sheetName,
    { timeout: 5000 },
  );
  await page.waitForSelector('.cell-filter-button[data-filter-col="1"].is-filtered', { timeout: 5000 });
  return page.evaluate((name) => {
    const sheet = window.__webSheetDevSheetModel?.(name);
    return {
      criteria: sheet?.autoFilter?.criteria || null,
      hiddenButtons: sheet?.autoFilter?.hiddenButtons || null,
      hasNameButton: Boolean(document.querySelector('.cell-filter-button[data-filter-col="1"]')),
      hasScoreButton: Boolean(document.querySelector('.cell-filter-button[data-filter-col="2"]')),
      rows: [2, 3, 4, 5].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        name: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    };
  }, sheetName);
}



async function importedTableSortOnlyState(page, sheetName) {
  await page.evaluate((name) => window.__webSheetDevSetActiveSheet?.(name), sheetName);
  await page.waitForFunction(
    (name) => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === name,
    sheetName,
    { timeout: 5000 },
  );
  await page.waitForSelector('.cell-filter-button[data-filter-col="2"].is-sorted[data-sort-direction="desc"]', { timeout: 5000 });
  const before = await importedAutoFilterSheetState(page, sheetName, [2, 3, 4, 5], 2);
  await page.locator('[data-ribbon-tab="data"]').click();
  const ribbonBeforeClear = await page.evaluate((name) => {
    const stateFor = (command) => {
      const button = document.querySelector(`[data-ribbon-command="${command}"]`);
      return {
        active: button?.classList.contains("is-active") || false,
        disabled: Boolean(button?.disabled),
      };
    };
    const sheet = window.__webSheetDevSheetModel?.(name);
    return {
      filter: stateFor("data-filter"),
      clear: stateFor("data-clear-filter"),
      reapply: stateFor("data-reapply-filter"),
      sheetSortState: sheet?.sortState || null,
      tableSortState: sheet?.tables?.[0]?.autoFilterState?.sortState || sheet?.tables?.[0]?.sortState || null,
      buttonSorted: document.querySelector('.cell-filter-button[data-filter-col="2"]')?.classList.contains("is-sorted") || false,
      buttonSortDirection: document.querySelector('.cell-filter-button[data-filter-col="2"]')?.getAttribute("data-sort-direction") || "",
    };
  }, sheetName);
  const exportedBeforeClear = await exportedTableSortWorkbookState(
    await page.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
  );
  await page.locator('[data-ribbon-command="data-clear-filter"]').click();
  await page.waitForFunction(
    () => !document.querySelector('.cell-filter-button[data-filter-col="2"]')?.classList.contains("is-sorted"),
    null,
    { timeout: 5000 },
  );
  const afterClear = await page.evaluate((name) => {
    const sheet = window.__webSheetDevSheetModel?.(name);
    const button = document.querySelector('.cell-filter-button[data-filter-col="2"]');
    return {
      sheetSortState: sheet?.sortState || null,
      tableSortState: sheet?.tables?.[0]?.autoFilterState?.sortState || sheet?.tables?.[0]?.sortState || null,
      buttonSorted: button?.classList.contains("is-sorted") || false,
      buttonSortDirection: button?.getAttribute("data-sort-direction") || "",
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetName);
  return { before, ribbonBeforeClear, exportedBeforeClear, afterClear };
}



async function exportedTableSortWorkbookState(base64) {
  const buffer = Buffer.from(String(base64 || ""), "base64");
  const zip = await JSZip.loadAsync(buffer);
  const tableXmlEntries = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^xl\/tables\/table\d+\.xml$/i.test(name))
      .map(async (name) => [name, await zip.file(name).async("text")]),
  );
  const worksheetXmlEntries = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
      .map(async (name) => [name, await zip.file(name).async("text")]),
  );
  const workbook = new ExcelJS.Workbook();
  let excelJsLoaded = false;
  let worksheetNames = [];
  try {
    await workbook.xlsx.load(await workbookBufferForExcelJsSmoke(buffer));
    excelJsLoaded = true;
    worksheetNames = workbook.worksheets.map((sheet) => sheet.name);
  } catch (error) {
    return {
      excelJsLoaded: false,
      error: error?.message || String(error),
      worksheetNames,
      tableXmlHasSortState: tableXmlEntries.some(([, xml]) => /<sortState\b/.test(xml)),
      worksheetSortStateCount: 0,
    };
  }
  return {
    excelJsLoaded,
    worksheetNames,
    tableXmlHasSortState: tableXmlEntries.some(([, xml]) => /<sortState\b/.test(xml)),
    worksheetSortStateCount: worksheetXmlEntries.filter(([, xml]) => /<sortState\b/.test(xml)).length,
  };
}



async function workbookBufferForExcelJsSmoke(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  let changed = false;
  await Promise.all(Object.keys(zip.files)
    .filter((name) => /^xl\/tables\/table\d+\.xml$/i.test(name))
    .map(async (name) => {
      const xml = await zip.file(name)?.async("text");
      const nextXml = String(xml || "")
        .replace(/<sortState\b[\s\S]*?<\/sortState>|<sortState\b[^>]*\/>/g, "")
        .replace(/<colorFilter\b[\s\S]*?<\/colorFilter>|<colorFilter\b[^>]*\/>/g, "");
      if (nextXml === xml) return;
      zip.file(name, nextXml);
      changed = true;
    }));
  return changed ? zip.generateAsync({ type: "nodebuffer" }) : buffer;
}



async function exerciseAutoFilterColorImportExport(context) {
  const filterWorkbookPath = await createAutoFilterMenuWorkbook({ includeMixedTextScore: true });
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?autoFilterColorExport=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", filterWorkbookPath);
    await waitForUploadedWorkbook(page, filterWorkbookPath);
    await page.waitForSelector('.sheet-cell[data-address="B6"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="B6"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-filter"]').click();
    await page.waitForSelector('.cell-filter-button[data-filter-col="2"]', { timeout: 5000 });
    await page.locator('.cell-filter-button[data-filter-col="2"]').click();
    await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await page.locator('[data-auto-filter-action="color-filter-menu"]').click();
    await page.waitForSelector(".auto-filter-submenu:not(.hidden) [data-auto-filter-action='color-filter'][data-auto-filter-color-target='fill']", { timeout: 5000 });
    const colorChoice = await page.locator("[data-auto-filter-action='color-filter'][data-auto-filter-color-target='fill']").first().getAttribute("data-auto-filter-color");
    await page.locator("[data-auto-filter-action='color-filter'][data-auto-filter-color-target='fill']").first().click();
    await page.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="2"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const beforeExport = await page.evaluate(() => {
      const sheet = window.__webSheetDevSheetModel?.("Filter");
      return {
        criteria: sheet?.autoFilter?.criteria?.[2] || null,
        rows: [2, 3, 4, 5, 6].map((row) => ({
          row,
          hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
          score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
        })),
      };
    });
    const exportedBase64 = await page.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || "");
    const exported = await exportedAutoFilterColorWorkbookState(exportedBase64);
    const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-color-filter-export-"));
    const exportedPath = path.join(dir, "color-filter-export.xlsx");
    await writeFile(exportedPath, Buffer.from(exportedBase64, "base64"));

    await page.locator('[data-ribbon-command="data-clear-filter"]').click();
    await page.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );
    await page.locator('.cell-filter-button[data-filter-col="2"]').click();
    await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
    await page.locator('[data-auto-filter-action="color-filter-menu"]').click();
    await page.waitForSelector(".auto-filter-submenu:not(.hidden) [data-auto-filter-action='color-filter'][data-auto-filter-color-target='font']", { timeout: 5000 });
    const fontColorChoice = await page.locator("[data-auto-filter-action='color-filter'][data-auto-filter-color-target='font']").first().getAttribute("data-auto-filter-color");
    await page.locator("[data-auto-filter-action='color-filter'][data-auto-filter-color-target='font']").first().click();
    await page.waitForFunction(
      () => document.querySelector('.cell-filter-button[data-filter-col="2"]')?.classList.contains("is-filtered"),
      null,
      { timeout: 5000 },
    );
    const fontBeforeExport = await page.evaluate(() => {
      const sheet = window.__webSheetDevSheetModel?.("Filter");
      return {
        criteria: sheet?.autoFilter?.criteria?.[2] || null,
        rows: [2, 3, 4, 5, 6].map((row) => ({
          row,
          hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
          score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
        })),
      };
    });
    const fontExportedBase64 = await page.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || "");
    const fontExported = await exportedAutoFilterColorWorkbookState(fontExportedBase64);
    const fontExportedPath = path.join(dir, "font-color-filter-export.xlsx");
    await writeFile(fontExportedPath, Buffer.from(fontExportedBase64, "base64"));

    const reloadExportedColorFilter = async (targetPath) => {
      const reloadPage = await context.newPage();
      try {
        await reloadPage.goto(`${appUrl}?autoFilterColorReload=${Date.now()}-${path.basename(targetPath)}`, { waitUntil: "load" });
        await reloadPage.setInputFiles("#fileInput", targetPath);
        await waitForUploadedWorkbook(reloadPage, targetPath);
        await reloadPage.waitForSelector('.cell-filter-button[data-filter-col="2"]', { timeout: 10000 });
        return reloadPage.evaluate(() => {
          const sheet = window.__webSheetDevSheetModel?.("Filter");
          const button = document.querySelector('.cell-filter-button[data-filter-col="2"]');
          return {
            buttonFiltered: button?.classList.contains("is-filtered") || false,
            criteria: sheet?.autoFilter?.criteria?.[2] || null,
            rows: [2, 3, 4, 5, 6].map((row) => ({
              row,
              hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
              score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
            })),
          };
        });
      } finally {
        await reloadPage.close();
      }
    };
    const reloaded = await reloadExportedColorFilter(exportedPath);
    const fontReloaded = await reloadExportedColorFilter(fontExportedPath);
    return { colorChoice, beforeExport, exported, reloaded, fontColorChoice, fontBeforeExport, fontExported, fontReloaded };
  } finally {
    await page.close();
  }
}



async function exportedAutoFilterColorWorkbookState(base64) {
  const buffer = Buffer.from(String(base64 || ""), "base64");
  const workbook = new ExcelJS.Workbook();
  let excelJsLoaded = false;
  let excelJsError = "";
  try {
    await workbook.xlsx.load(await workbookBufferForExcelJsSmoke(buffer));
    excelJsLoaded = true;
  } catch (error) {
    excelJsError = error?.message || String(error);
  }
  const zip = await JSZip.loadAsync(buffer);
  const stylesXml = await zip.file("xl/styles.xml")?.async("text");
  const worksheetXmlEntries = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
      .map(async (name) => [name, await zip.file(name).async("text")]),
  );
  const tableXmlEntries = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^xl\/tables\/table\d+\.xml$/i.test(name))
      .map(async (name) => [name, await zip.file(name).async("text")]),
  );
  const worksheetColorFilters = worksheetXmlEntries.flatMap(([, xml]) => [...String(xml || "").matchAll(/<colorFilter\b([^>]*)\/>/g)].map((match) => match[1] || ""));
  const tableColorFilters = tableXmlEntries.flatMap(([, xml]) => [...String(xml || "").matchAll(/<colorFilter\b([^>]*)\/>/g)].map((match) => match[1] || ""));
  const dxfCount = Number.parseInt(String(stylesXml || "").match(/<dxfs\b[^>]*\bcount="(\d+)"/)?.[1] || "0", 10);
  const dxfRgbColors = [...String(stylesXml || "").matchAll(/<dxf\b[\s\S]*?<\/dxf>/g)]
    .flatMap((match) => [...match[0].matchAll(/\brgb="([0-9A-Fa-f]{8})"/g)].map((colorMatch) => colorMatch[1].toUpperCase()));
  return {
    excelJsLoaded,
    excelJsError,
    worksheetColorFilterCount: worksheetColorFilters.length,
    tableColorFilterCount: tableColorFilters.length,
    colorFilterCount: worksheetColorFilters.length + tableColorFilters.length,
    fillColorFilterCount: [...worksheetColorFilters, ...tableColorFilters].filter((attrs) => !/\bcellColor="0"/.test(attrs)).length,
    fontColorFilterCount: [...worksheetColorFilters, ...tableColorFilters].filter((attrs) => /\bcellColor="0"/.test(attrs)).length,
    dxfCount,
    dxfRgbColors,
  };
}



async function exportedAutoFilterCustomWorkbookState(base64) {
  const buffer = Buffer.from(String(base64 || ""), "base64");
  const zip = await JSZip.loadAsync(buffer);
  const worksheetXmlEntries = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
      .map(async (name) => [name, await zip.file(name).async("text")]),
  );
  const tableXmlEntries = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^xl\/tables\/table\d+\.xml$/i.test(name))
      .map(async (name) => [name, await zip.file(name).async("text")]),
  );
  const customFilters = [...worksheetXmlEntries, ...tableXmlEntries].flatMap(([, xml]) =>
    [...String(xml || "").matchAll(/<customFilter\b([^>]*)\/>/g)].map((match) => {
      const attrs = match[1] || "";
      return {
        operator: attrs.match(/\boperator="([^"]*)"/)?.[1] || "equal",
        value: attrs.match(/\bval="([^"]*)"/)?.[1] || "",
      };
    }),
  );
  return {
    customFilterCount: customFilters.length,
    customFilters,
  };
}



async function exportedAutoFilterDateGroupWorkbookState(base64) {
  const buffer = Buffer.from(String(base64 || ""), "base64");
  const zip = await JSZip.loadAsync(buffer);
  const xmlEntries = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^xl\/(?:worksheets\/sheet|tables\/table)\d+\.xml$/i.test(name))
      .map(async (name) => [name, await zip.file(name).async("text")]),
  );
  const dateGroupItems = xmlEntries.flatMap(([, xml]) =>
    [...String(xml || "").matchAll(/<dateGroupItem\b([^>]*)\/>/g)].map((match) => {
      const attrs = match[1] || "";
      return {
        year: attrs.match(/\byear="([^"]*)"/)?.[1] || "",
        month: attrs.match(/\bmonth="([^"]*)"/)?.[1] || "",
        day: attrs.match(/\bday="([^"]*)"/)?.[1] || "",
        hour: attrs.match(/\bhour="([^"]*)"/)?.[1] || "",
        minute: attrs.match(/\bminute="([^"]*)"/)?.[1] || "",
        second: attrs.match(/\bsecond="([^"]*)"/)?.[1] || "",
        grouping: attrs.match(/\bdateTimeGrouping="([^"]*)"/)?.[1] || "",
      };
    }),
  );
  const dynamicFilters = xmlEntries.flatMap(([, xml]) =>
    [...String(xml || "").matchAll(/<dynamicFilter\b([^>]*)\/>/g)].map((match) => {
      const attrs = match[1] || "";
      return {
        type: attrs.match(/\btype="([^"]*)"/)?.[1] || "",
        val: attrs.match(/\bval="([^"]*)"/)?.[1] || "",
      };
    }),
  );
  const filterColumns = xmlEntries.flatMap(([pathName, xml]) =>
    [...String(xml || "").matchAll(/<filterColumn\b([^>]*)(?:\/>|>[\s\S]*?<\/filterColumn>)/g)].map((match) => {
      const attrs = match[1] || "";
      return {
        path: pathName,
        colId: attrs.match(/\bcolId="([^"]*)"/)?.[1] || "",
        hiddenButton: attrs.match(/\bhiddenButton="([^"]*)"/)?.[1] || "",
        showButton: attrs.match(/\bshowButton="([^"]*)"/)?.[1] || "",
      };
    }),
  );
  return {
    dateGroupItemCount: dateGroupItems.length,
    dateGroupItems,
    dynamicFilters,
    filterColumns,
  };
}



async function importedMultiTableFilterState(page, sheetName) {
  await page.evaluate((name) => window.__webSheetDevSetActiveSheet?.(name), sheetName);
  await page.waitForFunction(
    (name) => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === name,
    sheetName,
    { timeout: 5000 },
  );
  await page.waitForSelector('.cell-filter-button[data-filter-col="5"]', { timeout: 5000 });
  const before = await page.evaluate((name) => {
    const sheet = window.__webSheetDevSheetModel?.(name);
    return {
      sheetRange: sheet?.autoFilter?.range || null,
      tableRanges: (sheet?.tables || []).map((table) => ({
        name: table.name,
        showFilterButton: table.showFilterButton,
        range: table.range,
        autoFilterRange: table.autoFilterState?.range || null,
      })),
      buttons: [...document.querySelectorAll(".cell-filter-button")].map((button) => ({
        row: Number(button.getAttribute("data-filter-row")),
        col: Number(button.getAttribute("data-filter-col")),
        filtered: button.classList.contains("is-filtered"),
      })),
      rows: [2, 3, 4].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        left: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        right: document.querySelector(`.sheet-cell[data-address="D${row}"] .cell-content`)?.textContent || "",
      })),
    };
  }, sheetName);
  await page.locator('.cell-filter-button[data-filter-col="4"]').click();
  await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
  const afterRightMenu = await page.evaluate((name) => {
    const sheet = window.__webSheetDevSheetModel?.(name);
    return {
      sheetRange: sheet?.autoFilter?.range || null,
      menuValues: [...document.querySelectorAll("#autoFilterMenu .auto-filter-values .auto-filter-value")].map((item) =>
        item.textContent.replace(/\s+/g, " ").trim(),
      ),
    };
  }, sheetName);
  await page.keyboard.press("Escape");
  await page.waitForSelector("#autoFilterMenu", { state: "detached", timeout: 5000 });

  await page.locator('.cell-filter-button[data-filter-col="4"]').click();
  await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
  await page.evaluate(() => {
    const menu = document.querySelector("#autoFilterMenu");
    if (!menu) return;
    menu.dataset.valueSelectionDirty = "true";
    menu.querySelectorAll("[data-auto-filter-value]").forEach((input) => {
      input.checked = input.dataset.autoFilterValue === "Blue";
    });
    const selectAll = menu.querySelector("[data-auto-filter-select-all]");
    if (selectAll) {
      selectAll.checked = false;
      selectAll.indeterminate = true;
    }
  });
  await page.locator('#autoFilterMenu [data-auto-filter-action="ok"]').click();
  await page.waitForSelector("#autoFilterMenu", { state: "detached", timeout: 5000 });
  await page.waitForFunction(
    () => [2, 3, 4].every((row) => getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none"),
    null,
    { timeout: 5000 },
  );
  const afterRightBlue = await page.evaluate((name) => {
    const sheet = window.__webSheetDevSheetModel?.(name);
    return {
      tableCriteria: (sheet?.tables || []).map((table) => ({
        name: table.name,
        criteria: table.autoFilterState?.criteria || {},
      })),
      buttons: [...document.querySelectorAll(".cell-filter-button")].map((button) => ({
        col: Number(button.getAttribute("data-filter-col")),
        filtered: button.classList.contains("is-filtered"),
      })),
      rows: [2, 3, 4].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
      })),
    };
  }, sheetName);

  await page.locator('.sheet-cell[data-address="D1"]').click();
  await page.evaluate(() => document.querySelector('[data-ribbon-tab="data"]')?.click());
  await page.waitForSelector('[data-ribbon-panel="data"].is-active', { timeout: 5000 });
  await page.locator('[data-ribbon-command="data-clear-filter"]').click();
  await page.waitForFunction(
    () => getComputedStyle(document.querySelector('.row-header[data-row="2"]')).display !== "none" &&
      [3, 4].every((row) => getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none"),
    null,
    { timeout: 5000 },
  );
  const afterRightRibbonClear = await page.evaluate((name) => {
    const sheet = window.__webSheetDevSheetModel?.(name);
    return {
      tableCriteria: (sheet?.tables || []).map((table) => ({
        name: table.name,
        criteria: table.autoFilterState?.criteria || {},
      })),
      buttons: [...document.querySelectorAll(".cell-filter-button")].map((button) => ({
        col: Number(button.getAttribute("data-filter-col")),
        filtered: button.classList.contains("is-filtered"),
      })),
      rows: [2, 3, 4].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
      })),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetName);

  await page.locator('.cell-filter-button[data-filter-col="4"]').click();
  await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
  await page.evaluate(() => {
    const menu = document.querySelector("#autoFilterMenu");
    if (!menu) return;
    menu.dataset.valueSelectionDirty = "true";
    menu.querySelectorAll("[data-auto-filter-value]").forEach((input) => {
      input.checked = input.dataset.autoFilterValue === "Blue";
    });
    const selectAll = menu.querySelector("[data-auto-filter-select-all]");
    if (selectAll) {
      selectAll.checked = false;
      selectAll.indeterminate = true;
    }
  });
  await page.locator('#autoFilterMenu [data-auto-filter-action="ok"]').click();
  await page.waitForSelector("#autoFilterMenu", { state: "detached", timeout: 5000 });
  await page.waitForFunction(
    () => [2, 3, 4].every((row) => getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none"),
    null,
    { timeout: 5000 },
  );

  await page.locator('.cell-filter-button[data-filter-col="4"]').click();
  await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
  await page.locator('#autoFilterMenu [data-auto-filter-action="clear"]').click();
  await page.waitForSelector("#autoFilterMenu", { state: "detached", timeout: 5000 });
  await page.waitForFunction(
    () => getComputedStyle(document.querySelector('.row-header[data-row="2"]')).display !== "none" &&
      [3, 4].every((row) => getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none"),
    null,
    { timeout: 5000 },
  );
  const afterRightClear = await page.evaluate((name) => {
    const sheet = window.__webSheetDevSheetModel?.(name);
    return {
      tableCriteria: (sheet?.tables || []).map((table) => ({
        name: table.name,
        criteria: table.autoFilterState?.criteria || {},
      })),
      buttons: [...document.querySelectorAll(".cell-filter-button")].map((button) => ({
        col: Number(button.getAttribute("data-filter-col")),
        filtered: button.classList.contains("is-filtered"),
      })),
      rows: [2, 3, 4].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
      })),
    };
  }, sheetName);

  await page.locator('.cell-filter-button[data-filter-col="4"]').click();
  await page.waitForSelector("#autoFilterMenu", { timeout: 5000 });
  await page.evaluate(() => {
    const menu = document.querySelector("#autoFilterMenu");
    if (!menu) return;
    menu.dataset.valueSelectionDirty = "true";
    menu.querySelectorAll("[data-auto-filter-value]").forEach((input) => {
      input.checked = input.dataset.autoFilterValue === "Red";
    });
    const selectAll = menu.querySelector("[data-auto-filter-select-all]");
    if (selectAll) {
      selectAll.checked = false;
      selectAll.indeterminate = true;
    }
  });
  await page.locator('#autoFilterMenu [data-auto-filter-action="ok"]').click();
  await page.waitForSelector("#autoFilterMenu", { state: "detached", timeout: 5000 });
  await page.waitForFunction(
    () => getComputedStyle(document.querySelector('.row-header[data-row="2"]')).display !== "none" &&
      [3, 4].every((row) => getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none"),
    null,
    { timeout: 5000 },
  );
  await editCellWithFormulaBar(page, "D2", "Blue");
  await page.waitForFunction(
    () => document.querySelector('.sheet-cell[data-address="D2"] .cell-content')?.textContent === "Blue" &&
      getComputedStyle(document.querySelector('.row-header[data-row="2"]')).display !== "none",
    null,
    { timeout: 5000 },
  );
  await page.locator('[data-ribbon-command="data-reapply-filter"]').click();
  await page.waitForFunction(
    () => [2, 3, 4].every((row) => getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none"),
    null,
    { timeout: 5000 },
  );
  const afterRibbonReapply = await page.evaluate((name) => {
    const sheet = window.__webSheetDevSheetModel?.(name);
    return {
      tableCriteria: (sheet?.tables || []).map((table) => ({
        name: table.name,
        criteria: table.autoFilterState?.criteria || {},
      })),
      buttons: [...document.querySelectorAll(".cell-filter-button")].map((button) => ({
        col: Number(button.getAttribute("data-filter-col")),
        filtered: button.classList.contains("is-filtered"),
      })),
      rows: [2, 3, 4].map((row) => ({
        row,
        hidden: getComputedStyle(document.querySelector(`.row-header[data-row="${row}"]`)).display === "none",
        right: document.querySelector(`.sheet-cell[data-address="D${row}"] .cell-content`)?.textContent || "",
      })),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  }, sheetName);
  return { before, afterRightMenu, afterRightBlue, afterRightRibbonClear, afterRightClear, afterRibbonReapply };
}

export {
  createSortCustomListWorkbook,
  createSortDateWorkbook,
  createSortFormulaWorkbook,
  exerciseAutoFilterColorImportExport,
  exerciseAutoFilterConditionMenu,
  exerciseDataValidationMenu,
  exerciseHomeEditingCommands,
  exerciseImportedAutoFilterCriteria,
  exportedAutoFilterColorWorkbookState,
  exportedAutoFilterCustomWorkbookState,
  exportedAutoFilterDateGroupWorkbookState,
  exportedTableSortWorkbookState,
  importedAutoFilterSheetState,
  importedHiddenButtonAutoFilterState,
  importedMultiTableFilterState,
  importedTableSortOnlyState,
  workbookBufferForExcelJsSmoke,
};
