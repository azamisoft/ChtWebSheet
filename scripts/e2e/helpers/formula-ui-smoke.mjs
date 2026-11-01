// Sequential E2E feature module.
import {
  createFormulaUiWorkbook,
  exportedStandaloneModelFromHtml,
  looksLikeStandaloneWorkbookHtml,
} from "./fixtures-and-model.mjs";
import {
  appUrl,
  browser,
  chromePath,
  chromium,
  path,
  pathToFileURL,
  readFile,
} from "../../e2e-real-upload.mjs";


async function runFormulaUiSmokeOnly() {
  const formulaUiWorkbookPath = await createFormulaUiWorkbook();
  const browser = await chromium.launch({
    headless: true,
    executablePath: chromePath,
  });
  try {
    const context = await browser.newContext({
      viewport: { width: 1280, height: 720 },
      acceptDownloads: true,
      permissions: ["clipboard-read", "clipboard-write"],
    });
    const page = await context.newPage();
    const messages = [];
    page.on("console", (message) => {
      if (message.type() === "error" && !/Failed to load resource: the server responded with a status of 404/i.test(message.text())) {
        messages.push(`${message.type()}: ${message.text()}`);
      }
    });
    page.on("pageerror", (error) => {
      messages.push(`pageerror: ${error.stack || error.message}`);
    });
    const selectAllShortcut = process.platform === "darwin" ? "Meta+A" : "Control+A";
    const saveShortcut = process.platform === "darwin" ? "Meta+S" : "Control+S";
    const resetFormulaInput = async (value) => {
      await page.locator("#formulaInput").click();
      await page.keyboard.press(selectAllShortcut);
      await page.keyboard.press("Backspace");
      if (value) await page.keyboard.type(value);
    };
    const quoteSheetNameForFormulaTest = (name) => /^[A-Za-z_][A-Za-z0-9_]*$/.test(String(name || ""))
      ? String(name)
      : `'${String(name || "").replace(/'/g, "''")}'`;
    const popupSnapshot = () => page.evaluate(() => {
      const limit = 80;
      return {
        names: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].slice(0, limit).map((item) => item.textContent),
        categories: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-category")].slice(0, limit).map((item) => item.textContent),
        syntaxes: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-syntax")].slice(0, limit).map((item) => item.textContent),
        descriptions: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-description")].slice(0, limit).map((item) => item.textContent),
        active: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active .formula-autocomplete-name")?.textContent || "",
        itemCount: document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length,
        syntaxCount: document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-syntax").length,
        descriptionCount: document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-description").length,
        role: document.querySelector("#formulaAutocomplete")?.getAttribute("role") || "",
        ariaLabel: document.querySelector("#formulaAutocomplete")?.getAttribute("aria-label") || "",
        activeDetailName: document.querySelector("#formulaAutocomplete .formula-autocomplete-active-name")?.textContent || "",
        activeDetailCategory: document.querySelector("#formulaAutocomplete .formula-autocomplete-active-category")?.textContent || "",
        activeDetailSyntax: document.querySelector("#formulaAutocomplete .formula-autocomplete-active-syntax")?.textContent || "",
        activeDetailDescription: document.querySelector("#formulaAutocomplete .formula-autocomplete-active-description")?.textContent || "",
        inputValue: document.querySelector("#formulaInput")?.value || "",
      };
    });
    const functionArgumentsSnapshot = () => page.evaluate(() => {
      const rows = [...document.querySelectorAll("[data-function-argument-row]")].map((row) => ({
        index: Number(row.dataset.functionArgumentRow),
        label: row.querySelector(".function-argument-name")?.textContent?.trim().replace(/:$/u, "") || "",
        active: row.classList.contains("is-active"),
        inputDescribedBy: row.querySelector("[data-formula-argument-input]")?.getAttribute("aria-describedby") || "",
      }));
      return {
        title: document.querySelector("#functionArgumentsTitle")?.textContent || "",
        functionName: document.querySelector(".function-argument-heading strong")?.textContent || "",
        description: document.querySelector("[data-function-argument-description]")?.textContent?.replace(/\s+/g, " ").trim() || "",
        descriptionId: document.querySelector("[data-function-argument-description]")?.id || "",
        rows,
        activeLabels: rows.filter((row) => row.active).map((row) => row.label),
        syntax: document.querySelector(".function-argument-syntax")?.textContent || "",
      };
    });

    await page.goto(`${appUrl}?formula-ui-smoke=${Date.now()}`, { waitUntil: "load" });
    const formulaUiWorkbookBase64 = (await readFile(formulaUiWorkbookPath)).toString("base64");
    await page.evaluate(
      async ({ base64, name }) => {
        const binary = atob(base64);
        const bytes = new Uint8Array(binary.length);
        for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
        await window.__webSheetDevLoadBlob?.(
          new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
          name,
        );
      },
      { base64: formulaUiWorkbookBase64, name: path.basename(formulaUiWorkbookPath) },
    );
    await page.waitForFunction(
      () => {
        const alert = document.querySelector("#appAlert")?.textContent || "";
        const formulaInput = document.querySelector("#formulaInput");
        const workspace = document.querySelector("#sheetWorkspace");
        return /読み込みに失敗|读取失败/.test(alert) ||
          (Boolean(document.querySelector(".sheet-cell[data-address]")) && formulaInput && !formulaInput.disabled && !workspace?.classList.contains("hidden"));
      },
      null,
      { timeout: 60000 },
    );
    const loadState = await page.evaluate(() => ({
      status: document.querySelector("#statusText")?.textContent || "",
      alert: document.querySelector("#appAlert")?.textContent || "",
      formulaDisabled: document.querySelector("#formulaInput")?.disabled ?? true,
      workspaceHidden: document.querySelector("#sheetWorkspace")?.classList.contains("hidden") ?? true,
    }));
    if (/読み込みに失敗|读取失败/.test(loadState.alert)) {
      throw new Error(`Formula UI smoke workbook load failed: ${JSON.stringify(loadState)}`);
    }
    await page.waitForSelector(".sheet-cell[data-address]", { timeout: 60000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    try {
      await page.waitForFunction(
        () =>
          document.querySelector("#formulaInput") &&
          !document.querySelector("#formulaInput").disabled &&
          !document.querySelector("#sheetWorkspace")?.classList.contains("hidden"),
        null,
        { timeout: 15000 },
      );
    } catch (error) {
      const debugState = await page.evaluate(() => ({
        status: document.querySelector("#statusText")?.textContent || "",
        alert: document.querySelector("#appAlert")?.textContent || "",
        selectedLabel: document.querySelector("#selectedCellLabel")?.textContent || "",
        formulaDisabled: document.querySelector("#formulaInput")?.disabled ?? true,
        workspaceHidden: document.querySelector("#sheetWorkspace")?.classList.contains("hidden") ?? true,
        visibleCells: document.querySelectorAll(".sheet-cell[data-address]").length,
      }));
      throw new Error(`Formula UI smoke formula bar did not become editable: ${JSON.stringify(debugState)}`);
    }

    await page.evaluate(() => {
      window.__webSheetDevSetActiveSheet?.("CHT");
      window.__webSheetDevAddDefinedName?.({
        name: "SalesTotal",
        refersTo: "'CHT'!$B$9:$B$11",
        scope: "workbook",
      });
      window.__webSheetDevAddTable?.("CHT", {
        name: "SalesTable",
        displayName: "SalesTable",
        range: { sheetIndex: 0, top: 9, left: 2, bottom: 12, right: 22 },
        showTotalsRow: true,
        columns: [
          { name: "数量" },
          { name: "Sales Amount" },
          { name: "区分" },
          { name: "税#率" },
          { name: "O'Brien" },
          { name: "端]数" },
          { name: "#Data" },
          { name: "@担当" },
          { name: "#This Row" },
          { name: "金額,税込" },
          { name: "比率.%" },
          { name: "A:B" },
          { name: "費用?" },
          { name: "売上$" },
          { name: "A+B" },
          { name: "値(税)" },
          { name: "担当,@税込" },
          { name: "#All" },
          { name: "#Headers" },
          { name: "#Totals" },
          { name: "始[数" },
        ],
      });
    });
    await page.waitForFunction(() => window.__webSheetDevDefinedNames?.().some((definition) => definition.name === "SalesTotal"), null, { timeout: 5000 });

    const definitions = await page.evaluate(() => {
      const all = window.__webSheetDevFormulaAutocompleteDefinitions?.() || [];
      const wanted = ["SUM", "SUMPRODUCT", "XLOOKUP", "FILTER", "PMT", "LOGNORM.DIST", "LOGNORM.INV", "DELTA", "TBILLEQ", "ISPMT"];
      return Object.fromEntries(wanted.map((name) => [name, all.find((definition) => definition.name === name) || null]));
    });
    const sheetNames = await page.evaluate(() => [...document.querySelectorAll(".sheet-tab")].map((tab) => tab.textContent.trim()).filter(Boolean));
    const qualifiedSheetName = sheetNames.find((name) => /[^A-Za-z0-9_]/.test(name)) || sheetNames[0] || "CHT";
    const qualifiedSheetPrefix = quoteSheetNameForFormulaTest(qualifiedSheetName);
    const sheetAutocompleteQuery = qualifiedSheetName.slice(0, Math.max(1, Math.min(2, qualifiedSheetName.length)));
    const apostropheSheetName = sheetNames.find((name) => name.includes("'")) || "O'Brien Sheet";
    const apostropheSheetPrefix = quoteSheetNameForFormulaTest(apostropheSheetName);
    const apostropheSheetAutocompleteQuery = apostropheSheetName.slice(0, 1);
    const apostropheSheetEscapedQuery = apostropheSheetPrefix.slice(1, Math.min(apostropheSheetPrefix.length - 1, 5));

    await resetFormulaInput("=su");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const sumOpen = await popupSnapshot();
    await page.keyboard.press("ArrowDown");
    await page.waitForFunction(() => document.querySelector("#formulaAutocomplete .formula-autocomplete-active-name")?.textContent === "SUMIF", null, { timeout: 5000 });
    const sumAfterArrow = await popupSnapshot();
    await page.keyboard.press("Home");
    await page.waitForFunction(() => document.querySelector("#formulaAutocomplete .formula-autocomplete-active-name")?.textContent === "SUM", null, { timeout: 5000 });
    await page.keyboard.press("Tab");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SUM(", null, { timeout: 5000 });
    await page.waitForSelector("#formulaAutocomplete .formula-autocomplete-hint", { timeout: 5000 });
    const sumAccepted = await page.evaluate(() => ({
      inputValue: document.querySelector("#formulaInput")?.value || "",
      hint: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint")?.textContent?.replace(/\s+/g, " ").trim() || "",
      activeArgument: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint .is-active")?.textContent || "",
      activeDescription: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint-detail")?.textContent || "",
      listVisible: Boolean(document.querySelector("#formulaAutocomplete [data-formula-autocomplete-item]")),
    }));
    await resetFormulaInput("=av");
    await page.waitForFunction(() => document.querySelector("#formulaAutocomplete .formula-autocomplete-active-name")?.textContent === "AVERAGE", null, { timeout: 5000 });
    await page.keyboard.press("(");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=AVERAGE(", null, { timeout: 5000 });
    await page.waitForSelector("#formulaAutocomplete .formula-autocomplete-hint", { timeout: 5000 });
    const parenAccepted = await page.evaluate(() => ({
      inputValue: document.querySelector("#formulaInput")?.value || "",
      hint: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint")?.textContent?.replace(/\s+/g, " ").trim() || "",
      activeArgument: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint .is-active")?.textContent || "",
      listVisible: Boolean(document.querySelector("#formulaAutocomplete [data-formula-autocomplete-item]")),
    }));
    await resetFormulaInput("=");
    await page.keyboard.press("(");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=(", null, { timeout: 5000 });
    const plainParen = await page.evaluate(() => ({
      inputValue: document.querySelector("#formulaInput")?.value || "",
      popupVisible: Boolean(document.querySelector("#formulaAutocomplete")),
    }));
    await resetFormulaInput("=SUM(");
    await page.waitForSelector("#formulaAutocomplete .formula-autocomplete-hint", { timeout: 5000 });
    await page.keyboard.type("1,2,");
    await page.waitForFunction(
      () =>
        document.querySelector("#formulaInput")?.value === "=SUM(1,2," &&
        /number3/.test(document.querySelector("#formulaAutocomplete .formula-autocomplete-hint .is-active")?.textContent || ""),
      null,
      { timeout: 5000 },
    );
    const repeatedArgument = await page.evaluate(() => ({
      hint: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint")?.textContent?.replace(/\s+/g, " ").trim() || "",
      activeArgument: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint .is-active")?.textContent || "",
      activeDescription: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint-detail")?.textContent || "",
    }));

    await resetFormulaInput(`=${qualifiedSheetPrefix}!B2:C3`);
    await page.locator("#formulaInput").evaluate((input) => {
      const start = input.value.indexOf("B2");
      input.setSelectionRange(start, start);
    });
    const qualifiedReferenceF4 = [];
    for (let index = 0; index < 4; index += 1) {
      await page.keyboard.press("F4");
      qualifiedReferenceF4.push(await page.locator("#formulaInput").inputValue());
    }

    await resetFormulaInput(`=${sheetAutocompleteQuery}`);
    await page.waitForFunction(
      (sheetName) => [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].some((item) => item.textContent === sheetName),
      qualifiedSheetName,
      { timeout: 5000 },
    );
    const sheetReferenceOpen = await popupSnapshot();
    await page.keyboard.press("Tab");
    await page.waitForFunction(
      (expected) => document.querySelector("#formulaInput")?.value === expected,
      `=${qualifiedSheetPrefix}!`,
      { timeout: 5000 },
    );
    const sheetReferenceAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput(`='${sheetAutocompleteQuery}`);
    await page.waitForFunction(
      (sheetName) => [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].some((item) => item.textContent === sheetName),
      qualifiedSheetName,
      { timeout: 5000 },
    );
    const quotedSheetReferenceOpen = await popupSnapshot();
    await page.keyboard.press("Tab");
    await page.waitForFunction(
      (expected) => document.querySelector("#formulaInput")?.value === expected,
      `=${qualifiedSheetPrefix}!`,
      { timeout: 5000 },
    );
    const quotedSheetReferenceAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput(`=${apostropheSheetAutocompleteQuery}`);
    await page.waitForFunction(
      (sheetName) => [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].some((item) => item.textContent === sheetName),
      apostropheSheetName,
      { timeout: 5000 },
    );
    const apostropheSheetReferenceOpen = await popupSnapshot();
    if (!(await page.evaluate((sheetName) => {
      const items = [...document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]")];
      const item = items.find((button) => button.querySelector(".formula-autocomplete-name")?.textContent === sheetName);
      item?.click();
      return Boolean(item);
    }, apostropheSheetName))) throw new Error("apostrophe sheet candidate not found");
    await page.waitForFunction(
      (expected) => document.querySelector("#formulaInput")?.value === expected,
      `=${apostropheSheetPrefix}!`,
      { timeout: 5000 },
    );
    const apostropheSheetReferenceAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput(`='${apostropheSheetEscapedQuery}`);
    await page.waitForFunction(
      (sheetName) => [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].some((item) => item.textContent === sheetName),
      apostropheSheetName,
      { timeout: 5000 },
    );
    const quotedApostropheSheetReferenceOpen = await popupSnapshot();
    if (!(await page.evaluate((sheetName) => {
      const items = [...document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]")];
      const item = items.find((button) => button.querySelector(".formula-autocomplete-name")?.textContent === sheetName);
      item?.click();
      return Boolean(item);
    }, apostropheSheetName))) throw new Error("quoted apostrophe sheet candidate not found");
    await page.waitForFunction(
      (expected) => document.querySelector("#formulaInput")?.value === expected,
      `=${apostropheSheetPrefix}!`,
      { timeout: 5000 },
    );
    const quotedApostropheSheetReferenceAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput(`=${apostropheSheetPrefix}!B2:C3`);
    await page.locator("#formulaInput").evaluate((input) => {
      const start = input.value.indexOf("B2");
      input.setSelectionRange(start, start);
    });
    const apostropheReferenceF4 = [];
    for (let index = 0; index < 4; index += 1) {
      await page.keyboard.press("F4");
      apostropheReferenceF4.push(await page.locator("#formulaInput").inputValue());
    }

    await resetFormulaInput("=lognorm.");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const lognormOpen = await popupSnapshot();

    await resetFormulaInput("=delta");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const deltaOpen = await popupSnapshot();

    await resetFormulaInput("=xlookup");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const xlookupOpen = await popupSnapshot();
    await page.keyboard.press("Tab");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=XLOOKUP(", null, { timeout: 5000 });
    await page.keyboard.press("Control+A");
    await page.waitForSelector('[data-data-dialog="function-arguments"]', { timeout: 5000 });
    const xlookupDialogInitial = await functionArgumentsSnapshot();
    await page.locator('[data-formula-argument-input][data-formula-argument-index="2"]').focus();
    await page.waitForFunction(
      () => document.querySelector('[data-function-argument-row="2"]')?.classList.contains("is-active"),
      null,
      { timeout: 5000 },
    );
    const xlookupDialogReturnArray = await functionArgumentsSnapshot();
    await page.locator('[data-data-dialog-action="function-argument-reference-toggle"][data-formula-argument-index="2"]').click();
    await page.waitForFunction(
      () =>
        document.querySelector('[data-data-dialog="function-arguments"]')?.classList.contains("is-reference-collapsed") &&
        document.querySelector('[data-function-argument-row="2"]')?.classList.contains("is-active"),
      null,
      { timeout: 5000 },
    );
    const xlookupDialogReferenceCollapsed = await functionArgumentsSnapshot();
    await page.keyboard.press("Escape");
    await page.waitForFunction(
      () => !document.querySelector('[data-data-dialog="function-arguments"]')?.classList.contains("is-reference-collapsed"),
      null,
      { timeout: 5000 },
    );
    const xlookupReturnArrayInput = page.locator('[data-formula-argument-input][data-formula-argument-index="2"]');
    await xlookupReturnArrayInput.fill(`${qualifiedSheetPrefix}!B2:C3`);
    await xlookupReturnArrayInput.evaluate((input) => {
      const start = input.value.indexOf("B2");
      input.setSelectionRange(start, start);
    });
    await page.keyboard.press("F4");
    const xlookupDialogReferenceF4 = await xlookupReturnArrayInput.inputValue();
    await page.keyboard.press("Escape");
    await page.waitForFunction(() => !document.querySelector('[data-data-dialog="function-arguments"]'), null, { timeout: 5000 });

    await resetFormulaInput("=filter");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const filterOpen = await popupSnapshot();

    await resetFormulaInput("=pmt");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const pmtOpen = await popupSnapshot();

    await resetFormulaInput("=SalesTo");
    await page.waitForFunction(() => document.querySelector("#formulaAutocomplete .formula-autocomplete-active-name")?.textContent === "SalesTotal", null, { timeout: 5000 });
    const definedNameOpen = await popupSnapshot();
    await page.keyboard.press("Tab");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTotal", null, { timeout: 5000 });
    const definedNameAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTa");
    await page.waitForFunction(() => document.querySelector("#formulaAutocomplete .formula-autocomplete-active-name")?.textContent === "SalesTable", null, { timeout: 5000 });
    const tableNameOpen = await popupSnapshot();
    await page.keyboard.press("Tab");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable", null, { timeout: 5000 });
    const tableNameAccepted = await page.locator("#formulaInput").inputValue();

    const clickAutocompleteItemByName = async (name) => page.evaluate((targetName) => {
      const items = [...document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]")];
      const item = items.find((button) => button.querySelector(".formula-autocomplete-name")?.textContent === targetName);
      item?.click();
      return Boolean(item);
    }, name);
    const clickAutocompleteItemBySyntax = async (syntax) => page.evaluate((targetSyntax) => {
      const items = [...document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]")];
      const item = items.find((button) => button.querySelector(".formula-autocomplete-syntax")?.textContent === targetSyntax);
      item?.click();
      return Boolean(item);
    }, syntax);

    await resetFormulaInput("=SalesTable[");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const tableStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("Sales Amount"))) throw new Error("Sales Amount structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[Sales Amount]", null, { timeout: 5000 });
    const tableStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[税");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const escapedHashColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("税#率"))) throw new Error("escaped # structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[税'#率]", null, { timeout: 5000 });
    const escapedHashColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[O");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const escapedApostropheColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("O'Brien"))) throw new Error("escaped apostrophe structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[O''Brien]", null, { timeout: 5000 });
    const escapedApostropheColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[O''");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const escapedApostropheTypedEscapeColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[O''Brien]"))) throw new Error("escaped apostrophe typed structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[O''Brien]", null, { timeout: 5000 });
    const escapedApostropheTypedEscapeColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[端");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const escapedBracketColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("端]数"))) throw new Error("escaped bracket structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[端']数]", null, { timeout: 5000 });
    const escapedBracketColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[始");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const escapedOpenBracketColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("始[数"))) throw new Error("escaped open bracket structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[始'[数]", null, { timeout: 5000 });
    const escapedOpenBracketColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[始'[");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const escapedOpenBracketTypedEscapeColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[始'[数]"))) throw new Error("escaped open bracket typed structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[始'[数]", null, { timeout: 5000 });
    const escapedOpenBracketTypedEscapeColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[#D");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const escapedSpecialNameColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable['#Data]"))) throw new Error("escaped #Data column structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable['#Data]", null, { timeout: 5000 });
    const escapedSpecialNameColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[担");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const escapedAtColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable['@担当]"))) throw new Error("escaped @ structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable['@担当]", null, { timeout: 5000 });
    const escapedAtColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[#This");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const escapedThisRowNameColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable['#This Row]"))) throw new Error("escaped #This Row column structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable['#This Row]", null, { timeout: 5000 });
    const escapedThisRowNameColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[税込");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const doubleWrappedCommaColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[金額,税込]]"))) throw new Error("double-wrapped comma structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[金額,税込]]", null, { timeout: 5000 });
    const doubleWrappedCommaColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[比");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const doubleWrappedDotPercentColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[比率.%]]"))) throw new Error("double-wrapped dot/percent structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[比率.%]]", null, { timeout: 5000 });
    const doubleWrappedDotPercentColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[A");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const doubleWrappedColonColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[A:B]]"))) throw new Error("double-wrapped colon structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[A:B]]", null, { timeout: 5000 });
    const doubleWrappedColonColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[費");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const doubleWrappedQuestionColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[費用?]]"))) throw new Error("double-wrapped question structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[費用?]]", null, { timeout: 5000 });
    const doubleWrappedQuestionColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[売");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const doubleWrappedDollarColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[売上$]]"))) throw new Error("double-wrapped dollar structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[売上$]]", null, { timeout: 5000 });
    const doubleWrappedDollarColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[A+");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const doubleWrappedPlusColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[A+B]]"))) throw new Error("double-wrapped plus structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[A+B]]", null, { timeout: 5000 });
    const doubleWrappedPlusColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[値");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const doubleWrappedParenColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[値(税)]]"))) throw new Error("double-wrapped parenthesized structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[値(税)]]", null, { timeout: 5000 });
    const doubleWrappedParenColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[担当");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const doubleWrappedEscapedAtCommaColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[担当,'@税込]]"))) throw new Error("double-wrapped escaped @ comma structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[担当,'@税込]]", null, { timeout: 5000 });
    const doubleWrappedEscapedAtCommaColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[#A");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const escapedAllNameColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable['#All]"))) throw new Error("escaped #All column structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable['#All]", null, { timeout: 5000 });
    const escapedAllNameColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable['#A");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const escapedAllNameTypedEscapeColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable['#All]"))) throw new Error("escaped #All typed structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable['#All]", null, { timeout: 5000 });
    const escapedAllNameTypedEscapeColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[#H");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const escapedHeadersNameColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable['#Headers]"))) throw new Error("escaped #Headers column structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable['#Headers]", null, { timeout: 5000 });
    const escapedHeadersNameColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[#T");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const escapedTotalsNameColumnStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable['#Totals]"))) throw new Error("escaped #Totals column structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable['#Totals]", null, { timeout: 5000 });
    const escapedTotalsNameColumnStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const nestedSpecialStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("#All"))) throw new Error("#All nested structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#All]]", null, { timeout: 5000 });
    const nestedSpecialStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#This");
    await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });
    const nestedThisRowStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("#This Row"))) throw new Error("#This Row nested structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#This Row]]", null, { timeout: 5000 });
    const nestedThisRowStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await page.keyboard.press("Escape");
    await page.evaluate(() => window.__webSheetDevSelectRange?.(10, 2, 10, 2));
    await page.waitForFunction(() => document.querySelector("#selectedCellLabel")?.textContent === "B10", null, { timeout: 5000 });
    await resetFormulaInput("=[@");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const currentRowStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("Sales Amount"))) throw new Error("Sales Amount current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=[@[Sales Amount]]", null, { timeout: 5000 });
    const currentRowStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[@");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const namedCurrentRowStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("Sales Amount"))) throw new Error("Sales Amount named current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[@[Sales Amount]]", null, { timeout: 5000 });
    const namedCurrentRowStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[@税");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const namedCurrentRowEscapedHashOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("税#率"))) throw new Error("escaped # named current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[@[税'#率]]", null, { timeout: 5000 });
    const namedCurrentRowEscapedHashAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=[@O");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const currentRowEscapedApostropheOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("O'Brien"))) throw new Error("escaped apostrophe current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=[@[O''Brien]]", null, { timeout: 5000 });
    const currentRowEscapedApostropheAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=[@端");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const currentRowEscapedBracketOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("端]数"))) throw new Error("escaped bracket current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=[@[端']数]]", null, { timeout: 5000 });
    const currentRowEscapedBracketAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=[@始'[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const currentRowEscapedOpenBracketTypedEscapeOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("[@[始'[数]]"))) throw new Error("escaped open bracket typed current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=[@[始'[数]]", null, { timeout: 5000 });
    const currentRowEscapedOpenBracketTypedEscapeAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=[@担");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const currentRowEscapedAtOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("[@['@担当]]"))) throw new Error("escaped @ current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=[@['@担当]]", null, { timeout: 5000 });
    const currentRowEscapedAtAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=[@['@");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const currentRowEscapedAtTypedEscapeOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("[@['@担当]]"))) throw new Error("escaped @ typed current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=[@['@担当]]", null, { timeout: 5000 });
    const currentRowEscapedAtTypedEscapeAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[@担");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const namedCurrentRowEscapedAtOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[@['@担当]]"))) throw new Error("escaped @ named current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[@['@担当]]", null, { timeout: 5000 });
    const namedCurrentRowEscapedAtAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=[@税込");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const currentRowDoubleWrappedCommaOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("[@[金額,税込]]"))) throw new Error("double-wrapped comma current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=[@[金額,税込]]", null, { timeout: 5000 });
    const currentRowDoubleWrappedCommaAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[@比");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const namedCurrentRowDoubleWrappedDotPercentOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[@[比率.%]]"))) throw new Error("double-wrapped dot/percent named current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[@[比率.%]]", null, { timeout: 5000 });
    const namedCurrentRowDoubleWrappedDotPercentAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=[@費");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const currentRowDoubleWrappedQuestionOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("[@[費用?]]"))) throw new Error("double-wrapped question current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=[@[費用?]]", null, { timeout: 5000 });
    const currentRowDoubleWrappedQuestionAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[@値");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const namedCurrentRowDoubleWrappedParenOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[@[値(税)]]"))) throw new Error("double-wrapped parenthesized named current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[@[値(税)]]", null, { timeout: 5000 });
    const namedCurrentRowDoubleWrappedParenAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=[@担当");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const currentRowDoubleWrappedEscapedAtCommaOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("[@[担当,'@税込]]"))) throw new Error("double-wrapped escaped @ comma current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=[@[担当,'@税込]]", null, { timeout: 5000 });
    const currentRowDoubleWrappedEscapedAtCommaAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=[@#A");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const currentRowEscapedAllNameOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("[@['#All]]"))) throw new Error("escaped #All current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=[@['#All]]", null, { timeout: 5000 });
    const currentRowEscapedAllNameAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[@#T");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const namedCurrentRowEscapedTotalsNameOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[@['#Totals]]"))) throw new Error("escaped #Totals named current-row candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[@['#Totals]]", null, { timeout: 5000 });
    const namedCurrentRowEscapedTotalsNameAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[@[税'#率]:[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const namedCurrentRowEscapedRangeOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("端]数"))) throw new Error("escaped named current-row range candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[@[税'#率]:[端']数]]", null, { timeout: 5000 });
    const namedCurrentRowEscapedRangeAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=[@[O''Brien]:[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const currentRowEscapedRangeOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("端]数"))) throw new Error("escaped current-row range candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=[@[O''Brien]:[端']数]]", null, { timeout: 5000 });
    const currentRowEscapedRangeAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#Data],[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const dataQualifierStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("Sales Amount"))) throw new Error("Sales Amount #Data structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#Data],[Sales Amount]]", null, { timeout: 5000 });
    const dataQualifierStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#Headers],[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const headersQualifierStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("Sales Amount"))) throw new Error("Sales Amount #Headers structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#Headers],[Sales Amount]]", null, { timeout: 5000 });
    const headersQualifierStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#Totals],[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const totalsQualifierStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("Sales Amount"))) throw new Error("Sales Amount #Totals structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#Totals],[Sales Amount]]", null, { timeout: 5000 });
    const totalsQualifierStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#Data],[税");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const dataQualifierEscapedHashOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("税#率"))) throw new Error("escaped # #Data structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#Data],[税'#率]]", null, { timeout: 5000 });
    const dataQualifierEscapedHashAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#Totals],[端");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const totalsQualifierEscapedBracketOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("端]数"))) throw new Error("escaped bracket #Totals structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#Totals],[端']数]]", null, { timeout: 5000 });
    const totalsQualifierEscapedBracketAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#Data],[始'[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const dataQualifierEscapedOpenBracketTypedEscapeOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[#Data],[始'[数]]"))) throw new Error("escaped open bracket typed #Data structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#Data],[始'[数]]", null, { timeout: 5000 });
    const dataQualifierEscapedOpenBracketTypedEscapeAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#This Row],[担");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const thisRowQualifierEscapedAtOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[#This Row],['@担当]]"))) throw new Error("escaped @ #This Row structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#This Row],['@担当]]", null, { timeout: 5000 });
    const thisRowQualifierEscapedAtAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#Data],[税込");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const dataQualifierDoubleWrappedCommaOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[#Data],[金額,税込]]"))) throw new Error("double-wrapped comma #Data structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#Data],[金額,税込]]", null, { timeout: 5000 });
    const dataQualifierDoubleWrappedCommaAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#Data],[#A");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const dataQualifierEscapedAllNameOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[#Data],['#All]]"))) throw new Error("escaped #All #Data structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#Data],['#All]]", null, { timeout: 5000 });
    const dataQualifierEscapedAllNameAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#Data],['#A");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const dataQualifierEscapedAllNameTypedEscapeOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[#Data],['#All]]"))) throw new Error("escaped #All typed #Data structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#Data],['#All]]", null, { timeout: 5000 });
    const dataQualifierEscapedAllNameTypedEscapeAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#Headers],[#H");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const headersQualifierEscapedHeadersNameOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[#Headers],['#Headers]]"))) throw new Error("escaped #Headers #Headers structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#Headers],['#Headers]]", null, { timeout: 5000 });
    const headersQualifierEscapedHeadersNameAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#Totals],[#T");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const totalsQualifierEscapedTotalsNameOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[#Totals],['#Totals]]"))) throw new Error("escaped #Totals #Totals structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#Totals],['#Totals]]", null, { timeout: 5000 });
    const totalsQualifierEscapedTotalsNameAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[数量]:[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const columnRangeStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("Sales Amount"))) throw new Error("Sales Amount column-range structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[数量]:[Sales Amount]]", null, { timeout: 5000 });
    const columnRangeStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[税'#率]:[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const escapedRangeStructuredOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("端]数"))) throw new Error("escaped column-range structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[税'#率]:[端']数]]", null, { timeout: 5000 });
    const escapedRangeStructuredAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[金額,税込]:[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const doubleWrappedRangeStartOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("SalesTable[[金額,税込]:[A:B]]"))) throw new Error("double-wrapped range-start structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[金額,税込]:[A:B]]", null, { timeout: 5000 });
    const doubleWrappedRangeStartAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#Data],[税'#率]:[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const dataQualifierEscapedRangeOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("端]数"))) throw new Error("escaped #Data column-range structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#Data],[税'#率]:[端']数]]", null, { timeout: 5000 });
    const dataQualifierEscapedRangeAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#All],[O''Brien]:[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const allQualifierEscapedRangeOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("端]数"))) throw new Error("escaped #All column-range structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#All],[O''Brien]:[端']数]]", null, { timeout: 5000 });
    const allQualifierEscapedRangeAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#Headers],[#Data],[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const headersDataQualifierOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("Sales Amount"))) throw new Error("Sales Amount #Headers #Data structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#Headers],[#Data],[Sales Amount]]", null, { timeout: 5000 });
    const headersDataQualifierAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=SalesTable[[#Headers],[#Data],[税'#率]:[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const headersDataEscapedRangeOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("端]数"))) throw new Error("escaped #Headers #Data range structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#Headers],[#Data],[税'#率]:[端']数]]", null, { timeout: 5000 });
    const headersDataEscapedRangeAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=[[#Headers],[#Data],[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const currentTableHeadersDataQualifierOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("Sales Amount"))) throw new Error("Sales Amount current-table #Headers #Data structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=[[#Headers],[#Data],[Sales Amount]]", null, { timeout: 5000 });
    const currentTableHeadersDataQualifierAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=[[#Data],[");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const currentTableDataQualifierOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemByName("Sales Amount"))) throw new Error("Sales Amount current-table #Data structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=[[#Data],[Sales Amount]]", null, { timeout: 5000 });
    const currentTableDataQualifierAccepted = await page.locator("#formulaInput").inputValue();

    await resetFormulaInput("=[[#Data],[#A");
    await page.waitForFunction(() => document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length > 0, null, { timeout: 5000 });
    const currentTableDataQualifierEscapedAllNameOpen = await popupSnapshot();
    if (!(await clickAutocompleteItemBySyntax("[[#Data],['#All]]"))) throw new Error("escaped #All current-table #Data structured candidate not found");
    await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=[[#Data],['#All]]", null, { timeout: 5000 });
    const currentTableDataQualifierEscapedAllNameAccepted = await page.locator("#formulaInput").inputValue();

    await page.keyboard.press("Escape");
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(10, 4, "=[@[Sales Amount]]+1");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "21" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "31",
      null,
      { timeout: 5000 },
    );
    const shorthandStructuredReferenceCalculation = await page.evaluate(() => ({
      shorthand: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      propagated: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      shorthandEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(11, 4, "=SalesTable[@[Sales Amount]]+2");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "22" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "32",
      null,
      { timeout: 5000 },
    );
    const namedStructuredReferenceCalculation = await page.evaluate(() => ({
      propagated: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      named: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      namedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(10, 4, "=[@[税'#率]]+1");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "6" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "9",
      null,
      { timeout: 5000 },
    );
    const currentRowEscapedHashCalculation = await page.evaluate(() => ({
      shorthand: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      propagated: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      shorthandEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(11, 4, "=SalesTable[@[端']数]]+2");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "5" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "6",
      null,
      { timeout: 5000 },
    );
    const namedCurrentRowEscapedBracketCalculation = await page.evaluate(() => ({
      propagated: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      named: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      namedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(10, 4, "=SUM([@[税'#率]:[端']数]])");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "15" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "21",
      null,
      { timeout: 5000 },
    );
    const currentRowEscapedRangeCalculation = await page.evaluate(() => ({
      shorthand: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      propagated: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      shorthandEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(11, 4, "=SUM(SalesTable[@[税'#率]:[端']数]])");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "15" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "21",
      null,
      { timeout: 5000 },
    );
    const namedCurrentRowEscapedRangeCalculation = await page.evaluate(() => ({
      propagated: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      named: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      namedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(10, 4, "=[@['@担当]]+10");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "12" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "13",
      null,
      { timeout: 5000 },
    );
    const currentRowEscapedAtCalculation = await page.evaluate(() => ({
      shorthand: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      propagated: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      shorthandEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(10, 4, "=[@[金額,税込]]+1");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "14" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "18",
      null,
      { timeout: 5000 },
    );
    const currentRowDoubleWrappedCommaCalculation = await page.evaluate(() => ({
      shorthand: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      propagated: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      shorthandEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(11, 4, "=SalesTable[@[比率.%]]+30");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "31" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "32",
      null,
      { timeout: 5000 },
    );
    const namedCurrentRowDoubleWrappedDotPercentCalculation = await page.evaluate(() => ({
      propagated: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      named: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      namedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(10, 4, "=[@[費用?]]+40");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "42" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "45",
      null,
      { timeout: 5000 },
    );
    const currentRowDoubleWrappedQuestionCalculation = await page.evaluate(() => ({
      shorthand: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      propagated: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      shorthandEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(11, 4, "=SalesTable[@[値(税)]]+50");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "67" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "69",
      null,
      { timeout: 5000 },
    );
    const namedCurrentRowDoubleWrappedParenCalculation = await page.evaluate(() => ({
      propagated: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      named: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      namedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(10, 4, "=[@[担当,'@税込]]+70");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "93" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "99",
      null,
      { timeout: 5000 },
    );
    const currentRowDoubleWrappedEscapedAtCommaCalculation = await page.evaluate(() => ({
      shorthand: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      propagated: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      shorthandEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(10, 4, "=[@['#All]]+80");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "111" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "123",
      null,
      { timeout: 5000 },
    );
    const currentRowEscapedAllNameCalculation = await page.evaluate(() => ({
      shorthand: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      propagated: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      shorthandEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(11, 4, "=SalesTable[@['#Totals]]+90");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "131" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "143",
      null,
      { timeout: 5000 },
    );
    const namedCurrentRowEscapedTotalsNameCalculation = await page.evaluate(() => ({
      propagated: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      named: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      namedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(10, 4, "=SUM(SalesTable[[#This Row],[税'#率]:[端']数]])");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "15" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "21",
      null,
      { timeout: 5000 },
    );
    const thisRowQualifierEscapedRangeCalculation = await page.evaluate(() => ({
      shorthand: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      propagated: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      shorthandEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(11, 4, "=SalesTable[[#This Row],['@担当]]+20");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "22" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "23",
      null,
      { timeout: 5000 },
    );
    const thisRowQualifierEscapedAtCalculation = await page.evaluate(() => ({
      propagated: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      named: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      namedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(10, 4, "=SUM(SalesTable[[#This Row],[金額,税込]:[A:B]])");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "35" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "41",
      null,
      { timeout: 5000 },
    );
    const thisRowQualifierDoubleWrappedRangeCalculation = await page.evaluate(() => ({
      shorthand: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      propagated: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      shorthandEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(10, 4, "=COUNTA([[#Headers],[#Data],[Sales Amount]])");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "3" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "3",
      null,
      { timeout: 5000 },
    );
    const currentTableHeadersDataCalculation = await page.evaluate(() => ({
      shorthand: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      propagated: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      shorthandEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(10, 4, "=SUM([[#Data],[Sales Amount]])");
    });
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 10, 4) === "50" &&
        window.__webSheetDevCellDisplay?.("CHT", 11, 4) === "50",
      null,
      { timeout: 5000 },
    );
    const currentTableDataCalculation = await page.evaluate(() => ({
      shorthand: window.__webSheetDevCellDisplay?.("CHT", 10, 4) || "",
      propagated: window.__webSheetDevCellDisplay?.("CHT", 11, 4) || "",
      shorthandEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 10, 4),
      propagatedEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 11, 4),
    }));
    await page.evaluate(() => {
      window.__webSheetDevSetCell?.(13, 5, "=SUM(SalesTable[[#Data],[Sales Amount]])");
      window.__webSheetDevSetCell?.(14, 5, "=COUNTA(SalesTable[[#Headers],[Sales Amount]])");
      window.__webSheetDevSetCell?.(15, 5, "=SUM(SalesTable[[数量]:[Sales Amount]])");
      window.__webSheetDevSetCell?.(16, 5, "=SUM(SalesTable[[#Totals],[Sales Amount]])");
      window.__webSheetDevSetCell?.(17, 5, "=SUM(SalesTable[[#All],[Sales Amount]])");
      window.__webSheetDevSetCell?.(18, 5, "=SUM(SalesTable[税'#率])");
      window.__webSheetDevSetCell?.(19, 5, "=SUM(SalesTable[O''Brien])");
      window.__webSheetDevSetCell?.(20, 5, "=SUM(SalesTable[端']数])");
      window.__webSheetDevSetCell?.(21, 5, "=SUM(SalesTable[[税'#率]:[端']数]])");
      window.__webSheetDevSetCell?.(22, 5, "=SUM(SalesTable[[#Totals],[税'#率]])");
      window.__webSheetDevSetCell?.(23, 5, "=SUM(SalesTable[[#All],[端']数]])");
      window.__webSheetDevSetCell?.(24, 5, "=SUM(SalesTable[[#Data],[税'#率]:[端']数]])");
      window.__webSheetDevSetCell?.(25, 5, "=SUM(SalesTable[[#Totals],[税'#率]:[端']数]])");
      window.__webSheetDevSetCell?.(26, 5, "=SUM(SalesTable[[#All],[税'#率]:[端']数]])");
      window.__webSheetDevSetCell?.(27, 5, "=COUNTA(SalesTable[[#Headers],[税'#率]:[端']数]])");
      window.__webSheetDevSetCell?.(28, 5, "=COUNTA(SalesTable[[#Headers],[#Data],[Sales Amount]])");
      window.__webSheetDevSetCell?.(29, 5, "=SUM(SalesTable[[#Data],[#Totals],[Sales Amount]])");
      window.__webSheetDevSetCell?.(30, 5, "=SUM(SalesTable[[#Headers],[#Data],[#Totals],[Sales Amount]])");
      window.__webSheetDevSetCell?.(31, 5, "=SUM(SalesTable[[#Headers],[#Data],[税'#率]:[端']数]])");
      window.__webSheetDevSetCell?.(32, 5, "=SUM(SalesTable['#Data])");
      window.__webSheetDevSetCell?.(33, 5, "=SUM(SalesTable[[#Totals],['#Data]])");
      window.__webSheetDevSetCell?.(34, 5, "=SUM(SalesTable['@担当])");
      window.__webSheetDevSetCell?.(35, 5, "=SUM(SalesTable['#This Row])");
      window.__webSheetDevSetCell?.(36, 5, "=SUM(SalesTable[[#All],['#Data]])");
      window.__webSheetDevSetCell?.(37, 5, "=SUM(SalesTable[[金額,税込]])");
      window.__webSheetDevSetCell?.(38, 5, "=SUM(SalesTable[[#Totals],[金額,税込]])");
      window.__webSheetDevSetCell?.(39, 5, "=SUM(SalesTable[[#All],[金額,税込]])");
      window.__webSheetDevSetCell?.(40, 5, "=SUM(SalesTable[[比率.%]])");
      window.__webSheetDevSetCell?.(41, 5, "=SUM(SalesTable[[A:B]])");
      window.__webSheetDevSetCell?.(42, 5, "=SUM(SalesTable[[金額,税込]:[A:B]])");
      window.__webSheetDevSetCell?.(43, 5, "=SUM(SalesTable[[費用?]])");
      window.__webSheetDevSetCell?.(44, 5, "=SUM(SalesTable[[売上$]])");
      window.__webSheetDevSetCell?.(45, 5, "=SUM(SalesTable[[A+B]])");
      window.__webSheetDevSetCell?.(46, 5, "=SUM(SalesTable[[値(税)]])");
      window.__webSheetDevSetCell?.(47, 5, "=SUM(SalesTable[[費用?]:[値(税)]])");
      window.__webSheetDevSetCell?.(48, 5, "=SUM(SalesTable[[担当,'@税込]])");
      window.__webSheetDevSetCell?.(49, 5, "=SUM(SalesTable[[#All],[担当,'@税込]])");
      window.__webSheetDevSetCell?.(50, 5, "=SUM(SalesTable['#All])");
      window.__webSheetDevSetCell?.(51, 5, "=COUNTA(SalesTable[[#Headers],['#Headers]])");
      window.__webSheetDevSetCell?.(52, 5, "=SUM(SalesTable[[#Totals],['#Totals]])");
      window.__webSheetDevSetCell?.(53, 5, "=SUM(SalesTable[[#All],['#Totals]])");
      window.__webSheetDevSetCell?.(54, 5, "=SUM(SalesTable[始'[数])");
      window.__webSheetDevSetCell?.(55, 5, "=SUM(SalesTable[[#Totals],[始'[数]])");
      window.__webSheetDevSetCell?.(56, 5, "=SUM(SalesTable[[#All],[始'[数]])");
    });
    await page.evaluate((sheetPrefix) => {
      window.__webSheetDevSetCell?.(57, 5, `=SUM(${sheetPrefix}!B2:C3)`);
    }, apostropheSheetPrefix);
    await page.waitForFunction(
      () =>
        window.__webSheetDevCellDisplay?.("CHT", 13, 5) === "50" &&
        window.__webSheetDevCellDisplay?.("CHT", 14, 5) === "1" &&
        window.__webSheetDevCellDisplay?.("CHT", 15, 5) === "60" &&
        window.__webSheetDevCellDisplay?.("CHT", 16, 5) === "50" &&
        window.__webSheetDevCellDisplay?.("CHT", 17, 5) === "100" &&
        window.__webSheetDevCellDisplay?.("CHT", 18, 5) === "13" &&
        window.__webSheetDevCellDisplay?.("CHT", 19, 5) === "16" &&
        window.__webSheetDevCellDisplay?.("CHT", 20, 5) === "7" &&
        window.__webSheetDevCellDisplay?.("CHT", 21, 5) === "36" &&
        window.__webSheetDevCellDisplay?.("CHT", 22, 5) === "13" &&
        window.__webSheetDevCellDisplay?.("CHT", 23, 5) === "14" &&
        window.__webSheetDevCellDisplay?.("CHT", 24, 5) === "36" &&
        window.__webSheetDevCellDisplay?.("CHT", 25, 5) === "36" &&
        window.__webSheetDevCellDisplay?.("CHT", 26, 5) === "72" &&
        window.__webSheetDevCellDisplay?.("CHT", 27, 5) === "3" &&
        window.__webSheetDevCellDisplay?.("CHT", 28, 5) === "3" &&
        window.__webSheetDevCellDisplay?.("CHT", 29, 5) === "100" &&
        window.__webSheetDevCellDisplay?.("CHT", 30, 5) === "100" &&
        window.__webSheetDevCellDisplay?.("CHT", 31, 5) === "36" &&
        window.__webSheetDevCellDisplay?.("CHT", 32, 5) === "23" &&
        window.__webSheetDevCellDisplay?.("CHT", 33, 5) === "99" &&
        window.__webSheetDevCellDisplay?.("CHT", 34, 5) === "5" &&
        window.__webSheetDevCellDisplay?.("CHT", 35, 5) === "10" &&
        window.__webSheetDevCellDisplay?.("CHT", 36, 5) === "122" &&
        window.__webSheetDevCellDisplay?.("CHT", 37, 5) === "30" &&
        window.__webSheetDevCellDisplay?.("CHT", 38, 5) === "301" &&
        window.__webSheetDevCellDisplay?.("CHT", 39, 5) === "331" &&
        window.__webSheetDevCellDisplay?.("CHT", 40, 5) === "3" &&
        window.__webSheetDevCellDisplay?.("CHT", 41, 5) === "43" &&
        window.__webSheetDevCellDisplay?.("CHT", 42, 5) === "76" &&
        window.__webSheetDevCellDisplay?.("CHT", 43, 5) === "7" &&
        window.__webSheetDevCellDisplay?.("CHT", 44, 5) === "10" &&
        window.__webSheetDevCellDisplay?.("CHT", 45, 5) === "24" &&
        window.__webSheetDevCellDisplay?.("CHT", 46, 5) === "36" &&
        window.__webSheetDevCellDisplay?.("CHT", 47, 5) === "77" &&
        window.__webSheetDevCellDisplay?.("CHT", 48, 5) === "52" &&
        window.__webSheetDevCellDisplay?.("CHT", 49, 5) === "172" &&
        window.__webSheetDevCellDisplay?.("CHT", 50, 5) === "74" &&
        window.__webSheetDevCellDisplay?.("CHT", 51, 5) === "1" &&
        window.__webSheetDevCellDisplay?.("CHT", 52, 5) === "410" &&
        window.__webSheetDevCellDisplay?.("CHT", 53, 5) === "504" &&
        window.__webSheetDevCellDisplay?.("CHT", 54, 5) === "116" &&
        window.__webSheetDevCellDisplay?.("CHT", 55, 5) === "590" &&
        window.__webSheetDevCellDisplay?.("CHT", 56, 5) === "706" &&
        window.__webSheetDevCellDisplay?.("CHT", 57, 5) === "579",
      null,
      { timeout: 5000 },
    );
    const nestedStructuredReferenceCalculation = await page.evaluate(() => ({
      data: window.__webSheetDevCellDisplay?.("CHT", 13, 5) || "",
      headers: window.__webSheetDevCellDisplay?.("CHT", 14, 5) || "",
      columnRange: window.__webSheetDevCellDisplay?.("CHT", 15, 5) || "",
      totals: window.__webSheetDevCellDisplay?.("CHT", 16, 5) || "",
      all: window.__webSheetDevCellDisplay?.("CHT", 17, 5) || "",
      hash: window.__webSheetDevCellDisplay?.("CHT", 18, 5) || "",
      apostrophe: window.__webSheetDevCellDisplay?.("CHT", 19, 5) || "",
      bracket: window.__webSheetDevCellDisplay?.("CHT", 20, 5) || "",
      escapedRange: window.__webSheetDevCellDisplay?.("CHT", 21, 5) || "",
      escapedTotals: window.__webSheetDevCellDisplay?.("CHT", 22, 5) || "",
      escapedAll: window.__webSheetDevCellDisplay?.("CHT", 23, 5) || "",
      rowQualifiedEscapedRangeData: window.__webSheetDevCellDisplay?.("CHT", 24, 5) || "",
      rowQualifiedEscapedRangeTotals: window.__webSheetDevCellDisplay?.("CHT", 25, 5) || "",
      rowQualifiedEscapedRangeAll: window.__webSheetDevCellDisplay?.("CHT", 26, 5) || "",
      rowQualifiedEscapedRangeHeaders: window.__webSheetDevCellDisplay?.("CHT", 27, 5) || "",
      combinedHeadersData: window.__webSheetDevCellDisplay?.("CHT", 28, 5) || "",
      combinedDataTotals: window.__webSheetDevCellDisplay?.("CHT", 29, 5) || "",
      combinedHeadersDataTotals: window.__webSheetDevCellDisplay?.("CHT", 30, 5) || "",
      combinedHeadersDataEscapedRange: window.__webSheetDevCellDisplay?.("CHT", 31, 5) || "",
      specialNameColumn: window.__webSheetDevCellDisplay?.("CHT", 32, 5) || "",
      specialNameColumnTotals: window.__webSheetDevCellDisplay?.("CHT", 33, 5) || "",
      atNameColumn: window.__webSheetDevCellDisplay?.("CHT", 34, 5) || "",
      thisRowNameColumn: window.__webSheetDevCellDisplay?.("CHT", 35, 5) || "",
      specialNameColumnAll: window.__webSheetDevCellDisplay?.("CHT", 36, 5) || "",
      commaColumn: window.__webSheetDevCellDisplay?.("CHT", 37, 5) || "",
      commaColumnTotals: window.__webSheetDevCellDisplay?.("CHT", 38, 5) || "",
      commaColumnAll: window.__webSheetDevCellDisplay?.("CHT", 39, 5) || "",
      dotPercentColumn: window.__webSheetDevCellDisplay?.("CHT", 40, 5) || "",
      colonColumn: window.__webSheetDevCellDisplay?.("CHT", 41, 5) || "",
      doubleWrappedRange: window.__webSheetDevCellDisplay?.("CHT", 42, 5) || "",
      questionColumn: window.__webSheetDevCellDisplay?.("CHT", 43, 5) || "",
      dollarColumn: window.__webSheetDevCellDisplay?.("CHT", 44, 5) || "",
      plusColumn: window.__webSheetDevCellDisplay?.("CHT", 45, 5) || "",
      parenColumn: window.__webSheetDevCellDisplay?.("CHT", 46, 5) || "",
      punctuationRange: window.__webSheetDevCellDisplay?.("CHT", 47, 5) || "",
      escapedAtCommaColumn: window.__webSheetDevCellDisplay?.("CHT", 48, 5) || "",
      escapedAtCommaColumnAll: window.__webSheetDevCellDisplay?.("CHT", 49, 5) || "",
      allNameColumn: window.__webSheetDevCellDisplay?.("CHT", 50, 5) || "",
      headersNameColumnHeader: window.__webSheetDevCellDisplay?.("CHT", 51, 5) || "",
      totalsNameColumnTotals: window.__webSheetDevCellDisplay?.("CHT", 52, 5) || "",
      totalsNameColumnAll: window.__webSheetDevCellDisplay?.("CHT", 53, 5) || "",
      openBracketColumn: window.__webSheetDevCellDisplay?.("CHT", 54, 5) || "",
      openBracketColumnTotals: window.__webSheetDevCellDisplay?.("CHT", 55, 5) || "",
      openBracketColumnAll: window.__webSheetDevCellDisplay?.("CHT", 56, 5) || "",
      apostropheSheetRange: window.__webSheetDevCellDisplay?.("CHT", 57, 5) || "",
      dataEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 13, 5),
      headersEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 14, 5),
      columnRangeEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 15, 5),
      totalsEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 16, 5),
      allEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 17, 5),
      hashEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 18, 5),
      apostropheEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 19, 5),
      bracketEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 20, 5),
      escapedRangeEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 21, 5),
      escapedTotalsEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 22, 5),
      escapedAllEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 23, 5),
      rowQualifiedEscapedRangeDataEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 24, 5),
      rowQualifiedEscapedRangeTotalsEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 25, 5),
      rowQualifiedEscapedRangeAllEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 26, 5),
      rowQualifiedEscapedRangeHeadersEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 27, 5),
      combinedHeadersDataEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 28, 5),
      combinedDataTotalsEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 29, 5),
      combinedHeadersDataTotalsEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 30, 5),
      combinedHeadersDataEscapedRangeEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 31, 5),
      specialNameColumnEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 32, 5),
      specialNameColumnTotalsEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 33, 5),
      atNameColumnEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 34, 5),
      thisRowNameColumnEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 35, 5),
      specialNameColumnAllEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 36, 5),
      commaColumnEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 37, 5),
      commaColumnTotalsEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 38, 5),
      commaColumnAllEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 39, 5),
      dotPercentColumnEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 40, 5),
      colonColumnEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 41, 5),
      doubleWrappedRangeEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 42, 5),
      questionColumnEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 43, 5),
      dollarColumnEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 44, 5),
      plusColumnEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 45, 5),
      parenColumnEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 46, 5),
      punctuationRangeEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 47, 5),
      escapedAtCommaColumnEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 48, 5),
      escapedAtCommaColumnAllEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 49, 5),
      allNameColumnEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 50, 5),
      headersNameColumnHeaderEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 51, 5),
      totalsNameColumnTotalsEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 52, 5),
      totalsNameColumnAllEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 53, 5),
      openBracketColumnEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 54, 5),
      openBracketColumnTotalsEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 55, 5),
      openBracketColumnAllEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 56, 5),
      apostropheSheetRangeEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 57, 5),
    }));

    const formulaStandaloneViewSetup = await page.evaluate(() => {
      return window.__webSheetDevSetSheetViewOptions?.("CHT", {
        viewMode: "page-layout",
        zoomScale: 125,
        showGridLines: false,
        showHeadings: false,
        showZeros: false,
        showObjects: false,
        showOutlineSymbols: false,
        showWhiteSpace: false,
        rightToLeft: true,
        freezePanes: { rows: 2, columns: 1, activePane: "bottomRight" },
        splitPane: null,
      }) || null;
    });
    const formulaStandaloneDownloadPromise = page.waitForEvent("download");
    await page.keyboard.press(saveShortcut);
    const formulaStandaloneDownload = await formulaStandaloneDownloadPromise;
    const formulaStandaloneDownloadPath = await formulaStandaloneDownload.path();
    const formulaStandaloneHtmlPath = path.join(path.dirname(formulaStandaloneDownloadPath), formulaStandaloneDownload.suggestedFilename());
    await formulaStandaloneDownload.saveAs(formulaStandaloneHtmlPath);
    const formulaStandaloneHtml = await readFile(formulaStandaloneHtmlPath, "utf8");
    const formulaStandaloneModel = exportedStandaloneModelFromHtml(formulaStandaloneHtml);
    const formulaStandaloneModelSheet = formulaStandaloneModel.sheets?.find((sheet) => sheet.name === "CHT");
    const formulaStandaloneModelApostropheSheet = formulaStandaloneModel.sheets?.find((sheet) => sheet.name === apostropheSheetName);
    const formulaStandaloneModelCell = formulaStandaloneModelSheet?.cells?.["57:5"] || null;
    const formulaStandaloneExport = {
      suggestedFilename: formulaStandaloneDownload.suggestedFilename(),
      isHtml: looksLikeStandaloneWorkbookHtml(formulaStandaloneHtml),
      sourceName: formulaStandaloneModel.sourceName || "",
      saveType: formulaStandaloneModel.saveType || "",
      sheetNames: (formulaStandaloneModel.sheets || []).map((sheet) => sheet.name),
      formulaCell: formulaStandaloneModelCell,
      view: formulaStandaloneModelSheet ? {
        viewMode: formulaStandaloneModelSheet.viewMode || "normal",
        zoomScale: formulaStandaloneModelSheet.zoomScale || 0,
        showGridLines: formulaStandaloneModelSheet.showGridLines,
        showHeadings: formulaStandaloneModelSheet.showHeadings,
        showZeros: formulaStandaloneModelSheet.showZeros,
        showObjects: formulaStandaloneModelSheet.showObjects,
        showOutlineSymbols: formulaStandaloneModelSheet.showOutlineSymbols,
        showWhiteSpace: formulaStandaloneModelSheet.showWhiteSpace,
        rightToLeft: formulaStandaloneModelSheet.rightToLeft,
        freezePanes: formulaStandaloneModelSheet.freezePanes || null,
      } : null,
      apostropheSheetValues: {
        b2: formulaStandaloneModelApostropheSheet?.cells?.["2:2"]?.raw ?? formulaStandaloneModelApostropheSheet?.cells?.["2:2"]?.display ?? null,
        c3: formulaStandaloneModelApostropheSheet?.cells?.["3:3"]?.raw ?? formulaStandaloneModelApostropheSheet?.cells?.["3:3"]?.display ?? null,
      },
    };
    const formulaStandalonePage = await context.newPage();
    const formulaStandaloneMessages = [];
    const formulaStandaloneBlockedRequests = [];
    formulaStandalonePage.on("console", (message) => {
      if (message.type() === "error" && !/Failed to load resource: the server responded with a status of 404/i.test(message.text())) {
        formulaStandaloneMessages.push(`${message.type()}: ${message.text()}`);
      }
    });
    formulaStandalonePage.on("pageerror", (error) => {
      formulaStandaloneMessages.push(`pageerror: ${error.stack || error.message}`);
    });
    await formulaStandalonePage.route("**/*", async (route) => {
      const url = route.request().url();
      if (/^https?:/i.test(url)) {
        formulaStandaloneBlockedRequests.push(url);
        await route.abort();
        return;
      }
      await route.continue();
    });
    await formulaStandalonePage.goto(pathToFileURL(formulaStandaloneHtmlPath).href, { waitUntil: "load" });
    await formulaStandalonePage.waitForFunction(
      () => window.__WEBSHEET_STANDALONE_MODEL__?.sheets?.length && document.querySelector("#sheetTabs .sheet-tab")?.textContent?.trim() === "CHT",
      null,
      { timeout: 30000 },
    );
    await formulaStandalonePage.waitForFunction(
      () =>
        window.__webSheetStandaloneCell?.("CHT", 57, 5)?.display === "579" ||
        window.__webSheetDevCellDisplay?.("CHT", 57, 5) === "579",
      null,
      { timeout: 10000 },
    );
    const formulaStandaloneReload = await formulaStandalonePage.evaluate(() => {
      const cell = window.__webSheetStandaloneCell?.("CHT", 57, 5) || null;
      const diagnostics = window.__webSheetStandaloneDiagnostics?.() || null;
      return {
        diagnostics,
        cell,
        devCell: window.__webSheetDevCellModel?.("CHT", 57, 5) || null,
        devDisplay: window.__webSheetDevCellDisplay?.("CHT", 57, 5) || "",
        devEngine: window.__webSheetDevFormulaEngineCell?.("CHT", 57, 5) || null,
        status: document.querySelector("#statusText")?.textContent || "",
        firstTab: document.querySelector("#sheetTabs .sheet-tab")?.textContent?.trim() || "",
        sheetCount: window.__WEBSHEET_STANDALONE_MODEL__?.sheets?.length || 0,
        runtimeSource: window.__WEBSHEET_RUNTIME_SOURCE__ || "",
      };
    });
    formulaStandaloneReload.messages = formulaStandaloneMessages;
    formulaStandaloneReload.blockedRequests = formulaStandaloneBlockedRequests;
    await formulaStandalonePage.close();

    const result = {
      definitions,
      sumOpen,
      sumAfterArrow,
      sumAccepted,
      parenAccepted,
      plainParen,
      repeatedArgument,
      sheetNames,
      qualifiedSheetName,
      qualifiedReferenceF4,
      apostropheSheetName,
      apostropheReferenceF4,
      sheetReferenceOpen,
      sheetReferenceAccepted,
      quotedSheetReferenceOpen,
      quotedSheetReferenceAccepted,
      apostropheSheetReferenceOpen,
      apostropheSheetReferenceAccepted,
      quotedApostropheSheetReferenceOpen,
      quotedApostropheSheetReferenceAccepted,
      lognormOpen,
      deltaOpen,
      xlookupOpen,
      xlookupDialogInitial,
      xlookupDialogReturnArray,
      xlookupDialogReferenceCollapsed,
      xlookupDialogReferenceF4,
      filterOpen,
      pmtOpen,
      definedNameOpen,
      definedNameAccepted,
      tableNameOpen,
      tableNameAccepted,
      tableStructuredOpen,
      tableStructuredAccepted,
      escapedHashColumnStructuredOpen,
      escapedHashColumnStructuredAccepted,
      escapedApostropheColumnStructuredOpen,
      escapedApostropheColumnStructuredAccepted,
      escapedApostropheTypedEscapeColumnStructuredOpen,
      escapedApostropheTypedEscapeColumnStructuredAccepted,
      escapedBracketColumnStructuredOpen,
      escapedBracketColumnStructuredAccepted,
      escapedOpenBracketColumnStructuredOpen,
      escapedOpenBracketColumnStructuredAccepted,
      escapedOpenBracketTypedEscapeColumnStructuredOpen,
      escapedOpenBracketTypedEscapeColumnStructuredAccepted,
      escapedSpecialNameColumnStructuredOpen,
      escapedSpecialNameColumnStructuredAccepted,
      escapedAtColumnStructuredOpen,
      escapedAtColumnStructuredAccepted,
      escapedThisRowNameColumnStructuredOpen,
      escapedThisRowNameColumnStructuredAccepted,
      doubleWrappedCommaColumnStructuredOpen,
      doubleWrappedCommaColumnStructuredAccepted,
      doubleWrappedDotPercentColumnStructuredOpen,
      doubleWrappedDotPercentColumnStructuredAccepted,
      doubleWrappedColonColumnStructuredOpen,
      doubleWrappedColonColumnStructuredAccepted,
      doubleWrappedQuestionColumnStructuredOpen,
      doubleWrappedQuestionColumnStructuredAccepted,
      doubleWrappedDollarColumnStructuredOpen,
      doubleWrappedDollarColumnStructuredAccepted,
      doubleWrappedPlusColumnStructuredOpen,
      doubleWrappedPlusColumnStructuredAccepted,
      doubleWrappedParenColumnStructuredOpen,
      doubleWrappedParenColumnStructuredAccepted,
      doubleWrappedEscapedAtCommaColumnStructuredOpen,
      doubleWrappedEscapedAtCommaColumnStructuredAccepted,
      escapedAllNameColumnStructuredOpen,
      escapedAllNameColumnStructuredAccepted,
      escapedAllNameTypedEscapeColumnStructuredOpen,
      escapedAllNameTypedEscapeColumnStructuredAccepted,
      escapedHeadersNameColumnStructuredOpen,
      escapedHeadersNameColumnStructuredAccepted,
      escapedTotalsNameColumnStructuredOpen,
      escapedTotalsNameColumnStructuredAccepted,
      nestedSpecialStructuredOpen,
      nestedSpecialStructuredAccepted,
      nestedThisRowStructuredOpen,
      nestedThisRowStructuredAccepted,
      currentRowStructuredOpen,
      currentRowStructuredAccepted,
      namedCurrentRowStructuredOpen,
      namedCurrentRowStructuredAccepted,
      namedCurrentRowEscapedHashOpen,
      namedCurrentRowEscapedHashAccepted,
      currentRowEscapedApostropheOpen,
      currentRowEscapedApostropheAccepted,
      currentRowEscapedBracketOpen,
      currentRowEscapedBracketAccepted,
      currentRowEscapedOpenBracketTypedEscapeOpen,
      currentRowEscapedOpenBracketTypedEscapeAccepted,
      currentRowEscapedAtOpen,
      currentRowEscapedAtAccepted,
      currentRowEscapedAtTypedEscapeOpen,
      currentRowEscapedAtTypedEscapeAccepted,
      namedCurrentRowEscapedAtOpen,
      namedCurrentRowEscapedAtAccepted,
      currentRowDoubleWrappedCommaOpen,
      currentRowDoubleWrappedCommaAccepted,
      namedCurrentRowDoubleWrappedDotPercentOpen,
      namedCurrentRowDoubleWrappedDotPercentAccepted,
      currentRowDoubleWrappedQuestionOpen,
      currentRowDoubleWrappedQuestionAccepted,
      namedCurrentRowDoubleWrappedParenOpen,
      namedCurrentRowDoubleWrappedParenAccepted,
      currentRowDoubleWrappedEscapedAtCommaOpen,
      currentRowDoubleWrappedEscapedAtCommaAccepted,
      currentRowEscapedAllNameOpen,
      currentRowEscapedAllNameAccepted,
      namedCurrentRowEscapedTotalsNameOpen,
      namedCurrentRowEscapedTotalsNameAccepted,
      namedCurrentRowEscapedRangeOpen,
      namedCurrentRowEscapedRangeAccepted,
      currentRowEscapedRangeOpen,
      currentRowEscapedRangeAccepted,
      dataQualifierStructuredOpen,
      dataQualifierStructuredAccepted,
      headersQualifierStructuredOpen,
      headersQualifierStructuredAccepted,
      totalsQualifierStructuredOpen,
      totalsQualifierStructuredAccepted,
      dataQualifierEscapedHashOpen,
      dataQualifierEscapedHashAccepted,
      totalsQualifierEscapedBracketOpen,
      totalsQualifierEscapedBracketAccepted,
      dataQualifierEscapedOpenBracketTypedEscapeOpen,
      dataQualifierEscapedOpenBracketTypedEscapeAccepted,
      thisRowQualifierEscapedAtOpen,
      thisRowQualifierEscapedAtAccepted,
      dataQualifierDoubleWrappedCommaOpen,
      dataQualifierDoubleWrappedCommaAccepted,
      dataQualifierEscapedAllNameOpen,
      dataQualifierEscapedAllNameAccepted,
      dataQualifierEscapedAllNameTypedEscapeOpen,
      dataQualifierEscapedAllNameTypedEscapeAccepted,
      headersQualifierEscapedHeadersNameOpen,
      headersQualifierEscapedHeadersNameAccepted,
      totalsQualifierEscapedTotalsNameOpen,
      totalsQualifierEscapedTotalsNameAccepted,
      columnRangeStructuredOpen,
      columnRangeStructuredAccepted,
      escapedRangeStructuredOpen,
      escapedRangeStructuredAccepted,
      doubleWrappedRangeStartOpen,
      doubleWrappedRangeStartAccepted,
      dataQualifierEscapedRangeOpen,
      dataQualifierEscapedRangeAccepted,
      allQualifierEscapedRangeOpen,
      allQualifierEscapedRangeAccepted,
      headersDataQualifierOpen,
      headersDataQualifierAccepted,
      headersDataEscapedRangeOpen,
      headersDataEscapedRangeAccepted,
      currentTableHeadersDataQualifierOpen,
      currentTableHeadersDataQualifierAccepted,
      currentTableDataQualifierOpen,
      currentTableDataQualifierAccepted,
      currentTableDataQualifierEscapedAllNameOpen,
      currentTableDataQualifierEscapedAllNameAccepted,
      shorthandStructuredReferenceCalculation,
      namedStructuredReferenceCalculation,
      currentRowEscapedHashCalculation,
      namedCurrentRowEscapedBracketCalculation,
      currentRowEscapedRangeCalculation,
      namedCurrentRowEscapedRangeCalculation,
      currentRowEscapedAtCalculation,
      currentRowDoubleWrappedCommaCalculation,
      namedCurrentRowDoubleWrappedDotPercentCalculation,
      currentRowDoubleWrappedQuestionCalculation,
      namedCurrentRowDoubleWrappedParenCalculation,
      currentRowDoubleWrappedEscapedAtCommaCalculation,
      currentRowEscapedAllNameCalculation,
      namedCurrentRowEscapedTotalsNameCalculation,
      thisRowQualifierEscapedRangeCalculation,
      thisRowQualifierEscapedAtCalculation,
      thisRowQualifierDoubleWrappedRangeCalculation,
      currentTableHeadersDataCalculation,
      currentTableDataCalculation,
      nestedStructuredReferenceCalculation,
      formulaStandaloneViewSetup,
      formulaStandaloneExport,
      formulaStandaloneReload,
      messages,
    };
    const qualifiedExpected = [
      `=${qualifiedSheetPrefix}!$B$2:$C$3`,
      `=${qualifiedSheetPrefix}!B$2:C$3`,
      `=${qualifiedSheetPrefix}!$B2:$C3`,
      `=${qualifiedSheetPrefix}!B2:C3`,
    ];
    const apostropheExpected = [
      `=${apostropheSheetPrefix}!$B$2:$C$3`,
      `=${apostropheSheetPrefix}!B$2:C$3`,
      `=${apostropheSheetPrefix}!$B2:$C3`,
      `=${apostropheSheetPrefix}!B2:C3`,
    ];
    const failures = [];
    if (messages.length) failures.push(`browser errors: ${messages.join(" | ")}`);
    if (definitions.SUM?.syntax !== "SUM(number1, [number2], ...)") failures.push("SUM definition syntax");
    if (definitions.SUMPRODUCT?.syntax !== "SUMPRODUCT(array1, [array2], ...)") failures.push("SUMPRODUCT definition syntax");
    if (definitions.XLOOKUP?.syntax !== "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])") failures.push("XLOOKUP definition syntax");
    if (definitions.FILTER?.syntax !== "FILTER(array, include, [if_empty])") failures.push("FILTER definition syntax");
    if (definitions.PMT?.category !== "財務") failures.push("PMT definition category");
    if (definitions["LOGNORM.DIST"]?.syntax !== "LOGNORM.DIST(x, mean, standard_dev, cumulative)") failures.push("LOGNORM.DIST definition syntax");
    if (definitions.DELTA?.category !== "エンジニアリング") failures.push("DELTA definition category");
    if (definitions.TBILLEQ?.category !== "財務") failures.push("TBILLEQ definition category");
    if (definitions.ISPMT?.syntax !== "ISPMT(rate, per, nper, pv)") failures.push("ISPMT definition syntax");
    if (!sumOpen.names.includes("SUM") || !sumOpen.names.includes("SUMIF")) failures.push("SUM candidate list");
    if (!sumOpen.syntaxes.some((syntax) => String(syntax || "").startsWith("SUM("))) failures.push("SUM candidate syntax column");
    if (!sumOpen.descriptions.some((description) => String(description || "").includes("合計"))) failures.push("SUM candidate description");
    if (sumOpen.activeDetailName !== "SUM" || !String(sumOpen.activeDetailSyntax || "").startsWith("SUM(")) failures.push("SUM active detail");
    if (sumAfterArrow.active !== "SUMIF" || sumAfterArrow.activeDetailName !== "SUMIF" || !String(sumAfterArrow.activeDetailSyntax || "").startsWith("SUMIF(")) failures.push("autocomplete active detail follows selection");
    if (sumOpen.role !== "listbox" || sumOpen.ariaLabel !== "数式オートコンプリート") failures.push("autocomplete accessibility");
    if (sumAccepted.inputValue !== "=SUM(" || !sumAccepted.hint.includes("SUM(number1")) failures.push("SUM accepted hint");
    if (!String(sumAccepted.activeDescription || "").includes("引数 1")) failures.push("SUM first argument ordinal");
    if (sumAccepted.listVisible) failures.push("candidate list should close after accepting SUM");
    if (parenAccepted.inputValue !== "=AVERAGE(" || !parenAccepted.hint.includes("AVERAGE(number1") || !String(parenAccepted.activeArgument || "").includes("number1") || parenAccepted.listVisible) failures.push("open paren accepts active function candidate");
    if (plainParen.inputValue !== "=(" || plainParen.popupVisible) failures.push("plain open paren should not accept an empty formula candidate");
    if (!String(repeatedArgument.activeArgument || "").includes("number3")) failures.push("SUM repeated argument active");
    if (!String(repeatedArgument.activeDescription || "").includes("引数 3")) failures.push("SUM repeated argument ordinal");
    if (qualifiedReferenceF4.join("|") !== qualifiedExpected.join("|")) failures.push("qualified sheet F4 cycle");
    if (!sheetReferenceOpen.names.includes(qualifiedSheetName) || !sheetReferenceOpen.categories.includes("シート")) failures.push("sheet reference candidate list");
    if (sheetReferenceAccepted !== `=${qualifiedSheetPrefix}!`) failures.push("sheet reference candidate insertion");
    if (!quotedSheetReferenceOpen.names.includes(qualifiedSheetName) || !quotedSheetReferenceOpen.categories.includes("シート")) failures.push("quoted sheet reference candidate list");
    if (quotedSheetReferenceAccepted !== `=${qualifiedSheetPrefix}!`) failures.push("quoted sheet reference candidate insertion");
    if (apostropheReferenceF4.join("|") !== apostropheExpected.join("|")) failures.push("apostrophe sheet F4 cycle");
    if (!apostropheSheetReferenceOpen.names.includes(apostropheSheetName) || !apostropheSheetReferenceOpen.syntaxes.includes(`${apostropheSheetPrefix}!`)) failures.push("apostrophe sheet reference candidate list");
    if (apostropheSheetReferenceAccepted !== `=${apostropheSheetPrefix}!`) failures.push("apostrophe sheet reference candidate insertion");
    if (!quotedApostropheSheetReferenceOpen.names.includes(apostropheSheetName) || !quotedApostropheSheetReferenceOpen.syntaxes.includes(`${apostropheSheetPrefix}!`)) failures.push("quoted apostrophe sheet reference candidate list");
    if (quotedApostropheSheetReferenceAccepted !== `=${apostropheSheetPrefix}!`) failures.push("quoted apostrophe sheet reference candidate insertion");
    if (!lognormOpen.names.includes("LOGNORM.DIST") || !lognormOpen.names.includes("LOGNORM.INV")) failures.push("LOGNORM candidate list");
    if (!lognormOpen.categories.includes("統計")) failures.push("LOGNORM category");
    if (!lognormOpen.syntaxes.includes("LOGNORM.DIST(x, mean, standard_dev, cumulative)")) failures.push("LOGNORM syntax column");
    if (!lognormOpen.descriptions.some((description) => String(description || "").includes("対数正規分布"))) failures.push("LOGNORM description");
    if (deltaOpen.active !== "DELTA" || !deltaOpen.names.includes("DELTA")) failures.push("DELTA active candidate");
    if (!deltaOpen.categories.includes("エンジニアリング")) failures.push("DELTA category");
    if (!deltaOpen.syntaxes.includes("DELTA(number1, [number2])")) failures.push("DELTA syntax column");
    if (xlookupOpen.active !== "XLOOKUP" || !xlookupOpen.syntaxes.includes("XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])")) failures.push("XLOOKUP candidate");
    if (!xlookupOpen.descriptions.some((description) => String(description || "").includes("検索"))) failures.push("XLOOKUP description");
    if (xlookupDialogInitial.title !== "関数の引数" || xlookupDialogInitial.functionName !== "XLOOKUP") failures.push("XLOOKUP arguments dialog title");
    if (!xlookupDialogInitial.activeLabels.includes("lookup_value")) failures.push("XLOOKUP initial active argument");
    if (!xlookupDialogInitial.description.includes("引数 1") || !xlookupDialogInitial.description.includes("lookup_value")) failures.push("XLOOKUP initial argument detail");
    if (xlookupDialogInitial.descriptionId !== "functionArgumentDescription") failures.push("function arguments description id");
    if (!xlookupDialogInitial.rows.every((row) => row.inputDescribedBy === "functionArgumentDescription")) failures.push("function arguments aria description linkage");
    if (!xlookupDialogReturnArray.activeLabels.includes("return_array")) failures.push("XLOOKUP focused active argument");
    if (!xlookupDialogReturnArray.description.includes("引数 3") || !xlookupDialogReturnArray.description.includes("return_array")) failures.push("XLOOKUP focused argument detail");
    if (!xlookupDialogReferenceCollapsed.activeLabels.includes("return_array") || !xlookupDialogReferenceCollapsed.description.includes("引数 3")) failures.push("XLOOKUP reference picker active argument");
    if (xlookupDialogReferenceF4 !== `${qualifiedSheetPrefix}!$B$2:$C$3`) failures.push("function argument F4 reference cycle");
    if (!filterOpen.names.includes("FILTER") || !filterOpen.syntaxes.includes("FILTER(array, include, [if_empty])")) failures.push("FILTER candidate");
    if (!filterOpen.descriptions.some((description) => String(description || "").includes("条件"))) failures.push("FILTER description");
    if (!pmtOpen.names.includes("PMT") || !pmtOpen.categories.includes("財務") || !pmtOpen.syntaxes.includes("PMT(rate, nper, pv, [fv], [type])")) failures.push("PMT candidate");
    if (!pmtOpen.descriptions.some((description) => String(description || "").includes("支払"))) failures.push("PMT description");
    if (definedNameOpen.active !== "SalesTotal" || !definedNameOpen.categories.includes("名前")) failures.push("defined name autocomplete");
    if (definedNameAccepted !== "=SalesTotal") failures.push("defined name insertion");
    if (tableNameOpen.active !== "SalesTable" || !tableNameOpen.categories.includes("テーブル")) failures.push("table name autocomplete");
    if (tableNameAccepted !== "=SalesTable") failures.push("table name insertion");
    if (!tableStructuredOpen.names.includes("#All") || !tableStructuredOpen.names.includes("Sales Amount") || !tableStructuredOpen.names.includes("値(税)")) failures.push("structured reference candidate list");
    if (!tableStructuredOpen.syntaxes.includes("SalesTable[#All]")) failures.push("structured reference special-item syntax");
    if (!tableStructuredOpen.syntaxes.includes("SalesTable[Sales Amount]")) failures.push("structured reference syntax");
    if (tableStructuredAccepted !== "=SalesTable[Sales Amount]") failures.push("structured reference insertion");
    if (!escapedHashColumnStructuredOpen.names.includes("税#率") || !escapedHashColumnStructuredOpen.syntaxes.includes("SalesTable[税'#率]")) failures.push("escaped # structured reference syntax");
    if (escapedHashColumnStructuredAccepted !== "=SalesTable[税'#率]") failures.push("escaped # structured reference insertion");
    if (!escapedApostropheColumnStructuredOpen.names.includes("O'Brien") || !escapedApostropheColumnStructuredOpen.syntaxes.includes("SalesTable[O''Brien]")) failures.push("escaped apostrophe structured reference syntax");
    if (escapedApostropheColumnStructuredAccepted !== "=SalesTable[O''Brien]") failures.push("escaped apostrophe structured reference insertion");
    if (!escapedApostropheTypedEscapeColumnStructuredOpen.names.includes("O'Brien") || !escapedApostropheTypedEscapeColumnStructuredOpen.syntaxes.includes("SalesTable[O''Brien]")) failures.push("escaped apostrophe typed structured reference syntax");
    if (escapedApostropheTypedEscapeColumnStructuredAccepted !== "=SalesTable[O''Brien]") failures.push("escaped apostrophe typed structured reference insertion");
    if (!escapedBracketColumnStructuredOpen.names.includes("端]数") || !escapedBracketColumnStructuredOpen.syntaxes.includes("SalesTable[端']数]")) failures.push("escaped bracket structured reference syntax");
    if (escapedBracketColumnStructuredAccepted !== "=SalesTable[端']数]") failures.push("escaped bracket structured reference insertion");
    if (!escapedOpenBracketColumnStructuredOpen.names.includes("始[数") || !escapedOpenBracketColumnStructuredOpen.syntaxes.includes("SalesTable[始'[数]")) failures.push("escaped open bracket structured reference syntax");
    if (escapedOpenBracketColumnStructuredAccepted !== "=SalesTable[始'[数]") failures.push("escaped open bracket structured reference insertion");
    if (!escapedOpenBracketTypedEscapeColumnStructuredOpen.names.includes("始[数") || !escapedOpenBracketTypedEscapeColumnStructuredOpen.syntaxes.includes("SalesTable[始'[数]")) failures.push("escaped open bracket typed structured reference syntax");
    if (escapedOpenBracketTypedEscapeColumnStructuredAccepted !== "=SalesTable[始'[数]") failures.push("escaped open bracket typed structured reference insertion");
    if (!escapedSpecialNameColumnStructuredOpen.names.includes("#Data") || !escapedSpecialNameColumnStructuredOpen.syntaxes.includes("SalesTable['#Data]")) failures.push("escaped #Data column-name structured reference syntax");
    if (escapedSpecialNameColumnStructuredAccepted !== "=SalesTable['#Data]") failures.push("escaped #Data column-name structured reference insertion");
    if (!escapedAtColumnStructuredOpen.names.includes("@担当") || !escapedAtColumnStructuredOpen.syntaxes.includes("SalesTable['@担当]")) failures.push("escaped @ structured reference syntax");
    if (escapedAtColumnStructuredAccepted !== "=SalesTable['@担当]") failures.push("escaped @ structured reference insertion");
    if (!escapedThisRowNameColumnStructuredOpen.names.includes("#This Row") || !escapedThisRowNameColumnStructuredOpen.syntaxes.includes("SalesTable['#This Row]")) failures.push("escaped #This Row column-name structured reference syntax");
    if (escapedThisRowNameColumnStructuredAccepted !== "=SalesTable['#This Row]") failures.push("escaped #This Row column-name structured reference insertion");
    if (!doubleWrappedCommaColumnStructuredOpen.names.includes("金額,税込") || !doubleWrappedCommaColumnStructuredOpen.syntaxes.includes("SalesTable[[金額,税込]]")) failures.push("double-wrapped comma structured reference syntax");
    if (doubleWrappedCommaColumnStructuredAccepted !== "=SalesTable[[金額,税込]]") failures.push("double-wrapped comma structured reference insertion");
    if (!doubleWrappedDotPercentColumnStructuredOpen.names.includes("比率.%") || !doubleWrappedDotPercentColumnStructuredOpen.syntaxes.includes("SalesTable[[比率.%]]")) failures.push("double-wrapped dot/percent structured reference syntax");
    if (doubleWrappedDotPercentColumnStructuredAccepted !== "=SalesTable[[比率.%]]") failures.push("double-wrapped dot/percent structured reference insertion");
    if (!doubleWrappedColonColumnStructuredOpen.names.includes("A:B") || !doubleWrappedColonColumnStructuredOpen.syntaxes.includes("SalesTable[[A:B]]")) failures.push("double-wrapped colon structured reference syntax");
    if (doubleWrappedColonColumnStructuredAccepted !== "=SalesTable[[A:B]]") failures.push("double-wrapped colon structured reference insertion");
    if (!doubleWrappedQuestionColumnStructuredOpen.names.includes("費用?") || !doubleWrappedQuestionColumnStructuredOpen.syntaxes.includes("SalesTable[[費用?]]")) failures.push("double-wrapped question structured reference syntax");
    if (doubleWrappedQuestionColumnStructuredAccepted !== "=SalesTable[[費用?]]") failures.push("double-wrapped question structured reference insertion");
    if (!doubleWrappedDollarColumnStructuredOpen.names.includes("売上$") || !doubleWrappedDollarColumnStructuredOpen.syntaxes.includes("SalesTable[[売上$]]")) failures.push("double-wrapped dollar structured reference syntax");
    if (doubleWrappedDollarColumnStructuredAccepted !== "=SalesTable[[売上$]]") failures.push("double-wrapped dollar structured reference insertion");
    if (!doubleWrappedPlusColumnStructuredOpen.names.includes("A+B") || !doubleWrappedPlusColumnStructuredOpen.syntaxes.includes("SalesTable[[A+B]]")) failures.push("double-wrapped plus structured reference syntax");
    if (doubleWrappedPlusColumnStructuredAccepted !== "=SalesTable[[A+B]]") failures.push("double-wrapped plus structured reference insertion");
    if (!doubleWrappedParenColumnStructuredOpen.names.includes("値(税)") || !doubleWrappedParenColumnStructuredOpen.syntaxes.includes("SalesTable[[値(税)]]")) failures.push("double-wrapped parenthesized structured reference syntax");
    if (doubleWrappedParenColumnStructuredAccepted !== "=SalesTable[[値(税)]]") failures.push("double-wrapped parenthesized structured reference insertion");
    if (!doubleWrappedEscapedAtCommaColumnStructuredOpen.names.includes("担当,@税込") || !doubleWrappedEscapedAtCommaColumnStructuredOpen.syntaxes.includes("SalesTable[[担当,'@税込]]")) failures.push("double-wrapped escaped @ comma structured reference syntax");
    if (doubleWrappedEscapedAtCommaColumnStructuredAccepted !== "=SalesTable[[担当,'@税込]]") failures.push("double-wrapped escaped @ comma structured reference insertion");
    if (!escapedAllNameColumnStructuredOpen.names.includes("#All") || !escapedAllNameColumnStructuredOpen.syntaxes.includes("SalesTable[#All]") || !escapedAllNameColumnStructuredOpen.syntaxes.includes("SalesTable['#All]")) failures.push("escaped #All column-name structured reference syntax");
    if (escapedAllNameColumnStructuredAccepted !== "=SalesTable['#All]") failures.push("escaped #All column-name structured reference insertion");
    if (!escapedAllNameTypedEscapeColumnStructuredOpen.names.includes("#All") || !escapedAllNameTypedEscapeColumnStructuredOpen.syntaxes.includes("SalesTable['#All]")) failures.push("escaped #All typed column-name structured reference syntax");
    if (escapedAllNameTypedEscapeColumnStructuredAccepted !== "=SalesTable['#All]") failures.push("escaped #All typed column-name structured reference insertion");
    if (!escapedHeadersNameColumnStructuredOpen.names.includes("#Headers") || !escapedHeadersNameColumnStructuredOpen.syntaxes.includes("SalesTable[#Headers]") || !escapedHeadersNameColumnStructuredOpen.syntaxes.includes("SalesTable['#Headers]")) failures.push("escaped #Headers column-name structured reference syntax");
    if (escapedHeadersNameColumnStructuredAccepted !== "=SalesTable['#Headers]") failures.push("escaped #Headers column-name structured reference insertion");
    if (!escapedTotalsNameColumnStructuredOpen.names.includes("#Totals") || !escapedTotalsNameColumnStructuredOpen.syntaxes.includes("SalesTable[#Totals]") || !escapedTotalsNameColumnStructuredOpen.syntaxes.includes("SalesTable['#Totals]")) failures.push("escaped #Totals column-name structured reference syntax");
    if (escapedTotalsNameColumnStructuredAccepted !== "=SalesTable['#Totals]") failures.push("escaped #Totals column-name structured reference insertion");
    if (!nestedSpecialStructuredOpen.names.includes("#All") || !nestedSpecialStructuredOpen.syntaxes.includes("SalesTable[[#All]]")) failures.push("nested special structured reference syntax");
    if (nestedSpecialStructuredAccepted !== "=SalesTable[[#All]]") failures.push("nested special structured reference insertion");
    if (!nestedThisRowStructuredOpen.names.includes("#This Row") || !nestedThisRowStructuredOpen.syntaxes.includes("SalesTable[[#This Row]]")) failures.push("nested #This Row structured reference syntax");
    if (nestedThisRowStructuredAccepted !== "=SalesTable[[#This Row]]") failures.push("nested #This Row structured reference insertion");
    if (!currentRowStructuredOpen.names.includes("Sales Amount") || !currentRowStructuredOpen.syntaxes.includes("[@[Sales Amount]]")) failures.push("current-row structured reference candidate");
    if (currentRowStructuredAccepted !== "=[@[Sales Amount]]") failures.push("current-row structured reference insertion");
    if (!namedCurrentRowStructuredOpen.names.includes("Sales Amount") || !namedCurrentRowStructuredOpen.syntaxes.includes("SalesTable[@[Sales Amount]]")) failures.push("named current-row structured reference candidate");
    if (namedCurrentRowStructuredAccepted !== "=SalesTable[@[Sales Amount]]") failures.push("named current-row structured reference insertion");
    if (!namedCurrentRowEscapedHashOpen.names.includes("税#率") || !namedCurrentRowEscapedHashOpen.syntaxes.includes("SalesTable[@[税'#率]]")) failures.push("escaped # named current-row structured reference candidate");
    if (namedCurrentRowEscapedHashAccepted !== "=SalesTable[@[税'#率]]") failures.push("escaped # named current-row structured reference insertion");
    if (!currentRowEscapedApostropheOpen.names.includes("O'Brien") || !currentRowEscapedApostropheOpen.syntaxes.includes("[@[O''Brien]]")) failures.push("escaped apostrophe current-row structured reference candidate");
    if (currentRowEscapedApostropheAccepted !== "=[@[O''Brien]]") failures.push("escaped apostrophe current-row structured reference insertion");
    if (!currentRowEscapedBracketOpen.names.includes("端]数") || !currentRowEscapedBracketOpen.syntaxes.includes("[@[端']数]]")) failures.push("escaped bracket current-row structured reference candidate");
    if (currentRowEscapedBracketAccepted !== "=[@[端']数]]") failures.push("escaped bracket current-row structured reference insertion");
    if (!currentRowEscapedOpenBracketTypedEscapeOpen.names.includes("始[数") || !currentRowEscapedOpenBracketTypedEscapeOpen.syntaxes.includes("[@[始'[数]]")) failures.push("escaped open bracket typed current-row structured reference candidate");
    if (currentRowEscapedOpenBracketTypedEscapeAccepted !== "=[@[始'[数]]") failures.push("escaped open bracket typed current-row structured reference insertion");
    if (!currentRowEscapedAtOpen.names.includes("@担当") || !currentRowEscapedAtOpen.syntaxes.includes("[@['@担当]]")) failures.push("escaped @ current-row structured reference candidate");
    if (currentRowEscapedAtAccepted !== "=[@['@担当]]") failures.push("escaped @ current-row structured reference insertion");
    if (!currentRowEscapedAtTypedEscapeOpen.names.includes("@担当") || !currentRowEscapedAtTypedEscapeOpen.syntaxes.includes("[@['@担当]]")) failures.push("escaped @ typed current-row structured reference candidate");
    if (currentRowEscapedAtTypedEscapeAccepted !== "=[@['@担当]]") failures.push("escaped @ typed current-row structured reference insertion");
    if (!namedCurrentRowEscapedAtOpen.names.includes("@担当") || !namedCurrentRowEscapedAtOpen.syntaxes.includes("SalesTable[@['@担当]]")) failures.push("escaped @ named current-row structured reference candidate");
    if (namedCurrentRowEscapedAtAccepted !== "=SalesTable[@['@担当]]") failures.push("escaped @ named current-row structured reference insertion");
    if (!currentRowDoubleWrappedCommaOpen.names.includes("金額,税込") || !currentRowDoubleWrappedCommaOpen.syntaxes.includes("[@[金額,税込]]")) failures.push("double-wrapped comma current-row structured reference candidate");
    if (currentRowDoubleWrappedCommaAccepted !== "=[@[金額,税込]]") failures.push("double-wrapped comma current-row structured reference insertion");
    if (!namedCurrentRowDoubleWrappedDotPercentOpen.names.includes("比率.%") || !namedCurrentRowDoubleWrappedDotPercentOpen.syntaxes.includes("SalesTable[@[比率.%]]")) failures.push("double-wrapped dot/percent named current-row structured reference candidate");
    if (namedCurrentRowDoubleWrappedDotPercentAccepted !== "=SalesTable[@[比率.%]]") failures.push("double-wrapped dot/percent named current-row structured reference insertion");
    if (!currentRowDoubleWrappedQuestionOpen.names.includes("費用?") || !currentRowDoubleWrappedQuestionOpen.syntaxes.includes("[@[費用?]]")) failures.push("double-wrapped question current-row structured reference candidate");
    if (currentRowDoubleWrappedQuestionAccepted !== "=[@[費用?]]") failures.push("double-wrapped question current-row structured reference insertion");
    if (!namedCurrentRowDoubleWrappedParenOpen.names.includes("値(税)") || !namedCurrentRowDoubleWrappedParenOpen.syntaxes.includes("SalesTable[@[値(税)]]")) failures.push("double-wrapped parenthesized named current-row structured reference candidate");
    if (namedCurrentRowDoubleWrappedParenAccepted !== "=SalesTable[@[値(税)]]") failures.push("double-wrapped parenthesized named current-row structured reference insertion");
    if (!currentRowDoubleWrappedEscapedAtCommaOpen.names.includes("担当,@税込") || !currentRowDoubleWrappedEscapedAtCommaOpen.syntaxes.includes("[@[担当,'@税込]]")) failures.push("double-wrapped escaped @ comma current-row structured reference candidate");
    if (currentRowDoubleWrappedEscapedAtCommaAccepted !== "=[@[担当,'@税込]]") failures.push("double-wrapped escaped @ comma current-row structured reference insertion");
    if (!currentRowEscapedAllNameOpen.names.includes("#All") || !currentRowEscapedAllNameOpen.syntaxes.includes("[@['#All]]")) failures.push("escaped #All current-row structured reference candidate");
    if (currentRowEscapedAllNameAccepted !== "=[@['#All]]") failures.push("escaped #All current-row structured reference insertion");
    if (!namedCurrentRowEscapedTotalsNameOpen.names.includes("#Totals") || !namedCurrentRowEscapedTotalsNameOpen.syntaxes.includes("SalesTable[@['#Totals]]")) failures.push("escaped #Totals named current-row structured reference candidate");
    if (namedCurrentRowEscapedTotalsNameAccepted !== "=SalesTable[@['#Totals]]") failures.push("escaped #Totals named current-row structured reference insertion");
    if (!namedCurrentRowEscapedRangeOpen.names.includes("端]数") || !namedCurrentRowEscapedRangeOpen.syntaxes.includes("SalesTable[@[税'#率]:[端']数]]")) failures.push("escaped named current-row range structured reference candidate");
    if (namedCurrentRowEscapedRangeAccepted !== "=SalesTable[@[税'#率]:[端']数]]") failures.push("escaped named current-row range structured reference insertion");
    if (!currentRowEscapedRangeOpen.names.includes("端]数") || !currentRowEscapedRangeOpen.syntaxes.includes("[@[O''Brien]:[端']数]]")) failures.push("escaped current-row range structured reference candidate");
    if (currentRowEscapedRangeAccepted !== "=[@[O''Brien]:[端']数]]") failures.push("escaped current-row range structured reference insertion");
    if (!dataQualifierStructuredOpen.names.includes("Sales Amount") || !dataQualifierStructuredOpen.syntaxes.includes("SalesTable[[#Data],[Sales Amount]]")) failures.push("#Data structured reference candidate");
    if (dataQualifierStructuredAccepted !== "=SalesTable[[#Data],[Sales Amount]]") failures.push("#Data structured reference insertion");
    if (!headersQualifierStructuredOpen.names.includes("Sales Amount") || !headersQualifierStructuredOpen.syntaxes.includes("SalesTable[[#Headers],[Sales Amount]]")) failures.push("#Headers structured reference candidate");
    if (headersQualifierStructuredAccepted !== "=SalesTable[[#Headers],[Sales Amount]]") failures.push("#Headers structured reference insertion");
    if (!totalsQualifierStructuredOpen.names.includes("Sales Amount") || !totalsQualifierStructuredOpen.syntaxes.includes("SalesTable[[#Totals],[Sales Amount]]")) failures.push("#Totals structured reference candidate");
    if (totalsQualifierStructuredAccepted !== "=SalesTable[[#Totals],[Sales Amount]]") failures.push("#Totals structured reference insertion");
    if (!dataQualifierEscapedHashOpen.names.includes("税#率") || !dataQualifierEscapedHashOpen.syntaxes.includes("SalesTable[[#Data],[税'#率]]")) failures.push("escaped # #Data structured reference candidate");
    if (dataQualifierEscapedHashAccepted !== "=SalesTable[[#Data],[税'#率]]") failures.push("escaped # #Data structured reference insertion");
    if (!totalsQualifierEscapedBracketOpen.names.includes("端]数") || !totalsQualifierEscapedBracketOpen.syntaxes.includes("SalesTable[[#Totals],[端']数]]")) failures.push("escaped bracket #Totals structured reference candidate");
    if (totalsQualifierEscapedBracketAccepted !== "=SalesTable[[#Totals],[端']数]]") failures.push("escaped bracket #Totals structured reference insertion");
    if (!dataQualifierEscapedOpenBracketTypedEscapeOpen.names.includes("始[数") || !dataQualifierEscapedOpenBracketTypedEscapeOpen.syntaxes.includes("SalesTable[[#Data],[始'[数]]")) failures.push("escaped open bracket typed #Data structured reference candidate");
    if (dataQualifierEscapedOpenBracketTypedEscapeAccepted !== "=SalesTable[[#Data],[始'[数]]") failures.push("escaped open bracket typed #Data structured reference insertion");
    if (!thisRowQualifierEscapedAtOpen.names.includes("@担当") || !thisRowQualifierEscapedAtOpen.syntaxes.includes("SalesTable[[#This Row],['@担当]]")) failures.push("escaped @ #This Row structured reference candidate");
    if (thisRowQualifierEscapedAtAccepted !== "=SalesTable[[#This Row],['@担当]]") failures.push("escaped @ #This Row structured reference insertion");
    if (!dataQualifierDoubleWrappedCommaOpen.names.includes("金額,税込") || !dataQualifierDoubleWrappedCommaOpen.syntaxes.includes("SalesTable[[#Data],[金額,税込]]")) failures.push("double-wrapped comma #Data structured reference candidate");
    if (dataQualifierDoubleWrappedCommaAccepted !== "=SalesTable[[#Data],[金額,税込]]") failures.push("double-wrapped comma #Data structured reference insertion");
    if (!dataQualifierEscapedAllNameOpen.names.includes("#All") || !dataQualifierEscapedAllNameOpen.syntaxes.includes("SalesTable[[#Data],['#All]]")) failures.push("escaped #All #Data structured reference candidate");
    if (dataQualifierEscapedAllNameAccepted !== "=SalesTable[[#Data],['#All]]") failures.push("escaped #All #Data structured reference insertion");
    if (!dataQualifierEscapedAllNameTypedEscapeOpen.names.includes("#All") || !dataQualifierEscapedAllNameTypedEscapeOpen.syntaxes.includes("SalesTable[[#Data],['#All]]")) failures.push("escaped #All typed #Data structured reference candidate");
    if (dataQualifierEscapedAllNameTypedEscapeAccepted !== "=SalesTable[[#Data],['#All]]") failures.push("escaped #All typed #Data structured reference insertion");
    if (!headersQualifierEscapedHeadersNameOpen.names.includes("#Headers") || !headersQualifierEscapedHeadersNameOpen.syntaxes.includes("SalesTable[[#Headers],['#Headers]]")) failures.push("escaped #Headers #Headers structured reference candidate");
    if (headersQualifierEscapedHeadersNameAccepted !== "=SalesTable[[#Headers],['#Headers]]") failures.push("escaped #Headers #Headers structured reference insertion");
    if (!totalsQualifierEscapedTotalsNameOpen.names.includes("#Totals") || !totalsQualifierEscapedTotalsNameOpen.syntaxes.includes("SalesTable[[#Totals],['#Totals]]")) failures.push("escaped #Totals #Totals structured reference candidate");
    if (totalsQualifierEscapedTotalsNameAccepted !== "=SalesTable[[#Totals],['#Totals]]") failures.push("escaped #Totals #Totals structured reference insertion");
    if (!columnRangeStructuredOpen.names.includes("Sales Amount") || !columnRangeStructuredOpen.syntaxes.includes("SalesTable[[数量]:[Sales Amount]]")) failures.push("column-range structured reference candidate");
    if (columnRangeStructuredAccepted !== "=SalesTable[[数量]:[Sales Amount]]") failures.push("column-range structured reference insertion");
    if (!escapedRangeStructuredOpen.names.includes("端]数") || !escapedRangeStructuredOpen.syntaxes.includes("SalesTable[[税'#率]:[端']数]]")) failures.push("escaped column-range structured reference candidate");
    if (escapedRangeStructuredAccepted !== "=SalesTable[[税'#率]:[端']数]]") failures.push("escaped column-range structured reference insertion");
    if (!doubleWrappedRangeStartOpen.names.includes("A:B") || !doubleWrappedRangeStartOpen.syntaxes.includes("SalesTable[[金額,税込]:[A:B]]")) failures.push("double-wrapped range-start structured reference candidate");
    if (doubleWrappedRangeStartAccepted !== "=SalesTable[[金額,税込]:[A:B]]") failures.push("double-wrapped range-start structured reference insertion");
    if (!dataQualifierEscapedRangeOpen.names.includes("端]数") || !dataQualifierEscapedRangeOpen.syntaxes.includes("SalesTable[[#Data],[税'#率]:[端']数]]")) failures.push("escaped #Data column-range structured reference candidate");
    if (dataQualifierEscapedRangeAccepted !== "=SalesTable[[#Data],[税'#率]:[端']数]]") failures.push("escaped #Data column-range structured reference insertion");
    if (!allQualifierEscapedRangeOpen.names.includes("端]数") || !allQualifierEscapedRangeOpen.syntaxes.includes("SalesTable[[#All],[O''Brien]:[端']数]]")) failures.push("escaped #All column-range structured reference candidate");
    if (allQualifierEscapedRangeAccepted !== "=SalesTable[[#All],[O''Brien]:[端']数]]") failures.push("escaped #All column-range structured reference insertion");
    if (!headersDataQualifierOpen.names.includes("Sales Amount") || !headersDataQualifierOpen.syntaxes.includes("SalesTable[[#Headers],[#Data],[Sales Amount]]")) failures.push("#Headers #Data structured reference candidate");
    if (headersDataQualifierAccepted !== "=SalesTable[[#Headers],[#Data],[Sales Amount]]") failures.push("#Headers #Data structured reference insertion");
    if (!headersDataEscapedRangeOpen.names.includes("端]数") || !headersDataEscapedRangeOpen.syntaxes.includes("SalesTable[[#Headers],[#Data],[税'#率]:[端']数]]")) failures.push("escaped #Headers #Data range structured reference candidate");
    if (headersDataEscapedRangeAccepted !== "=SalesTable[[#Headers],[#Data],[税'#率]:[端']数]]") failures.push("escaped #Headers #Data range structured reference insertion");
    if (!currentTableHeadersDataQualifierOpen.names.includes("Sales Amount") || !currentTableHeadersDataQualifierOpen.syntaxes.includes("[[#Headers],[#Data],[Sales Amount]]")) failures.push("current-table #Headers #Data structured reference candidate");
    if (currentTableHeadersDataQualifierAccepted !== "=[[#Headers],[#Data],[Sales Amount]]") failures.push("current-table #Headers #Data structured reference insertion");
    if (!currentTableDataQualifierOpen.names.includes("Sales Amount") || !currentTableDataQualifierOpen.syntaxes.includes("[[#Data],[Sales Amount]]")) failures.push("current-table #Data structured reference candidate");
    if (currentTableDataQualifierAccepted !== "=[[#Data],[Sales Amount]]") failures.push("current-table #Data structured reference insertion");
    if (!currentTableDataQualifierEscapedAllNameOpen.names.includes("#All") || !currentTableDataQualifierEscapedAllNameOpen.syntaxes.includes("[[#Data],['#All]]")) failures.push("escaped #All current-table #Data structured reference candidate");
    if (currentTableDataQualifierEscapedAllNameAccepted !== "=[[#Data],['#All]]") failures.push("escaped #All current-table #Data structured reference insertion");
    if (shorthandStructuredReferenceCalculation.shorthand !== "21" || shorthandStructuredReferenceCalculation.propagated !== "31") failures.push("shorthand structured reference formula calculation");
    if (namedStructuredReferenceCalculation.propagated !== "22" || namedStructuredReferenceCalculation.named !== "32") failures.push("named structured reference formula calculation");
    if (currentRowEscapedHashCalculation.shorthand !== "6" || currentRowEscapedHashCalculation.propagated !== "9") failures.push("escaped # current-row structured reference formula calculation");
    if (namedCurrentRowEscapedBracketCalculation.propagated !== "5" || namedCurrentRowEscapedBracketCalculation.named !== "6") failures.push("escaped bracket named current-row structured reference formula calculation");
    if (currentRowEscapedRangeCalculation.shorthand !== "15" || currentRowEscapedRangeCalculation.propagated !== "21") failures.push("escaped current-row range structured reference formula calculation");
    if (namedCurrentRowEscapedRangeCalculation.propagated !== "15" || namedCurrentRowEscapedRangeCalculation.named !== "21") failures.push("escaped named current-row range structured reference formula calculation");
    if (currentRowEscapedAtCalculation.shorthand !== "12" || currentRowEscapedAtCalculation.propagated !== "13") failures.push("escaped @ current-row structured reference formula calculation");
    if (currentRowDoubleWrappedCommaCalculation.shorthand !== "14" || currentRowDoubleWrappedCommaCalculation.propagated !== "18") failures.push("double-wrapped comma current-row structured reference formula calculation");
    if (namedCurrentRowDoubleWrappedDotPercentCalculation.propagated !== "31" || namedCurrentRowDoubleWrappedDotPercentCalculation.named !== "32") failures.push("double-wrapped dot/percent named current-row structured reference formula calculation");
    if (currentRowDoubleWrappedQuestionCalculation.shorthand !== "42" || currentRowDoubleWrappedQuestionCalculation.propagated !== "45") failures.push("double-wrapped question current-row structured reference formula calculation");
    if (namedCurrentRowDoubleWrappedParenCalculation.propagated !== "67" || namedCurrentRowDoubleWrappedParenCalculation.named !== "69") failures.push("double-wrapped parenthesized named current-row structured reference formula calculation");
    if (currentRowDoubleWrappedEscapedAtCommaCalculation.shorthand !== "93" || currentRowDoubleWrappedEscapedAtCommaCalculation.propagated !== "99") failures.push("double-wrapped escaped @ comma current-row structured reference formula calculation");
    if (currentRowEscapedAllNameCalculation.shorthand !== "111" || currentRowEscapedAllNameCalculation.propagated !== "123") failures.push("escaped #All current-row structured reference formula calculation");
    if (namedCurrentRowEscapedTotalsNameCalculation.propagated !== "131" || namedCurrentRowEscapedTotalsNameCalculation.named !== "143") failures.push("escaped #Totals named current-row structured reference formula calculation");
    if (thisRowQualifierEscapedRangeCalculation.shorthand !== "15" || thisRowQualifierEscapedRangeCalculation.propagated !== "21") failures.push("#This Row escaped range structured reference formula calculation");
    if (thisRowQualifierEscapedAtCalculation.propagated !== "22" || thisRowQualifierEscapedAtCalculation.named !== "23") failures.push("#This Row escaped @ structured reference formula calculation");
    if (thisRowQualifierDoubleWrappedRangeCalculation.shorthand !== "35" || thisRowQualifierDoubleWrappedRangeCalculation.propagated !== "41") failures.push("#This Row double-wrapped range structured reference formula calculation");
    if (currentTableHeadersDataCalculation.shorthand !== "3" || currentTableHeadersDataCalculation.propagated !== "3") failures.push("current-table #Headers #Data structured reference formula calculation");
    if (currentTableDataCalculation.shorthand !== "50" || currentTableDataCalculation.propagated !== "50") failures.push("current-table #Data structured reference formula calculation");
    if (
      nestedStructuredReferenceCalculation.data !== "50" ||
      nestedStructuredReferenceCalculation.headers !== "1" ||
      nestedStructuredReferenceCalculation.columnRange !== "60" ||
      nestedStructuredReferenceCalculation.totals !== "50" ||
      nestedStructuredReferenceCalculation.all !== "100" ||
      nestedStructuredReferenceCalculation.hash !== "13" ||
      nestedStructuredReferenceCalculation.apostrophe !== "16" ||
      nestedStructuredReferenceCalculation.bracket !== "7" ||
      nestedStructuredReferenceCalculation.escapedRange !== "36" ||
      nestedStructuredReferenceCalculation.escapedTotals !== "13" ||
      nestedStructuredReferenceCalculation.escapedAll !== "14" ||
      nestedStructuredReferenceCalculation.rowQualifiedEscapedRangeData !== "36" ||
      nestedStructuredReferenceCalculation.rowQualifiedEscapedRangeTotals !== "36" ||
      nestedStructuredReferenceCalculation.rowQualifiedEscapedRangeAll !== "72" ||
      nestedStructuredReferenceCalculation.rowQualifiedEscapedRangeHeaders !== "3" ||
      nestedStructuredReferenceCalculation.combinedHeadersData !== "3" ||
      nestedStructuredReferenceCalculation.combinedDataTotals !== "100" ||
      nestedStructuredReferenceCalculation.combinedHeadersDataTotals !== "100" ||
      nestedStructuredReferenceCalculation.combinedHeadersDataEscapedRange !== "36" ||
      nestedStructuredReferenceCalculation.specialNameColumn !== "23" ||
      nestedStructuredReferenceCalculation.specialNameColumnTotals !== "99" ||
      nestedStructuredReferenceCalculation.atNameColumn !== "5" ||
      nestedStructuredReferenceCalculation.thisRowNameColumn !== "10" ||
      nestedStructuredReferenceCalculation.specialNameColumnAll !== "122" ||
      nestedStructuredReferenceCalculation.commaColumn !== "30" ||
      nestedStructuredReferenceCalculation.commaColumnTotals !== "301" ||
      nestedStructuredReferenceCalculation.commaColumnAll !== "331" ||
      nestedStructuredReferenceCalculation.dotPercentColumn !== "3" ||
      nestedStructuredReferenceCalculation.colonColumn !== "43" ||
      nestedStructuredReferenceCalculation.doubleWrappedRange !== "76" ||
      nestedStructuredReferenceCalculation.questionColumn !== "7" ||
      nestedStructuredReferenceCalculation.dollarColumn !== "10" ||
      nestedStructuredReferenceCalculation.plusColumn !== "24" ||
      nestedStructuredReferenceCalculation.parenColumn !== "36" ||
      nestedStructuredReferenceCalculation.punctuationRange !== "77" ||
      nestedStructuredReferenceCalculation.escapedAtCommaColumn !== "52" ||
      nestedStructuredReferenceCalculation.escapedAtCommaColumnAll !== "172" ||
      nestedStructuredReferenceCalculation.allNameColumn !== "74" ||
      nestedStructuredReferenceCalculation.headersNameColumnHeader !== "1" ||
      nestedStructuredReferenceCalculation.totalsNameColumnTotals !== "410" ||
      nestedStructuredReferenceCalculation.totalsNameColumnAll !== "504" ||
      nestedStructuredReferenceCalculation.openBracketColumn !== "116" ||
      nestedStructuredReferenceCalculation.openBracketColumnTotals !== "590" ||
      nestedStructuredReferenceCalculation.openBracketColumnAll !== "706" ||
      nestedStructuredReferenceCalculation.apostropheSheetRange !== "579"
    ) failures.push("nested structured reference formula calculation");
    if (!formulaStandaloneExport.isHtml) failures.push("formula standalone quick save should download WebSheet HTML");
    if (formulaStandaloneExport.sourceName !== formulaStandaloneExport.suggestedFilename) failures.push("formula standalone exported sourceName");
    if (formulaStandaloneExport.saveType !== "html") failures.push("formula standalone exported saveType");
    if (!formulaStandaloneExport.sheetNames.includes(apostropheSheetName)) failures.push("formula standalone exported apostrophe sheet");
    if (formulaStandaloneExport.formulaCell?.formula !== `=SUM(${apostropheSheetPrefix}!B2:C3)`) failures.push("formula standalone exported apostrophe formula");
    if (Number(formulaStandaloneExport.apostropheSheetValues.b2) !== 123 || Number(formulaStandaloneExport.apostropheSheetValues.c3) !== 456) failures.push("formula standalone exported apostrophe sheet values");
    if (formulaStandaloneViewSetup?.viewMode !== "page-layout" || formulaStandaloneViewSetup?.zoomScale !== 125) failures.push("formula standalone view setup");
    if (formulaStandaloneExport.view?.viewMode !== "page-layout" || formulaStandaloneExport.view?.zoomScale !== 125) failures.push("formula standalone exported view mode/zoom");
    if (
      formulaStandaloneExport.view?.showGridLines !== false ||
      formulaStandaloneExport.view?.showHeadings !== false ||
      formulaStandaloneExport.view?.showZeros !== false ||
      formulaStandaloneExport.view?.showObjects !== false ||
      formulaStandaloneExport.view?.showOutlineSymbols !== false ||
      formulaStandaloneExport.view?.showWhiteSpace !== false ||
      formulaStandaloneExport.view?.rightToLeft !== true
    ) failures.push("formula standalone exported view toggles");
    if (formulaStandaloneExport.view?.freezePanes?.rows !== 2 || formulaStandaloneExport.view?.freezePanes?.columns !== 1) failures.push("formula standalone exported freeze panes");
    if (formulaStandaloneReload.messages?.length) failures.push(`formula standalone reload browser errors: ${formulaStandaloneReload.messages.join(" | ")}`);
    if (formulaStandaloneReload.blockedRequests?.length) failures.push(`formula standalone reload tried remote runtime: ${formulaStandaloneReload.blockedRequests.join(" | ")}`);
    if (formulaStandaloneReload.runtimeSource !== "local" && formulaStandaloneReload.diagnostics?.runtimeSource !== "local") failures.push("formula standalone reload local runtime source");
    if (formulaStandaloneReload.firstTab !== "CHT" || formulaStandaloneReload.sheetCount < 3) failures.push("formula standalone reload workbook shell");
    if (formulaStandaloneReload.diagnostics?.activeSheetView?.viewMode !== "page-layout" || formulaStandaloneReload.diagnostics?.activeSheetView?.zoomScale !== 125) failures.push("formula standalone reload view mode/zoom");
    if (
      formulaStandaloneReload.diagnostics?.activeSheetView?.showGridLines !== false ||
      formulaStandaloneReload.diagnostics?.activeSheetView?.showHeadings !== false ||
      formulaStandaloneReload.diagnostics?.activeSheetView?.showZeros !== false ||
      formulaStandaloneReload.diagnostics?.activeSheetView?.showObjects !== false ||
      formulaStandaloneReload.diagnostics?.activeSheetView?.showOutlineSymbols !== false ||
      formulaStandaloneReload.diagnostics?.activeSheetView?.showWhiteSpace !== false ||
      formulaStandaloneReload.diagnostics?.activeSheetView?.rightToLeft !== true
    ) failures.push("formula standalone reload view toggles");
    if (formulaStandaloneReload.diagnostics?.activeSheetView?.freezePanes?.rows !== 2 || formulaStandaloneReload.diagnostics?.activeSheetView?.freezePanes?.columns !== 1) failures.push("formula standalone reload freeze panes");
    if (formulaStandaloneReload.diagnostics?.workbookView?.windowCount < 1 || formulaStandaloneReload.diagnostics?.workbookView?.showSheetTabs !== true) failures.push("formula standalone reload workbook view diagnostics");
    if (formulaStandaloneReload.cell?.display !== "579" && formulaStandaloneReload.devDisplay !== "579") failures.push("formula standalone reload display value");
    if (Number(formulaStandaloneReload.cell?.engine?.value ?? formulaStandaloneReload.devEngine?.value) !== 579) failures.push("formula standalone reload formula engine value");
    if (formulaStandaloneReload.cell?.cell?.formula !== `=SUM(${apostropheSheetPrefix}!B2:C3)` && formulaStandaloneReload.devCell?.formula !== `=SUM(${apostropheSheetPrefix}!B2:C3)`) failures.push("formula standalone reload formula raw");
    if (failures.length) throw new Error(`Formula UI smoke failed: ${failures.join(", ")} ${JSON.stringify(result)}`);
    console.log(JSON.stringify({ ok: true, result }, null, 2));
  } finally {
    await browser.close();
  }
}

export {
  runFormulaUiSmokeOnly,
};
