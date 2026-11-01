// Sequential E2E feature module.
import {
  appUrl,
  browser,
  chromePath,
  chromium,
  mkdtemp,
  os,
  path,
  pathToFileURL,
  readFile,
  workbookPath,
  writeFile,
} from "../../e2e-real-upload.mjs";
import {
  createMacroVbaSubsetWorkbook,
  exportedStandaloneModelFromHtml,
  looksLikeStandaloneWorkbookHtml,
  replaceStandaloneModelPayload,
  waitForUploadedWorkbook,
} from "./fixtures-and-model.mjs";
import {
  exerciseAutoFilterColorImportExport,
  exerciseAutoFilterConditionMenu,
  exerciseDataValidationMenu,
  exerciseImportedAutoFilterCriteria,
} from "./editing-and-validation.mjs";
import {
  exerciseOutlineGrouping,
  exerciseQuickSortActiveColumn,
  exerciseQuickSortPartialSelection,
  exerciseSortDialogColorOrder,
  exerciseSortDialogCustomList,
  exerciseSortDialogLevels,
} from "./connections-and-analysis.mjs";
import {
  exerciseAutoFilterDateMenu,
  exerciseAutoFilterReapplyAndClear,
} from "./views-and-macros.mjs";
import {
  exerciseDataCurrentRegion,
} from "./query-transform.mjs";
import {
  exerciseCustomViewImportExport,
} from "./formula-smoke.mjs";


async function runDatePersistenceSmokeOnly() {
  const browser = await chromium.launch({
    headless: true,
    executablePath: chromePath,
  });
  try {
    const context = await browser.newContext({
      viewport: { width: 960, height: 540 },
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
    await page.goto(`${appUrl}?datePersistence=${Date.now()}`, { waitUntil: "load" });
    await page.waitForFunction(() => typeof window.__webSheetDevNormalizeDateCell === "function", null, { timeout: 10000 });
    const result = await page.evaluate(() => {
      const jsDate = "Tue Feb 01 2005 09:00:00 GMT+0900 (日本標準時)";
      const normalize = (cell) => window.__webSheetDevNormalizeDateCell(cell);
      return {
        jsDateFormatted: normalize({
          raw: jsDate,
          cached: jsDate,
          display: jsDate,
          numFmt: "m/d/yy h:mm",
        }),
        jsDateNoFormat: normalize({
          raw: jsDate,
          cached: jsDate,
          display: jsDate,
          numFmt: "",
        }),
        japaneseLiteralDate: normalize({
          raw: "2005年2月1日 09:00:00",
          cached: "2005年2月1日 09:00:00",
          display: "2005年2月1日 09:00:00",
          numFmt: 'yyyy"年"m"月"d"日 h:mm',
        }),
        isoDateFormatted: normalize({
          raw: "2005-02-01T09:00:00+09:00",
          cached: "2005-02-01T09:00:00+09:00",
          display: "2005-02-01T09:00:00+09:00",
          numFmt: "yyyy/m/d h:mm",
        }),
      };
    });
    const failures = [];
    const serialClose = (value) => Math.abs(Number(value) - 38384.375) < 0.00001;
    if (!serialClose(result.jsDateFormatted.raw) || result.jsDateFormatted.rendered !== "2/1/05 9:00") {
      failures.push("JS Date string with Japanese timezone should restore to Excel serial and formatted display");
    }
    if (!serialClose(result.jsDateNoFormat.raw) || result.jsDateNoFormat.rendered !== "2005/2/1 9:00") {
      failures.push("JS Date string without numFmt should not render the long Date.toString text");
    }
    if (!serialClose(result.japaneseLiteralDate.raw) || result.japaneseLiteralDate.rendered !== "2005年2月1日 9:00") {
      failures.push("Japanese literal persisted date should restore with Japanese Excel date format");
    }
    if (!serialClose(result.isoDateFormatted.raw) || result.isoDateFormatted.rendered !== "2005/2/1 9:00") {
      failures.push("ISO persisted date should restore with Excel date format");
    }
    if (messages.length) failures.push(`browser errors: ${messages.join(" | ")}`);
    if (failures.length) throw new Error(`Date persistence smoke failed: ${failures.join("; ")}\n${JSON.stringify(result, null, 2)}`);
    console.log(JSON.stringify({ ok: true, datePersistence: result }, null, 2));
  } finally {
    await browser.close();
  }
}



async function runClipboardSessionSmokeOnly() {
  const browser = await chromium.launch({
    headless: true,
    executablePath: chromePath,
  });
  try {
    const context = await browser.newContext({
      viewport: { width: 1280, height: 720 },
      acceptDownloads: true,
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
    await page.goto(`${appUrl}?clipboardSession=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForFunction(
      () =>
        typeof window.__webSheetDevSelectRanges === "function" &&
        typeof window.__webSheetDevClipboardSessionRoundTrip === "function",
      null,
      { timeout: 10000 },
    );
    const result = await page.evaluate(async () => {
      window.__webSheetDevSelectRanges([{ top: 5, left: 2, bottom: 5, right: 2 }]);
      return window.__webSheetDevClipboardSessionRoundTrip();
    });
    const failures = [];
    if (!result?.restored) failures.push("internal clipboard should restore from session storage");
    if (result?.sourceLabel !== "B5:G5") failures.push(`expected merged B5:G5 source label, got ${result?.sourceLabel || ""}`);
    if (result?.rowCount !== 1 || result?.colCount !== 6) failures.push("restored clipboard should keep merged 1x6 size");
    if (result?.cellCount < 1) failures.push("restored clipboard should include copied cell model");
    if (String(result?.text || "").trim() === "") failures.push("restored clipboard text should not be empty");
    if (messages.length) failures.push(`browser errors: ${messages.join(" | ")}`);
    if (failures.length) throw new Error(`Clipboard session smoke failed: ${failures.join("; ")}\n${JSON.stringify(result, null, 2)}`);
    console.log(JSON.stringify({ ok: true, clipboardSession: result }, null, 2));
  } finally {
    await browser.close();
  }
}



async function runAltEnterSmokeOnly() {
  const browser = await chromium.launch({
    headless: true,
    executablePath: chromePath,
  });
  try {
    const context = await browser.newContext({
      viewport: { width: 1280, height: 720 },
      acceptDownloads: true,
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
    await page.goto(`${appUrl}?altEnter=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForFunction(() => typeof window.__webSheetDevSelectRanges === "function", null, { timeout: 10000 });
    await page.evaluate(() => window.__webSheetDevSelectRanges([{ top: 5, left: 2, bottom: 5, right: 2 }]));
    await page.keyboard.press("F2");
    const editor = page.locator(".sheet-cell[contenteditable='true']");
    await editor.waitFor({ timeout: 10000 });
    const beforeMetrics = await editor.evaluate((element) => ({
      scrollHeight: element.scrollHeight,
      clientHeight: element.clientHeight,
    }));
    await page.keyboard.press("Alt+Enter");
    const first = await editor.evaluate((element) => ({
      text: (element.textContent || "").replace(/\u200b/g, ""),
      rawText: element.textContent || "",
      scrollHeight: element.scrollHeight,
      clientHeight: element.clientHeight,
    }));
    await page.keyboard.press("Alt+Enter");
    const second = await editor.evaluate((element) => ({
      text: (element.textContent || "").replace(/\u200b/g, ""),
      rawText: element.textContent || "",
      scrollHeight: element.scrollHeight,
      clientHeight: element.clientHeight,
    }));
    await page.keyboard.press("Enter");
    const committedCell = await page.evaluate(() => window.__webSheetDevCellModel?.("CHT", 5, 2) || null);
    const committedText = String(committedCell?.raw ?? committedCell?.formula ?? committedCell?.display ?? "");

    await page.evaluate(() => window.__webSheetDevSelectRanges([{ top: 5, left: 2, bottom: 5, right: 2 }]));
    await page.keyboard.press("F2");
    await editor.waitFor({ timeout: 10000 });
    await editor.evaluate((element) => {
      element.textContent = "=S";
      const range = document.createRange();
      range.selectNodeContents(element);
      range.collapse(false);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      element.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: "S" }));
    });
    await page.waitForSelector("#formulaAutocomplete", { timeout: 3000 });
    await page.keyboard.press("Alt+Enter");
    const autocomplete = await editor.evaluate((element) => ({
      text: (element.textContent || "").replace(/\u200b/g, ""),
      rawText: element.textContent || "",
      scrollHeight: element.scrollHeight,
      clientHeight: element.clientHeight,
    }));
    const autocompletePopupCount = await page.locator("#formulaAutocomplete").count();
    await page.keyboard.press("Escape");

    const result = {
      beforeMetrics,
      firstText: first.text,
      firstRawText: first.rawText,
      firstNewlines: (first.text.match(/\n/g) || []).length,
      firstVisibleLineHeight: first.scrollHeight,
      secondText: second.text,
      secondRawText: second.rawText,
      secondNewlines: (second.text.match(/\n/g) || []).length,
      secondVisibleLineHeight: second.scrollHeight,
      committedText,
      committedNewlines: (committedText.match(/\n/g) || []).length,
      autocompleteText: autocomplete.text,
      autocompleteRawText: autocomplete.rawText,
      autocompleteNewlines: (autocomplete.text.match(/\n/g) || []).length,
      autocompleteVisibleLineHeight: autocomplete.scrollHeight,
      autocompletePopupCount,
    };
    const failures = [];
    if (result.firstNewlines !== 1) failures.push("first Alt+Enter should insert one line break");
    if (result.firstVisibleLineHeight <= result.beforeMetrics.scrollHeight) failures.push("first Alt+Enter should create a visible second line");
    if (result.secondNewlines !== 2) failures.push("second Alt+Enter should insert a second line break");
    if (result.committedText.includes("\u200b") || result.committedNewlines !== 2) failures.push("committed cell text should keep line breaks without edit sentinels");
    if (result.autocompleteText !== "=S\n") failures.push("Alt+Enter should not be consumed by formula autocomplete");
    if (messages.length) failures.push(`browser errors: ${messages.join(" | ")}`);
    if (failures.length) throw new Error(`Alt+Enter smoke failed: ${failures.join("; ")}\n${JSON.stringify(result, null, 2)}`);
    console.log(JSON.stringify({ ok: true, altEnter: result }, null, 2));
  } finally {
    await browser.close();
  }
}



async function runMacroVbaSubsetOnly() {
  const macroWorkbookPath = await createMacroVbaSubsetWorkbook();
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
    await page.goto(`${appUrl}?macroVbaSubset=${Date.now()}`);
    await page.setInputFiles("#fileInput", macroWorkbookPath);
    await waitForUploadedWorkbook(page, macroWorkbookPath);
    await page.evaluate(() => window.__webSheetDevSetActiveSheet?.("CHT"));
    const macroCode = `Sub MacroVbaSubsetSmoke()
  Dim i
  total = WorksheetFunction.Sum(Range("J6:J6"), 8)
  Worksheets(1).Cells(6, 22).Value = "CellsValue"
  Worksheets("CHT").Range(Cells(6, 23), Cells(6, 24)).Value = Array("Array1", "Array2")
  Range("Y6").FormulaR1C1 = "=RC[-15]+1"
  Range("Z6").Value = total
  Range("AB6").Value = Application.WorksheetFunction.Average(Range("J6:J6"), 46)
  Range("AC6").Value = total + 2
  Range("AA6").Value = "Indexed"
  Range("AA6").Font.ColorIndex = 2
  Range("AA6").Interior.ColorIndex = 5
  Range("AA6").HorizontalAlignment = xlHAlignCenter
  Range("AA6").VerticalAlignment = xlVAlignBottom
  Range("AA6").Font.Underline = xlUnderlineStyleDouble
  For i = 1 To 3
    Cells(i + 8, 2).Value = i * 2
  Next i
  If total >= 50 And Range("W6").Value = "Array1" Then
    Range("AD6").Value = "IfTrue"
  Else
    Range("AD6").Value = "IfFalse"
  End If
  If total < 50 Then
    Range("AE6").Value = "Bad"
  ElseIf Range("AB6").Value = 44 Then
    Range("AE6").Value = "ElseIfTrue"
  Else
    Range("AE6").Value = "ElseFalse"
  End If
  If Not (Range("AB6").Value < 44) Then Range("AF6").Value = "InlineIf"
  Range("AG6").Value = WorksheetFunction.CountIf(Range("B9:B11"), ">2")
  Range("AH6").Value = WorksheetFunction.SumIf(Range("B9:B11"), ">2", Range("B9:B11"))
  Range("AI6").Value = WorksheetFunction.Round(2.345, 2)
  Range("AJ6").Value = WorksheetFunction.Left("ABCDE", 2) & WorksheetFunction.Right("ABCDE", 2)
  Range("AK6").Value = WorksheetFunction.Mid("ABCDE", 2, 3)
  i = 0
  doTotal = 0
  Do While i < 10
    i = i + 1
    If i = 4 Then Exit Do
    doTotal = doTotal + i
  Loop
  Range("AL6").Value = doTotal
  j = 0
  Do
    j = j + 1
  Loop Until j >= 3
  Range("AM6").Value = j
  evenTotal = 0
  For i = 1 To 6
    If i = 5 Then Exit For
    If i Mod 2 = 0 Then evenTotal = evenTotal + i
  Next i
  Range("AN6").Value = evenTotal
  Select Case Range("AN6").Value
    Case 1 To 5
      Range("AO6").Value = "CaseLow"
    Case Is >= 6
      Range("AO6").Value = "CaseHigh"
    Case Else
      Range("AO6").Value = "CaseElse"
  End Select
  Range("C9").Value = "Low"
  Range("C10").Value = "Mid"
  Range("C11").Value = "High"
  Range("AP6").Value = WorksheetFunction.VLookup(4, Range("B9:C11"), 2, False)
  Range("AQ6").Value = WorksheetFunction.Match(6, Range("B9:B11"), 0)
  Range("AR6").Value = WorksheetFunction.Index(Range("C9:C11"), 3, 1)
  Range("AS6").Value = WorksheetFunction.CountIfs(Range("B9:B11"), ">=4", Range("C9:C11"), "H*")
  Range("AT6").Value = WorksheetFunction.SumIfs(Range("B9:B11"), Range("B9:B11"), ">=4", Range("C9:C11"), "<>Mid")
  Range("AU6").Value = WorksheetFunction.AverageIf(Range("B9:B11"), ">2", Range("B9:B11"))
  words = Array("AA", "BB", "CC")
  Range("AV6").Value = LBound(words) & ":" & UBound(words) & ":" & words(1)
  Range("AW6").Value = InStr(1, "ABCDE", "CD") & ":" & Replace("A-B-C", "-", "/")
  eachTotal = 0
  For Each cell In Range("B9:B11")
    eachTotal = eachTotal + cell.Value
    cell.Offset(0, 2).Value = cell.Row
  Next cell
  Range("AX6").Value = eachTotal
  Range("AY6").Value = Range("B9:B11").Rows.Count & ":" & Range("B9:B11").Count & ":" & Range("B9").Address
End Sub`;
    const runResult = await page.evaluate((code) => {
      window.__webSheetDevSetMacros?.([{ name: "MacroVbaSubsetSmoke", code }]);
      return window.__webSheetDevRunMacro?.("MacroVbaSubsetSmoke", { showMessages: false });
    }, macroCode);
    await page.waitForFunction(
      () => window.__webSheetDevCellModel?.("CHT", 6, 22)?.raw === "CellsValue",
      null,
      { timeout: 10000 },
    );
    const afterRun = await page.evaluate(() => {
      const get = (row, col) => window.__webSheetDevCellModel?.("CHT", row, col) || null;
      return {
        v6: get(6, 22),
        w6: get(6, 23),
        x6: get(6, 24),
        y6: get(6, 25),
        z6: get(6, 26),
        aa6: get(6, 27),
        ab6: get(6, 28),
        ac6: get(6, 29),
        b9: get(9, 2),
        b10: get(10, 2),
        b11: get(11, 2),
        ad6: get(6, 30),
        ae6: get(6, 31),
        af6: get(6, 32),
        ag6: get(6, 33),
        ah6: get(6, 34),
        ai6: get(6, 35),
        aj6: get(6, 36),
        ak6: get(6, 37),
        al6: get(6, 38),
        am6: get(6, 39),
        an6: get(6, 40),
        ao6: get(6, 41),
        ap6: get(6, 42),
        aq6: get(6, 43),
        ar6: get(6, 44),
        as6: get(6, 45),
        at6: get(6, 46),
        au6: get(6, 47),
        av6: get(6, 48),
        aw6: get(6, 49),
        ax6: get(6, 50),
        ay6: get(6, 51),
        d9: get(9, 4),
        d10: get(10, 4),
        d11: get(11, 4),
        history: window.__webSheetDevHistoryState?.(),
      };
    });
    const undoShortcut = process.platform === "darwin" ? "Meta+Z" : "Control+Z";
    const redoShortcut = process.platform === "darwin" ? "Meta+Y" : "Control+Y";
    await page.keyboard.press(undoShortcut);
    await page.waitForFunction(
      () => !window.__webSheetDevCellModel?.("CHT", 6, 22)?.raw &&
        !window.__webSheetDevCellModel?.("CHT", 6, 25)?.formula,
      null,
      { timeout: 10000 },
    );
    const afterUndo = await page.evaluate(() => ({
      v6: window.__webSheetDevCellModel?.("CHT", 6, 22) || null,
      y6: window.__webSheetDevCellModel?.("CHT", 6, 25) || null,
      aa6: window.__webSheetDevCellModel?.("CHT", 6, 27) || null,
      ac6: window.__webSheetDevCellModel?.("CHT", 6, 29) || null,
      b9: window.__webSheetDevCellModel?.("CHT", 9, 2) || null,
      ad6: window.__webSheetDevCellModel?.("CHT", 6, 30) || null,
      ao6: window.__webSheetDevCellModel?.("CHT", 6, 41) || null,
      ap6: window.__webSheetDevCellModel?.("CHT", 6, 42) || null,
      ax6: window.__webSheetDevCellModel?.("CHT", 6, 50) || null,
      d9: window.__webSheetDevCellModel?.("CHT", 9, 4) || null,
      history: window.__webSheetDevHistoryState?.(),
    }));
    await page.keyboard.press(redoShortcut);
    await page.waitForFunction(
      () => window.__webSheetDevCellModel?.("CHT", 6, 22)?.raw === "CellsValue" &&
        window.__webSheetDevCellModel?.("CHT", 6, 25)?.formula === "=J6+1",
      null,
      { timeout: 10000 },
    );
    const afterRedo = await page.evaluate(() => ({
      v6: window.__webSheetDevCellModel?.("CHT", 6, 22) || null,
      y6: window.__webSheetDevCellModel?.("CHT", 6, 25) || null,
      aa6: window.__webSheetDevCellModel?.("CHT", 6, 27) || null,
      ac6: window.__webSheetDevCellModel?.("CHT", 6, 29) || null,
      b11: window.__webSheetDevCellModel?.("CHT", 11, 2) || null,
      ak6: window.__webSheetDevCellModel?.("CHT", 6, 37) || null,
      ao6: window.__webSheetDevCellModel?.("CHT", 6, 41) || null,
      ap6: window.__webSheetDevCellModel?.("CHT", 6, 42) || null,
      ax6: window.__webSheetDevCellModel?.("CHT", 6, 50) || null,
      d11: window.__webSheetDevCellModel?.("CHT", 11, 4) || null,
      history: window.__webSheetDevHistoryState?.(),
    }));
    await context.close();

    const failures = [];
    if (messages.length) failures.push(`browser errors: ${messages.join(" | ")}`);
    if (runResult?.unsupported?.length) failures.push(`unsupported: ${runResult.unsupported.join(" | ")}`);
    if (runResult?.writeCount < 20) failures.push(`write count too low: ${runResult?.writeCount}`);
    if (afterRun.v6?.raw !== "CellsValue") failures.push("Cells(row,col).Value");
    if (afterRun.w6?.raw !== "Array1" || afterRun.x6?.raw !== "Array2") failures.push("Range(Cells,Cells).Value Array");
    if (afterRun.y6?.formula !== "=J6+1") failures.push(`FormulaR1C1 conversion: ${afterRun.y6?.formula}`);
    if (Number(afterRun.z6?.raw) !== 50) failures.push(`WorksheetFunction.Sum: ${afterRun.z6?.raw}`);
    if (Number(afterRun.ab6?.raw) !== 44) failures.push(`Application.WorksheetFunction.Average: ${afterRun.ab6?.raw}`);
    if (Number(afterRun.ac6?.raw) !== 52) failures.push(`VBA variable arithmetic: ${afterRun.ac6?.raw}`);
    if (Number(afterRun.b9?.raw) !== 2 || Number(afterRun.b10?.raw) !== 4 || Number(afterRun.b11?.raw) !== 6) failures.push("For/Next loop arithmetic");
    if (afterRun.ad6?.raw !== "IfTrue") failures.push(`If/Else block: ${afterRun.ad6?.raw}`);
    if (afterRun.ae6?.raw !== "ElseIfTrue") failures.push(`ElseIf block: ${afterRun.ae6?.raw}`);
    if (afterRun.af6?.raw !== "InlineIf") failures.push(`Inline If/Not comparison: ${afterRun.af6?.raw}`);
    if (Number(afterRun.ag6?.raw) !== 2) failures.push(`WorksheetFunction.CountIf: ${afterRun.ag6?.raw}`);
    if (Number(afterRun.ah6?.raw) !== 10) failures.push(`WorksheetFunction.SumIf: ${afterRun.ah6?.raw}`);
    if (Number(afterRun.ai6?.raw) !== 2.35) failures.push(`WorksheetFunction.Round: ${afterRun.ai6?.raw}`);
    if (afterRun.aj6?.raw !== "ABDE") failures.push(`WorksheetFunction.Left/Right concat: ${afterRun.aj6?.raw}`);
    if (afterRun.ak6?.raw !== "BCD") failures.push(`WorksheetFunction.Mid: ${afterRun.ak6?.raw}`);
    if (Number(afterRun.al6?.raw) !== 6) failures.push(`Do While / Exit Do: ${afterRun.al6?.raw}`);
    if (Number(afterRun.am6?.raw) !== 3) failures.push(`Do Loop Until: ${afterRun.am6?.raw}`);
    if (Number(afterRun.an6?.raw) !== 6) failures.push(`Exit For / Mod inline If: ${afterRun.an6?.raw}`);
    if (afterRun.ao6?.raw !== "CaseHigh") failures.push(`Select Case / Case Is: ${afterRun.ao6?.raw}`);
    if (afterRun.ap6?.raw !== "Mid") failures.push(`WorksheetFunction.VLookup: ${afterRun.ap6?.raw}`);
    if (Number(afterRun.aq6?.raw) !== 3) failures.push(`WorksheetFunction.Match: ${afterRun.aq6?.raw}`);
    if (afterRun.ar6?.raw !== "High") failures.push(`WorksheetFunction.Index: ${afterRun.ar6?.raw}`);
    if (Number(afterRun.as6?.raw) !== 1) failures.push(`WorksheetFunction.CountIfs wildcard: ${afterRun.as6?.raw}`);
    if (Number(afterRun.at6?.raw) !== 6) failures.push(`WorksheetFunction.SumIfs: ${afterRun.at6?.raw}`);
    if (Number(afterRun.au6?.raw) !== 5) failures.push(`WorksheetFunction.AverageIf: ${afterRun.au6?.raw}`);
    if (afterRun.av6?.raw !== "0:2:BB") failures.push(`VBA Array/LBound/UBound/index: ${afterRun.av6?.raw}`);
    if (afterRun.aw6?.raw !== "3:A/B/C") failures.push(`VBA InStr/Replace: ${afterRun.aw6?.raw}`);
    if (Number(afterRun.ax6?.raw) !== 12) failures.push(`For Each Range variable: ${afterRun.ax6?.raw}`);
    if (Number(afterRun.d9?.raw) !== 9 || Number(afterRun.d10?.raw) !== 10 || Number(afterRun.d11?.raw) !== 11) failures.push("For Each Range Offset/Row");
    if (afterRun.ay6?.raw !== "3:3:$B$9") failures.push(`Range Rows.Count/Count/Address: ${afterRun.ay6?.raw}`);
    if (afterRun.aa6?.raw !== "Indexed") failures.push("styled target raw value");
    if (afterRun.aa6?.hasFill !== true || !afterRun.aa6?.css?.backgroundColor) failures.push("Interior.ColorIndex");
    if (afterRun.aa6?.hasFontColor !== true || !afterRun.aa6?.css?.color) failures.push("Font.ColorIndex");
    if (afterRun.aa6?.css?.textAlign !== "center") failures.push("xlHAlignCenter");
    if (afterRun.aa6?.css?.verticalAlign !== "bottom") failures.push("xlVAlignBottom");
    if (!String(afterRun.aa6?.css?.textDecoration || "").includes("underline")) failures.push("xlUnderlineStyleDouble");
    if (afterUndo.v6?.raw || afterUndo.y6?.formula || afterUndo.aa6?.raw || afterUndo.ac6?.raw || afterUndo.b9?.raw || afterUndo.ad6?.raw || afterUndo.ao6?.raw || afterUndo.ap6?.raw || afterUndo.ax6?.raw || afterUndo.d9?.raw) failures.push("macro undo");
    if (afterRedo.v6?.raw !== "CellsValue" || afterRedo.y6?.formula !== "=J6+1" || afterRedo.aa6?.raw !== "Indexed" || Number(afterRedo.ac6?.raw) !== 52 || Number(afterRedo.b11?.raw) !== 6 || afterRedo.ak6?.raw !== "BCD" || afterRedo.ao6?.raw !== "CaseHigh" || afterRedo.ap6?.raw !== "Mid" || Number(afterRedo.ax6?.raw) !== 12 || Number(afterRedo.d11?.raw) !== 11) failures.push("macro redo");
    if (!afterRun.history?.undo?.some((label) => label.includes("マクロ実行 MacroVbaSubsetSmoke"))) failures.push("macro history label");
    if (failures.length) {
      throw new Error(`Macro VBA subset smoke failed: ${failures.join(", ")} ${JSON.stringify({ runResult, afterRun, afterUndo, afterRedo })}`);
    }
    console.log(JSON.stringify({ ok: true, macroVbaSubset: { runResult, afterRun, afterUndo, afterRedo } }, null, 2));
  } finally {
    await browser.close();
  }
}



function assertValidationMenuResult(result) {
  if (
    !result?.labels?.some((label) => label.includes("データの入力規則(V)...")) ||
    !result?.marked?.status?.includes("2 個") ||
    !result?.afterClearAllMark?.status?.includes("1 個") ||
    result?.importedHiddenDropdown?.validation?.showDropDown !== false ||
    result?.importedHiddenDropdown?.validation?.imeMode !== "disabled" ||
    result?.exportedHiddenDropdown?.showDropDown !== true ||
    result?.exportedHiddenDropdown?.imeMode !== true ||
    result?.importedBlankRangeValidation?.c30?.options?.join("|") !== "One|Two" ||
    result?.importedBlankRangeValidation?.e30?.formulae?.[0] !== "=A12" ||
    result?.importedBlankRangeValidation?.f30?.formulae?.[0] !== "=LEN(F30)<=3" ||
    result?.importedBlankRangeValidation?.h30?.options?.join("|") !== "East|West" ||
    result?.importedBlankRangeValidation?.h32?.options?.join("|") !== "East|West" ||
    result?.importedBlankRangeValidation?.i33?.options?.join("|") !== "East|West" ||
    result?.importedBlankRangeValidation?.j30?.type !== "decimal" ||
    result?.importedBlankRangeValidation?.j30?.operator !== "greaterThan" ||
    result?.importedBlankRangeValidation?.j30?.formulae?.[0] !== "1.5" ||
    result?.importedBlankRangeValidation?.j30?.showInputMessage !== true ||
    result?.importedBlankRangeValidation?.j30?.inputTitle !== "小数案内" ||
    result?.importedBlankRangeValidation?.j30?.errorStyle !== "warning" ||
    result?.importedBlankRangeValidation?.k30?.type !== "date" ||
    result?.importedBlankRangeValidation?.k30?.formulae?.join("|") !== "45413|45443" ||
    result?.importedBlankRangeValidation?.l30?.type !== "time" ||
    result?.importedBlankRangeValidation?.l30?.formulae?.join("|") !== "0.375|0.75" ||
    result?.importedBlankRangeValidation?.m30?.type !== "textLength" ||
    result?.importedBlankRangeValidation?.m30?.formulae?.[0] !== "5" ||
    !result?.exportedBlankRangeValidation?.multiRangeTag?.includes('sqref="H30 H32:I33"') ||
    !result?.exportedBlankRangeValidation?.decimalTag?.includes('promptTitle="小数案内"') ||
    !result?.exportedBlankRangeValidation?.decimalTag?.includes('errorStyle="warning"') ||
    !result?.exportedBlankRangeValidation?.dateTag ||
    !result?.exportedBlankRangeValidation?.timeTag ||
    !result?.exportedBlankRangeValidation?.textLengthTag
  ) {
    throw new Error(`Validation menu smoke failed: ${JSON.stringify(result)}`);
  }
}



async function runValidationMenuOnly() {
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
    const result = await exerciseDataValidationMenu(context);
    await context.close();
    assertValidationMenuResult(result);
    console.log(JSON.stringify({ ok: true, dataValidationMenu: result }, null, 2));
  } finally {
    await browser.close();
  }
}



async function runOutlineGroupingOnly() {
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
    const result = await exerciseOutlineGrouping(context);
    await context.close();
    console.log(JSON.stringify({ ok: true, outlineGrouping: result }, null, 2));
  } finally {
    await browser.close();
  }
}



function assertAutoFilterDateSmoke(autoFilterDateMenu) {
  const rowsMatch = (rows, expectedHiddenByRow) =>
    [2, 3, 4, 5, 6].every((row) => rows?.find((item) => item.row === row)?.hidden === expectedHiddenByRow.get(row));
  const expectedHiddenRows = (predicate) => new Map([
    [2, new Date(2026, 4, 20)],
    [3, new Date(2026, 4, 20)],
    [4, new Date(2026, 4, 23)],
    [5, new Date(2026, 5, 1)],
    [6, new Date(2025, 11, 31)],
  ].map(([row, date]) => [row, !predicate(date)]));
  const referenceDate = new Date();
  const referenceStartOfDay = new Date(referenceDate.getFullYear(), referenceDate.getMonth(), referenceDate.getDate());
  const nextMonthReference = new Date(referenceDate.getFullYear(), referenceDate.getMonth() + 1, 1);
  const thisMonthExpected = expectedHiddenRows((date) => date.getFullYear() === referenceDate.getFullYear() && date.getMonth() === referenceDate.getMonth());
  const nextMonthExpected = expectedHiddenRows((date) => date.getFullYear() === nextMonthReference.getFullYear() && date.getMonth() === nextMonthReference.getMonth());
  const thisYearExpected = expectedHiddenRows((date) => date.getFullYear() === referenceDate.getFullYear());
  const yearToDateExpected = expectedHiddenRows((date) =>
    date.getFullYear() === referenceDate.getFullYear() &&
    date >= new Date(referenceDate.getFullYear(), 0, 1) &&
    date <= referenceStartOfDay);
  const allMayExpected = expectedHiddenRows((date) => date.getMonth() === 4);
  const secondQuarterExpected = expectedHiddenRows((date) => date.getMonth() >= 3 && date.getMonth() <= 5);
  const failures = [];
  if (!autoFilterDateMenu.menu?.includes("日付フィルター")) failures.push("date menu label");
  if (!autoFilterDateMenu.submenu?.includes("指定の日付より後")) failures.push("date comparison command");
  if (!autoFilterDateMenu.submenu?.includes("今月")) failures.push("date preset command");
  if (!autoFilterDateMenu.submenu?.includes("第2四半期") || !autoFilterDateMenu.submenu?.includes("5月")) failures.push("date period commands");
  if (autoFilterDateMenu.valueLabels?.join("|") !== "2025/12/31|May 20, 2026|05/23/2026|06/01/2026") failures.push("date value grouping labels");
  if (!rowsMatch(autoFilterDateMenu.thisMonthRows, thisMonthExpected)) failures.push("this month rows");
  if (!rowsMatch(autoFilterDateMenu.nextMonthRows, nextMonthExpected)) failures.push("next month rows");
  if (!rowsMatch(autoFilterDateMenu.thisYearRows, thisYearExpected)) failures.push("this year rows");
  if (!rowsMatch(autoFilterDateMenu.yearToDateRows, yearToDateExpected)) failures.push("year-to-date rows");
  if (!rowsMatch(autoFilterDateMenu.allMayRows, allMayExpected)) failures.push("all May period rows");
  if (!rowsMatch(autoFilterDateMenu.secondQuarterRows, secondQuarterExpected)) failures.push("second quarter period rows");
  if (!autoFilterDateMenu.periodExported?.dynamicFilters?.some((item) => item.type === "Q2")) failures.push("date period dynamicFilter export");
  if (failures.length) throw new Error(`AutoFilter date smoke failed: ${failures.join(", ")}`);
}



async function runAutoFilterDateOnly() {
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
    const autoFilterDateMenu = await exerciseAutoFilterDateMenu(context);
    await context.close();
    assertAutoFilterDateSmoke(autoFilterDateMenu);
    console.log(JSON.stringify({ ok: true, autoFilterDateMenu }, null, 2));
  } finally {
    await browser.close();
  }
}



function assertDataFilterSortSmoke(result) {
  const failures = [];
  const filter = result.autoFilterReapplyClear;
  const imported = result.importedAutoFilterCriteria;
  const currentRegion = result.dataCurrentRegion;
  const sortDialog = result.sortDialogLevels;
  const sortColor = result.sortDialogColorOrder;
  const quickSort = result.quickSortActiveColumn;
  const partialSort = result.quickSortPartialSelection;
  const colorImportExport = result.autoFilterColorImportExport;
  const autoFilterConditionMenu = result.autoFilterConditionMenu;

  if (filter.sortHeaderState?.sorted !== true) failures.push("filter header shows sorted state");
  if (filter.sortHeaderState?.desc !== true || filter.sortHeaderState?.direction !== "desc") failures.push("filter header desc direction");
  if (filter.sortHeaderState?.sortState?.keys?.[0]?.col !== 2) failures.push("filter header sort key column");
  if (filter.sortMenuState?.direction !== "desc") failures.push("filter menu sort direction");
  if (filter.sortMenuState?.descPressed !== "true") failures.push("filter menu desc button active");
  if (!filter.sortMenuState?.current?.includes("降順")) failures.push("filter menu current sort description");
  if (filter.sortClear?.header?.sorted !== false || filter.sortClear?.header?.direction) failures.push("clear sort removes header state");
  if (!filter.sortAfterClearReapply?.status?.includes("フィルターを再適用")) failures.push("reapply after clear keeps filter behavior");

  const savedSort = imported.savedSort;
  if (savedSort?.buttonSorted !== true || savedSort?.buttonSortDirection !== "desc") failures.push("imported saved sort button state");
  if (savedSort?.sortState?.keys?.[0]?.col !== 2 || savedSort?.sortState?.keys?.[0]?.direction !== "desc") failures.push("imported saved sort model state");
  if (imported.standaloneSort?.hasFilterButton !== false) failures.push("standalone sort does not create filter button");
  const tableSortOnly = imported.tableSortOnly;
  if (tableSortOnly?.before?.buttonSorted !== true || tableSortOnly?.before?.buttonSortDirection !== "desc") failures.push("imported table sort button state");
  if (tableSortOnly?.before?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "C:20|B:12|D:8|A:5") failures.push("imported table sort reorders rows");
  if (tableSortOnly?.ribbonBeforeClear?.sheetSortState) failures.push("table-only sort smoke clears sheet sort before ribbon check");
  if (tableSortOnly?.ribbonBeforeClear?.tableSortState?.keys?.[0]?.direction !== "desc") failures.push("table-only sort model state");
  if (tableSortOnly?.ribbonBeforeClear?.clear?.disabled !== false || tableSortOnly?.ribbonBeforeClear?.reapply?.disabled !== false) failures.push("table-only sort enables clear and reapply");
  if (tableSortOnly?.exportedBeforeClear?.excelJsLoaded !== true) failures.push("table-only sort exported xlsx reloads");
  if (tableSortOnly?.exportedBeforeClear?.tableXmlHasSortState) failures.push("table-only sort export avoids table sortState xml");
  if (!tableSortOnly?.exportedBeforeClear?.worksheetSortStateCount) failures.push("table-only sort export writes worksheet sortState");
  if (tableSortOnly?.afterClear?.buttonSorted !== false || tableSortOnly?.afterClear?.buttonSortDirection) failures.push("table-only sort clear removes button state");
  if (tableSortOnly?.afterClear?.tableSortState) failures.push("table-only sort clear removes table sort state");
  if (imported.importedCellColorFilter?.criteria?.kind !== "color" || imported.importedCellColorFilter?.criteria?.target !== "fill") failures.push("imported worksheet cell color filter restores criteria");
  if (imported.importedCellColorFilter?.buttonFiltered !== true) failures.push("imported worksheet cell color filter shows filtered button");
  if (imported.importedCellColorFilter?.rows?.find((row) => row.row === 4)?.hidden !== false) failures.push("imported worksheet cell color filter keeps matching row visible");
  if (imported.importedCellColorFilter?.rows?.filter((row) => row.row !== 4).some((row) => row.hidden !== true)) failures.push("imported worksheet cell color filter hides non-matching rows");
  if (imported.importedFontColorFilter?.criteria?.kind !== "color" || imported.importedFontColorFilter?.criteria?.target !== "font") failures.push("imported worksheet font color filter restores criteria");
  if (imported.importedFontColorFilter?.buttonFiltered !== true) failures.push("imported worksheet font color filter shows filtered button");
  if (imported.importedFontColorFilter?.rows?.find((row) => row.row === 3)?.hidden !== false) failures.push("imported worksheet font color filter keeps matching row visible");
  if (imported.importedFontColorFilter?.rows?.filter((row) => row.row !== 3).some((row) => row.hidden !== true)) failures.push("imported worksheet font color filter hides non-matching rows");
  if (imported.importedTableColorFilter?.criteria?.kind !== "color" || imported.importedTableColorFilter?.criteria?.target !== "fill") failures.push("imported table cell color filter restores criteria");
  if (imported.importedTableColorFilter?.buttonFiltered !== true) failures.push("imported table cell color filter shows filtered button");
  if (imported.importedTableColorFilter?.rows?.find((row) => row.row === 4)?.hidden !== false) failures.push("imported table cell color filter keeps matching row visible");
  if (imported.importedTableColorFilter?.rows?.filter((row) => row.row !== 4).some((row) => row.hidden !== true)) failures.push("imported table cell color filter hides non-matching rows");
  if (imported.importedCellColorSort?.buttonSorted !== true || imported.importedCellColorSort?.buttonSortDirection !== "asc") failures.push("imported worksheet cell color sort button state");
  if (imported.importedCellColorSort?.sortState?.keys?.[0]?.sortOn !== "cellColor" || imported.importedCellColorSort?.sortState?.keys?.[0]?.color !== "#ff0000") failures.push("imported worksheet cell color sort model state");
  if (imported.importedCellColorSort?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "B:12|A:5|C:20|D:8") failures.push("imported worksheet cell color sort reorders rows");
  if (imported.importedFontColorSort?.buttonSorted !== true || imported.importedFontColorSort?.buttonSortDirection !== "asc") failures.push("imported worksheet font color sort button state");
  if (imported.importedFontColorSort?.sortState?.keys?.[0]?.sortOn !== "fontColor" || imported.importedFontColorSort?.sortState?.keys?.[0]?.color !== "#ff0000") failures.push("imported worksheet font color sort model state");
  if (imported.importedFontColorSort?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "D:8|A:5|B:12|C:20") failures.push("imported worksheet font color sort reorders rows");
  if (imported.importedCustomListSort?.buttonSorted !== true || imported.importedCustomListSort?.buttonSortDirection !== "asc") failures.push("imported worksheet custom-list sort button state");
  if (imported.importedCustomListSort?.sortState?.keys?.[0]?.customList !== "month-ja") failures.push("imported worksheet custom-list sort model state");
  if (imported.importedCustomListSort?.rows?.map((row) => `${row.name}:${row.score}`).join("|") !== "1月:M1|2月:M2|3月:M3|12月:M12") failures.push("imported worksheet custom-list sort reorders rows");
  if (imported.importedTextContainsFilter?.criteria?.operator1 !== "contains" || imported.importedTextContainsFilter?.criteria?.value1 !== "ta") failures.push("imported text contains wildcard filter restores criteria");
  if (imported.importedTextContainsFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "alpha:hidden|Alpine:hidden|Beta:visible|delta:visible") failures.push("imported text contains wildcard filter applies row visibility");
  if (imported.importedTextBeginsFilter?.criteria?.operator1 !== "beginsWith" || imported.importedTextBeginsFilter?.criteria?.value1 !== "Al") failures.push("imported text begins wildcard filter restores criteria");
  if (imported.importedTextBeginsFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "alpha:visible|Alpine:visible|Beta:hidden|delta:hidden") failures.push("imported text begins wildcard filter applies row visibility");
  if (imported.importedTextNotEndsFilter?.criteria?.operator1 !== "notEndsWith" || imported.importedTextNotEndsFilter?.criteria?.value1 !== "ta") failures.push("imported text not-ends wildcard filter restores criteria");
  if (imported.importedTextNotEndsFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "alpha:visible|Alpine:visible|Beta:hidden|delta:hidden") failures.push("imported text not-ends wildcard filter applies row visibility");
  if (imported.importedLiteralWildcardFilter?.criteria?.operator1 !== "equal" || imported.importedLiteralWildcardFilter?.criteria?.value1 !== "A~*") failures.push("imported literal wildcard filter keeps escaped pattern criteria");
  if (imported.importedLiteralWildcardFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "A*:visible|A?:hidden|Alpha:hidden|Star:hidden") failures.push("imported literal wildcard filter treats escaped star as text");
  if (imported.importedMultiDateGroupFilter?.criteria?.kind !== "dateGroups" || imported.importedMultiDateGroupFilter?.criteria?.dateGroups?.length !== 2) failures.push("imported multiple date group filter restores criteria");
  if (imported.importedMultiDateGroupFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "2026/05/20:visible|2026/06/01:visible|2025/05/10:hidden|2026/07/01:hidden") failures.push("imported multiple date group filter applies row visibility");
  if ((imported.importedDateGroupExported?.dateGroupItemCount || 0) < 3) failures.push("date group filter export preserves dateGroupItem xml");
  if (imported.importedBetweenFilter?.criteria?.logic !== "and" || imported.importedBetweenFilter?.criteria?.operator1 !== "greaterThanOrEqual" || imported.importedBetweenFilter?.criteria?.operator2 !== "lessThanOrEqual") failures.push("imported between filter restores AND criteria");
  if (imported.importedBetweenFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "A:hidden|B:visible|C:visible|D:hidden") failures.push("imported between filter applies row visibility");
  if (imported.importedNotBetweenFilter?.criteria?.logic !== "or" || imported.importedNotBetweenFilter?.criteria?.operator1 !== "lessThan" || imported.importedNotBetweenFilter?.criteria?.operator2 !== "greaterThan") failures.push("imported not-between filter restores OR criteria");
  if (imported.importedNotBetweenFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "A:visible|B:hidden|C:hidden|D:visible") failures.push("imported not-between filter applies row visibility");
  if (imported.importedMonthPeriodFilter?.criteria?.datePeriod !== "M5") failures.push("imported month period dynamic filter restores criteria");
  if (imported.importedMonthPeriodFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "2026/05/20:visible|2025/05/10:visible|2026/06/01:hidden|2026/04/30:hidden") failures.push("imported month period dynamic filter applies row visibility");
  if (imported.importedQuarterPeriodFilter?.criteria?.datePeriod !== "Q2") failures.push("imported quarter period dynamic filter restores criteria");
  if (imported.importedQuarterPeriodFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "2026/04/01:visible|2025/05/10:visible|2026/07/01:hidden|2026/03/31:hidden") failures.push("imported quarter period dynamic filter applies row visibility");
  if (imported.importedAboveStdDevFilter?.criteria?.stdDevOperator !== "aboveStdDev" || imported.importedAboveStdDevFilter?.criteria?.stdDevCount !== 1) failures.push("imported above standard deviation dynamic filter restores criteria");
  if (imported.importedAboveStdDevFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "A:hidden|B:hidden|C:hidden|D:visible") failures.push("imported above standard deviation dynamic filter applies row visibility");
  if (imported.importedBelowStdDevFilter?.criteria?.stdDevOperator !== "belowStdDev" || imported.importedBelowStdDevFilter?.criteria?.stdDevCount !== 1) failures.push("imported below standard deviation dynamic filter restores criteria");
  if (imported.importedBelowStdDevFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "A:visible|B:hidden|C:hidden|D:hidden") failures.push("imported below standard deviation dynamic filter applies row visibility");
  if (imported.importedHiddenButtonFilter?.hasNameButton !== true || imported.importedHiddenButtonFilter?.hasScoreButton !== false) failures.push("imported hidden auto filter button controls header button visibility");
  if (imported.importedHiddenButtonFilter?.hiddenButtons?.["2"] !== true) failures.push("imported hidden auto filter button restores hiddenButtons model");
  if (imported.importedHiddenButtonFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "Alpha:visible|Beta:hidden|Alpha:visible|Gamma:hidden") failures.push("imported hidden auto filter button sheet keeps active criteria");
  if (imported.importedDateHourGroupFilter?.criteria?.dateGroup?.grouping !== "hour" || imported.importedDateHourGroupFilter?.criteria?.dateGroup?.hour !== 13) failures.push("imported date hour group filter restores hour criteria");
  if (imported.importedDateHourGroupFilter?.rows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "2026/05/20 13:10:visible|2026/05/20 14:10:hidden|2026/05/21 13:10:hidden|2026/05/20 13:55:visible") failures.push("imported date hour group filter applies time row visibility");
  if (imported.importedCustomNonBlankFilter?.criteria?.operator1 !== "nonBlank") failures.push("imported custom not-equal blank restores nonBlank operator");
  if (imported.importedCustomNonBlankFilter?.rows?.map((row) => `${row.name || "(空白)"}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "Filled:visible|(空白):hidden|Zero:visible|(空白):hidden") failures.push("imported custom not-equal blank applies nonBlank row visibility");
  if (!imported.importedDateGroupExported?.dynamicFilters?.some((item) => item.type === "M5") || !imported.importedDateGroupExported?.dynamicFilters?.some((item) => item.type === "Q2")) failures.push("dynamic date period filters export as Excel dynamicFilter types");
  if (!imported.importedDateGroupExported?.dynamicFilters?.some((item) => item.type === "aboveStdDev" && item.val === "1") || !imported.importedDateGroupExported?.dynamicFilters?.some((item) => item.type === "belowStdDev" && item.val === "1")) failures.push("dynamic standard deviation filters export with Excel val attribute");
  if (!imported.importedDateGroupExported?.filterColumns?.some((item) => item.colId === "1" && item.hiddenButton === "1" && item.showButton === "0")) failures.push("hidden auto filter button exports hiddenButton/showButton attributes");
  if (!imported.importedDateGroupExported?.dateGroupItems?.some((item) => item.grouping === "hour" && item.hour === "13")) failures.push("date hour group filter exports hour dateGroupItem xml");
  if (!imported.importedCustomFilterExported?.customFilters?.some((item) => item.operator === "notEqual" && item.value === "")) failures.push("custom nonBlank filter exports Excel notEqual blank customFilter");

  if (currentRegion.chtHeaderFilter?.modelRange?.top !== 15) failures.push("CHT row header filter starts on selected header row");
  if (currentRegion.chtHeaderFilter?.modelRange?.left !== 2 || currentRegion.chtHeaderFilter?.modelRange?.right !== 42) failures.push("CHT row header filter expands through contiguous header cells");
  if (currentRegion.tablePriorityFilter?.range?.left !== 2 || currentRegion.tablePriorityFilter?.range?.right !== 4) failures.push("table range has priority for auto filter");
  if (currentRegion.toggleFilter?.hasAutoFilter !== false || currentRegion.toggleFilter?.buttonCount !== 0) failures.push("filter toggle removes auto filter");

  if (sortDialog.reopened?.levelCount !== 2) failures.push("sort dialog keeps multiple levels");
  if (sortDialog.reopened?.orders?.join("|") !== "asc|desc") failures.push("sort dialog keeps level orders");
  if (!sortColor?.colorOptions?.includes("#ff0000")) failures.push("sort dialog offers cell color order choice");
  if (sortColor?.rows?.[0]?.score !== "8") failures.push("sort dialog cell color order sorts matching color first");
  if (!sortColor?.colorExported?.cellColorSortCount || !sortColor?.colorExported?.dxfSortCount) failures.push("sort dialog cell color order exports dxf sort condition");
  if (!sortColor?.fontColorOptions?.includes("#ff0000")) failures.push("sort dialog offers font color order choice");
  if (sortColor?.fontRows?.[0]?.score !== "5") failures.push("sort dialog font color order sorts matching color first");
  if (!sortColor?.fontExported?.fontColorSortCount || !sortColor?.fontExported?.dxfSortCount) failures.push("sort dialog font color order exports dxf sort condition");
  if (!result.sortDialogCustomList?.months?.exported?.customListSortCount) failures.push("sort dialog custom list order exports customList sort condition");
  if (result.sortDialogCustomList?.months?.exported?.customListTexts?.[0] !== "1月,2月,3月,4月,5月,6月,7月,8月,9月,10月,11月,12月") failures.push("sort dialog custom list order exports Japanese month list");
  if (!quickSort.status?.includes("並べ替えました")) failures.push("quick sort active column sorts current region");
  if (!partialSort.status?.includes("並べ替えました")) failures.push("quick sort partial selection expands current region");
  if (!colorImportExport?.colorChoice) failures.push("color filter chooses a source color");
  if (colorImportExport?.beforeExport?.criteria?.kind !== "color") failures.push("color filter criteria is captured before export");
  if (colorImportExport?.beforeExport?.rows?.find((row) => row.row === 4)?.hidden !== false) failures.push("color filter keeps matching row visible before export");
  if (colorImportExport?.beforeExport?.rows?.filter((row) => row.row !== 4).some((row) => row.hidden !== true)) failures.push("color filter hides non-matching rows before export");
  if (colorImportExport?.exported?.excelJsLoaded !== true) failures.push(`color filter exported xlsx reloads in ExcelJS${colorImportExport?.exported?.excelJsError ? `: ${colorImportExport.exported.excelJsError}` : ""}`);
  if (!colorImportExport?.exported?.colorFilterCount) failures.push("color filter export writes colorFilter xml");
  if (!colorImportExport?.exported?.fillColorFilterCount) failures.push("color filter export marks cell-color filter");
  if (!colorImportExport?.exported?.dxfCount || !colorImportExport?.exported?.dxfRgbColors?.length) failures.push("color filter export writes dxf style color");
  if (colorImportExport?.reloaded?.buttonFiltered !== true) failures.push("color filter reload shows filtered header");
  if (colorImportExport?.reloaded?.criteria?.kind !== "color") failures.push("color filter reload restores criteria");
  if (colorImportExport?.reloaded?.rows?.find((row) => row.row === 4)?.hidden !== false) failures.push("color filter reload keeps matching row visible");
  if (colorImportExport?.reloaded?.rows?.filter((row) => row.row !== 4).some((row) => row.hidden !== true)) failures.push("color filter reload hides non-matching rows");
  if (!colorImportExport?.fontColorChoice) failures.push("font color filter chooses a source color");
  if (colorImportExport?.fontBeforeExport?.criteria?.kind !== "color" || colorImportExport?.fontBeforeExport?.criteria?.target !== "font") failures.push("font color filter criteria is captured before export");
  if (colorImportExport?.fontBeforeExport?.rows?.find((row) => row.row === 3)?.hidden !== false) failures.push("font color filter keeps matching row visible before export");
  if (colorImportExport?.fontBeforeExport?.rows?.filter((row) => row.row !== 3).some((row) => row.hidden !== true)) failures.push("font color filter hides non-matching rows before export");
  if (colorImportExport?.fontExported?.excelJsLoaded !== true) failures.push(`font color filter exported xlsx reloads in ExcelJS${colorImportExport?.fontExported?.excelJsError ? `: ${colorImportExport.fontExported.excelJsError}` : ""}`);
  if (!colorImportExport?.fontExported?.fontColorFilterCount) failures.push("font color filter export marks font-color filter");
  if (!colorImportExport?.fontExported?.dxfCount || !colorImportExport?.fontExported?.dxfRgbColors?.length) failures.push("font color filter export writes dxf style color");
  if (colorImportExport?.fontReloaded?.buttonFiltered !== true) failures.push("font color filter reload shows filtered header");
  if (colorImportExport?.fontReloaded?.criteria?.kind !== "color" || colorImportExport?.fontReloaded?.criteria?.target !== "font") failures.push("font color filter reload restores criteria");
  if (colorImportExport?.fontReloaded?.rows?.find((row) => row.row === 3)?.hidden !== false) failures.push("font color filter reload keeps matching row visible");
  if (colorImportExport?.fontReloaded?.rows?.filter((row) => row.row !== 3).some((row) => row.hidden !== true)) failures.push("font color filter reload hides non-matching rows");
  if (autoFilterConditionMenu?.textCaseExported?.customFilters?.[0]?.operator !== "equal" || autoFilterConditionMenu?.textCaseExported?.customFilters?.[0]?.value !== "AL*") failures.push("text begins filter exports Excel wildcard pattern");
  if (autoFilterConditionMenu?.textWildcardExported?.customFilters?.[0]?.operator !== "equal" || autoFilterConditionMenu?.textWildcardExported?.customFilters?.[0]?.value !== "d*") failures.push("text equal wildcard filter exports raw wildcard pattern");
  if (autoFilterConditionMenu?.textNotContainsRows?.map((row) => `${row.name}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "alpha:visible|Alpine:visible|Beta:hidden|delta:hidden") failures.push("text not-contains menu applies row visibility");
  if (autoFilterConditionMenu?.textNotContainsExported?.customFilters?.[0]?.operator !== "notEqual" || autoFilterConditionMenu?.textNotContainsExported?.customFilters?.[0]?.value !== "*ta*") failures.push("text not-contains filter exports Excel wildcard pattern");
  if (!autoFilterConditionMenu?.numberSubmenu?.includes("指定の範囲外")) failures.push("number filter menu shows not-between command");
  if (!autoFilterConditionMenu?.numberSubmenu?.includes("標準偏差より上") || !autoFilterConditionMenu?.numberSubmenu?.includes("標準偏差より下")) failures.push("number filter menu shows standard deviation commands");
  if (autoFilterConditionMenu?.notBetweenRows?.map((row) => `${row.score}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "5:visible|12:hidden|20:hidden|8:visible|ZZZ:hidden") failures.push("number not-between menu applies OR range criteria");
  if (autoFilterConditionMenu?.notBetweenExported?.customFilters?.map((item) => `${item.operator}:${item.value}`).join("|") !== "lessThan:10|greaterThan:20") failures.push("number not-between filter exports OR customFilters");
  if (autoFilterConditionMenu?.aboveStdDevRows?.map((row) => `${row.score}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "5:hidden|12:hidden|20:visible|8:hidden|ZZZ:hidden") failures.push("number above standard deviation menu applies dynamic filter");
  if (!autoFilterConditionMenu?.aboveStdDevExported?.dynamicFilters?.some((item) => item.type === "aboveStdDev" && item.val === "1")) failures.push("number above standard deviation menu exports dynamicFilter");
  if (autoFilterConditionMenu?.belowStdDevRows?.map((row) => `${row.score}:${row.hidden ? "hidden" : "visible"}`).join("|") !== "5:hidden|12:hidden|20:hidden|8:hidden|ZZZ:hidden") failures.push("number below standard deviation menu applies dynamic filter");
  if (!autoFilterConditionMenu?.belowStdDevExported?.dynamicFilters?.some((item) => item.type === "belowStdDev" && item.val === "1")) failures.push("number below standard deviation menu exports dynamicFilter");

  if (failures.length) {
    throw new Error(`Data/filter/sort smoke failed: ${failures.join(", ")}`);
  }
}



async function runDataFilterSortOnly() {
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
    const autoFilterReapplyClear = await exerciseAutoFilterReapplyAndClear(context);
    const importedAutoFilterCriteria = await exerciseImportedAutoFilterCriteria(context);
    const autoFilterConditionMenu = await exerciseAutoFilterConditionMenu(context);
    const dataCurrentRegion = await exerciseDataCurrentRegion(context);
    const autoFilterColorImportExport = await exerciseAutoFilterColorImportExport(context);
    const sortDialogLevels = await exerciseSortDialogLevels(context);
    const sortDialogColorOrder = await exerciseSortDialogColorOrder(context);
    const sortDialogCustomList = await exerciseSortDialogCustomList(context);
    const quickSortActiveColumn = await exerciseQuickSortActiveColumn(context);
    const quickSortPartialSelection = await exerciseQuickSortPartialSelection(context);
    await context.close();
    const result = {
      ok: true,
      autoFilterReapplyClear,
      importedAutoFilterCriteria,
      autoFilterConditionMenu,
      dataCurrentRegion,
      autoFilterColorImportExport,
      sortDialogLevels,
      sortDialogColorOrder,
      sortDialogCustomList,
      quickSortActiveColumn,
      quickSortPartialSelection,
    };
    assertDataFilterSortSmoke(result);
    console.log(JSON.stringify(result, null, 2));
  } finally {
    await browser.close();
  }
}



function assertCustomViewImportExportSmoke(result) {
  if (
    result?.imported?.name !== "ImportedView" ||
    result?.imported?.autoUpdate !== true ||
    result?.imported?.mergeInterval !== 17 ||
    result?.imported?.changesSavedWin !== true ||
    result?.imported?.onlySync !== true ||
    result?.imported?.personalView !== true ||
    result?.imported?.customMainWindowView?.showHorizontalScroll !== false ||
    result?.imported?.customMainWindowView?.showVerticalScroll !== false ||
    result?.imported?.customMainWindowView?.showSheetTabs !== false ||
    result?.exported?.workbookCustomAutoUpdate !== true ||
    result?.exported?.workbookCustomMergeInterval !== true ||
    result?.exported?.workbookCustomChangesSaved !== true ||
    result?.exported?.workbookCustomOnlySync !== true ||
    result?.exported?.workbookCustomPersonalView !== true ||
    result?.exported?.workbookCustomHorizontalScrollHidden !== true ||
    result?.exported?.workbookCustomVerticalScrollHidden !== true ||
    result?.exported?.workbookCustomSheetTabsHidden !== true ||
    result?.exported?.workbookPrDate1904 !== true ||
    result?.exported?.workbookPrCodeName !== true ||
    result?.exported?.calcPrManual !== true ||
    result?.exported?.calcPrIterate !== true ||
    result?.exported?.sheetAGridHidden !== true ||
    result?.exported?.sheetAHeadingsHidden !== true ||
    result?.exported?.sheetAFormulaShown !== true ||
    result?.exported?.sheetAZeroHidden !== true ||
    result?.exported?.sheetACustomWhiteSpaceShown !== true ||
    result?.exported?.sheetATopLeft !== true ||
    result?.exported?.sheetASelection !== true ||
    result?.exported?.sheetBPageLayout !== true ||
    result?.exported?.sheetBZoom !== true ||
    result?.exported?.sheetBViewWhiteSpaceHidden !== true ||
    result?.exported?.sheetBViewRightToLeft !== true ||
    result?.exported?.sheetBCustomWhiteSpaceHidden !== true ||
    result?.exported?.sheetBCustomRightToLeft !== true ||
    result?.exported?.sheetBTopLeft !== true ||
    result?.exported?.sheetBPane !== true ||
    result?.exported?.sheetBSelection !== true ||
    result?.deletedExport?.stateCount !== 0 ||
    result?.deletedExport?.workbookCustomViewsRemoved !== true ||
    result?.deletedExport?.sheetACustomViewsRemoved !== true ||
    result?.deletedExport?.sheetBCustomViewsRemoved !== true ||
    result?.deletedExport?.workbookPrKept !== true ||
    result?.deletedExport?.calcPrKept !== true
  ) {
    throw new Error(`Custom view import/export smoke failed: ${JSON.stringify(result)}`);
  }
}



async function runCustomViewImportExportOnly() {
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
    const result = await exerciseCustomViewImportExport(context);
    await context.close();
    assertCustomViewImportExportSmoke(result);
    console.log(JSON.stringify({ ok: true, customViewImportExport: result }, null, 2));
  } finally {
    await browser.close();
  }
}



async function runSaveSmokeOnly() {
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
    await page.goto(`${appUrl}?save-smoke=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    const loadedState = await page.evaluate(() => window.__webSheetDevWorkbookState?.() || {});

    const ctrlDownloadPromise = page.waitForEvent("download");
    await page.keyboard.press("Control+S");
    const ctrlDownload = await ctrlDownloadPromise;
    const ctrlHtml = await readFile(await ctrlDownload.path(), "utf8");
    const ctrlState = await page.evaluate(() => window.__webSheetDevWorkbookState?.() || {});

    await page.locator("#formulaInput").focus();
    const metaDownloadPromise = page.waitForEvent("download");
    await page.keyboard.press("Meta+S");
    const metaDownload = await metaDownloadPromise;
    const metaHtml = await readFile(await metaDownload.path(), "utf8");
    const formulaFocusedAfterMeta = await page.evaluate(() => document.activeElement?.id === "formulaInput");

    await page.click('[data-ribbon-tab="file"]');
    await page.waitForSelector("#fileSideMenu:not(.hidden)", { timeout: 10000 });
    await page.click('#fileSideMenu [data-ribbon-command="save-as"]');
    await page.waitForSelector('[data-data-dialog="save-as"]', { timeout: 10000 });
    const saveAsInitial = await page.evaluate(() => ({
      type: document.querySelector("[data-data-field='saveAsType']")?.value || "",
      fileName: document.querySelector("[data-data-field='saveAsFileName']")?.value || "",
      note: document.querySelector("[data-save-as-note]")?.textContent || "",
    }));
    await page.selectOption("[data-data-field='saveAsType']", "xlsx");
    const xlsxDownloadPromise = page.waitForEvent("download");
    await page.click('[data-data-dialog-action="save-as-ok"]');
    const xlsxDownload = await xlsxDownloadPromise;
    const xlsxBytes = await readFile(await xlsxDownload.path());

    const tempDir = await mkdtemp(path.join(os.tmpdir(), "websheet-save-smoke-"));
    const standaloneHtmlPath = path.join(tempDir, "save-smoke.html");
    await writeFile(standaloneHtmlPath, ctrlHtml, "utf8");
    const standaloneLocalSaveRequests = [];
    const standalonePage = await context.newPage();
    await standalonePage.route("**/*", async (route) => {
      const url = route.request().url();
      if (/^http:\/\/(?:127\.0\.0\.1|localhost):\d+\/__websheet\/local-save/i.test(url)) {
        let payload = {};
        try {
          payload = JSON.parse(route.request().postData() || "{}");
        } catch {
          payload = {};
        }
        standaloneLocalSaveRequests.push({
          url,
          path: payload.path || "",
          fileName: payload.fileName || "",
          hasWorkbookModel: String(payload.html || "").includes('id="websheet-model"'),
          isHtml: looksLikeStandaloneWorkbookHtml(String(payload.html || "")),
        });
        await route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ ok: true }) });
        return;
      }
      await route.continue();
    });
    await standalonePage.goto(pathToFileURL(standaloneHtmlPath).href, { waitUntil: "load" });
    await standalonePage.waitForFunction(
      () => document.querySelector("#sheetTabs .sheet-tab")?.textContent?.trim() === "CHT",
      null,
      { timeout: 30000 },
    );
    await standalonePage.evaluate(() => {
      window.__saveSmokePickerCalls = [];
      Object.defineProperty(window, "showSaveFilePicker", {
        configurable: true,
        value: async (options = {}) => {
          window.__saveSmokePickerCalls.push({ suggestedName: String(options.suggestedName || "") });
          throw new DOMException("Picker should not be used for direct local save", "AbortError");
        },
      });
    });
    await standalonePage.keyboard.press("Control+S");
    await standalonePage.waitForFunction(() => (document.querySelector("#statusText")?.textContent || "").includes("上書き保存"), null, { timeout: 10000 });
    const standaloneState = await standalonePage.evaluate(() => ({
      workbookState: {
        sourceName: document.querySelector("#workbookName")?.textContent || "",
        saveType: window.__WEBSHEET_STANDALONE_MODEL__?.saveType || "",
        currentSaveType: window.__WEBSHEET_STANDALONE_MODEL__?.saveType || "",
        activeSheetName: document.querySelector("#sheetTabs .sheet-tab.is-active")?.textContent?.trim() || "",
        isStandalone: window.__WEBSHEET_STANDALONE__ === true,
      },
      pickerCalls: window.__saveSmokePickerCalls || [],
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    const exportedModel = exportedStandaloneModelFromHtml(ctrlHtml);
    const wrappedStringPayloadPath = path.join(tempDir, "wrapped-string-payload.html");
    await writeFile(
      wrappedStringPayloadPath,
      replaceStandaloneModelPayload(ctrlHtml, { workbook: { payload: JSON.stringify(exportedModel) } }),
      "utf8",
    );
    const wrappedPage = await context.newPage();
    await wrappedPage.goto(pathToFileURL(wrappedStringPayloadPath).href, { waitUntil: "load" });
    await wrappedPage.waitForFunction(
      () => document.querySelector("#sheetTabs .sheet-tab.is-active")?.textContent?.trim() === "CHT",
      null,
      { timeout: 30000 },
    );
    const wrappedState = await wrappedPage.evaluate(() => ({
      activeSheetName: document.querySelector("#sheetTabs .sheet-tab.is-active")?.textContent?.trim() || "",
      sourceName: document.querySelector("#workbookName")?.textContent || "",
      saveType: window.__WEBSHEET_STANDALONE_MODEL__?.saveType || "",
      isStandalone: window.__WEBSHEET_STANDALONE__ === true,
    }));

    await wrappedPage.close();
    await standalonePage.close();
    await page.close();
    await context.close();

    const result = {
      loadedState: {
        sourceName: loadedState.sourceName || "",
        saveType: loadedState.saveType || "",
        currentSaveType: loadedState.currentSaveType || "",
        isStandalone: loadedState.isStandalone === true,
      },
      ctrlSave: {
        suggestedFilename: ctrlDownload.suggestedFilename(),
        isHtml: looksLikeStandaloneWorkbookHtml(ctrlHtml),
        hasWorkbookModel: ctrlHtml.includes('id="websheet-model"'),
        state: {
          sourceName: ctrlState.sourceName || "",
          saveType: ctrlState.saveType || "",
          currentSaveType: ctrlState.currentSaveType || "",
          isStandalone: ctrlState.isStandalone === true,
        },
      },
      metaSave: {
        suggestedFilename: metaDownload.suggestedFilename(),
        isHtml: looksLikeStandaloneWorkbookHtml(metaHtml),
        formulaFocusedAfterMeta,
      },
      saveAs: {
        initial: saveAsInitial,
        xlsxSuggestedFilename: xlsxDownload.suggestedFilename(),
        xlsxZipHeader: xlsxBytes.subarray(0, 4).toString("hex"),
      },
      standalone: {
        state: standaloneState.workbookState,
        status: standaloneState.status,
        pickerCalls: standaloneState.pickerCalls,
        localSaveRequests: standaloneLocalSaveRequests,
      },
      wrappedStringPayload: {
        activeSheetName: wrappedState.activeSheetName || "",
        sourceName: wrappedState.sourceName || "",
        saveType: wrappedState.saveType || "",
        isStandalone: wrappedState.isStandalone === true,
      },
    };

    const failures = [];
    if (result.loadedState.saveType !== "html" || result.loadedState.currentSaveType !== "html" || result.loadedState.isStandalone) {
      failures.push("dev server workbook should default quick save to html");
    }
    if (!/\.html$/i.test(result.ctrlSave.suggestedFilename) || !result.ctrlSave.isHtml || !result.ctrlSave.hasWorkbookModel) {
      failures.push("Ctrl+S from dev server should download standalone html");
    }
    if (result.ctrlSave.state.saveType !== "html" || result.ctrlSave.state.currentSaveType !== "html") {
      failures.push("Ctrl+S should keep workbook save state as html");
    }
    if (!/\.html$/i.test(result.metaSave.suggestedFilename) || !result.metaSave.isHtml || !result.metaSave.formulaFocusedAfterMeta) {
      failures.push("Meta+S should save html and keep formula input focus");
    }
    if (result.saveAs.initial.type !== "html" || !/\.html$/i.test(result.saveAs.initial.fileName)) {
      failures.push("Save As should initially select html on dev server");
    }
    if (!/\.xlsx$/i.test(result.saveAs.xlsxSuggestedFilename) || result.saveAs.xlsxZipHeader !== "504b0304") {
      failures.push("Save As xlsx should download a zip-based Excel workbook");
    }
    if (result.standalone.pickerCalls.length || result.standalone.localSaveRequests.length !== 1) {
      failures.push("file:// standalone quick save should use local-save bridge without picker");
    }
    if (!result.standalone.localSaveRequests[0]?.isHtml || !result.standalone.localSaveRequests[0]?.path.endsWith("save-smoke.html")) {
      failures.push("local-save bridge should receive the current html path and workbook html");
    }
    if (result.wrappedStringPayload.activeSheetName !== "CHT" || !result.wrappedStringPayload.isStandalone) {
      failures.push("standalone loader should parse nested string JSON payloads");
    }
    if (failures.length) throw new Error(`Save smoke failed: ${failures.join("; ")}\n${JSON.stringify(result, null, 2)}`);
    console.log(JSON.stringify({ ok: true, saveSmoke: result }, null, 2));
  } finally {
    await browser.close();
  }
}

export {
  assertAutoFilterDateSmoke,
  assertCustomViewImportExportSmoke,
  assertDataFilterSortSmoke,
  assertValidationMenuResult,
  runAltEnterSmokeOnly,
  runAutoFilterDateOnly,
  runClipboardSessionSmokeOnly,
  runCustomViewImportExportOnly,
  runDataFilterSortOnly,
  runDatePersistenceSmokeOnly,
  runMacroVbaSubsetOnly,
  runOutlineGroupingOnly,
  runSaveSmokeOnly,
  runValidationMenuOnly,
};
