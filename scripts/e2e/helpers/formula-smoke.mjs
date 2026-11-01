// Sequential E2E feature module.
import {
  JSZip,
  appUrl,
  browser,
  chromePath,
  chromium,
  customViewWorkbookPath,
  path,
  readFile,
  workbookPath,
} from "../../e2e-real-upload.mjs";
import {
  clickRibbonTab,
  e2eProgress,
  exportedDefinedNamesWorkbookState,
  exportedPageBreakWorkbookState,
  exportedSheetViewWorkbookState,
  waitForUploadedWorkbook,
} from "./fixtures-and-model.mjs";


async function exerciseCustomViewImportExport(context) {
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?customView=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", customViewWorkbookPath);
    await waitForUploadedWorkbook(page, customViewWorkbookPath);
    await page.waitForFunction(() => (window.__webSheetDevWorkbookState?.().customViews || []).length > 0, null, { timeout: 10000 });
    const imported = await page.evaluate(() => {
      const view = (window.__webSheetDevWorkbookState?.().customViews || [])[0] || null;
      const state = window.__webSheetDevWorkbookState?.() || {};
      return {
        name: view?.name || "",
        guid: view?.guid || "",
        autoUpdate: view?.autoUpdate,
        mergeInterval: view?.mergeInterval,
        changesSavedWin: view?.changesSavedWin,
        onlySync: view?.onlySync,
        personalView: view?.personalView,
        showFormulaBar: view?.showFormulaBar,
        activeSheetIndex: view?.activeSheetIndex,
        activeScroll: view?.activeScroll || null,
        customMainWindowView: view?.mainWindowView || null,
        workbookProperties: state.workbookProperties || null,
        calcProperties: state.calcProperties || null,
        workbookView: state.workbookView || null,
        mainWindowView: state.mainWindowView || null,
        mainWindowHidden: state.mainWindowHidden,
        windowViews: state.windowViews || [],
        activeWindowId: state.activeWindowId || "",
        date1904Cell: window.__webSheetDevCellModel?.("CustomA", 5, 5) || null,
        sheetA: view?.sheets?.[0] || null,
        sheetB: view?.sheets?.[1] || null,
      };
    });

    await page.click('[data-ribbon-tab="view"]');
    await page.click('[data-ribbon-command="view-custom-views"]');
    await page.waitForSelector('[data-data-dialog="custom-views"]', { timeout: 5000 });
    const dialog = await page.evaluate(() => ({
      title: document.querySelector("#customViewsTitle")?.textContent || "",
      names: [...document.querySelectorAll("[data-data-field='customViewIndex'] option")].map((option) => option.textContent.trim()),
      showDisabled: document.querySelector("[data-data-dialog-action='custom-view-show']")?.disabled,
    }));
    await page.selectOption("[data-data-field='customViewIndex']", { label: "ImportedView" });
    await page.click("[data-data-dialog-action='custom-view-show']");
    await page.waitForFunction(() => window.__webSheetDevWorkbookState?.().activeSheetName === "CustomB", null, { timeout: 5000 });
    const applied = await page.evaluate(() => ({
      activeSheetName: window.__webSheetDevWorkbookState?.().activeSheetName || "",
      selectionLabel: document.querySelector("#selectedCellLabel")?.textContent || "",
      showFormulaBar: !document.querySelector(".formula-bar")?.classList.contains("hidden"),
      viewMode: window.__webSheetDevSheetModel?.("CustomB")?.viewMode || "",
      zoom: window.__webSheetDevSheetModel?.("CustomB")?.zoomScale || 0,
      freezePanes: window.__webSheetDevSheetModel?.("CustomB")?.freezePanes || null,
      showWhiteSpace: window.__webSheetDevSheetModel?.("CustomB")?.showWhiteSpace,
      rightToLeft: window.__webSheetDevSheetModel?.("CustomB")?.rightToLeft,
      workbookView: window.__webSheetDevWorkbookState?.().workbookView || null,
      mainWindowView: window.__webSheetDevWorkbookState?.().mainWindowView || null,
      hideHorizontalScrollbar: document.body.classList.contains("workbook-hide-horizontal-scrollbar"),
      hideVerticalScrollbar: document.body.classList.contains("workbook-hide-vertical-scrollbar"),
      hideSheetTabs: document.body.classList.contains("workbook-hide-sheet-tabs"),
      scrollTop: document.querySelector(".sheet-host")?.scrollTop || 0,
      scrollLeft: document.querySelector(".sheet-host")?.scrollLeft || 0,
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.click('[data-ribbon-tab="file"]');
    await page.waitForFunction(() => !document.querySelector("#fileSideMenu")?.classList.contains("hidden"), null, { timeout: 5000 });
    await page.click('#fileSideMenu [data-ribbon-command="save-as"]');
    await page.waitForSelector('[data-data-dialog="save-as"]', { timeout: 5000 });
    await page.selectOption("[data-data-field='saveAsType']", "xlsx");
    const downloadPromise = page.waitForEvent("download");
    await page.click('[data-data-dialog-action="save-as-ok"]');
    const download = await downloadPromise;
    const zip = await JSZip.loadAsync(await readFile(await download.path()));
    const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
    const sheet1Xml = await zip.file("xl/worksheets/sheet1.xml")?.async("text");
    const sheet2Xml = await zip.file("xl/worksheets/sheet2.xml")?.async("text");
    const workbookViewTags = workbookXml?.match(/<workbookView\b[^>]*\/?>/g) || [];
    const sheet1ViewTag = sheet1Xml?.match(/<sheetView\b[^>]*\/?>/)?.[0] || "";
    const sheet2ViewTag = sheet2Xml?.match(/<sheetView\b[^>]*\/?>/)?.[0] || "";
    const sheet1CustomViewTag = sheet1Xml?.match(/<customSheetView\b[^>]*>/)?.[0] || "";
    const sheet2CustomViewTag = sheet2Xml?.match(/<customSheetView\b[^>]*>/)?.[0] || "";
    const hasWorkbookView = (...attributes) => workbookViewTags.some((tag) => attributes.every((attribute) => tag.includes(attribute)));
    const exported = {
      workbookHasCustomViews: /<customWorkbookViews\b/.test(workbookXml || ""),
      workbookName: /name="ImportedView"/.test(workbookXml || ""),
      workbookActiveSheet: /activeSheetId="1"/.test(workbookXml || ""),
      workbookFormulaBarHidden: /showFormulaBar="0"/.test(workbookXml || ""),
      workbookCustomAutoUpdate: /<customWorkbookView\b[^>]*name="ImportedView"[^>]*autoUpdate="1"/.test(workbookXml || ""),
      workbookCustomMergeInterval: /<customWorkbookView\b[^>]*name="ImportedView"[^>]*mergeInterval="17"/.test(workbookXml || ""),
      workbookCustomChangesSaved: /<customWorkbookView\b[^>]*name="ImportedView"[^>]*changesSavedWin="1"/.test(workbookXml || ""),
      workbookCustomOnlySync: /<customWorkbookView\b[^>]*name="ImportedView"[^>]*onlySync="1"/.test(workbookXml || ""),
      workbookCustomPersonalView: /<customWorkbookView\b[^>]*name="ImportedView"[^>]*personalView="1"/.test(workbookXml || ""),
      workbookCustomHorizontalScrollHidden: /<customWorkbookView\b[^>]*name="ImportedView"[^>]*showHorizontalScroll="0"/.test(workbookXml || ""),
      workbookCustomVerticalScrollHidden: /<customWorkbookView\b[^>]*name="ImportedView"[^>]*showVerticalScroll="0"/.test(workbookXml || ""),
      workbookCustomSheetTabsHidden: /<customWorkbookView\b[^>]*name="ImportedView"[^>]*showSheetTabs="0"/.test(workbookXml || ""),
      workbookPrDate1904: /<workbookPr\b[^>]*date1904="1"/.test(workbookXml || ""),
      workbookPrCodeName: /<workbookPr\b[^>]*codeName="ThisWorkbook"/.test(workbookXml || ""),
      workbookPrFilterPrivacy: /<workbookPr\b[^>]*filterPrivacy="1"/.test(workbookXml || ""),
      calcPrManual: /<calcPr\b[^>]*calcMode="manual"/.test(workbookXml || ""),
      calcPrR1C1: /<calcPr\b[^>]*refMode="R1C1"/.test(workbookXml || ""),
      calcPrIterate: /<calcPr\b[^>]*iterate="1"/.test(workbookXml || "") && /<calcPr\b[^>]*iterateCount="42"/.test(workbookXml || ""),
      sheetADate1904RoundTrip: /<c\b[^>]*r="E5"[\s\S]*?<v>36922\.375<\/v>/.test(sheet1Xml || ""),
      workbookViewCount: workbookViewTags.length >= 2,
      workbookMainViewVisible: hasWorkbookView('visibility="visible"'),
      workbookSecondViewHidden: hasWorkbookView('activeTab="1"', 'visibility="hidden"'),
      workbookSecondViewSize: hasWorkbookView('activeTab="1"', 'xWindow="410"', 'windowWidth="18000"'),
      sheetAHasCustomView: /<customSheetView\b/.test(sheet1Xml || ""),
      sheetAGridHidden: /showGridLines="0"/.test(sheet1Xml || ""),
      sheetAHeadingsHidden: /showRowCol="0"/.test(sheet1Xml || ""),
      sheetAFormulaShown: /showFormulas="1"/.test(sheet1Xml || ""),
      sheetAZeroHidden: /zeroValues="0"/.test(sheet1Xml || ""),
      sheetACustomWhiteSpaceShown: sheet1CustomViewTag.includes('showWhiteSpace="1"'),
      sheetACustomRightToLeftOff: sheet1CustomViewTag.includes('rightToLeft="0"'),
      sheetATopLeft: /topLeftCell="D5"/.test(sheet1Xml || ""),
      sheetASelection: /activeCell="B2"/.test(sheet1Xml || "") && /sqref="B2:C3"/.test(sheet1Xml || ""),
      sheetBHasCustomView: /<customSheetView\b/.test(sheet2Xml || ""),
      sheetBPageLayout: /view="pageLayout"/.test(sheet2Xml || ""),
      sheetBZoom: /scale="125"/.test(sheet2Xml || ""),
      sheetBViewWhiteSpaceHidden: sheet2ViewTag.includes('showWhiteSpace="0"'),
      sheetBViewRightToLeft: sheet2ViewTag.includes('rightToLeft="1"'),
      sheetBCustomWhiteSpaceHidden: sheet2CustomViewTag.includes('showWhiteSpace="0"'),
      sheetBCustomRightToLeft: sheet2CustomViewTag.includes('rightToLeft="1"'),
      sheetBTopLeft: /topLeftCell="F8"/.test(sheet2Xml || ""),
      sheetBPane: /xSplit="1"/.test(sheet2Xml || "") && /ySplit="2"/.test(sheet2Xml || "") && /state="frozen"/.test(sheet2Xml || ""),
      sheetBSelection: /activeCell="D4"/.test(sheet2Xml || "") && /sqref="D4"/.test(sheet2Xml || ""),
    };
    await page.click('[data-ribbon-tab="view"]');
    await page.click('[data-ribbon-command="view-custom-views"]');
    await page.waitForSelector('[data-data-dialog="custom-views"]', { timeout: 5000 });
    await page.selectOption("[data-data-field='customViewIndex']", { label: "ImportedView" });
    await page.click("[data-data-dialog-action='custom-view-delete']");
    await page.waitForSelector("#messageBox", { timeout: 5000 });
    await page.click('[data-message-box-result="delete"]');
    await page.waitForFunction(() => (window.__webSheetDevWorkbookState?.().customViews || []).length === 0, null, { timeout: 5000 });
    const deletedState = await page.evaluate(() => ({
      count: (window.__webSheetDevWorkbookState?.().customViews || []).length,
      names: (window.__webSheetDevWorkbookState?.().customViews || []).map((view) => view.name),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.click('[data-data-dialog-action="cancel"]');
    await page.waitForSelector('[data-data-dialog="custom-views"]', { state: "detached", timeout: 5000 });
    await page.click('[data-ribbon-tab="file"]');
    await page.waitForFunction(() => !document.querySelector("#fileSideMenu")?.classList.contains("hidden"), null, { timeout: 5000 });
    await page.click('#fileSideMenu [data-ribbon-command="save-as"]');
    await page.waitForSelector('[data-data-dialog="save-as"]', { timeout: 5000 });
    await page.selectOption("[data-data-field='saveAsType']", "xlsx");
    const deletedDownloadPromise = page.waitForEvent("download");
    await page.click('[data-data-dialog-action="save-as-ok"]');
    const deletedDownload = await deletedDownloadPromise;
    const deletedZip = await JSZip.loadAsync(await readFile(await deletedDownload.path()));
    const deletedWorkbookXml = await deletedZip.file("xl/workbook.xml")?.async("text");
    const deletedSheet1Xml = await deletedZip.file("xl/worksheets/sheet1.xml")?.async("text");
    const deletedSheet2Xml = await deletedZip.file("xl/worksheets/sheet2.xml")?.async("text");
    const deletedExport = {
      stateCount: deletedState.count,
      stateNames: deletedState.names,
      status: deletedState.status,
      workbookCustomViewsRemoved: !/<customWorkbookViews\b/.test(deletedWorkbookXml || ""),
      sheetACustomViewsRemoved: !/<customSheetViews\b/.test(deletedSheet1Xml || ""),
      sheetBCustomViewsRemoved: !/<customSheetViews\b/.test(deletedSheet2Xml || ""),
      workbookPrKept: /<workbookPr\b[^>]*date1904="1"/.test(deletedWorkbookXml || ""),
      calcPrKept: /<calcPr\b[^>]*calcMode="manual"/.test(deletedWorkbookXml || ""),
    };
    return { imported, dialog, applied, exported, deletedExport };
  } finally {
    await page.close();
  }
}



async function runFormulaSmokeOnly() {
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
    await page.goto(`${appUrl}?t=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await page.waitForFunction(() => window.__webSheetDevCellDisplay && window.__webSheetDevSetCell && window.__webSheetDevFormulaEngineCell, null, { timeout: 30000 });
    await page.waitForTimeout(1000);
    const result = await page.evaluate(() => {
      const seed = [
        [90, 27, 1],
        [91, 27, 2],
        [92, 27, 2],
        [93, 27, 4],
        [94, 27, 5],
        [90, 28, 10],
        [91, 28, 20],
        [92, 28, 15],
        [93, 28, 5],
        [94, 28, 25],
        [90, 29, "alpha"],
        [91, 29, "bravo"],
        [92, 29, "apricot"],
        [93, 29, "delta"],
        [94, 29, "anchor"],
        [90, 30, 5],
        [91, 30, 4],
        [92, 30, 2],
        [93, 30, 1],
        [94, 30, 0],
        [90, 31, 50],
        [91, 31, 40],
        [92, 31, 20],
        [93, 31, 10],
        [94, 31, 0],
        [90, 32, ""],
        [91, 32, "x"],
        [92, 32, ""],
        [93, 32, 0],
        [94, 32, ""],
      ];
      seed.forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
      const formulas = [
        ["textsplit", 90, 115, '=TEXTSPLIT("A,B;C,D",",",";")'],
        ["textsplitIgnoreEmpty", 90, 121, '=TEXTSPLIT("A,,B",",",,TRUE)'],
        ["textbefore", 90, 119, '=TEXTBEFORE("A-B-C","-",2)'],
        ["textafter", 90, 120, '=TEXTAFTER("A-B-C","-",2)'],
        ["textafterLast", 90, 123, '=TEXTAFTER("A-B-C","-",-1)'],
        ["textbeforeFallback", 90, 124, '=TEXTBEFORE("ABC","-",1,0,FALSE,"NF")'],
        ["aggregateSum", 90, 126, "=AGGREGATE(9,0,AA90:AA94)"],
        ["aggregateLarge", 90, 127, "=AGGREGATE(14,0,AA90:AA94,2)"],
        ["aggregateMedian", 90, 128, "=AGGREGATE(12,0,AA90:AA94)"],
        ["aggregateIgnoreError", 90, 129, "=AGGREGATE(9,6,AA90:AA94,1/0)"],
        ["ifs", 90, 130, "=IFS(FALSE,1,TRUE,2)"],
        ["xor", 90, 131, "=XOR(TRUE,FALSE,FALSE)"],
        ["letSimple", 90, 132, "=LET(x,2,x+3)"],
        ["letMulti", 90, 133, "=LET(x,2,y,x+3,y*2)"],
        ["letRange", 90, 134, "=LET(nums,AA90:AA94,SUM(nums))"],
        ["makearray", 90, 135, "=MAKEARRAY(2,2,LAMBDA(r,c,r+c))"],
        ["byrow", 90, 138, "=BYROW(AA90:AB91,LAMBDA(row,SUM(row)))"],
        ["bycol", 90, 140, "=BYCOL(AA90:AB91,LAMBDA(col,SUM(col)))"],
        ["map", 90, 143, "=MAP(AA90:AA91,LAMBDA(x,x*2))"],
        ["reduce", 90, 145, "=REDUCE(0,AA90:AA94,LAMBDA(acc,value,acc+value))"],
        ["scan", 90, 147, "=SCAN(0,AA90:AA92,LAMBDA(acc,value,acc+value))"],
        ["sortby", 90, 149, "=SORTBY(AA90:AB94,AB90:AB94,-1)"],
        ["textsplitMultiDelimiter", 90, 152, '=TEXTSPLIT("A,B;C|D",{",",";","|"})'],
        ["xmatchWildcard", 90, 156, '=XMATCH("a*",AC90:AC94,2)'],
        ["xmatchWildcardReverse", 91, 156, '=XMATCH("a*",AC90:AC94,2,-1)'],
        ["takeLastRows", 90, 158, "=TAKE(AA90:AB94,-2)"],
        ["dropLastRows", 90, 161, "=DROP(AA90:AB94,-2)"],
        ["xlookupExact", 90, 164, '=XLOOKUP(4,AA90:AA94,AB90:AB94,"NF")'],
        ["xlookupReverseWildcard", 91, 164, '=XLOOKUP("a*",AC90:AC94,AA90:AA94,"NF",2,-1)'],
        ["xlookupRowSpill", 90, 166, '=XLOOKUP(4,AA90:AA94,AA90:AB94,"NF")'],
        ["filterRows", 90, 169, '=FILTER(AA90:AB94,AA90:AA94>2,"NF")'],
        ["filterEmptyFallback", 90, 172, '=FILTER(AA90:AB94,AA90:AA94>99,"NF")'],
        ["xlookupApproxSmaller", 90, 174, '=XLOOKUP(4.5,AA90:AA94,AB90:AB94,"NF",-1)'],
        ["xlookupApproxLarger", 90, 175, '=XLOOKUP(3,AA90:AA94,AB90:AB94,"NF",1)'],
        ["xmatchBinaryAsc", 90, 176, "=XMATCH(4,AA90:AA94,0,2)"],
        ["xmatchBinaryDesc", 90, 177, "=XMATCH(4,AD90:AD94,0,-2)"],
        ["xlookupBinaryDesc", 90, 178, '=XLOOKUP(3,AD90:AD94,AE90:AE94,"NF",1,-2)'],
        ["transpose", 90, 180, "=TRANSPOSE(AA90:AB91)"],
        ["chooseArray", 90, 183, "=CHOOSE(2,AA90:AA91,AB90:AB91)"],
        ["indexColumn", 90, 185, "=INDEX(AA90:AB94,,2)"],
        ["offsetRange", 90, 187, "=OFFSET(AA90,1,0,2,2)"],
        ["filterAndConditions", 90, 190, '=FILTER(AA90:AB94,(AA90:AA94>1)*(AB90:AB94<20),"NF")'],
        ["filterOrConditions", 90, 193, '=FILTER(AA90:AB94,(AA90:AA94=1)+(AB90:AB94=25),"NF")'],
        ["sumproductPairs", 90, 196, "=SUMPRODUCT(AA90:AA94,AB90:AB94)"],
        ["sumproductBoolean", 91, 196, "=SUMPRODUCT((AA90:AA94>1)*AB90:AB94)"],
        ["countblank", 92, 196, "=COUNTBLANK(AF90:AF94)"],
        ["countif", 93, 196, "=COUNTIF(AA90:AA94,2)"],
        ["countifs", 94, 196, '=COUNTIFS(AA90:AA94,">1",AB90:AB94,"<20")'],
        ["sumif", 95, 196, "=SUMIF(AA90:AA94,2,AB90:AB94)"],
        ["sumifs", 96, 196, '=SUMIFS(AB90:AB94,AA90:AA94,">1",AB90:AB94,"<20")'],
        ["averageif", 97, 196, "=AVERAGEIF(AA90:AA94,2,AB90:AB94)"],
        ["averageifs", 98, 196, '=AVERAGEIFS(AB90:AB94,AA90:AA94,">1",AB90:AB94,"<20")'],
        ["maxifs", 99, 196, '=MAXIFS(AB90:AB94,AA90:AA94,">1",AB90:AB94,"<20")'],
        ["minifs", 100, 196, '=MINIFS(AB90:AB94,AA90:AA94,">1",AB90:AB94,"<20")'],
        ["networkdays", 101, 196, "=NETWORKDAYS(DATE(2023,1,2),DATE(2023,1,6))"],
        ["networkdaysHoliday", 102, 196, "=NETWORKDAYS(DATE(2023,1,2),DATE(2023,1,6),DATE(2023,1,4))"],
        ["networkdaysIntl", 103, 196, "=NETWORKDAYS.INTL(DATE(2023,1,2),DATE(2023,1,8),11)"],
        ["workday", 104, 196, "=WORKDAY(DATE(2023,1,6),1)-DATE(2023,1,6)"],
        ["workdayIntl", 105, 196, "=WORKDAY.INTL(DATE(2023,1,6),1,11)-DATE(2023,1,6)"],
        ["ceilingMathPositive", 106, 196, "=CEILING.MATH(5.5,2)"],
        ["ceilingMathNegative", 107, 196, "=CEILING.MATH(-5.5,2)"],
        ["ceilingMathNegativeAway", 108, 196, "=CEILING.MATH(-5.5,2,1)"],
        ["imabs", 109, 196, '=IMABS("3+4i")'],
        ["datedifMonths", 110, 196, '=DATEDIF(DATE(2020,1,15),DATE(2021,3,14),"M")'],
        ["datedifYm", 111, 196, '=DATEDIF(DATE(2020,1,15),DATE(2021,3,14),"YM")'],
        ["switchMatch", 112, 196, '=SWITCH("B","A",1,"B",2,9)'],
        ["switchDefault", 113, 196, '=SWITCH("Z","A",1,"B",2,9)'],
        ["edate", 114, 196, "=EDATE(DATE(2023,1,31),1)-DATE(2023,2,28)"],
        ["eomonth", 115, 196, "=EOMONTH(DATE(2023,1,15),1)-DATE(2023,2,28)"],
        ["days", 116, 196, "=DAYS(DATE(2023,1,6),DATE(2023,1,2))"],
        ["days360", 117, 196, "=DAYS360(DATE(2023,1,1),DATE(2023,2,1))"],
        ["yearfracBasis2", 118, 196, "=ROUND(YEARFRAC(DATE(2023,1,1),DATE(2023,7,1),2),6)"],
        ["floorMathPositive", 119, 196, "=FLOOR.MATH(5.5,2)"],
        ["floorMathNegative", 120, 196, "=FLOOR.MATH(-5.5,2)"],
        ["floorMathNegativeToward", 121, 196, "=FLOOR.MATH(-5.5,2,1)"],
        ["mround", 122, 196, "=MROUND(5,2)"],
        ["isblank", 123, 196, "=ISBLANK(AF90)"],
        ["iserror", 124, 196, "=ISERROR(1/0)"],
        ["iserr", 125, 196, "=ISERR(1/0)"],
        ["isna", 126, 196, "=ISNA(NA())"],
        ["isnumber", 127, 196, "=ISNUMBER(AA90)"],
        ["istext", 128, 196, '=ISTEXT("abc")'],
        ["isnontext", 129, 196, "=ISNONTEXT(123)"],
        ["islogical", 130, 196, "=ISLOGICAL(TRUE)"],
        ["iseven", 131, 196, "=ISEVEN(2.9)"],
        ["isodd", 132, 196, "=ISODD(-3)"],
        ["nText", 133, 196, '=N("abc")'],
        ["nTrue", 134, 196, "=N(TRUE)"],
        ["imreal", 135, 196, '=IMREAL("3+4i")'],
        ["imaginary", 136, 196, '=IMAGINARY("3+4i")'],
        ["imsum", 137, 196, '=IMSUM("1+2i","3-4i")'],
        ["improduct", 138, 196, '=IMPRODUCT("1+2i","3+4i")'],
        ["imsub", 139, 196, '=IMSUB("3+4i","1+2i")'],
        ["imdiv", 140, 196, '=IMDIV("3+4i","1+0i")'],
        ["imconjugate", 141, 196, '=IMCONJUGATE("3+4i")'],
        ["datevalueIso", 142, 196, '=DATEVALUE("2023/1/2")-DATE(2023,1,2)'],
        ["datevalueJapanese", 143, 196, '=DATEVALUE("2023年1月2日")-DATE(2023,1,2)'],
        ["timevaluePm", 144, 196, '=TIMEVALUE("6:45 PM")'],
        ["hour", 145, 196, '=HOUR(TIMEVALUE("18:45:30"))'],
        ["minute", 146, 196, '=MINUTE(TIMEVALUE("18:45:30"))'],
        ["second", 147, 196, '=SECOND(TIMEVALUE("18:45:30"))'],
        ["weekdayDefault", 148, 196, '=WEEKDAY(DATE(2023,1,1))'],
        ["weekdayMonday", 149, 196, '=WEEKDAY(DATE(2023,1,2),2)'],
        ["weeknumSunday", 150, 196, '=WEEKNUM(DATE(2023,1,8),1)'],
        ["weeknumMonday", 151, 196, '=WEEKNUM(DATE(2023,1,2),2)'],
        ["isoweeknum", 152, 196, '=ISOWEEKNUM(DATE(2023,1,2))'],
        ["timevalueJapanesePm", 153, 196, '=HOUR(TIMEVALUE("午後 1:05"))'],
        ["concatenate", 154, 196, '=CONCATENATE("A","B","C")'],
        ["replaceText", 155, 196, '=REPLACE("abcde",2,2,"XX")'],
        ["proper", 156, 196, '=PROPER("hello WORLD")'],
        ["exactTrue", 157, 196, '=EXACT("WebSheet","WebSheet")'],
        ["exactFalse", 158, 196, '=EXACT("WebSheet","websheet")'],
        ["char", 159, 196, "=CHAR(65)"],
        ["code", 160, 196, '=CODE("Azami")'],
        ["clean", 161, 196, '=CLEAN(CHAR(7)&"A"&CHAR(10))'],
        ["rept", 162, 196, '=REPT("ab",3)'],
        ["averagea", 163, 196, '=AVERAGEA(TRUE,FALSE,"text",2)'],
        ["maxa", 164, 196, '=MAXA(TRUE,FALSE,"text",-5)'],
        ["mina", 165, 196, '=MINA(TRUE,FALSE,"text",-5)'],
        ["ceilingPrecise", 166, 196, "=CEILING.PRECISE(-5.5,2)"],
        ["floorPrecise", 167, 196, "=FLOOR.PRECISE(-5.5,2)"],
        ["quotient", 168, 196, "=QUOTIENT(7,3)"],
        ["sign", 169, 196, "=SIGN(-10)"],
        ["pmt", 170, 196, "=ROUND(PMT(0.05/12,12,1000),2)"],
        ["pv", 171, 196, "=ROUND(PV(0.05/12,12,-85.6074817884674),2)"],
        ["fv", 172, 196, "=ROUND(ABS(FV(0.05/12,12,-85.6074817884674,1000)),2)"],
        ["nper", 173, 196, "=ROUND(NPER(0.05/12,-85.6074817884674,1000),0)"],
        ["rate", 174, 196, "=ROUND(RATE(12,-85.6074817884674,1000)*1200,2)"],
        ["npv", 175, 196, "=ROUND(NPV(0.1,100,100),2)"],
        ["irr", 176, 196, "=ROUND(IRR({-100,60,60}),6)"],
        ["mirr", 177, 196, "=ROUND(MIRR({-100,60,60},0.1,0.1),6)"],
        ["ipmt", 178, 196, "=ROUND(IPMT(0.05/12,1,12,1000),2)"],
        ["ppmt", 179, 196, "=ROUND(PPMT(0.05/12,1,12,1000),2)"],
        ["cumipmt", 180, 196, "=ROUND(CUMIPMT(0.05/12,12,1000,1,1,0),2)"],
        ["cumprinc", 181, 196, "=ROUND(CUMPRINC(0.05/12,12,1000,1,1,0),2)"],
        ["effect", 182, 196, "=ROUND(EFFECT(0.12,12),6)"],
        ["nominal", 183, 196, "=ROUND(NOMINAL(EFFECT(0.12,12),12),6)"],
        ["roman", 184, 196, "=ROMAN(1999)"],
        ["arabic", 185, 196, '=ARABIC("MCMXCIX")'],
        ["base", 186, 196, "=BASE(255,16,4)"],
        ["decimal", 187, 196, '=DECIMAL("00FF",16)'],
        ["unichar", 188, 196, "=UNICHAR(9731)"],
        ["unicode", 189, 196, '=UNICODE("雪")'],
        ["stdevS", 190, 196, "=ROUND(STDEV.S({1,2,3}),6)"],
        ["stdevP", 191, 196, "=ROUND(STDEV.P({1,2,3}),6)"],
        ["stdevLegacy", 192, 196, "=ROUND(STDEV({1,2,3}),6)"],
        ["stdevpLegacy", 193, 196, "=ROUND(STDEVP({1,2,3}),6)"],
        ["varS", 194, 196, "=ROUND(VAR.S({1,2,3}),6)"],
        ["varP", 195, 196, "=ROUND(VAR.P({1,2,3}),6)"],
        ["vara", 196, 196, '=ROUND(VARA(1,TRUE,"x"),6)'],
        ["varpa", 197, 196, '=ROUND(VARPA(1,TRUE,"x"),6)'],
        ["stdeva", 198, 196, '=ROUND(STDEVA(1,TRUE,"x"),6)'],
        ["stdevpa", 199, 196, '=ROUND(STDEVPA(1,TRUE,"x"),6)'],
        ["correl", 200, 196, "=ROUND(CORREL({1,2,3},{2,4,6}),6)"],
        ["covarianceS", 201, 196, "=ROUND(COVARIANCE.S({1,2,3},{2,4,6}),6)"],
        ["covarianceP", 202, 196, "=ROUND(COVARIANCE.P({1,2,3},{2,4,6}),6)"],
        ["covar", 203, 196, "=ROUND(COVAR({1,2,3},{2,4,6}),6)"],
        ["slope", 204, 196, "=ROUND(SLOPE({2,4,6},{1,2,3}),6)"],
        ["avedev", 205, 196, "=AVEDEV({1,2,3,4})"],
        ["devsq", 206, 196, "=DEVSQ({1,2,3})"],
        ["geomean", 207, 196, "=GEOMEAN({4,16})"],
        ["harmean", 208, 196, "=HARMEAN({2,6})"],
        ["pearson", 209, 196, "=ROUND(PEARSON({1,2,3},{2,4,6}),6)"],
        ["rsq", 210, 196, "=ROUND(RSQ({2,4,6},{1,2,3}),6)"],
        ["steyx", 211, 196, "=ROUND(STEYX({2,4,6},{1,2,3}),6)"],
        ["skew", 212, 196, "=ROUND(SKEW({1,2,3,6}),6)"],
        ["skewP", 213, 196, "=ROUND(SKEW.P({1,2,3,6}),6)"],
        ["normDistCumulative", 214, 196, "=ROUND(NORM.DIST(0,0,1,TRUE),6)"],
        ["normDistDensity", 215, 196, "=ROUND(NORM.DIST(0,0,1,FALSE),6)"],
        ["normdistLegacy", 216, 196, "=ROUND(NORMDIST(0,0,1,FALSE),6)"],
        ["normSDist", 217, 196, "=ROUND(NORM.S.DIST(1,TRUE),6)"],
        ["normSDistDensity", 218, 196, "=ROUND(NORM.S.DIST(1,FALSE),6)"],
        ["normsdistLegacy", 219, 196, "=ROUND(NORMSDIST(1),6)"],
        ["normInv", 220, 196, "=ROUND(NORM.INV(0.5,10,2),6)"],
        ["normSInv", 221, 196, "=ROUND(NORM.S.INV(0.841344746),6)"],
        ["standardize", 222, 196, "=ROUND(STANDARDIZE(42,40,1.5),6)"],
        ["exponDistCumulative", 223, 196, "=ROUND(EXPON.DIST(0.2,10,TRUE),6)"],
        ["exponDistDensity", 224, 196, "=ROUND(EXPONDIST(0.2,10,FALSE),6)"],
        ["poissonDistMass", 225, 196, "=ROUND(POISSON.DIST(2,5,FALSE),6)"],
        ["poissonDistCumulative", 226, 196, "=ROUND(POISSON(2,5,TRUE),6)"],
        ["binomDistMass", 227, 196, "=ROUND(BINOM.DIST(6,10,0.5,FALSE),6)"],
        ["binomDistCumulative", 228, 196, "=ROUND(BINOMDIST(6,10,0.5,TRUE),6)"],
        ["binomInv", 229, 196, "=BINOM.INV(10,0.5,0.75)"],
        ["critbinom", 230, 196, "=CRITBINOM(10,0.5,0.75)"],
        ["fisher", 231, 196, "=ROUND(FISHER(0.5),6)"],
        ["fisherinv", 232, 196, "=ROUND(FISHERINV(FISHER(0.5)),6)"],
        ["modNegative", 233, 196, "=MOD(-3,2)"],
        ["power", 234, 196, "=POWER(2,3)"],
        ["sqrt", 235, 196, "=SQRT(16)"],
        ["lnExp", 236, 196, "=ROUND(LN(EXP(1)),6)"],
        ["logDefault", 237, 196, "=LOG(100)"],
        ["logBase", 238, 196, "=LOG(8,2)"],
        ["log10", 239, 196, "=LOG10(100)"],
        ["fact", 240, 196, "=FACT(5)"],
        ["factdouble", 241, 196, "=FACTDOUBLE(6)"],
        ["combin", 242, 196, "=COMBIN(5,2)"],
        ["combina", 243, 196, "=COMBINA(3,2)"],
        ["permut", 244, 196, "=PERMUT(5,2)"],
        ["permutationa", 245, 196, "=PERMUTATIONA(3,2)"],
        ["multinomial", 246, 196, "=MULTINOMIAL(2,3,4)"],
        ["gcd", 247, 196, "=GCD(24,36,60)"],
        ["lcm", 248, 196, "=LCM(4,6,10)"],
        ["exp", 249, 196, "=ROUND(EXP(1),6)"],
        ["pi", 250, 196, "=ROUND(PI(),6)"],
        ["sin", 251, 196, "=ROUND(SIN(PI()/2),6)"],
        ["cos", 252, 196, "=ROUND(COS(0),6)"],
        ["tan", 253, 196, "=ROUND(TAN(PI()/4),6)"],
        ["asin", 254, 196, "=ROUND(ASIN(1),6)"],
        ["acos", 255, 196, "=ROUND(ACOS(0),6)"],
        ["atan", 256, 196, "=ROUND(ATAN(1),6)"],
        ["atan2", 257, 196, "=ROUND(ATAN2(1,1),6)"],
        ["degrees", 258, 196, "=ROUND(DEGREES(PI()),6)"],
        ["radians", 259, 196, "=ROUND(RADIANS(180),6)"],
        ["sinh", 260, 196, "=ROUND(SINH(1),6)"],
        ["cosh", 261, 196, "=ROUND(COSH(1),6)"],
        ["tanh", 262, 196, "=ROUND(TANH(1),6)"],
        ["asinh", 263, 196, "=ROUND(ASINH(1),6)"],
        ["acosh", 264, 196, "=ROUND(ACOSH(2),6)"],
        ["atanh", 265, 196, "=ROUND(ATANH(0.5),6)"],
        ["cot", 266, 196, "=ROUND(COT(1),6)"],
        ["sec", 267, 196, "=ROUND(SEC(0),6)"],
        ["csc", 268, 196, "=ROUND(CSC(PI()/2),6)"],
        ["acot", 269, 196, "=ROUND(ACOT(1),6)"],
        ["coth", 270, 196, "=ROUND(COTH(1),6)"],
        ["sech", 271, 196, "=ROUND(SECH(1),6)"],
        ["csch", 272, 196, "=ROUND(CSCH(1),6)"],
        ["seriessum", 273, 196, "=SERIESSUM(2,0,1,{1,2,3})"],
        ["sumsq", 274, 196, "=SUMSQ(1,2,3)"],
        ["evenPositive", 275, 196, "=EVEN(2.1)"],
        ["evenNegative", 276, 196, "=EVEN(-2.1)"],
        ["oddPositive", 277, 196, "=ODD(2.1)"],
        ["oddNegative", 278, 196, "=ODD(-2.1)"],
        ["isoCeilingNegative", 279, 196, "=ISO.CEILING(-5.5,2)"],
        ["gamma", 280, 196, "=ROUND(GAMMA(5),6)"],
        ["gammaln", 281, 196, "=ROUND(GAMMALN(5),6)"],
        ["gammaDistCumulative", 282, 196, "=ROUND(GAMMA.DIST(2,3,2,TRUE),6)"],
        ["gammaDistDensity", 283, 196, "=ROUND(GAMMA.DIST(2,3,2,FALSE),6)"],
        ["gammaInv", 284, 196, "=ROUND(GAMMA.INV(GAMMA.DIST(2,3,2,TRUE),3,2),6)"],
        ["betaDistCumulative", 285, 196, "=ROUND(BETA.DIST(0.5,2,3,TRUE),6)"],
        ["betaDistDensity", 286, 196, "=ROUND(BETA.DIST(0.5,2,3,FALSE),6)"],
        ["betaInv", 287, 196, "=ROUND(BETA.INV(0.6875,2,3),6)"],
        ["chisqDistCumulative", 288, 196, "=ROUND(CHISQ.DIST(2,4,TRUE),6)"],
        ["chisqDistDensity", 289, 196, "=ROUND(CHISQ.DIST(2,4,FALSE),6)"],
        ["chisqInv", 290, 196, "=ROUND(CHISQ.INV(CHISQ.DIST(2,4,TRUE),4),6)"],
        ["chidistLegacy", 291, 196, "=ROUND(CHIDIST(2,4),6)"],
        ["weibullCumulative", 292, 196, "=ROUND(WEIBULL.DIST(2,3,4,TRUE),6)"],
        ["weibullDensity", 293, 196, "=ROUND(WEIBULL.DIST(2,3,4,FALSE),6)"],
        ["hypgeomMass", 294, 196, "=ROUND(HYPGEOM.DIST(1,4,3,10,FALSE),6)"],
        ["hypgeomCumulative", 295, 196, "=ROUND(HYPGEOM.DIST(1,4,3,10,TRUE),6)"],
        ["confidenceNorm", 296, 196, "=ROUND(CONFIDENCE.NORM(0.05,2,64),6)"],
        ["confidenceLegacy", 297, 196, "=ROUND(CONFIDENCE(0.05,2,64),6)"],
        ["confidenceT", 298, 196, "=ROUND(CONFIDENCE.T(0.05,2,64),6)"],
        ["gammadistLegacy", 299, 196, "=ROUND(GAMMADIST(2,3,2,TRUE),6)"],
        ["betadistLegacy", 300, 196, "=ROUND(BETADIST(0.5,2,3),6)"],
        ["chiinvLegacy", 301, 196, "=ROUND(CHIINV(CHIDIST(2,4),4),6)"],
        ["weibullLegacy", 302, 196, "=ROUND(WEIBULL(2,3,4,TRUE),6)"],
        ["hypgeomdistLegacy", 303, 196, "=ROUND(HYPGEOMDIST(1,4,3,10),6)"],
        ["fDistInvRoundTrip", 304, 196, "=ROUND(F.INV(F.DIST(2,5,10,TRUE),5,10),6)"],
        ["fDistRtInvRoundTrip", 305, 196, "=ROUND(F.INV.RT(F.DIST.RT(2,5,10),5,10),6)"],
        ["finvLegacyRoundTrip", 306, 196, "=ROUND(FINV(FDIST(2,5,10),5,10),6)"],
        ["tDistCumulative", 307, 196, "=ROUND(T.DIST(0,10,TRUE),6)"],
        ["tDistDensity", 308, 196, "=ROUND(T.DIST(0,10,FALSE),6)"],
        ["tInv", 309, 196, "=ROUND(T.INV(0.975,10),6)"],
        ["tInv2t", 310, 196, "=ROUND(T.INV.2T(0.05,10),6)"],
        ["tDist2tRoundTrip", 311, 196, "=ROUND(T.DIST.2T(T.INV.2T(0.05,10),10),6)"],
        ["tdistLegacyRoundTrip", 312, 196, "=ROUND(TDIST(TINV(0.05,10),10,2),6)"],
        ["tdistRtLegacy", 313, 196, "=ROUND(TDISTRT(T.INV.2T(0.05,10),10),6)"],
        ["lognormDistCumulative", 314, 196, "=ROUND(LOGNORM.DIST(1,0,1,TRUE),6)"],
        ["lognormDistDensity", 315, 196, "=ROUND(LOGNORM.DIST(1,0,1,FALSE),6)"],
        ["lognormInv", 316, 196, "=ROUND(LOGNORM.INV(0.5,0,1),6)"],
        ["lognormdistLegacy", 317, 196, "=ROUND(LOGNORMDIST(1,0,1),6)"],
        ["loginvLegacy", 318, 196, "=ROUND(LOGINV(0.5,0,1),6)"],
        ["lognorminvLegacy", 319, 196, "=ROUND(LOGNORMINV(0.5,0,1),6)"],
        ["negbinomDistMass", 320, 196, "=ROUND(NEGBINOM.DIST(2,3,0.5,FALSE),6)"],
        ["negbinomDistCumulative", 321, 196, "=ROUND(NEGBINOM.DIST(2,3,0.5,TRUE),6)"],
        ["negbinomdistLegacy", 322, 196, "=ROUND(NEGBINOMDIST(2,3,0.5),6)"],
        ["erf", 323, 196, "=ROUND(ERF(1),6)"],
        ["erfRange", 324, 196, "=ROUND(ERF(0,1),6)"],
        ["erfc", 325, 196, "=ROUND(ERFC(1),6)"],
        ["phi", 326, 196, "=ROUND(PHI(0),6)"],
        ["gauss", 327, 196, "=ROUND(GAUSS(1),6)"],
        ["zTest", 328, 196, "=ROUND(Z.TEST({1,2,3},2),6)"],
        ["ztestLegacy", 329, 196, "=ROUND(ZTEST({1,2,3},2),6)"],
        ["fTest", 330, 196, "=ROUND(F.TEST({1,2,3},{1,2,3}),6)"],
        ["ftestLegacy", 331, 196, "=ROUND(FTEST({1,2,3},{1,2,3}),6)"],
        ["chisqTest", 332, 196, "=ROUND(CHISQ.TEST({1,2;3,4},{1,2;3,4}),6)"],
        ["chitestLegacy", 333, 196, "=ROUND(CHITEST({1,2;3,4},{1,2;3,4}),6)"],
        ["tTest", 334, 196, "=ROUND(T.TEST({1,2,3},{1,2,3},2,2),6)"],
        ["ttestLegacy", 335, 196, "=ROUND(TTEST({1,2,3},{1,2,3},2,2),6)"],
        ["bitand", 336, 196, "=BITAND(13,7)"],
        ["bitor", 337, 196, "=BITOR(13,7)"],
        ["bitxor", 338, 196, "=BITXOR(13,7)"],
        ["bitlshift", 339, 196, "=BITLSHIFT(3,2)"],
        ["bitrshift", 340, 196, "=BITRSHIFT(8,2)"],
        ["dec2bin", 341, 196, '=DEC2BIN(10,8)'],
        ["dec2binNegative", 342, 196, '=DEC2BIN(-1)'],
        ["dec2oct", 343, 196, '=DEC2OCT(64)'],
        ["dec2hex", 344, 196, '=DEC2HEX(255,4)'],
        ["bin2decNegative", 345, 196, '=BIN2DEC("1111111111")'],
        ["bin2hexNegative", 346, 196, '=BIN2HEX("1111111111")'],
        ["bin2octPlaces", 347, 196, '=BIN2OCT("1010",4)'],
        ["oct2binPlaces", 348, 196, '=OCT2BIN("10",6)'],
        ["oct2decNegative", 349, 196, '=OCT2DEC("7777777777")'],
        ["oct2hexPlaces", 350, 196, '=OCT2HEX("10",4)'],
        ["hex2binPlaces", 351, 196, '=HEX2BIN("F",8)'],
        ["hex2octPlaces", 352, 196, '=HEX2OCT("F",4)'],
        ["hex2decNegative", 353, 196, '=HEX2DEC("FFFFFFFFFF")'],
        ["deltaEqual", 354, 196, "=DELTA(5,5)"],
        ["deltaDifferent", 355, 196, "=DELTA(5,4)"],
        ["gestepTrue", 356, 196, "=GESTEP(5,4)"],
        ["gestepFalse", 357, 196, "=GESTEP(3,4)"],
        ["complex", 358, 196, '=COMPLEX(3,4)'],
        ["complexJ", 359, 196, '=COMPLEX(3,-4,"j")'],
        ["imargument", 360, 196, '=ROUND(IMARGUMENT("1+i"),6)'],
        ["imexpReal", 361, 196, '=ROUND(IMREAL(IMEXP("1+i")),6)'],
        ["imexpImag", 362, 196, '=ROUND(IMAGINARY(IMEXP("1+i")),6)'],
        ["imlnReal", 363, 196, '=ROUND(IMREAL(IMLN("1+i")),6)'],
        ["imlnImag", 364, 196, '=ROUND(IMAGINARY(IMLN("1+i")),6)'],
        ["imlog10Real", 365, 196, '=ROUND(IMREAL(IMLOG10("1+i")),6)'],
        ["imlog2Real", 366, 196, '=ROUND(IMREAL(IMLOG2("1+i")),6)'],
        ["impowerImag", 367, 196, '=ROUND(IMAGINARY(IMPOWER("1+i",2)),6)'],
        ["imsqrtReal", 368, 196, '=ROUND(IMREAL(IMSQRT("3+4i")),6)'],
        ["imsqrtImag", 369, 196, '=ROUND(IMAGINARY(IMSQRT("3+4i")),6)'],
        ["imsinZero", 370, 196, '=ROUND(IMREAL(IMSIN("0+0i")),6)'],
        ["imcosZero", 371, 196, '=ROUND(IMREAL(IMCOS("0+0i")),6)'],
        ["imtanZero", 372, 196, '=ROUND(IMREAL(IMTAN("0+0i")),6)'],
        ["imsecReal", 373, 196, '=ROUND(IMREAL(IMSEC("1+0i")),6)'],
        ["imcscReal", 374, 196, '=ROUND(IMREAL(IMCSC("1+0i")),6)'],
        ["imcotReal", 375, 196, '=ROUND(IMREAL(IMCOT("1+0i")),6)'],
        ["imsinhZero", 376, 196, '=ROUND(IMREAL(IMSINH("0+0i")),6)'],
        ["imcoshZero", 377, 196, '=ROUND(IMREAL(IMCOSH("0+0i")),6)'],
        ["imsechZero", 378, 196, '=ROUND(IMREAL(IMSECH("0+0i")),6)'],
        ["imcschReal", 379, 196, '=ROUND(IMREAL(IMCSCH("1+0i")),6)'],
        ["sln", 380, 196, "=ROUND(SLN(1000,100,5),6)"],
        ["syd", 381, 196, "=ROUND(SYD(1000,100,5,1),6)"],
        ["ddb", 382, 196, "=ROUND(DDB(1000,100,5,1),6)"],
        ["db", 383, 196, "=ROUND(DB(1000,100,5,1,12),6)"],
        ["vdb", 427, 196, "=ROUND(VDB(1000,100,5,1,2,2,FALSE),6)"],
        ["amorLinc", 428, 196, "=ROUND(AMORLINC(1000,DATE(2023,1,1),DATE(2023,7,1),100,1,0.15,3),6)"],
        ["amorDegrc", 429, 196, "=ROUND(AMORDEGRC(1000,DATE(2023,1,1),DATE(2023,7,1),100,1,0.15,3),6)"],
        ["formulaTextSource", 430, 195, "=SUM(1,2)"],
        ["formulaText", 430, 196, "=FORMULATEXT(GM430)"],
        ["isFormula", 431, 196, "=ISFORMULA(GM430)"],
        ["isRef", 432, 196, "=ISREF(GM430)"],
        ["sheetNumber", 433, 196, "=SHEET(GM430)"],
        ["sheetsInReference", 434, 196, "=SHEETS(GM430)"],
        ["cellAddress", 435, 196, '=CELL("address",GM430)'],
        ["cellContents", 436, 196, '=CELL("contents",GM430)'],
        ["cellType", 437, 196, '=CELL("type",GM430)'],
        ["infoRecalc", 438, 196, '=INFO("recalc")'],
        ["erfcPrecise", 439, 196, "=ROUND(ERFC.PRECISE(1),6)"],
        ["besselI0", 440, 196, "=ROUND(BESSELI(1,0),6)"],
        ["besselJ0", 441, 196, "=ROUND(BESSELJ(1,0),6)"],
        ["besselK0", 442, 196, "=ROUND(BESSELK(1,0),6)"],
        ["besselY0", 443, 196, "=ROUND(BESSELY(1,0),6)"],
        ["erfPrecise", 444, 196, "=ROUND(ERF.PRECISE(1),6)"],
        ["gammalnPrecise", 445, 196, "=ROUND(GAMMALN.PRECISE(5),6)"],
        ["mmult", 446, 196, "=MMULT({1,2;3,4},{5,6;7,8})"],
        ["mdeterm", 448, 196, "=MDETERM({1,2;3,4})"],
        ["minverse", 449, 196, "=MINVERSE({1,2;3,4})"],
        ["munit", 451, 196, "=MUNIT(2)"],
        ["rowCurrent", 453, 196, "=ROW()"],
        ["columnCurrent", 454, 196, "=COLUMN()"],
        ["rowsRange", 455, 196, "=ROWS(GM446:GN447)"],
        ["columnsRange", 456, 196, "=COLUMNS(GM446:GN447)"],
        ["hyperlinkFriendly", 457, 196, '=HYPERLINK("https://example.com","リンク")'],
        ["tText", 458, 196, '=T("abc")'],
        ["tNumber", 459, 196, "=T(123)"],
        ["linest", 460, 196, "=LINEST({2;4;6},{1;2;3})"],
        ["linestStats", 461, 196, "=LINEST({1;2;4;8},{1;2;3;4},TRUE,TRUE)"],
        ["trend", 467, 196, "=TREND({2;4;6},{1;2;3},{4;5})"],
        ["logest", 470, 196, "=LOGEST({2;4;8},{1;2;3})"],
        ["growth", 471, 196, "=GROWTH({2;4;8},{1;2;3},{4;5})"],
        ["prob", 473, 196, "=PROB({1;2;3},{0.2;0.3;0.5},2,3)"],
        ["trimmean", 474, 196, "=TRIMMEAN({1,2,3,100},0.5)"],
        ["binomDistRange", 475, 196, "=ROUND(BINOM.DIST.RANGE(10,0.5,3,5),6)"],
        ["fvSchedule", 384, 196, "=ROUND(FVSCHEDULE(100,{0.1,0.2}),6)"],
        ["xnpv", 385, 196, "=ROUND(XNPV(0.1,{-1000,1100},{43831,44197}),6)"],
        ["dollarde", 386, 196, "=ROUND(DOLLARDE(1.02,16),6)"],
        ["dollarfr", 387, 196, "=ROUND(DOLLARFR(1.125,16),6)"],
        ["rri", 388, 196, "=ROUND(RRI(12,1000,2000),6)"],
        ["pduration", 389, 196, "=ROUND(PDURATION(0.05,1000,2000),6)"],
        ["ispmt", 390, 196, "=ROUND(ISPMT(0.1/12,1,36,8000),6)"],
        ["sqrtpi", 391, 196, "=ROUND(SQRTPI(2),6)"],
        ["tbillEq", 392, 196, "=ROUND(TBILLEQ(DATE(2023,1,1),DATE(2023,4,1),0.05),6)"],
        ["tbillPrice", 393, 196, "=ROUND(TBILLPRICE(DATE(2023,1,1),DATE(2023,4,1),0.05),6)"],
        ["tbillYield", 394, 196, "=ROUND(TBILLYIELD(DATE(2023,1,1),DATE(2023,4,1),98.75),6)"],
        ["kurt", 395, 196, "=ROUND(KURT({1,2,3,4,5}),6)"],
        ["rankAvg", 396, 196, "=RANK.AVG(2,{1,2,2,3,4},0)"],
        ["rankLegacy", 397, 196, "=RANK(2,{1,2,2,3,4},0)"],
        ["percentRankInc", 398, 196, "=ROUND(PERCENTRANK.INC({1,2,3,4,5},3.5,6),6)"],
        ["percentRankExc", 399, 196, "=ROUND(PERCENTRANK.EXC({1,2,3,4,5},3.5,6),6)"],
        ["percentileLegacy", 400, 196, "=PERCENTILE({1,2,3,4,5},0.5)"],
        ["quartileLegacy", 401, 196, "=QUARTILE({1,2,3,4,5},3)"],
        ["forecastLegacy", 402, 196, "=FORECAST(3,{10,20,30},{1,2,3})"],
        ["modeLegacy", 403, 196, "=MODE({1,2,2,3})"],
        ["xirr", 404, 196, "=ROUND(XIRR({-1000,1100},{DATE(2020,1,1),DATE(2021,1,1)}),6)"],
        ["disc", 405, 196, "=ROUND(DISC(DATE(2023,1,1),DATE(2023,7,1),97,100,3),6)"],
        ["priceDisc", 406, 196, "=ROUND(PRICEDISC(DATE(2023,1,1),DATE(2023,7,1),0.05,100,3),6)"],
        ["yieldDisc", 407, 196, "=ROUND(YIELDDISC(DATE(2023,1,1),DATE(2023,7,1),97,100,3),6)"],
        ["intrate", 408, 196, "=ROUND(INTRATE(DATE(2023,1,1),DATE(2023,7,1),970,1000,3),6)"],
        ["received", 409, 196, "=ROUND(RECEIVED(DATE(2023,1,1),DATE(2023,7,1),970,0.05,3),6)"],
        ["sumx2my2", 410, 196, "=SUMX2MY2({1,2,3},{4,5,6})"],
        ["sumx2py2", 411, 196, "=SUMX2PY2({1,2,3},{4,5,6})"],
        ["sumxmy2", 412, 196, "=SUMXMY2({1,2,3},{4,5,6})"],
        ["coupDayBs", 413, 196, "=COUPDAYBS(DATE(2023,1,15),DATE(2023,7,31),2,3)"],
        ["coupDays", 414, 196, "=COUPDAYS(DATE(2023,1,15),DATE(2023,7,31),2,3)"],
        ["coupDaysNc", 415, 196, "=COUPDAYSNC(DATE(2023,1,15),DATE(2023,7,31),2,3)"],
        ["coupNcd", 416, 196, "=COUPNCD(DATE(2023,1,15),DATE(2023,7,31),2,3)-DATE(2023,1,31)"],
        ["coupPcd", 417, 196, "=COUPPCD(DATE(2023,1,15),DATE(2023,7,31),2,3)-DATE(2022,7,31)"],
        ["coupNum", 418, 196, "=COUPNUM(DATE(2023,1,15),DATE(2023,7,31),2,3)"],
        ["price", 419, 196, "=ROUND(PRICE(DATE(2023,1,16),DATE(2025,1,15),0.05,0.04,100,2,3),6)"],
        ["yieldSecurity", 420, 196, "=ROUND(YIELD(DATE(2023,1,16),DATE(2025,1,15),0.05,PRICE(DATE(2023,1,16),DATE(2025,1,15),0.05,0.04,100,2,3),100,2,3),6)"],
        ["duration", 421, 196, "=ROUND(DURATION(DATE(2023,1,16),DATE(2025,1,15),0.05,0.04,2,3),6)"],
        ["mduration", 422, 196, "=ROUND(MDURATION(DATE(2023,1,16),DATE(2025,1,15),0.05,0.04,2,3),6)"],
        ["accrint", 423, 196, "=ROUND(ACCRINT(DATE(2023,1,1),DATE(2023,7,1),DATE(2023,4,1),0.05,1000,2,3,TRUE),6)"],
        ["accrintM", 424, 196, "=ROUND(ACCRINTM(DATE(2023,1,1),DATE(2023,7,1),0.05,1000,3),6)"],
        ["priceMat", 425, 196, "=ROUND(PRICEMAT(DATE(2023,2,1),DATE(2023,7,1),DATE(2023,1,1),0.05,0.04,3),6)"],
        ["yieldMat", 426, 196, "=ROUND(YIELDMAT(DATE(2023,2,1),DATE(2023,7,1),DATE(2023,1,1),0.05,PRICEMAT(DATE(2023,2,1),DATE(2023,7,1),DATE(2023,1,1),0.05,0.04,3),3),6)"],
        ["modeMult", 395, 198, "=MODE.MULT({1,2,2,3,3,4})"],
        ["frequency", 395, 200, "=FREQUENCY({1,2,3,4,5},{2,4})"],
      ];
      [
        [90, 115], [90, 116], [91, 115], [91, 116],
        [90, 121], [90, 122],
        [90, 135], [90, 136], [91, 135], [91, 136],
        [90, 138], [91, 138],
        [90, 140], [90, 141],
        [90, 143], [91, 143],
        [90, 147], [91, 147], [92, 147],
        [90, 149], [90, 150], [91, 149], [91, 150],
        [90, 152], [90, 153], [90, 154], [90, 155],
        [90, 158], [90, 159], [91, 158], [91, 159],
        [90, 161], [90, 162], [91, 161], [91, 162], [92, 161], [92, 162],
        [90, 164], [91, 164],
        [90, 166], [90, 167],
        [90, 169], [90, 170], [91, 169], [91, 170],
        [90, 172],
        [90, 174], [90, 175], [90, 176], [90, 177], [90, 178],
        [90, 180], [90, 181], [91, 180], [91, 181],
        [90, 183], [91, 183],
        [90, 185], [91, 185], [92, 185], [93, 185], [94, 185],
        [90, 187], [90, 188], [91, 187], [91, 188],
        [90, 190], [90, 191], [91, 190], [91, 191],
        [90, 193], [90, 194], [91, 193], [91, 194],
        [90, 196], [91, 196], [92, 196], [93, 196], [94, 196],
        [95, 196], [96, 196], [97, 196], [98, 196], [99, 196],
        [100, 196], [101, 196], [102, 196], [103, 196], [104, 196],
        [105, 196], [106, 196], [107, 196], [108, 196], [109, 196],
        [110, 196], [111, 196], [112, 196], [113, 196],
        [114, 196], [115, 196], [116, 196], [117, 196], [118, 196],
        [119, 196], [120, 196], [121, 196], [122, 196], [123, 196],
        [124, 196], [125, 196], [126, 196], [127, 196], [128, 196],
        [129, 196], [130, 196], [131, 196], [132, 196], [133, 196],
        [134, 196], [135, 196], [136, 196], [137, 196], [138, 196],
        [139, 196], [140, 196], [141, 196],
        [142, 196], [143, 196], [144, 196], [145, 196], [146, 196],
        [147, 196], [148, 196], [149, 196], [150, 196], [151, 196],
        [152, 196], [153, 196],
        [154, 196], [155, 196], [156, 196], [157, 196], [158, 196],
        [159, 196], [160, 196], [161, 196], [162, 196],
        [163, 196], [164, 196], [165, 196], [166, 196], [167, 196],
        [168, 196], [169, 196], [170, 196], [171, 196], [172, 196],
        [173, 196], [174, 196], [175, 196], [176, 196], [177, 196],
        [178, 196], [179, 196], [180, 196], [181, 196], [182, 196],
        [183, 196],
        [184, 196], [185, 196], [186, 196], [187, 196], [188, 196],
        [189, 196],
        [190, 196], [191, 196], [192, 196], [193, 196], [194, 196],
        [195, 196], [196, 196], [197, 196], [198, 196], [199, 196],
        [200, 196], [201, 196], [202, 196], [203, 196], [204, 196],
        [205, 196], [206, 196], [207, 196], [208, 196], [209, 196],
        [210, 196], [211, 196], [212, 196], [213, 196],
        [214, 196], [215, 196], [216, 196], [217, 196], [218, 196],
        [219, 196], [220, 196], [221, 196], [222, 196], [223, 196],
        [224, 196], [225, 196], [226, 196], [227, 196], [228, 196],
        [229, 196], [230, 196], [231, 196], [232, 196],
        [233, 196], [234, 196], [235, 196], [236, 196], [237, 196],
        [238, 196], [239, 196], [240, 196], [241, 196], [242, 196],
        [243, 196], [244, 196], [245, 196], [246, 196], [247, 196],
        [248, 196], [249, 196],
        [250, 196], [251, 196], [252, 196], [253, 196], [254, 196],
        [255, 196], [256, 196], [257, 196], [258, 196], [259, 196],
        [260, 196], [261, 196], [262, 196], [263, 196], [264, 196],
        [265, 196], [266, 196], [267, 196], [268, 196], [269, 196],
        [270, 196], [271, 196], [272, 196],
        [273, 196], [274, 196], [275, 196], [276, 196], [277, 196],
        [278, 196], [279, 196],
        [280, 196], [281, 196], [282, 196], [283, 196], [284, 196],
        [285, 196], [286, 196], [287, 196], [288, 196], [289, 196],
        [290, 196], [291, 196], [292, 196], [293, 196], [294, 196],
        [295, 196], [296, 196], [297, 196], [298, 196], [299, 196],
        [300, 196], [301, 196], [302, 196], [303, 196],
        [304, 196], [305, 196], [306, 196], [307, 196], [308, 196],
        [309, 196], [310, 196], [311, 196], [312, 196], [313, 196],
        [314, 196], [315, 196], [316, 196], [317, 196], [318, 196],
        [319, 196], [320, 196], [321, 196], [322, 196], [323, 196],
        [324, 196], [325, 196], [326, 196], [327, 196],
        [328, 196], [329, 196], [330, 196], [331, 196], [332, 196],
        [333, 196], [334, 196], [335, 196], [336, 196], [337, 196],
        [338, 196], [339, 196], [340, 196], [341, 196], [342, 196],
        [343, 196], [344, 196], [345, 196], [346, 196], [347, 196],
        [348, 196], [349, 196], [350, 196], [351, 196], [352, 196],
        [353, 196], [354, 196], [355, 196], [356, 196], [357, 196],
        [358, 196], [359, 196], [360, 196], [361, 196], [362, 196],
        [363, 196], [364, 196], [365, 196], [366, 196], [367, 196],
        [368, 196], [369, 196], [370, 196], [371, 196], [372, 196],
        [373, 196], [374, 196], [375, 196], [376, 196], [377, 196],
        [378, 196], [379, 196], [380, 196], [381, 196], [382, 196],
        [383, 196], [384, 196], [385, 196], [386, 196], [387, 196],
        [388, 196], [389, 196], [390, 196], [391, 196], [392, 196],
        [393, 196], [394, 196], [395, 196], [396, 196], [397, 196],
        [398, 196], [399, 196], [400, 196], [401, 196], [402, 196],
        [403, 196], [404, 196], [405, 196], [406, 196], [407, 196],
        [408, 196], [409, 196], [410, 196], [411, 196], [412, 196],
        [413, 196], [414, 196], [415, 196], [416, 196], [417, 196],
        [418, 196], [419, 196], [420, 196], [421, 196], [422, 196],
        [423, 196], [424, 196], [425, 196], [426, 196],
        [427, 196], [428, 196], [429, 196],
        [430, 195], [430, 196], [431, 196], [432, 196], [433, 196],
        [434, 196], [435, 196], [436, 196], [437, 196], [438, 196],
        [439, 196], [440, 196], [441, 196], [442, 196], [443, 196],
        [444, 196], [445, 196], [446, 196], [446, 197], [447, 196], [447, 197],
        [448, 196], [449, 196], [449, 197], [450, 196], [450, 197],
        [451, 196], [451, 197], [452, 196], [452, 197],
        [453, 196], [454, 196], [455, 196], [456, 196],
        [457, 196], [458, 196], [459, 196],
        [460, 196], [460, 197],
        [461, 196], [461, 197], [462, 196], [462, 197], [463, 196], [463, 197], [464, 196], [464, 197], [465, 196], [465, 197],
        [467, 196], [468, 196],
        [470, 196], [470, 197],
        [471, 196], [472, 196],
        [473, 196], [474, 196], [475, 196],
        [395, 198], [396, 198], [397, 198],
        [395, 200], [396, 200], [397, 200],
      ].forEach(([row, col]) => window.__webSheetDevSetCell?.(row, col, ""));
      formulas.forEach(([, row, col, formula]) => window.__webSheetDevSetCell?.(row, col, formula));
      const display = (row, col) => window.__webSheetDevCellDisplay?.("CHT", row, col);
      const values = Object.fromEntries(formulas.map(([key, row, col]) => [key, {
        display: display(row, col),
        engine: window.__webSheetDevFormulaEngineCell?.("CHT", row, col),
      }]));
      return {
        values,
        spills: {
          textsplit: [display(90, 115), display(90, 116), display(91, 115), display(91, 116)],
          textsplitIgnoreEmpty: [display(90, 121), display(90, 122)],
          makearray: [display(90, 135), display(90, 136), display(91, 135), display(91, 136)],
          byrow: [display(90, 138), display(91, 138)],
          bycol: [display(90, 140), display(90, 141)],
          map: [display(90, 143), display(91, 143)],
          scan: [display(90, 147), display(91, 147), display(92, 147)],
          sortby: [display(90, 149), display(90, 150), display(91, 149), display(91, 150)],
          textsplitMultiDelimiter: [display(90, 152), display(90, 153), display(90, 154), display(90, 155)],
          takeLastRows: [display(90, 158), display(90, 159), display(91, 158), display(91, 159)],
          dropLastRows: [display(90, 161), display(90, 162), display(91, 161), display(91, 162), display(92, 161), display(92, 162)],
          xlookupRowSpill: [display(90, 166), display(90, 167)],
          filterRows: [display(90, 169), display(90, 170), display(91, 169), display(91, 170)],
          transpose: [display(90, 180), display(90, 181), display(91, 180), display(91, 181)],
          chooseArray: [display(90, 183), display(91, 183)],
          indexColumn: [display(90, 185), display(91, 185), display(92, 185), display(93, 185), display(94, 185)],
          offsetRange: [display(90, 187), display(90, 188), display(91, 187), display(91, 188)],
          filterAndConditions: [display(90, 190), display(90, 191), display(91, 190), display(91, 191)],
          filterOrConditions: [display(90, 193), display(90, 194), display(91, 193), display(91, 194)],
          mmult: [display(446, 196), display(446, 197), display(447, 196), display(447, 197)],
          minverse: [display(449, 196), display(449, 197), display(450, 196), display(450, 197)],
          munit: [display(451, 196), display(451, 197), display(452, 196), display(452, 197)],
          linest: [display(460, 196), display(460, 197)],
          linestStats: [display(461, 196), display(461, 197), display(463, 196), display(463, 197), display(465, 196), display(465, 197)],
          trend: [display(467, 196), display(468, 196)],
          logest: [display(470, 196), display(470, 197)],
          growth: [display(471, 196), display(472, 196)],
          modeMult: [display(395, 198), display(396, 198)],
          frequency: [display(395, 200), display(396, 200), display(397, 200)],
        },
      };
    });
    const failures = [];
    if (result.spills.textsplit.join("|") !== "A|B|C|D") failures.push("TEXTSPLIT basic spill");
    if (result.spills.textsplitIgnoreEmpty.join("|") !== "A|B") failures.push("TEXTSPLIT omitted row delimiter + ignore_empty");
    if (result.values.textbefore.display !== "A-B") failures.push("TEXTBEFORE positive instance");
    if (result.values.textafter.display !== "C") failures.push("TEXTAFTER positive instance");
    if (result.values.textafterLast.display !== "C") failures.push("TEXTAFTER negative instance");
    if (result.values.textbeforeFallback.display !== "NF") failures.push("TEXTBEFORE if_not_found");
    if (result.values.aggregateSum.display !== "14") failures.push("AGGREGATE SUM");
    if (result.values.aggregateLarge.display !== "4") failures.push("AGGREGATE LARGE");
    if (result.values.aggregateMedian.display !== "2") failures.push("AGGREGATE MEDIAN");
    if (result.values.aggregateIgnoreError.display !== "14") failures.push("AGGREGATE ignore errors");
    if (result.values.ifs.display !== "2") failures.push("IFS first true branch");
    if (result.values.xor.display !== "TRUE") failures.push("XOR parity");
    if (result.values.letSimple.display !== "5") failures.push("LET simple binding");
    if (result.values.letMulti.display !== "10") failures.push("LET multiple binding");
    if (result.values.letRange.display !== "14") failures.push("LET range binding");
    if (result.spills.makearray.join("|") !== "2|3|3|4") failures.push("MAKEARRAY lambda spill");
    if (result.spills.byrow.join("|") !== "11|22") failures.push("BYROW lambda spill");
    if (result.spills.bycol.join("|") !== "3|30") failures.push("BYCOL lambda spill");
    if (result.spills.map.join("|") !== "2|4") failures.push("MAP lambda spill");
    if (result.values.reduce.display !== "14") failures.push("REDUCE lambda result");
    if (result.spills.scan.join("|") !== "1|3|5") failures.push("SCAN lambda spill");
    if (result.spills.sortby.join("|") !== "5|25|2|20") failures.push("SORTBY descending spill");
    if (result.spills.textsplitMultiDelimiter.join("|") !== "A|B|C|D") failures.push("TEXTSPLIT array delimiter spill");
    if (result.values.xmatchWildcard.display !== "1") failures.push("XMATCH wildcard");
    if (result.values.xmatchWildcardReverse.display !== "5") failures.push("XMATCH reverse wildcard");
    if (result.spills.takeLastRows.join("|") !== "4|5|5|25") failures.push("TAKE negative row spill");
    if (result.spills.dropLastRows.join("|") !== "1|10|2|20|2|15") failures.push("DROP negative row spill");
    if (result.values.xlookupExact.display !== "5") failures.push("XLOOKUP exact");
    if (result.values.xlookupReverseWildcard.display !== "5") failures.push("XLOOKUP reverse wildcard");
    if (result.spills.xlookupRowSpill.join("|") !== "4|5") failures.push("XLOOKUP row spill");
    if (result.spills.filterRows.join("|") !== "4|5|5|25") failures.push("FILTER row spill");
    if (result.values.filterEmptyFallback.display !== "NF") failures.push("FILTER if_empty fallback");
    if (result.values.xlookupApproxSmaller.display !== "5") failures.push("XLOOKUP approximate smaller");
    if (result.values.xlookupApproxLarger.display !== "5") failures.push("XLOOKUP approximate larger");
    if (result.values.xmatchBinaryAsc.display !== "4") failures.push("XMATCH ascending binary search");
    if (result.values.xmatchBinaryDesc.display !== "2") failures.push("XMATCH descending binary search");
    if (result.values.xlookupBinaryDesc.display !== "40") failures.push("XLOOKUP descending binary approximate");
    if (result.spills.transpose.join("|") !== "1|2|10|20") failures.push("TRANSPOSE spill display");
    if (result.spills.chooseArray.join("|") !== "10|20") failures.push("CHOOSE range spill display");
    if (result.spills.indexColumn.join("|") !== "10|20|15|5|25") failures.push("INDEX column spill display");
    if (result.spills.offsetRange.join("|") !== "2|20|2|15") failures.push("OFFSET range spill display");
    if (result.spills.filterAndConditions.join("|") !== "2|15|4|5") failures.push("FILTER AND conditions");
    if (result.spills.filterOrConditions.join("|") !== "1|10|5|25") failures.push("FILTER OR conditions");
    if (result.values.sumproductPairs.display !== "225") failures.push("SUMPRODUCT paired arrays");
    if (result.values.sumproductBoolean.display !== "65") failures.push("SUMPRODUCT boolean arithmetic");
    if (result.values.countblank.display !== "3") failures.push("COUNTBLANK blank cells");
    if (result.values.countif.display !== "2") failures.push("COUNTIF criteria");
    if (result.values.countifs.display !== "2") failures.push("COUNTIFS multiple criteria");
    if (result.values.sumif.display !== "35") failures.push("SUMIF criteria sum range");
    if (result.values.sumifs.display !== "20") failures.push("SUMIFS multiple criteria");
    if (result.values.averageif.display !== "17.5") failures.push("AVERAGEIF criteria average range");
    if (result.values.averageifs.display !== "10") failures.push("AVERAGEIFS multiple criteria");
    if (result.values.maxifs.display !== "15") failures.push("MAXIFS multiple criteria");
    if (result.values.minifs.display !== "5") failures.push("MINIFS multiple criteria");
    if (result.values.networkdays.display !== "5") failures.push("NETWORKDAYS inclusive weekdays");
    if (result.values.networkdaysHoliday.display !== "4") failures.push("NETWORKDAYS holiday");
    if (result.values.networkdaysIntl.display !== "6") failures.push("NETWORKDAYS.INTL weekend code");
    if (result.values.workday.display !== "3") failures.push("WORKDAY skips weekend");
    if (result.values.workdayIntl.display !== "1") failures.push("WORKDAY.INTL custom weekend");
    if (result.values.ceilingMathPositive.display !== "6") failures.push("CEILING.MATH positive");
    if (result.values.ceilingMathNegative.display !== "-4") failures.push("CEILING.MATH negative default");
    if (result.values.ceilingMathNegativeAway.display !== "-6") failures.push("CEILING.MATH negative away mode");
    if (result.values.imabs.display !== "5") failures.push("IMABS complex magnitude");
    if (result.values.datedifMonths.display !== "13") failures.push("DATEDIF full months");
    if (result.values.datedifYm.display !== "1") failures.push("DATEDIF months ignoring years");
    if (result.values.switchMatch.display !== "2") failures.push("SWITCH matched branch");
    if (result.values.switchDefault.display !== "9") failures.push("SWITCH default branch");
    if (result.values.edate.display !== "0") failures.push("EDATE month-end clamp");
    if (result.values.eomonth.display !== "0") failures.push("EOMONTH end of month");
    if (result.values.days.display !== "4") failures.push("DAYS date difference");
    if (result.values.days360.display !== "30") failures.push("DAYS360 30/360 count");
    if (result.values.yearfracBasis2.display !== "0.502778") failures.push("YEARFRAC actual/360");
    if (result.values.floorMathPositive.display !== "4") failures.push("FLOOR.MATH positive");
    if (result.values.floorMathNegative.display !== "-6") failures.push("FLOOR.MATH negative default");
    if (result.values.floorMathNegativeToward.display !== "-4") failures.push("FLOOR.MATH negative toward zero mode");
    if (result.values.mround.display !== "6") failures.push("MROUND nearest multiple");
    if (result.values.isblank.display !== "TRUE") failures.push("ISBLANK blank cell");
    if (result.values.iserror.display !== "TRUE") failures.push("ISERROR detects errors");
    if (result.values.iserr.display !== "TRUE") failures.push("ISERR detects non-NA errors");
    if (result.values.isna.display !== "TRUE") failures.push("ISNA detects #N/A");
    if (result.values.isnumber.display !== "TRUE") failures.push("ISNUMBER detects numbers");
    if (result.values.istext.display !== "TRUE") failures.push("ISTEXT detects strings");
    if (result.values.isnontext.display !== "TRUE") failures.push("ISNONTEXT detects non-text");
    if (result.values.islogical.display !== "TRUE") failures.push("ISLOGICAL detects booleans");
    if (result.values.iseven.display !== "TRUE") failures.push("ISEVEN truncates decimals");
    if (result.values.isodd.display !== "TRUE") failures.push("ISODD negative odd");
    if (result.values.nText.display !== "0") failures.push("N converts text to zero");
    if (result.values.nTrue.display !== "1") failures.push("N converts TRUE to one");
    if (result.values.imreal.display !== "3") failures.push("IMREAL real component");
    if (result.values.imaginary.display !== "4") failures.push("IMAGINARY imaginary component");
    if (result.values.imsum.display !== "4-2i") failures.push("IMSUM complex addition");
    if (result.values.improduct.display !== "-5+10i") failures.push("IMPRODUCT complex multiplication");
    if (result.values.imsub.display !== "2+2i") failures.push("IMSUB complex subtraction");
    if (result.values.imdiv.display !== "3+4i") failures.push("IMDIV complex division");
    if (result.values.imconjugate.display !== "3-4i") failures.push("IMCONJUGATE complex conjugate");
    if (result.values.datevalueIso.display !== "0") failures.push("DATEVALUE yyyy/m/d");
    if (result.values.datevalueJapanese.display !== "0") failures.push("DATEVALUE Japanese date");
    if (result.values.timevaluePm.display !== "0.78125") failures.push("TIMEVALUE PM text");
    if (result.values.hour.display !== "18") failures.push("HOUR extracts hour");
    if (result.values.minute.display !== "45") failures.push("MINUTE extracts minute");
    if (result.values.second.display !== "30") failures.push("SECOND extracts second");
    if (result.values.weekdayDefault.display !== "1") failures.push("WEEKDAY default Sunday start");
    if (result.values.weekdayMonday.display !== "1") failures.push("WEEKDAY Monday start");
    if (result.values.weeknumSunday.display !== "2") failures.push("WEEKNUM Sunday start");
    if (result.values.weeknumMonday.display !== "2") failures.push("WEEKNUM Monday start");
    if (result.values.isoweeknum.display !== "1") failures.push("ISOWEEKNUM ISO week");
    if (result.values.timevalueJapanesePm.display !== "13") failures.push("TIMEVALUE Japanese PM text");
    if (result.values.concatenate.display !== "ABC") failures.push("CONCATENATE joins text");
    if (result.values.replaceText.display !== "aXXde") failures.push("REPLACE text segment");
    if (result.values.proper.display !== "Hello World") failures.push("PROPER title casing");
    if (result.values.exactTrue.display !== "TRUE") failures.push("EXACT true");
    if (result.values.exactFalse.display !== "FALSE") failures.push("EXACT false");
    if (result.values.char.display !== "A") failures.push("CHAR code conversion");
    if (result.values.code.display !== "65") failures.push("CODE first character");
    if (result.values.clean.display !== "A") failures.push("CLEAN removes control characters");
    if (result.values.rept.display !== "ababab") failures.push("REPT repeats text");
    if (result.values.averagea.display !== "0.75") failures.push("AVERAGEA includes logical/text values");
    if (result.values.maxa.display !== "1") failures.push("MAXA includes logical/text values");
    if (result.values.mina.display !== "-5") failures.push("MINA includes logical/text values");
    if (result.values.ceilingPrecise.display !== "-4") failures.push("CEILING.PRECISE negative");
    if (result.values.floorPrecise.display !== "-6") failures.push("FLOOR.PRECISE negative");
    if (result.values.quotient.display !== "2") failures.push("QUOTIENT integer division");
    if (result.values.sign.display !== "-1") failures.push("SIGN negative");
    if (result.values.pmt.display !== "-85.61") failures.push("PMT loan payment");
    if (result.values.pv.display !== "1,000") failures.push("PV present value");
    if (result.values.fv.display !== "0") failures.push("FV future value");
    if (result.values.nper.display !== "12") failures.push("NPER period count");
    if (result.values.rate.display !== "5") failures.push("RATE loan rate");
    if (result.values.npv.display !== "173.55") failures.push("NPV discounted cash flow");
    if (result.values.irr.display !== "0.130662") failures.push("IRR internal rate");
    if (result.values.mirr.display !== "0.122497") failures.push("MIRR modified internal rate");
    if (result.values.ipmt.display !== "-4.17") failures.push("IPMT interest payment");
    if (result.values.ppmt.display !== "-81.44") failures.push("PPMT principal payment");
    if (result.values.cumipmt.display !== "-4.17") failures.push("CUMIPMT cumulative interest");
    if (result.values.cumprinc.display !== "-81.44") failures.push("CUMPRINC cumulative principal");
    if (result.values.effect.display !== "0.126825") failures.push("EFFECT effective rate");
    if (result.values.nominal.display !== "0.12") failures.push("NOMINAL nominal rate");
    if (result.values.roman.display !== "MCMXCIX") failures.push("ROMAN conversion");
    if (result.values.arabic.display !== "1,999") failures.push("ARABIC conversion");
    if (result.values.base.display !== "00FF") failures.push("BASE conversion");
    if (result.values.decimal.display !== "255") failures.push("DECIMAL conversion");
    if (result.values.unichar.display !== "☃") failures.push("UNICHAR code point");
    if (result.values.unicode.display !== "38,634") failures.push("UNICODE code point");
    if (result.values.stdevS.display !== "1") failures.push("STDEV.S sample standard deviation");
    if (result.values.stdevP.display !== "0.816497") failures.push("STDEV.P population standard deviation");
    if (result.values.stdevLegacy.display !== "1") failures.push("STDEV legacy sample standard deviation");
    if (result.values.stdevpLegacy.display !== "0.816497") failures.push("STDEVP legacy population standard deviation");
    if (result.values.varS.display !== "1") failures.push("VAR.S sample variance");
    if (result.values.varP.display !== "0.666667") failures.push("VAR.P population variance");
    if (result.values.vara.display !== "0.333333") failures.push("VARA logical/text variance");
    if (result.values.varpa.display !== "0.222222") failures.push("VARPA logical/text population variance");
    if (result.values.stdeva.display !== "0.57735") failures.push("STDEVA logical/text standard deviation");
    if (result.values.stdevpa.display !== "0.471405") failures.push("STDEVPA logical/text population standard deviation");
    if (result.values.correl.display !== "1") failures.push("CORREL paired arrays");
    if (result.values.covarianceS.display !== "2") failures.push("COVARIANCE.S paired arrays");
    if (result.values.covarianceP.display !== "1.333333") failures.push("COVARIANCE.P paired arrays");
    if (result.values.covar.display !== "1.333333") failures.push("COVAR legacy paired arrays");
    if (result.values.slope.display !== "2") failures.push("SLOPE regression gradient");
    if (result.values.avedev.display !== "1") failures.push("AVEDEV mean absolute deviation");
    if (result.values.devsq.display !== "2") failures.push("DEVSQ squared deviations");
    if (result.values.geomean.display !== "8") failures.push("GEOMEAN positive values");
    if (result.values.harmean.display !== "3") failures.push("HARMEAN positive values");
    if (result.values.pearson.display !== "1") failures.push("PEARSON paired arrays");
    if (result.values.rsq.display !== "1") failures.push("RSQ paired arrays");
    if (result.values.steyx.display !== "0") failures.push("STEYX regression standard error");
    if (result.values.skew.display !== "1.19034") failures.push("SKEW sample skewness");
    if (result.values.skewP.display !== "0.687243") failures.push("SKEW.P population skewness");
    if (result.values.normDistCumulative.display !== "0.5") failures.push("NORM.DIST cumulative");
    if (result.values.normDistDensity.display !== "0.398942") failures.push("NORM.DIST density");
    if (result.values.normdistLegacy.display !== "0.398942") failures.push("NORMDIST legacy density");
    if (result.values.normSDist.display !== "0.841345") failures.push("NORM.S.DIST cumulative");
    if (result.values.normSDistDensity.display !== "0.241971") failures.push("NORM.S.DIST density");
    if (result.values.normsdistLegacy.display !== "0.841345") failures.push("NORMSDIST legacy cumulative");
    if (result.values.normInv.display !== "10") failures.push("NORM.INV median");
    if (result.values.normSInv.display !== "1") failures.push("NORM.S.INV inverse");
    if (result.values.standardize.display !== "1.333333") failures.push("STANDARDIZE z score");
    if (result.values.exponDistCumulative.display !== "0.864665") failures.push("EXPON.DIST cumulative");
    if (result.values.exponDistDensity.display !== "1.353353") failures.push("EXPONDIST legacy density");
    if (result.values.poissonDistMass.display !== "0.084224") failures.push("POISSON.DIST mass");
    if (result.values.poissonDistCumulative.display !== "0.124652") failures.push("POISSON legacy cumulative");
    if (result.values.binomDistMass.display !== "0.205078") failures.push("BINOM.DIST mass");
    if (result.values.binomDistCumulative.display !== "0.828125") failures.push("BINOMDIST legacy cumulative");
    if (result.values.binomInv.display !== "6") failures.push("BINOM.INV threshold");
    if (result.values.critbinom.display !== "6") failures.push("CRITBINOM threshold");
    if (result.values.fisher.display !== "0.549306") failures.push("FISHER transform");
    if (result.values.fisherinv.display !== "0.5") failures.push("FISHERINV inverse");
    if (result.values.modNegative.display !== "1") failures.push("MOD follows divisor sign");
    if (result.values.power.display !== "8") failures.push("POWER exponent");
    if (result.values.sqrt.display !== "4") failures.push("SQRT square root");
    if (result.values.lnExp.display !== "1") failures.push("LN/EXP inverse");
    if (result.values.logDefault.display !== "2") failures.push("LOG default base");
    if (result.values.logBase.display !== "3") failures.push("LOG explicit base");
    if (result.values.log10.display !== "2") failures.push("LOG10 base ten");
    if (result.values.fact.display !== "120") failures.push("FACT factorial");
    if (result.values.factdouble.display !== "48") failures.push("FACTDOUBLE double factorial");
    if (result.values.combin.display !== "10") failures.push("COMBIN combinations");
    if (result.values.combina.display !== "6") failures.push("COMBINA repeated combinations");
    if (result.values.permut.display !== "20") failures.push("PERMUT permutations");
    if (result.values.permutationa.display !== "9") failures.push("PERMUTATIONA repeated permutations");
    if (result.values.multinomial.display !== "1,260") failures.push("MULTINOMIAL coefficient");
    if (result.values.gcd.display !== "12") failures.push("GCD common divisor");
    if (result.values.lcm.display !== "60") failures.push("LCM common multiple");
    if (result.values.exp.display !== "2.718282") failures.push("EXP natural exponent");
    if (result.values.pi.display !== "3.141593") failures.push("PI constant");
    if (result.values.sin.display !== "1") failures.push("SIN trigonometric sine");
    if (result.values.cos.display !== "1") failures.push("COS trigonometric cosine");
    if (result.values.tan.display !== "1") failures.push("TAN trigonometric tangent");
    if (result.values.asin.display !== "1.570796") failures.push("ASIN inverse sine");
    if (result.values.acos.display !== "1.570796") failures.push("ACOS inverse cosine");
    if (result.values.atan.display !== "0.785398") failures.push("ATAN inverse tangent");
    if (result.values.atan2.display !== "0.785398") failures.push("ATAN2 coordinate angle");
    if (result.values.degrees.display !== "180") failures.push("DEGREES radians to degrees");
    if (result.values.radians.display !== "3.141593") failures.push("RADIANS degrees to radians");
    if (result.values.sinh.display !== "1.175201") failures.push("SINH hyperbolic sine");
    if (result.values.cosh.display !== "1.543081") failures.push("COSH hyperbolic cosine");
    if (result.values.tanh.display !== "0.761594") failures.push("TANH hyperbolic tangent");
    if (result.values.asinh.display !== "0.881374") failures.push("ASINH inverse hyperbolic sine");
    if (result.values.acosh.display !== "1.316958") failures.push("ACOSH inverse hyperbolic cosine");
    if (result.values.atanh.display !== "0.549306") failures.push("ATANH inverse hyperbolic tangent");
    if (result.values.cot.display !== "0.642093") failures.push("COT cotangent");
    if (result.values.sec.display !== "1") failures.push("SEC secant");
    if (result.values.csc.display !== "1") failures.push("CSC cosecant");
    if (result.values.acot.display !== "0.785398") failures.push("ACOT inverse cotangent");
    if (result.values.coth.display !== "1.313035") failures.push("COTH hyperbolic cotangent");
    if (result.values.sech.display !== "0.648054") failures.push("SECH hyperbolic secant");
    if (result.values.csch.display !== "0.850918") failures.push("CSCH hyperbolic cosecant");
    if (result.values.seriessum.display !== "17") failures.push("SERIESSUM power series");
    if (result.values.sumsq.display !== "14") failures.push("SUMSQ square sum");
    if (result.values.evenPositive.display !== "4") failures.push("EVEN positive away from zero");
    if (result.values.evenNegative.display !== "-4") failures.push("EVEN negative away from zero");
    if (result.values.oddPositive.display !== "3") failures.push("ODD positive away from zero");
    if (result.values.oddNegative.display !== "-3") failures.push("ODD negative away from zero");
    if (result.values.isoCeilingNegative.display !== "-4") failures.push("ISO.CEILING negative toward positive infinity");
    if (result.values.gamma.display !== "24") failures.push("GAMMA value");
    if (result.values.gammaln.display !== "3.178054") failures.push("GAMMALN value");
    if (result.values.gammaDistCumulative.display !== "0.080301") failures.push("GAMMA.DIST cumulative");
    if (result.values.gammaDistDensity.display !== "0.09197") failures.push("GAMMA.DIST density");
    if (result.values.gammaInv.display !== "2") failures.push("GAMMA.INV inverse");
    if (result.values.betaDistCumulative.display !== "0.6875") failures.push("BETA.DIST cumulative");
    if (result.values.betaDistDensity.display !== "1.5") failures.push("BETA.DIST density");
    if (result.values.betaInv.display !== "0.5") failures.push("BETA.INV inverse");
    if (result.values.chisqDistCumulative.display !== "0.264241") failures.push("CHISQ.DIST cumulative");
    if (result.values.chisqDistDensity.display !== "0.18394") failures.push("CHISQ.DIST density");
    if (result.values.chisqInv.display !== "2") failures.push("CHISQ.INV inverse");
    if (result.values.chidistLegacy.display !== "0.735759") failures.push("CHIDIST legacy right tail");
    if (result.values.weibullCumulative.display !== "0.117503") failures.push("WEIBULL.DIST cumulative");
    if (result.values.weibullDensity.display !== "0.165468") failures.push("WEIBULL.DIST density");
    if (result.values.hypgeomMass.display !== "0.5") failures.push("HYPGEOM.DIST mass");
    if (result.values.hypgeomCumulative.display !== "0.666667") failures.push("HYPGEOM.DIST cumulative");
    if (result.values.confidenceNorm.display !== "0.489991") failures.push("CONFIDENCE.NORM interval");
    if (result.values.confidenceLegacy.display !== "0.489991") failures.push("CONFIDENCE legacy interval");
    if (result.values.confidenceT.display !== "0.499585") failures.push("CONFIDENCE.T interval");
    if (result.values.gammadistLegacy.display !== "0.080301") failures.push("GAMMADIST legacy cumulative");
    if (result.values.betadistLegacy.display !== "0.6875") failures.push("BETADIST legacy cumulative");
    if (result.values.chiinvLegacy.display !== "2") failures.push("CHIINV legacy inverse");
    if (result.values.weibullLegacy.display !== "0.117503") failures.push("WEIBULL legacy cumulative");
    if (result.values.hypgeomdistLegacy.display !== "0.5") failures.push("HYPGEOMDIST legacy mass");
    if (result.values.fDistInvRoundTrip.display !== "2") failures.push("F.DIST/F.INV round trip");
    if (result.values.fDistRtInvRoundTrip.display !== "2") failures.push("F.DIST.RT/F.INV.RT round trip");
    if (result.values.finvLegacyRoundTrip.display !== "2") failures.push("FDIST/FINV legacy round trip");
    if (result.values.tDistCumulative.display !== "0.5") failures.push("T.DIST cumulative");
    if (result.values.tDistDensity.display !== "0.389108") failures.push("T.DIST density");
    if (result.values.tInv.display !== "2.228139") failures.push("T.INV inverse");
    if (result.values.tInv2t.display !== "2.228139") failures.push("T.INV.2T inverse");
    if (result.values.tDist2tRoundTrip.display !== "0.05") failures.push("T.DIST.2T round trip");
    if (result.values.tdistLegacyRoundTrip.display !== "0.05") failures.push("TDIST/TINV legacy round trip");
    if (result.values.tdistRtLegacy.display !== "0.025") failures.push("TDISTRT legacy right tail");
    if (result.values.lognormDistCumulative.display !== "0.5") failures.push("LOGNORM.DIST cumulative");
    if (result.values.lognormDistDensity.display !== "0.398942") failures.push("LOGNORM.DIST density");
    if (result.values.lognormInv.display !== "1") failures.push("LOGNORM.INV inverse");
    if (result.values.lognormdistLegacy.display !== "0.5") failures.push("LOGNORMDIST legacy cumulative");
    if (result.values.loginvLegacy.display !== "1") failures.push("LOGINV legacy inverse");
    if (result.values.lognorminvLegacy.display !== "1") failures.push("LOGNORMINV legacy inverse");
    if (result.values.negbinomDistMass.display !== "0.1875") failures.push("NEGBINOM.DIST mass");
    if (result.values.negbinomDistCumulative.display !== "0.5") failures.push("NEGBINOM.DIST cumulative");
    if (result.values.negbinomdistLegacy.display !== "0.1875") failures.push("NEGBINOMDIST legacy mass");
    if (result.values.erf.display !== "0.842701") failures.push("ERF value");
    if (result.values.erfRange.display !== "0.842701") failures.push("ERF range value");
    if (result.values.erfc.display !== "0.157299") failures.push("ERFC value");
    if (result.values.erfcPrecise.display !== "0.157299") failures.push("ERFC.PRECISE value");
    if (result.values.besselI0.display !== "1.266066") failures.push("BESSELI order zero");
    if (result.values.besselJ0.display !== "0.765198") failures.push("BESSELJ order zero");
    if (result.values.besselK0.display !== "0.421024") failures.push("BESSELK order zero");
    if (result.values.besselY0.display !== "0.088257") failures.push("BESSELY order zero");
    if (result.values.erfPrecise.display !== "0.842701") failures.push("ERF.PRECISE value");
    if (result.values.gammalnPrecise.display !== "3.178054") failures.push("GAMMALN.PRECISE value");
    if (result.spills.mmult.join("|") !== "19|22|43|50") failures.push("MMULT spill result");
    if (result.values.mdeterm.display !== "-2") failures.push("MDETERM determinant");
    if (result.spills.minverse.join("|") !== "-2|1|1.5|-0.5") failures.push("MINVERSE spill result");
    if (result.spills.munit.join("|") !== "1|0|0|1") failures.push("MUNIT identity spill");
    if (result.values.rowCurrent.display !== "453") failures.push("ROW current formula row");
    if (result.values.columnCurrent.display !== "196") failures.push("COLUMN current formula column");
    if (result.values.rowsRange.display !== "2") failures.push("ROWS range height");
    if (result.values.columnsRange.display !== "2") failures.push("COLUMNS range width");
    if (result.values.hyperlinkFriendly.display !== "リンク") failures.push("HYPERLINK friendly text");
    if (result.values.tText.display !== "abc") failures.push("T text value");
    if (result.values.tNumber.display !== "") failures.push("T non-text empty value");
    if (result.spills.linest.join("|") !== "2|0") failures.push("LINEST coefficient spill");
    if (result.spills.linestStats[0] !== "2.3" || result.spills.linestStats[1] !== "-2") failures.push("LINEST stats coefficient row");
    if (result.spills.trend.join("|") !== "8|10") failures.push("TREND prediction spill");
    if (result.spills.logest.join("|") !== "2|1") failures.push("LOGEST exponential coefficient spill");
    if (result.spills.growth.join("|") !== "16|32") failures.push("GROWTH prediction spill");
    if (result.values.prob.display !== "0.8") failures.push("PROB probability range");
    if (result.values.trimmean.display !== "2.5") failures.push("TRIMMEAN trimmed average");
    if (result.values.binomDistRange.display !== "0.568359") failures.push("BINOM.DIST.RANGE interval probability");
    if (result.values.phi.display !== "0.398942") failures.push("PHI density");
    if (result.values.gauss.display !== "0.341345") failures.push("GAUSS value");
    if (result.values.zTest.display !== "0.5") failures.push("Z.TEST one-tail probability");
    if (result.values.ztestLegacy.display !== "0.5") failures.push("ZTEST legacy probability");
    if (result.values.fTest.display !== "1") failures.push("F.TEST identical variances");
    if (result.values.ftestLegacy.display !== "1") failures.push("FTEST legacy identical variances");
    if (result.values.chisqTest.display !== "1") failures.push("CHISQ.TEST identical ranges");
    if (result.values.chitestLegacy.display !== "1") failures.push("CHITEST legacy identical ranges");
    if (result.values.tTest.display !== "1") failures.push("T.TEST identical arrays");
    if (result.values.ttestLegacy.display !== "1") failures.push("TTEST legacy identical arrays");
    if (result.values.bitand.display !== "5") failures.push("BITAND bitwise and");
    if (result.values.bitor.display !== "15") failures.push("BITOR bitwise or");
    if (result.values.bitxor.display !== "10") failures.push("BITXOR bitwise xor");
    if (result.values.bitlshift.display !== "12") failures.push("BITLSHIFT left shift");
    if (result.values.bitrshift.display !== "2") failures.push("BITRSHIFT right shift");
    if (result.values.dec2bin.display !== "00001010") failures.push("DEC2BIN places");
    if (result.values.dec2binNegative.display !== "1111111111") failures.push("DEC2BIN negative twos complement");
    if (result.values.dec2oct.display !== "100") failures.push("DEC2OCT conversion");
    if (result.values.dec2hex.display !== "00FF") failures.push("DEC2HEX places");
    if (result.values.bin2decNegative.display !== "-1") failures.push("BIN2DEC negative twos complement");
    if (result.values.bin2hexNegative.display !== "FFFFFFFFFF") failures.push("BIN2HEX negative twos complement");
    if (result.values.bin2octPlaces.display !== "0012") failures.push("BIN2OCT places");
    if (result.values.oct2binPlaces.display !== "001000") failures.push("OCT2BIN places");
    if (result.values.oct2decNegative.display !== "-1") failures.push("OCT2DEC negative twos complement");
    if (result.values.oct2hexPlaces.display !== "0008") failures.push("OCT2HEX places");
    if (result.values.hex2binPlaces.display !== "00001111") failures.push("HEX2BIN places");
    if (result.values.hex2octPlaces.display !== "0017") failures.push("HEX2OCT places");
    if (result.values.hex2decNegative.display !== "-1") failures.push("HEX2DEC negative twos complement");
    if (result.values.deltaEqual.display !== "1") failures.push("DELTA equal");
    if (result.values.deltaDifferent.display !== "0") failures.push("DELTA different");
    if (result.values.gestepTrue.display !== "1") failures.push("GESTEP true");
    if (result.values.gestepFalse.display !== "0") failures.push("GESTEP false");
    if (result.values.complex.display !== "3+4i") failures.push("COMPLEX default suffix");
    if (result.values.complexJ.display !== "3-4j") failures.push("COMPLEX j suffix");
    if (result.values.imargument.display !== "0.785398") failures.push("IMARGUMENT angle");
    if (result.values.imexpReal.display !== "1.468694") failures.push("IMEXP real");
    if (result.values.imexpImag.display !== "2.287355") failures.push("IMEXP imaginary");
    if (result.values.imlnReal.display !== "0.346574") failures.push("IMLN real");
    if (result.values.imlnImag.display !== "0.785398") failures.push("IMLN imaginary");
    if (result.values.imlog10Real.display !== "0.150515") failures.push("IMLOG10 real");
    if (result.values.imlog2Real.display !== "0.5") failures.push("IMLOG2 real");
    if (result.values.impowerImag.display !== "2") failures.push("IMPOWER imaginary");
    if (result.values.imsqrtReal.display !== "2") failures.push("IMSQRT real");
    if (result.values.imsqrtImag.display !== "1") failures.push("IMSQRT imaginary");
    if (result.values.imsinZero.display !== "0") failures.push("IMSIN zero");
    if (result.values.imcosZero.display !== "1") failures.push("IMCOS zero");
    if (result.values.imtanZero.display !== "0") failures.push("IMTAN zero");
    if (result.values.imsecReal.display !== "1.850816") failures.push("IMSEC real");
    if (result.values.imcscReal.display !== "1.188395") failures.push("IMCSC real");
    if (result.values.imcotReal.display !== "0.642093") failures.push("IMCOT real");
    if (result.values.imsinhZero.display !== "0") failures.push("IMSINH zero");
    if (result.values.imcoshZero.display !== "1") failures.push("IMCOSH zero");
    if (result.values.imsechZero.display !== "1") failures.push("IMSECH zero");
    if (result.values.imcschReal.display !== "0.850918") failures.push("IMCSCH real");
    if (result.values.sln.display !== "180") failures.push("SLN depreciation");
    if (result.values.syd.display !== "300") failures.push("SYD depreciation");
    if (result.values.ddb.display !== "400") failures.push("DDB depreciation");
    if (result.values.db.display !== "369") failures.push("DB depreciation");
    if (result.values.vdb.display !== "240") failures.push("VDB variable declining balance");
    if (result.values.amorLinc.display !== "150") failures.push("AMORLINC accounting period depreciation");
    if (result.values.amorDegrc.display !== "305.265411") failures.push("AMORDEGRC coefficient depreciation");
    if (result.values.formulaTextSource.display !== "3") failures.push("FORMULATEXT source formula result");
    if (result.values.formulaText.display !== "=SUM(1,2)") failures.push("FORMULATEXT reference formula text");
    if (result.values.isFormula.display !== "TRUE") failures.push("ISFORMULA reference detection");
    if (result.values.isRef.display !== "TRUE") failures.push("ISREF reference detection");
    if (result.values.sheetNumber.display !== "1") failures.push("SHEET reference number");
    if (result.values.sheetsInReference.display !== "1") failures.push("SHEETS reference count");
    if (result.values.cellAddress.display !== "$GM$430") failures.push("CELL address info");
    if (result.values.cellContents.display !== "3") failures.push("CELL contents info");
    if (result.values.cellType.display !== "v") failures.push("CELL type info");
    if (result.values.infoRecalc.display !== "Automatic") failures.push("INFO recalc");
    if (result.values.fvSchedule.display !== "132") failures.push("FVSCHEDULE rates");
    if (result.values.xnpv.display !== "-0.26109") failures.push("XNPV dated cash flows");
    if (result.values.dollarde.display !== "1.125") failures.push("DOLLARDE fractional price");
    if (result.values.dollarfr.display !== "1.02") failures.push("DOLLARFR fractional price");
    if (result.values.rri.display !== "0.059463") failures.push("RRI equivalent rate");
    if (result.values.pduration.display !== "14.206699") failures.push("PDURATION periods");
    if (result.values.ispmt.display !== "-64.814815") failures.push("ISPMT interest");
    if (result.values.sqrtpi.display !== "2.506628") failures.push("SQRTPI value");
    if (result.values.tbillEq.display !== "0.051336") failures.push("TBILLEQ equivalent yield");
    if (result.values.tbillPrice.display !== "98.75") failures.push("TBILLPRICE price");
    if (result.values.tbillYield.display !== "0.050633") failures.push("TBILLYIELD yield");
    if (result.values.kurt.display !== "-1.2") failures.push("KURT sample kurtosis");
    if (result.values.rankAvg.display !== "3.5") failures.push("RANK.AVG tie average");
    if (result.values.rankLegacy.display !== "3") failures.push("RANK legacy rank");
    if (result.values.percentRankInc.display !== "0.625") failures.push("PERCENTRANK.INC interpolation");
    if (result.values.percentRankExc.display !== "0.583333") failures.push("PERCENTRANK.EXC interpolation");
    if (result.values.percentileLegacy.display !== "3") failures.push("PERCENTILE legacy percentile");
    if (result.values.quartileLegacy.display !== "4") failures.push("QUARTILE legacy quartile");
    if (result.values.forecastLegacy.display !== "30") failures.push("FORECAST legacy linear prediction");
    if (result.values.modeLegacy.display !== "2") failures.push("MODE legacy mode");
    if (result.spills.modeMult.join("|") !== "2|3") failures.push("MODE.MULT spill");
    if (result.spills.frequency.join("|") !== "2|2|1") failures.push("FREQUENCY spill");
    if (result.values.xirr.display !== "0.099714") failures.push("XIRR dated internal rate");
    if (result.values.disc.display !== "0.060497") failures.push("DISC discount rate");
    if (result.values.priceDisc.display !== "97.520548") failures.push("PRICEDISC price");
    if (result.values.yieldDisc.display !== "0.062368") failures.push("YIELDDISC yield");
    if (result.values.intrate.display !== "0.062368") failures.push("INTRATE rate");
    if (result.values.received.display !== "994.662172") failures.push("RECEIVED amount");
    if (result.values.sumx2my2.display !== "-63") failures.push("SUMX2MY2 pair squares");
    if (result.values.sumx2py2.display !== "91") failures.push("SUMX2PY2 pair squares");
    if (result.values.sumxmy2.display !== "27") failures.push("SUMXMY2 pair square differences");
    if (result.values.coupDayBs.display !== "168") failures.push("COUPDAYBS days before settlement");
    if (result.values.coupDays.display !== "182.5") failures.push("COUPDAYS coupon period days");
    if (result.values.coupDaysNc.display !== "16") failures.push("COUPDAYSNC days to next coupon");
    if (result.values.coupNcd.display !== "0") failures.push("COUPNCD next coupon date");
    if (result.values.coupPcd.display !== "0") failures.push("COUPPCD previous coupon date");
    if (result.values.coupNum.display !== "2") failures.push("COUPNUM coupon count");
    if (result.values.price.display !== "101.917813") failures.push("PRICE periodic security price");
    if (result.values.yieldSecurity.display !== "0.04") failures.push("YIELD periodic security yield");
    if (result.values.duration.display !== "1.921934") failures.push("DURATION Macaulay duration");
    if (result.values.mduration.display !== "1.884249") failures.push("MDURATION modified duration");
    if (result.values.accrint.display !== "12.328767") failures.push("ACCRINT accrued interest");
    if (result.values.accrintM.display !== "24.794521") failures.push("ACCRINTM accrued interest at maturity");
    if (result.values.priceMat.display !== "100.397445") failures.push("PRICEMAT maturity interest price");
    if (result.values.yieldMat.display !== "0.04") failures.push("YIELDMAT maturity interest yield");
    if (messages.length) failures.push(`browser errors: ${messages.join(" | ")}`);
    if (failures.length) {
      throw new Error(`Formula smoke failed: ${failures.join(", ")} ${JSON.stringify(result)}`);
    }
    console.log(JSON.stringify({ ok: true, result }, null, 2));
  } finally {
    await browser.close();
  }
}



async function runViewSmokeOnly() {
  e2eProgress("view smoke: start");
  const browser = await chromium.launch({
    headless: true,
    executablePath: chromePath,
  });
  e2eProgress("view smoke: browser launched");
  try {
    const context = await browser.newContext({
      viewport: { width: 1280, height: 720 },
      acceptDownloads: true,
      permissions: ["clipboard-read", "clipboard-write"],
    });
    e2eProgress("view smoke: context created");
    const page = await context.newPage();
    e2eProgress("view smoke: page created");
    const messages = [];
    page.on("console", (message) => {
      if (message.type() === "error" && !/Failed to load resource: the server responded with a status of 404/i.test(message.text())) {
        messages.push(`${message.type()}: ${message.text()}`);
      }
    });
    page.on("pageerror", (error) => {
      messages.push(`pageerror: ${error.stack || error.message}`);
    });
    await page.goto(`${appUrl}?t=${Date.now()}`, { waitUntil: "load" });
    e2eProgress("view smoke: app loaded");
    await page.setInputFiles("#fileInput", workbookPath);
    e2eProgress("view smoke: workbook file selected");
    await waitForUploadedWorkbook(page, workbookPath);
    e2eProgress("view smoke: workbook status loaded");
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT"), null, { timeout: 30000 });
    e2eProgress("view smoke: CHT sheet model available");
    await page.evaluate(() => window.__webSheetDevSetActiveSheet?.("CHT"));
    e2eProgress("view smoke: CHT sheet activation requested");
    await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "CHT", null, { timeout: 10000 });
    e2eProgress("view smoke: workbook loaded");
    await page.evaluate(() => {
      const sheet = window.__webSheetDevSheetModel?.("CHT");
      if (!sheet) return;
      sheet.pageLayout = { ...(sheet.pageLayout || {}), pageBreaks: { rows: [], cols: [] } };
    });
    await clickRibbonTab(page, "view");
    const viewState = () => page.evaluate(() => {
      const sheet = window.__webSheetDevSheetModel?.("CHT") || {};
      const grid = document.querySelector(".sheet-grid");
      const formulaBar = document.querySelector(".formula-bar");
      const commandState = (command) => {
        const button = document.querySelector(`[data-ribbon-command="${command}"]`);
        return {
          active: Boolean(button?.classList.contains("is-active")),
          checked: button?.getAttribute("aria-checked") || "",
          pressed: button?.getAttribute("aria-pressed") || "",
          disabled: Boolean(button?.disabled),
          title: button?.getAttribute("title") || "",
        };
      };
      return {
        viewMode: sheet.viewMode || "normal",
        zoomScale: Math.round(Number(sheet.zoomScale || 100)),
        showGridLines: sheet.showGridLines,
        showHeadings: sheet.showHeadings,
        showZeros: sheet.showZeros,
        showRuler: sheet.showRuler,
        showObjects: sheet.showObjects,
        showOutlineSymbols: sheet.showOutlineSymbols,
        rightToLeft: sheet.rightToLeft,
        splitPane: sheet.splitPane || null,
        freezePanes: sheet.freezePanes || null,
        rowCount: sheet.rowCount || 0,
        colCount: sheet.colCount || 0,
        formulaBarHidden: Boolean(formulaBar?.classList.contains("hidden")),
        gridClasses: grid ? [...grid.classList] : [],
        status: document.querySelector("#statusText")?.textContent || "",
        commands: Object.fromEntries([
          "view-gridlines",
          "view-headings",
          "view-formula-bar",
          "view-zero-values",
          "view-ruler",
          "view-objects",
          "view-outline-symbols",
          "view-right-to-left",
          "view-horizontal-scrollbar",
          "view-vertical-scrollbar",
          "view-sheet-tabs",
          "view-split",
          "view-freeze-panes",
          "view-100",
        ].map((command) => [command, commandState(command)])),
        bodyClasses: [...document.body.classList],
        workbookView: window.__webSheetDevWorkbookState?.().workbookView || null,
        mainWindowView: window.__webSheetDevWorkbookState?.().mainWindowView || null,
      };
    });
    const clickViewCommand = async (command) => {
      await page.click(`[data-ribbon-command="${command}"]`);
      await page.waitForTimeout(50);
      return viewState();
    };
    const ensureViewState = async (predicate, command) => {
      if (await page.evaluate(predicate)) return;
      await page.click(`[data-ribbon-command="${command}"]`);
      await page.waitForFunction(predicate, null, { timeout: 5000 });
    };
    await page.click('[data-ribbon-command="view-normal"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.viewMode === "normal", null, { timeout: 5000 });
    await ensureViewState(() => window.__webSheetDevSheetModel?.("CHT")?.showGridLines !== false, "view-gridlines");
    await ensureViewState(() => window.__webSheetDevSheetModel?.("CHT")?.showHeadings !== false, "view-headings");
    await ensureViewState(() => !document.querySelector(".formula-bar")?.classList.contains("hidden"), "view-formula-bar");
    await ensureViewState(() => window.__webSheetDevSheetModel?.("CHT")?.showZeros !== false, "view-zero-values");
    await ensureViewState(() => window.__webSheetDevSheetModel?.("CHT")?.showObjects !== false, "view-objects");
    await ensureViewState(() => window.__webSheetDevSheetModel?.("CHT")?.showOutlineSymbols !== false, "view-outline-symbols");
    await ensureViewState(() => window.__webSheetDevSheetModel?.("CHT")?.rightToLeft !== true, "view-right-to-left");
    await ensureViewState(() => !document.body.classList.contains("workbook-hide-horizontal-scrollbar"), "view-horizontal-scrollbar");
    await ensureViewState(() => !document.body.classList.contains("workbook-hide-vertical-scrollbar"), "view-vertical-scrollbar");
    await ensureViewState(() => !document.body.classList.contains("workbook-hide-sheet-tabs"), "view-sheet-tabs");
    await ensureViewState(() => !window.__webSheetDevSheetModel?.("CHT")?.splitPane, "view-split");
    await ensureViewState(() => !window.__webSheetDevSheetModel?.("CHT")?.freezePanes, "view-freeze-panes");
    await ensureViewState(() => Math.round(Number(window.__webSheetDevSheetModel?.("CHT")?.zoomScale || 100)) === 100, "view-100");
    await ensureViewState(() => window.__webSheetDevSheetModel?.("CHT")?.showRuler !== true, "view-ruler");
    const viewInitial = await viewState();
    e2eProgress("view smoke: initial state normalized");
    const viewCoreToggles = {};
    viewCoreToggles.gridlinesOff = await clickViewCommand("view-gridlines");
    viewCoreToggles.gridlinesOn = await clickViewCommand("view-gridlines");
    viewCoreToggles.headingsOff = await clickViewCommand("view-headings");
    viewCoreToggles.headingsOn = await clickViewCommand("view-headings");
    viewCoreToggles.formulaBarOff = await clickViewCommand("view-formula-bar");
    viewCoreToggles.formulaBarOn = await clickViewCommand("view-formula-bar");
    viewCoreToggles.zeroValuesOff = await clickViewCommand("view-zero-values");
    viewCoreToggles.zeroValuesOn = await clickViewCommand("view-zero-values");
    viewCoreToggles.objectsOff = await clickViewCommand("view-objects");
    viewCoreToggles.objectsOn = await clickViewCommand("view-objects");
    viewCoreToggles.outlineOff = await clickViewCommand("view-outline-symbols");
    viewCoreToggles.outlineOn = await clickViewCommand("view-outline-symbols");
    viewCoreToggles.rightToLeftOn = await clickViewCommand("view-right-to-left");
    viewCoreToggles.rightToLeftOff = await clickViewCommand("view-right-to-left");
    viewCoreToggles.horizontalScrollbarOff = await clickViewCommand("view-horizontal-scrollbar");
    viewCoreToggles.horizontalScrollbarOn = await clickViewCommand("view-horizontal-scrollbar");
    viewCoreToggles.verticalScrollbarOff = await clickViewCommand("view-vertical-scrollbar");
    viewCoreToggles.verticalScrollbarOn = await clickViewCommand("view-vertical-scrollbar");
    viewCoreToggles.sheetTabsOff = await clickViewCommand("view-sheet-tabs");
    viewCoreToggles.sheetTabsOn = await clickViewCommand("view-sheet-tabs");
    e2eProgress("view smoke: core visibility toggles");
    await page.evaluate(() => window.__webSheetDevSelectRange?.(4, 4, 4, 4));
    viewCoreToggles.splitOn = await clickViewCommand("view-split");
    viewCoreToggles.splitOff = await clickViewCommand("view-split");
    await page.evaluate(() => window.__webSheetDevSelectRange?.(4, 4, 4, 4));
    viewCoreToggles.freezeOn = await clickViewCommand("view-freeze-panes");
    viewCoreToggles.freezeOff = await clickViewCommand("view-freeze-panes");
    await page.click('[data-ribbon-command="view-zoom"]');
    await page.waitForSelector("#zoomDialog", { timeout: 5000 });
    await page.locator('input[name="zoomDialogOption"][value="custom"]').check();
    await page.locator("[data-zoom-custom-input]").fill("137");
    await page.click('[data-zoom-dialog-action="ok"]');
    await page.waitForFunction(() => Math.round(Number(window.__webSheetDevSheetModel?.("CHT")?.zoomScale || 0)) === 137, null, { timeout: 5000 });
    viewCoreToggles.zoomBefore100 = await viewState();
    viewCoreToggles.zoom100 = await clickViewCommand("view-100");
    await page.evaluate(() => window.__webSheetDevSelectRanges?.([
      { top: 116, left: 2, bottom: 117, right: 3 },
      { top: 120, left: 5, bottom: 121, right: 6 },
    ]));
    viewCoreToggles.zoomSelectionMulti = await page.evaluate(() => window.__webSheetDevZoomSelectionToFit?.());
    await page.click('[data-ribbon-command="view-100"]');
    await page.waitForFunction(() => Math.round(Number(window.__webSheetDevSheetModel?.("CHT")?.zoomScale || 0)) === 100, null, { timeout: 5000 });
    await page.evaluate(() => window.__webSheetDevSelectFullRows?.(4, 5));
    viewCoreToggles.zoomSelectionRows = await page.evaluate(() => window.__webSheetDevZoomSelectionToFit?.());
    await page.click('[data-ribbon-command="view-100"]');
    await page.waitForFunction(() => Math.round(Number(window.__webSheetDevSheetModel?.("CHT")?.zoomScale || 0)) === 100, null, { timeout: 5000 });
    await page.evaluate(() => window.__webSheetDevSelectFullColumns?.(2, 3));
    viewCoreToggles.zoomSelectionColumns = await page.evaluate(() => window.__webSheetDevZoomSelectionToFit?.());
    await page.click('[data-ribbon-command="view-100"]');
    await page.waitForFunction(() => Math.round(Number(window.__webSheetDevSheetModel?.("CHT")?.zoomScale || 0)) === 100, null, { timeout: 5000 });
    e2eProgress("view smoke: zoom commands");
    await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 1, 1, 1));
    await page.evaluate(() => {
      const host = document.querySelector(".sheet-host");
      if (host) {
        host.scrollTop = 0;
        host.scrollLeft = 0;
      }
      const horizontal = document.querySelector(".excel-horizontal-scrollbar");
      if (horizontal) horizontal.scrollLeft = 0;
    });
    await page.click('[data-ribbon-command="view-page-layout"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.viewMode === "page-layout", null, { timeout: 5000 });
    viewCoreToggles.pageLayout = await viewState();
    viewCoreToggles.rulerOn = await clickViewCommand("view-ruler");
    await page.click('[data-ribbon-command="view-normal"]');
    await page.waitForFunction(() => window.__webSheetDevSheetModel?.("CHT")?.viewMode === "normal", null, { timeout: 5000 });
    viewCoreToggles.normalRestored = await viewState();
    await page.click('[data-ribbon-command="view-page-break"]');
    await page.waitForFunction(() => document.querySelector(".sheet-grid")?.classList.contains("view-page-break"), null, { timeout: 5000 });
    await page.waitForSelector(".sheet-view-guide.is-page-break.is-automatic", { timeout: 5000 });
    e2eProgress("view smoke: page break view entered");
    const target = await page.evaluate(() => {
      const guide = document.querySelector(".sheet-view-guide.is-page-break.is-automatic");
      const rect = guide?.getBoundingClientRect();
      return {
        axis: guide?.dataset.pageBreakAxis || "",
        index: Number(guide?.dataset.pageBreakIndex || 0),
        x: rect ? rect.left + Math.max(2, Math.min(6, rect.width / 2 || 2)) : 0,
        y: rect ? rect.top + Math.max(2, Math.min(12, rect.height / 2 || 2)) : 0,
      };
    });
    await page.mouse.click(target.x, target.y, { button: "right" });
    await page.waitForSelector("#viewCommandMenu.is-page-break", { timeout: 5000 });
    await page.keyboard.press("I");
    await page.waitForFunction(({ axis, index }) => {
      const breaks = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {};
      const list = axis === "row" ? breaks.rows || [] : breaks.cols || [];
      return list.includes(index);
    }, target, { timeout: 5000 });
    await page.locator(`.sheet-view-guide.is-page-break.is-manual[data-page-break-axis="${target.axis}"][data-page-break-index="${target.index}"]`).first().click({ button: "right" });
    await page.waitForSelector("#viewCommandMenu.is-page-break", { timeout: 5000 });
    const menu = await page.evaluate(() => ({
      activeElement: document.activeElement?.outerHTML?.slice(0, 240) || "",
      actions: [...document.querySelectorAll("#viewCommandMenu [data-view-menu-action]")].map((button) => ({
        action: button.dataset.viewMenuAction || "",
        disabled: button.disabled,
        accessKey: button.dataset.accessKey || "",
        text: button.textContent.replace(/\s+/g, " ").trim(),
      })),
    }));
    await page.keyboard.press("R");
    await page.waitForFunction(({ axis, index }) => {
      const breaks = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {};
      const list = axis === "row" ? breaks.rows || [] : breaks.cols || [];
      return !list.includes(index);
    }, target, { timeout: 5000 });
    const result = await page.evaluate(({ axis, index }) => {
      const breaks = window.__webSheetDevSheetModel?.("CHT")?.pageLayout?.pageBreaks || {};
      const list = axis === "row" ? breaks.rows || [] : breaks.cols || [];
      return {
        breaks,
        removed: !list.includes(index),
        status: document.querySelector("#statusText")?.textContent || "",
      };
    }, target);
    e2eProgress("view smoke: page break context menu");
    await page.evaluate(() => window.__webSheetDevSetPageBreaks?.("CHT", [8, 11], [8, 13]));
    const pageBreakExportWithBreaks = await exportedPageBreakWorkbookState(
      await page.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
      "CHT",
    );
    e2eProgress("view smoke: page break export with breaks");
    await page.evaluate(() => window.__webSheetDevSetPageBreaks?.("CHT", [], []));
    const pageBreakExportCleared = await exportedPageBreakWorkbookState(
      await page.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
      "CHT",
    );
    e2eProgress("view smoke: page break export cleared");
    const pageBreakExport = {
      withBreaks: pageBreakExportWithBreaks,
      cleared: pageBreakExportCleared,
    };
    await page.evaluate(() => window.__webSheetDevSetDefinedNames?.([
      { name: "ExportKeep", refersTo: "=CHT!$A$1", scope: "workbook", comment: "keep" },
      { name: "ExportLocal", refersTo: "=CHT!$B$1", scope: "sheet", sheetIndex: 0 },
      { name: "ExportHidden", refersTo: "=CHT!$C$1", scope: "workbook", hidden: true },
    ]));
    const definedNameExportWithNames = await exportedDefinedNamesWorkbookState(
      await page.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
    );
    e2eProgress("view smoke: defined name export with names");
    await page.evaluate(() => window.__webSheetDevSetDefinedNames?.([
      { name: "ExportKeep", refersTo: "=CHT!$D$1", scope: "workbook", comment: "updated" },
    ]));
    const definedNameExportAfterDelete = await exportedDefinedNamesWorkbookState(
      await page.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
    );
    e2eProgress("view smoke: defined name export after delete");
    const definedNameExport = {
      withNames: definedNameExportWithNames,
      afterDelete: definedNameExportAfterDelete,
    };
    await page.evaluate(() => window.__webSheetDevSelectRanges?.([
      { top: 8, left: 2, bottom: 9, right: 4 },
      { top: 11, left: 6, bottom: 11, right: 7 },
    ]));
    const sheetViewSelectionExport = await exportedSheetViewWorkbookState(
      await page.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
      "CHT",
    );
    e2eProgress("view smoke: sheet view selection export");
    await page.evaluate(() => window.__webSheetDevSelectRange?.(4, 4, 4, 4));
    const sheetViewFreezeOn = await clickViewCommand("view-freeze-panes");
    const sheetViewFrozenSelectionExport = await exportedSheetViewWorkbookState(
      await page.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
      "CHT",
    );
    e2eProgress("view smoke: sheet view frozen export");
    const sheetViewFreezeOff = await clickViewCommand("view-freeze-panes");
    const sheetViewExport = {
      selection: sheetViewSelectionExport,
      frozenSelection: sheetViewFrozenSelectionExport,
      freezeOn: sheetViewFreezeOn,
      freezeOff: sheetViewFreezeOff,
    };
    const failures = [];
    if (messages.length) failures.push(`browser errors: ${messages.join(" | ")}`);
    if (viewInitial.rowCount !== 794 || viewInitial.colCount !== 81) failures.push("workbook dimension ignores styled XFD tail columns");
    if (!result.removed) failures.push("manual page break should be removed");
    if (viewCoreToggles.gridlinesOff.showGridLines !== false || !viewCoreToggles.gridlinesOff.gridClasses.includes("hide-gridlines") || viewCoreToggles.gridlinesOff.commands["view-gridlines"].checked !== "false") failures.push("gridlines off");
    if (viewCoreToggles.gridlinesOn.showGridLines !== true || viewCoreToggles.gridlinesOn.gridClasses.includes("hide-gridlines") || viewCoreToggles.gridlinesOn.commands["view-gridlines"].checked !== "true") failures.push("gridlines on");
    if (!String(viewInitial.commands["view-gridlines"].title || "").includes("非表示") || !String(viewCoreToggles.gridlinesOff.commands["view-gridlines"].title || "").includes("表示")) failures.push("gridlines command title");
    if (viewCoreToggles.headingsOff.showHeadings !== false || !viewCoreToggles.headingsOff.gridClasses.includes("hide-headings") || viewCoreToggles.headingsOff.commands["view-headings"].checked !== "false") failures.push("headings off");
    if (viewCoreToggles.headingsOn.showHeadings !== true || viewCoreToggles.headingsOn.gridClasses.includes("hide-headings") || viewCoreToggles.headingsOn.commands["view-headings"].checked !== "true") failures.push("headings on");
    if (!String(viewInitial.commands["view-headings"].title || "").includes("非表示") || !String(viewCoreToggles.headingsOff.commands["view-headings"].title || "").includes("表示")) failures.push("headings command title");
    if (!viewCoreToggles.formulaBarOff.formulaBarHidden || viewCoreToggles.formulaBarOff.commands["view-formula-bar"].checked !== "false") failures.push("formula bar off");
    if (viewCoreToggles.formulaBarOn.formulaBarHidden || viewCoreToggles.formulaBarOn.commands["view-formula-bar"].checked !== "true") failures.push("formula bar on");
    if (!String(viewInitial.commands["view-formula-bar"].title || "").includes("非表示") || !String(viewCoreToggles.formulaBarOff.commands["view-formula-bar"].title || "").includes("表示")) failures.push("formula bar command title");
    if (viewCoreToggles.zeroValuesOff.showZeros !== false || viewCoreToggles.zeroValuesOff.commands["view-zero-values"].checked !== "false") failures.push("zero values off");
    if (viewCoreToggles.zeroValuesOn.showZeros !== true || viewCoreToggles.zeroValuesOn.commands["view-zero-values"].checked !== "true") failures.push("zero values on");
    if (viewCoreToggles.objectsOff.showObjects !== false || viewCoreToggles.objectsOff.commands["view-objects"].checked !== "false") failures.push("objects off");
    if (viewCoreToggles.objectsOn.showObjects !== true || viewCoreToggles.objectsOn.commands["view-objects"].checked !== "true") failures.push("objects on");
    if (viewCoreToggles.outlineOff.showOutlineSymbols !== false || viewCoreToggles.outlineOff.commands["view-outline-symbols"].checked !== "false") failures.push("outline symbols off");
    if (viewCoreToggles.outlineOn.showOutlineSymbols !== true || viewCoreToggles.outlineOn.commands["view-outline-symbols"].checked !== "true") failures.push("outline symbols on");
    if (viewCoreToggles.rightToLeftOn.rightToLeft !== true || !viewCoreToggles.rightToLeftOn.gridClasses.includes("is-right-to-left")) failures.push("right-to-left on");
    if (viewCoreToggles.rightToLeftOff.rightToLeft === true || viewCoreToggles.rightToLeftOff.gridClasses.includes("is-right-to-left")) failures.push("right-to-left off");
    if (
      viewCoreToggles.horizontalScrollbarOff.mainWindowView?.showHorizontalScroll !== false ||
      !viewCoreToggles.horizontalScrollbarOff.bodyClasses.includes("workbook-hide-horizontal-scrollbar") ||
      viewCoreToggles.horizontalScrollbarOff.commands["view-horizontal-scrollbar"].checked !== "false"
    ) failures.push("horizontal scrollbar off");
    if (
      viewCoreToggles.horizontalScrollbarOn.mainWindowView?.showHorizontalScroll !== true ||
      viewCoreToggles.horizontalScrollbarOn.bodyClasses.includes("workbook-hide-horizontal-scrollbar") ||
      viewCoreToggles.horizontalScrollbarOn.commands["view-horizontal-scrollbar"].checked !== "true"
    ) failures.push("horizontal scrollbar on");
    if (
      viewCoreToggles.verticalScrollbarOff.mainWindowView?.showVerticalScroll !== false ||
      !viewCoreToggles.verticalScrollbarOff.bodyClasses.includes("workbook-hide-vertical-scrollbar") ||
      viewCoreToggles.verticalScrollbarOff.commands["view-vertical-scrollbar"].checked !== "false"
    ) failures.push("vertical scrollbar off");
    if (
      viewCoreToggles.verticalScrollbarOn.mainWindowView?.showVerticalScroll !== true ||
      viewCoreToggles.verticalScrollbarOn.bodyClasses.includes("workbook-hide-vertical-scrollbar") ||
      viewCoreToggles.verticalScrollbarOn.commands["view-vertical-scrollbar"].checked !== "true"
    ) failures.push("vertical scrollbar on");
    if (
      viewCoreToggles.sheetTabsOff.mainWindowView?.showSheetTabs !== false ||
      !viewCoreToggles.sheetTabsOff.bodyClasses.includes("workbook-hide-sheet-tabs") ||
      viewCoreToggles.sheetTabsOff.commands["view-sheet-tabs"].checked !== "false"
    ) failures.push("sheet tabs off");
    if (
      viewCoreToggles.sheetTabsOn.mainWindowView?.showSheetTabs !== true ||
      viewCoreToggles.sheetTabsOn.bodyClasses.includes("workbook-hide-sheet-tabs") ||
      viewCoreToggles.sheetTabsOn.commands["view-sheet-tabs"].checked !== "true"
    ) failures.push("sheet tabs on");
    if (viewCoreToggles.splitOn.splitPane?.row !== 4 || viewCoreToggles.splitOn.splitPane?.col !== 4 || viewCoreToggles.splitOn.commands["view-split"].pressed !== "true") failures.push("split pane on");
    if (viewCoreToggles.splitOff.splitPane || viewCoreToggles.splitOff.commands["view-split"].pressed !== "false") failures.push("split pane off");
    if (viewCoreToggles.freezeOn.freezePanes?.rows !== 3 || viewCoreToggles.freezeOn.freezePanes?.columns !== 3 || viewCoreToggles.freezeOn.commands["view-freeze-panes"].pressed !== "true") failures.push("freeze panes on");
    if (viewCoreToggles.freezeOff.freezePanes || viewCoreToggles.freezeOff.commands["view-freeze-panes"].pressed !== "false") failures.push("freeze panes off");
    if (viewCoreToggles.zoomBefore100.zoomScale !== 137 || viewCoreToggles.zoom100.zoomScale !== 100 || viewCoreToggles.zoom100.commands["view-100"].pressed !== "true") failures.push("zoom 100");
    if (viewCoreToggles.zoomSelectionMulti?.fitRange !== "B116:T121" || viewCoreToggles.zoomSelectionMulti?.result !== true || viewCoreToggles.zoomSelectionMulti?.after === 100) failures.push("zoom selection multi-range");
    if (
      viewCoreToggles.zoomSelectionRows?.range !== "4:5"
      || viewCoreToggles.zoomSelectionRows?.fitRangeObject?.top !== 4
      || viewCoreToggles.zoomSelectionRows?.fitRangeObject?.bottom !== 5
      || !Number.isFinite(Number(viewCoreToggles.zoomSelectionRows?.fitRangeObject?.right))
      || viewCoreToggles.zoomSelectionRows.fitRangeObject.right > viewInitial.colCount
      || viewCoreToggles.zoomSelectionRows.fitRangeObject.left < 1
    ) failures.push("zoom selection full rows");
    if (
      viewCoreToggles.zoomSelectionColumns?.range !== "B:C"
      || viewCoreToggles.zoomSelectionColumns?.fitRangeObject?.left !== 2
      || viewCoreToggles.zoomSelectionColumns?.fitRangeObject?.right !== 3
      || !Number.isFinite(Number(viewCoreToggles.zoomSelectionColumns?.fitRangeObject?.bottom))
      || viewCoreToggles.zoomSelectionColumns.fitRangeObject.bottom > viewInitial.rowCount
      || viewCoreToggles.zoomSelectionColumns.fitRangeObject.top < 1
    ) failures.push("zoom selection full columns");
    if (viewCoreToggles.pageLayout.viewMode !== "page-layout" || !viewCoreToggles.pageLayout.gridClasses.includes("view-page-layout")) failures.push("page layout view");
    if (viewCoreToggles.rulerOn.showRuler !== true || !viewCoreToggles.rulerOn.gridClasses.includes("show-ruler")) failures.push("ruler on");
    if (!String(viewInitial.commands["view-ruler"].title || "").includes("ページ レイアウト") || !String(viewCoreToggles.pageLayout.commands["view-ruler"].title || "").includes("表示") || !String(viewCoreToggles.rulerOn.commands["view-ruler"].title || "").includes("非表示")) failures.push("ruler command title");
    if (viewCoreToggles.normalRestored.viewMode !== "normal" || viewCoreToggles.normalRestored.gridClasses.includes("view-page-layout")) failures.push("normal view restored");
    if (
      pageBreakExport.withBreaks.rowBreaks.breakIndexes.join("|") !== "8|11" ||
      pageBreakExport.withBreaks.rowBreaks.count !== 2 ||
      pageBreakExport.withBreaks.rowBreaks.manualBreakCount !== 2 ||
      pageBreakExport.withBreaks.colBreaks.breakIndexes.join("|") !== "8|13" ||
      pageBreakExport.withBreaks.colBreaks.count !== 2 ||
      pageBreakExport.withBreaks.colBreaks.manualBreakCount !== 2
    ) failures.push("page break export writes row and column breaks");
    if (pageBreakExport.cleared.rowBreaks.present || pageBreakExport.cleared.colBreaks.present) failures.push("page break export clears stale breaks");
    if (
      definedNameExport.withNames.counts.ExportKeep !== 1 ||
      definedNameExport.withNames.counts.ExportLocal !== 1 ||
      definedNameExport.withNames.counts.ExportHidden !== 1 ||
      definedNameExport.withNames.entries.find((entry) => entry.name === "ExportLocal")?.localSheetId !== "0" ||
      definedNameExport.withNames.entries.find((entry) => entry.name === "ExportHidden")?.hidden !== true
    ) failures.push("defined name export writes workbook, sheet, and hidden names");
    if (
      definedNameExport.afterDelete.counts.ExportKeep !== 1 ||
      definedNameExport.afterDelete.counts.ExportLocal ||
      definedNameExport.afterDelete.counts.ExportHidden ||
      definedNameExport.afterDelete.entries.find((entry) => entry.name === "ExportKeep")?.text !== "CHT!$D$1"
    ) failures.push("defined name export removes stale names");
    if (
      sheetViewExport.selection.selections[0]?.activeCell !== "B8" ||
      sheetViewExport.selection.selections[0]?.sqref !== "B7:F9 B11:AP12"
    ) failures.push("sheet view export writes active cell and multi-range sqref");
    if (
      sheetViewExport.frozenSelection.pane?.state !== "frozen" ||
      sheetViewExport.frozenSelection.pane?.activePane !== "bottomRight" ||
      sheetViewExport.frozenSelection.selections[0]?.pane !== "bottomRight" ||
      sheetViewExport.frozenSelection.selections[0]?.activeCell !== "D4" ||
      sheetViewExport.frozenSelection.selections[0]?.sqref !== "B4:G4"
    ) failures.push("sheet view export writes frozen pane selection");
    if (failures.length) {
      throw new Error(`View smoke failed: ${failures.join(", ")} ${JSON.stringify({ viewInitial, viewCoreToggles, target, menu, result, pageBreakExport, definedNameExport, sheetViewExport })}`);
    }
    e2eProgress("view smoke: assertions passed");
    console.log(JSON.stringify({ ok: true, viewInitial, viewCoreToggles, target, menu, result, pageBreakExport, definedNameExport, sheetViewExport }, null, 2));
  } finally {
    await browser.close();
  }
}

export {
  exerciseCustomViewImportExport,
  runFormulaSmokeOnly,
  runViewSmokeOnly,
};
