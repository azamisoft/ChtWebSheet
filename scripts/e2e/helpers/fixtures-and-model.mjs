// Sequential E2E feature module.
import {
  ExcelJS,
  JSZip,
  WEBSHEET_MODEL_SCRIPT_RE,
  e2eProgressEnabled,
  gunzipSync,
  mkdtemp,
  os,
  path,
  readFile,
  workbookPath,
  writeFile,
} from "../../e2e-real-upload.mjs";


function looksLikeStandaloneWorkbookHtml(source = "") {
  return (
    source.includes('id="websheet-model"') &&
    source.includes("websheet-standalone.js") &&
    source.includes("websheet-standalone.css")
  );
}



function e2eProgress(label) {
  if (e2eProgressEnabled) console.error(`[websheet-e2e] ${new Date().toISOString()} ${label}`);
}



function jsonScriptPayload(value) {
  return JSON.stringify(value).replace(/</g, "\\u003c").replace(/\u2028/g, "\\u2028").replace(/\u2029/g, "\\u2029");
}



function websheetModelScript(payload) {
  return `<script id="websheet-model" type="application/json" data-cws-model="workbook" data-cws-format="CWS_HTML" data-cws-schema-version="1" data-cws-editable="true">${payload}</script>`;
}



function exportedStandaloneModelFromHtml(source) {
  const match = String(source || "").match(WEBSHEET_MODEL_SCRIPT_RE);
  if (!match) throw new Error("Cannot find standalone model script.");
  return decodeStandaloneModelPayload(JSON.parse(match[1] || "{}"));
}



function decodeStandaloneModelPayload(value) {
  let decoded = value;
  if (value && typeof value === "object" && value.encoding === "gzip-base64" && typeof value.payload === "string") {
    decoded = JSON.parse(gunzipSync(Buffer.from(value.payload, "base64")).toString("utf8"));
  }
  decodeCompressedImageAssets(decoded);
  return decoded;
}



function decodeCompressedImageAssets(model) {
  const assets = model?.assets?.images;
  if (!assets || typeof assets !== "object") return;
  Object.values(assets).forEach((asset) => {
    if (!asset || asset.dataEncoding !== "gzip-base64" || typeof asset.data !== "string") return;
    asset.data = gunzipSync(Buffer.from(asset.data, "base64")).toString("base64");
    delete asset.dataEncoding;
    delete asset.uncompressedByteLength;
  });
}



function replaceStandaloneModelPayload(source, payload) {
  return String(source || "").replace(
    WEBSHEET_MODEL_SCRIPT_RE,
    websheetModelScript(jsonScriptPayload(payload)),
  );
}



function replaceStandaloneModelScriptRaw(source, rawJson) {
  return String(source || "").replace(
    WEBSHEET_MODEL_SCRIPT_RE,
    websheetModelScript(rawJson),
  );
}



function xmlAttr(tag, name) {
  const escapedName = String(name || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = String(tag || "").match(new RegExp(`\\b${escapedName}="([^"]*)"`));
  return match ? match[1] : "";
}



async function worksheetXmlForSheetName(zip, sheetName) {
  const worksheetPath = await worksheetPathForSheetName(zip, sheetName);
  return worksheetPath ? zip.file(worksheetPath)?.async("text") || "" : "";
}



async function worksheetPathForSheetName(zip, sheetName) {
  const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("text");
  if (!workbookXml) return "";
  const sheetTags = [...String(workbookXml).matchAll(/<sheet\b[^>]*>/g)].map((match) => match[0]);
  const targetSheet = sheetTags.find((tag) => xmlAttr(tag, "name") === sheetName) || sheetTags[0] || "";
  const relationshipId = xmlAttr(targetSheet, "r:id");
  let target = "";
  if (relationshipId && relsXml) {
    const relationshipTag = [...String(relsXml).matchAll(/<Relationship\b[^>]*\/>/g)]
      .map((match) => match[0])
      .find((tag) => xmlAttr(tag, "Id") === relationshipId);
    target = xmlAttr(relationshipTag, "Target");
  }
  const worksheetPath = target
    ? `xl/${target.replace(/^\/?xl\//, "").replace(/^\//, "")}`.replace(/\/+/g, "/")
    : "xl/worksheets/sheet1.xml";
  return worksheetPath;
}



async function exportedPageBreakWorkbookState(base64, sheetName = "CHT") {
  const zip = await JSZip.loadAsync(Buffer.from(String(base64 || ""), "base64"));
  const worksheetXml = await worksheetXmlForSheetName(zip, sheetName);
  const blockState = (tag) => {
    const block = worksheetXml.match(new RegExp(`<${tag}\\b[\\s\\S]*?<\\/${tag}>|<${tag}\\b[^>]*\\/>`))?.[0] || "";
    const breakIndexes = [...block.matchAll(/<brk\b[^>]*\bid="(\d+)"/g)]
      .map((match) => Number(match[1]) + 1)
      .filter(Number.isFinite)
      .sort((a, b) => a - b);
    return {
      present: Boolean(block),
      count: Number(xmlAttr(block, "count")) || 0,
      manualBreakCount: Number(xmlAttr(block, "manualBreakCount")) || 0,
      breakIndexes,
    };
  };
  return {
    rowBreaks: blockState("rowBreaks"),
    colBreaks: blockState("colBreaks"),
  };
}



async function exportedOutlineWorkbookState(base64, sheetName = "Outline") {
  const zip = await JSZip.loadAsync(Buffer.from(String(base64 || ""), "base64"));
  const worksheetXml = await worksheetXmlForSheetName(zip, sheetName);
  const outlinePr = worksheetXml.match(/<outlinePr\b[^>]*\/>|<outlinePr\b[^>]*>[\s\S]*?<\/outlinePr>/)?.[0] || "";
  const rowTags = [...worksheetXml.matchAll(/<row\b[^>]*>/g)].map((match) => match[0]);
  const colTags = [...worksheetXml.matchAll(/<col\b[^>]*\/>/g)].map((match) => match[0]);
  const rows = rowTags
    .map((tag) => ({
      index: Number(xmlAttr(tag, "r")) || 0,
      hidden: xmlAttr(tag, "hidden") === "1",
      collapsed: xmlAttr(tag, "collapsed") === "1",
      outlineLevel: Number(xmlAttr(tag, "outlineLevel")) || 0,
    }))
    .filter((row) => row.index > 0);
  const columns = [];
  colTags.forEach((tag) => {
    const min = Math.max(1, Number(xmlAttr(tag, "min")) || 1);
    const max = Math.max(min, Number(xmlAttr(tag, "max")) || min);
    for (let index = min; index <= max; index += 1) {
      columns.push({
        index,
        hidden: xmlAttr(tag, "hidden") === "1",
        collapsed: xmlAttr(tag, "collapsed") === "1",
        outlineLevel: Number(xmlAttr(tag, "outlineLevel")) || 0,
      });
    }
  });
  return {
    rowHidden: rows.filter((row) => row.hidden).map((row) => row.index).sort((a, b) => a - b),
    rowCollapsed: rows.filter((row) => row.collapsed).map((row) => row.index).sort((a, b) => a - b),
    columnHidden: columns.filter((column) => column.hidden).map((column) => column.index).sort((a, b) => a - b),
    columnCollapsed: columns.filter((column) => column.collapsed).map((column) => column.index).sort((a, b) => a - b),
    outlinePr: {
      present: Boolean(outlinePr),
      summaryBelow: xmlAttr(outlinePr, "summaryBelow") || "",
      summaryRight: xmlAttr(outlinePr, "summaryRight") || "",
      applyStyles: xmlAttr(outlinePr, "applyStyles") || "",
    },
  };
}



async function exportedDefinedNamesWorkbookState(base64) {
  const zip = await JSZip.loadAsync(Buffer.from(String(base64 || ""), "base64"));
  const workbookXml = await zip.file("xl/workbook.xml")?.async("text") || "";
  const block = workbookXml.match(/<definedNames\b[^>]*>([\s\S]*?)<\/definedNames>/)?.[1] || "";
  const entries = [...block.matchAll(/<definedName\b[^>]*(?:\/>|>[\s\S]*?<\/definedName>)/g)]
    .map((match) => {
      const xml = match[0];
      const text = xml.includes("</definedName>")
        ? xml.replace(/^<definedName\b[^>]*>/, "").replace(/<\/definedName>$/, "")
        : "";
      return {
        name: xmlAttr(xml, "name"),
        localSheetId: xmlAttr(xml, "localSheetId"),
        hidden: xmlAttr(xml, "hidden") === "1",
        text,
      };
    });
  const names = entries.map((entry) => entry.name);
  const counts = names.reduce((acc, name) => {
    acc[name] = (acc[name] || 0) + 1;
    return acc;
  }, {});
  return {
    hasDefinedNamesBlock: Boolean(block),
    entries,
    names,
    counts,
    reservedNames: names.filter((name) => /^_xlnm\./i.test(name)),
  };
}



async function exportedSheetViewWorkbookState(base64, sheetName = "CHT") {
  const zip = await JSZip.loadAsync(Buffer.from(String(base64 || ""), "base64"));
  const worksheetXml = await worksheetXmlForSheetName(zip, sheetName);
  const sheetView = worksheetXml.match(/<sheetView\b[^>]*>[\s\S]*?<\/sheetView>|<sheetView\b[^>]*\/>/)?.[0] || "";
  const sheetViewOpen = sheetView.match(/^<sheetView\b[^>]*>/)?.[0] || "";
  const paneTag = sheetView.match(/<pane\b[^>]*(?:\/>|>[\s\S]*?<\/pane>)/)?.[0] || "";
  const selectionTags = [...sheetView.matchAll(/<selection\b[^>]*(?:\/>|>[\s\S]*?<\/selection>)/g)].map((match) => match[0]);
  return {
    hasSheetView: Boolean(sheetView),
    topLeftCell: xmlAttr(sheetViewOpen, "topLeftCell"),
    pane: paneTag
      ? {
          state: xmlAttr(paneTag, "state"),
          activePane: xmlAttr(paneTag, "activePane"),
          topLeftCell: xmlAttr(paneTag, "topLeftCell"),
          xSplit: xmlAttr(paneTag, "xSplit"),
          ySplit: xmlAttr(paneTag, "ySplit"),
        }
      : null,
    selections: selectionTags.map((tag) => ({
      pane: xmlAttr(tag, "pane"),
      activeCell: xmlAttr(tag, "activeCell"),
      sqref: xmlAttr(tag, "sqref"),
    })),
  };
}



async function clickRibbonTab(page, tabId) {
  await page.evaluate((id) => {
    document.body.classList.remove("is-websheet-fullscreen");
    document.querySelector(".ribbon")?.classList.remove("is-collapsed");
    document.querySelector(`[data-ribbon-tab="${id}"]`)?.click();
    document.body.classList.remove("is-websheet-fullscreen");
    document.querySelector(".ribbon")?.classList.remove("is-collapsed");
  }, tabId);
  await page.waitForFunction((id) => document.querySelector(`[data-ribbon-panel="${id}"]`)?.classList.contains("is-active"), tabId, { timeout: 5000 });
}



async function createPageBackgroundImage() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-bg-"));
  const filePath = path.join(dir, "page-background.png");
  const pngBase64 =
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAMUlEQVR4nGNkYGD4z0AEYBxVSFUwCkbBqBiFQyEaYDSQGgYGBjIyMiJqFgYAAJ09Ax8u3QKMAAAAAElFTkSuQmCC";
  await writeFile(filePath, Buffer.from(pngBase64, "base64"));
  return filePath;
}



async function createRichNativeShapeWorkbook(sourcePath) {
  const buffer = await readFile(sourcePath);
  const zip = await JSZip.loadAsync(buffer);
  const drawingPath = "xl/drawings/drawing2.xml";
  const drawingFile = zip.file(drawingPath);
  if (!drawingFile) return sourcePath;
  let drawingXml = await drawingFile.async("text");
  const richTextBody = `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr kumimoji="1" lang="ja-JP" sz="1600" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>赤Bold</a:t></a:r><a:r><a:rPr kumimoji="1" lang="ja-JP" sz="1400" i="1"><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>青Italic</a:t></a:r></a:p></xdr:txBody>`;
  drawingXml = drawingXml.replace(/<xdr:txBody>[\s\S]*?<\/xdr:txBody>/, richTextBody);
  drawingXml = drawingXml.replace(
    /<\/xdr:spPr><xdr:style>/,
    `<a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="5400000"><a:srgbClr val="808080"><a:alpha val="60000"/></a:srgbClr></a:outerShdw></a:effectLst></xdr:spPr><xdr:style>`,
  );
  drawingXml = drawingXml.replace(
    "<xdr:cNvCxnSpPr/>",
    '<xdr:cNvCxnSpPr><a:stCxn id="2" idx="9"/><a:endCxn id="3" idx="12"/></xdr:cNvCxnSpPr>',
  );
  if (!drawingXml.includes('name="グループ画像 8"')) {
    const groupedPictureXml = `<xdr:oneCellAnchor><xdr:from><xdr:col>15</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="2500000" cy="1200000"/><xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="20" name="グループ 8"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2500000" cy="1200000"/><a:chOff x="0" y="0"/><a:chExt cx="2500000" cy="1200000"/></a:xfrm></xdr:grpSpPr><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="21" name="グループ画像 8"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdGroupPicture"/><a:srcRect l="10000" t="5000" r="10000" b="5000"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm rot="1800000" flipH="1"><a:off x="100000" y="150000"/><a:ext cx="900000" cy="650000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="22" name="グループ四角形 9"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="1300000" y="240000"/><a:ext cx="850000" cy="560000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style></xdr:sp></xdr:grpSp><xdr:clientData/></xdr:oneCellAnchor>`;
    drawingXml = drawingXml.replace("</xdr:wsDr>", `${groupedPictureXml}</xdr:wsDr>`);
  }
  if (!drawingXml.includes('name="単独画像 10"')) {
    const singlePictureXml = `<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>200000</xdr:colOff><xdr:row>11</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="1200000" cy="720000"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="23" name="単独画像 10"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdSinglePicture"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm rot="900000"><a:off x="0" y="0"/><a:ext cx="1200000" cy="720000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="25400" cmpd="sng"><a:solidFill><a:srgbClr val="00B050"/></a:solidFill><a:prstDash val="dash"/></a:ln><a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="5400000"><a:srgbClr val="000000"><a:alpha val="45000"/></a:srgbClr></a:outerShdw></a:effectLst></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
    drawingXml = drawingXml.replace("</xdr:wsDr>", `${singlePictureXml}</xdr:wsDr>`);
  }
  if (!drawingXml.includes('name="グラフ 11"')) {
    const graphicFrameXml = `<xdr:oneCellAnchor><xdr:from><xdr:col>2</xdr:col><xdr:colOff>160000</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>80000</xdr:rowOff></xdr:from><xdr:ext cx="1800000" cy="900000"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="24" name="グラフ 11"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm rot="600000"><a:off x="0" y="0"/><a:ext cx="1800000" cy="900000"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdChartPlaceholder"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`;
    drawingXml = drawingXml.replace("</xdr:wsDr>", `${graphicFrameXml}</xdr:wsDr>`);
  }
  if (!drawingXml.includes('name="グラフ 12"')) {
    const lineGraphicFrameXml = `<xdr:oneCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>120000</xdr:colOff><xdr:row>17</xdr:row><xdr:rowOff>60000</xdr:rowOff></xdr:from><xdr:ext cx="1800000" cy="900000"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="25" name="グラフ 12"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="1800000" cy="900000"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdChartLine"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`;
    drawingXml = drawingXml.replace("</xdr:wsDr>", `${lineGraphicFrameXml}</xdr:wsDr>`);
  }
  if (!drawingXml.includes('name="グラフ 13"')) {
    const pieGraphicFrameXml = `<xdr:oneCellAnchor><xdr:from><xdr:col>8</xdr:col><xdr:colOff>120000</xdr:colOff><xdr:row>17</xdr:row><xdr:rowOff>60000</xdr:rowOff></xdr:from><xdr:ext cx="1800000" cy="900000"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="26" name="グラフ 13"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="1800000" cy="900000"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdChartPie"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`;
    drawingXml = drawingXml.replace("</xdr:wsDr>", `${pieGraphicFrameXml}</xdr:wsDr>`);
  }
  zip.file("xl/charts/chart1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:strRef><c:f>Sheet1!$D$1</c:f><c:strCache><c:pt idx="0"><c:v>売上推移</c:v></c:pt></c:strCache></c:strRef></c:tx><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1800" b="1"><a:solidFill><a:srgbClr val="990000"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr></c:title>
    <c:plotArea>
      <c:spPr><a:solidFill><a:srgbClr val="E2F0D9"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></a:ln></c:spPr>
      <c:barChart>
        <c:barDir val="bar"/>
        <c:grouping val="stacked"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$E$1</c:f><c:strCache><c:pt idx="0"><c:v>売上</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$1:$A$3</c:f><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$1:$B$3</c:f><c:numCache><c:pt idx="0"><c:v>120</c:v></c:pt><c:pt idx="1"><c:v>180</c:v></c:pt><c:pt idx="2"><c:v>150</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/><c:order val="1"/><c:tx><c:strRef><c:f>Sheet1!$F$1</c:f><c:strCache><c:pt idx="0"><c:v>利益</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="FFC000"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$1:$A$3</c:f><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$C$1:$C$3</c:f><c:numCache><c:pt idx="0"><c:v>45</c:v></c:pt><c:pt idx="1"><c:v>72</c:v></c:pt><c:pt idx="2"><c:v>66</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:dLbls><c:dLblPos val="outEnd"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr><c:showVal val="1"/></c:dLbls>
      </c:barChart>
      <c:catAx><c:axId val="10"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200" i="1"><a:solidFill><a:srgbClr val="00B050"/></a:solidFill></a:rPr><a:t>四半期</a:t></a:r></a:p></c:rich></c:tx></c:title><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="5B9BD5"/></a:solidFill><a:prstDash val="dashDot"/></a:ln></c:spPr></c:catAx>
      <c:valAx><c:axId val="20"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200" b="1"><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill></a:rPr><a:t>金額</a:t></a:r></a:p></c:rich></c:tx></c:title><c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="A5A5A5"/></a:solidFill><a:prstDash val="dot"/></a:ln></c:spPr></c:majorGridlines><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></a:ln></c:spPr></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="r"/></c:legend>
  </c:chart>
  <c:spPr><a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="C65911"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>
</c:chartSpace>`);
  zip.file("xl/charts/chart2.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>マーカー確認</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:lineChart>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/><c:tx><c:v>実績</c:v></c:tx>
          <c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="7030A0"/></a:solidFill></a:ln></c:spPr>
          <c:marker><c:symbol val="diamond"/><c:size val="9"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></c:spPr></c:marker>
          <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>4月</c:v></c:pt><c:pt idx="1"><c:v>5月</c:v></c:pt><c:pt idx="2"><c:v>6月</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>20</c:v></c:pt><c:pt idx="1"><c:v>35</c:v></c:pt><c:pt idx="2"><c:v>28</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/></c:legend>
  </c:chart>
</c:chartSpace>`);
  zip.file("xl/charts/chart3.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>構成比</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/><c:tx><c:v>構成</c:v></c:tx>
          <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="2"/><c:spPr><a:solidFill><a:srgbClr val="A5A5A5"/></a:solidFill></c:spPr></c:dPt>
          <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>A部門</c:v></c:pt><c:pt idx="1"><c:v>B部門</c:v></c:pt><c:pt idx="2"><c:v>C部門</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>40</c:v></c:pt><c:pt idx="1"><c:v>35</c:v></c:pt><c:pt idx="2"><c:v>25</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:dLbls><c:dLblPos val="bestFit"/><c:showCatName val="1"/><c:showVal val="1"/><c:showPercent val="1"/><c:separator>, </c:separator></c:dLbls>
      </c:pieChart>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/></c:legend>
  </c:chart>
</c:chartSpace>`);
  zip.file("xl/media/websheet-test-picture.png", Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAGElEQVR42mP8z8AARLJgYGBgYmAAAgAJAAOJlX5IAAAAAElFTkSuQmCC", "base64"));
  const drawingRelsPath = "xl/drawings/_rels/drawing2.xml.rels";
  const drawingRelsFile = zip.file(drawingRelsPath);
  let drawingRelsXml = drawingRelsFile
    ? await drawingRelsFile.async("text")
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  if (!drawingRelsXml.includes('Id="rIdGroupPicture"')) {
    drawingRelsXml = drawingRelsXml.replace(
      "</Relationships>",
      '<Relationship Id="rIdGroupPicture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>',
    );
    zip.file(drawingRelsPath, drawingRelsXml);
  }
  if (!drawingRelsXml.includes('Id="rIdSinglePicture"')) {
    drawingRelsXml = drawingRelsXml.replace(
      "</Relationships>",
      '<Relationship Id="rIdSinglePicture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/websheet-test-picture.png"/></Relationships>',
    );
    zip.file(drawingRelsPath, drawingRelsXml);
  }
  if (!drawingRelsXml.includes('Id="rIdChartPlaceholder"')) {
    drawingRelsXml = drawingRelsXml.replace(
      "</Relationships>",
      '<Relationship Id="rIdChartPlaceholder" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>',
    );
    zip.file(drawingRelsPath, drawingRelsXml);
  }
  if (!drawingRelsXml.includes('Id="rIdChartLine"')) {
    drawingRelsXml = drawingRelsXml.replace(
      "</Relationships>",
      '<Relationship Id="rIdChartLine" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/></Relationships>',
    );
    zip.file(drawingRelsPath, drawingRelsXml);
  }
  if (!drawingRelsXml.includes('Id="rIdChartPie"')) {
    drawingRelsXml = drawingRelsXml.replace(
      "</Relationships>",
      '<Relationship Id="rIdChartPie" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/></Relationships>',
    );
    zip.file(drawingRelsPath, drawingRelsXml);
  }
  const contentTypesPath = "[Content_Types].xml";
  const contentTypesFile = zip.file(contentTypesPath);
  if (contentTypesFile) {
    let contentTypesXml = await contentTypesFile.async("text");
    if (!contentTypesXml.includes('/xl/charts/chart1.xml')) {
      contentTypesXml = contentTypesXml.replace(
        "</Types>",
        '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
      );
      zip.file(contentTypesPath, contentTypesXml);
    }
    if (!contentTypesXml.includes('/xl/charts/chart2.xml')) {
      contentTypesXml = contentTypesXml.replace(
        "</Types>",
        '<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
      );
      zip.file(contentTypesPath, contentTypesXml);
    }
    if (!contentTypesXml.includes('/xl/charts/chart3.xml')) {
      contentTypesXml = contentTypesXml.replace(
        "</Types>",
        '<Override PartName="/xl/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
      );
      zip.file(contentTypesPath, contentTypesXml);
    }
  }
  zip.file(drawingPath, drawingXml);
  const output = await zip.generateAsync({ type: "nodebuffer" });
  const tempPath = path.join(os.tmpdir(), `websheet-native-rich-shape-${Date.now()}.xlsx`);
  await writeFile(tempPath, output);
  return tempPath;
}



async function waitForUploadedWorkbook(page, workbookPath) {
  const workbookName = path.basename(workbookPath);
  await page.waitForFunction(
    (name) => {
      const status = document.querySelector("#statusText")?.textContent || "";
      return document.title.includes(name) && /読み込みました|読み込み済み|已加载/.test(status);
    },
    workbookName,
    { timeout: 10000 },
  );
}



async function createFormulaUiWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-formula-ui-"));
  const outputPath = path.join(dir, "formula-ui.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("CHT");
  sheet.getCell("A1").value = "Formula UI";
  sheet.getCell("B9").value = "数量";
  sheet.getCell("C9").value = "Sales Amount";
  sheet.getCell("D9").value = "区分";
  sheet.getCell("E9").value = "税#率";
  sheet.getCell("F9").value = "O'Brien";
  sheet.getCell("G9").value = "端]数";
  sheet.getCell("H9").value = "#Data";
  sheet.getCell("I9").value = "@担当";
  sheet.getCell("J9").value = "#This Row";
  sheet.getCell("K9").value = "金額,税込";
  sheet.getCell("L9").value = "比率.%";
  sheet.getCell("M9").value = "A:B";
  sheet.getCell("N9").value = "費用?";
  sheet.getCell("O9").value = "売上$";
  sheet.getCell("P9").value = "A+B";
  sheet.getCell("Q9").value = "値(税)";
  sheet.getCell("R9").value = "担当,@税込";
  sheet.getCell("S9").value = "#All";
  sheet.getCell("T9").value = "#Headers";
  sheet.getCell("U9").value = "#Totals";
  sheet.getCell("V9").value = "始[数";
  sheet.getCell("B10").value = 4;
  sheet.getCell("C10").value = 20;
  sheet.getCell("D10").value = "B";
  sheet.getCell("E10").value = 5;
  sheet.getCell("F10").value = 7;
  sheet.getCell("G10").value = 3;
  sheet.getCell("H10").value = 11;
  sheet.getCell("I10").value = 2;
  sheet.getCell("J10").value = 4;
  sheet.getCell("K10").value = 13;
  sheet.getCell("L10").value = 1;
  sheet.getCell("M10").value = 21;
  sheet.getCell("N10").value = 2;
  sheet.getCell("O10").value = 3;
  sheet.getCell("P10").value = 11;
  sheet.getCell("Q10").value = 17;
  sheet.getCell("R10").value = 23;
  sheet.getCell("S10").value = 31;
  sheet.getCell("T10").value = 37;
  sheet.getCell("U10").value = 41;
  sheet.getCell("V10").value = 57;
  sheet.getCell("B11").value = 6;
  sheet.getCell("C11").value = 30;
  sheet.getCell("D11").value = "C";
  sheet.getCell("E11").value = 8;
  sheet.getCell("F11").value = 9;
  sheet.getCell("G11").value = 4;
  sheet.getCell("H11").value = 12;
  sheet.getCell("I11").value = 3;
  sheet.getCell("J11").value = 6;
  sheet.getCell("K11").value = 17;
  sheet.getCell("L11").value = 2;
  sheet.getCell("M11").value = 22;
  sheet.getCell("N11").value = 5;
  sheet.getCell("O11").value = 7;
  sheet.getCell("P11").value = 13;
  sheet.getCell("Q11").value = 19;
  sheet.getCell("R11").value = 29;
  sheet.getCell("S11").value = 43;
  sheet.getCell("T11").value = 47;
  sheet.getCell("U11").value = 53;
  sheet.getCell("V11").value = 59;
  sheet.getCell("B12").value = "合計";
  sheet.getCell("C12").value = 50;
  sheet.getCell("D12").value = "";
  sheet.getCell("E12").value = 13;
  sheet.getCell("F12").value = 16;
  sheet.getCell("G12").value = 7;
  sheet.getCell("H12").value = 99;
  sheet.getCell("I12").value = 88;
  sheet.getCell("J12").value = 77;
  sheet.getCell("K12").value = 301;
  sheet.getCell("L12").value = 9;
  sheet.getCell("M12").value = 99;
  sheet.getCell("N12").value = 70;
  sheet.getCell("O12").value = 80;
  sheet.getCell("P12").value = 90;
  sheet.getCell("Q12").value = 100;
  sheet.getCell("R12").value = 120;
  sheet.getCell("S12").value = 310;
  sheet.getCell("T12").value = 370;
  sheet.getCell("U12").value = 410;
  sheet.getCell("V12").value = 590;
  const second = workbook.addWorksheet("AI 製品紹介");
  second.getCell("A1").value = "Sheet reference";
  const apostropheSheet = workbook.addWorksheet("O'Brien Sheet");
  apostropheSheet.getCell("A1").value = "Apostrophe sheet reference";
  apostropheSheet.getCell("B2").value = 123;
  apostropheSheet.getCell("C3").value = 456;
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createMacroVbaSubsetWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-macro-vba-"));
  const outputPath = path.join(dir, "macro-vba-subset.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("CHT");
  sheet.getCell("A1").value = "Macro VBA subset";
  sheet.getCell("J6").value = 42;
  sheet.getCell("J6").numFmt = "0";
  sheet.getColumn(10).width = 12;
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createValidationMenuWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-validation-menu-"));
  const outputPath = path.join(dir, "validation-menu.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("入力規則");
  sheet.getCell("A1").value = "通常セル";
  sheet.getCell("A4").value = "注意スタイル";
  sheet.getCell("A5").value = "情報スタイル";
  sheet.getCell("A6").value = "警告なし";
  sheet.getCell("A12").value = 5;
  sheet.getCell("A13").value = 9;
  sheet.getCell("D1").value = "Alpha";
  sheet.getCell("D2").value = "Beta";
  sheet.getCell("D3").value = "Gamma";
  sheet.getCell("E1").value = "候補リスト";
  sheet.getCell("E2").value = "'入力規則'!$D$1:$D$3";
  sheet.getCell("E3").value = "部署2";
  sheet.getCell("A25").value = "ExcelJS が落とす属性";
  sheet.getCell("B25").value = "Alpha";
  sheet.getCell("F1").value = "部署1";
  sheet.getCell("G1").value = "部署2";
  sheet.getCell("F2").value = "札幌";
  sheet.getCell("F3").value = "仙台";
  sheet.getCell("F4").value = "福岡";
  sheet.getCell("G2").value = "Red";
  sheet.getCell("G3").value = "Blue";
  sheet.getCell("G4").value = "Yellow";
  sheet.getCell("B1").value = 20;
  sheet.getCell("B1").dataValidation = {
    type: "whole",
    operator: "between",
    formulae: [1, 10],
    allowBlank: true,
    showErrorMessage: true,
    errorTitle: "入力エラー",
    error: "1 から 10 までの整数を入力してください。",
  };
  sheet.getCell("B25").dataValidation = {
    type: "list",
    formulae: ['"Alpha,Beta,Gamma"'],
    allowBlank: true,
    showErrorMessage: true,
  };
  await workbook.xlsx.writeFile(outputPath);
  await addWorkbookDefinedNames(outputPath, [
    { name: "候補リスト", localSheetId: 0, refersTo: "'入力規則'!$D$1:$D$3" },
    { name: "動的候補", refersTo: "OFFSET('入力規則'!$D$1,0,0,COUNTA('入力規則'!$D:$D),1)" },
  ]);
  await patchValidationMenuWorkbookDataValidations(outputPath);
  return outputPath;
}



async function patchValidationMenuWorkbookDataValidations(filePath) {
  const buffer = await readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);
  const worksheetPath = "xl/worksheets/sheet1.xml";
  let xml = await zip.file(worksheetPath)?.async("text");
  if (!xml) return;
  xml = xml.replace(/<dataValidation\b(?=[^>]*\bsqref="B25")[^>]*>/, (tag) => {
    const clean = tag.replace(/\s+showDropDown="[^"]*"/, "").replace(/\s+imeMode="[^"]*"/, "");
    return clean.replace(/>$/, ' showDropDown="1" imeMode="disabled">');
  });
  const blankRangeRule = '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="C30:D31"><formula1>"One,Two"</formula1></dataValidation>';
  const referenceRule = '<dataValidation type="whole" operator="lessThanOrEqual" allowBlank="1" showErrorMessage="1" sqref="E30:E31"><formula1>A12</formula1></dataValidation>';
  const customRule = '<dataValidation type="custom" allowBlank="1" showErrorMessage="1" sqref="F30"><formula1>LEN(F30)&lt;=3</formula1></dataValidation>';
  const multiRangeRule = '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="H30 H32:I33"><formula1>"East,West"</formula1></dataValidation>';
  const decimalRule = '<dataValidation type="decimal" operator="greaterThan" allowBlank="0" showInputMessage="1" promptTitle="小数案内" prompt="1.5 より大きい値" showErrorMessage="1" errorStyle="warning" errorTitle="小数エラー" error="小数を確認してください" sqref="J30"><formula1>1.5</formula1></dataValidation>';
  const dateRule = '<dataValidation type="date" operator="between" allowBlank="1" showErrorMessage="1" sqref="K30"><formula1>45413</formula1><formula2>45443</formula2></dataValidation>';
  const timeRule = '<dataValidation type="time" operator="between" allowBlank="1" showErrorMessage="1" sqref="L30"><formula1>0.375</formula1><formula2>0.75</formula2></dataValidation>';
  const textLengthRule = '<dataValidation type="textLength" operator="equal" allowBlank="1" showErrorMessage="1" sqref="M30"><formula1>5</formula1></dataValidation>';
  if (!xml.includes('sqref="C30:D31"')) {
    xml = xml.replace("</dataValidations>", `${blankRangeRule}${referenceRule}${customRule}${multiRangeRule}${decimalRule}${dateRule}${timeRule}${textLengthRule}</dataValidations>`);
    xml = xml.replace(/<dataValidations\b([^>]*)\bcount="(\d+)"([^>]*)>/, (_match, before, count, after) =>
      `<dataValidations${before}count="${Number(count) + 8}"${after}>`,
    );
  }
  zip.file(worksheetPath, xml);
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
}



async function addWorkbookDefinedNames(filePath, definitions) {
  const buffer = await readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);
  const workbookPath = "xl/workbook.xml";
  const workbookFile = zip.file(workbookPath);
  if (!workbookFile) return;
  let xml = await workbookFile.async("text");
  const entries = definitions
    .filter((definition) => definition?.name && definition?.refersTo)
    .map((definition) => {
      const localSheetId = Number.isInteger(definition.localSheetId) ? ` localSheetId="${definition.localSheetId}"` : "";
      return `<definedName name="${xmlAttributeEscape(definition.name)}"${localSheetId}>${xmlTextEscape(definition.refersTo)}</definedName>`;
    })
    .join("");
  if (!entries) return;
  if (xml.includes("<definedNames>")) {
    xml = xml.replace("</definedNames>", `${entries}</definedNames>`);
  } else {
    xml = xml.replace("</workbook>", `<definedNames>${entries}</definedNames></workbook>`);
  }
  zip.file(workbookPath, xml);
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
}



function xmlAttributeEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}



function xmlTextEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}



async function createCustomViewWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-custom-view-"));
  const outputPath = path.join(dir, "custom-views.xlsx");
  const workbook = new ExcelJS.Workbook();
  const first = workbook.addWorksheet("CustomA");
  const second = workbook.addWorksheet("CustomB");
  first.getCell("A1").value = "Custom View A";
  first.getCell("B2").value = "Selection Start";
  first.getCell("C3").value = "Selection End";
  first.getCell("E5").value = 36922.375;
  first.getCell("E5").numFmt = "yyyy-mm-dd hh:mm";
  second.getCell("A1").value = "Custom View B";
  second.getCell("D4").value = "Active Cell";
  await workbook.xlsx.writeFile(outputPath);
  await addWorkbookCustomViews(outputPath);
  return outputPath;
}



async function addWorkbookCustomViews(filePath) {
  const customViewGuid = "{11111111-2222-3333-4444-555555555555}";
  const buffer = await readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);
  const workbookPath = "xl/workbook.xml";
  const workbookFile = zip.file(workbookPath);
  if (!workbookFile) return;
  let workbookXml = await workbookFile.async("text");
  const bookViews = `<bookViews><workbookView xWindow="10" yWindow="20" windowWidth="24000" windowHeight="12000" firstSheet="0" activeTab="0" visibility="visible"/><workbookView xWindow="410" yWindow="520" windowWidth="18000" windowHeight="9000" firstSheet="1" activeTab="1" visibility="hidden"/></bookViews>`;
  const workbookPr = `<workbookPr date1904="1" codeName="ThisWorkbook" defaultThemeVersion="164011" filterPrivacy="1"/>`;
  const calcPr = `<calcPr calcId="191029" calcMode="manual" refMode="R1C1" fullCalcOnLoad="0" calcOnSave="0" concurrentCalc="0" iterate="1" iterateCount="42" iterateDelta="0.0005"/>`;
  const customWorkbookViews = `<customWorkbookViews><customWorkbookView name="ImportedView" guid="${customViewGuid}" autoUpdate="1" mergeInterval="17" changesSavedWin="1" onlySync="1" personalView="1" includePrintSettings="1" includeHiddenRowCol="1" maximized="1" minimized="0" showHorizontalScroll="0" showVerticalScroll="0" showSheetTabs="0" showFormulaBar="0" xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000" activeSheetId="1"/></customWorkbookViews>`;
  workbookXml = workbookXml.replace(/<workbookPr\b[^>]*(?:\/>|>[\s\S]*?<\/workbookPr>)/, "");
  workbookXml = workbookXml.replace(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/, "");
  workbookXml = workbookXml.replace(/<bookViews\b[\s\S]*?<\/bookViews>/, "");
  workbookXml = workbookXml.replace(/<customWorkbookViews\b[\s\S]*?<\/customWorkbookViews>/, "");
  workbookXml = workbookXml.includes("<bookViews")
    ? workbookXml.replace("<bookViews", `${workbookPr}<bookViews`)
    : workbookXml.replace("<sheets", `${workbookPr}<sheets`);
  workbookXml = workbookXml.includes("<sheets")
    ? workbookXml.replace("<sheets", `${bookViews}<sheets`)
    : workbookXml.replace("</workbook>", `${bookViews}</workbook>`);
  workbookXml = workbookXml.includes("</workbook>")
    ? workbookXml.replace("</workbook>", `${calcPr}</workbook>`)
    : `${workbookXml}${calcPr}`;
  workbookXml = workbookXml.includes("<sheets")
    ? workbookXml.replace("<sheets", `${customWorkbookViews}<sheets`)
    : workbookXml.replace("</workbook>", `${customWorkbookViews}</workbook>`);
  zip.file(workbookPath, workbookXml);
  const sheetCustomViews = [
    `<customSheetViews><customSheetView guid="${customViewGuid}" scale="80" colorId="64" showPageBreaks="0" showFormulas="1" showGridLines="0" showRowCol="0" outlineSymbols="1" zeroValues="0" showWhiteSpace="1" rightToLeft="0" fitToPage="0" topLeftCell="D5"><selection activeCell="B2" sqref="B2:C3"/></customSheetView></customSheetViews>`,
    `<customSheetViews><customSheetView guid="${customViewGuid}" scale="125" colorId="64" showPageBreaks="0" showFormulas="0" showGridLines="1" showRowCol="1" outlineSymbols="1" zeroValues="1" showWhiteSpace="0" rightToLeft="1" fitToPage="0" topLeftCell="F8" view="pageLayout"><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="D4" sqref="D4"/></customSheetView></customSheetViews>`,
  ];
  for (const [index, block] of sheetCustomViews.entries()) {
    const sheetPath = `xl/worksheets/sheet${index + 1}.xml`;
    const file = zip.file(sheetPath);
    if (!file) continue;
    let xml = await file.async("text");
    xml = xml.replace(/<customSheetViews\b[\s\S]*?<\/customSheetViews>/, "");
    xml = xml.includes("</sheetViews>")
      ? xml.replace("</sheetViews>", `</sheetViews>${block}`)
      : xml.replace("</worksheet>", `${block}</worksheet>`);
    zip.file(sheetPath, xml);
  }
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
}



async function createAutoFilterMenuWorkbook(options = {}) {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-autofilter-menu-"));
  const outputPath = path.join(dir, "autofilter-menu.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Filter");
  sheet.addRow(["Name", "Score"]);
  sheet.addRow(["alpha", 5]);
  sheet.addRow(["Alpine", 12]);
  sheet.addRow(["Beta", 20]);
  sheet.addRow(["delta", 8]);
  if (options.includeMixedTextScore) sheet.addRow(["Zeta", "ZZZ"]);
  sheet.getCell("D1").value = "Other";
  sheet.getCell("E1").value = "Value";
  sheet.getCell("D2").value = "X";
  sheet.getCell("E2").value = 100;
  sheet.getCell("D3").value = "Y";
  sheet.getCell("E3").value = 200;
  sheet.getCell("B4").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFF0000" },
  };
  sheet.getCell("B3").font = { name: "Calibri", size: 11, color: { argb: "FFFF0000" } };
  sheet.getCell("G2").value = { formula: "B2>10" };
  sheet.getCell("H1").value = "既存";
  const criteria = workbook.addWorksheet("Criteria");
  criteria.getCell("A1").value = "Score";
  criteria.getCell("A2").value = ">10";
  criteria.getCell("B1").value = "Name";
  criteria.getCell("C1").value = "Score";
  criteria.getCell("B2").value = "a*";
  criteria.getCell("C2").value = ">=10";
  criteria.getCell("B3").value = "delta";
  criteria.getCell("D1").value = "Name";
  criteria.getCell("D2").value = "Al";
  workbook.addWorksheet("Extract");
  const blankHeader = workbook.addWorksheet("BlankHeader");
  blankHeader.getCell("A1").value = "空白見出し";
  blankHeader.getCell("C1").value = "数量";
  blankHeader.getCell("A2").value = "A";
  blankHeader.getCell("B2").value = "部署1";
  blankHeader.getCell("C2").value = 10;
  blankHeader.getCell("A3").value = "B";
  blankHeader.getCell("B3").value = "部署2";
  blankHeader.getCell("C3").value = 20;
  blankHeader.getCell("A4").value = "C";
  blankHeader.getCell("B4").value = "部署3";
  blankHeader.getCell("C4").value = 30;
  const tablePriority = workbook.addWorksheet("TablePriority");
  tablePriority.addTable({
    name: "PriorityTable",
    ref: "B2",
    headerRow: true,
    totalsRow: false,
    columns: [{ name: "品名" }, { name: "数量" }, { name: "担当" }],
    rows: [
      ["Alpha", 10, "A"],
      ["Beta", 20, "B"],
      ["Gamma", 30, "C"],
    ],
  });
  tablePriority.getCell("E2").value = "隣接列";
  tablePriority.getCell("E3").value = "外側1";
  tablePriority.getCell("E4").value = "外側2";
  tablePriority.getCell("E5").value = "外側3";
  workbook.definedNames.add("Filter!$A$1:$B$5", "Database");
  workbook.definedNames.add("Criteria!$D$1:$D$2", "Criteria");
  workbook.definedNames.add("Extract!$G$1", "Extract");
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createRelationshipsWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-relationships-"));
  const outputPath = path.join(dir, "relationships.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Relations");
  sheet.getCell("A1").value = "CustomerID";
  sheet.getCell("B1").value = "Name";
  sheet.getCell("A2").value = "C001";
  sheet.getCell("B2").value = "Azami";
  sheet.getCell("A3").value = "C002";
  sheet.getCell("B3").value = "Biz";
  sheet.getCell("A4").value = "C003";
  sheet.getCell("B4").value = "Trust";
  sheet.getCell("D1").value = "OrderID";
  sheet.getCell("E1").value = "CustomerID";
  sheet.getCell("F1").value = "Amount";
  sheet.getCell("D2").value = "O001";
  sheet.getCell("E2").value = "C001";
  sheet.getCell("F2").value = 10;
  sheet.getCell("D3").value = "O002";
  sheet.getCell("E3").value = "C001";
  sheet.getCell("F3").value = 15;
  sheet.getCell("D4").value = "O003";
  sheet.getCell("E4").value = "C002";
  sheet.getCell("F4").value = 20;
  sheet.getCell("D5").value = "O004";
  sheet.getCell("E5").value = "C003";
  sheet.getCell("F5").value = 30;
  sheet.getColumn(1).width = 14;
  sheet.getColumn(2).width = 12;
  sheet.getColumn(4).width = 12;
  sheet.getColumn(5).width = 14;
  sheet.getColumn(6).width = 10;
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createAutoFilterBlankWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-autofilter-blank-"));
  const outputPath = path.join(dir, "autofilter-blank.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("BlankFilter");
  sheet.addRow(["Memo"]);
  sheet.addRow(["Filled"]);
  sheet.addRow([]);
  sheet.addRow([]);
  sheet.addRow(["Note"]);
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createImportedAutoFilterWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-imported-autofilter-"));
  const outputPath = path.join(dir, "imported-autofilter.xlsx");
  const workbook = new ExcelJS.Workbook();
  const valueSheet = workbook.addWorksheet("値フィルター");
  valueSheet.addRow(["Name", "Score"]);
  valueSheet.addRow(["Alpha", 10]);
  valueSheet.addRow(["Beta", 20]);
  valueSheet.addRow(["Alpha", 30]);
  valueSheet.addRow(["Gamma", 40]);
  valueSheet.autoFilter = "A1:B5";
  const customSheet = workbook.addWorksheet("条件フィルター");
  customSheet.addRow(["Name", "Score"]);
  customSheet.addRow(["A", 5]);
  customSheet.addRow(["B", 12]);
  customSheet.addRow(["C", 20]);
  customSheet.addRow(["D", 8]);
  customSheet.autoFilter = "A1:B5";
  const topSheet = workbook.addWorksheet("トップテン");
  topSheet.addRow(["Name", "Score"]);
  topSheet.addRow(["A", 5]);
  topSheet.addRow(["B", 12]);
  topSheet.addRow(["C", 20]);
  topSheet.addRow(["D", 8]);
  topSheet.autoFilter = "A1:B5";
  const averageSheet = workbook.addWorksheet("平均");
  averageSheet.addRow(["Name", "Score"]);
  averageSheet.addRow(["A", 5]);
  averageSheet.addRow(["B", 10]);
  averageSheet.addRow(["C", 20]);
  averageSheet.addRow(["D", 25]);
  averageSheet.autoFilter = "A1:B5";
  const dateSheet = workbook.addWorksheet("日付グループ");
  dateSheet.addRow(["Date", "Amount"]);
  dateSheet.addRow(["2026/05/20", 10]);
  dateSheet.addRow(["2026/05/25", 20]);
  dateSheet.addRow(["2026/06/01", 30]);
  dateSheet.addRow(["2025/12/31", 40]);
  dateSheet.autoFilter = "A1:B5";
  const blankSheet = workbook.addWorksheet("空白インポート");
  blankSheet.addRow(["Memo"]);
  blankSheet.addRow(["Filled"]);
  blankSheet.addRow([null]);
  blankSheet.addRow([""]);
  blankSheet.addRow(["Note"]);
  blankSheet.autoFilter = "A1:A5";
  const autoSortSheet = workbook.addWorksheet("保存済み並べ替え");
  autoSortSheet.addRow(["Name", "Score"]);
  autoSortSheet.addRow(["A", 5]);
  autoSortSheet.addRow(["C", 20]);
  autoSortSheet.addRow(["B", 12]);
  autoSortSheet.addRow(["D", 8]);
  autoSortSheet.autoFilter = "A1:B5";
  const standaloneSortSheet = workbook.addWorksheet("単独並べ替え");
  standaloneSortSheet.addRow(["Name", "Score"]);
  standaloneSortSheet.addRow(["Delta", 4]);
  standaloneSortSheet.addRow(["Alpha", 1]);
  standaloneSortSheet.addRow(["Charlie", 3]);
  standaloneSortSheet.addRow(["Bravo", 2]);
  const tableSheet = workbook.addWorksheet("テーブルフィルター");
  tableSheet.addTable({
    name: "TableFilter1",
    ref: "A1",
    headerRow: true,
    totalsRow: false,
    columns: [{ name: "Name" }, { name: "Score" }],
    rows: [
      ["Alpha", 10],
      ["Beta", 20],
      ["Alpha", 30],
      ["Gamma", 40],
    ],
  });
  const tableMetaSheet = workbook.addWorksheet("テーブルメタ");
  tableMetaSheet.addTable({
    name: "SalesTable",
    ref: "A1",
    headerRow: true,
    totalsRow: true,
    style: {
      theme: "TableStyleMedium9",
      showFirstColumn: true,
      showLastColumn: false,
      showRowStripes: true,
      showColumnStripes: false,
    },
    columns: [
      { name: "地域", totalsRowLabel: "合計" },
      { name: "売上", totalsRowFunction: "sum" },
      { name: "担当", totalsRowFunction: "none" },
      { name: "倍額", totalsRowFunction: "sum" },
    ],
    rows: [
      ["東", 10, "A", { formula: "[@売上]*2", result: 20 }],
      ["西", 20, "B", { formula: "[@売上]*2", result: 40 }],
      ["東", 30, "C", { formula: "[@売上]*2", result: 60 }],
    ],
  });
  const structuredFormulaSheet = workbook.addWorksheet("構造化参照");
  structuredFormulaSheet.addRow(["合計", { formula: "SUM(SalesTable[売上])", result: 60 }]);
  structuredFormulaSheet.addRow(["データ", { formula: "SUM(SalesTable[[#Data],[売上]])", result: 60 }]);
  structuredFormulaSheet.addRow(["ヘッダー", { formula: "SalesTable[[#Headers],[売上]]", result: "売上" }]);
  structuredFormulaSheet.addRow(["集計", { formula: "SUM(SalesTable[[#Totals],[売上]])", result: 60 }]);
  structuredFormulaSheet.addRow(["列範囲", { formula: "SUM(SalesTable[[売上]:[倍額]])", result: 180 }]);
  const structureShiftSheet = workbook.addWorksheet("構造シフト参照");
  structureShiftSheet.addRow(["データ合計", { formula: "SUM('テーブルメタ'!B2:B4)", result: 60 }]);
  structureShiftSheet.addRow(["合計セル", { formula: "'テーブルメタ'!B5", result: 60 }]);
  structureShiftSheet.addRow(["最終計算列", { formula: "'テーブルメタ'!D4", result: 60 }]);
  structureShiftSheet.addRow(["文字列保持", '=TEXT("B2:B4","@")']);
  structureShiftSheet.addRow(["列全体", { formula: "SUM('テーブルメタ'!B:B)", result: 60 }]);
  structureShiftSheet.addRow(["行全体", { formula: "SUM('テーブルメタ'!2:4)", result: 180 }]);
  structureShiftSheet.addRow(["列範囲", { formula: "SUM('テーブルメタ'!B:D)", result: 180 }]);
  structureShiftSheet.getCell("B8").value = "検証";
  structureShiftSheet.getCell("B8").dataValidation = {
    type: "list",
    formulae: ["'テーブルメタ'!$B$2:$B$4"],
    allowBlank: true,
  };
  const expansionSheet = workbook.addWorksheet("テーブル拡張");
  expansionSheet.addTable({
    name: "ExpandTable",
    ref: "A1",
    headerRow: true,
    totalsRow: false,
    style: {
      theme: "TableStyleMedium2",
      showRowStripes: true,
    },
    columns: [{ name: "Name" }, { name: "Amount" }, { name: "Twice" }],
    rows: [
      ["A", 1, { formula: "[@Amount]*2", result: 2 }],
      ["B", 2, { formula: "[@Amount]*2", result: 4 }],
    ],
  });
  const multiTableSheet = workbook.addWorksheet("複数テーブルフィルター");
  multiTableSheet.addTable({
    name: "LeftFilterTable",
    ref: "A1",
    headerRow: true,
    totalsRow: false,
    columns: [{ name: "LeftName" }, { name: "LeftScore" }],
    rows: [
      ["Alpha", 10],
      ["Beta", 20],
      ["Gamma", 30],
    ],
  });
  multiTableSheet.addTable({
    name: "RightFilterTable",
    ref: "D1",
    headerRow: true,
    totalsRow: false,
    columns: [{ name: "RightName" }, { name: "RightScore" }],
    rows: [
      ["Red", 1],
      ["Blue", 2],
      ["Green", 3],
    ],
  });
  const tableSortSheet = workbook.addWorksheet("テーブル並べ替え");
  tableSortSheet.addTable({
    name: "TableSortOnly",
    ref: "A1",
    headerRow: true,
    totalsRow: false,
    columns: [{ name: "Name" }, { name: "Score" }],
    rows: [
      ["A", 5],
      ["C", 20],
      ["B", 12],
      ["D", 8],
    ],
  });
  const cellColorSheet = workbook.addWorksheet("セル色フィルター");
  cellColorSheet.addRow(["Name", "Score"]);
  cellColorSheet.addRow(["A", 5]);
  cellColorSheet.addRow(["B", 12]);
  cellColorSheet.addRow(["C", 20]);
  cellColorSheet.addRow(["D", 8]);
  cellColorSheet.getCell("B4").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } };
  cellColorSheet.autoFilter = "A1:B5";
  const fontColorSheet = workbook.addWorksheet("フォント色フィルター");
  fontColorSheet.addRow(["Name", "Score"]);
  fontColorSheet.addRow(["A", 5]);
  fontColorSheet.addRow(["B", 12]);
  fontColorSheet.addRow(["C", 20]);
  fontColorSheet.addRow(["D", 8]);
  fontColorSheet.getCell("B3").font = { name: "Calibri", size: 11, color: { argb: "FFFF0000" } };
  fontColorSheet.autoFilter = "A1:B5";
  const tableColorSheet = workbook.addWorksheet("テーブル色フィルター");
  tableColorSheet.addTable({
    name: "TableColorFilter",
    ref: "A1",
    headerRow: true,
    totalsRow: false,
    columns: [{ name: "Name" }, { name: "Score" }],
    rows: [
      ["A", 5],
      ["B", 12],
      ["C", 20],
      ["D", 8],
    ],
  });
  tableColorSheet.getCell("B4").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } };
  const cellColorSortSheet = workbook.addWorksheet("セル色並べ替え");
  cellColorSortSheet.addRow(["Name", "Score"]);
  cellColorSortSheet.addRow(["A", 5]);
  cellColorSortSheet.addRow(["B", 12]);
  cellColorSortSheet.addRow(["C", 20]);
  cellColorSortSheet.addRow(["D", 8]);
  cellColorSortSheet.getCell("B3").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } };
  cellColorSortSheet.autoFilter = "A1:B5";
  const fontColorSortSheet = workbook.addWorksheet("フォント色並べ替え");
  fontColorSortSheet.addRow(["Name", "Score"]);
  fontColorSortSheet.addRow(["A", 5]);
  fontColorSortSheet.addRow(["B", 12]);
  fontColorSortSheet.addRow(["C", 20]);
  fontColorSortSheet.addRow(["D", 8]);
  fontColorSortSheet.getCell("B5").font = { name: "Calibri", size: 11, color: { argb: "FFFF0000" } };
  fontColorSortSheet.autoFilter = "A1:B5";
  const customListSortSheet = workbook.addWorksheet("リスト並べ替え");
  customListSortSheet.addRow(["Month", "Code"]);
  customListSortSheet.addRow(["3月", "M3"]);
  customListSortSheet.addRow(["12月", "M12"]);
  customListSortSheet.addRow(["1月", "M1"]);
  customListSortSheet.addRow(["2月", "M2"]);
  customListSortSheet.autoFilter = "A1:B5";
  const textContainsSheet = workbook.addWorksheet("文字含むフィルター");
  textContainsSheet.addRow(["Name", "Score"]);
  textContainsSheet.addRow(["alpha", 5]);
  textContainsSheet.addRow(["Alpine", 12]);
  textContainsSheet.addRow(["Beta", 20]);
  textContainsSheet.addRow(["delta", 8]);
  textContainsSheet.autoFilter = "A1:B5";
  const textBeginsSheet = workbook.addWorksheet("文字始まりフィルター");
  textBeginsSheet.addRow(["Name", "Score"]);
  textBeginsSheet.addRow(["alpha", 5]);
  textBeginsSheet.addRow(["Alpine", 12]);
  textBeginsSheet.addRow(["Beta", 20]);
  textBeginsSheet.addRow(["delta", 8]);
  textBeginsSheet.autoFilter = "A1:B5";
  const textNotEndsSheet = workbook.addWorksheet("文字終わり除外");
  textNotEndsSheet.addRow(["Name", "Score"]);
  textNotEndsSheet.addRow(["alpha", 5]);
  textNotEndsSheet.addRow(["Alpine", 12]);
  textNotEndsSheet.addRow(["Beta", 20]);
  textNotEndsSheet.addRow(["delta", 8]);
  textNotEndsSheet.autoFilter = "A1:B5";
  const textLiteralWildcardSheet = workbook.addWorksheet("文字リテラルフィルター");
  textLiteralWildcardSheet.addRow(["Name", "Score"]);
  textLiteralWildcardSheet.addRow(["A*", 1]);
  textLiteralWildcardSheet.addRow(["A?", 2]);
  textLiteralWildcardSheet.addRow(["Alpha", 3]);
  textLiteralWildcardSheet.addRow(["Star", 4]);
  textLiteralWildcardSheet.autoFilter = "A1:B5";
  const multiDateGroupSheet = workbook.addWorksheet("日付複数グループ");
  multiDateGroupSheet.addRow(["Date", "Amount"]);
  multiDateGroupSheet.addRow(["2026/05/20", 10]);
  multiDateGroupSheet.addRow(["2026/06/01", 20]);
  multiDateGroupSheet.addRow(["2025/05/10", 30]);
  multiDateGroupSheet.addRow(["2026/07/01", 40]);
  multiDateGroupSheet.autoFilter = "A1:B5";
  const betweenSheet = workbook.addWorksheet("範囲内フィルター");
  betweenSheet.addRow(["Name", "Score"]);
  betweenSheet.addRow(["A", 5]);
  betweenSheet.addRow(["B", 12]);
  betweenSheet.addRow(["C", 20]);
  betweenSheet.addRow(["D", 8]);
  betweenSheet.autoFilter = "A1:B5";
  const notBetweenSheet = workbook.addWorksheet("範囲外フィルター");
  notBetweenSheet.addRow(["Name", "Score"]);
  notBetweenSheet.addRow(["A", 5]);
  notBetweenSheet.addRow(["B", 12]);
  notBetweenSheet.addRow(["C", 20]);
  notBetweenSheet.addRow(["D", 8]);
  notBetweenSheet.autoFilter = "A1:B5";
  const monthPeriodSheet = workbook.addWorksheet("月期間フィルター");
  monthPeriodSheet.addRow(["Date", "Amount"]);
  monthPeriodSheet.addRow(["2026/05/20", 10]);
  monthPeriodSheet.addRow(["2025/05/10", 20]);
  monthPeriodSheet.addRow(["2026/06/01", 30]);
  monthPeriodSheet.addRow(["2026/04/30", 40]);
  monthPeriodSheet.autoFilter = "A1:B5";
  const quarterPeriodSheet = workbook.addWorksheet("四半期期間フィルター");
  quarterPeriodSheet.addRow(["Date", "Amount"]);
  quarterPeriodSheet.addRow(["2026/04/01", 10]);
  quarterPeriodSheet.addRow(["2025/05/10", 20]);
  quarterPeriodSheet.addRow(["2026/07/01", 30]);
  quarterPeriodSheet.addRow(["2026/03/31", 40]);
  quarterPeriodSheet.autoFilter = "A1:B5";
  const aboveStdDevSheet = workbook.addWorksheet("標準偏差上フィルター");
  aboveStdDevSheet.addRow(["Name", "Score"]);
  aboveStdDevSheet.addRow(["A", 1]);
  aboveStdDevSheet.addRow(["B", 10]);
  aboveStdDevSheet.addRow(["C", 20]);
  aboveStdDevSheet.addRow(["D", 40]);
  aboveStdDevSheet.autoFilter = "A1:B5";
  const belowStdDevSheet = workbook.addWorksheet("標準偏差下フィルター");
  belowStdDevSheet.addRow(["Name", "Score"]);
  belowStdDevSheet.addRow(["A", 1]);
  belowStdDevSheet.addRow(["B", 10]);
  belowStdDevSheet.addRow(["C", 20]);
  belowStdDevSheet.addRow(["D", 40]);
  belowStdDevSheet.autoFilter = "A1:B5";
  const hiddenButtonSheet = workbook.addWorksheet("非表示ボタンフィルター");
  hiddenButtonSheet.addRow(["Name", "Score"]);
  hiddenButtonSheet.addRow(["Alpha", 10]);
  hiddenButtonSheet.addRow(["Beta", 20]);
  hiddenButtonSheet.addRow(["Alpha", 30]);
  hiddenButtonSheet.addRow(["Gamma", 40]);
  hiddenButtonSheet.autoFilter = "A1:B5";
  const dateHourGroupSheet = workbook.addWorksheet("日付時刻グループ");
  dateHourGroupSheet.addRow(["Date", "Amount"]);
  dateHourGroupSheet.addRow(["2026/05/20 13:10", 10]);
  dateHourGroupSheet.addRow(["2026/05/20 14:10", 20]);
  dateHourGroupSheet.addRow(["2026/05/21 13:10", 30]);
  dateHourGroupSheet.addRow(["2026/05/20 13:55", 40]);
  dateHourGroupSheet.autoFilter = "A1:B5";
  const customNonBlankSheet = workbook.addWorksheet("空白以外カスタム");
  customNonBlankSheet.addRow(["Name", "Score"]);
  customNonBlankSheet.addRow(["Filled", 10]);
  customNonBlankSheet.addRow(["", 20]);
  customNonBlankSheet.addRow(["Zero", 0]);
  customNonBlankSheet.addRow(["", 40]);
  customNonBlankSheet.autoFilter = "A1:B5";
  await workbook.xlsx.writeFile(outputPath);
  const zip = await JSZip.loadAsync(await readFile(outputPath));
  const [importedFillDxfId, importedFontDxfId] = await appendImportedAutoFilterDxfs(zip, [
    '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill></dxf>',
    '<dxf><font><color rgb="FFFF0000"/></font></dxf>',
  ]);
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet1.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="0"><filters><filter val="Alpha"/></filters></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet2.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="10"/></customFilters></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet3.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="1"><top10 top="1" percent="0" val="2"/></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet4.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="1"><dynamicFilter type="aboveAverage"/></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet5.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="0"><filters><dateGroupItem year="2026" month="5" dateTimeGrouping="month"/></filters></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet6.xml", "A1:A5", '<autoFilter ref="A1:A5"><filterColumn colId="0"><filters blank="1"/></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet7.xml", "A1:B5", '<autoFilter ref="A1:B5"><sortState ref="A2:B5"><sortCondition descending="1" ref="B2:B5"/></sortState></autoFilter>');
  await insertWorksheetSortState(zip, "xl/worksheets/sheet8.xml", '<sortState ref="A2:B5"><sortCondition ref="A2:A5"/></sortState>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet16.xml", "A1:B5", `<autoFilter ref="A1:B5"><filterColumn colId="1"><colorFilter dxfId="${importedFillDxfId}" cellColor="1"/></filterColumn></autoFilter>`);
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet17.xml", "A1:B5", `<autoFilter ref="A1:B5"><filterColumn colId="1"><colorFilter dxfId="${importedFontDxfId}" cellColor="0"/></filterColumn></autoFilter>`);
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet19.xml", "A1:B5", `<autoFilter ref="A1:B5"><sortState ref="A2:B5"><sortCondition sortBy="cellColor" dxfId="${importedFillDxfId}" ref="B2:B5"/></sortState></autoFilter>`);
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet20.xml", "A1:B5", `<autoFilter ref="A1:B5"><sortState ref="A2:B5"><sortCondition sortBy="fontColor" dxfId="${importedFontDxfId}" ref="B2:B5"/></sortState></autoFilter>`);
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet21.xml", "A1:B5", '<autoFilter ref="A1:B5"><sortState ref="A2:B5"><sortCondition customList="1月,2月,3月,4月,5月,6月,7月,8月,9月,10月,11月,12月" ref="A2:A5"/></sortState></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet22.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="0"><customFilters><customFilter operator="equal" val="*ta*"/></customFilters></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet23.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="0"><customFilters><customFilter operator="equal" val="Al*"/></customFilters></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet24.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="0"><customFilters><customFilter operator="notEqual" val="*ta"/></customFilters></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet25.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="0"><customFilters><customFilter operator="equal" val="A~*"/></customFilters></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet26.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="0"><filters><dateGroupItem year="2026" month="5" dateTimeGrouping="month"/><dateGroupItem year="2026" month="6" dateTimeGrouping="month"/></filters></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet27.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="1"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="10"/><customFilter operator="lessThanOrEqual" val="20"/></customFilters></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet28.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="1"><customFilters><customFilter operator="lessThan" val="10"/><customFilter operator="greaterThan" val="20"/></customFilters></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet29.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="0"><dynamicFilter type="M5"/></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet30.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="0"><dynamicFilter type="Q2"/></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet31.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="1"><dynamicFilter type="aboveStdDev" val="1"/></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet32.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="1"><dynamicFilter type="belowStdDev" val="1"/></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet33.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="0"><filters><filter val="Alpha"/></filters></filterColumn><filterColumn colId="1" hiddenButton="1" showButton="0"/></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet34.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="0"><filters><dateGroupItem year="2026" month="5" day="20" hour="13" dateTimeGrouping="hour"/></filters></filterColumn></autoFilter>');
  await replaceWorksheetAutoFilter(zip, "xl/worksheets/sheet35.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="0"><customFilters><customFilter operator="notEqual" val=""/></customFilters></filterColumn></autoFilter>');
  await replaceTableAutoFilter(zip, "xl/tables/table1.xml", "A1:B5", '<autoFilter ref="A1:B5"><filterColumn colId="0"><filters><filter val="Alpha"/></filters></filterColumn></autoFilter>');
  await replaceTableAutoFilter(zip, "xl/tables/table2.xml", "A1:D4", '<autoFilter ref="A1:D4"/>');
  await replaceTableAutoFilter(zip, "xl/tables/table3.xml", "A1:C3", '<autoFilter ref="A1:C3"/>');
  await replaceTableAutoFilter(zip, "xl/tables/table4.xml", "A1:B4", '<autoFilter ref="A1:B4"><filterColumn colId="0"><filters><filter val="Alpha"/></filters></filterColumn></autoFilter>');
  await replaceTableAutoFilter(zip, "xl/tables/table5.xml", "D1:E4", '<autoFilter ref="D1:E4"><filterColumn colId="0"><filters><filter val="Red"/></filters></filterColumn></autoFilter>');
  await replaceTableAutoFilter(zip, "xl/tables/table6.xml", "A1:B5", '<autoFilter ref="A1:B5"><sortState ref="A2:B5"><sortCondition descending="1" ref="B2:B5"/></sortState></autoFilter>');
  await replaceTableAutoFilter(zip, "xl/tables/table7.xml", "A1:B5", `<autoFilter ref="A1:B5"><filterColumn colId="1"><colorFilter dxfId="${importedFillDxfId}" cellColor="1"/></filterColumn></autoFilter>`);
  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
  await addWorkbookDefinedNames(outputPath, [
    { name: "SalesDataAmounts", refersTo: "'テーブルメタ'!$B$2:$B$4" },
    { name: "SalesTotalCell", refersTo: "'テーブルメタ'!$B$5" },
    { name: "SalesWholeAmountColumn", refersTo: "'テーブルメタ'!$B:$B" },
    { name: "SalesDataRows", refersTo: "'テーブルメタ'!$2:$4" },
  ]);
  return outputPath;
}



async function replaceWorksheetAutoFilter(zip, path, ref, autoFilterXml) {
  const file = zip.file(path);
  if (!file) return;
  let xml = await file.async("text");
  const escapedRef = ref.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`<autoFilter\\s+ref="${escapedRef}"\\s*/>`);
  xml = pattern.test(xml)
    ? xml.replace(pattern, autoFilterXml)
    : xml.replace("</worksheet>", `${autoFilterXml}</worksheet>`);
  zip.file(path, xml);
}



async function appendImportedAutoFilterDxfs(zip, dxfXmls = []) {
  const file = zip.file("xl/styles.xml");
  if (!file || !dxfXmls.length) return dxfXmls.map((_, index) => index);
  const stylesXml = await file.async("text");
  const existingCount = importedStyleDxfCount(stylesXml);
  const totalCount = existingCount + dxfXmls.length;
  const appended = dxfXmls.join("");
  let nextXml = stylesXml;
  if (/<dxfs\b[^>]*\/>/.test(nextXml)) {
    nextXml = nextXml.replace(/<dxfs\b([^>]*)\/>/, (match, attrs) => {
      const cleanAttrs = String(attrs || "").replace(/\s+count="[^"]*"/, "");
      return `<dxfs${cleanAttrs} count="${totalCount}">${appended}</dxfs>`;
    });
  } else if (/<dxfs\b[\s\S]*?<\/dxfs>/.test(nextXml)) {
    nextXml = nextXml.replace(/<dxfs\b[^>]*>[\s\S]*?<\/dxfs>/, (block) => {
      const open = block.match(/^<dxfs\b[^>]*>/)?.[0] || "<dxfs>";
      const body = block.replace(/^<dxfs\b[^>]*>/, "").replace(/<\/dxfs>$/, "");
      const nextOpen = /\bcount=/.test(open)
        ? open.replace(/\bcount="[^"]*"/, `count="${totalCount}"`)
        : open.replace(/>$/, ` count="${totalCount}">`);
      return `${nextOpen}${body}${appended}</dxfs>`;
    });
  } else {
    const dxfsXml = `<dxfs count="${totalCount}">${appended}</dxfs>`;
    nextXml = /<tableStyles\b/.test(nextXml)
      ? nextXml.replace(/<tableStyles\b/, `${dxfsXml}<tableStyles`)
      : nextXml.replace("</styleSheet>", `${dxfsXml}</styleSheet>`);
  }
  zip.file("xl/styles.xml", nextXml);
  return dxfXmls.map((_, index) => existingCount + index);
}



function importedStyleDxfCount(stylesXml) {
  const count = Number.parseInt(String(stylesXml || "").match(/<dxfs\b[^>]*\bcount="(\d+)"/)?.[1] || "", 10);
  if (Number.isInteger(count) && count >= 0) return count;
  const block = String(stylesXml || "").match(/<dxfs\b[\s\S]*?<\/dxfs>/)?.[0] || "";
  return (block.match(/<dxf\b/g) || []).length;
}



async function replaceTableAutoFilter(zip, path, ref, autoFilterXml) {
  const file = zip.file(path);
  if (!file) return;
  let xml = await file.async("text");
  if (/<autoFilter\b[^>]*\/>/.test(xml)) {
    xml = xml.replace(/<autoFilter\b[^>]*\/>/, autoFilterXml);
  } else if (/<autoFilter\b[\s\S]*?<\/autoFilter>/.test(xml)) {
    xml = xml.replace(/<autoFilter\b[\s\S]*?<\/autoFilter>/, autoFilterXml);
  } else {
    xml = xml.replace("</table>", `${autoFilterXml}</table>`);
  }
  zip.file(path, xml);
}



async function insertWorksheetSortState(zip, path, sortStateXml) {
  const file = zip.file(path);
  if (!file) return;
  let xml = await file.async("text");
  xml = xml.includes("<sortState")
    ? xml.replace(/<sortState[\s\S]*?<\/sortState>/, sortStateXml)
    : xml.replace("</worksheet>", `${sortStateXml}</worksheet>`);
  zip.file(path, xml);
}



async function createHomeToolbarWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-home-toolbar-"));
  const outputPath = path.join(dir, "home-toolbar.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("HomeTools");
  sheet.addRow(["Name", "Score", "", "Values", "", "Total", ""]);
  sheet.addRow(["Beta", 2, "", 1]);
  sheet.addRow(["Alpha", 1, "", 2]);
  sheet.addRow(["Gamma", 3, "", 3]);
  sheet.getCell("F1").value = { formula: "SUM(B2:B4)", result: 6 };
  sheet.getCell("G2").note = "確認用メモ";
  sheet.getCell("H1").value = "Average";
  sheet.getCell("H2").value = 4;
  sheet.getCell("H3").value = 8;
  sheet.getCell("H4").value = 12;
  sheet.getCell("J2").value = "Seed";
  sheet.getCell("J2").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
  sheet.getCell("K2").value = { formula: "H2*2", result: 8 };
  sheet.getCell("L2").value = "Format";
  sheet.getCell("L2").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00B050" } };
  sheet.getCell("M2").value = { text: "Link", hyperlink: "https://example.com" };
  sheet.getCell("N2").value = "A";
  sheet.getCell("N2").dataValidation = { type: "list", allowBlank: true, formulae: ['"A,B"'] };
  sheet.getCell("N3").value = "B";
  sheet.getCell("N3").dataValidation = { type: "list", allowBlank: true, formulae: ['"A,B"'] };
  sheet.getCell("O2").value = "find foo";
  sheet.getCell("O3").value = "foo foo";
  sheet.getCell("O4").value = "bar";
  sheet.getCell("P2").value = "A";
  sheet.getCell("Q2").value = "A";
  sheet.getCell("R2").value = "B";
  sheet.getCell("P3").value = 1;
  sheet.getCell("Q3").value = 2;
  sheet.getCell("R3").value = 1;
  sheet.getCell("P5").value = 1;
  sheet.getCell("P6").value = 1;
  sheet.getCell("P7").value = 2;
  sheet.getCell("Q5").value = "A";
  sheet.getCell("Q6").value = "B";
  sheet.getCell("Q7").value = "A";
  sheet.getCell("S2").value = 1;
  sheet.getCell("S3").value = 2;
  sheet.getCell("S4").value = { formula: "SUM(S2:S3)", result: 3 };
  sheet.getColumn(1).width = 12;
  sheet.getColumn(2).width = 10;
  sheet.getColumn(4).width = 10;
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createSortDialogWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-sort-dialog-"));
  const outputPath = path.join(dir, "sort-dialog.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Sort");
  sheet.addRow(["Group", "Score"]);
  sheet.addRow(["B", 12]);
  sheet.addRow(["A", 5]);
  sheet.addRow(["B", 8]);
  sheet.addRow(["A", 20]);
  sheet.getCell("B4").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFF0000" },
  };
  sheet.getCell("B3").font = { name: "Calibri", size: 11, color: { argb: "FFFF0000" } };
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createRemoveDuplicatesWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-remove-duplicates-"));
  const outputPath = path.join(dir, "remove-duplicates.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Duplicates");
  sheet.addRow(["Name", "Score"]);
  sheet.addRow(["A", 1]);
  sheet.addRow(["B", 2]);
  sheet.addRow(["A", 1]);
  sheet.addRow(["C", 3]);
  sheet.addRow(["a", 1]);
  sheet.getCell("D1").value = "Calc";
  sheet.getCell("E1").value = "Group";
  sheet.getCell("D2").value = { formula: "1+1" };
  sheet.getCell("E2").value = "Same";
  sheet.getCell("D3").value = { formula: "3-1" };
  sheet.getCell("E3").value = "Same";
  sheet.getCell("D4").value = { formula: "2+3" };
  sheet.getCell("E4").value = "Other";
  sheet.getCell("F1").value = "Date";
  sheet.getCell("G1").value = "Group";
  sheet.getCell("F2").value = new Date(Date.UTC(2026, 4, 20));
  sheet.getCell("F2").numFmt = 'yyyy/mm/dd;@';
  sheet.getCell("G2").value = "Same";
  sheet.getCell("F3").value = new Date(Date.UTC(2026, 4, 20));
  sheet.getCell("F3").numFmt = 'mm/dd/yyyy;@';
  sheet.getCell("G3").value = "Same";
  sheet.getCell("F4").value = new Date(Date.UTC(2026, 4, 21));
  sheet.getCell("F4").numFmt = 'yyyy/mm/dd;@';
  sheet.getCell("G4").value = "Other";
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createAutoFilterDateWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-date-filter-"));
  const outputPath = path.join(dir, "date-filter.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("DateFilter");
  sheet.addRow(["Date", "Amount"]);
  sheet.addRow([new Date(Date.UTC(2026, 4, 20)), 10]);
  sheet.addRow([new Date(Date.UTC(2026, 4, 20)), 11]);
  sheet.addRow(["05/23/2026", 20]);
  sheet.addRow(["06/01/2026", 30]);
  sheet.addRow(["2025/12/31", 40]);
  sheet.getCell("A2").numFmt = 'mmm d, yyyy;@';
  sheet.getCell("A3").numFmt = 'm/d/yy;@';
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createAdvancedFilterUniqueWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-advanced-unique-"));
  const outputPath = path.join(dir, "advanced-unique.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("AdvancedUnique");
  sheet.addRow(["Name", "Score", "Date"]);
  sheet.addRow(["Alpha", 10, new Date(Date.UTC(2026, 4, 20))]);
  sheet.addRow(["alpha", 10, new Date(Date.UTC(2026, 4, 20))]);
  sheet.addRow(["Beta", 20, "05/23/2026"]);
  sheet.addRow(["Gamma", 30, "12/31/2025"]);
  sheet.getCell("C2").numFmt = 'mmm d, yyyy;@';
  sheet.getCell("C3").numFmt = 'm/d/yy;@';
  sheet.getCell("E1").value = "Date";
  sheet.getCell("E2").value = ">=May 1, 2026";
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createAdvancedFilterUniqueSubsetWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-advanced-subset-"));
  const outputPath = path.join(dir, "advanced-subset.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Subset");
  sheet.addRow(["Name", "Score"]);
  sheet.addRow(["Alpha", 10]);
  sheet.addRow(["Alpha", 20]);
  sheet.addRow(["Beta", 20]);
  sheet.addRow(["Alpha", 10]);
  sheet.getCell("F1").value = "Name";
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createDelimitedImportWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-import-target-"));
  const outputPath = path.join(dir, "import-target.xlsx");
  const csvPath = path.join(dir, "import-source.csv");
  const textCsvPath = path.join(dir, "import-text-source.csv");
  const transformCsvPath = path.join(dir, "import-transform-source.csv");
  const utf16CsvPath = path.join(dir, "import-utf16-source.csv");
  const shiftJisCsvPath = path.join(dir, "import-shiftjis-source.csv");
  const jsonPath = path.join(dir, "import-structured-source.json");
  const xmlPath = path.join(dir, "import-structured-source.xml");
  const workbookDataPath = path.join(dir, "import-workbook-source.xlsx");
  const folderCsvPath = path.join(dir, "folder-source.csv");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Import");
  sheet.getCell("A1").value = "既存";
  sheet.getCell("B2").value = "保持";
  await workbook.xlsx.writeFile(outputPath);
  await writeFile(csvPath, "Name,Score\nAlpha,10\nBeta,20\n", "utf8");
  await writeFile(textCsvPath, "Code,Qty\n0012,09\n", "utf8");
  await writeFile(transformCsvPath, "Name,Score,City\nAlpha,10,Tokyo\nBeta,20,Osaka\nGamma,30,Tokyo\n", "utf8");
  await writeFile(
    utf16CsvPath,
    Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from("名前,点数\n東京,10\n大阪,20\n", "utf16le")]),
  );
  await writeFile(shiftJisCsvPath, Buffer.from("96bc914f2c935f90940a938c8b9e2c31300a91e58de32c32300a", "hex"));
  await writeFile(folderCsvPath, "Name,Score,City\nFolderAlpha,51,Tokyo\nFolderBeta,62,Osaka\n", "utf8");
  await writeFile(
    jsonPath,
    JSON.stringify([
      { Name: "JsonAlpha", Score: 11, Detail: { City: "Tokyo" } },
      { Name: "JsonBeta", Score: 22, Detail: { City: "Osaka" } },
    ]),
    "utf8",
  );
  await writeFile(
    xmlPath,
    `<catalog>
      <entry code="X1"><Name>XmlAlpha</Name><Qty>4</Qty><City>Tokyo</City></entry>
      <entry code="X2"><Name>XmlBeta</Name><Qty>8</Qty><City>Kyoto</City></entry>
    </catalog>`,
    "utf8",
  );
  const sourceWorkbook = new ExcelJS.Workbook();
  const summary = sourceWorkbook.addWorksheet("Summary");
  summary.addRow(["Metric", "Value"]);
  summary.addRow(["Count", 2]);
  const rows = sourceWorkbook.addWorksheet("Rows");
  rows.addRow(["Name", "Score", "City"]);
  rows.addRow(["BookAlpha", 31, "Tokyo"]);
  rows.addRow(["BookBeta", 42, "Osaka"]);
  await sourceWorkbook.xlsx.writeFile(workbookDataPath);
  return { workbookPath: outputPath, csvPath, textCsvPath, transformCsvPath, utf16CsvPath, shiftJisCsvPath, jsonPath, xmlPath, workbookDataPath, folderCsvPath };
}



async function createPowerQueryTransformWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-power-query-"));
  const outputPath = path.join(dir, "power-query.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("PQ");
  sheet.addRow(["Name", "Score", "City"]);
  sheet.addRow(["Alpha", "10", "Tokyo"]);
  sheet.addRow(["Beta", "20", "Osaka"]);
  sheet.addRow(["Alpha", "10", "Tokyo"]);
  sheet.addRow(["Gamma", "", "Tokyo"]);
  sheet.addRow(["", "", ""]);
  sheet.getCell("A8").value = "FullName";
  sheet.getCell("B8").value = "Amount";
  sheet.getCell("C8").value = "Category";
  sheet.getCell("A9").value = " Alpha - East ";
  sheet.getCell("B9").value = "1";
  sheet.getCell("C9").value = "A";
  sheet.getCell("A10").value = "";
  sheet.getCell("B10").value = "2";
  sheet.getCell("C10").value = "A";
  sheet.getCell("A11").value = "Beta - West\n";
  sheet.getCell("B11").value = "3";
  sheet.getCell("C11").value = "B";
  sheet.getCell("A12").value = "";
  sheet.getCell("B12").value = "4";
  sheet.getCell("C12").value = "B";
  sheet.getCell("A15").value = "Item";
  sheet.getCell("B15").value = "Qty";
  sheet.getCell("C15").value = "Region";
  sheet.getCell("D15").value = "Code";
  sheet.getCell("A16").value = "Alpha";
  sheet.getCell("B16").value = "5";
  sheet.getCell("C16").value = "Tokyo";
  sheet.getCell("D16").value = "A-001";
  sheet.getCell("A17").value = "Beta";
  sheet.getCell("B17").value = "12";
  sheet.getCell("C17").value = "Osaka";
  sheet.getCell("D17").value = "B-020";
  sheet.getCell("A18").value = "Gamma";
  sheet.getCell("B18").value = "8";
  sheet.getCell("C18").value = "Tokyo";
  sheet.getCell("D18").value = "C-003";
  sheet.getCell("A19").value = "Delta";
  sheet.getCell("B19").value = "20";
  sheet.getCell("C19").value = "Kyoto";
  sheet.getCell("D19").value = "D-100";
  sheet.getCell("A22").value = "Dept";
  sheet.getCell("B22").value = "Q1";
  sheet.getCell("C22").value = "Q2";
  sheet.getCell("D22").value = "City";
  sheet.getCell("E22").value = "Code";
  sheet.getCell("A23").value = "East";
  sheet.getCell("B23").value = "10";
  sheet.getCell("C23").value = "20";
  sheet.getCell("D23").value = "tokyo";
  sheet.getCell("E23").value = "aa-001";
  sheet.getCell("A24").value = "West";
  sheet.getCell("B24").value = "30";
  sheet.getCell("C24").value = "40";
  sheet.getCell("D24").value = "osaka";
  sheet.getCell("E24").value = "bb-002";
  sheet.getCell("A25").value = "North";
  sheet.getCell("B25").value = "50";
  sheet.getCell("C25").value = "60";
  sheet.getCell("D25").value = "kyoto";
  sheet.getCell("E25").value = "cc-003";
  sheet.getCell("A26").value = "South";
  sheet.getCell("B26").value = "70";
  sheet.getCell("C26").value = "80";
  sheet.getCell("D26").value = "nagoya";
  sheet.getCell("E26").value = "dd-004";
  sheet.getCell("A29").value = "Date";
  sheet.getCell("B29").value = "Amount";
  sheet.getCell("C29").value = "Err";
  sheet.getCell("D29").value = "Empty";
  sheet.getCell("A30").value = "2026-05-24";
  sheet.getCell("B30").value = "10.4";
  sheet.getCell("C30").value = "#N/A";
  sheet.getCell("A31").value = "2026-06-01";
  sheet.getCell("B31").value = "20.5";
  sheet.getCell("C31").value = "OK";
  sheet.getCell("A32").value = "2025-12-31";
  sheet.getCell("B32").value = "30.6";
  sheet.getCell("C32").value = "#DIV/0!";
  sheet.getCell("A35").value = "Metric";
  sheet.getCell("B35").value = "Q1";
  sheet.getCell("C35").value = "Q2";
  sheet.getCell("A36").value = "Sales";
  sheet.getCell("B36").value = "10";
  sheet.getCell("C36").value = "20";
  sheet.getCell("A37").value = "Cost";
  sheet.getCell("B37").value = "3";
  sheet.getCell("C37").value = "4";
  sheet.getCell("A40").value = "RowID";
  sheet.getCell("B40").value = "Value";
  sheet.getCell("C40").value = "Note";
  sheet.getCell("A41").value = "1";
  sheet.getCell("B41").value = "10";
  sheet.getCell("C41").value = "top";
  sheet.getCell("A42").value = "2";
  sheet.getCell("B42").value = "20";
  sheet.getCell("C42").value = "mid";
  sheet.getCell("A43").value = "3";
  sheet.getCell("B43").value = "30";
  sheet.getCell("C43").value = "keep";
  sheet.getCell("A44").value = "4";
  sheet.getCell("B44").value = "40";
  sheet.getCell("C44").value = "tail";
  sheet.getCell("A45").value = "5";
  sheet.getCell("B45").value = "50";
  sheet.getCell("C45").value = "bottom";
  sheet.getCell("A48").value = "Region";
  sheet.getCell("B48").value = "Plan";
  sheet.getCell("C48").value = "Amount";
  sheet.getCell("A49").value = "East";
  sheet.getCell("B49").value = "Basic";
  sheet.getCell("C49").value = "10";
  sheet.getCell("A50").value = "East";
  sheet.getCell("B50").value = "Pro";
  sheet.getCell("C50").value = "25";
  sheet.getCell("A51").value = "West";
  sheet.getCell("B51").value = "Basic";
  sheet.getCell("C51").value = "7";
  sheet.getCell("A52").value = "West";
  sheet.getCell("B52").value = "Pro";
  sheet.getCell("C52").value = "12";
  sheet.getCell("A53").value = "North";
  sheet.getCell("B53").value = "Basic";
  sheet.getCell("C53").value = "5";
  sheet.getCell("A54").value = "North";
  sheet.getCell("B54").value = "Pro";
  sheet.getCell("C54").value = "9";
  sheet.getCell("A57").value = "Item";
  sheet.getCell("B57").value = "Value";
  sheet.getCell("A58").value = "N1";
  sheet.getCell("B58").value = "-25";
  sheet.getCell("A59").value = "N2";
  sheet.getCell("B59").value = "50";
  sheet.getCell("A60").value = "N3";
  sheet.getCell("B60").value = "75";
  const dateHeaders = [
    "Quarter",
    "MonthName",
    "DayName",
    "DayOfWeek",
    "DayOfYear",
    "StartMonth",
    "EndMonth",
    "StartQuarter",
    "EndQuarter",
    "StartYear",
    "EndYear",
  ];
  dateHeaders.forEach((header, index) => {
    sheet.getCell(63, index + 1).value = header;
    sheet.getCell(64, index + 1).value = "2026-05-24";
    sheet.getCell(65, index + 1).value = "2026-12-31";
  });
  sheet.getCell("A68").value = "Text";
  sheet.getCell("A69").value = "ID-[ab12]-Tokyo";
  sheet.getCell("A70").value = "ID-[cd34]-Osaka";
  sheet.getCell("A71").value = "ID-[ef56]-Kyoto";
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createFlashFillWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-flash-fill-"));
  const outputPath = path.join(dir, "flash-fill.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("FlashFill");
  sheet.getCell("A1").value = "Name";
  sheet.getCell("B1").value = "Initials";
  sheet.getCell("A2").value = "John Smith";
  sheet.getCell("B2").value = "JS";
  sheet.getCell("A3").value = "Ada Lovelace";
  sheet.getCell("A4").value = "Grace Hopper";
  sheet.getCell("C1").value = "Email";
  sheet.getCell("D1").value = "User";
  sheet.getCell("C2").value = "alpha.one@example.com";
  sheet.getCell("D2").value = "alpha.one";
  sheet.getCell("C3").value = "beta.long@example.jp";
  sheet.getCell("C4").value = "gamma@example.net";
  sheet.getCell("E1").value = "Full";
  sheet.getCell("F1").value = "Reversed";
  sheet.getCell("E2").value = "John Smith";
  sheet.getCell("F2").value = "Smith, John";
  sheet.getCell("E3").value = "Ada Lovelace";
  sheet.getCell("E4").value = "Grace Hopper";
  sheet.getCell("G1").value = "Code";
  sheet.getCell("H1").value = "Number";
  sheet.getCell("G2").value = "INV001-Tokyo";
  sheet.getCell("H2").value = "001";
  sheet.getCell("G3").value = "INV120-Osaka";
  sheet.getCell("G4").value = "INV045-Kyoto";
  sheet.getCell("I1").value = "AlphaNum";
  sheet.getCell("J1").value = "Letters";
  sheet.getCell("I2").value = "AB123";
  sheet.getCell("J2").value = "AB";
  sheet.getCell("I3").value = "CD456";
  sheet.getCell("I4").value = "EF789";
  sheet.getCell("K1").value = "Area";
  sheet.getCell("L1").value = "Code";
  sheet.getCell("K2").value = "東京（JP）";
  sheet.getCell("L2").value = "JP";
  sheet.getCell("K3").value = "大阪（OS）";
  sheet.getCell("K4").value = "京都（KY）";
  sheet.getCell("M1").value = "SlashDate";
  sheet.getCell("N1").value = "HyphenDate";
  sheet.getCell("M2").value = "2026/05/24";
  sheet.getCell("N2").value = "2026-05-24";
  sheet.getCell("M3").value = "2026/06/01";
  sheet.getCell("M4").value = "2025/12/31";
  sheet.getCell("O1").value = "Phone";
  sheet.getCell("P1").value = "PhoneFmt";
  sheet.getCell("O2").value = "09012345678";
  sheet.getCell("P2").value = "090-1234-5678";
  sheet.getCell("O3").value = "08099887766";
  sheet.getCell("O4").value = "07011112222";
  sheet.getCell("Q1").value = "CompactDate";
  sheet.getCell("R1").value = "DateFmt";
  sheet.getCell("Q2").value = "20260524";
  sheet.getCell("R2").value = "2026/05/24";
  sheet.getCell("Q3").value = "20260601";
  sheet.getCell("Q4").value = "20251231";
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createWhatIfWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-what-if-"));
  const outputPath = path.join(dir, "what-if.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("WhatIf");
  sheet.getCell("A1").value = "Input";
  sheet.getCell("B1").value = 10;
  sheet.getCell("A2").value = "Output";
  sheet.getCell("B2").value = { formula: "B1*2" };
  sheet.getCell("A3").value = "Input2";
  sheet.getCell("B3").value = 3;
  sheet.getCell("A4").value = "Combo";
  sheet.getCell("B4").value = { formula: "B1+B3" };
  sheet.getCell("D1").value = "Input";
  sheet.getCell("E1").value = { formula: "B2" };
  sheet.getCell("D2").value = 5;
  sheet.getCell("D3").value = 10;
  sheet.getCell("D4").value = 15;
  sheet.getCell("G2").value = 5;
  sheet.getCell("G3").value = 10;
  sheet.getCell("H2").value = "No formula";
  sheet.getCell("J1").value = { formula: "B4" };
  sheet.getCell("K1").value = 5;
  sheet.getCell("L1").value = 10;
  sheet.getCell("J2").value = 1;
  sheet.getCell("J3").value = 2;
  sheet.getCell("N1").value = "External input";
  sheet.getCell("O1").value = { formula: "C2" };
  sheet.getCell("N2").value = 2;
  sheet.getCell("N3").value = 4;
  sheet.getCell("N4").value = 6;
  sheet.getCell("C1").value = "Cross sheet";
  sheet.getCell("C2").value = { formula: "Inputs!B1*3" };
  const inputs = workbook.addWorksheet("Inputs");
  inputs.getCell("A1").value = "External input";
  inputs.getCell("B1").value = 4;
  workbook.definedNames.add("WhatIf!$B$1", "MainInput");
  workbook.definedNames.add("WhatIf!$B$2", "MainResult");
  workbook.definedNames.add("Inputs!$B$1", "ExternalInput");
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createForecastWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-forecast-"));
  const outputPath = path.join(dir, "forecast.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("ForecastSource");
  sheet.addRow(["Period", "Value"]);
  sheet.addRow([1, 10]);
  sheet.addRow([2, 20]);
  sheet.addRow([3, 32]);
  sheet.addRow([4, 41]);
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createDataAnalysisWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-analysis-"));
  const outputPath = path.join(dir, "data-analysis.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("分析");
  sheet.addRow(["X", "Y", "GroupA", "GroupB"]);
  sheet.addRow([1, 3, 10, 8]);
  sheet.addRow([2, 5, 12, 9]);
  sheet.addRow([3, 7, 15, 11]);
  sheet.addRow([4, 9, 18, 14]);
  sheet.addRow([5, 11, 20, 17]);
  sheet.addRow(["Low", "Prob"]);
  sheet.addRow([100, 0.25]);
  sheet.addRow([200, 0.5]);
  sheet.addRow([300, 0.25]);
  sheet.getCell("F1").value = "Bin";
  sheet.getCell("F2").value = 10;
  sheet.getCell("F3").value = 15;
  sheet.getCell("F4").value = 20;
  sheet.getCell("H1").value = "Fourier";
  sheet.getCell("H2").value = 1;
  sheet.getCell("H3").value = 0;
  sheet.getCell("H4").value = -1;
  sheet.getCell("H5").value = 0;
  sheet.getCell("I1").value = "Complex";
  sheet.getCell("I2").value = "1+i";
  sheet.getCell("I3").value = "2";
  sheet.getCell("I4").value = "1-i";
  sheet.getCell("I5").value = "0";
  sheet.getCell("K1").value = "Inverse";
  sheet.getCell("K2").value = "4";
  sheet.getCell("K3").value = "0";
  sheet.getCell("K4").value = "0";
  sheet.getCell("K5").value = "4i";
  sheet.getCell("M1").value = "条件";
  sheet.getCell("N1").value = "装置A";
  sheet.getCell("O1").value = "装置B";
  sheet.getCell("P1").value = "装置C";
  sheet.getCell("M2").value = "朝";
  sheet.getCell("N2").value = 10;
  sheet.getCell("O2").value = 12;
  sheet.getCell("P2").value = 13;
  sheet.getCell("M3").value = "昼";
  sheet.getCell("N3").value = 11;
  sheet.getCell("O3").value = 15;
  sheet.getCell("P3").value = 16;
  sheet.getCell("M4").value = "夜";
  sheet.getCell("N4").value = 14;
  sheet.getCell("O4").value = 18;
  sheet.getCell("P4").value = 19;
  sheet.getCell("R1").value = "標本";
  sheet.getCell("S1").value = "装置A";
  sheet.getCell("T1").value = "装置B";
  sheet.getCell("R2").value = "朝";
  sheet.getCell("S2").value = 10;
  sheet.getCell("T2").value = 12;
  sheet.getCell("R3").value = "朝";
  sheet.getCell("S3").value = 11;
  sheet.getCell("T3").value = 13;
  sheet.getCell("R4").value = "夜";
  sheet.getCell("S4").value = 14;
  sheet.getCell("T4").value = 18;
  sheet.getCell("R5").value = "夜";
  sheet.getCell("S5").value = 15;
  sheet.getCell("T5").value = 19;
  sheet.getCell("U1").value = "PairA";
  sheet.getCell("V1").value = "PairB";
  sheet.getCell("U2").value = 1;
  sheet.getCell("V2").value = 10;
  sheet.getCell("U3").value = 2;
  sheet.getCell("V4").value = 30;
  sheet.getCell("U5").value = 4;
  sheet.getCell("V5").value = 40;
  sheet.getCell("U6").value = 5;
  sheet.getCell("V6").value = 50;
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createConsolidateWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-consolidate-"));
  const outputPath = path.join(dir, "consolidate.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Consolidate");
  sheet.getCell("A1").value = "Region";
  sheet.getCell("B1").value = "Q1";
  sheet.getCell("C1").value = "Q2";
  sheet.getCell("A2").value = "East";
  sheet.getCell("B2").value = 10;
  sheet.getCell("C2").value = 20;
  sheet.getCell("A3").value = "West";
  sheet.getCell("B3").value = 5;
  sheet.getCell("C3").value = 7;
  sheet.getCell("E1").value = "Region";
  sheet.getCell("F1").value = "Q2";
  sheet.getCell("G1").value = "Q1";
  sheet.getCell("E2").value = "East";
  sheet.getCell("F2").value = 3;
  sheet.getCell("G2").value = 2;
  sheet.getCell("E3").value = "North";
  sheet.getCell("F3").value = 8;
  sheet.getCell("G3").value = 9;
  const second = workbook.addWorksheet("追加");
  second.getCell("A1").value = "Region";
  second.getCell("B1").value = "Q1";
  second.getCell("C1").value = "Q2";
  second.getCell("A2").value = "East";
  second.getCell("B2").value = 11;
  second.getCell("C2").value = 13;
  second.getCell("A3").value = "South";
  second.getCell("B3").value = 4;
  second.getCell("C3").value = 6;
  workbook.definedNames.add("Consolidate!$A$1:$C$3", "ConsolidateBase");
  workbook.definedNames.add("'追加'!$A$1:$C$3", "ConsolidateExtra");
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createSubtotalWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-subtotal-"));
  const outputPath = path.join(dir, "subtotal.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Subtotal");
  sheet.addRow(["Group", "Amount"]);
  sheet.addRow(["A", 5]);
  sheet.addRow(["A", 20]);
  sheet.addRow(["B", 12]);
  sheet.addRow(["B", 8]);
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createOutlineWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-outline-"));
  const outputPath = path.join(dir, "outline.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Outline");
  sheet.addRow(["Region", "Jan", "Feb", "Mar", "Apr"]);
  sheet.addRow(["East", 10, 20, 30, 40]);
  sheet.addRow(["West", 15, 25, 35, 45]);
  sheet.addRow(["North", 12, 22, 32, 42]);
  sheet.addRow(["South", 18, 28, 38, 48]);
  sheet.addRow(["Total", 55, 95, 135, 175]);
  const imported = workbook.addWorksheet("ImportedOutline");
  imported.getCell("A1").value = "Imported outline";
  imported.getCell("B2").value = "Visible cells stop here";
  imported.getRow(9).height = 21;
  imported.getRow(10).height = 21;
  imported.getColumn(8).width = 12;
  imported.getColumn(9).width = 12;
  imported.getColumn(10).width = 12;
  await workbook.xlsx.writeFile(outputPath);
  await patchImportedOutlineWorksheet(outputPath);
  return outputPath;
}



async function patchImportedOutlineWorksheet(outputPath) {
  const zip = await JSZip.loadAsync(await readFile(outputPath));
  const worksheetPath = await worksheetPathForSheetName(zip, "ImportedOutline");
  if (!worksheetPath) return;
  const file = zip.file(worksheetPath);
  if (!file) return;
  let xml = await file.async("text");
  const outlinePrXml = '<outlinePr applyStyles="1"/>';
  if (/<outlinePr\b/.test(xml)) {
    xml = xml.replace(/<outlinePr\b[^>]*\/>|<outlinePr\b[^>]*>[\s\S]*?<\/outlinePr>/, outlinePrXml);
  } else if (/<sheetPr\b[^>]*\/>/.test(xml)) {
    xml = xml.replace(/<sheetPr\b([^>]*)\/>/, `<sheetPr$1>${outlinePrXml}</sheetPr>`);
  } else if (/<sheetPr\b[^>]*>/.test(xml)) {
    xml = xml.replace(/<sheetPr\b[^>]*>/, (tag) => `${tag}${outlinePrXml}`);
  } else {
    xml = xml.replace(/<worksheet\b[^>]*>/, (tag) => `${tag}<sheetPr>${outlinePrXml}</sheetPr>`);
  }
  const colsXml = '<cols><col min="8" max="9" width="12" customWidth="1" hidden="1" outlineLevel="1"/><col min="10" max="10" width="12" customWidth="1" collapsed="1"/></cols>';
  if (/<cols\b[\s\S]*?<\/cols>|<cols\b[^>]*\/>/.test(xml)) {
    xml = xml.replace(/<cols\b[\s\S]*?<\/cols>|<cols\b[^>]*\/>/, colsXml);
  } else {
    xml = xml.replace(/<sheetData\b/, `${colsXml}<sheetData`);
  }
  xml = xml.replace(/<row\b(?=[^>]*\br="(?:9|10)")[\s\S]*?<\/row>|<row\b(?=[^>]*\br="(?:9|10)")[^>]*\/>/g, "");
  const outlineRowsXml = '<row r="9" ht="21" customHeight="1" hidden="1" outlineLevel="1"/><row r="10" ht="21" customHeight="1" collapsed="1"/>';
  xml = xml.replace("</sheetData>", `${outlineRowsXml}</sheetData>`);
  zip.file(worksheetPath, xml);
  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
}



async function createTextToColumnsWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-text-columns-"));
  const outputPath = path.join(dir, "text-columns.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("TextColumns");
  sheet.getCell("A1").value = "05/24/2026,ZJ,100";
  sheet.getCell("A2").value = "06/01/2026,AA,200";
  sheet.getCell("A3").value = '"May 20, 2026",BB,300';
  sheet.getCell("B1").value = "既存";
  sheet.getCell("C2").value = "保持";
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createTextToColumnsAdvancedWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-text-columns-advanced-"));
  const outputPath = path.join(dir, "text-columns-advanced.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("TextColumnsAdvanced");
  sheet.getCell("A1").value = "1.234,50;123-";
  sheet.getCell("A2").value = "2.000,00;45-";
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createTextToColumnsFixedWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-text-columns-fixed-"));
  const outputPath = path.join(dir, "text-columns-fixed.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("FixedWidth");
  sheet.getCell("A1").value = "ALPHA001TOKYO";
  sheet.getCell("A2").value = "GAMMA010OSAKA";
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createSortColumnsWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-sort-columns-"));
  const outputPath = path.join(dir, "sort-columns.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("SortColumns");
  sheet.getCell("A1").value = "Field";
  sheet.getCell("B1").value = "West";
  sheet.getCell("C1").value = "East";
  sheet.getCell("D1").value = "North";
  sheet.getCell("A2").value = "Score";
  sheet.getCell("B2").value = 30;
  sheet.getCell("C2").value = 10;
  sheet.getCell("D2").value = 20;
  sheet.getCell("A3").value = "Code";
  sheet.getCell("B3").value = "W";
  sheet.getCell("C3").value = "E";
  sheet.getCell("D3").value = "N";
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createSortCaseWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-sort-case-"));
  const outputPath = path.join(dir, "sort-case.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("SortCase");
  sheet.addRow(["Key", "Value"]);
  sheet.addRow(["A", 1]);
  sheet.addRow(["a", 2]);
  sheet.addRow(["B", 3]);
  sheet.addRow(["b", 4]);
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}



async function createSortBlankWorkbook() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "websheet-sort-blank-"));
  const outputPath = path.join(dir, "sort-blank.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("SortBlank");
  sheet.addRow(["Name", "Score"]);
  sheet.addRow(["Blank", null]);
  sheet.addRow(["Low", 5]);
  sheet.addRow(["High", 20]);
  sheet.addRow(["Mid", 10]);
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}

export {
  addWorkbookCustomViews,
  addWorkbookDefinedNames,
  appendImportedAutoFilterDxfs,
  clickRibbonTab,
  createAdvancedFilterUniqueSubsetWorkbook,
  createAdvancedFilterUniqueWorkbook,
  createAutoFilterBlankWorkbook,
  createAutoFilterDateWorkbook,
  createAutoFilterMenuWorkbook,
  createConsolidateWorkbook,
  createCustomViewWorkbook,
  createDataAnalysisWorkbook,
  createDelimitedImportWorkbook,
  createFlashFillWorkbook,
  createForecastWorkbook,
  createFormulaUiWorkbook,
  createHomeToolbarWorkbook,
  createImportedAutoFilterWorkbook,
  createMacroVbaSubsetWorkbook,
  createOutlineWorkbook,
  createPageBackgroundImage,
  createPowerQueryTransformWorkbook,
  createRelationshipsWorkbook,
  createRemoveDuplicatesWorkbook,
  createRichNativeShapeWorkbook,
  createSortBlankWorkbook,
  createSortCaseWorkbook,
  createSortColumnsWorkbook,
  createSortDialogWorkbook,
  createSubtotalWorkbook,
  createTextToColumnsAdvancedWorkbook,
  createTextToColumnsFixedWorkbook,
  createTextToColumnsWorkbook,
  createValidationMenuWorkbook,
  createWhatIfWorkbook,
  decodeCompressedImageAssets,
  decodeStandaloneModelPayload,
  e2eProgress,
  exportedDefinedNamesWorkbookState,
  exportedOutlineWorkbookState,
  exportedPageBreakWorkbookState,
  exportedSheetViewWorkbookState,
  exportedStandaloneModelFromHtml,
  importedStyleDxfCount,
  insertWorksheetSortState,
  jsonScriptPayload,
  looksLikeStandaloneWorkbookHtml,
  patchImportedOutlineWorksheet,
  patchValidationMenuWorkbookDataValidations,
  replaceStandaloneModelPayload,
  replaceStandaloneModelScriptRaw,
  replaceTableAutoFilter,
  replaceWorksheetAutoFilter,
  waitForUploadedWorkbook,
  websheetModelScript,
  worksheetPathForSheetName,
  worksheetXmlForSheetName,
  xmlAttr,
  xmlAttributeEscape,
  xmlTextEscape,
};
