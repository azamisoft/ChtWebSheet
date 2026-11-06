// Return offsets so rich-text runs can retain their own fonts and styles even
// when a word crosses a run boundary. Fall back to characters for an overwide
// token instead of clipping it or forcing a whole word beyond the cell.
import { graphemes } from "../../workbook-runtime-state.js";

export function measuredLineRanges(text, maxWidth, measure) {
  if (!text) return [[0, 0]];
  const tokens = /\p{Sc}[ \t]*[+−-]?\p{N}+(?:[.,]\p{N}+)*(?:[ \t]+[\p{Script=Latin}]{1,4}(?![\p{Script=Latin}]))?|[\p{Script=Latin}\p{N}_][\p{Script=Latin}\p{M}\p{N}'’]*|[ \t]+|\P{M}\p{M}*|\p{M}/gu;
  const lines = [];
  let start = 0;
  let end = 0;
  for (const match of text.matchAll(tokens)) {
    const token = match[0];
    const next = match.index + token.length;
    if (/^[ \t]+$/.test(token) || measure(start, next) <= maxWidth) {
      end = next;
      continue;
    }
    if (start < end) {
      lines.push([start, end]);
      start = end;
    }
    if (measure(start, next) <= maxWidth) {
      end = next;
      continue;
    }
    const parts = graphemes ? graphemes.segment(token) : token;
    for (const part of parts) {
      const size = (part.segment ?? part).length;
      if (end > start && measure(start, end + size) > maxWidth) {
        lines.push([start, end]);
        start = end;
      }
      end += size;
    }
  }
  lines.push([start, end]);
  return lines;
}

export function sliceStyledLine(runs, start, end) {
  const result = [];
  let offset = 0;
  for (const run of runs) {
    const next = offset + run.text.length;
    if (offset < end && next > start) {
      result.push({...run, text: run.text.slice(Math.max(0, start - offset), Math.min(run.text.length, end - offset))});
    }
    offset = next;
    if (offset >= end) break;
  }
  return result;
}
