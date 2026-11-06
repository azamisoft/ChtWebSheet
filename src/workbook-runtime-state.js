// Shared workbook runtime bindings, initialized outside the feature-module graph.
export const tableDeclarationScope = { current: null };
export const graphemes = typeof Intl.Segmenter === "function"
  ? new Intl.Segmenter(undefined, {granularity: "grapheme"})
  : null;
