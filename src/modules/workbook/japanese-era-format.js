// Only unquoted date tokens participate. In particular, E+00 in a numeric format is an exponent.
export function hasJapaneseEraTokens(section) {
  if (!/[ge]/i.test(section)) return false;
  const tokens = section.replace(/"[^"]*"|\\./g, " ");
  const locale = tokens.match(/\[\$[^\]]*?-([^\]]+)\]/i)?.[1];
  const japanese = /^(?:[0-9a-f]*0411|411|ja(?:-JP)?)$/i.test(locale || "");
  if (locale && !japanese) return false;
  const words = tokens.replace(/\[[^\]]*\]|[eE][+-]0+/g, " ").match(/[a-z]+/gi) || [];
  return words.some(word => /^[ymdhseg]+$/i.test(word) && (/g/i.test(word) || (japanese && /e/i.test(word))));
}

export function expandJapaneseEraTokens(section, date) {
  const long = new Intl.DateTimeFormat("ja-JP-u-ca-japanese-nu-latn", { era: "long", year: "numeric" }).formatToParts(date);
  const name = long.find(part => part.type === "era").value;
  const year = long.find(part => part.type === "year").value.replace(/^元$/, "1");
  return section.replace(/"[^"]*"|\\.|\[[^\]]*\]|[a-z]+/gi, word => {
    if (!/^[ymdhseg]+$/i.test(word)) return word;
    return word.replace(/g{1,3}|e{1,2}/gi, token => {
      if (/e/i.test(token)) return '"' + year.padStart(token.length, "0") + '"';
      if (token.length === 3) return '"' + name + '"';
      if (token.length === 2) return '"' + Array.from(name)[0] + '"';
      const narrow = new Intl.DateTimeFormat("ja-JP-u-ca-japanese", { era: "narrow", year: "numeric" }).formatToParts(date);
      return '"' + narrow.find(part => part.type === "era").value + '"';
    });
  });
}
