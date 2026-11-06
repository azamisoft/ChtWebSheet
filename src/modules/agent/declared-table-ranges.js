// Table declarations apply to one CWS batch. The sheet index is resolved when
// createTable runs, after any preceding sheet creation, rename or reorder.
import { tableDeclarationScope } from "../../workbook-runtime-state.js";

export function declaredTableRanges(previousOps = []) {
  const declarations = new Map();
  for (const op of previousOps) {
    if (op.type === "createTable" && op.name && Number.isInteger(op.sheetIndex)) {
      addDeclaredTable(declarations, op.sheetIndex, op.name);
    }
  }
  return declarations;
}

function addDeclaredTable(declarations, index, name) {
  if (!declarations.has(index)) declarations.set(index, new Set());
  declarations.get(index).add(name.toLowerCase());
}

export function registerDeclaredTable(index, name) {
  if (tableDeclarationScope.current) addDeclaredTable(tableDeclarationScope.current, index, name);
}

export function withDeclaredTableRanges(declarations, apply) {
  const previous = tableDeclarationScope.current;
  tableDeclarationScope.current = declarations;
  try {
    return apply();
  } finally {
    tableDeclarationScope.current = previous;
  }
}

export function isDeclaredTableRange(sheetIndex, table) {
  const names = tableDeclarationScope.current?.get(sheetIndex);
  return Boolean(names && [table.name, table.displayName].some(name => typeof name === "string" && names.has(name.toLowerCase())));
}
