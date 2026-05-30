/* global console */

// Extracts all cell references from a formula string.
// Works entirely with text — no Excel API calls needed.
// Examples:
//   parseFormula("=A1+B2")           → ["A1", "B2"]
//   parseFormula("=Sheet2!C3*D4")    → ["Sheet2!C3", "D4"]
//   parseFormula("=SUM(A1:B5)")      → ["A1:B5"]
//   parseFormula("=42")              → []
export function parseFormula(formula) {
  if (!formula || !formula.startsWith("=")) {
    return [];
  }

  // Matches optional sheet prefix ('My Sheet'!A1 or Sheet2!A1)
  // followed by a cell address like A1, $A$1, A1:B3, $A$1:$B$3
  const pattern =
    /(?:'[^']+'|[A-Za-z0-9_]+)?!?\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?/g;

  const matches = formula.match(pattern) || [];

  // Filter out plain function names that accidentally matched the pattern.
  // A real cell ref always contains at least one digit.
  const cellRefs = matches.filter((m) => /[0-9]/.test(m));

  return cellRefs;
}
