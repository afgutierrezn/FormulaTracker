/* global console */

import { parseFormula } from "./fallback-parser.js";

// Set to true to skip the Excel API and always use the regex parser.
// Useful for testing the fallback without needing an older Excel version.
const FORCE_FALLBACK = false;

// Main function. Builds the dependency graph starting from one cell.
// startAddress: fully-qualified address like "Sheet1!C1"
// mode: "precedents" (what feeds into this cell) or "dependents" (what uses this cell)
// context: the Excel context object from inside an Excel.run() call
// Returns { nodes: [...], edges: [...] }
export async function buildGraph(startAddress, mode, context) {
  const nodes = [];
  const edges = [];
  const visited = new Set();
  const rootSheet = splitAddress(startAddress).sheet;

  await traverse(startAddress, mode, context, nodes, edges, visited, 5, rootSheet);

  return { nodes, edges };
}

// Visits one cell, records it as a node, finds its connections, and recurses into each one.
async function traverse(address, mode, context, nodes, edges, visited, depth, rootSheet) {
  if (visited.has(address)) return;
  visited.add(address);

  const { sheet, cell } = splitAddress(address);

  // Load this cell's formula from Excel
  const range = context.workbook.worksheets.getItem(sheet).getRange(cell);
  range.load(["formulas", "address"]);
  await context.sync();

  const rawValue = range.formulas[0][0];
  const formula = typeof rawValue === "string" && rawValue.startsWith("=") ? rawValue : "";

  // Record this cell as a node in the graph
  if (!nodes.some((n) => n.id === address)) {
    nodes.push({
      id: address,
      label: makeLabel(sheet, cell, rootSheet),
      sheet: sheet,
      formula: formula,
    });
  }

  // Stop here if we've hit the depth limit
  if (depth === 0) return;

  // Find the cells connected to this one
  const connected = await getConnectedCells(range, address, sheet, formula, mode, context);

  // Add an edge for each connection, then recurse into it.
  // Edges always point in the direction of data flow:
  //   precedents mode: precedent → current cell  (A1 feeds C1)
  //   dependents mode: current cell → dependent  (C1 feeds D2)
  for (const connectedAddress of connected) {
    const edgeSource = mode === "precedents" ? connectedAddress : address;
    const edgeTarget = mode === "precedents" ? address : connectedAddress;

    if (visited.has(connectedAddress)) {
      // Already visited — this is a circular reference
      edges.push({ source: edgeSource, target: edgeTarget, label: "circular" });
    } else {
      edges.push({ source: edgeSource, target: edgeTarget, label: "" });
      await traverse(connectedAddress, mode, context, nodes, edges, visited, depth - 1, rootSheet);
    }
  }
}

// Returns the list of connected cell addresses.
// Tries the Excel API first; if unavailable, falls back to the regex parser.
async function getConnectedCells(range, address, sheet, formula, mode, context) {
  if (!FORCE_FALLBACK) {
    try {
      return await getConnectedViaApi(range, sheet, mode, context);
    } catch (e) {
      // API not available on this Excel version — use the regex fallback
      console.log("Excel API unavailable, switching to fallback parser:", e.message);
    }
  }
  return getConnectedViaFallback(address, sheet, formula, mode, context);
}

// Gets connected cells using the Excel getDirectPrecedents / getDirectDependents API.
async function getConnectedViaApi(range, sheet, mode, context) {
  const connected = mode === "precedents"
    ? range.getDirectPrecedents()
    : range.getDirectDependents();

  connected.areas.load("address");
  await context.sync();

  return connected.areas.items.map((area) => normalizeAddress(area.address, sheet));
}

// Gets connected cells using the regex formula parser (no Excel API needed).
// For precedents: parses the formula string directly.
// For dependents: scans all cells in the sheet for formulas that reference this cell.
async function getConnectedViaFallback(address, sheet, formula, mode, context) {
  if (mode === "precedents") {
    const refs = parseFormula(formula);
    return refs.map((ref) => normalizeAddress(ref, sheet));
  } else {
    return await scanForDependents(address, sheet, context);
  }
}

// Scans every used cell in the sheet and returns those whose formula references targetAddress.
// Note: only scans the active sheet — cross-sheet dependents are not detected in fallback mode.
async function scanForDependents(targetAddress, sheet, context) {
  const worksheet = context.workbook.worksheets.getItem(sheet);
  let usedRange;
  try {
    usedRange = worksheet.getUsedRange();
    usedRange.load(["formulas", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await context.sync();
  } catch (e) {
    return []; // sheet is empty or getUsedRange failed
  }

  const dependents = [];

  for (let row = 0; row < usedRange.rowCount; row++) {
    for (let col = 0; col < usedRange.columnCount; col++) {
      const cellFormula = String(usedRange.formulas[row][col]);
      if (!cellFormula.startsWith("=")) continue;

      const refs = parseFormula(cellFormula).map((r) => normalizeAddress(r, sheet));

      if (refs.includes(targetAddress)) {
        const colLetter = columnIndexToLetter(usedRange.columnIndex + col);
        const rowNumber = usedRange.rowIndex + row + 1;
        dependents.push(sheet + "!" + colLetter + rowNumber);
      }
    }
  }

  return dependents;
}

// Splits "Sheet1!A1" into { sheet: "Sheet1", cell: "A1" }.
// Strips surrounding quotes from sheet names like 'My Sheet'.
function splitAddress(address) {
  const bangIndex = address.lastIndexOf("!");
  if (bangIndex === -1) {
    return { sheet: "Sheet1", cell: address };
  }
  return {
    sheet: address.slice(0, bangIndex).replace(/^'|'$/g, ""),
    cell: address.slice(bangIndex + 1),
  };
}

// Ensures an address always includes a sheet name.
// "A1" → "Sheet1!A1". "Sheet2!B3" → "Sheet2!B3". "'My Sheet'!C4" → "My Sheet!C4".
function normalizeAddress(address, defaultSheet) {
  if (!address.includes("!")) return defaultSheet + "!" + address;
  // Strip surrounding quotes from the sheet name part e.g. 'My Sheet'!A1 → My Sheet!A1
  return address.replace(/^'(.*)'!/, "$1!");
}

// Makes the display label for a node.
// Same-sheet cells show just the cell address (e.g. "A1").
// Cross-sheet cells show sheet and cell (e.g. "Sheet2 · B3").
function makeLabel(sheet, cell, rootSheet) {
  if (sheet === rootSheet) return cell;
  return sheet + " · " + cell;
}

// Converts a zero-based column index to a column letter (0 → "A", 25 → "Z", 26 → "AA").
function columnIndexToLetter(index) {
  let letter = "";
  let n = index + 1;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}
