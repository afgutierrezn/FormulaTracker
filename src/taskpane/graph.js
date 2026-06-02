/* global console */

import { parseFormula } from "./fallback-parser.js";

// Loads the formula and computed value for a single cell.
// address: fully-qualified like "Sheet1!A1"
// Returns { formula, value } where formula is the "=..." string (or "" if the cell has no formula)
// and value is the number, string, or boolean Excel displays in the cell.
export async function getPrecedentCell(address, context) {
  const { sheet, cell } = splitAddress(address);
  const range = context.workbook.worksheets.getItem(sheet).getRange(cell);
  range.load(["formulas", "values"]);
  await context.sync();
  const formula = String(range.formulas[0][0]);
  return {
    formula: formula.startsWith("=") ? formula : "",
    value: range.values[0][0],
  };
}

// Given a list of fully-qualified cell addresses, finds all cells that reference
// any of them (their "dependents") and returns a flat array of results.
// Each result has: address (full), sheet name, and count (how many times the
// dependent's formula references any address from the input list).
export async function getDependentsRows(addressList, context) {
  const addressSet = new Set(addressList);
  const allDeps = new Set();
  // scanDeps tracks only addresses found by the text-scan fallback, which are
  // always valid. Used as a safe retry set if the API returns a bad address.
  const scanDeps = new Set();

  // Step 1: for each source cell, try getDirectDependents() AND always also
  // run the text scan. On Mac, the API can return invalid addresses for cells
  // that contain a hardcoded value (no formula), causing Step 2 to fail.
  // Running both ensures we always have a known-good fallback set (scanDeps).
  for (const address of addressList) {
    const { sheet, cell } = splitAddress(address);
    const range = context.workbook.worksheets.getItem(sheet).getRange(cell);

    try {
      const deps = range.getDirectDependents();
      deps.areas.load("address");
      await context.sync();
      for (const area of deps.areas.items) {
        const normalized = normalizeAddress(area.address, sheet);
        if (normalized.includes(":")) {
          expandRange(normalized, sheet).forEach((a) => allDeps.add(a));
        } else {
          allDeps.add(normalized);
        }
      }
    } catch (_) {
      // getDirectDependents not available or threw — scan below will cover this.
    }

    // Always run the text scan regardless of whether the API found results.
    // The scan is the authoritative source for same-sheet dependents.
    try {
      const fallback = await scanForDependents(address, sheet, context);
      fallback.forEach((a) => {
        allDeps.add(a);
        scanDeps.add(a);
      });
    } catch (_) {
      // Scan also failed — skip this address.
    }
  }

  // Scan every other sheet for cross-sheet dependents.
  // Needed for hardcoded-value cells where getDirectDependents() is unreliable on Mac.
  const allSheets = context.workbook.worksheets;
  allSheets.load("items/name");
  await context.sync();

  for (const ws of allSheets.items) {
    for (const address of addressList) {
      const { sheet: sourceSheet } = splitAddress(address);
      if (ws.name === sourceSheet) continue; // already scanned above
      try {
        const crossSheet = await scanForDependents(address, ws.name, context);
        crossSheet.forEach((a) => {
          allDeps.add(a);
          scanDeps.add(a);
        });
      } catch (_) {}
    }
  }

  if (allDeps.size === 0) return [];

  // Step 2: batch-load the formula for every unique dependent cell.
  // If this throws (because the API returned an invalid address on Mac),
  // fall back to scan-only results which are guaranteed to be valid addresses.
  let depRanges = loadDepFormulas(allDeps, context);
  try {
    await context.sync();
  } catch (_) {
    // Reload using only the scan-derived addresses (always valid).
    depRanges = loadDepFormulas(scanDeps, context);
    await context.sync();
  }

  // Step 3: count how many times each dependent references any source address.
  const result = [];
  for (const [depAddr, range] of Object.entries(depRanges)) {
    const { sheet } = splitAddress(depAddr);
    const formula = String(range.formulas[0][0]);
    const refs = parseFormula(formula);

    let count = 0;
    for (const ref of refs) {
      const normalized = normalizeAddress(ref, sheet);
      if (normalized.includes(":")) {
        count += expandRange(normalized, sheet).filter((a) => addressSet.has(a)).length;
      } else {
        if (addressSet.has(normalized)) count++;
      }
    }

    result.push({ address: depAddr, sheet, count: Math.max(count, 1) });
  }

  return result;
}

// Issues range.load("formulas") for each address and returns the range map.
// Does NOT call context.sync() — the caller must do that.
function loadDepFormulas(deps, context) {
  const ranges = {};
  for (const depAddr of deps) {
    const { sheet, cell } = splitAddress(depAddr);
    const r = context.workbook.worksheets.getItem(sheet).getRange(cell);
    r.load("formulas");
    ranges[depAddr] = r;
  }
  return ranges;
}

// Expands a range address like "Sheet1!A1:C2" into individual cell addresses:
// ["Sheet1!A1", "Sheet1!B1", "Sheet1!C1", "Sheet1!A2", "Sheet1!B2", "Sheet1!C2"]
export function expandRange(rangeAddr, defaultSheet) {
  const normalized = normalizeAddress(rangeAddr, defaultSheet);
  const bangPos = normalized.lastIndexOf("!");
  const sheet = bangPos !== -1 ? normalized.slice(0, bangPos) : defaultSheet;
  const rangePart = bangPos !== -1 ? normalized.slice(bangPos + 1) : normalized;

  if (!rangePart.includes(":")) return [normalized]; // single cell, not a range

  const [startCell, endCell] = rangePart.split(":");
  const startCol = colLetterToIndex(startCell.replace(/[\$0-9]/g, ""));
  const startRow = parseInt(startCell.replace(/[\$A-Z]/g, ""), 10);
  const endCol = colLetterToIndex(endCell.replace(/[\$0-9]/g, ""));
  const endRow = parseInt(endCell.replace(/[\$A-Z]/g, ""), 10);

  const cells = [];
  for (let row = startRow; row <= endRow; row++) {
    for (let col = startCol; col <= endCol; col++) {
      cells.push(sheet + "!" + columnIndexToLetter(col) + row);
    }
  }
  return cells;
}

// Scans every used cell in the sheet and returns those whose formula references targetAddress.
// Used as a fallback when getDirectDependents() is not supported by this Excel version.
// Note: only scans the active sheet — cross-sheet dependents won't appear in fallback mode.
async function scanForDependents(targetAddress, sheet, context) {
  const worksheet = context.workbook.worksheets.getItem(sheet);
  let usedRange;
  try {
    usedRange = worksheet.getUsedRange();
    usedRange.load(["formulas", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await context.sync();
  } catch (_) {
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
export function splitAddress(address) {
  const bangIndex = address.lastIndexOf("!");
  if (bangIndex === -1) return { sheet: "Sheet1", cell: address };
  return {
    sheet: address.slice(0, bangIndex).replace(/^'|'$/g, ""),
    cell: address.slice(bangIndex + 1),
  };
}

// Ensures an address always includes a sheet name.
// "A1" → "Sheet1!A1".  "Sheet2!B3" → "Sheet2!B3".  "'My Sheet'!C4" → "My Sheet!C4".
export function normalizeAddress(address, defaultSheet) {
  if (!address.includes("!")) return defaultSheet + "!" + address;
  return address.replace(/^'(.*)'!/, "$1!");
}

// Converts a zero-based column index to a letter. 0 → "A", 25 → "Z", 26 → "AA".
export function columnIndexToLetter(index) {
  let letter = "";
  let n = index + 1;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

// Converts a column letter like "A", "Z", "AA" to a zero-based index.
function colLetterToIndex(letters) {
  let index = 0;
  for (let i = 0; i < letters.length; i++) {
    index = index * 26 + (letters.charCodeAt(i) - 64);
  }
  return index - 1;
}
