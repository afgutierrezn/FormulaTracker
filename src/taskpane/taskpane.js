/* global console, Excel, Office, window */

import { parseFormula as parseFormulaTokens } from "./formula-parser.js";
import { buildAllFunctionResults } from "./fn-evaluator.js";
import { getPrecedentCell, getDependentsRows, expandRange, splitAddress, normalizeAddress } from "./graph.js";
import { renderPrecedents, renderDependents, renderError } from "./renderer.js";

// When Office.js is ready, show the main UI and wire up the tab buttons.
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    const btnP = document.getElementById("btn-precedents");
    const btnD = document.getElementById("btn-dependents");

    btnP.addEventListener("click", () => {
      setActiveTab("precedents");
      runMode("precedents");
    });
    btnD.addEventListener("click", () => {
      setActiveTab("dependents");
      runMode("dependents");
    });

    // Esc clears the row selection and resets the formula bar highlight.
    // (Office.addin.hide() requires SharedRuntime which is not active on this Excel version.)
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") {
        e.preventDefault();
        document.querySelectorAll("tr.selected-row").forEach((r) => r.classList.remove("selected-row"));
        const formulaText = document.getElementById("formula-text");
        if (formulaText) formulaText.innerHTML = formulaText.textContent; // strip <mark> highlight
      }
    });

    // If a keyboard shortcut fired before the task pane was open,
    // commands.js will have stored the mode in window._pendingMode.
    if (window._pendingMode) {
      const mode = window._pendingMode;
      window._pendingMode = null;
      setActiveTab(mode);
      runMode(mode);
    }
  }
});

// Highlights the correct tab button.
function setActiveTab(mode) {
  document.getElementById("btn-precedents").classList.toggle("active", mode === "precedents");
  document.getElementById("btn-dependents").classList.toggle("active", mode === "dependents");
}

// Reads the selected cell(s), fetches data from Excel, and renders the table.
async function runMode(mode) {
  try {
    if (mode === "precedents") {
      await runPrecedents();
    } else {
      await runDependents();
    }
  } catch (err) {
    console.error("FormulaTracker error:", err);
    renderError("Could not load data. Try a different cell.");
  }
}

// Runs Explore Precedents for the active cell.
async function runPrecedents() {
  await Excel.run(async (context) => {
    // Read the selected cell (use only the top-left if a range is selected)
    const selection = context.workbook.getSelectedRange();
    selection.load("address");
    await context.sync();

    const rawAddress = selection.address.split(",")[0].split(":")[0];
    const defaultSheet = splitAddress(rawAddress).sheet;
    const address = normalizeAddress(rawAddress, defaultSheet);

    // Load this cell's formula and value
    const { formula, value } = await getPrecedentCell(address, context);

    // Parse the formula into a token tree
    const tokens = parseFormulaTokens(formula);

    // Enrich the token tree: load values for cell refs and evaluate functions
    await enrichTokenTree(tokens, defaultSheet, context);

    // Render the table
    renderPrecedents({ activeAddress: address, formula, value, tokens });
  });
}

// Runs Explore Dependents for the selected cell(s).
async function runDependents() {
  await Excel.run(async (context) => {
    // Read the full selection — may be a single cell, a range, or multiple areas
    const selection = context.workbook.getSelectedRange();
    selection.load("address");
    await context.sync();

    // Split into individual cell addresses (handles "Sheet1!A1:C3" and "Sheet1!A1,Sheet1!B5")
    const rawAddr = selection.address;
    const defaultSheet = splitAddress(rawAddr.split(",")[0].split(":")[0]).sheet;
    const addressList = parseSelectionAddress(rawAddr, defaultSheet);

    // Use the first address as the label for row 1
    const firstAddress = addressList[0];

    // Find all cells that depend on any of the selected cells
    const rows = await getDependentsRows(addressList, context);

    renderDependents(firstAddress, rows);
  });
}

// Fills in cellValue/cellFormula on cell_ref tokens and computedValue on function tokens.
// All cell ref values are loaded in one batch sync; functions are evaluated in a second batch.
async function enrichTokenTree(tokens, defaultSheet, context) {
  // --- Batch 1: load values and formulas for all cell_ref tokens ---
  const cellRefTokens = [];
  collectCellRefs(tokens, cellRefTokens);

  const rangeMap = {}; // normalized address → { range, tokens[] }
  for (const token of cellRefTokens) {
    const full = normalizeAddress(token.address, defaultSheet);
    if (!rangeMap[full]) {
      const { sheet, cell } = splitAddress(full);
      const range = context.workbook.worksheets.getItem(sheet).getRange(cell);
      range.load(["values", "formulas"]);
      rangeMap[full] = { range, tokens: [] };
    }
    rangeMap[full].tokens.push(token);
  }

  if (Object.keys(rangeMap).length > 0) {
    await context.sync();
    for (const { range, tokens: toks } of Object.values(rangeMap)) {
      const cellValue = range.values[0][0];
      const cellFormula = String(range.formulas[0][0]);
      for (const tok of toks) {
        tok.cellValue = cellValue;
        tok.cellFormula = cellFormula.startsWith("=") ? cellFormula : "";
      }
    }
  }

  // --- Batch 2: evaluate all function tokens using workbook.functions ---
  let fnEvals;
  try {
    fnEvals = buildAllFunctionResults(tokens, defaultSheet, context);
  } catch (e) {
    // workbook.functions not available on this Excel version — skip
    return;
  }

  if (fnEvals.length > 0) {
    try {
      fnEvals.forEach(({ fnResult }) => fnResult.load("value"));
      await context.sync();
      fnEvals.forEach(({ token, fnResult }) => {
        token.computedValue = fnResult.value;
      });
    } catch (_) {
      // Evaluation failed (API unavailable or unsupported function) — leave computedValue as null
    }
  }
}

// Recursively collects all cell_ref tokens from the token tree into the result array.
function collectCellRefs(tokens, result) {
  for (const token of tokens) {
    if (token.type === "cell_ref") {
      result.push(token);
    } else if (token.type === "function") {
      collectCellRefs(token.args, result);
    }
  }
}

// Converts a selection address string into a list of fully-qualified individual cell addresses.
// Handles: "Sheet1!A1", "Sheet1!A1:C3", "Sheet1!A1,Sheet1!B5", mixed ranges and singles.
function parseSelectionAddress(rawAddr, defaultSheet) {
  const addresses = [];
  // Split on commas for multi-area selections
  const areas = rawAddr.split(",");
  for (const area of areas) {
    const trimmed = area.trim();
    if (trimmed.includes(":")) {
      // It's a range — expand to individual cells
      const normalized = normalizeAddress(trimmed, defaultSheet);
      expandRange(normalized, defaultSheet).forEach((a) => addresses.push(a));
    } else {
      addresses.push(normalizeAddress(trimmed, defaultSheet));
    }
  }
  return addresses;
}

// Exposed on window so commands.js (keyboard shortcut handler) can call them
// from the shared runtime when the task pane is already open.
window.showPrecedents = () => {
  setActiveTab("precedents");
  runMode("precedents");
};
window.showDependents = () => {
  setActiveTab("dependents");
  runMode("dependents");
};
