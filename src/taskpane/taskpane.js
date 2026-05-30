/* global console, document, Excel, Office, window */

import { buildGraph } from "./graph.js";
import { renderGraph } from "./renderer.js";

// When Office.js is ready, show the main UI and wire up the buttons.
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";

    document.getElementById("btn-precedents").onclick = () => runMode("precedents");
    document.getElementById("btn-dependents").onclick = () => runMode("dependents");

    // If a keyboard shortcut fired before the task pane was open,
    // commands.js will have stored the requested mode in window._pendingMode.
    // Run it now that the pane is ready.
    if (window._pendingMode) {
      runMode(window._pendingMode);
      window._pendingMode = null;
    }
  }
});

// Reads the selected cell, builds the graph, and draws it.
// mode is "precedents" or "dependents".
async function runMode(mode) {
  try {
    await Excel.run(async (context) => {
      const selection = context.workbook.getSelectedRange();
      selection.load(["address", "formulas"]);
      await context.sync();

      // If a multi-cell range is selected, use only the top-left cell.
      const address = selection.address.split(":")[0];
      const formula = String(selection.formulas[0][0]);

      // Update the tracking labels at the top of the pane
      document.getElementById("tracking-label").textContent = "Tracking: " + address;
      document.getElementById("formula-label").textContent =
        formula.startsWith("=") ? "Formula: " + formula : "";

      // Build the dependency data and draw the graph
      const graphData = await buildGraph(address, mode, context);
      renderGraph(graphData);
    });
  } catch (error) {
    console.error("FormulaTracker error:", error);
  }
}

// Exposed on window so commands.js (keyboard shortcut handler) can call these
// from the shared runtime when the task pane is already open.
window.showPrecedents = () => runMode("precedents");
window.showDependents = () => runMode("dependents");
