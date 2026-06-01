/* global Office, window */

Office.onReady(() => {
  // Office.js is ready in the shared runtime.
});

// Associate actions outside onReady so they register synchronously at script load,
// before Office.js completes async initialisation.
Office.actions.associate("ShowPrecedents", showPrecedentsAction);
Office.actions.associate("ShowDependents", showDependentsAction);

// Called when the user presses Cmd+[ (Mac) or Ctrl+[ (Windows).
// Shows the precedents of the selected cell.
async function showPrecedentsAction(event) {
  if (typeof window.showPrecedents === "function") {
    // Task pane is already loaded — call the function directly.
    window.showPrecedents();
  } else {
    // Task pane hasn't loaded yet — store the mode so taskpane.js picks it up on init.
    window._pendingMode = "precedents";
  }
  await Office.addin.showAsTaskpane();
  event.completed();
}

// Called when the user presses Cmd+] (Mac) or Ctrl+] (Windows).
// Shows the dependents of the selected cell.
async function showDependentsAction(event) {
  if (typeof window.showDependents === "function") {
    window.showDependents();
  } else {
    window._pendingMode = "dependents";
  }
  await Office.addin.showAsTaskpane();
  event.completed();
}

