# FormulaTracker

An Excel add-in for Mac that shows you how your formulas connect to each other as a visual diagram.

Select a cell and click **Show Precedents** to see which cells feed into it, or **Show Dependents** to see which cells use it. The diagram traces up to 5 levels deep and handles circular references without crashing.

![FormulaTracker task pane diagram](assets/icon-128.png)

---

## What it does

- **Precedents** — shows all the cells that a formula depends on, traced recursively
- **Dependents** — shows all the cells that reference a given cell
- **Cross-sheet references** — labels nodes with their sheet name (e.g. `Sheet2 · B3`)
- **Circular references** — marks the offending node in red with a `↩` symbol instead of freezing
- **Dynamic reference warning** — shows a warning if a formula uses `INDIRECT`, `OFFSET`, or `INDEX`, which can't be fully traced

---

## Requirements

- Excel for Mac **16.55 or later**
- Node.js (install via [nodejs.org](https://nodejs.org) or `brew install node`)

---

## Running the add-in locally

```bash
# Install dependencies (first time only)
npm install

# Start the dev server and sideload into Excel
npm start
```

This opens Excel with the add-in already loaded. Click **Show Task Pane** in the Home tab ribbon to open the panel.

To stop:
```bash
npm stop
```

---

## How to use it

1. Open the FormulaTracker panel via the **Home → Show Task Pane** ribbon button
2. Click any cell that contains a formula
3. Click **Show Precedents** to see what feeds into it, or **Show Dependents** to see what uses it
4. The panel updates with a top-down diagram and shows the cell address and formula at the top

---

## File structure

```
src/
  taskpane/
    taskpane.html        — task pane layout and buttons
    taskpane.js          — wires up Office.js, reads selected cell, calls graph + renderer
    fallback-parser.js   — regex-based formula parser for older Excel versions
    graph.js             — builds the node/edge data by traversing dependencies
    renderer.js          — draws the graph using Cytoscape.js
  commands/
    commands.js          — keyboard shortcut action handlers (registered but Mac support pending)
    commands.html        — function file entry point
manifest.xml             — Office add-in manifest
shortcuts.json           — keyboard shortcut definitions (Cmd+Shift+[ and Cmd+Shift+])
```

---

## Known limitations

- **Keyboard shortcuts** — `Cmd+Shift+[` and `Cmd+Shift+]` are registered but currently not firing on Excel for Mac in the local development setup. The task pane buttons are the primary interface.
- **Dynamic references** — formulas using `INDIRECT()`, `OFFSET()`, or `INDEX()` as references cannot be fully traced. The add-in shows a warning when it encounters these.
- **Dependents fallback** — on older Excel versions (pre-16.55) that don't support `getDirectDependents()`, the fallback scans only the active sheet. Cross-sheet dependents won't appear.

---

## Tech stack

- [Office.js](https://learn.microsoft.com/en-us/office/dev/add-ins/) — Excel add-in API
- [Cytoscape.js](https://cytoscape.org/) — graph rendering
- [Webpack](https://webpack.js.org/) — bundler
- Vanilla JavaScript (no framework)
