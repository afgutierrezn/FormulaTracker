/* global document, Excel */

// Guard: only attach the keyboard listener once per page load.
let _navAttached = false;

// Draws the Explore Precedents table and formula bar.
// enrichedData has:
//   activeAddress — the cell being explored, e.g. "Sheet1!A1"
//   formula       — full formula string like "=SUM(B1, C1)"
//   value         — computed value of the active cell
//   tokens        — token tree from formula-parser.js, enriched with cellValue / computedValue
export function renderPrecedents(enrichedData) {
  const { activeAddress, formula, value, tokens } = enrichedData;

  clearAll();
  document.getElementById("formula-bar").style.display = "block";
  setFormulaBar(formula, null); // show formula, no highlight yet

  const table = buildTable(["Cell / ▶", "Formula / Function", "Value"]);
  const tbody = table.querySelector("tbody");

  // Row 1: the active cell, always highlighted blue
  appendRow(tbody, {
    col1: activeAddress,
    col2: formula || "(no formula)",
    col3: formatValue(value),
    depth: 0,
    isActive: true,
    path: "", // root row has no path
    token: { type: "cell_ref", address: activeAddress },
    formula: formula,
  });

  // Render each top-level token from the formula.
  // depth=0 keeps them visible; their children start at depth=1 and are hidden until expanded.
  let fnCounter = 0;
  for (const token of tokens) {
    fnCounter = renderToken(token, tbody, 0, "t" + fnCounter, formula, fnCounter);
  }

  document.getElementById("table-container").appendChild(table);
  attachKeyboardNav();
}

// Draws the Explore Dependents table.
// activeAddress — the source cell (or label for multi-cell selection)
// rows — [{ address, sheet, count }] from getDependentsRows()
export function renderDependents(activeAddress, rows) {
  clearAll();
  document.getElementById("formula-bar").style.display = "none";

  if (rows.length > 50) {
    const bar = document.getElementById("warning-bar");
    bar.textContent = "Large number of dependents detected. This may take a moment.";
    bar.style.display = "block";
  }

  const table = buildTable(["Cell address", "Sheet", "Count"]);
  const tbody = table.querySelector("tbody");

  // Row 1: the active cell/selection, blue
  const { sheet: activeSheet } = splitAddressLocal(activeAddress);
  appendRow(tbody, {
    col1: activeAddress,
    col2: activeSheet,
    col3: "",
    depth: 0,
    isActive: true,
    path: "",
    token: { type: "cell_ref", address: activeAddress },
    formula: null,
  });

  if (rows.length === 0) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="3" style="padding:10px;color:#888;font-style:italic;">No dependents found.</td>`;
    tbody.appendChild(tr);
  }

  for (const row of rows) {
    const { sheet } = splitAddressLocal(row.address);
    appendRow(tbody, {
      col1: row.address,
      col2: sheet,
      col3: String(row.count),
      depth: 0,
      isActive: false,
      path: "",
      token: { type: "cell_ref", address: row.address },
      formula: null,
    });
  }

  document.getElementById("table-container").appendChild(table);
  attachKeyboardNav();
}

// Shows a plain error message, replacing whatever is in the table area.
export function renderError(message) {
  clearAll();
  const div = document.createElement("div");
  div.style.cssText = "padding:16px;color:#cc0000;font-size:12px;";
  div.textContent = message;
  document.getElementById("table-container").appendChild(div);
}

// ------- internal helpers -------

// Renders one token as one or more rows. Returns the updated fnCounter.
function renderToken(token, tbody, depth, path, fullFormula, fnCounter) {
  if (token.type === "function") {
    // Function row: has expand button, shows function name and computed value
    const valueText = token.computedValue != null ? formatValue(token.computedValue) : "";
    appendRow(tbody, {
      col1: null, // will be replaced by expand button
      col2: token.name,
      col3: valueText,
      depth,
      isActive: false,
      isFn: true,
      path,
      token,
      formula: fullFormula,
      tooltipIfEmpty: token.computedValue == null
        ? "Value unavailable — function not supported by evaluation API"
        : null,
    });

    // Child rows (hidden by default, revealed when user clicks ▶)
    let childCounter = 0;
    for (const arg of token.args) {
      fnCounter = renderToken(arg, tbody, depth + 1, path + "." + childCounter, fullFormula, fnCounter);
      childCounter++;
    }
    fnCounter++;
  } else if (token.type === "cell_ref") {
    const cellValue = token.cellValue != null ? formatValue(token.cellValue) : "";
    const cellFormula = token.cellFormula && token.cellFormula.startsWith("=") ? token.cellFormula : "";
    appendRow(tbody, {
      col1: shortAddress(token.address),
      col2: cellFormula,
      col3: cellValue,
      depth,
      isActive: false,
      isFn: false,
      path,
      token,
      formula: fullFormula,
    });
  } else if (token.type === "range_ref") {
    appendRow(tbody, {
      col1: token.address,
      col2: "",
      col3: "",
      depth,
      isActive: false,
      isFn: false,
      path,
      token,
      formula: fullFormula,
    });
  } else {
    // expression: literal, comparison, named range, etc.
    const exprValue = resolveExpressionValue(token.raw);
    appendRow(tbody, {
      col1: "",
      col2: token.raw,
      col3: exprValue !== null ? formatValue(exprValue) : "",
      depth,
      isActive: false,
      isFn: false,
      path,
      token,
      formula: fullFormula,
    });
  }
  return fnCounter;
}

// Creates and appends a single table row. Wires up expand/collapse and navigation.
function appendRow(tbody, opts) {
  const { col1, col2, col3, depth, isActive, isFn, path, token, formula, tooltipIfEmpty } = opts;
  const tr = document.createElement("tr");
  tr.setAttribute("data-path", path);

  // All non-root rows are hidden initially if they have a parent path (i.e. depth > 0 and
  // they were added as children of a function row). Visibility is toggled by the parent's ▶ button.
  if (depth > 0) {
    tr.style.display = "none";
  }

  if (isActive) tr.classList.add("active-row");

  const indentPx = depth * 14;

  // ---- col1: address, expand button, or blank ----
  const td1 = document.createElement("td");
  td1.style.paddingLeft = indentPx + "px";

  if (isFn) {
    const btn = document.createElement("button");
    btn.className = "expand-btn";
    btn.textContent = "▶";
    btn.setAttribute("aria-label", "Expand " + col2);
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      toggleExpand(path, btn, tbody);
    });
    td1.appendChild(btn);
  } else {
    td1.textContent = col1 || "";
  }

  // ---- col2: formula/function name/expression ----
  const td2 = document.createElement("td");
  td2.textContent = col2 || "";

  // ---- col3: value ----
  const td3 = document.createElement("td");
  td3.textContent = col3 || "";
  if (tooltipIfEmpty && !col3) {
    td3.title = tooltipIfEmpty;
    td3.textContent = "—";
    td3.style.color = "#bbb";
  }

  tr.appendChild(td1);
  tr.appendChild(td2);
  tr.appendChild(td3);
  tbody.appendChild(tr);

  // Click the row itself: navigate to the cell (cell_ref and range_ref only), and highlight formula
  tr.addEventListener("click", () => {
    // Deselect all rows
    tbody.querySelectorAll("tr.selected-row").forEach((r) => r.classList.remove("selected-row"));
    tr.classList.add("selected-row");

    // Highlight the matching part of the formula bar
    if (formula && token && token.start != null) {
      setFormulaBar(formula, token);
    }

    // Navigate Excel to this cell (only for cell_ref tokens)
    if (token && token.type === "cell_ref") {
      navigateTo(token.address);
    }
  });
}

// Shows/hides all rows whose path starts with `fnPath + "."`.
// Also resets any nested expand buttons to ▶ when collapsing.
function toggleExpand(fnPath, btn, tbody) {
  const expanding = btn.textContent === "▶";
  btn.textContent = expanding ? "▼" : "▶";
  btn.setAttribute("aria-label", (expanding ? "Collapse " : "Expand ") + btn.closest("tr").querySelector("td:nth-child(2)").textContent);

  tbody.querySelectorAll("tr[data-path]").forEach((row) => {
    const rowPath = row.getAttribute("data-path");
    if (!rowPath.startsWith(fnPath + ".")) return;

    if (expanding) {
      // Only show direct children (one dot beyond the parent path)
      const childDepth = fnPath.split(".").length;
      const rowDepth = rowPath.split(".").length - 1;
      if (rowDepth === childDepth) {
        row.style.display = "";
      }
      // Deeper rows stay hidden until their parent is expanded
    } else {
      // Collapse all descendants, reset their expand buttons
      row.style.display = "none";
      const nestedBtn = row.querySelector(".expand-btn");
      if (nestedBtn) nestedBtn.textContent = "▶";
    }
  });
}

// Opens a new Excel.run() to navigate to the given cell address.
// Skips activate() when the target sheet is already active — on Mac, activating the
// current sheet cancels the subsequent select().
async function navigateTo(address) {
  try {
    await Excel.run(async (ctx) => {
      const { sheet, cell } = splitAddressLocal(address);
      const activeSheet = ctx.workbook.worksheets.getActiveWorksheet();
      activeSheet.load("name");
      await ctx.sync();
      if (activeSheet.name !== sheet) {
        ctx.workbook.worksheets.getItem(sheet).activate();
        await ctx.sync();
      }
      ctx.workbook.worksheets.getItem(sheet).getRange(cell).select();
      await ctx.sync();
    });
  } catch (e) {
    console.log("Navigation failed:", e.message);
  }
}

// Updates the formula bar text. If token is provided, wraps the matching substring in <mark>.
function setFormulaBar(formula, token) {
  const el = document.getElementById("formula-text");
  if (!formula) {
    el.innerHTML = "";
    return;
  }

  if (!token || token.start == null) {
    el.innerHTML = esc(formula);
    return;
  }

  // token.start/end are offsets into the body (after "="); add 1 for the leading "="
  const hlStart = token.start + 1;
  const hlEnd = token.end + 2; // +1 for "=", +1 to make the slice inclusive

  const before = formula.slice(0, hlStart);
  const highlighted = formula.slice(hlStart, hlEnd);
  const after = formula.slice(hlEnd);

  el.innerHTML = esc(before) + "<mark>" + esc(highlighted) + "</mark>" + esc(after);
}

// Builds a <table> element with column headers, colgroup for widths, and an empty <tbody>.
function buildTable(headers) {
  const table = document.createElement("table");
  table.className = "explorer-table";

  // colgroup sets proportional column widths (col1 widest since it shows addresses)
  const colgroup = document.createElement("colgroup");
  [34, 44, 22].forEach((pct) => {
    const col = document.createElement("col");
    col.style.width = pct + "%";
    colgroup.appendChild(col);
  });
  table.appendChild(colgroup);

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  headers.forEach((h) => {
    const th = document.createElement("th");
    th.textContent = h;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);
  table.appendChild(document.createElement("tbody"));
  return table;
}

// Clears the table container, warning bar, and formula bar.
function clearAll() {
  document.getElementById("table-container").innerHTML = "";
  const wb = document.getElementById("warning-bar");
  wb.style.display = "none";
  wb.textContent = "";
}

// Formats a cell value for display: numbers as-is, booleans as TRUE/FALSE, etc.
function formatValue(v) {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") return String(v);
  return String(v);
}

// For simple expression tokens that are plain literals, return the parsed value.
// Returns null for complex expressions that can't be reduced to a primitive.
function resolveExpressionValue(raw) {
  if (!raw) return null;
  const t = raw.trim();
  if (t === "TRUE") return true;
  if (t === "FALSE") return false;
  if (t.startsWith('"') && t.endsWith('"')) return t.slice(1, -1);
  const num = Number(t);
  if (!isNaN(num) && t !== "") return num;
  return null;
}

// Returns the cell-only part of an address for compact display.
// "Sheet1!B3" → "B3". "B3" → "B3".
function shortAddress(address) {
  if (!address) return "";
  const bang = address.lastIndexOf("!");
  return bang !== -1 ? address.slice(bang + 1) : address;
}

// Escapes HTML special characters so text can be inserted with innerHTML.
function esc(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Attaches arrow-key navigation to the table. Called once per page load.
// Up/Down move between visible rows and trigger the same action as clicking.
// Right expands a collapsed function row; Left collapses an expanded one (or
// jumps to the parent row and collapses it).
function attachKeyboardNav() {
  if (_navAttached) return;
  _navAttached = true;

  document.addEventListener("keydown", (e) => {
    if (!["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(e.key)) return;

    const tbody = document.querySelector("#table-container tbody");
    if (!tbody) return;

    const rows = Array.from(tbody.querySelectorAll("tr")).filter(
      (r) => r.style.display !== "none"
    );
    if (rows.length === 0) return;

    e.preventDefault();

    // Prefer an explicitly selected row over the always-present active-row (row 1).
    const currentRow =
      rows.find((r) => r.classList.contains("selected-row")) ||
      rows.find((r) => r.classList.contains("active-row"));
    const idx = currentRow ? rows.indexOf(currentRow) : -1;

    if (e.key === "ArrowDown") {
      const next = rows[Math.min(idx + 1, rows.length - 1)];
      next.click();
      next.scrollIntoView({ block: "nearest" });
    } else if (e.key === "ArrowUp") {
      const prev = rows[Math.max(idx - 1, 0)];
      prev.click();
      prev.scrollIntoView({ block: "nearest" });
    } else if (e.key === "ArrowRight") {
      const row = rows[idx];
      if (!row) return;
      const btn = row.querySelector(".expand-btn");
      if (btn && btn.textContent.trim() === "▶") btn.click();
    } else if (e.key === "ArrowLeft") {
      const row = rows[idx];
      if (!row) return;
      const btn = row.querySelector(".expand-btn");
      if (btn && btn.textContent.trim() === "▼") {
        btn.click();
      } else {
        // Child row: jump to parent function row and collapse it
        const path = row.getAttribute("data-path");
        if (path && path.includes(".")) {
          const parentPath = path.slice(0, path.lastIndexOf("."));
          const parentRow = tbody.querySelector(`tr[data-path="${parentPath}"]`);
          if (parentRow) {
            const parentBtn = parentRow.querySelector(".expand-btn");
            if (parentBtn && parentBtn.textContent.trim() === "▼") parentBtn.click();
            parentRow.click();
            parentRow.scrollIntoView({ block: "nearest" });
          }
        }
      }
    }
  });
}

// Local copy of splitAddress so renderer.js doesn't import graph.js (avoids circular deps).
function splitAddressLocal(address) {
  const bangIndex = address.lastIndexOf("!");
  if (bangIndex === -1) return { sheet: "Sheet1", cell: address };
  return {
    sheet: address.slice(0, bangIndex).replace(/^'|'$/g, ""),
    cell: address.slice(bangIndex + 1),
  };
}
