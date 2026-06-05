# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project context

Building an Excel add-in for Mac that traces formula dependencies
(precedents and dependents) and renders them as a list/table in a task pane.
The developer is non-technical — clarity and simplicity take priority over
elegance or performance.

---

## Commands

```bash
npm start          # build dev bundle, sideload into Excel, open dev server
npm stop           # remove the sideloaded add-in from Excel
npm run build      # production build → dist/
npm run build:dev  # development build → dist/
npm run watch      # rebuild on file change (no dev server)
npm run lint       # lint with office-addin-lint
```

There are no tests. Verify changes by running `npm start` and exercising the
feature directly in Excel.

The dev server runs on `https://localhost:3000`. SSL certs are managed by
`office-addin-dev-certs` and generated automatically on first run.

---

## Architecture

The add-in is plain JavaScript with no framework. Webpack bundles two entry
points from `src/`: `taskpane` (the visible panel) and `commands` (keyboard
shortcut handlers). Both run in a **shared runtime** so they can communicate
via `window.*`.

### Data flow for Precedents

```
taskpane.js → graph.js (getPrecedentCell)
           → formula-parser.js (parseFormula → token tree)
           → fn-evaluator.js (buildAllFunctionResults → computed values via workbook.functions)
           → renderer.js (renderPrecedents → DOM table)
```

### Data flow for Dependents

```
taskpane.js → graph.js (getDependentsRows → tries getDirectDependents() API,
                         always also runs scanForDependents() text scan as fallback)
           → fallback-parser.js (parseFormula → cell refs for counting)
           → renderer.js (renderDependents → DOM table)
```

### Two parsers — different jobs

| File | Input | Output | Used for |
|---|---|---|---|
| `formula-parser.js` | `=SUM(A1,B1)` | Rich token tree with `start`/`end` offsets and nested `args` | Precedents UI — rows, formula-bar highlights, function expansion |
| `fallback-parser.js` | `=SUM(A1,B1)` | Flat array of cell/range ref strings | Scanning dependent formulas to count how many times they reference a source cell |

### Key files

- **`taskpane.js`** — Office.js entry point; reads selection, calls graph + parser + evaluator, calls renderer. Also registers keyboard shortcut actions via `Office.actions.associate` and exposes `window.showPrecedents` / `window.showDependents` for the shared runtime.
- **`graph.js`** — all Excel API calls for fetching cell data. Exports address utilities (`splitAddress`, `normalizeAddress`, `expandRange`) used across files.
- **`renderer.js`** — pure DOM manipulation, no Excel API calls. Has a local copy of `splitAddress` to avoid a circular import with `graph.js`.
- **`commands.js`** — keyboard shortcut handlers. If the task pane is already open, calls `window.showPrecedents/Dependents`; otherwise stores the mode in `window._pendingMode` and `taskpane.js` reads it on init.
- **`fn-evaluator.js`** — maps Excel function names (e.g. `IF`) to their `workbook.functions` API equivalents (e.g. `if_`). Skipped silently if the API is unavailable.

### Shared runtime communication

`commands.js` and `taskpane.js` share a JS context. `commands.js` calls
`window.showPrecedents()` directly if the task pane is open. If the pane
hasn't loaded yet, it writes `window._pendingMode` and `taskpane.js`
reads it in `Office.onReady`.

---

## 1. Think Before Coding

Before writing anything:

- State your assumptions explicitly. If uncertain, ask.
- If multiple approaches exist, present them briefly — don't pick silently.
- If a simpler approach exists, say so.
- If the request is ambiguous, stop and ask. Don't guess and proceed.

## 2. Simplicity First

- Write the minimum code that solves the problem. Nothing more.
- No extra features, abstractions, or "future-proofing" unless asked.
- No error handling for scenarios that can't realistically happen.
- Add a short plain-English comment above every function explaining
  what it does. Assume the reader has no coding background.
- If a solution is 100 lines and it could be 40, rewrite it.

## 3. Surgical Changes

- Only modify the file explicitly mentioned in the request.
- Do not "improve" adjacent code, comments, or formatting.
- Do not refactor working code unless asked.
- Match the existing style and structure of the file.
- If you notice something unrelated that looks wrong, mention it —
  don't fix it without asking.
- Every changed line should trace directly to what was requested.

## 4. Goal-Driven Execution

For any task, state a brief plan before coding:

1. [Step] → verify: [how to check it worked]
2. [Step] → verify: [how to check it worked]

For bug fixes: describe what's causing the bug before touching code.
For new functions: confirm the expected inputs and outputs before writing.

## 5. Mac compatibility

- `getDirectPrecedents()` and `getDirectDependents()` may not be
  available on older Excel for Mac versions.
- Always check for API support before calling these. If unsupported,
  fall back to the regex-based formula parser in `fallback-parser.js`.
- `getDirectDependents()` is also unreliable for hardcoded-value cells
  (cells with no formula) even when the API is present — the text-scan
  fallback in `graph.js` is the authoritative source for those.
- Calling `worksheet.activate()` on the already-active sheet cancels
  any subsequent `range.select()` in the same sync batch. Always check
  the active sheet name and skip `activate()` if already on that sheet.
- Do not assume Windows behaviour applies on Mac.

## 6. Scope boundary

This add-in does one thing: formula dependency tracing. Do not suggest
or add features outside this scope (error auditing, formula complexity
scoring, etc.) unless explicitly asked.
