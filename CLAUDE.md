# CLAUDE.md

Behavioural guidelines for this project. Read before every task.

## Project context

Building an Excel add-in for Mac that traces formula dependencies 
(precedents and dependents) and renders them as a visual graph in 
a task pane. The developer is non-technical — clarity and 
simplicity take priority over elegance or performance.

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
- Do not assume Windows behaviour applies on Mac.

## 6. Scope boundary

This add-in does one thing: formula dependency tracing. Do not suggest 
or add features outside this scope (error auditing, formula complexity 
scoring, etc.) unless explicitly asked.