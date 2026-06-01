/* global console */

// Parses an Excel formula string into a tree of typed tokens.
// Each token represents one component: a function call, a cell reference,
// a range reference, or a raw expression (number, string, operator, comparison).
//
// Token types:
//   { type:"function",    name:"SUM",   start, end, args:[...] }
//   { type:"cell_ref",   address:"B1", start, end }
//   { type:"range_ref",  address:"B1:C5", start, end }
//   { type:"expression", raw:"A1>50",  start, end }
//
// start/end are character offsets into the formula body (the part after the leading "=").
// They are used to highlight the matching portion in the formula bar when a row is clicked.

export function parseFormula(formula) {
  if (!formula || !formula.startsWith("=")) return [];
  const body = formula.slice(1); // strip the leading "="
  return parseArgList(body, 0, body.length);
}

// Parses a comma-separated list of tokens within str[start..end).
function parseArgList(str, start, end) {
  const tokens = [];
  let pos = start;

  while (pos < end) {
    while (pos < end && str[pos] === " ") pos++; // skip leading whitespace
    if (pos >= end) break;

    const result = parseToken(str, pos, end);
    if (result.token) {
      if (result.token.type === "expression") {
        // Try to pull out individual cell/range refs from arithmetic expressions
        // like "A1+B1" or "A1*2". Each ref becomes its own row in the UI.
        splitExpression(result.token.raw, result.token.start).forEach((t) => tokens.push(t));
      } else {
        tokens.push(result.token);
      }
    }
    pos = result.nextPos;

    // Skip trailing whitespace and the comma separator
    while (pos < end && str[pos] === " ") pos++;
    if (pos < end && str[pos] === ",") pos++;
    while (pos < end && str[pos] === " ") pos++;
  }

  return tokens;
}

// Parses a single token starting at str[pos] within str[pos..end).
function parseToken(str, pos, end) {
  // Check for a function call: LETTERS( — e.g. SUM(, INDEX(, IF(
  const fnMatch = str.slice(pos).match(/^([A-Z][A-Z0-9._]*)\(/);
  if (fnMatch) {
    const name = fnMatch[1];
    const openParen = pos + name.length;
    const closeParen = findMatchingParen(str, openParen);
    const args = parseArgList(str, openParen + 1, closeParen);
    return {
      token: { type: "function", name, start: pos, end: closeParen, args },
      nextPos: closeParen + 1,
    };
  }

  // For everything else: find where this argument ends, then classify the whole chunk.
  const argEnd = findArgEnd(str, pos, end);
  const raw = str.slice(pos, argEnd).trim();

  if (!raw) return { token: null, nextPos: argEnd };

  // Range reference: "A1:B5", "Sheet2!A1:B5", "'My Sheet'!A1:B5"
  if (/^(?:(?:'[^']+'|[A-Za-z0-9_]+)!)?(\$?[A-Z]{1,3}\$?[0-9]{1,7}:\$?[A-Z]{1,3}\$?[0-9]{1,7})$/.test(raw)) {
    return {
      token: { type: "range_ref", address: raw, start: pos, end: argEnd - 1 },
      nextPos: argEnd,
    };
  }

  // Single cell reference: "B1", "$B$1", "Sheet2!C3", "'My Sheet'!D4"
  if (/^(?:(?:'[^']+'|[A-Za-z0-9_]+)!)?(\$?[A-Z]{1,3}\$?[0-9]{1,7})$/.test(raw)) {
    return {
      token: { type: "cell_ref", address: raw, start: pos, end: argEnd - 1 },
      nextPos: argEnd,
    };
  }

  // Everything else: numbers, quoted strings, comparisons like "A1>50", named ranges, etc.
  return {
    token: { type: "expression", raw, start: pos, end: argEnd - 1 },
    nextPos: argEnd,
  };
}

// Returns the position of the closing ")" that matches the "(" at openPos.
function findMatchingParen(str, openPos) {
  let depth = 0;
  let inString = false;

  for (let i = openPos; i < str.length; i++) {
    if (str[i] === '"' && !inString) { inString = true; continue; }
    if (str[i] === '"' && inString) { inString = false; continue; }
    if (inString) continue;
    if (str[i] === "(") depth++;
    if (str[i] === ")") {
      depth--;
      if (depth === 0) return i;
    }
  }
  return str.length - 1; // fallback: treat end of string as close
}

// Splits an expression like "A1+B1" or "A1>50" into individual cell/range ref tokens
// plus any non-ref expression parts. Segments that are only operator characters
// (+, -, *, /, ^, &, <, >, =, !) are dropped — they add no useful information as rows.
// startOffset is the character position of raw[0] within the formula body (used for
// formula-bar highlighting). If no cell refs are found, returns the original as-is.
function splitExpression(raw, startOffset) {
  // Matches a cell or range ref, optionally prefixed by a sheet name
  const cellPat = /(?:(?:'[^']+'|[A-Za-z0-9_]+)!)?\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?/g;
  const parts = [];
  let lastEnd = 0;
  let match;

  while ((match = cellPat.exec(raw)) !== null) {
    const mStart = match.index;
    const mEnd = mStart + match[0].length;

    // Any text between the previous match and this one (usually an operator)
    if (mStart > lastEnd) {
      const between = raw.slice(lastEnd, mStart).trim();
      // Drop segments that are only operator/comparison characters
      if (between && !/^[+\-*/^&<>=!%\s]+$/.test(between)) {
        parts.push({ type: "expression", raw: between, start: startOffset + lastEnd, end: startOffset + mStart - 1 });
      }
    }

    const addr = match[0];
    if (addr.includes(":")) {
      parts.push({ type: "range_ref", address: addr, start: startOffset + mStart, end: startOffset + mEnd - 1 });
    } else {
      parts.push({ type: "cell_ref", address: addr, start: startOffset + mStart, end: startOffset + mEnd - 1 });
    }

    lastEnd = mEnd;
  }

  // Any trailing text after the last match (e.g. ">50" or "*2")
  if (lastEnd < raw.length) {
    const trailing = raw.slice(lastEnd).trim();
    if (trailing && !/^[+\-*/^&<>=!%\s]+$/.test(trailing)) {
      parts.push({ type: "expression", raw: trailing, start: startOffset + lastEnd, end: startOffset + raw.length - 1 });
    }
  }

  return parts.length > 0 ? parts : [{ type: "expression", raw, start: startOffset, end: startOffset + raw.length - 1 }];
}

// Returns the index where the current argument ends:
// either the next top-level comma or closing ")" within str[pos..end).
function findArgEnd(str, pos, end) {
  let depth = 0;
  let inString = false;

  for (let i = pos; i < end; i++) {
    if (str[i] === '"' && !inString) { inString = true; continue; }
    if (str[i] === '"' && inString) { inString = false; continue; }
    if (inString) continue;
    if (str[i] === "(") depth++;
    if (str[i] === ")") {
      if (depth === 0) return i;
      depth--;
    }
    if (str[i] === "," && depth === 0) return i;
  }
  return end;
}
