/* global console */

// Evaluates an Excel function using the workbook.functions API built into Office.js.
// This lets us show the computed value of a function row (like SUM or AVERAGE)
// without writing anything to the workbook.
//
// How to use:
//   const results = buildAllFunctionResults(tokens, defaultSheet, context);
//   results.forEach(r => r.fnResult.load("value"));
//   await context.sync();
//   results.forEach(r => { r.token.computedValue = r.fnResult.value; });
//
// Note: IF, AND, OR, NOT are reserved words in JavaScript.
// The Office JS API names them if_, and_, or_, not_ (with underscores).

import { splitAddress, normalizeAddress } from "./graph.js";

// Maps Excel function names to their context.workbook.functions method names.
const FN_MAP = {
  SUM: "sum",
  AVERAGE: "average",
  COUNT: "count",
  COUNTA: "countA",
  MIN: "min",
  MAX: "max",
  ABS: "abs",
  INT: "int",
  ROUND: "round",
  ROUNDUP: "roundUp",
  ROUNDDOWN: "roundDown",
  IF: "if_",
  AND: "and_",
  OR: "or_",
  NOT: "not_",
  IFERROR: "ifError",
  ISNUMBER: "isNumber",
  ISTEXT: "isText",
  ISBLANK: "isBlank",
  SUMIF: "sumIf",
  COUNTIF: "countIf",
  VLOOKUP: "vlookup",
  HLOOKUP: "hlookup",
  INDEX: "index",
  MATCH: "match",
  LEN: "len",
  LEFT: "left",
  RIGHT: "right",
  MID: "mid",
  TRIM: "trim",
  UPPER: "upper",
  LOWER: "lower",
  CONCATENATE: "concatenate",
};

// Walks the token tree and returns one { token, fnResult } object per function token.
// fnResult is a FunctionResult object — call result.load("value") then context.sync()
// to read its computed value. Returns an empty array if workbook.functions is unavailable.
export function buildAllFunctionResults(tokens, defaultSheet, context) {
  const results = [];
  collectFunctionResults(tokens, defaultSheet, context, results);
  return results;
}

function collectFunctionResults(tokens, defaultSheet, context, results) {
  for (const token of tokens) {
    if (token.type === "function") {
      const fnResult = buildFunctionResult(token.name, token.args, defaultSheet, context);
      if (fnResult) {
        results.push({ token, fnResult });
      }
      // Recurse into nested function args regardless of whether this fn could be evaluated
      collectFunctionResults(token.args, defaultSheet, context, results);
    }
  }
}

// Builds a FunctionResult for one function token. Returns null if the function
// is unsupported or if any argument can't be resolved (e.g. a complex expression like A1>50).
// Does NOT call context.sync() — the caller must do that.
function buildFunctionResult(fnName, argTokens, defaultSheet, context) {
  const apiName = FN_MAP[fnName.toUpperCase()];
  if (!apiName) return null;

  const resolvedArgs = resolveArgs(argTokens, defaultSheet, context);
  if (resolvedArgs === null) return null;

  try {
    return context.workbook.functions[apiName](...resolvedArgs);
  } catch (e) {
    console.log("buildFunctionResult failed for", fnName, e.message);
    return null;
  }
}

// Converts an array of tokens into arguments the API accepts: Range objects, FunctionResult
// objects (for nested functions), or primitives (numbers, strings, booleans).
// Returns null if any token can't be resolved.
function resolveArgs(argTokens, defaultSheet, context) {
  const resolved = [];
  for (const token of argTokens) {
    const val = resolveArg(token, defaultSheet, context);
    if (val === null) return null;
    resolved.push(val);
  }
  return resolved;
}

// Converts a single token to a value the workbook.functions API can accept.
// Returns null if the token can't be resolved.
function resolveArg(token, defaultSheet, context) {
  // Cell or range reference → Excel.Range object
  if (token.type === "cell_ref" || token.type === "range_ref") {
    const full = normalizeAddress(token.address, defaultSheet);
    const { sheet, cell } = splitAddress(full);
    return context.workbook.worksheets.getItem(sheet).getRange(cell);
  }

  // Nested function → build its FunctionResult (can be passed directly to an outer function)
  if (token.type === "function") {
    return buildFunctionResult(token.name, token.args, defaultSheet, context);
  }

  // Expression: try to parse as a plain primitive value
  if (token.type === "expression") {
    const raw = token.raw.trim();
    if (raw === "TRUE") return true;
    if (raw === "FALSE") return false;
    if (raw.startsWith('"') && raw.endsWith('"')) return raw.slice(1, -1); // quoted string
    const num = Number(raw);
    if (!isNaN(num) && raw !== "") return num; // number or integer
    return null; // complex expression like "A1>50" — can't resolve
  }

  return null;
}
