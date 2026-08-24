/**
 * Local, dependency-backed shim for @oai/artifact-tool (codex runtime).
 * Copyright © 2026 David YE
 * SPDX-License-Identifier: MIT
 *
 * Reproduces the subset of the FileBlob/SpreadsheetFile API used by
 * TA-Analysis pipeline scripts (analyze_source.mjs, build_enriched_workbook.mjs,
 * verify_enriched_workbook.mjs) on top of exceljs, so the pipeline runs outside
 * the codex sandbox. Supported surface:
 *   FileBlob.load(path)
 *   SpreadsheetFile.importXlsx(blob) / exportXlsx(workbook)
 *   workbook.worksheets.getItem/add, workbook.inspect, workbook.render*
 *   sheet.name, sheet.showGridLines, sheet.getUsedRange(true).values,
 *   sheet.getRange(addr) -> {values, formulasR1C1, fillDown, merge, format},
 *   sheet.tables.add(ref, hasHeaders, name), sheet.freezePanes.freezeRows/Columns
 *   range.format.{fill,font,wrapText,verticalAlignment,horizontalAlignment,
 *                 borders,columnWidth,rowHeight,numberFormat}
 *   render() is NOT supported (no rasterizer outside the codex runtime).
 * Formula support: R1C1 templates with relative RC[k] refs, translated to A1,
 * evaluated and cached with a small interpreter for IF/AND/ISNUMBER/IFERROR/
 * VALUE/LEFT/TRIM/FIND/SUBSTITUTE and arithmetic/comparison operators.
 */
import fs from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";

const ERR = Symbol("formula-error");

function makeError(code) {
  const e = new Error(code);
  e[ERR] = true;
  e.code = code;
  return e;
}

const COLUMN_CACHE = new Map();
function columnName(index) {
  let cached = COLUMN_CACHE.get(index);
  if (cached) return cached;
  let value = index;
  let out = "";
  while (value > 0) {
    value -= 1;
    out = String.fromCharCode(65 + (value % 26)) + out;
    value = Math.floor(value / 26);
  }
  COLUMN_CACHE.set(index, out);
  return out;
}

function argb(color) {
  const hex = String(color || "").replace("#", "").toUpperCase();
  if (/^[0-9A-F]{6}$/.test(hex)) return `FF${hex}`;
  if (/^[0-9A-F]{8}$/.test(hex)) return hex;
  return undefined;
}

// ---------- address helpers ----------
const ADDR_RE = /^\$?([A-Za-z]+)\$?(\d+)?(?::\$?([A-Za-z]+)\$?(\d+)?)?$/;
export function parseAddress(address) {
  const match = ADDR_RE.exec(String(address).trim());
  if (!match) throw new Error(`Unsupported range address: ${address}`);
  const colFrom = columnIndex(match[1].toUpperCase());
  const rowFrom = match[2] ? Number(match[2]) : 1;
  const colTo = match[3] ? columnIndex(match[3].toUpperCase()) : colFrom;
  const rowTo = match[4] ? Number(match[4]) : match[3] ? (match[4] === undefined && match[2] ? rowFrom : rowFrom) : rowFrom;
  // "A:B" style (columns only) and "1:3" handled by caller flags below
  return { colFrom, rowFrom, colTo, rowTo, fullColumn: !match[2] && !!match[1], fullRow: false };
}

export function columnIndex(letters) {
  let value = 0;
  for (const char of letters) value = value * 26 + (char.charCodeAt(0) - 64);
  return value;
}

const ROWCOL_RE = /^(\$?\d+):(\$?\d+)$/;
function parseAddressEx(address) {
  const rowCol = ROWCOL_RE.exec(String(address).trim());
  if (rowCol) {
    const rowFrom = Number(rowCol[1].replace("$", ""));
    const rowTo = Number(rowCol[2].replace("$", ""));
    return { kind: "rows", rowFrom, rowTo };
  }
  const base = parseAddress(address);
  if (base.fullColumn) return { kind: "columns", colFrom: base.colFrom, colTo: base.colTo };
  return { kind: "range", colFrom: base.colFrom, rowFrom: base.rowFrom, colTo: base.colTo, rowTo: base.rowTo };
}

// ---------- formula evaluation ----------
function tokenize(input) {
  const tokens = [];
  let i = 0;
  while (i < input.length) {
    const char = input[i];
    if (char === " ") { i += 1; continue; }
    if (char === '"') {
      let j = i + 1;
      let text = "";
      while (j < input.length) {
        if (input[j] === '"' && input[j + 1] === '"') { text += '"'; j += 2; continue; }
        if (input[j] === '"') break;
        text += input[j];
        j += 1;
      }
      tokens.push({ type: "string", value: text });
      i = j + 1;
      continue;
    }
    if (/[0-9]/.test(char) || (char === "." && /[0-9]/.test(input[i + 1] || ""))) {
      let j = i;
      while (j < input.length && /[0-9.]/.test(input[j])) j += 1;
      tokens.push({ type: "number", value: Number(input.slice(i, j)) });
      i = j;
      continue;
    }
    const two = input.slice(i, i + 2);
    if (["<>", ">=", "<="].includes(two)) { tokens.push({ type: "op", value: two }); i += 2; continue; }
    if ("+-*/()<>=,&".includes(char)) { tokens.push({ type: "op", value: char }); i += 1; continue; }
    if (/^[Rr]/.test(char)) {
      // R1C1 relative refs: RC[k] (column offset), R[k]C, R[k]C[m], bare RC/R nC m
      const bracketed = /^[Rr]\[(-?\d+)\][Cc](?:\[(-?\d+)\])?/.exec(input.slice(i));
      const plain = !bracketed ? /^[Rr][Cc](?:\[(-?\d+)\])?/.exec(input.slice(i)) : null;
      if (bracketed || plain) {
        let deltaRow = 0;
        let deltaCol = 0;
        if (bracketed) {
          deltaRow = Number(bracketed[1]);
          if (bracketed[2] !== undefined) deltaCol = Number(bracketed[2]);
        } else {
          if (plain[1] !== undefined) deltaCol = Number(plain[1]);
        }
        tokens.push({ type: "relref", deltaRow, deltaCol });
        i += bracketed ? bracketed[0].length : plain[0].length;
        continue;
      }
    }
    if (/[A-Za-z$]/.test(char)) {
      let j = i;
      while (j < input.length && /[A-Za-z0-9$_.]/.test(input[j])) j += 1;
      const word = input.slice(i, j);
      if (/^\$?[A-Za-z]{1,3}\$?\d+$/.test(word)) tokens.push({ type: "absref", value: word });
      else tokens.push({ type: "ident", value: word.toUpperCase() });
      i = j;
      continue;
    }
    throw makeError("#NAME?");
  }
  return tokens;
}

function parse(tokens) {
  let pos = 0;
  const peek = () => tokens[pos];
  const eat = (type, value) => {
    const token = tokens[pos];
    if (!token || token.type !== type || (value !== undefined && token.value !== value)) {
      const found = token ? `${token.type}:${token.type === "string" ? JSON.stringify(token.value) : String(token.value)}` : "EOF";
      throw contextualError(`#NAME? eat expected ${type}:${value ?? ""} at ${pos}, found ${found}`);
    }
    pos += 1;
    return token;
  };
  const parseExpr = () => parseComparison();
  const parseComparison = () => {
    let left = parseAdditive();
    while (peek()?.type === "op" && ["=", "<>", ">=", "<=", ">", "<"].includes(peek().value)) {
      const op = tokens[pos++].value;
      const right = parseAdditive();
      left = { kind: "binop", op, left, right };
    }
    return left;
  };
  const parseAdditive = () => {
    let left = parseMultiplicative();
    while (peek()?.type === "op" && ["+", "-", "&"].includes(peek().value)) {
      const op = tokens[pos++].value;
      const right = parseMultiplicative();
      left = { kind: "binop", op, left, right };
    }
    return left;
  };
  const parseMultiplicative = () => {
    let left = parseUnary();
    while (peek()?.type === "op" && ["*", "/"].includes(peek().value)) {
      const op = tokens[pos++].value;
      const right = parseUnary();
      left = { kind: "binop", op, left, right };
    }
    return left;
  };
  const parseUnary = () => {
    if (peek()?.type === "op" && peek().value === "-") { pos += 1; return { kind: "negate", operand: parseUnary() }; }
    return parsePrimary();
  };
  const parsePrimary = () => {
    const token = peek();
    if (!token) throw makeError("#NAME?");
    if (token.type === "number") { pos += 1; return { kind: "literal", value: token.value }; }
    if (token.type === "string") { pos += 1; return { kind: "literal", value: token.value }; }
    if (token.type === "relref") { pos += 1; return { kind: "relref", deltaRow: token.deltaRow, deltaCol: token.deltaCol }; }
    if (token.type === "absref") { pos += 1; const ref = parseAddress(token.value.replace(/\$/g, "") + (/\d/.test(token.value.slice(-1)) ? "" : "1")); return { kind: "absref", row: ref.rowFrom, col: ref.colFrom }; }
    if (token.type === "op" && token.value === "(") { pos += 1; const inner = parseExpr(); eat("op", ")"); return inner; }
    if (token.type === "ident") {
      pos += 1;
      eat("op", "(");
      const args = [];
      if (!(peek()?.type === "op" && peek().value === ")")) {
        args.push(parseExpr());
        while (peek()?.type === "op" && peek().value === ",") { pos += 1; args.push(parseExpr()); }
      }
      eat("op", ")");
      return { kind: "call", name: token.value, args };
    }
    const detail = token.type === "string" ? JSON.stringify(token.value) : String(token.value);
    throw contextualError(`#NAME? at token ${pos} (${token.type}:${detail})`);
  };
  const ast = parseExpr();
  if (pos !== tokens.length) {
    const rest = tokens.slice(pos, pos + 4).map((t) => `${t.type}:${String(t.value).slice(0, 24)}`).join(" | ");
    throw contextualError(`#NAME? trailing ${tokens.length - pos} tokens: ${rest}`);
  }
  return ast;
}

function contextualError(message) {
  const error = makeError("#NAME?");
  error.message = message;
  return error;
}

function coerceNumber(value) {
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (value === null || value === undefined || value === "") return 0;
  const text = String(value).trim().replace(/,/g, "");
  if (/^-?\d+(\.\d+)?$/.test(text)) return Number(text);
  throw makeError("#VALUE!");
}

const FUNCS = {
  IF: (args, ctx) => {
    const condition = ctx.eval(args[0]);
    const truthy = typeof condition === "boolean" ? condition : !(condition === 0 || condition === "" || condition === null || condition === undefined);
    return truthy ? ctx.eval(args[1]) : (args[2] !== undefined ? ctx.eval(args[2]) : false);
  },
  AND: (args, ctx) => args.every((arg) => { const v = ctx.eval(arg); return typeof v === "boolean" ? v : !(v === 0 || v === "" || v === null || v === undefined); }),
  ISNUMBER: (args, ctx) => typeof ctx.evalLoose(args[0]) === "number",
  IFERROR: (args, ctx) => { try { const v = ctx.eval(args[0]); return v; } catch (error) { if (error[ERR]) return ctx.eval(args[1]); throw error; } },
  VALUE: (args, ctx) => coerceNumber(ctx.eval(args[0])),
  LEFT: (args, ctx) => {
    const text = stringify(ctx.eval(args[0]));
    const count = args.length > 1 ? coerceNumber(ctx.eval(args[1])) : 1;
    return text.slice(0, Math.max(0, Math.floor(count)));
  },
  TRIM: (args, ctx) => stringify(ctx.eval(args[0])).replace(/\s+/g, " ").trim(),
  FIND: (args, ctx) => {
    const needle = stringify(ctx.eval(args[0]));
    const haystack = stringify(ctx.eval(args[1]));
    const index = haystack.indexOf(needle) + 1;
    if (index === 0) throw makeError("#VALUE!");
    return index;
  },
  SUBSTITUTE: (args, ctx) => {
    const text = stringify(ctx.eval(args[0]));
    const oldText = stringify(ctx.eval(args[1]));
    const newText = stringify(ctx.eval(args[2]));
    if (!oldText) return text;
    return text.split(oldText).join(newText);
  },
};

function stringify(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

function evaluateAst(ast, ctx) {
  switch (ast.kind) {
    case "literal": return ast.value;
    case "negate": return -coerceNumber(evaluateAst(ast.operand, ctx));
    case "relref": return ctx.cell(ctx.row + ast.deltaRow, ctx.col + ast.deltaCol);
    case "absref": return ctx.cell(ast.row, ast.col);
    case "binop": {
      const left = ctx.eval(ast.left);
      const right = ctx.eval(ast.right);
      switch (ast.op) {
        case "+": return coerceNumber(left) + coerceNumber(right);
        case "-": return coerceNumber(left) - coerceNumber(right);
        case "&": return stringify(left) + stringify(right);
        case "*": return coerceNumber(left) * coerceNumber(right);
        case "/": { const d = coerceNumber(right); if (d === 0) throw makeError("#DIV/0!"); return coerceNumber(left) / d; }
        case "=": return equals(left, right);
        case "<>": return !equals(left, right);
        case ">": return compare(left, right) > 0;
        case "<": return compare(left, right) < 0;
        case ">=": return compare(left, right) >= 0;
        case "<=": return compare(left, right) <= 0;
      }
      throw makeError("#NAME?");
    }
    case "call": {
      const func = FUNCS[ast.name];
      if (!func) throw makeError(`#NAME? ${ast.name}`);
      return func(ast.args, ctx);
    }
  }
  throw makeError("#NAME?");
}

function equals(left, right) {
  if (typeof left === "number" && typeof right === "number") return left === right;
  return stringify(left).toLowerCase() === stringify(right).toLowerCase();
}

function compare(left, right) {
  if (typeof left === "number" && typeof right === "number") return left < right ? -1 : left > right ? 1 : 0;
  const a = stringify(left);
  const b = stringify(right);
  return a < b ? -1 : a > b ? 1 : 0;
}

// ---------- public classes ----------
class BlobImpl {
  constructor(bytes, filePath) {
    this.bytes = bytes;
    this.filePath = filePath;
  }
  async arrayBuffer() {
    return this.bytes.buffer.slice(this.bytes.byteOffset, this.bytes.byteOffset + this.bytes.byteLength);
  }
  async save(filePath) {
    await fs.writeFile(filePath, this.bytes);
  }
}

class RangeImpl {
  constructor(sheet, address) {
    this._sheet = sheet;
    this._address = address;
    const parsed = parseAddressEx(address);
    this._parsed = parsed;
    if (parsed.kind === "range") {
      this._top = parsed.rowFrom;
      this._left = parsed.colFrom;
      this._bottom = parsed.rowTo;
      this._right = parsed.colTo;
    } else {
      this._top = 1;
      this._left = 1;
      this._bottom = Infinity;
      this._right = Infinity;
    }
  }

  _cells() {
    const { kind } = this._parsed;
    if (kind === "range") {
      const cells = [];
      for (let row = this._top; row <= this._bottom; row += 1)
        for (let col = this._left; col <= this._right; col += 1)
          cells.push(this._sheet._ws.getRow(row).getCell(col));
      return cells;
    }
    if (kind === "columns")
      return this._sheet._ws.columns.slice(this._parsed.colFrom - 1, this._parsed.colTo).filter(Boolean).flatMap((c) => c.cells ?? []);
    return [];
  }

  get values() {
    const { kind } = this._parsed;
    if (kind !== "range") throw new Error(".values read requires a bounded range");
    const rows = [];
    for (let row = this._top; row <= this._bottom; row += 1) {
      const out = [];
      for (let col = this._left; col <= this._right; col += 1) {
        const raw = this._sheet._readCell(row, col);
        out.push(raw);
      }
      rows.push(out);
    }
    return rows;
  }

  set values(matrix) {
    const { kind } = this._parsed;
    if (kind !== "range") throw new Error(".values write requires a bounded range");
    matrix.forEach((rowValues, rowIndex) => {
      rowValues.forEach((value, colIndex) => {
        const row = this._top + rowIndex;
        const col = this._left + colIndex;
        this._sheet._writeValue(row, col, value);
      });
    });
  }

  set formulasR1C1(matrix) {
    const template = matrix[0][0];
    this._sheet._lastFormula = { template, row: this._top, col: this._left };
    this._applyFormulaRow(template, this._top, this._left);
  }

  _applyFormulaRow(template, row, leftCol) {
    const ast = this._sheet._compileFormula(template);
    const formulaA1 = translateR1C1ToA1(template, row, leftCol);
    const sheet = this._sheet;
    const ctx = {
      row,
      col: leftCol,
      cell: (r, c) => sheet._rawCellValue(r, c),
    };
    ctx.eval = (node) => evaluateAst(node, ctx);
    ctx.evalLoose = (node) => {
      try { return evaluateAst(node, ctx); } catch (error) { if (error[ERR]) return error.code; throw error; }
    };
    let result = null;
    try {
      const evaluated = evaluateAst(ast, ctx);
      if (typeof evaluated === "boolean") result = evaluated ? "TRUE" : "FALSE";
      else if (typeof evaluated === "number" && Number.isFinite(evaluated)) result = evaluated;
      else if (typeof evaluated === "string") result = evaluated;
      else result = null;
    } catch (error) {
      result = error[ERR] ? error.code : "";
      if (!error[ERR]) {
        console.error(`[shim-eval] unexpected ${row},${leftCol}:`, error.message);
        if (process.env.SHIM_DEBUG) console.error(error.stack.split("\n").slice(0, 6).join("\n"));
      }
    }
    if (process.env.SHIM_DEBUG && row <= 6) {
      console.error(`[shim-eval] r${row} c${leftCol} formula=${formulaA1.slice(0, 160)} ->`, JSON.stringify(result));
    }
    sheet._ws.getRow(row).getCell(leftCol).value = { formula: formulaA1, result };
  }

  async fillDown() {
    const last = this._sheet._lastFormula;
    if (!last || last.col !== this._left || last.row !== this._top) throw new Error("fillDown(): no matching formulasR1C1 anchor");
    const bottom = this._bottom === Infinity ? this._sheet._usedRowCount() : this._bottom;
    for (let row = last.row + 1; row <= bottom; row += 1) this._applyFormulaRow(last.template, row, last.col);
  }

  merge(across) {
    const { kind } = this._parsed;
    if (kind !== "range") return;
    if (across === true) {
      for (let row = this._top; row <= this._bottom; row += 1) {
        try { this._sheet._ws.mergeCells(row, this._left, row, this._right); } catch { /* already merged */ }
      }
    } else {
      try { this._sheet._ws.mergeCells(this._top, this._left, this._bottom, this._right); } catch { /* already merged */ }
    }
  }

  get format() {
    return this._formatProxy ||= createFormatProxy(() => this._parsed, (fn) => fn(this._sheet, this._parsed, this._bounds()));
  }

  set format(spec) {
    applyFormat(this._sheet, this._parsed, this._bounds(), spec);
  }

  _bounds() {
    const { kind } = this._parsed;
    if (kind === "range") return { top: this._top, left: this._left, bottom: this._bottom, right: this._right };
    return null;
  }
}

function createFormatProxy(getParsed, apply) {
  const target = {};
  return new Proxy(target, {
    set(_t, key, value) {
      if (key === "borders" || key === "numberFormat" || key === "fill" || key === "font" || key === "wrapText" ||
          key === "verticalAlignment" || key === "horizontalAlignment" || key === "columnWidth" || key === "rowHeight") {
        apply((sheet, parsed, bounds) => applyFormat(sheet, parsed, bounds, { [key]: value }));
        return true;
      }
      target[key] = value;
      return true;
    },
    get(_t, key) { return target[key]; },
  });
}

function applyFormat(sheet, parsed, bounds, spec) {
  const ws = sheet._ws;
  if (spec.columnWidth !== undefined) {
    const from = parsed.kind === "columns" ? parsed.colFrom : bounds?.left;
    const to = parsed.kind === "columns" ? parsed.colTo : bounds?.right;
    if (from && to) for (let col = from; col <= to; col += 1) ws.getColumn(col).width = spec.columnWidth;
    return;
  }
  if (spec.rowHeight !== undefined) {
    const from = parsed.kind === "rows" ? parsed.rowFrom : bounds?.top;
    const to = parsed.kind === "rows" ? parsed.rowTo : bounds?.bottom;
    if (from && to) for (let row = from; row <= to; row += 1) ws.getRow(row).height = spec.rowHeight;
    return;
  }
  if (!bounds) return;
  const { top, left, bottom, right } = bounds;
  for (let row = top; row <= bottom; row += 1) {
    for (let col = left; col <= right; col += 1) {
      const cell = ws.getRow(row).getCell(col);
      if (spec.fill !== undefined) {
        const fillArgb = argb(spec.fill);
        if (fillArgb) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fillArgb } };
      }
      if (spec.font !== undefined) {
        const font = {};
        if (spec.font.bold !== undefined) font.bold = spec.font.bold;
        if (spec.font.italic !== undefined) font.italic = spec.font.italic;
        if (spec.font.size !== undefined) font.size = spec.font.size;
        const color = argb(spec.font.color);
        if (color) font.color = { argb: color };
        if (Object.keys(font).length) cell.font = { ...cell.font, ...font };
      }
      if (spec.wrapText !== undefined || spec.verticalAlignment !== undefined || spec.horizontalAlignment !== undefined) {
        cell.alignment = {
          ...(cell.alignment || {}),
          ...(spec.wrapText !== undefined ? { wrapText: !!spec.wrapText } : {}),
          ...(spec.verticalAlignment !== undefined ? { vertical: spec.verticalAlignment } : {}),
          ...(spec.horizontalAlignment !== undefined ? { horizontal: spec.horizontalAlignment } : {}),
        };
      }
      if (spec.borders !== undefined && spec.borders.preset === "all") {
        const borderStyle = { style: spec.borders.style || "thin" };
        const color = argb(spec.borders.color);
        if (color) borderStyle.color = { argb: color };
        cell.border = { top: borderStyle, bottom: borderStyle, left: borderStyle, right: borderStyle };
      }
      if (spec.numberFormat !== undefined) cell.numFmt = spec.numberFormat;
    }
  }
}

function translateR1C1ToA1(template, row, col) {
  return template.replace(/RC\[(-?\d+)\]/g, (_m, dc) => `${columnName(col + Number(dc))}${row}`)
    .replace(/R\[-?(\d+)\]C/g, (_m, dr) => `A${row + Number(dr)}`);
}

class FreezePanesImpl {
  constructor(sheet) { this._sheet = sheet; }
  _view() {
    const ws = this._sheet._ws;
    if (!ws.views?.length) ws.views = [{}];
    const view = ws.views[0];
    if (view.state !== "frozen") { view.state = "frozen"; view.xSplit = view.xSplit || 0; view.ySplit = view.ySplit || 0; }
    return view;
  }
  freezeRows(count) { this._view().ySplit = count; }
  freezeColumns(count) { this._view().xSplit = count; }
}

export class ShimWorksheet {
  constructor(workbook, ws) {
    this._workbook = workbook;
    this._ws = ws;
    this.freezePanes = new FreezePanesImpl(this);
    this.tables = {
      add: (reference, hasHeaders, name) => this._addTable(reference, hasHeaders, name),
    };
    this._lastFormula = null;
  }

  get name() { return this._ws.name; }
  set showGridLines(value) {
    if (!this._ws.views?.length) this._ws.views = [{}];
    this._ws.views[0].showGridLines = !!value;
  }

  getRange(address) { return new RangeImpl(this, address); }

  getUsedRange(valuesOnly) {
    const shim = this;
    const rowCount = shim._usedRowCount();
    const colCount = shim._usedColumnCount();
    return {
      get values() {
        const rows = [];
        for (let row = 1; row <= rowCount; row += 1) {
          const out = [];
          for (let col = 1; col <= colCount; col += 1) out.push(shim._readCell(row, col));
          rows.push(out);
        }
        return rows;
      },
    };
  }

  _addTable(reference, hasHeaders, name) {
    const parsed = parseAddressEx(reference);
    const headerRow = parsed.rowFrom;
    const columns = [];
    for (let col = parsed.colFrom; col <= parsed.colTo; col += 1) {
      const label = hasHeaders ? this._readCell(headerRow, col) : `列${col}`;
      columns.push({ name: stringify(label) || `列${col}`, filterButton: false });
    }
    this._workbook._registryTables.push({ name, sheet: this._ws.name, ref: reference, columns: columns.length });
    try {
      this._ws.addTable({
        name,
        displayName: name,
        reference,
        headerRow: !!hasHeaders,
        totalsRow: false,
        style: { theme: "TableStyleMedium2", showRowStripes: true },
        columns: columns.map((c) => ({ name: c.name, filterButton: false })),
        rows: [],
      });
    } catch {
      /* exceljs table constraints (duplicate names etc.) – registration kept for inspect */
    }
    return { name };
  }

  _usedRowCount() {
    let max = 0;
    this._ws.eachRow({ includeEmpty: false }, (row, number) => {
      if (number > max) max = number;
    });
    return max;
  }

  _usedColumnCount() {
    let max = 0;
    this._ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (colNumber > max) max = colNumber;
      });
    });
    return Math.max(max, 1);
  }

  _rawCellValue(row, col) {
    if (process.env.SHIM_DEBUG && (typeof col !== "number" || typeof row !== "number")) {
      console.error(`[shim] bad cell args row=${JSON.stringify(row)} col=${JSON.stringify(col)}`);
      console.error(new Error().stack.split("\n").slice(1, 8).join("\n"));
    }
    const cell = this._ws.getRow(row)?.getCell?.(col);
    if (!cell) return null;
    const value = cell.value;
    if (value && typeof value === "object" && !Array.isArray(value) && !(value instanceof Date)) {
      if ("result" in value) return value.result ?? null;
      if ("text" in value) return value.text;
      if ("richText" in value) return value.richText.map((part) => part.text).join("");
      return null;
    }
    if (value instanceof Date) return value;
    return value ?? null;
  }

  _writeValue(row, col, value) {
    const cell = this._ws.getRow(row).getCell(col);
    if (value === undefined) cell.value = null;
    else cell.value = value;
  }

  _compileFormulaCache() { return this._formulaCache ||= new Map(); }
  _compileFormula(template) {
    const cache = this._compileFormulaCache();
    let ast = cache.get(template);
    if (!ast) { ast = parse(tokenize(template.startsWith("=") ? template.slice(1) : template)); cache.set(template, ast); }
    return ast;
  }

  _readCell(row, col) {
    const raw = this._rawCellValue(row, col);
    if (raw instanceof Date) return raw.toISOString().slice(0, 10);
    return raw === undefined ? null : raw;
  }
}

export class SpreadsheetWorkbook {
  constructor(excelWorkbook) {
    this._wb = excelWorkbook;
    this._registryTables = [];
  }
  get worksheets() {
    const workbook = this;
    return {
      getItem: (name) => {
        const ws = workbook._wb.getWorksheet(name);
        if (!ws) throw new Error(`Worksheet not found: ${name}`);
        return new ShimWorksheet(workbook, ws);
      },
      add: (name) => new ShimWorksheet(workbook, workbook._wb.addWorksheet(name)),
    };
  }
  async inspect() {
    const lines = [];
    lines.push(JSON.stringify({ type: "workbook", sheets: this._wb.worksheets.length, tables: this._registryTables.length }));
    for (const ws of this._wb.worksheets) {
      const shim = new ShimWorksheet(this, ws);
      lines.push(JSON.stringify({ type: "sheet", name: ws.name, rows: shim._usedRowCount(), cols: shim._usedColumnCount() }));
    }
    for (const table of this._registryTables) lines.push(JSON.stringify({ type: "table", ...table }));
    return { ndjson: lines.join("\n") };
  }
  async render() {
    throw new Error("workbook.render() is not supported by the local artifact-tool shim (no rasterizer outside codex runtime)");
  }
}

export const FileBlob = {
  async load(filePath) {
    const bytes = await fs.readFile(filePath);
    return new BlobImpl(bytes, path.resolve(String(filePath)));
  },
};

export const SpreadsheetFile = {
  async importXlsx(blobLike) {
    const workbook = new ExcelJS.Workbook();
    const bytes = blobLike?.bytes ?? Buffer.from(await blobLike.arrayBuffer());
    await workbook.xlsx.load(bytes);
    return new SpreadsheetWorkbook(workbook);
  },
  async exportXlsx(workbook) {
    if (!(workbook instanceof SpreadsheetWorkbook)) throw new Error("exportXlsx expects a workbook from importXlsx");
    const buffer = await workbook._wb.xlsx.writeBuffer();
    return new BlobImpl(Buffer.from(buffer));
  },
};
