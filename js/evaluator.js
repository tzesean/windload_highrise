import { parse, colToNum, numToCol, parseCellRef } from './formula-engine.js';

const EXCEL_ERROR = Symbol('error');

class Workbook {
  constructor(bundle) {
    this.sheets = bundle.sheets; // {sheetName: {cellRef: raw}}
    this.names = bundle.names;   // {name: {sheet, ref}}
    this.overrides = {};         // key `${sheet}!${ref}` -> value (mimics macro paste)
    this.cache = {};             // memoized computed values, cleared per run
    this.astCache = {};          // parsed formula cache (persists across runs)
    this.stack = new Set();
  }

  reset() { this.cache = {}; }

  setOverride(sheet, ref, value) {
    this.overrides[`${sheet}!${ref}`] = value;
  }
  clearOverrides() { this.overrides = {}; }

  resolveName(name) {
    // named range -> {sheet, ref}
    const nm = this.names[name];
    if (nm) return nm;
    return null;
  }

  getRaw(sheet, ref) {
    const ov = this.overrides[`${sheet}!${ref}`];
    if (ov !== undefined) return ov;
    const s = this.sheets[sheet];
    if (!s) throw new Error('Unknown sheet ' + sheet);
    if (!(ref in s)) return null; // blank cell
    return s[ref];
  }

  getCellValue(sheet, ref) {
    const key = `${sheet}!${ref}`;
    if (key in this.cache) {
      const v = this.cache[key];
      if (v === EXCEL_ERROR) return { error: true, message: 'circular/err at ' + key };
      return v;
    }
    if (this.stack.has(key)) {
      // circular reference
      return { error: true, message: 'circular ' + key };
    }
    this.stack.add(key);
    let raw = this.getRaw(sheet, ref);
    let result;
    if (typeof raw === 'string' && raw.startsWith('=')) {
      const astKey = raw; // formulas identical text share AST
      let ast = this.astCache[astKey];
      if (!ast) { ast = parse(raw); this.astCache[astKey] = ast; }
      try {
        result = evalNode(ast, this, sheet);
      } catch (e) {
        result = { error: true, message: e.message };
      }
    } else {
      result = raw;
    }
    this.stack.delete(key);
    this.cache[key] = result;
    return result;
  }

  // Expand a range into a flat list of {sheet, ref} cells, row-major
  expandRange(sheet, fromRef, toRef) {
    const a = parseCellRef(fromRef);
    const b = parseCellRef(toRef);
    const c1 = Math.min(a.col, b.col), c2 = Math.max(a.col, b.col);
    const r1 = Math.min(a.row, b.row), r2 = Math.max(a.row, b.row);
    const cells = [];
    for (let r = r1; r <= r2; r++) {
      const row = [];
      for (let c = c1; c <= c2; c++) {
        row.push({ sheet, ref: numToCol(c) + r });
      }
      cells.push(row);
    }
    return cells; // 2D array of {sheet, ref}
  }
}

function isErr(v) { return v && typeof v === 'object' && v.error; }
function toNum(v) {
  if (isErr(v)) throw new Error(v.message);
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string') {
    const n = parseFloat(v);
    if (!isNaN(n)) return n;
    throw new Error('Cannot coerce "' + v + '" to number');
  }
  throw new Error('Cannot coerce to number: ' + JSON.stringify(v));
}
function toStr(v) {
  if (isErr(v)) throw new Error(v.message);
  if (v === null || v === undefined) return '';
  return String(v);
}
function truthy(v) {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (typeof v === 'string') return v.toUpperCase() === 'TRUE';
  return !!v;
}

function resolveRefTarget(node, wb, currentSheet) {
  // node: {type:'ref', sheet, name} -- name could be a cell ref or a defined name
  let sheet = node.sheet || currentSheet;
  let ref = node.name;
  if (!/^\$?[A-Z]+\$?\d+$/.test(ref)) {
    const nm = wb.resolveName(ref);
    if (nm) return { sheet: nm.sheet, ref: nm.ref, isRange: nm.ref.includes(':') };
    throw new Error('Unknown name ' + ref);
  }
  return { sheet, ref: ref.replace(/\$/g, ''), isRange: false };
}

function flattenGrid(grid) {
  const out = [];
  for (const row of grid) for (const cell of row) out.push(cell);
  return out;
}

function getRangeGrid(node, wb, currentSheet) {
  if (node.type === 'range') {
    const sheet = node.sheet || currentSheet;
    return wb.expandRange(sheet, node.from.replace(/\$/g, ''), node.to.replace(/\$/g, ''));
  }
  if (node.type === 'ref') {
    const target = resolveRefTarget(node, wb, currentSheet);
    if (target.isRange) {
      const [f, t] = target.ref.split(':');
      return wb.expandRange(target.sheet, f, t);
    }
    return [[{ sheet: target.sheet, ref: target.ref }]];
  }
  throw new Error('Not a range node: ' + node.type);
}

function evalNode(node, wb, sheet) {
  switch (node.type) {
    case 'num': return node.v;
    case 'str': return node.v;
    case 'neg': return -toNum(evalNode(node.expr, wb, sheet));
    case 'pct': return toNum(evalNode(node.expr, wb, sheet)) / 100;
    case 'binop': return evalBinop(node, wb, sheet);
    case 'ref': {
      const target = resolveRefTarget(node, wb, sheet);
      if (target.isRange) {
        // used where a single-cell context expected but name is a range -> return grid
        return getRangeGrid(node, wb, sheet);
      }
      return wb.getCellValue(target.sheet, target.ref);
    }
    case 'range': return getRangeGrid(node, wb, sheet);
    case 'call': return evalCall(node, wb, sheet);
    default: throw new Error('Unknown node type ' + node.type);
  }
}

function evalBinop(node, wb, sheet) {
  const op = node.op;
  if (op === '&') return toStr(evalNode(node.left, wb, sheet)) + toStr(evalNode(node.right, wb, sheet));
  const l = evalNode(node.left, wb, sheet);
  const r = evalNode(node.right, wb, sheet);
  if (['=', '<>'].includes(op)) {
    let eq;
    if (typeof l === 'string' || typeof r === 'string') eq = toStr(l).toUpperCase() === toStr(r).toUpperCase();
    else eq = toNum(l) === toNum(r);
    return op === '=' ? eq : !eq;
  }
  if (['<', '>', '<=', '>='].includes(op)) {
    const ln = toNum(l), rn = toNum(r);
    switch (op) {
      case '<': return ln < rn;
      case '>': return ln > rn;
      case '<=': return ln <= rn;
      case '>=': return ln >= rn;
    }
  }
  const ln = toNum(l), rn = toNum(r);
  switch (op) {
    case '+': return ln + rn;
    case '-': return ln - rn;
    case '*': return ln * rn;
    case '/': if (rn === 0) throw new Error('#DIV/0!'); return ln / rn;
    case '^': return Math.pow(ln, rn);
  }
  throw new Error('Unknown op ' + op);
}

function flatValues(node, wb, sheet) {
  // For SUM/AVERAGE/MAX/COUNTA: accept ranges or scalar args, return flat array of raw values
  if (node.type === 'range' || (node.type === 'ref' && resolveRefTarget(node, wb, sheet).isRange)) {
    const grid = getRangeGrid(node, wb, sheet);
    return flattenGrid(grid).map(c => wb.getCellValue(c.sheet, c.ref));
  }
  return [evalNode(node, wb, sheet)];
}

function evalCall(node, wb, sheet) {
  const name = node.name;
  const A = node.args;
  switch (name) {
    case 'IF': {
      const cond = truthy(evalNode(A[0], wb, sheet));
      if (cond) return A[1] !== undefined ? evalNode(A[1], wb, sheet) : true;
      return A[2] !== undefined ? evalNode(A[2], wb, sheet) : false;
    }
    case 'AND': return A.every(a => truthy(evalNode(a, wb, sheet)));
    case 'OR': return A.some(a => truthy(evalNode(a, wb, sheet)));
    case 'ABS': return Math.abs(toNum(evalNode(A[0], wb, sheet)));
    case 'LN': return Math.log(toNum(evalNode(A[0], wb, sheet)));
    case 'ROW': {
      // ROW(ref) -> row number of ref
      const ref = A[0];
      const target = resolveRefTarget(ref, wb, sheet);
      return parseCellRef(target.ref).row;
    }
    case 'RIGHT': {
      const s = toStr(evalNode(A[0], wb, sheet));
      const n = A[1] ? toNum(evalNode(A[1], wb, sheet)) : 1;
      return s.slice(Math.max(0, s.length - n));
    }
    case 'CEILING': {
      const x = toNum(evalNode(A[0], wb, sheet));
      const sig = toNum(evalNode(A[1], wb, sheet));
      if (sig === 0) return 0;
      return Math.ceil(x / sig - 1e-9) * sig;
    }
    case 'SUM': {
      let vals = [];
      for (const a of A) vals = vals.concat(flatValues(a, wb, sheet));
      return vals.reduce((s, v) => s + (typeof v === 'number' ? v : toNum(v) || 0), 0);
    }
    case 'AVERAGE': {
      let vals = [];
      for (const a of A) vals = vals.concat(flatValues(a, wb, sheet));
      const nums = vals.filter(v => typeof v === 'number' || (typeof v === 'string' && v !== '' && !isNaN(parseFloat(v))));
      const ns = nums.map(toNum);
      return ns.reduce((s, v) => s + v, 0) / ns.length;
    }
    case 'MAX': {
      let vals = [];
      for (const a of A) vals = vals.concat(flatValues(a, wb, sheet));
      const ns = vals.filter(v => typeof v === 'number' || (typeof v === 'string' && v !== '' && !isNaN(parseFloat(v)))).map(toNum);
      return Math.max(...ns);
    }
    case 'MIN': {
      let vals = [];
      for (const a of A) vals = vals.concat(flatValues(a, wb, sheet));
      const ns = vals.filter(v => typeof v === 'number' || (typeof v === 'string' && v !== '' && !isNaN(parseFloat(v)))).map(toNum);
      return Math.min(...ns);
    }
    case 'COUNTA': {
      let vals = [];
      for (const a of A) vals = vals.concat(flatValues(a, wb, sheet));
      return vals.filter(v => v !== null && v !== undefined && v !== '').length;
    }
    case 'OFFSET': {
      const ref = A[0];
      const target = resolveRefTarget(ref, wb, sheet);
      const dRow = toNum(evalNode(A[1], wb, sheet));
      const dCol = A[2] ? toNum(evalNode(A[2], wb, sheet)) : 0;
      const base = parseCellRef(target.ref);
      const newRef = numToCol(base.col + dCol) + (base.row + dRow);
      return wb.getCellValue(target.sheet, newRef);
    }
    case 'INDEX': {
      const grid = getRangeGrid(A[0], wb, sheet);
      const rowNum = A[1] ? toNum(evalNode(A[1], wb, sheet)) : 0;
      const colNum = A[2] ? toNum(evalNode(A[2], wb, sheet)) : 0;
      let cell;
      if (grid.length === 1 && rowNum >= 1 && !A[2]) {
        // 1-row range indexed by "row" arg acting as column (Excel INDEX semantics when omitted col)
        cell = grid[0][rowNum - 1];
      } else if (grid[0].length === 1 && colNum === 0) {
        cell = grid[rowNum - 1][0];
      } else {
        const r = rowNum >= 1 ? rowNum - 1 : 0;
        const c = colNum >= 1 ? colNum - 1 : 0;
        cell = grid[r][c];
      }
      if (!cell) return { error: true, message: '#REF! INDEX out of range' };
      return wb.getCellValue(cell.sheet, cell.ref);
    }
    case 'MATCH': {
      const lookup = evalNode(A[0], wb, sheet);
      const grid = getRangeGrid(A[1], wb, sheet);
      const flat = flattenGrid(grid).map(c => wb.getCellValue(c.sheet, c.ref));
      const matchType = A[2] !== undefined ? toNum(evalNode(A[2], wb, sheet)) : 1;
      if (matchType === 0) {
        for (let i = 0; i < flat.length; i++) {
          const v = flat[i];
          const eq = (typeof lookup === 'string' || typeof v === 'string')
            ? toStr(v).toUpperCase() === toStr(lookup).toUpperCase()
            : toNum(v) === toNum(lookup);
          if (eq) return i + 1;
        }
        return { error: true, message: '#N/A MATCH exact' };
      } else if (matchType === 1) {
        // largest value <= lookup, assumes ascending order (numeric or text)
        let best = -1;
        const isText = typeof lookup === 'string' && isNaN(parseFloat(lookup));
        for (let i = 0; i < flat.length; i++) {
          const le = isText
            ? toStr(flat[i]).toUpperCase() <= toStr(lookup).toUpperCase()
            : toNum(flat[i]) <= toNum(lookup);
          if (le) best = i + 1; else break;
        }
        if (best === -1) return { error: true, message: '#N/A MATCH asc' };
        return best;
      } else {
        // -1: smallest value >= lookup, assumes descending order (numeric or text)
        let best = -1;
        const isText = typeof lookup === 'string' && isNaN(parseFloat(lookup));
        for (let i = 0; i < flat.length; i++) {
          const ge = isText
            ? toStr(flat[i]).toUpperCase() >= toStr(lookup).toUpperCase()
            : toNum(flat[i]) >= toNum(lookup);
          if (ge) best = i + 1; else break;
        }
        if (best === -1) return { error: true, message: '#N/A MATCH desc' };
        return best;
      }
    }
    case 'VLOOKUP': {
      const lookup = evalNode(A[0], wb, sheet);
      const grid = getRangeGrid(A[1], wb, sheet);
      const colIndex = toNum(evalNode(A[2], wb, sheet));
      const approx = A[3] !== undefined ? truthy(evalNode(A[3], wb, sheet)) : true;
      const firstCol = grid.map(row => wb.getCellValue(row[0].sheet, row[0].ref));
      let rowIdx = -1;
      if (!approx) {
        for (let i = 0; i < firstCol.length; i++) {
          const v = firstCol[i];
          const eq = (typeof lookup === 'string' || typeof v === 'string')
            ? toStr(v).toUpperCase() === toStr(lookup).toUpperCase()
            : toNum(v) === toNum(lookup);
          if (eq) { rowIdx = i; break; }
        }
      } else {
        const ln = toNum(lookup);
        for (let i = 0; i < firstCol.length; i++) {
          if (toNum(firstCol[i]) <= ln) rowIdx = i; else break;
        }
      }
      if (rowIdx === -1) return { error: true, message: '#N/A VLOOKUP' };
      const cell = grid[rowIdx][colIndex - 1];
      return wb.getCellValue(cell.sheet, cell.ref);
    }
    case 'INTERPOLATE': {
      // Ported 1:1 from the workbook's VBA UDF
      const x1 = toNum(evalNode(A[0], wb, sheet));
      const y1 = toNum(evalNode(A[1], wb, sheet));
      const x2 = toNum(evalNode(A[2], wb, sheet));
      const y2 = toNum(evalNode(A[3], wb, sheet));
      const inputXY = toNum(evalNode(A[4], wb, sheet));
      const type = toStr(evalNode(A[5], wb, sheet));
      if (type === 'y') {
        return y1 - ((y1 - y2) * (x1 - inputXY) / (x1 - x2));
      } else {
        return x1 - ((y1 - inputXY) * (x1 - x2) / (y1 - y2));
      }
    }
    default:
      throw new Error('Unsupported function ' + name);
  }
}

export { Workbook, evalNode };
