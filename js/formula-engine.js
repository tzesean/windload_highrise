/* Generic Excel formula parser + evaluator.
 * Operates directly on the extracted {sheet: {cellRef: formulaOrValue}} map
 * pulled from the real workbook -- no hand-retyped business logic.
 */

// ---------- Cell ref helpers ----------
function colToNum(col) {
  let n = 0;
  for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
function numToCol(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function parseCellRef(ref) {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref);
  if (!m) throw new Error('Bad cell ref: ' + ref);
  return { col: colToNum(m[1]), row: parseInt(m[2], 10) };
}
function cellKey(col, row) { return numToCol(col) + row; }

// ---------- Tokenizer ----------
function tokenize(formula) {
  const src = formula;
  const tokens = [];
  let i = 0;
  const n = src.length;
  const isDigit = c => c >= '0' && c <= '9';
  const isAlpha = c => /[A-Za-z_]/.test(c);
  while (i < n) {
    const c = src[i];
    if (c === ' ') { i++; continue; }
    if (c === "'") {
      // quoted sheet name
      let j = i + 1, s = '';
      while (j < n && src[j] !== "'") { s += src[j]; j++; }
      j++; // skip closing quote
      tokens.push({ t: 'SHEET', v: s });
      i = j;
      continue;
    }
    if (c === '"') {
      let j = i + 1, s = '';
      while (j < n && src[j] !== '"') { s += src[j]; j++; }
      j++;
      tokens.push({ t: 'STRING', v: s });
      i = j;
      continue;
    }
    if (isDigit(c) || (c === '.' && isDigit(src[i + 1]))) {
      let j = i, s = '';
      while (j < n && /[0-9.]/.test(src[j])) { s += src[j]; j++; }
      if (src[j] === 'E' || src[j] === 'e') {
        s += src[j]; j++;
        if (src[j] === '+' || src[j] === '-') { s += src[j]; j++; }
        while (j < n && isDigit(src[j])) { s += src[j]; j++; }
      }
      tokens.push({ t: 'NUMBER', v: parseFloat(s) });
      i = j;
      continue;
    }
    if (isAlpha(c) || c === '$') {
      let j = i, s = '';
      while (j < n && /[A-Za-z0-9_.$]/.test(src[j])) { s += src[j]; j++; }
      tokens.push({ t: 'NAME', v: s });
      i = j;
      continue;
    }
    if (c === '<' && src[i + 1] === '=') { tokens.push({ t: 'OP', v: '<=' }); i += 2; continue; }
    if (c === '>' && src[i + 1] === '=') { tokens.push({ t: 'OP', v: '>=' }); i += 2; continue; }
    if (c === '<' && src[i + 1] === '>') { tokens.push({ t: 'OP', v: '<>' }); i += 2; continue; }
    if ('+-*/^=<>&,(){}!:%'.includes(c)) { tokens.push({ t: 'OP', v: c }); i++; continue; }
    throw new Error('Unexpected char ' + c + ' in ' + formula);
  }
  return tokens;
}

// ---------- Parser: produces AST ----------
// Grammar (low to high precedence):
// expr := concat (('='|'<>'|'<'|'>'|'<='|'>=') concat)*
// concat := add ('&' add)*
// add := mul (('+'|'-') mul)*
// mul := unary (('*'|'/') unary)*
// unary := '-' unary | power
// power := postfix ('^' unary)*
// postfix := atom ('%')*
// atom := NUMBER | STRING | funccall | reference | '(' expr ')'
function parseFormula(tokens) {
  let pos = 0;
  const peek = () => tokens[pos];
  const next = () => tokens[pos++];

  function parseExpr() {
    let left = parseConcat();
    while (peek() && peek().t === 'OP' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) {
      const op = next().v;
      const right = parseConcat();
      left = { type: 'binop', op, left, right };
    }
    return left;
  }
  function parseConcat() {
    let left = parseAdd();
    while (peek() && peek().t === 'OP' && peek().v === '&') {
      next();
      const right = parseAdd();
      left = { type: 'binop', op: '&', left, right };
    }
    return left;
  }
  function parseAdd() {
    let left = parseMul();
    while (peek() && peek().t === 'OP' && (peek().v === '+' || peek().v === '-')) {
      const op = next().v;
      const right = parseMul();
      left = { type: 'binop', op, left, right };
    }
    return left;
  }
  function parseMul() {
    let left = parseUnary();
    while (peek() && peek().t === 'OP' && (peek().v === '*' || peek().v === '/')) {
      const op = next().v;
      const right = parseUnary();
      left = { type: 'binop', op, left, right };
    }
    return left;
  }
  function parseUnary() {
    if (peek() && peek().t === 'OP' && peek().v === '-') {
      next();
      return { type: 'neg', expr: parseUnary() };
    }
    return parsePower();
  }
  function parsePower() {
    let left = parsePostfix();
    if (peek() && peek().t === 'OP' && peek().v === '^') {
      next();
      const right = parseUnary();
      left = { type: 'binop', op: '^', left, right };
    }
    return left;
  }
  function parsePostfix() {
    let node = parseAtom();
    while (peek() && peek().t === 'OP' && peek().v === '%') {
      next();
      node = { type: 'pct', expr: node };
    }
    return node;
  }
  function parseAtom() {
    const tok = peek();
    if (!tok) throw new Error('Unexpected end of formula');
    if (tok.t === 'NUMBER') { next(); return { type: 'num', v: tok.v }; }
    if (tok.t === 'STRING') { next(); return { type: 'str', v: tok.v }; }
    if (tok.t === 'OP' && tok.v === '(') {
      next();
      const e = parseExpr();
      if (!(peek() && peek().v === ')')) throw new Error('Expected )');
      next();
      return e;
    }
    if (tok.t === 'SHEET') {
      next();
      if (!(peek() && peek().t === 'OP' && peek().v === '!')) throw new Error('Expected ! after sheet name');
      next();
      return parseRefOrCall(tok.v);
    }
    if (tok.t === 'NAME') {
      next();
      // could be Sheet!ref (unquoted), func call, named range, or bare cell ref
      if (peek() && peek().t === 'OP' && peek().v === '!') {
        next();
        return parseRefOrCall(tok.v);
      }
      return parseRefOrCall(null, tok.v);
    }
    throw new Error('Unexpected token ' + JSON.stringify(tok));
  }
  function parseRefOrCall(sheet, name) {
    if (name === undefined) {
      // sheet given, next token is the ref part (NAME)
      const tok = next();
      name = tok.v;
    }
    if (peek() && peek().t === 'OP' && peek().v === '(') {
      // function call
      next();
      const args = [];
      if (!(peek() && peek().v === ')')) {
        args.push(parseExpr());
        while (peek() && peek().v === ',') { next(); args.push(parseExpr()); }
      }
      if (!(peek() && peek().v === ')')) throw new Error('Expected ) closing ' + name);
      next();
      return { type: 'call', name: name.toUpperCase(), args, sheet };
    }
    // reference (cell or range) or named range
    if (peek() && peek().t === 'OP' && peek().v === ':') {
      next();
      const tok2 = next();
      let name2 = tok2.v;
      return { type: 'range', sheet, from: name, to: name2 };
    }
    return { type: 'ref', sheet, name };
  }

  const ast = parseExpr();
  return ast;
}

function parse(formula) {
  const f = formula.startsWith('=') ? formula.slice(1) : formula;
  return parseFormula(tokenize(f));
}

export { parse, colToNum, numToCol, parseCellRef, cellKey };
