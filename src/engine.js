/** Gridline calculation and document core. No DOM, network calls, eval, or dependencies. */
export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;
export const MAX_RANGE_CELLS = 200000;
const clone = value => value === undefined ? undefined : structuredClone(value);
export const keyOf = (r, c) => `${r},${c}`;
export function colName(c) {
  let s = ''; for (c++; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + (c - 1) % 26) + s;
  return s;
}
export function address(r, c) { return `${colName(c)}${r + 1}`; }
export function parseAddress(s) {
  const m = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/.exec(s);
  if (!m) return null;
  let c = 0; for (const x of m[1].toUpperCase()) c = c * 26 + x.charCodeAt(0) - 64;
  const r = Number(m[2]) - 1;
  return r < MAX_ROWS && c <= MAX_COLS ? { r, c: c - 1 } : null;
}
export function normalizedRange(a, b = a) {
  return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
}
export function parseRange(s) {
  const [a, b = a] = s.trim().split(':'); const x = parseAddress(a), y = parseAddress(b);
  return x && y ? normalizedRange(x, y) : null;
}
export function rangeAddress(q) { return q.r1 === q.r2 && q.c1 === q.c2 ? address(q.r1, q.c1) : `${address(q.r1, q.c1)}:${address(q.r2, q.c2)}`; }
export function* cellsIn(q, limit = MAX_RANGE_CELLS) {
  if ((q.r2 - q.r1 + 1) * (q.c2 - q.c1 + 1) > limit) throw new Error(`Selection exceeds the ${limit.toLocaleString()}-cell operation limit.`);
  for (let r = q.r1; r <= q.r2; r++) for (let c = q.c1; c <= q.c2; c++) yield { r, c };
}
export class FormulaError extends Error {
  constructor(code, detail = '') { super(detail || code); this.code = code; this.name = 'FormulaError'; }
  toString() { return this.code; }
}
const fail = (code, detail) => { throw new FormulaError(code, detail); };
const checkError = x => { if (x instanceof FormulaError) throw x; return x; };
export class RangeValue {
  constructor(rows, reference = null) { this.rows = rows; this.reference = reference; }
  flat() { return this.rows.flat(); }
}
const scalar = x => checkError(x instanceof RangeValue ? x.rows[0]?.[0] ?? null : x);
const number = v => {
  v = scalar(v);
  if (v === null || v === '' || v === undefined) return 0;
  if (typeof v === 'number') return Number.isFinite(v) ? v : fail('#NUM!');
  if (typeof v === 'boolean') return +v;
  const n = Number(v); return Number.isFinite(n) ? n : fail('#VALUE!');
};
const str = v => { v = scalar(v); return v == null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v); };
const logical = v => { v = scalar(v); if (typeof v === 'string' && !/^(true|false)$/i.test(v)) fail('#VALUE!'); return typeof v === 'string' ? /^true$/i.test(v) : !!v; };
const flatten = args => args.flatMap(x => x instanceof RangeValue ? x.flat() : [x]);
const numeric = args => {
  const out = [];
  for (const arg of args) {
    if (arg instanceof RangeValue) { for (const x of arg.flat()) { checkError(x); if (typeof x === 'number') out.push(x); } }
    else if (arg != null && arg !== '') out.push(number(arg));
  }
  return out;
};
const sum = xs => xs.reduce((a, b) => a + b, 0);
const safeResult = v => typeof v === 'number' && !Number.isFinite(v) ? fail('#NUM!') : v;
const errorRegex = /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|CYCLE!|SPILL!|ERROR!)/i;
const cellPattern = /^\$?([A-Z]{1,3})(\$?)([1-9]\d*)$/i;
export function tokenizeFormula(source) {
  const tokens = []; let i = source.startsWith('=') ? 1 : 0;
  while (i < source.length) {
    const start = i, ch = source[i];
    if (/\s/.test(ch)) { i++; continue; }
    if (ch === '"' || ch === "'") {
      const quote = ch; let s = ''; i++;
      let closed = false;
      while (i < source.length) {
        if (source[i] === quote) { if (source[i + 1] === quote) { s += quote; i += 2; } else { i++; closed = true; break; } }
        else s += source[i++];
      }
      if (!closed) fail('#ERROR!', 'Unterminated quoted text.');
      tokens.push({ type: quote === '"' ? 'string' : 'sheet', value: s, start, end: i }); continue;
    }
    const rest = source.slice(i);
    let m;
    if ((m = errorRegex.exec(rest))) { i += m[0].length; tokens.push({ type: 'error', value: m[0].toUpperCase(), start, end: i }); continue; }
    if ((m = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(rest))) { i += m[0].length; tokens.push({ type: 'number', value: Number(m[0]), start, end: i }); continue; }
    if ((m = /^\$?[A-Za-z_][A-Za-z0-9_.$]*/.exec(rest))) { i += m[0].length; tokens.push({ type: 'id', value: m[0], start, end: i }); continue; }
    if ((m = /^(<=|>=|<>|[+\-*/^&%=<>,;():!])/.exec(rest))) { i += m[0].length; tokens.push({ type: 'op', value: m[0] === ';' ? ',' : m[0], start, end: i }); continue; }
    fail('#ERROR!', `Unexpected character at ${i + 1}: ${ch}`);
  }
  tokens.push({ type: 'eof', value: '', start: i, end: i }); return tokens;
}
const precedence = { '=': 10, '<>': 10, '<': 10, '>': 10, '<=': 10, '>=': 10, '&': 20, '+': 30, '-': 30, '*': 40, '/': 40, '^': 50 };
export class FormulaParser {
  constructor(source) { this.tokens = tokenizeFormula(source); this.i = 0; this.depth = 0; }
  peek(value) { return value === undefined ? this.tokens[this.i] : this.tokens[this.i].value === value; }
  take() { return this.tokens[this.i++]; }
  expect(value) { if (!this.peek(value)) fail('#ERROR!', `Expected ${value}.`); this.take(); }
  parse() { const node = this.expression(0); if (this.peek().type !== 'eof') fail('#ERROR!', 'Unexpected trailing expression.'); return node; }
  expression(min) {
    if (++this.depth > 128) fail('#NUM!', 'Formula nesting limit exceeded.');
    let left = this.atom();
    while (true) {
      if (this.peek('%')) { this.take(); left = { type: 'unary', op: '%', node: left }; continue; }
      const op = this.peek().value, p = precedence[op]; if (p === undefined || p < min) break;
      this.take(); const right = this.expression(p + 1); left = { type: 'binary', op, left, right };
    }
    this.depth--; return left;
  }
  atom() {
    const t = this.take();
    if (t.type === 'number' || t.type === 'string') return { type: 'literal', value: t.value };
    if (t.type === 'error') return { type: 'error', value: t.value };
    if (t.value === '+' || t.value === '-') return { type: 'unary', op: t.value, node: this.expression(55) };
    if (t.value === '(') { const n = this.expression(0); this.expect(')'); return n; }
    if (t.type !== 'id' && t.type !== 'sheet') fail('#ERROR!', 'Expected a value, reference, or function.');
    if (this.peek('(')) {
      this.take(); const args = [];
      if (!this.peek(')')) while (true) {
        args.push(this.peek(',') || this.peek(')') ? { type: 'literal', value: null } : this.expression(0));
        if (!this.peek(',')) break; this.take();
      }
      this.expect(')'); return { type: 'call', name: t.value.toUpperCase(), args };
    }
    let sheet = null, value = t.value;
    if (this.peek('!')) { sheet = value; this.take(); const ref = this.take(); if (ref.type === 'error') return { type: 'error', value: ref.value }; value = ref.value; }
    if (this.peek(':')) {
      this.take(); const end = this.take().value;
      let a = parseAddress(value), b = parseAddress(end), whole = false;
      if (!a && !b && /^\$?[A-Z]{1,3}$/i.test(value) && /^\$?[A-Z]{1,3}$/i.test(end)) {
        a = parseAddress(value.replace('$', '') + '1'); b = parseAddress(end.replace('$', '') + MAX_ROWS); whole = true;
      }
      if (!a || !b) fail('#REF!'); return { type: 'range', sheet, ...normalizedRange(a, b), whole };
    }
    const ref = parseAddress(value);
    if (ref) return { type: 'ref', sheet, ...ref };
    if (/^TRUE$/i.test(value)) return { type: 'literal', value: true };
    if (/^FALSE$/i.test(value)) return { type: 'literal', value: false };
    return { type: 'name', value, sheet };
  }
}
function referenceTokens(source) {
  let tokens; try { tokens = tokenizeFormula(source); } catch { return []; }
  const result = [];
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]; if (t.type !== 'id' || !cellPattern.test(t.value) || tokens[i + 1]?.value === '!' || tokens[i + 1]?.value === '(') continue;
    let sheet = tokens[i - 1]?.value === '!' ? tokens[i - 2]?.value : null;
    if (tokens[i - 1]?.value === ':') sheet = result.at(-1)?.sheet ?? null;
    result.push({ ...t, sheet, index: i, pair: tokens[i - 1]?.value === ':' });
  }
  return result;
}
function replaceSpans(source, edits) { for (const e of edits.sort((a, b) => b.start - a.start)) source = source.slice(0, e.start) + e.text + source.slice(e.end); return source; }
export function shiftFormula(source, dr, dc) {
  if (!source.startsWith('=')) return source;
  return replaceSpans(source, referenceTokens(source).map(t => {
    const a = parseAddress(t.value); if (!a) return { ...t, text: '#REF!' };
    const r = a.r + (/\$\d/.test(t.value) ? 0 : dr), c = a.c + (t.value[0] === '$' ? 0 : dc);
    const text = r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS ? '#REF!' : `${t.value[0] === '$' ? '$' : ''}${colName(c)}${/\$\d/.test(t.value) ? '$' : ''}${r + 1}`;
    return { ...t, text };
  }));
}
function rewriteStructure(source, formulaSheet, targetSheet, axis, at, delta) {
  if (!source.startsWith('=')) return source;
  const refs = referenceTokens(source), edits = [];
  const format = (t, p) => `${t.value.startsWith('$') ? '$' : ''}${colName(p.c)}${/\$\d/.test(t.value) ? '$' : ''}${p.r + 1}`;
  for (let i = 0; i < refs.length; i++) {
    const t = refs[i]; if ((t.sheet ?? formulaSheet).toLowerCase() !== targetSheet.toLowerCase()) continue;
    const a = parseAddress(t.value); if (!a) continue;
    const coord = axis === 'row' ? 'r' : 'c';
    if (refs[i + 1]?.pair) {
      const end = refs[++i], b = parseAddress(end.value); if (!b) continue;
      const low = Math.min(a[coord], b[coord]), high = Math.max(a[coord], b[coord]);
      if (delta < 0 && low === high && low === at) { edits.push({ start: t.start, end: end.end, text: '#REF!' }); continue; }
      if (delta > 0) { if (a[coord] >= at) a[coord]++; if (b[coord] >= at) b[coord]++; }
      else {
        const reversed = a[coord] > b[coord];
        let lo = low > at ? low - 1 : low, hi = high >= at ? high - 1 : high;
        a[coord] = reversed ? hi : lo; b[coord] = reversed ? lo : hi;
      }
      edits.push({ ...t, text: format(t, a) }, { ...end, text: format(end, b) }); continue;
    }
    if (delta < 0 && a[coord] === at) { edits.push({ ...t, text: '#REF!' }); continue; }
    if (a[coord] >= at) a[coord] += delta;
    edits.push({ ...t, text: a.r >= MAX_ROWS || a.c >= MAX_COLS ? '#REF!' : format(t, a) });
  }
  return replaceSpans(source, edits);
}
function compare(a, b) {
  a = scalar(a); b = scalar(b);
  if (a == null) a = typeof b === 'string' ? '' : 0;
  if (b == null) b = typeof a === 'string' ? '' : 0;
  if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase().localeCompare(b.toLowerCase());
  if (typeof a !== typeof b) return (typeof a === 'number' ? 0 : typeof a === 'string' ? 1 : 2) - (typeof b === 'number' ? 0 : typeof b === 'string' ? 1 : 2);
  return a < b ? -1 : a > b ? 1 : 0;
}
function criteriaPredicate(crit) {
  crit = scalar(crit);
  if (typeof crit !== 'string') return v => compare(v, crit) === 0;
  const m = /^(<=|>=|<>|=|<|>)(.*)$/.exec(crit), op = m?.[1] ?? '=', text = m?.[2] ?? crit;
  const target = text !== '' && Number.isFinite(Number(text)) ? Number(text) : text;
  if (typeof target === 'string' && /[*?~]/.test(target) && (op === '=' || op === '<>')) {
    let pattern = '';
    for (let i = 0; i < target.length; i++) {
      let ch = target[i];
      if (ch === '~' && i + 1 < target.length) { ch = target[++i]; pattern += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
      else pattern += ch === '*' ? '.*' : ch === '?' ? '.' : ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    const re = new RegExp(`^${pattern}$`, 'i'); return v => re.test(str(v)) === (op === '=');
  }
  return v => { const c = compare(v, target); return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[op]; };
}
const DAY = 86400000;
const EPOCH = Date.UTC(1899, 11, 31);
export function dateSerial(date) { const days = Math.floor((date.getTime() - EPOCH) / DAY); return days >= 60 ? days + 1 : days; }
export function serialDate(serial) { return new Date(EPOCH + (serial >= 60 ? serial - 1 : serial) * DAY); }
export const FUNCTIONS = new Map();
function register(name, min, max, execute, description) { FUNCTIONS.set(name, { min, max, execute, description }); }
register('SUM', 1, Infinity, a => sum(numeric(a)), 'Adds numbers in a range.');
register('AVERAGE', 1, Infinity, a => { const v = numeric(a); return v.length ? sum(v) / v.length : fail('#DIV/0!'); }, 'Returns the arithmetic mean.');
register('MIN', 1, Infinity, a => { const v = numeric(a); return v.length ? v.reduce((a, b) => Math.min(a, b), Infinity) : 0; }, 'Returns the smallest number.');
register('MAX', 1, Infinity, a => { const v = numeric(a); return v.length ? v.reduce((a, b) => Math.max(a, b), -Infinity) : 0; }, 'Returns the largest number.');
register('COUNT', 1, Infinity, a => flatten(a).filter(x => typeof x === 'number').length, 'Counts numeric cells.');
register('COUNTA', 1, Infinity, a => flatten(a).filter(x => x !== null && x !== undefined).length, 'Counts nonempty values.');
register('COUNTBLANK', 1, 1, a => flatten(a).filter(x => x === null || x === '').length, 'Counts blank cells.');
register('PRODUCT', 1, Infinity, a => { const v = numeric(a); return v.length ? v.reduce((x, y) => x * y, 1) : 0; }, 'Multiplies numbers.');
register('MEDIAN', 1, Infinity, a => { const v = numeric(a).sort((x, y) => x - y); const n = v.length; return n ? n % 2 ? v[n >> 1] : (v[n / 2 - 1] + v[n / 2]) / 2 : fail('#NUM!'); }, 'Returns the middle value.');
register('SUMPRODUCT', 1, Infinity, a => {
  const arr = a.map(x => x instanceof RangeValue ? x.flat() : [x]); if (arr.some(x => x.length !== arr[0].length)) fail('#VALUE!');
  return sum(arr[0].map((_, i) => arr.reduce((v, xs) => v * (typeof checkError(xs[i]) === 'number' ? xs[i] : 0), 1)));
}, 'Sums products of corresponding values.');
for (const [n, f] of Object.entries({ ABS: Math.abs, SQRT: Math.sqrt, INT: Math.floor, EXP: Math.exp, LN: Math.log, LOG10: Math.log10, SIN: Math.sin, COS: Math.cos, TAN: Math.tan, SIGN: Math.sign })) register(n, 1, 1, a => safeResult(f(number(a[0]))), `${n.toLowerCase()} of a number.`);
register('PI', 0, 0, () => Math.PI, 'Returns pi.');
register('POWER', 2, 2, a => safeResult(number(a[0]) ** number(a[1])), 'Raises a number to a power.');
register('MOD', 2, 2, a => { const x = number(a[0]), y = number(a[1]); return y ? x - y * Math.floor(x / y) : fail('#DIV/0!'); }, 'Returns the remainder.');
for (const name of ['ROUND', 'ROUNDUP', 'ROUNDDOWN', 'TRUNC']) register(name, name === 'TRUNC' ? 1 : 2, 2, a => {
  const x = number(a[0]), d = Math.trunc(number(a[1] ?? 0)); if (Math.abs(d) > 308) fail('#NUM!');
  const factor = 10 ** d, abs = Math.abs(x) * factor;
  const rounded = name === 'ROUND' ? Math.floor(abs + 0.5 + Number.EPSILON * abs) : name === 'ROUNDUP' ? Math.ceil(abs) : Math.floor(abs);
  return safeResult(Math.sign(x) * rounded / factor);
}, 'Rounds a number.');
for (const name of ['AND', 'OR', 'XOR']) register(name, 1, Infinity, a => {
  const v = flatten(a).filter(x => typeof x !== 'string' && x !== null).map(logical); if (!v.length) fail('#VALUE!');
  return name === 'AND' ? v.every(Boolean) : name === 'OR' ? v.some(Boolean) : v.filter(Boolean).length % 2 === 1;
}, 'Combines logical tests.');
register('NOT', 1, 1, a => !logical(a[0]), 'Reverses a logical value.');
for (const [name, test] of Object.entries({ ISNUMBER: x => typeof x === 'number', ISTEXT: x => typeof x === 'string', ISBLANK: x => x === null, ISLOGICAL: x => typeof x === 'boolean' })) register(name, 1, 1, a => test(scalar(a[0])), 'Tests a value’s type.');
register('N', 1, 1, a => { const x = scalar(a[0]); return typeof x === 'number' ? x : typeof x === 'boolean' ? +x : 0; }, 'Converts a value to a number.');
register('NA', 0, 0, () => fail('#N/A'), 'Returns the unavailable error.');
register('LEN', 1, 1, a => str(a[0]).length, 'Counts characters.');
register('LEFT', 1, 2, a => { const n = number(a[1] ?? 1); return n < 0 ? fail('#VALUE!') : str(a[0]).slice(0, n); }, 'Returns leading characters.');
register('RIGHT', 1, 2, a => { const n = number(a[1] ?? 1); return n < 0 ? fail('#VALUE!') : n ? str(a[0]).slice(-n) : ''; }, 'Returns trailing characters.');
register('MID', 3, 3, a => { const s = number(a[1]), n = number(a[2]); return s < 1 || n < 0 ? fail('#VALUE!') : str(a[0]).slice(s - 1, s - 1 + n); }, 'Returns text from a position.');
register('TRIM', 1, 1, a => str(a[0]).trim().replace(/ +/g, ' '), 'Removes extra spaces.');
register('UPPER', 1, 1, a => str(a[0]).toUpperCase(), 'Converts text to uppercase.');
register('LOWER', 1, 1, a => str(a[0]).toLowerCase(), 'Converts text to lowercase.');
register('PROPER', 1, 1, a => str(a[0]).toLowerCase().replace(/\b\w/g, x => x.toUpperCase()), 'Capitalizes words.');
register('CONCAT', 1, Infinity, a => flatten(a).map(str).join(''), 'Joins text and ranges.');
register('CONCATENATE', 1, Infinity, a => a.map(str).join(''), 'Joins text values.');
register('TEXTJOIN', 3, Infinity, a => flatten(a.slice(2)).filter(x => !logical(a[1]) || (x !== null && x !== '')).map(str).join(str(a[0])), 'Joins text with a delimiter.');
register('SUBSTITUTE', 3, 4, a => {
  const s = str(a[0]), old = str(a[1]), next = str(a[2]); if (!old) return s;
  if (a.length < 4) return s.split(old).join(next);
  const at = number(a[3]); if (at < 1) fail('#VALUE!'); let n = 0;
  return s.split(old).reduce((all, p, i) => i ? all + (++n === at ? next : old) + p : p, '');
}, 'Replaces matching text.');
for (const name of ['FIND', 'SEARCH']) register(name, 2, 3, a => {
  let needle = str(a[0]), text = str(a[1]); const start = number(a[2] ?? 1); if (start < 1) fail('#VALUE!');
  if (name === 'SEARCH') { needle = needle.toLowerCase(); text = text.toLowerCase(); }
  const i = text.indexOf(needle, start - 1); return i < 0 ? fail('#VALUE!') : i + 1;
}, 'Finds a text position.');
register('VALUE', 1, 1, a => number(str(a[0]).replace(/[$,€£]/g, '')), 'Converts numeric text.');
register('REPT', 2, 2, a => { const n = Math.trunc(number(a[1])); const text = str(a[0]); return n < 0 || n * text.length > 32767 ? fail('#VALUE!') : text.repeat(n); }, 'Repeats text.');
register('COUNTIF', 2, 2, a => { const p = criteriaPredicate(a[1]); return flatten([a[0]]).filter(p).length; }, 'Counts matching cells.');
for (const name of ['SUMIF', 'AVERAGEIF']) register(name, 2, 3, a => {
  const test = flatten([a[0]]), values = flatten([a[2] ?? a[0]]), p = criteriaPredicate(a[1]); if (test.length !== values.length) fail('#VALUE!');
  const v = values.filter((x, i) => p(test[i]) && typeof checkError(x) === 'number');
  return name === 'SUMIF' ? sum(v) : v.length ? sum(v) / v.length : fail('#DIV/0!');
}, 'Aggregates values matching a condition.');
for (const name of ['SUMIFS', 'COUNTIFS', 'AVERAGEIFS']) register(name, name === 'COUNTIFS' ? 2 : 3, Infinity, a => {
  const count = name === 'COUNTIFS', first = count ? 0 : 1;
  if ((a.length - first) % 2 !== 0) fail('#VALUE!');
  const pairs = []; for (let i = first; i < a.length; i += 2) pairs.push([flatten([a[i]]), criteriaPredicate(a[i + 1])]);
  const v = count ? pairs[0][0] : flatten([a[0]]); if (pairs.some(p => p[0].length !== v.length)) fail('#VALUE!');
  const match = v.filter((x, i) => pairs.every(([xs, pred]) => pred(xs[i])));
  if (count) return match.length;
  const nums = match.filter(x => typeof checkError(x) === 'number');
  return name === 'SUMIFS' ? sum(nums) : nums.length ? sum(nums) / nums.length : fail('#DIV/0!');
}, 'Aggregates with multiple conditions.');
register('INDEX', 2, 3, a => {
  const rows = a[0] instanceof RangeValue ? a[0].rows : [[a[0]]], r = number(a[1]) - 1, c = number(a[2] ?? 1) - 1;
  if (r < 0 || c < 0 || r >= rows.length || c >= rows[0].length) fail('#REF!'); return checkError(rows[r][c]);
}, 'Returns a value at a row and column.');
register('MATCH', 2, 3, a => {
  const values = flatten([a[1]]), mode = number(a[2] ?? 1); let best = -1;
  for (let i = 0; i < values.length; i++) { const c = compare(values[i], a[0]); if (c === 0) return i + 1; if (mode === 1 && c <= 0 || mode === -1 && c >= 0) best = i; }
  return best < 0 ? fail('#N/A') : best + 1;
}, 'Finds the position of a value.');
register('VLOOKUP', 3, 4, a => {
  if (!(a[1] instanceof RangeValue)) fail('#VALUE!'); const rows = a[1].rows, col = number(a[2]) - 1, approx = a.length === 3 || logical(a[3]);
  if (col < 0 || col >= (rows[0]?.length ?? 0)) fail('#REF!'); let found;
  for (const row of rows) { const c = compare(row[0], a[0]); if (c === 0) return checkError(row[col]); if (approx && c <= 0) found = row[col]; }
  return found === undefined ? fail('#N/A') : checkError(found);
}, 'Looks up a value in the first column.');
register('XLOOKUP', 3, 6, a => {
  const keys = flatten([a[1]]), vals = flatten([a[2]]); if (keys.length !== vals.length) fail('#VALUE!');
  const mode = number(a[4] ?? 0), search = number(a[5] ?? 1); if (![0, -1, 1, 2].includes(mode) || ![1, -1].includes(search)) fail('#VALUE!', 'Binary XLOOKUP search modes are not supported.');
  const indices = keys.map((_, i) => i); if (search === -1) indices.reverse(); let candidate = -1;
  for (const i of indices) {
    const cmp = compare(keys[i], a[0]); if (mode === 2 ? criteriaPredicate(a[0])(keys[i]) : cmp === 0) return checkError(vals[i]);
    if (mode === -1 && cmp < 0 && (candidate < 0 || compare(keys[i], keys[candidate]) > 0) || mode === 1 && cmp > 0 && (candidate < 0 || compare(keys[i], keys[candidate]) < 0)) candidate = i;
  }
  return candidate >= 0 ? checkError(vals[candidate]) : a.length >= 4 ? a[3] : fail('#N/A');
}, 'Looks up an exact or nearest match.');
register('DATE', 3, 3, a => { let y = number(a[0]); if (y >= 0 && y < 1900) y += 1900; return dateSerial(new Date(Date.UTC(y, number(a[1]) - 1, number(a[2])))); }, 'Creates a spreadsheet date serial.');
for (const name of ['YEAR', 'MONTH', 'DAY']) register(name, 1, 1, a => { const d = serialDate(number(a[0])); return name === 'YEAR' ? d.getUTCFullYear() : name === 'MONTH' ? d.getUTCMonth() + 1 : d.getUTCDate(); }, 'Extracts a calendar component.');
register('TODAY', 0, 0, () => dateSerial(new Date()), 'Returns today’s date.');
register('NOW', 0, 0, () => { const d = new Date(); return dateSerial(d) + (d.getUTCHours() * 3600 + d.getUTCMinutes() * 60 + d.getUTCSeconds()) / 86400; }, 'Returns the current date and UTC time.');
register('RAND', 0, 0, () => Math.random(), 'Returns a random number between 0 and 1.');
register('RANDBETWEEN', 2, 2, a => { const lo = Math.ceil(number(a[0])), hi = Math.floor(number(a[1])); return hi < lo ? fail('#NUM!') : Math.floor(Math.random() * (hi - lo + 1)) + lo; }, 'Returns a random integer.');
register('TEXT', 2, 2, a => formatValue(scalar(a[0]), { format: str(a[1]) }), 'Formats a number as text.');
register('ROWS', 1, 1, a => a[0] instanceof RangeValue ? a[0].rows.length : 1, 'Counts rows in a range.');
register('COLUMNS', 1, 1, a => a[0] instanceof RangeValue ? a[0].rows[0]?.length ?? 0 : 1, 'Counts columns in a range.');
for (const n of ['IF', 'IFERROR', 'IFNA', 'IFS', 'ISERROR', 'ISNA', 'CHOOSE', 'ROW', 'COLUMN']) register(n, 0, Infinity, () => null, 'Evaluated lazily by the calculation engine.');

export class CalculationEngine {
  constructor(workbook) { this.workbook = workbook; this.cache = new Map(); this.astCache = new Map(); this.dependencies = new Map(); this.dependents = new Map(); this.active = new Set(); this.stack = []; this.evaluations = 0; this.rangeDependencies = new Map(); this.volatile = new Set(); }
  id(sheet, r, c) { return `${sheet.id}!${r},${c}`; }
  invalidate(ids = null) {
    if (!ids) { this.cache.clear(); this.dependencies.clear(); this.dependents.clear(); this.active.clear(); this.rangeDependencies.clear(); this.volatile.clear(); return; }
    const queue = [...ids, ...this.volatile], visited = new Set();
    // Whole-column ranges are evaluated only to the used row, but must observe future rows too.
    for (const id of ids) {
      const bang = id.indexOf('!'), sid = id.slice(0, bang), [r, c] = id.slice(bang + 1).split(',').map(Number);
      for (const [owner, ranges] of this.rangeDependencies) if (ranges.some(q => q.sid === sid && r >= q.r1 && r <= q.r2 && c >= q.c1 && c <= q.c2)) queue.push(owner);
    }
    for (let i = 0; i < queue.length; i++) {
      const id = queue[i]; if (visited.has(id)) continue; visited.add(id); this.cache.delete(id);
      for (const d of this.dependents.get(id) ?? []) queue.push(d);
    }
  }
  depend(on) {
    const current = this.stack.at(-1); if (!current) return;
    if (!this.dependencies.has(current)) this.dependencies.set(current, new Set()); this.dependencies.get(current).add(on);
    if (!this.dependents.has(on)) this.dependents.set(on, new Set()); this.dependents.get(on).add(current);
  }
  get(sheet, r, c) {
    const id = this.id(sheet, r, c); this.depend(id);
    if (this.cache.has(id)) return this.cache.get(id);
    if (this.active.has(id)) return new FormulaError('#CYCLE!', 'Circular reference.');
    if (this.stack.length > 256) return new FormulaError('#NUM!', 'Calculation depth limit exceeded.');
    const raw = sheet.cells.get(keyOf(r, c))?.raw ?? '';
    if (!raw.startsWith('=')) return rawValue(raw);
    for (const old of this.dependencies.get(id) ?? []) this.dependents.get(old)?.delete(id);
    this.dependencies.set(id, new Set()); this.rangeDependencies.delete(id); this.volatile.delete(id); if (/\b(TODAY|NOW|RAND|RANDBETWEEN)\s*\(/i.test(raw)) this.volatile.add(id); this.active.add(id); this.stack.push(id); this.evaluations++;
    let value;
    try {
      let ast = this.astCache.get(raw);
      if (!ast) { ast = new FormulaParser(raw).parse(); if (this.astCache.size > 20000) this.astCache.clear(); this.astCache.set(raw, ast); }
      value = scalar(this.evaluate(ast, sheet, { r, c })); value = safeResult(value === null ? 0 : value);
    } catch (error) { value = error instanceof FormulaError ? error : new FormulaError('#ERROR!', error.message); }
    finally { this.stack.pop(); this.active.delete(id); }
    this.cache.set(id, value); return value;
  }
  resolveSheet(name, current) { const sheet = name ? this.workbook.sheetByName(name) : current; if (!sheet) fail('#REF!', `Unknown sheet: ${name}`); return sheet; }
  evaluate(n, sheet, here) {
    if (n.type === 'literal') return n.value;
    if (n.type === 'error') fail(n.value);
    if (n.type === 'ref') return checkError(this.get(this.resolveSheet(n.sheet, sheet), n.r, n.c));
    if (n.type === 'name') {
      const target = this.workbook.names[n.value.toUpperCase()]; if (!target) fail('#NAME?', `Unknown name: ${n.value}`);
      return this.evaluate(new FormulaParser('=' + target).parse(), sheet, here);
    }
    if (n.type === 'range') {
      const source = this.resolveSheet(n.sheet, sheet); let last = n.r2;
      if (n.whole) {
        last = Math.max(source.usedRange().r2, 0);
        const owner = this.stack.at(-1); if (owner) { if (!this.rangeDependencies.has(owner)) this.rangeDependencies.set(owner, []); this.rangeDependencies.get(owner).push({ sid: source.id, r1: n.r1, r2: n.r2, c1: n.c1, c2: n.c2 }); }
      }
      if ((last - n.r1 + 1) * (n.c2 - n.c1 + 1) > MAX_RANGE_CELLS) fail('#NUM!', 'Range evaluation limit exceeded.');
      const rows = []; for (let r = n.r1; r <= last; r++) { const row = []; for (let c = n.c1; c <= n.c2; c++) row.push(this.get(source, r, c)); rows.push(row); }
      return new RangeValue(rows, { sheet: source, ...n });
    }
    if (n.type === 'unary') { const x = number(this.evaluate(n.node, sheet, here)); return n.op === '-' ? -x : n.op === '%' ? x / 100 : x; }
    if (n.type === 'binary') {
      const a = scalar(this.evaluate(n.left, sheet, here)), b = scalar(this.evaluate(n.right, sheet, here));
      if (['=', '<>', '<', '>', '<=', '>='].includes(n.op)) { const c = compare(a, b); return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[n.op]; }
      if (n.op === '&') return str(a) + str(b); const x = number(a), y = number(b);
      if (n.op === '/' && y === 0) fail('#DIV/0!');
      return safeResult(n.op === '+' ? x + y : n.op === '-' ? x - y : n.op === '*' ? x * y : n.op === '/' ? x / y : x ** y);
    }
    if (n.type !== 'call') fail('#ERROR!');
    const name = n.name, args = n.args, evalAt = i => i < args.length ? this.evaluate(args[i], sheet, here) : null;
    const arity = (lo, hi = lo) => { if (args.length < lo || args.length > hi) fail('#VALUE!', `${name}: invalid argument count.`); };
    if (name === 'IF') { arity(2, 3); return logical(evalAt(0)) ? evalAt(1) : args.length === 3 ? evalAt(2) : false; }
    if (name === 'IFERROR' || name === 'IFNA') { arity(2); try { return checkError(scalar(evalAt(0))); } catch (e) { if (e instanceof FormulaError && (name === 'IFERROR' || e.code === '#N/A')) return evalAt(1); throw e; } }
    if (['ISNUMBER','ISTEXT','ISBLANK','ISLOGICAL'].includes(name)) { arity(1); try { const x = scalar(evalAt(0)); return name === 'ISNUMBER' ? typeof x === 'number' : name === 'ISTEXT' ? typeof x === 'string' : name === 'ISBLANK' ? x === null : typeof x === 'boolean'; } catch (e) { if (e instanceof FormulaError) return false; throw e; } }
    if (name === 'ISERROR' || name === 'ISNA') { arity(1); try { scalar(evalAt(0)); return false; } catch (e) { return name === 'ISERROR' || e.code === '#N/A'; } }
    if (name === 'IFS') { if (args.length < 2 || args.length % 2) fail('#VALUE!'); for (let i = 0; i < args.length; i += 2) if (logical(evalAt(i))) return evalAt(i + 1); fail('#N/A'); }
    if (name === 'CHOOSE') { arity(2, 255); const i = Math.trunc(number(evalAt(0))); if (i < 1 || i >= args.length) fail('#VALUE!'); return evalAt(i); }
    if (name === 'ROW' || name === 'COLUMN') { arity(0, 1); const ref = args[0]; if (ref && !['ref', 'range'].includes(ref.type)) fail('#VALUE!'); return name === 'ROW' ? (ref?.r ?? ref?.r1 ?? here.r) + 1 : (ref?.c ?? ref?.c1 ?? here.c) + 1; }
    const fn = FUNCTIONS.get(name); if (!fn) fail('#NAME?', `Unsupported function: ${name}`);
    arity(fn.min, fn.max); return safeResult(fn.execute(args.map((_, i) => evalAt(i))));
  }
}
export function rawValue(raw) {
  if (raw === '' || raw == null) return null;
  if (raw[0] === "'") return raw.slice(1);
  if (/^(true|false)$/i.test(raw)) return /^true$/i.test(raw);
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?%?$/i.test(raw.trim())) { const n = Number(raw.trim().replace('%', '')); if (!Number.isFinite(n)) return new FormulaError('#NUM!'); return raw.trim().endsWith('%') ? n / 100 : n; }
  if (/^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|CYCLE!)$/.test(raw)) return new FormulaError(raw);
  return raw;
}
const formatters = new Map();
function numFormat(locale, options, value) { const key = JSON.stringify([locale, options]); if (!formatters.has(key)) formatters.set(key, new Intl.NumberFormat(locale, options)); return formatters.get(key).format(value); }
export function formatValue(value, style = {}) {
  if (value instanceof FormulaError) return value.code;
  if (value === null || value === undefined) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value !== 'number') return String(value);
  const fmt = style.format || 'general', dp = style.decimals;
  if (fmt === 'date' || /[ymd]/i.test(fmt) && !/[Ee][+-]/.test(fmt) && fmt !== 'number' && fmt !== 'currency') return serialDate(value).toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric', timeZone: 'UTC' });
  if (fmt === 'percent' || fmt.includes('%')) return numFormat('en-US', { style: 'percent', minimumFractionDigits: dp ?? (fmt === 'percent' ? 1 : (fmt.split('.')[1]?.match(/[0#]/g)?.length ?? 0)), maximumFractionDigits: dp ?? (fmt === 'percent' ? 1 : (fmt.split('.')[1]?.match(/[0#]/g)?.length ?? 0)) }, value);
  if (fmt === 'currency' || /[$€£]/.test(fmt)) return numFormat('en-US', { style: 'currency', currency: fmt.includes('€') ? 'EUR' : fmt.includes('£') ? 'GBP' : 'USD', minimumFractionDigits: dp ?? 0, maximumFractionDigits: dp ?? 0 }, value);
  if (fmt === 'number' || fmt === 'integer' || /[0#]/.test(fmt)) { const d = dp ?? (fmt === 'integer' ? 0 : fmt === 'number' ? 2 : (fmt.split('.')[1]?.match(/[0#]/g)?.length ?? 0)); return numFormat('en-US', { useGrouping: true, minimumFractionDigits: d, maximumFractionDigits: d }, value); }
  if (dp !== undefined) return numFormat('en-US', { useGrouping: false, minimumFractionDigits: dp, maximumFractionDigits: dp }, value);
  return Math.abs(value) >= 1e12 || Math.abs(value) < 1e-8 && value !== 0 ? value.toExponential(5) : String(Number(value.toPrecision(12)));
}
let nextSheetId = 1;
export class Sheet {
  constructor(name = 'Sheet1', data = null) {
    this.id = `s${Date.now().toString(36)}${nextSheetId++}`; this.name = name; this.cells = new Map(); this.colWidths = new Map(); this.rowHeights = new Map();
    this.merges = []; this.conditionalRules = []; this.hiddenRows = new Set(); this.filters = null; this.freezeRows = 0; this.freezeCols = 0; this.charts = []; this.gridlines = true; this.color = '#18835a'; this.revision = 0; this._used = null;
    if (data) { for (const prop of ['id','name','merges','conditionalRules','filters','freezeRows','freezeCols','charts','gridlines','color','revision','protected','dataRegion']) if (Object.hasOwn(data, prop)) this[prop] = data[prop]; this.cells = new Map(data.cells ?? []); this.colWidths = new Map(data.colWidths ?? []); this.rowHeights = new Map(data.rowHeights ?? []); this.hiddenRows = new Set(data.hiddenRows ?? []); this._used = null; }
  }
  get(r, c) { return this.cells.get(keyOf(r, c)); }
  raw(r, c) { return this.get(r, c)?.raw ?? ''; }
  usedRange() {
    if (this._used) return this._used;
    let r2 = 0, c2 = 0; for (const [key, cell] of this.cells) { if (!cell.raw && !cell.style) continue; const [r, c] = key.split(',').map(Number); r2 = Math.max(r2, r); c2 = Math.max(c2, c); }
    return this._used = { r1: 0, c1: 0, r2, c2 };
  }
  mergeAt(r, c) { return this.merges.find(q => r >= q.r1 && r <= q.r2 && c >= q.c1 && c <= q.c2); }
  toJSON() {
    const { _used, ...rest } = this; return { ...rest, cells: [...this.cells], colWidths: [...this.colWidths], rowHeights: [...this.rowHeights], hiddenRows: [...this.hiddenRows] };
  }
}
export class Workbook {
  constructor() { this.title = 'Untitled workbook'; this.sheets = [new Sheet()]; this.activeSheetId = this.sheets[0].id; this.names = {}; this.listeners = new Set(); this.engine = new CalculationEngine(this); this.undoStack = []; this.redoStack = []; this._transaction = null; this.revision = 0; }
  get activeSheet() { return this.sheets.find(x => x.id === this.activeSheetId) ?? this.sheets[0]; }
  sheetByName(name) { return this.sheets.find(x => x.name.toLowerCase() === name.toLowerCase()); }
  onChange(fn) { this.listeners.add(fn); return () => this.listeners.delete(fn); }
  emit(label, full = false) { this.revision++; for (const fn of this.listeners) fn({ label, full, revision: this.revision }); }
  toJSON() { return { format: 'gridline', version: 1, title: this.title, activeSheetId: this.activeSheetId, names: this.names, sheets: this.sheets.map(s => s.toJSON()) }; }
  static fromJSON(data) {
    if (!data || data.format !== 'gridline' || data.version !== 1 || !Array.isArray(data.sheets) || !data.sheets.length || data.sheets.length > 256) throw new Error('Not a supported Gridline workbook.');
    const wb = new Workbook(); wb.title = String(data.title ?? 'Workbook').slice(0, 200); wb.names = data.names && typeof data.names === 'object' ? { ...data.names } : {};
    wb.sheets = data.sheets.map(d => {
      if (!Array.isArray(d.cells) || d.cells.length > 1000000) throw new Error('Workbook cell limit exceeded.');
      const sheet = new Sheet(String(d.name).slice(0, 31), d); sheet.name = String(d.name).slice(0, 31);
      for (const [k, cell] of sheet.cells) { const [r, c] = k.split(',').map(Number); if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS || typeof cell.raw !== 'string') throw new Error('Invalid cell record.'); }
      return sheet;
    });
    if (new Set(wb.sheets.map(s => s.id)).size !== wb.sheets.length || new Set(wb.sheets.map(s => s.name.toLowerCase())).size !== wb.sheets.length) throw new Error('Duplicate sheet identities.');
    wb.activeSheetId = data.activeSheetId; return wb;
  }
  restore(data) { const w = Workbook.fromJSON(data); this.title = w.title; this.sheets = w.sheets; this.activeSheetId = w.activeSheetId; this.names = w.names; this.engine.invalidate(); }
  value(sheet, r, c) { return this.engine.get(sheet, r, c); }
  display(sheet, r, c) { const cell = sheet.get(r, c); let style = cell?.style ?? {}; if (!style.format && cell?.raw.trim().endsWith('%')) style = { ...style, format: 'percent' }; return formatValue(this.value(sheet, r, c), style); }
  transaction(label, fn) {
    if (this._transaction) return fn();
    const tx = { label, changes: new Map() }; this._transaction = tx;
    try { fn(); } catch (e) {
      for (const change of tx.changes.values()) { const s = this.sheets.find(s => s.id === change.sid); if (change.before === undefined) s.cells.delete(change.key); else s.cells.set(change.key, change.before); s._used = null; }
      this.engine.invalidate(); throw e;
    } finally { this._transaction = null; }
    if (!tx.changes.size) return;
    const changes = [...tx.changes.values()]; for (const x of changes) x.after = clone(this.sheets.find(s => s.id === x.sid).cells.get(x.key));
    this.undoStack.push({ kind: 'cells', label, changes }); this.trimHistory(); this.redoStack = [];
    this.engine.invalidate(changes.map(x => `${x.sid}!${x.key}`)); this.emit(label);
  }
  trimHistory() { if (this.undoStack.length > 100) this.undoStack.shift(); }
  setCell(sheet, r, c, patch) {
    if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS) throw new Error('Cell is outside the sheet.');
    if (!this._transaction) return this.transaction('Edit cell', () => this.setCell(sheet, r, c, patch));
    const key = keyOf(r, c), id = `${sheet.id}!${key}`, before = sheet.cells.get(key);
    if (!this._transaction.changes.has(id)) this._transaction.changes.set(id, { sid: sheet.id, key, before: clone(before) });
    if (patch === null) sheet.cells.delete(key);
    else {
      const next = { raw: '', ...clone(before), ...clone(patch) }; next.raw = String(next.raw).slice(0, 32767);
      if (patch.style) next.style = { ...before?.style, ...clone(patch.style) };
      if (!next.raw && !next.style && !next.note) sheet.cells.delete(key); else sheet.cells.set(key, next);
    }
    sheet.revision++; sheet._used = null;
    this.engine.invalidate([id]);
  }
  setRaw(sheet, r, c, raw) { this.setCell(sheet, r, c, { raw: String(raw) }); }
  applyStyle(sheet, q, style) { this.transaction('Format cells', () => { for (const { r, c } of cellsIn(q)) this.setCell(sheet, r, c, { style }); }); }
  clear(sheet, q, all = false) { this.transaction('Clear cells', () => { for (const [key] of sheet.cells) { const [r, c] = key.split(',').map(Number); if (r >= q.r1 && r <= q.r2 && c >= q.c1 && c <= q.c2) this.setCell(sheet, r, c, all ? null : { raw: '' }); } }); }
  mutate(label, fn) {
    const before = clone(this.toJSON());
    try { fn(); } catch (e) { this.restore(before); throw e; }
    this.engine.invalidate(); for (const s of this.sheets) { s._used = null; s.revision++; }
    const after = clone(this.toJSON()); this.undoStack.push({ kind: 'snapshot', label, before, after }); this.trimHistory(); this.redoStack = []; this.emit(label, true);
  }
  historyStep(from, to, mode) {
    const cmd = from.pop(); if (!cmd) return false;
    if (cmd.kind === 'snapshot') this.restore(cmd[mode]);
    else { for (const x of cmd.changes) { const s = this.sheets.find(s => s.id === x.sid); if (!s) continue; if (x[mode] === undefined) s.cells.delete(x.key); else s.cells.set(x.key, clone(x[mode])); s._used = null; s.revision++; } this.engine.invalidate(cmd.changes.map(x => `${x.sid}!${x.key}`)); }
    to.push(cmd); this.emit(`${mode === 'before' ? 'Undo' : 'Redo'} ${cmd.label}`, cmd.kind === 'snapshot'); return true;
  }
  undo() { return this.historyStep(this.undoStack, this.redoStack, 'before'); }
  redo() { return this.historyStep(this.redoStack, this.undoStack, 'after'); }
  uniqueSheetName(base = 'Sheet') { let name = base.slice(0, 31), i = 2; while (this.sheetByName(name)) name = `${base.slice(0, 26)} ${i++}`; return name; }
  addSheet(name = 'Sheet') { let sheet; this.mutate('Add sheet', () => { sheet = new Sheet(this.uniqueSheetName(name)); this.sheets.push(sheet); this.activeSheetId = sheet.id; }); return sheet; }
  renameSheet(sheet, name) {
    name = name.trim(); if (!name || name.length > 31 || /[\\/*?:\[\]]/.test(name)) throw new Error('Use a unique name of 1–31 characters without \\ / * ? : [ ].');
    if (this.sheets.some(s => s !== sheet && s.name.toLowerCase() === name.toLowerCase())) throw new Error('A sheet already has that name.');
    this.mutate('Rename sheet', () => {
      const old = sheet.name; sheet.name = name;
      const renameReferences = source => {
        let tokens; try { tokens = tokenizeFormula(source); } catch { return source; }
        return replaceSpans(source, tokens.filter((t, i) => tokens[i + 1]?.value === '!' && t.value.toLowerCase() === old.toLowerCase()).map(t => ({ ...t, text: `'${name.replaceAll("'", "''")}'` })));
      };
      for (const s of this.sheets) for (const cell of s.cells.values()) if (cell.raw.startsWith('=')) cell.raw = renameReferences(cell.raw);
      for (const [key, target] of Object.entries(this.names)) this.names[key] = renameReferences('=' + target).slice(1);
    });
  }
  duplicateSheet(sheet) { this.mutate('Duplicate sheet', () => { const copy = new Sheet('', clone(sheet.toJSON())); copy.id = `s${Date.now().toString(36)}${nextSheetId++}`; copy.name = this.uniqueSheetName(sheet.name + ' copy'); this.sheets.push(copy); this.activeSheetId = copy.id; }); }
  deleteSheet(sheet) { if (this.sheets.length === 1) throw new Error('A workbook needs at least one sheet.'); this.mutate('Delete sheet', () => { this.sheets = this.sheets.filter(s => s !== sheet); if (this.activeSheetId === sheet.id) this.activeSheetId = this.sheets[0].id; }); }
  fill(sheet, source, target) {
    this.transaction('Fill cells', () => {
      const h = source.r2 - source.r1 + 1, w = source.c2 - source.c1 + 1;
      const originals = new Map(); for (const p of cellsIn(source)) originals.set(keyOf(p.r, p.c), clone(sheet.get(p.r, p.c) ?? { raw: '' }));
      const verticalSeries = w === 1 && h === 2 && [source.r1, source.r2].every(r => typeof rawValue(sheet.raw(r, source.c1)) === 'number');
      const horizontalSeries = h === 1 && w === 2 && [source.c1, source.c2].every(c => typeof rawValue(sheet.raw(source.r1, c)) === 'number');
      for (const { r, c } of cellsIn(target)) {
        if (r >= source.r1 && r <= source.r2 && c >= source.c1 && c <= source.c2) continue;
        const sr = source.r1 + ((r - source.r1) % h + h) % h, sc = source.c1 + ((c - source.c1) % w + w) % w;
        const cell = clone(originals.get(keyOf(sr, sc))); cell.raw = shiftFormula(cell.raw, r - sr, c - sc);
        if (verticalSeries) { const start = number(rawValue(originals.get(keyOf(source.r1, source.c1)).raw)), end = number(rawValue(originals.get(keyOf(source.r2, source.c1)).raw)); cell.raw = String(start + (end - start) * (r - source.r1)); }
        if (horizontalSeries) { const start = number(rawValue(originals.get(keyOf(source.r1, source.c1)).raw)), end = number(rawValue(originals.get(keyOf(source.r1, source.c2)).raw)); cell.raw = String(start + (end - start) * (c - source.c1)); }
        this.setCell(sheet, r, c, cell);
      }
    });
  }
  structuralEdit(sheet, axis, at, delta) {
    this.mutate(`${delta > 0 ? 'Insert' : 'Delete'} ${axis}`, () => {
      const coord = axis === 'row' ? 0 : 1, limit = axis === 'row' ? MAX_ROWS : MAX_COLS, next = new Map();
      for (const [key, cell] of sheet.cells) { const p = key.split(',').map(Number); if (delta < 0 && p[coord] === at) continue; if (p[coord] >= at) p[coord] += delta; if (p[coord] < limit) next.set(keyOf(...p), cell); }
      sheet.cells = next;
      const sizes = axis === 'row' ? 'rowHeights' : 'colWidths', mapped = new Map();
      for (const [i, size] of sheet[sizes]) { if (delta < 0 && i === at) continue; const n = i >= at ? i + delta : i; if (n < limit) mapped.set(n, size); } sheet[sizes] = mapped;
      sheet.hiddenRows.clear(); sheet.filters = null; sheet.dataRegion = null;
      const a = axis === 'row' ? 'r1' : 'c1', b = axis === 'row' ? 'r2' : 'c2';
      sheet.merges = sheet.merges.filter(q => !(delta < 0 && q[a] === at && q[b] === at)).map(q => { const n = { ...q }; if (delta > 0) { if (n[a] >= at) n[a]++; if (n[b] >= at) n[b]++; } else { if (n[a] > at) n[a]--; if (n[b] >= at) n[b]--; } return n; });
      for (const s of this.sheets) for (const cell of s.cells.values()) cell.raw = rewriteStructure(cell.raw, s.name, sheet.name, axis, at, delta);
      for (const [name, value] of Object.entries(this.names)) this.names[name] = rewriteStructure('=' + value, sheet.name, sheet.name, axis, at, delta).slice(1);
      // Chart/conditional ranges are layout metadata; conservatively discard rules whose coordinates would become stale.
      sheet.conditionalRules = []; sheet.charts = [];
    });
  }
  sort(sheet, q, column, descending = false, header = true) {
    const first = q.r1 + (header ? 1 : 0); if (first > q.r2) return;
    const rows = []; for (let r = first; r <= q.r2; r++) rows.push({ r, value: this.value(sheet, r, column), cells: Array.from({ length: q.c2 - q.c1 + 1 }, (_, i) => clone(sheet.get(r, q.c1 + i) ?? { raw: '' })) });
    rows.sort((a, b) => { if (a.value == null) return b.value == null ? a.r - b.r : 1; if (b.value == null) return -1; return (compare(a.value instanceof FormulaError ? a.value.code : a.value, b.value instanceof FormulaError ? b.value.code : b.value) * (descending ? -1 : 1)) || a.r - b.r; });
    this.transaction('Sort range', () => rows.forEach((row, i) => row.cells.forEach((cell, j) => { cell.raw = shiftFormula(cell.raw, first + i - row.r, 0); this.setCell(sheet, first + i, q.c1 + j, cell); })));
  }
}
