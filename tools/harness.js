const XLSX = require('/tmp/xlsxbuild/node_modules/xlsx');
const fs = require('fs');
const assert = require('assert');

const html = fs.readFileSync('/home/claude/my-game/index.html', 'utf8');
const wb = XLSX.readFile(process.env.HARNESS_XLSX_PATH || '/home/claude/my-game/data/不朽之旅.xlsx');

function sheetRows(name, opts={}) {
  const ws = wb.Sheets[name];
  if (!ws) throw new Error('no sheet ' + name);
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, ...opts });
}

function currentBlock(name) {
  const re = new RegExp(`##AUTO-GENERATED:${name}:START##([\\s\\S]*?)##AUTO-GENERATED:${name}:END##`);
  const m = html.match(re);
  if (!m) throw new Error('no block for ' + name);
  return m[1];
}

function findMatchingLiteral(str, startIdx) {
  // str[startIdx] is the opening '{' or '['. Scan forward respecting string/template
  // literals (with backslash-escapes) to find the matching close, so literals that
  // contain raw ';' or nested brackets inside string values (e.g. inline CSS/HTML) don't
  // truncate the match.
  const open = str[startIdx];
  const close = open === '{' ? '}' : ']';
  let depth = 0;
  let inString = null; // one of "'", '"', '`' or null
  for (let i = startIdx; i < str.length; i++) {
    const c = str[i];
    if (inString) {
      if (c === '\\') { i++; continue; }
      if (c === inString) inString = null;
      continue;
    }
    if (c === '"' || c === "'" || c === '`') { inString = c; continue; }
    if (c === open) depth++;
    else if (c === close) {
      depth--;
      if (depth === 0) return str.slice(startIdx, i + 1);
    }
  }
  throw new Error('unterminated literal starting at ' + startIdx);
}

function currentValue(name) {
  // some consts share one marker region with siblings (e.g. EQUIPMENT/REFINE/COST/BUILD),
  // so search the whole file for "const/let NAME = <literal>" and then bracket-match the
  // literal itself rather than cutting at the first ';' (data literals can contain raw
  // ';' inside string values, e.g. inline CSS/HTML in icon fields).
  const re = new RegExp(`(?:const|let)\\s+${name}\\s*=\\s*`);
  const m = re.exec(html);
  if (!m) throw new Error('no const literal for ' + name);
  const startIdx = m.index + m[0].length;
  const literal = findMatchingLiteral(html, startIdx);
  // eslint-disable-next-line no-eval
  return eval('(' + literal + ')');
}

function check(name, generated) {
  const current = currentValue(name);
  try {
    assert.deepStrictEqual(generated, current);
    console.log('OK  ', name);
    return true;
  } catch (e) {
    console.log('FAIL', name);
    console.log('  expected (current):', JSON.stringify(current).slice(0,500));
    console.log('  got (generated):   ', JSON.stringify(generated).slice(0,500));
    return false;
  }
}

module.exports = { sheetRows, currentBlock, currentValue, check, wb };
