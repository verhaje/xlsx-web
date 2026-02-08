// Helpers to parse and expand A1-style ranges
export function parseRef(ref) {
  if (!ref) return null;
  const m = /^([A-Z]+)(\d+)$/i.exec(ref);
  if (!m) return null;
  const col = m[1].toUpperCase();
  let colNum = 0;
  for (let i = 0; i < col.length; i++) colNum = colNum * 26 + (col.charCodeAt(i) - 64);
  const row = parseInt(m[2], 10);
  return { col: colNum, row };
}

export function expandRange(range) {
  if (!range) return [];
  const parts = range.split(':');
  if (parts.length === 1) return [parts[0]];
  const start = parts[0];
  const end = parts[1];
  const s = parseRef(start);
  const e = parseRef(end);
  if (!s || !e) return [start];
  const refs = [];
  for (let r = s.row; r <= e.row; r++) {
    for (let c = s.col; c <= e.col; c++) {
      let colStr = '';
      let cc = c;
      while (cc > 0) { const rem = (cc - 1) % 26; colStr = String.fromCharCode(65 + rem) + colStr; cc = Math.floor((cc - 1) / 26); }
      refs.push(`${colStr}${r}`);
    }
  }
  return refs;
}

export function getRangeDimensions(rangeStr) {
  const parts = rangeStr.split(':');
  if (parts.length === 1) return { rows: 1, cols: 1 };
  const s = parseRef(parts[0]);
  const e = parseRef(parts[1]);
  if (!s || !e) return { rows: 1, cols: 1 };
  return {
    rows: e.row - s.row + 1,
    cols: e.col - s.col + 1
  };
}
