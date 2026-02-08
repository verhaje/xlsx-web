// parser.js - ZIP/XML helpers and cell value utilities
export async function readZipText(zip, path) {
  const entry = zip.file(path);
  if (!entry) {
    throw new Error(`Missing ${path}`);
  }
  return entry.async('text');
}

export function parseXml(xmlString) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlString, 'application/xml');
  const parserError = doc.getElementsByTagName('parsererror');
  if (parserError && parserError.length > 0) {
    throw new Error('Failed to parse XML.');
  }
  return doc;
}

export function buildRelationshipMap(relsDoc) {
  const rels = new Map();
  const relationshipNodes = Array.from(relsDoc.getElementsByTagName('Relationship'));
  relationshipNodes.forEach((rel) => {
    const id = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    if (id && target) rels.set(id, target);
  });
  return rels;
}

export function normalizeTargetPath(target) {
  if (target.startsWith('/')) return target.slice(1);
  if (target.startsWith('xl/')) return target;
  return `xl/${target}`;
}

export async function loadSharedStrings(zip) {
  const shared = zip.file('xl/sharedStrings.xml');
  if (!shared) return [];
  const sharedXml = await shared.async('text');
  const sharedDoc = parseXml(sharedXml);
  const items = Array.from(sharedDoc.getElementsByTagName('si'));
  return items.map((item) => {
    const texts = Array.from(item.getElementsByTagName('t'));
    return texts.map((t) => t.textContent || '').join('');
  });
}

// cell reference utilities
export function parseCellRef(ref) {
  const match = /([A-Z]+)(\d+)/i.exec(ref);
  if (!match) return { col: 0, row: 0 };
  const colLetters = match[1].toUpperCase();
  const row = parseInt(match[2], 10);
  return { col: columnNameToIndex(colLetters), row };
}

export function columnNameToIndex(name) {
  let index = 0;
  for (let i = 0; i < name.length; i += 1) index = index * 26 + (name.charCodeAt(i) - 64);
  return index;
}

export function columnIndexToName(index) {
  let result = '';
  let current = index;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    current = Math.floor((current - 1) / 26);
  }
  return result;
}

export function extractCellValue(cell, sharedStrings) {
  const type = cell.getAttribute('t');
  if (type === 'inlineStr') {
    const inlineText = cell.getElementsByTagName('t')[0];
    return inlineText ? inlineText.textContent || '' : '';
  }
  const valueNode = cell.getElementsByTagName('v')[0];
  if (!valueNode) return '';
  const rawValue = valueNode.textContent || '';
  if (type === 's') {
    const index = parseInt(rawValue, 10);
    return Number.isNaN(index) ? rawValue : sharedStrings[index] || '';
  }
  if (type === 'b') return rawValue === '1' ? 'TRUE' : 'FALSE';
  return rawValue;
}

export function expandRange(range) {
  const [start, end] = range.split(':');
  const s = parseCellRef(start);
  const e = end ? parseCellRef(end) : s;
  const cells = [];
  for (let r = s.row; r <= e.row; r++) for (let c = s.col; c <= e.col; c++) cells.push({ row: r, col: c });
  return cells;
}

export function parseSqref(sqref) {
  const parts = sqref.trim().split(/\s+/);
  const cells = [];
  parts.forEach((p) => { cells.push(...expandRange(p)); });
  return cells;
}

export function parseRangeRef(range) {
  const [startRef, endRef] = range.split(':');
  const start = parseCellRef(startRef);
  const end = endRef ? parseCellRef(endRef) : start;
  return { start, end };
}

export function parseMergeCells(sheetDoc) {
  const mergeCellsNode = sheetDoc.getElementsByTagName('mergeCells')[0];
  if (!mergeCellsNode) return [];
  const mergeNodes = Array.from(mergeCellsNode.getElementsByTagName('mergeCell'));
  const ranges = [];
  mergeNodes.forEach((node) => {
    const sqref = node.getAttribute('ref') || node.getAttribute('sqref');
    if (!sqref) return;
    const parts = sqref.trim().split(/\s+/);
    parts.forEach((part) => {
      const { start, end } = parseRangeRef(part);
      if (!start.row || !start.col || !end.row || !end.col) return;
      ranges.push({ start, end });
    });
  });
  return ranges;
}

export function parseColumnStyles(sheetDoc) {
  const colsNode = sheetDoc.getElementsByTagName('cols')[0];
  if (!colsNode) return [];
  const colNodes = Array.from(colsNode.getElementsByTagName('col'));
  const ranges = [];
  colNodes.forEach((col) => {
    const min = parseInt(col.getAttribute('min') || '0', 10);
    const max = parseInt(col.getAttribute('max') || '0', 10);
    const style = col.getAttribute('style');
    if (!min || !max || style == null) return;
    ranges.push({ min, max, style });
  });
  return ranges;
}

export function buildMergedMaps(mergedRanges) {
  const anchorMap = new Map();
  const coveredMap = new Map();
  mergedRanges.forEach(({ start, end }) => {
    const rowSpan = end.row - start.row + 1;
    const colSpan = end.col - start.col + 1;
    const anchorKey = `${start.row}-${start.col}`;
    anchorMap.set(anchorKey, { rowSpan, colSpan });
    for (let r = start.row; r <= end.row; r += 1) {
      for (let c = start.col; c <= end.col; c += 1) {
        if (r === start.row && c === start.col) continue;
        coveredMap.set(`${r}-${c}`, anchorKey);
      }
    }
  });
  return { anchorMap, coveredMap };
}

export function mapMergedRef(refA1, sheetModel) {
  if (!sheetModel || !sheetModel.coveredMap) return refA1;
  const parsed = parseCellRef(refA1);
  if (!parsed.row || !parsed.col) return refA1;
  const anchorKey = sheetModel.coveredMap.get(`${parsed.row}-${parsed.col}`);
  if (!anchorKey) return refA1;
  const [row, col] = anchorKey.split('-').map(Number);
  return `${columnIndexToName(col)}${row}`;
}

export function parseSharedFormulas(sheetDoc) {
  const sharedFormulas = new Map();
  const sharedCells = new Map();
  const fNodes = Array.from(sheetDoc.getElementsByTagName('f'));
  fNodes.forEach((fNode) => {
    if (fNode.getAttribute('t') !== 'shared') return;
    const si = fNode.getAttribute('si');
    if (!si) return;
    const parent = fNode.parentNode;
    if (!parent || parent.nodeName !== 'c') return;
    const cellRef = parent.getAttribute('r');
    if (!cellRef) return;
    const text = fNode.textContent || '';
    if (text && !sharedFormulas.has(si)) {
      sharedFormulas.set(si, { si, anchorRef: cellRef, baseFormula: text });
    }
    sharedCells.set(cellRef, { si });
  });

  sharedCells.forEach((info, cellRef) => {
    const entry = sharedFormulas.get(info.si);
    if (!entry) return;
    sharedCells.set(cellRef, { ...info, anchorRef: entry.anchorRef, baseFormula: entry.baseFormula });
  });

  return { sharedFormulas, sharedCells };
}

export async function loadSheetRelationships(zip, sheetTarget) {
  const relsPath = sheetTarget.replace('.xml', '.xml.rels');
  try {
    const relsXml = await readZipText(zip, relsPath);
    const relsDoc = parseXml(relsXml);
    return buildRelationshipMap(relsDoc);
  } catch (e) {
    return new Map();
  }
}

export async function loadDrawing(zip, drawingTarget) {
  try {
    const drawingXml = await readZipText(zip, drawingTarget);
    const drawingDoc = parseXml(drawingXml);
    return parseDrawing(drawingDoc);
  } catch (e) {
    return [];
  }
}

function parseDrawing(drawingDoc) {
  const images = [];
  const twoCellAnchors = Array.from(drawingDoc.getElementsByTagName('xdr:twoCellAnchor'));
  const oneCellAnchors = Array.from(drawingDoc.getElementsByTagName('xdr:oneCellAnchor'));

  const parseAnchor = (anchorEl) => {
    const fromEl = anchorEl.getElementsByTagName('xdr:from')[0];
    const toEl = anchorEl.getElementsByTagName('xdr:to')[0];
    const picEl = anchorEl.getElementsByTagName('xdr:pic')[0];
    if (!picEl) return null;

    const blipEl = picEl.getElementsByTagName('a:blip')[0];
    if (!blipEl) return null;

    const embed = blipEl.getAttribute('r:embed');
    if (!embed) return null;

    let from = { col: 0, row: 0 };
    let to = { col: 1, row: 1 };

    if (fromEl) {
      const colEl = fromEl.getElementsByTagName('xdr:col')[0];
      const rowEl = fromEl.getElementsByTagName('xdr:row')[0];
      from.col = parseInt(colEl?.textContent || '0', 10) + 1; // 0-based to 1-based
      from.row = parseInt(rowEl?.textContent || '0', 10) + 1;
    }

    if (toEl) {
      const colEl = toEl.getElementsByTagName('xdr:col')[0];
      const rowEl = toEl.getElementsByTagName('xdr:row')[0];
      to.col = parseInt(colEl?.textContent || '0', 10) + 1;
      to.row = parseInt(rowEl?.textContent || '0', 10) + 1;
    }

    return {
      embed,
      from,
      to,
      type: 'twoCellAnchor'
    };
  };

  twoCellAnchors.forEach(anchor => {
    const image = parseAnchor(anchor);
    if (image) images.push(image);
  });

  oneCellAnchors.forEach(anchor => {
    const image = parseAnchor(anchor);
    if (image) {
      image.type = 'oneCellAnchor';
      images.push(image);
    }
  });

  return images;
}

export async function loadMedia(zip, embedId) {
  try {
    const mediaPath = `xl/media/${embedId}`;
    const mediaFile = zip.file(mediaPath);
    if (!mediaFile) return null;
    const data = await mediaFile.async('base64');
    return `data:image/png;base64,${data}`; // Assume PNG for now, could detect from file extension
  } catch (e) {
    return null;
  }
}

export function evalSimpleFormula(formula) {
  if (!formula) return null;
  const f = formula.trim();
  const cleaned = f.replace(/^\(|\)$/g, '');
  const num = Number(cleaned);
  return Number.isNaN(num) ? null : num;
}
