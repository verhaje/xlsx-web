// renderer.js - DOM rendering helpers
import { parseCellRef, columnIndexToName, extractCellValue, evalSimpleFormula, readZipText, parseXml, parseMergeCells, parseSharedFormulas, buildMergedMaps, mapMergedRef, loadSheetRelationships, loadDrawing, loadMedia, normalizeTargetPath, parseColumnStyles, parseRangeRef, expandRange } from './parser.js';
import { getCellStyle, dxfToCss, applyCssToElement, normalizeRgb } from './styles.js';

function buildSheetModel(sheetDoc) {
  const mergedRanges = parseMergeCells(sheetDoc);
  const { anchorMap, coveredMap } = buildMergedMaps(mergedRanges);
  const { sharedFormulas, sharedCells } = parseSharedFormulas(sheetDoc);
  return { mergedRanges, anchorMap, coveredMap, sharedFormulas, sharedCells };
}


function shiftA1Ref(ref, rowOffset, colOffset) {
  const match = /^(\$?)([A-Z]{1,3})(\$?)(\d+)$/.exec(ref);
  if (!match) return ref;
  const colAbs = match[1] === '$';
  const colLetters = match[2];
  const rowAbs = match[3] === '$';
  const rowNum = parseInt(match[4], 10);
  let colIndex = parseCellRef(`${colLetters}1`).col;
  let rowIndex = rowNum;
  if (!colAbs) colIndex += colOffset;
  if (!rowAbs) rowIndex += rowOffset;
  if (colIndex < 1 || rowIndex < 1) return '#REF!';
  const colName = columnIndexToName(colIndex);
  return `${colAbs ? '$' : ''}${colName}${rowAbs ? '$' : ''}${rowIndex}`;
}

function shiftFormulaRefs(formula, rowOffset, colOffset) {
  if (!rowOffset && !colOffset) return formula;
  const refRegex = /((?:'[^']+'|[A-Za-z0-9_]+)!)?(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?/g;

  const adjustChunk = (chunk) => chunk.replace(refRegex, (match, sheetPrefix, startRef, endRef) => {
    const shiftedStart = shiftA1Ref(startRef, rowOffset, colOffset);
    const shiftedEnd = endRef ? shiftA1Ref(endRef, rowOffset, colOffset) : null;
    return `${sheetPrefix || ''}${shiftedStart}${shiftedEnd ? `:${shiftedEnd}` : ''}`;
  });

  let output = '';
  let buffer = '';
  let inString = false;

  for (let i = 0; i < formula.length; i += 1) {
    const ch = formula[i];
    if (ch === '"') {
      if (inString && formula[i + 1] === '"') {
        output += '""';
        i += 1;
        continue;
      }
      if (!inString) {
        output += adjustChunk(buffer);
        buffer = '';
        inString = true;
        output += '"';
      } else {
        output += '"';
        inString = false;
      }
      continue;
    }
    if (inString) output += ch; else buffer += ch;
  }

  if (buffer) output += adjustChunk(buffer);
  return output;
}

function deriveSharedFormula(baseFormula, anchorRef, targetRef) {
  if (!baseFormula || !anchorRef || !targetRef) return baseFormula || '';
  const anchor = parseCellRef(anchorRef);
  const target = parseCellRef(targetRef);
  if (!anchor.row || !anchor.col || !target.row || !target.col) return baseFormula;
  const rowOffset = target.row - anchor.row;
  const colOffset = target.col - anchor.col;
  return shiftFormulaRefs(baseFormula, rowOffset, colOffset);
}

function getCellFormula(cellEl, cellRef, sheetModel) {
  const fNode = cellEl.getElementsByTagName('f')[0];
  if (!fNode) return null;
  const fType = fNode.getAttribute('t');
  if (fType === 'shared' && sheetModel && sheetModel.sharedCells) {
    const sharedInfo = sheetModel.sharedCells.get(cellRef);
    if (sharedInfo && sharedInfo.baseFormula && sharedInfo.anchorRef) {
      return deriveSharedFormula(sharedInfo.baseFormula, sharedInfo.anchorRef, cellRef);
    }
  }
  return fNode.textContent || '';
}

function getSharedCacheKey(sheetName, sharedInfo, targetRef) {
  if (!sharedInfo || !sharedInfo.anchorRef) return null;
  const anchor = parseCellRef(sharedInfo.anchorRef);
  const target = parseCellRef(targetRef);
  if (!anchor.row || !anchor.col || !target.row || !target.col) return null;
  const rowOffset = target.row - anchor.row;
  const colOffset = target.col - anchor.col;
  return `${sheetName}::${sharedInfo.si}::${rowOffset},${colOffset}`;
}

export function renderTabs(container, sheets, onSelect) {
  container.innerHTML = '';
  sheets.forEach((sheet, index) => {
    const button = document.createElement('button');
    button.className = 'tab-button';
    button.type = 'button';
    button.textContent = sheet.name;
    button.addEventListener('click', async () => { setActiveTab(container, index); await onSelect(sheet); });
    container.appendChild(button);
  });
}

export function setActiveTab(container, index) {
  const buttons = Array.from(container.querySelectorAll('.tab-button'));
  buttons.forEach((button, idx) => { if (idx === index) button.classList.add('active'); else button.classList.remove('active'); });
}

export async function renderSheet({ zip, sheet, sharedStrings, styles, tableContainer, sheetNameEl, formulaEngine, sheets }) {
  if (!sheet.target) return;
  const sheetXml = await readZipText(zip, sheet.target);
  const sheetDoc = parseXml(sheetXml);
  const sheetData = sheetDoc.getElementsByTagName('sheetData')[0];
  const sheetModel = buildSheetModel(sheetDoc);
  const columnStyleRanges = parseColumnStyles(sheetDoc);

  // Load images for this sheet
  const sheetRels = await loadSheetRelationships(zip, sheet.target);
  const images = [];
  for (const [relId, target] of sheetRels) {
    if (target.includes('drawings/')) {
      const drawingImages = await loadDrawing(zip, normalizeTargetPath(target));
      for (const image of drawingImages) {
        const mediaData = await loadMedia(zip, image.embed);
        if (mediaData) {
          images.push({ ...image, dataUrl: mediaData });
        }
      }
    }
  }

  // conditional formatting map
  const cfMap = new Map();
  let cfRuleOrder = 0;
  const parseCfvo = (node) => {
    if (!node) return null;
    return {
      type: node.getAttribute('type') || 'num',
      val: node.getAttribute('val'),
      gte: node.getAttribute('gte') !== '0'
    };
  };

  const parseColorNode = (node) => {
    if (!node) return null;
    return {
      rgb: node.getAttribute('rgb'),
      theme: node.getAttribute('theme'),
      tint: node.getAttribute('tint')
    };
  };

  const cfNodes = Array.from(sheetDoc.getElementsByTagName('conditionalFormatting'));
  cfNodes.forEach((cfNode) => {
    const sqref = cfNode.getAttribute('sqref'); if (!sqref) return;
    const rangeParts = sqref.trim().split(/\s+/);
    const cfRules = Array.from(cfNode.getElementsByTagName('cfRule'));
    cfRules.forEach((rule) => {
      const dxfId = rule.getAttribute('dxfId');
      const type = rule.getAttribute('type');
      const operator = rule.getAttribute('operator');
      const formulaNodes = Array.from(rule.getElementsByTagName('formula'));
      const formula = (formulaNodes[0] || {}).textContent || null;
      const formula1 = (rule.getElementsByTagName('formula1')[0] || {}).textContent || null;
      const formula2 = (rule.getElementsByTagName('formula2')[0] || {}).textContent || null;
      const priorityRaw = rule.getAttribute('priority');
      const priorityParsed = priorityRaw == null ? Number.MAX_SAFE_INTEGER : parseInt(priorityRaw, 10);
      const priority = Number.isNaN(priorityParsed) ? Number.MAX_SAFE_INTEGER : priorityParsed;
      const stopIfTrue = rule.getAttribute('stopIfTrue') === '1' || rule.getAttribute('stopIfTrue') === 'true';
      const text = rule.getAttribute('text');
      const top = rule.getAttribute('top');
      const bottom = rule.getAttribute('bottom');
      const percent = rule.getAttribute('percent');
      const rank = rule.getAttribute('rank');
      const aboveAverage = rule.getAttribute('aboveAverage');
      const equalAverage = rule.getAttribute('equalAverage');
      const colorScaleNode = rule.getElementsByTagName('colorScale')[0];
      const dataBarNode = rule.getElementsByTagName('dataBar')[0];
      const iconSetNode = rule.getElementsByTagName('iconSet')[0];
      const colorScale = colorScaleNode ? {
        cfvos: Array.from(colorScaleNode.getElementsByTagName('cfvo')).map(parseCfvo).filter(Boolean),
        colors: Array.from(colorScaleNode.getElementsByTagName('color')).map(parseColorNode).filter(Boolean)
      } : null;
      const dataBar = dataBarNode ? {
        cfvos: Array.from(dataBarNode.getElementsByTagName('cfvo')).map(parseCfvo).filter(Boolean),
        color: parseColorNode(dataBarNode.getElementsByTagName('color')[0]),
        showValue: dataBarNode.getAttribute('showValue') !== '0'
      } : null;
      const iconSet = iconSetNode ? {
        cfvos: Array.from(iconSetNode.getElementsByTagName('cfvo')).map(parseCfvo).filter(Boolean),
        iconSet: iconSetNode.getAttribute('iconSet') || '3TrafficLights1',
        showValue: iconSetNode.getAttribute('showValue') !== '0'
      } : null;
      let css = null;
      if (dxfId != null && styles.dxfs && styles.dxfs[parseInt(dxfId, 10)]) css = dxfToCss(styles.dxfs[parseInt(dxfId, 10)], styles.themeColors);
      const hasVisualRule = !!(colorScale || dataBar || iconSet);
      if (!css && !hasVisualRule) return;

      const ruleInfo = {
        type,
        operator,
        formula,
        formula1,
        formula2,
        formulas: formulaNodes.map((node) => node.textContent || '').filter(Boolean),
        css,
        priority,
        stopIfTrue,
        text,
        top,
        bottom,
        percent,
        rank,
        aboveAverage,
        equalAverage,
        colorScale,
        dataBar,
        iconSet,
        order: cfRuleOrder,
        targets: new Map()
      };
      cfRuleOrder += 1;

      rangeParts.forEach((part) => {
        const { start, end } = parseRangeRef(part);
        if (!start.row || !start.col || !end.row || !end.col) return;
        const anchor = start;
        const targetCells = expandRange(part);
        targetCells.forEach(({ row, col }) => {
          const key = `${row}-${col}`;
          if (!ruleInfo.targets.has(key)) ruleInfo.targets.set(key, { row, col, key });
          if (!cfMap.has(key)) cfMap.set(key, []);
          cfMap.get(key).push({ rule: ruleInfo, anchor });
        });
      });
    });
  });

  cfMap.forEach((entries) => {
    entries.sort((a, b) => {
      if (a.rule.priority !== b.rule.priority) return a.rule.priority - b.rule.priority;
      return a.rule.order - b.rule.order;
    });
  });

  if (!sheetData) { tableContainer.innerHTML = '<div class="status">No data found in sheet.</div>'; return; }

  const rows = Array.from(sheetData.getElementsByTagName('row'));
  const cellMap = new Map();
  const rowStyleMap = new Map();
  let maxRow = 0; let maxCol = 0;
  // build cell map and extents
  rows.forEach((row) => {
    const rowIndex = parseInt(row.getAttribute('r') || '0', 10);
    const actualRow = Number.isNaN(rowIndex) || rowIndex === 0 ? maxRow + 1 : rowIndex;
    maxRow = Math.max(maxRow, actualRow);
    const rowStyle = row.getAttribute('s');
    if (rowStyle != null) rowStyleMap.set(actualRow, rowStyle);
    const cells = Array.from(row.getElementsByTagName('c'));
    cells.forEach((cell) => {
      const ref = cell.getAttribute('r') || '';
      const { col } = parseCellRef(ref);
      if (col === 0) return;
      maxCol = Math.max(maxCol, col);
      const key = `${actualRow}-${col}`;
      cellMap.set(key, cell);
    });
  });

    // evaluation cache across sheets: key = `${sheetName}::${row}-${col}`
    const sheetCache = new Map();
    // Promise-coalescing cache: prevents thundering herd when multiple
    // concurrent async calls request the same sheet before the first completes.
    const _buildingSheets = new Map();

    // store current sheet map in cache
    sheetCache.set(sheet.name, { cellMap, sheetDoc, maxRow, maxCol, sheetModel });

    async function buildCellMapForSheet(sheetObj) {
      if (!sheetObj || !sheetObj.target) return { cellMap: new Map(), sheetDoc: null };
      if (sheetCache.has(sheetObj.name)) return sheetCache.get(sheetObj.name);
      // Prevent thundering herd: if already building this sheet, reuse the promise
      if (_buildingSheets.has(sheetObj.name)) return _buildingSheets.get(sheetObj.name);
      const promise = (async () => {
        try {
          const xml = await readZipText(zip, sheetObj.target);
          const doc = parseXml(xml);
          const data = doc.getElementsByTagName('sheetData')[0];
          const rows = data ? Array.from(data.getElementsByTagName('row')) : [];
          const cmap = new Map();
          let mRow = 0; let mCol = 0;
          rows.forEach((r) => {
            const rowIndex = parseInt(r.getAttribute('r') || '0', 10);
            const actualRow = Number.isNaN(rowIndex) || rowIndex === 0 ? mRow + 1 : rowIndex;
            mRow = Math.max(mRow, actualRow);
            const cells = Array.from(r.getElementsByTagName('c'));
            cells.forEach((cell) => {
              const ref = cell.getAttribute('r') || '';
              const { col } = parseCellRef(ref);
              if (col === 0) return;
              mCol = Math.max(mCol, col);
              const key = `${actualRow}-${col}`;
              cmap.set(key, cell);
            });
          });
          const model = buildSheetModel(doc);
          const info = { cellMap: cmap, sheetDoc: doc, maxRow: mRow, maxCol: mCol, sheetModel: model };
          sheetCache.set(sheetObj.name, info);
          return info;
        } catch (e) {
          return { cellMap: new Map(), sheetDoc: null };
        } finally {
          _buildingSheets.delete(sheetObj.name);
        }
      })();
      _buildingSheets.set(sheetObj.name, promise);
      return promise;
    }

    // Track cells currently being evaluated to detect circular references
    const evaluatingCells = new Set();
    // Promise-coalescing cache for cell evaluation: prevents duplicate
    // evaluation when multiple formulas reference the same cell concurrently.
    const _pendingCells = new Map();

    function _stripSheetQuotes(name) {
      if (name.length >= 2 && name.startsWith("'") && name.endsWith("'")) return name.slice(1, -1);
      return name;
    }

    async function evaluateCellByRef(refA1, contextSheetName) {
      // handle cross-sheet refs like SheetName!A1
      // contextSheetName is the sheet the calling formula lives on;
      // non-prefixed refs resolve on that sheet (defaults to the rendered sheet).
      let targetSheetName = contextSheetName || sheet.name;
      let a1 = refA1;
      if (typeof refA1 === 'string' && refA1.includes('!')) {
        const parts = refA1.split('!');
        // join all but last in case sheet name contains ! (unlikely)
        targetSheetName = parts.slice(0, -1).join('!');
        a1 = parts[parts.length - 1];
      }
      targetSheetName = _stripSheetQuotes(targetSheetName);

      // get or build cellMap for target sheet
      let targetSheetObj = null;
      if (targetSheetName === sheet.name) targetSheetObj = sheet; else {
        if (Array.isArray(sheets)) targetSheetObj = sheets.find(s => s.name === targetSheetName);
      }

      const info = targetSheetObj ? await buildCellMapForSheet(targetSheetObj) : { cellMap: new Map(), sheetDoc: null, sheetModel: null };
      a1 = mapMergedRef(a1, info.sheetModel);
      const parsed = parseCellRef(a1);
      if (!parsed.col || !parsed.row) return '';
      const key = `${targetSheetName}::${parsed.row}-${parsed.col}`;
      // simple per-call cache
      if (sheetCache.has('__values') && sheetCache.get('__values').has(key)) return sheetCache.get('__values').get(key);
      // Prevent thundering herd: if this cell is already being evaluated concurrently, reuse the promise
      if (_pendingCells.has(key)) return _pendingCells.get(key);
      // ensure values map
      if (!sheetCache.has('__values')) sheetCache.set('__values', new Map());

      // Circular reference guard: if this cell is already being evaluated, return error
      if (evaluatingCells.has(key)) return '#REF!';
      evaluatingCells.add(key);

      const promise = (async () => {
        const cmap = info.cellMap || new Map();
        const cellEl = cmap.get(`${parsed.row}-${parsed.col}`);
        if (!cellEl) { sheetCache.get('__values').set(key, ''); return ''; }

        let val;
        const cellRef = `${columnIndexToName(parsed.col)}${parsed.row}`;
        const formulaText = getCellFormula(cellEl, cellRef, info.sheetModel);
        if (formulaText !== null && formulaEngine) {
          const sharedInfo = info.sheetModel?.sharedCells?.get(cellRef) || null;
          const cacheKey = getSharedCacheKey(targetSheetName, sharedInfo, cellRef);
          if (cacheKey && sheetCache.get('__sharedValues')?.has(cacheKey)) {
            val = sheetCache.get('__sharedValues').get(cacheKey);
          } else {
            // Pass targetSheetName so non-prefixed refs in the formula resolve on the correct sheet
            val = await formulaEngine.evaluateFormula(formulaText, { resolveCell: async (r) => await evaluateCellByRef(r, targetSheetName), sharedStrings, zip, sheetDoc: info.sheetDoc });
            if (cacheKey) {
              if (!sheetCache.has('__sharedValues')) sheetCache.set('__sharedValues', new Map());
              sheetCache.get('__sharedValues').set(cacheKey, val);
            }
          }
        } else {
          val = extractCellValue(cellEl, sharedStrings);
        }
        sheetCache.get('__values').set(key, val);
        return val;
      })();

      _pendingCells.set(key, promise);
      try {
        return await promise;
      } finally {
        evaluatingCells.delete(key);
        _pendingCells.delete(key);
      }
    }

    function normalizeCfFormula(raw) {
      if (raw == null) return null;
      const text = String(raw).trim();
      if (!text) return null;
      return text.replace(/^=/, '');
    }

    function isTruthyCfValue(value) {
      if (value == null) return false;
      if (typeof value === 'boolean') return value;
      if (typeof value === 'number') return value !== 0;
      const text = String(value).trim();
      if (!text) return false;
      const upper = text.toUpperCase();
      if (upper === 'TRUE') return true;
      if (upper === 'FALSE') return false;
      const num = Number(text);
      if (!Number.isNaN(num)) return num !== 0;
      return true;
    }

    async function getRuleTargetCells(rule) {
      if (!rule || !rule.targets) return [];
      if (rule._targetCells) return rule._targetCells;
      const cells = Array.from(rule.targets.values());
      rule._targetCells = cells;
      return cells;
    }

    async function getRuleNumericValues(rule) {
      if (!rule) return [];
      if (rule._numericValues) return rule._numericValues;
      const cells = await getRuleTargetCells(rule);
      const numeric = [];
      for (const cell of cells) {
        const a1 = `${columnIndexToName(cell.col)}${cell.row}`;
        const val = await evaluateCellByRef(a1);
        const num = Number(val);
        if (!Number.isNaN(num)) numeric.push(num);
      }
      rule._numericValues = numeric;
      return numeric;
    }

    function resolveThemeColor(color, themeColors) {
      if (!color) return null;
      if (color.rgb) return normalizeRgb(color.rgb) || null;
      if (color.theme != null) {
        const idx = parseInt(color.theme, 10);
        if (!Number.isNaN(idx) && themeColors && themeColors[idx]) return themeColors[idx];
      }
      return null;
    }

    function clamp(value, min, max) {
      if (value < min) return min;
      if (value > max) return max;
      return value;
    }

    function interpolateColor(hexA, hexB, t) {
      if (!hexA || !hexB) return null;
      const parse = (hex) => {
        const cleaned = hex.replace('#', '');
        return [
          parseInt(cleaned.slice(0, 2), 16),
          parseInt(cleaned.slice(2, 4), 16),
          parseInt(cleaned.slice(4, 6), 16)
        ];
      };
      const [r1, g1, b1] = parse(hexA);
      const [r2, g2, b2] = parse(hexB);
      const r = Math.round(r1 + (r2 - r1) * t);
      const g = Math.round(g1 + (g2 - g1) * t);
      const b = Math.round(b1 + (b2 - b1) * t);
      return `#${[r, g, b].map(v => v.toString(16).padStart(2, '0')).join('')}`;
    }

    async function resolveCfvoValue(cfvo, rule, row, col, anchor) {
      if (!cfvo) return null;
      const type = cfvo.type || 'num';
      const numeric = await getRuleNumericValues(rule);
      const min = numeric.length ? Math.min(...numeric) : 0;
      const max = numeric.length ? Math.max(...numeric) : 0;
      if (type === 'min') return min;
      if (type === 'max') return max;
      if (type === 'percent') {
        const p = Number(cfvo.val || 0) / 100;
        return min + (max - min) * p;
      }
      if (type === 'percentile') {
        const p = Number(cfvo.val || 0) / 100;
        if (!numeric.length) return 0;
        const sorted = [...numeric].sort((a, b) => a - b);
        const idx = clamp(Math.ceil((sorted.length - 1) * p), 0, sorted.length - 1);
        return sorted[idx];
      }
      if (type === 'formula') return resolveCfValue(cfvo.val, row, col, anchor);
      return Number(cfvo.val);
    }

    async function resolveCfvoValues(cfvos, rule, row, col, anchor) {
      const values = [];
      for (const cfvo of cfvos) {
        const v = await resolveCfvoValue(cfvo, rule, row, col, anchor);
        values.push(Number(v));
      }
      return values;
    }

    function getIconSetMeta(name) {
      const map = {
        '3Arrows': [
          { icon: '▼', color: '#d93025' },
          { icon: '▶', color: '#f29900' },
          { icon: '▲', color: '#188038' }
        ],
        '3ArrowsGray': [
          { icon: '▼', color: '#6e6e6e' },
          { icon: '▶', color: '#9e9e9e' },
          { icon: '▲', color: '#4a4a4a' }
        ],
        '3TrafficLights1': [
          { icon: '●', color: '#d93025' },
          { icon: '●', color: '#f29900' },
          { icon: '●', color: '#188038' }
        ],
        '3TrafficLights2': [
          { icon: '●', color: '#d93025' },
          { icon: '●', color: '#f29900' },
          { icon: '●', color: '#188038' }
        ],
        '3Signs': [
          { icon: '■', color: '#d93025' },
          { icon: '■', color: '#f29900' },
          { icon: '■', color: '#188038' }
        ],
        '3Symbols': [
          { icon: '✖', color: '#d93025' },
          { icon: '●', color: '#f29900' },
          { icon: '✔', color: '#188038' }
        ],
        '3Symbols2': [
          { icon: '✖', color: '#d93025' },
          { icon: '●', color: '#f29900' },
          { icon: '✔', color: '#188038' }
        ]
      };
      return map[name] || map['3TrafficLights1'];
    }

    async function evaluateColorScale(rule, cellValue, row, col, anchor) {
      if (!rule.colorScale) return null;
      const num = Number(cellValue);
      if (Number.isNaN(num)) return null;
      const colors = rule.colorScale.colors || [];
      const cfvos = rule.colorScale.cfvos || [];
      if (colors.length < 2 || cfvos.length < 2) return null;
      const values = await resolveCfvoValues(cfvos, rule, row, col, anchor);
      if (values.some(v => Number.isNaN(v))) return null;
      const resolvedColors = colors.map(c => resolveThemeColor(c, styles.themeColors)).filter(Boolean);
      if (resolvedColors.length < 2) return null;
      const min = values[0];
      const max = values[values.length - 1];
      if (max === min) return { backgroundColor: resolvedColors[resolvedColors.length - 1] };
      let color = resolvedColors[resolvedColors.length - 1];
      if (values.length >= 3 && resolvedColors.length >= 3) {
        const mid = values[1];
        if (num <= mid) {
          const t = clamp((num - min) / (mid - min), 0, 1);
          color = interpolateColor(resolvedColors[0], resolvedColors[1], t) || resolvedColors[1];
        } else {
          const t = clamp((num - mid) / (max - mid), 0, 1);
          color = interpolateColor(resolvedColors[1], resolvedColors[2], t) || resolvedColors[2];
        }
      } else {
        const t = clamp((num - min) / (max - min), 0, 1);
        color = interpolateColor(resolvedColors[0], resolvedColors[1], t) || resolvedColors[1];
      }
      return { backgroundColor: color };
    }

    async function evaluateDataBar(rule, cellValue, row, col, anchor) {
      if (!rule.dataBar) return null;
      const num = Number(cellValue);
      if (Number.isNaN(num)) return null;
      const cfvos = rule.dataBar.cfvos || [];
      if (cfvos.length < 2) return null;
      const values = await resolveCfvoValues(cfvos, rule, row, col, anchor);
      if (values.some(v => Number.isNaN(v))) return null;
      const min = values[0];
      const max = values[values.length - 1];
      const color = resolveThemeColor(rule.dataBar.color, styles.themeColors) || '#638ec6';
      const percent = max === min ? 100 : clamp(((num - min) / (max - min)) * 100, 0, 100);
      return {
        css: {
          backgroundImage: `linear-gradient(90deg, ${color} ${percent}%, transparent ${percent}%)`,
          backgroundRepeat: 'no-repeat',
          backgroundSize: '100% 100%'
        },
        hideValue: !rule.dataBar.showValue
      };
    }

    async function evaluateIconSet(rule, cellValue, row, col, anchor) {
      if (!rule.iconSet) return null;
      const num = Number(cellValue);
      if (Number.isNaN(num)) return null;
      const cfvos = rule.iconSet.cfvos || [];
      if (!cfvos.length) return null;
      const values = await resolveCfvoValues(cfvos, rule, row, col, anchor);
      if (values.some(v => Number.isNaN(v))) return null;
      const icons = getIconSetMeta(rule.iconSet.iconSet);
      let idx = 0;
      for (let i = values.length - 1; i >= 0; i -= 1) {
        const gte = cfvos[i]?.gte !== false;
        if (gte ? num >= values[i] : num > values[i]) {
          idx = i;
          break;
        }
      }
      const icon = icons[Math.min(idx, icons.length - 1)] || icons[0];
      return {
        icon,
        hideValue: !rule.iconSet.showValue
      };
    }

    async function resolveCfValue(raw, row, col, anchor) {
      const normalized = normalizeCfFormula(raw);
      if (!normalized) return null;
      const simple = evalSimpleFormula(normalized);
      if (simple != null) return simple;
      if (!formulaEngine) return simple;
      const rowOffset = row - anchor.row;
      const colOffset = col - anchor.col;
      const shifted = shiftFormulaRefs(normalized, rowOffset, colOffset);
      return formulaEngine.evaluateFormula(shifted, { resolveCell: async (r) => await evaluateCellByRef(r), sharedStrings, zip, sheetDoc });
    }

    async function evaluateConditionalFormatting(entry, cellValue, row, col) {
      const { rule, anchor } = entry;
      if (!rule) return { matched: false };
      if (rule.type === 'cellIs') {
        const v1 = await resolveCfValue(rule.formula1 || rule.formula, row, col, anchor);
        const v2 = await resolveCfValue(rule.formula2, row, col, anchor);
        const num = Number(cellValue);
        const n1 = Number(v1);
        const n2 = Number(v2);
        if (Number.isNaN(num) || Number.isNaN(n1)) {
          const s = String(cellValue ?? '');
          const s1 = String(v1 ?? '');
          if (rule.operator === 'equal') return { matched: s === s1, css: rule.css };
          if (rule.operator === 'notEqual') return { matched: s !== s1, css: rule.css };
          return { matched: false };
        }
        switch (rule.operator) {
          case 'greaterThan': return { matched: num > n1, css: rule.css };
          case 'lessThan': return { matched: num < n1, css: rule.css };
          case 'greaterThanOrEqual': return { matched: num >= n1, css: rule.css };
          case 'lessThanOrEqual': return { matched: num <= n1, css: rule.css };
          case 'equal': return { matched: num === n1, css: rule.css };
          case 'notEqual': return { matched: num !== n1, css: rule.css };
          case 'between': return { matched: !Number.isNaN(n2) && num >= n1 && num <= n2, css: rule.css };
          default: return { matched: false };
        }
      }
      if (rule.type === 'containsText' || rule.type === 'notContainsText' || rule.type === 'beginsWith' || rule.type === 'endsWith') {
        const txt = rule.text != null ? rule.text : (rule.formula || '');
        const valueText = String(cellValue ?? '');
        let matched = false;
        if (rule.type === 'beginsWith') matched = valueText.startsWith(txt);
        else if (rule.type === 'endsWith') matched = valueText.endsWith(txt);
        else matched = valueText.indexOf(txt) !== -1;
        return { matched: rule.type === 'notContainsText' ? !matched : matched, css: rule.css };
      }
      if (rule.type === 'expression') {
        const expr = rule.formula || rule.formula1 || (rule.formulas && rule.formulas[0]) || null;
        const result = await resolveCfValue(expr, row, col, anchor);
        return { matched: isTruthyCfValue(result), css: rule.css };
      }
      if (rule.type === 'top10') {
        const numeric = await getRuleNumericValues(rule);
        if (!numeric.length) return { matched: false };
        const rankRaw = parseInt(rule.rank || '10', 10);
        const rank = Number.isNaN(rankRaw) ? 10 : rankRaw;
        const isPercent = rule.percent === '1';
        const isBottom = rule.bottom === '1' || rule.top === '0';
        let count = isPercent ? Math.ceil((numeric.length * rank) / 100) : rank;
        count = Math.max(1, Math.min(count, numeric.length));
        const sorted = [...numeric].sort((a, b) => (isBottom ? a - b : b - a));
        const threshold = sorted[count - 1];
        const num = Number(cellValue);
        if (Number.isNaN(num)) return { matched: false };
        return { matched: isBottom ? num <= threshold : num >= threshold, css: rule.css };
      }
      if (rule.type === 'aboveAverage') {
        const numeric = await getRuleNumericValues(rule);
        if (!numeric.length) return { matched: false };
        const avg = numeric.reduce((sum, v) => sum + v, 0) / numeric.length;
        const num = Number(cellValue);
        if (Number.isNaN(num)) return { matched: false };
        const above = rule.aboveAverage !== '0';
        const equal = rule.equalAverage === '1';
        if (above) return { matched: equal ? num >= avg : num > avg, css: rule.css };
        return { matched: equal ? num <= avg : num < avg, css: rule.css };
      }
      if (rule.type === 'colorScale') {
        const css = await evaluateColorScale(rule, cellValue, row, col, anchor);
        return css ? { matched: true, css } : { matched: false };
      }
      if (rule.type === 'dataBar') {
        const result = await evaluateDataBar(rule, cellValue, row, col, anchor);
        return result ? { matched: true, css: result.css, hideValue: result.hideValue } : { matched: false };
      }
      if (rule.type === 'iconSet') {
        const result = await evaluateIconSet(rule, cellValue, row, col, anchor);
        return result ? { matched: true, icon: result.icon, hideValue: result.hideValue } : { matched: false };
      }
      return { matched: false };
    }

  

  const colStyleByIndex = [];
  if (columnStyleRanges.length) {
    columnStyleRanges.forEach(({ min, max, style }) => {
      for (let c = min; c <= max; c += 1) colStyleByIndex[c] = style;
    });
  }

  const table = document.createElement('table'); table.className = 'sheet-table';
  // create colgroup for explicit column widths and easier resizing
  const colgroup = document.createElement('colgroup');
  for (let c = 1; c <= maxCol; c += 1) {
    const colEl = document.createElement('col');
    colEl.dataset.colIndex = String(c);
    colgroup.appendChild(colEl);
  }
  table.appendChild(colgroup);
  const thead = document.createElement('thead'); const headerRow = document.createElement('tr');
  const corner = document.createElement('th'); corner.textContent = ''; headerRow.appendChild(corner);
  for (let col = 1; col <= maxCol; col += 1) {
    const th = document.createElement('th');
    th.textContent = columnIndexToName(col);
    th.dataset.colIndex = String(col);
    headerRow.appendChild(th);
  }
  thead.appendChild(headerRow); table.appendChild(thead);

  const tbody = document.createElement('tbody');

  // yield helper to keep UI responsive during large renders
  const yieldToEventLoop = () => new Promise((res) => setTimeout(res, 0));
  const batchSize = 50; // number of rows per chunk (larger batches to reduce yield overhead)

  // --- Show a lightweight loading overlay while rendering ---
  const loadingOverlay = document.createElement('div');
  loadingOverlay.className = 'sheet-loading-overlay';
  loadingOverlay.innerHTML = '<div class="sheet-loading-spinner"></div><div class="sheet-loading-text">Rendering sheet\u2026</div>';
  tableContainer.innerHTML = '';
  tableContainer.appendChild(loadingOverlay);
  // yield so the browser paints the overlay before heavy work begins
  await yieldToEventLoop();

  // --- Helper: applies value + style + conditional formatting to a td ---
  function applyCellContent(td, value, cellStyle, cellEl, key) {
    const anchorSpan = sheetModel.anchorMap.get(key);
    if (anchorSpan) {
      if (anchorSpan.colSpan > 1) td.colSpan = anchorSpan.colSpan;
      if (anchorSpan.rowSpan > 1) td.rowSpan = anchorSpan.rowSpan;
    }
    if (cellStyle && cellStyle.type === 'runs') {
      if (cellStyle.baseline) applyCssToElement(td, cellStyle.baseline);
      if (value === undefined || value === null || value === '') {
        td.textContent = '';
        td.classList.add('empty-cell');
      } else {
        cellStyle.runs.forEach((run) => {
          const span = document.createElement('span');
          span.textContent = run.text;
          applyCssToElement(span, run.style);
          td.appendChild(span);
        });
      }
    } else {
      if (value === undefined || value === null || value === '') {
        td.textContent = '';
        td.classList.add('empty-cell');
      } else {
        td.textContent = value;
      }
      if (cellStyle && cellStyle.style) applyCssToElement(td, cellStyle.style);
    }
  }

  async function applyCfToTd(td, cfEntries, value, row, col) {
    if (!cfEntries || !cfEntries.length) return;
    for (const entry of cfEntries) {
      const result = await evaluateConditionalFormatting(entry, value, row, col);
      if (result.matched) {
        if (result.css) applyCssToElement(td, result.css);
        if (result.icon) {
          const existing = td.querySelector('.cf-icon');
          if (existing) existing.remove();
          const iconSpan = document.createElement('span');
          iconSpan.className = 'cf-icon';
          iconSpan.textContent = result.icon.icon;
          iconSpan.style.marginRight = '4px';
          iconSpan.style.fontSize = '12px';
          iconSpan.style.color = result.icon.color || '#000';
          iconSpan.style.verticalAlign = 'middle';
          td.insertBefore(iconSpan, td.firstChild);
        }
        if (result.hideValue) {
          td.style.color = 'transparent';
          const children = Array.from(td.querySelectorAll('span'));
          children.forEach((span) => {
            if (!span.classList.contains('cf-icon')) span.style.color = 'transparent';
          });
        }
        if (entry.rule.stopIfTrue) break;
      }
    }
  }

  for (let startRow = 1; startRow <= maxRow; startRow += batchSize) {
    const endRow = Math.min(maxRow, startRow + batchSize - 1);
    const frag = document.createDocumentFragment();
    // Collect deferred formula work items for this batch
    const deferredWork = [];

    for (let row = startRow; row <= endRow; row += 1) {
      const tr = document.createElement('tr');
      const rowHeader = document.createElement('th');
      rowHeader.textContent = row.toString();
      tr.appendChild(rowHeader);
      rowHeader.dataset.rowIndex = String(row);

      for (let col = 1; col <= maxCol; col += 1) {
        const key = `${row}-${col}`;
        if (sheetModel.coveredMap.has(key)) continue;
        const td = document.createElement('td');
        td.dataset.rowIndex = String(row);
        td.dataset.colIndex = String(col);
        const cellEl = cellMap.get(key);
        let value = '';
        let cellStyle = null;
        const rowStyleIndex = rowStyleMap.get(row);
        const colStyleIndex = colStyleByIndex[col];
        const fallbackStyleIndex = rowStyleIndex != null ? rowStyleIndex : colStyleIndex;
        let needsFormulaEval = false;
        if (cellEl) {
          const cellRef = `${columnIndexToName(col)}${row}`;
          const formulaText = getCellFormula(cellEl, cellRef, sheetModel);
          if (formulaText !== null && formulaEngine) {
            // Check shared-formula cache synchronously first
            const sharedInfo = sheetModel.sharedCells.get(cellRef) || null;
            const cacheKey = getSharedCacheKey(sheet.name, sharedInfo, cellRef);
            if (cacheKey && sheetCache.get('__sharedValues')?.has(cacheKey)) {
              value = sheetCache.get('__sharedValues').get(cacheKey);
            } else {
              // Defer formula evaluation to parallel batch below
              needsFormulaEval = true;
              deferredWork.push({ td, row, col, key, cellEl, cellRef, formulaText, cacheKey, fallbackStyleIndex });
            }
            td.dataset.formula = formulaText;
          } else {
            value = extractCellValue(cellEl, sharedStrings);
          }
          td.dataset.value = value;
          cellStyle = getCellStyle(cellEl, styles.cellXfs, styles.fonts, styles.themeColors, styles.fills, styles.borders, styles.cellStyleXfs, fallbackStyleIndex);
        } else if (fallbackStyleIndex != null) {
          cellStyle = getCellStyle(null, styles.cellXfs, styles.fonts, styles.themeColors, styles.fills, styles.borders, styles.cellStyleXfs, fallbackStyleIndex);
        }

        if (!needsFormulaEval) {
          // Apply content immediately for non-formula cells
          applyCellContent(td, value, cellStyle, cellEl, key);
          // apply conditional formatting (only if not deferred)
          const cfEntries = cfMap.get(key) || [];
          if (cfEntries.length) {
            deferredWork.push({ td, row, col, key, cfOnly: true, value, cfEntries });
          }
        }
        tr.appendChild(td);
      }
      frag.appendChild(tr);
    }

    // --- Evaluate all deferred formulas in this batch concurrently ---
    if (deferredWork.length > 0) {
      await Promise.all(deferredWork.map(async (item) => {
        if (item.cfOnly) {
          // Only conditional formatting needed
          await applyCfToTd(item.td, item.cfEntries, item.value, item.row, item.col);
          return;
        }
        const { td, row, col, key, cellEl, cellRef, formulaText, cacheKey, fallbackStyleIndex } = item;
        let value;
        // Re-check shared cache (might have been populated by a parallel eval)
        if (cacheKey && sheetCache.get('__sharedValues')?.has(cacheKey)) {
          value = sheetCache.get('__sharedValues').get(cacheKey);
        } else {
          value = await formulaEngine.evaluateFormula(formulaText, { resolveCell: async (r) => await evaluateCellByRef(r), sharedStrings, zip, sheetDoc });
          if (cacheKey) {
            if (!sheetCache.has('__sharedValues')) sheetCache.set('__sharedValues', new Map());
            sheetCache.get('__sharedValues').set(cacheKey, value);
          }
        }
        td.dataset.value = value;
        const cellStyle = getCellStyle(cellEl, styles.cellXfs, styles.fonts, styles.themeColors, styles.fills, styles.borders, styles.cellStyleXfs, fallbackStyleIndex);
        applyCellContent(td, value, cellStyle, cellEl, key);
        // apply conditional formatting
        const cfEntries = cfMap.get(key) || [];
        if (cfEntries.length) {
          await applyCfToTd(td, cfEntries, value, row, col);
        }
      }));
    }

    tbody.appendChild(frag);
    // yield to the event loop between batches so the UI stays responsive
    await yieldToEventLoop();
  }
  table.appendChild(tbody);

  sheetNameEl.textContent = sheet.name;
  // Remove loading overlay and show the rendered table
  tableContainer.innerHTML = ''; tableContainer.appendChild(table);

  // Render images over the table
  if (images.length > 0) {
    const imageContainer = document.createElement('div');
    imageContainer.className = 'image-overlay';
    imageContainer.style.position = 'absolute';
    imageContainer.style.top = '0';
    imageContainer.style.left = '0';
    imageContainer.style.pointerEvents = 'none';
    imageContainer.style.zIndex = '10';

    images.forEach(image => {
      const img = document.createElement('img');
      img.src = image.dataUrl;
      img.style.position = 'absolute';
      img.style.pointerEvents = 'auto';

      // Calculate position based on cell coordinates
      // This is a simplified positioning - in a real implementation you'd need to
      // calculate pixel positions based on column widths and row heights
      const startCol = image.from.col;
      const startRow = image.from.row;
      const endCol = image.to.col;
      const endRow = image.to.row;

      // For now, position images relative to the table cells
      // This assumes default column width of ~64px and row height of ~20px
      const colWidth = 64;
      const rowHeight = 20;

      const left = (startCol - 1) * colWidth;
      const top = (startRow - 1) * rowHeight;
      const width = (endCol - startCol + 1) * colWidth;
      const height = (endRow - startRow + 1) * rowHeight;

      img.style.left = `${left}px`;
      img.style.top = `${top}px`;
      img.style.width = `${width}px`;
      img.style.height = `${height}px`;
      img.style.objectFit = 'contain';

      imageContainer.appendChild(img);
    });

    // Position the image container over the table
    tableContainer.style.position = 'relative';
    tableContainer.appendChild(imageContainer);
  }

  // return the created table element so callers can attach resizers/handlers
  return table;
}
