// XlsxWriter.ts - XLSX save / export engine
//
// Supports two modes:
// 1. Save existing workbook: clones the original ZIP, merges CellStore
//    changes (values/formulas) into existing sheet XML while preserving
//    all other setup (styles, formatting, drawings, tables, etc.).
// 2. Create new workbook: builds a minimal valid XLSX from scratch.

import { CellReference } from '../core/CellReference';
import { XmlParser } from '../core/XmlParser';
import { CellStore, CellData } from '../data/CellStore';
import type { SheetInfo } from '../types';

declare const JSZip: any;
type JSZipInstance = any;

// ---- XML escaping helper ----

function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ---- Shared strings builder ----

interface SharedStringTable {
  /** Map from string value to index */
  map: Map<string, number>;
  /** Ordered list of unique strings */
  strings: string[];
}

function buildSharedStrings(stores: Map<string, CellStore>): SharedStringTable {
  const map = new Map<string, number>();
  const strings: string[] = [];

  for (const [, store] of stores) {
    store.forEach((_key, data) => {
      if (data.type === 'string' && data.value !== '' && data.value != null) {
        const v = String(data.value);
        if (!map.has(v)) {
          map.set(v, strings.length);
          strings.push(v);
        }
      }
    });
  }

  return { map, strings };
}

function serializeSharedStrings(sst: SharedStringTable): string {
  const items = sst.strings
    .map((s) => `<si><t>${escapeXml(s)}</t></si>`)
    .join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sst.strings.length}" uniqueCount="${sst.strings.length}">${items}</sst>`;
}

// ---- Cell XML serialization ----

function cellToXml(row: number, col: number, data: CellData, sst: SharedStringTable): string {
  const ref = CellReference.toA1(row, col);
  const styleAttr = data.styleIndex !== undefined ? ` s="${data.styleIndex}"` : '';

  if (data.formula) {
    const val = data.value != null && data.value !== '' ? `<v>${escapeXml(String(data.value))}</v>` : '';
    return `<c r="${ref}"${styleAttr}><f>${escapeXml(data.formula)}</f>${val}</c>`;
  }

  if (data.type === 'string' && data.value !== '' && data.value != null) {
    const idx = sst.map.get(String(data.value));
    if (idx !== undefined) {
      return `<c r="${ref}"${styleAttr} t="s"><v>${idx}</v></c>`;
    }
    // Fallback to inline string
    return `<c r="${ref}"${styleAttr} t="inlineStr"><is><t>${escapeXml(String(data.value))}</t></is></c>`;
  }

  if (data.type === 'boolean') {
    return `<c r="${ref}"${styleAttr} t="b"><v>${data.value ? '1' : '0'}</v></c>`;
  }

  if (data.type === 'number' || data.type === 'date') {
    const numVal = data.value != null ? String(data.value) : '';
    return `<c r="${ref}"${styleAttr}><v>${numVal}</v></c>`;
  }

  if (data.type === 'error') {
    return `<c r="${ref}"${styleAttr} t="e"><v>${escapeXml(String(data.value))}</v></c>`;
  }

  // Empty cell with style (preserve formatting)
  if (data.styleIndex !== undefined) {
    return `<c r="${ref}"${styleAttr}/>`;
  }

  return '';
}

// ---- Sheet XML merge (preserving existing structure) ----

/**
 * Merge CellStore changes into existing sheet XML.
 * Only updates <c> elements' values and formulas; preserves everything else
 * (merge cells, conditional formatting, data validations, drawings, etc.).
 */
function mergeSheetXml(existingXml: string, store: CellStore, sst: SharedStringTable): string {
  const doc = XmlParser.parseXml(existingXml);
  const sheetData = doc.getElementsByTagName('sheetData')[0];
  if (!sheetData) return existingXml;

  // Build a map of existing rows/cells for quick lookup
  const rowElements = new Map<number, Element>();
  const existingRows = Array.from(sheetData.getElementsByTagName('row'));
  for (const rowEl of existingRows) {
    const rowIdx = parseInt(rowEl.getAttribute('r') || '0', 10);
    if (rowIdx > 0) rowElements.set(rowIdx, rowEl);
  }

  // Track which cells in the store have been processed
  const processedKeys = new Set<string>();

  // Update existing cells with CellStore values
  for (const [rowIdx, rowEl] of rowElements) {
    const cells = Array.from(rowEl.getElementsByTagName('c'));
    for (const cellEl of cells) {
      const ref = cellEl.getAttribute('r') || '';
      const parsed = CellReference.parse(ref);
      if (parsed.col === 0) continue;

      const storeData = store.get(rowIdx, parsed.col);
      const key = CellReference.cellKey(rowIdx, parsed.col);
      processedKeys.add(key);

      if (!storeData || (storeData.type === 'empty' && storeData.value === '' && !storeData.formula)) {
        // Cell was cleared - remove formula/value but keep the element for style
        const existingFormula = cellEl.getElementsByTagName('f')[0];
        if (existingFormula) cellEl.removeChild(existingFormula);
        const existingValue = cellEl.getElementsByTagName('v')[0];
        if (existingValue) cellEl.removeChild(existingValue);
        const existingIs = cellEl.getElementsByTagName('is')[0];
        if (existingIs) cellEl.removeChild(existingIs);
        cellEl.removeAttribute('t');
        continue;
      }

      // Update formula
      const existingFormula = cellEl.getElementsByTagName('f')[0];
      if (storeData.formula) {
        if (existingFormula) {
          existingFormula.textContent = storeData.formula;
          // Remove shared formula attributes since we're writing the resolved formula
          existingFormula.removeAttribute('si');
          existingFormula.removeAttribute('ref');
          existingFormula.removeAttribute('t');
        } else {
          const fEl = doc.createElement('f');
          fEl.textContent = storeData.formula;
          // Insert formula before value
          const vEl = cellEl.getElementsByTagName('v')[0];
          if (vEl) cellEl.insertBefore(fEl, vEl);
          else cellEl.appendChild(fEl);
        }
      } else if (existingFormula) {
        cellEl.removeChild(existingFormula);
      }

      // Update value
      const existingValue = cellEl.getElementsByTagName('v')[0];
      const existingIs = cellEl.getElementsByTagName('is')[0];

      if (storeData.type === 'string' && storeData.value != null && storeData.value !== '') {
        const idx = sst.map.get(String(storeData.value));
        if (idx !== undefined) {
          cellEl.setAttribute('t', 's');
          if (existingIs) cellEl.removeChild(existingIs);
          if (existingValue) {
            existingValue.textContent = String(idx);
          } else {
            const vEl = doc.createElement('v');
            vEl.textContent = String(idx);
            cellEl.appendChild(vEl);
          }
        } else {
          // Inline string
          cellEl.setAttribute('t', 'inlineStr');
          if (existingValue) cellEl.removeChild(existingValue);
          if (existingIs) {
            const tEl = existingIs.getElementsByTagName('t')[0];
            if (tEl) tEl.textContent = String(storeData.value);
          } else {
            const isEl = doc.createElement('is');
            const tEl = doc.createElement('t');
            tEl.textContent = String(storeData.value);
            isEl.appendChild(tEl);
            cellEl.appendChild(isEl);
          }
        }
      } else if (storeData.type === 'boolean') {
        cellEl.setAttribute('t', 'b');
        if (existingIs) cellEl.removeChild(existingIs);
        if (existingValue) {
          existingValue.textContent = storeData.value ? '1' : '0';
        } else {
          const vEl = doc.createElement('v');
          vEl.textContent = storeData.value ? '1' : '0';
          cellEl.appendChild(vEl);
        }
      } else if (storeData.type === 'error') {
        cellEl.setAttribute('t', 'e');
        if (existingIs) cellEl.removeChild(existingIs);
        if (existingValue) {
          existingValue.textContent = String(storeData.value);
        } else {
          const vEl = doc.createElement('v');
          vEl.textContent = String(storeData.value);
          cellEl.appendChild(vEl);
        }
      } else {
        // Number / date / formula result
        cellEl.removeAttribute('t');
        if (existingIs) cellEl.removeChild(existingIs);
        const val = storeData.value != null ? String(storeData.value) : '';
        if (existingValue) {
          existingValue.textContent = val;
        } else if (val) {
          const vEl = doc.createElement('v');
          vEl.textContent = val;
          cellEl.appendChild(vEl);
        }
      }
    }
  }

  // Add new cells that weren't in the original XML
  store.forEach((key, data, row, col) => {
    if (processedKeys.has(key)) return;
    if (data.type === 'empty' && !data.formula && (data.value === '' || data.value == null)) return;

    let rowEl = rowElements.get(row);
    if (!rowEl) {
      rowEl = doc.createElement('row');
      rowEl.setAttribute('r', String(row));

      // Insert in row order
      let inserted = false;
      for (const child of Array.from(sheetData.children)) {
        const childRow = parseInt(child.getAttribute?.('r') || '0', 10);
        if (childRow > row) {
          sheetData.insertBefore(rowEl, child);
          inserted = true;
          break;
        }
      }
      if (!inserted) sheetData.appendChild(rowEl);
      rowElements.set(row, rowEl);
    }

    const ref = CellReference.toA1(row, col);
    const cellXmlStr = cellToXml(row, col, data, sst);
    if (!cellXmlStr) return;

    // Parse the cell XML and import it
    const tempDoc = XmlParser.parseXml(`<root xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${cellXmlStr}</root>`);
    const newCell = tempDoc.documentElement.firstChild;
    if (newCell) {
      const imported = doc.importNode(newCell, true);
      // Insert in column order
      let inserted = false;
      for (const child of Array.from(rowEl.children)) {
        const childRef = child.getAttribute?.('r') || '';
        const childParsed = CellReference.parse(childRef);
        if (childParsed.col > col) {
          rowEl.insertBefore(imported, child);
          inserted = true;
          break;
        }
      }
      if (!inserted) rowEl.appendChild(imported);
    }
  });

  // Update dimension ref if needed
  const dimensionEl = doc.getElementsByTagName('dimension')[0];
  if (dimensionEl && store.maxRow > 0 && store.maxCol > 0) {
    const dimRef = `A1:${CellReference.toA1(store.maxRow, store.maxCol)}`;
    dimensionEl.setAttribute('ref', dimRef);
  }

  return new XMLSerializer().serializeToString(doc);
}

// ---- Full sheet XML generation (for new sheets) ----

function generateSheetXml(store: CellStore, sst: SharedStringTable): string {
  const maxRow = Math.max(store.maxRow, 1);
  const maxCol = Math.max(store.maxCol, 1);
  const dimRef = `A1:${CellReference.toA1(maxRow, maxCol)}`;

  const rows: string[] = [];
  // Collect cells by row
  const rowMap = new Map<number, Array<{ col: number; data: CellData }>>();
  store.forEach((_key, data, row, col) => {
    if (data.type === 'empty' && !data.formula && (data.value === '' || data.value == null)) return;
    if (!rowMap.has(row)) rowMap.set(row, []);
    rowMap.get(row)!.push({ col, data });
  });

  // Sort rows
  const sortedRows = Array.from(rowMap.keys()).sort((a, b) => a - b);
  for (const rowIdx of sortedRows) {
    const cells = rowMap.get(rowIdx)!.sort((a, b) => a.col - b.col);
    const cellXmls = cells
      .map((c) => cellToXml(rowIdx, c.col, c.data, sst))
      .filter((x) => x);
    if (cellXmls.length > 0) {
      rows.push(`<row r="${rowIdx}">${cellXmls.join('')}</row>`);
    }
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="${dimRef}"/>
<sheetViews><sheetView workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<sheetData>${rows.join('')}</sheetData>
</worksheet>`;
}

// ---- Workbook XML update ----

function updateWorkbookXml(existingXml: string, sheets: SheetInfo[]): string {
  const doc = XmlParser.parseXml(existingXml);
  const sheetsEl = doc.getElementsByTagName('sheets')[0];
  if (!sheetsEl) return existingXml;

  // Remove existing sheet elements
  while (sheetsEl.firstChild) sheetsEl.removeChild(sheetsEl.firstChild);

  // Add updated sheets
  sheets.forEach((sheet, i) => {
    const sheetEl = doc.createElementNS(
      'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'sheet'
    );
    sheetEl.setAttribute('name', sheet.name);
    sheetEl.setAttribute('sheetId', String(i + 1));
    sheetEl.setAttribute('r:id', sheet.relId || `rId${i + 1}`);
    sheetsEl.appendChild(sheetEl);
  });

  return new XMLSerializer().serializeToString(doc);
}

// ---- Relationship XML generation ----

function generateRelsXml(sheets: SheetInfo[], hasStyles: boolean, hasSharedStrings: boolean, hasTheme: boolean): string {
  const rels: string[] = [];
  sheets.forEach((sheet, i) => {
    const relId = sheet.relId || `rId${i + 1}`;
    const target = sheet.target || `worksheets/sheet${i + 1}.xml`;
    // Strip xl/ prefix for rels (target is relative to xl/)
    const relTarget = target.startsWith('xl/') ? target.slice(3) : target;
    rels.push(`<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${relTarget}"/>`);
  });

  let nextId = sheets.length + 1;
  if (hasTheme) {
    rels.push(`<Relationship Id="rId${nextId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`);
    nextId++;
  }
  if (hasStyles) {
    rels.push(`<Relationship Id="rId${nextId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`);
    nextId++;
  }
  if (hasSharedStrings) {
    rels.push(`<Relationship Id="rId${nextId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`);
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`;
}

// ---- Content Types generation ----

function generateContentTypes(sheets: SheetInfo[]): string {
  const overrides = sheets.map((_s, i) =>
    `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
  ).join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
${overrides}
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
}

// ---- Minimal theme XML ----

const MINIMAL_THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2>
<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>
<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
<a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
<a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
</a:theme>`;

// ---- Minimal styles XML ----

const MINIMAL_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

// ---- Core properties ----

function generateCoreProps(): string {
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
<dc:creator>XLSX Reader</dc:creator>
</cp:coreProperties>`;
}

function generateAppProps(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<Application>XLSX Reader</Application>
</Properties>`;
}

function generateTopLevelRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

// ===== Public API =====

export class XlsxWriter {

  /**
   * Save an existing workbook by cloning the original ZIP and merging
   * CellStore changes. Preserves all formatting, styles, drawings, etc.
   */
  static async saveExisting(
    originalZip: JSZipInstance,
    sheets: SheetInfo[],
    cellStores: Map<string, CellStore>,
  ): Promise<Blob> {
    // Build shared strings from all CellStores
    const sst = buildSharedStrings(cellStores);

    // Clone the original ZIP
    const rawData = await originalZip.generateAsync({ type: 'arraybuffer' });
    const zip = await JSZip.loadAsync(rawData);

    // Update shared strings
    zip.file('xl/sharedStrings.xml', serializeSharedStrings(sst));

    // Merge each sheet
    for (const sheet of sheets) {
      const store = cellStores.get(sheet.name);
      if (!store) continue;

      if (sheet.target) {
        // Existing sheet - merge changes into original XML
        const entry = zip.file(sheet.target);
        if (entry) {
          const existingXml = await entry.async('text');
          const mergedXml = mergeSheetXml(existingXml, store, sst);
          zip.file(sheet.target, mergedXml);
        } else {
          // Target doesn't exist in ZIP - generate fresh
          const sheetXml = generateSheetXml(store, sst);
          zip.file(sheet.target, sheetXml);
        }
      } else {
        // New sheet (added locally) - generate fresh sheet XML
        const sheetIndex = sheets.indexOf(sheet);
        const target = `xl/worksheets/sheet${sheetIndex + 1}.xml`;
        sheet.target = target;
        const sheetXml = generateSheetXml(store, sst);
        zip.file(target, sheetXml);
      }
    }

    // Update workbook.xml with current sheet list
    const wbEntry = zip.file('xl/workbook.xml');
    if (wbEntry) {
      const wbXml = await wbEntry.async('text');
      const updatedWb = updateWorkbookXml(wbXml, sheets);
      zip.file('xl/workbook.xml', updatedWb);
    }

    // Update workbook.xml.rels to include new sheets
    const relsEntry = zip.file('xl/_rels/workbook.xml.rels');
    if (relsEntry) {
      const relsXml = await relsEntry.async('text');
      const relsDoc = XmlParser.parseXml(relsXml);
      const existingRels = new Set<string>();
      const relEls = Array.from(relsDoc.getElementsByTagName('Relationship'));
      for (const rel of relEls) {
        existingRels.add(rel.getAttribute('Id') || '');
      }

      // Add rels for any new sheets not already present
      for (const sheet of sheets) {
        const relId = sheet.relId || '';
        if (!existingRels.has(relId) && sheet.target) {
          const relTarget = sheet.target.startsWith('xl/') ? sheet.target.slice(3) : sheet.target;
          const newRel = relsDoc.createElementNS(
            'http://schemas.openxmlformats.org/package/2006/relationships',
            'Relationship'
          );
          newRel.setAttribute('Id', relId);
          newRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
          newRel.setAttribute('Target', relTarget);
          relsDoc.documentElement.appendChild(newRel);
        }
      }

      zip.file('xl/_rels/workbook.xml.rels', new XMLSerializer().serializeToString(relsDoc));
    }

    // Update [Content_Types].xml to include new sheets
    const ctEntry = zip.file('[Content_Types].xml');
    if (ctEntry) {
      const ctXml = await ctEntry.async('text');
      const ctDoc = XmlParser.parseXml(ctXml);
      const existingParts = new Set<string>();
      const overrides = Array.from(ctDoc.getElementsByTagName('Override'));
      for (const ov of overrides) {
        existingParts.add(ov.getAttribute('PartName') || '');
      }

      for (const sheet of sheets) {
        if (sheet.target) {
          const partName = '/' + sheet.target;
          if (!existingParts.has(partName)) {
            const newOverride = ctDoc.createElementNS(
              'http://schemas.openxmlformats.org/package/2006/content-types',
              'Override'
            );
            newOverride.setAttribute('PartName', partName);
            newOverride.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
            ctDoc.documentElement.appendChild(newOverride);
          }
        }
      }

      zip.file('[Content_Types].xml', new XMLSerializer().serializeToString(ctDoc));
    }

    // Remove calcChain.xml since we're not maintaining it (Excel will rebuild it)
    if (zip.file('xl/calcChain.xml')) {
      zip.remove('xl/calcChain.xml');
      // Also remove reference from content types
      const ct2 = zip.file('[Content_Types].xml');
      if (ct2) {
        let ct2Xml = await ct2.async('text');
        ct2Xml = ct2Xml.replace(/<Override[^>]*calcChain[^>]*\/>/g, '');
        zip.file('[Content_Types].xml', ct2Xml);
      }
    }

    return zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
  }

  /**
   * Create a brand new XLSX workbook from scratch.
   */
  static async createNew(
    sheets: SheetInfo[],
    cellStores: Map<string, CellStore>,
  ): Promise<Blob> {
    const sst = buildSharedStrings(cellStores);
    const zip = new JSZip();

    // Top-level rels
    zip.file('_rels/.rels', generateTopLevelRels());

    // Content types
    zip.file('[Content_Types].xml', generateContentTypes(sheets));

    // Doc properties
    zip.file('docProps/core.xml', generateCoreProps());
    zip.file('docProps/app.xml', generateAppProps());

    // Workbook
    const newSheets = sheets.map((s, i) => ({
      ...s,
      relId: `rId${i + 1}`,
      target: `xl/worksheets/sheet${i + 1}.xml`,
    }));

    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>${newSheets.map((s, i) =>
      `<sheet name="${escapeXml(s.name)}" sheetId="${i + 1}" r:id="${s.relId}"/>`
    ).join('')}</sheets>
</workbook>`;
    zip.file('xl/workbook.xml', workbookXml);

    // Workbook rels
    zip.file('xl/_rels/workbook.xml.rels', generateRelsXml(newSheets, true, true, true));

    // Theme
    zip.file('xl/theme/theme1.xml', MINIMAL_THEME);

    // Styles
    zip.file('xl/styles.xml', MINIMAL_STYLES);

    // Shared strings
    zip.file('xl/sharedStrings.xml', serializeSharedStrings(sst));

    // Sheets
    for (let i = 0; i < newSheets.length; i++) {
      const sheet = newSheets[i];
      const store = cellStores.get(sheets[i].name);
      const xml = store ? generateSheetXml(store, sst) : generateSheetXml(new CellStore(sheet.name), sst);
      zip.file(sheet.target!, xml);
    }

    return zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
  }

  /**
   * Trigger a download of a Blob as a file.
   */
  static downloadBlob(blob: Blob, filename: string): void {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 100);
  }
}
