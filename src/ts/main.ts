// main.ts - Application entry point

declare const JSZip: any;

import { XmlParser } from './core/XmlParser';
import { WorkbookLoader } from './workbook/WorkbookLoader';
import { StyleParser } from './styles/StyleParser';
import { SheetRenderer } from './renderer/SheetRenderer';
import { TabRenderer } from './renderer/TabRenderer';
import { CellReference } from './core/CellReference';
import { TableResizer } from './ui/TableResizer';
import { ThemeManager } from './ui/ThemeManager';
import { CellEditor } from './ui/CellEditor';
import { createFormulaEngine } from './formula/FormulaEngine';
import { WorkbookManager } from './data/WorkbookManager';
import { StyleApplicator } from './styles/StyleApplicator';
import { XlsxWriter } from './workbook/XlsxWriter';
import type { SheetInfo, IFormulaEngine } from './types';

// Bundle English locale map to avoid fetch/CORS when opening via file://
import enLocale from '../formula/locales/functions.en.json';

// ---- DOM element references ----

const welcomeScreen = document.getElementById('welcomeScreen')!;
const viewerScreen = document.getElementById('viewerScreen')!;

const fileInput = document.getElementById('fileInput') as HTMLInputElement;
const uploadZone = document.getElementById('uploadZone')!;

const tabsEl = document.getElementById('tabs')!;
const tableContainer = document.getElementById('tableContainer')!;
const sheetNameEl = document.getElementById('sheetName')!;
const statusEl = document.getElementById('status');
const backToWelcomeBtn = document.getElementById('backToWelcome');
const viewerFileName = document.getElementById('viewerFileName')!;
const viewerLastEdited = document.getElementById('viewerLastEdited')!;
const formulaCellRef = document.getElementById('formulaCellRef')!;
const formulaInput = document.getElementById('formulaInput') as HTMLInputElement;
const formulaBar = document.querySelector('.formula-bar') as HTMLElement;
const btnSave = document.getElementById('btnSave');
const btnNewWorkbook = document.getElementById('btnNewWorkbook');

// ---- Theme initialization ----

ThemeManager.init();

const themeToggleWelcome = document.getElementById('themeToggleWelcome');
const themeToggleViewer = document.getElementById('themeToggleViewer');
if (themeToggleWelcome) themeToggleWelcome.addEventListener('click', () => ThemeManager.toggleTheme());
if (themeToggleViewer) themeToggleViewer.addEventListener('click', () => ThemeManager.toggleTheme());

// ---- Screen navigation ----

function showScreen(screenId: string): void {
  document.querySelectorAll('.screen').forEach((s) => s.classList.remove('active'));
  document.getElementById(screenId)?.classList.add('active');
}

function updateStatus(message: string): void {
  if (statusEl) statusEl.textContent = message;
}

function clearView(): void {
  tabsEl.replaceChildren();
  tableContainer.replaceChildren();
  sheetNameEl.textContent = '';
}

function createBlankSheetTable(maxRow = 50, maxCol = 26): HTMLTableElement {
  const table = document.createElement('table');
  table.className = 'sheet-table';

  const colgroup = document.createElement('colgroup');
  for (let c = 1; c <= maxCol; c += 1) {
    const colEl = document.createElement('col');
    colEl.dataset.colIndex = String(c);
    colgroup.appendChild(colEl);
  }
  table.appendChild(colgroup);

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  const corner = document.createElement('th');
  corner.textContent = '';
  headerRow.appendChild(corner);
  for (let col = 1; col <= maxCol; col += 1) {
    const th = document.createElement('th');
    th.textContent = CellReference.columnIndexToName(col);
    th.dataset.colIndex = String(col);
    headerRow.appendChild(th);
  }
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  for (let row = 1; row <= maxRow; row += 1) {
    const tr = document.createElement('tr');
    const rowHeader = document.createElement('th');
    rowHeader.textContent = String(row);
    rowHeader.dataset.rowIndex = String(row);
    tr.appendChild(rowHeader);

    for (let col = 1; col <= maxCol; col += 1) {
      const td = document.createElement('td');
      td.dataset.rowIndex = String(row);
      td.dataset.colIndex = String(col);
      td.dataset.value = '';
      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  tableContainer.replaceChildren(table);
  return table;
}

function generateSheetName(sheets: SheetInfo[]): string {
  const existing = new Set(sheets.map((s) => s.name));
  let index = 1;
  while (existing.has(`Sheet${index}`)) index += 1;
  return `Sheet${index}`;
}

// ---- Utility helpers ----

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function formatDate(date: Date): string {
  const now = new Date();
  const diffMs = now.getTime() - date.getTime();
  const diffMins = Math.floor(diffMs / 60000);
  const diffHours = Math.floor(diffMs / 3600000);
  const diffDays = Math.floor(diffMs / 86400000);

  if (diffMins < 1) return 'Just now';
  if (diffMins < 60) return `${diffMins}m ago`;
  if (diffHours < 24) return `${diffHours}h ago`;
  if (diffDays < 7) return `${diffDays}d ago`;
  return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

function getFileExtension(filename: string): string {
  return filename.split('.').pop()?.toLowerCase() || '';
}

// ---- Drag and drop ----

if (uploadZone) {
  uploadZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadZone.classList.add('dragover');
  });
  uploadZone.addEventListener('dragleave', () => {
    uploadZone.classList.remove('dragover');
  });
  uploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadZone.classList.remove('dragover');
    const files = (e as DragEvent).dataTransfer?.files;
    if (files && files.length > 0) handleFile(files[0]);
  });
}

// ---- Back button ----

if (backToWelcomeBtn) {
  backToWelcomeBtn.addEventListener('click', () => {
    showScreen('welcomeScreen');
    clearView();
  });
}

// ---- Cell editor ----

let cellEditor: CellEditor | null = null;
let workbookManager: WorkbookManager | null = null;
let sheetList: SheetInfo[] = [];
let activeSheetIndex = 0;
let tabStartIndex = 0;
let rerenderTabs: (() => void) | null = null;
let tabResizeHandler: (() => void) | null = null;
let currentFileName = 'workbook.xlsx';
let isNewWorkbook = false;

// ---- Main file handler ----

let resizeDisposer: (() => void) | null = null;

async function handleFile(file: File): Promise<void> {
  if (!file) {
    updateStatus('No file loaded.');
    return;
  }

  const ext = getFileExtension(file.name);
  if (!['xlsx', 'xls', 'csv'].includes(ext)) {
    alert('Please select a valid Excel file (.xlsx, .xls, .csv)');
    return;
  }

  currentFileName = file.name;
  isNewWorkbook = false;
  viewerFileName.textContent = file.name;
  viewerLastEdited.textContent = 'Last edited ' + formatDate(new Date());

  showScreen('viewerScreen');
  updateStatus(`Loading ${file.name}...`);
  clearView();

  try {
    // Clear caches from previous workbook loads
    XmlParser.clearCache();
    StyleApplicator.clearStyleCache();

    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    const loader = new WorkbookLoader(zip);
    const relMap = await loader.loadRelationshipMap();
    const sharedStrings = await loader.loadSharedStrings();
    const styles = await StyleParser.parseStyles(zip);

    const formulaEngine = createFormulaEngine({ localeMap: enLocale });

    const workbookDoc = await XmlParser.readZipXml(zip, 'xl/workbook.xml');

    const sheets: SheetInfo[] = Array.from(workbookDoc.getElementsByTagName('sheet'))
      .map((sheetEl) => {
        const name = sheetEl.getAttribute('name') || 'Sheet';
        const relId = sheetEl.getAttribute('r:id') || '';
        const target = relMap.get(relId);
        return { name, relId, target: target ? XmlParser.normalizeTargetPath(target) : null };
      })
      .filter((s) => s.target);

    if (sheets.length === 0) {
      updateStatus('No sheets found in workbook.');
      return;
    }

    // Create WorkbookManager for in-memory cell data & background loading
    const wbManager = new WorkbookManager({ zip, sheets, sharedStrings, formulaEngine });
    workbookManager = wbManager;
  sheetList = sheets;
  activeSheetIndex = 0;
  tabStartIndex = 0;

    // Clean up previous editor
    if (cellEditor) { cellEditor.dispose(); cellEditor = null; }

    // Cache rendered sheet tables to avoid expensive re-renders on tab switch.
    // Key: sheet name, Value: { table, scrollTop, scrollLeft }
    const renderedSheetCache = new Map<string, { table: HTMLTableElement; scrollTop: number; scrollLeft: number }>();

    const selectSheet = async (sheetIndex: number): Promise<void> => {
      const sheet = sheetList[sheetIndex];
      if (!sheet) return;

      // Save scroll position of the current sheet before switching away
      const prevSheet = sheetList[activeSheetIndex];
      if (prevSheet) {
        const prevCached = renderedSheetCache.get(prevSheet.name);
        if (prevCached) {
          prevCached.scrollTop = tableContainer.scrollTop;
          prevCached.scrollLeft = tableContainer.scrollLeft;
        }
      }

      activeSheetIndex = sheetIndex;
      updateStatus(`Loading ${sheet.name}...`);
      if (resizeDisposer) { try { resizeDisposer(); } catch {} resizeDisposer = null; }

      wbManager.setActiveSheet(sheet.name);

      let tableEl: HTMLTableElement | void;

      // Check if we already rendered this sheet and can reuse the cached table
      const cached = renderedSheetCache.get(sheet.name);
      if (cached) {
        tableEl = cached.table;
        sheetNameEl.textContent = sheet.name;
        tableContainer.replaceChildren(tableEl);
        // Restore scroll position
        tableContainer.scrollTop = cached.scrollTop;
        tableContainer.scrollLeft = cached.scrollLeft;
      } else if (sheet.target) {
        if (!wbManager.isSheetLoaded(sheet.name)) {
          await wbManager.loadSheet(sheet);
        }

        // Pass the pre-built SheetModel from WorkbookManager so the renderer
        // doesn't re-parse merge cells and shared formulas from the raw XML.
        const cachedModel = wbManager.sheetModels.get(sheet.name);

        tableEl = await SheetRenderer.renderSheet({
          zip, sheet, sharedStrings, styles, tableContainer, sheetNameEl, formulaEngine, sheets: sheetList,
          sheetModel: cachedModel,
        });
        if (tableEl) {
          renderedSheetCache.set(sheet.name, { table: tableEl, scrollTop: 0, scrollLeft: 0 });
        }
      } else {
        wbManager.ensureSheetStore(sheet.name);
        sheetNameEl.textContent = sheet.name;
        tableEl = createBlankSheetTable();
        if (tableEl) {
          renderedSheetCache.set(sheet.name, { table: tableEl, scrollTop: 0, scrollLeft: 0 });
        }
      }

      if (tableEl) resizeDisposer = TableResizer.attach(tableEl);
      updateStatus(`Loaded ${sheet.name}.`);

      // Refresh editor for new sheet and overlay any user edits
      if (cellEditor) {
        cellEditor.refresh();
        cellEditor.applyDirtyCells();
      }
      else {
        cellEditor = new CellEditor({
          tableContainer, formulaInput, formulaCellRef, formulaBar, workbookManager: wbManager,
        });
      }
    };

    const renderTabs = (): void => {
      TabRenderer.renderTabs(tabsEl, {
        sheets: sheetList,
        activeIndex: activeSheetIndex,
        startIndex: tabStartIndex,
        onSelect: async (_sheet, index) => {
          try {
            await selectSheet(index);
          } catch (err) {
            console.error('Error selecting sheet:', err);
            updateStatus('Error loading sheet.');
          }
          renderTabs();
        },
        onNext: () => {
          tabStartIndex += 1;
          renderTabs();
        },
        onPrev: () => {
          tabStartIndex = Math.max(0, tabStartIndex - 1);
          renderTabs();
        },
        onAdd: async () => {
          const name = generateSheetName(sheetList);
          sheetList.push({
            name,
            relId: `local-${Date.now()}-${sheetList.length + 1}`,
            target: null,
          });
          wbManager.ensureSheetStore(name);

          activeSheetIndex = sheetList.length - 1;
          tabStartIndex = Math.max(0, activeSheetIndex);

          await selectSheet(activeSheetIndex);
          renderTabs();
        },
        onRename: async (sheet, index, newName) => {
          if (!newName) {
            alert('Worksheet name cannot be empty.');
            return;
          }
          if (sheetList.some((s, i) => i !== index && s.name === newName)) {
            alert(`A worksheet named "${newName}" already exists.`);
            return;
          }
          if (newName === sheet.name) return;

          const wasActive = activeSheetIndex === index;
          // Move cached rendered table to new name
          const cachedTable = renderedSheetCache.get(sheet.name);
          if (cachedTable) {
            renderedSheetCache.delete(sheet.name);
            renderedSheetCache.set(newName, cachedTable);
          }
          wbManager.renameSheet(sheet.name, newName);

          if (wasActive) {
            wbManager.setActiveSheet(newName);
            sheetNameEl.textContent = newName;
          }

          renderTabs();
          updateStatus(`Renamed worksheet to ${newName}.`);
        },
        onRemove: async (sheet, index) => {
          if (sheetList.length <= 1) {
            alert('At least one worksheet is required.');
            return;
          }

          // Remove cached rendered table
          renderedSheetCache.delete(sheet.name);
          wbManager.removeSheet(sheet.name);

          if (activeSheetIndex >= sheetList.length) {
            activeSheetIndex = Math.max(0, sheetList.length - 1);
          }
          if (index < activeSheetIndex) {
            activeSheetIndex = Math.max(0, activeSheetIndex - 1);
          }
          if (tabStartIndex > activeSheetIndex) {
            tabStartIndex = activeSheetIndex;
          }

          await selectSheet(activeSheetIndex);
          renderTabs();
        },
      });
    };

    rerenderTabs = renderTabs;

    if (tabResizeHandler) window.removeEventListener('resize', tabResizeHandler);
    tabResizeHandler = () => {
      rerenderTabs?.();
    };
    window.addEventListener('resize', tabResizeHandler);

    // Load first sheet into CellStore and render
    await wbManager.loadActiveSheet(sheetList[0]);
    await selectSheet(0);
    renderTabs();
  } catch (err) {
    console.error(err);
    updateStatus('Failed to load file. Make sure it is a valid .xlsx file.');
  }
}

// ---- File input handler ----

function initFileInputHandler(): void {
  const onChange = (event: Event) => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = '';
    if (file) handleFile(file).catch((e) => console.error(e));
  };
  fileInput.removeEventListener('change', onChange);
  fileInput.addEventListener('change', onChange);
}

initFileInputHandler();

// ---- Upload zone click handler ----

function initUiHandlers(): void {
  const uploadClick = (e: Event) => {
    const target = e.target as HTMLElement;
    if (target.tagName !== 'LABEL' && target.tagName !== 'INPUT') {
      fileInput.click();
    }
  };
  uploadZone.removeEventListener('click', uploadClick);
  uploadZone.addEventListener('click', uploadClick);

  if (backToWelcomeBtn) {
    backToWelcomeBtn.addEventListener('click', () => {
      showScreen('welcomeScreen');
      clearView();
    });
  }
}

initUiHandlers();

// ---- Save button handler ----

async function handleSave(): Promise<void> {
  if (!workbookManager) {
    alert('No workbook is loaded.');
    return;
  }

  const saveBtn = btnSave;
  if (saveBtn) {
    saveBtn.classList.add('saving');
    saveBtn.setAttribute('disabled', 'true');
  }
  updateStatus('Saving...');

  try {
    let blob: Blob;
    if (isNewWorkbook) {
      // New workbook — generate fresh XLSX
      blob = await XlsxWriter.createNew(sheetList, workbookManager.cellStores);
    } else {
      // Existing workbook — merge changes into original ZIP
      blob = await XlsxWriter.saveExisting(
        workbookManager.zip,
        sheetList,
        workbookManager.cellStores,
      );
    }

    const saveName = currentFileName.endsWith('.xlsx')
      ? currentFileName
      : currentFileName + '.xlsx';
    XlsxWriter.downloadBlob(blob, saveName);
    updateStatus(`Saved ${saveName}.`);
  } catch (err) {
    console.error('Save failed:', err);
    updateStatus('Save failed. See console for details.');
    alert('Failed to save workbook. See console for details.');
  } finally {
    if (saveBtn) {
      saveBtn.classList.remove('saving');
      saveBtn.removeAttribute('disabled');
    }
  }
}

if (btnSave) {
  btnSave.addEventListener('click', () => {
    handleSave().catch((e) => console.error(e));
  });
}

// ---- New Workbook handler ----

async function handleNewWorkbook(): Promise<void> {
  // Clean up previous state
  XmlParser.clearCache();
  StyleApplicator.clearStyleCache();
  clearView();

  if (cellEditor) { cellEditor.dispose(); cellEditor = null; }

  isNewWorkbook = true;
  currentFileName = 'New Workbook.xlsx';

  viewerFileName.textContent = currentFileName;
  viewerLastEdited.textContent = 'Just now';
  showScreen('viewerScreen');

  const formulaEngine = createFormulaEngine({ localeMap: enLocale });

  sheetList = [{ name: 'Sheet1', relId: 'rId1', target: null }];
  activeSheetIndex = 0;
  tabStartIndex = 0;

  // Create a "dummy" zip for WorkbookManager (empty zip for new workbooks)
  const emptyZip = new JSZip();
  const wbManager = new WorkbookManager({
    zip: emptyZip,
    sheets: sheetList,
    sharedStrings: [],
    formulaEngine,
  });
  workbookManager = wbManager;

  wbManager.ensureSheetStore('Sheet1');
  wbManager.setActiveSheet('Sheet1');

  // Render blank sheet
  sheetNameEl.textContent = 'Sheet1';
  const table = createBlankSheetTable();
  if (resizeDisposer) { try { resizeDisposer(); } catch {} resizeDisposer = null; }
  resizeDisposer = TableResizer.attach(table);

  cellEditor = new CellEditor({
    tableContainer, formulaInput, formulaCellRef, formulaBar, workbookManager: wbManager,
  });

  const renderTabs = (): void => {
    TabRenderer.renderTabs(tabsEl, {
      sheets: sheetList,
      activeIndex: activeSheetIndex,
      startIndex: tabStartIndex,
      onSelect: async (_sheet, index) => {
        const sheet = sheetList[index];
        if (!sheet) return;
        activeSheetIndex = index;
        wbManager.setActiveSheet(sheet.name);
        wbManager.ensureSheetStore(sheet.name);
        sheetNameEl.textContent = sheet.name;
        const tbl = createBlankSheetTable();
        if (resizeDisposer) { try { resizeDisposer(); } catch {} resizeDisposer = null; }
        resizeDisposer = TableResizer.attach(tbl);
        if (cellEditor) {
          cellEditor.refresh();
          cellEditor.applyDirtyCells();
        }
        renderTabs();
      },
      onNext: () => { tabStartIndex += 1; renderTabs(); },
      onPrev: () => { tabStartIndex = Math.max(0, tabStartIndex - 1); renderTabs(); },
      onAdd: async () => {
        const name = generateSheetName(sheetList);
        sheetList.push({ name, relId: `local-${Date.now()}-${sheetList.length + 1}`, target: null });
        wbManager.ensureSheetStore(name);
        activeSheetIndex = sheetList.length - 1;
        tabStartIndex = Math.max(0, activeSheetIndex);
        wbManager.setActiveSheet(name);
        sheetNameEl.textContent = name;
        const tbl = createBlankSheetTable();
        if (resizeDisposer) { try { resizeDisposer(); } catch {} resizeDisposer = null; }
        resizeDisposer = TableResizer.attach(tbl);
        if (cellEditor) {
          cellEditor.refresh();
          cellEditor.applyDirtyCells();
        }
        renderTabs();
      },
      onRename: async (sheet, index, newName) => {
        if (!newName) { alert('Worksheet name cannot be empty.'); return; }
        if (sheetList.some((s, i) => i !== index && s.name === newName)) {
          alert(`A worksheet named "${newName}" already exists.`);
          return;
        }
        if (newName === sheet.name) return;
        const wasActive = activeSheetIndex === index;
        wbManager.renameSheet(sheet.name, newName);
        if (wasActive) { wbManager.setActiveSheet(newName); sheetNameEl.textContent = newName; }
        renderTabs();
      },
      onRemove: async (sheet, index) => {
        if (sheetList.length <= 1) { alert('At least one worksheet is required.'); return; }
        wbManager.removeSheet(sheet.name);
        if (activeSheetIndex >= sheetList.length) activeSheetIndex = Math.max(0, sheetList.length - 1);
        if (index < activeSheetIndex) activeSheetIndex = Math.max(0, activeSheetIndex - 1);
        if (tabStartIndex > activeSheetIndex) tabStartIndex = activeSheetIndex;
        const s = sheetList[activeSheetIndex];
        wbManager.setActiveSheet(s.name);
        wbManager.ensureSheetStore(s.name);
        sheetNameEl.textContent = s.name;
        const tbl = createBlankSheetTable();
        if (resizeDisposer) { try { resizeDisposer(); } catch {} resizeDisposer = null; }
        resizeDisposer = TableResizer.attach(tbl);
        if (cellEditor) { cellEditor.refresh(); cellEditor.applyDirtyCells(); }
        renderTabs();
      },
    });
  };

  rerenderTabs = renderTabs;
  if (tabResizeHandler) window.removeEventListener('resize', tabResizeHandler);
  tabResizeHandler = () => { rerenderTabs?.(); };
  window.addEventListener('resize', tabResizeHandler);

  renderTabs();
  updateStatus('New workbook created.');
}

if (btnNewWorkbook) {
  btnNewWorkbook.addEventListener('click', () => {
    handleNewWorkbook().catch((e) => console.error(e));
  });
}
