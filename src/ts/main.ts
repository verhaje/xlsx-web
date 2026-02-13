// main.ts - Application entry point

declare const JSZip: any;

import { XmlParser } from './core/XmlParser';
import { WorkbookLoader } from './workbook/WorkbookLoader';
import { StyleParser } from './styles/StyleParser';
import { SheetRenderer } from './renderer/SheetRenderer';
import { TabRenderer } from './renderer/TabRenderer';
import { TableResizer } from './ui/TableResizer';
import { ThemeManager } from './ui/ThemeManager';
import { createFormulaEngine } from './formula/FormulaEngine';
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
const formulaCellRef = document.getElementById('formulaCellRef');
const formulaInput = document.getElementById('formulaInput') as HTMLInputElement | null;

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
  tabsEl.innerHTML = '';
  tableContainer.innerHTML = '';
  sheetNameEl.textContent = '';
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

// ---- Cell selection ----

let selectedCell: HTMLTableCellElement | null = null;
let cellSelectionHandler: ((e: Event) => void) | null = null;

function setupCellSelection(): void {
  if (cellSelectionHandler) {
    tableContainer.removeEventListener('click', cellSelectionHandler);
    cellSelectionHandler = null;
  }

  cellSelectionHandler = (e: Event) => {
    try {
      const target = e.target as HTMLElement;
      const td = target.closest('td') as HTMLTableCellElement | null;
      if (!td) return;

      if (selectedCell) selectedCell.classList.remove('selected');
      td.classList.add('selected');
      selectedCell = td;

      const rowIndex = td.dataset.rowIndex ? Number(td.dataset.rowIndex) : (td.parentElement as HTMLTableRowElement)?.rowIndex ?? 0;
      const colIndex = td.dataset.colIndex ? Number(td.dataset.colIndex) : td.cellIndex;

      if (colIndex > 0 && rowIndex > 0) {
        const colLetter = String.fromCharCode(64 + colIndex);
        const cellRef = `${colLetter}${rowIndex}`;
        if (formulaCellRef) formulaCellRef.textContent = cellRef;

        const formula = td.dataset.formula;
        if (formulaInput) {
          if (formula !== undefined) {
            formulaInput.value = formula;
          } else if (td.dataset.value !== undefined) {
            formulaInput.value = td.dataset.value;
          } else {
            formulaInput.value = td.textContent || '';
          }
        }
      }
    } catch (err) {
      console.error('cell selection handler error', err);
    }
  };

  tableContainer.addEventListener('click', cellSelectionHandler);
}

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

  viewerFileName.textContent = file.name;
  viewerLastEdited.textContent = 'Last edited ' + formatDate(new Date());

  showScreen('viewerScreen');
  updateStatus(`Loading ${file.name}...`);
  clearView();

  try {
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    const loader = new WorkbookLoader(zip);
    const relMap = await loader.loadRelationshipMap();
    const sharedStrings = await loader.loadSharedStrings();
    const styles = await StyleParser.parseStyles(zip);

    const formulaEngine = createFormulaEngine({ localeMap: enLocale });

    const workbookXml = await XmlParser.readZipText(zip, 'xl/workbook.xml');
    const workbookDoc = XmlParser.parseXml(workbookXml);

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

    TabRenderer.renderTabs(tabsEl, sheets, async (sheet: SheetInfo) => {
      updateStatus(`Loading ${sheet.name}...`);
      if (resizeDisposer) { try { resizeDisposer(); } catch {} resizeDisposer = null; }
      const tableEl = await SheetRenderer.renderSheet({
        zip, sheet, sharedStrings, styles, tableContainer, sheetNameEl, formulaEngine, sheets,
      });
      if (tableEl) resizeDisposer = TableResizer.attach(tableEl);
      updateStatus(`Loaded ${sheet.name}.`);
      setupCellSelection();
    });

    // Initial sheet render
    if (resizeDisposer) { try { resizeDisposer(); } catch {} resizeDisposer = null; }
    const initialTable = await SheetRenderer.renderSheet({
      zip, sheet: sheets[0], sharedStrings, styles, tableContainer, sheetNameEl, formulaEngine, sheets,
    });
    if (initialTable) resizeDisposer = TableResizer.attach(initialTable);
    TabRenderer.setActiveTab(tabsEl, 0);
    updateStatus(`Loaded ${sheets[0].name}.`);
    setupCellSelection();
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
