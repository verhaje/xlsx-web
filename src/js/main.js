// main.js - entrypoint module
import { readZipText, parseXml, buildRelationshipMap, normalizeTargetPath, loadSharedStrings } from './parser.js';
import { parseStyles } from './styles.js';
import { renderTabs, renderSheet, setActiveTab } from './renderer.js';
import { attachResizers } from './resize.js';
import { createFormulaEngine } from './formula/engine/index.js';
// Bundle English locale map to avoid fetch/CORS when opening via file://
import enLocale from '../formula/locales/functions.en.json';

// DOM Elements - Screens
const welcomeScreen = document.getElementById('welcomeScreen');
const viewerScreen = document.getElementById('viewerScreen');

// DOM Elements - Welcome Screen
const fileInput = document.getElementById('fileInput');
const uploadZone = document.getElementById('uploadZone');
// recent files list removed

// DOM Elements - Viewer Screen
const tabsEl = document.getElementById('tabs');
const tableContainer = document.getElementById('tableContainer');
const sheetNameEl = document.getElementById('sheetName');
const statusEl = document.getElementById('status');
const backToWelcomeBtn = document.getElementById('backToWelcome');
const viewerFileName = document.getElementById('viewerFileName');
const viewerLastEdited = document.getElementById('viewerLastEdited');
const formulaCellRef = document.getElementById('formulaCellRef');
const formulaInput = document.getElementById('formulaInput');

const THEME_KEY = 'xlsx_reader_theme';

// Theme Management
function getSystemTheme() {
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

function getStoredTheme() {
  try {
    return localStorage.getItem(THEME_KEY);
  } catch {
    return null;
  }
}

function setTheme(theme) {
  if (theme === 'auto') {
    document.documentElement.removeAttribute('data-theme');
    try {
      localStorage.removeItem(THEME_KEY);
    } catch {}
  } else {
    document.documentElement.setAttribute('data-theme', theme);
    try {
      localStorage.setItem(THEME_KEY, theme);
    } catch {}
  }
}

function toggleTheme() {
  const currentTheme = document.documentElement.getAttribute('data-theme');
  const systemTheme = getSystemTheme();
  
  if (!currentTheme) {
    // Currently auto, switch to opposite of system
    setTheme(systemTheme === 'dark' ? 'light' : 'dark');
  } else if (currentTheme === 'dark') {
    setTheme('light');
  } else {
    setTheme('dark');
  }
}

// Initialize theme
const storedTheme = getStoredTheme();
if (storedTheme) {
  setTheme(storedTheme);
}

// Listen for system theme changes when in auto mode
window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', (e) => {
  if (!document.documentElement.getAttribute('data-theme')) {
    // Only react if we're in auto mode
    // CSS will handle the update automatically
  }
});

// Theme toggle buttons
const themeToggleWelcome = document.getElementById('themeToggleWelcome');
const themeToggleViewer = document.getElementById('themeToggleViewer');

if (themeToggleWelcome) {
  themeToggleWelcome.addEventListener('click', toggleTheme);
}

if (themeToggleViewer) {
  themeToggleViewer.addEventListener('click', toggleTheme);
}

function showScreen(screenId) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById(screenId).classList.add('active');
}

function updateStatus(message) {
  if (statusEl) statusEl.textContent = message;
}

function clearView() {
  tabsEl.innerHTML = '';
  tableContainer.innerHTML = '';
  sheetNameEl.textContent = '';
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function formatDate(date) {
  const now = new Date();
  const diffMs = now - date;
  const diffMins = Math.floor(diffMs / 60000);
  const diffHours = Math.floor(diffMs / 3600000);
  const diffDays = Math.floor(diffMs / 86400000);

  if (diffMins < 1) return 'Just now';
  if (diffMins < 60) return `${diffMins}m ago`;
  if (diffHours < 24) return `${diffHours}h ago`;
  if (diffDays < 7) return `${diffDays}d ago`;
  return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

function getFileExtension(filename) {
  const ext = filename.split('.').pop().toLowerCase();
  return ext;
}

function getFileIconSvg(ext) {
  const icons = {
    xlsx: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><path d="M8 13h2M14 13h2M8 17h2M14 17h2"/></svg>',
    xls: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/></svg>',
    csv: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>'
  };
  return icons[ext] || icons.xlsx;
}

// Recent-files feature removed

// Drag and drop handling
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
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      handleFile(files[0]);
    }
  });

}

// Back button
if (backToWelcomeBtn) {
  backToWelcomeBtn.addEventListener('click', () => {
    showScreen('welcomeScreen');
    clearView();
  });
}

// Cell selection handling
let selectedCell = null;
let cellSelectionHandler = null;

function setupCellSelection() {
  // Remove previous handler if already attached
  if (cellSelectionHandler) {
    tableContainer.removeEventListener('click', cellSelectionHandler);
    cellSelectionHandler = null;
  }

  cellSelectionHandler = (e) => {
    try {
      const td = e.target.closest('td');
      if (!td) return;

      // Remove previous selection
      if (selectedCell) {
        selectedCell.classList.remove('selected');
      }

      // Add new selection
      td.classList.add('selected');
      selectedCell = td;

      // Get cell reference
      const row = td.parentElement;
      const rowIndex = td.dataset.rowIndex ? Number(td.dataset.rowIndex) : row.rowIndex;
      const colIndex = td.dataset.colIndex ? Number(td.dataset.colIndex) : td.cellIndex;

      if (colIndex > 0 && rowIndex > 0) {
        const colLetter = String.fromCharCode(64 + colIndex);
        const cellRef = `${colLetter}${rowIndex}`;
        if (formulaCellRef) formulaCellRef.textContent = cellRef;

        // Prefer showing the formula if present, otherwise the computed value
        const formula = td.dataset.formula;
        if (formula !== undefined) {
          formulaInput.value = formula;
        } else if (td.dataset.value !== undefined) {
          formulaInput.value = td.dataset.value;
        } else {
          formulaInput.value = td.textContent || '';
        }
      }
    } catch (err) {
      // swallow any errors to avoid breaking other UI handlers
      console.error('cell selection handler error', err);
    }
  };

  tableContainer.addEventListener('click', cellSelectionHandler);
}

// active resizer disposer for current table
let resizeDisposer = null;

async function handleFile(file) {
  if (!file) {
    updateStatus('No file loaded.');
    return;
  }

  const ext = getFileExtension(file.name);
  if (!['xlsx', 'xls', 'csv'].includes(ext)) {
    alert('Please select a valid Excel file (.xlsx, .xls, .csv)');
    return;
  }

  // recent-files handling removed

  // Update viewer header
  viewerFileName.textContent = file.name;
  viewerLastEdited.textContent = 'Last edited ' + formatDate(new Date());

  // Switch to viewer
  showScreen('viewerScreen');
  updateStatus(`Loading ${file.name}...`);
  clearView();

  try {
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    const workbookXml = await readZipText(zip, 'xl/workbook.xml');
    const workbookDoc = parseXml(workbookXml);
    const relsXml = await readZipText(zip, 'xl/_rels/workbook.xml.rels');
    const relsDoc = parseXml(relsXml);
    const relMap = buildRelationshipMap(relsDoc);
    const sharedStrings = await loadSharedStrings(zip);
    const styles = await parseStyles(zip);

    // Initialize formula engine with bundled English locale map
    let formulaEngine = createFormulaEngine({ localeMap: enLocale });

    const sheets = Array.from(workbookDoc.getElementsByTagName('sheet')).map((sheet) => {
      const name = sheet.getAttribute('name') || 'Sheet';
      const relId = sheet.getAttribute('r:id');
      const target = relMap.get(relId);
      return { name, relId, target: target ? normalizeTargetPath(target) : null };
    }).filter(s => s.target);

    if (sheets.length === 0) {
      updateStatus('No sheets found in workbook.');
      return;
    }

    renderTabs(tabsEl, sheets, async (sheet) => {
      updateStatus(`Loading ${sheet.name}...`);
      // render sheet and attach resizers
      if (resizeDisposer) { try { resizeDisposer(); } catch (e) {} resizeDisposer = null; }
      const tableEl = await renderSheet({ zip, sheet, sharedStrings, styles, tableContainer, sheetNameEl, formulaEngine, sheets });
      if (tableEl) resizeDisposer = attachResizers(tableEl);
      updateStatus(`Loaded ${sheet.name}.`);
      setupCellSelection();
    });

    // initial sheet render
    if (resizeDisposer) { try { resizeDisposer(); } catch (e) {} resizeDisposer = null; }
    const initialTable = await renderSheet({ zip, sheet: sheets[0], sharedStrings, styles, tableContainer, sheetNameEl, formulaEngine, sheets });
    if (initialTable) resizeDisposer = attachResizers(initialTable);
    setActiveTab(tabsEl, 0);
    updateStatus(`Loaded ${sheets[0].name}.`);
    setupCellSelection();
  } catch (err) {
    console.error(err);
    updateStatus('Failed to load file. Make sure it is a valid .xlsx file.');
  }
}

// Attach file input handler idempotently and reset input after use
function initFileInputHandler() {
  function onChange(event) {
    const file = event.target.files[0];
    // Reset input so selecting same file again will fire change
    event.target.value = '';
    if (file) handleFile(file).catch((e) => console.error(e));
  }

  // Ensure we don't attach multiple identical listeners
  fileInput.removeEventListener('change', onChange);
  fileInput.addEventListener('change', onChange);
}

initFileInputHandler();

// recent-files removed; no initialization required

// Ensure upload zone and back button handlers are (re)attached
function initUiHandlers() {
  // Upload zone click -> open file picker
  const uploadClick = (e) => {
    if (e.target.tagName !== 'LABEL' && e.target.tagName !== 'INPUT') {
      fileInput.click();
    }
  };
  uploadZone.removeEventListener('click', uploadClick);
  uploadZone.addEventListener('click', uploadClick);

  // Back button handler: reattach cleanly
  if (backToWelcomeBtn) {
    // remove any anonymous handler by replacing with a stable one
    backToWelcomeBtn.addEventListener('click', () => {
      showScreen('welcomeScreen');
      clearView();
    });
  }
}

initUiHandlers();
