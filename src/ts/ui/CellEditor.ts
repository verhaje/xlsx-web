// CellEditor.ts - Inline cell editing and formula bar integration
//
// Manages the editing lifecycle:
// - Formula bar: shows "=FORMULA" for formula cells, value for data cells
// - Inline editing: double-click a cell to edit its value (non-formula)
// - Formula edit: click formula bar to edit formulas
// - Commits changes via RecalcEngine which cascades to dependents

import { CellReference } from '../core/CellReference';
import { WorkbookManager } from '../data/WorkbookManager';
import { FormulaAutocomplete } from './FormulaAutocomplete';
import type { CellStore, CellData } from '../data/CellStore';

export interface CellEditorOptions {
  /** The table container element */
  tableContainer: HTMLElement;
  /** The formula bar input element */
  formulaInput: HTMLInputElement;
  /** The cell reference display element */
  formulaCellRef: HTMLElement;
  /** The formula bar container (for autocomplete positioning) */
  formulaBar: HTMLElement;
  /** The workbook manager */
  workbookManager: WorkbookManager;
}

/**
 * CellEditor handles all cell editing interactions.
 */
export class CellEditor {
  private tableContainer: HTMLElement;
  private formulaInput: HTMLInputElement;
  private formulaCellRef: HTMLElement;
  private formulaBar: HTMLElement;
  private manager: WorkbookManager;
  private autocomplete: FormulaAutocomplete;

  /** Currently selected cell */
  private selectedTd: HTMLTableCellElement | null = null;
  private selectedRow = 0;
  private selectedCol = 0;

  /** Edit state */
  private isEditing = false;
  private editMode: 'inline' | 'formulaBar' | null = null;
  private inlineInput: HTMLInputElement | null = null;

  /** Event handler references for cleanup */
  private cellClickHandler: ((e: Event) => void) | null = null;
  private cellDblClickHandler: ((e: Event) => void) | null = null;
  private formulaInputKeyHandler: ((e: KeyboardEvent) => void) | null = null;
  private formulaInputHandler: ((e: Event) => void) | null = null;
  private formulaFocusHandler: ((e: Event) => void) | null = null;
  private formulaBlurHandler: ((e: Event) => void) | null = null;
  private keydownHandler: ((e: KeyboardEvent) => void) | null = null;

  constructor(options: CellEditorOptions) {
    this.tableContainer = options.tableContainer;
    this.formulaInput = options.formulaInput;
    this.formulaCellRef = options.formulaCellRef;
    this.formulaBar = options.formulaBar;
    this.manager = options.workbookManager;

    // Create autocomplete
    this.autocomplete = new FormulaAutocomplete(this.formulaBar);
    this.autocomplete.onSelect = (name: string) => this.insertAutocomplete(name);

    // Set up the workbook manager's DOM update callback
    this.manager.onCellUpdated = (sheetName, row, col, value) => {
      this.updateCellDOM(sheetName, row, col, value);
    };

    this.setupEventListeners();
  }

  /**
   * Refresh the editor for a new active sheet.
   */
  refresh(): void {
    this.cancelEdit();
    this.selectedTd = null;
    this.selectedRow = 0;
    this.selectedCol = 0;
    this.formulaCellRef.textContent = 'A1';
    this.formulaInput.value = '';
  }

  /**
   * Apply dirty (user-edited) cell values from the CellStore onto the rendered DOM.
   * Call this after SheetRenderer.renderSheet to overlay edits on the re-rendered table.
   */
  applyDirtyCells(): void {
    const sheetName = this.manager.activeSheet;
    const store = this.manager.cellStores.get(sheetName);
    if (!store) return;

    const dirtyCells = store.getDirtyCells();
    for (const { row, col, data } of dirtyCells) {
      const td = this.findTdByRowCol(row, col);
      if (!td) continue;

      const displayValue = data.value !== undefined && data.value !== null ? String(data.value) : '';
      td.textContent = displayValue;
      td.dataset.value = displayValue;
      if (data.formula) {
        td.dataset.formula = data.formula;
      } else {
        delete td.dataset.formula;
      }
    }
  }

  /**
   * Clean up all event listeners.
   */
  dispose(): void {
    if (this.cellClickHandler) this.tableContainer.removeEventListener('click', this.cellClickHandler);
    if (this.cellDblClickHandler) this.tableContainer.removeEventListener('dblclick', this.cellDblClickHandler);
    if (this.formulaInputKeyHandler) this.formulaInput.removeEventListener('keydown', this.formulaInputKeyHandler);
    if (this.formulaInputHandler) this.formulaInput.removeEventListener('input', this.formulaInputHandler);
    if (this.formulaFocusHandler) this.formulaInput.removeEventListener('focus', this.formulaFocusHandler);
    if (this.formulaBlurHandler) this.formulaInput.removeEventListener('blur', this.formulaBlurHandler);
    if (this.keydownHandler) document.removeEventListener('keydown', this.keydownHandler);
    this.autocomplete.dispose();
  }

  // ---- Private: Event setup ----

  private setupEventListeners(): void {
    // Cell click → select cell, show in formula bar
    this.cellClickHandler = (e: Event) => {
      const target = e.target as HTMLElement;
      // Don't interfere with inline editing
      if (target.tagName === 'INPUT') return;
      const td = target.closest('td') as HTMLTableCellElement | null;
      if (!td) return;

      // If we're editing a different cell, commit first
      if (this.isEditing) {
        this.commitEdit();
      }

      this.selectCell(td);
    };
    this.tableContainer.addEventListener('click', this.cellClickHandler);

    // Double-click → inline edit
    this.cellDblClickHandler = (e: Event) => {
      e.preventDefault(); // Prevent default text selection and scroll
      const target = e.target as HTMLElement;
      const td = target.closest('td') as HTMLTableCellElement | null;
      if (!td) return;

      this.selectCell(td);
      this.startInlineEdit();
    };
    this.tableContainer.addEventListener('dblclick', this.cellDblClickHandler);

    // Formula bar: make it editable
    this.formulaInput.removeAttribute('readonly');

    // Formula bar focus → start formula bar edit
    this.formulaFocusHandler = () => {
      if (!this.isEditing && this.selectedTd) {
        this.editMode = 'formulaBar';
        this.isEditing = true;
        this.formulaInput.classList.add('editing');
      }
    };
    this.formulaInput.addEventListener('focus', this.formulaFocusHandler);

    // Formula bar blur → commit (unless autocomplete took focus)
    this.formulaBlurHandler = () => {
      // Small delay to allow autocomplete click to fire
      setTimeout(() => {
        if (this.isEditing && this.editMode === 'formulaBar' && !this.autocomplete.isVisible()) {
          this.commitEdit();
        }
      }, 150);
    };
    this.formulaInput.addEventListener('blur', this.formulaBlurHandler);

    // Formula bar input → update autocomplete
    this.formulaInputHandler = () => {
      const val = this.formulaInput.value;
      if (val.startsWith('=')) {
        this.autocomplete.update(val, this.formulaInput.selectionStart || val.length);
      } else {
        this.autocomplete.hide();
      }
    };
    this.formulaInput.addEventListener('input', this.formulaInputHandler);

    // Formula bar keydown → handle Enter/Escape/autocomplete
    this.formulaInputKeyHandler = async (e: KeyboardEvent) => {
      // Let autocomplete handle first
      if (this.autocomplete.handleKey(e)) return;

      if (e.key === 'Enter') {
        e.preventDefault();
        await this.commitEdit();
        this.formulaInput.blur();
        // Move to next row
        this.moveSelection(1, 0);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.cancelEdit();
        this.formulaInput.blur();
      } else if (e.key === 'Tab') {
        e.preventDefault();
        await this.commitEdit();
        this.formulaInput.blur();
        this.moveSelection(0, e.shiftKey ? -1 : 1);
      }
    };
    this.formulaInput.addEventListener('keydown', this.formulaInputKeyHandler);

    // Global keydown for typing into selected cell
    this.keydownHandler = (e: KeyboardEvent) => {
      // Only if not already editing and a cell is selected
      if (this.isEditing || !this.selectedTd) return;
      // Ignore modifier keys, function keys, etc
      if (e.ctrlKey || e.metaKey || e.altKey) return;
      if (e.key.length !== 1 && e.key !== 'Backspace' && e.key !== 'Delete' && e.key !== 'F2') return;

      // F2 → edit current cell
      if (e.key === 'F2') {
        e.preventDefault();
        const cellData = this.getSelectedCellData();
        if (cellData?.formula) {
          // Edit formula in formula bar
          this.formulaInput.focus();
        } else {
          this.startInlineEdit();
        }
        return;
      }

      // Delete/Backspace → clear cell
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        this.clearSelectedCell();
        return;
      }

      // Typing starts → begin editing
      // If it starts with '=', use formula bar
      if (e.key === '=') {
        e.preventDefault();
        this.formulaInput.value = '=';
        this.formulaInput.focus();
        // Place cursor at end
        this.formulaInput.setSelectionRange(1, 1);
        return;
      }

      // Otherwise, start inline edit with the typed character
      this.startInlineEdit(e.key);
      e.preventDefault();
    };
    document.addEventListener('keydown', this.keydownHandler);
  }

  // ---- Cell Selection ----

  private selectCell(td: HTMLTableCellElement): void {
    // Deselect previous
    if (this.selectedTd) this.selectedTd.classList.remove('selected');

    td.classList.add('selected');
    this.selectedTd = td;

    const rowIndex = td.dataset.rowIndex ? Number(td.dataset.rowIndex) : 0;
    const colIndex = td.dataset.colIndex ? Number(td.dataset.colIndex) : 0;
    this.selectedRow = rowIndex;
    this.selectedCol = colIndex;

    if (colIndex > 0 && rowIndex > 0) {
      const colLetter = CellReference.columnIndexToName(colIndex);
      const cellRef = `${colLetter}${rowIndex}`;
      this.formulaCellRef.textContent = cellRef;

      // Show formula or value in formula bar
      const cellData = this.getSelectedCellData();
      if (cellData?.formula) {
        // Show with leading '=' for formulas
        this.formulaInput.value = '=' + cellData.formula;
      } else if (cellData) {
        this.formulaInput.value = cellData.value !== undefined && cellData.value !== null
          ? String(cellData.value) : '';
      } else {
        // Fall back to td data attributes (for compatibility during initial render)
        const formula = td.dataset.formula;
        if (formula) {
          this.formulaInput.value = '=' + formula;
        } else {
          this.formulaInput.value = td.dataset.value ?? td.textContent ?? '';
        }
      }
    }
  }

  // ---- Inline Editing ----

  private startInlineEdit(initialChar?: string): void {
    if (!this.selectedTd || this.selectedRow === 0 || this.selectedCol === 0) return;

    const cellData = this.getSelectedCellData();

    // If it's a formula cell, edit in formula bar instead
    if (cellData?.formula) {
      this.formulaInput.value = '=' + cellData.formula;
      this.formulaInput.focus();
      return;
    }

    this.isEditing = true;
    this.editMode = 'inline';

    // Create an inline input
    const td = this.selectedTd;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'inline-cell-editor';

    if (initialChar) {
      input.value = initialChar;
    } else {
      input.value = cellData?.value !== undefined && cellData?.value !== null
        ? String(cellData.value) : (td.textContent || '');
    }

    // Style to match cell
    input.style.width = '100%';
    input.style.height = '100%';
    input.style.boxSizing = 'border-box';
    input.style.padding = '0 4px';
    input.style.border = 'none';
    input.style.outline = 'none';
    input.style.font = 'inherit';
    input.style.background = 'transparent';

    // If the user starts typing '=', switch to formula bar
    input.addEventListener('input', () => {
      if (input.value.startsWith('=')) {
        this.formulaInput.value = input.value;
        this.removeInlineInput();
        this.editMode = 'formulaBar';
        this.formulaInput.focus();
        // Position cursor at end
        this.formulaInput.setSelectionRange(input.value.length, input.value.length);
      } else {
        // Sync to formula bar
        this.formulaInput.value = input.value;
      }
    });

    input.addEventListener('keydown', async (e: KeyboardEvent) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        await this.commitEdit();
        this.moveSelection(1, 0);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.cancelEdit();
      } else if (e.key === 'Tab') {
        e.preventDefault();
        await this.commitEdit();
        this.moveSelection(0, e.shiftKey ? -1 : 1);
      }
    });

    td.textContent = '';
    td.appendChild(input);
    this.inlineInput = input;
    input.focus({ preventScroll: true });

    if (initialChar) {
      input.setSelectionRange(1, 1);
    } else {
      input.select();
    }
  }

  private removeInlineInput(): void {
    if (this.inlineInput && this.inlineInput.parentElement) {
      const td = this.inlineInput.parentElement as HTMLTableCellElement;
      td.removeChild(this.inlineInput);
    }
    this.inlineInput = null;
  }

  // ---- Edit Commit / Cancel ----

  private async commitEdit(): Promise<void> {
    if (!this.isEditing || this.selectedRow === 0 || this.selectedCol === 0) {
      this.isEditing = false;
      this.editMode = null;
      return;
    }

    let newValue: string;

    if (this.editMode === 'inline' && this.inlineInput) {
      newValue = this.inlineInput.value;
      this.removeInlineInput();
    } else {
      newValue = this.formulaInput.value;
    }

    // Capture the cell references BEFORE any async work, because moveSelection
    // (called by the caller after commitEdit) may change them.
    const commitTd = this.selectedTd;
    const commitRow = this.selectedRow;
    const commitCol = this.selectedCol;

    this.isEditing = false;
    this.editMode = null;
    this.formulaInput.classList.remove('editing');
    this.autocomplete.hide();

    const sheetName = this.manager.activeSheet;
    if (!sheetName) return;

    // Ensure the dependency graph is built before committing edits
    await this.manager.ensureGraphReady();

    // Determine if it's a formula
    if (newValue.startsWith('=')) {
      // Set formula
      const computedValue = await this.manager.recalcEngine.setCellFormula(
        sheetName, commitRow, commitCol, newValue
      );

      // Update the formula bar to show the computed value  
      // (formula bar keeps showing the formula, DOM shows computed)
      if (commitTd) {
        commitTd.textContent = computedValue !== undefined && computedValue !== null
          ? String(computedValue) : '';
        commitTd.dataset.formula = newValue.slice(1);
        commitTd.dataset.value = String(computedValue ?? '');
      }
    } else {
      // Set plain value — update DOM immediately before async work
      if (commitTd) {
        commitTd.textContent = newValue;
        delete commitTd.dataset.formula;
        commitTd.dataset.value = newValue;
      }

      await this.manager.recalcEngine.setCellValue(
        sheetName, commitRow, commitCol, newValue
      );
    }
  }

  private cancelEdit(): void {
    if (this.editMode === 'inline') {
      this.removeInlineInput();
      // Restore original content
      if (this.selectedTd) {
        const cellData = this.getSelectedCellData();
        this.selectedTd.textContent = cellData?.value !== undefined && cellData?.value !== null
          ? String(cellData.value) : '';
      }
    }

    // Restore formula bar
    if (this.selectedTd) {
      const cellData = this.getSelectedCellData();
      if (cellData?.formula) {
        this.formulaInput.value = '=' + cellData.formula;
      } else {
        this.formulaInput.value = cellData?.value !== undefined && cellData?.value !== null
          ? String(cellData.value) : '';
      }
    }

    this.isEditing = false;
    this.editMode = null;
    this.formulaInput.classList.remove('editing');
    this.autocomplete.hide();
  }

  // ---- Helpers ----

  private getSelectedCellData(): CellData | undefined {
    if (this.selectedRow === 0 || this.selectedCol === 0) return undefined;
    const store = this.manager.cellStores.get(this.manager.activeSheet);
    return store?.get(this.selectedRow, this.selectedCol);
  }

  private clearSelectedCell(): void {
    if (this.selectedRow === 0 || this.selectedCol === 0) return;
    const sheetName = this.manager.activeSheet;
    const store = this.manager.cellStores.get(sheetName);
    if (!store) return;

    store.set(this.selectedRow, this.selectedCol, '');
    this.manager.graph.clearCell(sheetName, this.selectedRow, this.selectedCol);

    if (this.selectedTd) {
      this.selectedTd.textContent = '';
      delete this.selectedTd.dataset.formula;
      this.selectedTd.dataset.value = '';
    }
    this.formulaInput.value = '';

    // Ensure graph is ready before recalculating dependents
    this.manager.ensureGraphReady().then(() => {
      this.manager.recalcEngine.setCellValue(sheetName, this.selectedRow, this.selectedCol, '');
    });
  }

  /**
   * Update a cell's DOM representation after recalculation.
   */
  private updateCellDOM(sheetName: string, row: number, col: number, value: any): void {
    // Only update DOM cells on the active sheet
    if (sheetName !== this.manager.activeSheet) return;

    const td = this.findTdByRowCol(row, col);
    if (!td) return;

    td.textContent = value !== undefined && value !== null ? String(value) : '';
    td.dataset.value = String(value ?? '');

    // If this is the selected cell, update formula bar too
    if (row === this.selectedRow && col === this.selectedCol) {
      const cellData = this.getSelectedCellData();
      if (cellData?.formula) {
        this.formulaInput.value = '=' + cellData.formula;
      } else {
        this.formulaInput.value = value !== undefined && value !== null ? String(value) : '';
      }
    }
  }

  private findTdByRowCol(row: number, col: number): HTMLTableCellElement | null {
    return this.tableContainer.querySelector(
      `td[data-row-index="${row}"][data-col-index="${col}"]`
    ) as HTMLTableCellElement | null;
  }

  /**
   * Move selection to an adjacent cell.
   */
  private moveSelection(rowDelta: number, colDelta: number): void {
    const newRow = this.selectedRow + rowDelta;
    const newCol = this.selectedCol + colDelta;
    if (newRow < 1 || newCol < 1) return;

    const td = this.findTdByRowCol(newRow, newCol);
    if (td) {
      this.selectCell(td);
    }
  }

  /**
   * Insert an autocomplete suggestion into the formula bar.
   */
  private insertAutocomplete(functionName: string): void {
    const input = this.formulaInput;
    const cursorPos = input.selectionStart || input.value.length;
    const text = input.value;

    // Find the token we're replacing
    const beforeCursor = text.slice(0, cursorPos);
    const match = /([A-Za-z_][A-Za-z0-9_]*)$/.exec(beforeCursor);

    if (match) {
      const tokenStart = cursorPos - match[1].length;
      const afterCursor = text.slice(cursorPos);
      const newText = text.slice(0, tokenStart) + functionName + '(' + afterCursor;
      input.value = newText;

      // Position cursor inside the parentheses
      const newCursorPos = tokenStart + functionName.length + 1;
      input.setSelectionRange(newCursorPos, newCursorPos);
      input.focus();
    }
  }
}
