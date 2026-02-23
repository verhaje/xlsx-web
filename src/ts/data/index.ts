// data/index.ts - Barrel export for data module

export { CellStore } from './CellStore';
export type { CellData, CellDataType } from './CellStore';
export { DependencyGraph } from './DependencyGraph';
export { RecalcEngine } from './RecalcEngine';
export type { RecalcContext } from './RecalcEngine';
export { WorkbookManager } from './WorkbookManager';
export type { WorkbookManagerOptions, SheetLoadStatus } from './WorkbookManager';
