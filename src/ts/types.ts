// types.ts - All shared types and interfaces for the xlsx_reader application

import type JSZip from 'jszip';

// ---- Cell Reference Types ----

export interface CellRef {
  row: number;
  col: number;
}

export interface RangeRef {
  start: CellRef;
  end: CellRef;
}

// ---- Workbook / Sheet Types ----

export interface SheetInfo {
  name: string;
  relId: string;
  target: string | null;
}

// ---- Token Types (Formula Engine) ----

export type TokenType = 'NUMBER' | 'STRING' | 'BOOL' | 'CELL' | 'RANGE' | 'FUNC' | 'OP' | 'IDENT' | 'EOF';

export interface Token {
  type: TokenType;
  value: string | number | boolean;
  sheet?: string;
}

// ---- Font / Color / Style Types ----

export interface ColorObject {
  rgb?: string | null;
  theme?: string | null;
  indexed?: string | null;
  tint?: string | null;
}

export interface FontInfo {
  name?: string;
  pt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  rgb?: string | null;
  theme?: string | null;
  indexed?: string | null;
  tint?: string | null;
}

export interface FillInfo {
  patternType?: string | null;
  fg?: ColorObject;
  bg?: ColorObject;
}

export interface BorderSide {
  style: string | null;
  rgb?: string | null;
  theme?: string | null;
  indexed?: string | null;
}

export interface BorderDef {
  left?: BorderSide;
  right?: BorderSide;
  top?: BorderSide;
  bottom?: BorderSide;
  diagonal?: BorderSide;
  [key: string]: BorderSide | undefined;
}

export interface CellXf {
  fontId: string | null;
  applyFont: string | null;
  fillId: string | null;
  applyFill: string | null;
  borderId: string | null;
  applyBorder: string | null;
  xfId?: string | null;
}

export interface DxfStyle {
  font?: FontInfo;
  fill?: ColorObject;
  border?: BorderDef;
}

export interface StyleResult {
  fonts: FontInfo[];
  cellXfs: CellXf[];
  cellStyleXfs: CellXf[];
  themeColors: Record<number, string>;
  dxfs: DxfStyle[];
  fills: FillInfo[];
  borders: BorderDef[];
}

export interface CellStyleRuns {
  type: 'runs';
  runs: Array<{ text: string; style: CssProperties }>;
  baseline: CssProperties;
  style?: CssProperties;
}

export interface CellStyleSingle {
  type: 'single';
  style: CssProperties;
}

export type CellStyleResult = CellStyleRuns | CellStyleSingle;

export type CssProperties = Record<string, string>;

// ---- Shared Formula Types ----

export interface SharedFormulaEntry {
  si: string;
  anchorRef: string;
  baseFormula: string;
}

export interface SharedCellEntry {
  si: string;
  anchorRef?: string;
  baseFormula?: string;
}

// ---- Sheet Model ----

export interface MergeAnchor {
  rowSpan: number;
  colSpan: number;
}

export interface SheetModel {
  mergedRanges: RangeRef[];
  anchorMap: Map<string, MergeAnchor>;
  coveredMap: Map<string, string>;
  sharedFormulas: Map<string, SharedFormulaEntry>;
  sharedCells: Map<string, SharedCellEntry>;
}

export interface ColumnStyleRange {
  min: number;
  max: number;
  style: string;
}

// ---- Image / Drawing Types ----

export interface ImageAnchor {
  embed: string;
  from: CellRef;
  to: CellRef;
  type: string;
  dataUrl?: string;
}

// ---- Formula Engine Types ----

export type FormulaValue = string | number | boolean;

export type BuiltinFunction = (args: any[], meta?: any[]) => Promise<any>;

export interface FormulaContext {
  resolveCell: (ref: string) => Promise<any>;
  resolveCellsBatch?: (refs: string[]) => Promise<any[]>;
  sharedStrings?: string[];
  zip?: JSZip;
  sheetDoc?: Document;
}

export interface IFormulaEngine {
  evaluateFormula(text: string, context?: FormulaContext): Promise<any>;
  registerFunction(name: string, fn: BuiltinFunction): void;
  clearCache(): void;
  ERRORS: Record<string, string>;
}

export interface FormulaEngineOptions {
  localeMap?: Record<string, string>;
}

export interface ParserOptions {
  resolveFunction: (name: string) => string;
  customFunctions: Record<string, BuiltinFunction>;
  builtins: Record<string, BuiltinFunction>;
  resolveCellsBatch?: (refs: string[]) => Promise<any[]>;
}

// ---- Conditional Formatting Types ----

export interface CfvoInfo {
  type: string;
  val: string | null;
  gte: boolean;
}

export interface CfColorInfo {
  rgb?: string | null;
  theme?: string | null;
  tint?: string | null;
}

export interface ColorScaleInfo {
  cfvos: CfvoInfo[];
  colors: CfColorInfo[];
}

export interface DataBarInfo {
  cfvos: CfvoInfo[];
  color: CfColorInfo | null;
  showValue: boolean;
}

export interface IconSetInfo {
  cfvos: CfvoInfo[];
  iconSet: string;
  showValue: boolean;
}

export interface CfRule {
  type: string | null;
  operator: string | null;
  formula: string | null;
  formula1: string | null;
  formula2: string | null;
  formulas: string[];
  css: CssProperties | null;
  priority: number;
  stopIfTrue: boolean;
  text: string | null;
  top: string | null;
  bottom: string | null;
  percent: string | null;
  rank: string | null;
  aboveAverage: string | null;
  equalAverage: string | null;
  colorScale: ColorScaleInfo | null;
  dataBar: DataBarInfo | null;
  iconSet: IconSetInfo | null;
  order: number;
  targets: Map<string, { row: number; col: number; key: string }>;
  _targetCells?: Array<{ row: number; col: number; key: string }>;
  _numericValues?: number[];
}

export interface CfEntry {
  rule: CfRule;
  anchor: CellRef;
}

export interface CfResult {
  matched: boolean;
  css?: CssProperties;
  icon?: { icon: string; color: string };
  hideValue?: boolean;
}

// ---- Render Types ----

export interface RenderSheetOptions {
  zip: JSZip;
  sheet: SheetInfo;
  sharedStrings: string[];
  styles: StyleResult;
  tableContainer: HTMLElement;
  sheetNameEl: HTMLElement;
  formulaEngine: IFormulaEngine;
  sheets: SheetInfo[];
}

export interface SheetCacheEntry {
  cellMap: Map<string, Element>;
  sheetDoc: Document | null;
  maxRow: number;
  maxCol: number;
  sheetModel?: SheetModel;
}

// ---- Range Dimensions (formula engine) ----

export interface RangeDimensions {
  rows: number;
  cols: number;
}

export interface ParsedRef {
  col: number;
  row: number;
}
