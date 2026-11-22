export type HorizontalAlign = 'left' | 'center' | 'right' | 'general';
export type VerticalAlign = 'top' | 'middle' | 'bottom';

export type NumberFormatCategory =
  | 'general'
  | 'number'
  | 'currency'
  | 'accounting'
  | 'date'
  | 'longDate'
  | 'time'
  | 'percentage'
  | 'fraction'
  | 'scientific'
  | 'text';

export interface NumberFormatConfig {
  category: NumberFormatCategory;
  decimals?: number;
  currencySymbol?: string;
  useThousands?: boolean;
}

export interface CellBorderSide {
  style: 'thin' | 'medium' | 'thick' | 'dashed' | 'double';
  color: string;
}

export interface CellBorders {
  top?: CellBorderSide;
  right?: CellBorderSide;
  bottom?: CellBorderSide;
  left?: CellBorderSide;
}

export interface CellStyle {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  color?: string;
  backgroundColor?: string;
  horizontalAlign?: HorizontalAlign;
  verticalAlign?: VerticalAlign;
  wrapText?: boolean;
  indent?: number;
  borders?: CellBorders;
  numberFormat?: NumberFormatConfig;
  locked?: boolean;
}

export type CellValue = string | number | boolean | null;

export interface SparklineConfig {
  type: 'line' | 'column' | 'winloss';
  range: string;
  color?: string;
  negativeColor?: string;
}

export interface DataValidationRule {
  type: 'list' | 'whole' | 'decimal' | 'textLength';
  operator?: 'between' | 'greaterThan' | 'lessThan' | 'equal' | 'notEqual';
  formula1: string; // comma-separated items for list, or min/target
  formula2?: string; // max for between
  showDropdown?: boolean;
  inputTitle?: string;
  inputMessage?: string;
  errorTitle?: string;
  errorMessage?: string;
}

export interface CellData {
  raw: string; // e.g., "=SUM(A1:A5)", "1200", "Hello"
  style?: CellStyle;
  comment?: string;
  hyperlink?: string;
  isCheckbox?: boolean;
  sparkline?: SparklineConfig;
  validation?: DataValidationRule;
}

export interface EvaluatedCell {
  value: CellValue;
  error?: string; // e.g., "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#CIRC!"
  formatted: string;
}

export interface MergedRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export type ConditionalRuleType =
  | 'greaterThan'
  | 'lessThan'
  | 'between'
  | 'equal'
  | 'textContains'
  | 'duplicate'
  | 'dataBar'
  | 'colorScale';

export interface ConditionalRule {
  id: string;
  range: string; // e.g. "B2:B20"
  type: ConditionalRuleType;
  value1?: string;
  value2?: string;
  backgroundColor?: string;
  textColor?: string;
  barColor?: string;
}

export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'donut' | 'scatter';

export interface SheetChart {
  id: string;
  type: ChartType;
  title: string;
  dataRange: string; // e.g. "A1:D6"
  x: number;
  y: number;
  width: number;
  height: number;
  showLegend: boolean;
  showGridlines: boolean;
  switchRowCol?: boolean;
  palette?: 'excel' | 'ocean' | 'warm' | 'monochrome';
}

export interface AutoFilterConfig {
  range: string; // e.g., "A1:F1" (header row range) or "A1:F20"
  hiddenValuesByCol: Record<number, string[]>; // colIndex -> hidden string values
}

export interface TableRegion {
  id: string;
  name: string;
  range: string; // e.g., "A1:E10"
  stylePreset: 'green' | 'blue' | 'slate' | 'amber';
  bandedRows: boolean;
  hasHeaderRow: boolean;
}

export interface FreezePanes {
  row: number; // number of frozen top rows (0 = none)
  col: number; // number of frozen left cols (0 = none)
}

export interface PageSetup {
  orientation: 'portrait' | 'landscape';
  margins: 'normal' | 'wide' | 'narrow';
  showGridlines: boolean;
  showHeadings: boolean;
  printGridlines: boolean;
  printHeadings: boolean;
  printArea?: string;
}

export interface Sheet {
  id: string;
  name: string;
  tabColor?: string;
  rowCount: number;
  colCount: number;
  cells: Record<string, CellData>; // key: "r,c" e.g. "0,0"
  colWidths: Record<number, number>;
  rowHeights: Record<number, number>;
  hiddenRows: Record<number, boolean>;
  hiddenCols: Record<number, boolean>;
  mergedRanges: MergedRange[];
  conditionalRules: ConditionalRule[];
  charts: SheetChart[];
  autoFilter: AutoFilterConfig | null;
  tables: TableRegion[];
  freezePanes: FreezePanes;
  protected: boolean;
  pageSetup: PageSetup;
}

export interface NamedRange {
  name: string;
  sheetId: string;
  range: string; // e.g. "B2:B10"
  comment?: string;
}

export interface Workbook {
  id: string;
  title: string;
  sheets: Sheet[];
  activeSheetId: string;
  namedRanges: NamedRange[];
  calcMode: 'automatic' | 'manual';
  showFormulas: boolean;
  showFormulaBar: boolean;
  zoom: number; // percentage e.g. 100
  theme: 'light' | 'dark';
  currencySymbol: string;
}

export interface CellCoord {
  row: number;
  col: number;
}

export interface SelectionRange {
  start: CellCoord;
  end: CellCoord;
}

export interface ClipboardData {
  mode: 'copy' | 'cut';
  sheetId: string;
  range: SelectionRange;
  cells: Record<string, CellData>; // relative key "dr,dc"
}

export type RibbonTabId =
  | 'file'
  | 'home'
  | 'insert'
  | 'pageLayout'
  | 'formulas'
  | 'data'
  | 'review'
  | 'view';

export type ExcelMode = 'Ready' | 'Enter' | 'Edit' | 'Point';
