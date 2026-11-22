import React, { useEffect, useRef, useState } from 'react';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDownAZ,
  ArrowDownToLine,
  ArrowRightToLine,
  ArrowUpZA,
  BarChart2,
  BarChart3,
  Bold,
  BookOpen,
  Calculator,
  CheckSquare,
  ChevronDown,
  Clipboard,
  ClipboardPaste,
  Columns3,
  Copy,
  Database,
  DollarSign,
  Eraser,
  Eye,
  FileSpreadsheet,
  Filter,
  GitBranch,
  Grid3x3,
  Highlighter,
  IndentDecrease,
  IndentIncrease,
  Italic,
  LayoutGrid,
  LineChart,
  Link2,
  ListFilter,
  Lock,
  Maximize2,
  Merge,
  MessageSquarePlus,
  MessageSquareX,
  PaintBucket,
  Paintbrush,
  Percent,
  PieChart,
  Pin,
  Plus,
  Printer,
  Rows3,
  ScatterChart,
  Scissors,
  Search,
  ShieldCheck,
  Sigma,
  SlidersHorizontal,
  Sparkles,
  SplitSquareHorizontal,
  Strikethrough,
  Table,
  Tag,
  Target,
  Trash2,
  TrendingUp,
  Type,
  Underline,
  Unlock,
  Upload,
  WrapText,
  ZoomIn,
} from 'lucide-react';
import {
  CellStyle,
  ChartType,
  ConditionalRuleType,
  NumberFormatCategory,
  PageSetup,
  RibbonTabId,
  Sheet,
  Workbook,
} from '../types/spreadsheet';
import { FUNCTION_CATALOG } from '../engine/formulaEngine';

interface RibbonProps {
  activeTab: RibbonTabId;
  onSelectTab: (tab: RibbonTabId) => void;
  onOpenFileBackstage: () => void;
  workbook: Workbook;
  sheet: Sheet;
  activeCellStyle: CellStyle;
  formatPainterActive: boolean;
  showAllComments: boolean;
  pageBreakView: boolean;
  onTogglePageBreakView: () => void;
  onToggleTheme: () => void;

  // Clipboard
  onCut: () => void;
  onCopy: () => void;
  onPaste: () => void;
  onPasteValues: () => void;
  onToggleFormatPainter: () => void;

  // Formatting
  onUpdateStyle: (patch: Partial<CellStyle>) => void;
  onApplyBorder: (preset: 'all' | 'outer' | 'bottom' | 'top' | 'left' | 'right' | 'thickBottom' | 'none') => void;
  onMergeCenter: () => void;
  onUnmerge: () => void;
  onSetNumberFormat: (cat: NumberFormatCategory) => void;
  onAdjustDecimals: (delta: 1 | -1) => void;
  onApplyCellStylePreset: (preset: 'normal' | 'good' | 'bad' | 'neutral' | 'heading1' | 'heading2' | 'total') => void;

  // Styles & Tables
  onOpenConditionalModal: (initialType?: ConditionalRuleType) => void;
  onClearConditionalRules: () => void;
  onFormatAsTable: (preset: 'green' | 'blue' | 'slate' | 'amber') => void;

  // Cells & Structure
  onInsertRow: () => void;
  onInsertCol: () => void;
  onAddSheet: () => void;
  onDeleteRow: () => void;
  onDeleteCol: () => void;
  onDeleteSheet: () => void;
  onAutoFitCol: () => void;
  onOpenSizeModal: (kind: 'col' | 'row') => void;
  onHideRow: () => void;
  onUnhideRows: () => void;
  onHideCol: () => void;
  onUnhideCols: () => void;

  // Editing
  onAutoSum: (fn: 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN') => void;
  onFillDown: () => void;
  onFillRight: () => void;
  onFillSeries: () => void;
  onClear: (mode: 'all' | 'formats' | 'contents') => void;
  onSort: (asc: boolean) => void;
  onOpenCustomSort: () => void;
  onToggleFilter: () => void;
  onClearFilterValues: () => void;
  onOpenFindReplace: () => void;

  // Insert
  onOpenPivotModal: () => void;
  onInsertChart: (type: ChartType) => void;
  onInsertSparkline: (type: 'line' | 'column' | 'winloss') => void;
  onOpenCommentModal: () => void;
  onDeleteComment: () => void;
  onToggleShowAllComments: () => void;
  onOpenHyperlinkModal: () => void;
  onToggleCheckbox: () => void;
  onOpenValidationModal: () => void;

  // Page Layout
  onUpdatePageSetup: (patch: Partial<PageSetup>) => void;
  onSetPrintArea: () => void;
  onClearPrintArea: () => void;
  onPrint: () => void;

  // Formulas
  onOpenFunctionWizard: () => void;
  onInsertFormulaTemplate: (fnName: string) => void;
  onOpenNameManager: () => void;
  onDefineNameQuick: () => void;
  onToggleShowFormulas: () => void;
  onTracePrecedents: () => void;
  onTraceDependents: () => void;
  onClearAuditArrows: () => void;
  onToggleCalcMode: () => void;
  onCalculateNow: () => void;

  // Data
  onImportFileClick: () => void;
  onOpenTextToColumns: () => void;
  onOpenRemoveDuplicates: () => void;
  onOpenGoalSeek: () => void;
  onSetFreezePanes: (mode: 'topRow' | 'firstCol' | 'selection' | 'unfreeze') => void;

  // Review & View
  onToggleProtectSheet: () => void;
  onOpenWorkbookStats: () => void;
  onToggleFormulaBar: () => void;
  onSetZoom: (z: number) => void;
  onOpenFormatCellsModal: () => void;
}

const FONT_FAMILIES = [
  'Plus Jakarta Sans',
  'DM Sans',
  'Arial',
  'Georgia',
  'Times New Roman',
  'JetBrains Mono',
  'Courier New',
];

const FONT_SIZES = [9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];

const SWATCH_COLORS = [
  '#ffffff', '#f8fafc', '#f1f5f9', '#e2e8f0', '#94a3b8', '#475569', '#0f172a', '#000000',
  '#fee2e2', '#fecaca', '#ef4444', '#b91c1c', '#fef3c7', '#fde68a', '#f59e0b', '#b45309',
  '#dcfce7', '#bbf7d0', '#107c41', '#166534', '#dbeafe', '#bfdbfe', '#2563eb', '#1e40af',
  '#ede9fe', '#ddd6fe', '#7c3aed', '#5b21b6', '#fce7f3', '#fbcfe8', '#db2777', '#9d174d',
];

export const Ribbon: React.FC<RibbonProps> = (props) => {
  const {
    activeTab,
    onSelectTab,
    onOpenFileBackstage,
    workbook,
    sheet,
    activeCellStyle,
  } = props;

  const [openMenu, setOpenMenu] = useState<string | null>(null);
  const ribbonRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const close = (e: MouseEvent) => {
      if (ribbonRef.current && !ribbonRef.current.contains(e.target as Node)) {
        setOpenMenu(null);
      }
    };
    document.addEventListener('mousedown', close);
    return () => document.removeEventListener('mousedown', close);
  }, []);

  const toggleMenu = (name: string) => setOpenMenu((m) => (m === name ? null : name));
  const runAndClose = (fn: () => void) => {
    fn();
    setOpenMenu(null);
  };

  const tabs: { id: RibbonTabId; label: string }[] = [
    { id: 'file', label: 'File' },
    { id: 'home', label: 'Home' },
    { id: 'insert', label: 'Insert' },
    { id: 'pageLayout', label: 'Page Layout' },
    { id: 'formulas', label: 'Formulas' },
    { id: 'data', label: 'Data' },
    { id: 'review', label: 'Review' },
    { id: 'view', label: 'View' },
  ];

  const btnCls =
    'px-2 py-1 rounded hover:bg-slate-200/70 dark:hover:bg-slate-700/70 active:bg-slate-300/60 text-slate-700 dark:text-slate-200 flex items-center gap-1 text-xs transition-colors whitespace-nowrap shrink-0';
  const activeBtnCls =
    'bg-emerald-100 dark:bg-emerald-900/50 text-[#107c41] dark:text-emerald-300 font-semibold';
  const groupCls =
    'flex flex-col justify-between px-2.5 py-1 border-r border-slate-200 dark:border-slate-700 shrink-0 h-[76px]';
  const groupLabelCls = 'text-[10px] text-slate-400 dark:text-slate-500 text-center leading-none pt-0.5';

  const bigBtnCls =
    'px-2 py-1 rounded hover:bg-slate-200/70 dark:hover:bg-slate-700/70 text-slate-700 dark:text-slate-200 flex flex-col items-center justify-center gap-1 text-[11px] leading-tight transition-colors whitespace-nowrap shrink-0 h-[56px] min-w-[46px]';

  const numberFormatLabel: Record<NumberFormatCategory, string> = {
    general: 'General',
    number: 'Number',
    currency: 'Currency',
    accounting: 'Accounting',
    date: 'Short Date',
    longDate: 'Long Date',
    time: 'Time',
    percentage: 'Percentage',
    fraction: 'Fraction',
    scientific: 'Scientific',
    text: 'Text',
  };

  const currentFormat = activeCellStyle.numberFormat?.category || 'general';

  return (
    <div ref={ribbonRef} className="no-print bg-slate-50 dark:bg-slate-900 border-b border-slate-200 dark:border-slate-700 select-none shrink-0 z-30 relative">
      {/* Tab Strip */}
      <div className="flex items-center px-2 pt-1 gap-0.5 bg-white dark:bg-slate-950 border-b border-slate-200 dark:border-slate-800 overflow-x-auto">
        {tabs.map((t) => {
          const isFile = t.id === 'file';
          const active = activeTab === t.id;
          return (
            <button
              key={t.id}
              type="button"
              onClick={() => {
                if (isFile) {
                  onOpenFileBackstage();
                } else {
                  onSelectTab(t.id);
                }
              }}
              className={`px-3.5 py-1.5 text-xs font-medium rounded-t transition-colors whitespace-nowrap shrink-0 ${
                isFile
                  ? 'bg-[#107c41] text-white hover:bg-[#0c5e31] mr-1'
                  : active
                  ? 'bg-slate-50 dark:bg-slate-900 text-[#107c41] dark:text-emerald-400 border-t-2 border-t-[#107c41] border-x border-slate-200 dark:border-slate-700 -mb-px'
                  : 'text-slate-600 dark:text-slate-300 hover:text-slate-900 dark:hover:text-white hover:bg-slate-100 dark:hover:bg-slate-800'
              }`}
            >
              {t.label}
            </button>
          );
        })}
      </div>

      {/* Ribbon Content Area */}
      <div className="flex items-center px-1.5 py-1 overflow-x-auto overflow-y-visible min-h-[84px]">
        {activeTab === 'home' && (
          <>
            {/* Clipboard */}
            <div className={groupCls}>
              <div className="flex items-center gap-1">
                <div className="relative">
                  <button
                    type="button"
                    onClick={props.onPaste}
                    className={bigBtnCls}
                    title="Paste (Ctrl+V)"
                  >
                    <ClipboardPaste className="w-5 h-5 text-[#107c41]" />
                    <span>Paste</span>
                  </button>
                </div>
                <div className="flex flex-col gap-0.5">
                  <button type="button" onClick={props.onCut} className={btnCls} title="Cut (Ctrl+X)">
                    <Scissors className="w-3.5 h-3.5" />
                    <span>Cut</span>
                  </button>
                  <button type="button" onClick={props.onCopy} className={btnCls} title="Copy (Ctrl+C)">
                    <Copy className="w-3.5 h-3.5" />
                    <span>Copy</span>
                  </button>
                  <button
                    type="button"
                    onClick={props.onToggleFormatPainter}
                    className={`${btnCls} ${props.formatPainterActive ? activeBtnCls : ''}`}
                    title="Format Painter — click then click target cell"
                  >
                    <Paintbrush className="w-3.5 h-3.5" />
                    <span>Painter</span>
                  </button>
                </div>
                <button
                  type="button"
                  onClick={props.onPasteValues}
                  title="Paste Values Only"
                  className="p-1 rounded hover:bg-slate-200/70 dark:hover:bg-slate-700 text-[10px] text-slate-500 dark:text-slate-300 flex flex-col items-center"
                >
                  <Clipboard className="w-3.5 h-3.5" />
                  <span>123</span>
                </button>
              </div>
              <div className={groupLabelCls}>Clipboard</div>
            </div>

            {/* Font */}
            <div className={groupCls}>
              <div className="flex flex-col gap-1">
                <div className="flex items-center gap-1">
                  <select
                    value={activeCellStyle.fontFamily || 'DM Sans'}
                    onChange={(e) => props.onUpdateStyle({ fontFamily: e.target.value })}
                    className="h-6 text-xs px-1.5 bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-600 rounded w-32 text-slate-800 dark:text-slate-100"
                    title="Font Family"
                  >
                    {FONT_FAMILIES.map((f) => (
                      <option key={f} value={f}>
                        {f}
                      </option>
                    ))}
                  </select>
                  <select
                    value={activeCellStyle.fontSize || 12}
                    onChange={(e) => props.onUpdateStyle({ fontSize: Number(e.target.value) })}
                    className="h-6 text-xs px-1 bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-600 rounded w-12 text-slate-800 dark:text-slate-100 font-mono-num"
                    title="Font Size"
                  >
                    {FONT_SIZES.map((s) => (
                      <option key={s} value={s}>
                        {s}
                      </option>
                    ))}
                  </select>
                  <button
                    type="button"
                    onClick={() => props.onUpdateStyle({ fontSize: Math.min(72, (activeCellStyle.fontSize || 12) + 1) })}
                    className={btnCls}
                    title="Increase Font Size"
                  >
                    <span className="text-[11px] font-bold">A+</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => props.onUpdateStyle({ fontSize: Math.max(8, (activeCellStyle.fontSize || 12) - 1) })}
                    className={btnCls}
                    title="Decrease Font Size"
                  >
                    <span className="text-[10px] font-bold">A-</span>
                  </button>
                </div>

                <div className="flex items-center gap-0.5">
                  <button
                    type="button"
                    onClick={() => props.onUpdateStyle({ bold: !activeCellStyle.bold })}
                    className={`${btnCls} ${activeCellStyle.bold ? activeBtnCls : ''}`}
                    title="Bold (Ctrl+B)"
                  >
                    <Bold className="w-3.5 h-3.5" />
                  </button>
                  <button
                    type="button"
                    onClick={() => props.onUpdateStyle({ italic: !activeCellStyle.italic })}
                    className={`${btnCls} ${activeCellStyle.italic ? activeBtnCls : ''}`}
                    title="Italic (Ctrl+I)"
                  >
                    <Italic className="w-3.5 h-3.5" />
                  </button>
                  <button
                    type="button"
                    onClick={() => props.onUpdateStyle({ underline: !activeCellStyle.underline })}
                    className={`${btnCls} ${activeCellStyle.underline ? activeBtnCls : ''}`}
                    title="Underline (Ctrl+U)"
                  >
                    <Underline className="w-3.5 h-3.5" />
                  </button>
                  <button
                    type="button"
                    onClick={() => props.onUpdateStyle({ strikethrough: !activeCellStyle.strikethrough })}
                    className={`${btnCls} ${activeCellStyle.strikethrough ? activeBtnCls : ''}`}
                    title="Strikethrough"
                  >
                    <Strikethrough className="w-3.5 h-3.5" />
                  </button>

                  <div className="h-4 w-px bg-slate-200 dark:bg-slate-700 mx-0.5" />

                  {/* Borders dropdown */}
                  <div className="relative">
                    <button
                      type="button"
                      onClick={() => toggleMenu('borders')}
                      className={btnCls}
                      title="Borders"
                    >
                      <Grid3x3 className="w-3.5 h-3.5" />
                      <ChevronDown className="w-3 h-3" />
                    </button>
                    {openMenu === 'borders' && (
                      <div className="fixed mt-1 w-44 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl py-1 z-50 text-xs">
                        {[
                          { id: 'all', label: 'All Borders' },
                          { id: 'outer', label: 'Outside Borders' },
                          { id: 'bottom', label: 'Bottom Border' },
                          { id: 'thickBottom', label: 'Thick Bottom Border' },
                          { id: 'top', label: 'Top Border' },
                          { id: 'left', label: 'Left Border' },
                          { id: 'right', label: 'Right Border' },
                          { id: 'none', label: 'No Border' },
                        ].map((b) => (
                          <button
                            key={b.id}
                            onClick={() => runAndClose(() => props.onApplyBorder(b.id as any))}
                            className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200"
                          >
                            {b.label}
                          </button>
                        ))}
                      </div>
                    )}
                  </div>

                  {/* Fill Color */}
                  <div className="relative">
                    <button
                      type="button"
                      onClick={() => toggleMenu('fill')}
                      className={btnCls}
                      title="Fill Color"
                    >
                      <PaintBucket className="w-3.5 h-3.5" />
                      <span
                        className="w-2.5 h-1 rounded-xs"
                        style={{ backgroundColor: activeCellStyle.backgroundColor || '#fef08a' }}
                      />
                    </button>
                    {openMenu === 'fill' && (
                      <div className="fixed mt-1 w-52 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl p-2 z-50">
                        <div className="text-[11px] font-semibold text-slate-500 mb-1.5">Fill Color</div>
                        <div className="grid grid-cols-8 gap-1 mb-2">
                          {SWATCH_COLORS.map((col) => (
                            <button
                              key={col}
                              onClick={() => runAndClose(() => props.onUpdateStyle({ backgroundColor: col }))}
                              style={{ backgroundColor: col }}
                              className="w-5 h-5 rounded-xs border border-slate-300 dark:border-slate-600 hover:scale-110 transition-transform"
                            />
                          ))}
                        </div>
                        <button
                          onClick={() => runAndClose(() => props.onUpdateStyle({ backgroundColor: undefined }))}
                          className="w-full py-1 text-xs border border-slate-200 dark:border-slate-600 rounded hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200"
                        >
                          No Fill
                        </button>
                      </div>
                    )}
                  </div>

                  {/* Font Color */}
                  <div className="relative">
                    <button
                      type="button"
                      onClick={() => toggleMenu('fontColor')}
                      className={btnCls}
                      title="Font Color"
                    >
                      <Type className="w-3.5 h-3.5" />
                      <span
                        className="w-2.5 h-1 rounded-xs"
                        style={{ backgroundColor: activeCellStyle.color || '#dc2626' }}
                      />
                    </button>
                    {openMenu === 'fontColor' && (
                      <div className="fixed mt-1 w-52 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl p-2 z-50">
                        <div className="text-[11px] font-semibold text-slate-500 mb-1.5">Font Color</div>
                        <div className="grid grid-cols-8 gap-1 mb-2">
                          {SWATCH_COLORS.map((col) => (
                            <button
                              key={col}
                              onClick={() => runAndClose(() => props.onUpdateStyle({ color: col }))}
                              style={{ backgroundColor: col }}
                              className="w-5 h-5 rounded-xs border border-slate-300 dark:border-slate-600 hover:scale-110 transition-transform"
                            />
                          ))}
                        </div>
                        <button
                          onClick={() => runAndClose(() => props.onUpdateStyle({ color: undefined }))}
                          className="w-full py-1 text-xs border border-slate-200 dark:border-slate-600 rounded hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200"
                        >
                          Automatic
                        </button>
                      </div>
                    )}
                  </div>

                  <button
                    type="button"
                    onClick={props.onOpenFormatCellsModal}
                    className={btnCls}
                    title="Format Cells Dialog"
                  >
                    <SlidersHorizontal className="w-3 h-3" />
                  </button>
                </div>
              </div>
              <div className={groupLabelCls}>Font</div>
            </div>

            {/* Alignment */}
            <div className={groupCls}>
              <div className="flex flex-col gap-1">
                <div className="flex items-center gap-0.5">
                  <button
                    type="button"
                    onClick={() => props.onUpdateStyle({ horizontalAlign: 'left' })}
                    className={`${btnCls} ${activeCellStyle.horizontalAlign === 'left' ? activeBtnCls : ''}`}
                    title="Align Left"
                  >
                    <AlignLeft className="w-3.5 h-3.5" />
                  </button>
                  <button
                    type="button"
                    onClick={() => props.onUpdateStyle({ horizontalAlign: 'center' })}
                    className={`${btnCls} ${activeCellStyle.horizontalAlign === 'center' ? activeBtnCls : ''}`}
                    title="Align Center"
                  >
                    <AlignCenter className="w-3.5 h-3.5" />
                  </button>
                  <button
                    type="button"
                    onClick={() => props.onUpdateStyle({ horizontalAlign: 'right' })}
                    className={`${btnCls} ${activeCellStyle.horizontalAlign === 'right' ? activeBtnCls : ''}`}
                    title="Align Right"
                  >
                    <AlignRight className="w-3.5 h-3.5" />
                  </button>
                  <div className="h-4 w-px bg-slate-200 dark:bg-slate-700 mx-0.5" />
                  <button
                    type="button"
                    onClick={() => props.onUpdateStyle({ verticalAlign: 'top' })}
                    className={`${btnCls} ${activeCellStyle.verticalAlign === 'top' ? activeBtnCls : ''}`}
                    title="Top Align"
                  >
                    <span className="text-[10px] font-semibold">Top</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => props.onUpdateStyle({ verticalAlign: 'middle' })}
                    className={`${btnCls} ${activeCellStyle.verticalAlign === 'middle' ? activeBtnCls : ''}`}
                    title="Middle Align"
                  >
                    <span className="text-[10px] font-semibold">Mid</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => props.onUpdateStyle({ verticalAlign: 'bottom' })}
                    className={`${btnCls} ${activeCellStyle.verticalAlign === 'bottom' ? activeBtnCls : ''}`}
                    title="Bottom Align"
                  >
                    <span className="text-[10px] font-semibold">Bot</span>
                  </button>
                </div>

                <div className="flex items-center gap-0.5">
                  <button
                    type="button"
                    onClick={() => props.onUpdateStyle({ wrapText: !activeCellStyle.wrapText })}
                    className={`${btnCls} ${activeCellStyle.wrapText ? activeBtnCls : ''}`}
                    title="Wrap Text"
                  >
                    <WrapText className="w-3.5 h-3.5" />
                    <span>Wrap</span>
                  </button>
                  <button
                    type="button"
                    onClick={props.onMergeCenter}
                    className={btnCls}
                    title="Merge & Center selected cells"
                  >
                    <Merge className="w-3.5 h-3.5" />
                    <span>Merge</span>
                  </button>
                  <button
                    type="button"
                    onClick={props.onUnmerge}
                    className={btnCls}
                    title="Unmerge Cells"
                  >
                    <span className="text-[10px]">Unmerge</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => props.onUpdateStyle({ indent: Math.max(0, (activeCellStyle.indent || 0) - 1) })}
                    className={btnCls}
                    title="Decrease Indent"
                  >
                    <IndentDecrease className="w-3.5 h-3.5" />
                  </button>
                  <button
                    type="button"
                    onClick={() => props.onUpdateStyle({ indent: Math.min(8, (activeCellStyle.indent || 0) + 1) })}
                    className={btnCls}
                    title="Increase Indent"
                  >
                    <IndentIncrease className="w-3.5 h-3.5" />
                  </button>
                </div>
              </div>
              <div className={groupLabelCls}>Alignment</div>
            </div>

            {/* Number */}
            <div className={groupCls}>
              <div className="flex flex-col gap-1">
                <select
                  value={currentFormat}
                  onChange={(e) => props.onSetNumberFormat(e.target.value as NumberFormatCategory)}
                  className="h-6 text-xs px-1.5 bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-600 rounded w-32 text-slate-800 dark:text-slate-100"
                  title="Number Format"
                >
                  {(Object.keys(numberFormatLabel) as NumberFormatCategory[]).map((cat) => (
                    <option key={cat} value={cat}>
                      {numberFormatLabel[cat]}
                    </option>
                  ))}
                </select>
                <div className="flex items-center gap-0.5">
                  <button
                    type="button"
                    onClick={() => props.onSetNumberFormat('currency')}
                    className={`${btnCls} ${currentFormat === 'currency' ? activeBtnCls : ''}`}
                    title="Currency Format"
                  >
                    <DollarSign className="w-3.5 h-3.5" />
                  </button>
                  <button
                    type="button"
                    onClick={() => props.onSetNumberFormat('percentage')}
                    className={`${btnCls} ${currentFormat === 'percentage' ? activeBtnCls : ''}`}
                    title="Percent Style (%)"
                  >
                    <Percent className="w-3.5 h-3.5" />
                  </button>
                  <button
                    type="button"
                    onClick={() => props.onSetNumberFormat('number')}
                    className={`${btnCls} ${currentFormat === 'number' ? activeBtnCls : ''}`}
                    title="Comma Style (1,000.00)"
                  >
                    <span className="font-mono-num font-bold text-xs">,</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => props.onAdjustDecimals(1)}
                    className={btnCls}
                    title="Increase Decimal"
                  >
                    <span className="font-mono-num text-[11px]">.00</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => props.onAdjustDecimals(-1)}
                    className={btnCls}
                    title="Decrease Decimal"
                  >
                    <span className="font-mono-num text-[11px]">.0</span>
                  </button>
                </div>
              </div>
              <div className={groupLabelCls}>Number</div>
            </div>

            {/* Styles */}
            <div className={groupCls}>
              <div className="flex items-center gap-1">
                <div className="relative">
                  <button type="button" onClick={() => toggleMenu('condFmt')} className={bigBtnCls}>
                    <Highlighter className="w-4 h-4 text-amber-600" />
                    <span className="flex items-center gap-0.5">
                      Cond. Fmt <ChevronDown className="w-2.5 h-2.5" />
                    </span>
                  </button>
                  {openMenu === 'condFmt' && (
                    <div className="fixed mt-1 w-52 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl py-1 z-50 text-xs">
                      <button onClick={() => runAndClose(() => props.onOpenConditionalModal('greaterThan'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Greater Than...</button>
                      <button onClick={() => runAndClose(() => props.onOpenConditionalModal('lessThan'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Less Than...</button>
                      <button onClick={() => runAndClose(() => props.onOpenConditionalModal('between'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Between...</button>
                      <button onClick={() => runAndClose(() => props.onOpenConditionalModal('textContains'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Text That Contains...</button>
                      <button onClick={() => runAndClose(() => props.onOpenConditionalModal('duplicate'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Duplicate Values</button>
                      <div className="border-t border-slate-100 dark:border-slate-700 my-1" />
                      <button onClick={() => runAndClose(() => props.onOpenConditionalModal('dataBar'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Data Bars</button>
                      <button onClick={() => runAndClose(() => props.onOpenConditionalModal('colorScale'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Color Scales (Red-Yellow-Green)</button>
                      <div className="border-t border-slate-100 dark:border-slate-700 my-1" />
                      <button onClick={() => runAndClose(() => props.onOpenConditionalModal())} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 font-medium text-[#107c41]">Manage Rules...</button>
                      <button onClick={() => runAndClose(props.onClearConditionalRules)} className="w-full px-3 py-1.5 text-left hover:bg-red-50 dark:hover:bg-red-900/30 text-red-600">Clear Rules from Sheet</button>
                    </div>
                  )}
                </div>

                <div className="relative">
                  <button type="button" onClick={() => toggleMenu('fmtTable')} className={bigBtnCls}>
                    <Table className="w-4 h-4 text-[#107c41]" />
                    <span className="flex items-center gap-0.5">
                      Table <ChevronDown className="w-2.5 h-2.5" />
                    </span>
                  </button>
                  {openMenu === 'fmtTable' && (
                    <div className="fixed mt-1 w-44 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl py-1 z-50 text-xs">
                      <button onClick={() => runAndClose(() => props.onFormatAsTable('green'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Excel Green Table</button>
                      <button onClick={() => runAndClose(() => props.onFormatAsTable('blue'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Corporate Blue Table</button>
                      <button onClick={() => runAndClose(() => props.onFormatAsTable('slate'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Neutral Slate Table</button>
                      <button onClick={() => runAndClose(() => props.onFormatAsTable('amber'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Warm Amber Table</button>
                    </div>
                  )}
                </div>

                <div className="relative">
                  <button type="button" onClick={() => toggleMenu('cellStyles')} className={bigBtnCls}>
                    <Sparkles className="w-4 h-4 text-sky-600" />
                    <span className="flex items-center gap-0.5">
                      Styles <ChevronDown className="w-2.5 h-2.5" />
                    </span>
                  </button>
                  {openMenu === 'cellStyles' && (
                    <div className="fixed mt-1 w-44 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl py-1 z-50 text-xs">
                      <button onClick={() => runAndClose(() => props.onApplyCellStylePreset('normal'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Normal</button>
                      <button onClick={() => runAndClose(() => props.onApplyCellStylePreset('good'))} className="w-full px-3 py-1.5 text-left bg-emerald-50 text-emerald-800 hover:opacity-90">Good</button>
                      <button onClick={() => runAndClose(() => props.onApplyCellStylePreset('bad'))} className="w-full px-3 py-1.5 text-left bg-red-50 text-red-800 hover:opacity-90">Bad</button>
                      <button onClick={() => runAndClose(() => props.onApplyCellStylePreset('neutral'))} className="w-full px-3 py-1.5 text-left bg-amber-50 text-amber-800 hover:opacity-90">Neutral</button>
                      <button onClick={() => runAndClose(() => props.onApplyCellStylePreset('heading1'))} className="w-full px-3 py-1.5 text-left font-bold text-slate-900 dark:text-white border-b-2 border-[#107c41] hover:bg-slate-100">Heading 1</button>
                      <button onClick={() => runAndClose(() => props.onApplyCellStylePreset('heading2'))} className="w-full px-3 py-1.5 text-left font-semibold text-slate-700 dark:text-slate-200 hover:bg-slate-100">Heading 2</button>
                      <button onClick={() => runAndClose(() => props.onApplyCellStylePreset('total'))} className="w-full px-3 py-1.5 text-left font-bold border-t border-b-2 border-slate-800 hover:bg-slate-100">Total</button>
                    </div>
                  )}
                </div>
              </div>
              <div className={groupLabelCls}>Styles</div>
            </div>

            {/* Cells */}
            <div className={groupCls}>
              <div className="flex items-center gap-1">
                <div className="relative">
                  <button type="button" onClick={() => toggleMenu('insertCells')} className={bigBtnCls}>
                    <Plus className="w-4 h-4 text-[#107c41]" />
                    <span className="flex items-center gap-0.5">
                      Insert <ChevronDown className="w-2.5 h-2.5" />
                    </span>
                  </button>
                  {openMenu === 'insertCells' && (
                    <div className="fixed mt-1 w-44 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl py-1 z-50 text-xs">
                      <button onClick={() => runAndClose(props.onInsertRow)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Insert Sheet Rows</button>
                      <button onClick={() => runAndClose(props.onInsertCol)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Insert Sheet Columns</button>
                      <button onClick={() => runAndClose(props.onAddSheet)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Insert New Sheet</button>
                    </div>
                  )}
                </div>

                <div className="relative">
                  <button type="button" onClick={() => toggleMenu('deleteCells')} className={bigBtnCls}>
                    <Trash2 className="w-4 h-4 text-red-600" />
                    <span className="flex items-center gap-0.5">
                      Delete <ChevronDown className="w-2.5 h-2.5" />
                    </span>
                  </button>
                  {openMenu === 'deleteCells' && (
                    <div className="fixed mt-1 w-44 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl py-1 z-50 text-xs">
                      <button onClick={() => runAndClose(props.onDeleteRow)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Delete Sheet Rows</button>
                      <button onClick={() => runAndClose(props.onDeleteCol)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Delete Sheet Columns</button>
                      <button onClick={() => runAndClose(props.onDeleteSheet)} className="w-full px-3 py-1.5 text-left hover:bg-red-50 text-red-600">Delete Sheet</button>
                    </div>
                  )}
                </div>

                <div className="relative">
                  <button type="button" onClick={() => toggleMenu('formatCells')} className={bigBtnCls}>
                    <Rows3 className="w-4 h-4 text-slate-600 dark:text-slate-300" />
                    <span className="flex items-center gap-0.5">
                      Format <ChevronDown className="w-2.5 h-2.5" />
                    </span>
                  </button>
                  {openMenu === 'formatCells' && (
                    <div className="fixed mt-1 w-48 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl py-1 z-50 text-xs">
                      <button onClick={() => runAndClose(() => props.onOpenSizeModal('row'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Row Height...</button>
                      <button onClick={() => runAndClose(() => props.onOpenSizeModal('col'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Column Width...</button>
                      <button onClick={() => runAndClose(props.onAutoFitCol)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">AutoFit Column Width</button>
                      <div className="border-t border-slate-100 dark:border-slate-700 my-1" />
                      <button onClick={() => runAndClose(props.onHideRow)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Hide Rows</button>
                      <button onClick={() => runAndClose(props.onUnhideRows)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Unhide All Rows</button>
                      <button onClick={() => runAndClose(props.onHideCol)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Hide Columns</button>
                      <button onClick={() => runAndClose(props.onUnhideCols)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Unhide All Columns</button>
                      <div className="border-t border-slate-100 dark:border-slate-700 my-1" />
                      <button onClick={() => runAndClose(props.onOpenFormatCellsModal)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 font-medium text-[#107c41]">Format Cells...</button>
                    </div>
                  )}
                </div>
              </div>
              <div className={groupLabelCls}>Cells</div>
            </div>

            {/* Editing */}
            <div className={`${groupCls} border-r-0`}>
              <div className="flex items-center gap-1">
                <div className="flex flex-col gap-0.5">
                  <div className="relative">
                    <button
                      type="button"
                      onClick={() => toggleMenu('autoSum')}
                      className={btnCls}
                      title="AutoSum"
                    >
                      <Sigma className="w-3.5 h-3.5 text-[#107c41]" />
                      <span>AutoSum</span>
                      <ChevronDown className="w-3 h-3" />
                    </button>
                    {openMenu === 'autoSum' && (
                      <div className="fixed mt-1 w-40 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl py-1 z-50 text-xs">
                        {(['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'] as const).map((fn) => (
                          <button
                            key={fn}
                            onClick={() => runAndClose(() => props.onAutoSum(fn))}
                            className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700"
                          >
                            {fn}
                          </button>
                        ))}
                      </div>
                    )}
                  </div>

                  <div className="relative">
                    <button type="button" onClick={() => toggleMenu('fillMenu')} className={btnCls}>
                      <ArrowDownToLine className="w-3.5 h-3.5 text-sky-600" />
                      <span>Fill</span>
                      <ChevronDown className="w-3 h-3" />
                    </button>
                    {openMenu === 'fillMenu' && (
                      <div className="fixed mt-1 w-40 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl py-1 z-50 text-xs">
                        <button onClick={() => runAndClose(props.onFillDown)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Fill Down (Ctrl+D)</button>
                        <button onClick={() => runAndClose(props.onFillRight)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Fill Right (Ctrl+R)</button>
                        <button onClick={() => runAndClose(props.onFillSeries)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Smart Fill Series</button>
                      </div>
                    )}
                  </div>

                  <div className="relative">
                    <button type="button" onClick={() => toggleMenu('clearMenu')} className={btnCls}>
                      <Eraser className="w-3.5 h-3.5 text-amber-600" />
                      <span>Clear</span>
                      <ChevronDown className="w-3 h-3" />
                    </button>
                    {openMenu === 'clearMenu' && (
                      <div className="fixed mt-1 w-40 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl py-1 z-50 text-xs">
                        <button onClick={() => runAndClose(() => props.onClear('all'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Clear All</button>
                        <button onClick={() => runAndClose(() => props.onClear('formats'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Clear Formats</button>
                        <button onClick={() => runAndClose(() => props.onClear('contents'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Clear Contents (Del)</button>
                      </div>
                    )}
                  </div>
                </div>

                <div className="relative">
                  <button type="button" onClick={() => toggleMenu('sortFilter')} className={bigBtnCls}>
                    <ListFilter className="w-4 h-4 text-[#107c41]" />
                    <span className="flex items-center gap-0.5">
                      Sort & Filter <ChevronDown className="w-2.5 h-2.5" />
                    </span>
                  </button>
                  {openMenu === 'sortFilter' && (
                    <div className="fixed mt-1 w-44 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl py-1 z-50 text-xs">
                      <button onClick={() => runAndClose(() => props.onSort(true))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><ArrowDownAZ className="w-3.5 h-3.5" /> Sort A to Z</button>
                      <button onClick={() => runAndClose(() => props.onSort(false))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><ArrowUpZA className="w-3.5 h-3.5" /> Sort Z to A</button>
                      <button onClick={() => runAndClose(props.onOpenCustomSort)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Custom Sort...</button>
                      <div className="border-t border-slate-100 dark:border-slate-700 my-1" />
                      <button onClick={() => runAndClose(props.onToggleFilter)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><Filter className="w-3.5 h-3.5" /> {sheet.autoFilter ? 'Remove Filter' : 'Filter'}</button>
                      {sheet.autoFilter && (
                        <button onClick={() => runAndClose(props.onClearFilterValues)} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Clear Filters</button>
                      )}
                    </div>
                  )}
                </div>

                <button
                  type="button"
                  onClick={props.onOpenFindReplace}
                  className={bigBtnCls}
                  title="Find & Replace (Ctrl+F)"
                >
                  <Search className="w-4 h-4 text-slate-600 dark:text-slate-300" />
                  <span>Find</span>
                </button>
              </div>
              <div className={groupLabelCls}>Editing</div>
            </div>
          </>
        )}

        {activeTab === 'insert' && (
          <>
            <div className={groupCls}>
              <div className="flex items-center gap-1">
                <button type="button" onClick={props.onOpenPivotModal} className={bigBtnCls}>
                  <LayoutGrid className="w-5 h-5 text-[#107c41]" />
                  <span>PivotTable</span>
                </button>
                <button type="button" onClick={() => props.onFormatAsTable('green')} className={bigBtnCls}>
                  <Table className="w-5 h-5 text-sky-600" />
                  <span>Table</span>
                </button>
              </div>
              <div className={groupLabelCls}>Tables</div>
            </div>

            <div className={groupCls}>
              <div className="flex items-center gap-1">
                <button type="button" onClick={() => props.onInsertChart('column')} className={bigBtnCls} title="Insert Column Chart">
                  <BarChart3 className="w-5 h-5 text-[#107c41]" />
                  <span>Column</span>
                </button>
                <button type="button" onClick={() => props.onInsertChart('bar')} className={bigBtnCls} title="Insert Horizontal Bar Chart">
                  <BarChart2 className="w-5 h-5 text-sky-600 rotate-90" />
                  <span>Bar</span>
                </button>
                <button type="button" onClick={() => props.onInsertChart('line')} className={bigBtnCls} title="Insert Line Chart">
                  <LineChart className="w-5 h-5 text-indigo-600" />
                  <span>Line</span>
                </button>
                <button type="button" onClick={() => props.onInsertChart('area')} className={bigBtnCls} title="Insert Area Chart">
                  <TrendingUp className="w-5 h-5 text-teal-600" />
                  <span>Area</span>
                </button>
                <button type="button" onClick={() => props.onInsertChart('pie')} className={bigBtnCls} title="Insert Pie Chart">
                  <PieChart className="w-5 h-5 text-amber-600" />
                  <span>Pie</span>
                </button>
                <button type="button" onClick={() => props.onInsertChart('donut')} className={bigBtnCls} title="Insert Donut Chart">
                  <PieChart className="w-5 h-5 text-rose-600" />
                  <span>Donut</span>
                </button>
                <button type="button" onClick={() => props.onInsertChart('scatter')} className={bigBtnCls} title="Insert Scatter Plot">
                  <ScatterChart className="w-5 h-5 text-purple-600" />
                  <span>Scatter</span>
                </button>
              </div>
              <div className={groupLabelCls}>Charts (from Selection)</div>
            </div>

            <div className={groupCls}>
              <div className="flex items-center gap-1">
                <button type="button" onClick={() => props.onInsertSparkline('line')} className={bigBtnCls} title="Insert Line Sparkline into cell">
                  <TrendingUp className="w-4 h-4 text-[#107c41]" />
                  <span>Line</span>
                </button>
                <button type="button" onClick={() => props.onInsertSparkline('column')} className={bigBtnCls} title="Insert Column Sparkline into cell">
                  <BarChart3 className="w-4 h-4 text-sky-600" />
                  <span>Column</span>
                </button>
                <button type="button" onClick={() => props.onInsertSparkline('winloss')} className={bigBtnCls} title="Insert Win/Loss Sparkline into cell">
                  <BarChart2 className="w-4 h-4 text-amber-600" />
                  <span>Win/Loss</span>
                </button>
              </div>
              <div className={groupLabelCls}>Sparklines</div>
            </div>

            <div className={`${groupCls} border-r-0`}>
              <div className="flex items-center gap-1">
                <button type="button" onClick={props.onToggleCheckbox} className={bigBtnCls} title="Toggle Interactive Checkbox in selected cells">
                  <CheckSquare className="w-4 h-4 text-[#107c41]" />
                  <span>Checkbox</span>
                </button>
                <button type="button" onClick={props.onOpenValidationModal} className={bigBtnCls} title="Insert Dropdown List / Validation">
                  <ListFilter className="w-4 h-4 text-sky-600" />
                  <span>Dropdown</span>
                </button>
                <button type="button" onClick={props.onOpenHyperlinkModal} className={bigBtnCls} title="Insert Hyperlink">
                  <Link2 className="w-4 h-4 text-indigo-600" />
                  <span>Link</span>
                </button>
                <button type="button" onClick={props.onOpenCommentModal} className={bigBtnCls} title="Insert Comment / Note">
                  <MessageSquarePlus className="w-4 h-4 text-amber-600" />
                  <span>Comment</span>
                </button>
              </div>
              <div className={groupLabelCls}>Controls & Links</div>
            </div>
          </>
        )}

        {activeTab === 'pageLayout' && (
          <>
            <div className={groupCls}>
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={() =>
                    props.onUpdatePageSetup({
                      orientation: sheet.pageSetup.orientation === 'portrait' ? 'landscape' : 'portrait',
                    })
                  }
                  className={bigBtnCls}
                >
                  <FileSpreadsheet className="w-4 h-4 text-[#107c41]" />
                  <span>{sheet.pageSetup.orientation === 'portrait' ? 'Portrait' : 'Landscape'}</span>
                </button>

                <div className="relative">
                  <button type="button" onClick={() => toggleMenu('margins')} className={bigBtnCls}>
                    <Columns3 className="w-4 h-4 text-sky-600" />
                    <span className="flex items-center gap-0.5">
                      Margins <ChevronDown className="w-2.5 h-2.5" />
                    </span>
                  </button>
                  {openMenu === 'margins' && (
                    <div className="fixed mt-1 w-40 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl py-1 z-50 text-xs">
                      {(['normal', 'wide', 'narrow'] as const).map((m) => (
                        <button
                          key={m}
                          onClick={() => runAndClose(() => props.onUpdatePageSetup({ margins: m }))}
                          className={`w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 capitalize ${
                            sheet.pageSetup.margins === m ? 'font-semibold text-[#107c41]' : ''
                          }`}
                        >
                          {m}
                        </button>
                      ))}
                    </div>
                  )}
                </div>

                <button type="button" onClick={props.onSetPrintArea} className={bigBtnCls}>
                  <Grid3x3 className="w-4 h-4 text-indigo-600" />
                  <span>Set Print Area</span>
                </button>
                <button type="button" onClick={props.onClearPrintArea} className={bigBtnCls}>
                  <Eraser className="w-4 h-4 text-slate-500" />
                  <span>Clear Area</span>
                </button>
                <button type="button" onClick={props.onPrint} className={bigBtnCls}>
                  <Printer className="w-4 h-4 text-[#107c41]" />
                  <span>Print</span>
                </button>
              </div>
              <div className={groupLabelCls}>Page Setup</div>
            </div>

            <div className={`${groupCls} border-r-0`}>
              <div className="flex flex-col gap-1 px-1">
                <div className="flex items-center gap-3">
                  <label className="flex items-center gap-1.5 text-xs cursor-pointer">
                    <input
                      type="checkbox"
                      checked={sheet.pageSetup.showGridlines}
                      onChange={(e) => props.onUpdatePageSetup({ showGridlines: e.target.checked })}
                    />
                    <span>View Gridlines</span>
                  </label>
                  <label className="flex items-center gap-1.5 text-xs cursor-pointer">
                    <input
                      type="checkbox"
                      checked={sheet.pageSetup.printGridlines}
                      onChange={(e) => props.onUpdatePageSetup({ printGridlines: e.target.checked })}
                    />
                    <span>Print Gridlines</span>
                  </label>
                </div>
                <div className="flex items-center gap-3">
                  <label className="flex items-center gap-1.5 text-xs cursor-pointer">
                    <input
                      type="checkbox"
                      checked={sheet.pageSetup.showHeadings}
                      onChange={(e) => props.onUpdatePageSetup({ showHeadings: e.target.checked })}
                    />
                    <span>View Headings</span>
                  </label>
                  <label className="flex items-center gap-1.5 text-xs cursor-pointer">
                    <input
                      type="checkbox"
                      checked={sheet.pageSetup.printHeadings}
                      onChange={(e) => props.onUpdatePageSetup({ printHeadings: e.target.checked })}
                    />
                    <span>Print Headings</span>
                  </label>
                </div>
              </div>
              <div className={groupLabelCls}>Sheet Options</div>
            </div>
          </>
        )}

        {activeTab === 'formulas' && (
          <>
            <div className={groupCls}>
              <div className="flex items-center gap-1">
                <button type="button" onClick={props.onOpenFunctionWizard} className={bigBtnCls}>
                  <span className="font-serif italic font-bold text-base text-[#107c41]">fx</span>
                  <span>Insert Function</span>
                </button>
                <button type="button" onClick={() => props.onAutoSum('SUM')} className={bigBtnCls}>
                  <Sigma className="w-4 h-4 text-[#107c41]" />
                  <span>AutoSum</span>
                </button>
                {(
                  [
                    { key: 'Financial', label: 'Financial' },
                    { key: 'Logical', label: 'Logical' },
                    { key: 'Text', label: 'Text' },
                    { key: 'Date & Time', label: 'Date & Time' },
                    { key: 'Lookup & Reference', label: 'Lookup' },
                    { key: 'Math & Trig', label: 'Math' },
                    { key: 'Statistical', label: 'Stats' },
                  ] as const
                ).map((cat) => (
                  <div key={cat.key} className="relative">
                    <button type="button" onClick={() => toggleMenu(`fn-${cat.key}`)} className={bigBtnCls}>
                      <BookOpen className="w-4 h-4 text-sky-600" />
                      <span className="flex items-center gap-0.5">
                        {cat.label} <ChevronDown className="w-2.5 h-2.5" />
                      </span>
                    </button>
                    {openMenu === `fn-${cat.key}` && (
                      <div className="fixed mt-1 w-56 max-h-64 overflow-y-auto bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl py-1 z-50 text-xs">
                        {FUNCTION_CATALOG.filter((f) => f.category === cat.key).map((fn) => (
                          <button
                            key={fn.name}
                            onClick={() => runAndClose(() => props.onInsertFormulaTemplate(fn.name))}
                            className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex flex-col"
                          >
                            <span className="font-mono-num font-semibold text-[#107c41] dark:text-emerald-400">
                              {fn.name}
                            </span>
                            <span className="text-[10px] text-slate-500 truncate">{fn.description}</span>
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                ))}
              </div>
              <div className={groupLabelCls}>Function Library</div>
            </div>

            <div className={groupCls}>
              <div className="flex items-center gap-1">
                <button type="button" onClick={props.onOpenNameManager} className={bigBtnCls}>
                  <Tag className="w-4 h-4 text-amber-600" />
                  <span>Name Manager</span>
                </button>
                <button type="button" onClick={props.onDefineNameQuick} className={bigBtnCls}>
                  <Plus className="w-4 h-4 text-[#107c41]" />
                  <span>Define Name</span>
                </button>
              </div>
              <div className={groupLabelCls}>Defined Names</div>
            </div>

            <div className={groupCls}>
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={props.onToggleShowFormulas}
                  className={`${bigBtnCls} ${workbook.showFormulas ? activeBtnCls : ''}`}
                  title="Show Formulas instead of calculated results (Ctrl+`)"
                >
                  <Eye className="w-4 h-4 text-indigo-600" />
                  <span>Show Formulas</span>
                </button>
                <button type="button" onClick={props.onTracePrecedents} className={bigBtnCls} title="Highlight cells referenced by active cell">
                  <GitBranch className="w-4 h-4 text-sky-600" />
                  <span>Precedents</span>
                </button>
                <button type="button" onClick={props.onTraceDependents} className={bigBtnCls} title="Highlight cells that depend on active cell">
                  <ArrowRightToLine className="w-4 h-4 text-[#107c41]" />
                  <span>Dependents</span>
                </button>
                <button type="button" onClick={props.onClearAuditArrows} className={bigBtnCls} title="Remove auditing highlights">
                  <Eraser className="w-4 h-4 text-slate-500" />
                  <span>Clear Arrows</span>
                </button>
              </div>
              <div className={groupLabelCls}>Formula Auditing</div>
            </div>

            <div className={`${groupCls} border-r-0`}>
              <div className="flex items-center gap-1">
                <button type="button" onClick={props.onCalculateNow} className={bigBtnCls} title="Calculate Now (F9)">
                  <Calculator className="w-4 h-4 text-[#107c41]" />
                  <span>Calculate Now</span>
                </button>
                <button
                  type="button"
                  onClick={props.onToggleCalcMode}
                  className={`${bigBtnCls} ${workbook.calcMode === 'manual' ? activeBtnCls : ''}`}
                >
                  <SlidersHorizontal className="w-4 h-4 text-slate-600 dark:text-slate-300" />
                  <span>{workbook.calcMode === 'automatic' ? 'Auto Calc' : 'Manual Calc'}</span>
                </button>
              </div>
              <div className={groupLabelCls}>Calculation</div>
            </div>
          </>
        )}

        {activeTab === 'data' && (
          <>
            <div className={groupCls}>
              <div className="flex items-center gap-1">
                <button type="button" onClick={props.onImportFileClick} className={bigBtnCls}>
                  <Upload className="w-4 h-4 text-[#107c41]" />
                  <span>From XLSX/CSV</span>
                </button>
              </div>
              <div className={groupLabelCls}>Get Data</div>
            </div>

            <div className={groupCls}>
              <div className="flex items-center gap-1">
                <button type="button" onClick={() => props.onSort(true)} className={bigBtnCls}>
                  <ArrowDownAZ className="w-4 h-4 text-[#107c41]" />
                  <span>Sort A-Z</span>
                </button>
                <button type="button" onClick={() => props.onSort(false)} className={bigBtnCls}>
                  <ArrowUpZA className="w-4 h-4 text-sky-600" />
                  <span>Sort Z-A</span>
                </button>
                <button type="button" onClick={props.onOpenCustomSort} className={bigBtnCls}>
                  <SlidersHorizontal className="w-4 h-4 text-slate-600 dark:text-slate-300" />
                  <span>Custom Sort</span>
                </button>
                <button
                  type="button"
                  onClick={props.onToggleFilter}
                  className={`${bigBtnCls} ${sheet.autoFilter ? activeBtnCls : ''}`}
                >
                  <Filter className="w-4 h-4 text-[#107c41]" />
                  <span>Filter</span>
                </button>
                {sheet.autoFilter && (
                  <button type="button" onClick={props.onClearFilterValues} className={bigBtnCls}>
                    <Eraser className="w-4 h-4 text-amber-600" />
                    <span>Clear Filter</span>
                  </button>
                )}
              </div>
              <div className={groupLabelCls}>Sort & Filter</div>
            </div>

            <div className={groupCls}>
              <div className="flex items-center gap-1">
                <button type="button" onClick={props.onOpenTextToColumns} className={bigBtnCls}>
                  <SplitSquareHorizontal className="w-4 h-4 text-sky-600" />
                  <span>Text to Cols</span>
                </button>
                <button type="button" onClick={props.onOpenRemoveDuplicates} className={bigBtnCls}>
                  <Database className="w-4 h-4 text-red-600" />
                  <span>Remove Dupes</span>
                </button>
                <button type="button" onClick={props.onOpenValidationModal} className={bigBtnCls}>
                  <ShieldCheck className="w-4 h-4 text-[#107c41]" />
                  <span>Validation</span>
                </button>
                <button type="button" onClick={props.onOpenGoalSeek} className={bigBtnCls}>
                  <Target className="w-4 h-4 text-indigo-600" />
                  <span>Goal Seek</span>
                </button>
              </div>
              <div className={groupLabelCls}>Data & Forecast Tools</div>
            </div>

            <div className={`${groupCls} border-r-0`}>
              <div className="relative">
                <button type="button" onClick={() => toggleMenu('freezeData')} className={bigBtnCls}>
                  <Pin className="w-4 h-4 text-sky-600" />
                  <span className="flex items-center gap-0.5">
                    Freeze Panes <ChevronDown className="w-2.5 h-2.5" />
                  </span>
                </button>
                {openMenu === 'freezeData' && (
                  <div className="fixed mt-1 w-48 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl py-1 z-50 text-xs">
                    <button onClick={() => runAndClose(() => props.onSetFreezePanes('selection'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Freeze at Current Cell</button>
                    <button onClick={() => runAndClose(() => props.onSetFreezePanes('topRow'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Freeze Top Row</button>
                    <button onClick={() => runAndClose(() => props.onSetFreezePanes('firstCol'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Freeze First Column</button>
                    <button onClick={() => runAndClose(() => props.onSetFreezePanes('unfreeze'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 text-red-600">Unfreeze Panes</button>
                  </div>
                )}
              </div>
              <div className={groupLabelCls}>Outline</div>
            </div>
          </>
        )}

        {activeTab === 'review' && (
          <>
            <div className={groupCls}>
              <div className="flex items-center gap-1">
                <button type="button" onClick={props.onOpenCommentModal} className={bigBtnCls}>
                  <MessageSquarePlus className="w-4 h-4 text-[#107c41]" />
                  <span>New Comment</span>
                </button>
                <button type="button" onClick={props.onDeleteComment} className={bigBtnCls}>
                  <MessageSquareX className="w-4 h-4 text-red-600" />
                  <span>Delete</span>
                </button>
                <button
                  type="button"
                  onClick={props.onToggleShowAllComments}
                  className={`${bigBtnCls} ${props.showAllComments ? activeBtnCls : ''}`}
                >
                  <Eye className="w-4 h-4 text-amber-600" />
                  <span>Show All</span>
                </button>
              </div>
              <div className={groupLabelCls}>Comments</div>
            </div>

            <div className={`${groupCls} border-r-0`}>
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={props.onToggleProtectSheet}
                  className={`${bigBtnCls} ${sheet.protected ? activeBtnCls : ''}`}
                >
                  {sheet.protected ? (
                    <Unlock className="w-4 h-4 text-amber-600" />
                  ) : (
                    <Lock className="w-4 h-4 text-[#107c41]" />
                  )}
                  <span>{sheet.protected ? 'Unprotect Sheet' : 'Protect Sheet'}</span>
                </button>
                <button type="button" onClick={props.onOpenWorkbookStats} className={bigBtnCls}>
                  <BarChart3 className="w-4 h-4 text-sky-600" />
                  <span>Statistics</span>
                </button>
              </div>
              <div className={groupLabelCls}>Protect & Inspect</div>
            </div>
          </>
        )}

        {activeTab === 'view' && (
          <>
            <div className={groupCls}>
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={() => props.pageBreakView && props.onTogglePageBreakView()}
                  className={`${bigBtnCls} ${!props.pageBreakView ? activeBtnCls : ''}`}
                >
                  <Grid3x3 className="w-4 h-4 text-[#107c41]" />
                  <span>Normal</span>
                </button>
                <button
                  type="button"
                  onClick={() => !props.pageBreakView && props.onTogglePageBreakView()}
                  className={`${bigBtnCls} ${props.pageBreakView ? activeBtnCls : ''}`}
                >
                  <FileSpreadsheet className="w-4 h-4 text-sky-600" />
                  <span>Page Break</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    if (!document.fullscreenElement) document.documentElement.requestFullscreen?.();
                    else document.exitFullscreen?.();
                  }}
                  className={bigBtnCls}
                >
                  <Maximize2 className="w-4 h-4 text-slate-600 dark:text-slate-300" />
                  <span>Fullscreen</span>
                </button>
              </div>
              <div className={groupLabelCls}>Workbook Views</div>
            </div>

            <div className={groupCls}>
              <div className="flex flex-col gap-1 px-2">
                <label className="flex items-center gap-1.5 text-xs cursor-pointer">
                  <input
                    type="checkbox"
                    checked={workbook.showFormulaBar}
                    onChange={props.onToggleFormulaBar}
                  />
                  <span>Formula Bar</span>
                </label>
                <label className="flex items-center gap-1.5 text-xs cursor-pointer">
                  <input
                    type="checkbox"
                    checked={sheet.pageSetup.showGridlines}
                    onChange={(e) => props.onUpdatePageSetup({ showGridlines: e.target.checked })}
                  />
                  <span>Gridlines</span>
                </label>
                <label className="flex items-center gap-1.5 text-xs cursor-pointer">
                  <input
                    type="checkbox"
                    checked={sheet.pageSetup.showHeadings}
                    onChange={(e) => props.onUpdatePageSetup({ showHeadings: e.target.checked })}
                  />
                  <span>Headings</span>
                </label>
              </div>
              <div className={groupLabelCls}>Show</div>
            </div>

            <div className={groupCls}>
              <div className="flex items-center gap-1">
                {[75, 100, 125, 150].map((z) => (
                  <button
                    key={z}
                    type="button"
                    onClick={() => props.onSetZoom(z)}
                    className={`${bigBtnCls} ${workbook.zoom === z ? activeBtnCls : ''}`}
                  >
                    <ZoomIn className="w-3.5 h-3.5" />
                    <span className="font-mono-num">{z}%</span>
                  </button>
                ))}
              </div>
              <div className={groupLabelCls}>Zoom</div>
            </div>

            <div className={`${groupCls} border-r-0`}>
              <div className="flex items-center gap-1">
                <div className="relative">
                  <button type="button" onClick={() => toggleMenu('freezeView')} className={bigBtnCls}>
                    <Pin className="w-4 h-4 text-sky-600" />
                    <span>Freeze Panes</span>
                  </button>
                  {openMenu === 'freezeView' && (
                    <div className="fixed mt-1 w-48 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-xl py-1 z-50 text-xs">
                      <button onClick={() => runAndClose(() => props.onSetFreezePanes('selection'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Freeze Panes at Selection</button>
                      <button onClick={() => runAndClose(() => props.onSetFreezePanes('topRow'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Freeze Top Row</button>
                      <button onClick={() => runAndClose(() => props.onSetFreezePanes('firstCol'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700">Freeze First Column</button>
                      <button onClick={() => runAndClose(() => props.onSetFreezePanes('unfreeze'))} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 text-red-600">Unfreeze Panes</button>
                    </div>
                  )}
                </div>
                <button type="button" onClick={props.onToggleTheme} className={bigBtnCls}>
                  <SlidersHorizontal className="w-4 h-4 text-[#107c41]" />
                  <span>{workbook.theme === 'dark' ? 'Light Mode' : 'Dark Mode'}</span>
                </button>
              </div>
              <div className={groupLabelCls}>Window</div>
            </div>
          </>
        )}
      </div>
    </div>
  );
};
