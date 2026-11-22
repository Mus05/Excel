import React, { useMemo, useState } from 'react';
import {
  Download,
  FilePlus2,
  FileSpreadsheet,
  FolderOpen,
  Home,
  Printer,
  Search,
  Sparkles,
  Trash2,
  X,
} from 'lucide-react';
import {
  CellStyle,
  ConditionalRule,
  ConditionalRuleType,
  DataValidationRule,
  NamedRange,
  NumberFormatCategory,
  Sheet,
  Workbook,
} from '../types/spreadsheet';
import { FUNCTION_CATALOG, WorkbookEvaluator, solveGoalSeek } from '../engine/formulaEngine';
import { WORKBOOK_TEMPLATES } from '../utils/templates';
import {
  cellKey,
  colIndexToLetter,
  formatRangeRef,
  normalizeRange,
  parseCellRef,
  parseRangeRef,
} from '../utils/cellCoords';

const ModalShell: React.FC<{
  title: string;
  onClose: () => void;
  widthClass?: string;
  children: React.ReactNode;
}> = ({ title, onClose, widthClass = 'max-w-lg', children }) => (
  <div
    onClick={onClose}
    className="no-print fixed inset-0 z-50 bg-black/40 backdrop-blur-[1px] flex items-center justify-center p-4"
  >
    <div
      onClick={(e) => e.stopPropagation()}
      className={`w-full ${widthClass} bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded-lg shadow-2xl overflow-hidden flex flex-col max-h-[88vh]`}
    >
      <div className="flex items-center justify-between px-4 py-3 bg-slate-50 dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700">
        <h3 className="font-display font-semibold text-sm text-slate-900 dark:text-white">{title}</h3>
        <button
          type="button"
          onClick={onClose}
          className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-500"
        >
          <X className="w-4 h-4" />
        </button>
      </div>
      <div className="p-4 overflow-y-auto text-xs text-slate-700 dark:text-slate-200">{children}</div>
    </div>
  </div>
);

/* 1. FILE BACKSTAGE OVERLAY */
export const FileBackstageModal: React.FC<{
  workbook: Workbook;
  onClose: () => void;
  onLoadWorkbook: (wb: Workbook) => void;
  onImportFile: () => void;
  onExportXLSX: () => void;
  onExportCSV: () => void;
  onExportJSON: () => void;
  onPrint: () => void;
  onGoHome: () => void;
}> = ({
  workbook,
  onClose,
  onLoadWorkbook,
  onImportFile,
  onExportXLSX,
  onExportCSV,
  onExportJSON,
  onPrint,
  onGoHome,
}) => {
  const stats = useMemo(() => {
    let cells = 0;
    let formulas = 0;
    let charts = 0;
    for (const s of workbook.sheets) {
      charts += s.charts.length;
      for (const c of Object.values(s.cells)) {
        if (c.raw !== '') cells++;
        if (c.raw.startsWith('=')) formulas++;
      }
    }
    return { sheets: workbook.sheets.length, cells, formulas, charts };
  }, [workbook]);

  return (
    <div className="no-print fixed inset-0 z-50 bg-black/50 flex">
      <div className="w-60 bg-[#107c41] text-white flex flex-col py-4 select-none shrink-0">
        <div className="px-5 pb-4 mb-2 border-b border-white/15 flex items-center justify-between">
          <span className="font-display font-bold text-base">File</span>
          <button onClick={onClose} className="p-1 rounded hover:bg-white/15">
            <X className="w-4 h-4" />
          </button>
        </div>
        <button onClick={() => { onClose(); onGoHome(); }} className="px-5 py-2.5 text-left text-xs font-medium hover:bg-white/15 flex items-center gap-2.5">
          <Home className="w-4 h-4" /> Home Dashboard
        </button>
        <button onClick={onImportFile} className="px-5 py-2.5 text-left text-xs font-medium hover:bg-white/15 flex items-center gap-2.5">
          <FolderOpen className="w-4 h-4" /> Open / Import File...
        </button>
        <button onClick={() => { onExportXLSX(); onClose(); }} className="px-5 py-2.5 text-left text-xs font-medium hover:bg-white/15 flex items-center gap-2.5">
          <Download className="w-4 h-4" /> Save As Excel (.xlsx)
        </button>
        <button onClick={() => { onExportCSV(); onClose(); }} className="px-5 py-2.5 text-left text-xs font-medium hover:bg-white/15 flex items-center gap-2.5">
          <FileSpreadsheet className="w-4 h-4" /> Export Active Sheet (.csv)
        </button>
        <button onClick={() => { onExportJSON(); onClose(); }} className="px-5 py-2.5 text-left text-xs font-medium hover:bg-white/15 flex items-center gap-2.5">
          <Download className="w-4 h-4" /> Export Workbook (.json)
        </button>
        <button onClick={() => { onClose(); setTimeout(onPrint, 120); }} className="px-5 py-2.5 text-left text-xs font-medium hover:bg-white/15 flex items-center gap-2.5">
          <Printer className="w-4 h-4" /> Print Sheet
        </button>

        <div className="mt-auto px-5 pt-4 border-t border-white/15 text-[11px] text-white/80 space-y-1">
          <div className="font-semibold text-white">Workbook Info</div>
          <div>Sheets: {stats.sheets}</div>
          <div>Populated Cells: {stats.cells}</div>
          <div>Formulas: {stats.formulas}</div>
          <div>Charts: {stats.charts}</div>
        </div>
      </div>

      <div className="flex-1 bg-slate-50 dark:bg-slate-900 p-8 overflow-y-auto" onClick={onClose}>
        <div className="max-w-4xl mx-auto" onClick={(e) => e.stopPropagation()}>
          <div className="flex items-center justify-between mb-6">
            <div>
              <h2 className="font-display font-bold text-xl text-slate-900 dark:text-white">
                New Workbook & Templates
              </h2>
              <p className="text-xs text-slate-500 mt-1">
                Choose a blank grid or start with a ready-made spreadsheet template.
              </p>
            </div>
            <button
              onClick={onClose}
              className="px-4 py-1.5 text-xs font-medium border border-slate-300 dark:border-slate-700 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
            >
              Back to Sheet
            </button>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
            {WORKBOOK_TEMPLATES.map((tpl) => (
              <button
                key={tpl.id}
                onClick={() => {
                  onLoadWorkbook(tpl.build());
                  onClose();
                }}
                className="text-left p-4 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg hover:border-[#107c41] hover:shadow-md transition-all flex flex-col gap-2"
              >
                <div className="flex items-center gap-2 text-[#107c41] dark:text-emerald-400 font-display font-semibold text-sm">
                  <FilePlus2 className="w-4 h-4 shrink-0" />
                  <span>{tpl.name}</span>
                </div>
                <p className="text-xs text-slate-500 dark:text-slate-400 leading-relaxed">
                  {tpl.description}
                </p>
              </button>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};

/* 2. FIND & REPLACE MODAL */
export const FindReplaceModal: React.FC<{
  sheet: Sheet;
  evaluator: WorkbookEvaluator;
  onClose: () => void;
  onJumpToCell: (row: number, col: number) => void;
  onReplaceCell: (row: number, col: number, newRaw: string) => void;
}> = ({ sheet, evaluator, onClose, onJumpToCell, onReplaceCell }) => {
  const [findText, setFindText] = useState('');
  const [replaceText, setReplaceText] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [entireCell, setEntireCell] = useState(false);
  const [matchIdx, setMatchIdx] = useState(0);
  const [statusMsg, setStatusMsg] = useState('');

  const matches = useMemo(() => {
    if (!findText) return [];
    const found: { row: number; col: number; raw: string; display: string }[] = [];
    const needle = matchCase ? findText : findText.toLowerCase();

    for (const [k, cell] of Object.entries(sheet.cells)) {
      if (!cell.raw) continue;
      const [r, c] = k.split(',').map(Number);
      const ev = evaluator.evaluateCell(sheet.id, r, c);
      const hayRaw = matchCase ? cell.raw : cell.raw.toLowerCase();
      const hayDisp = matchCase ? ev.formatted : ev.formatted.toLowerCase();

      const ok = entireCell
        ? hayRaw === needle || hayDisp === needle
        : hayRaw.includes(needle) || hayDisp.includes(needle);
      if (ok) found.push({ row: r, col: c, raw: cell.raw, display: ev.formatted });
    }
    return found.sort((a, b) => (a.row !== b.row ? a.row - b.row : a.col - b.col));
  }, [findText, matchCase, entireCell, sheet.cells, sheet.id, evaluator]);

  const handleFindNext = () => {
    if (matches.length === 0) {
      setStatusMsg('No matches found.');
      return;
    }
    const idx = matchIdx % matches.length;
    const target = matches[idx];
    onJumpToCell(target.row, target.col);
    setMatchIdx((idx + 1) % matches.length);
    setStatusMsg(`Match ${idx + 1} of ${matches.length} (${colIndexToLetter(target.col)}${target.row + 1})`);
  };

  const handleReplaceOne = () => {
    if (matches.length === 0) return;
    const idx = Math.max(0, (matchIdx - 1 + matches.length) % matches.length);
    const target = matches[idx];
    const flags = matchCase ? 'g' : 'gi';
    const escaped = findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const updated = target.raw.replace(new RegExp(escaped, flags), replaceText);
    onReplaceCell(target.row, target.col, updated);
    setStatusMsg(`Replaced in ${colIndexToLetter(target.col)}${target.row + 1}`);
  };

  const handleReplaceAll = () => {
    if (matches.length === 0) {
      setStatusMsg('No matches to replace.');
      return;
    }
    const flags = matchCase ? 'g' : 'gi';
    const escaped = findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(escaped, flags);
    let count = 0;
    for (const m of matches) {
      const updated = m.raw.replace(re, replaceText);
      if (updated !== m.raw) {
        onReplaceCell(m.row, m.col, updated);
        count++;
      }
    }
    setStatusMsg(`Replaced ${count} occurrence(s).`);
  };

  return (
    <ModalShell title="Find and Replace" onClose={onClose} widthClass="max-w-md">
      <div className="space-y-3">
        <div>
          <label className="block font-medium mb-1">Find what</label>
          <input
            autoFocus
            type="text"
            value={findText}
            onChange={(e) => { setFindText(e.target.value); setMatchIdx(0); }}
            onKeyDown={(e) => { if (e.key === 'Enter') handleFindNext(); }}
            className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-transparent"
          />
        </div>
        <div>
          <label className="block font-medium mb-1">Replace with</label>
          <input
            type="text"
            value={replaceText}
            onChange={(e) => setReplaceText(e.target.value)}
            className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-transparent"
          />
        </div>
        <div className="flex items-center gap-4">
          <label className="flex items-center gap-1.5 cursor-pointer">
            <input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} />
            <span>Match case</span>
          </label>
          <label className="flex items-center gap-1.5 cursor-pointer">
            <input type="checkbox" checked={entireCell} onChange={(e) => setEntireCell(e.target.checked)} />
            <span>Match entire cell</span>
          </label>
        </div>
        {statusMsg && <div className="text-[11px] text-[#107c41] dark:text-emerald-400 font-medium">{statusMsg}</div>}
        <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100 dark:border-slate-800">
          <button onClick={handleFindNext} className="px-3 py-1.5 bg-[#107c41] text-white rounded font-medium hover:bg-[#0c5e31]">Find Next</button>
          <button onClick={handleReplaceOne} className="px-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded hover:bg-slate-100 dark:hover:bg-slate-800">Replace</button>
          <button onClick={handleReplaceAll} className="px-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded hover:bg-slate-100 dark:hover:bg-slate-800">Replace All</button>
        </div>
      </div>
    </ModalShell>
  );
};

/* 3. INSERT FUNCTION WIZARD */
export const FunctionWizardModal: React.FC<{
  onClose: () => void;
  onInsertFunction: (fnName: string) => void;
}> = ({ onClose, onInsertFunction }) => {
  const [query, setQuery] = useState('');
  const [cat, setCat] = useState<string>('All');
  const [selected, setSelected] = useState(FUNCTION_CATALOG[0]);

  const categories = ['All', 'Math & Trig', 'Statistical', 'Logical', 'Lookup & Reference', 'Text', 'Date & Time', 'Financial', 'Information'];

  const filtered = useMemo(() => {
    return FUNCTION_CATALOG.filter((f) => {
      const catOk = cat === 'All' || f.category === cat;
      const q = query.trim().toLowerCase();
      const qOk = !q || f.name.toLowerCase().includes(q) || f.description.toLowerCase().includes(q);
      return catOk && qOk;
    });
  }, [query, cat]);

  return (
    <ModalShell title="Insert Function (fx)" onClose={onClose} widthClass="max-w-xl">
      <div className="space-y-3">
        <div className="grid grid-cols-2 gap-2">
          <input
            autoFocus
            type="text"
            placeholder="Search function (e.g. VLOOKUP, SUMIF, PMT)..."
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            className="px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-transparent"
          />
          <select
            value={cat}
            onChange={(e) => setCat(e.target.value)}
            className="px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-800"
          >
            {categories.map((c) => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
        </div>

        <div className="h-52 overflow-y-auto border border-slate-200 dark:border-slate-700 rounded divide-y divide-slate-100 dark:divide-slate-800">
          {filtered.map((fn) => (
            <button
              key={fn.name}
              onClick={() => setSelected(fn)}
              onDoubleClick={() => { onInsertFunction(fn.name); onClose(); }}
              className={`w-full px-3 py-2 text-left flex items-center justify-between ${
                selected.name === fn.name
                  ? 'bg-emerald-50 dark:bg-emerald-900/40'
                  : 'hover:bg-slate-50 dark:hover:bg-slate-800'
              }`}
            >
              <span className="font-mono-num font-semibold text-[#107c41] dark:text-emerald-400">{fn.name}</span>
              <span className="text-[11px] text-slate-400">{fn.category}</span>
            </button>
          ))}
        </div>

        <div className="p-3 bg-slate-50 dark:bg-slate-800 rounded border border-slate-200 dark:border-slate-700 space-y-1">
          <div className="font-mono-num font-semibold text-slate-900 dark:text-white">{selected.syntax}</div>
          <div className="text-slate-600 dark:text-slate-300">{selected.description}</div>
          <div className="font-mono-num text-[11px] text-slate-500">Example: {selected.example}</div>
        </div>

        <div className="flex justify-end gap-2">
          <button onClick={onClose} className="px-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded">Cancel</button>
          <button
            onClick={() => { onInsertFunction(selected.name); onClose(); }}
            className="px-4 py-1.5 bg-[#107c41] text-white rounded font-medium hover:bg-[#0c5e31]"
          >
            Insert {selected.name}()
          </button>
        </div>
      </div>
    </ModalShell>
  );
};

/* 4. FORMAT CELLS MODAL */
export const FormatCellsModal: React.FC<{
  initialStyle: CellStyle;
  onClose: () => void;
  onApply: (patch: Partial<CellStyle>) => void;
}> = ({ initialStyle, onClose, onApply }) => {
  const [tab, setTab] = useState<'number' | 'alignment' | 'font' | 'fill'>('number');
  const [draft, setDraft] = useState<CellStyle>({ ...initialStyle });

  return (
    <ModalShell title="Format Cells" onClose={onClose} widthClass="max-w-lg">
      <div className="flex gap-1 border-b border-slate-200 dark:border-slate-700 mb-3">
        {(['number', 'alignment', 'font', 'fill'] as const).map((t) => (
          <button
            key={t}
            onClick={() => setTab(t)}
            className={`px-3 py-1.5 capitalize font-medium border-b-2 ${
              tab === t
                ? 'border-[#107c41] text-[#107c41]'
                : 'border-transparent text-slate-500 hover:text-slate-800'
            }`}
          >
            {t}
          </button>
        ))}
      </div>

      {tab === 'number' && (
        <div className="space-y-3">
          <div>
            <label className="block font-medium mb-1">Category</label>
            <select
              value={draft.numberFormat?.category || 'general'}
              onChange={(e) =>
                setDraft({
                  ...draft,
                  numberFormat: { ...draft.numberFormat, category: e.target.value as NumberFormatCategory },
                })
              }
              className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-800"
            >
              <option value="general">General</option>
              <option value="number">Number</option>
              <option value="currency">Currency</option>
              <option value="accounting">Accounting</option>
              <option value="date">Short Date</option>
              <option value="longDate">Long Date</option>
              <option value="time">Time</option>
              <option value="percentage">Percentage</option>
              <option value="fraction">Fraction</option>
              <option value="scientific">Scientific</option>
              <option value="text">Text</option>
            </select>
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="block font-medium mb-1">Decimal places</label>
              <input
                type="number"
                min={0}
                max={10}
                value={draft.numberFormat?.decimals ?? 2}
                onChange={(e) =>
                  setDraft({
                    ...draft,
                    numberFormat: {
                      category: draft.numberFormat?.category || 'number',
                      ...draft.numberFormat,
                      decimals: Number(e.target.value),
                    },
                  })
                }
                className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-transparent"
              />
            </div>
            <div>
              <label className="block font-medium mb-1">Currency Symbol</label>
              <select
                value={draft.numberFormat?.currencySymbol || 'RM'}
                onChange={(e) =>
                  setDraft({
                    ...draft,
                    numberFormat: {
                      category: draft.numberFormat?.category || 'currency',
                      ...draft.numberFormat,
                      currencySymbol: e.target.value,
                    },
                  })
                }
                className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-800"
              >
                <option value="RM">RM (MYR)</option>
                <option value="$">$ (USD)</option>
                <option value="SGD">SGD</option>
                <option value="€">€ (EUR)</option>
                <option value="£">£ (GBP)</option>
              </select>
            </div>
          </div>
        </div>
      )}

      {tab === 'alignment' && (
        <div className="space-y-3">
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="block font-medium mb-1">Horizontal</label>
              <select
                value={draft.horizontalAlign || 'general'}
                onChange={(e) => setDraft({ ...draft, horizontalAlign: e.target.value as any })}
                className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-800"
              >
                <option value="general">General</option>
                <option value="left">Left</option>
                <option value="center">Center</option>
                <option value="right">Right</option>
              </select>
            </div>
            <div>
              <label className="block font-medium mb-1">Vertical</label>
              <select
                value={draft.verticalAlign || 'middle'}
                onChange={(e) => setDraft({ ...draft, verticalAlign: e.target.value as any })}
                className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-800"
              >
                <option value="top">Top</option>
                <option value="middle">Middle</option>
                <option value="bottom">Bottom</option>
              </select>
            </div>
          </div>
          <label className="flex items-center gap-2 cursor-pointer">
            <input
              type="checkbox"
              checked={Boolean(draft.wrapText)}
              onChange={(e) => setDraft({ ...draft, wrapText: e.target.checked })}
            />
            <span>Wrap text</span>
          </label>
        </div>
      )}

      {tab === 'font' && (
        <div className="space-y-3">
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="block font-medium mb-1">Size (px)</label>
              <input
                type="number"
                min={8}
                max={72}
                value={draft.fontSize || 12}
                onChange={(e) => setDraft({ ...draft, fontSize: Number(e.target.value) })}
                className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-transparent"
              />
            </div>
            <div>
              <label className="block font-medium mb-1">Text Color</label>
              <input
                type="color"
                value={draft.color || '#0f172a'}
                onChange={(e) => setDraft({ ...draft, color: e.target.value })}
                className="w-full h-8 border border-slate-300 dark:border-slate-600 rounded cursor-pointer"
              />
            </div>
          </div>
          <div className="flex flex-wrap gap-4">
            <label className="flex items-center gap-1.5"><input type="checkbox" checked={Boolean(draft.bold)} onChange={(e) => setDraft({ ...draft, bold: e.target.checked })} /> Bold</label>
            <label className="flex items-center gap-1.5"><input type="checkbox" checked={Boolean(draft.italic)} onChange={(e) => setDraft({ ...draft, italic: e.target.checked })} /> Italic</label>
            <label className="flex items-center gap-1.5"><input type="checkbox" checked={Boolean(draft.underline)} onChange={(e) => setDraft({ ...draft, underline: e.target.checked })} /> Underline</label>
            <label className="flex items-center gap-1.5"><input type="checkbox" checked={Boolean(draft.strikethrough)} onChange={(e) => setDraft({ ...draft, strikethrough: e.target.checked })} /> Strikethrough</label>
          </div>
        </div>
      )}

      {tab === 'fill' && (
        <div className="space-y-3">
          <div>
            <label className="block font-medium mb-1">Background Fill Color</label>
            <div className="flex items-center gap-3">
              <input
                type="color"
                value={draft.backgroundColor || '#ffffff'}
                onChange={(e) => setDraft({ ...draft, backgroundColor: e.target.value })}
                className="w-16 h-8 border border-slate-300 dark:border-slate-600 rounded cursor-pointer"
              />
              <button
                type="button"
                onClick={() => setDraft({ ...draft, backgroundColor: undefined })}
                className="px-3 py-1 border border-slate-300 dark:border-slate-600 rounded"
              >
                Clear Fill
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="flex justify-end gap-2 mt-4 pt-3 border-t border-slate-100 dark:border-slate-800">
        <button onClick={onClose} className="px-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded">Cancel</button>
        <button
          onClick={() => { onApply(draft); onClose(); }}
          className="px-4 py-1.5 bg-[#107c41] text-white rounded font-medium hover:bg-[#0c5e31]"
        >
          Apply
        </button>
      </div>
    </ModalShell>
  );
};

/* 5. CONDITIONAL FORMATTING MODAL */
export const ConditionalFormatModal: React.FC<{
  rules: ConditionalRule[];
  defaultRange: string;
  initialType?: ConditionalRuleType;
  onClose: () => void;
  onAddRule: (rule: ConditionalRule) => void;
  onDeleteRule: (id: string) => void;
}> = ({ rules, defaultRange, initialType, onClose, onAddRule, onDeleteRule }) => {
  const [range, setRange] = useState(defaultRange);
  const [type, setType] = useState<ConditionalRuleType>(initialType || 'greaterThan');
  const [value1, setValue1] = useState('1000');
  const [value2, setValue2] = useState('5000');
  const [bgColor, setBgColor] = useState('#dcfce7');
  const [textColor, setTextColor] = useState('#166534');

  const handleAdd = () => {
    onAddRule({
      id: `cf-${Date.now()}`,
      range: range.toUpperCase(),
      type,
      value1,
      value2,
      backgroundColor: bgColor,
      textColor,
      barColor: '#107c41',
    });
  };

  return (
    <ModalShell title="Conditional Formatting Rules" onClose={onClose} widthClass="max-w-lg">
      <div className="space-y-4">
        <div className="p-3 bg-slate-50 dark:bg-slate-800 rounded border border-slate-200 dark:border-slate-700 space-y-2.5">
          <div className="font-semibold text-slate-800 dark:text-slate-100">Add New Rule</div>
          <div className="grid grid-cols-2 gap-2">
            <div>
              <label className="block text-[11px] text-slate-500 mb-0.5">Applies to Range</label>
              <input
                type="text"
                value={range}
                onChange={(e) => setRange(e.target.value)}
                className="w-full px-2 py-1 font-mono-num border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-900"
              />
            </div>
            <div>
              <label className="block text-[11px] text-slate-500 mb-0.5">Rule Type</label>
              <select
                value={type}
                onChange={(e) => setType(e.target.value as ConditionalRuleType)}
                className="w-full px-2 py-1 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-900"
              >
                <option value="greaterThan">Greater Than</option>
                <option value="lessThan">Less Than</option>
                <option value="between">Between</option>
                <option value="equal">Equal To</option>
                <option value="textContains">Text Contains</option>
                <option value="duplicate">Duplicate Values</option>
                <option value="dataBar">Data Bar</option>
                <option value="colorScale">Color Scale (Red-Yellow-Green)</option>
              </select>
            </div>
          </div>

          {!['duplicate', 'dataBar', 'colorScale'].includes(type) && (
            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="block text-[11px] text-slate-500 mb-0.5">Value</label>
                <input
                  type="text"
                  value={value1}
                  onChange={(e) => setValue1(e.target.value)}
                  className="w-full px-2 py-1 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-900"
                />
              </div>
              {type === 'between' && (
                <div>
                  <label className="block text-[11px] text-slate-500 mb-0.5">And</label>
                  <input
                    type="text"
                    value={value2}
                    onChange={(e) => setValue2(e.target.value)}
                    className="w-full px-2 py-1 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-900"
                  />
                </div>
              )}
            </div>
          )}

          {!['dataBar', 'colorScale'].includes(type) && (
            <div className="flex items-center gap-4">
              <label className="flex items-center gap-1.5">
                <span>Fill:</span>
                <input type="color" value={bgColor} onChange={(e) => setBgColor(e.target.value)} className="w-8 h-6 rounded" />
              </label>
              <label className="flex items-center gap-1.5">
                <span>Text:</span>
                <input type="color" value={textColor} onChange={(e) => setTextColor(e.target.value)} className="w-8 h-6 rounded" />
              </label>
            </div>
          )}

          <button
            onClick={handleAdd}
            className="px-3 py-1.5 bg-[#107c41] text-white rounded font-medium hover:bg-[#0c5e31]"
          >
            Add Rule
          </button>
        </div>

        <div>
          <div className="font-semibold mb-1.5">Active Rules ({rules.length})</div>
          {rules.length === 0 ? (
            <div className="text-slate-400 py-3 text-center border border-dashed border-slate-200 dark:border-slate-700 rounded">
              No conditional formatting rules on this sheet.
            </div>
          ) : (
            <div className="space-y-1.5 max-h-40 overflow-y-auto">
              {rules.map((r) => (
                <div
                  key={r.id}
                  className="flex items-center justify-between px-3 py-2 border border-slate-200 dark:border-slate-700 rounded"
                >
                  <div>
                    <span className="font-semibold capitalize">{r.type}</span>
                    {r.value1 && <span className="ml-1.5 text-slate-500">({r.value1}{r.type === 'between' ? ` – ${r.value2}` : ''})</span>}
                    <div className="font-mono-num text-[11px] text-slate-400">Range: {r.range}</div>
                  </div>
                  <button onClick={() => onDeleteRule(r.id)} className="p-1 text-red-600 hover:bg-red-50 rounded">
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </ModalShell>
  );
};

/* 6. DATA VALIDATION MODAL */
export const DataValidationModal: React.FC<{
  existing?: DataValidationRule;
  onClose: () => void;
  onSave: (rule: DataValidationRule | undefined) => void;
}> = ({ existing, onClose, onSave }) => {
  const [type, setType] = useState<DataValidationRule['type']>(existing?.type || 'list');
  const [operator, setOperator] = useState<DataValidationRule['operator']>(existing?.operator || 'between');
  const [formula1, setFormula1] = useState(existing?.formula1 || 'Approved,Pending,Rejected');
  const [formula2, setFormula2] = useState(existing?.formula2 || '100');
  const [inputTitle, setInputTitle] = useState(existing?.inputTitle || '');
  const [inputMessage, setInputMessage] = useState(existing?.inputMessage || '');

  return (
    <ModalShell title="Data Validation" onClose={onClose} widthClass="max-w-md">
      <div className="space-y-3">
        <div>
          <label className="block font-medium mb-1">Allow</label>
          <select
            value={type}
            onChange={(e) => setType(e.target.value as DataValidationRule['type'])}
            className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-800"
          >
            <option value="list">List (Dropdown Menu)</option>
            <option value="whole">Whole Number</option>
            <option value="decimal">Decimal</option>
            <option value="textLength">Text Length</option>
          </select>
        </div>

        {type !== 'list' && (
          <div>
            <label className="block font-medium mb-1">Condition</label>
            <select
              value={operator}
              onChange={(e) => setOperator(e.target.value as any)}
              className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-800"
            >
              <option value="between">Between</option>
              <option value="greaterThan">Greater Than</option>
              <option value="lessThan">Less Than</option>
              <option value="equal">Equal To</option>
            </select>
          </div>
        )}

        <div>
          <label className="block font-medium mb-1">
            {type === 'list' ? 'Source Items (comma separated)' : operator === 'between' ? 'Minimum' : 'Target Value'}
          </label>
          <input
            type="text"
            value={formula1}
            onChange={(e) => setFormula1(e.target.value)}
            className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-transparent"
          />
        </div>

        {type !== 'list' && operator === 'between' && (
          <div>
            <label className="block font-medium mb-1">Maximum</label>
            <input
              type="text"
              value={formula2}
              onChange={(e) => setFormula2(e.target.value)}
              className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-transparent"
            />
          </div>
        )}

        <div className="pt-2 border-t border-slate-100 dark:border-slate-800 space-y-2">
          <div className="font-medium text-slate-600 dark:text-slate-300">Input Prompt (Optional)</div>
          <input
            type="text"
            placeholder="Title"
            value={inputTitle}
            onChange={(e) => setInputTitle(e.target.value)}
            className="w-full px-2.5 py-1 border border-slate-300 dark:border-slate-600 rounded bg-transparent"
          />
          <input
            type="text"
            placeholder="Message shown when cell is selected"
            value={inputMessage}
            onChange={(e) => setInputMessage(e.target.value)}
            className="w-full px-2.5 py-1 border border-slate-300 dark:border-slate-600 rounded bg-transparent"
          />
        </div>

        <div className="flex justify-between pt-2">
          <button
            onClick={() => { onSave(undefined); onClose(); }}
            className="px-3 py-1.5 text-red-600 border border-red-200 rounded hover:bg-red-50"
          >
            Clear Validation
          </button>
          <div className="flex gap-2">
            <button onClick={onClose} className="px-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded">Cancel</button>
            <button
              onClick={() => {
                onSave({ type, operator, formula1, formula2, showDropdown: true, inputTitle, inputMessage });
                onClose();
              }}
              className="px-4 py-1.5 bg-[#107c41] text-white rounded font-medium hover:bg-[#0c5e31]"
            >
              OK
            </button>
          </div>
        </div>
      </div>
    </ModalShell>
  );
};

/* 7. PIVOT TABLE BUILDER MODAL */
export const PivotTableModal: React.FC<{
  sheet: Sheet;
  evaluator: WorkbookEvaluator;
  defaultRange: string;
  onClose: () => void;
  onCreatePivot: (config: {
    headers: string[];
    rows: (string | number)[][];
    targetNewSheet: boolean;
  }) => void;
}> = ({ sheet, evaluator, defaultRange, onClose, onCreatePivot }) => {
  const [sourceRange, setSourceRange] = useState(defaultRange);

  const columnsMeta = useMemo(() => {
    const rng = parseRangeRef(sourceRange);
    if (!rng) return [];
    const cols: { colIdx: number; label: string }[] = [];
    for (let c = rng.start.col; c <= rng.end.col; c++) {
      const ev = evaluator.evaluateCell(sheet.id, rng.start.row, c);
      cols.push({ colIdx: c, label: ev.formatted || `Col ${colIndexToLetter(c)}` });
    }
    return cols;
  }, [sourceRange, sheet.id, evaluator]);

  const [rowField, setRowField] = useState<number>(columnsMeta[0]?.colIdx ?? 0);
  const [valField, setValField] = useState<number>(columnsMeta[columnsMeta.length - 1]?.colIdx ?? 1);
  const [agg, setAgg] = useState<'SUM' | 'COUNT' | 'AVERAGE' | 'MIN' | 'MAX'>('SUM');
  const [newSheet, setNewSheet] = useState(true);

  const handleGenerate = () => {
    const rng = parseRangeRef(sourceRange);
    if (!rng) return;
    const groups = new Map<string, number[]>();

    for (let r = rng.start.row + 1; r <= rng.end.row; r++) {
      const keyEv = evaluator.evaluateCell(sheet.id, r, rowField);
      const groupKey = keyEv.formatted || '(Blank)';
      const valEv = evaluator.evaluateCell(sheet.id, r, valField);
      const numVal = typeof valEv.value === 'number' ? valEv.value : Number(valEv.value);
      if (!groups.has(groupKey)) groups.set(groupKey, []);
      if (agg === 'COUNT' || !isNaN(numVal)) {
        groups.get(groupKey)!.push(isNaN(numVal) ? 1 : numVal);
      }
    }

    const rowFieldLabel = columnsMeta.find((c) => c.colIdx === rowField)?.label || 'Row Labels';
    const valFieldLabel = columnsMeta.find((c) => c.colIdx === valField)?.label || 'Value';
    const headers = [rowFieldLabel, `${agg} of ${valFieldLabel}`];
    const rows: (string | number)[][] = [];

    let grandVals: number[] = [];
    for (const [k, vals] of groups.entries()) {
      grandVals = grandVals.concat(vals);
      let res = 0;
      if (vals.length > 0) {
        if (agg === 'SUM') res = vals.reduce((a, b) => a + b, 0);
        else if (agg === 'COUNT') res = vals.length;
        else if (agg === 'AVERAGE') res = vals.reduce((a, b) => a + b, 0) / vals.length;
        else if (agg === 'MIN') res = Math.min(...vals);
        else if (agg === 'MAX') res = Math.max(...vals);
      }
      rows.push([k, Number(res.toFixed(2))]);
    }

    let grandTotal = 0;
    if (grandVals.length > 0) {
      if (agg === 'SUM') grandTotal = grandVals.reduce((a, b) => a + b, 0);
      else if (agg === 'COUNT') grandTotal = grandVals.length;
      else if (agg === 'AVERAGE') grandTotal = grandVals.reduce((a, b) => a + b, 0) / grandVals.length;
      else if (agg === 'MIN') grandTotal = Math.min(...grandVals);
      else if (agg === 'MAX') grandTotal = Math.max(...grandVals);
    }
    rows.push(['Grand Total', Number(grandTotal.toFixed(2))]);

    onCreatePivot({ headers, rows, targetNewSheet: newSheet });
    onClose();
  };

  return (
    <ModalShell title="Create PivotTable" onClose={onClose} widthClass="max-w-md">
      <div className="space-y-3">
        <div>
          <label className="block font-medium mb-1">Table / Range (include header row)</label>
          <input
            type="text"
            value={sourceRange}
            onChange={(e) => setSourceRange(e.target.value.toUpperCase())}
            className="w-full px-2.5 py-1.5 font-mono-num border border-slate-300 dark:border-slate-600 rounded bg-transparent"
          />
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="block font-medium mb-1">Group By (Rows)</label>
            <select
              value={rowField}
              onChange={(e) => setRowField(Number(e.target.value))}
              className="w-full px-2 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-800"
            >
              {columnsMeta.map((c) => (
                <option key={c.colIdx} value={c.colIdx}>{c.label}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="block font-medium mb-1">Values Column</label>
            <select
              value={valField}
              onChange={(e) => setValField(Number(e.target.value))}
              className="w-full px-2 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-800"
            >
              {columnsMeta.map((c) => (
                <option key={c.colIdx} value={c.colIdx}>{c.label}</option>
              ))}
            </select>
          </div>
        </div>
        <div>
          <label className="block font-medium mb-1">Summarize Values By</label>
          <select
            value={agg}
            onChange={(e) => setAgg(e.target.value as any)}
            className="w-full px-2 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-800"
          >
            <option value="SUM">Sum</option>
            <option value="COUNT">Count</option>
            <option value="AVERAGE">Average</option>
            <option value="MIN">Min</option>
            <option value="MAX">Max</option>
          </select>
        </div>
        <label className="flex items-center gap-2 cursor-pointer">
          <input type="checkbox" checked={newSheet} onChange={(e) => setNewSheet(e.target.checked)} />
          <span>Place PivotTable in New Worksheet</span>
        </label>
        <div className="flex justify-end gap-2 pt-2 border-t border-slate-100 dark:border-slate-800">
          <button onClick={onClose} className="px-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded">Cancel</button>
          <button onClick={handleGenerate} className="px-4 py-1.5 bg-[#107c41] text-white rounded font-medium hover:bg-[#0c5e31]">
            Generate PivotTable
          </button>
        </div>
      </div>
    </ModalShell>
  );
};

/* 8. NAME MANAGER MODAL */
export const NameManagerModal: React.FC<{
  namedRanges: NamedRange[];
  activeSheetId: string;
  defaultRange: string;
  sheets: Sheet[];
  onClose: () => void;
  onAdd: (nr: NamedRange) => void;
  onDelete: (name: string) => void;
}> = ({ namedRanges, activeSheetId, defaultRange, sheets, onClose, onAdd, onDelete }) => {
  const [name, setName] = useState('');
  const [range, setRange] = useState(defaultRange);
  const [err, setErr] = useState('');

  const handleAdd = () => {
    const clean = name.trim().replace(/\s+/g, '_');
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(clean)) {
      setErr('Name must start with a letter and contain no spaces.');
      return;
    }
    onAdd({ name: clean, sheetId: activeSheetId, range: range.toUpperCase() });
    setName('');
    setErr('');
  };

  return (
    <ModalShell title="Name Manager" onClose={onClose} widthClass="max-w-lg">
      <div className="space-y-3">
        <div className="flex gap-2">
          <input
            type="text"
            placeholder="Name (e.g. TaxRate)"
            value={name}
            onChange={(e) => setName(e.target.value)}
            className="flex-1 px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-transparent"
          />
          <input
            type="text"
            placeholder="Range (e.g. B2:B10)"
            value={range}
            onChange={(e) => setRange(e.target.value)}
            className="w-32 px-2.5 py-1.5 font-mono-num border border-slate-300 dark:border-slate-600 rounded bg-transparent"
          />
          <button onClick={handleAdd} className="px-3 py-1.5 bg-[#107c41] text-white rounded font-medium hover:bg-[#0c5e31]">
            New
          </button>
        </div>
        {err && <div className="text-red-600 text-[11px]">{err}</div>}
        <div className="border border-slate-200 dark:border-slate-700 rounded divide-y divide-slate-100 dark:divide-slate-800 max-h-56 overflow-y-auto">
          {namedRanges.length === 0 ? (
            <div className="p-4 text-center text-slate-400">No named ranges defined yet.</div>
          ) : (
            namedRanges.map((nr) => {
              const sName = sheets.find((s) => s.id === nr.sheetId)?.name || 'Sheet';
              return (
                <div key={nr.name} className="px-3 py-2 flex items-center justify-between">
                  <div>
                    <div className="font-semibold text-slate-800 dark:text-slate-100">{nr.name}</div>
                    <div className="font-mono-num text-[11px] text-slate-500">{sName}!{nr.range}</div>
                  </div>
                  <button onClick={() => onDelete(nr.name)} className="p-1 text-red-600 hover:bg-red-50 rounded">
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>
              );
            })
          )}
        </div>
      </div>
    </ModalShell>
  );
};

/* 9. GOAL SEEK MODAL */
export const GoalSeekModal: React.FC<{
  sheets: Sheet[];
  namedRanges: NamedRange[];
  activeSheetId: string;
  defaultSetCell: string;
  onClose: () => void;
  onApplyValue: (row: number, col: number, value: string) => void;
}> = ({ sheets, namedRanges, activeSheetId, defaultSetCell, onClose, onApplyValue }) => {
  const [setCell, setSetCell] = useState(defaultSetCell);
  const [toValue, setToValue] = useState('1000');
  const [byCell, setByCell] = useState('B3');
  const [resultMsg, setResultMsg] = useState('');

  const handleSolve = () => {
    const pSet = parseCellRef(setCell);
    const pBy = parseCellRef(byCell);
    const target = Number(toValue);
    if (!pSet || !pBy || isNaN(target)) {
      setResultMsg('Please enter valid cell addresses and numeric target.');
      return;
    }
    const res = solveGoalSeek(sheets, namedRanges, activeSheetId, pSet.row, pSet.col, target, pBy.row, pBy.col);
    if (res.success) {
      onApplyValue(pBy.row, pBy.col, String(Number(res.foundValue.toFixed(6))));
      setResultMsg(`Solution found! Set ${byCell.toUpperCase()} = ${res.foundValue.toFixed(4)} (${res.iterations} iterations).`);
    } else {
      setResultMsg('Could not converge to target value. Check that Set Cell depends on By Changing Cell.');
    }
  };

  return (
    <ModalShell title="Goal Seek (What-If Analysis)" onClose={onClose} widthClass="max-w-sm">
      <div className="space-y-3">
        <div>
          <label className="block font-medium mb-1">Set cell (formula cell)</label>
          <input type="text" value={setCell} onChange={(e) => setSetCell(e.target.value.toUpperCase())} className="w-full px-2.5 py-1.5 font-mono-num border border-slate-300 dark:border-slate-600 rounded bg-transparent" />
        </div>
        <div>
          <label className="block font-medium mb-1">To value</label>
          <input type="number" value={toValue} onChange={(e) => setToValue(e.target.value)} className="w-full px-2.5 py-1.5 font-mono-num border border-slate-300 dark:border-slate-600 rounded bg-transparent" />
        </div>
        <div>
          <label className="block font-medium mb-1">By changing cell</label>
          <input type="text" value={byCell} onChange={(e) => setByCell(e.target.value.toUpperCase())} className="w-full px-2.5 py-1.5 font-mono-num border border-slate-300 dark:border-slate-600 rounded bg-transparent" />
        </div>
        {resultMsg && <div className="p-2 bg-emerald-50 dark:bg-emerald-900/30 border border-emerald-200 dark:border-emerald-800 rounded text-[11px]">{resultMsg}</div>}
        <div className="flex justify-end gap-2 pt-2">
          <button onClick={onClose} className="px-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded">Close</button>
          <button onClick={handleSolve} className="px-4 py-1.5 bg-[#107c41] text-white rounded font-medium hover:bg-[#0c5e31]">Solve</button>
        </div>
      </div>
    </ModalShell>
  );
};

/* 10. TEXT TO COLUMNS / REMOVE DUPLICATES / CUSTOM SORT */
export const TextToColumnsModal: React.FC<{
  onClose: () => void;
  onExecute: (delimiter: string) => void;
}> = ({ onClose, onExecute }) => {
  const [preset, setPreset] = useState(',');
  const [custom, setCustom] = useState('|');

  return (
    <ModalShell title="Convert Text to Columns" onClose={onClose} widthClass="max-w-sm">
      <div className="space-y-3">
        <div className="text-slate-600 dark:text-slate-300">
          Splits text in the selected column across adjacent columns to the right.
        </div>
        <div>
          <label className="block font-medium mb-1">Delimiter</label>
          <select value={preset} onChange={(e) => setPreset(e.target.value)} className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-800">
            <option value=",">Comma (,)</option>
            <option value=";">Semicolon (;)</option>
            <option value="\t">Tab</option>
            <option value=" ">Space</option>
            <option value="custom">Custom</option>
          </select>
        </div>
        {preset === 'custom' && (
          <input type="text" value={custom} onChange={(e) => setCustom(e.target.value)} className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-transparent" />
        )}
        <div className="flex justify-end gap-2 pt-2">
          <button onClick={onClose} className="px-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded">Cancel</button>
          <button onClick={() => { onExecute(preset === 'custom' ? custom : preset === '\\t' ? '\t' : preset); onClose(); }} className="px-4 py-1.5 bg-[#107c41] text-white rounded font-medium hover:bg-[#0c5e31]">Split</button>
        </div>
      </div>
    </ModalShell>
  );
};

export const CustomSortModal: React.FC<{
  selectionCols: { colIdx: number; label: string }[];
  onClose: () => void;
  onExecute: (colIdx: number, asc: boolean, hasHeaders: boolean) => void;
}> = ({ selectionCols, onClose, onExecute }) => {
  const [colIdx, setColIdx] = useState(selectionCols[0]?.colIdx ?? 0);
  const [asc, setAsc] = useState(true);
  const [hasHeaders, setHasHeaders] = useState(true);

  return (
    <ModalShell title="Custom Sort" onClose={onClose} widthClass="max-w-sm">
      <div className="space-y-3">
        <label className="flex items-center gap-2 cursor-pointer">
          <input type="checkbox" checked={hasHeaders} onChange={(e) => setHasHeaders(e.target.checked)} />
          <span>My data has headers</span>
        </label>
        <div>
          <label className="block font-medium mb-1">Sort by Column</label>
          <select value={colIdx} onChange={(e) => setColIdx(Number(e.target.value))} className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-800">
            {selectionCols.map((c) => (
              <option key={c.colIdx} value={c.colIdx}>{c.label}</option>
            ))}
          </select>
        </div>
        <div>
          <label className="block font-medium mb-1">Order</label>
          <select value={asc ? 'asc' : 'desc'} onChange={(e) => setAsc(e.target.value === 'asc')} className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-800">
            <option value="asc">A to Z / Smallest to Largest</option>
            <option value="desc">Z to A / Largest to Smallest</option>
          </select>
        </div>
        <div className="flex justify-end gap-2 pt-2">
          <button onClick={onClose} className="px-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded">Cancel</button>
          <button onClick={() => { onExecute(colIdx, asc, hasHeaders); onClose(); }} className="px-4 py-1.5 bg-[#107c41] text-white rounded font-medium hover:bg-[#0c5e31]">Sort</button>
        </div>
      </div>
    </ModalShell>
  );
};

/* 11. COMMAND PALETTE / SEARCH */
export const CommandPaletteModal: React.FC<{
  onClose: () => void;
  commands: { id: string; label: string; category: string; run: () => void }[];
  onJumpToRef: (ref: string) => void;
}> = ({ onClose, commands, onJumpToRef }) => {
  const [query, setQuery] = useState('');

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return commands.slice(0, 14);
    return commands.filter((c) => c.label.toLowerCase().includes(q) || c.category.toLowerCase().includes(q)).slice(0, 14);
  }, [query, commands]);

  const isCellRef = /^[A-Za-z]{1,3}[1-9][0-9]*(:[A-Za-z]{1,3}[1-9][0-9]*)?$/.test(query.trim());

  return (
    <ModalShell title="Tell me what you want to do" onClose={onClose} widthClass="max-w-md">
      <div className="space-y-2.5">
        <div className="relative">
          <Search className="w-4 h-4 text-slate-400 absolute left-2.5 top-2" />
          <input
            autoFocus
            type="text"
            placeholder="Type a command (e.g. Pivot, Freeze, Chart, Currency) or cell (D20)..."
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                if (isCellRef) {
                  onJumpToRef(query.trim());
                  onClose();
                } else if (filtered[0]) {
                  filtered[0].run();
                  onClose();
                }
              }
            }}
            className="w-full pl-8 pr-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-transparent"
          />
        </div>
        <div className="max-h-64 overflow-y-auto divide-y divide-slate-100 dark:divide-slate-800 border border-slate-200 dark:border-slate-700 rounded">
          {isCellRef && (
            <button
              onClick={() => { onJumpToRef(query.trim()); onClose(); }}
              className="w-full px-3 py-2 text-left bg-emerald-50 dark:bg-emerald-900/30 hover:bg-emerald-100 flex items-center justify-between"
            >
              <span className="font-semibold text-[#107c41]">Go to {query.trim().toUpperCase()}</span>
              <span className="text-[11px] text-slate-400">Navigation</span>
            </button>
          )}
          {filtered.map((cmd) => (
            <button
              key={cmd.id}
              onClick={() => { cmd.run(); onClose(); }}
              className="w-full px-3 py-2 text-left hover:bg-slate-50 dark:hover:bg-slate-800 flex items-center justify-between"
            >
              <span className="font-medium text-slate-800 dark:text-slate-100">{cmd.label}</span>
              <span className="text-[11px] text-slate-400">{cmd.category}</span>
            </button>
          ))}
        </div>
      </div>
    </ModalShell>
  );
};

/* 12. SIMPLE INPUT MODAL (Comment, Hyperlink, Size, Sparkline, Stats) */
export const SimpleInputModal: React.FC<{
  title: string;
  label: string;
  initialValue: string;
  placeholder?: string;
  onClose: () => void;
  onConfirm: (val: string) => void;
}> = ({ title, label, initialValue, placeholder, onClose, onConfirm }) => {
  const [val, setVal] = useState(initialValue);
  return (
    <ModalShell title={title} onClose={onClose} widthClass="max-w-sm">
      <div className="space-y-3">
        <div>
          <label className="block font-medium mb-1">{label}</label>
          <input
            autoFocus
            type="text"
            value={val}
            placeholder={placeholder}
            onChange={(e) => setVal(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                onConfirm(val);
                onClose();
              }
            }}
            className="w-full px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 rounded bg-transparent"
          />
        </div>
        <div className="flex justify-end gap-2">
          <button onClick={onClose} className="px-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded">Cancel</button>
          <button
            onClick={() => { onConfirm(val); onClose(); }}
            className="px-4 py-1.5 bg-[#107c41] text-white rounded font-medium hover:bg-[#0c5e31]"
          >
            OK
          </button>
        </div>
      </div>
    </ModalShell>
  );
};

export const WorkbookStatsModal: React.FC<{
  workbook: Workbook;
  onClose: () => void;
}> = ({ workbook, onClose }) => {
  const stats = useMemo(() => {
    let totalCells = 0;
    let formulaCells = 0;
    let comments = 0;
    let charts = 0;
    for (const s of workbook.sheets) {
      charts += s.charts.length;
      for (const c of Object.values(s.cells)) {
        if (c.raw !== '') totalCells++;
        if (c.raw.startsWith('=')) formulaCells++;
        if (c.comment) comments++;
      }
    }
    return {
      sheets: workbook.sheets.length,
      totalCells,
      formulaCells,
      comments,
      charts,
      namedRanges: workbook.namedRanges.length,
    };
  }, [workbook]);

  return (
    <ModalShell title="Workbook Statistics" onClose={onClose} widthClass="max-w-sm">
      <div className="divide-y divide-slate-100 dark:divide-slate-800">
        {[
          ['Worksheets', stats.sheets],
          ['Cells with Data', stats.totalCells],
          ['Formulas', stats.formulaCells],
          ['Comments / Notes', stats.comments],
          ['Charts', stats.charts],
          ['Named Ranges', stats.namedRanges],
        ].map(([k, v]) => (
          <div key={String(k)} className="py-2 flex items-center justify-between">
            <span className="text-slate-500">{k}</span>
            <span className="font-mono-num font-semibold text-slate-900 dark:text-white">{v}</span>
          </div>
        ))}
      </div>
    </ModalShell>
  );
};

export const RemoveDuplicatesModal: React.FC<{
  selectionCols: { colIdx: number; label: string }[];
  onClose: () => void;
  onExecute: (selectedCols: number[], hasHeaders: boolean) => void;
}> = ({ selectionCols, onClose, onExecute }) => {
  const [selected, setSelected] = useState<number[]>(selectionCols.map((c) => c.colIdx));
  const [hasHeaders, setHasHeaders] = useState(true);

  const toggle = (cIdx: number) => {
    setSelected((prev) =>
      prev.includes(cIdx) ? prev.filter((x) => x !== cIdx) : [...prev, cIdx]
    );
  };

  return (
    <ModalShell title="Remove Duplicates" onClose={onClose} widthClass="max-w-sm">
      <div className="space-y-3">
        <label className="flex items-center gap-2 cursor-pointer">
          <input type="checkbox" checked={hasHeaders} onChange={(e) => setHasHeaders(e.target.checked)} />
          <span>My data has headers</span>
        </label>
        <div className="text-[11px] text-slate-500">Select columns to check for duplicate values:</div>
        <div className="max-h-44 overflow-y-auto border border-slate-200 dark:border-slate-700 rounded p-2 space-y-1.5">
          {selectionCols.map((c) => (
            <label key={c.colIdx} className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={selected.includes(c.colIdx)}
                onChange={() => toggle(c.colIdx)}
              />
              <span>{c.label}</span>
            </label>
          ))}
        </div>
        <div className="flex justify-end gap-2 pt-2">
          <button onClick={onClose} className="px-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded">Cancel</button>
          <button
            disabled={selected.length === 0}
            onClick={() => { onExecute(selected, hasHeaders); onClose(); }}
            className="px-4 py-1.5 bg-[#107c41] text-white rounded font-medium hover:bg-[#0c5e31] disabled:opacity-40"
          >
            OK
          </button>
        </div>
      </div>
    </ModalShell>
  );
};

