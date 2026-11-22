import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  CellCoord,
  CellData,
  CellStyle,
  ChartType,
  ClipboardData,
  ConditionalRule,
  ConditionalRuleType,
  DataValidationRule,
  ExcelMode,
  NumberFormatCategory,
  PageSetup,
  RibbonTabId,
  SelectionRange,
  Sheet,
  SheetChart,
  Workbook,
} from './types/spreadsheet';
import {
  cellKey,
  colIndexToLetter,
  formatCellRef,
  formatRangeRef,
  normalizeRange,
  parseCellKey,
  parseRangeRef,
  shiftFormulaRefs,
} from './utils/cellCoords';
import { extractFormulaRefs, WorkbookEvaluator } from './engine/formulaEngine';
import { buildDefaultWorkbook } from './utils/templates';
import {
  createEmptySheet,
  exportSheetToCSV,
  exportWorkbookToJSON,
  exportWorkbookToXLSX,
  importFileToWorkbook,
} from './utils/xlsxIO';
import { TitleBar } from './components/TitleBar';
import { Ribbon } from './components/Ribbon';
import { FormulaBar } from './components/FormulaBar';
import { SpreadsheetGrid } from './components/SpreadsheetGrid';
import { SheetTabsBar } from './components/SheetTabsBar';
import { SelectionSummary, StatusBar } from './components/StatusBar';
import { HomeDashboard } from './components/HomeDashboard';
import {
  LibraryWorkbookItem,
  loadWorkbookLibrary,
  saveWorkbookLibrary,
} from './utils/workbookLibrary';
import {
  CommandPaletteModal,
  ConditionalFormatModal,
  CustomSortModal,
  DataValidationModal,
  FileBackstageModal,
  FindReplaceModal,
  FormatCellsModal,
  FunctionWizardModal,
  GoalSeekModal,
  NameManagerModal,
  PivotTableModal,
  RemoveDuplicatesModal,
  SimpleInputModal,
  TextToColumnsModal,
  WorkbookStatsModal,
} from './components/ExcelModals';

const STORAGE_KEY = 'excelgrid_workbook_v1';

type ActiveModal =
  | null
  | { type: 'file' }
  | { type: 'find' }
  | { type: 'fx' }
  | { type: 'formatCells' }
  | { type: 'conditional'; initialType?: ConditionalRuleType }
  | { type: 'validation' }
  | { type: 'pivot' }
  | { type: 'names' }
  | { type: 'goalSeek' }
  | { type: 'textToCols' }
  | { type: 'removeDupes' }
  | { type: 'customSort' }
  | { type: 'commandPalette' }
  | { type: 'comment' }
  | { type: 'hyperlink' }
  | { type: 'size'; kind: 'col' | 'row' }
  | { type: 'sparkline'; sparkType: 'line' | 'column' | 'winloss' }
  | { type: 'stats' };

export default function App() {
  const [library, setLibrary] = useState<LibraryWorkbookItem[]>(() => loadWorkbookLibrary());
  const [viewMode, setViewMode] = useState<'home' | 'editor'>('home');
  const [workbook, setWorkbook] = useState<Workbook>(() => {
    const lib = loadWorkbookLibrary();
    return lib[0]?.workbook || buildDefaultWorkbook();
  });

  const [historyPast, setHistoryPast] = useState<Workbook[]>([]);
  const [historyFuture, setHistoryFuture] = useState<Workbook[]>([]);

  const [activeTab, setActiveTab] = useState<RibbonTabId>('home');
  const [activeCell, setActiveCell] = useState<CellCoord>({ row: 0, col: 0 });
  const [selection, setSelection] = useState<SelectionRange>({
    start: { row: 0, col: 0 },
    end: { row: 0, col: 0 },
  });
  const [extraSelections, setExtraSelections] = useState<SelectionRange[]>([]);
  const [isEditing, setIsEditing] = useState(false);
  const [editDraft, setEditDraft] = useState('');
  const [clipboard, setClipboard] = useState<ClipboardData | null>(null);
  const [formatPainterStyle, setFormatPainterStyle] = useState<CellStyle | null>(null);
  const [showAllComments, setShowAllComments] = useState(false);
  const [pageBreakView, setPageBreakView] = useState(false);
  const [selectedChartId, setSelectedChartId] = useState<string | null>(null);
  const [auditHighlights, setAuditHighlights] = useState<{
    precedents: CellCoord[];
    dependents: CellCoord[];
  }>({ precedents: [], dependents: [] });
  const [activeModal, setActiveModal] = useState<ActiveModal>(null);
  const [saveFlash, setSaveFlash] = useState(false);
  const [toastMsg, setToastMsg] = useState<string | null>(null);
  const [manualCalcVersion, setManualCalcVersion] = useState(0);

  const formulaInputRef = useRef<HTMLInputElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const lastInsertedRefSpan = useRef<{ start: number; len: number } | null>(null);

  const showToast = useCallback((msg: string) => {
    setToastMsg(msg);
    setTimeout(() => setToastMsg((cur) => (cur === msg ? null : cur)), 3200);
  }, []);

  // Apply dark mode class on <html>
  useEffect(() => {
    const root = document.documentElement;
    if (workbook.theme === 'dark') root.classList.add('dark');
    else root.classList.remove('dark');
  }, [workbook.theme]);

  // Auto-save active workbook into library and localStorage
  useEffect(() => {
    const t = setTimeout(() => {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(workbook));
      } catch {
        // storage quota full
      }
      setLibrary((prev) => {
        const exists = prev.some((item) => item.workbook.id === workbook.id);
        const next = exists
          ? prev.map((item) =>
              item.workbook.id === workbook.id
                ? { ...item, workbook, updatedAt: Date.now() }
                : item
            )
          : [{ workbook, updatedAt: Date.now(), pinned: false }, ...prev];
        saveWorkbookLibrary(next);
        return next;
      });
    }, 400);
    return () => clearTimeout(t);
  }, [workbook]);

  const activeSheet = useMemo(
    () => workbook.sheets.find((s) => s.id === workbook.activeSheetId) || workbook.sheets[0],
    [workbook.sheets, workbook.activeSheetId]
  );

  // Memoized Evaluator (respects calcMode)
  const evaluator = useMemo(() => {
    void manualCalcVersion;
    return new WorkbookEvaluator(workbook.sheets, workbook.namedRanges, workbook.currencySymbol);
  }, [workbook.sheets, workbook.namedRanges, workbook.currencySymbol, manualCalcVersion]);

  const commitWorkbookMutation = useCallback(
    (updater: (prev: Workbook) => Workbook, recordHistory = true) => {
      setWorkbook((prev) => {
        const next = updater(prev);
        if (next === prev) return prev;
        if (recordHistory) {
          setHistoryPast((h) => [...h.slice(-39), prev]);
          setHistoryFuture([]);
        }
        return next;
      });
    },
    []
  );

  const updateActiveSheet = useCallback(
    (updater: (sheet: Sheet) => Sheet, recordHistory = true) => {
      commitWorkbookMutation((wb) => {
        const idx = wb.sheets.findIndex((s) => s.id === wb.activeSheetId);
        if (idx === -1) return wb;
        const nextSheet = updater(wb.sheets[idx]);
        if (nextSheet === wb.sheets[idx]) return wb;
        const nextSheets = [...wb.sheets];
        nextSheets[idx] = nextSheet;
        return { ...wb, sheets: nextSheets };
      }, recordHistory);
    },
    [commitWorkbookMutation]
  );

  const handleUndo = useCallback(() => {
    setHistoryPast((past) => {
      if (past.length === 0) return past;
      const previous = past[past.length - 1];
      setWorkbook((cur) => {
        setHistoryFuture((f) => [cur, ...f.slice(0, 39)]);
        return previous;
      });
      return past.slice(0, -1);
    });
  }, []);

  const handleRedo = useCallback(() => {
    setHistoryFuture((fut) => {
      if (fut.length === 0) return fut;
      const next = fut[0];
      setWorkbook((cur) => {
        setHistoryPast((p) => [...p.slice(-39), cur]);
        return next;
      });
      return fut.slice(1);
    });
  }, []);

  const activeCellData = activeSheet.cells[cellKey(activeCell.row, activeCell.col)];
  const activeCellStyle = activeCellData?.style || {};

  // Sync formula bar when selection moves while not editing
  const formulaBarText = isEditing ? editDraft : activeCellData?.raw || '';

  const formulaHighlights = useMemo(() => {
    if (!isEditing || !editDraft.startsWith('=')) return [];
    return extractFormulaRefs(editDraft, activeSheet.id, workbook.sheets, workbook.namedRanges);
  }, [isEditing, editDraft, activeSheet.id, workbook.sheets, workbook.namedRanges]);

  // Live selection summary for StatusBar
  const selectionSummary = useMemo<SelectionSummary>(() => {
    const norm = normalizeRange(selection);
    let count = 0;
    let numericCount = 0;
    let sum = 0;
    let min = Infinity;
    let max = -Infinity;
    const cellTotal = (norm.end.row - norm.start.row + 1) * (norm.end.col - norm.start.col + 1);
    // Cap scan for huge full-column selections
    if (cellTotal > 10000) {
      for (const [k, cell] of Object.entries(activeSheet.cells)) {
        if (!cell.raw) continue;
        const { row, col } = parseCellKey(k);
        if (
          row >= norm.start.row &&
          row <= norm.end.row &&
          col >= norm.start.col &&
          col <= norm.end.col
        ) {
          count++;
          const ev = evaluator.evaluateCell(activeSheet.id, row, col);
          if (typeof ev.value === 'number') {
            numericCount++;
            sum += ev.value;
            if (ev.value < min) min = ev.value;
            if (ev.value > max) max = ev.value;
          }
        }
      }
    } else {
      for (let r = norm.start.row; r <= norm.end.row; r++) {
        for (let c = norm.start.col; c <= norm.end.col; c++) {
          const cell = activeSheet.cells[cellKey(r, c)];
          if (!cell || cell.raw === '') continue;
          count++;
          const ev = evaluator.evaluateCell(activeSheet.id, r, c);
          if (typeof ev.value === 'number') {
            numericCount++;
            sum += ev.value;
            if (ev.value < min) min = ev.value;
            if (ev.value > max) max = ev.value;
          }
        }
      }
    }
    return {
      count,
      numericCount,
      sum,
      avg: numericCount > 0 ? sum / numericCount : 0,
      min: numericCount > 0 ? min : 0,
      max: numericCount > 0 ? max : 0,
    };
  }, [selection, activeSheet.cells, activeSheet.id, evaluator]);

  // Validation checker
  const checkValidation = useCallback((rule: DataValidationRule | undefined, raw: string): string | null => {
    if (!rule || raw === '' || raw.startsWith('=')) return null;
    if (rule.type === 'list') {
      const items = rule.formula1.split(',').map((s) => s.trim().toLowerCase());
      if (!items.includes(raw.trim().toLowerCase())) {
        return rule.errorMessage || `Value must be one of: ${rule.formula1}`;
      }
      return null;
    }
    if (rule.type === 'whole' || rule.type === 'decimal') {
      const num = Number(raw);
      if (isNaN(num) || (rule.type === 'whole' && !Number.isInteger(num))) {
        return rule.errorMessage || `Please enter a valid ${rule.type === 'whole' ? 'whole number' : 'decimal number'}.`;
      }
      const v1 = Number(rule.formula1);
      const v2 = Number(rule.formula2 ?? 0);
      const op = rule.operator || 'between';
      if (op === 'between' && (num < v1 || num > v2)) return `Value must be between ${v1} and ${v2}.`;
      if (op === 'greaterThan' && !(num > v1)) return `Value must be greater than ${v1}.`;
      if (op === 'lessThan' && !(num < v1)) return `Value must be less than ${v1}.`;
      if (op === 'equal' && num !== v1) return `Value must equal ${v1}.`;
    }
    if (rule.type === 'textLength') {
      const len = raw.length;
      const v1 = Number(rule.formula1);
      const v2 = Number(rule.formula2 ?? 0);
      const op = rule.operator || 'between';
      if (op === 'between' && (len < v1 || len > v2)) return `Text length must be between ${v1} and ${v2} characters.`;
      if (op === 'lessThan' && !(len < v1)) return `Text length must be under ${v1} characters.`;
      if (op === 'greaterThan' && !(len > v1)) return `Text length must exceed ${v1} characters.`;
    }
    return null;
  }, []);

  const setCellRaw = useCallback(
    (row: number, col: number, raw: string) => {
      if (activeSheet.protected) {
        showToast('Sheet is protected. Unprotect in Review tab to edit.');
        return false;
      }
      const existing = activeSheet.cells[cellKey(row, col)];
      const err = checkValidation(existing?.validation, raw);
      if (err) {
        showToast(`Validation Error: ${err}`);
        return false;
      }
      updateActiveSheet((s) => {
        const nextCells = { ...s.cells };
        const key = cellKey(row, col);
        const prevCell = nextCells[key];
        if (raw === '' && !prevCell?.style && !prevCell?.comment && !prevCell?.validation && !prevCell?.sparkline && !prevCell?.isCheckbox) {
          delete nextCells[key];
        } else {
          nextCells[key] = { ...(prevCell || { raw: '' }), raw };
        }
        return { ...s, cells: nextCells };
      });
      return true;
    },
    [activeSheet.protected, activeSheet.cells, checkValidation, showToast, updateActiveSheet]
  );

  const handleStartEdit = useCallback(
    (initialChar?: string) => {
      if (activeSheet.protected) {
        showToast('Sheet is protected. Unprotect in Review tab to edit.');
        return;
      }
      lastInsertedRefSpan.current = null;
      setIsEditing(true);
      setEditDraft(initialChar !== undefined ? initialChar : activeCellData?.raw || '');
    },
    [activeSheet.protected, activeCellData?.raw, showToast]
  );

  const moveActiveCell = useCallback(
    (dRow: number, dCol: number, extend = false) => {
      const nr = Math.max(0, Math.min(activeSheet.rowCount - 1, activeCell.row + dRow));
      const nc = Math.max(0, Math.min(activeSheet.colCount - 1, activeCell.col + dCol));
      if (extend) {
        setSelection((sel) => ({
          start: sel.start,
          end: {
            row: Math.max(0, Math.min(activeSheet.rowCount - 1, sel.end.row + dRow)),
            col: Math.max(0, Math.min(activeSheet.colCount - 1, sel.end.col + dCol)),
          },
        }));
      } else {
        setActiveCell({ row: nr, col: nc });
        setSelection({ start: { row: nr, col: nc }, end: { row: nr, col: nc } });
        setExtraSelections([]);
      }
    },
    [activeCell, activeSheet.rowCount, activeSheet.colCount]
  );

  const handleCommitEdit = useCallback(
    (moveDir: 'down' | 'right' | 'left' | 'up' | 'none' = 'down') => {
      if (!isEditing) return;
      const ok = setCellRaw(activeCell.row, activeCell.col, editDraft);
      if (!ok) return;
      setIsEditing(false);
      lastInsertedRefSpan.current = null;
      if (moveDir === 'down') moveActiveCell(1, 0);
      else if (moveDir === 'up') moveActiveCell(-1, 0);
      else if (moveDir === 'right') moveActiveCell(0, 1);
      else if (moveDir === 'left') moveActiveCell(0, -1);
    },
    [isEditing, setCellRaw, activeCell.row, activeCell.col, editDraft, moveActiveCell]
  );

  const handleCancelEdit = useCallback(() => {
    setIsEditing(false);
    setEditDraft(activeCellData?.raw || '');
    lastInsertedRefSpan.current = null;
  }, [activeCellData?.raw]);

  const handleInsertRefIntoFormula = useCallback((refText: string) => {
    setEditDraft((prev) => {
      if (lastInsertedRefSpan.current) {
        const { start, len } = lastInsertedRefSpan.current;
        const next = prev.slice(0, start) + refText + prev.slice(start + len);
        lastInsertedRefSpan.current = { start, len: refText.length };
        return next;
      }
      const start = prev.length;
      lastInsertedRefSpan.current = { start, len: refText.length };
      return prev + refText;
    });
  }, []);

  // Format Painter or Normal Selection
  const handleSelectRange = useCallback(
    (range: SelectionRange, active?: CellCoord, appendMulti = false) => {
      const nextActive = active || range.start;
      if (formatPainterStyle && !appendMulti) {
        const norm = normalizeRange(range);
        updateActiveSheet((s) => {
          const nextCells = { ...s.cells };
          for (let r = norm.start.row; r <= norm.end.row; r++) {
            for (let c = norm.start.col; c <= norm.end.col; c++) {
              const k = cellKey(r, c);
              const prev = nextCells[k] || { raw: '' };
              nextCells[k] = { ...prev, style: { ...formatPainterStyle } };
            }
          }
          return { ...s, cells: nextCells };
        });
        setFormatPainterStyle(null);
      }
      setActiveCell(nextActive);
      if (appendMulti) {
        setExtraSelections((prev) => [...prev, selection]);
      } else {
        setExtraSelections([]);
      }
      setSelection(range);
    },
    [formatPainterStyle, selection, updateActiveSheet]
  );

  // Update Style across selection
  const handleUpdateStyle = useCallback(
    (patch: Partial<CellStyle>) => {
      if (activeSheet.protected) return showToast('Sheet is protected.');
      const ranges = [normalizeRange(selection), ...extraSelections.map(normalizeRange)];
      updateActiveSheet((s) => {
        const nextCells = { ...s.cells };
        for (const norm of ranges) {
          for (let r = norm.start.row; r <= norm.end.row; r++) {
            for (let c = norm.start.col; c <= norm.end.col; c++) {
              const k = cellKey(r, c);
              const prev = nextCells[k] || { raw: '' };
              nextCells[k] = {
                ...prev,
                style: { ...(prev.style || {}), ...patch },
              };
            }
          }
        }
        return { ...s, cells: nextCells };
      });
    },
    [activeSheet.protected, selection, extraSelections, showToast, updateActiveSheet]
  );

  const handleApplyBorder = useCallback(
    (preset: 'all' | 'outer' | 'bottom' | 'top' | 'left' | 'right' | 'thickBottom' | 'none') => {
      if (activeSheet.protected) return showToast('Sheet is protected.');
      const norm = normalizeRange(selection);
      const thin = { style: 'thin' as const, color: '#334155' };
      const thick = { style: 'thick' as const, color: '#0f172a' };

      updateActiveSheet((s) => {
        const nextCells = { ...s.cells };
        for (let r = norm.start.row; r <= norm.end.row; r++) {
          for (let c = norm.start.col; c <= norm.end.col; c++) {
            const k = cellKey(r, c);
            const prev = nextCells[k] || { raw: '' };
            const b = { ...(prev.style?.borders || {}) };
            if (preset === 'none') {
              nextCells[k] = { ...prev, style: { ...(prev.style || {}), borders: undefined } };
              continue;
            }
            if (preset === 'all') {
              b.top = thin; b.right = thin; b.bottom = thin; b.left = thin;
            } else if (preset === 'outer') {
              if (r === norm.start.row) b.top = thin;
              if (r === norm.end.row) b.bottom = thin;
              if (c === norm.start.col) b.left = thin;
              if (c === norm.end.col) b.right = thin;
            } else if (preset === 'bottom' && r === norm.end.row) b.bottom = thin;
            else if (preset === 'thickBottom' && r === norm.end.row) b.bottom = thick;
            else if (preset === 'top' && r === norm.start.row) b.top = thin;
            else if (preset === 'left' && c === norm.start.col) b.left = thin;
            else if (preset === 'right' && c === norm.end.col) b.right = thin;

            nextCells[k] = { ...prev, style: { ...(prev.style || {}), borders: b } };
          }
        }
        return { ...s, cells: nextCells };
      });
    },
    [activeSheet.protected, selection, showToast, updateActiveSheet]
  );

  // Clipboard Copy / Cut / Paste
  const handleCopyCut = useCallback(
    (mode: 'copy' | 'cut') => {
      const norm = normalizeRange(selection);
      const copied: Record<string, CellData> = {};
      const tsvRows: string[] = [];

      for (let r = norm.start.row; r <= norm.end.row; r++) {
        const tsvCols: string[] = [];
        for (let c = norm.start.col; c <= norm.end.col; c++) {
          const cell = activeSheet.cells[cellKey(r, c)];
          if (cell) {
            copied[cellKey(r - norm.start.row, c - norm.start.col)] = JSON.parse(JSON.stringify(cell));
          }
          const ev = evaluator.evaluateCell(activeSheet.id, r, c);
          tsvCols.push(ev.formatted);
        }
        tsvRows.push(tsvCols.join('\t'));
      }

      setClipboard({ mode, sheetId: activeSheet.id, range: norm, cells: copied });
      navigator.clipboard?.writeText(tsvRows.join('\n')).catch(() => {});
    },
    [selection, activeSheet.cells, activeSheet.id, evaluator]
  );

  const handlePaste = useCallback(
    async (valuesOnly = false) => {
      if (activeSheet.protected) return showToast('Sheet is protected.');
      const destStart = activeCell;

      if (clipboard) {
        const srcNorm = normalizeRange(clipboard.range);
        const rows = srcNorm.end.row - srcNorm.start.row + 1;
        const cols = srcNorm.end.col - srcNorm.start.col + 1;
        const dRow = destStart.row - srcNorm.start.row;
        const dCol = destStart.col - srcNorm.start.col;

        commitWorkbookMutation((wb) => {
          const nextSheets = wb.sheets.map((s) => {
            if (s.id !== activeSheet.id && !(clipboard.mode === 'cut' && s.id === clipboard.sheetId)) {
              return s;
            }
            const nextCells = { ...s.cells };
            if (clipboard.mode === 'cut' && s.id === clipboard.sheetId) {
              for (let r = srcNorm.start.row; r <= srcNorm.end.row; r++) {
                for (let c = srcNorm.start.col; c <= srcNorm.end.col; c++) {
                  delete nextCells[cellKey(r, c)];
                }
              }
            }
            if (s.id === activeSheet.id) {
              for (let dr = 0; dr < rows; dr++) {
                for (let dc = 0; dc < cols; dc++) {
                  const srcCell = clipboard.cells[cellKey(dr, dc)];
                  const tr = destStart.row + dr;
                  const tc = destStart.col + dc;
                  if (!srcCell) {
                    delete nextCells[cellKey(tr, tc)];
                    continue;
                  }
                  if (valuesOnly) {
                    const ev = evaluator.evaluateCell(clipboard.sheetId, srcNorm.start.row + dr, srcNorm.start.col + dc);
                    const prev = nextCells[cellKey(tr, tc)] || { raw: '' };
                    nextCells[cellKey(tr, tc)] = { ...prev, raw: ev.value !== null ? String(ev.value) : '' };
                  } else {
                    const cloned: CellData = JSON.parse(JSON.stringify(srcCell));
                    if (cloned.raw.startsWith('=') && clipboard.mode === 'copy') {
                      cloned.raw = shiftFormulaRefs(cloned.raw, dRow, dCol);
                    }
                    nextCells[cellKey(tr, tc)] = cloned;
                  }
                }
              }
            }
            return { ...s, cells: nextCells };
          });
          return { ...wb, sheets: nextSheets };
        });
        if (clipboard.mode === 'cut') setClipboard(null);
        return;
      }

      // Fallback to OS clipboard TSV
      try {
        const text = await navigator.clipboard.readText();
        if (!text) return;
        const lines = text.split(/\r?\n/);
        updateActiveSheet((s) => {
          const nextCells = { ...s.cells };
          lines.forEach((line, rOff) => {
            if (line === '' && rOff === lines.length - 1) return;
            line.split('\t').forEach((val, cOff) => {
              const k = cellKey(destStart.row + rOff, destStart.col + cOff);
              nextCells[k] = { ...(nextCells[k] || { raw: '' }), raw: val };
            });
          });
          return { ...s, cells: nextCells };
        });
      } catch {
        // clipboard permission denied
      }
    },
    [activeSheet.protected, activeSheet.id, activeCell, clipboard, commitWorkbookMutation, evaluator, showToast, updateActiveSheet]
  );

  // Clear Selection
  const handleClear = useCallback(
    (mode: 'all' | 'formats' | 'contents') => {
      if (activeSheet.protected) return showToast('Sheet is protected.');
      const ranges = [normalizeRange(selection), ...extraSelections.map(normalizeRange)];
      updateActiveSheet((s) => {
        const nextCells = { ...s.cells };
        for (const norm of ranges) {
          for (let r = norm.start.row; r <= norm.end.row; r++) {
            for (let c = norm.start.col; c <= norm.end.col; c++) {
              const k = cellKey(r, c);
              const prev = nextCells[k];
              if (!prev) continue;
              if (mode === 'all') delete nextCells[k];
              else if (mode === 'formats') nextCells[k] = { raw: prev.raw };
              else if (mode === 'contents') nextCells[k] = { ...prev, raw: '', sparkline: undefined };
            }
          }
        }
        return { ...s, cells: nextCells };
      });
    },
    [activeSheet.protected, selection, extraSelections, showToast, updateActiveSheet]
  );

  // Smart Fill Drag & Fill Down/Right
  const handleFillDrag = useCallback(
    (targetRange: SelectionRange) => {
      if (activeSheet.protected) return showToast('Sheet is protected.');
      const src = normalizeRange(selection);
      const dst = normalizeRange(targetRange);
      const isDown = dst.end.row > src.end.row;

      updateActiveSheet((s) => {
        const nextCells = { ...s.cells };
        if (isDown) {
          const srcRows = src.end.row - src.start.row + 1;
          for (let c = src.start.col; c <= src.end.col; c++) {
            // Detect arithmetic progression if >=2 numeric cells
            const srcNums: number[] = [];
            let allNums = true;
            for (let r = src.start.row; r <= src.end.row; r++) {
              const raw = s.cells[cellKey(r, c)]?.raw || '';
              if (raw === '' || raw.startsWith('=') || isNaN(Number(raw))) {
                allNums = false;
                break;
              }
              srcNums.push(Number(raw));
            }
            const step =
              allNums && srcNums.length >= 2
                ? srcNums[1] - srcNums[0]
                : allNums && srcNums.length === 1
                ? 0
                : 0;

            for (let r = src.end.row + 1; r <= dst.end.row; r++) {
              const offset = r - src.start.row;
              const templateRow = src.start.row + (offset % srcRows);
              const tpl = s.cells[cellKey(templateRow, c)];
              if (!tpl) continue;
              if (allNums && srcNums.length >= 2) {
                const val = srcNums[0] + step * offset;
                nextCells[cellKey(r, c)] = { ...JSON.parse(JSON.stringify(tpl)), raw: String(Number(val.toFixed(6))) };
              } else if (tpl.raw.startsWith('=')) {
                nextCells[cellKey(r, c)] = {
                  ...JSON.parse(JSON.stringify(tpl)),
                  raw: shiftFormulaRefs(tpl.raw, r - templateRow, 0),
                };
              } else {
                nextCells[cellKey(r, c)] = JSON.parse(JSON.stringify(tpl));
              }
            }
          }
        } else {
          // Fill Right
          const srcCols = src.end.col - src.start.col + 1;
          for (let r = src.start.row; r <= src.end.row; r++) {
            const srcNums: number[] = [];
            let allNums = true;
            for (let c = src.start.col; c <= src.end.col; c++) {
              const raw = s.cells[cellKey(r, c)]?.raw || '';
              if (raw === '' || raw.startsWith('=') || isNaN(Number(raw))) {
                allNums = false;
                break;
              }
              srcNums.push(Number(raw));
            }
            const step = allNums && srcNums.length >= 2 ? srcNums[1] - srcNums[0] : 0;

            for (let c = src.end.col + 1; c <= dst.end.col; c++) {
              const offset = c - src.start.col;
              const templateCol = src.start.col + (offset % srcCols);
              const tpl = s.cells[cellKey(r, templateCol)];
              if (!tpl) continue;
              if (allNums && srcNums.length >= 2) {
                const val = srcNums[0] + step * offset;
                nextCells[cellKey(r, c)] = { ...JSON.parse(JSON.stringify(tpl)), raw: String(Number(val.toFixed(6))) };
              } else if (tpl.raw.startsWith('=')) {
                nextCells[cellKey(r, c)] = {
                  ...JSON.parse(JSON.stringify(tpl)),
                  raw: shiftFormulaRefs(tpl.raw, 0, c - templateCol),
                };
              } else {
                nextCells[cellKey(r, c)] = JSON.parse(JSON.stringify(tpl));
              }
            }
          }
        }
        return { ...s, cells: nextCells };
      });
      setSelection(dst);
    },
    [activeSheet.protected, selection, showToast, updateActiveSheet]
  );

  // AutoSum
  const handleAutoSum = useCallback(
    (fn: 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN') => {
      if (activeSheet.protected) return showToast('Sheet is protected.');
      const norm = normalizeRange(selection);
      const isSingle = norm.start.row === norm.end.row && norm.start.col === norm.end.col;

      if (isSingle) {
        // Look upwards for contiguous numeric cells, else leftwards
        const r = norm.start.row;
        const c = norm.start.col;
        let top = r - 1;
        while (top >= 0 && typeof evaluator.evaluateCell(activeSheet.id, top, c).value === 'number') {
          top--;
        }
        top++;
        if (top <= r - 1) {
          setCellRaw(r, c, `=${fn}(${formatCellRef(top, c)}:${formatCellRef(r - 1, c)})`);
          return;
        }
        let left = c - 1;
        while (left >= 0 && typeof evaluator.evaluateCell(activeSheet.id, r, left).value === 'number') {
          left--;
        }
        left++;
        if (left <= c - 1) {
          setCellRaw(r, c, `=${fn}(${formatCellRef(r, left)}:${formatCellRef(r, c - 1)})`);
          return;
        }
        handleStartEdit(`=${fn}(`);
      } else {
        // Put total in row below each column
        const targetRow = norm.end.row + 1;
        updateActiveSheet((s) => {
          const nextCells = { ...s.cells };
          for (let c = norm.start.col; c <= norm.end.col; c++) {
            const formula = `=${fn}(${formatCellRef(norm.start.row, c)}:${formatCellRef(norm.end.row, c)})`;
            nextCells[cellKey(targetRow, c)] = {
              ...(nextCells[cellKey(targetRow, c)] || { raw: '' }),
              raw: formula,
              style: { ...(nextCells[cellKey(targetRow, c)]?.style || {}), bold: true },
            };
          }
          return { ...s, cells: nextCells };
        });
      }
    },
    [activeSheet.protected, activeSheet.id, selection, evaluator, setCellRaw, handleStartEdit, showToast, updateActiveSheet]
  );

  // Merge & Unmerge
  const handleMergeCenter = useCallback(() => {
    const norm = normalizeRange(selection);
    if (norm.start.row === norm.end.row && norm.start.col === norm.end.col) return;
    updateActiveSheet((s) => {
      // Remove overlapping merges
      const filtered = s.mergedRanges.filter(
        (m) =>
          m.endRow < norm.start.row ||
          m.startRow > norm.end.row ||
          m.endCol < norm.start.col ||
          m.startCol > norm.end.col
      );
      const nextCells = { ...s.cells };
      const masterKey = cellKey(norm.start.row, norm.start.col);
      nextCells[masterKey] = {
        ...(nextCells[masterKey] || { raw: '' }),
        style: { ...(nextCells[masterKey]?.style || {}), horizontalAlign: 'center', verticalAlign: 'middle' },
      };
      return { ...s, mergedRanges: [...filtered, { startRow: norm.start.row, startCol: norm.start.col, endRow: norm.end.row, endCol: norm.end.col }], cells: nextCells };
    });
  }, [selection, updateActiveSheet]);

  const handleUnmerge = useCallback(() => {
    const norm = normalizeRange(selection);
    updateActiveSheet((s) => ({
      ...s,
      mergedRanges: s.mergedRanges.filter(
        (m) =>
          m.endRow < norm.start.row ||
          m.startRow > norm.end.row ||
          m.endCol < norm.start.col ||
          m.startCol > norm.end.col
      ),
    }));
  }, [selection, updateActiveSheet]);

  // Insert / Delete Row & Col
  const handleInsertRowAt = useCallback(
    (targetRow: number) => {
      updateActiveSheet((s) => {
        const nextCells: Record<string, CellData> = {};
        for (const [k, cell] of Object.entries(s.cells)) {
          const { row, col } = parseCellKey(k);
          nextCells[cellKey(row >= targetRow ? row + 1 : row, col)] = cell;
        }
        return { ...s, rowCount: s.rowCount + 1, cells: nextCells };
      });
    },
    [updateActiveSheet]
  );

  const handleDeleteRowAt = useCallback(
    (targetRow: number) => {
      updateActiveSheet((s) => {
        const nextCells: Record<string, CellData> = {};
        for (const [k, cell] of Object.entries(s.cells)) {
          const { row, col } = parseCellKey(k);
          if (row === targetRow) continue;
          nextCells[cellKey(row > targetRow ? row - 1 : row, col)] = cell;
        }
        return { ...s, rowCount: Math.max(20, s.rowCount - 1), cells: nextCells };
      });
    },
    [updateActiveSheet]
  );

  const handleInsertColAt = useCallback(
    (targetCol: number) => {
      updateActiveSheet((s) => {
        const nextCells: Record<string, CellData> = {};
        for (const [k, cell] of Object.entries(s.cells)) {
          const { row, col } = parseCellKey(k);
          nextCells[cellKey(row, col >= targetCol ? col + 1 : col)] = cell;
        }
        return { ...s, colCount: s.colCount + 1, cells: nextCells };
      });
    },
    [updateActiveSheet]
  );

  const handleDeleteColAt = useCallback(
    (targetCol: number) => {
      updateActiveSheet((s) => {
        const nextCells: Record<string, CellData> = {};
        for (const [k, cell] of Object.entries(s.cells)) {
          const { row, col } = parseCellKey(k);
          if (col === targetCol) continue;
          nextCells[cellKey(row, col > targetCol ? col - 1 : col)] = cell;
        }
        return { ...s, colCount: Math.max(10, s.colCount - 1), cells: nextCells };
      });
    },
    [updateActiveSheet]
  );

  // Sorting
  const handleSort = useCallback(
    (sortCol: number, asc: boolean, hasHeadersOverride?: boolean) => {
      const norm = normalizeRange(selection);
      const isSingleCell = norm.start.row === norm.end.row && norm.start.col === norm.end.col;

      // Determine contiguous table bounds if single cell selected
      let rStart = norm.start.row;
      let rEnd = norm.end.row;
      let cStart = norm.start.col;
      let cEnd = norm.end.col;

      if (isSingleCell) {
        rStart = 0;
        cStart = 0;
        cEnd = 0;
        rEnd = 0;
        for (const k of Object.keys(activeSheet.cells)) {
          const { row, col } = parseCellKey(k);
          if (activeSheet.cells[k]?.raw) {
            if (row > rEnd) rEnd = row;
            if (col > cEnd) cEnd = col;
          }
        }
      }

      const hasHeaders = hasHeadersOverride ?? (rEnd > rStart && isSingleCell);
      const dataStartRow = hasHeaders ? rStart + 1 : rStart;
      if (dataStartRow >= rEnd) return;

      const rowBundles: { origRow: number; sortVal: any; cells: Record<number, CellData> }[] = [];
      for (let r = dataStartRow; r <= rEnd; r++) {
        const rowCells: Record<number, CellData> = {};
        let hasAny = false;
        for (let c = cStart; c <= cEnd; c++) {
          const cd = activeSheet.cells[cellKey(r, c)];
          if (cd) {
            rowCells[c] = cd;
            if (cd.raw !== '') hasAny = true;
          }
        }
        if (!hasAny && isSingleCell) continue;
        const ev = evaluator.evaluateCell(activeSheet.id, r, sortCol);
        rowBundles.push({ origRow: r, sortVal: ev.value ?? ev.formatted, cells: rowCells });
      }

      rowBundles.sort((a, b) => {
        const av = a.sortVal;
        const bv = b.sortVal;
        if (typeof av === 'number' && typeof bv === 'number') {
          return asc ? av - bv : bv - av;
        }
        const sa = String(av ?? '').toLowerCase();
        const sb = String(bv ?? '').toLowerCase();
        return asc ? sa.localeCompare(sb) : sb.localeCompare(sa);
      });

      updateActiveSheet((s) => {
        const nextCells = { ...s.cells };
        for (let r = dataStartRow; r <= rEnd; r++) {
          for (let c = cStart; c <= cEnd; c++) {
            delete nextCells[cellKey(r, c)];
          }
        }
        rowBundles.forEach((bundle, idx) => {
          const targetR = dataStartRow + idx;
          for (const [cStr, cell] of Object.entries(bundle.cells)) {
            nextCells[cellKey(targetR, Number(cStr))] = cell;
          }
        });
        return { ...s, cells: nextCells };
      });
    },
    [selection, activeSheet.cells, activeSheet.id, evaluator, updateActiveSheet]
  );

  // Format as Table
  const handleFormatAsTable = useCallback(
    (preset: 'green' | 'blue' | 'slate' | 'amber') => {
      const colors = {
        green: { hdr: '#107c41', band: '#ecfdf5', border: '#a7f3d0' },
        blue: { hdr: '#0284c7', band: '#f0f9ff', border: '#bae6fd' },
        slate: { hdr: '#334155', band: '#f8fafc', border: '#cbd5e1' },
        amber: { hdr: '#d97706', band: '#fffbeb', border: '#fde68a' },
      }[preset];

      const norm = normalizeRange(selection);
      updateActiveSheet((s) => {
        const nextCells = { ...s.cells };
        for (let r = norm.start.row; r <= norm.end.row; r++) {
          for (let c = norm.start.col; c <= norm.end.col; c++) {
            const k = cellKey(r, c);
            const prev = nextCells[k] || { raw: '' };
            const isHdr = r === norm.start.row;
            const isBand = !isHdr && (r - norm.start.row) % 2 === 1;
            nextCells[k] = {
              ...prev,
              style: {
                ...(prev.style || {}),
                bold: isHdr ? true : prev.style?.bold,
                color: isHdr ? '#ffffff' : prev.style?.color,
                backgroundColor: isHdr ? colors.hdr : isBand ? colors.band : '#ffffff',
                borders: {
                  bottom: { style: 'thin', color: colors.border },
                },
              },
            };
          }
        }
        return {
          ...s,
          cells: nextCells,
          autoFilter: { range: formatRangeRef(norm), hiddenValuesByCol: {} },
        };
      });
      showToast('Formatted selection as Filterable Table.');
    },
    [selection, showToast, updateActiveSheet]
  );

  // Insert Chart
  const handleInsertChart = useCallback(
    (type: ChartType) => {
      const norm = normalizeRange(selection);
      const isSingle = norm.start.row === norm.end.row && norm.start.col === norm.end.col;
      const dataRange = isSingle ? 'A1:F7' : formatRangeRef(norm);
      const chart: SheetChart = {
        id: `chart-${Date.now()}`,
        type,
        title: `${type.charAt(0).toUpperCase() + type.slice(1)} Chart`,
        dataRange,
        x: 240 + activeSheet.charts.length * 30,
        y: 140 + activeSheet.charts.length * 30,
        width: 480,
        height: 290,
        showLegend: true,
        showGridlines: true,
        palette: 'excel',
      };
      updateActiveSheet((s) => ({ ...s, charts: [...s.charts, chart] }));
      setSelectedChartId(chart.id);
    },
    [selection, activeSheet.charts.length, updateActiveSheet]
  );

  // Jump to reference (e.g. D20, A1:C5, or NamedRange)
  const handleJumpToReference = useCallback(
    (refOrName: string) => {
      const cleaned = refOrName.trim();
      const nr = workbook.namedRanges.find((n) => n.name.toLowerCase() === cleaned.toLowerCase());
      if (nr) {
        const rng = parseRangeRef(nr.range);
        if (rng) {
          setWorkbook((wb) => ({ ...wb, activeSheetId: nr.sheetId }));
          setActiveCell(rng.start);
          setSelection(rng);
        }
        return;
      }
      const rng = parseRangeRef(cleaned);
      if (rng) {
        setActiveCell(rng.start);
        setSelection(rng);
      } else {
        showToast(`Invalid reference or name: ${cleaned}`);
      }
    },
    [workbook.namedRanges, showToast]
  );

  // Global Keyboard Shortcuts
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (activeModal) return;
      const target = e.target as HTMLElement;
      const inInput = target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.tagName === 'SELECT';

      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        setActiveModal({ type: 'commandPalette' });
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') {
        e.preventDefault();
        setActiveModal({ type: 'find' });
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        localStorage.setItem(STORAGE_KEY, JSON.stringify(workbook));
        setSaveFlash(true);
        setTimeout(() => setSaveFlash(false), 1800);
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') {
        e.preventDefault();
        window.print();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key === '`') {
        e.preventDefault();
        commitWorkbookMutation((wb) => ({ ...wb, showFormulas: !wb.showFormulas }), false);
        return;
      }
      if (e.key === 'F9') {
        e.preventDefault();
        setManualCalcVersion((v) => v + 1);
        showToast('Workbook recalculated.');
        return;
      }

      if (inInput || isEditing) return;

      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        if (e.shiftKey) handleRedo();
        else handleUndo();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
        e.preventDefault();
        handleRedo();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
        e.preventDefault();
        handleCopyCut('copy');
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'x') {
        e.preventDefault();
        handleCopyCut('cut');
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v') {
        e.preventDefault();
        handlePaste(false);
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'b') {
        e.preventDefault();
        handleUpdateStyle({ bold: !activeCellStyle.bold });
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'i') {
        e.preventDefault();
        handleUpdateStyle({ italic: !activeCellStyle.italic });
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'u') {
        e.preventDefault();
        handleUpdateStyle({ underline: !activeCellStyle.underline });
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'd') {
        e.preventDefault();
        handleFillDrag({ start: normalizeRange(selection).start, end: { row: normalizeRange(selection).end.row + 1, col: normalizeRange(selection).end.col } });
        return;
      }

      if (e.key === 'F2') {
        e.preventDefault();
        handleStartEdit();
        return;
      }
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        if (selectedChartId) {
          updateActiveSheet((s) => ({ ...s, charts: s.charts.filter((ch) => ch.id !== selectedChartId) }));
          setSelectedChartId(null);
        } else {
          handleClear('contents');
        }
        return;
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        moveActiveCell(e.ctrlKey || e.metaKey ? -10 : -1, 0, e.shiftKey);
        return;
      }
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        moveActiveCell(e.ctrlKey || e.metaKey ? 10 : 1, 0, e.shiftKey);
        return;
      }
      if (e.key === 'ArrowLeft') {
        e.preventDefault();
        moveActiveCell(0, e.ctrlKey || e.metaKey ? -5 : -1, e.shiftKey);
        return;
      }
      if (e.key === 'ArrowRight') {
        e.preventDefault();
        moveActiveCell(0, e.ctrlKey || e.metaKey ? 5 : 1, e.shiftKey);
        return;
      }
      if (e.key === 'Enter') {
        e.preventDefault();
        moveActiveCell(e.shiftKey ? -1 : 1, 0);
        return;
      }
      if (e.key === 'Tab') {
        e.preventDefault();
        moveActiveCell(0, e.shiftKey ? -1 : 1);
        return;
      }
      if (e.key === 'Escape') {
        setClipboard(null);
        setFormatPainterStyle(null);
        setAuditHighlights({ precedents: [], dependents: [] });
        return;
      }

      // Direct typing enters cell edit mode
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        handleStartEdit(e.key);
      }
    };

    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [
    activeModal,
    isEditing,
    workbook,
    selection,
    activeCellStyle,
    selectedChartId,
    commitWorkbookMutation,
    handleUndo,
    handleRedo,
    handleCopyCut,
    handlePaste,
    handleUpdateStyle,
    handleFillDrag,
    handleStartEdit,
    handleClear,
    moveActiveCell,
    showToast,
    updateActiveSheet,
  ]);

  // Selection column metadata for modals
  const selectionColsMeta = useMemo(() => {
    const norm = normalizeRange(selection);
    const cols: { colIdx: number; label: string }[] = [];
    for (let c = norm.start.col; c <= norm.end.col; c++) {
      const hdr = evaluator.evaluateCell(activeSheet.id, norm.start.row, c).formatted;
      cols.push({ colIdx: c, label: hdr ? `${colIndexToLetter(c)} (${hdr})` : `Column ${colIndexToLetter(c)}` });
    }
    return cols;
  }, [selection, activeSheet.id, evaluator]);

  const commandList = useMemo(
    () => [
      { id: 'new-pivot', label: 'Insert PivotTable', category: 'Insert', run: () => setActiveModal({ type: 'pivot' }) },
      { id: 'chart-col', label: 'Insert Column Chart', category: 'Insert', run: () => handleInsertChart('column') },
      { id: 'chart-pie', label: 'Insert Pie Chart', category: 'Insert', run: () => handleInsertChart('pie') },
      { id: 'chart-line', label: 'Insert Line Chart', category: 'Insert', run: () => handleInsertChart('line') },
      { id: 'fx', label: 'Insert Function (fx)', category: 'Formulas', run: () => setActiveModal({ type: 'fx' }) },
      { id: 'cond-fmt', label: 'Conditional Formatting Rules', category: 'Home', run: () => setActiveModal({ type: 'conditional' }) },
      { id: 'data-val', label: 'Data Validation / Dropdown List', category: 'Data', run: () => setActiveModal({ type: 'validation' }) },
      { id: 'goal-seek', label: 'Goal Seek (What-If Analysis)', category: 'Data', run: () => setActiveModal({ type: 'goalSeek' }) },
      { id: 'rem-dupes', label: 'Remove Duplicates', category: 'Data', run: () => setActiveModal({ type: 'removeDupes' }) },
      { id: 'text-cols', label: 'Text to Columns', category: 'Data', run: () => setActiveModal({ type: 'textToCols' }) },
      { id: 'freeze-top', label: 'Freeze Top Row', category: 'View', run: () => updateActiveSheet((s) => ({ ...s, freezePanes: { row: 1, col: 0 } })) },
      { id: 'unfreeze', label: 'Unfreeze Panes', category: 'View', run: () => updateActiveSheet((s) => ({ ...s, freezePanes: { row: 0, col: 0 } })) },
      { id: 'find', label: 'Find & Replace', category: 'Home', run: () => setActiveModal({ type: 'find' }) },
      { id: 'name-mgr', label: 'Name Manager', category: 'Formulas', run: () => setActiveModal({ type: 'names' }) },
      { id: 'export-xlsx', label: 'Export Workbook (.xlsx)', category: 'File', run: () => exportWorkbookToXLSX(workbook) },
      { id: 'export-csv', label: 'Export Sheet (.csv)', category: 'File', run: () => exportSheetToCSV(workbook, activeSheet.id) },
      { id: 'print', label: 'Print Sheet', category: 'File', run: () => window.print() },
    ],
    [handleInsertChart, updateActiveSheet, workbook, activeSheet.id]
  );

  const openWorkbookInEditor = useCallback((wb: Workbook) => {
    setWorkbook(wb);
    setHistoryPast([]);
    setHistoryFuture([]);
    setIsEditing(false);
    setActiveCell({ row: 0, col: 0 });
    setSelection({ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
    setExtraSelections([]);
    setSelectedChartId(null);
    setAuditHighlights({ precedents: [], dependents: [] });
    setViewMode('editor');
  }, []);

  const hiddenFileInput = (
    <input
      ref={fileInputRef}
      type="file"
      accept=".xlsx,.xls,.csv,.json"
      className="hidden"
      onChange={async (e) => {
        const f = e.target.files?.[0];
        if (!f) return;
        try {
          const imported = await importFileToWorkbook(f);
          setLibrary((prev) => {
            const next = [{ workbook: imported, updatedAt: Date.now(), pinned: false }, ...prev];
            saveWorkbookLibrary(next);
            return next;
          });
          openWorkbookInEditor(imported);
          showToast(`Imported "${f.name}" successfully.`);
        } catch (err) {
          showToast(`Import failed: ${err instanceof Error ? err.message : 'Invalid file'}`);
        } finally {
          e.target.value = '';
        }
      }}
    />
  );

  if (viewMode === 'home') {
    return (
      <>
        {hiddenFileInput}
        <HomeDashboard
          library={library}
          theme={workbook.theme}
          onToggleTheme={() => {
            const nextTheme = workbook.theme === 'dark' ? 'light' : 'dark';
            setWorkbook((wb) => ({ ...wb, theme: nextTheme }));
          }}
          onOpenWorkbook={(wbId) => {
            const found = library.find((item) => item.workbook.id === wbId);
            if (found) openWorkbookInEditor(found.workbook);
          }}
          onCreateFromTemplate={(newWb) => {
            const themedWb: Workbook = { ...newWb, theme: workbook.theme };
            setLibrary((prev) => {
              const next = [{ workbook: themedWb, updatedAt: Date.now(), pinned: false }, ...prev];
              saveWorkbookLibrary(next);
              return next;
            });
            openWorkbookInEditor(themedWb);
          }}
          onImportFile={() => fileInputRef.current?.click()}
          onTogglePin={(wbId) => {
            setLibrary((prev) => {
              const next = prev.map((item) =>
                item.workbook.id === wbId ? { ...item, pinned: !item.pinned } : item
              );
              saveWorkbookLibrary(next);
              return next;
            });
          }}
          onDuplicateWorkbook={(wbId) => {
            const found = library.find((item) => item.workbook.id === wbId);
            if (!found) return;
            const cloned: Workbook = {
              ...JSON.parse(JSON.stringify(found.workbook)),
              id: `wb-${Date.now()}`,
              title: `${found.workbook.title.replace(/\.xlsx$/i, '')} (Copy).xlsx`,
            };
            setLibrary((prev) => {
              const next = [{ workbook: cloned, updatedAt: Date.now(), pinned: false }, ...prev];
              saveWorkbookLibrary(next);
              return next;
            });
            showToast(`Duplicated "${found.workbook.title}".`);
          }}
          onExportXLSX={(wb) => exportWorkbookToXLSX(wb)}
          onDeleteWorkbook={(wbId) => {
            setLibrary((prev) => {
              const next = prev.filter((item) => item.workbook.id !== wbId);
              saveWorkbookLibrary(next);
              return next;
            });
            showToast('Workbook removed from library.');
          }}
        />
        {toastMsg && (
          <div className="no-print fixed bottom-6 left-1/2 -translate-x-1/2 z-50 bg-slate-900 dark:bg-slate-800 text-white text-xs font-medium px-4 py-2 rounded-md shadow-lg border border-slate-700">
            {toastMsg}
          </div>
        )}
      </>
    );
  }

  return (
    <div className="h-screen w-screen flex flex-col overflow-hidden bg-white dark:bg-slate-950 text-slate-900 dark:text-slate-100">
      {hiddenFileInput}

      <TitleBar
        title={workbook.title}
        onRenameTitle={(t) => commitWorkbookMutation((wb) => ({ ...wb, title: t }), false)}
        canUndo={historyPast.length > 0}
        canRedo={historyFuture.length > 0}
        onUndo={handleUndo}
        onRedo={handleRedo}
        onSaveLocal={() => {
          localStorage.setItem(STORAGE_KEY, JSON.stringify(workbook));
          setSaveFlash(true);
          setTimeout(() => setSaveFlash(false), 1800);
        }}
        onExportXLSX={() => exportWorkbookToXLSX(workbook)}
        onPrint={() => window.print()}
        onOpenSearch={() => setActiveModal({ type: 'commandPalette' })}
        theme={workbook.theme}
        onToggleTheme={() =>
          commitWorkbookMutation((wb) => ({ ...wb, theme: wb.theme === 'dark' ? 'light' : 'dark' }), false)
        }
        saveFlash={saveFlash}
        onGoHome={() => {
          if (isEditing) handleCommitEdit('none');
          setViewMode('home');
        }}
      />

      <Ribbon
        activeTab={activeTab}
        onSelectTab={setActiveTab}
        onOpenFileBackstage={() => setActiveModal({ type: 'file' })}
        workbook={workbook}
        sheet={activeSheet}
        activeCellStyle={activeCellStyle}
        formatPainterActive={Boolean(formatPainterStyle)}
        showAllComments={showAllComments}
        pageBreakView={pageBreakView}
        onTogglePageBreakView={() => setPageBreakView((v) => !v)}
        onToggleTheme={() =>
          commitWorkbookMutation((wb) => ({ ...wb, theme: wb.theme === 'dark' ? 'light' : 'dark' }), false)
        }
        onCut={() => handleCopyCut('cut')}
        onCopy={() => handleCopyCut('copy')}
        onPaste={() => handlePaste(false)}
        onPasteValues={() => handlePaste(true)}
        onToggleFormatPainter={() =>
          setFormatPainterStyle((cur) => (cur ? null : { ...activeCellStyle }))
        }
        onUpdateStyle={handleUpdateStyle}
        onApplyBorder={handleApplyBorder}
        onMergeCenter={handleMergeCenter}
        onUnmerge={handleUnmerge}
        onSetNumberFormat={(cat: NumberFormatCategory) =>
          handleUpdateStyle({
            numberFormat: {
              category: cat,
              decimals: cat === 'percentage' ? 1 : 2,
              currencySymbol: workbook.currencySymbol,
            },
          })
        }
        onAdjustDecimals={(delta) => {
          const curFmt = activeCellStyle.numberFormat || { category: 'number' as const, decimals: 2 };
          const nextDec = Math.max(0, Math.min(8, (curFmt.decimals ?? 2) + delta));
          handleUpdateStyle({
            numberFormat: { ...curFmt, category: curFmt.category === 'general' ? 'number' : curFmt.category, decimals: nextDec },
          });
        }}
        onApplyCellStylePreset={(preset) => {
          const map: Record<string, Partial<CellStyle>> = {
            normal: { bold: false, color: undefined, backgroundColor: undefined, borders: undefined },
            good: { backgroundColor: '#dcfce7', color: '#166534' },
            bad: { backgroundColor: '#fee2e2', color: '#991b1b' },
            neutral: { backgroundColor: '#fef9c3', color: '#854d0e' },
            heading1: { bold: true, fontSize: 15, color: '#0f172a', borders: { bottom: { style: 'thick', color: '#107c41' } } },
            heading2: { bold: true, fontSize: 13, color: '#334155', borders: { bottom: { style: 'thin', color: '#94a3b8' } } },
            total: { bold: true, borders: { top: { style: 'thin', color: '#0f172a' }, bottom: { style: 'double', color: '#0f172a' } } },
          };
          handleUpdateStyle(map[preset] || {});
        }}
        onOpenConditionalModal={(initialType) => setActiveModal({ type: 'conditional', initialType })}
        onClearConditionalRules={() => updateActiveSheet((s) => ({ ...s, conditionalRules: [] }))}
        onFormatAsTable={handleFormatAsTable}
        onInsertRow={() => handleInsertRowAt(activeCell.row)}
        onInsertCol={() => handleInsertColAt(activeCell.col)}
        onAddSheet={() => {
          const idx = workbook.sheets.length + 1;
          const ns = createEmptySheet(`sheet-${Date.now()}`, `Sheet${idx}`);
          commitWorkbookMutation((wb) => ({
            ...wb,
            sheets: [...wb.sheets, ns],
            activeSheetId: ns.id,
          }));
        }}
        onDeleteRow={() => handleDeleteRowAt(activeCell.row)}
        onDeleteCol={() => handleDeleteColAt(activeCell.col)}
        onDeleteSheet={() => {
          if (workbook.sheets.length <= 1) return showToast('Cannot delete the only worksheet.');
          commitWorkbookMutation((wb) => {
            const next = wb.sheets.filter((s) => s.id !== activeSheet.id);
            return { ...wb, sheets: next, activeSheetId: next[0].id };
          });
        }}
        onAutoFitCol={() => {
          const c = activeCell.col;
          let maxLen = 6;
          for (let r = 0; r < activeSheet.rowCount; r++) {
            const ev = evaluator.evaluateCell(activeSheet.id, r, c);
            if (ev.formatted.length > maxLen) maxLen = ev.formatted.length;
          }
          updateActiveSheet((s) => ({
            ...s,
            colWidths: { ...s.colWidths, [c]: Math.min(360, Math.max(60, maxLen * 8 + 20)) },
          }));
        }}
        onOpenSizeModal={(kind) => setActiveModal({ type: 'size', kind })}
        onHideRow={() => {
          const norm = normalizeRange(selection);
          updateActiveSheet((s) => {
            const next = { ...s.hiddenRows };
            for (let r = norm.start.row; r <= norm.end.row; r++) next[r] = true;
            return { ...s, hiddenRows: next };
          });
        }}
        onUnhideRows={() => updateActiveSheet((s) => ({ ...s, hiddenRows: {} }))}
        onHideCol={() => {
          const norm = normalizeRange(selection);
          updateActiveSheet((s) => {
            const next = { ...s.hiddenCols };
            for (let c = norm.start.col; c <= norm.end.col; c++) next[c] = true;
            return { ...s, hiddenCols: next };
          });
        }}
        onUnhideCols={() => updateActiveSheet((s) => ({ ...s, hiddenCols: {} }))}
        onAutoSum={handleAutoSum}
        onFillDown={() => {
          const norm = normalizeRange(selection);
          handleFillDrag({ start: norm.start, end: { row: norm.end.row + 1, col: norm.end.col } });
        }}
        onFillRight={() => {
          const norm = normalizeRange(selection);
          handleFillDrag({ start: norm.start, end: { row: norm.end.row, col: norm.end.col + 1 } });
        }}
        onFillSeries={() => {
          const norm = normalizeRange(selection);
          if (norm.end.row > norm.start.row) {
            const startVal = Number(activeSheet.cells[cellKey(norm.start.row, norm.start.col)]?.raw || '1');
            if (!isNaN(startVal)) {
              updateActiveSheet((s) => {
                const nextCells = { ...s.cells };
                for (let r = norm.start.row; r <= norm.end.row; r++) {
                  nextCells[cellKey(r, norm.start.col)] = {
                    ...(nextCells[cellKey(r, norm.start.col)] || { raw: '' }),
                    raw: String(startVal + (r - norm.start.row)),
                  };
                }
                return { ...s, cells: nextCells };
              });
            }
          }
        }}
        onClear={handleClear}
        onSort={(asc) => handleSort(activeCell.col, asc)}
        onOpenCustomSort={() => setActiveModal({ type: 'customSort' })}
        onToggleFilter={() => {
          updateActiveSheet((s) => {
            if (s.autoFilter) return { ...s, autoFilter: null };
            const norm = normalizeRange(selection);
            const isSingle = norm.start.row === norm.end.row && norm.start.col === norm.end.col;
            let rangeStr = formatRangeRef(norm);
            if (isSingle) {
              let maxC = norm.start.col;
              let maxR = norm.start.row;
              for (const k of Object.keys(s.cells)) {
                const { row, col } = parseCellKey(k);
                if (s.cells[k]?.raw) {
                  if (col > maxC) maxC = col;
                  if (row > maxR) maxR = row;
                }
              }
              rangeStr = `${formatCellRef(norm.start.row, norm.start.col)}:${formatCellRef(maxR, maxC)}`;
            }
            return { ...s, autoFilter: { range: rangeStr, hiddenValuesByCol: {} } };
          });
        }}
        onClearFilterValues={() =>
          updateActiveSheet((s) =>
            s.autoFilter ? { ...s, autoFilter: { ...s.autoFilter, hiddenValuesByCol: {} } } : s
          )
        }
        onOpenFindReplace={() => setActiveModal({ type: 'find' })}
        onOpenPivotModal={() => setActiveModal({ type: 'pivot' })}
        onInsertChart={handleInsertChart}
        onInsertSparkline={(sparkType) => setActiveModal({ type: 'sparkline', sparkType })}
        onOpenCommentModal={() => setActiveModal({ type: 'comment' })}
        onDeleteComment={() =>
          updateActiveSheet((s) => {
            const k = cellKey(activeCell.row, activeCell.col);
            if (!s.cells[k]) return s;
            const next = { ...s.cells };
            next[k] = { ...next[k], comment: undefined };
            return { ...s, cells: next };
          })
        }
        onToggleShowAllComments={() => setShowAllComments((v) => !v)}
        onOpenHyperlinkModal={() => setActiveModal({ type: 'hyperlink' })}
        onToggleCheckbox={() => {
          const norm = normalizeRange(selection);
          updateActiveSheet((s) => {
            const nextCells = { ...s.cells };
            for (let r = norm.start.row; r <= norm.end.row; r++) {
              for (let c = norm.start.col; c <= norm.end.col; c++) {
                const k = cellKey(r, c);
                const prev = nextCells[k] || { raw: '' };
                const nextFlag = !prev.isCheckbox;
                nextCells[k] = {
                  ...prev,
                  isCheckbox: nextFlag,
                  raw: nextFlag ? (prev.raw.toUpperCase() === 'TRUE' ? 'TRUE' : 'FALSE') : prev.raw,
                };
              }
            }
            return { ...s, cells: nextCells };
          });
        }}
        onOpenValidationModal={() => setActiveModal({ type: 'validation' })}
        onUpdatePageSetup={(patch: Partial<PageSetup>) =>
          updateActiveSheet((s) => ({ ...s, pageSetup: { ...s.pageSetup, ...patch } }))
        }
        onSetPrintArea={() => {
          const ref = formatRangeRef(selection);
          updateActiveSheet((s) => ({ ...s, pageSetup: { ...s.pageSetup, printArea: ref } }));
          showToast(`Print Area set to ${ref}`);
        }}
        onClearPrintArea={() => {
          updateActiveSheet((s) => ({ ...s, pageSetup: { ...s.pageSetup, printArea: undefined } }));
          showToast('Print Area cleared.');
        }}
        onPrint={() => window.print()}
        onOpenFunctionWizard={() => setActiveModal({ type: 'fx' })}
        onInsertFormulaTemplate={(fnName) => {
          handleStartEdit(`=${fnName}(`);
          setTimeout(() => formulaInputRef.current?.focus(), 20);
        }}
        onOpenNameManager={() => setActiveModal({ type: 'names' })}
        onDefineNameQuick={() => setActiveModal({ type: 'names' })}
        onToggleShowFormulas={() =>
          commitWorkbookMutation((wb) => ({ ...wb, showFormulas: !wb.showFormulas }), false)
        }
        onTracePrecedents={() => {
          const precs = evaluator.getPrecedents(activeSheet.id, activeCell.row, activeCell.col);
          const sameSheet = precs
            .filter((p) => p.sheetId === activeSheet.id)
            .map((p) => ({ row: p.row, col: p.col }));
          setAuditHighlights((h) => ({ ...h, precedents: sameSheet }));
          showToast(`Found ${precs.length} precedent cell(s).`);
        }}
        onTraceDependents={() => {
          const deps = evaluator.getDependents(activeSheet.id, activeCell.row, activeCell.col);
          const sameSheet = deps
            .filter((d) => d.sheetId === activeSheet.id)
            .map((d) => ({ row: d.row, col: d.col }));
          setAuditHighlights((h) => ({ ...h, dependents: sameSheet }));
          showToast(`Found ${deps.length} dependent cell(s).`);
        }}
        onClearAuditArrows={() => setAuditHighlights({ precedents: [], dependents: [] })}
        onToggleCalcMode={() =>
          commitWorkbookMutation(
            (wb) => ({ ...wb, calcMode: wb.calcMode === 'automatic' ? 'manual' : 'automatic' }),
            false
          )
        }
        onCalculateNow={() => {
          setManualCalcVersion((v) => v + 1);
          showToast('Recalculation complete.');
        }}
        onImportFileClick={() => fileInputRef.current?.click()}
        onOpenTextToColumns={() => setActiveModal({ type: 'textToCols' })}
        onOpenRemoveDuplicates={() => setActiveModal({ type: 'removeDupes' })}
        onOpenGoalSeek={() => setActiveModal({ type: 'goalSeek' })}
        onSetFreezePanes={(mode) => {
          updateActiveSheet((s) => {
            if (mode === 'topRow') return { ...s, freezePanes: { row: 1, col: 0 } };
            if (mode === 'firstCol') return { ...s, freezePanes: { row: 0, col: 1 } };
            if (mode === 'selection') return { ...s, freezePanes: { row: activeCell.row, col: activeCell.col } };
            return { ...s, freezePanes: { row: 0, col: 0 } };
          });
        }}
        onToggleProtectSheet={() => {
          updateActiveSheet((s) => ({ ...s, protected: !s.protected }));
          showToast(activeSheet.protected ? 'Sheet unprotected.' : 'Sheet protected (read-only).');
        }}
        onOpenWorkbookStats={() => setActiveModal({ type: 'stats' })}
        onToggleFormulaBar={() =>
          commitWorkbookMutation((wb) => ({ ...wb, showFormulaBar: !wb.showFormulaBar }), false)
        }
        onSetZoom={(z) => commitWorkbookMutation((wb) => ({ ...wb, zoom: z }), false)}
        onOpenFormatCellsModal={() => setActiveModal({ type: 'formatCells' })}
      />

      <FormulaBar
        visible={workbook.showFormulaBar}
        cellAddress={formatCellRef(activeCell.row, activeCell.col)}
        formulaText={formulaBarText}
        isEditing={isEditing}
        namedRanges={workbook.namedRanges}
        onJumpToReference={handleJumpToReference}
        onStartEdit={() => handleStartEdit()}
        onChangeFormula={(val) => {
          lastInsertedRefSpan.current = null;
          setEditDraft(val);
        }}
        onCommitEdit={() => handleCommitEdit('down')}
        onCancelEdit={handleCancelEdit}
        onOpenFunctionWizard={() => setActiveModal({ type: 'fx' })}
        inputRef={formulaInputRef}
      />

      <SpreadsheetGrid
        sheet={activeSheet}
        sheets={workbook.sheets}
        namedRanges={workbook.namedRanges}
        evaluator={evaluator}
        currencySymbol={workbook.currencySymbol}
        showFormulas={workbook.showFormulas}
        zoom={workbook.zoom}
        pageBreakView={pageBreakView}
        showAllComments={showAllComments}
        selection={selection}
        extraSelections={extraSelections}
        activeCell={activeCell}
        isEditing={isEditing}
        editDraft={editDraft}
        clipboard={clipboard}
        formulaHighlights={formulaHighlights}
        auditHighlights={auditHighlights}
        selectedChartId={selectedChartId}
        onSelectChart={setSelectedChartId}
        onUpdateChart={(id, patch) =>
          updateActiveSheet((s) => ({
            ...s,
            charts: s.charts.map((ch) => (ch.id === id ? { ...ch, ...patch } : ch)),
          }))
        }
        onDeleteChart={(id) => {
          updateActiveSheet((s) => ({ ...s, charts: s.charts.filter((ch) => ch.id !== id) }));
          setSelectedChartId(null);
        }}
        onSelectRange={handleSelectRange}
        onStartEdit={handleStartEdit}
        onChangeEditDraft={(val) => {
          lastInsertedRefSpan.current = null;
          setEditDraft(val);
        }}
        onCommitEdit={handleCommitEdit}
        onCancelEdit={handleCancelEdit}
        onInsertRefIntoFormula={handleInsertRefIntoFormula}
        onFillDrag={handleFillDrag}
        onSetColWidth={(col, width) =>
          updateActiveSheet((s) => ({ ...s, colWidths: { ...s.colWidths, [col]: width } }), false)
        }
        onSetRowHeight={(row, height) =>
          updateActiveSheet((s) => ({ ...s, rowHeights: { ...s.rowHeights, [row]: height } }), false)
        }
        onAutoFitCol={(col) => {
          let maxLen = 6;
          for (let r = 0; r < activeSheet.rowCount; r++) {
            const ev = evaluator.evaluateCell(activeSheet.id, r, col);
            if (ev.formatted.length > maxLen) maxLen = ev.formatted.length;
          }
          updateActiveSheet((s) => ({
            ...s,
            colWidths: { ...s.colWidths, [col]: Math.min(360, Math.max(60, maxLen * 8 + 20)) },
          }));
        }}
        onToggleCheckboxCell={(row, col) => {
          const cur = activeSheet.cells[cellKey(row, col)];
          const nextVal = String(cur?.raw).toUpperCase() === 'TRUE' ? 'FALSE' : 'TRUE';
          setCellRaw(row, col, nextVal);
        }}
        onSetCellDirect={setCellRaw}
        onUpdateAutoFilterCol={(col, hiddenValues) =>
          updateActiveSheet((s) =>
            s.autoFilter
              ? {
                  ...s,
                  autoFilter: {
                    ...s.autoFilter,
                    hiddenValuesByCol: { ...s.autoFilter.hiddenValuesByCol, [col]: hiddenValues },
                  },
                }
              : s
          )
        }
        onSortCol={(col, asc) => handleSort(col, asc)}
        onCut={() => handleCopyCut('cut')}
        onCopy={() => handleCopyCut('copy')}
        onPaste={() => handlePaste(false)}
        onPasteValues={() => handlePaste(true)}
        onInsertRowAt={handleInsertRowAt}
        onInsertColAt={handleInsertColAt}
        onDeleteRowAt={handleDeleteRowAt}
        onDeleteColAt={handleDeleteColAt}
        onClearSelection={() => handleClear('contents')}
        onMergeCenter={handleMergeCenter}
        onOpenFormatCells={() => setActiveModal({ type: 'formatCells' })}
        onOpenValidation={() => setActiveModal({ type: 'validation' })}
        onOpenComment={() => setActiveModal({ type: 'comment' })}
        onDefineName={() => setActiveModal({ type: 'names' })}
        onAddMoreRows={() =>
          updateActiveSheet((s) => ({ ...s, rowCount: s.rowCount + 50 }))
        }
      />

      <SheetTabsBar
        sheets={workbook.sheets}
        activeSheetId={activeSheet.id}
        onSelectSheet={(id) => {
          if (isEditing) handleCommitEdit('none');
          setWorkbook((wb) => ({ ...wb, activeSheetId: id }));
          setActiveCell({ row: 0, col: 0 });
          setSelection({ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
        }}
        onAddSheet={() => {
          const idx = workbook.sheets.length + 1;
          const ns = createEmptySheet(`sheet-${Date.now()}`, `Sheet${idx}`);
          commitWorkbookMutation((wb) => ({
            ...wb,
            sheets: [...wb.sheets, ns],
            activeSheetId: ns.id,
          }));
        }}
        onRenameSheet={(id, name) =>
          commitWorkbookMutation((wb) => ({
            ...wb,
            sheets: wb.sheets.map((s) => (s.id === id ? { ...s, name } : s)),
          }))
        }
        onDuplicateSheet={(id) => {
          const orig = workbook.sheets.find((s) => s.id === id);
          if (!orig) return;
          const copy: Sheet = {
            ...JSON.parse(JSON.stringify(orig)),
            id: `sheet-${Date.now()}`,
            name: `${orig.name} (Copy)`,
          };
          commitWorkbookMutation((wb) => ({
            ...wb,
            sheets: [...wb.sheets, copy],
            activeSheetId: copy.id,
          }));
        }}
        onDeleteSheet={(id) => {
          if (workbook.sheets.length <= 1) return;
          commitWorkbookMutation((wb) => {
            const next = wb.sheets.filter((s) => s.id !== id);
            return {
              ...wb,
              sheets: next,
              activeSheetId: wb.activeSheetId === id ? next[0].id : wb.activeSheetId,
            };
          });
        }}
        onSetTabColor={(id, color) =>
          commitWorkbookMutation((wb) => ({
            ...wb,
            sheets: wb.sheets.map((s) => (s.id === id ? { ...s, tabColor: color } : s)),
          }))
        }
        onMoveSheet={(id, dir) => {
          commitWorkbookMutation((wb) => {
            const idx = wb.sheets.findIndex((s) => s.id === id);
            const nextIdx = idx + dir;
            if (idx === -1 || nextIdx < 0 || nextIdx >= wb.sheets.length) return wb;
            const next = [...wb.sheets];
            const [item] = next.splice(idx, 1);
            next.splice(nextIdx, 0, item);
            return { ...wb, sheets: next };
          });
        }}
      />

      <StatusBar
        mode={isEditing ? (editDraft.startsWith('=') ? 'Point' : 'Edit') : 'Ready'}
        sheetProtected={activeSheet.protected}
        calcMode={workbook.calcMode}
        summary={selectionSummary}
        zoom={workbook.zoom}
        onChangeZoom={(z) => commitWorkbookMutation((wb) => ({ ...wb, zoom: z }), false)}
        onCalculateNow={() => {
          setManualCalcVersion((v) => v + 1);
          showToast('Recalculated workbook.');
        }}
        showPageLayoutHint={pageBreakView}
        onTogglePageLayoutView={() => setPageBreakView((v) => !v)}
      />

      {/* Non-blocking Toast Banner */}
      {toastMsg && (
        <div className="no-print fixed bottom-16 left-1/2 -translate-x-1/2 z-50 bg-slate-900 dark:bg-slate-800 text-white text-xs font-medium px-4 py-2 rounded-md shadow-lg border border-slate-700">
          {toastMsg}
        </div>
      )}

      {/* Modals */}
      {activeModal?.type === 'file' && (
        <FileBackstageModal
          workbook={workbook}
          onClose={() => setActiveModal(null)}
          onLoadWorkbook={(wb) => {
            commitWorkbookMutation(() => wb);
            setActiveCell({ row: 0, col: 0 });
            setSelection({ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
          }}
          onImportFile={() => {
            setActiveModal(null);
            fileInputRef.current?.click();
          }}
          onExportXLSX={() => exportWorkbookToXLSX(workbook)}
          onExportCSV={() => exportSheetToCSV(workbook, activeSheet.id)}
          onExportJSON={() => exportWorkbookToJSON(workbook)}
          onPrint={() => window.print()}
          onGoHome={() => {
            if (isEditing) handleCommitEdit('none');
            setViewMode('home');
          }}
        />
      )}

      {activeModal?.type === 'find' && (
        <FindReplaceModal
          sheet={activeSheet}
          evaluator={evaluator}
          onClose={() => setActiveModal(null)}
          onJumpToCell={(row, col) => {
            setActiveCell({ row, col });
            setSelection({ start: { row, col }, end: { row, col } });
          }}
          onReplaceCell={setCellRaw}
        />
      )}

      {activeModal?.type === 'fx' && (
        <FunctionWizardModal
          onClose={() => setActiveModal(null)}
          onInsertFunction={(fnName) => {
            handleStartEdit(`=${fnName}(`);
            setTimeout(() => formulaInputRef.current?.focus(), 20);
          }}
        />
      )}

      {activeModal?.type === 'formatCells' && (
        <FormatCellsModal
          initialStyle={activeCellStyle}
          onClose={() => setActiveModal(null)}
          onApply={handleUpdateStyle}
        />
      )}

      {activeModal?.type === 'conditional' && (
        <ConditionalFormatModal
          rules={activeSheet.conditionalRules}
          defaultRange={formatRangeRef(selection)}
          initialType={activeModal.initialType}
          onClose={() => setActiveModal(null)}
          onAddRule={(rule: ConditionalRule) =>
            updateActiveSheet((s) => ({ ...s, conditionalRules: [...s.conditionalRules, rule] }))
          }
          onDeleteRule={(id) =>
            updateActiveSheet((s) => ({
              ...s,
              conditionalRules: s.conditionalRules.filter((r) => r.id !== id),
            }))
          }
        />
      )}

      {activeModal?.type === 'validation' && (
        <DataValidationModal
          existing={activeCellData?.validation}
          onClose={() => setActiveModal(null)}
          onSave={(rule: DataValidationRule | undefined) => {
            const norm = normalizeRange(selection);
            updateActiveSheet((s) => {
              const nextCells = { ...s.cells };
              for (let r = norm.start.row; r <= norm.end.row; r++) {
                for (let c = norm.start.col; c <= norm.end.col; c++) {
                  const k = cellKey(r, c);
                  nextCells[k] = { ...(nextCells[k] || { raw: '' }), validation: rule };
                }
              }
              return { ...s, cells: nextCells };
            });
          }}
        />
      )}

      {activeModal?.type === 'pivot' && (
        <PivotTableModal
          sheet={activeSheet}
          evaluator={evaluator}
          defaultRange={
            selection.start.row === selection.end.row && selection.start.col === selection.end.col
              ? 'A1:G7'
              : formatRangeRef(selection)
          }
          onClose={() => setActiveModal(null)}
          onCreatePivot={({ headers, rows, targetNewSheet }) => {
            const fillSheetCells = (target: Sheet, startR: number, startC: number) => {
              const nextCells = { ...target.cells };
              headers.forEach((h, cIdx) => {
                nextCells[cellKey(startR, startC + cIdx)] = {
                  raw: h,
                  style: { bold: true, color: '#ffffff', backgroundColor: '#107c41' },
                };
              });
              rows.forEach((rowArr, rIdx) => {
                const isGrand = rIdx === rows.length - 1;
                rowArr.forEach((val, cIdx) => {
                  nextCells[cellKey(startR + 1 + rIdx, startC + cIdx)] = {
                    raw: String(val),
                    style: {
                      bold: isGrand,
                      backgroundColor: isGrand ? '#ecfdf5' : undefined,
                    },
                  };
                });
              });
              return {
                ...target,
                colWidths: { ...target.colWidths, [startC]: 160, [startC + 1]: 140 },
                cells: nextCells,
              };
            };

            if (targetNewSheet) {
              const pSheet = createEmptySheet(`sheet-${Date.now()}`, `Pivot_${workbook.sheets.length + 1}`);
              const filled = fillSheetCells(pSheet, 1, 0);
              filled.cells[cellKey(0, 0)] = {
                raw: 'PivotTable Summary',
                style: { bold: true, fontSize: 14, color: '#107c41' },
              };
              commitWorkbookMutation((wb) => ({
                ...wb,
                sheets: [...wb.sheets, filled],
                activeSheetId: filled.id,
              }));
            } else {
              updateActiveSheet((s) => fillSheetCells(s, activeCell.row, activeCell.col));
            }
            showToast('PivotTable generated.');
          }}
        />
      )}

      {activeModal?.type === 'names' && (
        <NameManagerModal
          namedRanges={workbook.namedRanges}
          activeSheetId={activeSheet.id}
          defaultRange={formatRangeRef(selection)}
          sheets={workbook.sheets}
          onClose={() => setActiveModal(null)}
          onAdd={(nr) =>
            commitWorkbookMutation((wb) => ({
              ...wb,
              namedRanges: [...wb.namedRanges.filter((x) => x.name !== nr.name), nr],
            }))
          }
          onDelete={(name) =>
            commitWorkbookMutation((wb) => ({
              ...wb,
              namedRanges: wb.namedRanges.filter((x) => x.name !== name),
            }))
          }
        />
      )}

      {activeModal?.type === 'goalSeek' && (
        <GoalSeekModal
          sheets={workbook.sheets}
          namedRanges={workbook.namedRanges}
          activeSheetId={activeSheet.id}
          defaultSetCell={formatCellRef(activeCell.row, activeCell.col)}
          onClose={() => setActiveModal(null)}
          onApplyValue={(r, c, val) => setCellRaw(r, c, val)}
        />
      )}

      {activeModal?.type === 'textToCols' && (
        <TextToColumnsModal
          onClose={() => setActiveModal(null)}
          onExecute={(delimiter) => {
            const norm = normalizeRange(selection);
            updateActiveSheet((s) => {
              const nextCells = { ...s.cells };
              for (let r = norm.start.row; r <= norm.end.row; r++) {
                const raw = s.cells[cellKey(r, norm.start.col)]?.raw || '';
                if (!raw) continue;
                const parts = raw.split(delimiter);
                parts.forEach((p, idx) => {
                  const k = cellKey(r, norm.start.col + idx);
                  nextCells[k] = { ...(nextCells[k] || { raw: '' }), raw: p.trim() };
                });
              }
              return { ...s, cells: nextCells };
            });
            showToast('Split text into columns.');
          }}
        />
      )}

      {activeModal?.type === 'removeDupes' && (
        <RemoveDuplicatesModal
          selectionCols={selectionColsMeta}
          onClose={() => setActiveModal(null)}
          onExecute={(selectedCols, hasHeaders) => {
            const norm = normalizeRange(selection);
            const startR = hasHeaders ? norm.start.row + 1 : norm.start.row;
            const seen = new Set<string>();
            const keptRows: Record<number, CellData>[] = [];
            let removed = 0;

            for (let r = startR; r <= norm.end.row; r++) {
              const sig = selectedCols
                .map((c) => evaluator.evaluateCell(activeSheet.id, r, c).formatted)
                .join('||');
              if (seen.has(sig)) {
                removed++;
              } else {
                seen.add(sig);
                const rowMap: Record<number, CellData> = {};
                for (let c = norm.start.col; c <= norm.end.col; c++) {
                  const cd = activeSheet.cells[cellKey(r, c)];
                  if (cd) rowMap[c] = cd;
                }
                keptRows.push(rowMap);
              }
            }

            updateActiveSheet((s) => {
              const nextCells = { ...s.cells };
              for (let r = startR; r <= norm.end.row; r++) {
                for (let c = norm.start.col; c <= norm.end.col; c++) {
                  delete nextCells[cellKey(r, c)];
                }
              }
              keptRows.forEach((rowMap, idx) => {
                for (const [cStr, cd] of Object.entries(rowMap)) {
                  nextCells[cellKey(startR + idx, Number(cStr))] = cd;
                }
              });
              return { ...s, cells: nextCells };
            });
            showToast(`${removed} duplicate row(s) removed; ${keptRows.length} unique row(s) remain.`);
          }}
        />
      )}

      {activeModal?.type === 'customSort' && (
        <CustomSortModal
          selectionCols={selectionColsMeta}
          onClose={() => setActiveModal(null)}
          onExecute={(colIdx, asc, hasHeaders) => handleSort(colIdx, asc, hasHeaders)}
        />
      )}

      {activeModal?.type === 'commandPalette' && (
        <CommandPaletteModal
          onClose={() => setActiveModal(null)}
          commands={commandList}
          onJumpToRef={handleJumpToReference}
        />
      )}

      {activeModal?.type === 'comment' && (
        <SimpleInputModal
          title={`Comment for ${formatCellRef(activeCell.row, activeCell.col)}`}
          label="Note / Comment"
          initialValue={activeCellData?.comment || ''}
          placeholder="Add a note for this cell..."
          onClose={() => setActiveModal(null)}
          onConfirm={(text) => {
            updateActiveSheet((s) => {
              const k = cellKey(activeCell.row, activeCell.col);
              const prev = s.cells[k] || { raw: '' };
              return { ...s, cells: { ...s.cells, [k]: { ...prev, comment: text.trim() || undefined } } };
            });
          }}
        />
      )}

      {activeModal?.type === 'hyperlink' && (
        <SimpleInputModal
          title={`Insert Hyperlink in ${formatCellRef(activeCell.row, activeCell.col)}`}
          label="URL (https://...)"
          initialValue={activeCellData?.hyperlink || 'https://'}
          onClose={() => setActiveModal(null)}
          onConfirm={(url) => {
            updateActiveSheet((s) => {
              const k = cellKey(activeCell.row, activeCell.col);
              const prev = s.cells[k] || { raw: '' };
              return {
                ...s,
                cells: {
                  ...s.cells,
                  [k]: { ...prev, raw: prev.raw || url, hyperlink: url.trim() || undefined },
                },
              };
            });
          }}
        />
      )}

      {activeModal?.type === 'size' && (
        <SimpleInputModal
          title={activeModal.kind === 'col' ? 'Column Width' : 'Row Height'}
          label={activeModal.kind === 'col' ? 'Width (pixels)' : 'Height (pixels)'}
          initialValue={String(
            activeModal.kind === 'col'
              ? activeSheet.colWidths[activeCell.col] ?? 100
              : activeSheet.rowHeights[activeCell.row] ?? 24
          )}
          onClose={() => setActiveModal(null)}
          onConfirm={(val) => {
            const num = Math.max(20, Math.min(500, Number(val) || 100));
            const norm = normalizeRange(selection);
            updateActiveSheet((s) => {
              if (activeModal.kind === 'col') {
                const next = { ...s.colWidths };
                for (let c = norm.start.col; c <= norm.end.col; c++) next[c] = num;
                return { ...s, colWidths: next };
              } else {
                const next = { ...s.rowHeights };
                for (let r = norm.start.row; r <= norm.end.row; r++) next[r] = num;
                return { ...s, rowHeights: next };
              }
            });
          }}
        />
      )}

      {activeModal?.type === 'sparkline' && (
        <SimpleInputModal
          title={`Insert ${activeModal.sparkType.toUpperCase()} Sparkline`}
          label="Data Range (e.g. C2:F2)"
          initialValue="C2:F2"
          onClose={() => setActiveModal(null)}
          onConfirm={(range) => {
            updateActiveSheet((s) => {
              const k = cellKey(activeCell.row, activeCell.col);
              const prev = s.cells[k] || { raw: '' };
              return {
                ...s,
                cells: {
                  ...s.cells,
                  [k]: {
                    ...prev,
                    sparkline: { type: activeModal.sparkType, range: range.toUpperCase(), color: '#107c41' },
                  },
                },
              };
            });
          }}
        />
      )}

      {activeModal?.type === 'stats' && (
        <WorkbookStatsModal workbook={workbook} onClose={() => setActiveModal(null)} />
      )}
    </div>
  );
}
