import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  ArrowDownAZ,
  ArrowUpZA,
  ChevronDown,
  ClipboardPaste,
  Copy,
  Eraser,
  Filter,
  Merge,
  MessageSquarePlus,
  Plus,
  Scissors,
  ShieldCheck,
  SlidersHorizontal,
  Tag,
  Trash2,
} from 'lucide-react';
import {
  CellCoord,
  CellData,
  ClipboardData,
  EvaluatedCell,
  NamedRange,
  SelectionRange,
  Sheet,
  SheetChart,
} from '../types/spreadsheet';
import {
  cellKey,
  colIndexToLetter,
  formatCellRef,
  isCellInRange,
  normalizeRange,
  parseRangeRef,
} from '../utils/cellCoords';
import { FormulaRefHighlight, WorkbookEvaluator } from '../engine/formulaEngine';
import { FloatingChart } from './FloatingChart';

const DEFAULT_COL_WIDTH = 100;
const DEFAULT_ROW_HEIGHT = 24;
const ROW_HEADER_WIDTH = 46;
const COL_HEADER_HEIGHT = 24;

const REF_COLORS = ['#2563eb', '#dc2626', '#7c3aed', '#d97706', '#0d9488', '#db2777'];

interface SpreadsheetGridProps {
  sheet: Sheet;
  sheets: Sheet[];
  namedRanges: NamedRange[];
  evaluator: WorkbookEvaluator;
  currencySymbol: string;
  showFormulas: boolean;
  zoom: number;
  pageBreakView: boolean;
  showAllComments: boolean;
  selection: SelectionRange;
  extraSelections: SelectionRange[];
  activeCell: CellCoord;
  isEditing: boolean;
  editDraft: string;
  clipboard: ClipboardData | null;
  formulaHighlights: FormulaRefHighlight[];
  auditHighlights: { precedents: CellCoord[]; dependents: CellCoord[] };
  selectedChartId: string | null;
  onSelectChart: (id: string | null) => void;
  onUpdateChart: (id: string, patch: Partial<SheetChart>) => void;
  onDeleteChart: (id: string) => void;
  onSelectRange: (range: SelectionRange, active?: CellCoord, appendMulti?: boolean) => void;
  onStartEdit: (initialChar?: string) => void;
  onChangeEditDraft: (val: string) => void;
  onCommitEdit: (moveDir?: 'down' | 'right' | 'left' | 'up' | 'none') => void;
  onCancelEdit: () => void;
  onInsertRefIntoFormula: (refText: string) => void;
  onFillDrag: (targetRange: SelectionRange) => void;
  onSetColWidth: (col: number, width: number) => void;
  onSetRowHeight: (row: number, height: number) => void;
  onAutoFitCol: (col: number) => void;
  onToggleCheckboxCell: (row: number, col: number) => void;
  onSetCellDirect: (row: number, col: number, raw: string) => void;
  onUpdateAutoFilterCol: (col: number, hiddenValues: string[]) => void;
  onSortCol: (col: number, asc: boolean) => void;
  onCut: () => void;
  onCopy: () => void;
  onPaste: () => void;
  onPasteValues: () => void;
  onInsertRowAt: (row: number) => void;
  onInsertColAt: (col: number) => void;
  onDeleteRowAt: (row: number) => void;
  onDeleteColAt: (col: number) => void;
  onClearSelection: () => void;
  onMergeCenter: () => void;
  onOpenFormatCells: () => void;
  onOpenValidation: () => void;
  onOpenComment: () => void;
  onDefineName: () => void;
  onAddMoreRows: () => void;
}

export const SpreadsheetGrid: React.FC<SpreadsheetGridProps> = (props) => {
  const {
    sheet,
    sheets,
    namedRanges,
    evaluator,
    currencySymbol,
    showFormulas,
    zoom,
    pageBreakView,
    showAllComments,
    selection,
    extraSelections,
    activeCell,
    isEditing,
    editDraft,
    clipboard,
    formulaHighlights,
    auditHighlights,
    selectedChartId,
  } = props;

  const containerRef = useRef<HTMLDivElement>(null);
  const inlineInputRef = useRef<HTMLInputElement>(null);
  const [fillPreviewRange, setFillPreviewRange] = useState<SelectionRange | null>(null);
  const [validationDropdownCell, setValidationDropdownCell] = useState<CellCoord | null>(null);
  const [filterPopover, setFilterPopover] = useState<{ col: number; x: number; y: number } | null>(null);
  const [filterSearch, setFilterSearch] = useState('');
  const [hoveredCommentCell, setHoveredCommentCell] = useState<CellCoord | null>(null);
  const [ctxMenu, setCtxMenu] = useState<{ x: number; y: number; row: number; col: number; kind: 'cell' | 'row' | 'col' } | null>(null);

  const isDraggingSelect = useRef(false);
  const dragStartCell = useRef<CellCoord | null>(null);
  const isPointingFormula = useRef(false);

  const scale = zoom / 100;

  const getColWidth = useCallback((c: number) => sheet.colWidths[c] ?? DEFAULT_COL_WIDTH, [sheet.colWidths]);
  const getRowHeight = useCallback((r: number) => sheet.rowHeights[r] ?? DEFAULT_ROW_HEIGHT, [sheet.rowHeights]);

  // Precompute cumulative positions for visible cols & rows
  const colPositions = useMemo(() => {
    const lefts: number[] = [];
    const widths: number[] = [];
    let acc = 0;
    for (let c = 0; c < sheet.colCount; c++) {
      const w = sheet.hiddenCols[c] ? 0 : getColWidth(c);
      lefts.push(acc);
      widths.push(w);
      acc += w;
    }
    return { lefts, widths, totalWidth: acc };
  }, [sheet.colCount, sheet.hiddenCols, getColWidth]);

  // Check AutoFilter row hiding
  const autoFilterParsed = useMemo(() => {
    if (!sheet.autoFilter) return null;
    return parseRangeRef(sheet.autoFilter.range);
  }, [sheet.autoFilter]);

  const filteredHiddenRows = useMemo(() => {
    const hidden = new Set<number>();
    if (!sheet.autoFilter || !autoFilterParsed) return hidden;
    const headerRow = autoFilterParsed.start.row;
    const maxRow = Math.max(autoFilterParsed.end.row, sheet.rowCount - 1);
    const entries = Object.entries(sheet.autoFilter.hiddenValuesByCol);
    if (entries.length === 0) return hidden;

    for (let r = headerRow + 1; r <= maxRow; r++) {
      // Only filter contiguous data rows or up to end of autoFilter / non-empty rows
      let rowHasAny = false;
      for (let c = autoFilterParsed.start.col; c <= autoFilterParsed.end.col; c++) {
        if (sheet.cells[cellKey(r, c)]?.raw) {
          rowHasAny = true;
          break;
        }
      }
      if (!rowHasAny && r > autoFilterParsed.end.row) continue;

      for (const [colStr, hiddenList] of entries) {
        if (!hiddenList || hiddenList.length === 0) continue;
        const cIdx = Number(colStr);
        const ev = evaluator.evaluateCell(sheet.id, r, cIdx);
        const txt = ev.formatted || '(Blanks)';
        if (hiddenList.includes(txt)) {
          hidden.add(r);
          break;
        }
      }
    }
    return hidden;
  }, [sheet.autoFilter, autoFilterParsed, sheet.rowCount, sheet.cells, sheet.id, evaluator]);

  const rowPositions = useMemo(() => {
    const tops: number[] = [];
    const heights: number[] = [];
    let acc = 0;
    for (let r = 0; r < sheet.rowCount; r++) {
      const isHidden = sheet.hiddenRows[r] || filteredHiddenRows.has(r);
      const h = isHidden ? 0 : getRowHeight(r);
      tops.push(acc);
      heights.push(h);
      acc += h;
    }
    return { tops, heights, totalHeight: acc };
  }, [sheet.rowCount, sheet.hiddenRows, filteredHiddenRows, getRowHeight]);

  // Merged cell lookup
  const mergeMap = useMemo(() => {
    const masters = new Map<string, { rowSpan: number; colSpan: number; width: number; height: number }>();
    const slaves = new Set<string>();
    for (const m of sheet.mergedRanges) {
      let w = 0;
      for (let c = m.startCol; c <= m.endCol; c++) w += colPositions.widths[c] || 0;
      let h = 0;
      for (let r = m.startRow; r <= m.endRow; r++) h += rowPositions.heights[r] || 0;
      masters.set(cellKey(m.startRow, m.startCol), {
        rowSpan: m.endRow - m.startRow + 1,
        colSpan: m.endCol - m.startCol + 1,
        width: w,
        height: h,
      });
      for (let r = m.startRow; r <= m.endRow; r++) {
        for (let c = m.startCol; c <= m.endCol; c++) {
          if (r !== m.startRow || c !== m.startCol) slaves.add(cellKey(r, c));
        }
      }
    }
    return { masters, slaves };
  }, [sheet.mergedRanges, colPositions.widths, rowPositions.heights]);

  // Precompute Conditional Formatting stats (duplicates, min/max for data bars & color scales)
  const cfStats = useMemo(() => {
    const stats = new Map<
      string,
      { min: number; max: number; counts: Map<string, number> }
    >();
    for (const rule of sheet.conditionalRules) {
      const rng = parseRangeRef(rule.range);
      if (!rng) continue;
      let min = Infinity;
      let max = -Infinity;
      const counts = new Map<string, number>();
      for (let r = rng.start.row; r <= rng.end.row; r++) {
        for (let c = rng.start.col; c <= rng.end.col; c++) {
          const ev = evaluator.evaluateCell(sheet.id, r, c);
          if (typeof ev.value === 'number') {
            if (ev.value < min) min = ev.value;
            if (ev.value > max) max = ev.value;
          }
          const txt = ev.formatted;
          if (txt !== '') {
            counts.set(txt, (counts.get(txt) || 0) + 1);
          }
        }
      }
      stats.set(rule.id, { min, max, counts });
    }
    return stats;
  }, [sheet.conditionalRules, sheet.id, evaluator]);

  // Virtualized viewport window
  const [scrollPos, setScrollPos] = useState({ top: 0, left: 0, width: 1200, height: 700 });

  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    const updateSize = () => {
      setScrollPos((prev) => ({
        ...prev,
        width: el.clientWidth,
        height: el.clientHeight,
      }));
    };
    updateSize();
    const ro = new ResizeObserver(updateSize);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const handleScroll = (e: React.UIEvent<HTMLDivElement>) => {
    const el = e.currentTarget;
    setScrollPos((prev) => ({
      ...prev,
      top: el.scrollTop / scale,
      left: el.scrollLeft / scale,
    }));
  };

  // Determine visible row/col index ranges with generous overscan
  const visibleRowBounds = useMemo(() => {
    const overscan = 10;
    const topTarget = Math.max(0, scrollPos.top - 100);
    const botTarget = scrollPos.top + scrollPos.height / scale + 200;
    let start = sheet.freezePanes.row;
    while (start < sheet.rowCount && rowPositions.tops[start] + rowPositions.heights[start] < topTarget) {
      start++;
    }
    let end = start;
    while (end < sheet.rowCount && rowPositions.tops[end] < botTarget) {
      end++;
    }
    return {
      start: Math.max(sheet.freezePanes.row, start - overscan),
      end: Math.min(sheet.rowCount - 1, end + overscan),
    };
  }, [scrollPos.top, scrollPos.height, scale, sheet.freezePanes.row, sheet.rowCount, rowPositions]);

  const visibleColBounds = useMemo(() => {
    const overscan = 4;
    const leftTarget = Math.max(0, scrollPos.left - 120);
    const rightTarget = scrollPos.left + scrollPos.width / scale + 200;
    let start = sheet.freezePanes.col;
    while (start < sheet.colCount && colPositions.lefts[start] + colPositions.widths[start] < leftTarget) {
      start++;
    }
    let end = start;
    while (end < sheet.colCount && colPositions.lefts[end] < rightTarget) {
      end++;
    }
    return {
      start: Math.max(sheet.freezePanes.col, start - overscan),
      end: Math.min(sheet.colCount - 1, end + overscan),
    };
  }, [scrollPos.left, scrollPos.width, scale, sheet.freezePanes.col, sheet.colCount, colPositions]);

  // Focus inline input when editing starts
  useEffect(() => {
    if (isEditing && inlineInputRef.current) {
      // Only focus inline input if formula bar isn't focused
      if (document.activeElement?.tagName !== 'INPUT') {
        inlineInputRef.current.focus();
        const len = inlineInputRef.current.value.length;
        inlineInputRef.current.setSelectionRange(len, len);
      }
    }
  }, [isEditing]);

  // Close menus on outside click
  useEffect(() => {
    const close = () => {
      setCtxMenu(null);
      setFilterPopover(null);
      setValidationDropdownCell(null);
    };
    window.addEventListener('click', close);
    return () => window.removeEventListener('click', close);
  }, []);

  // Scroll active cell into view when moved via keyboard
  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    const cellTop = rowPositions.tops[activeCell.row] * scale;
    const cellBot = (rowPositions.tops[activeCell.row] + rowPositions.heights[activeCell.row]) * scale;
    const cellLeft = colPositions.lefts[activeCell.col] * scale;
    const cellRight = (colPositions.lefts[activeCell.col] + colPositions.widths[activeCell.col]) * scale;

    const freezeTopPx = (sheet.freezePanes.row > 0 ? rowPositions.tops[sheet.freezePanes.row] : 0) * scale;
    const freezeLeftPx = (sheet.freezePanes.col > 0 ? colPositions.lefts[sheet.freezePanes.col] : 0) * scale;

    const viewTop = el.scrollTop + COL_HEADER_HEIGHT * scale + freezeTopPx;
    const viewBot = el.scrollTop + el.clientHeight - 16;
    const viewLeft = el.scrollLeft + ROW_HEADER_WIDTH * scale + freezeLeftPx;
    const viewRight = el.scrollLeft + el.clientWidth - 16;

    if (activeCell.row >= sheet.freezePanes.row) {
      if (cellTop < viewTop) el.scrollTop = Math.max(0, cellTop - freezeTopPx);
      else if (cellBot > viewBot) el.scrollTop = cellBot - el.clientHeight + COL_HEADER_HEIGHT * scale + 20;
    }
    if (activeCell.col >= sheet.freezePanes.col) {
      if (cellLeft < viewLeft) el.scrollLeft = Math.max(0, cellLeft - freezeLeftPx);
      else if (cellRight > viewRight) el.scrollLeft = cellRight - el.clientWidth + ROW_HEADER_WIDTH * scale + 20;
    }
  }, [activeCell, scale, rowPositions, colPositions, sheet.freezePanes]);

  const handleCellMouseDown = (e: React.MouseEvent, row: number, col: number) => {
    if (e.button === 2) {
      if (!isCellInRange(row, col, selection)) {
        props.onSelectRange({ start: { row, col }, end: { row, col } }, { row, col });
      }
      return;
    }
    e.stopPropagation();
    props.onSelectChart(null);
    setCtxMenu(null);
    setFilterPopover(null);

    // If currently editing a formula starting with '=', point & click inserts cell reference!
    if (isEditing && editDraft.startsWith('=')) {
      e.preventDefault();
      isPointingFormula.current = true;
      dragStartCell.current = { row, col };
      props.onInsertRefIntoFormula(formatCellRef(row, col));
      const onUp = () => {
        isPointingFormula.current = false;
        dragStartCell.current = null;
        window.removeEventListener('mouseup', onUp);
      };
      window.addEventListener('mouseup', onUp);
      return;
    }

    if (isEditing) {
      props.onCommitEdit('none');
    }

    if (e.shiftKey) {
      props.onSelectRange({ start: activeCell, end: { row, col } }, activeCell);
      return;
    }

    const appendMulti = e.ctrlKey || e.metaKey;
    isDraggingSelect.current = true;
    dragStartCell.current = { row, col };
    props.onSelectRange({ start: { row, col }, end: { row, col } }, { row, col }, appendMulti);

    const onUp = () => {
      isDraggingSelect.current = false;
      dragStartCell.current = null;
      window.removeEventListener('mouseup', onUp);
    };
    window.addEventListener('mouseup', onUp);
  };

  const handleCellMouseEnter = (row: number, col: number) => {
    if (isPointingFormula.current && dragStartCell.current) {
      const s = dragStartCell.current;
      const refText =
        s.row === row && s.col === col
          ? formatCellRef(row, col)
          : `${formatCellRef(Math.min(s.row, row), Math.min(s.col, col))}:${formatCellRef(
              Math.max(s.row, row),
              Math.max(s.col, col)
            )}`;
      props.onInsertRefIntoFormula(refText);
      return;
    }

    if (isDraggingSelect.current && dragStartCell.current) {
      props.onSelectRange(
        { start: dragStartCell.current, end: { row, col } },
        dragStartCell.current
      );
    }
  };

  const startFillHandleDrag = (e: React.MouseEvent) => {
    e.stopPropagation();
    e.preventDefault();
    const norm = normalizeRange(selection);
    setFillPreviewRange(norm);
    let latestTarget = norm;

    const onMove = (me: MouseEvent) => {
      const el = containerRef.current;
      if (!el) return;
      const rect = el.getBoundingClientRect();
      const x = (me.clientX - rect.left + el.scrollLeft) / scale - ROW_HEADER_WIDTH;
      const y = (me.clientY - rect.top + el.scrollTop) / scale - COL_HEADER_HEIGHT;

      // Find hovered row/col
      let targetRow = norm.end.row;
      for (let r = 0; r < sheet.rowCount; r++) {
        if (y >= rowPositions.tops[r] && y <= rowPositions.tops[r] + rowPositions.heights[r]) {
          targetRow = r;
          break;
        }
      }
      let targetCol = norm.end.col;
      for (let c = 0; c < sheet.colCount; c++) {
        if (x >= colPositions.lefts[c] && x <= colPositions.lefts[c] + colPositions.widths[c]) {
          targetCol = c;
          break;
        }
      }

      // Decide vertical vs horizontal drag based on greater delta outside selection
      const dDown = Math.max(0, targetRow - norm.end.row);
      const dRight = Math.max(0, targetCol - norm.end.col);
      if (dDown >= dRight && targetRow > norm.end.row) {
        latestTarget = { start: norm.start, end: { row: targetRow, col: norm.end.col } };
      } else if (targetCol > norm.end.col) {
        latestTarget = { start: norm.start, end: { row: norm.end.row, col: targetCol } };
      } else {
        latestTarget = norm;
      }
      setFillPreviewRange(latestTarget);
    };

    const onUp = () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      setFillPreviewRange(null);
      if (
        latestTarget.end.row !== norm.end.row ||
        latestTarget.end.col !== norm.end.col
      ) {
        props.onFillDrag(latestTarget);
      }
    };

    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };

  const startColResize = (e: React.MouseEvent, col: number) => {
    e.stopPropagation();
    e.preventDefault();
    const startX = e.clientX;
    const origW = getColWidth(col);
    const onMove = (me: MouseEvent) => {
      const next = Math.max(32, Math.min(500, origW + (me.clientX - startX) / scale));
      props.onSetColWidth(col, Math.round(next));
    };
    const onUp = () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };

  const startRowResize = (e: React.MouseEvent, row: number) => {
    e.stopPropagation();
    e.preventDefault();
    const startY = e.clientY;
    const origH = getRowHeight(row);
    const onMove = (me: MouseEvent) => {
      const next = Math.max(18, Math.min(300, origH + (me.clientY - startY) / scale));
      props.onSetRowHeight(row, Math.round(next));
    };
    const onUp = () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };

  const normSelection = normalizeRange(selection);

  // Selection overlay box coordinates
  const selRect = useMemo(() => {
    const left = colPositions.lefts[normSelection.start.col] ?? 0;
    const top = rowPositions.tops[normSelection.start.row] ?? 0;
    const right =
      (colPositions.lefts[normSelection.end.col] ?? 0) +
      (colPositions.widths[normSelection.end.col] ?? 0);
    const bottom =
      (rowPositions.tops[normSelection.end.row] ?? 0) +
      (rowPositions.heights[normSelection.end.row] ?? 0);
    return { left, top, width: Math.max(0, right - left), height: Math.max(0, bottom - top) };
  }, [normSelection, colPositions, rowPositions]);

  // Evaluate Conditional Formatting for a single cell
  const getCellCFStyle = (row: number, col: number, ev: EvaluatedCell) => {
    let bg: string | undefined;
    let color: string | undefined;
    let dataBarPct: number | null = null;
    let dataBarColor = '#107c41';

    for (const rule of sheet.conditionalRules) {
      const rng = parseRangeRef(rule.range);
      if (!rng || !isCellInRange(row, col, rng)) continue;
      const st = cfStats.get(rule.id);

      if (rule.type === 'dataBar' && typeof ev.value === 'number' && st) {
        const span = Math.max(1, st.max - Math.min(0, st.min));
        dataBarPct = Math.max(0, Math.min(100, ((ev.value - Math.min(0, st.min)) / span) * 100));
        dataBarColor = rule.barColor || '#107c41';
        continue;
      }

      if (rule.type === 'colorScale' && typeof ev.value === 'number' && st && st.max > st.min) {
        const ratio = (ev.value - st.min) / (st.max - st.min);
        if (ratio >= 0.66) bg = '#dcfce7';
        else if (ratio >= 0.33) bg = '#fef9c3';
        else bg = '#fee2e2';
        continue;
      }

      let matched = false;
      const numVal = typeof ev.value === 'number' ? ev.value : Number(ev.value);
      const v1 = Number(rule.value1);
      const v2 = Number(rule.value2);

      switch (rule.type) {
        case 'greaterThan':
          matched = typeof ev.value === 'number' && !isNaN(v1) && numVal > v1;
          break;
        case 'lessThan':
          matched = typeof ev.value === 'number' && !isNaN(v1) && numVal < v1;
          break;
        case 'between':
          matched = typeof ev.value === 'number' && !isNaN(v1) && !isNaN(v2) && numVal >= v1 && numVal <= v2;
          break;
        case 'equal':
          matched = ev.formatted.toLowerCase() === (rule.value1 || '').toLowerCase();
          break;
        case 'textContains':
          matched = Boolean(rule.value1) && ev.formatted.toLowerCase().includes((rule.value1 || '').toLowerCase());
          break;
        case 'duplicate':
          matched = ev.formatted !== '' && (st?.counts.get(ev.formatted) || 0) > 1;
          break;
      }

      if (matched) {
        if (rule.backgroundColor) bg = rule.backgroundColor;
        if (rule.textColor) color = rule.textColor;
      }
    }

    return { bg, color, dataBarPct, dataBarColor };
  };

  // Render inline Sparkline SVG
  const renderSparkline = (cell: CellData, width: number, height: number) => {
    if (!cell.sparkline) return null;
    const rng = parseRangeRef(cell.sparkline.range);
    if (!rng) return null;
    const vals: number[] = [];
    for (let r = rng.start.row; r <= rng.end.row; r++) {
      for (let c = rng.start.col; c <= rng.end.col; c++) {
        const ev = evaluator.evaluateCell(sheet.id, r, c);
        vals.push(typeof ev.value === 'number' ? ev.value : 0);
      }
    }
    if (vals.length === 0) return null;

    const pad = 4;
    const w = Math.max(20, width - pad * 2);
    const h = Math.max(10, height - pad * 2);
    const min = Math.min(...vals);
    const max = Math.max(...vals);
    const span = max - min || 1;
    const col = cell.sparkline.color || '#107c41';

    if (cell.sparkline.type === 'line') {
      const pts = vals
        .map((v, i) => {
          const x = pad + (vals.length === 1 ? w / 2 : (i / (vals.length - 1)) * w);
          const y = pad + h - ((v - min) / span) * h;
          return `${x},${y}`;
        })
        .join(' ');
      return (
        <svg width={width} height={height} className="pointer-events-none">
          <polyline fill="none" stroke={col} strokeWidth={1.6} points={pts} />
        </svg>
      );
    }

    if (cell.sparkline.type === 'column') {
      const bw = Math.max(2, w / vals.length - 2);
      return (
        <svg width={width} height={height} className="pointer-events-none">
          {vals.map((v, i) => {
            const bh = Math.max(2, ((v - Math.min(0, min)) / (max - Math.min(0, min) || 1)) * h);
            return (
              <rect
                key={i}
                x={pad + i * (w / vals.length)}
                y={pad + h - bh}
                width={bw}
                height={bh}
                fill={v < 0 ? '#dc2626' : col}
                rx={0.5}
              />
            );
          })}
        </svg>
      );
    }

    // winloss
    const bw = Math.max(2, w / vals.length - 2);
    return (
      <svg width={width} height={height} className="pointer-events-none">
        <line x1={pad} y1={pad + h / 2} x2={pad + w} y2={pad + h / 2} stroke="#cbd5e1" strokeWidth={0.75} />
        {vals.map((v, i) => (
          <rect
            key={i}
            x={pad + i * (w / vals.length)}
            y={v >= 0 ? pad + 1 : pad + h / 2 + 1}
            width={bw}
            height={Math.max(2, h / 2 - 2)}
            fill={v >= 0 ? col : '#dc2626'}
          />
        ))}
      </svg>
    );
  };

  // List of rows & cols to render (frozen + visible window)
  const rowsToRender = useMemo(() => {
    const arr: number[] = [];
    for (let r = 0; r < sheet.freezePanes.row && r < sheet.rowCount; r++) {
      if (rowPositions.heights[r] > 0) arr.push(r);
    }
    for (let r = visibleRowBounds.start; r <= visibleRowBounds.end; r++) {
      if (r >= sheet.freezePanes.row && rowPositions.heights[r] > 0) arr.push(r);
    }
    return arr;
  }, [sheet.freezePanes.row, sheet.rowCount, visibleRowBounds, rowPositions.heights]);

  const colsToRender = useMemo(() => {
    const arr: number[] = [];
    for (let c = 0; c < sheet.freezePanes.col && c < sheet.colCount; c++) {
      if (colPositions.widths[c] > 0) arr.push(c);
    }
    for (let c = visibleColBounds.start; c <= visibleColBounds.end; c++) {
      if (c >= sheet.freezePanes.col && colPositions.widths[c] > 0) arr.push(c);
    }
    return arr;
  }, [sheet.freezePanes.col, sheet.colCount, visibleColBounds, colPositions.widths]);

  // Unique values for active AutoFilter popover
  const filterUniqueValues = useMemo(() => {
    if (!filterPopover || !autoFilterParsed) return [];
    const set = new Set<string>();
    const startR = autoFilterParsed.start.row + 1;
    for (let r = startR; r < sheet.rowCount; r++) {
      let hasRow = false;
      for (let c = autoFilterParsed.start.col; c <= autoFilterParsed.end.col; c++) {
        if (sheet.cells[cellKey(r, c)]?.raw) {
          hasRow = true;
          break;
        }
      }
      if (!hasRow && r > autoFilterParsed.end.row) continue;
      const ev = evaluator.evaluateCell(sheet.id, r, filterPopover.col);
      set.add(ev.formatted || '(Blanks)');
    }
    return Array.from(set).sort();
  }, [filterPopover, autoFilterParsed, sheet.rowCount, sheet.cells, sheet.id, evaluator]);

  const activeCellData = sheet.cells[cellKey(activeCell.row, activeCell.col)];

  const borderToCss = (side?: { style: string; color: string }) => {
    if (!side) return undefined;
    const width = side.style === 'thick' ? '3px' : side.style === 'medium' ? '2px' : '1px';
    const style = side.style === 'dashed' ? 'dashed' : side.style === 'double' ? 'double' : 'solid';
    return `${width} ${style} ${side.color}`;
  };

  return (
    <div
      ref={containerRef}
      onScroll={handleScroll}
      onContextMenu={(e) => e.preventDefault()}
      className="flex-1 relative overflow-auto excel-scroll bg-slate-100 dark:bg-slate-950 select-none"
    >
      {/* Scaled Canvas Wrapper */}
      <div
        style={{
          width: (ROW_HEADER_WIDTH + colPositions.totalWidth + 80) * scale,
          height: (COL_HEADER_HEIGHT + rowPositions.totalHeight + 60) * scale,
        }}
        className="relative"
      >
        <div
          style={{
            transform: `scale(${scale})`,
            transformOrigin: 'top left',
            width: ROW_HEADER_WIDTH + colPositions.totalWidth,
            height: COL_HEADER_HEIGHT + rowPositions.totalHeight,
          }}
          className="relative bg-white dark:bg-slate-900"
        >
          {/* Corner Select-All Cell */}
          {sheet.pageSetup.showHeadings && (
            <button
              type="button"
              onClick={() =>
                props.onSelectRange(
                  {
                    start: { row: 0, col: 0 },
                    end: { row: sheet.rowCount - 1, col: sheet.colCount - 1 },
                  },
                  { row: 0, col: 0 }
                )
              }
              style={{
                width: ROW_HEADER_WIDTH,
                height: COL_HEADER_HEIGHT,
                position: 'sticky',
                top: 0,
                left: 0,
              }}
              title="Select All"
              className="z-40 bg-slate-100 dark:bg-slate-800 border-r border-b border-slate-300 dark:border-slate-700 flex items-end justify-end p-1 hover:bg-slate-200 dark:hover:bg-slate-700"
            >
              <div className="w-0 h-0 border-l-[8px] border-l-transparent border-b-[8px] border-b-slate-400" />
            </button>
          )}

          {/* Column Headers */}
          {sheet.pageSetup.showHeadings &&
            colsToRender.map((c) => {
              const isFrozenCol = c < sheet.freezePanes.col;
              const isSelectedCol = c >= normSelection.start.col && c <= normSelection.end.col;
              const left = ROW_HEADER_WIDTH + colPositions.lefts[c];
              return (
                <div
                  key={`ch-${c}`}
                  onClick={(e) => {
                    if (e.shiftKey) {
                      props.onSelectRange(
                        {
                          start: { row: 0, col: activeCell.col },
                          end: { row: sheet.rowCount - 1, col: c },
                        },
                        { row: 0, col: activeCell.col }
                      );
                    } else {
                      props.onSelectRange(
                        { start: { row: 0, col: c }, end: { row: sheet.rowCount - 1, col: c } },
                        { row: 0, col: c }
                      );
                    }
                  }}
                  onContextMenu={(e) => {
                    e.preventDefault();
                    setCtxMenu({ x: e.clientX, y: e.clientY, row: 0, col: c, kind: 'col' });
                  }}
                  style={{
                    position: isFrozenCol ? 'sticky' : 'absolute',
                    top: 0,
                    left: isFrozenCol ? left : left,
                    width: colPositions.widths[c],
                    height: COL_HEADER_HEIGHT,
                  }}
                  className={`z-30 flex items-center justify-center text-[11px] font-semibold border-r border-b border-slate-300 dark:border-slate-700 cursor-pointer ${
                    isSelectedCol
                      ? 'bg-emerald-100 dark:bg-emerald-900/50 text-[#107c41] dark:text-emerald-300 border-b-2 border-b-[#107c41]'
                      : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-700'
                  } ${c === sheet.freezePanes.col - 1 ? 'border-r-2 border-r-slate-500' : ''}`}
                >
                  <span>{colIndexToLetter(c)}</span>
                  <div
                    onMouseDown={(e) => startColResize(e, c)}
                    onDoubleClick={(e) => {
                      e.stopPropagation();
                      props.onAutoFitCol(c);
                    }}
                    title="Drag to resize, double-click to AutoFit"
                    className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize hover:bg-[#107c41]"
                  />
                </div>
              );
            })}

          {/* Row Headers */}
          {sheet.pageSetup.showHeadings &&
            rowsToRender.map((r) => {
              const isFrozenRow = r < sheet.freezePanes.row;
              const isSelectedRow = r >= normSelection.start.row && r <= normSelection.end.row;
              const top = COL_HEADER_HEIGHT + rowPositions.tops[r];
              return (
                <div
                  key={`rh-${r}`}
                  onClick={(e) => {
                    if (e.shiftKey) {
                      props.onSelectRange(
                        {
                          start: { row: activeCell.row, col: 0 },
                          end: { row: r, col: sheet.colCount - 1 },
                        },
                        { row: activeCell.row, col: 0 }
                      );
                    } else {
                      props.onSelectRange(
                        { start: { row: r, col: 0 }, end: { row: r, col: sheet.colCount - 1 } },
                        { row: r, col: 0 }
                      );
                    }
                  }}
                  onContextMenu={(e) => {
                    e.preventDefault();
                    setCtxMenu({ x: e.clientX, y: e.clientY, row: r, col: 0, kind: 'row' });
                  }}
                  style={{
                    position: isFrozenRow ? 'sticky' : 'absolute',
                    left: 0,
                    top,
                    width: ROW_HEADER_WIDTH,
                    height: rowPositions.heights[r],
                  }}
                  className={`z-30 flex items-center justify-center text-[11px] font-mono-num font-medium border-r border-b border-slate-300 dark:border-slate-700 cursor-pointer ${
                    isSelectedRow
                      ? 'bg-emerald-100 dark:bg-emerald-900/50 text-[#107c41] dark:text-emerald-300 border-r-2 border-r-[#107c41]'
                      : 'bg-slate-100 dark:bg-slate-800 text-slate-500 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-slate-700'
                  } ${r === sheet.freezePanes.row - 1 ? 'border-b-2 border-b-slate-500' : ''}`}
                >
                  <span>{r + 1}</span>
                  <div
                    onMouseDown={(e) => startRowResize(e, r)}
                    title="Drag to resize row height"
                    className="absolute left-0 right-0 bottom-0 h-1.5 cursor-row-resize hover:bg-[#107c41]"
                  />
                </div>
              );
            })}

          {/* Cells Layer */}
          <div
            style={{
              position: 'absolute',
              left: sheet.pageSetup.showHeadings ? ROW_HEADER_WIDTH : 0,
              top: sheet.pageSetup.showHeadings ? COL_HEADER_HEIGHT : 0,
              width: colPositions.totalWidth,
              height: rowPositions.totalHeight,
            }}
          >
            {rowsToRender.map((r) =>
              colsToRender.map((c) => {
                const key = cellKey(r, c);
                if (mergeMap.slaves.has(key)) return null;

                const mergeInfo = mergeMap.masters.get(key);
                const cellW = mergeInfo ? mergeInfo.width : colPositions.widths[c];
                const cellH = mergeInfo ? mergeInfo.height : rowPositions.heights[r];

                const cell = sheet.cells[key];
                const ev = evaluator.evaluateCell(sheet.id, r, c);
                const isActive = activeCell.row === r && activeCell.col === c;
                const inSel =
                  isCellInRange(r, c, normSelection) ||
                  extraSelections.some((ex) => isCellInRange(r, c, ex));

                const cf = getCellCFStyle(r, c, ev);

                const isNum = typeof ev.value === 'number';
                const align =
                  cell?.style?.horizontalAlign && cell.style.horizontalAlign !== 'general'
                    ? cell.style.horizontalAlign
                    : isNum
                    ? 'right'
                    : typeof ev.value === 'boolean'
                    ? 'center'
                    : 'left';

                const vAlign =
                  cell?.style?.verticalAlign === 'top'
                    ? 'items-start'
                    : cell?.style?.verticalAlign === 'bottom'
                    ? 'items-end'
                    : 'items-center';

                const isFrozenRow = r < sheet.freezePanes.row;
                const isFrozenCol = c < sheet.freezePanes.col;

                // Formula editing highlight border
                const fHighlight = formulaHighlights.find(
                  (h) =>
                    h.sheetId === sheet.id &&
                    r >= h.startRow &&
                    r <= h.endRow &&
                    c >= h.startCol &&
                    c <= h.endCol
                );

                // Audit highlight
                const isPrec = auditHighlights.precedents.some((p) => p.row === r && p.col === c);
                const isDep = auditHighlights.dependents.some((p) => p.row === r && p.col === c);

                // AutoFilter button check
                const hasFilterArrow =
                  sheet.autoFilter &&
                  autoFilterParsed &&
                  r === autoFilterParsed.start.row &&
                  c >= autoFilterParsed.start.col &&
                  c <= autoFilterParsed.end.col;

                const filterActiveOnCol =
                  hasFilterArrow &&
                  (sheet.autoFilter?.hiddenValuesByCol[c]?.length || 0) > 0;

                const displayText = showFormulas && cell?.raw ? cell.raw : ev.formatted;

                return (
                  <div
                    key={key}
                    onMouseDown={(e) => handleCellMouseDown(e, r, c)}
                    onMouseEnter={() => {
                      handleCellMouseEnter(r, c);
                      if (cell?.comment) setHoveredCommentCell({ row: r, col: c });
                    }}
                    onMouseLeave={() => {
                      if (hoveredCommentCell?.row === r && hoveredCommentCell?.col === c) {
                        setHoveredCommentCell(null);
                      }
                    }}
                    onDoubleClick={(e) => {
                      e.stopPropagation();
                      props.onStartEdit();
                    }}
                    onContextMenu={(e) => {
                      e.preventDefault();
                      setCtxMenu({ x: e.clientX, y: e.clientY, row: r, col: c, kind: 'cell' });
                    }}
                    style={{
                      position: isFrozenRow || isFrozenCol ? 'sticky' : 'absolute',
                      left: colPositions.lefts[c],
                      top: rowPositions.tops[r],
                      width: cellW,
                      height: cellH,
                      zIndex: isFrozenRow && isFrozenCol ? 25 : isFrozenRow || isFrozenCol ? 18 : mergeInfo ? 5 : 1,
                      backgroundColor: cf.bg || cell?.style?.backgroundColor || undefined,
                      color: ev.error
                        ? '#dc2626'
                        : cf.color || cell?.style?.color || undefined,
                      fontFamily: cell?.style?.fontFamily || undefined,
                      fontSize: cell?.style?.fontSize ? `${cell.style.fontSize}px` : '12px',
                      fontWeight: cell?.style?.bold ? 700 : 400,
                      fontStyle: cell?.style?.italic ? 'italic' : 'normal',
                      textDecoration: [
                        cell?.style?.underline ? 'underline' : '',
                        cell?.style?.strikethrough ? 'line-through' : '',
                      ]
                        .filter(Boolean)
                        .join(' ') || undefined,
                      borderTop: borderToCss(cell?.style?.borders?.top),
                      borderRight: borderToCss(cell?.style?.borders?.right),
                      borderBottom: borderToCss(cell?.style?.borders?.bottom),
                      borderLeft: borderToCss(cell?.style?.borders?.left),
                      paddingLeft: cell?.style?.indent ? `${6 + cell.style.indent * 10}px` : undefined,
                    }}
                    className={`px-1.5 flex ${vAlign} overflow-hidden text-slate-900 dark:text-slate-100 ${
                      sheet.pageSetup.showGridlines
                        ? 'border-r border-b border-slate-200 dark:border-slate-800'
                        : ''
                    } ${inSel && !isActive ? 'bg-[#107c41]/10 dark:bg-emerald-500/15' : ''} ${
                      isNum ? 'font-mono-num' : ''
                    } ${fHighlight ? 'ring-1 ring-inset' : ''} ${
                      isPrec ? 'ring-2 ring-inset ring-sky-500 bg-sky-50/40' : ''
                    } ${isDep ? 'ring-2 ring-inset ring-indigo-500 bg-indigo-50/40' : ''}`}
                  >
                    {/* Conditional Formatting Data Bar */}
                    {cf.dataBarPct !== null && (
                      <div
                        style={{
                          width: `${cf.dataBarPct}%`,
                          backgroundColor: cf.dataBarColor,
                        }}
                        className="absolute left-0.5 top-1 bottom-1 opacity-20 rounded-xs pointer-events-none"
                      />
                    )}

                    {/* Formula color highlight overlay */}
                    {fHighlight && (
                      <div
                        style={{ borderColor: REF_COLORS[fHighlight.colorIndex % REF_COLORS.length] }}
                        className="absolute inset-0 border-2 border-dashed pointer-events-none"
                      />
                    )}

                    {/* Active Inline Cell Editor */}
                    {isActive && isEditing ? (
                      <input
                        ref={inlineInputRef}
                        type="text"
                        value={editDraft}
                        onChange={(e) => props.onChangeEditDraft(e.target.value)}
                        onMouseDown={(e) => e.stopPropagation()}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') {
                            e.preventDefault();
                            props.onCommitEdit(e.shiftKey ? 'up' : 'down');
                          } else if (e.key === 'Tab') {
                            e.preventDefault();
                            props.onCommitEdit(e.shiftKey ? 'left' : 'right');
                          } else if (e.key === 'Escape') {
                            e.preventDefault();
                            props.onCancelEdit();
                          }
                        }}
                        className="w-full h-full bg-white dark:bg-slate-950 text-slate-900 dark:text-white outline-none font-mono-num text-xs z-30"
                      />
                    ) : cell?.isCheckbox ? (
                      <div className="w-full flex items-center justify-center">
                        <input
                          type="checkbox"
                          checked={ev.value === true || String(cell.raw).toUpperCase() === 'TRUE'}
                          onChange={() => props.onToggleCheckboxCell(r, c)}
                          onMouseDown={(e) => e.stopPropagation()}
                          className="w-3.5 h-3.5 accent-[#107c41] cursor-pointer"
                        />
                      </div>
                    ) : cell?.sparkline ? (
                      renderSparkline(cell, cellW - 8, cellH - 2)
                    ) : (
                      <div
                        style={{
                          textAlign: align,
                        }}
                        className={`w-full relative z-10 ${
                          cell?.style?.wrapText ? 'whitespace-normal break-words leading-tight' : 'whitespace-nowrap truncate'
                        }`}
                      >
                        {cell?.hyperlink ? (
                          <a
                            href={cell.hyperlink}
                            target="_blank"
                            rel="noreferrer"
                            onClick={(e) => e.stopPropagation()}
                            className="text-sky-600 dark:text-sky-400 underline hover:text-sky-700"
                          >
                            {displayText || cell.hyperlink}
                          </a>
                        ) : (
                          displayText
                        )}
                      </div>
                    )}

                    {/* Data Validation List Dropdown Button */}
                    {isActive && !isEditing && cell?.validation?.type === 'list' && cell.validation.showDropdown !== false && (
                      <button
                        type="button"
                        onMouseDown={(e) => {
                          e.stopPropagation();
                          setValidationDropdownCell(
                            validationDropdownCell?.row === r && validationDropdownCell?.col === c
                              ? null
                              : { row: r, col: c }
                          );
                        }}
                        className="ml-auto shrink-0 w-4 h-4 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 rounded-xs flex items-center justify-center text-slate-700 dark:text-slate-200 z-20"
                      >
                        <ChevronDown className="w-3 h-3" />
                      </button>
                    )}

                    {/* AutoFilter Dropdown Button */}
                    {hasFilterArrow && (
                      <button
                        type="button"
                        onMouseDown={(e) => {
                          e.stopPropagation();
                          setFilterSearch('');
                          setFilterPopover({ col: c, x: e.clientX, y: e.clientY });
                        }}
                        className={`ml-auto shrink-0 w-4 h-4 rounded-xs flex items-center justify-center z-20 ${
                          filterActiveOnCol
                            ? 'bg-[#107c41] text-white'
                            : 'bg-white/85 dark:bg-slate-800/85 text-slate-700 dark:text-slate-200 hover:bg-white'
                        }`}
                        title="Filter column"
                      >
                        <Filter className="w-2.5 h-2.5" />
                      </button>
                    )}

                    {/* Comment Corner Triangle */}
                    {cell?.comment && (
                      <div
                        title={cell.comment}
                        className="absolute top-0 right-0 w-0 h-0 border-t-[6px] border-t-red-600 border-l-[6px] border-l-transparent pointer-events-none z-20"
                      />
                    )}

                    {/* Error Corner Triangle */}
                    {ev.error && (
                      <div
                        title={`Formula Error: ${ev.error}`}
                        className="absolute top-0 left-0 w-0 h-0 border-t-[6px] border-t-emerald-600 border-r-[6px] border-r-transparent pointer-events-none z-20"
                      />
                    )}

                    {/* Comment Popover (hover or Show All) */}
                    {cell?.comment &&
                      (showAllComments ||
                        (hoveredCommentCell?.row === r && hoveredCommentCell?.col === c)) && (
                        <div className="absolute left-full top-0 ml-1 w-48 p-2 bg-amber-50 dark:bg-slate-800 border border-amber-300 dark:border-slate-600 rounded shadow-lg text-[11px] text-slate-800 dark:text-slate-100 z-50 pointer-events-none whitespace-normal">
                          <div className="font-semibold text-[10px] text-amber-700 dark:text-amber-400 mb-0.5">
                            Note · {colIndexToLetter(c)}
                            {r + 1}
                          </div>
                          {cell.comment}
                        </div>
                      )}
                  </div>
                );
              })
            )}

            {/* Primary Selection Border & Fill Handle */}
            <div
              style={{
                left: selRect.left,
                top: selRect.top,
                width: selRect.width,
                height: selRect.height,
              }}
              className="absolute border-2 border-[#107c41] pointer-events-none z-20"
            >
              {!isEditing && (
                <div
                  onMouseDown={startFillHandleDrag}
                  title="Drag to auto-fill series or formulas"
                  className="absolute -right-[4px] -bottom-[4px] w-[8px] h-[8px] bg-[#107c41] border border-white cursor-crosshair pointer-events-auto"
                />
              )}
            </div>

            {/* Fill Handle Drag Preview */}
            {fillPreviewRange && (
              <div
                style={{
                  left: colPositions.lefts[fillPreviewRange.start.col] ?? 0,
                  top: rowPositions.tops[fillPreviewRange.start.row] ?? 0,
                  width:
                    (colPositions.lefts[fillPreviewRange.end.col] ?? 0) +
                    (colPositions.widths[fillPreviewRange.end.col] ?? 0) -
                    (colPositions.lefts[fillPreviewRange.start.col] ?? 0),
                  height:
                    (rowPositions.tops[fillPreviewRange.end.row] ?? 0) +
                    (rowPositions.heights[fillPreviewRange.end.row] ?? 0) -
                    (rowPositions.tops[fillPreviewRange.start.row] ?? 0),
                }}
                className="absolute border-2 border-dashed border-[#107c41] bg-[#107c41]/5 pointer-events-none z-25"
              />
            )}

            {/* Clipboard Marching Ants */}
            {clipboard && clipboard.sheetId === sheet.id && (
              <svg
                style={{
                  left: colPositions.lefts[normalizeRange(clipboard.range).start.col] ?? 0,
                  top: rowPositions.tops[normalizeRange(clipboard.range).start.row] ?? 0,
                  width:
                    (colPositions.lefts[normalizeRange(clipboard.range).end.col] ?? 0) +
                    (colPositions.widths[normalizeRange(clipboard.range).end.col] ?? 0) -
                    (colPositions.lefts[normalizeRange(clipboard.range).start.col] ?? 0),
                  height:
                    (rowPositions.tops[normalizeRange(clipboard.range).end.row] ?? 0) +
                    (rowPositions.heights[normalizeRange(clipboard.range).end.row] ?? 0) -
                    (rowPositions.tops[normalizeRange(clipboard.range).start.row] ?? 0),
                }}
                className="absolute pointer-events-none z-25 overflow-visible"
              >
                <rect x="0" y="0" width="100%" height="100%" className="marching-ants-rect" />
              </svg>
            )}

            {/* Page Break Dashed Guides when Page Break View is active */}
            {pageBreakView && (
              <>
                <div
                  style={{ left: 720, top: 0, bottom: 0 }}
                  className="absolute w-0 border-l-2 border-dashed border-sky-500/70 pointer-events-none z-20"
                />
                <div
                  style={{ top: 960, left: 0, right: 0 }}
                  className="absolute h-0 border-t-2 border-dashed border-sky-500/70 pointer-events-none z-20"
                />
              </>
            )}

            {/* Data Validation List Dropdown Popover */}
            {validationDropdownCell &&
              (() => {
                const vCell = sheet.cells[cellKey(validationDropdownCell.row, validationDropdownCell.col)];
                if (!vCell?.validation || vCell.validation.type !== 'list') return null;
                const items = vCell.validation.formula1
                  .split(',')
                  .map((s) => s.trim())
                  .filter(Boolean);
                return (
                  <div
                    onMouseDown={(e) => e.stopPropagation()}
                    style={{
                      left: colPositions.lefts[validationDropdownCell.col],
                      top:
                        rowPositions.tops[validationDropdownCell.row] +
                        rowPositions.heights[validationDropdownCell.row],
                      minWidth: Math.max(120, colPositions.widths[validationDropdownCell.col]),
                    }}
                    className="absolute z-40 bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-600 rounded shadow-lg py-1 text-xs"
                  >
                    {items.map((item) => (
                      <button
                        key={item}
                        onClick={() => {
                          props.onSetCellDirect(validationDropdownCell.row, validationDropdownCell.col, item);
                          setValidationDropdownCell(null);
                        }}
                        className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 dark:hover:bg-slate-700 text-slate-800 dark:text-slate-100"
                      >
                        {item}
                      </button>
                    ))}
                  </div>
                );
              })()}

            {/* Floating Charts */}
            {sheet.charts.map((chart) => (
              <FloatingChart
                key={chart.id}
                chart={chart}
                sheet={sheet}
                sheets={sheets}
                namedRanges={namedRanges}
                currencySymbol={currencySymbol}
                selected={selectedChartId === chart.id}
                onSelect={() => props.onSelectChart(chart.id)}
                onUpdate={(patch) => props.onUpdateChart(chart.id, patch)}
                onDelete={() => props.onDeleteChart(chart.id)}
              />
            ))}
          </div>
        </div>
      </div>

      {/* Add More Rows Button at bottom */}
      <div className="no-print px-12 py-2 flex items-center gap-3 text-xs text-slate-500">
        <button
          type="button"
          onClick={props.onAddMoreRows}
          className="px-3 py-1 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 font-medium"
        >
          + Add 50 Rows (Currently {sheet.rowCount} rows)
        </button>
      </div>

      {/* Data Validation Input Prompt Banner when active cell has inputMessage */}
      {activeCellData?.validation?.inputMessage && (
        <div className="fixed bottom-16 right-6 z-30 max-w-xs bg-amber-50 dark:bg-slate-800 border border-amber-300 dark:border-slate-600 rounded-md shadow-md px-3 py-2 text-xs">
          {activeCellData.validation.inputTitle && (
            <div className="font-semibold text-slate-900 dark:text-white mb-0.5">
              {activeCellData.validation.inputTitle}
            </div>
          )}
          <div className="text-slate-600 dark:text-slate-300">{activeCellData.validation.inputMessage}</div>
        </div>
      )}

      {/* AutoFilter Column Popover */}
      {filterPopover && (
        <div
          onClick={(e) => e.stopPropagation()}
          style={{ left: Math.min(window.innerWidth - 250, filterPopover.x), top: filterPopover.y + 8 }}
          className="fixed z-50 w-60 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-md shadow-xl p-2.5 text-xs space-y-2"
        >
          <div className="flex flex-col gap-1 border-b border-slate-100 dark:border-slate-700 pb-2">
            <button
              onClick={() => {
                props.onSortCol(filterPopover.col, true);
                setFilterPopover(null);
              }}
              className="w-full px-2 py-1 text-left rounded hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"
            >
              <ArrowDownAZ className="w-3.5 h-3.5 text-[#107c41]" /> Sort A to Z
            </button>
            <button
              onClick={() => {
                props.onSortCol(filterPopover.col, false);
                setFilterPopover(null);
              }}
              className="w-full px-2 py-1 text-left rounded hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"
            >
              <ArrowUpZA className="w-3.5 h-3.5 text-sky-600" /> Sort Z to A
            </button>
          </div>

          <input
            type="text"
            placeholder="Search values..."
            value={filterSearch}
            onChange={(e) => setFilterSearch(e.target.value)}
            className="w-full px-2 py-1 border border-slate-300 dark:border-slate-600 rounded bg-transparent text-slate-800 dark:text-slate-100"
          />

          <div className="max-h-40 overflow-y-auto space-y-1 border border-slate-100 dark:border-slate-700 rounded p-1.5">
            {(() => {
              const hiddenSet = new Set(sheet.autoFilter?.hiddenValuesByCol[filterPopover.col] || []);
              const filteredList = filterUniqueValues.filter((v) =>
                v.toLowerCase().includes(filterSearch.toLowerCase())
              );
              return (
                <>
                  <label className="flex items-center gap-2 font-semibold cursor-pointer pb-1 border-b border-slate-100 dark:border-slate-700">
                    <input
                      type="checkbox"
                      checked={hiddenSet.size === 0}
                      onChange={(e) => {
                        props.onUpdateAutoFilterCol(
                          filterPopover.col,
                          e.target.checked ? [] : [...filterUniqueValues]
                        );
                      }}
                    />
                    <span>(Select All)</span>
                  </label>
                  {filteredList.map((val) => {
                    const checked = !hiddenSet.has(val);
                    return (
                      <label key={val} className="flex items-center gap-2 cursor-pointer py-0.5">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={(e) => {
                            const next = new Set(hiddenSet);
                            if (e.target.checked) next.delete(val);
                            else next.add(val);
                            props.onUpdateAutoFilterCol(filterPopover.col, Array.from(next));
                          }}
                        />
                        <span className="truncate">{val}</span>
                      </label>
                    );
                  })}
                </>
              );
            })()}
          </div>
          <div className="flex justify-end">
            <button
              onClick={() => setFilterPopover(null)}
              className="px-3 py-1 bg-[#107c41] text-white rounded font-medium hover:bg-[#0c5e31]"
            >
              Done
            </button>
          </div>
        </div>
      )}

      {/* Right-Click Context Menu */}
      {ctxMenu && (
        <div
          onClick={(e) => e.stopPropagation()}
          style={{
            left: Math.min(window.innerWidth - 210, ctxMenu.x),
            top: Math.min(window.innerHeight - 340, ctxMenu.y),
          }}
          className="fixed z-50 w-52 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-md shadow-xl py-1 text-xs text-slate-700 dark:text-slate-200"
        >
          <button onClick={() => { props.onCut(); setCtxMenu(null); }} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><Scissors className="w-3.5 h-3.5" /> Cut</button>
          <button onClick={() => { props.onCopy(); setCtxMenu(null); }} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><Copy className="w-3.5 h-3.5" /> Copy</button>
          <button onClick={() => { props.onPaste(); setCtxMenu(null); }} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><ClipboardPaste className="w-3.5 h-3.5" /> Paste</button>
          <button onClick={() => { props.onPasteValues(); setCtxMenu(null); }} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><span className="w-3.5 text-center font-mono-num font-bold">123</span> Paste Values Only</button>
          <div className="border-t border-slate-100 dark:border-slate-700 my-1" />
          <button onClick={() => { props.onInsertRowAt(ctxMenu.row); setCtxMenu(null); }} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><Plus className="w-3.5 h-3.5" /> Insert Row Above</button>
          <button onClick={() => { props.onInsertColAt(ctxMenu.col); setCtxMenu(null); }} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><Plus className="w-3.5 h-3.5" /> Insert Column Left</button>
          <button onClick={() => { props.onDeleteRowAt(ctxMenu.row); setCtxMenu(null); }} className="w-full px-3 py-1.5 text-left hover:bg-red-50 dark:hover:bg-red-900/30 text-red-600 flex items-center gap-2"><Trash2 className="w-3.5 h-3.5" /> Delete Row</button>
          <button onClick={() => { props.onDeleteColAt(ctxMenu.col); setCtxMenu(null); }} className="w-full px-3 py-1.5 text-left hover:bg-red-50 dark:hover:bg-red-900/30 text-red-600 flex items-center gap-2"><Trash2 className="w-3.5 h-3.5" /> Delete Column</button>
          <div className="border-t border-slate-100 dark:border-slate-700 my-1" />
          <button onClick={() => { props.onClearSelection(); setCtxMenu(null); }} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><Eraser className="w-3.5 h-3.5" /> Clear Contents</button>
          <button onClick={() => { props.onMergeCenter(); setCtxMenu(null); }} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><Merge className="w-3.5 h-3.5" /> Merge & Center</button>
          <button onClick={() => { props.onSortCol(ctxMenu.col, true); setCtxMenu(null); }} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><ArrowDownAZ className="w-3.5 h-3.5" /> Sort A to Z</button>
          <div className="border-t border-slate-100 dark:border-slate-700 my-1" />
          <button onClick={() => { props.onOpenFormatCells(); setCtxMenu(null); }} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><SlidersHorizontal className="w-3.5 h-3.5" /> Format Cells...</button>
          <button onClick={() => { props.onOpenValidation(); setCtxMenu(null); }} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><ShieldCheck className="w-3.5 h-3.5" /> Data Validation...</button>
          <button onClick={() => { props.onOpenComment(); setCtxMenu(null); }} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><MessageSquarePlus className="w-3.5 h-3.5" /> Insert/Edit Comment</button>
          <button onClick={() => { props.onDefineName(); setCtxMenu(null); }} className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"><Tag className="w-3.5 h-3.5" /> Define Name...</button>
        </div>
      )}
    </div>
  );
};
