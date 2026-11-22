import React, { useMemo, useRef, useState } from 'react';
import { BarChart3, GripHorizontal, Settings2, Trash2, X } from 'lucide-react';
import { ChartType, NamedRange, Sheet, SheetChart } from '../types/spreadsheet';
import { WorkbookEvaluator } from '../engine/formulaEngine';
import { parseRangeRef } from '../utils/cellCoords';

interface FloatingChartProps {
  chart: SheetChart;
  sheet: Sheet;
  sheets: Sheet[];
  namedRanges: NamedRange[];
  currencySymbol: string;
  selected: boolean;
  onSelect: () => void;
  onUpdate: (patch: Partial<SheetChart>) => void;
  onDelete: () => void;
}

const PALETTES: Record<string, string[]> = {
  excel: ['#107c41', '#0284c7', '#d97706', '#7c3aed', '#dc2626', '#0d9488'],
  ocean: ['#0369a1', '#0ea5e9', '#14b8a6', '#6366f1', '#1e293b', '#38bdf8'],
  warm: ['#dc2626', '#ea580c', '#d97706', '#eab308', '#be123c', '#f97316'],
  monochrome: ['#0f172a', '#334155', '#64748b', '#94a3b8', '#cbd5e1', '#475569'],
};

export const FloatingChart: React.FC<FloatingChartProps> = ({
  chart,
  sheet,
  sheets,
  namedRanges,
  currencySymbol,
  selected,
  onSelect,
  onUpdate,
  onDelete,
}) => {
  const [editing, setEditing] = useState(false);
  const dragRef = useRef<{ startX: number; startY: number; origX: number; origY: number } | null>(null);
  const resizeRef = useRef<{ startX: number; startY: number; origW: number; origH: number } | null>(null);

  const colors = PALETTES[chart.palette || 'excel'];

  // Extract chart data matrix from sheet range
  const parsedData = useMemo(() => {
    const rng = parseRangeRef(chart.dataRange);
    if (!rng) return { categories: [] as string[], series: [] as { name: string; values: number[] }[] };
    const ev = new WorkbookEvaluator(sheets, namedRanges, currencySymbol);

    const rawMatrix: (string | number)[][] = [];
    for (let r = rng.start.row; r <= rng.end.row; r++) {
      const row: (string | number)[] = [];
      for (let c = rng.start.col; c <= rng.end.col; c++) {
        const res = ev.evaluateCell(sheet.id, r, c);
        if (typeof res.value === 'number') row.push(res.value);
        else row.push(res.formatted || '');
      }
      rawMatrix.push(row);
    }

    if (rawMatrix.length === 0) return { categories: [], series: [] };

    const matrix = chart.switchRowCol
      ? rawMatrix[0].map((_, cIdx) => rawMatrix.map((row) => row[cIdx]))
      : rawMatrix;

    const hasHeaderRow = matrix.length > 1 && matrix[0].slice(1).some((v) => typeof v === 'string' && v !== '');
    const startRow = hasHeaderRow ? 1 : 0;
    const colCount = matrix[0]?.length || 0;

    const categories: string[] = [];
    const series: { name: string; values: number[] }[] = [];

    for (let c = 1; c < colCount; c++) {
      const name = hasHeaderRow ? String(matrix[0][c] || `Series ${c}`) : `Series ${c}`;
      series.push({ name, values: [] });
    }
    if (series.length === 0 && colCount === 1) {
      series.push({ name: hasHeaderRow ? String(matrix[0][0] || 'Value') : 'Value', values: [] });
    }

    for (let r = startRow; r < matrix.length; r++) {
      const cat = String(matrix[r][0] ?? `Row ${r}`);
      categories.push(cat);
      if (colCount === 1) {
        const v = matrix[r][0];
        series[0].values.push(typeof v === 'number' ? v : Number(v) || 0);
      } else {
        for (let c = 1; c < colCount; c++) {
          const v = matrix[r][c];
          series[c - 1].values.push(typeof v === 'number' ? v : Number(v) || 0);
        }
      }
    }

    return { categories, series };
  }, [chart.dataRange, chart.switchRowCol, sheet, sheets, namedRanges, currencySymbol]);

  const startDrag = (e: React.MouseEvent) => {
    e.stopPropagation();
    onSelect();
    dragRef.current = { startX: e.clientX, startY: e.clientY, origX: chart.x, origY: chart.y };
    const onMove = (me: MouseEvent) => {
      if (!dragRef.current) return;
      onUpdate({
        x: Math.max(0, dragRef.current.origX + (me.clientX - dragRef.current.startX)),
        y: Math.max(0, dragRef.current.origY + (me.clientY - dragRef.current.startY)),
      });
    };
    const onUp = () => {
      dragRef.current = null;
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };

  const startResize = (e: React.MouseEvent) => {
    e.stopPropagation();
    onSelect();
    resizeRef.current = { startX: e.clientX, startY: e.clientY, origW: chart.width, origH: chart.height };
    const onMove = (me: MouseEvent) => {
      if (!resizeRef.current) return;
      onUpdate({
        width: Math.max(280, resizeRef.current.origW + (me.clientX - resizeRef.current.startX)),
        height: Math.max(200, resizeRef.current.origH + (me.clientY - resizeRef.current.startY)),
      });
    };
    const onUp = () => {
      resizeRef.current = null;
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };

  const renderChartSVG = () => {
    const svgW = Math.max(240, chart.width - 24);
    const svgH = Math.max(120, chart.height - (chart.showLegend ? 96 : 70));
    const { categories, series } = parsedData;

    if (categories.length === 0 || series.length === 0) {
      return (
        <div className="flex flex-col items-center justify-center h-full text-xs text-slate-400">
          <BarChart3 className="w-6 h-6 mb-1 opacity-50" />
          <span>No numeric data in range {chart.dataRange}</span>
        </div>
      );
    }

    if (chart.type === 'pie' || chart.type === 'donut') {
      const vals = series[0].values.map((v) => Math.max(0, v));
      const total = vals.reduce((a, b) => a + b, 0) || 1;
      const cx = svgW / 2;
      const cy = svgH / 2;
      const r = Math.min(cx, cy) - 12;
      const innerR = chart.type === 'donut' ? r * 0.52 : 0;
      let startAngle = -Math.PI / 2;

      return (
        <svg width={svgW} height={svgH} className="overflow-visible">
          {vals.map((val, idx) => {
            const sliceAngle = (val / total) * Math.PI * 2;
            const endAngle = startAngle + sliceAngle;
            const largeArc = sliceAngle > Math.PI ? 1 : 0;
            const x1 = cx + r * Math.cos(startAngle);
            const y1 = cy + r * Math.sin(startAngle);
            const x2 = cx + r * Math.cos(endAngle);
            const y2 = cy + r * Math.sin(endAngle);

            let d = '';
            if (innerR > 0) {
              const ix1 = cx + innerR * Math.cos(endAngle);
              const iy1 = cy + innerR * Math.sin(endAngle);
              const ix2 = cx + innerR * Math.cos(startAngle);
              const iy2 = cy + innerR * Math.sin(startAngle);
              d = `M ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} L ${ix1} ${iy1} A ${innerR} ${innerR} 0 ${largeArc} 0 ${ix2} ${iy2} Z`;
            } else {
              d = `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} Z`;
            }
            startAngle = endAngle;
            const pct = ((val / total) * 100).toFixed(1);
            return (
              <path
                key={idx}
                d={d}
                fill={colors[idx % colors.length]}
                stroke="#ffffff"
                strokeWidth={1.5}
              >
                <title>{`${categories[idx]}: ${val.toLocaleString()} (${pct}%)`}</title>
              </path>
            );
          })}
        </svg>
      );
    }

    const padLeft = 52;
    const padRight = 14;
    const padTop = 10;
    const padBottom = 26;
    const plotW = Math.max(40, svgW - padLeft - padRight);
    const plotH = Math.max(40, svgH - padTop - padBottom);

    const allVals = series.flatMap((s) => s.values);
    const maxVal = Math.max(0, ...allVals, 1);
    const minVal = Math.min(0, ...allVals);
    const span = maxVal - minVal || 1;

    const yForVal = (v: number) => padTop + plotH - ((v - minVal) / span) * plotH;

    if (chart.type === 'bar') {
      // Horizontal Bar
      const xForVal = (v: number) => padLeft + ((v - minVal) / span) * plotW;
      const groupH = plotH / categories.length;
      const barH = Math.max(4, (groupH * 0.72) / series.length);
      const zeroX = xForVal(0);

      return (
        <svg width={svgW} height={svgH}>
          {chart.showGridlines &&
            [0, 0.25, 0.5, 0.75, 1].map((t, i) => {
              const x = padLeft + t * plotW;
              const labelVal = minVal + t * span;
              return (
                <g key={i}>
                  <line x1={x} y1={padTop} x2={x} y2={padTop + plotH} stroke="currentColor" className="text-slate-200 dark:text-slate-700" strokeDasharray="3 3" />
                  <text x={x} y={padTop + plotH + 14} textAnchor="middle" className="fill-slate-500 text-[9px] font-mono-num">
                    {labelVal >= 1000 ? `${(labelVal / 1000).toFixed(1)}k` : Math.round(labelVal)}
                  </text>
                </g>
              );
            })}
          {categories.map((cat, cIdx) => {
            const groupY = padTop + cIdx * groupH;
            return (
              <g key={cIdx}>
                <text
                  x={padLeft - 6}
                  y={groupY + groupH / 2 + 3}
                  textAnchor="end"
                  className="fill-slate-600 dark:fill-slate-300 text-[9px]"
                >
                  {cat.length > 9 ? `${cat.slice(0, 8)}…` : cat}
                </text>
                {series.map((s, sIdx) => {
                  const val = s.values[cIdx] ?? 0;
                  const bx = val >= 0 ? zeroX : xForVal(val);
                  const bw = Math.max(1, Math.abs(xForVal(val) - zeroX));
                  const by = groupY + groupH * 0.14 + sIdx * barH;
                  return (
                    <rect
                      key={sIdx}
                      x={bx}
                      y={by}
                      width={bw}
                      height={Math.max(2, barH - 1)}
                      rx={1.5}
                      fill={colors[sIdx % colors.length]}
                    >
                      <title>{`${cat} · ${s.name}: ${val.toLocaleString()}`}</title>
                    </rect>
                  );
                })}
              </g>
            );
          })}
        </svg>
      );
    }

    // Column, Line, Area, Scatter
    const zeroY = yForVal(0);
    const groupW = plotW / categories.length;

    return (
      <svg width={svgW} height={svgH}>
        {chart.showGridlines &&
          [0, 0.25, 0.5, 0.75, 1].map((t, i) => {
            const y = padTop + (1 - t) * plotH;
            const labelVal = minVal + t * span;
            return (
              <g key={i}>
                <line x1={padLeft} y1={y} x2={padLeft + plotW} y2={y} stroke="currentColor" className="text-slate-200 dark:text-slate-700" strokeDasharray="3 3" />
                <text x={padLeft - 6} y={y + 3} textAnchor="end" className="fill-slate-500 text-[9px] font-mono-num">
                  {Math.abs(labelVal) >= 1000 ? `${(labelVal / 1000).toFixed(1)}k` : Math.round(labelVal)}
                </text>
              </g>
            );
          })}

        {/* X Axis labels */}
        {categories.map((cat, idx) => (
          <text
            key={idx}
            x={padLeft + idx * groupW + groupW / 2}
            y={padTop + plotH + 15}
            textAnchor="middle"
            className="fill-slate-600 dark:fill-slate-300 text-[9px]"
          >
            {cat.length > 10 ? `${cat.slice(0, 9)}…` : cat}
          </text>
        ))}

        {/* Column Bars */}
        {chart.type === 'column' &&
          categories.map((cat, cIdx) => {
            const barW = Math.max(4, (groupW * 0.72) / series.length);
            const gx = padLeft + cIdx * groupW + groupW * 0.14;
            return series.map((s, sIdx) => {
              const val = s.values[cIdx] ?? 0;
              const y = val >= 0 ? yForVal(val) : zeroY;
              const h = Math.max(1, Math.abs(yForVal(val) - zeroY));
              return (
                <rect
                  key={`${cIdx}-${sIdx}`}
                  x={gx + sIdx * barW}
                  y={y}
                  width={Math.max(2, barW - 1.5)}
                  height={h}
                  rx={1.5}
                  fill={colors[sIdx % colors.length]}
                >
                  <title>{`${cat} · ${s.name}: ${val.toLocaleString()}`}</title>
                </rect>
              );
            });
          })}

        {/* Line & Area */}
        {(chart.type === 'line' || chart.type === 'area') &&
          series.map((s, sIdx) => {
            const pts = s.values.map((v, cIdx) => ({
              x: padLeft + cIdx * groupW + groupW / 2,
              y: yForVal(v),
              v,
            }));
            const linePath = pts.map((p, i) => `${i === 0 ? 'M' : 'L'} ${p.x} ${p.y}`).join(' ');
            const areaPath =
              pts.length > 0
                ? `${linePath} L ${pts[pts.length - 1].x} ${zeroY} L ${pts[0].x} ${zeroY} Z`
                : '';
            const col = colors[sIdx % colors.length];
            return (
              <g key={sIdx}>
                {chart.type === 'area' && <path d={areaPath} fill={col} fillOpacity={0.18} />}
                <path d={linePath} fill="none" stroke={col} strokeWidth={2} />
                {pts.map((p, i) => (
                  <circle key={i} cx={p.x} cy={p.y} r={3} fill={col} stroke="#ffffff" strokeWidth={1}>
                    <title>{`${categories[i]} · ${s.name}: ${p.v.toLocaleString()}`}</title>
                  </circle>
                ))}
              </g>
            );
          })}

        {/* Scatter */}
        {chart.type === 'scatter' &&
          series.map((s, sIdx) => {
            const col = colors[sIdx % colors.length];
            return s.values.map((v, cIdx) => {
              const cx = padLeft + cIdx * groupW + groupW / 2;
              const cy = yForVal(v);
              return (
                <circle key={`${sIdx}-${cIdx}`} cx={cx} cy={cy} r={4.5} fill={col} fillOpacity={0.85}>
                  <title>{`${categories[cIdx]} · ${s.name}: ${v.toLocaleString()}`}</title>
                </circle>
              );
            });
          })}
      </svg>
    );
  };

  return (
    <div
      style={{
        left: chart.x,
        top: chart.y,
        width: chart.width,
        height: chart.height,
      }}
      onMouseDown={(e) => {
        e.stopPropagation();
        onSelect();
      }}
      className={`absolute z-20 bg-white dark:bg-slate-900 rounded-md shadow-md border flex flex-col select-none ${
        selected
          ? 'border-[#107c41] ring-2 ring-[#107c41]/30'
          : 'border-slate-300 dark:border-slate-700'
      }`}
    >
      {/* Header drag bar */}
      <div
        onMouseDown={startDrag}
        className="flex items-center justify-between px-3 py-1.5 border-b border-slate-100 dark:border-slate-800 cursor-move bg-slate-50/70 dark:bg-slate-800/50 rounded-t-md"
      >
        <div className="flex items-center gap-1.5 min-w-0">
          <GripHorizontal className="w-3.5 h-3.5 text-slate-400 shrink-0" />
          <span className="font-display font-semibold text-xs text-slate-800 dark:text-slate-100 truncate">
            {chart.title}
          </span>
        </div>
        <div className="flex items-center gap-1 shrink-0">
          <button
            type="button"
            onMouseDown={(e) => e.stopPropagation()}
            onClick={() => setEditing((v) => !v)}
            title="Chart Design"
            className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-600 dark:text-slate-300"
          >
            <Settings2 className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onMouseDown={(e) => e.stopPropagation()}
            onClick={onDelete}
            title="Delete Chart"
            className="p-1 rounded hover:bg-red-100 dark:hover:bg-red-900/40 text-red-600"
          >
            <Trash2 className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* Chart Canvas */}
      <div className="flex-1 p-3 flex flex-col items-center justify-center overflow-hidden">
        {renderChartSVG()}

        {chart.showLegend && (
          <div className="flex flex-wrap items-center justify-center gap-3 mt-1 text-[10px] text-slate-600 dark:text-slate-300">
            {(chart.type === 'pie' || chart.type === 'donut'
              ? parsedData.categories
              : parsedData.series.map((s) => s.name)
            ).map((label, idx) => (
              <div key={idx} className="flex items-center gap-1">
                <span
                  className="w-2.5 h-2.5 rounded-xs inline-block"
                  style={{ backgroundColor: colors[idx % colors.length] }}
                />
                <span className="truncate max-w-[90px]">{label}</span>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Chart Editor Popover */}
      {editing && (
        <div
          onMouseDown={(e) => e.stopPropagation()}
          className="absolute right-2 top-9 w-64 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-md shadow-xl p-3 z-30 text-xs space-y-2.5"
        >
          <div className="flex items-center justify-between border-b border-slate-100 dark:border-slate-700 pb-1.5">
            <span className="font-semibold text-slate-800 dark:text-slate-100">Chart Design</span>
            <button onClick={() => setEditing(false)} className="text-slate-400 hover:text-slate-700">
              <X className="w-3.5 h-3.5" />
            </button>
          </div>
          <div>
            <label className="block text-[11px] text-slate-500 mb-0.5">Chart Title</label>
            <input
              type="text"
              value={chart.title}
              onChange={(e) => onUpdate({ title: e.target.value })}
              className="w-full px-2 py-1 border border-slate-300 dark:border-slate-600 rounded bg-transparent text-slate-800 dark:text-slate-100"
            />
          </div>
          <div>
            <label className="block text-[11px] text-slate-500 mb-0.5">Data Range</label>
            <input
              type="text"
              value={chart.dataRange}
              onChange={(e) => onUpdate({ dataRange: e.target.value.toUpperCase() })}
              className="w-full px-2 py-1 font-mono-num border border-slate-300 dark:border-slate-600 rounded bg-transparent text-slate-800 dark:text-slate-100"
            />
          </div>
          <div className="grid grid-cols-2 gap-2">
            <div>
              <label className="block text-[11px] text-slate-500 mb-0.5">Type</label>
              <select
                value={chart.type}
                onChange={(e) => onUpdate({ type: e.target.value as ChartType })}
                className="w-full px-1.5 py-1 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-900 text-slate-800 dark:text-slate-100"
              >
                <option value="column">Column</option>
                <option value="bar">Bar</option>
                <option value="line">Line</option>
                <option value="area">Area</option>
                <option value="pie">Pie</option>
                <option value="donut">Donut</option>
                <option value="scatter">Scatter</option>
              </select>
            </div>
            <div>
              <label className="block text-[11px] text-slate-500 mb-0.5">Palette</label>
              <select
                value={chart.palette || 'excel'}
                onChange={(e) => onUpdate({ palette: e.target.value as SheetChart['palette'] })}
                className="w-full px-1.5 py-1 border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-900 text-slate-800 dark:text-slate-100"
              >
                <option value="excel">Excel Classic</option>
                <option value="ocean">Ocean Blue</option>
                <option value="warm">Warm Sunset</option>
                <option value="monochrome">Monochrome</option>
              </select>
            </div>
          </div>
          <div className="flex flex-col gap-1.5 pt-1">
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={chart.showLegend}
                onChange={(e) => onUpdate({ showLegend: e.target.checked })}
              />
              <span>Show Legend</span>
            </label>
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={chart.showGridlines}
                onChange={(e) => onUpdate({ showGridlines: e.target.checked })}
              />
              <span>Show Gridlines</span>
            </label>
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={Boolean(chart.switchRowCol)}
                onChange={(e) => onUpdate({ switchRowCol: e.target.checked })}
              />
              <span>Switch Row / Column</span>
            </label>
          </div>
        </div>
      )}

      {/* Resize handle */}
      <div
        onMouseDown={startResize}
        title="Drag to resize chart"
        className="absolute right-0 bottom-0 w-3.5 h-3.5 cursor-nwse-resize flex items-end justify-end p-0.5"
      >
        <div className="w-2 h-2 border-r-2 border-b-2 border-slate-400" />
      </div>
    </div>
  );
};
