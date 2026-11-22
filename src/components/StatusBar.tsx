import React from 'react';
import { Minus, Plus } from 'lucide-react';
import { ExcelMode } from '../types/spreadsheet';

export interface SelectionSummary {
  count: number;
  numericCount: number;
  sum: number;
  avg: number;
  min: number;
  max: number;
}

interface StatusBarProps {
  mode: ExcelMode;
  sheetProtected: boolean;
  calcMode: 'automatic' | 'manual';
  summary: SelectionSummary;
  zoom: number;
  onChangeZoom: (zoom: number) => void;
  onCalculateNow: () => void;
  showPageLayoutHint: boolean;
  onTogglePageLayoutView: () => void;
}

export const StatusBar: React.FC<StatusBarProps> = ({
  mode,
  sheetProtected,
  calcMode,
  summary,
  zoom,
  onChangeZoom,
  onCalculateNow,
  showPageLayoutHint,
  onTogglePageLayoutView,
}) => {
  const copyNum = (n: number) => {
    navigator.clipboard?.writeText(String(Number(n.toFixed(6))));
  };

  return (
    <footer className="no-print h-6 bg-[#107c41] dark:bg-[#0b532c] text-white text-[11px] flex items-center justify-between px-3 select-none shrink-0">
      {/* Left: Mode & Protection */}
      <div className="flex items-center gap-3 shrink-0">
        <span className="font-medium tracking-wide">{mode}</span>
        {sheetProtected && (
          <span className="text-amber-200 font-medium">Protected Sheet</span>
        )}
        {calcMode === 'manual' && (
          <button
            onClick={onCalculateNow}
            className="bg-white/20 hover:bg-white/30 px-1.5 py-0.5 rounded text-[10px] font-medium"
          >
            Calculate (F9)
          </button>
        )}
      </div>

      {/* Right: Live Quick Aggregations + View + Zoom */}
      <div className="flex items-center gap-4 min-w-0">
        {summary.count > 1 && (
          <div className="hidden md:flex items-center gap-3 font-mono-num text-[11px] text-white/95">
            {summary.numericCount > 0 && (
              <>
                <button onClick={() => copyNum(summary.avg)} title="Click to copy Average" className="hover:underline">
                  Average: {summary.avg.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                </button>
                <span aria-hidden="true" className="opacity-50">·</span>
              </>
            )}
            <span>Count: {summary.count}</span>
            {summary.numericCount > 0 && (
              <>
                <span aria-hidden="true" className="opacity-50">·</span>
                <button onClick={() => copyNum(summary.min)} title="Click to copy Min" className="hover:underline">
                  Min: {summary.min.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                </button>
                <span aria-hidden="true" className="opacity-50">·</span>
                <button onClick={() => copyNum(summary.max)} title="Click to copy Max" className="hover:underline">
                  Max: {summary.max.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                </button>
                <span aria-hidden="true" className="opacity-50">·</span>
                <button onClick={() => copyNum(summary.sum)} title="Click to copy Sum" className="hover:underline font-semibold">
                  Sum: {summary.sum.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                </button>
              </>
            )}
          </div>
        )}

        <div className="flex items-center gap-2 shrink-0 border-l border-white/20 pl-3">
          <button
            onClick={onTogglePageLayoutView}
            className={`px-1.5 py-0.5 rounded text-[10px] ${showPageLayoutHint ? 'bg-white/25 font-semibold' : 'hover:bg-white/15'}`}
            title="Toggle Page Break View"
          >
            {showPageLayoutHint ? 'Page View' : 'Normal'}
          </button>
          <button
            onClick={() => onChangeZoom(Math.max(50, zoom - 10))}
            className="p-0.5 hover:bg-white/15 rounded"
            title="Zoom Out"
          >
            <Minus className="w-3 h-3" />
          </button>
          <input
            type="range"
            min={50}
            max={160}
            step={5}
            value={zoom}
            onChange={(e) => onChangeZoom(Number(e.target.value))}
            className="w-16 accent-white cursor-pointer"
          />
          <button
            onClick={() => onChangeZoom(Math.min(160, zoom + 10))}
            className="p-0.5 hover:bg-white/15 rounded"
            title="Zoom In"
          >
            <Plus className="w-3 h-3" />
          </button>
          <button
            onClick={() => onChangeZoom(100)}
            title="Reset Zoom to 100%"
            className="font-mono-num w-9 text-right hover:underline"
          >
            {zoom}%
          </button>
        </div>
      </div>
    </footer>
  );
};
