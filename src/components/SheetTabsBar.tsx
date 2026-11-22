import React, { useEffect, useRef, useState } from 'react';
import { Copy, Palette, Plus, Trash2, ChevronLeft, ChevronRight, Edit3 } from 'lucide-react';
import { Sheet } from '../types/spreadsheet';

interface SheetTabsBarProps {
  sheets: Sheet[];
  activeSheetId: string;
  onSelectSheet: (id: string) => void;
  onAddSheet: () => void;
  onRenameSheet: (id: string, name: string) => void;
  onDuplicateSheet: (id: string) => void;
  onDeleteSheet: (id: string) => void;
  onSetTabColor: (id: string, color?: string) => void;
  onMoveSheet: (id: string, dir: -1 | 1) => void;
}

const TAB_COLORS = [
  undefined,
  '#107c41',
  '#0284c7',
  '#d97706',
  '#dc2626',
  '#7c3aed',
  '#0d9488',
  '#334155',
];

export const SheetTabsBar: React.FC<SheetTabsBarProps> = ({
  sheets,
  activeSheetId,
  onSelectSheet,
  onAddSheet,
  onRenameSheet,
  onDuplicateSheet,
  onDeleteSheet,
  onSetTabColor,
  onMoveSheet,
}) => {
  const [renamingId, setRenamingId] = useState<string | null>(null);
  const [draftName, setDraftName] = useState('');
  const [ctxMenu, setCtxMenu] = useState<{ sheetId: string; x: number; y: number } | null>(null);
  const menuRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const close = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        setCtxMenu(null);
      }
    };
    document.addEventListener('mousedown', close);
    return () => document.removeEventListener('mousedown', close);
  }, []);

  const commitRename = () => {
    if (renamingId && draftName.trim()) {
      onRenameSheet(renamingId, draftName.trim());
    }
    setRenamingId(null);
  };

  return (
    <div className="no-print h-8 bg-slate-100 dark:bg-slate-800 border-t border-slate-200 dark:border-slate-700 flex items-center px-2 gap-1 select-none shrink-0 overflow-x-auto">
      <div className="flex items-center gap-0.5 shrink-0">
        {sheets.map((s) => {
          const active = s.id === activeSheetId;
          const isRenaming = renamingId === s.id;
          return (
            <div
              key={s.id}
              onClick={() => onSelectSheet(s.id)}
              onDoubleClick={() => {
                setRenamingId(s.id);
                setDraftName(s.name);
              }}
              onContextMenu={(e) => {
                e.preventDefault();
                setCtxMenu({ sheetId: s.id, x: e.clientX, y: e.clientY });
              }}
              style={{
                borderBottomColor: s.tabColor || (active ? '#107c41' : 'transparent'),
              }}
              className={`group relative px-3 h-7 flex items-center gap-1.5 text-xs cursor-pointer rounded-t border-t border-x transition-colors whitespace-nowrap shrink-0 ${
                active
                  ? 'bg-white dark:bg-slate-900 border-slate-300 dark:border-slate-700 font-semibold text-[#107c41] dark:text-emerald-400 border-b-2'
                  : 'bg-slate-200/60 dark:bg-slate-800 border-transparent text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-700 border-b-2'
              }`}
            >
              {isRenaming ? (
                <input
                  autoFocus
                  value={draftName}
                  onChange={(e) => setDraftName(e.target.value)}
                  onBlur={commitRename}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') commitRename();
                    if (e.key === 'Escape') setRenamingId(null);
                  }}
                  onClick={(e) => e.stopPropagation()}
                  className="w-24 px-1 py-0.5 text-xs bg-white dark:bg-slate-950 border border-[#107c41] rounded outline-none text-slate-900 dark:text-slate-100"
                />
              ) : (
                <span>{s.name}</span>
              )}
              {s.protected && (
                <span className="text-[10px] text-amber-600 dark:text-amber-400" title="Protected Sheet">
                  Lock
                </span>
              )}
            </div>
          );
        })}
      </div>

      <button
        type="button"
        onClick={onAddSheet}
        title="New Sheet"
        className="w-6 h-6 rounded-full hover:bg-slate-200 dark:hover:bg-slate-700 flex items-center justify-center text-slate-600 dark:text-slate-300 shrink-0 ml-1"
      >
        <Plus className="w-3.5 h-3.5" />
      </button>

      {ctxMenu && (
        <div
          ref={menuRef}
          style={{ left: ctxMenu.x, top: Math.max(8, ctxMenu.y - 220) }}
          className="fixed z-50 w-48 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-md shadow-xl py-1 text-xs text-slate-700 dark:text-slate-200"
        >
          <button
            onClick={() => {
              const target = sheets.find((s) => s.id === ctxMenu.sheetId);
              if (target) {
                setRenamingId(target.id);
                setDraftName(target.name);
              }
              setCtxMenu(null);
            }}
            className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"
          >
            <Edit3 className="w-3.5 h-3.5" /> Rename
          </button>
          <button
            onClick={() => {
              onDuplicateSheet(ctxMenu.sheetId);
              setCtxMenu(null);
            }}
            className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"
          >
            <Copy className="w-3.5 h-3.5" /> Duplicate Sheet
          </button>
          <button
            onClick={() => {
              onMoveSheet(ctxMenu.sheetId, -1);
              setCtxMenu(null);
            }}
            className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"
          >
            <ChevronLeft className="w-3.5 h-3.5" /> Move Left
          </button>
          <button
            onClick={() => {
              onMoveSheet(ctxMenu.sheetId, 1);
              setCtxMenu(null);
            }}
            className="w-full px-3 py-1.5 text-left hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center gap-2"
          >
            <ChevronRight className="w-3.5 h-3.5" /> Move Right
          </button>
          <div className="px-3 py-1.5 border-y border-slate-100 dark:border-slate-700 my-1">
            <div className="flex items-center gap-1.5 text-[11px] text-slate-500 mb-1">
              <Palette className="w-3 h-3" /> Tab Color
            </div>
            <div className="flex items-center gap-1.5">
              {TAB_COLORS.map((col, idx) => (
                <button
                  key={idx}
                  onClick={() => {
                    onSetTabColor(ctxMenu.sheetId, col);
                    setCtxMenu(null);
                  }}
                  style={{ backgroundColor: col || '#e2e8f0' }}
                  className="w-4 h-4 rounded-xs border border-slate-300 dark:border-slate-600"
                  title={col || 'Default'}
                />
              ))}
            </div>
          </div>
          <button
            disabled={sheets.length <= 1}
            onClick={() => {
              onDeleteSheet(ctxMenu.sheetId);
              setCtxMenu(null);
            }}
            className="w-full px-3 py-1.5 text-left hover:bg-red-50 dark:hover:bg-red-900/30 text-red-600 disabled:opacity-40 flex items-center gap-2"
          >
            <Trash2 className="w-3.5 h-3.5" /> Delete Sheet
          </button>
        </div>
      )}
    </div>
  );
};
