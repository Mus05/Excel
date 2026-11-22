import React, { useState } from 'react';
import {
  Download,
  Home,
  Moon,
  Printer,
  Redo2,
  Save,
  Search,
  Sun,
  Table2,
  Undo2,
} from 'lucide-react';

interface TitleBarProps {
  title: string;
  onRenameTitle: (newTitle: string) => void;
  canUndo: boolean;
  canRedo: boolean;
  onUndo: () => void;
  onRedo: () => void;
  onSaveLocal: () => void;
  onExportXLSX: () => void;
  onPrint: () => void;
  onOpenSearch: () => void;
  theme: 'light' | 'dark';
  onToggleTheme: () => void;
  saveFlash: boolean;
  onGoHome: () => void;
}

export const TitleBar: React.FC<TitleBarProps> = ({
  title,
  onRenameTitle,
  canUndo,
  canRedo,
  onUndo,
  onRedo,
  onSaveLocal,
  onExportXLSX,
  onPrint,
  onOpenSearch,
  theme,
  onToggleTheme,
  saveFlash,
  onGoHome,
}) => {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(title);

  const commitTitle = () => {
    setEditing(false);
    const trimmed = draft.trim();
    if (trimmed) onRenameTitle(trimmed);
    else setDraft(title);
  };

  return (
    <header className="no-print h-10 bg-[#107c41] dark:bg-[#0b532c] text-white flex items-center justify-between px-3 select-none shrink-0 border-b border-black/15">
      {/* Left: Brand & Quick Access Toolbar */}
      <div className="flex items-center gap-2 min-w-0">
        <button
          type="button"
          onClick={onGoHome}
          title="Back to Home Dashboard"
          className="flex items-center gap-1.5 px-2 py-1 rounded hover:bg-white/15 transition-colors border-r border-white/20 pr-2.5 shrink-0"
        >
          <Home className="w-3.5 h-3.5 text-white" />
          <Table2 className="w-4 h-4 text-white hidden sm:block" />
          <span className="font-display font-semibold text-sm tracking-tight whitespace-nowrap">
            ExcelGrid
          </span>
        </button>

        <div className="flex items-center gap-0.5 shrink-0">
          <button
            onClick={onSaveLocal}
            title="Save Workbook (Ctrl+S)"
            className="p-1.5 rounded hover:bg-white/15 active:bg-white/25 transition-colors flex items-center gap-1 text-xs"
          >
            <Save className="w-3.5 h-3.5" />
            {saveFlash && <span className="text-[11px] font-medium">Saved</span>}
          </button>
          <button
            onClick={onUndo}
            disabled={!canUndo}
            title="Undo (Ctrl+Z)"
            className="p-1.5 rounded hover:bg-white/15 disabled:opacity-40 disabled:pointer-events-none transition-colors"
          >
            <Undo2 className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={onRedo}
            disabled={!canRedo}
            title="Redo (Ctrl+Y)"
            className="p-1.5 rounded hover:bg-white/15 disabled:opacity-40 disabled:pointer-events-none transition-colors"
          >
            <Redo2 className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={onPrint}
            title="Print (Ctrl+P)"
            className="p-1.5 rounded hover:bg-white/15 transition-colors"
          >
            <Printer className="w-3.5 h-3.5" />
          </button>
        </div>

        <div className="h-4 w-px bg-white/20 mx-1 hidden sm:block" />

        {editing ? (
          <input
            autoFocus
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            onBlur={commitTitle}
            onKeyDown={(e) => {
              if (e.key === 'Enter') commitTitle();
              if (e.key === 'Escape') {
                setDraft(title);
                setEditing(false);
              }
            }}
            className="bg-white/20 text-white placeholder-white/70 text-xs font-medium px-2 py-1 rounded outline-none border border-white/40 w-48 sm:w-64"
          />
        ) : (
          <button
            onClick={() => {
              setDraft(title);
              setEditing(true);
            }}
            title="Click to rename workbook"
            className="text-xs font-medium px-2 py-1 rounded hover:bg-white/15 truncate max-w-[180px] sm:max-w-[260px] text-left"
          >
            {title}
          </button>
        )}
      </div>

      {/* Center: Tell me what you want to do / Search */}
      <button
        onClick={onOpenSearch}
        className="hidden md:flex items-center gap-2 bg-white/15 hover:bg-white/25 transition-colors text-white/90 text-xs px-3 py-1 rounded w-72 lg:w-96 border border-white/15"
      >
        <Search className="w-3.5 h-3.5 opacity-85 shrink-0" />
        <span className="truncate">Search commands, functions, or Go To (Ctrl+K)</span>
      </button>

      {/* Right: Theme Toggle & Export */}
      <div className="flex items-center gap-1.5 shrink-0">
        <button
          onClick={onToggleTheme}
          title={theme === 'dark' ? 'Switch to Light Mode' : 'Switch to Dark Mode'}
          className="p-1.5 rounded hover:bg-white/15 transition-colors"
        >
          {theme === 'dark' ? <Sun className="w-3.5 h-3.5" /> : <Moon className="w-3.5 h-3.5" />}
        </button>
        <button
          onClick={onExportXLSX}
          title="Download as .xlsx"
          className="flex items-center gap-1.5 bg-white text-[#107c41] hover:bg-emerald-50 font-semibold text-xs px-2.5 py-1 rounded shadow-xs transition-colors whitespace-nowrap shrink-0"
        >
          <Download className="w-3.5 h-3.5" />
          <span className="hidden sm:inline">Export .xlsx</span>
        </button>
      </div>
    </header>
  );
};
