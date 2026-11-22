import React, { useMemo, useState } from 'react';
import {
  Copy,
  Download,
  FilePlus2,
  FileSpreadsheet,
  FolderOpen,
  Grid3x3,
  Home,
  LayoutTemplate,
  Moon,
  Pin,
  Plus,
  Search,
  Sun,
  Table2,
  Trash2,
} from 'lucide-react';
import { Workbook } from '../types/spreadsheet';
import { WORKBOOK_TEMPLATES } from '../utils/templates';
import {
  formatRelativeTime,
  getWorkbookStats,
  LibraryWorkbookItem,
} from '../utils/workbookLibrary';

interface HomeDashboardProps {
  library: LibraryWorkbookItem[];
  theme: 'light' | 'dark';
  onToggleTheme: () => void;
  onOpenWorkbook: (workbookId: string) => void;
  onCreateFromTemplate: (wb: Workbook) => void;
  onImportFile: () => void;
  onTogglePin: (workbookId: string) => void;
  onDuplicateWorkbook: (workbookId: string) => void;
  onExportXLSX: (wb: Workbook) => void;
  onDeleteWorkbook: (workbookId: string) => void;
}

type SidebarSection = 'home' | 'templates';
type WorkbookFilter = 'all' | 'pinned';

export const HomeDashboard: React.FC<HomeDashboardProps> = ({
  library,
  theme,
  onToggleTheme,
  onOpenWorkbook,
  onCreateFromTemplate,
  onImportFile,
  onTogglePin,
  onDuplicateWorkbook,
  onExportXLSX,
  onDeleteWorkbook,
}) => {
  const [section, setSection] = useState<SidebarSection>('home');
  const [filter, setFilter] = useState<WorkbookFilter>('all');
  const [searchQuery, setSearchQuery] = useState('');

  const totals = useMemo(() => {
    let sheets = 0;
    let formulas = 0;
    let cells = 0;
    for (const item of library) {
      const st = getWorkbookStats(item.workbook);
      sheets += st.sheets;
      formulas += st.formulas;
      cells += st.cells;
    }
    return { workbooks: library.length, sheets, formulas, cells };
  }, [library]);

  const filteredLibrary = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return [...library]
      .filter((item) => {
        if (filter === 'pinned' && !item.pinned) return false;
        if (!q) return true;
        const titleMatch = item.workbook.title.toLowerCase().includes(q);
        const sheetMatch = item.workbook.sheets.some((s) => s.name.toLowerCase().includes(q));
        return titleMatch || sheetMatch;
      })
      .sort((a, b) => {
        if (a.pinned !== b.pinned) return a.pinned ? -1 : 1;
        return b.updatedAt - a.updatedAt;
      });
  }, [library, filter, searchQuery]);

  const filteredTemplates = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    if (!q) return WORKBOOK_TEMPLATES;
    return WORKBOOK_TEMPLATES.filter(
      (t) => t.name.toLowerCase().includes(q) || t.description.toLowerCase().includes(q)
    );
  }, [searchQuery]);

  const handleNewBlank = () => {
    const blankTpl = WORKBOOK_TEMPLATES.find((t) => t.id === 'blank') || WORKBOOK_TEMPLATES[0];
    const wb = blankTpl.build();
    const num = library.length + 1;
    wb.title = `Book${num}.xlsx`;
    onCreateFromTemplate(wb);
  };

  const renderTemplateThumbnail = (tplId: string) => {
    if (tplId === 'blank') {
      return (
        <div className="h-24 w-full bg-white dark:bg-slate-950 border-b border-slate-200 dark:border-slate-800 p-2.5 flex flex-col justify-between">
          <div className="grid grid-cols-4 gap-px bg-slate-200 dark:bg-slate-800 border border-slate-200 dark:border-slate-800 flex-1">
            {Array.from({ length: 12 }).map((_, i) => (
              <div key={i} className="bg-white dark:bg-slate-950" />
            ))}
          </div>
        </div>
      );
    }

    const accent =
      tplId === 'sales-dashboard'
        ? '#107c41'
        : tplId === 'monthly-budget'
        ? '#0284c7'
        : tplId === 'invoice'
        ? '#0f766e'
        : '#4f46e5';

    return (
      <div className="h-24 w-full bg-slate-50 dark:bg-slate-950 border-b border-slate-200 dark:border-slate-800 p-2.5 flex flex-col gap-1.5">
        <div className="h-2.5 w-full rounded-xs" style={{ backgroundColor: accent }} />
        <div className="grid grid-cols-3 gap-1.5 flex-1">
          <div className="col-span-2 flex flex-col gap-1">
            <div className="h-2 w-full bg-slate-200 dark:bg-slate-800 rounded-xs" />
            <div className="h-2 w-4/5 bg-slate-200 dark:bg-slate-800 rounded-xs" />
            <div className="h-2 w-11/12 bg-slate-200 dark:bg-slate-800 rounded-xs" />
          </div>
          <div className="flex items-end gap-0.5 pt-1">
            <div className="flex-1 h-3/5 rounded-xs opacity-70" style={{ backgroundColor: accent }} />
            <div className="flex-1 h-full rounded-xs" style={{ backgroundColor: accent }} />
            <div className="flex-1 h-4/5 rounded-xs opacity-85" style={{ backgroundColor: accent }} />
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="h-screen w-screen flex bg-slate-50 dark:bg-slate-950 text-slate-900 dark:text-slate-100 overflow-hidden">
      {/* Left Navigation Sidebar */}
      <aside className="w-60 bg-[#0e6b38] dark:bg-[#093e21] text-white flex flex-col shrink-0 select-none border-r border-black/10 hidden md:flex">
        <div className="h-14 px-5 flex items-center gap-2.5 border-b border-white/15">
          <Table2 className="w-5 h-5 text-white shrink-0" />
          <span className="font-display font-bold text-base tracking-tight whitespace-nowrap">
            ExcelGrid
          </span>
        </div>

        <nav className="p-3 space-y-1 flex-1">
          <button
            type="button"
            onClick={() => setSection('home')}
            className={`w-full px-3 py-2 rounded-md text-xs font-medium flex items-center gap-2.5 transition-colors whitespace-nowrap ${
              section === 'home'
                ? 'bg-white text-[#107c41] font-semibold shadow-xs'
                : 'text-white/90 hover:bg-white/15'
            }`}
          >
            <Home className="w-4 h-4 shrink-0" />
            <span>Home</span>
          </button>

          <button
            type="button"
            onClick={handleNewBlank}
            className="w-full px-3 py-2 rounded-md text-xs font-medium text-white/90 hover:bg-white/15 flex items-center gap-2.5 transition-colors whitespace-nowrap"
          >
            <FilePlus2 className="w-4 h-4 shrink-0" />
            <span>New Blank Workbook</span>
          </button>

          <button
            type="button"
            onClick={() => setSection('templates')}
            className={`w-full px-3 py-2 rounded-md text-xs font-medium flex items-center gap-2.5 transition-colors whitespace-nowrap ${
              section === 'templates'
                ? 'bg-white text-[#107c41] font-semibold shadow-xs'
                : 'text-white/90 hover:bg-white/15'
            }`}
          >
            <LayoutTemplate className="w-4 h-4 shrink-0" />
            <span>Templates</span>
          </button>

          <button
            type="button"
            onClick={onImportFile}
            className="w-full px-3 py-2 rounded-md text-xs font-medium text-white/90 hover:bg-white/15 flex items-center gap-2.5 transition-colors whitespace-nowrap"
          >
            <FolderOpen className="w-4 h-4 shrink-0" />
            <span>Open .xlsx / .csv</span>
          </button>
        </nav>

        <div className="p-3 border-t border-white/15 space-y-2">
          <button
            type="button"
            onClick={onToggleTheme}
            className="w-full px-3 py-2 rounded-md text-xs font-medium text-white/90 hover:bg-white/15 flex items-center justify-between transition-colors"
          >
            <span className="flex items-center gap-2.5">
              {theme === 'dark' ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
              <span>{theme === 'dark' ? 'Light Appearance' : 'Dark Appearance'}</span>
            </span>
          </button>
          <div className="px-3 py-1.5 text-[11px] text-white/75 flex items-center justify-between font-mono-num">
            <span>Saved Workbooks</span>
            <span>{library.length}</span>
          </div>
        </div>
      </aside>

      {/* Main Content Area */}
      <div className="flex-1 flex flex-col min-w-0 overflow-hidden">
        {/* Top Bar */}
        <header className="h-14 bg-white dark:bg-slate-900 border-b border-slate-200 dark:border-slate-800 px-5 flex items-center justify-between gap-6 shrink-0">
          <div className="flex items-center gap-3 min-w-0">
            <div className="flex items-center gap-2 md:hidden text-[#107c41]">
              <Table2 className="w-5 h-5" />
              <span className="font-display font-bold text-sm">ExcelGrid</span>
            </div>
            <div className="hidden md:flex items-center gap-2 text-xs text-slate-500 dark:text-slate-400 whitespace-nowrap">
              <span className="font-medium text-slate-900 dark:text-white">Workspace</span>
              <span aria-hidden="true">/</span>
              <span>{section === 'home' ? 'Home Dashboard' : 'Template Gallery'}</span>
            </div>
          </div>

          <div className="flex-1 max-w-md relative">
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Search workbooks, sheets, or templates..."
              className="w-full h-8 pl-8 pr-3 text-xs bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-md text-slate-900 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:border-[#107c41]"
            />
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={onImportFile}
              className="hidden sm:flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium border border-slate-200 dark:border-slate-700 rounded-md hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 whitespace-nowrap shrink-0"
            >
              <FolderOpen className="w-3.5 h-3.5" />
              <span>Open File</span>
            </button>
            <button
              type="button"
              onClick={handleNewBlank}
              className="flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-semibold bg-[#107c41] hover:bg-[#0c5e31] text-white rounded-md transition-colors whitespace-nowrap shrink-0"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Blank Workbook</span>
            </button>
          </div>
        </header>

        {/* Scrollable Dashboard Body */}
        <main className="flex-1 overflow-y-auto excel-scroll p-6 lg:p-8">
          <div className="max-w-6xl mx-auto space-y-8">
            {/* Greeting + Summary Strip */}
            <div className="flex flex-col md:flex-row md:items-end justify-between gap-4 pb-6 border-b border-slate-200 dark:border-slate-800">
              <div>
                <h1 className="font-display font-bold text-2xl text-slate-900 dark:text-white tracking-tight">
                  {section === 'home' ? 'Spreadsheet Dashboard' : 'Workbook Templates'}
                </h1>
                <p className="text-sm text-slate-500 dark:text-slate-400 mt-1">
                  Start from a blank grid, launch a financial model, or open a saved workbook.
                </p>
              </div>

              <div className="flex items-center gap-6 text-xs">
                <div>
                  <div className="text-slate-500 dark:text-slate-400">Workbooks</div>
                  <div className="font-mono-num font-semibold text-lg text-slate-900 dark:text-white">
                    {totals.workbooks}
                  </div>
                </div>
                <div className="h-7 w-px bg-slate-200 dark:bg-slate-800" />
                <div>
                  <div className="text-slate-500 dark:text-slate-400">Worksheets</div>
                  <div className="font-mono-num font-semibold text-lg text-slate-900 dark:text-white">
                    {totals.sheets}
                  </div>
                </div>
                <div className="h-7 w-px bg-slate-200 dark:bg-slate-800" />
                <div>
                  <div className="text-slate-500 dark:text-slate-400">Active Formulas</div>
                  <div className="font-mono-num font-semibold text-lg text-[#107c41] dark:text-emerald-400">
                    {totals.formulas}
                  </div>
                </div>
              </div>
            </div>

            {/* Templates Grid */}
            <section className="space-y-3">
              <div className="flex items-center justify-between">
                <h2 className="font-display font-semibold text-sm text-slate-900 dark:text-white">
                  Create New
                </h2>
                {section === 'home' && (
                  <button
                    type="button"
                    onClick={() => setSection('templates')}
                    className="text-xs font-medium text-[#107c41] dark:text-emerald-400 hover:underline whitespace-nowrap"
                  >
                    View all templates
                  </button>
                )}
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3.5">
                {filteredTemplates.map((tpl) => (
                  <button
                    key={tpl.id}
                    type="button"
                    onClick={() => onCreateFromTemplate(tpl.build())}
                    className="group text-left bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-lg overflow-hidden hover:border-[#107c41] dark:hover:border-emerald-500 transition-colors flex flex-col"
                  >
                    {renderTemplateThumbnail(tpl.id)}
                    <div className="p-3 flex-1 flex flex-col justify-between gap-1">
                      <div className="font-display font-semibold text-xs text-slate-900 dark:text-white group-hover:text-[#107c41] dark:group-hover:text-emerald-400">
                        {tpl.name}
                      </div>
                      <p className="text-[11px] text-slate-500 dark:text-slate-400 line-clamp-2 leading-relaxed">
                        {tpl.description}
                      </p>
                    </div>
                  </button>
                ))}
              </div>
            </section>

            {/* Recent & Saved Workbooks Table */}
            {section === 'home' && (
              <section className="space-y-3">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <div className="flex items-center gap-1 p-1 bg-slate-200/70 dark:bg-slate-800 rounded-lg">
                    <button
                      type="button"
                      onClick={() => setFilter('all')}
                      className={`px-3 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap ${
                        filter === 'all'
                          ? 'bg-white dark:bg-slate-900 text-slate-900 dark:text-white shadow-xs'
                          : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white'
                      }`}
                    >
                      Recent Workbooks
                    </button>
                    <button
                      type="button"
                      onClick={() => setFilter('pinned')}
                      className={`px-3 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap ${
                        filter === 'pinned'
                          ? 'bg-white dark:bg-slate-900 text-slate-900 dark:text-white shadow-xs'
                          : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white'
                      }`}
                    >
                      Pinned
                    </button>
                  </div>

                  <div className="text-xs text-slate-500 dark:text-slate-400 font-mono-num">
                    {filteredLibrary.length} workbook{filteredLibrary.length === 1 ? '' : 's'}
                  </div>
                </div>

                {filteredLibrary.length === 0 ? (
                  <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-lg p-10 text-center space-y-3">
                    <Grid3x3 className="w-8 h-8 text-slate-300 dark:text-slate-700 mx-auto" />
                    <div className="text-sm font-medium text-slate-800 dark:text-slate-200">
                      {searchQuery
                        ? `No workbooks matching "${searchQuery}"`
                        : filter === 'pinned'
                        ? 'No pinned workbooks yet'
                        : 'No saved workbooks'}
                    </div>
                    <p className="text-xs text-slate-500 dark:text-slate-400 max-w-sm mx-auto">
                      Create a new blank workbook or import an existing Excel file to start working.
                    </p>
                    <div className="pt-1">
                      <button
                        type="button"
                        onClick={handleNewBlank}
                        className="px-4 py-2 text-xs font-semibold bg-[#107c41] hover:bg-[#0c5e31] text-white rounded-md transition-colors"
                      >
                        Create Blank Workbook
                      </button>
                    </div>
                  </div>
                ) : (
                  <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-lg overflow-hidden">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="border-b border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-800/50 text-[11px] font-semibold text-slate-500 dark:text-slate-400">
                          <th className="py-2.5 px-4">Workbook Name</th>
                          <th className="py-2.5 px-3 text-right hidden sm:table-cell">Sheets</th>
                          <th className="py-2.5 px-3 text-right hidden md:table-cell">Cells</th>
                          <th className="py-2.5 px-3 text-right hidden md:table-cell">Formulas</th>
                          <th className="py-2.5 px-3 text-right">Modified</th>
                          <th className="py-2.5 px-4 text-right">Actions</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 dark:divide-slate-800 text-xs">
                        {filteredLibrary.map((item) => {
                          const st = getWorkbookStats(item.workbook);
                          return (
                            <tr
                              key={item.workbook.id}
                              onClick={() => onOpenWorkbook(item.workbook.id)}
                              className="group hover:bg-slate-50 dark:hover:bg-slate-800/60 cursor-pointer transition-colors"
                            >
                              <td className="py-3 px-4">
                                <div className="flex items-center gap-3 min-w-0">
                                  <FileSpreadsheet className="w-4 h-4 text-[#107c41] dark:text-emerald-400 shrink-0" />
                                  <div className="min-w-0">
                                    <div className="font-medium text-slate-900 dark:text-white group-hover:text-[#107c41] dark:group-hover:text-emerald-400 truncate">
                                      {item.workbook.title}
                                    </div>
                                    <div className="text-[11px] text-slate-400 truncate">
                                      {item.workbook.sheets.map((s) => s.name).join(' · ')}
                                    </div>
                                  </div>
                                </div>
                              </td>
                              <td className="py-3 px-3 text-right font-mono-num text-slate-600 dark:text-slate-300 hidden sm:table-cell">
                                {st.sheets}
                              </td>
                              <td className="py-3 px-3 text-right font-mono-num text-slate-600 dark:text-slate-300 hidden md:table-cell">
                                {st.cells}
                              </td>
                              <td className="py-3 px-3 text-right font-mono-num text-slate-600 dark:text-slate-300 hidden md:table-cell">
                                {st.formulas}
                              </td>
                              <td className="py-3 px-3 text-right font-mono-num text-slate-500 dark:text-slate-400 whitespace-nowrap">
                                {formatRelativeTime(item.updatedAt)}
                              </td>
                              <td
                                className="py-3 px-4 text-right"
                                onClick={(e) => e.stopPropagation()}
                              >
                                <div className="inline-flex items-center gap-1">
                                  <button
                                    type="button"
                                    onClick={() => onTogglePin(item.workbook.id)}
                                    title={item.pinned ? 'Unpin workbook' : 'Pin workbook'}
                                    className={`p-1.5 rounded hover:bg-slate-200 dark:hover:bg-slate-700 transition-colors ${
                                      item.pinned
                                        ? 'text-[#107c41] dark:text-emerald-400'
                                        : 'text-slate-400 hover:text-slate-700 dark:hover:text-slate-200'
                                    }`}
                                  >
                                    <Pin className="w-3.5 h-3.5" />
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => onDuplicateWorkbook(item.workbook.id)}
                                    title="Duplicate workbook"
                                    className="p-1.5 rounded hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-500 hover:text-slate-800 dark:text-slate-400 dark:hover:text-slate-100 transition-colors"
                                  >
                                    <Copy className="w-3.5 h-3.5" />
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => onExportXLSX(item.workbook)}
                                    title="Download .xlsx"
                                    className="p-1.5 rounded hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-500 hover:text-slate-800 dark:text-slate-400 dark:hover:text-slate-100 transition-colors"
                                  >
                                    <Download className="w-3.5 h-3.5" />
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => onDeleteWorkbook(item.workbook.id)}
                                    title="Delete workbook"
                                    className="p-1.5 rounded hover:bg-red-100 dark:hover:bg-red-900/40 text-slate-400 hover:text-red-600 transition-colors"
                                  >
                                    <Trash2 className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </section>
            )}
          </div>
        </main>
      </div>
    </div>
  );
};
