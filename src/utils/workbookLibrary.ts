import { Workbook } from '../types/spreadsheet';
import { buildDefaultWorkbook, buildMonthlyBudgetWorkbook } from './templates';

const LIBRARY_STORAGE_KEY = 'excelgrid_library_v1';
const LEGACY_STORAGE_KEY = 'excelgrid_workbook_v1';

export interface LibraryWorkbookItem {
  workbook: Workbook;
  updatedAt: number;
  pinned: boolean;
}

export function getWorkbookStats(wb: Workbook): {
  sheets: number;
  cells: number;
  formulas: number;
  charts: number;
} {
  let cells = 0;
  let formulas = 0;
  let charts = 0;
  for (const s of wb.sheets) {
    charts += s.charts.length;
    for (const c of Object.values(s.cells)) {
      if (c.raw !== '') cells++;
      if (c.raw.startsWith('=')) formulas++;
    }
  }
  return {
    sheets: wb.sheets.length,
    cells,
    formulas,
    charts,
  };
}

export function loadWorkbookLibrary(): LibraryWorkbookItem[] {
  try {
    const raw = localStorage.getItem(LIBRARY_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as LibraryWorkbookItem[];
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed.filter((item) => item?.workbook?.id && Array.isArray(item.workbook.sheets));
      }
    }

    // Migrate legacy single workbook if present
    const legacy = localStorage.getItem(LEGACY_STORAGE_KEY);
    if (legacy) {
      const wb = JSON.parse(legacy) as Workbook;
      if (wb && Array.isArray(wb.sheets) && wb.sheets.length > 0) {
        const seeded: LibraryWorkbookItem[] = [
          { workbook: wb, updatedAt: Date.now(), pinned: true },
          {
            workbook: buildMonthlyBudgetWorkbook(),
            updatedAt: Date.now() - 1000 * 60 * 60 * 3,
            pinned: false,
          },
        ];
        saveWorkbookLibrary(seeded);
        return seeded;
      }
    }
  } catch {
    // ignore corrupted localStorage
  }

  const initial: LibraryWorkbookItem[] = [
    {
      workbook: buildDefaultWorkbook(),
      updatedAt: Date.now() - 1000 * 60 * 12,
      pinned: true,
    },
    {
      workbook: buildMonthlyBudgetWorkbook(),
      updatedAt: Date.now() - 1000 * 60 * 60 * 5,
      pinned: false,
    },
  ];
  saveWorkbookLibrary(initial);
  return initial;
}

export function saveWorkbookLibrary(items: LibraryWorkbookItem[]): void {
  try {
    localStorage.setItem(LIBRARY_STORAGE_KEY, JSON.stringify(items));
  } catch {
    // storage quota full
  }
}

export function formatRelativeTime(ts: number): string {
  const diffSec = Math.max(0, Math.floor((Date.now() - ts) / 1000));
  if (diffSec < 60) return 'Just now';
  const diffMin = Math.floor(diffSec / 60);
  if (diffMin < 60) return `${diffMin}m ago`;
  const diffHr = Math.floor(diffMin / 60);
  if (diffHr < 24) return `${diffHr}h ago`;
  return new Date(ts).toLocaleDateString('en-US', {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
  });
}
