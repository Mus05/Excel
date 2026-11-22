import * as XLSX from 'xlsx';
import { CellData, Sheet, Workbook } from '../types/spreadsheet';
import { cellKey, colIndexToLetter, parseCellKey } from './cellCoords';
import { WorkbookEvaluator } from '../engine/formulaEngine';

export function createEmptySheet(id: string, name: string, rowCount = 100, colCount = 26): Sheet {
  return {
    id,
    name,
    rowCount,
    colCount,
    cells: {},
    colWidths: {},
    rowHeights: {},
    hiddenRows: {},
    hiddenCols: {},
    mergedRanges: [],
    conditionalRules: [],
    charts: [],
    autoFilter: null,
    tables: [],
    freezePanes: { row: 0, col: 0 },
    protected: false,
    pageSetup: {
      orientation: 'portrait',
      margins: 'normal',
      showGridlines: true,
      showHeadings: true,
      printGridlines: true,
      printHeadings: false,
    },
  };
}

export function exportWorkbookToXLSX(workbook: Workbook) {
  const wb = XLSX.utils.book_new();
  const evaluator = new WorkbookEvaluator(workbook.sheets, workbook.namedRanges, workbook.currencySymbol);

  for (const sheet of workbook.sheets) {
    let maxR = 0;
    let maxC = 0;
    for (const k of Object.keys(sheet.cells)) {
      const { row, col } = parseCellKey(k);
      if (row > maxR) maxR = row;
      if (col > maxC) maxC = col;
    }

    const ws: XLSX.WorkSheet = {};
    ws['!ref'] = `A1:${colIndexToLetter(maxC)}${maxR + 1}`;

    for (const [k, cell] of Object.entries(sheet.cells)) {
      if (!cell || cell.raw === '') continue;
      const { row, col } = parseCellKey(k);
      const addr = `${colIndexToLetter(col)}${row + 1}`;
      const ev = evaluator.evaluateCell(sheet.id, row, col);
      const cellObj: XLSX.CellObject = { t: 's', v: '' };

      if (cell.raw.startsWith('=')) {
        cellObj.f = cell.raw.slice(1);
        if (typeof ev.value === 'number') {
          cellObj.t = 'n';
          cellObj.v = ev.value;
        } else if (typeof ev.value === 'boolean') {
          cellObj.t = 'b';
          cellObj.v = ev.value;
        } else {
          cellObj.t = 's';
          cellObj.v = ev.formatted;
        }
      } else if (typeof ev.value === 'number') {
        cellObj.t = 'n';
        cellObj.v = ev.value;
      } else if (typeof ev.value === 'boolean') {
        cellObj.t = 'b';
        cellObj.v = ev.value;
      } else {
        cellObj.t = 's';
        cellObj.v = cell.raw;
      }
      ws[addr] = cellObj;
    }

    // Export column widths
    const cols: XLSX.ColInfo[] = [];
    for (let c = 0; c <= maxC; c++) {
      const w = sheet.colWidths[c];
      cols.push({ wpx: w || 100 });
    }
    ws['!cols'] = cols;

    // Export merges
    if (sheet.mergedRanges.length > 0) {
      ws['!merges'] = sheet.mergedRanges.map((m) => ({
        s: { r: m.startRow, c: m.startCol },
        e: { r: m.endRow, c: m.endCol },
      }));
    }

    XLSX.utils.book_append_sheet(wb, ws, sheet.name.slice(0, 31));
  }

  const filename = workbook.title.toLowerCase().endsWith('.xlsx')
    ? workbook.title
    : `${workbook.title}.xlsx`;
  XLSX.writeFile(wb, filename);
}

export function exportSheetToCSV(workbook: Workbook, sheetId: string) {
  const sheet = workbook.sheets.find((s) => s.id === sheetId);
  if (!sheet) return;
  const evaluator = new WorkbookEvaluator(workbook.sheets, workbook.namedRanges, workbook.currencySymbol);

  let maxR = 0;
  let maxC = 0;
  for (const k of Object.keys(sheet.cells)) {
    const { row, col } = parseCellKey(k);
    if (row > maxR) maxR = row;
    if (col > maxC) maxC = col;
  }

  const lines: string[] = [];
  for (let r = 0; r <= maxR; r++) {
    const rowVals: string[] = [];
    for (let c = 0; c <= maxC; c++) {
      const ev = evaluator.evaluateCell(sheet.id, r, c);
      const val = ev.formatted;
      if (val.includes(',') || val.includes('"') || val.includes('\n')) {
        rowVals.push(`"${val.replace(/"/g, '""')}"`);
      } else {
        rowVals.push(val);
      }
    }
    lines.push(rowVals.join(','));
  }

  const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${sheet.name}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

export function exportWorkbookToJSON(workbook: Workbook) {
  const blob = new Blob([JSON.stringify(workbook, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${workbook.title.replace(/\.xlsx$/i, '')}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

export async function importFileToWorkbook(file: File): Promise<Workbook> {
  const ext = file.name.split('.').pop()?.toLowerCase();
  if (ext === 'json') {
    const text = await file.text();
    const parsed = JSON.parse(text) as Workbook;
    if (parsed && Array.isArray(parsed.sheets) && parsed.sheets.length > 0) {
      return parsed;
    }
    throw new Error('Invalid workbook JSON format.');
  }

  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array', cellFormula: true, cellStyles: true });

  const sheets: Sheet[] = wb.SheetNames.map((name, idx) => {
    const ws = wb.Sheets[name];
    const ref = ws['!ref'];
    const decoded = ref ? XLSX.utils.decode_range(ref) : { s: { r: 0, c: 0 }, e: { r: 20, c: 10 } };
    const rowCount = Math.max(100, decoded.e.r + 30);
    const colCount = Math.max(26, decoded.e.c + 10);

    const sheet = createEmptySheet(`sheet-${Date.now()}-${idx}`, name, rowCount, colCount);

    for (let r = decoded.s.r; r <= decoded.e.r; r++) {
      for (let c = decoded.s.c; c <= decoded.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cellObj = ws[addr];
        if (!cellObj) continue;
        let raw = '';
        if (cellObj.f) {
          raw = `=${cellObj.f}`;
        } else if (cellObj.v !== undefined && cellObj.v !== null) {
          raw = String(cellObj.v);
        }
        if (raw !== '') {
          const data: CellData = { raw };
          sheet.cells[cellKey(r, c)] = data;
        }
      }
    }

    if (ws['!merges']) {
      sheet.mergedRanges = ws['!merges'].map((m) => ({
        startRow: m.s.r,
        startCol: m.s.c,
        endRow: m.e.r,
        endCol: m.e.c,
      }));
    }

    if (ws['!cols']) {
      ws['!cols'].forEach((colInfo, colIdx) => {
        if (colInfo?.wpx) sheet.colWidths[colIdx] = Math.max(40, Math.min(400, colInfo.wpx));
      });
    }

    return sheet;
  });

  if (sheets.length === 0) {
    sheets.push(createEmptySheet(`sheet-${Date.now()}`, 'Sheet1'));
  }

  return {
    id: `wb-${Date.now()}`,
    title: file.name,
    sheets,
    activeSheetId: sheets[0].id,
    namedRanges: [],
    calcMode: 'automatic',
    showFormulas: false,
    showFormulaBar: true,
    zoom: 100,
    theme: 'light',
    currencySymbol: 'RM',
  };
}
