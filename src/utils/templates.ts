import { Workbook } from '../types/spreadsheet';
import { cellKey } from './cellCoords';
import { createEmptySheet } from './xlsxIO';

export interface WorkbookTemplateInfo {
  id: string;
  name: string;
  description: string;
  build: () => Workbook;
}

export function buildDefaultWorkbook(): Workbook {
  const s1 = createEmptySheet('sheet-sales', 'Sales Overview', 100, 26);
  s1.tabColor = '#107c41';
  s1.freezePanes = { row: 1, col: 0 };
  s1.colWidths = {
    0: 150,
    1: 120,
    2: 110,
    3: 110,
    4: 110,
    5: 110,
    6: 125,
    7: 110,
    8: 115,
    9: 95,
  };
  s1.rowHeights = { 0: 32 };

  const headerStyle = {
    bold: true,
    color: '#ffffff',
    backgroundColor: '#107c41',
    horizontalAlign: 'center' as const,
    verticalAlign: 'middle' as const,
  };

  const headers = [
    'Product Line',
    'Category',
    'Q1 Sales',
    'Q2 Sales',
    'Q3 Sales',
    'Q4 Sales',
    'Total Revenue',
    'Trend',
    'Status',
    'Reviewed',
  ];

  headers.forEach((h, c) => {
    s1.cells[cellKey(0, c)] = { raw: h, style: { ...headerStyle } };
  });

  const rows: [string, string, number, number, number, number, boolean][] = [
    ['ProBook Laptop 15', 'Hardware', 28400, 31200, 29800, 36500, true],
    ['CloudSync Suite', 'Software', 19200, 22500, 25400, 28900, true],
    ['ErgoDesk Dock', 'Accessories', 8400, 7900, 9200, 11400, true],
    ['UltraWide 34"', 'Hardware', 15600, 14200, 16800, 19500, false],
    ['SecuriKey Pro', 'Software', 11500, 13800, 15100, 17900, true],
    ['MechType Keyboard', 'Accessories', 6200, 6800, 6400, 8100, false],
  ];

  const currencyFmt = { category: 'currency' as const, decimals: 0, currencySymbol: 'RM' };

  rows.forEach((r, idx) => {
    const rowIdx = idx + 1;
    const excelRow = rowIdx + 1;
    const zebra = idx % 2 === 1 ? '#f8fafc' : undefined;

    s1.cells[cellKey(rowIdx, 0)] = {
      raw: r[0],
      style: { bold: true, backgroundColor: zebra },
    };
    s1.cells[cellKey(rowIdx, 1)] = {
      raw: r[1],
      style: { backgroundColor: zebra },
      validation: {
        type: 'list',
        formula1: 'Hardware,Software,Accessories,Services',
        showDropdown: true,
      },
    };
    [2, 3, 4, 5].forEach((cIdx, qIdx) => {
      s1.cells[cellKey(rowIdx, cIdx)] = {
        raw: String(r[2 + qIdx]),
        style: { numberFormat: currencyFmt, backgroundColor: zebra },
      };
    });
    s1.cells[cellKey(rowIdx, 6)] = {
      raw: `=SUM(C${excelRow}:F${excelRow})`,
      style: { bold: true, numberFormat: currencyFmt, backgroundColor: zebra },
    };
    s1.cells[cellKey(rowIdx, 7)] = {
      raw: '',
      style: { backgroundColor: zebra },
      sparkline: {
        type: 'line',
        range: `C${excelRow}:F${excelRow}`,
        color: '#107c41',
      },
    };
    s1.cells[cellKey(rowIdx, 8)] = {
      raw: `=IF(G${excelRow}>=75000,"Exceeds","On Track")`,
      style: { horizontalAlign: 'center', backgroundColor: zebra },
    };
    s1.cells[cellKey(rowIdx, 9)] = {
      raw: r[6] ? 'TRUE' : 'FALSE',
      isCheckbox: true,
      style: { horizontalAlign: 'center', backgroundColor: zebra },
    };
  });

  // Total row
  const totalStyle = {
    bold: true,
    backgroundColor: '#ecfdf5',
    numberFormat: currencyFmt,
    borders: {
      top: { style: 'thin' as const, color: '#107c41' },
      bottom: { style: 'double' as const, color: '#107c41' },
    },
  };
  s1.cells[cellKey(7, 0)] = { raw: 'Total', style: { ...totalStyle, numberFormat: undefined } };
  s1.cells[cellKey(7, 1)] = { raw: '=COUNTA(A2:A7)&" items"', style: { ...totalStyle, numberFormat: undefined } };
  ['C', 'D', 'E', 'F', 'G'].forEach((colLetter, idx) => {
    s1.cells[cellKey(7, 2 + idx)] = {
      raw: `=SUM(${colLetter}2:${colLetter}7)`,
      style: { ...totalStyle },
    };
  });
  s1.cells[cellKey(7, 8)] = {
    raw: '=AVERAGE(G2:G7)',
    style: { ...totalStyle },
    comment: 'Average annual revenue per product line',
  };
  s1.cells[cellKey(7, 9)] = {
    raw: '=COUNTIF(J2:J7,TRUE)&"/6"',
    style: { ...totalStyle, horizontalAlign: 'center', numberFormat: undefined },
  };

  // Summary KPI block below
  s1.cells[cellKey(9, 0)] = { raw: 'Key Metrics', style: { bold: true, fontSize: 13, color: '#0f172a' } };
  s1.cells[cellKey(10, 0)] = { raw: 'Highest Line Revenue', style: { color: '#475569' } };
  s1.cells[cellKey(10, 1)] = { raw: '=MAX(G2:G7)', style: { bold: true, numberFormat: currencyFmt } };
  s1.cells[cellKey(11, 0)] = { raw: 'Software Revenue (SUMIF)', style: { color: '#475569' } };
  s1.cells[cellKey(11, 1)] = { raw: '=SUMIF(B2:B7,"Software",G2:G7)', style: { bold: true, numberFormat: currencyFmt } };
  s1.cells[cellKey(12, 0)] = { raw: 'Tax Rate (Named Range)', style: { color: '#475569' } };
  s1.cells[cellKey(12, 1)] = { raw: '0.08', style: { bold: true, numberFormat: { category: 'percentage', decimals: 1 } } };
  s1.cells[cellKey(13, 0)] = { raw: 'Estimated Tax Due', style: { color: '#475569' } };
  s1.cells[cellKey(13, 1)] = { raw: '=G8*TaxRate', style: { bold: true, numberFormat: currencyFmt } };

  s1.conditionalRules = [
    {
      id: 'cf-1',
      range: 'G2:G7',
      type: 'dataBar',
      barColor: '#107c41',
    },
    {
      id: 'cf-2',
      range: 'I2:I7',
      type: 'equal',
      value1: 'Exceeds',
      backgroundColor: '#dcfce7',
      textColor: '#166534',
    },
  ];

  s1.charts = [
    {
      id: 'chart-sales-1',
      type: 'column',
      title: 'Quarterly Sales by Product Line',
      dataRange: 'A1:F7',
      x: 340,
      y: 280,
      width: 560,
      height: 310,
      showLegend: true,
      showGridlines: true,
      palette: 'excel',
    },
  ];

  // Second sheet: Product Catalog (used for VLOOKUP demo)
  const s2 = createEmptySheet('sheet-catalog', 'Product Catalog', 80, 20);
  s2.tabColor = '#0284c7';
  s2.colWidths = { 0: 155, 1: 110, 2: 115, 3: 120 };
  ['Product Name', 'SKU', 'Unit Cost', 'Margin %'].forEach((h, c) => {
    s2.cells[cellKey(0, c)] = {
      raw: h,
      style: { bold: true, color: '#ffffff', backgroundColor: '#0284c7', horizontalAlign: 'center' },
    };
  });
  const catRows: [string, string, number, number][] = [
    ['ProBook Laptop 15', 'PB-1500', 3200, 0.28],
    ['CloudSync Suite', 'CS-900', 450, 0.72],
    ['ErgoDesk Dock', 'ED-220', 180, 0.45],
    ['UltraWide 34"', 'UW-340', 1450, 0.31],
    ['SecuriKey Pro', 'SK-110', 120, 0.68],
    ['MechType Keyboard', 'MK-080', 210, 0.42],
  ];
  catRows.forEach((row, i) => {
    const r = i + 1;
    s2.cells[cellKey(r, 0)] = { raw: row[0], style: { bold: true } };
    s2.cells[cellKey(r, 1)] = { raw: row[1] };
    s2.cells[cellKey(r, 2)] = { raw: String(row[2]), style: { numberFormat: currencyFmt } };
    s2.cells[cellKey(r, 3)] = { raw: String(row[3]), style: { numberFormat: { category: 'percentage', decimals: 0 } } };
  });

  return {
    id: 'wb-default',
    title: 'Quarterly Sales & Budget.xlsx',
    sheets: [s1, s2],
    activeSheetId: s1.id,
    namedRanges: [{ name: 'TaxRate', sheetId: s1.id, range: 'B13', comment: 'Corporate SST rate (8%)' }],
    calcMode: 'automatic',
    showFormulas: false,
    showFormulaBar: true,
    zoom: 100,
    theme: 'light',
    currencySymbol: 'RM',
  };
}

export function buildBlankWorkbook(): Workbook {
  const s1 = createEmptySheet('sheet-1', 'Sheet1', 100, 26);
  return {
    id: `wb-${Date.now()}`,
    title: 'Book1.xlsx',
    sheets: [s1],
    activeSheetId: s1.id,
    namedRanges: [],
    calcMode: 'automatic',
    showFormulas: false,
    showFormulaBar: true,
    zoom: 100,
    theme: 'light',
    currencySymbol: 'RM',
  };
}

export function buildMonthlyBudgetWorkbook(): Workbook {
  const s1 = createEmptySheet('sheet-budget', 'Monthly Budget', 80, 20);
  s1.tabColor = '#107c41';
  s1.colWidths = { 0: 190, 1: 120, 2: 120, 3: 120, 4: 130 };
  const cur = { category: 'currency' as const, decimals: 2, currencySymbol: 'RM' };

  s1.cells[cellKey(0, 0)] = { raw: 'Personal Monthly Budget Planner', style: { bold: true, fontSize: 15, color: '#107c41' } };
  s1.mergedRanges = [{ startRow: 0, startCol: 0, endRow: 0, endCol: 4 }];

  ['Expense Category', 'Budgeted', 'Actual', 'Variance', 'Status'].forEach((h, c) => {
    s1.cells[cellKey(2, c)] = {
      raw: h,
      style: { bold: true, color: '#ffffff', backgroundColor: '#1e293b', horizontalAlign: 'center' },
    };
  });

  const items: [string, number, number][] = [
    ['Housing & Rent', 1800, 1800],
    ['Groceries & Dining', 950, 1085],
    ['Utilities & Internet', 320, 295],
    ['Transportation & Fuel', 450, 410],
    ['Savings & Investments', 1200, 1200],
    ['Entertainment & Travel', 400, 520],
  ];

  items.forEach(([cat, bud, act], i) => {
    const r = i + 3;
    const er = r + 1;
    s1.cells[cellKey(r, 0)] = { raw: cat };
    s1.cells[cellKey(r, 1)] = { raw: String(bud), style: { numberFormat: cur } };
    s1.cells[cellKey(r, 2)] = { raw: String(act), style: { numberFormat: cur } };
    s1.cells[cellKey(r, 3)] = { raw: `=B${er}-C${er}`, style: { numberFormat: cur, bold: true } };
    s1.cells[cellKey(r, 4)] = { raw: `=IF(D${er}>=0,"Under Budget","Over Budget")`, style: { horizontalAlign: 'center' } };
  });

  s1.cells[cellKey(9, 0)] = { raw: 'Total Expenses', style: { bold: true, backgroundColor: '#f1f5f9' } };
  s1.cells[cellKey(9, 1)] = { raw: '=SUM(B4:B9)', style: { bold: true, numberFormat: cur, backgroundColor: '#f1f5f9' } };
  s1.cells[cellKey(9, 2)] = { raw: '=SUM(C4:C9)', style: { bold: true, numberFormat: cur, backgroundColor: '#f1f5f9' } };
  s1.cells[cellKey(9, 3)] = { raw: '=SUM(D4:D9)', style: { bold: true, numberFormat: cur, backgroundColor: '#f1f5f9' } };
  s1.cells[cellKey(9, 4)] = { raw: '=IF(D10>=0,"Surplus","Deficit")', style: { bold: true, horizontalAlign: 'center', backgroundColor: '#f1f5f9' } };

  s1.conditionalRules = [
    { id: 'mb-1', range: 'D4:D9', type: 'lessThan', value1: '0', backgroundColor: '#fee2e2', textColor: '#991b1b' },
    { id: 'mb-2', range: 'D4:D9', type: 'greaterThan', value1: '0', backgroundColor: '#dcfce7', textColor: '#166534' },
  ];

  s1.charts = [
    {
      id: 'chart-budget',
      type: 'bar',
      title: 'Budgeted vs Actual Spending',
      dataRange: 'A3:C9',
      x: 20,
      y: 300,
      width: 560,
      height: 290,
      showLegend: true,
      showGridlines: true,
      palette: 'ocean',
    },
  ];

  return {
    id: `wb-${Date.now()}`,
    title: 'Monthly_Budget.xlsx',
    sheets: [s1],
    activeSheetId: s1.id,
    namedRanges: [],
    calcMode: 'automatic',
    showFormulas: false,
    showFormulaBar: true,
    zoom: 100,
    theme: 'light',
    currencySymbol: 'RM',
  };
}

export function buildInvoiceWorkbook(): Workbook {
  const s1 = createEmptySheet('sheet-inv', 'Invoice', 60, 15);
  s1.colWidths = { 0: 220, 1: 90, 2: 120, 3: 130 };
  const cur = { category: 'currency' as const, decimals: 2, currencySymbol: 'RM' };

  s1.cells[cellKey(0, 0)] = { raw: 'TAX INVOICE #INV-2026-041', style: { bold: true, fontSize: 16, color: '#107c41' } };
  s1.cells[cellKey(1, 0)] = { raw: 'Bill To: Horizon Digital Sdn Bhd', style: { bold: true } };
  s1.cells[cellKey(2, 0)] = { raw: 'Issue Date:', style: { color: '#64748b' } };
  s1.cells[cellKey(2, 1)] = { raw: '=TODAY()', style: { numberFormat: { category: 'date' } } };

  ['Item Description', 'Qty', 'Unit Price', 'Line Total'].forEach((h, c) => {
    s1.cells[cellKey(4, c)] = { raw: h, style: { bold: true, color: '#ffffff', backgroundColor: '#107c41', horizontalAlign: 'center' } };
  });

  const lines: [string, number, number][] = [
    ['Enterprise Cloud Setup & Migration', 1, 4500],
    ['Annual Support Retainer (Months)', 12, 650],
    ['Staff Onboarding Workshop', 2, 1200],
    ['Custom API Integration Module', 1, 2800],
  ];

  lines.forEach(([desc, qty, price], i) => {
    const r = 5 + i;
    const er = r + 1;
    s1.cells[cellKey(r, 0)] = { raw: desc };
    s1.cells[cellKey(r, 1)] = { raw: String(qty), style: { horizontalAlign: 'center' } };
    s1.cells[cellKey(r, 2)] = { raw: String(price), style: { numberFormat: cur } };
    s1.cells[cellKey(r, 3)] = { raw: `=B${er}*C${er}`, style: { numberFormat: cur, bold: true } };
  });

  s1.cells[cellKey(10, 2)] = { raw: 'Subtotal', style: { bold: true, horizontalAlign: 'right' } };
  s1.cells[cellKey(10, 3)] = { raw: '=SUM(D6:D9)', style: { numberFormat: cur, bold: true } };
  s1.cells[cellKey(11, 2)] = { raw: 'SST (8%)', style: { horizontalAlign: 'right' } };
  s1.cells[cellKey(11, 3)] = { raw: '=D11*0.08', style: { numberFormat: cur } };
  s1.cells[cellKey(12, 2)] = { raw: 'Grand Total', style: { bold: true, horizontalAlign: 'right', backgroundColor: '#ecfdf5' } };
  s1.cells[cellKey(12, 3)] = { raw: '=D11+D12', style: { numberFormat: cur, bold: true, backgroundColor: '#ecfdf5', borders: { top: { style: 'thin', color: '#107c41' }, bottom: { style: 'double', color: '#107c41' } } } };

  return {
    id: `wb-${Date.now()}`,
    title: 'Tax_Invoice.xlsx',
    sheets: [s1],
    activeSheetId: s1.id,
    namedRanges: [],
    calcMode: 'automatic',
    showFormulas: false,
    showFormulaBar: true,
    zoom: 100,
    theme: 'light',
    currencySymbol: 'RM',
  };
}

export function buildLoanCalculatorWorkbook(): Workbook {
  const s1 = createEmptySheet('sheet-loan', 'Loan Amortization', 60, 15);
  s1.colWidths = { 0: 190, 1: 140, 2: 140, 3: 140 };
  const cur = { category: 'currency' as const, decimals: 2, currencySymbol: 'RM' };

  s1.cells[cellKey(0, 0)] = { raw: 'Loan & Financing Calculator (PMT)', style: { bold: true, fontSize: 15, color: '#107c41' } };
  s1.cells[cellKey(2, 0)] = { raw: 'Loan Amount (Principal)', style: { bold: true } };
  s1.cells[cellKey(2, 1)] = { raw: '250000', style: { numberFormat: cur, backgroundColor: '#fef9c3' } };
  s1.cells[cellKey(3, 0)] = { raw: 'Annual Interest Rate', style: { bold: true } };
  s1.cells[cellKey(3, 1)] = { raw: '0.045', style: { numberFormat: { category: 'percentage', decimals: 2 }, backgroundColor: '#fef9c3' } };
  s1.cells[cellKey(4, 0)] = { raw: 'Loan Term (Years)', style: { bold: true } };
  s1.cells[cellKey(4, 1)] = { raw: '15', style: { backgroundColor: '#fef9c3' } };

  s1.cells[cellKey(6, 0)] = { raw: 'Monthly Payment (PMT)', style: { bold: true, backgroundColor: '#ecfdf5' } };
  s1.cells[cellKey(6, 1)] = { raw: '=-PMT(B4/12,B5*12,B3)', style: { bold: true, numberFormat: cur, backgroundColor: '#ecfdf5' } };
  s1.cells[cellKey(7, 0)] = { raw: 'Total Payment', style: { bold: true } };
  s1.cells[cellKey(7, 1)] = { raw: '=B7*B5*12', style: { numberFormat: cur } };
  s1.cells[cellKey(8, 0)] = { raw: 'Total Interest Paid', style: { bold: true } };
  s1.cells[cellKey(8, 1)] = { raw: '=B8-B3', style: { numberFormat: cur, color: '#dc2626' } };

  return {
    id: `wb-${Date.now()}`,
    title: 'Loan_Calculator.xlsx',
    sheets: [s1],
    activeSheetId: s1.id,
    namedRanges: [],
    calcMode: 'automatic',
    showFormulas: false,
    showFormulaBar: true,
    zoom: 100,
    theme: 'light',
    currencySymbol: 'RM',
  };
}

export const WORKBOOK_TEMPLATES: WorkbookTemplateInfo[] = [
  {
    id: 'blank',
    name: 'Blank Workbook',
    description: 'Start fresh with an empty spreadsheet grid.',
    build: buildBlankWorkbook,
  },
  {
    id: 'sales-dashboard',
    name: 'Quarterly Sales & Budget',
    description: 'Multi-sheet sales report with SUM/IF/SUMIF formulas, sparklines, checkboxes, and chart.',
    build: buildDefaultWorkbook,
  },
  {
    id: 'monthly-budget',
    name: 'Monthly Personal Budget',
    description: 'Track budgeted vs actual expenses with variance conditional formatting and bar chart.',
    build: buildMonthlyBudgetWorkbook,
  },
  {
    id: 'invoice',
    name: 'Tax Invoice Generator',
    description: 'Clean itemized invoice with automatic TODAY() date, subtotal, SST calculation, and totals.',
    build: buildInvoiceWorkbook,
  },
  {
    id: 'loan-calc',
    name: 'Loan Calculator (PMT)',
    description: 'Financial model using PMT() to compute monthly installments and total interest.',
    build: buildLoanCalculatorWorkbook,
  },
];
