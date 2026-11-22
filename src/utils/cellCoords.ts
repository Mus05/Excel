import { CellCoord, CellValue, NumberFormatConfig, SelectionRange } from '../types/spreadsheet';

export function cellKey(row: number, col: number): string {
  return `${row},${col}`;
}

export function parseCellKey(key: string): CellCoord {
  const [r, c] = key.split(',').map(Number);
  return { row: r, col: c };
}

export function colIndexToLetter(col: number): string {
  let letter = '';
  let c = col;
  while (c >= 0) {
    letter = String.fromCharCode((c % 26) + 65) + letter;
    c = Math.floor(c / 26) - 1;
  }
  return letter;
}

export function colLetterToIndex(letters: string): number {
  let col = 0;
  const upper = letters.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    col = col * 26 + (upper.charCodeAt(i) - 64);
  }
  return col - 1;
}

export function formatCellRef(row: number, col: number, absRow = false, absCol = false): string {
  return `${absCol ? '$' : ''}${colIndexToLetter(col)}${absRow ? '$' : ''}${row + 1}`;
}

export function formatRangeRef(range: SelectionRange): string {
  const norm = normalizeRange(range);
  if (norm.start.row === norm.end.row && norm.start.col === norm.end.col) {
    return formatCellRef(norm.start.row, norm.start.col);
  }
  return `${formatCellRef(norm.start.row, norm.start.col)}:${formatCellRef(norm.end.row, norm.end.col)}`;
}

export function parseCellRef(ref: string): { row: number; col: number; absRow: boolean; absCol: boolean } | null {
  const match = ref.trim().match(/^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)$/);
  if (!match) return null;
  const absCol = match[1] === '$';
  const col = colLetterToIndex(match[2]);
  const absRow = match[3] === '$';
  const row = parseInt(match[4], 10) - 1;
  if (row < 0 || col < 0 || col > 16383) return null;
  return { row, col, absRow, absCol };
}

export function parseRangeRef(rangeStr: string): SelectionRange | null {
  const cleaned = rangeStr.trim();
  const parts = cleaned.split(':');
  if (parts.length === 1) {
    const single = parseCellRef(parts[0]);
    if (!single) return null;
    return {
      start: { row: single.row, col: single.col },
      end: { row: single.row, col: single.col },
    };
  }
  if (parts.length === 2) {
    const start = parseCellRef(parts[0]);
    const end = parseCellRef(parts[1]);
    if (!start || !end) return null;
    return normalizeRange({
      start: { row: start.row, col: start.col },
      end: { row: end.row, col: end.col },
    });
  }
  return null;
}

export function normalizeRange(range: SelectionRange): SelectionRange {
  return {
    start: {
      row: Math.min(range.start.row, range.end.row),
      col: Math.min(range.start.col, range.end.col),
    },
    end: {
      row: Math.max(range.start.row, range.end.row),
      col: Math.max(range.start.col, range.end.col),
    },
  };
}

export function isCellInRange(row: number, col: number, range: SelectionRange): boolean {
  const n = normalizeRange(range);
  return row >= n.start.row && row <= n.end.row && col >= n.start.col && col <= n.end.col;
}

/**
 * Shifts relative cell references in a formula when copied or dragged by (dRow, dCol).
 * Absolute references ($A$1, $A1, A$1) preserve their locked axes.
 */
export function shiftFormulaRefs(formula: string, dRow: number, dCol: number): string {
  if (!formula.startsWith('=')) return formula;
  // Tokenize carefully so we don't touch string literals "..."
  let out = '';
  let i = 0;
  while (i < formula.length) {
    if (formula[i] === '"') {
      const start = i;
      i++;
      while (i < formula.length && formula[i] !== '"') {
        if (formula[i] === '\\') i++;
        i++;
      }
      if (i < formula.length) i++;
      out += formula.slice(start, i);
      continue;
    }
    // Check if preceded by identifier character or '!' (we still shift after '!' unless cross-sheet? Actually Excel shifts cross-sheet too!)
    const rest = formula.slice(i);
    // Match cell reference not immediately followed by '(' (function name) or alphanumeric
    const match = rest.match(/^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)(?![A-Za-z0-9_(])/);
    if (match && (i === 0 || !/[A-Za-z0-9_.]/.test(formula[i - 1]))) {
      const absCol = match[1] === '$';
      const col = colLetterToIndex(match[2]);
      const absRow = match[3] === '$';
      const row = parseInt(match[4], 10) - 1;
      // Sanity: ensure letters look like a reasonable column (<= ZZ = 701)
      if (col <= 701 && row >= 0) {
        const newCol = absCol ? col : col + dCol;
        const newRow = absRow ? row : row + dRow;
        if (newCol < 0 || newRow < 0) {
          out += '#REF!';
        } else {
          out += formatCellRef(newRow, newCol, absRow, absCol);
        }
        i += match[0].length;
        continue;
      }
    }
    out += formula[i];
    i++;
  }
  return out;
}

// Convert JS Date to Excel serial number (days since 1899-12-30)
export function dateToExcelSerial(date: Date): number {
  const epoch = Date.UTC(1899, 11, 30);
  const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), date.getHours(), date.getMinutes(), date.getSeconds());
  return (utc - epoch) / 86400000;
}

export function excelSerialToDate(serial: number): Date {
  const epoch = Date.UTC(1899, 11, 30);
  return new Date(epoch + serial * 86400000);
}

function toFraction(num: number): string {
  if (!Number.isFinite(num)) return String(num);
  const sign = num < 0 ? '-' : '';
  const abs = Math.abs(num);
  const whole = Math.floor(abs);
  const frac = abs - whole;
  if (frac < 1e-6) return `${sign}${whole}`;
  let bestNum = 1;
  let bestDen = 1;
  let bestErr = Math.abs(frac - 1);
  for (let den = 2; den <= 64; den++) {
    const n = Math.round(frac * den);
    const err = Math.abs(frac - n / den);
    if (err < bestErr) {
      bestErr = err;
      bestNum = n;
      bestDen = den;
    }
  }
  if (bestNum === 0) return `${sign}${whole}`;
  if (bestNum === bestDen) return `${sign}${whole + 1}`;
  return whole > 0 ? `${sign}${whole} ${bestNum}/${bestDen}` : `${sign}${bestNum}/${bestDen}`;
}

export function formatCellValue(
  val: CellValue,
  format?: NumberFormatConfig,
  defaultCurrency = 'RM'
): string {
  if (val === null || val === undefined || val === '') return '';
  if (typeof val === 'boolean') return val ? 'TRUE' : 'FALSE';
  if (typeof val === 'string') return val;

  const num = val;
  if (!Number.isFinite(num)) return '#NUM!';

  const cat = format?.category ?? 'general';
  const decimals = format?.decimals ?? 2;
  const currency = format?.currencySymbol ?? defaultCurrency;

  switch (cat) {
    case 'general': {
      // Keep clean precision without floating point artifacts
      const rounded = Math.abs(num) < 1e-12 ? 0 : Number(num.toPrecision(12));
      return String(rounded);
    }
    case 'number': {
      return num.toLocaleString('en-US', {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
        useGrouping: format?.useThousands ?? true,
      });
    }
    case 'currency':
    case 'accounting': {
      const absFormatted = Math.abs(num).toLocaleString('en-US', {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
        useGrouping: true,
      });
      if (num < 0) {
        return cat === 'accounting' ? `${currency} (${absFormatted})` : `-${currency} ${absFormatted}`;
      }
      return `${currency} ${absFormatted}`;
    }
    case 'percentage': {
      const pct = num * 100;
      const dec = format?.decimals ?? 1;
      return `${pct.toLocaleString('en-US', {
        minimumFractionDigits: dec,
        maximumFractionDigits: dec,
      })}%`;
    }
    case 'date': {
      const d = excelSerialToDate(num);
      if (isNaN(d.getTime())) return String(num);
      const dd = String(d.getUTCDate()).padStart(2, '0');
      const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
      const yyyy = d.getUTCFullYear();
      return `${dd}/${mm}/${yyyy}`;
    }
    case 'longDate': {
      const d = excelSerialToDate(num);
      if (isNaN(d.getTime())) return String(num);
      return d.toLocaleDateString('en-US', {
        timeZone: 'UTC',
        weekday: 'short',
        year: 'numeric',
        month: 'short',
        day: 'numeric',
      });
    }
    case 'time': {
      const d = excelSerialToDate(num);
      if (isNaN(d.getTime())) return String(num);
      const hh = String(d.getUTCHours()).padStart(2, '0');
      const mi = String(d.getUTCMinutes()).padStart(2, '0');
      const ss = String(d.getUTCSeconds()).padStart(2, '0');
      return `${hh}:${mi}:${ss}`;
    }
    case 'fraction':
      return toFraction(num);
    case 'scientific':
      return num.toExponential(decimals).toUpperCase();
    case 'text':
      return String(num);
    default:
      return String(num);
  }
}

export function tryParseInputDate(raw: string): number | null {
  const trimmed = raw.trim();
  // Match YYYY-MM-DD or DD/MM/YYYY
  const iso = trimmed.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    const y = Number(iso[1]), m = Number(iso[2]) - 1, d = Number(iso[3]);
    const dt = new Date(Date.UTC(y, m, d));
    if (dt.getUTCFullYear() === y && dt.getUTCMonth() === m && dt.getUTCDate() === d) {
      return (dt.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
    }
  }
  const dmy = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmy) {
    const d = Number(dmy[1]), m = Number(dmy[2]) - 1, y = Number(dmy[3]);
    const dt = new Date(Date.UTC(y, m, d));
    if (dt.getUTCFullYear() === y && dt.getUTCMonth() === m && dt.getUTCDate() === d) {
      return (dt.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
    }
  }
  return null;
}
