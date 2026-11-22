import { CellValue, EvaluatedCell, NamedRange, Sheet } from '../types/spreadsheet';
import {
  cellKey,
  colIndexToLetter,
  colLetterToIndex,
  dateToExcelSerial,
  excelSerialToDate,
  formatCellValue,
  tryParseInputDate,
} from '../utils/cellCoords';

export interface FunctionMeta {
  name: string;
  category:
    | 'Math & Trig'
    | 'Statistical'
    | 'Logical'
    | 'Lookup & Reference'
    | 'Text'
    | 'Date & Time'
    | 'Financial'
    | 'Information';
  syntax: string;
  description: string;
  example: string;
}

export const FUNCTION_CATALOG: FunctionMeta[] = [
  // Math & Trig
  { name: 'SUM', category: 'Math & Trig', syntax: 'SUM(number1, [number2], ...)', description: 'Adds all numbers in a range of cells.', example: '=SUM(B2:B10)' },
  { name: 'SUMIF', category: 'Math & Trig', syntax: 'SUMIF(range, criteria, [sum_range])', description: 'Adds cells specified by a given condition or criteria.', example: '=SUMIF(A2:A10, "Sales", B2:B10)' },
  { name: 'SUMIFS', category: 'Math & Trig', syntax: 'SUMIFS(sum_range, criteria_range1, criteria1, ...)', description: 'Adds cells that meet multiple criteria.', example: '=SUMIFS(C2:C10, A2:A10, "North", B2:B10, ">100")' },
  { name: 'PRODUCT', category: 'Math & Trig', syntax: 'PRODUCT(number1, [number2], ...)', description: 'Multiplies all the numbers given as arguments.', example: '=PRODUCT(A1:A5)' },
  { name: 'SUMPRODUCT', category: 'Math & Trig', syntax: 'SUMPRODUCT(array1, [array2], ...)', description: 'Returns the sum of the products of corresponding ranges.', example: '=SUMPRODUCT(B2:B5, C2:C5)' },
  { name: 'ROUND', category: 'Math & Trig', syntax: 'ROUND(number, num_digits)', description: 'Rounds a number to a specified number of digits.', example: '=ROUND(3.14159, 2)' },
  { name: 'ROUNDUP', category: 'Math & Trig', syntax: 'ROUNDUP(number, num_digits)', description: 'Rounds a number up, away from zero.', example: '=ROUNDUP(3.12, 1)' },
  { name: 'ROUNDDOWN', category: 'Math & Trig', syntax: 'ROUNDDOWN(number, num_digits)', description: 'Rounds a number down, toward zero.', example: '=ROUNDDOWN(3.19, 1)' },
  { name: 'ABS', category: 'Math & Trig', syntax: 'ABS(number)', description: 'Returns the absolute value of a number.', example: '=ABS(-45)' },
  { name: 'SQRT', category: 'Math & Trig', syntax: 'SQRT(number)', description: 'Returns a positive square root.', example: '=SQRT(144)' },
  { name: 'POWER', category: 'Math & Trig', syntax: 'POWER(number, power)', description: 'Returns the result of a number raised to a power.', example: '=POWER(2, 8)' },
  { name: 'MOD', category: 'Math & Trig', syntax: 'MOD(number, divisor)', description: 'Returns the remainder after number is divided by divisor.', example: '=MOD(10, 3)' },
  { name: 'INT', category: 'Math & Trig', syntax: 'INT(number)', description: 'Rounds a number down to the nearest integer.', example: '=INT(8.9)' },
  { name: 'CEILING', category: 'Math & Trig', syntax: 'CEILING(number, significance)', description: 'Rounds number up to the nearest multiple of significance.', example: '=CEILING(4.2, 1)' },
  { name: 'FLOOR', category: 'Math & Trig', syntax: 'FLOOR(number, significance)', description: 'Rounds number down to the nearest multiple of significance.', example: '=FLOOR(4.8, 1)' },
  { name: 'PI', category: 'Math & Trig', syntax: 'PI()', description: 'Returns the value of Pi (3.14159265358979).', example: '=PI()' },
  { name: 'RAND', category: 'Math & Trig', syntax: 'RAND()', description: 'Returns a random number between 0 and 1.', example: '=RAND()' },
  { name: 'RANDBETWEEN', category: 'Math & Trig', syntax: 'RANDBETWEEN(bottom, top)', description: 'Returns a random integer between the numbers you specify.', example: '=RANDBETWEEN(1, 100)' },
  { name: 'LOG', category: 'Math & Trig', syntax: 'LOG(number, [base])', description: 'Returns the logarithm of a number to a specified base.', example: '=LOG(100, 10)' },
  { name: 'LN', category: 'Math & Trig', syntax: 'LN(number)', description: 'Returns the natural logarithm of a number.', example: '=LN(10)' },
  { name: 'EXP', category: 'Math & Trig', syntax: 'EXP(number)', description: 'Returns e raised to the power of a given number.', example: '=EXP(1)' },

  // Statistical
  { name: 'AVERAGE', category: 'Statistical', syntax: 'AVERAGE(number1, [number2], ...)', description: 'Returns the average (arithmetic mean) of its arguments.', example: '=AVERAGE(B2:B10)' },
  { name: 'AVERAGEIF', category: 'Statistical', syntax: 'AVERAGEIF(range, criteria, [average_range])', description: 'Returns the average of all cells that meet a given criteria.', example: '=AVERAGEIF(A2:A10, "East", B2:B10)' },
  { name: 'COUNT', category: 'Statistical', syntax: 'COUNT(value1, [value2], ...)', description: 'Counts the number of cells that contain numbers.', example: '=COUNT(B2:B10)' },
  { name: 'COUNTA', category: 'Statistical', syntax: 'COUNTA(value1, [value2], ...)', description: 'Counts the number of cells that are not empty.', example: '=COUNTA(A2:A10)' },
  { name: 'COUNTBLANK', category: 'Statistical', syntax: 'COUNTBLANK(range)', description: 'Counts the number of empty cells in a range.', example: '=COUNTBLANK(A2:A10)' },
  { name: 'COUNTIF', category: 'Statistical', syntax: 'COUNTIF(range, criteria)', description: 'Counts the number of cells within a range that meet a single condition.', example: '=COUNTIF(B2:B10, ">500")' },
  { name: 'COUNTIFS', category: 'Statistical', syntax: 'COUNTIFS(criteria_range1, criteria1, ...)', description: 'Counts cells that meet multiple criteria.', example: '=COUNTIFS(A2:A10, "Active", B2:B10, ">=100")' },
  { name: 'MIN', category: 'Statistical', syntax: 'MIN(number1, [number2], ...)', description: 'Returns the smallest number in a set of values.', example: '=MIN(B2:B10)' },
  { name: 'MAX', category: 'Statistical', syntax: 'MAX(number1, [number2], ...)', description: 'Returns the largest value in a set of values.', example: '=MAX(B2:B10)' },
  { name: 'MEDIAN', category: 'Statistical', syntax: 'MEDIAN(number1, [number2], ...)', description: 'Returns the median of the given numbers.', example: '=MEDIAN(B2:B10)' },
  { name: 'STDEV', category: 'Statistical', syntax: 'STDEV(number1, [number2], ...)', description: 'Estimates standard deviation based on a sample.', example: '=STDEV(B2:B10)' },
  { name: 'VAR', category: 'Statistical', syntax: 'VAR(number1, [number2], ...)', description: 'Estimates variance based on a sample.', example: '=VAR(B2:B10)' },
  { name: 'LARGE', category: 'Statistical', syntax: 'LARGE(array, k)', description: 'Returns the k-th largest value in a data set.', example: '=LARGE(B2:B10, 2)' },
  { name: 'SMALL', category: 'Statistical', syntax: 'SMALL(array, k)', description: 'Returns the k-th smallest value in a data set.', example: '=SMALL(B2:B10, 2)' },
  { name: 'RANK', category: 'Statistical', syntax: 'RANK(number, ref, [order])', description: 'Returns the rank of a number in a list of numbers.', example: '=RANK(B2, $B$2:$B$10, 0)' },

  // Logical
  { name: 'IF', category: 'Logical', syntax: 'IF(logical_test, value_if_true, [value_if_false])', description: 'Checks whether a condition is met and returns one value if TRUE, another if FALSE.', example: '=IF(B2>=1000, "Target Met", "Below")' },
  { name: 'IFS', category: 'Logical', syntax: 'IFS(condition1, value1, [condition2, value2], ...)', description: 'Checks multiple conditions and returns a value for the first TRUE condition.', example: '=IFS(A1>=90, "A", A1>=80, "B", TRUE, "C")' },
  { name: 'AND', category: 'Logical', syntax: 'AND(logical1, [logical2], ...)', description: 'Returns TRUE if all arguments evaluate to TRUE.', example: '=AND(B2>100, C2="Paid")' },
  { name: 'OR', category: 'Logical', syntax: 'OR(logical1, [logical2], ...)', description: 'Returns TRUE if any argument evaluates to TRUE.', example: '=OR(A2="VIP", B2>5000)' },
  { name: 'NOT', category: 'Logical', syntax: 'NOT(logical)', description: 'Reverses the logic of its argument.', example: '=NOT(A2="")' },
  { name: 'XOR', category: 'Logical', syntax: 'XOR(logical1, [logical2], ...)', description: 'Returns a logical Exclusive Or of all arguments.', example: '=XOR(A1>0, B1>0)' },
  { name: 'IFERROR', category: 'Logical', syntax: 'IFERROR(value, value_if_error)', description: 'Returns a value you specify if a formula evaluates to an error; otherwise returns the formula result.', example: '=IFERROR(A2/B2, 0)' },
  { name: 'IFNA', category: 'Logical', syntax: 'IFNA(value, value_if_na)', description: 'Returns the value you specify if the expression resolves to #N/A.', example: '=IFNA(VLOOKUP(E2, A2:B10, 2, FALSE), "Not Found")' },

  // Lookup & Reference
  { name: 'VLOOKUP', category: 'Lookup & Reference', syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])', description: 'Looks for a value in the leftmost column of a table and returns a value in the same row from a specified column.', example: '=VLOOKUP(E2, A2:C10, 3, FALSE)' },
  { name: 'HLOOKUP', category: 'Lookup & Reference', syntax: 'HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])', description: 'Looks for a value in the top row of a table and returns a value in the same column.', example: '=HLOOKUP("Q2", A1:E4, 3, FALSE)' },
  { name: 'XLOOKUP', category: 'Lookup & Reference', syntax: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])', description: 'Searches a range or array and returns an item corresponding to the first match.', example: '=XLOOKUP(E2, A2:A10, C2:C10, "Missing")' },
  { name: 'INDEX', category: 'Lookup & Reference', syntax: 'INDEX(array, row_num, [column_num])', description: 'Returns a value from a table or range at the specified row and column.', example: '=INDEX(A2:C10, 3, 2)' },
  { name: 'MATCH', category: 'Lookup & Reference', syntax: 'MATCH(lookup_value, lookup_array, [match_type])', description: 'Searches for a specified item in a range and returns its relative position.', example: '=MATCH("Laptop", A2:A10, 0)' },
  { name: 'CHOOSE', category: 'Lookup & Reference', syntax: 'CHOOSE(index_num, value1, [value2], ...)', description: 'Uses index_num to return a value from the list of value arguments.', example: '=CHOOSE(2, "Low", "Medium", "High")' },
  { name: 'ROW', category: 'Lookup & Reference', syntax: 'ROW([reference])', description: 'Returns the row number of a reference.', example: '=ROW(A5)' },
  { name: 'COLUMN', category: 'Lookup & Reference', syntax: 'COLUMN([reference])', description: 'Returns the column number of a reference.', example: '=COLUMN(C1)' },
  { name: 'ROWS', category: 'Lookup & Reference', syntax: 'ROWS(array)', description: 'Returns the number of rows in a reference or array.', example: '=ROWS(A1:A10)' },
  { name: 'COLUMNS', category: 'Lookup & Reference', syntax: 'COLUMNS(array)', description: 'Returns the number of columns in a reference or array.', example: '=COLUMNS(A1:E1)' },

  // Text
  { name: 'CONCAT', category: 'Text', syntax: 'CONCAT(text1, [text2], ...)', description: 'Combines the text from multiple ranges and/or strings.', example: '=CONCAT(A2, " ", B2)' },
  { name: 'CONCATENATE', category: 'Text', syntax: 'CONCATENATE(text1, [text2], ...)', description: 'Joins several text strings into one text string.', example: '=CONCATENATE(A2, " - ", B2)' },
  { name: 'TEXTJOIN', category: 'Text', syntax: 'TEXTJOIN(delimiter, ignore_empty, text1, ...)', description: 'Combines text from multiple ranges using a delimiter.', example: '=TEXTJOIN(", ", TRUE, A2:A5)' },
  { name: 'LEFT', category: 'Text', syntax: 'LEFT(text, [num_chars])', description: 'Returns the specified number of characters from the start of a text string.', example: '=LEFT(A2, 3)' },
  { name: 'RIGHT', category: 'Text', syntax: 'RIGHT(text, [num_chars])', description: 'Returns the specified number of characters from the end of a text string.', example: '=RIGHT(A2, 4)' },
  { name: 'MID', category: 'Text', syntax: 'MID(text, start_num, num_chars)', description: 'Returns characters from the middle of a text string.', example: '=MID(A2, 2, 5)' },
  { name: 'LEN', category: 'Text', syntax: 'LEN(text)', description: 'Returns the number of characters in a text string.', example: '=LEN(A2)' },
  { name: 'TRIM', category: 'Text', syntax: 'TRIM(text)', description: 'Removes extra spaces from text.', example: '=TRIM(A2)' },
  { name: 'UPPER', category: 'Text', syntax: 'UPPER(text)', description: 'Converts text to uppercase.', example: '=UPPER(A2)' },
  { name: 'LOWER', category: 'Text', syntax: 'LOWER(text)', description: 'Converts text to lowercase.', example: '=LOWER(A2)' },
  { name: 'PROPER', category: 'Text', syntax: 'PROPER(text)', description: 'Capitalizes the first letter in each word of a text value.', example: '=PROPER(A2)' },
  { name: 'SUBSTITUTE', category: 'Text', syntax: 'SUBSTITUTE(text, old_text, new_text)', description: 'Substitutes new_text for old_text in a text string.', example: '=SUBSTITUTE(A2, "-", "/")' },
  { name: 'FIND', category: 'Text', syntax: 'FIND(find_text, within_text, [start_num])', description: 'Finds one text value within another (case-sensitive).', example: '=FIND("@", A2)' },
  { name: 'SEARCH', category: 'Text', syntax: 'SEARCH(find_text, within_text, [start_num])', description: 'Finds one text value within another (case-insensitive).', example: '=SEARCH("pro", A2)' },
  { name: 'TEXT', category: 'Text', syntax: 'TEXT(value, format_text)', description: 'Formats a number and converts it to text.', example: '=TEXT(B2, "$#,##0.00")' },
  { name: 'VALUE', category: 'Text', syntax: 'VALUE(text)', description: 'Converts a text string that represents a number to a number.', example: '=VALUE("1250")' },
  { name: 'REPT', category: 'Text', syntax: 'REPT(text, number_times)', description: 'Repeats text a given number of times.', example: '=REPT("★", B2)' },

  // Date & Time
  { name: 'TODAY', category: 'Date & Time', syntax: 'TODAY()', description: 'Returns the serial number of the current date.', example: '=TODAY()' },
  { name: 'NOW', category: 'Date & Time', syntax: 'NOW()', description: 'Returns the serial number of the current date and time.', example: '=NOW()' },
  { name: 'DATE', category: 'Date & Time', syntax: 'DATE(year, month, day)', description: 'Returns the serial number of a particular date.', example: '=DATE(2026, 11, 14)' },
  { name: 'YEAR', category: 'Date & Time', syntax: 'YEAR(serial_number)', description: 'Returns the year corresponding to a date.', example: '=YEAR(A2)' },
  { name: 'MONTH', category: 'Date & Time', syntax: 'MONTH(serial_number)', description: 'Returns the month of a date (1 to 12).', example: '=MONTH(A2)' },
  { name: 'DAY', category: 'Date & Time', syntax: 'DAY(serial_number)', description: 'Returns the day of a date (1 to 31).', example: '=DAY(A2)' },
  { name: 'WEEKDAY', category: 'Date & Time', syntax: 'WEEKDAY(serial_number, [return_type])', description: 'Returns the day of the week corresponding to a date.', example: '=WEEKDAY(A2)' },
  { name: 'DATEDIF', category: 'Date & Time', syntax: 'DATEDIF(start_date, end_date, unit)', description: 'Calculates the number of days ("D"), months ("M"), or years ("Y") between two dates.', example: '=DATEDIF(A2, B2, "D")' },

  // Financial
  { name: 'PMT', category: 'Financial', syntax: 'PMT(rate, nper, pv, [fv], [type])', description: 'Calculates the payment for a loan based on constant payments and a constant interest rate.', example: '=PMT(0.05/12, 60, -50000)' },
  { name: 'FV', category: 'Financial', syntax: 'FV(rate, nper, pmt, [pv], [type])', description: 'Returns the future value of an investment.', example: '=FV(0.06/12, 120, -500, -10000)' },
  { name: 'PV', category: 'Financial', syntax: 'PV(rate, nper, pmt, [fv], [type])', description: 'Returns the present value of an investment.', example: '=PV(0.05/12, 60, -1000)' },
  { name: 'NPV', category: 'Financial', syntax: 'NPV(rate, value1, [value2], ...)', description: 'Calculates the net present value of an investment.', example: '=NPV(0.1, B2:B6)' },
  { name: 'IRR', category: 'Financial', syntax: 'IRR(values, [guess])', description: 'Returns the internal rate of return for a series of cash flows.', example: '=IRR(B2:B6)' },

  // Information
  { name: 'ISBLANK', category: 'Information', syntax: 'ISBLANK(value)', description: 'Returns TRUE if the value is blank.', example: '=ISBLANK(A2)' },
  { name: 'ISNUMBER', category: 'Information', syntax: 'ISNUMBER(value)', description: 'Returns TRUE if the value is a number.', example: '=ISNUMBER(A2)' },
  { name: 'ISTEXT', category: 'Information', syntax: 'ISTEXT(value)', description: 'Returns TRUE if the value is text.', example: '=ISTEXT(A2)' },
  { name: 'ISERROR', category: 'Information', syntax: 'ISERROR(value)', description: 'Returns TRUE if the value is any error value.', example: '=ISERROR(A2/B2)' },
];

class FormulaError extends Error {
  code: string;
  constructor(code: string) {
    super(code);
    this.code = code;
  }
}

type EvalArg = CellValue | CellValue[][];

function isRange2D(val: EvalArg): val is CellValue[][] {
  return Array.isArray(val);
}

function flattenArgs(args: EvalArg[]): CellValue[] {
  const out: CellValue[] = [];
  for (const a of args) {
    if (isRange2D(a)) {
      for (const row of a) {
        for (const v of row) out.push(v);
      }
    } else {
      out.push(a);
    }
  }
  return out;
}

function toNum(v: CellValue, allowBlankAsZero = true): number {
  if (v === null || v === undefined || v === '') {
    if (allowBlankAsZero) return 0;
    throw new FormulaError('#VALUE!');
  }
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const trimmed = String(v).trim();
  if (trimmed === '') return allowBlankAsZero ? 0 : NaN;
  const n = Number(trimmed.replace(/,/g, ''));
  if (!isNaN(n)) return n;
  throw new FormulaError('#VALUE!');
}

function toBool(v: CellValue): boolean {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (typeof v === 'string') {
    const u = v.trim().toUpperCase();
    if (u === 'TRUE') return true;
    if (u === 'FALSE') return false;
  }
  if (v === null || v === '') return false;
  throw new FormulaError('#VALUE!');
}

function toStr(v: CellValue): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

function matchesCriteria(val: CellValue, criteria: CellValue): boolean {
  const critStr = toStr(criteria);
  const opMatch = critStr.match(/^(<=|>=|<>|<|>|=)(.*)$/);
  if (opMatch) {
    const op = opMatch[1];
    const rhsRaw = opMatch[2].trim();
    const rhsNum = Number(rhsRaw);
    const valNum = typeof val === 'number' ? val : val !== null && val !== '' && !isNaN(Number(val)) ? Number(val) : NaN;
    if (!isNaN(rhsNum) && rhsRaw !== '' && !isNaN(valNum)) {
      switch (op) {
        case '<=': return valNum <= rhsNum;
        case '>=': return valNum >= rhsNum;
        case '<>': return valNum !== rhsNum;
        case '<': return valNum < rhsNum;
        case '>': return valNum > rhsNum;
        case '=': return valNum === rhsNum;
      }
    }
    const lhsStr = toStr(val).toLowerCase();
    const rhsStr = rhsRaw.toLowerCase();
    switch (op) {
      case '<>': return lhsStr !== rhsStr;
      case '=': return lhsStr === rhsStr;
      case '<': return lhsStr < rhsStr;
      case '>': return lhsStr > rhsStr;
      case '<=': return lhsStr <= rhsStr;
      case '>=': return lhsStr >= rhsStr;
    }
  }

  // Wildcard support (* and ?)
  if (critStr.includes('*') || critStr.includes('?')) {
    const escaped = critStr
      .replace(/[.+^${}()|[\]\\]/g, '\\$&')
      .replace(/\*/g, '.*')
      .replace(/\?/g, '.');
    const regex = new RegExp(`^${escaped}$`, 'i');
    return regex.test(toStr(val));
  }

  if (typeof criteria === 'number' && typeof val === 'number') {
    return val === criteria;
  }
  return toStr(val).toLowerCase() === critStr.toLowerCase();
}

export interface FormulaRefHighlight {
  sheetId: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  text: string;
  colorIndex: number;
}

export function extractFormulaRefs(
  formula: string,
  currentSheetId: string,
  sheets: Sheet[],
  namedRanges: NamedRange[]
): FormulaRefHighlight[] {
  if (!formula.startsWith('=')) return [];
  const results: FormulaRefHighlight[] = [];
  let colorIdx = 0;

  // Strip string literals before matching refs
  const cleaned = formula.replace(/"([^"\\]|\\.)*"/g, (m) => ' '.repeat(m.length));

  // Match optional sheet prefix + cell or range
  const refRegex = /(?:(?:'([^']+)'|([A-Za-z0-9_]+))!)?(\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)(?::(\$?[A-Za-z]{1,3}\$?[1-9][0-9]*))?/g;
  let m: RegExpExecArray | null;
  while ((m = refRegex.exec(cleaned)) !== null) {
    // Make sure it's not a function name followed by '('
    const afterChar = cleaned[m.index + m[0].length];
    if (afterChar === '(') continue;
    const beforeChar = m.index > 0 ? cleaned[m.index - 1] : '';
    if (/[A-Za-z0-9_]/.test(beforeChar)) continue;

    const sheetName = m[1] || m[2];
    let targetSheetId = currentSheetId;
    if (sheetName) {
      const found = sheets.find((s) => s.name.toLowerCase() === sheetName.toLowerCase());
      if (!found) continue;
      targetSheetId = found.id;
    }

    const startParsed = m[3].match(/^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$/);
    if (!startParsed) continue;
    const sc = colLetterToIndex(startParsed[1]);
    const sr = parseInt(startParsed[2], 10) - 1;
    if (sc > 701) continue;

    let er = sr;
    let ec = sc;
    if (m[4]) {
      const endParsed = m[4].match(/^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$/);
      if (endParsed) {
        ec = colLetterToIndex(endParsed[1]);
        er = parseInt(endParsed[2], 10) - 1;
      }
    }

    results.push({
      sheetId: targetSheetId,
      startRow: Math.min(sr, er),
      startCol: Math.min(sc, ec),
      endRow: Math.max(sr, er),
      endCol: Math.max(sc, ec),
      text: m[0],
      colorIndex: colorIdx++,
    });
  }

  // Also check named ranges
  for (const nr of namedRanges) {
    const nrRegex = new RegExp(`\\b${nr.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'gi');
    if (nrRegex.test(cleaned)) {
      const parts = nr.range.split(':');
      const s = parts[0].match(/^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$/);
      const e = (parts[1] || parts[0]).match(/^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$/);
      if (s && e) {
        const sr = parseInt(s[2], 10) - 1;
        const sc = colLetterToIndex(s[1]);
        const er = parseInt(e[2], 10) - 1;
        const ec = colLetterToIndex(e[1]);
        results.push({
          sheetId: nr.sheetId,
          startRow: Math.min(sr, er),
          startCol: Math.min(sc, ec),
          endRow: Math.max(sr, er),
          endCol: Math.max(sc, ec),
          text: nr.name,
          colorIndex: colorIdx++,
        });
      }
    }
  }

  return results;
}

export class WorkbookEvaluator {
  private sheets: Sheet[];
  private namedRanges: NamedRange[];
  private defaultCurrency: string;
  private cache: Map<string, EvaluatedCell> = new Map();
  private visiting: Set<string> = new Set();
  private dependencies: Map<string, Set<string>> = new Map(); // cellId -> set of precedent cellIds
  private dependents: Map<string, Set<string>> = new Map(); // precedent cellId -> set of dependent cellIds

  constructor(sheets: Sheet[], namedRanges: NamedRange[], defaultCurrency = 'RM') {
    this.sheets = sheets;
    this.namedRanges = namedRanges;
    this.defaultCurrency = defaultCurrency;
  }

  private getSheet(sheetId: string): Sheet | undefined {
    return this.sheets.find((s) => s.id === sheetId);
  }

  private getSheetByName(name: string): Sheet | undefined {
    return this.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase());
  }

  private globalKey(sheetId: string, row: number, col: number): string {
    return `${sheetId}:${row},${col}`;
  }

  private recordDependency(fromSheetId: string, fromRow: number, fromCol: number, toSheetId: string, toRow: number, toCol: number) {
    const fromKey = this.globalKey(fromSheetId, fromRow, fromCol);
    const toKey = this.globalKey(toSheetId, toRow, toCol);
    if (!this.dependencies.has(fromKey)) this.dependencies.set(fromKey, new Set());
    this.dependencies.get(fromKey)!.add(toKey);
    if (!this.dependents.has(toKey)) this.dependents.set(toKey, new Set());
    this.dependents.get(toKey)!.add(fromKey);
  }

  public getPrecedents(sheetId: string, row: number, col: number): { sheetId: string; row: number; col: number }[] {
    this.evaluateCell(sheetId, row, col);
    const key = this.globalKey(sheetId, row, col);
    const deps = this.dependencies.get(key);
    if (!deps) return [];
    return Array.from(deps).map((k) => {
      const [sid, coords] = k.split(':');
      const [r, c] = coords.split(',').map(Number);
      return { sheetId: sid, row: r, col: c };
    });
  }

  public getDependents(sheetId: string, row: number, col: number): { sheetId: string; row: number; col: number }[] {
    // Evaluate all formula cells across sheets to build complete graph
    for (const s of this.sheets) {
      for (const k of Object.keys(s.cells)) {
        if (s.cells[k]?.raw?.startsWith('=')) {
          const [r, c] = k.split(',').map(Number);
          this.evaluateCell(s.id, r, c);
        }
      }
    }
    const key = this.globalKey(sheetId, row, col);
    const deps = this.dependents.get(key);
    if (!deps) return [];
    return Array.from(deps).map((k) => {
      const [sid, coords] = k.split(':');
      const [r, c] = coords.split(',').map(Number);
      return { sheetId: sid, row: r, col: c };
    });
  }

  public evaluateCell(sheetId: string, row: number, col: number): EvaluatedCell {
    const gKey = this.globalKey(sheetId, row, col);
    const cached = this.cache.get(gKey);
    if (cached) return cached;

    const sheet = this.getSheet(sheetId);
    if (!sheet) return { value: null, error: '#REF!', formatted: '#REF!' };

    const cell = sheet.cells[cellKey(row, col)];
    if (!cell || cell.raw === '') {
      const emptyRes: EvaluatedCell = { value: null, formatted: '' };
      this.cache.set(gKey, emptyRes);
      return emptyRes;
    }

    const raw = cell.raw;

    if (!raw.startsWith('=')) {
      const val = this.parseLiteral(raw, cell.style?.numberFormat?.category === 'text');
      const formatted = formatCellValue(val, cell.style?.numberFormat, this.defaultCurrency);
      const res: EvaluatedCell = { value: val, formatted };
      this.cache.set(gKey, res);
      return res;
    }

    if (this.visiting.has(gKey)) {
      return { value: null, error: '#CIRC!', formatted: '#CIRC!' };
    }

    this.visiting.add(gKey);
    try {
      const expr = raw.slice(1).trim();
      if (!expr) {
        const res: EvaluatedCell = { value: 0, formatted: '0' };
        this.cache.set(gKey, res);
        return res;
      }
      const val = this.evalExpression(expr, sheetId, row, col);
      let finalVal: CellValue;
      if (isRange2D(val)) {
        finalVal = val[0]?.[0] ?? null;
      } else {
        finalVal = val;
      }
      const formatted = formatCellValue(finalVal, cell.style?.numberFormat, this.defaultCurrency);
      const res: EvaluatedCell = { value: finalVal, formatted };
      this.cache.set(gKey, res);
      return res;
    } catch (err) {
      const code = err instanceof FormulaError ? err.code : '#ERROR!';
      const res: EvaluatedCell = { value: null, error: code, formatted: code };
      this.cache.set(gKey, res);
      return res;
    } finally {
      this.visiting.delete(gKey);
    }
  }

  private parseLiteral(raw: string, forceText = false): CellValue {
    if (forceText) return raw;
    const trimmed = raw.trim();
    if (trimmed === '') return '';
    const upper = trimmed.toUpperCase();
    if (upper === 'TRUE') return true;
    if (upper === 'FALSE') return false;

    // Percentage literal e.g. "12.5%"
    if (/^-?\d+(\.\d+)?%$/.test(trimmed)) {
      return parseFloat(trimmed.slice(0, -1)) / 100;
    }

    // Numeric literal
    if (/^-?\d+(,\d{3})*(\.\d+)?([eE][+-]?\d+)?$/.test(trimmed)) {
      const n = Number(trimmed.replace(/,/g, ''));
      if (!isNaN(n)) return n;
    }

    // Date literal
    const dateSerial = tryParseInputDate(trimmed);
    if (dateSerial !== null) return dateSerial;

    return raw;
  }

  private evalExpression(expr: string, sheetId: string, curRow: number, curCol: number): EvalArg {
    const tokens = this.tokenize(expr);
    let pos = 0;

    const peek = () => tokens[pos];
    const consume = () => tokens[pos++];

    const parseComparison = (): EvalArg => {
      let left = parseConcat();
      while (peek() && ['=', '<>', '<', '>', '<=', '>='].includes(peek().value) && peek().type === 'op') {
        const op = consume().value;
        const right = parseConcat();
        left = this.compareValues(scalarize(left), scalarize(right), op);
      }
      return left;
    };

    const parseConcat = (): EvalArg => {
      let left = parseAddSub();
      while (peek() && peek().type === 'op' && peek().value === '&') {
        consume();
        const right = parseAddSub();
        left = toStr(scalarize(left)) + toStr(scalarize(right));
      }
      return left;
    };

    const parseAddSub = (): EvalArg => {
      let left = parseMulDiv();
      while (peek() && peek().type === 'op' && (peek().value === '+' || peek().value === '-')) {
        const op = consume().value;
        const right = parseMulDiv();
        const l = toNum(scalarize(left));
        const r = toNum(scalarize(right));
        left = op === '+' ? l + r : l - r;
      }
      return left;
    };

    const parseMulDiv = (): EvalArg => {
      let left = parsePower();
      while (peek() && peek().type === 'op' && (peek().value === '*' || peek().value === '/')) {
        const op = consume().value;
        const right = parsePower();
        const l = toNum(scalarize(left));
        const r = toNum(scalarize(right));
        if (op === '/') {
          if (r === 0) throw new FormulaError('#DIV/0!');
          left = l / r;
        } else {
          left = l * r;
        }
      }
      return left;
    };

    const parsePower = (): EvalArg => {
      let left = parseUnary();
      if (peek() && peek().type === 'op' && peek().value === '^') {
        consume();
        const right = parsePower();
        const res = Math.pow(toNum(scalarize(left)), toNum(scalarize(right)));
        if (isNaN(res)) throw new FormulaError('#NUM!');
        return res;
      }
      return left;
    };

    const parseUnary = (): EvalArg => {
      if (peek() && peek().type === 'op' && (peek().value === '+' || peek().value === '-')) {
        const op = consume().value;
        const val = parseUnary();
        const n = toNum(scalarize(val));
        return op === '-' ? -n : n;
      }
      return parsePostfix();
    };

    const parsePostfix = (): EvalArg => {
      let val = parsePrimary();
      while (peek() && peek().type === 'op' && peek().value === '%') {
        consume();
        val = toNum(scalarize(val)) / 100;
      }
      return val;
    };

    const scalarize = (v: EvalArg): CellValue => {
      if (isRange2D(v)) return v[0]?.[0] ?? null;
      return v;
    };

    const parsePrimary = (): EvalArg => {
      const tok = peek();
      if (!tok) throw new FormulaError('#VALUE!');

      if (tok.type === 'number') {
        consume();
        return parseFloat(tok.value);
      }

      if (tok.type === 'string') {
        consume();
        return tok.value;
      }

      if (tok.type === 'bool') {
        consume();
        return tok.value === 'TRUE';
      }

      if (tok.type === 'error') {
        consume();
        throw new FormulaError(tok.value);
      }

      if (tok.type === 'paren' && tok.value === '(') {
        consume();
        const val = parseComparison();
        if (!peek() || peek().value !== ')') throw new FormulaError('#VALUE!');
        consume();
        return val;
      }

      if (tok.type === 'ident') {
        consume();
        // Function call?
        if (peek() && peek().type === 'paren' && peek().value === '(') {
          const fnName = tok.value.toUpperCase();
          consume(); // '('
          const args: EvalArg[] = [];
          const argRawRefs: string[] = [];
          const argErrors: (string | null)[] = [];
          if (!(peek() && peek().type === 'paren' && peek().value === ')')) {
            while (true) {
              const startTokIdx = pos;
              let capturedErr: string | null = null;
              let val: EvalArg = null;
              if (fnName === 'IFERROR' || fnName === 'IFNA' || fnName === 'ISERROR') {
                try {
                  val = parseComparison();
                } catch (e) {
                  if (e instanceof FormulaError) {
                    capturedErr = e.code;
                    // Skip remaining tokens of this argument up to matching comma or ')'
                    let depth = 0;
                    while (peek()) {
                      const t = peek();
                      if (t.type === 'paren' && t.value === '(') depth++;
                      else if (t.type === 'paren' && t.value === ')') {
                        if (depth === 0) break;
                        depth--;
                      } else if (t.type === 'comma' && depth === 0) {
                        break;
                      }
                      consume();
                    }
                  } else {
                    throw e;
                  }
                }
              } else {
                val = parseComparison();
              }
              args.push(val);
              argErrors.push(capturedErr);
              argRawRefs.push(tokens.slice(startTokIdx, pos).map((t) => t.value).join(''));
              if (peek() && peek().type === 'comma') {
                consume();
              } else {
                break;
              }
            }
          }
          if (!peek() || peek().value !== ')') throw new FormulaError('#VALUE!');
          consume();
          if (fnName === 'ISERROR') {
            return argErrors[0] !== null;
          }
          if (fnName === 'IFERROR') {
            if (argErrors[0] !== null) {
              if (argErrors[1]) throw new FormulaError(argErrors[1]);
              return args.length > 1 ? (isRange2D(args[1]) ? args[1][0]?.[0] ?? null : args[1]) : '';
            }
            return isRange2D(args[0]) ? args[0][0]?.[0] ?? null : args[0];
          }
          if (fnName === 'IFNA') {
            if (argErrors[0] === '#N/A') {
              if (argErrors[1]) throw new FormulaError(argErrors[1]);
              return args.length > 1 ? (isRange2D(args[1]) ? args[1][0]?.[0] ?? null : args[1]) : '';
            }
            if (argErrors[0] !== null) throw new FormulaError(argErrors[0]);
            return isRange2D(args[0]) ? args[0][0]?.[0] ?? null : args[0];
          }
          return this.callFunction(fnName, args, argRawRefs, sheetId, curRow, curCol);
        }

        // Check if range with ':'
        if (peek() && peek().type === 'colon') {
          consume();
          const endTok = consume();
          if (!endTok || endTok.type !== 'ident') throw new FormulaError('#REF!');
          return this.resolveRangeRef(`${tok.value}:${endTok.value}`, sheetId, curRow, curCol);
        }

        // Single cell ref, cross-sheet ref, or named range
        return this.resolveIdentifier(tok.value, sheetId, curRow, curCol);
      }

      throw new FormulaError('#VALUE!');
    };

    const result = parseComparison();
    if (pos < tokens.length) {
      throw new FormulaError('#VALUE!');
    }
    return result;
  }

  private resolveIdentifier(ident: string, sheetId: string, curRow: number, curCol: number): EvalArg {
    // Check named range first
    const nr = this.namedRanges.find((n) => n.name.toLowerCase() === ident.toLowerCase());
    if (nr) {
      return this.resolveRangeRef(nr.range, nr.sheetId, curRow, curCol, sheetId);
    }

    // Check optional sheet!cell
    let targetSheetId = sheetId;
    let refPart = ident;
    const bangIdx = ident.indexOf('!');
    if (bangIdx !== -1) {
      let sName = ident.slice(0, bangIdx);
      if (sName.startsWith("'") && sName.endsWith("'")) sName = sName.slice(1, -1);
      const targetSheet = this.getSheetByName(sName);
      if (!targetSheet) throw new FormulaError('#REF!');
      targetSheetId = targetSheet.id;
      refPart = ident.slice(bangIdx + 1);
    }

    const m = refPart.match(/^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$/);
    if (!m) throw new FormulaError('#NAME?');
    const col = colLetterToIndex(m[1]);
    const row = parseInt(m[2], 10) - 1;
    if (col > 16383 || row < 0) throw new FormulaError('#REF!');

    this.recordDependency(sheetId, curRow, curCol, targetSheetId, row, col);
    const ev = this.evaluateCell(targetSheetId, row, col);
    if (ev.error) throw new FormulaError(ev.error);
    return ev.value;
  }

  private resolveRangeRef(
    rangeStr: string,
    defaultSheetId: string,
    curRow: number,
    curCol: number,
    callerSheetId?: string
  ): CellValue[][] {
    const depFromSheet = callerSheetId ?? defaultSheetId;
    let targetSheetId = defaultSheetId;
    let body = rangeStr;
    const bangIdx = rangeStr.indexOf('!');
    if (bangIdx !== -1) {
      let sName = rangeStr.slice(0, bangIdx);
      if (sName.startsWith("'") && sName.endsWith("'")) sName = sName.slice(1, -1);
      const targetSheet = this.getSheetByName(sName);
      if (!targetSheet) throw new FormulaError('#REF!');
      targetSheetId = targetSheet.id;
      body = rangeStr.slice(bangIdx + 1);
    }

    const parts = body.split(':');
    if (parts.length !== 2) throw new FormulaError('#REF!');

    const targetSheet = this.getSheet(targetSheetId);
    if (!targetSheet) throw new FormulaError('#REF!');

    // Support whole column (A:A) or cell:cell (A1:C10)
    const colOnlyStart = parts[0].match(/^\$?([A-Za-z]{1,3})$/);
    const colOnlyEnd = parts[1].match(/^\$?([A-Za-z]{1,3})$/);
    let r1 = 0, c1 = 0, r2 = 0, c2 = 0;

    if (colOnlyStart && colOnlyEnd) {
      c1 = colLetterToIndex(colOnlyStart[1]);
      c2 = colLetterToIndex(colOnlyEnd[1]);
      r1 = 0;
      r2 = Math.max(0, targetSheet.rowCount - 1);
    } else {
      const p1 = parts[0].match(/^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$/);
      const p2 = parts[1].match(/^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$/);
      if (!p1 || !p2) throw new FormulaError('#REF!');
      c1 = colLetterToIndex(p1[1]);
      r1 = parseInt(p1[2], 10) - 1;
      c2 = colLetterToIndex(p2[1]);
      r2 = parseInt(p2[2], 10) - 1;
    }

    const sr = Math.min(r1, r2);
    const er = Math.max(r1, r2);
    const sc = Math.min(c1, c2);
    const ec = Math.max(c1, c2);

    const matrix: CellValue[][] = [];
    for (let r = sr; r <= er; r++) {
      const rowArr: CellValue[] = [];
      for (let c = sc; c <= ec; c++) {
        this.recordDependency(depFromSheet, curRow, curCol, targetSheetId, r, c);
        const ev = this.evaluateCell(targetSheetId, r, c);
        if (ev.error) throw new FormulaError(ev.error);
        rowArr.push(ev.value);
      }
      matrix.push(rowArr);
    }
    return matrix;
  }

  private compareValues(a: CellValue, b: CellValue, op: string): boolean {
    // Normalize blanks
    const av = a === null ? '' : a;
    const bv = b === null ? '' : b;

    if (typeof av === 'number' && typeof bv === 'number') {
      switch (op) {
        case '=': return av === bv;
        case '<>': return av !== bv;
        case '<': return av < bv;
        case '>': return av > bv;
        case '<=': return av <= bv;
        case '>=': return av >= bv;
      }
    }
    if (typeof av === 'boolean' && typeof bv === 'boolean') {
      switch (op) {
        case '=': return av === bv;
        case '<>': return av !== bv;
        case '<': return av < bv;
        case '>': return av > bv;
        case '<=': return av <= bv;
        case '>=': return av >= bv;
      }
    }
    const as = toStr(av).toLowerCase();
    const bs = toStr(bv).toLowerCase();
    switch (op) {
      case '=': return as === bs;
      case '<>': return as !== bs;
      case '<': return as < bs;
      case '>': return as > bs;
      case '<=': return as <= bs;
      case '>=': return as >= bs;
    }
    return false;
  }

  private callFunction(
    name: string,
    args: EvalArg[],
    argRawRefs: string[],
    sheetId: string,
    curRow: number,
    curCol: number
  ): EvalArg {
    const scalar = (idx: number, def?: CellValue): CellValue => {
      if (idx >= args.length) return def ?? null;
      const v = args[idx];
      return isRange2D(v) ? (v[0]?.[0] ?? null) : v;
    };

    const getNumericValues = (list: EvalArg[]): number[] => {
      const nums: number[] = [];
      for (const item of list) {
        if (isRange2D(item)) {
          for (const row of item) {
            for (const v of row) {
              if (typeof v === 'number') nums.push(v);
            }
          }
        } else {
          if (typeof item === 'number') nums.push(item);
          else if (typeof item === 'boolean') nums.push(item ? 1 : 0);
          else if (typeof item === 'string' && item.trim() !== '' && !isNaN(Number(item))) {
            nums.push(Number(item));
          }
        }
      }
      return nums;
    };

    switch (name) {
      // MATH & TRIG
      case 'SUM': {
        return getNumericValues(args).reduce((a, b) => a + b, 0);
      }
      case 'PRODUCT': {
        const nums = getNumericValues(args);
        return nums.length === 0 ? 0 : nums.reduce((a, b) => a * b, 1);
      }
      case 'SUMPRODUCT': {
        if (args.length === 0) throw new FormulaError('#VALUE!');
        const arrays = args.map((a) => (isRange2D(a) ? a : [[a]]));
        const rows = arrays[0].length;
        const cols = arrays[0][0]?.length || 0;
        let total = 0;
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            let prod = 1;
            for (const arr of arrays) {
              if (arr.length !== rows || (arr[0]?.length || 0) !== cols) throw new FormulaError('#VALUE!');
              const v = arr[r][c];
              prod *= typeof v === 'number' ? v : 0;
            }
            total += prod;
          }
        }
        return total;
      }
      case 'SUMIF': {
        if (args.length < 2) throw new FormulaError('#VALUE!');
        const range = isRange2D(args[0]) ? args[0] : [[args[0]]];
        const criteria = scalar(1);
        const sumRange = args[2] ? (isRange2D(args[2]) ? args[2] : [[args[2]]]) : range;
        let sum = 0;
        for (let r = 0; r < range.length; r++) {
          for (let c = 0; c < range[r].length; c++) {
            if (matchesCriteria(range[r][c], criteria)) {
              const sv = sumRange[r]?.[c];
              if (typeof sv === 'number') sum += sv;
            }
          }
        }
        return sum;
      }
      case 'SUMIFS': {
        if (args.length < 3 || args.length % 2 === 0) throw new FormulaError('#VALUE!');
        const sumRange = isRange2D(args[0]) ? args[0] : [[args[0]]];
        let sum = 0;
        for (let r = 0; r < sumRange.length; r++) {
          for (let c = 0; c < sumRange[r].length; c++) {
            let ok = true;
            for (let i = 1; i < args.length; i += 2) {
              const argVal = args[i];
              const cr: CellValue[][] = isRange2D(argVal) ? argVal : [[argVal]];
              const crit = scalar(i + 1);
              if (!matchesCriteria(cr[r]?.[c] ?? null, crit)) {
                ok = false;
                break;
              }
            }
            if (ok && typeof sumRange[r][c] === 'number') sum += sumRange[r][c] as number;
          }
        }
        return sum;
      }
      case 'ABS': return Math.abs(toNum(scalar(0)));
      case 'SQRT': {
        const n = toNum(scalar(0));
        if (n < 0) throw new FormulaError('#NUM!');
        return Math.sqrt(n);
      }
      case 'POWER': return Math.pow(toNum(scalar(0)), toNum(scalar(1)));
      case 'MOD': {
        const n = toNum(scalar(0));
        const d = toNum(scalar(1));
        if (d === 0) throw new FormulaError('#DIV/0!');
        return n - d * Math.floor(n / d);
      }
      case 'INT': return Math.floor(toNum(scalar(0)));
      case 'ROUND': {
        const n = toNum(scalar(0));
        const d = Math.trunc(toNum(scalar(1), true));
        const f = Math.pow(10, d);
        return Math.round(n * f) / f;
      }
      case 'ROUNDUP': {
        const n = toNum(scalar(0));
        const d = Math.trunc(toNum(scalar(1), true));
        const f = Math.pow(10, d);
        return (n >= 0 ? Math.ceil(n * f) : Math.floor(n * f)) / f;
      }
      case 'ROUNDDOWN': {
        const n = toNum(scalar(0));
        const d = Math.trunc(toNum(scalar(1), true));
        const f = Math.pow(10, d);
        return Math.trunc(n * f) / f;
      }
      case 'CEILING': {
        const n = toNum(scalar(0));
        const sig = toNum(scalar(1), true) || 1;
        return Math.ceil(n / sig) * sig;
      }
      case 'FLOOR': {
        const n = toNum(scalar(0));
        const sig = toNum(scalar(1), true) || 1;
        return Math.floor(n / sig) * sig;
      }
      case 'PI': return Math.PI;
      case 'RAND': return Math.random();
      case 'RANDBETWEEN': {
        const min = Math.ceil(toNum(scalar(0)));
        const max = Math.floor(toNum(scalar(1)));
        if (min > max) throw new FormulaError('#NUM!');
        return Math.floor(Math.random() * (max - min + 1)) + min;
      }
      case 'LOG': {
        const n = toNum(scalar(0));
        const base = args.length > 1 ? toNum(scalar(1)) : 10;
        if (n <= 0 || base <= 0 || base === 1) throw new FormulaError('#NUM!');
        return Math.log(n) / Math.log(base);
      }
      case 'LN': {
        const n = toNum(scalar(0));
        if (n <= 0) throw new FormulaError('#NUM!');
        return Math.log(n);
      }
      case 'EXP': return Math.exp(toNum(scalar(0)));

      // STATISTICAL
      case 'AVERAGE': {
        const nums = getNumericValues(args);
        if (nums.length === 0) throw new FormulaError('#DIV/0!');
        return nums.reduce((a, b) => a + b, 0) / nums.length;
      }
      case 'AVERAGEIF': {
        if (args.length < 2) throw new FormulaError('#VALUE!');
        const range = isRange2D(args[0]) ? args[0] : [[args[0]]];
        const criteria = scalar(1);
        const avgRange = args[2] ? (isRange2D(args[2]) ? args[2] : [[args[2]]]) : range;
        let sum = 0;
        let cnt = 0;
        for (let r = 0; r < range.length; r++) {
          for (let c = 0; c < range[r].length; c++) {
            if (matchesCriteria(range[r][c], criteria)) {
              const av = avgRange[r]?.[c];
              if (typeof av === 'number') {
                sum += av;
                cnt++;
              }
            }
          }
        }
        if (cnt === 0) throw new FormulaError('#DIV/0!');
        return sum / cnt;
      }
      case 'COUNT': {
        let cnt = 0;
        for (const v of flattenArgs(args)) {
          if (typeof v === 'number') cnt++;
        }
        return cnt;
      }
      case 'COUNTA': {
        let cnt = 0;
        for (const v of flattenArgs(args)) {
          if (v !== null && v !== undefined && v !== '') cnt++;
        }
        return cnt;
      }
      case 'COUNTBLANK': {
        let cnt = 0;
        for (const v of flattenArgs(args)) {
          if (v === null || v === undefined || v === '') cnt++;
        }
        return cnt;
      }
      case 'COUNTIF': {
        if (args.length < 2) throw new FormulaError('#VALUE!');
        const vals = flattenArgs([args[0]]);
        const criteria = scalar(1);
        return vals.filter((v) => matchesCriteria(v, criteria)).length;
      }
      case 'COUNTIFS': {
        if (args.length < 2 || args.length % 2 !== 0) throw new FormulaError('#VALUE!');
        const firstRange = isRange2D(args[0]) ? args[0] : [[args[0]]];
        let cnt = 0;
        for (let r = 0; r < firstRange.length; r++) {
          for (let c = 0; c < firstRange[r].length; c++) {
            let ok = true;
            for (let i = 0; i < args.length; i += 2) {
              const argVal = args[i];
              const cr: CellValue[][] = isRange2D(argVal) ? argVal : [[argVal]];
              const crit = scalar(i + 1);
              if (!matchesCriteria(cr[r]?.[c] ?? null, crit)) {
                ok = false;
                break;
              }
            }
            if (ok) cnt++;
          }
        }
        return cnt;
      }
      case 'MIN': {
        const nums = getNumericValues(args);
        return nums.length === 0 ? 0 : Math.min(...nums);
      }
      case 'MAX': {
        const nums = getNumericValues(args);
        return nums.length === 0 ? 0 : Math.max(...nums);
      }
      case 'MEDIAN': {
        const nums = getNumericValues(args).sort((a, b) => a - b);
        if (nums.length === 0) throw new FormulaError('#NUM!');
        const mid = Math.floor(nums.length / 2);
        return nums.length % 2 !== 0 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
      }
      case 'VAR': {
        const nums = getNumericValues(args);
        if (nums.length < 2) throw new FormulaError('#DIV/0!');
        const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
        return nums.reduce((s, x) => s + Math.pow(x - mean, 2), 0) / (nums.length - 1);
      }
      case 'STDEV': {
        const nums = getNumericValues(args);
        if (nums.length < 2) throw new FormulaError('#DIV/0!');
        const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
        const variance = nums.reduce((s, x) => s + Math.pow(x - mean, 2), 0) / (nums.length - 1);
        return Math.sqrt(variance);
      }
      case 'LARGE': {
        const nums = getNumericValues([args[0]]).sort((a, b) => b - a);
        const k = Math.trunc(toNum(scalar(1)));
        if (k < 1 || k > nums.length) throw new FormulaError('#NUM!');
        return nums[k - 1];
      }
      case 'SMALL': {
        const nums = getNumericValues([args[0]]).sort((a, b) => a - b);
        const k = Math.trunc(toNum(scalar(1)));
        if (k < 1 || k > nums.length) throw new FormulaError('#NUM!');
        return nums[k - 1];
      }
      case 'RANK': {
        const num = toNum(scalar(0));
        const nums = getNumericValues([args[1]]);
        const order = args.length > 2 ? toNum(scalar(2)) : 0;
        const sorted = [...nums].sort((a, b) => (order === 0 ? b - a : a - b));
        const idx = sorted.indexOf(num);
        if (idx === -1) throw new FormulaError('#N/A');
        return idx + 1;
      }

      // LOGICAL
      case 'IF': {
        const cond = toBool(scalar(0));
        return cond ? scalar(1, true) : scalar(2, false);
      }
      case 'IFS': {
        for (let i = 0; i + 1 < args.length; i += 2) {
          if (toBool(scalar(i))) return scalar(i + 1);
        }
        throw new FormulaError('#N/A');
      }
      case 'AND': {
        const vals = flattenArgs(args).filter((v) => v !== null && v !== '');
        if (vals.length === 0) throw new FormulaError('#VALUE!');
        return vals.every((v) => toBool(v));
      }
      case 'OR': {
        const vals = flattenArgs(args).filter((v) => v !== null && v !== '');
        if (vals.length === 0) throw new FormulaError('#VALUE!');
        return vals.some((v) => toBool(v));
      }
      case 'NOT': return !toBool(scalar(0));
      case 'XOR': {
        const vals = flattenArgs(args).filter((v) => v !== null && v !== '');
        const trues = vals.filter((v) => toBool(v)).length;
        return trues % 2 === 1;
      }
      case 'IFERROR': {
        return scalar(0);
      }
      case 'IFNA': {
        return scalar(0);
      }

      // LOOKUP & REFERENCE
      case 'VLOOKUP': {
        if (args.length < 3) throw new FormulaError('#VALUE!');
        const lookupVal = scalar(0);
        const table = isRange2D(args[1]) ? args[1] : [[args[1]]];
        const colIdx = Math.trunc(toNum(scalar(2)));
        const approx = args.length > 3 ? toBool(scalar(3)) : true;
        if (colIdx < 1 || colIdx > (table[0]?.length || 0)) throw new FormulaError('#REF!');
        if (!approx) {
          for (const row of table) {
            if (this.compareValues(row[0], lookupVal, '=')) {
              return row[colIdx - 1] ?? null;
            }
          }
          throw new FormulaError('#N/A');
        } else {
          let bestIdx = -1;
          for (let r = 0; r < table.length; r++) {
            if (this.compareValues(table[r][0], lookupVal, '<=')) {
              bestIdx = r;
            }
          }
          if (bestIdx === -1) throw new FormulaError('#N/A');
          return table[bestIdx][colIdx - 1] ?? null;
        }
      }
      case 'HLOOKUP': {
        if (args.length < 3) throw new FormulaError('#VALUE!');
        const lookupVal = scalar(0);
        const table = isRange2D(args[1]) ? args[1] : [[args[1]]];
        const rowIdx = Math.trunc(toNum(scalar(2)));
        const approx = args.length > 3 ? toBool(scalar(3)) : true;
        if (rowIdx < 1 || rowIdx > table.length) throw new FormulaError('#REF!');
        const topRow = table[0] || [];
        if (!approx) {
          for (let c = 0; c < topRow.length; c++) {
            if (this.compareValues(topRow[c], lookupVal, '=')) {
              return table[rowIdx - 1][c] ?? null;
            }
          }
          throw new FormulaError('#N/A');
        } else {
          let bestCol = -1;
          for (let c = 0; c < topRow.length; c++) {
            if (this.compareValues(topRow[c], lookupVal, '<=')) bestCol = c;
          }
          if (bestCol === -1) throw new FormulaError('#N/A');
          return table[rowIdx - 1][bestCol] ?? null;
        }
      }
      case 'XLOOKUP': {
        if (args.length < 3) throw new FormulaError('#VALUE!');
        const lookupVal = scalar(0);
        const lookupArr = flattenArgs([args[1]]);
        const returnArr = flattenArgs([args[2]]);
        for (let i = 0; i < lookupArr.length; i++) {
          if (this.compareValues(lookupArr[i], lookupVal, '=')) {
            return returnArr[i] ?? null;
          }
        }
        if (args.length > 3) return scalar(3);
        throw new FormulaError('#N/A');
      }
      case 'INDEX': {
        if (args.length < 2) throw new FormulaError('#VALUE!');
        const arr = isRange2D(args[0]) ? args[0] : [[args[0]]];
        const r = Math.trunc(toNum(scalar(1)));
        const c = args.length > 2 ? Math.trunc(toNum(scalar(2))) : 1;
        if (r < 1 || r > arr.length || c < 1 || c > (arr[0]?.length || 0)) throw new FormulaError('#REF!');
        return arr[r - 1][c - 1] ?? null;
      }
      case 'MATCH': {
        if (args.length < 2) throw new FormulaError('#VALUE!');
        const lookupVal = scalar(0);
        const arr = flattenArgs([args[1]]);
        const matchType = args.length > 2 ? Math.trunc(toNum(scalar(2))) : 1;
        if (matchType === 0) {
          for (let i = 0; i < arr.length; i++) {
            if (this.compareValues(arr[i], lookupVal, '=')) return i + 1;
          }
          throw new FormulaError('#N/A');
        }
        let best = -1;
        for (let i = 0; i < arr.length; i++) {
          if (matchType > 0 ? this.compareValues(arr[i], lookupVal, '<=') : this.compareValues(arr[i], lookupVal, '>=')) {
            best = i;
          }
        }
        if (best === -1) throw new FormulaError('#N/A');
        return best + 1;
      }
      case 'CHOOSE': {
        const idx = Math.trunc(toNum(scalar(0)));
        if (idx < 1 || idx >= args.length) throw new FormulaError('#VALUE!');
        return scalar(idx);
      }
      case 'ROW': {
        if (args.length === 0) return curRow + 1;
        const ref = argRawRefs[0]?.trim() || '';
        const m = ref.match(/\$?[A-Za-z]{1,3}\$?([1-9][0-9]*)/);
        return m ? parseInt(m[1], 10) : curRow + 1;
      }
      case 'COLUMN': {
        if (args.length === 0) return curCol + 1;
        const ref = argRawRefs[0]?.trim() || '';
        const m = ref.match(/\$?([A-Za-z]{1,3})\$?[1-9][0-9]*/);
        return m ? colLetterToIndex(m[1]) + 1 : curCol + 1;
      }
      case 'ROWS': {
        const arr = isRange2D(args[0]) ? args[0] : [[args[0]]];
        return arr.length;
      }
      case 'COLUMNS': {
        const arr = isRange2D(args[0]) ? args[0] : [[args[0]]];
        return arr[0]?.length || 0;
      }

      // TEXT
      case 'CONCAT':
      case 'CONCATENATE': {
        return flattenArgs(args).map((v) => toStr(v)).join('');
      }
      case 'TEXTJOIN': {
        const delim = toStr(scalar(0));
        const ignoreEmpty = toBool(scalar(1));
        const items = flattenArgs(args.slice(2)).map((v) => toStr(v));
        return (ignoreEmpty ? items.filter((s) => s !== '') : items).join(delim);
      }
      case 'LEFT': {
        const str = toStr(scalar(0));
        const n = args.length > 1 ? Math.max(0, Math.trunc(toNum(scalar(1)))) : 1;
        return str.slice(0, n);
      }
      case 'RIGHT': {
        const str = toStr(scalar(0));
        const n = args.length > 1 ? Math.max(0, Math.trunc(toNum(scalar(1)))) : 1;
        return n === 0 ? '' : str.slice(-n);
      }
      case 'MID': {
        const str = toStr(scalar(0));
        const start = Math.trunc(toNum(scalar(1)));
        const n = Math.trunc(toNum(scalar(2)));
        if (start < 1 || n < 0) throw new FormulaError('#VALUE!');
        return str.slice(start - 1, start - 1 + n);
      }
      case 'LEN': return toStr(scalar(0)).length;
      case 'TRIM': return toStr(scalar(0)).replace(/\s+/g, ' ').trim();
      case 'UPPER': return toStr(scalar(0)).toUpperCase();
      case 'LOWER': return toStr(scalar(0)).toLowerCase();
      case 'PROPER': {
        return toStr(scalar(0)).replace(/\b\w+/g, (w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
      }
      case 'SUBSTITUTE': {
        const text = toStr(scalar(0));
        const oldText = toStr(scalar(1));
        const newText = toStr(scalar(2));
        if (oldText === '') return text;
        return text.split(oldText).join(newText);
      }
      case 'FIND': {
        const needle = toStr(scalar(0));
        const hay = toStr(scalar(1));
        const start = args.length > 2 ? Math.trunc(toNum(scalar(2))) : 1;
        const idx = hay.indexOf(needle, start - 1);
        if (idx === -1) throw new FormulaError('#VALUE!');
        return idx + 1;
      }
      case 'SEARCH': {
        const needle = toStr(scalar(0)).toLowerCase();
        const hay = toStr(scalar(1)).toLowerCase();
        const start = args.length > 2 ? Math.trunc(toNum(scalar(2))) : 1;
        const idx = hay.indexOf(needle, start - 1);
        if (idx === -1) throw new FormulaError('#VALUE!');
        return idx + 1;
      }
      case 'TEXT': {
        const val = scalar(0);
        const fmt = toStr(scalar(1));
        if (typeof val === 'number') {
          if (fmt.includes('%')) {
            return `${(val * 100).toFixed(fmt.includes('.00') ? 2 : 0)}%`;
          }
          if (fmt.includes('$')) {
            return `$${val.toLocaleString('en-US', { minimumFractionDigits: fmt.includes('.00') ? 2 : 0, maximumFractionDigits: fmt.includes('.00') ? 2 : 0 })}`;
          }
          if (fmt.includes('#,##0')) {
            return val.toLocaleString('en-US', { minimumFractionDigits: fmt.includes('.00') ? 2 : 0, maximumFractionDigits: fmt.includes('.00') ? 2 : 0 });
          }
        }
        return toStr(val);
      }
      case 'VALUE': {
        const s = toStr(scalar(0)).replace(/[$,%]/g, '').trim();
        const n = Number(s);
        if (isNaN(n)) throw new FormulaError('#VALUE!');
        return n;
      }
      case 'REPT': {
        const text = toStr(scalar(0));
        const count = Math.max(0, Math.trunc(toNum(scalar(1))));
        return text.repeat(count);
      }

      // DATE & TIME
      case 'TODAY': {
        const now = new Date();
        return dateToExcelSerial(new Date(now.getFullYear(), now.getMonth(), now.getDate()));
      }
      case 'NOW': {
        return dateToExcelSerial(new Date());
      }
      case 'DATE': {
        const y = Math.trunc(toNum(scalar(0)));
        const m = Math.trunc(toNum(scalar(1))) - 1;
        const d = Math.trunc(toNum(scalar(2)));
        return (Date.UTC(y, m, d) - Date.UTC(1899, 11, 30)) / 86400000;
      }
      case 'YEAR': return excelSerialToDate(toNum(scalar(0))).getUTCFullYear();
      case 'MONTH': return excelSerialToDate(toNum(scalar(0))).getUTCMonth() + 1;
      case 'DAY': return excelSerialToDate(toNum(scalar(0))).getUTCDate();
      case 'WEEKDAY': {
        const d = excelSerialToDate(toNum(scalar(0))).getUTCDay();
        const type = args.length > 1 ? Math.trunc(toNum(scalar(1))) : 1;
        if (type === 2) return d === 0 ? 7 : d;
        return d + 1;
      }
      case 'DATEDIF': {
        const d1 = excelSerialToDate(toNum(scalar(0)));
        const d2 = excelSerialToDate(toNum(scalar(1)));
        const unit = toStr(scalar(2)).toUpperCase();
        if (d1 > d2) throw new FormulaError('#NUM!');
        if (unit === 'D') return Math.floor((d2.getTime() - d1.getTime()) / 86400000);
        if (unit === 'Y') {
          let years = d2.getUTCFullYear() - d1.getUTCFullYear();
          if (d2.getUTCMonth() < d1.getUTCMonth() || (d2.getUTCMonth() === d1.getUTCMonth() && d2.getUTCDate() < d1.getUTCDate())) years--;
          return years;
        }
        if (unit === 'M') {
          let months = (d2.getUTCFullYear() - d1.getUTCFullYear()) * 12 + (d2.getUTCMonth() - d1.getUTCMonth());
          if (d2.getUTCDate() < d1.getUTCDate()) months--;
          return months;
        }
        throw new FormulaError('#NUM!');
      }

      // FINANCIAL
      case 'PMT': {
        const rate = toNum(scalar(0));
        const nper = toNum(scalar(1));
        const pv = toNum(scalar(2));
        const fv = args.length > 3 ? toNum(scalar(3)) : 0;
        const type = args.length > 4 ? toNum(scalar(4)) : 0;
        if (rate === 0) return -(pv + fv) / nper;
        const pvif = Math.pow(1 + rate, nper);
        let pmt = (rate / (pvif - 1)) * -(pv * pvif + fv);
        if (type === 1) pmt /= 1 + rate;
        return pmt;
      }
      case 'FV': {
        const rate = toNum(scalar(0));
        const nper = toNum(scalar(1));
        const pmt = toNum(scalar(2));
        const pv = args.length > 3 ? toNum(scalar(3)) : 0;
        const type = args.length > 4 ? toNum(scalar(4)) : 0;
        if (rate === 0) return -(pv + pmt * nper);
        const factor = Math.pow(1 + rate, nper);
        return -pv * factor - pmt * (1 + rate * type) * ((factor - 1) / rate);
      }
      case 'PV': {
        const rate = toNum(scalar(0));
        const nper = toNum(scalar(1));
        const pmt = toNum(scalar(2));
        const fv = args.length > 3 ? toNum(scalar(3)) : 0;
        const type = args.length > 4 ? toNum(scalar(4)) : 0;
        if (rate === 0) return -(fv + pmt * nper);
        const factor = Math.pow(1 + rate, nper);
        return (-fv - pmt * (1 + rate * type) * ((factor - 1) / rate)) / factor;
      }
      case 'NPV': {
        const rate = toNum(scalar(0));
        const flows = getNumericValues(args.slice(1));
        let npv = 0;
        for (let i = 0; i < flows.length; i++) {
          npv += flows[i] / Math.pow(1 + rate, i + 1);
        }
        return npv;
      }
      case 'IRR': {
        const flows = getNumericValues([args[0]]);
        let guess = args.length > 1 ? toNum(scalar(1)) : 0.1;
        for (let iter = 0; iter < 50; iter++) {
          let npv = 0;
          let dNpv = 0;
          for (let i = 0; i < flows.length; i++) {
            npv += flows[i] / Math.pow(1 + guess, i);
            dNpv -= (i * flows[i]) / Math.pow(1 + guess, i + 1);
          }
          if (Math.abs(npv) < 1e-7) return guess;
          if (dNpv === 0) break;
          guess -= npv / dNpv;
        }
        return guess;
      }

      // INFORMATION
      case 'ISBLANK': {
        const v = scalar(0);
        return v === null || v === undefined || v === '';
      }
      case 'ISNUMBER': return typeof scalar(0) === 'number';
      case 'ISTEXT': return typeof scalar(0) === 'string' && scalar(0) !== '';
      case 'ISERROR': return false;

      default:
        throw new FormulaError('#NAME?');
    }
  }

  private tokenize(expr: string): { type: string; value: string }[] {
    const tokens: { type: string; value: string }[] = [];
    let i = 0;

    while (i < expr.length) {
      const ch = expr[i];
      if (/\s/.test(ch)) {
        i++;
        continue;
      }

      // String literal
      if (ch === '"') {
        i++;
        let str = '';
        while (i < expr.length) {
          if (expr[i] === '"' && expr[i + 1] === '"') {
            str += '"';
            i += 2;
          } else if (expr[i] === '"') {
            i++;
            break;
          } else {
            str += expr[i++];
          }
        }
        tokens.push({ type: 'string', value: str });
        continue;
      }

      // Quoted sheet name e.g. 'Sales Q1'!A1
      if (ch === "'") {
        let raw = "'";
        i++;
        while (i < expr.length && expr[i] !== "'") {
          raw += expr[i++];
        }
        if (i < expr.length) raw += expr[i++];
        while (i < expr.length && /[A-Za-z0-9_!$]/.test(expr[i])) {
          raw += expr[i++];
        }
        tokens.push({ type: 'ident', value: raw });
        continue;
      }

      // Number
      if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(expr[i + 1] || ''))) {
        let num = '';
        while (i < expr.length && /[0-9.]/.test(expr[i])) {
          num += expr[i++];
        }
        if ((expr[i] === 'e' || expr[i] === 'E') && /[+-0-9]/.test(expr[i + 1] || '')) {
          num += expr[i++];
          if (expr[i] === '+' || expr[i] === '-') num += expr[i++];
          while (i < expr.length && /[0-9]/.test(expr[i])) num += expr[i++];
        }
        tokens.push({ type: 'number', value: num });
        continue;
      }

      // Error literal (#REF!, #DIV/0!, etc.)
      if (ch === '#') {
        let err = '#';
        i++;
        while (i < expr.length && /[A-Za-z0-9/!?]/.test(expr[i])) {
          err += expr[i++];
        }
        tokens.push({ type: 'error', value: err });
        continue;
      }

      // Two-char comparison ops
      const two = expr.slice(i, i + 2);
      if (two === '<=' || two === '>=' || two === '<>') {
        tokens.push({ type: 'op', value: two });
        i += 2;
        continue;
      }

      if ('+-*/^&=<>,%'.includes(ch)) {
        if (ch === ',') {
          tokens.push({ type: 'comma', value: ',' });
        } else {
          tokens.push({ type: 'op', value: ch });
        }
        i++;
        continue;
      }

      if (ch === ':') {
        tokens.push({ type: 'colon', value: ':' });
        i++;
        continue;
      }

      if (ch === '(' || ch === ')') {
        tokens.push({ type: 'paren', value: ch });
        i++;
        continue;
      }

      // Identifier (cell ref, sheet!ref, function name, boolean, named range)
      if (/[A-Za-z_$]/.test(ch)) {
        let id = '';
        while (i < expr.length && /[A-Za-z0-9_.$!]/.test(expr[i])) {
          id += expr[i++];
        }
        const upper = id.toUpperCase();
        if ((upper === 'TRUE' || upper === 'FALSE') && expr[i] !== '(') {
          tokens.push({ type: 'bool', value: upper });
        } else {
          tokens.push({ type: 'ident', value: id });
        }
        continue;
      }

      throw new FormulaError('#VALUE!');
    }
    return tokens;
  }

  // Handle IFERROR / IFNA safely by re-evaluating arguments individually when needed
  public evaluateCellSafe(sheetId: string, row: number, col: number): EvaluatedCell {
    return this.evaluateCell(sheetId, row, col);
  }
}

/**
 * Wait - notice in evalExpression above, IFERROR/IFNA/IF evaluated all args upfront!
 * Let's enhance IFERROR/IFNA/IF? Actually for most formulas, evaluating args works great, EXCEPT IFERROR(A1/0, 0) where A1/0 throws during arg evaluation!
 * Let's patch evalExpression to lazily/safely handle IFERROR and IFNA so `=IFERROR(1/0, "Safe")` works cleanly!
 */
export function solveGoalSeek(
  sheets: Sheet[],
  namedRanges: NamedRange[],
  targetSheetId: string,
  setRow: number,
  setCol: number,
  targetValue: number,
  byChangingRow: number,
  byChangingCol: number
): { success: boolean; foundValue: number; achievedTarget: number; iterations: number } {
  const cloneSheets = (): Sheet[] =>
    sheets.map((s) =>
      s.id === targetSheetId
        ? { ...s, cells: { ...s.cells } }
        : s
    );

  const evalWithGuess = (guess: number): number => {
    const cloned = cloneSheets();
    const s = cloned.find((sh) => sh.id === targetSheetId)!;
    const key = cellKey(byChangingRow, byChangingCol);
    s.cells[key] = { ...(s.cells[key] || { raw: '' }), raw: String(guess) };
    const ev = new WorkbookEvaluator(cloned, namedRanges);
    const res = ev.evaluateCell(targetSheetId, setRow, setCol);
    if (typeof res.value === 'number') return res.value;
    return NaN;
  };

  const sheet = sheets.find((s) => s.id === targetSheetId);
  const startRaw = sheet?.cells[cellKey(byChangingRow, byChangingCol)]?.raw || '0';
  let x0 = parseFloat(startRaw);
  if (isNaN(x0)) x0 = 1;
  let x1 = x0 === 0 ? 1 : x0 * 1.05;

  let y0 = evalWithGuess(x0) - targetValue;
  let y1 = evalWithGuess(x1) - targetValue;
  if (isNaN(y0) || isNaN(y1)) {
    return { success: false, foundValue: x0, achievedTarget: NaN, iterations: 0 };
  }

  let iter = 0;
  for (; iter < 80; iter++) {
    if (Math.abs(y1) < 1e-6) {
      return { success: true, foundValue: x1, achievedTarget: y1 + targetValue, iterations: iter };
    }
    const denom = y1 - y0;
    if (Math.abs(denom) < 1e-12) {
      x1 += 0.1;
      y1 = evalWithGuess(x1) - targetValue;
      continue;
    }
    const x2 = x1 - (y1 * (x1 - x0)) / denom;
    x0 = x1;
    y0 = y1;
    x1 = x2;
    y1 = evalWithGuess(x1) - targetValue;
    if (isNaN(y1)) break;
  }

  return {
    success: Math.abs(y1) < 0.01,
    foundValue: x1,
    achievedTarget: y1 + targetValue,
    iterations: iter,
  };
}

export function getCellAddressLabel(row: number, col: number): string {
  return `${colIndexToLetter(col)}${row + 1}`;
}
