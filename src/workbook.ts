import type { DataCell, DataRow, SheetData, TableData, WorkbookData } from './types';
import { parseNumeric } from './model';

interface ExcelCell {
  t?: string;
  v?: unknown;
  w?: string;
  z?: string;
  f?: string;
  l?: { Target?: string };
}
type ExcelSheet = Record<string, unknown>;
interface ExcelBook {
  SheetNames: string[];
  Sheets: Record<string, ExcelSheet>;
  Workbook?: { WBProps?: { date1904?: boolean } };
}
interface SheetJs {
  read: (data: Uint8Array | string, options: Record<string, unknown>) => ExcelBook;
  write: (book: ExcelBook, options: Record<string, unknown>) => Uint8Array | ArrayBuffer;
  utils: {
    book_new: () => ExcelBook;
    book_append_sheet: (book: ExcelBook, sheet: ExcelSheet, name: string) => void;
    encode_cell: (address: { r: number; c: number }) => string;
    encode_range: (range: { s: { r: number; c: number }; e: { r: number; c: number } }) => string;
  };
  SSF: {
    is_date: (format: string) => boolean;
    parse_date_code: (value: number, options?: { date1904?: boolean }) =>
      { y: number; m: number; d: number; H: number; M: number; S: number } | null;
    format: (format: string, value: number, options?: { date1904?: boolean }) => string;
  };
}

declare global {
  // Loaded by the local vendor script before the application starts.
  var XLSX: SheetJs | undefined;
}

const MAX_BYTES = 25 * 1024 * 1024;
const MAX_ROWS = 100_000;
const MAX_COLUMNS = 250;
const EMPTY: DataCell = { value: null, text: '', kind: 'empty', numeric: null, date: null };
const CELL_ADDRESS = /^([A-Z]{1,3})([1-9]\d*)$/;
const SAFE_LINK = /^(https?:\/\/|mailto:)/i;

function xlsx(): SheetJs {
  if (!globalThis.XLSX) throw new Error('本地 Excel 解析组件未加载，请重新打开页面');
  return globalThis.XLSX;
}

function columnIndex(letters: string): number {
  let index = 0;
  for (const character of letters) index = index * 26 + character.charCodeAt(0) - 64;
  return index - 1;
}

function isoDate(year: number, month: number, day: number): string | null {
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return null;
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
  return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function asDate(value: number, date1904: boolean): string | null {
  const parsed = xlsx().SSF.parse_date_code(value, { date1904 });
  return parsed ? isoDate(parsed.y, parsed.m, parsed.d) : null;
}

function dateString(text: string): string | null {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text.trim());
  return match ? isoDate(Number(match[1]), Number(match[2]), Number(match[3])) : null;
}

function formatText(cell: ExcelCell, raw: string | number | boolean, date1904: boolean): string {
  if (typeof cell.w === 'string') return cell.w;
  if (typeof raw === 'number' && cell.z) {
    try { return xlsx().SSF.format(cell.z, raw, { date1904 }); } catch { /* use raw below */ }
  }
  return String(raw);
}

function isSafeLink(value: string | undefined): value is string {
  return typeof value === 'string' && SAFE_LINK.test(value.trim());
}

function convertCell(cell: ExcelCell, date1904: boolean): DataCell {
  const format = typeof cell.z === 'string' ? cell.z : undefined;
  const link = isSafeLink(cell.l?.Target) ? cell.l.Target : undefined;
  const formula = typeof cell.f === 'string';
  const common = { ...(format ? { format } : {}), ...(link ? { link } : {}), ...(formula ? { formula: true } : {}) };
  if (formula && (cell.v === undefined || cell.v === null)) {
    return { value: null, text: '公式缺少缓存值', kind: 'error', numeric: null, date: null, ...common };
  }
  if (cell.t === 'e') {
    return { value: null, text: typeof cell.w === 'string' ? cell.w : 'Excel 错误值', kind: 'error', numeric: null, date: null, ...common };
  }
  if (cell.v === undefined || cell.v === null || cell.t === 'z') return { ...EMPTY, ...common };
  if (cell.t === 'b' || typeof cell.v === 'boolean') {
    const value = Boolean(cell.v);
    return { value, text: formatText(cell, value, date1904), kind: 'boolean', numeric: null, date: null, ...common };
  }
  if (cell.v instanceof Date) {
    const date = isoDate(cell.v.getUTCFullYear(), cell.v.getUTCMonth() + 1, cell.v.getUTCDate());
    return { value: date ?? cell.v.toISOString(), text: formatText(cell, date ?? cell.v.toISOString(), date1904), kind: 'date', numeric: null, date, ...common };
  }
  if (typeof cell.v === 'number' && Number.isFinite(cell.v)) {
    const value = cell.v;
    const text = formatText(cell, value, date1904);
    if (format && xlsx().SSF.is_date(format)) {
      const date = asDate(value, date1904);
      return { value, text, kind: date ? 'date' : 'number', numeric: date ? null : value, date, ...common };
    }
    if (format?.includes('%')) return { value, text, kind: 'percent', numeric: value * 100, date: null, ...common };
    return { value, text, kind: 'number', numeric: value, date: null, ...common };
  }
  const value = String(cell.v);
  const text = formatText(cell, value, date1904);
  const date = dateString(value);
  return { value, text, kind: date ? 'date' : 'text', numeric: date ? null : parseNumeric(value), date, ...common };
}

function sheetData(name: string, sheet: ExcelSheet, date1904: boolean, warnings: string[]): SheetData {
  const entries: { row: number; column: number; value: DataCell }[] = [];
  let minRow = Infinity, maxRow = -1, maxColumn = -1;
  let missingFormulas = 0;
  for (const [address, raw] of Object.entries(sheet)) {
    const match = CELL_ADDRESS.exec(address);
    if (!match || !raw || typeof raw !== 'object') continue;
    const row = Number(match[2]) - 1;
    const column = columnIndex(match[1]);
    if (!Number.isSafeInteger(row) || column >= MAX_COLUMNS) {
      throw new Error(`工作表「${name}」超过 ${MAX_COLUMNS} 列或含无效行号`);
    }
    const value = convertCell(raw as ExcelCell, date1904);
    if (value.kind === 'empty') continue;
    if (value.kind === 'error' && value.formula) missingFormulas++;
    entries.push({ row, column, value });
    minRow = Math.min(minRow, row);
    maxRow = Math.max(maxRow, row);
    maxColumn = Math.max(maxColumn, column);
  }
  if (!entries.length) return { name, rows: [], firstRow: 1, date1904 };
  if (maxRow - minRow > MAX_ROWS) {
    throw new Error(`工作表「${name}」超过 ${MAX_ROWS} 个数据行跨度，请缩小后再导入`);
  }
  const rows: DataCell[][] = Array.from({ length: maxRow - minRow + 1 }, () => []);
  for (const entry of entries) rows[entry.row - minRow][entry.column] = entry.value;
  if (missingFormulas) warnings.push(`工作表「${name}」有 ${missingFormulas} 个公式缺少缓存值，无法显示或导出计算结果`);
  if (maxColumn + 1 > MAX_COLUMNS) throw new Error(`工作表「${name}」超过 ${MAX_COLUMNS} 列`);
  return { name, rows, firstRow: minRow + 1, date1904 };
}

function sniff(data: Uint8Array, name: string): string | null {
  const extension = name.split('.').pop()?.toLowerCase();
  if (extension !== 'xlsx' && extension !== 'xls' && extension !== 'csv') {
    throw new Error('请选择 .xlsx、.xls 或 .csv 文件');
  }
  const zip = data.length >= 4 && data[0] === 0x50 && data[1] === 0x4b && [3, 5, 7].includes(data[2]);
  const cfb = data.length >= 8 && [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((byte, index) => data[index] === byte);
  if (extension === 'xlsx' && !zip) throw new Error('文件内容不是有效的 XLSX 压缩包');
  if (extension === 'xls' && !cfb) throw new Error('文件内容不是有效的 XLS 文件');
  if (extension === 'csv') {
    if (zip || cfb || data.includes(0)) throw new Error('CSV 文件包含二进制内容');
    try { return new TextDecoder('utf-8', { fatal: true }).decode(data); }
    catch {
      try { return new TextDecoder('gb18030', { fatal: true }).decode(data); }
      catch { throw new Error('CSV 文本编码无法识别'); }
    }
  }
  return null;
}

export function readWorkbook(data: ArrayBuffer, name: string): WorkbookData {
  if (!data.byteLength) throw new Error('文件为空');
  if (data.byteLength > MAX_BYTES) throw new Error('文件超过 25 MB 导入上限');
  const bytes = new Uint8Array(data);
  const csv = sniff(bytes, name);
  let book: ExcelBook;
  try {
    book = xlsx().read(csv ?? bytes, {
      type: csv === null ? 'array' : 'string', raw: csv !== null,
      cellText: true, cellNF: true, cellDates: false, WTF: true,
    });
  } catch (error) {
    throw new Error(`无法解析 Excel 文件：${error instanceof Error ? error.message : String(error)}`);
  }
  if (!book.SheetNames?.length) throw new Error('文件没有可用工作表');
  const date1904 = Boolean(book.Workbook?.WBProps?.date1904);
  const warnings: string[] = [];
  const sheets = book.SheetNames.map(sheetName => sheetData(sheetName, book.Sheets[sheetName] ?? {}, date1904, warnings));
  return { name, size: data.byteLength, sheets, warnings };
}

function toExcelCell(cell: DataCell): ExcelCell {
  const format = cell.format ? { z: cell.format } : {};
  const link = isSafeLink(cell.link) ? { l: { Target: cell.link } } : {};
  if (cell.kind === 'empty' || cell.kind === 'error' || cell.value === null) return { t: 'z' };
  if (typeof cell.value === 'number') return { t: 'n', v: cell.value, ...format, ...link };
  if (typeof cell.value === 'boolean') return { t: 'b', v: cell.value, ...link };
  return { t: 's', v: cell.value, ...link };
}

/** Export passed rows in passed order. Formulas are deliberately written as cached constants. */
export function exportWorkbook(table: TableData, rows: DataRow[], columnIds: string[]): Uint8Array {
  const api = xlsx();
  if (!columnIds.length) throw new Error('至少选择一列导出');
  const columns = columnIds.map(id => {
    const column = table.columns.find(item => item.id === id);
    if (!column) throw new Error(`导出列不存在：${id}`);
    return column;
  });
  const sheet: ExcelSheet = {};
  const writeCell = (r: number, c: number, cell: DataCell) => {
    sheet[api.utils.encode_cell({ r, c })] = toExcelCell(cell);
  };
  columns.forEach((column, index) => writeCell(0, index, table.headerCells[column.index] ?? EMPTY));
  rows.forEach((row, rowIndex) => columns.forEach((column, columnIndex) => {
    writeCell(rowIndex + 1, columnIndex, row.cells[column.index] ?? EMPTY);
  }));
  sheet['!ref'] = api.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: columns.length - 1 } });
  sheet['!autofilter'] = { ref: sheet['!ref'] };
  sheet['!cols'] = columns.map(column => ({ wch: Math.max(12, Math.min(48, column.label.length + 4)) }));
  const book = api.utils.book_new();
  book.Workbook = { WBProps: { date1904: table.date1904 } };
  api.utils.book_append_sheet(book, sheet, table.sheetName.slice(0, 31) || '筛选结果');
  const result = api.write(book, { bookType: 'xlsx', type: 'array', compression: true, cellNF: true });
  return result instanceof Uint8Array ? result : new Uint8Array(result);
}

export function downloadWorkbook(bytes: Uint8Array, filename: string): void {
  const blob = new Blob([new Uint8Array(bytes)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename.toLowerCase().endsWith('.xlsx') ? filename : `${filename}.xlsx`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}
