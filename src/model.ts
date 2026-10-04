import type {
  Column, ColumnKind, DataCell, DataRow, FilterRule, MatchMode,
  SheetData, SortRule, TableData,
} from './types';

const EMPTY: DataCell = { value: null, text: '', kind: 'empty', numeric: null, date: null };
const HISTOGRAM_BINS = 20;

/** Strictly parse a whole visible number; the returned percent is in displayed units. */
export function parseNumeric(input: string): number | null {
  const text = input.trim().replace(/\u00a0/g, ' ');
  if (!text) return null;
  const parenthesized = /^\((.*)\)$/.exec(text);
  const body = (parenthesized ? parenthesized[1] : text).trim();
  const match = /^([+-]?)(?:(\d{1,3}(?:,\d{3})+|\d+)(?:\.(\d+))?|\.(\d+))\s*([kKmMbB万亿]?)\s*(%?)$/.exec(body);
  if (!match) return null;
  if (parenthesized && match[1]) return null;
  const mantissa = Number((match[2] || '0').replaceAll(',', '') + (match[3] ? '.' + match[3] : match[4] ? '.' + match[4] : ''));
  const multiplier: Record<string, number> = { '': 1, k: 1e3, m: 1e6, b: 1e9, '万': 1e4, '亿': 1e8 };
  const result = mantissa * multiplier[match[5].toLowerCase()] * (parenthesized || match[1] === '-' ? -1 : 1);
  return Number.isFinite(result) ? result : null;
}

const filled = (cell: DataCell): boolean => cell.kind !== 'empty' && cell.kind !== 'error' && cell.value !== null && cell.text !== '';
const at = (row: DataRow, index: number): DataCell => row.cells[index] ?? EMPTY;

function columnKind(cells: DataCell[]): ColumnKind {
  const present = cells.filter(filled);
  if (present.length === 0) return 'text';
  if (present.every(cell => cell.kind === 'number')) return 'number';
  if (present.every(cell => cell.kind === 'percent')) return 'percent';
  if (present.every(cell => cell.kind === 'date')) return 'date';
  return 'text';
}

function columnSummary(index: number, cells: DataCell[], label: string): Column {
  const kind = columnKind(cells);
  const numeric = cells.map(cell => cell.numeric).filter((value): value is number => value !== null && Number.isFinite(value));
  let min: number | null = null;
  let max: number | null = null;
  for (const value of numeric) {
    min = min === null ? value : Math.min(min, value);
    max = max === null ? value : Math.max(max, value);
  }
  const histogram = Array.from({ length: HISTOGRAM_BINS }, () => 0);
  if (min !== null && max !== null) {
    for (const value of numeric) {
      const bin = max === min ? 0 : Math.min(HISTOGRAM_BINS - 1, Math.floor((value - min) / (max - min) * HISTOGRAM_BINS));
      histogram[bin] += 1;
    }
  }
  const counts = new Map<string, number>();
  for (const cell of cells) if (filled(cell)) counts.set(cell.text, (counts.get(cell.text) ?? 0) + 1);
  return {
    id: `col-${index}`, index, label, kind,
    emptyCount: cells.length - cells.filter(filled).length,
    validCount: kind === 'number' || kind === 'percent' ? numeric.length : cells.filter(filled).length,
    min, max, histogram,
    categories: [...counts].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'zh-CN'))
      .slice(0, 50).map(([value, count]) => ({ value, count })),
  };
}

/** Turn one sheet into a position-keyed table without changing source cells. */
export function buildTable(sheet: SheetData, headerRow: number): TableData {
  const headerIndex = headerRow - sheet.firstRow;
  if (!Number.isInteger(headerRow) || headerIndex < 0 || headerIndex >= sheet.rows.length) {
    throw new Error('表头行超出工作表范围');
  }
  const width = sheet.rows.reduce((max, row) => Math.max(max, row.length), 0);
  const headerCells = Array.from({ length: width }, (_, index) => sheet.rows[headerIndex]?.[index] ?? EMPTY);
  const rows: DataRow[] = [];
  for (let offset = headerIndex + 1; offset < sheet.rows.length; offset++) {
    const source = sheet.rows[offset] ?? [];
    if (!source.some(cell => cell && cell.kind !== 'empty')) continue;
    const sourceRow = sheet.firstRow + offset;
    rows.push({
      id: sourceRow, sourceRow,
      cells: Array.from({ length: width }, (_, index) => source[index] ?? EMPTY),
    });
  }
  const labels = new Map<string, number>();
  const columns = headerCells.map((header, index) => {
    const original = header.text.trim();
    const count = (labels.get(original) ?? 0) + 1;
    labels.set(original, count);
    const label = original ? `${original}${count > 1 ? ` (${count})` : ''}` : `未命名列 ${index + 1}`;
    return columnSummary(index, rows.map(row => row.cells[index]), label);
  });
  return {
    sheetName: sheet.name, headerRow, columns, headerCells, rows,
    date1904: sheet.date1904, warnings: [],
  };
}

const numericInput = (value: string): number | null => parseNumeric(value);
const dateInput = (value: string): string | null => {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value.trim());
  if (!match) return null;
  const year = Number(match[1]), month = Number(match[2]), day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day ? value.trim() : null;
};

function ruleState(rule: FilterRule, column: Column): { active: boolean; error: string | null } {
  if (!column) return { active: false, error: '条件对应的列不存在' };
  if (rule.op === 'empty' || rule.op === 'notEmpty') return { active: true, error: null };
  if (rule.op === 'oneOf') return rule.values.length
    ? { active: true, error: null } : { active: false, error: '请选择至少一个分类' };
  if (rule.op === 'contains' || rule.op === 'notContains' || rule.op === 'equals') return rule.value.trim()
    ? { active: true, error: null } : { active: false, error: '请输入匹配文字' };
  if (rule.op === 'range' || rule.op === 'dateRange') {
    const min = rule.min.trim(), max = rule.max.trim();
    if (!min && !max) return { active: false, error: '请填写范围边界' };
    const parse = rule.op === 'range' ? numericInput : dateInput;
    if (min && parse(min) === null) return { active: true, error: rule.op === 'range' ? '最小值不是有效数字' : '最早日期应为 YYYY-MM-DD' };
    if (max && parse(max) === null) return { active: true, error: rule.op === 'range' ? '最大值不是有效数字' : '最晚日期应为 YYYY-MM-DD' };
    if (min && max && (parse(min) as string | number) > (parse(max) as string | number)) {
      return { active: true, error: '范围下界不能大于上界' };
    }
    return { active: true, error: null };
  }
  return { active: true, error: '未知筛选操作' };
}

export function validateFilter(rule: FilterRule, column: Column): string | null {
  return ruleState(rule, column).error;
}

function matches(cell: DataCell, rule: FilterRule): boolean {
  if (rule.op === 'empty') return !filled(cell);
  if (rule.op === 'notEmpty') return filled(cell);
  if (rule.op === 'oneOf') return filled(cell) && rule.values.includes(cell.text);
  if (rule.op === 'contains') return filled(cell) && cell.text.toLocaleLowerCase().includes(rule.value.trim().toLocaleLowerCase());
  if (rule.op === 'notContains') return filled(cell) && !cell.text.toLocaleLowerCase().includes(rule.value.trim().toLocaleLowerCase());
  if (rule.op === 'equals') return filled(cell) && cell.text.toLocaleLowerCase() === rule.value.trim().toLocaleLowerCase();
  if (rule.op === 'range') {
    const value = cell.numeric;
    if (value === null || !Number.isFinite(value)) return false;
    const min = rule.min.trim() ? numericInput(rule.min) : null;
    const max = rule.max.trim() ? numericInput(rule.max) : null;
    return (min === null || value >= min) && (max === null || value <= max);
  }
  if (rule.op === 'dateRange') {
    const value = cell.date;
    if (value === null) return false;
    const min = rule.min.trim() ? dateInput(rule.min) : null;
    const max = rule.max.trim() ? dateInput(rule.max) : null;
    return (min === null || value >= min) && (max === null || value <= max);
  }
  return false;
}

function compareCells(left: DataCell, right: DataCell, kind: ColumnKind): number {
  const leftEmpty = !filled(left), rightEmpty = !filled(right);
  if (leftEmpty || rightEmpty) return Number(leftEmpty) - Number(rightEmpty);
  if ((kind === 'number' || kind === 'percent') && left.numeric !== null && right.numeric !== null) return left.numeric - right.numeric;
  if (kind === 'date' && left.date !== null && right.date !== null) return left.date.localeCompare(right.date);
  return left.text.localeCompare(right.text, 'zh-CN', { numeric: true, sensitivity: 'base' });
}

/** Filter and sort the entire table. Caller owns pagination and export selection. */
export function queryRows(
  table: TableData, filters: FilterRule[], matchMode: MatchMode, search: string, sort: SortRule | null,
): DataRow[] {
  const active: { rule: FilterRule; column: Column }[] = [];
  for (const rule of filters) {
    const column = table.columns.find(item => item.id === rule.columnId);
    if (!column) throw new Error('条件对应的列不存在');
    const state = ruleState(rule, column);
    if (state.active && state.error) throw new Error(state.error);
    if (state.active) active.push({ rule, column });
  }
  const needle = search.trim().toLocaleLowerCase();
  const result = table.rows.filter(row => {
    if (needle && !row.cells.some(cell => cell.text.toLocaleLowerCase().includes(needle))) return false;
    if (!active.length) return true;
    const test = ({ rule, column }: { rule: FilterRule; column: Column }) => matches(at(row, column.index), rule);
    return matchMode === 'any' ? active.some(test) : active.every(test);
  });
  if (sort) {
    const column = table.columns.find(item => item.id === sort.columnId);
    if (!column) throw new Error('排序列不存在');
    result.sort((left, right) => {
      const first = at(left, column.index), second = at(right, column.index);
      if (!filled(first) || !filled(second)) return Number(!filled(first)) - Number(!filled(second)) || left.sourceRow - right.sourceRow;
      const value = compareCells(first, second, column.kind);
      return (sort.direction === 'desc' ? -value : value) || left.sourceRow - right.sourceRow;
    });
  }
  return result;
}
