import { describe, expect, it } from 'vitest';
import { createDemoBytes, createDemoWorkbook } from '../src/demo';
import { buildTable, parseNumeric, queryRows, validateFilter } from '../src/model';
import { exportWorkbook, readWorkbook } from '../src/workbook';
import type { DataCell, FilterOp, FilterRule, SheetData } from '../src/types';

const bufferOf = (bytes: Uint8Array): ArrayBuffer =>
  bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
const rule = (columnId: string, op: FilterOp, fields: Partial<FilterRule> = {}): FilterRule => ({
  id: 'rule', columnId, op, value: '', min: '', max: '', values: [], ...fields,
});
const cell = (value: string | number | null, kind: DataCell['kind'] = 'text'): DataCell => ({
  value, kind, text: value === null ? '' : String(value),
  numeric: typeof value === 'number' ? value : null, date: null,
});

describe('numbers and table model', () => {
  it('parses complete numeric input with grouped digits and suffixes', () => {
    expect(parseNumeric('1,234.5')).toBe(1234.5);
    expect(parseNumeric('2K')).toBe(2000);
    expect(parseNumeric('3.5万')).toBe(35000);
    expect(parseNumeric('(1.2亿)')).toBe(-120000000);
    expect(parseNumeric('2%')).toBe(2);
    for (const input of ['', 'true', '2026-10-03', 'foo12bar', '12,34', 'Infinity']) {
      expect(parseNumeric(input)).toBeNull();
    }
  });

  it('keeps source positions and disambiguates duplicate and blank headings', () => {
    const sheet: SheetData = {
      name: '测试', firstRow: 3, date1904: false,
      rows: [
        [cell('编号'), cell('名称'), cell('名称'), cell(null, 'empty')],
        [cell('001'), cell('甲'), cell('乙'), cell('备注')],
        [],
        [cell('002'), cell('丙'), cell('丁'), cell(null, 'empty')],
      ],
    };
    const table = buildTable(sheet, 3);
    expect(table.columns.map(column => [column.id, column.label])).toEqual([
      ['col-0', '编号'], ['col-1', '名称'], ['col-2', '名称 (2)'], ['col-3', '未命名列 4'],
    ]);
    expect(table.rows.map(row => row.sourceRow)).toEqual([4, 6]);
    expect(table.rows[1].cells[3].value).toBeNull();
    expect(() => buildTable(sheet, 2)).toThrow('表头行');
  });

  it('filters closed ranges, blanks, categories, search and stable sorting', () => {
    const demo = createDemoWorkbook();
    const table = buildTable(demo.sheets[0], demo.sheets[0].firstRow);
    expect(table.rows).toHaveLength(128);
    expect(table.columns[0].kind).toBe('text');
    expect(table.columns[4].kind).toBe('percent');
    const target = table.rows[0];
    const followers = target.cells[3].numeric!;
    const exact = rule('col-3', 'range', { min: String(followers), max: String(followers) });
    expect(queryRows(table, [exact], 'all', '', null).map(row => row.id)).toEqual([target.id]);
    const china = rule('col-1', 'oneOf', { values: ['中国'] });
    expect(queryRows(table, [china, exact], 'all', '', null)).toHaveLength(1);
    expect(queryRows(table, [china, exact], 'any', '', null).length).toBeGreaterThan(1);
    const empty = rule('col-4', 'empty');
    expect(queryRows(table, [empty], 'all', '', null).length).toBeGreaterThan(0);
    expect(queryRows(table, [], 'all', 'DEMO-0001', null)).toHaveLength(1);
    const sorted = queryRows(table, [], 'all', '', { columnId: 'col-3', direction: 'desc' });
    expect(sorted[0].cells[3].numeric).toBeGreaterThan(sorted[1].cells[3].numeric!);
    expect(table.rows[0].id).toBe(target.id);
    expect(validateFilter(rule('col-3', 'range', { min: '5K', max: '2K' }), table.columns[3])).toContain('下界');
    expect(() => queryRows(table, [rule('col-3', 'range', { min: 'oops' })], 'all', '', null)).toThrow();
    expect(queryRows(table, [rule('col-3', 'range')], 'all', '', null)).toHaveLength(128);
  });
});

describe('workbook input and output', () => {
  it('reads two demo sheets and preserves dates, percentages, text IDs and blanks', () => {
    const demo = readWorkbook(bufferOf(createDemoBytes()), 'example.xlsx');
    expect(demo.sheets.map(sheet => sheet.name)).toEqual(['虚构创作者', '虚构市场计划']);
    const row = demo.sheets[0].rows[1];
    expect(row[0].value).toBe('DEMO-0001');
    expect(row[4].kind).toBe('percent');
    expect(row[4].numeric).toBeCloseTo((row[4].value as number) * 100);
    expect(row[5].kind).toBe('date');
    expect(row[5].date).toMatch(/^\d{4}-\d{2}-\d{2}$/);
    expect(demo.sheets[0].rows[17][4]).toBeUndefined();
  });

  it('rejects incorrect signatures, oversized bytes and actual out-of-range cells', () => {
    expect(() => readWorkbook(new TextEncoder().encode('hello').buffer, 'wrong.xlsx')).toThrow('XLSX');
    expect(() => readWorkbook(new Uint8Array(25 * 1024 * 1024 + 1).buffer, 'large.csv')).toThrow('25 MB');
    const api = globalThis.XLSX!;
    const book = api.utils.book_new();
    api.utils.book_append_sheet(book, { A1: { t: 's', v: '标题' }, IQ2: { t: 'n', v: 1 }, '!ref': 'A1:IQ2' }, '太宽');
    const bytes = api.write(book, { bookType: 'xlsx', type: 'array' });
    expect(() => readWorkbook(bufferOf(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)), 'wide.xlsx')).toThrow('250');
  });

  it('keeps CSV text identifiers while allowing numeric and date filters', () => {
    const csv = new TextEncoder().encode('编号,数量,日期\n001,2万,2026-10-03\n002,,2026-10-04');
    const parsed = readWorkbook(bufferOf(csv), 'records.csv');
    const table = buildTable(parsed.sheets[0], 1);
    expect(table.rows[0].cells[0].value).toBe('001');
    expect(table.rows[0].cells[1].value).toBe('2万');
    expect(table.rows[0].cells[1].numeric).toBe(20_000);
    expect(table.rows[1].cells[1].numeric).toBeNull();
    expect(queryRows(table, [rule('col-1', 'range', { min: '20K', max: '20K' })], 'all', '', null)).toHaveLength(1);
    expect(queryRows(table, [rule('col-2', 'dateRange', { min: '2026-10-04' })], 'all', '', null)).toHaveLength(1);
  });

  it('preserves 1904 date system, cached formula values and warns on missing cache', () => {
    const api = globalThis.XLSX!;
    const book = api.utils.book_new();
    book.Workbook = { WBProps: { date1904: true } };
    const sheet = {
      A1: { t: 's', v: '日期' }, B1: { t: 's', v: '结果' },
      A2: { t: 'n', v: 1, z: 'yyyy-mm-dd' },
      B2: { t: 'n', v: 42, f: '40+2' },
      B3: { t: 'n', f: '1+1' },
      '!ref': 'A1:B3',
    };
    api.utils.book_append_sheet(book, sheet, '公式');
    const bytes = api.write(book, { bookType: 'xlsx', type: 'array' });
    const data = readWorkbook(bufferOf(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)), 'formula.xlsx');
    expect(data.sheets[0].date1904).toBe(true);
    expect(data.sheets[0].rows[1][0].date).toBe('1904-01-02');
    expect(data.sheets[0].rows[1][1].value).toBe(42);
    expect(data.sheets[0].rows[1][1].formula).toBe(true);
    expect(data.warnings.join(' ')).toContain('缺少缓存值');
    const table = buildTable(data.sheets[0], 1);
    expect(table.rows).toHaveLength(2);
    const exported = exportWorkbook(table, table.rows, ['col-1', 'col-0']);
    const roundtrip = api.read(exported, { type: 'array', cellNF: true, cellText: true });
    const outputSheet = roundtrip.Sheets['公式'] as Record<string, { v?: unknown; f?: string }>;
    expect(outputSheet.A2.v).toBe(42);
    expect(outputSheet.A2.f).toBeUndefined();
    expect(outputSheet.B2.v).toBe(1);
    expect(roundtrip.Workbook?.WBProps?.date1904).toBe(true);
  });

  it('exports all passed rows in their order and only requested columns', () => {
    const demo = createDemoWorkbook();
    const table = buildTable(demo.sheets[0], 1);
    const selected = queryRows(table, [rule('col-1', 'oneOf', { values: ['中国'] })], 'all', '', { columnId: 'col-3', direction: 'desc' });
    const exported = exportWorkbook(table, selected, ['col-0', 'col-4']);
    const back = readWorkbook(bufferOf(exported), 'selected.xlsx');
    const result = buildTable(back.sheets[0], 1);
    expect(result.columns).toHaveLength(2);
    expect(result.rows).toHaveLength(selected.length);
    expect(result.rows[0].cells[0].value).toBe(selected[0].cells[0].value);
    expect(result.rows[0].cells[1].value).toBe(selected[0].cells[4].value);
    expect(result.rows[0].cells[1].kind).toBe('percent');
  });

  it('exports duplicate original headings and drops unsafe hyperlinks without mutating input', () => {
    const sheet: SheetData = {
      name: '链接', firstRow: 1, date1904: false,
      rows: [
        [cell('标题'), cell('标题'), cell(null, 'empty')],
        [
          { ...cell('安全'), link: 'https://example.org/row' },
          { ...cell('危险'), link: 'javascript:alert(1)' },
          cell('第三列'),
        ],
      ],
    };
    const table = buildTable(sheet, 1);
    const bytes = exportWorkbook(table, table.rows, ['col-1', 'col-0', 'col-2']);
    const parsed = readWorkbook(bufferOf(bytes), 'links.xlsx');
    expect(parsed.sheets[0].rows[0][0].text).toBe('标题');
    expect(parsed.sheets[0].rows[0][1].text).toBe('标题');
    expect(parsed.sheets[0].rows[0][2]).toBeUndefined();
    expect(parsed.sheets[0].rows[1][0].link).toBeUndefined();
    expect(parsed.sheets[0].rows[1][1].link).toBe('https://example.org/row');
    expect(table.rows[0].cells[1].link).toBe('javascript:alert(1)');
  });

  it('queries and exports 10,000 synthetic records', () => {
    const rows: DataCell[][] = [[cell('ID'), cell('数量'), cell('分类')]];
    for (let index = 0; index < 10_000; index++) {
      rows.push([cell(`ID-${index}`), cell(index, 'number'), cell(index % 2 ? '奇' : '偶')]);
    }
    const table = buildTable({ name: '合成', firstRow: 1, date1904: false, rows }, 1);
    const selected = queryRows(table, [rule('col-1', 'range', { min: '5,000', max: '5,099' })], 'all', '', null);
    expect(selected).toHaveLength(100);
    const bytes = exportWorkbook(table, selected, ['col-0', 'col-1']);
    const back = readWorkbook(bufferOf(bytes), 'ten-thousand.xlsx');
    expect(buildTable(back.sheets[0], 1).rows).toHaveLength(100);
  });
});
