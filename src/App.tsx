import { useEffect, useMemo, useRef, useState } from 'react';
import type { ChangeEvent, DragEvent, KeyboardEvent, ReactNode } from 'react';
import { buildTable, parseNumeric, queryRows, validateFilter } from './model';
import { createDemoBytes, createDemoWorkbook } from './demo';
import { downloadWorkbook, exportWorkbook, readWorkbook } from './workbook';
import type { Column, DataCell, DataRow, FilterOp, FilterRule, MatchMode, SortRule, TableData, ViewPreset, WorkbookData } from './types';

const STORAGE_KEY = 'excel-workbench-presets-v1';
const PAGE_SIZE = 50;
const MAX_BYTES = 25 * 1024 * 1024;
const filterOptions: { value: FilterOp; label: string }[] = [
  { value: 'range', label: '数值范围' }, { value: 'dateRange', label: '日期范围' },
  { value: 'contains', label: '包含文字' }, { value: 'notContains', label: '不包含文字' },
  { value: 'equals', label: '等于' }, { value: 'oneOf', label: '属于分类' },
  { value: 'empty', label: '为空' }, { value: 'notEmpty', label: '非空' },
];
const numberFormat = new Intl.NumberFormat('zh-CN');

function isFilterOp(value: unknown): value is FilterOp {
  return filterOptions.some((item) => item.value === value);
}

function emptyRule(column: Column): FilterRule {
  return {
    id: typeof crypto !== 'undefined' && 'randomUUID' in crypto ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`,
    columnId: column.id,
    op: column.kind === 'date' ? 'dateRange' : column.kind === 'number' || column.kind === 'percent' ? 'range' : 'contains',
    value: '', min: '', max: '', values: [],
  };
}

function isIncomplete(rule: FilterRule): boolean {
  if (rule.op === 'range' || rule.op === 'dateRange') return !rule.min.trim() && !rule.max.trim();
  if (rule.op === 'oneOf') return rule.values.length === 0;
  if (rule.op === 'empty' || rule.op === 'notEmpty') return false;
  return !rule.value.trim();
}

function safeLink(link?: string): string | null {
  if (!link) return null;
  try {
    const url = new URL(link);
    return url.protocol === 'https:' || url.protocol === 'http:' ? url.href : null;
  } catch { return null; }
}

function displayCell(cell?: DataCell): string {
  if (!cell || cell.kind === 'empty') return '—';
  if (cell.kind === 'error') return cell.text || '公式结果不可用';
  return cell.text || (cell.value === null ? '—' : String(cell.value));
}

function compactSize(bytes: number): string {
  return bytes < 1024 * 1024 ? `${Math.max(1, Math.round(bytes / 1024))} KB` : `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function readPresets(): ViewPreset[] {
  try {
    const raw = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]') as unknown;
    if (!Array.isArray(raw)) return [];
    return raw.filter((item): item is ViewPreset => {
      if (!item || typeof item !== 'object') return false;
      const preset = item as Partial<ViewPreset>;
      return preset.version === 1 && typeof preset.name === 'string' && Array.isArray(preset.signature)
        && Array.isArray(preset.filters) && Array.isArray(preset.visibleColumns)
        && (preset.matchMode === 'all' || preset.matchMode === 'any') && typeof preset.search === 'string';
    });
  } catch { return []; }
}

function headerChoices(workbook: WorkbookData, index: number, current: number): number[] {
  const sheet = workbook.sheets[index];
  if (!sheet) return [];
  const quick = sheet.rows.length ? sheet.rows.slice(0, 20).map((_, offset) => sheet.firstRow + offset) : [sheet.firstRow];
  return quick.includes(current) ? quick : [...quick, current].sort((a, b) => a - b);
}

function columnIds(sheet: WorkbookData['sheets'][number] | undefined, headerRow: number): string[] {
  if (!sheet?.rows.length) return [];
  try { return buildTable(sheet, headerRow).columns.map((column) => column.id); }
  catch { return []; }
}

function FilterCard({ rule, column, onChange, onRemove, error }: {
  rule: FilterRule; column: Column; onChange: (next: FilterRule) => void; onRemove: () => void; error: string | null;
}) {
  const [categorySearch, setCategorySearch] = useState('');
  const categories = column.categories.filter((item) => item.value.toLocaleLowerCase().includes(categorySearch.toLocaleLowerCase())).slice(0, 40);
  const numeric = rule.op === 'range';
  const date = rule.op === 'dateRange';
  const ticks = column.histogram;
  const maxTick = Math.max(1, ...ticks);
  const low = numeric && rule.min.trim() ? parseNumeric(rule.min) : null;
  const high = numeric && rule.max.trim() ? parseNumeric(rule.max) : null;
  return <div className="filter-card">
    <div className="filter-card-title"><span title={column.label}>{column.label}</span><button className="icon-button" type="button" onClick={onRemove} aria-label={`删除${column.label}条件`}>×</button></div>
    <div className="filter-meta"><span>{column.kind === 'date' ? '日期列' : column.kind === 'percent' ? '百分比列' : column.kind === 'number' ? '数值列' : '文本列'}</span><span>有效 {numberFormat.format(column.validCount)} · 空 {numberFormat.format(column.emptyCount)}</span></div>
    <label className="sr-only" htmlFor={`op-${rule.id}`}>筛选方式</label>
    <select id={`op-${rule.id}`} value={rule.op} onChange={(event) => onChange({ ...rule, op: event.target.value as FilterOp, value: '', min: '', max: '', values: [] })}>
      {filterOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
    </select>
    {(numeric || date) && <>
      {numeric && ticks.length > 0 && <div className="distribution" role="img" aria-label={`${column.label}的数值分布`}>
        {ticks.map((count, index) => {
          const binMin = column.min === null || column.max === null ? null : column.min + (column.max - column.min) * index / ticks.length;
          const binMax = column.min === null || column.max === null ? null : column.min + (column.max - column.min) * (index + 1) / ticks.length;
          const inRange = (low === null || binMax === null || binMax >= low) && (high === null || binMin === null || binMin <= high);
          return <i key={index} className={inRange ? 'in-range' : ''} style={{ height: `${Math.max(4, (count / maxTick) * 100)}%` }} />;
        })}
      </div>}
      {numeric && column.min !== null && column.max !== null && <div className="range-hints"><span>{numberFormat.format(column.min)}</span><span>{numberFormat.format(column.max)}{column.kind === 'percent' ? '%' : ''}</span></div>}
      <div className="range-fields">
        <label>最小值<input type={date ? 'date' : 'text'} inputMode={date ? undefined : 'decimal'} value={rule.min} placeholder={date ? undefined : '不限'} onChange={(event) => onChange({ ...rule, min: event.target.value })} /></label>
        <span aria-hidden="true">—</span>
        <label>最大值<input type={date ? 'date' : 'text'} inputMode={date ? undefined : 'decimal'} value={rule.max} placeholder={date ? undefined : '不限'} onChange={(event) => onChange({ ...rule, max: event.target.value })} /></label>
      </div>
      {numeric && <small className="field-hint">可输入 12,000、2.5万、3M；百分比按显示值填写。</small>}
    </>}
    {['contains', 'notContains', 'equals'].includes(rule.op) && <label className="sr-only" htmlFor={`value-${rule.id}`}>筛选文字</label>}
    {['contains', 'notContains', 'equals'].includes(rule.op) && <input id={`value-${rule.id}`} value={rule.value} placeholder="输入文字" onChange={(event) => onChange({ ...rule, value: event.target.value })} />}
    {rule.op === 'oneOf' && <div className="category-control">
      <label className="sr-only" htmlFor={`cat-${rule.id}`}>搜索分类值</label>
      <input id={`cat-${rule.id}`} value={categorySearch} placeholder="搜索或输入分类值" onChange={(event) => setCategorySearch(event.target.value)} />
      <div className="category-options">
        {categories.map((item) => <label key={item.value}><input type="checkbox" checked={rule.values.includes(item.value)} onChange={(event) => onChange({ ...rule, values: event.target.checked ? [...rule.values, item.value] : rule.values.filter((value) => value !== item.value) })} /><span title={item.value}>{item.value}</span><em>{numberFormat.format(item.count)}</em></label>)}
      </div>
      {categorySearch.trim() && !rule.values.includes(categorySearch.trim()) && <button className="text-button" type="button" onClick={() => { onChange({ ...rule, values: [...rule.values, categorySearch.trim()] }); setCategorySearch(''); }}>加入“{categorySearch.trim()}”</button>}
      {rule.values.length > 0 && <div className="selected-values">已选 {rule.values.length} 项 · <button type="button" className="text-button" onClick={() => onChange({ ...rule, values: [] })}>清除分类</button></div>}
    </div>}
    {error && <p className={isIncomplete(rule) ? 'filter-note' : 'field-error'} role={isIncomplete(rule) ? undefined : 'alert'}>{isIncomplete(rule) ? '填写条件后生效' : error}</p>}
  </div>;
}

function Modal({ title, onClose, children, className = '' }: { title: string; onClose: () => void; children: ReactNode; className?: string }) {
  const closeRef = useRef<HTMLButtonElement>(null);
  const dialogRef = useRef<HTMLElement>(null);
  const onCloseRef = useRef(onClose);
  onCloseRef.current = onClose;
  useEffect(() => {
    const prior = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    closeRef.current?.focus();
    const handleKey = (event: globalThis.KeyboardEvent) => {
      if (event.key === 'Escape') { onCloseRef.current(); return; }
      if (event.key !== 'Tab') return;
      const items = dialogRef.current?.querySelectorAll<HTMLElement>('button:not(:disabled),input:not(:disabled),select:not(:disabled),a[href],[tabindex="0"]');
      if (!items?.length) return;
      const first = items[0], last = items[items.length - 1];
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
    };
    document.addEventListener('keydown', handleKey);
    return () => { document.removeEventListener('keydown', handleKey); prior?.focus(); };
  }, []);
  return <div className="modal-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}>
    <section ref={dialogRef} role="dialog" aria-modal="true" aria-label={title} className={`modal ${className}`}>
      <div className="modal-head"><h2>{title}</h2><button type="button" className="icon-button" ref={closeRef} aria-label="关闭对话框" onClick={onClose}>×</button></div>
      {children}
    </section>
  </div>;
}

export default function App() {
  const fileRef = useRef<HTMLInputElement>(null);
  const importSeq = useRef(0);
  const [workbook, setWorkbook] = useState<WorkbookData | null>(null);
  const [sheetIndex, setSheetIndex] = useState(0);
  const [headerRow, setHeaderRow] = useState(1);
  const [filters, setFilters] = useState<FilterRule[]>([]);
  const [matchMode, setMatchMode] = useState<MatchMode>('all');
  const [search, setSearch] = useState('');
  const [sort, setSort] = useState<SortRule | null>(null);
  const [visibleColumns, setVisibleColumns] = useState<string[]>([]);
  const [selected, setSelected] = useState<number[]>([]);
  const [page, setPage] = useState(1);
  const [filterOpen, setFilterOpen] = useState(false);
  const [columnOpen, setColumnOpen] = useState(false);
  const [dialog, setDialog] = useState<'export' | 'presets' | 'detail' | 'header' | null>(null);
  const [customHeader, setCustomHeader] = useState('');
  const [detailRow, setDetailRow] = useState<DataRow | null>(null);
  const [exportScope, setExportScope] = useState<'matches' | 'selected'>('matches');
  const [exportColumns, setExportColumns] = useState<'all' | 'visible'>('all');
  const [exportName, setExportName] = useState('筛选结果');
  const [presetName, setPresetName] = useState('');
  const [presets, setPresets] = useState<ViewPreset[]>(readPresets);
  const [status, setStatus] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const [dragging, setDragging] = useState(false);

  const sheet = workbook?.sheets[sheetIndex];
  const table = useMemo(() => {
    if (!sheet) return null;
    try { return buildTable(sheet, headerRow); }
    catch { return null; }
  }, [sheet, headerRow]);
  const columnMap = useMemo(() => new Map(table?.columns.map((column) => [column.id, column]) || []), [table]);
  const labels = useMemo(() => {
    const used = new Map<string, number>();
    return (table?.columns || []).map((column) => {
      const base = column.label.trim() || `未命名列 ${column.index + 1}`;
      const count = (used.get(base) || 0) + 1;
      used.set(base, count);
      return { id: column.id, label: count > 1 ? `${base}（第 ${count} 列）` : base };
    });
  }, [table]);
  const nameOf = (id: string) => labels.find((item) => item.id === id)?.label || id;
  const validation = useMemo(() => filters.map((rule) => {
    const column = columnMap.get(rule.columnId);
    return column ? validateFilter(rule, column) : '列已不存在';
  }), [filters, columnMap]);
  const badFilter = validation.some((item, index) => item && !isIncomplete(filters[index]));
  const pendingFilter = filters.some(isIncomplete);
  const activeFilters = filters.filter((rule, index) => !validation[index] && !isIncomplete(rule));
  const query = useMemo(() => {
    if (!table) return { rows: [] as DataRow[], error: '' };
    try { return { rows: queryRows(table, activeFilters, matchMode, search, sort), error: '' }; }
    catch (caught) { return { rows: [] as DataRow[], error: caught instanceof Error ? caught.message : '筛选条件无法处理' }; }
  }, [table, filters, matchMode, search, sort]);
  const rows = query.rows;
  const shownColumns = (table?.columns || []).filter((column) => visibleColumns.includes(column.id));
  const pageCount = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
  const safePage = Math.min(page, pageCount);
  const pageRows = rows.slice((safePage - 1) * PAGE_SIZE, safePage * PAGE_SIZE);
  const matchedSelection = rows.filter((row) => selected.includes(row.id));
  const pageAllSelected = pageRows.length > 0 && pageRows.every((row) => selected.includes(row.id));
  const signature = table?.columns.map((column) => `${column.id}:${column.label}:${column.kind}`) || [];
  const headerMin = sheet?.firstRow ?? 1;
  const headerMax = sheet?.rows.length ? headerMin + sheet.rows.length - 1 : headerMin;
  const headerValue = customHeader.trim();
  const headerInputError = !/^\d+$/.test(headerValue) ? '请输入整数行号' : Number(headerValue) < headerMin || Number(headerValue) > headerMax ? `行号须在 ${headerMin}–${headerMax} 之间` : '';

  useEffect(() => { setPage(1); }, [filters, matchMode, search, sort]);
  useEffect(() => { if (page > pageCount) setPage(pageCount); }, [page, pageCount]);

  function resetView() {
    setFilters([]); setMatchMode('all'); setSearch(''); setSort(null); setSelected([]); setPage(1);
    setExportScope('matches');
    setColumnOpen(false); setDialog(null); setStatus('');
  }

  function useWorkbook(next: WorkbookData, message: string) {
    setWorkbook(next); setSheetIndex(0); setHeaderRow(next.sheets[0]?.firstRow || 1);
    setVisibleColumns(columnIds(next.sheets[0], next.sheets[0]?.firstRow || 1));
    resetView(); setError(''); setStatus(message);
  }

  async function importFile(file?: File) {
    if (!file) return;
    if (!/\.(xlsx|xls|csv)$/i.test(file.name)) { setError('请选择 .xlsx、.xls 或 .csv 文件。'); return; }
    if (file.size > MAX_BYTES) { setError('文件超过 25 MB。请先拆分文件后重试。'); return; }
    const seq = ++importSeq.current;
    setLoading(true); setStatus('正在读取文件…'); setError('');
    try {
      const data = await file.arrayBuffer();
      if (seq !== importSeq.current) return;
      const next = readWorkbook(data, file.name);
      if (!next.sheets.length) throw new Error('文件中没有可读取的工作表');
      useWorkbook(next, `已载入 ${file.name}`);
    } catch (caught) {
      if (seq === importSeq.current) { setError(caught instanceof Error ? caught.message : '文件解析失败，请检查格式。'); setStatus(''); }
    } finally { if (seq === importSeq.current) setLoading(false); }
  }

  function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    void importFile(event.target.files?.[0]);
    event.target.value = '';
  }

  function handleDrop(event: DragEvent<HTMLElement>) {
    event.preventDefault(); setDragging(false);
    void importFile(event.dataTransfer.files[0]);
  }

  function loadDemo() {
    importSeq.current++;
    try { useWorkbook(createDemoWorkbook(), '已载入虚构示例，可直接试用筛选。'); setLoading(false); }
    catch (caught) { setError(caught instanceof Error ? caught.message : '示例无法载入'); }
  }

  function switchSheet(index: number) {
    if (!workbook || index === sheetIndex) return;
    setSheetIndex(index); setHeaderRow(workbook.sheets[index]?.firstRow || 1);
    const nextSheet = workbook.sheets[index];
    setVisibleColumns(columnIds(nextSheet, nextSheet?.firstRow || 1)); resetView();
  }

  function updateHeader(value: number) {
    if (value === headerRow) return;
    setHeaderRow(value); setVisibleColumns(columnIds(sheet, value)); resetView();
  }

  function applyCustomHeader() {
    if (headerInputError) return;
    const value = Number(headerValue);
    if (value === headerRow) setDialog(null);
    else updateHeader(value);
  }

  function updateRule(next: FilterRule) { setFilters((current) => current.map((rule) => rule.id === next.id ? next : rule)); }

  function toggleSort(columnId: string) {
    setSort((current) => !current || current.columnId !== columnId ? { columnId, direction: 'asc' } : current.direction === 'asc' ? { columnId, direction: 'desc' } : null);
  }

  function togglePageSelection() {
    const ids = pageRows.map((row) => row.id);
    setSelected((current) => pageAllSelected ? current.filter((id) => !ids.includes(id)) : [...new Set([...current, ...ids])]);
  }

  function doExport() {
    if (!table || badFilter || pendingFilter || query.error) return;
    const exportRows = exportScope === 'selected' ? matchedSelection : rows;
    const ids = exportColumns === 'visible' ? shownColumns.map((column) => column.id) : table.columns.map((column) => column.id);
    if (!exportRows.length || !ids.length) { setError('导出范围为空，请调整行或显示列。'); setDialog(null); return; }
    try {
      downloadWorkbook(exportWorkbook(table, exportRows, ids), (exportName.trim() || '筛选结果').replace(/[\\/:*?"<>|]/g, '_').replace(/\.xlsx$/i, '') + '.xlsx');
      setStatus(`已导出 ${numberFormat.format(exportRows.length)} 行。`); setError(''); setDialog(null);
    } catch (caught) { setError(caught instanceof Error ? caught.message : '导出失败'); setDialog(null); }
  }

  function persistPresets(next: ViewPreset[]) {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(next)); setPresets(next); return true; }
    catch { setError('浏览器无法保存方案，请检查本地存储设置。'); return false; }
  }

  function savePreset() {
    if (!table || !presetName.trim() || badFilter || pendingFilter) return;
    const name = presetName.trim().slice(0, 50);
    const preset: ViewPreset = { version: 1, name, signature, filters, matchMode, search, visibleColumns, sort };
    if (persistPresets([...presets.filter((item) => item.name !== name), preset])) { setStatus(`已保存方案“${name}”。`); setPresetName(''); }
  }

  function restorePreset(preset: ViewPreset) {
    if (!table || preset.signature.length !== signature.length || preset.signature.some((value, index) => value !== signature[index])) {
      setError(`方案“${preset.name}”与当前列不兼容，请打开对应工作表。`); return;
    }
    const ids = new Set(table.columns.map((column) => column.id));
    if (!preset.filters.every((rule) => rule && typeof rule.id === 'string' && ids.has(rule.columnId) && isFilterOp(rule.op)
      && typeof rule.value === 'string' && typeof rule.min === 'string' && typeof rule.max === 'string' && Array.isArray(rule.values) && rule.values.every((value) => typeof value === 'string'))
      || !preset.visibleColumns.every((id) => ids.has(id)) || (preset.sort && !ids.has(preset.sort.columnId))) {
      setError(`方案“${preset.name}”内容无效，无法恢复。`); return;
    }
    setFilters(preset.filters); setMatchMode(preset.matchMode); setSearch(preset.search); setVisibleColumns(preset.visibleColumns);
    setSort(preset.sort); setSelected([]); setPage(1); setDialog(null); setError(''); setStatus(`已恢复方案“${preset.name}”。`);
  }

  function showDetail(row: DataRow) { setDetailRow(row); setDialog('detail'); }
  async function copyCell(text: string) {
    try { await navigator.clipboard.writeText(text); setStatus('已复制单元格完整内容。'); }
    catch { setError('复制失败。可在详情中选中文字手动复制。'); }
  }

  const exportDisabled = !table || rows.length === 0 || shownColumns.length === 0 || badFilter || pendingFilter || !!query.error;
  const filterMessage = badFilter ? '请修正标红的筛选条件。' : pendingFilter ? '有未填写的条件，请填写或删除后导出。' : query.error;

  return <div className="app-shell" onDragOver={(event) => { if (event.dataTransfer.types.includes('Files')) { event.preventDefault(); setDragging(true); } }} onDragLeave={(event) => { if (!event.currentTarget.contains(event.relatedTarget as Node)) setDragging(false); }} onDrop={handleDrop}>
    <input ref={fileRef} className="sr-only" type="file" accept=".xlsx,.xls,.csv" onChange={handleFileChange} aria-label="选择 Excel 或 CSV 文件" />
    {dragging && <div className="drop-overlay">松开以导入表格</div>}
    <header className="topbar">
      <div className="brand"><span className="brand-mark" aria-hidden="true"><i /><i /><i /></span><div><strong>表格工作台</strong><small>本地浏览 · 筛选 · 导出</small></div></div>
      <div className="top-actions"><button type="button" className="button secondary" onClick={() => fileRef.current?.click()}>导入文件</button><button type="button" className="button ghost" onClick={loadDemo}>试用示例</button><button type="button" className="button primary" disabled={exportDisabled} onClick={() => { if (!matchedSelection.length) setExportScope('matches'); setExportName(`${workbook?.name.replace(/\.[^.]+$/, '') || '筛选结果'}-筛选结果`); setDialog('export'); }}>导出 Excel</button></div>
    </header>
    {(error || status) && <div className={error ? 'notice error-notice' : 'notice success-notice'} role={error ? 'alert' : 'status'}><span>{error || status}</span><button type="button" className="icon-button" aria-label="关闭提示" onClick={() => { setError(''); setStatus(''); }}>×</button></div>}
    {!workbook ? <main className="empty-main"><section className="import-zone" onClick={() => fileRef.current?.click()} onKeyDown={(event: KeyboardEvent<HTMLElement>) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); fileRef.current?.click(); } }} role="button" tabIndex={0} aria-label="导入 Excel 或 CSV 文件">
      <div className="import-art" aria-hidden="true"><span /><span /><span /><span /></div>
      <p className="eyebrow">从文件开始</p><h1>把长表展开，再找到需要的行。</h1><p>拖入 Excel 或 CSV，或点击选择文件。筛选、搜索和导出都在此浏览器中完成。</p>
      <span className="button primary">选择文件</span><small>支持 .xlsx / .xls / .csv · 最大 25 MB · 单表最多 10 万行、250 列</small>
    </section><div className="empty-footer"><button type="button" className="text-button" onClick={loadDemo}>先用虚构示例看看 →</button><span>文件不会上传。主动保存的方案会在当前浏览器保留表头、筛选值和搜索词，用完可删除。</span></div></main> : <main className="workbench">
      <section className="workspace-meta">
        <div className="file-identity"><span className="file-badge">{workbook.name.toLowerCase().endsWith('.csv') ? 'CSV' : 'XLS'}</span><div><strong title={workbook.name}>{workbook.name}</strong><small>{compactSize(workbook.size)} · {workbook.sheets.length} 个工作表 · 当前 {numberFormat.format(table?.rows.length || 0)} 行</small></div></div>
        <div className="source-controls"><label>工作表<select value={sheetIndex} onChange={(event) => switchSheet(Number(event.target.value))}>{workbook.sheets.map((item, index) => <option key={`${item.name}-${index}`} value={index}>{item.name}</option>)}</select></label><label>表头所在行<select value={headerRow} onChange={(event) => { if (event.target.value === 'custom') { setCustomHeader(String(headerRow)); setDialog('header'); } else updateHeader(Number(event.target.value)); }}>{headerChoices(workbook, sheetIndex, headerRow).map((value) => <option key={value} value={value}>第 {value} 行</option>)}<option value="custom">自定义行号…</option></select></label></div>
      </section>
      {[...workbook.warnings, ...(table?.warnings || [])].length > 0 && <div className="warning-strip" role="status">{[...workbook.warnings, ...(table?.warnings || [])].join('；')}</div>}
      <div className="workspace-grid">
        <aside className={`filter-panel ${filterOpen ? 'is-open' : ''}`} aria-label="筛选条件">
          <div className="panel-head"><div><p className="eyebrow">筛选器</p><h2>缩小结果范围</h2></div><button type="button" className="text-button" onClick={() => { setFilters([]); setMatchMode('all'); }}>清空</button></div>
          <div className="filter-count"><strong>{numberFormat.format(rows.length)}</strong><span>/ {numberFormat.format(table?.rows.length || 0)} 行匹配</span></div>
          <div className="mode-control"><span>多个条件</span><div role="group" aria-label="条件组合"><button type="button" className={matchMode === 'all' ? 'active' : ''} onClick={() => setMatchMode('all')}>全部满足</button><button type="button" className={matchMode === 'any' ? 'active' : ''} onClick={() => setMatchMode('any')}>任意满足</button></div></div>
          <div className="filter-list">{filters.map((rule, index) => { const column = columnMap.get(rule.columnId); return column && <FilterCard key={rule.id} rule={rule} column={{ ...column, label: nameOf(column.id) }} onChange={updateRule} onRemove={() => setFilters((current) => current.filter((item) => item.id !== rule.id))} error={validation[index]} />; })}{filters.length === 0 && <p className="filter-empty">还没有条件。添加条件后，右侧表格会立即更新。</p>}</div>
          {table && table.columns.length > 0 && <label className="add-filter">添加条件<select value="" onChange={(event) => { const column = columnMap.get(event.target.value); if (column) setFilters((current) => [...current, emptyRule(column)]); }}><option value="">选择一列…</option>{table.columns.map((column) => <option key={column.id} value={column.id}>{nameOf(column.id)}</option>)}</select></label>}
          {filterMessage && <p className="panel-alert" role="alert">{filterMessage}</p>}
          <div className="panel-bottom"><button type="button" className="button secondary full" disabled={!table?.columns.length} onClick={() => setDialog('presets')}>保存或恢复方案</button><small>主动保存的方案会在本浏览器保留表头、筛选值、搜索词和视图设置；不保存完整工作簿。</small></div>
        </aside>
        <section className="results-panel">
          <div className="result-toolbar"><div className="result-heading"><button className="button secondary mobile-filter" type="button" onClick={() => setFilterOpen((value) => !value)} aria-expanded={filterOpen}>筛选 {filters.length ? `(${filters.length})` : ''}</button><div><p className="eyebrow">数据预览</p><h2>{table?.sheetName || '当前工作表'} <span>{numberFormat.format(rows.length)} 行</span></h2></div></div><div className="table-tools"><label className="search-field"><span className="sr-only">搜索全表</span><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="10.8" cy="10.8" r="6.6"/><path d="m16 16 5 5"/></svg><input value={search} placeholder="搜索全表内容" onChange={(event) => setSearch(event.target.value)} /></label><div className="columns-wrap"><button type="button" className="button secondary" onClick={() => setColumnOpen((value) => !value)} aria-expanded={columnOpen}>显示列 {shownColumns.length}/{table?.columns.length || 0}</button>{columnOpen && <div className="columns-menu"><div className="columns-menu-head"><strong>显示列</strong><button type="button" className="text-button" onClick={() => setVisibleColumns(table?.columns.map((column) => column.id) || [])}>显示全部</button></div>{table?.columns.map((column) => <label key={column.id}><input type="checkbox" checked={visibleColumns.includes(column.id)} onChange={(event) => setVisibleColumns((current) => event.target.checked ? [...current, column.id] : current.filter((id) => id !== column.id))} />{nameOf(column.id)}</label>)}</div>}</div></div></div>
          {table && table.columns.length > 0 && visibleColumns.length === 0 && <div className="inline-hint" role="status">所有列已隐藏。请在“显示列”中勾选至少一列。</div>}
          {table && table.columns.length > 0 && <div className="table-scroll" role="region" aria-label="可横向滚动的数据表" tabIndex={0}><table><thead><tr><th className="sticky-first check-header"><input type="checkbox" aria-label="勾选当前页全部行" checked={pageAllSelected} onChange={togglePageSelection} disabled={!pageRows.length} /></th>{shownColumns.map((column) => <th key={column.id} scope="col"><button type="button" className="sort-button" onClick={() => toggleSort(column.id)} title={`按${nameOf(column.id)}排序`}>{nameOf(column.id)}<span aria-hidden="true">{sort?.columnId === column.id ? sort.direction === 'asc' ? '↑' : '↓' : '↕'}</span></button></th>)}<th className="detail-heading">详情</th></tr></thead><tbody>{pageRows.map((row) => <tr key={row.id}><td className="sticky-first check-cell"><input type="checkbox" aria-label={`勾选源表第 ${row.sourceRow} 行`} checked={selected.includes(row.id)} onChange={(event) => setSelected((current) => event.target.checked ? [...current, row.id] : current.filter((id) => id !== row.id))} /><span>{row.sourceRow}</span></td>{shownColumns.map((column) => { const cell = row.cells[column.index]; const href = safeLink(cell?.link); return <td key={column.id} title={cell?.text || ''} className={cell?.kind === 'empty' ? 'empty-cell' : cell?.kind === 'error' ? 'error-cell' : column.kind === 'number' || column.kind === 'percent' ? 'numeric-cell' : ''}>{href ? <a href={href} target="_blank" rel="noopener noreferrer" onClick={(event) => event.stopPropagation()}>{displayCell(cell)}</a> : displayCell(cell)}</td>; })}<td><button type="button" className="row-detail" onClick={() => showDetail(row)} aria-label={`查看源表第 ${row.sourceRow} 行详情`}>查看</button></td></tr>)}</tbody></table></div>}
          {(!table || table.columns.length === 0) && <div className="result-empty"><h3>当前工作表没有可用列</h3><p>试着调整表头所在行，或导入包含数据的工作表。</p></div>}
          {table && table.columns.length > 0 && rows.length === 0 && <div className="result-empty"><h3>{table.rows.length === 0 ? '表头下没有数据行' : '没有匹配的记录'}</h3><p>{table.rows.length === 0 ? '更换工作表或调整表头所在行。' : '调整筛选条件或清空搜索词后重试。'}</p></div>}
          <div className="table-footer"><span>已勾选 {numberFormat.format(selected.length)} 行 · 当前匹配 {numberFormat.format(matchedSelection.length)} 行</span><div className="pagination"><label>每页 {PAGE_SIZE} 行</label><button type="button" disabled={safePage <= 1} onClick={() => setPage(safePage - 1)} aria-label="上一页">‹</button><span>{safePage} / {pageCount}</span><button type="button" disabled={safePage >= pageCount} onClick={() => setPage(safePage + 1)} aria-label="下一页">›</button></div></div>
        </section>
      </div>
      <footer className="workbench-foot"><span>浏览器本地处理 · 最多 25 MB / 单表 10 万行 / 250 列</span><button type="button" className="text-button" onClick={() => downloadWorkbook(createDemoBytes(), '表格工作台-虚构示例.xlsx')}>下载虚构示例</button></footer>
    </main>}
    {loading && <div className="loading-screen" role="status"><span className="spinner" />正在解析文件…</div>}
    {dialog === 'export' && table && <Modal title="导出筛选结果" onClose={() => setDialog(null)}><div className="modal-body"><p className="modal-intro">按当前排序导出，覆盖所有匹配页。原文件不会改变。</p><fieldset><legend>导出行</legend><label><input type="radio" name="scope" checked={exportScope === 'matches'} onChange={() => setExportScope('matches')} />全部匹配行 <strong>{numberFormat.format(rows.length)}</strong></label><label><input type="radio" name="scope" checked={exportScope === 'selected'} onChange={() => setExportScope('selected')} disabled={matchedSelection.length === 0} />已勾选且仍匹配 <strong>{numberFormat.format(matchedSelection.length)}</strong></label></fieldset><fieldset><legend>导出列</legend><label><input type="radio" name="columns" checked={exportColumns === 'all'} onChange={() => setExportColumns('all')} />全部列 <strong>{table.columns.length}</strong></label><label><input type="radio" name="columns" checked={exportColumns === 'visible'} onChange={() => setExportColumns('visible')} disabled={!shownColumns.length} />当前显示列 <strong>{shownColumns.length}</strong></label></fieldset><label className="modal-field">文件名<input value={exportName} onChange={(event) => setExportName(event.target.value)} /></label><p className="field-hint">导出为数据表：公式使用已保存的结果值，不保留图表、宏和原排版。</p></div><div className="modal-actions"><button type="button" className="button secondary" onClick={() => setDialog(null)}>取消</button><button type="button" className="button primary" onClick={doExport} disabled={exportScope === 'selected' && matchedSelection.length === 0}>导出 .xlsx</button></div></Modal>}
    {dialog === 'presets' && <Modal title="筛选方案" onClose={() => setDialog(null)}><div className="modal-body"><p className="modal-intro">主动保存的方案会在当前浏览器保留表头、筛选值、搜索词和视图设置，不保存完整工作簿；用完可删除。恢复时需匹配当前列。</p><label className="modal-field">方案名称<input value={presetName} maxLength={50} placeholder="例如：本周待复核" onChange={(event) => setPresetName(event.target.value)} /></label><button className="button primary" type="button" onClick={savePreset} disabled={!presetName.trim() || badFilter || pendingFilter}>保存当前方案</button><div className="preset-list"><h3>已保存的方案</h3>{presets.length === 0 && <p>还没有保存方案。</p>}{presets.map((preset) => <div className="preset-item" key={preset.name}><strong>{preset.name}</strong><div><button type="button" className="text-button" onClick={() => restorePreset(preset)}>恢复</button><button type="button" className="text-button danger" onClick={() => persistPresets(presets.filter((item) => item.name !== preset.name))}>删除</button></div></div>)}</div></div></Modal>}
    {dialog === 'detail' && detailRow && table && <Modal title={`源表第 ${detailRow.sourceRow} 行`} onClose={() => setDialog(null)} className="detail-modal"><div className="detail-list">{table.columns.map((column) => { const cell = detailRow.cells[column.index]; const href = safeLink(cell?.link); return <div className="detail-item" key={column.id}><div><strong>{nameOf(column.id)}</strong><button type="button" className="text-button" onClick={() => void copyCell(cell?.text || '')}>复制</button></div><p>{displayCell(cell)}</p>{href && <a href={href} target="_blank" rel="noopener noreferrer">打开链接 ↗</a>}{cell?.kind === 'error' && <small>公式没有可用的缓存结果。</small>}</div>; })}</div></Modal>}
    {dialog === 'header' && <Modal title="自定义表头行" onClose={() => setDialog(null)}><div className="modal-body"><p className="modal-intro">可选源表第 {headerMin}–{headerMax} 行。</p><label className="modal-field">表头行号<input type="text" inputMode="numeric" value={customHeader} onChange={(event) => setCustomHeader(event.target.value)} onKeyDown={(event) => { if (event.key === 'Enter') applyCustomHeader(); }} /></label>{headerInputError && <p className="field-error" role="alert">{headerInputError}</p>}</div><div className="modal-actions"><button type="button" className="button secondary" onClick={() => setDialog(null)}>取消</button><button type="button" className="button primary" onClick={applyCustomHeader} disabled={!!headerInputError}>应用表头</button></div></Modal>}
  </div>;
}
