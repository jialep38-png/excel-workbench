export type Value = string | number | boolean | null;
export type ColumnKind = 'number' | 'percent' | 'date' | 'text';
export type CellKind = ColumnKind | 'boolean' | 'empty' | 'error';

export interface DataCell {
  value: Value;
  text: string;
  kind: CellKind;
  numeric: number | null;
  date: string | null;
  format?: string;
  link?: string;
  formula?: boolean;
}

export interface DataRow {
  id: number;
  sourceRow: number;
  cells: DataCell[];
}

export interface Column {
  id: string;
  index: number;
  label: string;
  kind: ColumnKind;
  emptyCount: number;
  validCount: number;
  min: number | null;
  max: number | null;
  categories: { value: string; count: number }[];
  histogram: number[];
}

export interface SheetData {
  name: string;
  rows: DataCell[][];
  firstRow: number;
  date1904: boolean;
}

export interface WorkbookData {
  name: string;
  sheets: SheetData[];
  size: number;
  warnings: string[];
}

export interface TableData {
  sheetName: string;
  headerRow: number;
  columns: Column[];
  headerCells: DataCell[];
  rows: DataRow[];
  date1904: boolean;
  warnings: string[];
}

export type FilterOp = 'range' | 'dateRange' | 'contains' | 'notContains' | 'equals' | 'oneOf' | 'empty' | 'notEmpty';

export interface FilterRule {
  id: string;
  columnId: string;
  op: FilterOp;
  value: string;
  min: string;
  max: string;
  values: string[];
}

export type MatchMode = 'all' | 'any';

export interface SortRule {
  columnId: string;
  direction: 'asc' | 'desc';
}

export interface ViewPreset {
  version: 1;
  name: string;
  signature: string[];
  filters: FilterRule[];
  matchMode: MatchMode;
  search: string;
  visibleColumns: string[];
  sort: SortRule | null;
}
