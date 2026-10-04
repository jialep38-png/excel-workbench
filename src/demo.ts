import type { WorkbookData } from './types';
import { readWorkbook } from './workbook';

interface DemoSheetJs {
  utils: {
    book_new: () => { Workbook?: { WBProps?: { date1904?: boolean } } };
    book_append_sheet: (book: object, sheet: Record<string, unknown>, name: string) => void;
    aoa_to_sheet: (rows: unknown[][]) => Record<string, unknown>;
    encode_cell: (address: { r: number; c: number }) => string;
  };
  write: (book: object, options: Record<string, unknown>) => Uint8Array | ArrayBuffer;
}

function api(): DemoSheetJs {
  const sheetJs = globalThis.XLSX;
  if (!sheetJs) throw new Error('本地 Excel 解析组件未加载');
  return sheetJs as unknown as DemoSheetJs;
}

/** Entirely invented records for trying the workbench without personal data. */
export function createDemoBytes(): Uint8Array {
  const xlsx = api();
  const book = xlsx.utils.book_new();
  const countries = ['中国', '美国', '英国', '日本', '德国', '法国', '新加坡', '加拿大'];
  const roles = ['科技测评', '生活方式', '旅行记录', '知识分享'];
  const first: unknown[][] = [['编号', '国家', '角色', '关注人数', '互动率', '记录日期', '简介', '备注']];
  for (let index = 1; index <= 128; index++) {
    const intro = index % 19 === 0
      ? `这是一条用于检查长文本展示的虚构介绍，第 ${index} 行。它包含多个句子、标点和足够长的内容，以便在行详情中完整查看和复制，不对应任何真实账号、姓名、社交平台或联系方式。`.repeat(3)
      : `虚构创作者 ${index}，分享${roles[(index - 1) % roles.length]}相关内容。`;
    first.push([
      `DEMO-${String(index).padStart(4, '0')}`,
      countries[(index - 1) % countries.length],
      roles[(index - 1) % roles.length],
      1250 + index * 1387,
      index % 17 === 0 ? null : ((index * 7) % 85 + 5) / 1000,
      45100 + index,
      intro,
      index % 11 === 0 ? null : index % 7 === 0 ? '待复核' : '演示数据',
    ]);
  }
  const creators = xlsx.utils.aoa_to_sheet(first);
  for (let row = 1; row < first.length; row++) {
    const percentage = creators[xlsx.utils.encode_cell({ r: row, c: 4 })] as Record<string, unknown> | undefined;
    const date = creators[xlsx.utils.encode_cell({ r: row, c: 5 })] as Record<string, unknown> | undefined;
    if (percentage) percentage.z = '0.0%';
    if (date) date.z = 'yyyy-mm-dd';
  }
  xlsx.utils.book_append_sheet(book, creators, '虚构创作者');

  const markets: unknown[][] = [
    ['地区', '预算', '目标完成率', '开始日期', '说明'],
    ['东亚', 24000, 0.72, 45200, '完全虚构的市场计划'],
    ['北美', 31000, 0.64, 45214, '用于演示第二工作表'],
    ['欧洲', null, null, 45228, '空值用于检查筛选'],
    ['东南亚', 18000, 0.81, 45242, '预算与日期仅为示例'],
  ];
  const second = xlsx.utils.aoa_to_sheet(markets);
  for (let row = 1; row < markets.length; row++) {
    const percentage = second[xlsx.utils.encode_cell({ r: row, c: 2 })] as Record<string, unknown> | undefined;
    const date = second[xlsx.utils.encode_cell({ r: row, c: 3 })] as Record<string, unknown> | undefined;
    if (percentage) percentage.z = '0%';
    if (date) date.z = 'yyyy-mm-dd';
  }
  xlsx.utils.book_append_sheet(book, second, '虚构市场计划');
  const written = xlsx.write(book, { bookType: 'xlsx', type: 'array', compression: true });
  return written instanceof Uint8Array ? written : new Uint8Array(written);
}

export function createDemoWorkbook(): WorkbookData {
  const bytes = createDemoBytes();
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  return readWorkbook(buffer, '虚构示例.xlsx');
}
