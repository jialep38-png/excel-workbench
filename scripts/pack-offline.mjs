import { readFile, writeFile, mkdir, copyFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, resolve, relative, sep } from 'node:path';

const projectDir = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const distDir = resolve(projectDir, 'dist');
const releaseDir = resolve(projectDir, 'release');
const outputName = 'Excel筛选工作台.html';

function assetPath(reference) {
  const path = resolve(distDir, reference.replace(/^\.\//, ''));
  const local = relative(distDir, path);
  if (local.startsWith(`..${sep}`) || local === '..') throw new Error('构建资源超出 dist 目录');
  return path;
}

function inlineScript(source) {
  return source.replace(/<\/script/gi, '<\\/script');
}

function inlineStyle(source) {
  return source.replace(/<\/style/gi, '<\\/style');
}

let html = await readFile(resolve(distDir, 'index.html'), 'utf8');
const scriptTags = [...html.matchAll(/<script\b([^>]*)\bsrc="([^"]+)"([^>]*)><\/script>/g)];
const styleTags = [...html.matchAll(/<link\b[^>]*rel="stylesheet"[^>]*href="([^"]+)"[^>]*>/g)];
const preloadTags = [...html.matchAll(/<link\b[^>]*rel="modulepreload"[^>]*>/g)];
const replacements = [];
if (scriptTags.length !== 2) throw new Error(`预期应用和 SheetJS 两个脚本，实际 ${scriptTags.length}`);

for (const tag of scriptTags) {
  const source = await readFile(assetPath(tag[2]), 'utf8');
  const isModule = /type="module"/.test(tag[1] + tag[3]);
  if (isModule && /(?:\bimport\s*\(|\bfrom\s*["']\.\/)/.test(source)) {
    throw new Error('离线构建出现额外模块依赖，请保持单入口且不使用动态导入');
  }
  replacements.push({ start: tag.index, length: tag[0].length, value: `<script${isModule ? ' type="module"' : ''}>\n${inlineScript(source)}\n</script>` });
}

for (const tag of styleTags) {
  replacements.push({ start: tag.index, length: tag[0].length, value: `<style>\n${inlineStyle(await readFile(assetPath(tag[1]), 'utf8'))}\n</style>` });
}
for (const tag of preloadTags) replacements.push({ start: tag.index, length: tag[0].length, value: '' });
for (const replacement of replacements.sort((left, right) => right.start - left.start)) {
  // Slice the original markup so dollar sequences in JavaScript stay literal.
  html = html.slice(0, replacement.start) + replacement.value + html.slice(replacement.start + replacement.length);
}

// Include full runtime licenses so the HTML remains distributable on its own.
const licenseFiles = ['LICENSE', 'THIRD_PARTY_NOTICES.md', 'public/vendor/LICENSE.sheetjs', 'public/vendor/LICENSE.react'];
const notices = (await Promise.all(licenseFiles.map((path) => readFile(resolve(projectDir, path), 'utf8')))).join('\n\n');
html = html.replace('</head>', () => `<script type="text/plain" id="third-party-licenses">\n${inlineScript(notices)}\n</script>\n</head>`);

const markup = html
  .replace(/(<script\b[^>]*>)[\s\S]*?<\/script>/gi, '$1</script>')
  .replace(/(<style\b[^>]*>)[\s\S]*?<\/style>/gi, '$1</style>');
if (/<(?:script|link)\b[^>]*(?:src|href)=/.test(markup)) throw new Error('离线 HTML 仍含外部资源引用');
await mkdir(releaseDir, { recursive: true });
await writeFile(resolve(releaseDir, outputName), html, 'utf8');
await copyFile(resolve(projectDir, 'LICENSE'), resolve(releaseDir, 'LICENSE'));
await copyFile(resolve(projectDir, 'THIRD_PARTY_NOTICES.md'), resolve(releaseDir, 'THIRD_PARTY_NOTICES.md'));
await writeFile(resolve(releaseDir, '使用说明.txt'), [
  'Excel 筛选工作台',
  '',
  `双击 ${outputName}，用 Edge 或 Chrome 打开即可使用。无需安装 Node.js 或 Excel。`,
  '导入文件在浏览器内处理，页面可以断网使用。',
  '选择文件或拖入 Excel，选工作表和表头行，设置筛选，点击导出。',
  '导出全部匹配记录，不受当前分页影响；也可导出勾选记录或只导出可见列。',
  '保存的筛选方案仅含设置，不包含表格数据。关闭后需重新导入文件。',
  '导出的是数据表，公式使用原文件已保存的结果，原始图表、宏和复杂样式不保留。',
  '导入限制：25 MB、单工作表 100,000 条数据、250 列。',
  '',
  '项目采用 MIT 许可，第三方组件保留各自许可，详见附带文件。',
].join('\r\n'), 'utf8');
console.log(`Offline HTML: release/${outputName} (${(Buffer.byteLength(html) / 1024 / 1024).toFixed(2)} MB)`);
