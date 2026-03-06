/**
 * 模板解析辅助脚本
 * 用于从解包后的模板 XML 中提取关键格式参数
 *
 * 用法：
 *   1. 先用 unpack.py 解包模板 docx
 *   2. node scripts/analyze_template.js <unpacked_dir>
 */

const fs = require('fs');
const path = require('path');

const dir = process.argv[2];
if (!dir) {
  console.log('用法: node analyze_template.js <unpacked_directory>');
  console.log('示例: node analyze_template.js unpacked_template');
  process.exit(1);
}

const docPath = path.join(dir, 'word', 'document.xml');
if (!fs.existsSync(docPath)) {
  console.error(`文件不存在: ${docPath}`);
  process.exit(1);
}

const xml = fs.readFileSync(docPath, 'utf-8');

// 提取表格列宽
const gridColPattern = /<w:gridCol w:w="(\d+)"\/>/g;
const tables = [];
let currentTable = [];
let lastEnd = 0;

// 按 <w:tblGrid> 分组
const tblGridPattern = /<w:tblGrid>([\s\S]*?)<\/w:tblGrid>/g;
let match;
while ((match = tblGridPattern.exec(xml)) !== null) {
  const cols = [];
  let colMatch;
  const colPattern = /<w:gridCol w:w="(\d+)"\/>/g;
  while ((colMatch = colPattern.exec(match[1])) !== null) {
    cols.push(parseInt(colMatch[1]));
  }
  tables.push(cols);
}

console.log('=== 模板格式参数分析 ===\n');

// 页面设置
const pgSzMatch = xml.match(/<w:pgSz[^>]*w:w="(\d+)"[^>]*w:h="(\d+)"/);
if (pgSzMatch) {
  console.log(`页面宽度: ${pgSzMatch[1]} DXA (${(parseInt(pgSzMatch[1]) / 567).toFixed(1)} cm)`);
  console.log(`页面高度: ${pgSzMatch[2]} DXA (${(parseInt(pgSzMatch[2]) / 567).toFixed(1)} cm)`);
} else {
  console.log('页面尺寸: 默认 A4 (11906 x 16838 DXA)');
}

const pgMarMatch = xml.match(/<w:pgMar[^>]*w:top="(\d+)"[^>]*w:right="(\d+)"[^>]*w:bottom="(\d+)"[^>]*w:left="(\d+)"/);
if (pgMarMatch) {
  console.log(`页边距: 上=${pgMarMatch[1]} 右=${pgMarMatch[2]} 下=${pgMarMatch[3]} 左=${pgMarMatch[4]} DXA`);
}

console.log('');

// 表格结构
tables.forEach((cols, i) => {
  const total = cols.reduce((a, b) => a + b, 0);
  console.log(`表格${i + 1} (${cols.length}列, 总宽${total} DXA):`);
  console.log(`  列宽: [${cols.join(', ')}]`);
  console.log(`  JS常量: const TABLE${i + 1}_COLS = [${cols.join(', ')}];`);
  console.log('');
});

// 表格缩进
const tblIndMatch = xml.match(/<w:tblInd w:w="(\d+)"/);
if (tblIndMatch) {
  console.log(`表格缩进: ${tblIndMatch[1]} DXA`);
}

// 字号
const szPattern = /<w:sz w:val="(\d+)"\/>/g;
const sizes = new Set();
while ((match = szPattern.exec(xml)) !== null) {
  sizes.add(parseInt(match[1]));
}
console.log(`\n使用的字号(half-point): ${[...sizes].sort((a,b)=>a-b).join(', ')}`);
console.log(`对应磅值(pt): ${[...sizes].sort((a,b)=>a-b).map(s=>s/2).join(', ')}`);

// 字体
const fontPattern = /<w:rFonts[^>]*w:ascii="([^"]+)"/g;
const fonts = new Set();
while ((match = fontPattern.exec(xml)) !== null) {
  fonts.add(match[1]);
}
if (fonts.size > 0) {
  console.log(`使用的字体: ${[...fonts].join(', ')}`);
}

// 行距
const linePattern = /<w:spacing w:line="(\d+)"/g;
const lines = new Set();
while ((match = linePattern.exec(xml)) !== null) {
  lines.add(parseInt(match[1]));
}
if (lines.size > 0) {
  console.log(`行距值: ${[...lines].join(', ')}`);
}

// 首行缩进
const indentPattern = /<w:ind[^>]*w:firstLine="(\d+)"/g;
const indents = new Set();
while ((match = indentPattern.exec(xml)) !== null) {
  indents.add(parseInt(match[1]));
}
if (indents.size > 0) {
  console.log(`首行缩进值: ${[...indents].join(', ')}`);
}

// gridSpan 统计
const gridSpanPattern = /<w:gridSpan w:val="(\d+)"\/>/g;
const spans = new Set();
while ((match = gridSpanPattern.exec(xml)) !== null) {
  spans.add(parseInt(match[1]));
}
if (spans.size > 0) {
  console.log(`跨列合并(gridSpan): ${[...spans].join(', ')}`);
}

// 行高
const trHeightPattern = /<w:trHeight w:val="(\d+)"\/>/g;
const heights = new Set();
while ((match = trHeightPattern.exec(xml)) !== null) {
  heights.add(parseInt(match[1]));
}
if (heights.size > 0) {
  console.log(`表格行高: ${[...heights].sort((a,b)=>a-b).join(', ')}`);
}

console.log('\n=== 分析完成 ===');
console.log('请将以上参数填入 generate.js 顶部的常量区。');
