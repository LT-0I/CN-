const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, AlignmentType, BorderStyle, WidthType, VerticalAlign,
  PageBreak, SectionType
} = require('docx');

// === 模板格式参数（从六上数学.docx解析） ===
const PAGE_WIDTH = 11906;  // A4
const PAGE_HEIGHT = 16838;
const MARGIN_TOP = 1440;
const MARGIN_BOTTOM = 1440;
const MARGIN_LEFT = 1800;
const MARGIN_RIGHT = 1800;

// 表格1列宽（课时信息表，6列）
const TABLE1_COLS = [1362, 1356, 1356, 1358, 1219, 1756];
const TABLE1_WIDTH = TABLE1_COLS.reduce((a, b) => a + b, 0); // 8407
const TABLE1_INDENT = 109;

// 表格2列宽（导学流程表，3列：导学流程 | 初备 | 二度备）
// 第一页的导学流程表
const TABLE2_FIRST_COLS = [1482, 5774, 1941];
const TABLE2_FIRST_WIDTH = TABLE2_FIRST_COLS.reduce((a, b) => a + b, 0);

// 第二页起的导学流程表
const TABLE2_CONT_COLS = [1572, 5845, 1894];
const TABLE2_CONT_WIDTH = TABLE2_CONT_COLS.reduce((a, b) => a + b, 0);

// 边框样式
const BORDER = { style: BorderStyle.SINGLE, size: 1, color: "auto" };
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };

// 通用样式
const SONG_FONT = { ascii: "宋体", eastAsia: "宋体", hAnsi: "宋体" };
const HEADER_RUN_PROPS = { font: "宋体", bold: true, size: 28 };  // 14pt
const CONTENT_RUN_PROPS = { font: "宋体", size: 28 };  // 14pt
const TITLE_SIZE = 32;  // 16pt 导学案标题

// 创建标题行的文本run
function boldRun(text, size = 28) {
  return new TextRun({ text, bold: true, size, sizeCs: size, font: { hint: "eastAsia" } });
}

function normalRun(text, size = 28) {
  return new TextRun({ text, size, sizeCs: size, font: { hint: "eastAsia", ...SONG_FONT } });
}

function songRun(text, size = 28, bold = false) {
  return new TextRun({
    text, size, sizeCs: size, bold,
    font: { ascii: "宋体", eastAsia: "宋体", hAnsi: "宋体", hint: "eastAsia" }
  });
}

// 创建"导学案"标题段落
function createDaoxueAnTitle() {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({
        text: "导学案",
        bold: true,
        size: TITLE_SIZE,
        sizeCs: TITLE_SIZE,
        font: { hint: "eastAsia" }
      })
    ]
  });
}

// 创建表格单元格
function createCell(content, width, opts = {}) {
  const {
    bold = false, gridSpan, vMerge, vMergeRestart = false,
    alignment = AlignmentType.CENTER, font = null,
    spacing, indent, children: customChildren
  } = opts;

  const tcPr = {
    width: { size: width, type: WidthType.DXA },
    borders: BORDERS,
    verticalAlign: VerticalAlign.CENTER
  };
  if (gridSpan) tcPr.columnSpan = gridSpan;
  if (vMergeRestart) tcPr.rowSpan = opts.rowSpan || 2;

  const paragraphs = [];

  if (customChildren) {
    // 多段落内容
    customChildren.forEach(child => {
      const pOpts = { children: [] };
      if (child.alignment) pOpts.alignment = child.alignment;
      if (child.spacing) pOpts.spacing = child.spacing;
      if (child.indent) pOpts.indent = child.indent;

      child.runs.forEach(run => {
        pOpts.children.push(songRun(run.text, run.size || 28, run.bold || false));
      });
      paragraphs.push(new Paragraph(pOpts));
    });
  } else if (typeof content === 'string') {
    const pOpts = {
      alignment,
      children: [
        bold ? boldRun(content) : normalRun(content)
      ]
    };
    if (spacing) pOpts.spacing = spacing;
    if (indent) pOpts.indent = indent;
    paragraphs.push(new Paragraph(pOpts));
  } else {
    paragraphs.push(new Paragraph({ alignment, children: [] }));
  }

  return new TableCell({
    ...tcPr,
    children: paragraphs
  });
}

// 创建表格1（课时信息表）
function createInfoTable(lesson) {
  const rows = [];

  // Row 1: 年级 | 六 | 科目 | 数学 | 总课时 | [编号]
  rows.push(new TableRow({
    height: { value: 377, rule: "atLeast" },
    children: [
      createCell("年级", TABLE1_COLS[0], { bold: true }),
      createCell("六", TABLE1_COLS[1]),
      createCell("科目", TABLE1_COLS[2], { bold: true }),
      createCell("数学", TABLE1_COLS[3]),
      createCell("总课时", TABLE1_COLS[4], { bold: true }),
      createCell(String(lesson.totalLessonNum), TABLE1_COLS[5])
    ]
  }));

  // Row 2: 课题 | [title, span 3] | 课时 | [unitLessonNum]
  rows.push(new TableRow({
    height: { value: 377, rule: "atLeast" },
    children: [
      createCell("课题", TABLE1_COLS[0], { bold: true }),
      createCell(lesson.title, TABLE1_COLS[1] + TABLE1_COLS[2] + TABLE1_COLS[3], { gridSpan: 3 }),
      createCell("课时", TABLE1_COLS[4], { bold: true }),
      createCell(lesson.unitLessonNum, TABLE1_COLS[5])
    ]
  }));

  // Row 3: 主备教师 | [空] | 审核教师 | [空] | 课型 | 新授课
  rows.push(new TableRow({
    height: { value: 377, rule: "atLeast" },
    children: [
      createCell("主备教师", TABLE1_COLS[0], { bold: true }),
      createCell("", TABLE1_COLS[1]),
      createCell("审核教师", TABLE1_COLS[2], { bold: true }),
      createCell("", TABLE1_COLS[3]),
      createCell("课型", TABLE1_COLS[4], { bold: true }),
      createCell(lesson.courseType, TABLE1_COLS[5])
    ]
  }));

  // Row 4: 学习目标 | [content, span 5]
  const objContent = [];
  objContent.push({
    alignment: AlignmentType.LEFT,
    spacing: { line: 360, lineRule: "auto" },
    indent: { firstLine: 560 },
    runs: [{ text: "一、知识与能力", size: 28 }]
  });
  objContent.push({
    alignment: AlignmentType.LEFT,
    spacing: { line: 360, lineRule: "auto" },
    indent: { firstLine: 560 },
    runs: [{ text: lesson.objectives.knowledge, size: 28 }]
  });
  objContent.push({
    alignment: AlignmentType.LEFT,
    spacing: { line: 360, lineRule: "auto" },
    indent: { firstLine: 560 },
    runs: [{ text: "二、过程与方法", size: 28 }]
  });
  objContent.push({
    alignment: AlignmentType.LEFT,
    spacing: { line: 360, lineRule: "auto" },
    indent: { firstLine: 560 },
    runs: [{ text: lesson.objectives.process, size: 28 }]
  });
  objContent.push({
    alignment: AlignmentType.LEFT,
    spacing: { line: 360, lineRule: "auto" },
    indent: { firstLine: 560 },
    runs: [{ text: "三、情感态度与价值观", size: 28 }]
  });
  objContent.push({
    alignment: AlignmentType.LEFT,
    spacing: { line: 360, lineRule: "auto" },
    indent: { firstLine: 560 },
    runs: [{ text: lesson.objectives.emotion, size: 28 }]
  });

  rows.push(new TableRow({
    height: { value: 1746, rule: "atLeast" },
    children: [
      createCell("学习目标", TABLE1_COLS[0], { bold: true }),
      createCell(null, TABLE1_COLS[1] + TABLE1_COLS[2] + TABLE1_COLS[3] + TABLE1_COLS[4] + TABLE1_COLS[5], {
        gridSpan: 5,
        children: objContent
      })
    ]
  }));

  // Row 5: 重点
  rows.push(new TableRow({
    height: { value: 521, rule: "atLeast" },
    children: [
      createCell("重点", TABLE1_COLS[0], { bold: true }),
      createCell(null, TABLE1_COLS[1] + TABLE1_COLS[2] + TABLE1_COLS[3] + TABLE1_COLS[4] + TABLE1_COLS[5], {
        gridSpan: 5,
        children: [{
          spacing: { line: 360, lineRule: "auto" },
          indent: { firstLine: 560 },
          runs: [{ text: lesson.keyPoint, size: 28 }]
        }]
      })
    ]
  }));

  // Row 6: 难点
  rows.push(new TableRow({
    height: { value: 521, rule: "atLeast" },
    children: [
      createCell("难点", TABLE1_COLS[0], { bold: true }),
      createCell(null, TABLE1_COLS[1] + TABLE1_COLS[2] + TABLE1_COLS[3] + TABLE1_COLS[4] + TABLE1_COLS[5], {
        gridSpan: 5,
        children: [{
          spacing: { line: 360, lineRule: "auto" },
          indent: { firstLine: 560 },
          runs: [{ text: lesson.difficulty, size: 28 }]
        }]
      })
    ]
  }));

  // Row 7: 课前准备
  rows.push(new TableRow({
    height: { value: 521, rule: "atLeast" },
    children: [
      createCell("课前准备", TABLE1_COLS[0], { bold: true }),
      createCell(null, TABLE1_COLS[1] + TABLE1_COLS[2] + TABLE1_COLS[3] + TABLE1_COLS[4] + TABLE1_COLS[5], {
        gridSpan: 5,
        children: [{
          indent: { firstLine: 420 },
          runs: [{ text: lesson.preparation, size: 28 }]
        }]
      })
    ]
  }));

  return new Table({
    width: { size: TABLE1_WIDTH, type: WidthType.DXA },
    indent: { size: TABLE1_INDENT, type: WidthType.DXA },
    columnWidths: TABLE1_COLS,
    rows
  });
}

// 将教学内容文本分成段落
function textToParagraphs(text, bold = false) {
  const lines = text.split('\n');
  return lines.map(line => {
    const isSectionTitle = /^(活动|一、|二、|三、|四、)/.test(line.trim());
    return {
      spacing: { line: 360, lineRule: "auto" },
      indent: { firstLine: 562 },
      runs: [{ text: line.trim(), size: 28, bold: isSectionTitle || bold }]
    };
  }).filter(p => p.runs[0].text.length > 0);
}

// 创建导学流程表（包含标题行和内容行）
function createTeachingTable(lesson) {
  const cols = TABLE2_FIRST_COLS;
  const totalWidth = TABLE2_FIRST_WIDTH;

  // 构建初备列的所有段落
  const contentParagraphs = [];

  // 一、激趣导入
  contentParagraphs.push({
    spacing: { line: 360, lineRule: "auto" },
    indent: { firstLine: 562 },
    runs: [{ text: "一、激趣导入", size: 28, bold: true }]
  });
  textToParagraphs(lesson.teaching.introduction).forEach(p => contentParagraphs.push(p));

  // 二、快乐导学
  contentParagraphs.push({
    spacing: { line: 360, lineRule: "auto" },
    indent: { firstLine: 562 },
    runs: [{ text: "二、快乐导学", size: 28, bold: true }]
  });
  textToParagraphs(lesson.teaching.mainTeaching).forEach(p => contentParagraphs.push(p));

  // 三、学以致用
  contentParagraphs.push({
    spacing: { line: 360, lineRule: "auto" },
    indent: { firstLine: 562 },
    runs: [{ text: "三、学以致用", size: 28, bold: true }]
  });
  textToParagraphs(lesson.teaching.practice).forEach(p => contentParagraphs.push(p));

  // 四、总结提升
  contentParagraphs.push({
    spacing: { line: 360, lineRule: "auto" },
    indent: { firstLine: 562 },
    runs: [{ text: "四、总结提升", size: 28, bold: true }]
  });
  textToParagraphs(lesson.teaching.summary).forEach(p => contentParagraphs.push(p));

  // 导学流程列 - 竖排显示"导学流程"
  const flowCellChildren = [
    new Paragraph({ alignment: AlignmentType.CENTER, children: [boldRun("导", 32)] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [boldRun("学", 32)] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [boldRun("流", 32)] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [boldRun("程", 32)] })
  ];

  // 标题行
  const headerRow = new TableRow({
    height: { value: 540, rule: "atLeast" },
    children: [
      new TableCell({
        width: { size: cols[0], type: WidthType.DXA },
        borders: BORDERS,
        verticalAlign: VerticalAlign.CENTER,
        rowSpan: 2,
        children: flowCellChildren
      }),
      createCell("初备", cols[1], { bold: true, gridSpan: undefined }),
      createCell("二度备", cols[2], { bold: true })
    ]
  });

  // 内容行
  const contentRow = new TableRow({
    children: [
      // 导学流程列被合并（rowSpan）
      createCell(null, cols[1], {
        children: contentParagraphs
      }),
      createCell("", cols[2])  // 二度备留空
    ]
  });

  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: cols,
    rows: [headerRow, contentRow]
  });
}

// 创建板书设计和反思表
function createBoardTable(lesson) {
  const cols = TABLE2_CONT_COLS;
  const totalWidth = TABLE2_CONT_WIDTH;

  const boardParagraphs = lesson.boardDesign.split('\n').map(line => ({
    spacing: { line: 360, lineRule: "auto" },
    indent: { firstLine: 560 },
    runs: [{ text: line.trim(), size: 28 }]
  }));

  const rows = [
    // 板书设计
    new TableRow({
      height: { value: 2000, rule: "atLeast" },
      children: [
        createCell("板书设计", cols[0], { bold: true }),
        createCell(null, cols[1], { children: boardParagraphs }),
        createCell("", cols[2])
      ]
    }),
    // 导学反思
    new TableRow({
      height: { value: 1500, rule: "atLeast" },
      children: [
        createCell("导学反思", cols[0], { bold: true }),
        createCell("", cols[1]),
        createCell("", cols[2])
      ]
    })
  ];

  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: cols,
    rows
  });
}

// 为一个课时生成所有内容（多个section用于分页）
// compact模式：2页/课时，去掉"导学案"标题和页眉
function createLessonSections(lesson, isFirst, compact = false) {
  const children = [];

  if (!compact) {
    // 原始3页模式：添加"导学案"标题
    children.push(createDaoxueAnTitle());
  }

  // 表格1：课时信息表
  children.push(createInfoTable(lesson));

  // 空行分隔
  children.push(new Paragraph({
    spacing: { before: 0, after: 0 },
    children: []
  }));

  // 表格2：导学流程表
  children.push(createTeachingTable(lesson));

  // 空行分隔
  children.push(new Paragraph({
    spacing: { before: 0, after: 0 },
    children: []
  }));

  // 表格3：板书设计与反思
  children.push(createBoardTable(lesson));

  const sectionProps = {
    properties: {
      page: {
        size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
        margin: {
          top: compact ? 1000 : MARGIN_TOP,
          bottom: compact ? 1000 : MARGIN_BOTTOM,
          left: MARGIN_LEFT,
          right: MARGIN_RIGHT
        }
      },
      ...(isFirst ? {} : { type: SectionType.NEXT_PAGE })
    },
    children
  };

  if (!compact) {
    // 原始模式：添加页眉"导学案"
    sectionProps.headers = {
      default: new Header({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: "导学案",
                bold: true,
                size: TITLE_SIZE,
                sizeCs: TITLE_SIZE,
                font: { hint: "eastAsia" }
              })
            ]
          })
        ]
      })
    };
  }

  return sectionProps;
}

// 主函数
async function generate(dataFiles, outputPath) {
  const allLessons = [];

  for (const file of dataFiles) {
    const data = JSON.parse(fs.readFileSync(file, 'utf-8'));
    allLessons.push(...data);
  }

  console.log(`正在生成 ${allLessons.length} 个课时的导学案...`);

  const sections = allLessons.map((lesson, i) => createLessonSections(lesson, i === 0, compact));

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: "宋体",
            size: 28  // 14pt (= 28 half-points)
          }
        }
      }
    },
    sections
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  console.log(`文档已生成：${outputPath}`);
}

// 命令行参数处理
const args = process.argv.slice(2);
const compact = args.includes('--compact');
const filteredArgs = args.filter(a => a !== '--compact');

if (filteredArgs.length < 2) {
  console.log('用法: node generate.js <output.docx> <data1.json> [data2.json ...] [--compact]');
  console.log('  --compact: 2页/课时，去掉导学案标题和页眉');
  console.log('示例: node generate.js output/unit2.docx data/unit2.json --compact');
  process.exit(1);
}

const outputPath = filteredArgs[0];
const dataFiles = filteredArgs.slice(1);

// 确保输出目录存在
const outputDir = path.dirname(outputPath);
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

generate(dataFiles, outputPath).catch(err => {
  console.error('生成失败:', err);
  process.exit(1);
});
