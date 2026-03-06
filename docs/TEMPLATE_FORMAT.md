# 模板格式参数说明

本文档说明如何从 Word 模板中提取格式参数，以及各参数的含义。

---

## 解包模板

```bash
python scripts/office/unpack.py 模板文件.docx unpacked_template/
```

解包后重点查看 `unpacked_template/word/document.xml`。

---

## 关键格式参数

### 页面设置（从 `<w:sectPr>` 或默认值推断）

| 参数 | 字段 | 示例值 | 说明 |
|------|------|--------|------|
| 页面宽度 | `<w:pgSz w:w>` | 11906 | A4 宽度（DXA 单位） |
| 页面高度 | `<w:pgSz w:h>` | 16838 | A4 高度 |
| 上边距 | `<w:pgMar w:top>` | 1440 | 约 2.54cm |
| 下边距 | `<w:pgMar w:bottom>` | 1440 | |
| 左边距 | `<w:pgMar w:left>` | 1800 | 约 3.17cm |
| 右边距 | `<w:pgMar w:right>` | 1800 | |

**单位换算：** 1 inch = 1440 DXA, 1 cm ≈ 567 DXA

### 表格结构

从 `<w:tblGrid>` 子元素 `<w:gridCol w:w="xxx"/>` 提取列宽。

**表格1（课时信息表，6列）：**
```xml
<w:tblGrid>
  <w:gridCol w:w="1362"/>
  <w:gridCol w:w="1356"/>
  <w:gridCol w:w="1356"/>
  <w:gridCol w:w="1358"/>
  <w:gridCol w:w="1219"/>
  <w:gridCol w:w="1756"/>
</w:tblGrid>
```

**表格2（导学流程表，3列）：**
```xml
<w:tblGrid>
  <w:gridCol w:w="1482"/>  <!-- 导学流程（竖排） -->
  <w:gridCol w:w="5774"/>  <!-- 初备 -->
  <w:gridCol w:w="1941"/>  <!-- 二度备 -->
</w:tblGrid>
```

### 单元格合并

| 类型 | XML 标记 | 说明 |
|------|----------|------|
| 跨列合并 | `<w:gridSpan w:val="3"/>` | 合并3列 |
| 跨行合并（起始） | `<w:vMerge w:val="restart"/>` | 纵向合并开始 |
| 跨行合并（继续） | `<w:vMerge/>` | 被合并行 |

### 字体和字号

| 元素 | XML | 说明 |
|------|-----|------|
| 页面标题"导学案" | `<w:sz w:val="32"/>` | 16pt，加粗 |
| 表头（年级/科目等） | `<w:sz w:val="28"/>` `<w:b/>` | 14pt，加粗 |
| 正文内容 | `<w:sz w:val="28"/>` | 14pt |
| 宋体指定 | `<w:rFonts w:ascii="宋体" w:hAnsi="宋体"/>` | |
| 东亚字体提示 | `<w:rFonts w:hint="eastAsia"/>` | |

注意：`<w:sz>` 的值是 **半磅(half-point)**，28 = 14pt。

### 段落格式

| 属性 | XML | 说明 |
|------|-----|------|
| 行距1.5倍 | `<w:spacing w:line="360" w:lineRule="auto"/>` | |
| 首行缩进 | `<w:ind w:firstLine="560"/>` | 约2个字符 |
| 居中对齐 | `<w:jc w:val="center"/>` | |

### 表格行高

| 行类型 | 值 | 说明 |
|--------|-----|------|
| 普通信息行 | 377 | atLeast |
| 学习目标行 | 1746 | 多段落内容 |
| 重点/难点行 | 521 | |
| 导学流程标题行 | 540 | |
| 板书设计行 | 2000 | |
| 导学反思行 | 1500 | |

---

## 如何适配新模板

1. 解包新模板 `.docx`
2. 打开 `word/document.xml`
3. 按上述对照表提取各参数值
4. 修改 `generate.js` 顶部的常量区
