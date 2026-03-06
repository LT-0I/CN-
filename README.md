# 教师导学案自动生成工具 (Teaching Plan Generator)

基于 Word 模板自动生成中小学教师导学案（教案），支持任意科目、学期和教材版本。

## 功能特性

- 严格复制 Word 模板的表格结构、字体、字号、列宽等格式参数
- 分单元/分批生成，内置防幻觉机制（知识点锚定、自检验算、易错提醒）
- 支持 `--compact` 模式（2页/课时，无标题）和标准模式（3页/课时，含"导学案"标题+页眉）
- 数据与逻辑分离：课时内容（JSON）与文档生成脚本独立，便于复用和修正
- 一键合并全册文档

## 快速开始

```bash
# 1. 安装依赖
npm install

# 2. 解析你的模板（可选，已内置默认参数）
# 将你的模板 .docx 放入 templates/ 目录

# 3. 准备课时数据
# 在 data/ 目录下创建 JSON 文件（参考 examples/data/ 中的示例）

# 4. 生成单个单元
node scripts/generate.js output/unit1.docx data/unit1.json

# 5. 生成紧凑版（2页/课时，无标题）
node scripts/generate.js output/unit2.docx data/unit2.json --compact

# 6. 合并全册
node scripts/generate.js output/全册导学案.docx data/unit1.json data/unit2.json data/unit3.json
```

## 项目结构

```
teaching-plan-generator/
├── README.md                    # 本文件
├── package.json                 # Node.js 依赖
├── SKILL.md                     # AI Skill 定义（供 Claude Code 等 AI 工具使用）
├── PROMPT_TEMPLATE.md           # 分轮提示词模板
├── scripts/
│   ├── generate.js              # 核心生成脚本
│   └── analyze_template.js      # 模板解析辅助脚本
├── templates/
│   └── README.md                # 模板文件放置说明
├── examples/
│   ├── data/                    # 示例数据文件
│   │   └── sample_unit.json
│   └── output/                  # 示例输出（gitignore）
├── docs/
│   ├── ANTI_HALLUCINATION.md    # 防幻觉策略详解
│   ├── TEMPLATE_FORMAT.md       # 模板格式参数说明
│   └── JSON_SCHEMA.md           # 数据文件格式说明
└── .gitignore
```

## 输入与输出

### 输入
1. **模板文件**（.docx）：一份已有的导学案 Word 模板
2. **科目/学期/教材版本**：如"小学数学六年级下册（北师大版）"

### 输出
- 按单元拆分的 .docx 文件
- 合并后的全册 .docx 文件

## 防幻觉机制

详见 [docs/ANTI_HALLUCINATION.md](docs/ANTI_HALLUCINATION.md)

- 分轮生成（每轮6-8课时）
- 知识点锚定清单
- 练习题答案自检验算
- 易错点标注
- 数据与逻辑分离

## 许可证

MIT License
