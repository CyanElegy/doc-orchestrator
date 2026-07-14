# Doc-Orchestrator — 央企信息化项目文档编排 Skill

## 一、概述

Doc-Orchestrator 是一个面向央企（国有企业）信息化建设项目的文档生成编排 Skill。它从 Word/Excel 模板和优秀范例出发，结合项目关键信息，自动生成符合《过程管理办法》标准的技术文档。

### 核心理念

**格式权威来源只有模板。AI 只负责内容（Markdown），Python 精确装配格式（docx/xlsx）。**

这是一个**宣言式的工作流编排器**，而非代码密集的应用。AI 在此 Skill 中的角色是内容生成与质量评审，不参与任何格式化决策。

### 设计原则

1. **SKILL.md 是唯一入口。** 不存在平台适配器（adapters/）。Claude Code、OpenCode、Codex 等 Agent 平台通过各自的技能发现机制直接读取本文件。平台适配由平台方负责。
2. **Python 脚本是加速器，不是依赖。** 所有脚本都有手动替代路径。AI 在无法运行脚本时可手动编写等价代码在线执行。
3. **references/ 按需加载。** 引用文件仅在对应流程触发时读取，不预加载到上下文。
4. **AI 永不触碰格式决策。** 字体、字号、编号样式、页边距——全部继承自模板。AI 的输出仅限于 Markdown + YAML front matter。

---

## 二、工作模式

Skill 支持两种核心工作模式，通过初始化路径选择：

### 2.1 完整流程模式（Full Process Mode）

操作人提供一份「管理办法」文档（regulation doc）—— 描述项目阶段划分和每阶段应产出文档的规范性文件。AI 从管理办法中提取阶段结构、文档清单和交付顺序，然后操作人为每个阶段提供模板+范例，Skill 按阶段顺序依次生成全部文档，并在最后执行文件清单校验。

**适用场景：** 新项目启动，需要全套过程管理文档。

**流程：**
1. 操作人上传管理办法文档
2. AI 分析管理办法，提取阶段与文档清单
3. 显示阶段清单和文档依赖关系，操作人确认/调整
4. 操作人为每个阶段提供模板文件和优秀范例
5. 按阶段顺序逐文档生成
6. 全部完成后进行文件清单完整性校验

### 2.2 直接生成模式（Direct Generation Mode）

操作人直接提供：模板文件 + 优秀范例 + 项目信息。不经过流程拆解阶段，立即生成单份文档。

**适用场景：** 临时需要某份文档，或已有现成模板。

**流程：**
1. 操作人指定目标文档类型和模板
2. 提供优秀范例
3. 补充项目信息
4. 立即生成

---

## 三、粒度级别

Skill 支持 5 种粒度级别，操作人在初始化时选择：

| 级别 | 范围 | 适用场景 |
|------|------|----------|
| 全流程（full-process） | 管理办法定义的全部阶段和文档 | 项目整体启动 |
| 多阶段（multi-phase） | 操作人指定的 2~3 个连续阶段 | 阶段性交付 |
| 单阶段（single-phase） | 一个评审里程碑下的全部文档 | 里程碑节点 |
| 单文档（single-doc） | 一个阶段内的单份特定文档 | 文档补全 |
| 单章节（single-chapter） | 一份文档内的单个章节 | 局部修改/完善 |

单章节再生流程：
- 仅重新生成该章节的 Markdown 内容
- 文档其他章节的 Markdown 不变
- 整个文档重新执行格式装配（assemble_docx.py）
- 其他章节的 AI 评审意见保留

---

## 四、五层架构

### Layer 0: 引导与发现（Guide & Discover）

**职责：** 扫描 phases/ 目录，加载可用阶段列表，呈现选项给操作人，收集操作人输入。

**操作步骤：**
1. 检查 phases/ 目录是否存在阶段定义
2. 若无阶段定义：询问操作人是否有管理办法文档
   - 有：启动管理办法解析流程
   - 无：加载 `references/phase-preset.md` 使用预置标准流程
3. 呈现阶段选项和文档清单
4. 询问粒度级别（全流程/多阶段/单阶段/单文档/单章节）
5. 确认选择后进入 Layer 1

**输出：** 操作人确认的选择结果（target manifest）

### Layer 1: 资源解析与缓存（Resource Ingestion）

**职责：** 解析模板文件、读取范例文档、管理办法，缓存解析结果避免重复解析。

**处理的资源类型：**

| 资源 | 解析工具 | 输出 |
|------|----------|------|
| .docx 模板 | `scripts/extract_template.py` | JSON：标题树、样式定义、页面设置、占位符 |
| .xlsx 模板 | `scripts/assemble_xlsx.py`（parse mode） | JSON：列结构、公式、样式 |
| 优秀范例 .docx | `scripts/extract_template.py`（text-only mode） | Markdown：内容结构参考 |
| 管理办法文档 | AI 提炼结构 | YAML：阶段划分 + 文档清单 |

**缓存策略：**
- 每个资源的解析结果哈希存储在 `cache/` 目录
- 下次请求时比对修改时间，未变更则使用缓存
- 缓存 hash: sha256 of file content

### Layer 2: 内容生成（Content Generation）

**职责：** AI 逐章节生成 Markdown 内容。

**生成约束：**
- AI 输出**仅限** Markdown，可包含 YAML front matter
- 格式信息（字体、字号、缩进、间距）不在生成内容中出现
- 章节结构匹配模板提取出的标题树
- 表格仅用 Markdown 表格语法，无样式定义
- 图片引用格式：`![alt text](relative/path.png)`
- 交叉引用格式：`见 [章节标题](#chapter-identifier)`
- 每章内容独立生成，上下文仅包含该章节参考信息和前文概要

**逐章生成协议：**
1. 从模板提取的标题树中获取当前章节标题和层级
2. 加载该章节的参考内容（范例对应章节 + 项目信息 + 业务指导）
3. AI 生成该章节 Markdown
4. 追加到文档 Markdown 缓冲区
5. 重复直到所有章节完成

**输出：** 完整文档的 Markdown 文件（含 YAML front matter）

YAML front matter 结构：
```yaml
---
doc_id: requirements-spec
title: 需求规格说明书
phase: 01-requirements-review
version: 1.0
generated: 2026-07-14
chapters:
  - id: ch01-introduction
    title: 引言
    level: 1
  - id: ch02-scope
    title: 项目范围
    level: 1
  - id: ch02-01-business-goals
    title: 业务目标
    level: 2
templates_used:
  - requirements-template-v3.docx
examples_used:
  - example-requirements.docx
---
```

### Layer 3: 格式装配（Format Assembly）

**职责：** Python 脚本从模板继承样式，将 Markdown 内容装配为格式精确的 docx/xlsx。

**核心原则：**
- 模板文件被完整复制，作为样式源文件
- 所有样式（字体、字号、颜色、对齐、间距、编号）继承自模板
- Markdown 中的 `#` `##` `###` 映射到模板中的 Heading 1/2/3 样式
- 正文段落映射到模板中的 Normal 或 Body Text 样式
- 表格映射到模板中的 Table Grid 样式
- AI 不做任何样式映射决策，映射规则在脚本中定义

详见 `scripts/assemble_docx.py` 和 `scripts/assemble_xlsx.py`。

### Layer 4: 校验与归档（Validation & Archive）

**职责：** 两步校验 + 按管理办法规定的顺序归档。

**校验步骤：**

1. **规则检查（Rule Checker）：** 由 `scripts/assemble_docx.py --validate` 执行
   - 章节完整性：模板标题树的所有章节是否都有对应 Markdown
   - 编号连续性：Heading numbering 是否连续无跳号
   - 必填字段：项目信息占位符是否全部替换
   - 引用完整性：文档内交叉引用是否有效
   - 文件大小：是否非空、无异常

2. **AI 内容评审（AI Review）：** AI 对生成文档内容进行定性评审
   - 术语一致性：与项目信息中的业务术语是否一致
   - 内容深度：是否满足该文档类型的标准深度
   - 逻辑连贯性：章节之间逻辑衔接是否顺畅
   - 合规性：是否符合央企文档标准要求

**归档规则：**
- 产出文件按管理办法中定义的文档顺序编号
- 归档目录结构：`{output_dir}/{regulation_name}/{phase_id}/{doc_id}_v{version}.docx`
- 同步生成归档清单 `archive_manifest.yaml`

---

## 五、脚本（Accelerators）

脚本是加速器，不是硬依赖。AI 可以手动编写等价逻辑在线运行。

### 5.1 scripts/extract_template.py

**用途：** 从 .docx 模板中提取结构化信息，输出 JSON。

```bash
python3 scripts/extract_template.py template.docx > structure.json
```

**依赖：** python-docx

**安装：** `pip install python-docx`

**手动替代路径：** 由 AI 直接编写 python-docx 内联代码从模板读取样式信息。

**输出 JSON 结构：**
```json
{
  "document_name": "需求规格说明书模板",
  "page_setup": {
    "paper_size": "A4",
    "orientation": "portrait",
    "margins": {"top": 2.54, "bottom": 2.54, "left": 3.17, "right": 3.17}
  },
  "heading_tree": [
    {"level": 1, "text": "第一章 引言", "numbering_style": "第1章", "guidance": "简述项目背景"},
    {"level": 2, "text": "1.1 编写目的", "numbering_style": "1.1"},
    {"level": 1, "text": "第二章 项目范围", "numbering_style": "第2章"}
  ],
  "styles": {
    "Heading 1": {
      "font_name": "黑体",
      "font_size": 16,
      "bold": true,
      "alignment": "center",
      "space_before": 12,
      "space_after": 6
    },
    "Heading 2": {
      "font_name": "黑体",
      "font_size": 14,
      "bold": true,
      "alignment": "left",
      "space_before": 8,
      "space_after": 4
    },
    "Normal": {
      "font_name": "仿宋_GB2312",
      "font_size": 12,
      "line_spacing": 1.5,
      "first_line_indent": 0.74
    }
  },
  "header": "国家电网信息化项目 — 需求规格说明书",
  "footer": "第 {PAGE} 页 / 共 {NUMPAGES} 页",
  "placeholders": ["《项目名称》", "《建设单位》", "《编制日期》"],
  "tables": [
    {
      "rows": 5,
      "cols": 4,
      "header_row": true,
      "sample_data": ["序号", "功能模块", "功能描述", "优先级"]
    }
  ],
  "images": [
    {"relationship_id": "rId8", "inline": true, "width_cm": 14.5}
  ]
}
```

### 5.2 scripts/assemble_docx.py

**用途：** 将 Markdown 内容 + 模板 .docx + 项目信息 YAML 装配为格式精确的 .docx。

```bash
python3 scripts/assemble_docx.py content.md template.docx project.yaml output.docx
```

**参数：**
- `content.md` — AI 生成的 Markdown 文件（含 YAML front matter）
- `template.docx` — 作为样式源的模板文件
- `project.yaml` — 项目信息结构化数据
- `output.docx` — 输出文件路径

**依赖：** python-docx, Pillow

**安装：** `pip install python-docx Pillow`

**手动替代路径：** AI 编写等价 python-docx 内联代码。

**装配流程：**
1. 复制模板文件作为输出基础（继承全部样式定义、页面设置、页眉页脚）
2. 清除原模板的正文内容，保留样式定义
3. 解析 Markdown 为结构化文档树：
   - `# → Heading 1`
   - `## → Heading 2`
   - `### → Heading 3`
   - 段落 → Normal/Body Text
   - 表格 → Table Grid
   - 列表 → List Bullet/List Number
   - 代码块 → Code（monospace preserved）
   - 图片 → 嵌入 Image 对象，从模板获取尺寸参考
4. 应用对应的模板样式
5. 替换占位符：
   - `<PROJECT_NAME>` → 项目名称
   - `<DATE>` → 当前日期
   - 模板中 `《项目名称》` → 项目名称
   - 页眉/页脚中的占位符一并替换
6. 插入 TOC（Table of Contents）Field Code：`TOC \o "1-3" \h \z \u`
7. 嵌入图片（如有引用），从模板获取引用尺寸
8. 保存输出文件

### 5.3 scripts/assemble_xlsx.py

**用途：** 将结构化数据 + 模板 .xlsx 装配为格式精确的 .xlsx。

```bash
python3 scripts/assemble_xlsx.py data.yaml template.xlsx output.xlsx
```

**参数：**
- `data.yaml` — 结构化数据 YAML
- `template.xlsx` — 样式模板文件
- `output.xlsx` — 输出文件路径

**依赖：** openpyxl

**安装：** `pip install openpyxl`

**手动替代路径：** AI 编写等价 openpyxl 内联代码。

**装配流程：**
1. 复制模板文件作为输出基础（继承列宽、单元格样式、合并单元格、条件格式）
2. 清除数据区域单元格内容，保留以下元素：
   - 公式单元格（`=SUM(...)` 等）
   - 汇总行（summary rows）
   - 冻结窗格（freeze panes）
   - 数据验证（data validation）
3. 将数据 YAML 按 key 匹配到模板的列标题行
4. 逐行写入数据，继承每列对应的模板样式
5. 自动将空数据区域格式为与有数据区域一致（防止格式截断）
6. 保存输出文件

---

## 六、引用文件（references/）

以下文件按流程需求按需加载，不预加载到 AI 上下文。

### references/phase-schema.md

**用途：** 定义阶段（phase）YAML 的 Schema 规范。

**加载时机：** 当 AI 需要创建或验证阶段定义文件时。

**内容概要：**
- Phase YAML 的字段定义（id, name, description, order, documents[]）
- 每个文档条目（doc）的字段定义
- 文档间的依赖关系声明语法
- 自定义字段扩展规则

### references/phase-preset.md

**用途：** 内置的标准央企信息化项目阶段预设。

**加载时机：** 操作人提供了管理办法文档（改用管理办法解析），或没有管理办法文档时作为备选流程。

**内容概要：**
- 标准 5 阶段划分（需求分析评审、概要设计评审、详细设计评审、测试验收评审、上线部署评审）
- 各阶段的规范文档清单（20+ 种文档类型）
- 文档产出顺序依赖关系
- 默认模板匹配规则

### references/pitfalls.md

**用途：** 常见失败模式与恢复策略。

**加载时机：** 错误发生时，或 AI 在生成前检查可能的风险。

**内容概要：**
- 模板样式命名不一致导致装配失败
- 中文字体在模板中未内嵌导致跨平台渲染差异
- 图片引用路径丢失
- 页眉/页脚占位符未被替换
- Markdown 标题层级与模板 Heading 层级不匹配
- 表格行列数超出模板预设
- 每个问题的恢复操作步骤

---

## 七、硬约束规则（Hard Constraints）

这些规则是 Skill 不可违反的设计约束：

**规则 1：AI 永不做出格式化决策**
- 字体名称、字号、粗体/斜体、颜色、对齐方式、行距、段间距、编号样式、页边距——所有这些信息全部来自模板。
- AI 生成的 Markdown 中不应出现 `style=""`、`<font>`、CSS 类名或其他格式指示。
- 违反此规则意味着文档的格式无法保证与模板一致，必须重新生成。

**规则 2：AI 输出仅限 Markdown + YAML front matter**
- AI 的输出格式只有两种：Markdown（含 YAML front matter）和 YAML。
- AI 不直接生成 .docx、.xlsx 或其他二进制格式。
- AI 不调用脚本执行（脚本由框架在后续步骤调用）。

**规则 3：格式装配全在 Python 脚本中完成**
- 所有 docx/xlsx 的创建、格式应用、样式映射操作在 Python 脚本中完成。
- AI 上下文不包含 python-docx/openpyxl 的完整 API 文档，只需了解脚本的输入/输出契约。

**规则 4：模板是唯一的样式源**
- 模板文件被完整复制，而不是被读取样式信息后重新创建。
- 脚本清空模板正文但保留全部样式定义、页面设置、页眉页脚。
- 不可从零创建文档——必须始终基于模板。

**规则 5：单章节再生不重新生成整篇文档**
- 当触发单章节再生时，仅该章节对应的 Markdown 被重新生成。
- 文档其他章节的 Markdown 内容从缓冲区读取，保持不变。
- 整篇文档重新执行格式装配（assemble_docx.py），以保持 TOC 编号连续。

---

## 八、交互流程（完整步骤）

### Step 0: 初始化（Greet & Discover）

技能被唤起时：

1. 显示技能名称和版本：「Doc-Orchestrator — 央企信息化项目文档编排」
2. 用一句话解释技能能力：Skill 能从模板和范例出发，自动生成符合管理办法要求的技术文档
3. 呈现 3 条初始化路径：

```
请选择初始化方式：

[1] 使用管理办法文档 — 我有《过程管理办法》，请从中提取阶段和文档清单
[2] 使用内置预置流程 — 我没有管理办法，使用标准央企项目 5 阶段流程
[3] 直接生成单文档 — 我已有模板，只想生成一份文档

输入数字选择（1 / 2 / 3）：
```

### Step 1: 资源检查（Resource Check）

根据操作人的选择，检查所需资源是否齐全：

**路径 1（管理办法）：**
- 需要：管理办法文档（.docx 或 .pdf）
- AI 询问操作人上传
- 接收到文档后，分析提取阶段结构和文档清单
- 显示提取结果，操作人确认/修改

**路径 2（预置流程）：**
- 加载 `references/phase-preset.md`
- 显示标准 5 阶段结构和文档清单
- 操作人确认/裁剪阶段

**路径 3（单文档）：**
- 跳过阶段选择，直接进入模板选择

**所有路径通用检查：**
- Python 环境检查：`python3 --version`
- 依赖包检查：`pip list | grep python-docx` 等（非阻断，仅提示）
- 字体目录检查（非阻断，仅提示）

### Step 2: 收集项目信息

询问操作人提供项目基本信息（自由文本 + 结构化字段）：

**自由文本描述：**
- 项目概况（2~3 句）
- 主要业务目标
- 系统架构简介
- 关键约束条件

**结构化字段（AI 从自由文本中提取并请操作人确认）：**

```yaml
project:
  name: 国家电网配电自动化系统
  alias: 配网自动化
  client: 国家电网有限公司
  contractor: XXX 科技有限公司
  phase: 需求分析评审
  date: 2026-07-14
  reviewers: []
  version: V1.0
  key_terms:
    配电网: 10kV 及以下配电线路和设备
    FTU: 馈线终端单元
    DTU: 配变终端单元
```

操作人可手工录入项目信息，或提供一份项目信息文档让 AI 提取。

### Step 3: 逐文档生成

对每个需要生成的文档，执行以下子流程：

**文档生成子流程：**

1. **提取模板结构**
   - 操作人提供该文档的模板文件（.docx / .xlsx）
   - 运行 `scripts/extract_template.py` 提取标题树和样式定义
   - 显示模板结构给操作人确认

2. **范例分析**（可选但有最好）
   - 操作人提供优秀范例
   - 分析范例的内容结构和写作风格
   - 提取可复用的表达模板

3. **逐章生成 Markdown**
   - 按模板标题树，逐章节生成
   - 每章生成前提供本章节的上下文（模板指导文字、范例对应章节、项目信息）
   - 生成后追加到文档 Markdown 缓冲区
   - 全部章节完成后输出完整 `content.md`

4. **格式装配**
   - 运行 `scripts/assemble_docx.py` 或 `scripts/assemble_xlsx.py`
   - 输出最终 .docx / .xlsx 文件
   - 报告输出路径

5. **操作人审阅**
   - 显示生成结果的文件路径
   - 询问是否预览效果
   - 提供「修改某章节」「整体重生成」「继续下一份」等选项

### Step 4: 校验（Validation）

文档生成完毕后，执行两步校验：

**规则检查：**
- 章节完整性检查（模板 vs 生成的 Markdown）
- 编号连续性检查
- 占位符替换完整性检查
- 文件完整性检查（非空、可打开）

**AI 内容评审：**
- AI 通读生成文档的内容
- 评估内容质量、术语一致性、逻辑连贯性
- 输出评审报告（问题列表 + 严重程度 + 修改建议）

评审报告格式：
```yaml
review_summary:
  overall_score: 85/100
  issues:
    - severity: minor
      location: 第二章 2.3 节
      description: "术语'FTU'在首次出现时未给出全称"
      suggestion: "在首次出现处添加：馈线终端单元（FTU）"
    - severity: warning
      location: 第三章 3.1 节
      description: "功能列表与范例相比缺少非功能性需求"
      suggestion: "建议补充性能指标和安全要求小节"
  passed: true  # false 表示需要修改后再归档
```

### Step 5: 归档（Archive）

校验通过后：

1. 输出文件按管理办法中定义的文档顺序编号排序
2. 创建归档目录结构
3. 复制文件到归档目录
4. 同步生成归档清单 `archive_manifest.yaml`
5. 输出最终归档摘要给操作人

归档目录结构示例：
```
output/
└── 国家电网配电网管理办法-2026/
    ├── 01-需求分析评审/
    │   ├── 01_需求规格说明书_v1.0.docx
    │   ├── 02_项目计划书_v1.0.docx
    │   └── 03_评审纪要_v1.0.docx
    ├── 02-概要设计评审/
    │   ├── 01_概要设计说明书_v1.0.docx
    │   └── 02_接口规范_v1.0.docx
    └── archive_manifest.yaml
```

---

## 九、配置文件（config.yaml）

最小配置示例（安装 Skill 时框架自动检测）：

```yaml
output:
  base_dir: ./output

fonts:
  dir: ~/.fonts
```

完整配置项参考（所有字段均有默认值，可部分覆盖）：

```yaml
# ===== 输出目录 =====
output:
  base_dir: ./output             # 归档根目录

# ===== 字体 =====
fonts:
  dir: ~/.fonts                  # 中文字体目录
  required: []                   # 可指定必需字体列表（仅警告，不阻断）

# ===== 外部工具路径 =====
tools:
  libreoffice:
    binary: soffice              # LibreOffice 用于 TOC 渲染
  mermaid:
    binary: mmdc                 # mermaid-cli 用于流程图
  node:
    binary: node                 # Node.js 运行时
```

**注意：** LLM 配置不在 config.yaml 中。Agent 平台原生处理 LLM 调用。本 Skill 不需要也不控制 LLM 的选择和端点配置。

---

## 十、安装与依赖

### 全局依赖（一次性安装）

```bash
# Python 包
pip install python-docx openpyxl Pillow pyyaml docxtpl mammoth

# 可选工具
brew install --cask libreoffice            # TOC 渲染
npm install -g @mermaid-js/mermaid-cli      # 流程图生成
```

### 安装 Skill

```bash
# 通过 skills CLI
npx skills add https://github.com/CyanElegy/doc-orchestrator

# 或手动克隆到 skills 目录
git clone https://github.com/CyanElegy/doc-orchestrator
# 将 doc-orchestrator 目录放入 Agent 平台的 skills 路径
```

### 内网部署方案

| 依赖 | 内网方案 |
|------|----------|
| Python 包 | `pip download -r tools/requirements.txt` 在可联网机器下载，离线安装 |
| LibreOffice | IT 预装或离线安装包分发 |
| mermaid-cli | npm 离线包 |
| 字体（仿宋_GB2312、黑体、楷体_GB2312） | IT 部门字体库分发至 `fonts.dir` |

---

## 十一、目录结构

```
doc-orchestrator/
├── SKILL.md                   # ← 本文件，Skill 唯一入口
├── config.yaml                # 全局配置
├── phases/                    # 阶段定义（初始化时自动生成）
│   └── 01-requirements-review/
│       ├── phase.yaml         # 阶段元数据 + 文档清单
│       ├── templates/         # Word/Excel 模板
│       ├── examples/          # 优秀范例
│       ├── prompts/           # 生成提示词（可选）
│       └── assets/            # 标准图例、流程图
├── cache/                     # 解析缓存（自动管理，可删除）
├── scripts/                   # 加速器脚本
│   ├── extract_template.py    # 模板结构提取
│   ├── assemble_docx.py       # Word 格式装配
│   └── assemble_xlsx.py       # Excel 格式装配
├── shared/                    # 可选：共享库（如不提供 Python 包则无需）
├── references/                # 引用文件
│   ├── phase-schema.md        # 阶段 YAML Schema
│   ├── phase-preset.md        # 标准预设流程（5 阶段）
│   └── pitfalls.md            # 常见故障与恢复
├── tools/
│   └── requirements.txt       # Python 依赖列表
└── output/                    # 生成产物目录（自动创建）
```

**与 README.md 的关系：** README.md 是项目级别的概述文档，面向人工阅读；SKILL.md 是 Agent 平台使用的技能描述文档，是技能的唯一程序入口。SKILL.md 包含完整的流程定义、约束规则和执行协议，READMD.md 不包含这些内容。

---

## 十二、API 契约（内部接口定义）

### 12.1 模板提取契约

`scripts/extract_template.py` 的输入输出约定：

```
Input:  .docx 模板文件路径
Output: stdout 输出 JSON（含标题树、样式、页面设置）
Exit:   0 success, 1 error
```

JSON 结构已在 5.1 节定义。脚本必须输出合法的 JSON，不应输出多余日志到 stdout。

### 12.2 格式装配契约

`scripts/assemble_docx.py` 的输入输出约定：

```
Input:
  argv[1]: Markdown 文件路径（含 YAML front matter）
  argv[2]: 模板 .docx 路径
  argv[3]: 项目信息 YAML 路径
  argv[4]: 输出 .docx 路径

Output: 指定路径的 .docx 文件
Exit:   0 success, 1 error
```

`scripts/assemble_xlsx.py` 的输入输出约定：

```
Input:
  argv[1]: 数据 YAML 路径
  argv[2]: 模板 .xlsx 路径
  argv[3]: 输出 .xlsx 路径

Output: 指定路径的 .xlsx 文件
Exit:   0 success, 1 error
```

### 12.3 缓存契约

缓存文件存储在 `cache/` 目录，使用 sha256 哈希命名：

```
cache/
├── {sha256_of_file}.json       # 模板结构缓存
├── {sha256_of_analysis}.yaml   # 管理办法分析缓存
└── cache_manifest.yaml         # 缓存索引（file_path → hash 映射）
```

缓存失效条件：源文件的修改时间 > 缓存文件的创建时间。

### 12.4 Phase YAML 契约

`phases/` 目录下的阶段定义 YAML 结构：

```yaml
phase:
  id: 01-requirements-review        # 阶段唯一标识
  name: 需求分析评审                  # 阶段名称
  order: 1                          # 阶段顺序序号
  description: 项目需求分析与文档编制阶段
  regulation_ref: 国家电网信息化建设管理办法

documents:
  - id: requirements-spec
    name: 需求规格说明书
    type: docx                        # docx | xlsx
    template: templates/requirements-template.docx
    required: true                    # 是否必需文档
    depends_on: []                    # 依赖的文档 ID 列表
    chapters:                         # 文档章节标题树
      - id: ch01
        title: 第一章 引言
        level: 1
      - id: ch02
        title: 第二章 项目范围
        level: 1
      - id: ch02-01
        title: 2.1 业务目标
        level: 2

  - id: review-meeting-minutes
    name: 评审会议纪要
    type: docx
    template: templates/minutes-template.docx
    required: true
    depends_on: [requirements-spec]  # 在需求规格之后编写
```

---

## 十三、故障恢复与边缘情况

### 13.1 模板解析失败

**现象：** `extract_template.py` 输出错误或空 JSON。

**处理：**
1. 检查模板文件是否损坏（可手动在 Word 中打开）
2. 检查是否包含不受支持的格式（如宏、ActiveX 控件）
3. 失败时切换到手动替代路径：AI 读取模板文件结构，手工提取样式信息
4. 恢复后继续生成流程

### 13.2 装配脚本缺少依赖

**现象：** `ModuleNotFoundError: No module named 'docx'`

**处理：**
1. 提示操作人安装依赖：`pip install python-docx`
2. 提供手动替代路径：AI 在线编写 python-docx 代码
3. 安装完成后重新执行装配

### 13.3 章节内容过长

**现象：** AI 单次输出的 Markdown 超过上下文长度限制。

**处理：**
1. 将章节进一步拆分为子章节
2. 每子节单独生成，最后合并
3. 合并时确保章节编号连续

### 13.4 操作人中途退出

**现象：** 生成流程进行到一半，操作人需要暂停。

**处理：**
1. 已生成的 Markdown 内容保存到缓存
2. 已生成的 docx 保留在输出目录
3. 下次调用时检测到缓存，询问是否继续
4. 提供「查看已生成列表」「继续生成」「重新开始」选项

### 13.5 图片嵌入失败

**现象：** assemble_docx.py 报告图片文件未找到。

**处理：**
1. 检查图片路径是相对路径还是绝对路径
2. 相对于 Markdown 文件所在目录查找
3. 相对路径找不到时，提示操作人提供图片文件位置
4. 图片缺失不影响文字内容生成，在最终输出中保留占位

### 13.6 字体缺失

**现象：** 生成的 docx 在打开时字体显示异常（回退到默认字体）。

**处理：**
1. 提示操作人安装所需字体到 `fonts.dir` 指定的目录
2. 将字体嵌入 docx（assemble_docx.py 支持嵌入选项）
3. 提示操作人在本机安装该字体并重新打开文档

---

## 十四、提示词协议（Prompt Protocol）

AI 在生成文档内容时使用以下协议与自身交互：

### 章节生成提示词模板

```
你正在为文档「{doc_title}」生成第 {chapter_number} 章「{chapter_title}」的内容。

## 文档目的
{doc_description}

## 本章在文档中的位置
- 层级：Level {heading_level}
- 父章节：{parent_chapter}
- 模板指导文字：{guidance_text}

## 项目信息
{project_info_yaml}

## 参考范例
{example_excerpt}

## 生成约束
1. 输出格式仅为 Markdown（无样式、无格式信息）
2. 标题使用 # 标记，与指定层级对应
3. 表格使用 Markdown 表格语法
4. 图片引用格式：![alt](relative_path)
5. 交叉引用格式：见 [章节标题](#anchor)
6. 正文字号、字体、间距——不要在内容中指定
7. 术语使用需与项目信息中的 key_terms 一致
8. 语言风格：正式、简洁、技术文档风格

## 输出要求
- 仅输出 Markdown 正文，不要额外对话或解释
- 如果是第一章，包含 YAML front matter
```

### 评审提示词模板

```
你正在评审文档「{doc_title}」的内容质量。

## 评审标准
1. 术语一致性：与项目信息中的 key_terms 一致吗？
2. 章节完整性：是否覆盖了模板中定义的全部章节内容？
3. 内容深度：技术描述是否达到该文档类型应有的详细程度？
4. 逻辑连贯性：章节之间的逻辑过渡是否自然？
5. 合规性：符合央企文档的一般规范吗？

## 输出要求
以 YAML 格式输出评审报告（包含 overall_score、issues 列表、passed 布尔值）。
```

---

## 十五、版本与维护

### 版本记录

| 版本 | 日期 | 变更说明 |
|------|------|----------|
| 1.0 | 2026-07-14 | 初始版本 |

### 扩展指南

如需新增文档类型支持：

1. 在 `references/phase-preset.md` 中注册新的文档条目
2. 提供对应模板文件和范例（操作人提供或从已有项目迁移）
3. 如有新的复杂格式需求，扩展 `scripts/assemble_docx.py` 或 `scripts/assemble_xlsx.py`

### 废弃策略

- 版本 1.0 不做向后兼容保证
- 阶段定义 YAML 格式变更时，提供迁移脚本 `scripts/migrate_phase.sh`

---

## 十六、术语表

| 术语 | 英文 | 说明 |
|------|------|------|
| 管理办法 | regulation document | 央企发布的项目建设过程管理规范 |
| 模板 | template | 提供格式样式的 .docx/.xlsx 源文件 |
| 范例 | example | 已完成的优秀文档，作内容参考 |
| 阶段 | phase | 项目评审周期（需求、设计、测试等） |
| 文档清单 | document manifest | 一个阶段应产出的全部文档列表 |
| 占位符 | placeholder | 模板中待替换的项目信息标记（如《项目名称》） |
| 标题树 | heading tree | 模板中各级标题的结构化列表 |
| 格式装配 | format assembly | 将 Markdown 内容应用到模板样式的过程 |
| 规则检查 | rule check | 对生成文档的自动化格式和结构检查 |
| AI 评审 | AI review | AI 对文档内容的定性质量评估 |
| 归档清单 | archive manifest | 组织输出文件的索引 YAML |
| 全流程 | full process | 覆盖所有阶段的完整文档生成过程 |
| 直接生成 | direct generation | 单文档立即生成模式 |
| 单章节再生 | single-chapter regeneration | 仅重新生成文档的某一章节 |
