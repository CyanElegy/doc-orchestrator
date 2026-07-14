# Doc-Orchestrator

面向央企信息化建设项目的文档生成编排器。从 Word/Excel 模板和优秀范例出发，结合项目关键信息，自动生成符合《过程管理办法》标准的技术文档。

## 核心理念

> **格式权威来源只有模板 — AI 只负责内容，Python 精确装配格式。**

操作人只需提供：模板文件 + 优秀范例 + 项目关键信息。Skill 自动完成文档生成、格式装配、质量验收、归档排序。

## 功能概览

- 支持全流程 20+ 种文档类型，各评审阶段可独立运行
- 逐章节 AI 内容生成，精确控制上下文
- 从模板自动提取样式（字体、字号、编号风格、页边距、页眉页脚）
- 生成内容与格式彻底解耦（AI 产出纯内容 Markdown，Python 装配 docx/xlsx）
- 混合验收引擎：规则引擎（章节完整性、编号连续性）+ AI 内容评审
- 按管理办法规定的文件顺序归档，缺漏自动提示
- 内网友好：全部工具可本地离线运行

## 工作模式

| 模式 | 说明 | 适用场景 |
|------|------|----------|
| 完整流程模式 | 提供管理办法 → 提取阶段/文档清单 → 逐阶段生成全流程文档 | 新项目启动，需全流程管控 |
| 直接生成模式 | 提供模板 + 范例 → 直接生成单份文档 | 临时需要某份文档 |

## 架构

```
操作人输入 → Layer 0: 引导与发现 → Layer 1: 资源解析与缓存
→ Layer 2: AI 逐章内容生成(Markdown) → Layer 3: Python 格式装配(docx/xlsx)
→ Layer 4: 混合验收 + 归档输出
```

## 目录结构

```
doc-orchestrator/
├── SKILL.md                      # Skill 入口（AI 行为定义）
├── config.yaml                   # 全局配置
├── phases/                       # 阶段定义（初始化时自动生成）
│   └── 01-requirements-review/
│       ├── phase.yaml            # 阶段元数据 + 文档清单
│       ├── templates/            # Word/Excel 模板
│       ├── examples/             # 优秀范例
│       ├── prompts/              # 生成提示词
│       └── assets/               # 标准图例
├── scripts/                      # 加速脚本（3 个）
│   ├── extract_template.py       # 从 Word 模板提取样式结构
│   ├── assemble_docx.py          # Markdown + 模板 → 带样式的 .docx
│   └── assemble_xlsx.py          # 结构化数据 + 模板 → 带样式的 .xlsx
├── tools/
│   └── requirements.txt          # Python 依赖
└── output/                       # 生成产物
```

### 加速脚本

| 脚本 | 用途 | 调用方式 |
|------|------|----------|
| `scripts/extract_template.py` | 从 Word 模板提取章节结构/样式为 JSON | `python3 scripts/extract_template.py <模板文件>` |
| `scripts/assemble_docx.py` | 将 Markdown 内容按模板样式装配为 .docx | `python3 scripts/assemble_docx.py <content.md> <模板.docx> <项目.yaml> <输出.docx>` |
| `scripts/assemble_xlsx.py` | 将结构化数据按模板样式装配为 .xlsx | `python3 scripts/assemble_xlsx.py <数据.yaml> <模板.xlsx> <输出.xlsx>` |

## 快速开始

### 1. 环境要求

- Python 3.11+
- Claude Code 或其他 Agent 平台（LLM 配置由平台管理）

### 2. 安装依赖

```bash
cd doc-orchestrator
pip install -r tools/requirements.txt
```

### 3. 安装 Skill

```bash
npx skills add https://github.com/CyanElegy/doc-orchestrator
```

### 4. 可选工具（按需安装）

| 工具 | 用途 | 安装方式 |
|------|------|----------|
| LibreOffice 7.6+ | TOC 目录渲染、.vsdx 转换 | `brew install --cask libreoffice` |
| mermaid-cli | 流程图生成 | `npm install -g @mermaid-js/mermaid-cli` |
| 字体文件 | 公文标准字体 | 放入 `~/.fonts/` |

### 5. 配置

编辑 `config.yaml`，设置输出目录和字体目录：

```yaml
output_dir: ./output
fonts_dir: ~/.fonts
```

LLM 配置由 Agent 平台统一管理，不在 `config.yaml` 中设置。

### 6. 使用 SKILL.md

`SKILL.md` 是 Skill 的入口文件，定义了 AI 的行为规则、对话流程和工具调用方式。在 Claude Code 中加载该 Skill 后，AI 会自动按照 SKILL.md 定义的流程工作。

### 7. 在其他 Agent 平台使用

参考各平台的 Skill/Superpower 加载机制，将本项目注册为 Skill 即可。

## 使用流程

```
操作人                      Skill
  │                          │
  │  唤起 skill               │
  ├─────────────────────────►│  扫描可用阶段
  │                          │
  │◄─────────────────────────┤  展示可选文档清单
  │  选择目标                 │
  │  (全流程/阶段/单文档)      │
  ├─────────────────────────►│
  │                          │  环境预检
  │◄─────────────────────────┤  报告缺失项 + 安装指南
  │                          │
  │  提供资源文件              │
  │  (模板/范例/管理办法)      │
  ├─────────────────────────►│  解析 + 缓存
  │                          │
  │  补充项目信息              │
  ├─────────────────────────►│
  │                          │  逐章 AI 生成 Markdown
  │                          │  Python 装配 docx/xlsx
  │                          │  验收检查
  │                          │
  │◄─────────────────────────┤  验收报告 + 产物路径
  │  确认 / 修改              │
  │                          │
```

## 内网部署

全部工具链支持离线运行：

| 依赖 | 内网方案 |
|------|----------|
| Python 包 | `pip download -r requirements.txt` 离线安装 |
| LLM | Ollama + 离线导入模型文件，或内网 API |
| LibreOffice | IT 预装或离线安装包 |
| mermaid-cli | npm 离线包 |
| 字体 | IT 部门字体库 |
| Skill | 内网 Git 仓库直接拉取，`npx skills add <内网仓库地址>` |

## License

MIT
