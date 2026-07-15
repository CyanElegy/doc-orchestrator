# Doc-Orchestrator

从 Word/Excel 模板和上下文材料生成成品文档的 Agent Skill。

## 核心理念

> **确定的事代码做，不确定的事 Agent 做，Skill 是两者之间的契约。**

操作人只需提供：模板文件 + 上下文材料（范例文档、技术方案、需求说明、自然语言描述等）。Agent 负责理解上下文、生成内容，Python 脚本负责提取骨架和装配格式。

## 安装

```bash
npx skills add https://github.com/CyanElegy/doc-orchestrator --skill doc-orchestrator
```

## 使用

在 Claude Code 中 @ 唤起：

```
@doc-orchestrator 帮我生成需求规格说明书
```

Agent 会引导你提供模板和上下文材料，然后自动完成骨架提取、内容生成、格式装配和质量审阅。

## 环境要求

- Python 3.11+（首次使用时 Agent 会自动检测并给出安装指引）
- pip 包：python-docx、openpyxl、Pillow、pyyaml（Agent 会在缺失时提示安装命令）
- LibreOffice（可选，仅用于 .doc 文件转换）

## 目录结构

```
doc-orchestrator/
├── SKILL.md                      # Skill 定义（Agent 行为规范）
├── references/
│   ├── skeleton-schema.md        # skeleton.json / content.json 数据契约
│   └── pitfalls.md               # 常见问题与恢复
└── scripts/
    ├── extract.py                # 骨架提取（.docx/.xlsx/.doc/.xls）
    ├── assemble_docx.py          # Word 装配
    ├── assemble_xlsx.py          # Excel 装配
    └── numbering_patterns.py     # 编号检测共享模块
```

## License

MIT
