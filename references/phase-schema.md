# Phase Schema — phase.yaml 规范

## 概述

`phase.yaml` 是阶段定义的核心配置文件，位于 `phases/NN-phase-name/phase.yaml`。它描述评审阶段的元数据、输入信息采集规则、输出文档清单及阶段间依赖。Skill 通过扫描 `phases/` 目录下的所有 `phase.yaml` 发现可用阶段。

## 文件位置与编号

```
phases/
├── 01-requirements-review/
│   └── phase.yaml
├── 02-trial-review/
│   └── phase.yaml
```

目录命名必须以 `NN-` 开头（两位数字 + 连字符），控制阶段显示顺序和依赖解析顺序。编号缺漏不会报错但阶段列表可能乱序。

---

## YAML Schema（完整字段）

```yaml
id: string                  # [必填] 阶段唯一标识，kebab-case
name: string                # [必填] 阶段中文名称
version: string             # [必填] 语义化版本号，如 "1.0.0"
description: string         # [可选] 阶段描述
order: integer              # [可选] 排序序号，覆盖目录编号

# --- 输入信息采集 ---
input_questions:
  - id: string              # [必填] 问题唯一标识
    name: string            # [必填] 问题中文名称
    type: text | select | multiselect | number  # [必填]
    required: boolean       # [可选] 默认 true
    default: any            # [可选] 默认值
    placeholder: string     # [可选] 输入提示
    description: string     # [可选] 对AI的补充说明
    options:                # select/multiselect 必填
      - id: string          # 选项标识
        label: string       # 选项显示文本
    validation:             # [可选] 校验规则
      min: number           # 最小值/最小长度
      max: number           # 最大值/最大长度
      pattern: string       # 正则（text类型）
      min_items: integer    # 最少选择数（multiselect）
      max_items: integer    # 最大选择数（multiselect）
    conditions:             # [可选] 显示条件
      field: string         # 依赖的input_question id
      operator: eq | neq | in
      value: any

# --- 输出文档清单 ---
outputs:
  - id: string              # [必填] 文档唯一标识
    name: string            # [必填] 文档中文名称
    format: docx | xlsx | pptx  # [必填]
    template: string        # [必填] 模板路径（相对 phase.yaml 目录）
    prompt_template: string # [可选] 提示词模板路径
    chapters:               # docx 章节定义
      - id: string          # [必填] 章节标识
        title: string       # [必填] 章节标题
        description: string # [必填] 对AI的章节内容指引
        required: boolean   # [可选] 默认 true
        max_length: integer # [可选] 最大字符数
        examples:           # [可选] 范例引用
          - file: string    # 范例文件路径
            chapter: string # 范例中对应章节id
    examples:               # [可选] 完整文档范例
      - file: string
    excel_sheets:           # xlsx 定义
      - name: string        # Sheet名称
        columns:
          - header: string  # 列头
            type: text | number
            width: integer
            required: boolean
    shared_fields:          # [可选] 引用上游输出字段
      - from_phase: string
        field: string

# --- 跨文档共享配置 ---
shared_outputs:
  - id: string
    name: string
    description: string

# --- 依赖声明 ---
depends_on:
  - phase: string           # 上游阶段 id
    outputs:                # 需要继承的输出文档
      - string              # output id
      - string

# --- 高级配置 ---
settings:
  concurrency: integer      # 并行生成文档数，默认 1
  llm:
    model: string           # 覆盖全局 LLM 模型
    temperature: float      # 默认 0.7
    max_tokens: integer     # 默认 4096
  numbering:
    heading: string         # 一级标题编号，如"一、"
    sub_heading: string     # 二级标题编号，如"1.1"
  retry:
    max_attempts: integer   # 默认 3
    backoff: float          # 默认 2.0
```

---

## 字段详解

### input_questions — types 行为

| Type | 行为 | 关键配置 |
|------|------|---------|
| `text` | 自由文本输入 | `validation.pattern`, `validation.min/max`（字符长度） |
| `select` | 单选下拉 | 必须提供 `options[]` |
| `multiselect` | 多选复选框 | 必须提供 `options[]`，可用 `min_items/max_items` |
| `number` | 数字输入 | `validation.min/max` 控制范围 |

#### conditions 条件显示示例

```yaml
- id: tech_change
  name: 技术架构是否变更
  type: select
  options:
    - id: "yes"
      label: "是"
    - id: "no"
      label: "否"
- id: change_description
  name: 变更说明
  type: text
  conditions:
    field: tech_change
    operator: eq
    value: "yes"   # 仅当选择"是"时显示
```

### outputs — 章节与模板

`chapters[]` 定义 docx 文档的逐章节内容。AI 逐章生成，确保每章聚焦且上下文受限。

```yaml
chapters:
  - id: background
    title: 项目背景
    description: 项目立项背景、业务现状和建设目标
    required: true
    max_length: 2000
    examples:
      - file: examples/需求分析报告_范例.docx
        chapter: background
```

`excel_sheets` 定义 xlsx 的多 Sheet 结构：

```yaml
excel_sheets:
  - name: 工作量估算
    columns:
      - header: 模块名称; type: text; width: 30; required: true
      - header: 人天数;   type: number; width: 15; required: true
```

### depends_on — 依赖与文档继承

1. **DAG 构建**：Skill 根据 `depends_on` 构建有向无环图，检测循环依赖并报错
2. **文档继承（Carry-over）**：上游 `outputs` 中声明的文档 id 如果在当前阶段也有定义（同名 id），该文档成品自动继承到当前阶段，无需重新生成
3. **上下文注入**：上游 `input_questions` 的回答和 `shared_fields` 自动注入下游 AI 提示词

```yaml
depends_on:
  - phase: 01-requirements-review
    outputs:
      - requirements_report
      - tech_solution
```

- 不支持循环依赖（启动时环检测）
- 一个阶段可依赖零个（起始阶段）、一个或多个上游
- 多个阶段可同时依赖同一个上游

### shared_outputs — 跨文档复用字段

在统一位置定义项目级字段（如项目名称、项目编号、编制单位），所有文档的 AI 提示词自动包含。

```yaml
shared_outputs:
  - id: project_name
    name: 项目名称
    description: 项目全称，如"某集团信息化建设一期工程"
  - id: prepared_by
    name: 编制单位
    description: 文档编制单位名称
```

### settings — 高级运行配置

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| concurrency | integer | 1 | 并行生成的文档数 |
| llm.model | string | 继承全局 | 覆盖该阶段的 LLM 模型 |
| llm.temperature | float | 0.7 | AI 生成温度 |
| llm.max_tokens | integer | 4096 | 最大生成 token |
| numbering.heading | string | 从模板推断 | 一级标题编号风格 |
| numbering.sub_heading | string | 从模板推断 | 二级标题编号风格 |
| retry.max_attempts | integer | 3 | 生成失败重试次数 |
| retry.backoff | float | 2.0 | 重试退避因子 |

---

## 完整示例：需求分析评审

```yaml
id: requirements-review
name: 需求分析评审
version: "1.0.0"
description: 对项目需求进行评审，确认需求的完整性和可行性

input_questions:
  - id: module_list
    name: 系统功能模块清单
    type: text
    required: true
    placeholder: 请列出系统包含的所有功能模块
  - id: user_roles
    name: 用户角色分类
    type: text
    required: true
    placeholder: 如：系统管理员、普通用户、审计员
  - id: security_level
    name: 安全等级
    type: select
    required: true
    options:
      - id: non_classified;  label: 非密
      - id: secret;          label: 秘密
      - id: classified;      label: 机密
      - id: top_secret;      label: 绝密
  - id: user_scale
    name: 预计用户规模
    type: text
    required: false
    placeholder: 如：500-1000人

outputs:
  - id: requirements_report
    name: 需求分析报告
    format: docx
    template: templates/需求分析报告模板.docx
    chapters:
      - id: background;              title: 项目背景
      - id: scope;                   title: 项目范围
      - id: functional_requirements; title: 功能需求
      - id: non_functional_req;      title: 非功能需求
  - id: tech_solution
    name: 总体技术方案
    format: docx
    template: templates/总体技术方案模板.docx
    chapters:
      - id: architecture;    title: 系统架构
      - id: technology_stack; title: 技术选型
  - id: workload_estimate
    name: 工作量评估表
    format: xlsx
    template: templates/工作量评估表模板.xlsx
    excel_sheets:
      - name: 工作量估算
        columns:
          - header: 模块名称; type: text;   width: 30; required: true
          - header: 功能描述; type: text;   width: 50; required: true
          - header: 人天数;   type: number; width: 15; required: true
          - header: 备注;     type: text;   width: 40; required: false

shared_outputs:
  - id: project_name;  name: 项目名称; description: 项目全称
  - id: prepared_by;   name: 编制单位; description: 文档编制单位名称
```

---

## 完整示例：试用评审（含 depends_on）

```yaml
id: trial-review
name: 试用评审
version: "1.0.0"
depends_on:
  - phase: requirements-review
    outputs:
      - tech_solution

input_questions:
  - id: trial_scope
    name: 试用范围
    type: text; required: true
  - id: trial_environment
    name: 试用环境（硬件/软件/网络）
    type: text; required: true
  - id: trial_team
    name: 试用人员及角色
    type: text; required: true
  - id: trial_dates
    name: 试用起止时间
    type: text; required: true

outputs:
  - id: trial_application
    name: 试用评审申请表
    format: docx
    template: templates/试用评审申请表模板.docx
    chapters:
      - id: basic_info;    title: 基本信息
      - id: trial_content; title: 试用内容
  - id: trial_plan
    name: 试用方案
    format: docx
    template: templates/试用方案模板.docx
    chapters:
      - id: objectives;  title: 试用目标
      - id: plan;        title: 试用计划
      - id: criteria;    title: 通过标准
  - id: emergency_plan
    name: 应急处置方案
    format: docx
    template: templates/应急处置方案模板.docx
    chapters:
      - id: risk_list;         title: 风险列表
      - id: response_measures; title: 应对措施
  - id: tech_solution
    name: 总体技术方案
    format: docx
    template: templates/总体技术方案模板.docx
    chapters:
      - id: architecture; title: 系统架构
      - id: deployment;   title: 部署方案
```

---

## 常见错误速查

| 错误 | 后果 | 检查方法 |
|------|------|---------|
| 目录名缺 `NN-` 前缀 | 阶段不被识别 | `ls phases/` 检查命名 |
| depends_on 循环依赖 | 启动时报错退出 | 确保 DAG 无环 |
| select/multiselect 缺 options | 采集时无选项 | 检查 options 列表 |
| template 路径用绝对路径 | 找不到模板 | 路径应相对 phase.yaml 目录 |
| input_questions id 重复 | 采集时覆盖 | 同一阶段内 id 唯一 |
| outputs 同名 id 跨阶段 | 用于 carry-over 继承 | 确认与上游一致 |
