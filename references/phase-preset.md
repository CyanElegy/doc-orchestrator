# 阶段预设：央企信息化建设标准流程

## 概述

定义央企信息化建设项目的标准 4 阶段评审流程。当操作人未提供《过程管理办法》且不清楚评审流程时，Skill 自动加载此预设作为默认回退方案。如果后续提供了正式管理办法，预设被覆盖，优先按管理办法解析。

## 阶段总览

| 编号 | 阶段 | 说明 |
|------|------|------|
| 01 | 需求分析评审 | 确认需求的完整性和可行性 |
| 02 | 试用评审 | 验证系统是否满足试用要求 |
| 03 | 技术评审 | 检查技术方案和实现情况 |
| 04 | 验收评审 | 最终验收，确认项目交付物 |

## DAG 依赖关系

```
  01 需求分析评审
       │ tech_solution
       ▼
  02 试用评审
       │ trial_plan
       ▼
  03 技术评审 ◄──── requirements_report, tech_solution
       │ development_report, tech_report, test_report
       ▼
  04 验收评审
```

---

## 01-requirements-review（需求分析评审）

```yaml
id: requirements-review
name: 需求分析评审
description: 对项目需求进行评审，确认需求的完整性和可行性
```

### 输入问题

| ID | 名称 | 类型 | 必填 | 说明 |
|----|------|------|------|------|
| module_list | 系统功能模块清单 | text | 是 | 列出系统包含的所有功能模块 |
| user_roles | 用户角色分类 | text | 是 | 用户角色和权限划分 |
| security_level | 安全等级 | select | 是 | 选项：非密/秘密/机密/绝密 |
| user_scale | 预计用户规模 | text | 否 | 如"500-1000人" |
| system_relation | 与现有系统的关系 | text | 否 | 如"独立新建"、"替换旧系统" |

安全等级选项：`non_classified:非密`、`secret:秘密`、`classified:机密`、`top_secret:绝密`

### 输出文档

| ID | 名称 | 格式 | 章节 |
|----|------|------|------|
| requirements_report | 需求分析报告 | docx | 项目背景、项目范围、用户角色、功能需求、非功能需求、系统接口 |
| tech_solution | 总体技术方案 | docx | 总体架构、技术选型、部署架构、安全设计 |
| workload_estimate | 工作量评估表 | xlsx | Sheet:工作量估算，列:序号/模块名称/功能描述/人天数/负责人/备注 |

---

## 02-trial-review（试用评审）

```yaml
id: trial-review
name: 试用评审
description: 对系统进行试用评审，验证系统是否满足试用要求
depends_on:
  - phase: requirements-review
    outputs:
      - tech_solution
```

依赖说明：01 阶段的 `总体技术方案` 自动继承到本阶段，操作人审阅确认后直接使用。

### 输入问题

| ID | 名称 | 类型 | 必填 | 说明 |
|----|------|------|------|------|
| trial_scope | 试用范围 | text | 是 | 试用的功能模块和业务范围 |
| trial_environment | 试用环境（硬件/软件/网络） | text | 是 | 硬件配置、操作系统、中间件、网络条件 |
| trial_team | 试用人员及角色 | text | 是 | 参与试用的人员及其角色分工 |
| trial_dates | 试用起止时间 | text | 是 | 试用开始和结束日期 |

### 输出文档

| ID | 名称 | 格式 | 章节 |
|----|------|------|------|
| trial_application | 试用评审申请表 | docx | 基本信息、试用内容、申请理由 |
| trial_plan | 试用方案 | docx | 试用目标、组织架构、试用计划、通过标准 |
| emergency_plan | 应急处置方案 | docx | 风险识别、应对措施、恢复方案 |
| tech_solution | 总体技术方案 | docx | [从01继承] 系统架构、部署方案 |

---

## 03-tech-review（技术评审）

```yaml
id: tech-review
name: 技术评审
description: 检查技术方案和实现情况，确认技术可行性
depends_on:
  - phase: requirements-review
    outputs:
      - requirements_report
      - tech_solution
  - phase: trial-review
    outputs:
      - trial_plan
```

### 输入问题

| ID | 名称 | 类型 | 必填 | 条件 |
|----|------|------|------|------|
| tech_change | 技术架构是否变更 | select | 是 | 永远显示。选项：是/否 |
| change_description | 变更说明 | text | 是 | 仅当 tech_change=是 时显示 |

### 输出文档

| ID | 名称 | 格式 | 章节 |
|----|------|------|------|
| development_report | 研制报告 | docx | 研制过程、关键技术、实现细节、质量保障 |
| tech_report | 技术报告 | docx | 技术概述、性能评估、安全性评估、兼容性说明 |
| test_report | 测试报告 | docx | 测试概述、测试用例、缺陷统计、测试结论 |
| requirements_report | 需求分析报告 | docx | [从01继承] |
| tech_solution | 总体技术方案 | docx | [从01继承] |
| trial_plan | 试用方案 | docx | [从02继承] |

---

## 04-acceptance-review（验收评审）

```yaml
id: acceptance-review
name: 验收评审
description: 最终验收，确认项目交付物完整并满足合同要求
depends_on:
  - phase: tech-review
    outputs:
      - development_report
      - tech_report
      - test_report
```

### 输入问题

| ID | 名称 | 类型 | 必填 | 说明 |
|----|------|------|------|------|
| acceptance_method | 验收方式 | select | 是 | 选项：会议验收/现场验收/函审验收 |
| expert_list | 验收专家组名单 | text | 否 | 专家姓名、单位、职称 |

### 输出文档

| ID | 名称 | 格式 | 章节 |
|----|------|------|------|
| acceptance_application | 验收申请报告 | docx | 项目概况、完成情况、验收申请、附件清单 |
| project_summary | 项目总结报告 | docx | 项目背景、实施过程、建设成果、经验与教训、后续工作 |
| development_report | 研制报告 | docx | [从03继承] |
| tech_report | 技术报告 | docx | [从03继承] |
| test_report | 测试报告 | docx | [从03继承] |

---

## 预设加载与目录结构

当 Skill 加载此预设时，自动在 `phases/` 下创建 4 个阶段目录：

```
phases/
├── 01-requirements-review/
│   ├── phase.yaml            # 由预设数据生成
│   ├── templates/            # 操作人需放置模板
│   ├── examples/             # 操作人可放置范例
│   └── assets/
├── 02-trial-review/...
├── 03-tech-review/...
└── 04-acceptance-review/...
```

### 生成规则

1. 根据本文档数据为每个阶段创建 `phase.yaml`，`depends_on` 自动配置
2. 输出文档的 `template` 路径留空（操作人暂未提供），Skill 提示操作人放置模板文件
3. 如果操作人提供了管理办法，预设文件被覆盖，按管理办法重新解析

### 模板文件命名规范

```
templates/
├── 需求分析报告模板.docx
├── 总体技术方案模板.docx
├── 工作量评估表模板.xlsx
├── 试用评审申请表模板.docx
├── 试用方案模板.docx
├── 应急处置方案模板.docx
├── 研制报告模板.docx
├── 技术报告模板.docx
├── 测试报告模板.docx
├── 验收申请报告模板.docx
└── 项目总结报告模板.docx
```

如果操作人暂无模板，Skill 使用内置通用模板（仅含标题和基本样式）初步生成。

### 范例文件建议

```
examples/
├── 需求分析报告_范例.docx
├── 总体技术方案_范例.docx
├── 研制报告_范例.docx
└── 项目总结报告_范例.docx
```

范例帮助 AI 准确理解文档风格、术语体系和行文规范。

---

## 阶段交互流程

```
Skill: "检测到您没有提供过程管理办法，是否使用标准央企评审流程？
        预设流程包含以下 4 个阶段：
        1. 需求分析评审 → 2. 试用评审 → 3. 技术评审 → 4. 验收评审"

操作人: "好的，使用标准流程。"

Skill: "开始需求分析评审阶段。
        请输入以下信息：
        [必填] 系统功能模块清单
        [必填] 用户角色分类
        [必填] 安全等级（非密/秘密/机密/绝密）
        [可选] 预计用户规模
        [可选] 与现有系统的关系"

操作人: (提供信息)
Skill: "信息已收集，开始生成需求分析评审阶段文档..."
```

---

## 自定义与扩展

- **增减阶段**：手动在 `phases/` 中添加或删除阶段目录
- **修改阶段**：编辑对应阶段的 `phase.yaml`，增改 `input_questions` 和 `outputs`
- **调整流程**：修改 `depends_on` 改变文档继承关系
- **完全替换**：提供正式管理办法文件后，预设自动退化为替补方案
