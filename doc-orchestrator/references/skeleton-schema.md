# Skeleton & Content JSON Schema

`skeleton.json` 和 `content.json` 是 extract.py 和 assemble 脚本之间的数据契约。

---

## skeleton.json

由 `extract.py` 从模板文件中提取。Agent 和装配脚本都读取此文件。

### 顶层结构

```json
{
  "meta": {
    "type": "docx | xlsx",
    "source": "模板文件名.docx",
    "extracted_at": null
  },
  "warnings": [
    "模板包含 SmartArt 图形，其中文本无法自动提取"
  ],
  "units": [...],
  "unit_count": 18,
  "summary": {
    "total_units": 18,
    "headings": 12,
    "tables": 1,
    "placeholders": 3,
    "fixed_units": 8,
    "generated_units": 10
  }
}
```

`warnings` 为 `null` 时表示无警告。

### Unit 类型

每个 unit 描述模板中的一个结构元素，按文档顺序排列。

**所有 unit 都有 `para_index` 字段**，记录该元素在 `document.body` 子元素中的位置索引。`assemble_docx.py` 使用此字段定位对应的 XML 元素进行原地修改。

#### heading — 标题

```json
{
  "id": "h-001",
  "type": "fixed",
  "element": "heading",
  "level": 1,
  "text": "引言",
  "numbering": "第一章",
  "para_index": 22
}
```

- `type`: 始终为 `"fixed"`——标题结构来自模板，不可变
- `level`: 1-9，对应 Word 标题层级
- `numbering`: 编号前缀（如 `"第一章"`、`"1.1"`、`"一、"`）。无编号时为空字符串

#### paragraph (fixed) — 固定文字

```json
{
  "id": "p-002",
  "type": "fixed",
  "element": "paragraph",
  "text": "本报告依据《国家电网信息化建设管理办法》编制。",
  "para_index": 25,
  "spacer": false,
  "cover": false
}
```

模板中的固定文字、封面间距段落等。**Agent 不需要为这些 unit 生成 content**——装配脚本保留原始段落不变。

- `spacer`: 可选，封面页中的空白间距段落
- `cover`: 可选，标记该段落属于封面页区域（第一个 H1 之前的所有内容）

#### paragraph (generated) — 待生成段落

```json
{
  "id": "p-003",
  "type": "generated",
  "element": "paragraph",
  "description": "编写目的正文",
  "para_index": 27,
  "cover": false,
  "has_x_placeholder": false
}
```

需要 Agent 创作内容的段落。`description` 来自模板文本，可能是简短提示、`【...】` 指引文字、`XX/××××` 占位符模式或封面信息。

- `cover`: 可选，标记该段落属于封面页
- `has_x_placeholder`: 可选，标记该段落包含 `XX` 或 `××××` 等占位符模式

Agent 在 content.json 中为这些 unit 提供内容：
```json
{"id": "p-003", "content": "本文档旨在明确..."}
```

#### placeholder — 占位符

```json
{
  "id": "ph-004",
  "type": "placeholder",
  "element": "placeholder",
  "pattern": "《项目名称》",
  "key": "项目名称"
}
```

模板中的《...》标记。Agent 应从上下文材料中找到对应值替换。

Agent 在 content.json 中为这些 unit 提供替换值：
```json
{"id": "ph-004", "content": "配电自动化系统二期"}
```

#### table — 表格

```json
{
  "id": "t-005",
  "type": "generated",
  "element": "table",
  "table_index": 0,
  "headers": ["序号", "功能模块", "功能描述", "优先级"],
  "rows": 5
}
```

- `headers`: 列标题列表
- `rows`: 数据行数（不含标题行），0 表示模板中无预填数据

Agent 在 content.json 中为表格提供数据（JSON 二维数组）：
```json
{
  "id": "t-005",
  "content": [
    ["1", "配电终端管理", "实现FTU/DTU远程监控和配置下发", "高"],
    ["2", "数据采集服务", "实时采集配电网运行数据", "高"]
  ]
}
```

#### image — 图片占位符

```json
{
  "id": "img-006",
  "type": "generated",
  "element": "image",
  "alt_text": "系统架构图"
}
```

模板中的图片位置。Agent 可在 content 中提供图片路径，或使用 `{{asset:path}}` 在段落内容中引用图片。

#### sheet — Excel 工作表

```json
{
  "id": "sheet-001",
  "type": "generated",
  "element": "sheet",
  "sheet_name": "工作量估算",
  "header_row": 1,
  "headers": ["模块名称", "功能描述", "人天数", "负责人"],
  "data_rows": 0
}
```

Agent 在 content.json 中提供表格数据（与 table unit 格式相同）。

---

## content.json

由 Agent 生成。只包含需要填充的 unit——`fixed` 类型的 unit 不需要出现在 content.json 中。

### 顶层结构

```json
{
  "units": [
    {"id": "p-003", "content": "本文档旨在明确国家电网配电自动化系统二期..."},
    {"id": "ph-004", "content": "配电自动化系统二期"},
    {"id": "t-005", "content": [["1","配电终端管理","实现FTU/DTU远程监控","高"]]}
  ]
}
```

每个 unit 通过 `id` 与 skeleton.json 中的 unit 匹配。

### content 字段格式

| skeleton 中 unit 的 element | content 字段类型 | 说明 |
|---------------------------|-----------------|------|
| `heading` (fixed) | 不需要出现在 content 中 | 标题结构不可变 |
| `paragraph` (generated) | string | 该段落的完整文本 |
| `placeholder` | string | 替换值 |
| `table` | 二维数组 `[[...], [...]]` | 数据行（不含标题行） |
| `image` | string | 图片文件路径 |
| `sheet` | 二维数组 `[[...], [...]]` | Excel 数据行 |

### 部分替换

content.json 可以只包含部分 unit——例如只修改一个章节时：

```json
{
  "units": [
    {"id": "p-015", "content": "修改后的第三章内容..."}
  ]
}
```

装配脚本只替换 `p-015`，其他 unit 保持不变。`fixed` unit 始终使用 skeleton 中的原始文本。
