# Common Pitfalls — 常见问题与恢复

当发生错误或用户报告质量问题时，根据症状匹配对应条目。

## 1. 模板格式污染

**症状**：`extract.py` 输出 warnings 字段包含 SmartArt/文本框警告，或 skeleton 中缺少章节。

**原因**：模板含文本框、SmartArt 等复杂元素，`python-docx` 无法读取其中的文本。

**恢复**：
1. 在 Word 中打开模板，将文本框内容剪切粘贴为标准段落
2. SmartArt 内容在文本窗格中复制，删除 SmartArt 后粘贴为正文
3. 组织结构图等无法转换的元素截图，在 skeleton 中标记为 image unit
4. 重新运行 `python3 scripts/extract.py 模板.docx --output skeleton.json`

## 2. 骨架提取为空

**症状**：skeleton.json 的 units 数组为空或只含极少数 unit。

**原因**：模板使用自定义样式名（非 "Heading 1/2/3"）或内容全部在表格中。

**恢复**：
1. 检查模板是否使用 Word 内置标题样式
2. 若使用自定义样式：在 Word 中将样式改为内置 Heading 1-9
3. 若内容全在表格中：提取逻辑会自动将表格识别为 table unit

## 3. 上下文冲突

**症状**：Agent 发现多个上下文材料对同一信息描述不一致。

**原因**：不同来源的材料版本不一致或覆盖范围不同。

**恢复**：
1. Agent 暂停并列出冲突信息及各自来源
2. 用户确认以哪个来源为准
3. Agent 继续生成

## 4. 关键信息缺失

**症状**：skeleton 中的 placeholder unit 在上下文材料中找不到对应值。

**原因**：用户提供的上下文材料未覆盖所有占位符。

**恢复**：
1. Agent 列出所有未能填充的占位符
2. 用户手动提供对应值
3. Agent 更新 content.json 后继续装配

## 5. 脚本依赖缺失

**症状**：`ModuleNotFoundError: No module named 'docx'` 或类似错误。

**恢复**：
```bash
pip install python-docx openpyxl Pillow pyyaml
```

## 6. 表格列映射不确定

**症状**：Agent 无法确定上下文中的字段对应模板表格的哪一列。

**原因**：模板列标题与上下文材料中的字段名不一致。

**恢复**：
1. Agent 列出模板列标题和上下文中的候选字段
2. 用户手动指定映射关系
3. Agent 按映射生成表格数据

## 7. .doc 文件无法处理

**症状**：`extract.py` 报告无法转换 .doc 文件。

**原因**：系统未安装 LibreOffice。

**恢复**：
```bash
# macOS
brew install --cask libreoffice

# Ubuntu
sudo apt install libreoffice

# Windows
# 从 https://www.libreoffice.org/download/ 下载安装
```

## 快速诊断

| 症状关键词 | 匹配条目 |
|-----------|---------|
| SmartArt、文本框、章节缺失 | 1. 模板格式污染 |
| 骨架为空、无标题 | 2. 骨架提取为空 |
| 信息不一致、冲突 | 3. 上下文冲突 |
| 占位符未填充、《...》残留 | 4. 关键信息缺失 |
| ImportError、ModuleNotFoundError | 5. 脚本依赖缺失 |
| 表格列对应不上 | 6. 表格列映射不确定 |
| .doc 转换失败 | 7. .doc 文件无法处理 |
