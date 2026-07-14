# Common Pitfalls — 常见问题与恢复策略

## 概述

本文档记录 doc-orchestrator 的所有已知故障模式、症状、根因和恢复策略。当发生错误或操作人报告质量问题时，Skill 根据症状匹配条目给出修复方案。

每条结构：**故障名称** → **触发条件** → **症状** → **根因** → **恢复步骤** → **预防**

---

## 1. 模板格式污染

**触发条件**：模板中包含文本框（Text Box）、SmartArt、艺术字、ActiveX 控件等复杂元素。

**症状**：`docx_structure.py` 提取章节后部分章节缺失；生成的文档某些章节为空。

**根因**：`python-docx` 只能读取标准段落（Paragraph）和表格（Table），无法读取文本框、SmartArt 中的文本。

**恢复步骤**：
1. 在 Word 中打开模板，确认缺失内容是否在文本框或 SmartArt 中
2. 文本框内容：剪切后粘贴为标准正文段落
3. SmartArt 内容：在文本窗格复制所有文本 → 删除 SmartArt → 粘贴为编号段落
4. 组织结构图等无法转换的元素：截图放入 `assets/` 目录，在 phase.yaml 中用 `{{asset:文件名.png}}` 引用
5. 重新提取验证：`python shared/parsers/docx_structure.py --template templates/模板.docx`

**预防**：模板仅使用标准段落和表格；页眉页脚只放标题/Logo；不使用文本框作为内容容器。

---

## 2. 编号风格误判

**触发条件**：模板使用多级编号（如中文数字"一、二、三、"或混合格式），`extract_template.py` 未能正确识别。

**症状**：生成的文档标题编号不符合预期（如"1."而非"一、"）；多级标题编号混乱。

**根因**：Word 多级编号在不同版本和语言环境下存储格式不同，`python-docx` 对中文编号识别准确度有限。

**恢复步骤**：
1. 检查提取结果：`python shared/parsers/docx_structure.py --template 模板.docx --verbose`
2. 在 phase.yaml 中手动指定编号风格：
   ```yaml
   settings:
     numbering:
       heading: "一、"       # 一级标题
       sub_heading: "1.1"    # 二级标题
   ```
3. 常用编号参考：`一、/1.1`、`1/1.1`、`第1章/1.1`、`1./1.1`

**预防**：模板使用 Word 原生"多级列表"功能，而非手动编号；用 `--verbose` 验证提取结果。

---

## 3. 范例过时

**触发条件**：`examples/` 目录下的范例文档使用了旧术语、过期技术栈名称或已废弃的标准编号。

**症状**：AI 生成的文档中出现过时术语，引用不存在的标准，或模仿已改变的范例结构。

**根因**：`content_generator.py` 以范例为行文风格和术语参考。过时范例会被 AI 忠实模仿。

**恢复步骤**：
1. 确认过时内容：`python shared/parsers/example_distiller.py --example examples/范例.docx --list-chapters`
2. 替换为最新范例文件，重新运行蒸馏：`python shared/parsers/example_distiller.py --example examples/新范例.docx --output cache/examples/`
3. 暂无新范例时临时注释 phase.yaml 中的 examples 引用
4. 重新生成：`python shared/generators/content_generator.py --phase phases/01-requirements-review/ --regenerate`

**预防**：范例文件名包含日期（如`需求分析报告_范例_2026.docx`），每个项目周期更新范例。

---

## 4. 上下文溢出

**触发条件**：章节提示词包含过多范例文本、过长模板章节描述或过多上游上下文，导致 AI 上下文窗口超限。

**症状**：AI 返回截断内容（末尾章节缺失）；API 返回 `context_length_exceeded`；生成文件明显比预期短。

**根因**：提示词由章节描述 + 模板格式信息 + 范例蒸馏文本 + 上游信息 + 共享字段拼接而成，合计 token 超出 LLM 限制。

**恢复步骤**：
1. 定位超限章节：`grep "context_length\|token\|truncat" output/generation.log`
2. 在 phase.yaml 中减少 examples 引用数量（从 3 个减到 1 个）
3. 使用 `--max-chars 2000` 限制蒸馏文本长度
4. 章节数控制在 15 个以内，为每个 chapter 设置 `max_length`
5. 重新生成：`python shared/generators/content_generator.py --phase phases/01-requirements-review/ --output-doc requirements_report --regenerate`

**预防**：每章蒸馏控制在 2000 字符内；文档章节不超过 15 个；为每个 chapter 设 `max_length`。

---

## 5. TOC 不渲染

**触发条件**：生成 docx 中的目录域代码未被 Word/LibreOffice 自动更新。

**症状**：目录位置显示 `{ TOC \o "1-3" \h \z \u }` 而不是实际目录；在 macOS Quick Look 中不可见。

**根因**：`python-docx` 只能插入域代码无法触发更新。不同 Office 软件对域代码自动处理行为不同。

**恢复步骤**：
- **Word 手动**：右键目录 → "更新域" → "更新整个目录" → 保存
- **LibreOffice 批量**：`soffice --headless --convert-to docx:"MS Word 2007 XML" output/文档.docx`
- **toc_renderer.py 自动化**：`python shared/assemblers/toc_renderer.py --input output/文档.docx --output output/最终版.docx`

**预防**：在 phase.yaml 启用 `settings.toc.mode: static`（生成静态目录）；管线中加入 TOC 渲染步骤；验收检查中加入 TOC 验证。

---

## 6. 图片路径问题

**触发条件**：AI 内容中引用了 `{{asset:图片.png}}` 占位符，但对应图片不存在或格式不支持。

**症状**：`assemble_docx.py` 输出 "Image not found" 警告；生成 docx 中显示 `{{asset:xxx.png}}` 文本。

**根因**：`assemble_docx.py` 在 `shared/assets/` 和阶段 `assets/` 目录查找图片。文件缺失、路径错误或格式不支持（SVG/WebP）导致替换失败。

**恢复步骤**：
1. 查找缺失图片清单：`grep -rn "{{asset:" output/ --include="*.md"` 或查看日志 `grep "Image not found" output/generation.log`
2. 将缺失图片放入 `phases/NN-phase/assets/` 或 `shared/assets/`
3. 支持的格式：PNG(推荐)、JPEG、GIF、BMP、EMF/WMF。SVG 先转 PNG：
   ```bash
   pip install cairosvg
   cairosvg shared/assets/架构图.svg -o shared/assets/架构图.png
   ```
4. 重新装配：`python shared/assemblers/docx_assembler.py --markdown output/xxx.md --template 模板.docx --output output/修正版.docx`

**预防**：所有图片放在统一 assets 目录；装配前运行 `image_handler.py verify` 预检。

---

## 7. 验收标准与实际产出不匹配

**触发条件**：phase.yaml 定义需生成 N 份文档但实际只生成了 M 份（M < N），常见于 carry-over 文档场景。

**症状**：Layer 4 验收报告显示 checklist 缺失；输出目录文件数量少于预期；"carry-over 文档未找到来源"。

**根因**：上游阶段未执行；`depends_on` 中引用的 output id 与上游不一致；AI 或装配器出错；操作人跳过了阶段。

**恢复步骤**：
1. 运行验收检查确认缺失项：`python shared/validators/rule_checker.py --phase phases/03-tech-review/ --output-dir output/`
2. carry-over 缺失：先生成上游阶段 `python main.py --phase 01-requirements-review`
3. id 不匹配：检查 `depends_on` 引用的 output id 在上游 phase.yaml 是否存在且大小写一致
4. 单文档重试：`python shared/generators/content_generator.py --phase phases/03-tech-review/ --output-doc test_report --retry`
5. 重新装配验收：`python main.py --phase 03-tech-review --outputs test_report --skip-existing`

**预防**：每阶段完成后运行验收检查；保持 `depends_on` 和上游 `outputs[].id` 一致；使用流水线顺序生成。

---

## 8. Phase YAML 格式错误

**触发条件**：手动编辑 `phase.yaml` 引入 YAML 语法错误。

**症状**：`phase_scanner.py` 无法发现阶段；Skill 提示"未找到可用阶段"；抛出 `yaml.YAMLError`。

**根因**：缩进不一致、冒号后缺空格、列表项缩进错误、字符串含特殊字符未加引号。

**恢复步骤**：
1. yamllint 快速检查：`pip install yamllint && yamllint phases/01-requirements-review/phase.yaml`
2. Python 验证：`python -c "import yaml; yaml.safe_load(open('phases/01-requirements-review/phase.yaml'))"`
3. 常见错误速查：

   | 错误写法 | 正确写法 |
   |---------|---------|
   | `name:需求分析` | `name: 需求分析` |
   | 缩进使用 Tab | 缩进使用 2 空格 |
   | `value: 是:否`（冒号在值内未引） | `value: "是:否"` |
   | options 下缩进 4 格 | options 下缩进 2 格 |

4. 修复后重新扫描：`python shared/lib/phase_scanner.py --scan`

**预防**：使用 YAML 专用编辑器；编辑后立即运行 `yaml.safe_load()` 验证；预提交钩子中加 yamllint。

---

## 9. 依赖缺失

**触发条件**：Python 依赖包未安装或版本不兼容。

**症状**：`ImportError` / `ModuleNotFoundError`：`No module named 'docx'`（docx_assembler）、`openpyxl`（xlsx_assembler）、`PIL`（image_handler）。

**恢复步骤**：
1. 安装全部依赖：`pip install -r tools/requirements.txt --upgrade`
2. 创建干净虚拟环境：`python -m venv .venv && source .venv/bin/activate && pip install -r tools/requirements.txt`
3. 检查 Python 版本：需 3.11+
4. 子模块依赖速查：docx_assembler→python-docx、xlsx_assembler→openpyxl、image_handler→Pillow

**预防**：项目 README 标注 `pip install -r tools/requirements.txt`；requirements.txt 固定版本号；Skill 启动时运行 `preflight_checker.py` 检测依赖。

---

## 10. 内网环境 LLM 不可用

**触发条件**：无互联网访问的内网环境中，API 类型 LLM 后端无法连接。

**症状**：`llm_client.py` 报 `ConnectionError` / `timeout`；内容生成阶段长时间无响应后失败。

**根因**：央企内网严格网络控制。Ollama 服务未在内网部署，或 config.yaml 配置了需外网的 API 端点。

**恢复步骤**：
1. **本地部署 Ollama**：
   - 外网下载模型：`ollama pull qwen2.5:72b` → `ollama export qwen2.5:72b -o model.tar.gz`
   - 或从 Modelscope 下载 GGUF 文件
   - 通过 U 盘拷贝到内网机器
   - 内网导入：`ollama create qwen2.5:72b -f ./model.tar.gz`
   - 启动：`OLLAMA_HOST=0.0.0.0:11434 ollama serve`
   - 修改 `config.yaml`：`backend: ollama; endpoint: http://内网IP:11434; model: qwen2.5:72b`

2. **内网 API 代理**：
   ```yaml
   llm:
     backend: openai
     endpoint: http://内网API网关/v1
     api_key: 内网分配的Key
   ```

3. **无 LLM 降级模式**（仅装配模板，AI 生成跳过）：
   ```bash
   python main.py --offline-mode --phase 01-requirements-review
   ```

**预防**：项目初始化时运行 `preflight_checker.py --check-llm`；准备内外网两套 config 文件；预装 Ollama 和模型；保持纯 Python 依赖可离线运行。

---

## 故障快速诊断索引

| 症状关键词 | 匹配条目 | 优先级 |
|-----------|---------|--------|
| 章节缺失、提取为空 | 1. 模板格式污染 | 高 |
| 编号错误、标题编号不对 | 2. 编号风格误判 | 中 |
| 术语过时、模仿旧文档 | 3. 范例过时 | 中 |
| 截断、context_length_exceeded | 4. 上下文溢出 | 高 |
| {TOC}、目录不显示 | 5. TOC 不渲染 | 中 |
| {{asset}}、Image not found | 6. 图片路径问题 | 低 |
| 验收报缺文档 | 7. 验收标准与实际不匹配 | 高 |
| YAML error、找不到阶段 | 8. Phase YAML 格式错误 | 高 |
| ImportError、ModuleNotFoundError | 9. 依赖缺失 | 高 |
| timeout、ConnectionError | 10. 内网 LLM 不可用 | 高 |

**优先级**：高 = 阻止流程继续，需立即修复；中 = 影响产出质量，建议修复后重跑；低 = 影响外观，可在最终输出前修复。

---

## 故障升级路径

当以上步骤无法解决时：
1. **查看日志**：检查 `output/generation.log` 完整错误栈
2. **收集现场**：`python tools/collect_diagnostics.py --output ./diag.zip`（收集 config、phase.yaml、模板摘要、错误日志、依赖版本）
3. **联系支持**：将 `diag.zip` 发送给工具链维护团队
4. **临时绕过**：将该阶段标记为"手动处理"，直接编辑生成的 Markdown 后手动用模板装配
