# 故障排查与环境配置

## 依赖安装

### 基础 Python 依赖

```bash
pip install -r requirements.txt
```

包含：`openai`, `python-docx`, `pyyaml`, `pymupdf`, `pillow`, `pytesseract`

### OCR 依赖（如需处理图片或扫描版 PDF）

```bash
pip install pytesseract pillow
```

同时需要安装 Tesseract-OCR 引擎及中文语言包：
- **macOS**: `brew install tesseract tesseract-lang`（确保包含 `chi_sim`）
- **Windows**: 从 [GitHub tesseract-ocr](https://github.com/UB-Mannheim/tesseract/wiki) 下载安装包，安装时勾选中文语言包
- **Ubuntu/Debian**: `sudo apt install tesseract-ocr tesseract-ocr-chi-sim`
- **CentOS/RHEL**: `sudo yum install tesseract tesseract-langpack-chi_sim`

### Pandoc（处理 .doc 转换）

- **macOS**: `brew install pandoc`（macOS 处理 .doc 默认使用内置 `textutil`，无需 pandoc）
- **Windows**: 从 [pandoc 官网](https://pandoc.org/installing.html) 下载安装包
- **Linux**: `sudo apt install pandoc` 或 `sudo yum install pandoc`

---

## 常见问题

### Q1: 运行脚本时提示缺少某个 Python 库

根据提示执行对应的 `pip install` 命令即可。建议直接使用：
```bash
pip install -r requirements.txt
```

各脚本核心依赖：
- `extract_document.py` → `python-docx`, `pymupdf`, `pillow`, `pytesseract`
- `ai_structure_analyzer.py` → `openai`, `pyyaml`, `python-docx`
- `docx_formatter.py` → `python-docx`

### Q2: Hosted 版是否需要配置 API Key？

**在 WorkBuddy / CodeBuddy 中使用**：**不需要**。Hosted 版的设计目标就是消除"脚本内再配一层 API Key"的套娃体验。AI 结构标注由宿主 AI 直接完成，Python 脚本只负责提取和排版。

**在独立环境（命令行 / CI）中使用**：如果需要一步完成提取 → AI 标注 → 排版，则需要自备 API Key。脚本会按以下优先级获取 API 配置：
1. 命令行参数 `--api-key`
2. 环境变量 `OFFICIAL_FORMATTER_API_KEY`
3. 配置文件 `~/.config/official-document-formatter/api_config.yaml`
4. 常见宿主环境变量（`OPENAI_API_KEY`、`ANTHROPIC_API_KEY`、`DEEPSEEK_API_KEY`）
5. 若均未找到，脚本会提示你使用 `--extract-only` 进入宿主 AI 模式，或直接提供 `--api-key`

### Q3: AI 标注结果（structured.md）识别不准确

可以在 `structured.md` 中直接修改钩子标记。常见需要手动修正的场景：
- 公文标题被误标为 `<!--HOOK:H1-->`
- 附件说明被误标为 `<!--HOOK:BODY-->`
- 落款和日期顺序颠倒
- 二级标题未拆分（应确保 `<!--HOOK:H2-->` 只包含小标题本身）

修改后保存 `structured.md`，再运行：
```bash
python scripts/docx_formatter.py 你的文件_structured.md
```

### Q4: 生成的 Word 文档中字体显示不正确

请检查系统是否安装了以下字体：
- `方正小标宋简体`（标题）
- `仿宋`（正文）
- `楷体`（二级标题）
- `黑体`（一级标题）
- `宋体`（页码）

如果缺少某个字体，`docx_formatter.py` 会自动降级到系统已有的兼容字体（如 `FangSong`、`SimHei`、`SimSun`）。如需手动调整降级链，请修改 `docx_formatter.py` 中的 `resolve_font()` 调用。

### Q5: 提取的 PDF 文字错乱或分段过度

- 如果是**扫描版 PDF**（文字无法选中），必须使用 OCR。提取后请重点检查 `*_raw.md`。
- 如果是**电子版 PDF**（文字可选中），`extract_document.py` 已经内置了过度断行合并逻辑。若效果不佳，可以手动调整 `clean_pdf_breaks()` 函数中的合并规则。

### Q6: 页码显示为 `{ PAGE }` 或空白

这是 Word 域代码的正常现象。打开文档后按 `Ctrl+A` 全选，再按 `F9` 更新域即可显示正确页码。

### Q7: 原文中的表格在输出后丢失

v3.1 已支持表格保留：
- **docx 源文件**：会按原文顺序提取表格，并保留在最终 Word 中。
- **PDF 源文件**：若使用 PyMuPDF >= 1.23，会尝试提取表格并追加到文档末尾（可能与正文部分重复，但不会丢失）。

如果表格列宽不理想，可在生成后手动微调。

### Q8: 如何一次性运行完整流水线

**独立模式**（已配置 API Key）：
```bash
python scripts/official_formatter.py -i 报告.docx
```

**宿主 AI 模式**（WorkBuddy / CodeBuddy）：
```bash
# 步骤 1：提取
python scripts/official_formatter.py -i 报告.docx --extract-only

# 步骤 2：由宿主 AI 生成 structured.md（Agent 自动完成）

# 步骤 3：排版
python scripts/official_formatter.py -i 报告.docx --structured 报告_structured.md
```

会在同一目录下生成：
- `公文排版后-报告.docx`（最终排版文档）
- `报告_raw.md` / `报告_structured.md`（中间文件）
- 独立模式下额外生成 `勘误建议表.docx`

---

## 获取帮助

如果问题仍未解决，请检查以下文件是否生成正常：
- `*_raw.md`：原始提取文本
- `*_structured.md`：AI 标注后的中间文件
- `公文排版后-*.docx`：最终排版的 Word 文档

通常情况下，问题出在阶段 1（提取失败）或阶段 2（AI 识别偏差）。阶段 3 的 Python 直接排版逻辑稳定，出错的概率较低。
