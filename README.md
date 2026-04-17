# Official Document Formatter (Hosted Edition)

> 中文公文智能排版工具（Hosted 版）。在 WorkBuddy、CodeBuddy、Claude Code、Kimi CLI 等 AI 宿主环境中**零配置**即可使用，也可在命令行中以独立模式运行。

## 功能

- **零配置 Hosted 模式**：AI Agent 直接负责结构分析，Python 负责精确国标排版，无需用户手动申请/填写 API Key。
- **多格式输入**：支持 `.docx`、`.doc`、`.pdf`、`.txt`、`.md`、图片（OCR）。
- **符合 GB/T 9704**：自动处理页边距、文档网格（28 字 × 22 行）、字体、行距、页码、落款对齐。
- **字体降级**：若系统缺少 `方正小标宋简体` 或 `仿宋`，自动降级为兼容字体。
- **表格保留**：支持从 docx/pdf 提取表格并在 Word 中还原。

## 两种使用方式

### 方式一：Hosted 模式（WorkBuddy / CodeBuddy / Claude Code 等）

Agent 接到用户请求后，按 [`SKILL.md`](SKILL.md) 中的 **CRITICAL — Agent 操作指令** 执行即可：

1. 提取原始文本：
   ```bash
   python scripts/official_formatter.py -i "文档.docx" --extract-only
   ```
2. Agent 读取生成的 `*_raw.md`，按 `references/annotation_spec.md` 的钩子协议生成 `*_structured.md`。
3. 执行排版：
   ```bash
   python scripts/official_formatter.py -i "文档.docx" --structured "文档_structured.md"
   ```
4. 向用户交付 `公文排版后-文档.docx`。

### 方式二：独立模式（命令行 / CI）

若已有外部 LLM API Key，可一步完成：

```bash
pip install -r requirements.txt
python scripts/official_formatter.py -i "文档.docx" \
  --api-key "$API_KEY" \
  --base-url "https://api.openai.com/v1" \
  --model "gpt-4o"
```

或通过环境变量配置：

```bash
export OFFICIAL_FORMATTER_API_KEY="sk-xxx"
export OFFICIAL_FORMATTER_BASE_URL="https://api.openai.com/v1"
python scripts/official_formatter.py -i "文档.docx"
```

## 钩子协议

AI 与 Python 排版器之间通过 **HTML 注释钩子** 通信，每段以 `<!--HOOK:TYPE-->` 开头：

- `<!--HOOK:TITLE-->` — 公文主标题
- `<!--HOOK:BODY-->` — 正文段落
- `<!--HOOK:H1-->` / `<!--HOOK:H2-->` / `<!--HOOK:H3-->` / `<!--HOOK:H4-->` — 一至四级标题
- `<!--HOOK:ATTACHMENT-->` — 附件说明
- `<!--HOOK:SIGNATURE-->` — 发文单位
- `<!--HOOK:DATE-->` — 成文日期

详见 [`references/annotation_spec.md`](references/annotation_spec.md)。

## 文件结构

```
.
├── scripts/
│   ├── official_formatter.py      # 编排器（提取 / 排版 / 独立模式）
│   ├── extract_document.py        # 文档提取
│   ├── ai_structure_analyzer.py   # AI 语义分析（独立模式）
│   └── docx_formatter.py          # Python 直接国标排版
├── references/
│   ├── annotation_spec.md         # 钩子协议规范
│   ├── api_config.yaml.example    # 独立模式配置示例
│   ├── troubleshooting.md         # 常见问题
│   └── 格式要求.md                 # 格式参数来源
├── SKILL.md                       # Agent 必须遵守的操作指令
├── README.md                      # 本文件
└── requirements.txt               # Python 依赖
```

## 开源许可

本项目采用 **Apache License 2.0**。你可以在 GitHub 创建仓库时直接选择该许可证模板。
