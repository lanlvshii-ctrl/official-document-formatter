---
name: official-document-formatter
description: 中文公文智能排版工具——将任何非结构化的中文文稿一键生成严格符合国标格式的 Word 文档。当用户需要处理公文、通知、意见、报告、请示、制度、管理办法、红头文件、GB/T 9704、党政机关公文格式、方正小标宋、仿宋排版时，必须使用此技能。不用于普通博客、学术论文、非中文文档或技术文档排版。
---

# Official Document Formatter

将非结构化中文文稿，通过 **宿主 AI 语义识别 + Python 国标排版**，一键生成符合 GB/T 9704 标准的 Word 文档。

## 触发信号

**适用场景**：公文、通知、意见、报告、请示、制度、管理办法、红头文件、GB/T 9704、方正小标宋、仿宋排版

**不适用场景**：普通博客、个人日记、学术论文、技术文档、非中文文档

## CRITICAL — 核心规则

1. **钩子格式**：必须使用 `<!--HOOK:TYPE-->`，机器易识别，人眼不混淆。
2. **宿主 AI 优先**：在 WorkBuddy / CodeBuddy 中运行时，AI 结构标注由宿主 AI 完成，无需用户配置 API Key。
3. **一键交付**：运行 `official_formatter.py`，直接输出 `公文排版后-原文件名.docx`。
4. **字体降级**：无 `仿宋` → `FangSong`；无 `方正小标宋简体` → `黑体`。
5. **页码居中**：宋体四号 "- {页码} -"，不区分奇偶页。

## CRITICAL — Agent 操作指令（必须遵守）

当用户要求排版公文时，**必须**按以下 5 步执行：

### 步骤 1：提取原始文本
```bash
python scripts/official_formatter.py -i "<用户输入文件路径>" --extract-only
```
生成 `<文件名>_raw.md`。

### 步骤 2：读取 raw.md 并生成 structured.md
1. 读取步骤 1 生成的 `*_raw.md` 内容。
2. 根据 `references/annotation_spec.md` 的 **HTML 注释钩子协议**，为每个段落打上 `<!--HOOK:TYPE-->` 标记。
3. **必须遵守**：钩子后不加空格；`<!--HOOK:H2-->` 只含小标题本身，后跟正文必须拆分为独立 `<!--HOOK:BODY-->`；落款和日期分别用 `<!--HOOK:SIGNATURE-->` 和 `<!--HOOK:DATE-->` 独立成段。
4. 写入 `<文件名>_structured.md`。

**⚠️ IMPORTANT — 编号→钩子映射（极易出错！）：**
- `一、` `二、` `三、` → 必须标 `<!--HOOK:H1-->`（黑体），**绝对不能**标 H2
- `（一）` `（二）` `（三）` → 必须标 `<!--HOOK:H2-->`（楷体），**绝对不能**标 H3
- `1．` `2．` `3．` → 必须标 `<!--HOOK:H3-->`（仿宋加粗），**绝对不能**标 H4
- `（1）` `（2）` `（3）` → 必须标 `<!--HOOK:H4-->`（仿宋加粗）
- **判断依据是编号格式，不是内容层级！** 即使公文只有两级标题，"一、"也是 H1，"（一）"也是 H2

### 步骤 3：生成勘误建议（可选但推荐）
检查原文错别字、句法错误、逻辑错误、时间/数字错误、文风问题，在回复中呈现或生成 `勘误建议表.docx`。

### 步骤 4：执行排版
```bash
python scripts/official_formatter.py -i "<用户输入文件路径>" --structured "<生成的_structured.md路径>"
```
输出 `公文排版后-<原文件名>.docx`。

### 步骤 5：交付
- `公文排版后-<原文件名>.docx`（主交付物）
- `<文件名>_structured.md`（供复核）
- 勘误建议

**注意**：若环境配置了外部 API Key（如 `--api-key` 传入），可走独立模式一步完成。否则**默认走宿主 AI 模式**，不要让用户配置 API Key。

## 核心技术路线

### 模式一：宿主 AI 模式（推荐）

```
源文档 → [阶段1] --extract-only → raw.md
  → [阶段2] 宿主 AI 读取 raw.md，生成带 <!--HOOK:TYPE--> 的 structured.md
  → [阶段3] --structured → 公文排版后-原文件名.docx
```

### 模式二：独立模式（需自备 API Key）

`python scripts/official_formatter.py -i 输入.docx --api-key $KEY` → 一步完成提取 → AI 标注 → 排版。

## 钩子协议

| 钩子 | 含义 | 格式处理 |
|------|------|----------|
| `<!--HOOK:TITLE-->` | 公文主标题 | 方正小标宋简体 二号 居中 行距 32 磅 |
| `<!--HOOK:BODY-->` | 普通正文 | 仿宋 三号 首行缩进 2 字符 行距 28 磅 两端对齐 |
| `<!--HOOK:H1-->` | 一级标题（一、二、） | 黑体三号 首行缩进 2 字符 |
| `<!--HOOK:H2-->` | 二级标题（（一）（二）） | 楷体 三号 首行缩进 2 字符 |
| `<!--HOOK:H3-->` | 三级标题（1．2．） | 仿宋 三号 首行缩进 2 字符 |
| `<!--HOOK:H4-->` | 四级标题（（1）（2）） | 仿宋 三号 首行缩进 2 字符 |
| `<!--HOOK:ATTACHMENT-->` | 附件说明 | 仿宋 三号 首行缩进 2 字符 |
| `<!--HOOK:SIGNATURE-->` | 发文单位 | 仿宋 三号 右对齐 与下段同页 |
| `<!--HOOK:DATE-->` | 发文日期 | 仿宋 三号 右对齐 |

## 示例

**输入（raw.md 片段）**：
```markdown
关于修订《XX集团有限公司公务用车管理办法》的通知

各子公司、各部门：

为进一步规范集团公务用车管理，依据《党政机关公务用车管理办法》及相关规定，结合集团实际，对原办法进行修订。

一、总则

（一）适用范围。本办法适用于集团总部及下属全资、控股子公司的公务用车管理。

XX集团有限公司办公室
2025年3月15日
```

**AI 标注输出（structured.md 片段）**：
```markdown
<!--HOOK:TITLE-->关于修订《XX集团有限公司公务用车管理办法》的通知

<!--HOOK:BODY-->各子公司、各部门：

<!--HOOK:BODY-->为进一步规范集团公务用车管理，依据《党政机关公务用车管理办法》及相关规定，结合集团实际，对原办法进行修订。

<!--HOOK:H1-->一、总则

<!--HOOK:H2-->（一）适用范围。
<!--HOOK:BODY-->本办法适用于集团总部及下属全资、控股子公司的公务用车管理。

<!--HOOK:SIGNATURE-->XX集团有限公司办公室

<!--HOOK:DATE-->2025年3月15日
```

## 文件结构

| 文件 | 说明 |
|------|------|
| `scripts/official_formatter.py` | 编排器：支持 `--extract-only` / `--structured` / 独立模式 |
| `scripts/extract_document.py` | 文档提取（docx/doc/pdf/txt/md/图片/表格） |
| `scripts/ai_structure_analyzer.py` | AI 语义分析（独立模式专用，宿主模式不使用） |
| `scripts/docx_formatter.py` | Python 直接国标排版 |
| `references/annotation_spec.md` | 钩子协议规范 |
| `references/格式要求.md` | 用户提供的具体格式参数 |
| `references/troubleshooting.md` | 环境配置与故障排查 |
| `references/api_config.yaml.example` | API 配置示例（独立模式） |
| `requirements.txt` | Python 依赖 |

## 交付物

| 文件 | 说明 |
|------|------|
| `公文排版后-原文件名.docx` | **主交付物**：已排版的 Word 文档 |
| `*_structured.md` | 带钩子的中间文件，便于复核 |
| `*_raw.md` | 原始提取文本 |
| `勘误建议表.docx` | 勘误建议（独立模式自动生成 / 宿主模式由 AI 提供） |

## 常见问题

| 现象 | 解决方案 |
|------|----------|
| 字体显示为替代字体 | 安装字体包；`docx_formatter.py` 自动降级到 `仿宋`/`黑体`/`楷体` |
| PDF 文字错乱 | 优先提供 Word 源文件；扫描版检查 `*_raw.md` |
| 段落格式不对 | 检查 `*_structured.md` 中该段的 `<!--HOOK:TYPE-->` 是否正确 |
| 页码不显示 | Word 中 `Ctrl+A` → `F9` 更新域 |
| 提示缺少 API Key | Hosted 版正常行为，让 Agent 读取 raw.md 生成 structured.md 后再排版 |

详细排查见 `references/troubleshooting.md`。
