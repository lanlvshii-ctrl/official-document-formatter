# MEMORY.md — official-document-formatter

## 项目概况
- 中文公文智能排版技能（Hosted 版 v4.0），专为 WorkBuddy/CodeBuddy 设计
- 核心流程：文档提取(--extract-only) → 宿主AI结构标注(<!--HOOK:TYPE-->) → Python国标排版(--structured)
- 钩子协议：HTML注释 `<!--HOOK:TYPE-->`，9种类型：TITLE/H1/H2/H3/H4/BODY/ATTACHMENT/SIGNATURE/DATE
- 独立模式：支持 --api-key 一步完成（使用 ai_structure_analyzer.py 调用外部 LLM）

## 代码-文档同步状态（2026-04-17 确认）
- annotation_spec.md 与 docx_formatter.py 协议一致 ✅
- official_formatter.py 支持 --extract-only / --structured ✅
- ai_structure_analyzer.py 已移除 input() 非TTY挂起问题 ✅
- is_hosted_environment() 自动检测 ✅
- 2026-04-17 修复：H3/H4 加 bold=True（协议规定 H3仿宋加粗/H4仿宋加粗）；H1/H2 保持 bold=False（黑体字形已够粗不需加粗，楷体不需加粗）
- 2026-04-17 修复：split_h2_body → split_heading_body（同时处理 H1+H2 标题+正文拆分）
- 2026-04-17 新增：validate_hook_mapping() 自动校验钩子标注（根据编号格式检测并修正错误钩子），三层防线防 AI 层级偏移
- 用户确认：各级标题缩进2字符（0.74cm）是正确要求，不是 bug

## SKILL.md 优化记录
- 2026-04-17：从294行精简至153行；删除版本记录；文件结构和组件说明合并为表格；增加负向触发清单（博客/论文/技术文档/非中文）；description frontmatter 增加触发词覆盖
- 评估结果（skill-optimizer）：v4.0 工作区合并后评分 B+→A- 潜力

## evals 状态
- 8个用例（E001-E008）：含正触发2、负触发2、PDF+表格1、容错3
- 尚无 benchmark 运行记录

## 右键工具集成
- CLI 包装脚本：`~/.tools/bin/official_formatter_cli.py`
- Automator workflow：`~/Library/Services/公文格式整理.workflow`
- API 配置：`~/.config/official-document-formatter/api_config.yaml`（DeepSeek）
- 模式选择：弹对话框选择「精细/快速」
- 日志：`~/.tools/logs/official_formatter_YYYYMMDD.log`
