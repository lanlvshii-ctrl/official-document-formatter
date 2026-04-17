#!/usr/bin/env python3
"""
official_formatter.py
一键公文格式化：输入任意源文档，直接输出符合国标格式的 final.docx
内部流程：提取 → AI 三轮标注/自检 → Python 直接排版 → 生成勘误建议表
"""

import argparse
import os
import sys
from pathlib import Path

# 导入同级模块
import extract_document
import ai_structure_analyzer
import docx_formatter


def is_hosted_environment() -> bool:
    """检测当前是否运行在 WorkBuddy / CodeBuddy 等宿主环境中"""
    return bool(os.environ.get("CODEBUDDY_COPILOT_INTERNET_ENVIRONMENT"))


def main():
    parser = argparse.ArgumentParser(
        description="官方文档格式化器：源文档 → final.docx（+ correction.docx）"
    )
    parser.add_argument("input", help="源文档路径（支持 docx/doc/pdf/txt/md/图片）")
    parser.add_argument("-o", "--output-dir", help="输出目录（默认与源文件同目录）")
    parser.add_argument("--fast", action="store_true", help="使用单轮快速模式（跳过三轮自检，不生成勘误表）")
    parser.add_argument("--no-correction", action="store_true", help="即使使用精细模式，也不生成 correction.docx")
    parser.add_argument("--api-key", help="直接指定 API Key（独立模式使用）")
    parser.add_argument("--base-url", help="直接指定 API Base URL")
    parser.add_argument("--model", help="直接指定模型名称")
    parser.add_argument("--model-reasoning", help="直接指定深度思考模型名称（用于 Round 1）")
    parser.add_argument("--extract-only", action="store_true", help="仅执行文档提取，输出 raw.md（宿主 AI 模式）")
    parser.add_argument("--structured", help="传入已生成的 structured.md 路径，直接跳过 AI 分析进入排版阶段")
    args = parser.parse_args()

    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists():
        print(f"❌ 文件不存在: {input_path}")
        sys.exit(1)

    if args.output_dir:
        output_dir = Path(args.output_dir).expanduser().resolve()
    else:
        output_dir = input_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    base_name = input_path.stem
    raw_path = output_dir / f"{base_name}_raw.md"
    structured_path = output_dir / f"{base_name}_structured.md"
    final_path = output_dir / f"公文排版后-{base_name}.docx"
    correction_path = output_dir / "勘误建议表.docx"

    # -----------------------------------------------------------------------
    # 步骤 1：提取
    # -----------------------------------------------------------------------
    print("=" * 50)
    print("步骤 1/3：文档提取")
    print("=" * 50)
    try:
        text = extract_document.classify_and_extract(input_path)
    except Exception as e:
        print(f"❌ 提取失败: {e}")
        sys.exit(1)

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    while "\n\n\n" in text:
        text = text.replace("\n\n\n", "\n\n")

    raw_path.write_text(text, encoding="utf-8")
    cn_chars = extract_document.count_chinese_chars(text)
    print(f"✅ 提取完成: {raw_path}（中文字符数: {cn_chars}）")
    if cn_chars < 50:
        print("⚠️  警告: 中文字符数过少，请检查源文件质量。")

    if args.extract_only:
        print("\n" + "=" * 50)
        print("提取阶段已完成（--extract-only）")
        print("=" * 50)
        print(f"中间文件: {raw_path}")
        print("下一步：请使用宿主 AI 生成 structured.md，或直接提供 --structured 参数继续排版。")
        sys.exit(0)

    # -----------------------------------------------------------------------
    # 步骤 2：AI 标注（三轮收敛）
    # -----------------------------------------------------------------------
    # 如果传入了 --structured，直接跳过 AI 分析
    if args.structured:
        structured_path = Path(args.structured).expanduser().resolve()
        if not structured_path.exists():
            print(f"❌ 指定的 structured.md 不存在: {structured_path}")
            sys.exit(1)
        structured_text = structured_path.read_text(encoding="utf-8")
        corrections = []
    else:
        print("\n" + "=" * 50)
        print("步骤 2/3：AI 结构标注")
        print("=" * 50)

        if args.api_key:
            credentials = {
                "api_key": args.api_key.strip(),
                "base_url": (args.base_url or "https://api.openai.com/v1"),
                "model": (args.model or "gpt-4o"),
                "model_reasoning": args.model_reasoning or None,
            }
        else:
            credentials = ai_structure_analyzer.get_api_credentials()
            # 在宿主环境中，若未显式传入 --api-key，优先使用宿主 AI 模式
            # 避免使用仅对 Coding Agent 有效的 key（如 kimi-for-coding）导致 403
            if credentials is None or (is_hosted_environment() and not args.api_key):
                if credentials and is_hosted_environment():
                    print("\nℹ️  检测到当前处于 WorkBuddy/CodeBuddy 宿主环境，且未显式提供 --api-key。")
                    print("   为避免使用仅限宿主内部使用的 API 凭证，将自动进入宿主 AI 模式。")
                else:
                    print("\n⚠️  未配置外部 API Key，无法本地完成 AI 结构标注。")
                print("=" * 50)
                print("宿主 AI 模式指引：")
                print(f"  1. 当前已生成中间文件: {raw_path}")
                print("  2. 请使用 WorkBuddy/CodeBuddy 宿主 AI 读取该 raw.md 文件")
                print("  3. 让宿主 AI 根据 hooks 协议生成 structured.md 并保存到:")
                print(f"     {structured_path}")
                print("  4. 然后运行:")
                print(f"     python scripts/official_formatter.py \"{input_path}\" --structured \"{structured_path}\"")
                print("=" * 50)
                sys.exit(0)
            if args.base_url:
                credentials["base_url"] = args.base_url
            if args.model:
                credentials["model"] = args.model
            if args.model_reasoning:
                credentials["model_reasoning"] = args.model_reasoning

        if args.fast:
            structured_text = ai_structure_analyzer.run_fast_mode(text, credentials)
            corrections = []
        else:
            structured_text, corrections = ai_structure_analyzer.run_thorough_mode(text, credentials)

        structured_path.write_text(structured_text, encoding="utf-8")
        print(f"✅ 标注完成: {structured_path}")

        if not args.fast and not args.no_correction and corrections is not None:
            ai_structure_analyzer.generate_correction_docx(
                corrections, correction_path, doc_title="公文勘误建议表"
            )
            if ai_structure_analyzer.HAS_DOCX:
                print(f"✅ 勘误表生成: {correction_path}")

    # -----------------------------------------------------------------------
    # 步骤 3：Python 直接排版
    # -----------------------------------------------------------------------
    print("\n" + "=" * 50)
    print("步骤 3/3：Python 直接排版")
    print("=" * 50)
    if not docx_formatter.HAS_DOCX:
        print("❌ 缺少 python-docx，无法生成 Word 文档")
        sys.exit(1)

    paragraphs = docx_formatter.parse_structured_md(structured_text)
    if not paragraphs:
        print("❌ 没有解析到任何段落")
        sys.exit(1)

    doc = docx_formatter.Document()
    docx_formatter.apply_format(doc, paragraphs)
    doc.save(str(final_path))
    print(f"✅ 排版完成: {final_path}")
    print(f"   共 {len(paragraphs)} 个段落")

    print("\n" + "=" * 50)
    print("全部完成！")
    print("=" * 50)
    print(f"最终文档: {final_path}")
    if not args.fast and not args.no_correction and ai_structure_analyzer.HAS_DOCX:
        print(f"勘误建议: {correction_path}")
    print(f"中间文件: {raw_path}, {structured_path}")
    print('\n💡 提示: 若打开 Word 后页码显示为空白或 "{ PAGE }"，请按 Ctrl+A 再按 F9 更新域。')


if __name__ == "__main__":
    main()
