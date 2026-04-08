#!/usr/bin/env python3
# run_convert.py
"""
将 Word 文档中的公式转为 MathType 可用的 MathML 与 LaTeX。
用法:
  python run_convert.py <input.docx> [--output-dir DIR] [--latex] [--html] [--no-mathml]
"""
import argparse
import logging
import sys
from pathlib import Path

from word_formula_to_mathtype.extract import extract_omml_from_docx
from word_formula_to_mathtype.convert import convert_all_to_mathml, convert_all_to_latex
from word_formula_to_mathtype.output import write_mathml_file, write_latex_file

logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s: %(message)s",
)
logger = logging.getLogger(__name__)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="将 Word(.docx) 中的公式转为 MathType 可用的 MathML/LaTeX。"
    )
    parser.add_argument(
        "input_docx",
        type=str,
        help="输入的 .docx 文件路径",
    )
    parser.add_argument(
        "--output-dir", "-o",
        type=str,
        default=None,
        help="输出目录，默认与输入文件同目录",
    )
    parser.add_argument(
        "--no-mathml",
        action="store_true",
        help="不输出 MathML（默认输出）",
    )
    parser.add_argument(
        "--latex",
        action="store_true",
        help="同时输出 LaTeX 文件",
    )
    parser.add_argument(
        "--html",
        action="store_true",
        help="额外输出一个 HTML 预览（含 MathML）",
    )
    args = parser.parse_args()

    # MathML 默认开启，仅 --no-mathml 可关闭；若全部关闭则回退到 MathML
    do_mathml = not args.no_mathml
    if not do_mathml and not args.latex:
        logger.info("未指定任何输出格式，默认输出 MathML")
        do_mathml = True

    input_path = Path(args.input_docx)
    if not input_path.exists():
        logger.error("文件不存在: %s", input_path)
        return 1
    if input_path.suffix.lower() != ".docx":
        logger.error("仅支持 .docx 格式")
        return 1

    out_dir = Path(args.output_dir) if args.output_dir else input_path.parent
    base_name = input_path.stem
    logger.info("输入: %s，输出目录: %s，MathML=%s LaTeX=%s HTML=%s",
                input_path, out_dir, do_mathml, args.latex, args.html)

    try:
        omml_blocks = extract_omml_from_docx(str(input_path))
    except Exception as e:
        logger.exception("提取公式失败: %s", e)
        return 1

    if not omml_blocks:
        logger.warning("未在文档中发现公式（仅支持 Word 内置 OMML 公式，不支持 MathType 对象）")
        if do_mathml:
            write_mathml_file([], str(out_dir / f"{base_name}_formulas.xml"))
        if args.latex:
            write_latex_file([], str(out_dir / f"{base_name}_formulas.tex"))
        return 0

    logger.info("共提取 %d 个公式", len(omml_blocks))

    if do_mathml:
        mathml_result = convert_all_to_mathml(omml_blocks)
        write_mathml_file(
            mathml_result.items,
            str(out_dir / f"{base_name}_formulas.xml"),
            single_document=True,
        )
        logger.info("已写入 MathML: %s", out_dir / f"{base_name}_formulas.xml")
        logger.info("MathML 转换统计: %s", mathml_result.summary)
        if args.html:
            write_mathml_file(
                mathml_result.items,
                str(out_dir / f"{base_name}_formulas.html"),
                single_document=True,
            )
            logger.info("已写入 HTML 预览: %s", out_dir / f"{base_name}_formulas.html")

    if args.latex:
        latex_result = convert_all_to_latex(omml_blocks)
        write_latex_file(latex_result.items, str(out_dir / f"{base_name}_formulas.tex"))
        logger.info("已写入 LaTeX: %s", out_dir / f"{base_name}_formulas.tex")
        logger.info("LaTeX 转换统计: %s", latex_result.summary)

    return 0


if __name__ == "__main__":
    sys.exit(main())
