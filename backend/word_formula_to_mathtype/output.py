# word_formula_to_mathtype/output.py
"""
将转换后的 MathML/LaTeX 写入文件，便于在 MathType 中使用。
"""
import logging
from pathlib import Path
from typing import List, Optional, Tuple

logger = logging.getLogger(__name__)


def _extract_value(item: Tuple) -> Tuple[int, Optional[str]]:
    """
    从结果项中提取索引和值，兼容旧格式 (index, value) 和新格式 (index, value, format_type)。
    """
    if len(item) >= 3:
        return item[0], item[1]
    return item[0], item[1]


def write_mathml_file(
    results: List[Tuple],
    output_path: str,
    single_document: bool = True,
) -> None:
    """
    将 MathML 结果写入文件。
    :param results: [(index, mathml_string or None), ...] 或 [(index, mathml_string or None, format_type), ...]
    :param output_path: 输出路径（.xml 或 .html）
    :param single_document: True 则输出一个 XML/HTML 包含所有公式；False 则按公式分文件
    """
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    suffix = path.suffix.lower()
    valid_count = sum(1 for item in results if _extract_value(item)[1] is not None)
    logger.debug("写入 MathML 到 %s，共 %d 条有效公式", path, valid_count)

    if single_document:
        if suffix == ".html":
            _write_mathml_html(results, path)
        else:
            _write_mathml_xml(results, path)
    else:
        base = path.with_suffix("")
        for item in results:
            i, mml = _extract_value(item)
            if mml is None:
                continue
            out = Path(f"{base}_eq{i}.xml")
            out.write_text(_wrap_mathml(mml), encoding="utf-8")
            logger.info("已写入 %s", out)


def _wrap_mathml(mml: str) -> str:
    """包装为完整 XML 文档便于 MathType 打开。"""
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<math xmlns="http://www.w3.org/1998/Math/MathML">\n'
        f"{mml}\n"
        "</math>"
    )


def _normalize_mml(mml: str) -> str:
    """去掉嵌入的 XML 声明与首尾空白，便于嵌入到外层 XML。"""
    s = mml.strip()
    if s.startswith("<?xml"):
        idx = s.find("?>")
        if idx != -1:
            s = s[idx + 2 :].strip()
    return s


def _write_mathml_xml(
    results: List[Tuple],
    path: Path,
) -> None:
    valid = []
    for item in results:
        i, m = _extract_value(item)
        if m is not None:
            valid.append((i, _normalize_mml(m)))
    
    if not valid:
        path.write_text(
            '<?xml version="1.0" encoding="UTF-8"?>\n<formulas count="0"/>\n',
            encoding="utf-8",
        )
        return
    parts = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        f'<formulas count="{len(valid)}">',
    ]
    for idx, mml in valid:
        parts.append(f'  <formula index="{idx}">')
        parts.append(f"    {mml}")
        parts.append("  </formula>")
    parts.append("</formulas>")
    path.write_text("\n".join(parts), encoding="utf-8")


_HTML_STYLE = """\
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
       background: #f5f7fa; color: #1d2939; line-height: 1.6; padding: 2rem; }
.container { max-width: 860px; margin: 0 auto; }
header { background: linear-gradient(135deg, #1e3a5f, #2563eb); color: #fff;
         padding: 1.5rem 2rem; border-radius: 12px; margin-bottom: 1.5rem; }
header h1 { font-size: 1.5rem; font-weight: 600; }
header p  { font-size: 0.875rem; opacity: 0.85; margin-top: 0.25rem; }
.card { background: #fff; border: 1px solid #e5e7eb; border-radius: 10px;
        padding: 1.25rem 1.5rem; margin-bottom: 1rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04); }
.card-title { font-size: 0.8rem; font-weight: 600; color: #6b7280;
              text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.75rem; }
.card-index { display: inline-block; background: #2563eb; color: #fff;
              font-size: 0.75rem; font-weight: 700; width: 1.5rem; height: 1.5rem;
              line-height: 1.5rem; text-align: center; border-radius: 50%;
              margin-right: 0.5rem; vertical-align: middle; }
.formula-render { background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px;
                  padding: 1rem 1.25rem; margin: 0.75rem 0; text-align: center;
                  font-size: 1.25rem; overflow-x: auto; }
.formula-render math { display: block; }
details { margin-top: 0.5rem; }
summary { cursor: pointer; font-size: 0.8rem; color: #6b7280; user-select: none; }
summary:hover { color: #2563eb; }
pre { background: #1e293b; color: #e2e8f0; padding: 0.75rem 1rem; border-radius: 6px;
      font-size: 0.78rem; line-height: 1.5; overflow-x: auto; margin-top: 0.5rem;
      white-space: pre-wrap; word-break: break-all; }
footer { text-align: center; font-size: 0.75rem; color: #9ca3af; margin-top: 2rem; }
"""


def _write_mathml_html(
    results: List[Tuple],
    path: Path,
) -> None:
    valid = []
    for item in results:
        i, m = _extract_value(item)
        if m is not None:
            valid.append((i, _normalize_mml(m)))
    
    total = len(results)
    success = len(valid)
    parts = [
        "<!DOCTYPE html>",
        '<html lang="zh-CN">',
        "<head>",
        '<meta charset="UTF-8"/>',
        '<meta name="viewport" content="width=device-width, initial-scale=1.0"/>',
        "<title>Word 公式转 MathType — 转换结果</title>",
        f"<style>{_HTML_STYLE}</style>",
        "</head>",
        "<body>",
        '<div class="container">',
        "<header>",
        "<h1>Word 公式转 MathType — 转换结果</h1>",
        f"<p>共 {total} 个公式，成功转换 {success} 个 · 点击「查看 MathML 源码」可复制到 MathType</p>",
        "</header>",
    ]
    for idx, mml in valid:
        parts.append('<div class="card">')
        parts.append(f'<div class="card-title"><span class="card-index">{idx}</span>公式 {idx}</div>')
        parts.append(f'<div class="formula-render"><math xmlns="http://www.w3.org/1998/Math/MathML">{mml}</math></div>')
        parts.append("<details><summary>查看 MathML 源码</summary>")
        parts.append(f"<pre>{_escape(mml)}</pre>")
        parts.append("</details>")
        parts.append("</div>")
    parts.append("<footer>由 Word 公式转 MathType 工具自动生成</footer>")
    parts.append("</div>")
    parts.append("</body></html>")
    path.write_text("\n".join(parts), encoding="utf-8")


def _escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def write_latex_file(
    results: List[Tuple],
    output_path: str,
) -> None:
    """将 LaTeX 结果写入 .tex 或 .txt 文件（每个公式一行或带编号）。"""
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    logger.debug("写入 LaTeX 到 %s，共 %d 条", path, len(results))
    lines = []
    for item in results:
        idx, tex = _extract_value(item)
        if tex is None:
            lines.append(f"% 公式 {idx}: (转换失败)")
        else:
            lines.append(f"% 公式 {idx}")
            lines.append(tex)
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")
