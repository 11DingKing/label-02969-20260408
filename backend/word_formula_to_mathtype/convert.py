# word_formula_to_mathtype/convert.py
"""
将 OMML 转为 MathType 可用的 MathML 或 LaTeX。
依赖 docx-equation 的 OMML2MML.XSL 与 mml2tex。
支持 fallback 机制：MathML 失败尝试 LaTeX，都失败保存原始 OMML。
"""
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Tuple

from docx_equation.omml import omml2mml, omml2tex

logger = logging.getLogger(__name__)


@dataclass
class ConvertResult:
    """批量转换的结果与统计。"""
    items: List[Tuple[int, Optional[str], Optional[str]]] = field(default_factory=list)
    total: int = 0
    success: int = 0
    failed: int = 0
    failed_indices: List[int] = field(default_factory=list)
    failed_omml: List[Tuple[int, str]] = field(default_factory=list)

    @property
    def success_rate(self) -> str:
        if self.total == 0:
            return "0%"
        return f"{self.success / self.total * 100:.1f}%"

    @property
    def summary(self) -> str:
        if self.total == 0:
            return "无公式需要转换"
        if self.failed == 0:
            return f"全部成功：{self.success}/{self.total}"
        return (
            f"成功 {self.success}/{self.total}（{self.success_rate}），"
            f"失败 {self.failed} 个（编号：{', '.join(str(i) for i in self.failed_indices)}）"
        )


def omml_to_mathml(omml_xml: str) -> Optional[str]:
    """
    将单段 OMML XML 转为 MathML 字符串（MathType 可导入/粘贴）。
    :param omml_xml: 单个 m:oMath 或 m:oMathPara 的 XML 字符串
    :return: MathML 字符串，失败或空输入返回 None
    """
    if not omml_xml or not omml_xml.strip():
        logger.debug("跳过空 OMML 输入")
        return None
    try:
        result = omml2mml(omml_xml)
        if result and result.strip():
            return result.strip()
        logger.debug("OMML -> MathML 结果为空")
    except Exception as e:
        logger.warning("OMML -> MathML 转换失败: %s", e)
    return None


def omml_to_latex(omml_xml: str) -> Optional[str]:
    """
    将单段 OMML 转为 LaTeX（MathType 支持粘贴 LaTeX）。
    :param omml_xml: 单个 OMML XML 字符串
    :return: LaTeX 字符串，失败或空输入返回 None
    """
    if not omml_xml or not omml_xml.strip():
        logger.debug("跳过空 OMML 输入")
        return None
    try:
        result = omml2tex(omml_xml)
        s = str(result).strip() if result is not None else ""
        if s and s not in ("$", ""):
            return s
        logger.debug("OMML -> LaTeX 结果为空或占位: %r", s)
    except Exception as e:
        logger.warning("OMML -> LaTeX 转换失败: %s", e)
    return None


def convert_all_to_mathml(omml_blocks: List[str]) -> ConvertResult:
    """将多段 OMML 转为 MathML，返回带统计的结果。"""
    result = ConvertResult(total=len(omml_blocks))
    for i, omml in enumerate(omml_blocks, start=1):
        mml = omml_to_mathml(omml)
        result.items.append((i, mml, 'mathml' if mml is not None else None))
        if mml is not None:
            result.success += 1
        else:
            result.failed += 1
            result.failed_indices.append(i)
            result.failed_omml.append((i, omml))
    return result


def convert_all_to_latex(omml_blocks: List[str]) -> ConvertResult:
    """将多段 OMML 转为 LaTeX，返回带统计的结果。"""
    result = ConvertResult(total=len(omml_blocks))
    for i, omml in enumerate(omml_blocks, start=1):
        tex = omml_to_latex(omml)
        result.items.append((i, tex, 'latex' if tex is not None else None))
        if tex is not None:
            result.success += 1
        else:
            result.failed += 1
            result.failed_indices.append(i)
            result.failed_omml.append((i, omml))
    return result


def convert_all_with_fallback(omml_blocks: List[str]) -> ConvertResult:
    """
    将多段 OMML 转换，带 fallback 机制：
    1. 先尝试 MathML
    2. MathML 失败则尝试 LaTeX
    3. 都失败则保存原始 OMML
    返回带统计的结果，items 格式为 (index, mathml_or_latex, format_type)
    format_type 为 'mathml'、'latex' 或 None（失败）
    """
    result = ConvertResult(total=len(omml_blocks))
    for i, omml in enumerate(omml_blocks, start=1):
        mml = omml_to_mathml(omml)
        if mml is not None:
            result.items.append((i, mml, 'mathml'))
            result.success += 1
            continue
        
        tex = omml_to_latex(omml)
        if tex is not None:
            result.items.append((i, tex, 'latex'))
            result.success += 1
            continue
        
        result.items.append((i, None, None))
        result.failed += 1
        result.failed_indices.append(i)
        result.failed_omml.append((i, omml))
    
    return result


def save_failed_omml(failed_omml: List[Tuple[int, str]], output_dir: str, base_name: str = "failed") -> None:
    """
    将失败的 OMML 保存到 failed 目录。
    :param failed_omml: [(index, omml_xml), ...]
    :param output_dir: 输出目录
    :param base_name: 基础文件名
    """
    if not failed_omml:
        return
    
    out_dir = Path(output_dir) / "failed"
    out_dir.mkdir(parents=True, exist_ok=True)
    
    for idx, omml in failed_omml:
        file_path = out_dir / f"{base_name}_eq{idx}.xml"
        content = (
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<!-- 转换失败的 OMML 公式，可用于后续手动处理 -->\n'
            f"{omml}\n"
        )
        file_path.write_text(content, encoding="utf-8")
        logger.info("已保存失败公式: %s", file_path)
    
    logger.info("共保存 %d 个失败公式到 %s", len(failed_omml), out_dir)
