# word_formula_to_mathtype/convert.py
"""
将 OMML 转为 MathType 可用的 MathML 或 LaTeX。
依赖 docx-equation 的 OMML2MML.XSL 与 mml2tex。
"""
import logging
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

from docx_equation.omml import omml2mml, omml2tex

logger = logging.getLogger(__name__)


@dataclass
class ConvertResult:
    """批量转换的结果与统计。"""
    items: List[Tuple[int, Optional[str]]] = field(default_factory=list)
    total: int = 0
    success: int = 0
    failed: int = 0
    failed_indices: List[int] = field(default_factory=list)

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
        result.items.append((i, mml))
        if mml is not None:
            result.success += 1
        else:
            result.failed += 1
            result.failed_indices.append(i)
    return result


def convert_all_to_latex(omml_blocks: List[str]) -> ConvertResult:
    """将多段 OMML 转为 LaTeX，返回带统计的结果。"""
    result = ConvertResult(total=len(omml_blocks))
    for i, omml in enumerate(omml_blocks, start=1):
        tex = omml_to_latex(omml)
        result.items.append((i, tex))
        if tex is not None:
            result.success += 1
        else:
            result.failed += 1
            result.failed_indices.append(i)
    return result
