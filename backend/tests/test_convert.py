# tests/test_convert.py
"""转换模块测试：OMML -> MathML / LaTeX，含 ConvertResult 统计。"""
import pytest

from word_formula_to_mathtype.convert import (
    omml_to_mathml,
    omml_to_latex,
    convert_all_to_mathml,
    convert_all_to_latex,
    ConvertResult,
)


class TestOmmlToMathml:
    """omml_to_mathml 单公式转换。"""

    def test_valid_fraction(self, sample_omml):
        result = omml_to_mathml(sample_omml)
        assert result is not None
        assert "math" in result.lower() or "mfrac" in result
        assert "1" in result and "2" in result

    def test_empty_string(self):
        assert omml_to_mathml("") is None

    def test_invalid_xml_returns_none(self):
        assert omml_to_mathml("<m:oMath>broken") is None


class TestOmmlToLatex:
    """omml_to_latex 单公式转换。"""

    def test_valid_fraction(self, sample_omml):
        result = omml_to_latex(sample_omml)
        assert result is not None
        assert "frac" in result or "1" in result

    def test_empty_string(self):
        assert omml_to_latex("") is None


class TestConvertAll:
    """批量转换与统计。"""

    def test_convert_all_to_mathml(self, sample_omml):
        blocks = [sample_omml, sample_omml]
        result = convert_all_to_mathml(blocks)
        assert result.total == 2
        assert result.success == 2
        assert result.failed == 0
        assert len(result.items) == 2
        assert result.items[0][0] == 1 and result.items[1][0] == 2
        assert result.items[0][1] is not None and result.items[1][1] is not None

    def test_convert_all_to_latex(self, sample_omml):
        blocks = [sample_omml]
        result = convert_all_to_latex(blocks)
        assert result.total == 1
        assert result.success == 1
        assert result.failed == 0
        assert result.items[0][0] == 1
        assert result.items[0][1] is not None

    def test_empty_blocks(self):
        r1 = convert_all_to_mathml([])
        r2 = convert_all_to_latex([])
        assert r1.items == [] and r1.total == 0
        assert r2.items == [] and r2.total == 0

    def test_mixed_valid_invalid(self, sample_omml):
        blocks = [sample_omml, "<m:oMath>broken", sample_omml]
        result = convert_all_to_mathml(blocks)
        assert result.total == 3
        assert result.success == 2
        assert result.failed == 1
        assert result.failed_indices == [2]
        assert result.items[0][1] is not None
        assert result.items[1][1] is None
        assert result.items[2][1] is not None


class TestConvertResult:
    """ConvertResult 统计与摘要。"""

    def test_summary_all_success(self):
        r = ConvertResult(total=3, success=3, failed=0)
        assert "全部成功" in r.summary
        assert "3/3" in r.summary

    def test_summary_with_failures(self):
        r = ConvertResult(total=5, success=3, failed=2, failed_indices=[2, 4])
        assert "成功 3/5" in r.summary
        assert "失败 2" in r.summary
        assert "2, 4" in r.summary

    def test_summary_empty(self):
        r = ConvertResult(total=0, success=0, failed=0)
        assert "无公式" in r.summary

    def test_success_rate(self):
        r = ConvertResult(total=10, success=8, failed=2)
        assert r.success_rate == "80.0%"
