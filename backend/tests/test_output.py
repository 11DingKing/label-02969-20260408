# tests/test_output.py
"""输出模块测试：write_mathml_file、write_latex_file、归一化与转义。"""
from pathlib import Path

import pytest

from word_formula_to_mathtype.output import (
    write_mathml_file,
    write_latex_file,
    _normalize_mml,
    _escape,
    _wrap_mathml,
)


# 模拟 MathML 片段（与 docx-equation 输出风格一致）
SAMPLE_MATHML = "<mml:mfrac><mml:mn>1</mml:mn><mml:mn>2</mml:mn></mml:mfrac>"
SAMPLE_MATHML_WITH_DECL = '<?xml version="1.0"?>\n<mml:math xmlns:mml="http://www.w3.org/1998/Math/MathML">' + SAMPLE_MATHML + "</mml:math>"


class TestNormalizeMml:
    """_normalize_mml 行为。"""

    def test_strips_whitespace(self):
        assert _normalize_mml("  \n  <math/>  ") == "<math/>"

    def test_removes_xml_declaration(self):
        out = _normalize_mml(SAMPLE_MATHML_WITH_DECL)
        assert out.startswith("<") and "<?xml" not in out

    def test_passthrough_no_declaration(self):
        assert _normalize_mml(SAMPLE_MATHML) == SAMPLE_MATHML


class TestEscape:
    """_escape HTML 转义。"""

    def test_escapes_amp_lt_gt(self):
        assert "&amp;" in _escape("a & b")
        assert "&lt;" in _escape("<tag>")
        assert "&gt;" in _escape(">")


class TestWrapMathml:
    """_wrap_mathml 包装。"""

    def test_wraps_with_xml_and_math(self):
        out = _wrap_mathml("<mfrac/>")
        assert out.startswith("<?xml")
        assert "<math " in out or "<math>" in out
        assert "<mfrac/>" in out


class TestWriteMathmlFile:
    """write_mathml_file 写入。"""

    def test_empty_results_xml(self, tmp_path):
        path = tmp_path / "out.xml"
        write_mathml_file([], str(path), single_document=True)
        assert path.exists()
        content = path.read_text(encoding="utf-8")
        assert "count=\"0\"" in content or "formulas" in content

    def test_single_result_xml(self, tmp_path):
        path = tmp_path / "out.xml"
        write_mathml_file([(1, SAMPLE_MATHML)], str(path), single_document=True)
        assert path.exists()
        content = path.read_text(encoding="utf-8")
        assert "formula" in content and "1" in content

    def test_single_result_html(self, tmp_path):
        path = tmp_path / "out.html"
        write_mathml_file([(1, SAMPLE_MATHML)], str(path), single_document=True)
        assert path.exists()
        content = path.read_text(encoding="utf-8")
        assert "<!DOCTYPE html>" in content and "公式" in content

    def test_skip_none_in_single_doc(self, tmp_path):
        path = tmp_path / "out.xml"
        write_mathml_file([(1, None), (2, SAMPLE_MATHML)], str(path), single_document=True)
        content = path.read_text(encoding="utf-8")
        assert "count=\"1\"" in content

    def test_split_files(self, tmp_path):
        path = tmp_path / "out.xml"
        write_mathml_file([(1, SAMPLE_MATHML), (2, SAMPLE_MATHML)], str(path), single_document=False)
        eq1 = tmp_path / "out_eq1.xml"
        eq2 = tmp_path / "out_eq2.xml"
        assert eq1.exists() and eq2.exists()

    def test_creates_parent_dir(self, tmp_path):
        path = tmp_path / "sub" / "dir" / "out.xml"
        write_mathml_file([(1, SAMPLE_MATHML)], str(path), single_document=True)
        assert path.exists()


class TestWriteLatexFile:
    """write_latex_file 写入。"""

    def test_writes_index_and_latex(self, tmp_path):
        path = tmp_path / "out.tex"
        write_latex_file([(1, r"$\frac{1}{2}$"), (2, None)], str(path))
        assert path.exists()
        content = path.read_text(encoding="utf-8")
        assert "公式 1" in content and "frac" in content
        assert "公式 2" in content and "转换失败" in content

    def test_empty_results(self, tmp_path):
        path = tmp_path / "out.tex"
        write_latex_file([], str(path))
        assert path.exists()
