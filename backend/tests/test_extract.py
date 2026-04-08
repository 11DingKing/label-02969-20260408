# tests/test_extract.py
"""提取模块测试：_extract_omml_blocks、extract_omml_from_docx、边界与异常。"""
import pytest

from word_formula_to_mathtype.extract import (
    _extract_omml_blocks,
    extract_omml_from_docx,
    extract_omml_from_docx_with_positions,
)


class TestExtractOmmlBlocks:
    """_extract_omml_blocks 单元测试。"""

    def test_empty_content(self):
        assert _extract_omml_blocks("") == []
        assert _extract_omml_blocks("<w:p></w:p>") == []

    def test_no_math(self):
        xml = """<w:document><w:body><w:p><w:t>text</w:t></w:p></w:body></w:document>"""
        assert _extract_omml_blocks(xml) == []

    def test_single_omath(self):
        xml = """<w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>"""
        blocks = _extract_omml_blocks(xml)
        assert len(blocks) == 1
        assert "<m:oMath>" in blocks[0] and "</m:oMath>" in blocks[0]
        assert "x" in blocks[0]

    def test_single_fraction_omath(self):
        xml = """<w:p><m:oMath><m:f><m:fPr/><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath></w:p>"""
        blocks = _extract_omml_blocks(xml)
        assert len(blocks) == 1
        assert "m:f" in blocks[0] and "1" in blocks[0] and "2" in blocks[0]

    def test_omath_para(self):
        xml = """<w:p><m:oMathPara><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></m:oMathPara></w:p>"""
        blocks = _extract_omml_blocks(xml)
        assert len(blocks) == 1
        assert "m:oMathPara" in blocks[0] and "y" in blocks[0]

    def test_two_formulas(self):
        xml = """<w:body>
<w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>
<w:p><m:oMathPara><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></m:oMathPara></w:p>
</w:body>"""
        blocks = _extract_omml_blocks(xml)
        assert len(blocks) == 2
        texts = [blocks[0], blocks[1]]
        assert any("x" in b for b in texts) and any("y" in b for b in texts)


class TestExtractFromDocx:
    """extract_omml_from_docx 集成测试。"""

    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError, match="文件不存在"):
            extract_omml_from_docx("/nonexistent/path.docx")

    def test_wrong_extension(self, tmp_path):
        f = tmp_path / "file.doc"
        f.write_text("x")
        with pytest.raises(ValueError, match="仅支持 .docx"):
            extract_omml_from_docx(str(f))

    def test_invalid_docx_no_document_xml(self, tmp_path):
        import zipfile
        p = tmp_path / "bad.docx"
        with zipfile.ZipFile(p, "w") as z:
            z.writestr("other.xml", b"<x/>")
        with pytest.raises(ValueError, match="缺少 word/document.xml"):
            extract_omml_from_docx(str(p))

    def test_docx_no_math(self, docx_no_math):
        blocks = extract_omml_from_docx(docx_no_math)
        assert blocks == []

    def test_docx_one_math(self, docx_one_math):
        blocks = extract_omml_from_docx(docx_one_math)
        assert len(blocks) == 1
        assert "m:oMath" in blocks[0] and "1" in blocks[0] and "2" in blocks[0]

    def test_docx_two_math(self, docx_two_math):
        blocks = extract_omml_from_docx(docx_two_math)
        assert len(blocks) == 2


class TestExtractWithPositions:
    """extract_omml_from_docx_with_positions 测试。"""

    def test_returns_indexed(self, docx_one_math):
        result = extract_omml_from_docx_with_positions(docx_one_math)
        assert len(result) == 1
        assert result[0][0] == 1
        assert "m:oMath" in result[0][1]

    def test_empty_returns_empty(self, docx_no_math):
        result = extract_omml_from_docx_with_positions(docx_no_math)
        assert result == []
