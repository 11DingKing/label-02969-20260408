# tests/conftest.py
"""Pytest fixtures: 最小 docx 文件（无公式 / 含公式）、临时目录。"""
import zipfile
from pathlib import Path

import pytest


# 最小有效 document.xml：无公式
DOCUMENT_XML_NO_MATH = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body>
</w:document>"""

# 含一个简单分数公式的 document.xml
DOCUMENT_XML_ONE_MATH = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
<w:body>
<w:p><w:r><w:t>Formula: </w:t></w:r><m:oMath><m:f><m:fPr/><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath></w:p>
</w:body>
</w:document>"""

# 含两个公式：一个独立 oMath，一个 oMathPara
DOCUMENT_XML_TWO_MATH = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
<w:body>
<w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>
<w:p><m:oMathPara><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></m:oMathPara></w:p>
</w:body>
</w:document>"""

# 单段 OMML 字符串，用于 convert 测试
SAMPLE_OMML_FRACTION = """<m:oMath><m:f><m:fPr/><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>"""


def _write_docx(path: Path, document_xml: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", document_xml.encode("utf-8"))


@pytest.fixture
def docx_no_math(tmp_path):
    """无公式的 .docx 路径。"""
    p = tmp_path / "no_math.docx"
    _write_docx(p, DOCUMENT_XML_NO_MATH)
    return str(p)


@pytest.fixture
def docx_one_math(tmp_path):
    """含一个公式的 .docx 路径。"""
    p = tmp_path / "one_math.docx"
    _write_docx(p, DOCUMENT_XML_ONE_MATH)
    return str(p)


@pytest.fixture
def docx_two_math(tmp_path):
    """含两个公式的 .docx 路径。"""
    p = tmp_path / "two_math.docx"
    _write_docx(p, DOCUMENT_XML_TWO_MATH)
    return str(p)


@pytest.fixture
def sample_omml():
    """单段 OMML（分数 1/2）。"""
    return SAMPLE_OMML_FRACTION
