# word_formula_to_mathtype/extract.py
"""
从 .docx 中提取 OMML（Office Math Markup）公式。
.docx 本质为 ZIP，主内容在 word/document.xml 中，公式为 m:oMath 或 m:oMathPara 节点。
"""
import logging
import re
import zipfile
from pathlib import Path
from typing import List, Tuple

logger = logging.getLogger(__name__)


# 匹配 <m:oMath>...</m:oMath>（需考虑嵌套标签）
_OMML_OPEN = re.compile(r"<m:oMath[^>]*>", re.IGNORECASE)
_OMML_CLOSE = re.compile(r"</m:oMath>", re.IGNORECASE)
# 匹配 <m:oMathPara>...</m:oMathPara>，内容为 oMath
_OMATH_PARA_OPEN = re.compile(r"<m:oMathPara[^>]*>", re.IGNORECASE)
_OMATH_PARA_CLOSE = re.compile(r"</m:oMathPara>", re.IGNORECASE)


def _extract_omml_blocks(content: str) -> List[str]:
    """从 document.xml 的字符串中提取所有 OMML 块（每个块为完整 XML 字符串）。"""
    blocks: List[str] = []
    # 先处理 oMathPara：内部包含 oMath，整段视为一个公式
    rest = content
    while True:
        m_para = _OMATH_PARA_OPEN.search(rest)
        if not m_para:
            break
        start = m_para.end()
        depth = 1
        i = start
        while i < len(rest) and depth > 0:
            next_open = _OMATH_PARA_OPEN.search(rest, i)
            next_close = _OMATH_PARA_CLOSE.search(rest, i)
            if next_close and (not next_open or next_close.start() < next_open.start()):
                depth -= 1
                if depth == 0:
                    blocks.append(rest[m_para.start() : next_close.end()])
                    rest = rest[next_close.end() :]
                    break
                i = next_close.end()
            elif next_open:
                depth += 1
                i = next_open.end()
            else:
                break
        else:
            rest = rest[m_para.end() :]
    # 再处理独立的 m:oMath（不在 oMathPara 内的）
    rest = content
    pos = 0
    while True:
        m_open = _OMML_OPEN.search(rest, pos)
        if not m_open:
            break
        start = m_open.end()
        depth = 1
        i = start
        while i < len(rest) and depth > 0:
            next_close = _OMML_CLOSE.search(rest, i)
            next_open = _OMML_OPEN.search(rest, i)
            if next_close and (not next_open or next_close.start() < next_open.start()):
                depth -= 1
                if depth == 0:
                    candidate = rest[m_open.start() : next_close.end()]
                    # 若已在 oMathPara 中提取过，则跳过（简单判断：是否被我们已取块包含）
                    if not any(candidate in b for b in blocks):
                        blocks.append(candidate)
                    pos = next_close.end()
                    break
                i = next_close.end()
            elif next_open:
                depth += 1
                i = next_open.end()
            else:
                pos = m_open.end()
                break
        else:
            pos = m_open.end()
    return blocks


def extract_omml_from_docx(docx_path: str) -> List[str]:
    """
    从 .docx 文件中提取所有 OMML 公式块。
    :param docx_path: Word 文档路径
    :return: OMML XML 字符串列表
    """
    path = Path(docx_path)
    if not path.exists():
        logger.error("文件不存在: %s", docx_path)
        raise FileNotFoundError(f"文件不存在: {docx_path}")
    if path.suffix.lower() != ".docx":
        logger.error("仅支持 .docx 格式: %s", path.suffix)
        raise ValueError("仅支持 .docx 格式")

    logger.debug("正在读取 docx: %s", path)
    with zipfile.ZipFile(path, "r") as z:
        try:
            data = z.read("word/document.xml")
        except KeyError:
            logger.error("无效的 docx，缺少 word/document.xml: %s", path)
            raise ValueError("无效的 docx：缺少 word/document.xml")
        text = data.decode("utf-8", errors="replace")

    blocks = _extract_omml_blocks(text)
    logger.info("从 %s 提取到 %d 个 OMML 公式块", path.name, len(blocks))
    return blocks


def extract_omml_from_docx_with_positions(docx_path: str) -> List[Tuple[int, str]]:
    """
    从 .docx 中提取 OMML，并返回在文档中的大致顺序索引（用于输出时编号）。
    :return: [(index, omml_xml), ...]
    """
    blocks = extract_omml_from_docx(docx_path)
    return list(enumerate(blocks, start=1))
