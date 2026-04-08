# word_formula_to_mathtype/extract.py
"""
从 .docx 或 .doc 中提取 OMML（Office Math Markup）公式。
.docx 本质为 ZIP，主内容在 word/document.xml 中，公式为 m:oMath 或 m:oMathPara 节点。
.doc 格式需要先转换为 .docx 格式。
"""
import logging
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import List, Optional, Tuple

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


def _convert_doc_to_docx_libreoffice(doc_path: str, output_dir: str) -> Optional[str]:
    """
    使用 LibreOffice 将 .doc 转换为 .docx。
    :param doc_path: 输入 .doc 文件路径
    :param output_dir: 输出目录
    :return: 转换后的 .docx 文件路径，失败返回 None
    """
    doc_path = Path(doc_path)
    output_dir = Path(output_dir)
    
    try:
        cmd = [
            "libreoffice",
            "--headless",
            "--convert-to", "docx",
            "--outdir", str(output_dir),
            str(doc_path)
        ]
        
        logger.info("使用 LibreOffice 转换 .doc 到 .docx: %s", doc_path)
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120
        )
        
        if result.returncode != 0:
            logger.error("LibreOffice 转换失败: %s", result.stderr)
            return None
        
        output_file = output_dir / f"{doc_path.stem}.docx"
        if output_file.exists():
            logger.info("转换成功: %s", output_file)
            return str(output_file)
        else:
            logger.error("转换后文件不存在: %s", output_file)
            return None
            
    except FileNotFoundError:
        logger.error("未找到 LibreOffice，请确保已安装并添加到 PATH")
        return None
    except subprocess.TimeoutExpired:
        logger.error("LibreOffice 转换超时")
        return None
    except Exception as e:
        logger.error("转换 .doc 到 .docx 失败: %s", e)
        return None


def _convert_doc_to_docx(doc_path: str) -> Optional[str]:
    """
    将 .doc 文件转换为 .docx 格式。
    优先使用 LibreOffice（跨平台）。
    :param doc_path: 输入 .doc 文件路径
    :return: 转换后的 .docx 文件路径（临时文件），失败返回 None
    """
    temp_dir = tempfile.mkdtemp(prefix="word_convert_")
    
    try:
        result = _convert_doc_to_docx_libreoffice(doc_path, temp_dir)
        if result:
            return result
        
        shutil.rmtree(temp_dir, ignore_errors=True)
        return None
        
    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise


def extract_omml_from_doc(doc_path: str) -> List[str]:
    """
    从 .doc 文件中提取所有 OMML 公式块。
    先将 .doc 转换为 .docx，然后提取公式。
    :param doc_path: Word 文档路径
    :return: OMML XML 字符串列表
    """
    path = Path(doc_path)
    if not path.exists():
        logger.error("文件不存在: %s", doc_path)
        raise FileNotFoundError(f"文件不存在: {doc_path}")
    if path.suffix.lower() != ".doc":
        logger.error("仅支持 .doc 格式: %s", path.suffix)
        raise ValueError("仅支持 .doc 格式")
    
    logger.info("正在处理 .doc 文件: %s", path)
    docx_path = _convert_doc_to_docx(doc_path)
    
    if not docx_path:
        logger.error("无法将 .doc 转换为 .docx: %s", doc_path)
        raise RuntimeError(f"无法将 .doc 转换为 .docx: {doc_path}")
    
    try:
        blocks = extract_omml_from_docx(docx_path)
        return blocks
    finally:
        docx_path_obj = Path(docx_path)
        if docx_path_obj.exists():
            temp_dir = docx_path_obj.parent
            shutil.rmtree(temp_dir, ignore_errors=True)


def extract_omml_from_file(file_path: str) -> List[str]:
    """
    从 Word 文件中提取所有 OMML 公式块。
    自动识别 .docx 或 .doc 格式。
    :param file_path: Word 文档路径
    :return: OMML XML 字符串列表
    """
    path = Path(file_path)
    suffix = path.suffix.lower()
    
    if suffix == ".docx":
        return extract_omml_from_docx(file_path)
    elif suffix == ".doc":
        return extract_omml_from_doc(file_path)
    else:
        logger.error("不支持的文件格式: %s", suffix)
        raise ValueError(f"不支持的文件格式: {suffix}")


def extract_omml_from_file_with_positions(file_path: str) -> List[Tuple[int, str]]:
    """
    从 Word 文件中提取 OMML，并返回在文档中的大致顺序索引（用于输出时编号）。
    自动识别 .docx 或 .doc 格式。
    :return: [(index, omml_xml), ...]
    """
    blocks = extract_omml_from_file(file_path)
    return list(enumerate(blocks, start=1))
