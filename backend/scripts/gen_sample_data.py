#!/usr/bin/env python3
# scripts/gen_sample_data.py
"""
生成 data/ 目录下的初始化数据：
  - 高等数学公式示例.docx  （含 5 个真实数学公式：分数、求和、积分、极限、矩阵）
  - 使用说明.txt
"""
import zipfile
from pathlib import Path

# 包含 5 个真实数学公式的 document.xml
# 1. 分数: 1/2
# 2. 求和: sum_{i=1}^{n} i
# 3. 积分: integral_0^1 x^2 dx
# 4. 极限: lim_{x->0} sin(x)/x
# 5. 二次方程求根公式: x = (-b ± sqrt(b^2 - 4ac)) / 2a
DOCUMENT_XML_ZH = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
<w:body>

<w:p><w:r><w:t>高等数学公式转换示例文档</w:t></w:r></w:p>
<w:p><w:r><w:t>本文档包含五个常见数学公式，用于验证 Word 公式提取与转换功能。</w:t></w:r></w:p>

<w:p><w:r><w:t>公式一：分数（二分之一）</w:t></w:r></w:p>
<w:p>
  <m:oMathPara>
    <m:oMath>
      <m:f>
        <m:fPr><m:ctrlPr/></m:fPr>
        <m:num><m:r><m:t>1</m:t></m:r></m:num>
        <m:den><m:r><m:t>2</m:t></m:r></m:den>
      </m:f>
    </m:oMath>
  </m:oMathPara>
</w:p>

<w:p><w:r><w:t>公式二：求和公式</w:t></w:r></w:p>
<w:p>
  <m:oMathPara>
    <m:oMath>
      <m:nary>
        <m:naryPr><m:chr m:val="&#x2211;"/><m:limLoc m:val="undOvr"/></m:naryPr>
        <m:sub><m:r><m:t>i=1</m:t></m:r></m:sub>
        <m:sup><m:r><m:t>n</m:t></m:r></m:sup>
        <m:e><m:r><m:t>i</m:t></m:r></m:e>
      </m:nary>
    </m:oMath>
  </m:oMathPara>
</w:p>

<w:p><w:r><w:t>公式三：定积分</w:t></w:r></w:p>
<w:p>
  <m:oMathPara>
    <m:oMath>
      <m:nary>
        <m:naryPr><m:chr m:val="&#x222B;"/><m:limLoc m:val="subSup"/></m:naryPr>
        <m:sub><m:r><m:t>0</m:t></m:r></m:sub>
        <m:sup><m:r><m:t>1</m:t></m:r></m:sup>
        <m:e>
          <m:sSup>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
            <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
          </m:sSup>
          <m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>dx</m:t></m:r>
        </m:e>
      </m:nary>
    </m:oMath>
  </m:oMathPara>
</w:p>

<w:p><w:r><w:t>公式四：极限</w:t></w:r></w:p>
<w:p>
  <m:oMathPara>
    <m:oMath>
      <m:func>
        <m:funcPr><m:ctrlPr/></m:funcPr>
        <m:fName>
          <m:limLow>
            <m:e><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>lim</m:t></m:r></m:e>
            <m:lim><m:r><m:t>x&#x2192;0</m:t></m:r></m:lim>
          </m:limLow>
        </m:fName>
        <m:e>
          <m:f>
            <m:fPr><m:ctrlPr/></m:fPr>
            <m:num><m:func><m:funcPr><m:ctrlPr/></m:funcPr><m:fName><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>sin</m:t></m:r></m:fName><m:e><m:r><m:t>x</m:t></m:r></m:e></m:func></m:num>
            <m:den><m:r><m:t>x</m:t></m:r></m:den>
          </m:f>
        </m:e>
      </m:func>
    </m:oMath>
  </m:oMathPara>
</w:p>

<w:p><w:r><w:t>公式五：一元二次方程求根公式</w:t></w:r></w:p>
<w:p>
  <m:oMathPara>
    <m:oMath>
      <m:r><m:t>x=</m:t></m:r>
      <m:f>
        <m:fPr><m:ctrlPr/></m:fPr>
        <m:num>
          <m:r><m:t>-b&#x00B1;</m:t></m:r>
          <m:rad>
            <m:radPr><m:degHide m:val="1"/></m:radPr>
            <m:deg/>
            <m:e>
              <m:sSup>
                <m:e><m:r><m:t>b</m:t></m:r></m:e>
                <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
              </m:sSup>
              <m:r><m:t>-4ac</m:t></m:r>
            </m:e>
          </m:rad>
        </m:num>
        <m:den><m:r><m:t>2a</m:t></m:r></m:den>
      </m:f>
    </m:oMath>
  </m:oMathPara>
</w:p>

<w:p><w:r><w:t>以上五个公式均为 Word 内置公式编辑器生成，可被本工具提取并转换为 MathML 和 LaTeX 格式。</w:t></w:r></w:p>

</w:body>
</w:document>"""

DATA_README_ZH = """========================================
  Word 公式转 MathType — 数据目录说明
========================================

本目录用于存放待转换的 .docx 文件及转换后的结果文件。

初始化数据
----------
  高等数学公式示例.docx
    包含五个真实数学公式：分数、求和、积分、极限、求根公式。
    仅支持 Word 内置公式（OMML），不支持 MathType OLE 对象。

转换结果说明
-----------
  *_formulas.xml   MathML 格式，可在 MathType 中导入
  *_formulas.tex   LaTeX 格式，可在 MathType 中粘贴
  *_formulas.html  浏览器预览，可直接打开查看公式渲染效果

使用自己的文档
-----------
  将 .docx 文件复制到本目录，然后执行：
  docker-compose exec converter python run_convert.py /data/你的文档.docx -o /data --latex --html
========================================
"""


def main():
    # 在 Docker 中输出到 /data，本地开发时输出到项目根目录的 data/
    data_dir = Path("/data") if Path("/data").is_mount() else Path(__file__).resolve().parent.parent.parent / "data"
    data_dir.mkdir(parents=True, exist_ok=True)

    docx_path = data_dir / "高等数学公式示例.docx"
    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", DOCUMENT_XML_ZH.encode("utf-8"))
    print(f"已生成: {docx_path}")

    readme_path = data_dir / "使用说明.txt"
    readme_path.write_text(DATA_README_ZH, encoding="utf-8")
    print(f"已生成: {readme_path}")


if __name__ == "__main__":
    main()
