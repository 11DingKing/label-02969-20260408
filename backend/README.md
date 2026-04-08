# Backend — Word 公式转 MathType 工具

将 Word 文档（.docx）中的内置公式（OMML）批量转换为 MathType 可用的 MathML 与 LaTeX。

## 命令行参数

| 参数 | 说明 |
|------|------|
| `input_docx` | 输入的 .docx 文件路径（必填） |
| `-o, --output-dir` | 输出目录，默认与输入文件同目录 |
| `--mathml` | 输出 MathML（默认开启） |
| `--no-mathml` | 不输出 MathML |
| `--latex` | 同时输出 LaTeX 文件 |
| `--html` | 额外输出 HTML 预览（含公式渲染与 MathML 源码） |

## 本地开发

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# 运行转换
python run_convert.py ../data/高等数学公式示例.docx -o ../data --latex --html

# 运行测试
pytest tests/ -v
```

## Docker 使用

容器启动时会自动执行初始化脚本（`scripts/entrypoint.sh`）：
1. 检查并生成示例文档（5 个数学公式）
2. 自动执行一次转换，生成 MathML + LaTeX + HTML
3. 保持容器运行，等待后续 `exec` 调用

## 项目结构

```
backend/
├── word_formula_to_mathtype/   # 核心包
│   ├── extract.py              # 从 docx 提取 OMML
│   ├── convert.py              # OMML → MathML / LaTeX
│   └── output.py               # 写入 XML / LaTeX / HTML
├── tests/                      # 单元测试
├── scripts/
│   ├── gen_sample_data.py      # 生成示例数据
│   └── entrypoint.sh           # Docker 容器入口脚本
├── run_convert.py              # CLI 入口
├── requirements.txt
└── Dockerfile
```

## 限制

- 仅处理 Word 内置编辑器写入的公式（OMML）
- 通过"插入 → 对象 → MathType"嵌入的 OLE 公式无法提取
