#!/bin/bash
# entrypoint.sh — 容器启动时自动初始化示例数据并执行一次演示转换
set -e

echo "========================================"
echo "  Word 公式转 MathType 工具 — 初始化中"
echo "========================================"

# 1. 生成示例数据（如果尚未存在）
if [ ! -f /data/高等数学公式示例.docx ]; then
    echo "[初始化] 生成示例文档..."
    python scripts/gen_sample_data.py
else
    echo "[初始化] 示例文档已存在，跳过生成"
fi

# 2. 自动执行一次转换，生成 MathML + LaTeX + HTML
echo ""
echo "[转换] 正在转换示例文档..."
python run_convert.py /data/高等数学公式示例.docx -o /data --latex --html

echo ""
echo "========================================"
echo "  初始化完成"
echo "========================================"
echo ""
echo "生成的文件位于 data/ 目录："
ls -la /data/*.xml /data/*.tex /data/*.html 2>/dev/null || true
echo ""
echo "如需转换其他文档，请将 .docx 放入 data/ 目录后执行："
echo "  docker-compose exec converter python run_convert.py /data/你的文档.docx -o /data --latex --html"
echo ""

# 3. 保持容器运行
exec tail -f /dev/null
