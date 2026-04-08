# tests/test_cli.py
"""命令行入口测试：run_convert.py 参数与退出码。"""
import subprocess
import sys
from pathlib import Path


def run_cli(args, cwd=None):
    """执行 run_convert.py，返回 (returncode, stdout, stderr)。"""
    if cwd is None:
        cwd = Path(__file__).resolve().parent.parent
    cmd = [sys.executable, "run_convert.py"] + args
    r = subprocess.run(
        cmd,
        cwd=cwd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return r.returncode, r.stdout, r.stderr


class TestCliHelp:
    """--help 与无参数。"""

    def test_help_exits_zero(self):
        code, out, err = run_cli(["--help"])
        assert code == 0
        assert "input_docx" in out or "用法" in out or "docx" in out

    def test_no_args_exits_nonzero(self):
        code, _, err = run_cli([])
        assert code != 0
        assert "required" in err or "input_docx" in err or "error" in err.lower()


class TestCliValidation:
    """文件存在性、扩展名校验。"""

    def test_nonexistent_file_exits_one(self):
        code, _, err = run_cli(["/nonexistent/file.docx"])
        assert code == 1
        assert "不存在" in err or "not found" in err.lower() or "error" in err.lower()

    def test_wrong_extension_exits_one(self, tmp_path):
        f = tmp_path / "x.doc"
        f.write_text("x")
        code, _, err = run_cli([str(f)])
        assert code == 1
        assert "docx" in err or "格式" in err


class TestCliE2e:
    """端到端：无公式 / 有公式。"""

    def test_docx_no_math_exits_zero_and_writes_empty(self, docx_no_math, tmp_path):
        out_dir = tmp_path / "out"
        code, out, err = run_cli([docx_no_math, "-o", str(out_dir)])
        assert code == 0
        assert out_dir.exists()
        stem = Path(docx_no_math).stem
        xml_file = out_dir / f"{stem}_formulas.xml"
        assert xml_file.exists()
        content = xml_file.read_text(encoding="utf-8")
        assert "count=\"0\"" in content or "formulas" in content

    def test_docx_one_math_exits_zero_and_writes_mathml(self, docx_one_math, tmp_path):
        out_dir = tmp_path / "out"
        code, out, err = run_cli([docx_one_math, "-o", str(out_dir)])
        assert code == 0, (out, err)
        stem = Path(docx_one_math).stem
        xml_file = out_dir / f"{stem}_formulas.xml"
        assert xml_file.exists()
        content = xml_file.read_text(encoding="utf-8")
        assert "formula" in content and ("count=\"1\"" in content or "index=" in content)

    def test_latex_flag_writes_tex(self, docx_one_math, tmp_path):
        out_dir = tmp_path / "out"
        code, _, _ = run_cli([docx_one_math, "-o", str(out_dir), "--latex"])
        assert code == 0
        stem = Path(docx_one_math).stem
        tex_file = out_dir / f"{stem}_formulas.tex"
        assert tex_file.exists()
