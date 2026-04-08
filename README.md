# Word 公式转 MathType 工具

将 Word 文档（.docx）中的内置公式（OMML）批量转换为 MathType 可用的 MathML 与 LaTeX，便于在 MathType 中粘贴或导入。

## How to Run

### 一键启动

```bash
docker-compose up --build -d
```

启动后容器会自动完成以下操作：
1. 生成示例文档（高等数学公式示例.docx，含分数、求和、积分、极限、求根公式共 5 个公式）
2. 自动执行一次转换，在 `data/` 目录生成 MathML、LaTeX、HTML 三种格式的结果文件

查看初始化日志：

```bash
docker-compose logs converter
```

预期输出：

```
Word 公式转 MathType 工具 — 初始化中
[初始化] 生成示例文档...
[转换] 正在转换示例文档...
INFO: 从 高等数学公式示例.docx 提取到 5 个 OMML 公式块
INFO: 共提取 5 个公式
INFO: 已写入 MathML: /data/高等数学公式示例_formulas.xml
INFO: 已写入 HTML 预览: /data/高等数学公式示例_formulas.html
INFO: 已写入 LaTeX: /data/高等数学公式示例_formulas.tex
初始化完成
```

### 转换自己的文档

将 `.docx` 文件放入项目根目录的 `data/` 文件夹，然后执行：

```bash
docker-compose exec converter python run_convert.py /data/你的文档.docx -o /data --latex --html
```

### 非 Docker 场景（本地 Python）

```bash
cd backend
pip install -r requirements.txt
python run_convert.py your.docx -o ./output --latex --html
```

### 停止服务

```bash
docker-compose down
```

## Services

| 服务 | 说明 | 端口 |
|------|------|------|
| `converter` | Word 公式转换后端（CLI 工具，容器常驻） | 无（通过 `docker-compose exec` 调用） |

本项目为命令行工具，无 HTTP 服务端口。转换通过 `docker-compose exec` 按需执行。

## 测试账号

不涉及登录与账号，无需测试账号。

## 题目内容

> 我的 word 文章中公式很多，帮我写一个 python 脚本，用于将 word 中的公式转为 math type 公式。

实现说明：
- 输入：Word `.docx` 文件（仅支持 Word 内置公式 OMML，不支持以 OLE 嵌入的 MathType 对象）
- 输出：MathML（MathType 可导入）、LaTeX（可选）、HTML 预览（可选）
- 访问地址：本地 CLI 工具，无 Web 访问地址；输出为本地文件

---

## 质检人员测试指南

### 环境要求

- 已安装 Docker 与 Docker Compose
- 终端编码为 UTF-8

### 第一步：一键启动

```bash
docker-compose up --build -d
```

等待约 30 秒，执行 `docker ps` 确认容器 `word-formula-converter` 状态为 `Up`。

### 第二步：查看初始化日志

```bash
docker-compose logs converter
```

验收要点：
- 日志中显示"提取到 5 个 OMML 公式块"
- 日志中显示"已写入 MathML"、"已写入 HTML 预览"、"已写入 LaTeX"
- 中文无乱码

### 第三步：检查生成的结果文件

启动完成后 `data/` 目录下应自动生成以下文件：

| 文件 | 验收要点 |
|------|----------|
| `高等数学公式示例.docx` | 示例文档，含 5 个数学公式 |
| `高等数学公式示例_formulas.xml` | 包含 `count="5"` 及 5 个 `<formula>` 节点 |
| `高等数学公式示例_formulas.tex` | 包含 `\frac{1}{2}`、`\sum`、`\int`、`\lim`、`\sqrt` 等 LaTeX |
| `高等数学公式示例_formulas.html` | 浏览器打开可见 5 个公式的渲染效果，标题为中文无乱码 |

快速验证命令：

```bash
# 检查 XML 中公式数量
grep -c "formula index" data/高等数学公式示例_formulas.xml
# 预期输出: 5

# 检查 LaTeX 中的关键公式
cat data/高等数学公式示例_formulas.tex
# 预期包含: \frac{1}{2}, \sum, \int, \lim, \sqrt
```

### 第四步：手动执行转换（验证 exec 可用）

```bash
docker-compose exec converter python run_convert.py /data/高等数学公式示例.docx -o /data --latex --html
```

预期：终端输出"共提取 5 个公式"，无报错。

### 第五步：测试自定义文档（可选）

将任意含 Word 内置公式的 `.docx` 文件复制到 `data/` 目录：

```bash
cp 你的文档.docx data/
docker-compose exec converter python run_convert.py /data/你的文档.docx -o /data --latex --html
```

检查 `data/` 下生成的对应 `_formulas.xml`、`_formulas.tex`、`_formulas.html` 文件。

### 第六步：停止环境

```bash
docker-compose down
```

### 验收标准

- [ ] `docker-compose up --build -d` 一次性完成构建、启动、初始化，无报错
- [ ] 容器启动后自动生成示例文档并完成转换
- [ ] `data/` 目录下存在 `.xml`、`.tex`、`.html` 三个结果文件
- [ ] XML 文件包含 5 个公式的 MathML
- [ ] LaTeX 文件包含分数、求和、积分、极限、求根公式
- [ ] HTML 文件浏览器打开可见公式渲染，中文无乱码
- [ ] `docker-compose exec` 可手动执行转换

---

## 项目结构

```
.
├── backend/                    # 后端 — Python CLI 转换工具
│   ├── word_formula_to_mathtype/
│   ├── tests/
│   ├── scripts/
│   ├── run_convert.py
│   ├── requirements.txt
│   ├── Dockerfile
│   └── README.md
├── data/                       # 数据目录（挂载到容器 /data）
├── docker-compose.yml
├── .gitignore
└── README.md
```

后端详细说明见 [backend/README.md](backend/README.md)。

## 内置示例公式

| 编号 | 公式 | LaTeX |
|------|------|-------|
| 1 | 分数 | `\frac{1}{2}` |
| 2 | 求和 | `\sum_{i=1}^{n} i` |
| 3 | 定积分 | `\int_0^1 x^2 dx` |
| 4 | 极限 | `\lim_{x \to 0} \frac{\sin x}{x}` |
| 5 | 求根公式 | `x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}` |
