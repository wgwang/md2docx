# md2docx

`md2docx` 是一个 Python 工具，用于将 Markdown 文件转换为美观的 Docx 文档。它支持自定义字体、数学公式、表格、图片嵌入等功能，特别针对中文字体支持进行了优化。

## 功能特性

- **中文字体优化**：默认使用 "Microsoft YaHei" (微软雅黑)，解决生成文档中中文字体显示问题。
- **数学公式支持**：支持 LaTeX 格式的行内公式 ($...$) 和块级公式 ($$...$$)。
- **图片嵌入**：支持本地图片和网络图片（自动下载并嵌入）。
- **表格支持**：支持标准 Markdown 表格语法。
- **代码高亮**：支持代码块和行内代码样式。
- **换行控制**：提供选项保留 Markdown 中的硬换行。

## 安装

### 从源码安装

```bash
git clone https://github.com/yourusername/md2docx.git
cd md2docx
pip install .
```

或者使用开发者模式（推荐，便于调试）：

```bash
pip install -e .
```

## 使用方法

### 命令行工具 (CLI)

安装后，您可以直接在终端使用 `md2docx` 命令：

```bash
md2docx <输入文件.md> <输出文件.docx> [选项]
```

**示例：**

1. **基本转换**：
   ```bash
   md2docx input.md output.docx
   ```

2. **指定字体** (例如使用宋体)：
   ```bash
   md2docx input.md output.docx --font "SimSun"
   ```

3. **保留换行符** (将 Markdown 中的单次换行转换为文档中的换行)：
   ```bash
   md2docx input.md output.docx --preserve-breaks
   ```

### Python SDK

您也可以在 Python 代码中作为库使用：

```python
from md2docx import MarkdownToDocx

# 读取 Markdown 内容
with open("input.md", "r", encoding="utf-8") as f:
    md_content = f.read()

# 初始化转换器
# font_name: 指定使用的字体，默认为 "Microsoft YaHei"
# preserve_breaks: 是否保留单次换行，默认为 False
converter = MarkdownToDocx(font_name="Microsoft YaHei", preserve_breaks=True)

# 执行转换
converter.convert(md_content, "output.docx")
```

## 目录结构说明

```text
md2docx/
├── md2docx/            # 源代码目录
│   ├── __init__.py
│   ├── cli.py          # 命令行入口
│   └── converter.py    # 转换核心逻辑
├── examples/           # 示例文件
├── tests/              # 测试用例
├── setup.py            # 安装脚本
├── requirements.txt    # 依赖列表
└── README.md           # 说明文档
```

## 许可证

本项目采用 [MIT 许可证](LICENSE)。