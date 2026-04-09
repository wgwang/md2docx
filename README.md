# md2docx

[![Python Tests](https://github.com/yourusername/md2docx/actions/workflows/test.yml/badge.svg)](https://github.com/yourusername/md2docx/actions/workflows/test.yml)
[![PyPI version](https://badge.fury.io/py/md2docx-pro.svg)](https://badge.fury.io/py/md2docx-pro)

`md2docx` 是一个 Python 工具，用于将 Markdown 文件转换为美观的 Docx 文档。它支持自定义字体、数学公式、表格、图片嵌入等功能，特别针对中文字体支持进行了优化。

## 功能特性

- **多风格模版支持**：内置多种专业风格模版（专业、学术、创意），支持使用现有 `.docx` 作为模版继承样式。
- **中文字体优化**：默认使用 "Microsoft YaHei" (微软雅黑) 或 "SimSun" (宋体)，解决生成文档中中文字体显示问题。
- **数学公式支持**：支持 LaTeX 格式的行内公式 ($...$) 和块级公式 ($$...$$)，完美转换为 Word 原生公式。
- **图片嵌入**：支持本地图片和网络图片（自动下载并嵌入）。
- **表格支持**：支持标准 Markdown 表格语法。
- **Agent 友好**：提供 `convert_to_bytes` 接口，支持在内存中生成文档流，适配 AI Agent 和 Web API。
- **Agent Skill**：提供符合标准规范的 Skill 封装，方便 OpenClaw 等 AI 框架直接调用。

## 安装

### 通过 PyPI 安装 (推荐)

```bash
pip install md2docx-pro
```

### 从源码安装

```bash
git clone https://github.com/yourusername/md2docx.git
cd md2docx
pip install .
```

## 使用方法

### 命令行工具 (CLI)

```bash
md2docx <输入文件.md> <输出文件.docx> [选项]
```

**示例：**

1. **基本转换**：
   ```bash
   md2docx input.md output.docx
   ```

2. **使用内置专业模版**：
   ```bash
   md2docx input.md output.docx --template professional_template.docx
   ```

3. **指定字体**：
   ```bash
   md2docx input.md output.docx --font "SimSun"
   ```

### Python SDK

```python
from md2docx import MarkdownToDocx

md_content = "# Hello World"
converter = MarkdownToDocx(template_path="professional_template.docx")

# 1. 转换为本地文件
converter.convert(md_content, "output.docx")

# 2. 转换为内存流 (适用于 AI Agent)
docx_bytes = converter.convert_to_bytes(md_content)
```

## 内置模版

- `professional_template.docx`: 商务专业风格，现代无衬线字体 (Arial + 微软雅黑)。
- `academic_template.docx`: 严谨学术风格，衬线字体 (Times New Roman + 宋体)，1.5 倍行距。
- `creative_template.docx`: 活泼创意风格，橙色/蓝色主题，适合演示文档。

## AI Agent 集成

本项目包含一个 `md2docx-skill` 目录，符合 Gemini CLI 等 AI 框架的 Skill 规范。AI Agent 可以通过该 Skill 快速了解如何调用 `md2docx` 生成文档。

## 许可证

本项目采用 [MIT 许可证](LICENSE)。
