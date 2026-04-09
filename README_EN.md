# md2docx

[![Python Tests](https://github.com/yourusername/md2docx/actions/workflows/test.yml/badge.svg)](https://github.com/yourusername/md2docx/actions/workflows/test.yml)
[![PyPI version](https://badge.fury.io/py/md2docx.svg)](https://badge.fury.io/py/md2docx)

`md2docx` is a Python tool designed to convert Markdown files into elegantly formatted Docx documents. It features support for custom fonts, mathematical equations, tables, and embedded images, with specific optimizations for CJK (Chinese, Japanese, Korean) font support.

## Features

- **Multi-Style Templates**: Built-in professional, academic, and creative templates. Supports custom `.docx` templates to inherit styles.
- **CJK Font Optimization**: Defaults to "Microsoft YaHei" or "SimSun" to ensure proper rendering of Chinese characters in Word documents.
- **Math Support**: Renders LaTeX-style inline math ($...$) and block equations ($$...$$) into native Word equations.
- **Image Embedding**: Supports both local and remote images (automatically downloaded and embedded).
- **Agent Friendly**: Provides `convert_to_bytes` API for in-memory document generation, perfect for AI Agents and Web APIs.
- **Agent Skill**: Includes a standard-compliant Skill package for seamless integration with AI frameworks like OpenClaw.

## Installation

### Via PyPI (Recommended)

```bash
pip install md2docx
```

### From Source

```bash
git clone https://github.com/yourusername/md2docx.git
cd md2docx
pip install .
```

## Usage

### Command Line Interface (CLI)

```bash
md2docx <input_file.md> <output_file.docx> [options]
```

**Examples:**

1. **Basic Conversion**:
   ```bash
   md2docx input.md output.docx
   ```

2. **Use Professional Template**:
   ```bash
   md2docx input.md output.docx --template professional_template.docx
   ```

### Python SDK

```python
from md2docx import MarkdownToDocx

md_content = "# Hello World"
converter = MarkdownToDocx(template_path="professional_template.docx")

# 1. Convert to local file
converter.convert(md_content, "output.docx")

# 2. Convert to in-memory byte stream (ideal for Agents)
docx_bytes = converter.convert_to_bytes(md_content)
```

## Built-in Templates

- `professional_template.docx`: Business professional style, modern sans-serif (Arial + YaHei).
- `academic_template.docx`: Rigorous academic style, serif fonts (Times New Roman + SimSun), 1.5 line spacing.
- `creative_template.docx`: Vibrant creative style, colorful themes (Orange/Blue).

## AI Agent Integration

This repository includes a `md2docx-skill` directory following the Gemini CLI skill standard. AI agents can use this skill to quickly learn how to leverage `md2docx` for document generation.

## License

This project is licensed under the [MIT License](LICENSE).
