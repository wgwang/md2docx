# md2docx

`md2docx` is a Python tool designed to convert Markdown files into elegantly formatted Docx documents. It features support for custom fonts, mathematical equations, tables, and embedded images, with specific optimizations for CJK (Chinese, Japanese, Korean) font support.

## Features

- **CJK Font Optimization**: Defaults to "Microsoft YaHei" to ensure proper rendering of Chinese characters in Word documents.
- **Math Support**: Renders LaTeX-style inline math ($...$) and block equations ($$...$$).
- **Image Embedding**: Supports both local and remote images (automatically downloaded and embedded).
- **Tables**: Supports standard Markdown table syntax.
- **Code Styling**: clear styling for code blocks and inline code.
- **Line Break Control**: Option to preserve soft line breaks from Markdown as hard breaks in the document.

## Installation

### From Source

```bash
git clone https://github.com/yourusername/md2docx.git
cd md2docx
pip install .
```

Or for development:

```bash
pip install -e .
```

## Usage

### Command Line Interface (CLI)

After installation, use the `md2docx` command in your terminal:

```bash
md2docx <input_file.md> <output_file.docx> [options]
```

**Examples:**

1. **Basic Conversion**:
   ```bash
   md2docx input.md output.docx
   ```

2. **Specify Font** (e.g., using Arial):
   ```bash
   md2docx input.md output.docx --font "Arial"
   ```

3. **Preserve Line Breaks** (treat single newlines in Markdown as line breaks in Docx):
   ```bash
   md2docx input.md output.docx --preserve-breaks
   ```

### Python SDK

You can also use it as a library in your Python scripts:

```python
from md2docx import MarkdownToDocx

# Read Markdown content
with open("input.md", "r", encoding="utf-8") as f:
    md_content = f.read()

# Initialize converter
# font_name: Specify the font family (default: "Microsoft YaHei")
# preserve_breaks: Whether to preserve single line breaks (default: False)
converter = MarkdownToDocx(font_name="Microsoft YaHei", preserve_breaks=True)

# Convert
converter.convert(md_content, "output.docx")
```

## Project Structure

```text
md2docx/
├── md2docx/            # Source code
│   ├── __init__.py
│   ├── cli.py          # CLI entry point
│   └── converter.py    # Core conversion logic
├── examples/           # Example files
├── tests/              # Unit tests
├── setup.py            # Installation script
├── requirements.txt    # Dependencies
└── README.md           # Documentation
```

## License

This project is licensed under the [MIT License](LICENSE).
